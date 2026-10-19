# -*- coding: utf-8 -*-
"""Openings: geometria. Sezioni degli elementi MEP, famiglie e inserimento dei fori, area, ricerca degli
attraversamenti con muri, solai e strutture (anche nei link), lettura dei fori esistenti."""
from openings_core import *  # noqa: F401,F403


# ----------------------------------------------------------------------------------------------------------------
# Geometria
# ----------------------------------------------------------------------------------------------------------------
def solids_of(element):
    options = DB.Options()
    options.DetailLevel = DB.ViewDetailLevel.Fine
    result = []
    try:
        pending = list(element.get_Geometry(options) or [])
    except Exception:
        return result
    while pending:
        item = pending.pop()
        if isinstance(item, DB.GeometryInstance):
            pending.extend(item.GetInstanceGeometry())
        elif isinstance(item, DB.Solid) and item.Volume > 1e-9:
            result.append(item)
    return result


def solid_points(solids):
    points = []
    for solid in solids:
        for edge in solid.Edges:
            curve = edge.AsCurve()
            points.extend((curve.GetEndPoint(0), curve.GetEndPoint(1)))
    return points


def unit(vector):
    return vector.Normalize() if vector.GetLength() > 1e-9 else None


# ----------------------------------------------------------------------------------------------------------------
# Modelli letti (host e link RVT caricati) e riferimenti agli elementi MEP (MepRef)
# ----------------------------------------------------------------------------------------------------------------
class Source(object):
    """Modello da cui si leggono elementi: l'host (link_id 0, transform None, key "") o un'istanza di link caricata.
    key: nome del modello nei riferimenti ("PRJ_MEC"; "PRJ_MEC#2" se il link e' inserito piu' volte)."""

    def __init__(self, doc, transform, link_id, key, path):
        self.doc, self.transform, self.link_id, self.key, self.path = doc, transform, link_id, key, path


def doc_path(document):
    """Percorso del file: quello del centrale per i modelli condivisi, altrimenti quello del file."""
    try:
        if document.IsWorkshared:
            central = document.GetWorksharingCentralModelPath()
            if central is not None:
                return DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(central)
    except Exception:
        pass
    return document.PathName or ""


def element_name(element):
    try:
        return element.Name
    except Exception:
        return DB.Element.Name.GetValue(element)


def link_sources():
    """(modelli letti [Source], host per primo; {nome: [istanze]} dei link non caricati; titoli di link diversi
    con lo stesso nome; link caricati ma non letti [Source], tolti in Place: ctx.skipped_links). Letti una volta per
    controllo (reset_models). Chiave = titolo del file linkato, uguale a doc.Title nei modelli MEP / ARC; con piu'
    istanze dello stesso link "<titolo>#<n>", n in ordine di UniqueId."""
    if ctx.sources is None:
        loaded = [Source(ctx.doc, None, 0, "", doc_path(ctx.doc))]
        by_title, unloaded = {}, {}
        for link in DB.FilteredElementCollector(ctx.doc).OfClass(DB.RevitLinkInstance):
            link_doc = link.GetLinkDocument()
            if link_doc is None:
                name = element_name(ctx.doc.GetElement(link.GetTypeId())) or "?"
                if name.lower().endswith(".rvt"):
                    name = name[:-4]
                unloaded.setdefault(name, []).append(id_value(link.Id))
            else:
                by_title.setdefault(link_doc.Title, []).append((link, link_doc))
        clashes, skipped = [], []
        for title in sorted(by_title):
            links = sorted(by_title[title], key=lambda pair: pair[0].UniqueId)
            if len(set(doc_path(d) for _, d in links)) > 1:
                clashes.append(title)  # file diversi con lo stesso nome: i riferimenti non li distinguono
            for number, (link, link_doc) in enumerate(links, 1):
                key = title if len(links) == 1 else "{}#{}".format(title, number)
                source = Source(link_doc, link.GetTotalTransform(), id_value(link.Id), key, doc_path(link_doc))
                (skipped if key in ctx.skipped_links else loaded).append(source)
        ctx.sources = (loaded, unloaded, clashes, skipped)
    return ctx.sources


def loaded_links():
    """[(chiave, percorso)] di tutti i link caricati, anche quelli tolti in Place (per la finestra di Place)."""
    loaded, _, _, skipped = link_sources()
    return sorted((s.key, s.path) for s in loaded[1:] + skipped)


def model_sources():
    """(documento, transform o None, id del link o 0) del modello host e dei link RVT caricati."""
    return [(s.doc, s.transform, s.link_id) for s in link_sources()[0]]


def source_by_key(key):
    for source in link_sources()[0]:
        if source.key == key:
            return source
    return None


def source_by_link(link_id):
    for source in link_sources()[0]:
        if source.link_id == link_id:
            return source
    return None


def models_lines():
    """Righe del report: modelli letti (con il percorso), link non caricati, nomi doppi."""
    loaded, unloaded, clashes, skipped = link_sources()
    lines = ["Models read: {}.".format("; ".join(
        "{} ({})".format(s.key or "this model", s.path or "not saved") for s in loaded))]
    if skipped:
        lines.append("Links not read (unticked in Place): {}.".format(", ".join(sorted(s.key for s in skipped))))
    if unloaded:
        lines.append("Links not loaded (their elements are not read): {}.".format(", ".join(sorted(unloaded))))
    if clashes:
        lines.append("Different linked files with the same name: {}. References use the file name, so they cannot "
                     "tell them apart: rename one of the files.".format(", ".join(clashes)))
    return lines


class MepRef(object):
    """Elemento MEP rettilineo dell'host o di un link, con curva e sezione in coordinate host. text e' il
    riferimento scritto sul foro (ref_text); due MepRef sono uguali se hanno lo stesso text."""

    def __init__(self, element, source):
        self.element, self.source = element, source
        self.doc, self.transform = source.doc, source.transform
        self.text = ref_text(source.key, id_value(element.Id))
        self._curve = None
        self._section = None

    def __eq__(self, other):
        return isinstance(other, MepRef) and other.text == self.text

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash(self.text)

    @property
    def is_host(self):
        return self.source.link_id == 0

    @property
    def IsValidObject(self):
        return self.element.IsValidObject

    @property
    def curve(self):
        if self._curve is None:
            curve = self.element.Location.Curve
            self._curve = curve.CreateTransformed(self.transform) if self.transform is not None else curve
        return self._curve

    @property
    def direction(self):
        return self.curve.Direction

    @property
    def category_name(self):
        return self.element.Category.Name if self.element.Category else "Element"

    def box(self):
        """Bounding box in coordinate host (allineato agli assi host), o None."""
        box = self.element.get_BoundingBox(None)
        if box is None or self.transform is None:
            return box
        corners = [self.transform.OfPoint(DB.XYZ(x, y, z)) for x in (box.Min.X, box.Max.X)
                   for y in (box.Min.Y, box.Max.Y) for z in (box.Min.Z, box.Max.Z)]
        result = DB.BoundingBoxXYZ()
        result.Min = DB.XYZ(min(c.X for c in corners), min(c.Y for c in corners), min(c.Z for c in corners))
        result.Max = DB.XYZ(max(c.X for c in corners), max(c.Y for c in corners), max(c.Z for c in corners))
        return result

    def select_ids(self):
        """Id selezionabili nel modello aperto: solo gli elementi dell'host (quelli dei link non si selezionano per
        Id)."""
        return [id_value(self.element.Id)] if self.is_host else []


def mep_ref(element, source):
    """MepRef dell'elemento, uno solo per riferimento finche' non si rileggono i modelli (sezione calcolata una
    volta)."""
    text = ref_text(source.key, id_value(element.Id))
    ref = ctx.mep_cache.get(text)
    if ref is None:
        ref = ctx.mep_cache[text] = MepRef(element, source)
    return ref


def host_source():
    return link_sources()[0][0]


def mep_by_ref(text):
    """MepRef di un riferimento scritto su un foro, o None se l'elemento non c'e' (cancellato, non piu' MEP, link
    non caricato: vedi link_not_loaded)."""
    model, number = split_ref(text)
    source = source_by_key(model)
    if source is None:
        return None
    try:
        element = source.doc.GetElement(DB.ElementId(int(number)))
    except Exception:
        return None
    return mep_ref(element, source) if isinstance(element, MEP_CLASSES) else None


def link_not_loaded(text):
    """Il riferimento e' di un link non caricato, o non piu' inserito nel modello: l'elemento non si puo' leggere,
    ma non e' detto che sia stato cancellato. Non si toglie dal foro."""
    model, _ = split_ref(text)
    return bool(model) and source_by_key(model) is None


def mep_section(mep):
    """Sezione esterna, isolamento compreso, letta dalla geometria: (asse X, asse Y, semi-larghezza, semi-altezza,
    tonda). Gli assi sono quelli del connettore, perpendicolari al tratto; le misure vengono dai solidi di elemento
    e isolamento, cosi' valgono anche per canali ruotati e passerelle di costa. Misure lette nel documento
    dell'elemento, assi riportati in coordinate host (il transform del link e' rigido). Calcolata una volta per
    MepRef."""
    if mep._section is not None:
        return mep._section
    element, source_doc = mep.element, mep.doc
    conn = [c for c in element.ConnectorManager.Connectors if c.ConnectorType == DB.ConnectorType.End][0]
    bx, by = conn.CoordinateSystem.BasisX, conn.CoordinateSystem.BasisY
    solids = solids_of(element)
    if isinstance(element, (Pipe, Duct)):
        for insulation_id in DB.InsulationLiningBase.GetInsulationIds(source_doc, element.Id):
            solids.extend(solids_of(source_doc.GetElement(insulation_id)))
    origin = element.Location.Curve.GetEndPoint(0)
    points = solid_points(solids)
    if not points:
        raise ValueError("no geometry")
    xs = [p.Subtract(origin).DotProduct(bx) for p in points]
    ys = [p.Subtract(origin).DotProduct(by) for p in points]
    if mep.transform is not None:
        bx, by = mep.transform.OfVector(bx), mep.transform.OfVector(by)
    mep._section = (bx, by, (max(xs) - min(xs)) / 2.0, (max(ys) - min(ys)) / 2.0,
                    conn.Shape == DB.ConnectorProfileType.Round)
    return mep._section


def section_points(center, section):
    bx, by, hx, hy, round_shape = section
    if round_shape:
        r = max(hx, hy)
        return [center.Add(bx.Multiply(r * math.cos(a))).Add(by.Multiply(r * math.sin(a)))
                for a in [i * math.pi / 8.0 for i in range(16)]]
    return [center.Add(bx.Multiply(sx * hx)).Add(by.Multiply(sy * hy)) for sx in (-1, 1) for sy in (-1, 1)]


def make_frame(kind, w, hint):
    """(u, v, w): w attraverso l'elemento, u orizzontale (asse X della famiglia), v la verticale o w x u."""
    if kind == "wall":
        w = unit(DB.XYZ(w.X, w.Y, 0))
        return Z.CrossProduct(w), Z, w
    u = unit(DB.XYZ(hint.X, hint.Y, 0)) if hint is not None else None
    u = u or DB.XYZ.BasisX
    return u, Z.CrossProduct(u), Z


def opening_shape(items, frame):
    """Rettangolo che contiene tutti gli attraversamenti nel piano (u, v) e spessore lungo w. Ritorna (centro,
    estensione u, estensione v, spessore) in piedi."""
    u, v, w = frame
    o = items[0]["mid"]
    pu, pv, pw = [], [], []
    for item in items:
        for p in item["points"]:
            d = p.Subtract(o)
            pu.append(d.DotProduct(u))
            pv.append(d.DotProduct(v))
        for p in (item["p0"], item["p1"]):
            pw.append(p.Subtract(o).DotProduct(w))
    center = o.Add(u.Multiply((max(pu) + min(pu)) / 2.0)).Add(v.Multiply((max(pv) + min(pv)) / 2.0)).Add(
        w.Multiply((max(pw) + min(pw)) / 2.0))
    return center, max(pu) - min(pu), max(pv) - min(pv), max(pw) - min(pw)


def sizes(items, frame):
    """Dati del foro: centro, forma e misure in mm (franco, arrotondamento e sporgenza compresi)."""
    center, ext_u, ext_v, thickness = opening_shape(items, frame)
    clear = 2.0 * ctx.settings["clearance_mm"]
    depth = mm(thickness) + 2.0 * ctx.settings["extra_depth_mm"]
    if len(items) == 1 and items[0]["round"]:
        return {"center": center, "shape": "round", "diameter": round_up(mm(max(ext_u, ext_v)) + clear),
                "depth": depth}
    return {"center": center, "shape": "rect", "width": round_up(mm(ext_u) + clear),
            "height": round_up(mm(ext_v) + clear), "depth": depth}


def crossing_item(mep, section, p0, p1, frame, kind):
    direction = mep.direction
    return {"mep": mep, "p0": p0, "p1": p1, "mid": p0.Add(p1).Divide(2.0), "frame": frame, "kind": kind,
            "points": section_points(p0, section) + section_points(p1, section), "discipline": discipline(mep),
            "round": section[4] and abs(direction.DotProduct(frame[2])) > ROUND_COS, "section": section}


# ----------------------------------------------------------------------------------------------------------------
# Famiglie, fori esistenti, inserimento
# ----------------------------------------------------------------------------------------------------------------
def family_symbols():
    """{(muro|solaio, rect|round): primo tipo della famiglia scelta in Settings}; mancanti: elenco dei nomi."""
    by_name = {}
    for symbol in DB.FilteredElementCollector(ctx.doc).OfClass(DB.FamilySymbol).OfCategory(
            DB.BuiltInCategory.OST_GenericModel):
        by_name.setdefault(symbol.Family.Name, symbol)
    symbols, missing = {}, []
    for key, setting in FAMILY_KEYS.items():
        name = ctx.settings[setting]
        if name in by_name:
            symbols[key] = by_name[name]
        else:
            missing.append(name)
    return symbols, sorted(set(missing))


def existing_openings():
    """Fori gia' gestiti: Generic Model con e_GEN_Void_Status compilato (qualunque famiglia)."""
    return [i for i in DB.FilteredElementCollector(ctx.doc).OfClass(DB.FamilyInstance).OfCategory(
        DB.BuiltInCategory.OST_GenericModel) if get_text(i, ctx.status_param)]


def pick_level(z):
    levels = sorted(DB.FilteredElementCollector(ctx.doc).OfClass(DB.Level), key=level_elevation)
    below = [l for l in levels if level_elevation(l) <= z + 1e-6]
    return below[-1] if below else levels[0]


def create_opening(data, frame, kind, symbols, meps, state, missing):
    """Inserisce il foro con il solido centrato sull'attraversamento (famiglie rettangolari: inserimento sotto il
    centro di meta' altezza), con l'asse X della famiglia lungo u (muro circolare: lungo w, l'asse del cilindro);
    scrive misure, stato, Id e categorie dei MEP."""
    symbol = symbols[(kind, data["shape"])]
    if not symbol.IsActive:
        symbol.Activate()
        ctx.doc.Regenerate()
    length = LENGTH_PARAMS[(kind, data["shape"])]
    values = ((PARAM_DIAMETER, data["diameter"]), (length, data["depth"])) if data["shape"] == "round" else \
        ((PARAM_WIDTH, data["width"]), (PARAM_HEIGHT, data["height"]), (length, data["depth"]))
    base = BASE_PARAMS.get((kind, data["shape"]))
    lift = dict(values)[base] / 2.0 / MM_PER_FT if base else 0.0
    point = data["center"].Subtract(Z.Multiply(lift))
    instance = ctx.doc.Create.NewFamilyInstance(point, symbol, pick_level(point.Z), StructuralType.NonStructural)
    ctx.doc.Regenerate()
    current = instance.GetTransform().BasisX
    u = frame[2] if (kind, data["shape"]) == ("wall", "round") else frame[0]
    angle = math.atan2(current.CrossProduct(u).Z, current.DotProduct(u))
    if abs(angle) > 1e-9:
        axis = DB.Line.CreateBound(instance.Location.Point, instance.Location.Point.Add(Z))
        DB.ElementTransformUtils.RotateElement(ctx.doc, instance.Id, axis, angle)
    delta = point.Subtract(instance.Location.Point)
    if delta.GetLength() > 1e-9:
        DB.ElementTransformUtils.MoveElement(ctx.doc, instance.Id, delta)
    for name, value in values:
        param = instance.LookupParameter(name)
        if param is None or param.IsReadOnly:
            missing.add("{} on family {}".format(name, symbol.Family.Name))
        else:
            param.Set(value / MM_PER_FT)
    write_opening(instance, meps, state, missing)
    return instance


def is_straight_mep(element):
    """Tubo, canale, passerella o conduit rettilineo (elemento Revit o MepRef)."""
    element = element_of(element)
    return isinstance(element, MEP_CLASSES) and isinstance(element.Location.Curve, DB.Line)


def meps_in_outline(outline_of):
    """MepRef degli elementi MEP rettilinei, dell'host e dei link caricati, che toccano l'outline. outline_of(transform
    del modello o None) -> DB.Outline nelle coordinate di quel modello."""
    categories = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory](MEP_CATEGORIES))
    found = []
    for source in link_sources()[0]:
        for element in DB.FilteredElementCollector(source.doc).WherePasses(categories)\
                .WherePasses(DB.BoundingBoxIntersectsFilter(outline_of(source.transform)))\
                .WhereElementIsNotElementType():
            if is_straight_mep(element):
                found.append(mep_ref(element, source))
    return found


def outline_in(low, high, transform):
    """Outline (assi del modello) che contiene il box host low-high, portato nel modello con transform."""
    corners = [DB.XYZ(x, y, z) for x in (low.X, high.X) for y in (low.Y, high.Y) for z in (low.Z, high.Z)]
    if transform is not None:
        inverse = transform.Inverse
        corners = [inverse.OfPoint(c) for c in corners]
    return DB.Outline(DB.XYZ(min(c.X for c in corners), min(c.Y for c in corners), min(c.Z for c in corners)),
                      DB.XYZ(max(c.X for c in corners), max(c.Y for c in corners), max(c.Z for c in corners)))


def hosts_from_meps(meps):
    """Modalita' inversa: muri, solai e strutture del modello host e dei link caricati che toccano il bounding box
    di almeno un elemento MEP scelto. Stesso formato di pick_hosts."""
    categories = List[DB.BuiltInCategory]([DB.BuiltInCategory.OST_Walls, DB.BuiltInCategory.OST_Floors,
                                           DB.BuiltInCategory.OST_StructuralFraming,
                                           DB.BuiltInCategory.OST_StructuralColumns,
                                           DB.BuiltInCategory.OST_StructuralFoundation])
    boxes = [m.box() for m in meps]  # coordinate host
    boxes = [b for b in boxes if b is not None]
    hosts, seen = [], set()
    for source_doc, transform, link_id in model_sources():
        inverse = transform.Inverse if transform is not None else None
        model = source_doc.Title if link_id else "Host model"
        for box in boxes:
            corners = [DB.XYZ(x, y, z) for x in (box.Min.X, box.Max.X) for y in (box.Min.Y, box.Max.Y)
                       for z in (box.Min.Z, box.Max.Z)]
            if inverse is not None:
                corners = [inverse.OfPoint(c) for c in corners]
            outline = DB.Outline(DB.XYZ(min(c.X for c in corners), min(c.Y for c in corners),
                                        min(c.Z for c in corners)),
                                 DB.XYZ(max(c.X for c in corners), max(c.Y for c in corners),
                                        max(c.Z for c in corners)))
            for element in DB.FilteredElementCollector(source_doc).WherePasses(
                    DB.ElementMulticategoryFilter(categories)).WhereElementIsNotElementType()\
                    .WherePasses(DB.BoundingBoxIntersectsFilter(outline)):
                key = (link_id, id_value(element.Id))
                if key not in seen:
                    seen.add(key)
                    hosts.append((key, element, transform, model))
    return hosts


# ----------------------------------------------------------------------------------------------------------------
# Area: riquadro trascinato in pianta (quote dal view range, come Quick SectionBox) o section box della vista 3D
# ----------------------------------------------------------------------------------------------------------------
MIN_AREA_MM = 10.0           # riquadro piu' stretto di cosi': e' un clic, non un trascinamento
FALLBACK_HEIGHT_MM = 4000.0  # altezza se il view range non e' leggibile (livelli illimitati)
PLAN_TYPES = (DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.EngineeringPlan)
HOST_CATEGORIES = (DB.BuiltInCategory.OST_Walls, DB.BuiltInCategory.OST_Floors,
                   DB.BuiltInCategory.OST_StructuralFraming, DB.BuiltInCategory.OST_StructuralColumns,
                   DB.BuiltInCategory.OST_StructuralFoundation)


def flat_unit(vector):
    flat = DB.XYZ(vector.X, vector.Y, 0)
    return flat.Normalize() if flat.GetLength() > 1e-9 else None


def area_axes(area):
    right = DB.XYZ(area["right"][0], area["right"][1], 0)
    return right, Z.CrossProduct(right)


def make_area(right, points, z_min, z_max, source):
    """Area allineata agli assi della vista: min / max = (destra, alto, quota) dei punti proiettati. Solo numeri: si
    salva cosi' com'e' nel file del Last Report e si rilegge per il Refresh."""
    up = Z.CrossProduct(right)
    us = [p.DotProduct(right) for p in points]
    vs = [p.DotProduct(up) for p in points]
    return {"right": [right.X, right.Y], "min": [min(us), min(vs), z_min], "max": [max(us), max(vs), z_max],
            "source": source}


def area_hit(area, point, z_low, z_high):
    """Il punto sta nel rettangolo in pianta e l'intervallo di quote [z_low, z_high] tocca quello dell'area."""
    right, up = area_axes(area)
    tol = MOVE_TOL_MM / MM_PER_FT
    u, v = point.DotProduct(right), point.DotProduct(up)
    return area["min"][0] - tol <= u <= area["max"][0] + tol and area["min"][1] - tol <= v <= area["max"][1] + tol \
        and z_high >= area["min"][2] - tol and z_low <= area["max"][2] + tol


def crossing_in_area(area, p0, p1):
    """Attraversamento nell'area: punto medio nel rettangolo, tratto che tocca le quote. Il tratto e non il solo punto
    medio: con il view range consueto (fondo al livello) il solaio sotto il livello ha il centro fuori quota."""
    return area_hit(area, p0.Add(p1).Divide(2.0), min(p0.Z, p1.Z), max(p0.Z, p1.Z))


def opening_in_area(area, opening):
    box = opening.get_BoundingBox(None)
    if box is None:
        return False
    return area_hit(area, box.Min.Add(box.Max).Divide(2.0), box.Min.Z, box.Max.Z)


def area_outline(area, transform=None):
    """Outline allineato agli assi del documento (del link, con transform) che contiene l'area: filtro grossolano per
    i collector, il test vero e' area_hit."""
    right, up = area_axes(area)
    inverse = transform.Inverse if transform is not None else None
    corners = []
    for u in (area["min"][0], area["max"][0]):
        for v in (area["min"][1], area["max"][1]):
            for z in (area["min"][2], area["max"][2]):
                corner = right.Multiply(u).Add(up.Multiply(v)).Add(Z.Multiply(z))
                corners.append(inverse.OfPoint(corner) if inverse is not None else corner)
    return DB.Outline(DB.XYZ(min(c.X for c in corners), min(c.Y for c in corners), min(c.Z for c in corners)),
                      DB.XYZ(max(c.X for c in corners), max(c.Y for c in corners), max(c.Z for c in corners)))


def area_text(area):
    return "{}, {:.0f} x {:.0f} mm, height {:.0f} to {:.0f} mm".format(
        area["source"], (area["max"][0] - area["min"][0]) * MM_PER_FT, (area["max"][1] - area["min"][1]) * MM_PER_FT,
        area["min"][2] * MM_PER_FT, area["max"][2] * MM_PER_FT)


def hosts_in_area(area):
    """Muri, solai e strutture del modello host e dei link caricati che toccano l'area. Stesso formato di
    pick_hosts."""
    categories = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory](HOST_CATEGORIES))
    hosts = []
    for source_doc, transform, link_id in model_sources():
        model = source_doc.Title if link_id else "Host model"
        for element in DB.FilteredElementCollector(source_doc).WherePasses(categories).WhereElementIsNotElementType()\
                .WherePasses(DB.BoundingBoxIntersectsFilter(area_outline(area, transform))):
            hosts.append(((link_id, id_value(element.Id)), element, transform, model))
    return hosts


def meps_in_area(area):
    """Elementi MEP rettilinei (MepRef) dell'host e dei link caricati che toccano l'area (anche solo in parte)."""
    return meps_in_outline(lambda transform: area_outline(area, transform))


def host_normal(element, transform, point):
    """Normale del muro (o della trave) nel punto, anche curvo, in coordinate host."""
    curve = element.Location.Curve
    if transform is not None:
        curve = curve.CreateTransformed(transform)
    projection = curve.Project(point)
    tangent = curve.ComputeDerivatives(projection.Parameter, False).BasisX
    return Z.CrossProduct(DB.XYZ(tangent.X, tangent.Y, 0).Normalize())


def crossing_normal(element, transform, point, kind, direction):
    """Asse di attraversamento del foro: solaio, la verticale; muro o trave, la normale nel punto; pilastri e casi
    senza asse orizzontale leggibile, la direzione dell'elemento MEP in pianta."""
    if kind == "floor":
        return Z
    try:
        return host_normal(element, transform, point)
    except Exception:
        return DB.XYZ(direction.X, direction.Y, 0).Normalize()


def mep_near(solids):
    """Elementi MEP rettilinei (MepRef) dell'host e dei link caricati vicini ai solidi (filtro sul bounding box;
    solidi in coordinate host)."""
    points = solid_points(solids)
    low = DB.XYZ(min(p.X for p in points), min(p.Y for p in points), min(p.Z for p in points))
    high = DB.XYZ(max(p.X for p in points), max(p.Y for p in points), max(p.Z for p in points))
    return meps_in_outline(lambda transform: outline_in(low, high, transform))


def inside_existing(point, boxes):
    """Foro gestito il cui bounding box contiene il punto, o None."""
    for low, high, opening in boxes:
        if low.X <= point.X <= high.X and low.Y <= point.Y <= high.Y and low.Z <= point.Z <= high.Z:
            return opening
    return None


ADJACENT_SEARCH = 1000.0 / MM_PER_FT  # margine attorno all'elemento scelto in cui si cercano gli strati vicini
MAX_LAYER = 1000.0 / MM_PER_FT  # strato attraversato per piu' di cosi': l'elemento MEP ci corre dentro, non conta
PARALLEL_COS = 0.98             # strati di muro paralleli all'elemento scelto entro circa 11 gradi


def neighbours(solids, kind):
    """Muri (o solai, per i fori a solaio) del modello host e dei link vicini all'elemento scelto: gli strati di una
    struttura composita modellati come elementi separati. Ritorna [(chiave, etichetta, solidi in coordinate host,
    elemento, transform)]."""
    points = solid_points(solids)
    low = DB.XYZ(min(p.X for p in points) - ADJACENT_SEARCH, min(p.Y for p in points) - ADJACENT_SEARCH,
                 min(p.Z for p in points) - ADJACENT_SEARCH)
    high = DB.XYZ(max(p.X for p in points) + ADJACENT_SEARCH, max(p.Y for p in points) + ADJACENT_SEARCH,
                  max(p.Z for p in points) + ADJACENT_SEARCH)
    category = DB.BuiltInCategory.OST_Walls if kind == "wall" else DB.BuiltInCategory.OST_Floors
    found = []
    for source_doc, transform, link_id in model_sources():
        corners = [DB.XYZ(x, y, z) for x in (low.X, high.X) for y in (low.Y, high.Y) for z in (low.Z, high.Z)]
        if transform is not None:
            inverse = transform.Inverse
            corners = [inverse.OfPoint(c) for c in corners]
        outline = DB.Outline(DB.XYZ(min(c.X for c in corners), min(c.Y for c in corners), min(c.Z for c in corners)),
                             DB.XYZ(max(c.X for c in corners), max(c.Y for c in corners), max(c.Z for c in corners)))
        for element in DB.FilteredElementCollector(source_doc).OfCategory(category).WhereElementIsNotElementType()\
                .WherePasses(DB.BoundingBoxIntersectsFilter(outline)):
            key = (link_id, id_value(element.Id))
            if key not in ctx.neighbour_cache:
                own = solids_of(element)
                ctx.neighbour_cache[key] = [DB.SolidUtils.CreateTransformed(s, transform) for s in own] \
                    if transform is not None else own
            label = "{} {} ({})".format(element.Category.Name, id_value(element.Id),
                                        source_doc.Title if link_id else "Host model")
            found.append((key, label, ctx.neighbour_cache[key], element, transform))
    return found


def parallel_layer(element, transform, kind, w, point):
    """Lo strato e' parallelo all'elemento scelto (muri: normale nel punto quasi parallela a w; solai: sempre). Un
    muro a T o d'angolo non e' uno strato della stessa struttura."""
    if kind != "wall":
        return True
    try:
        return abs(host_normal(element, transform, point).DotProduct(w)) > PARALLEL_COS
    except Exception:
        return True  # muro senza linea di posizione leggibile: si giudica solo dalla distanza


def extend_through(p0, p1, layers, options, kind, w):
    """Allunga l'attraversamento attraverso tutti gli strati della struttura composita (elementi separati): ogni
    strato parallelo che, lungo l'asse dell'elemento MEP, sta entro l'intercapedine massima di Settings da quanto
    gia' attraversato viene incluso, intercapedine compresa. Si ripete finche' non se ne trovano altri. Ritorna
    (p0, p1, chiavi degli strati attraversati)."""
    gap = max(ctx.settings["layer_gap_mm"], 5.0) / MM_PER_FT
    direction = p1.Subtract(p0).Normalize()
    t0, t1 = 0.0, p1.DistanceTo(p0)
    used = set()
    changed = True
    while changed:
        changed = False
        probe = DB.Line.CreateBound(p0.Add(direction.Multiply(t0 - gap - MAX_LAYER)),
                                    p0.Add(direction.Multiply(t1 + gap + MAX_LAYER)))
        for key, label, solids, element, transform in layers:
            for solid in solids:
                try:
                    result = solid.IntersectWithCurve(probe, options)
                except Exception:
                    continue
                for i in range(result.SegmentCount):
                    segment = result.GetCurveSegment(i)
                    a = segment.GetEndPoint(0).Subtract(p0).DotProduct(direction)
                    b = segment.GetEndPoint(1).Subtract(p0).DotProduct(direction)
                    a, b = min(a, b), max(a, b)
                    if b - a < 1e-6 or b - a > MAX_LAYER or b < t0 - gap or a > t1 + gap:
                        continue  # troppo lontano da quanto gia' attraversato, o l'elemento MEP ci corre dentro
                    if key not in used:
                        middle = p0.Add(direction.Multiply((a + b) / 2.0))
                        if not parallel_layer(element, transform, kind, w, middle):
                            continue
                    used.add(key)
                    if a < t0 - 1e-6 or b > t1 + 1e-6:
                        t0, t1 = min(t0, a), max(t1, b)
                        changed = True
    return p0.Add(direction.Multiply(t0)), p0.Add(direction.Multiply(t1)), used


def find_crossings(hosts, stats, only=None, keep_existing=False, progress=None, area=None):
    """Attraversamenti degli elementi MEP con muri, solai e strutture, allungati attraverso gli strati a contatto
    (muri o solai separati di una struttura composita). only: riferimenti (MepRef.text) degli elementi MEP da
    considerare (modalita' inversa), None = tutti quelli vicini, dell'host e dei link. Quelli che hanno gia' un
    foro si saltano (Place) o, con keep_existing (check), restano con item["opening"] = il foro che li contiene.
    progress(i, n): avanzamento sugli elementi attraversati. area: solo gli attraversamenti dentro l'area (crossing_in_area), gli altri contati in stats["outside_area"]."""
    boxes = []
    for opening in existing_openings():
        box = opening.get_BoundingBox(None)
        if box is not None:
            boxes.append((box.Min, box.Max, opening))
    items = []
    options = DB.SolidCurveIntersectionOptions()
    for index, (key, element, transform, model) in enumerate(hosts):
        if progress is not None:
            progress(index, len(hosts))
        host_kind = HOST_KINDS[id_value(element.Category.Id)]
        solids = solids_of(element)
        if transform is not None:
            solids = [DB.SolidUtils.CreateTransformed(s, transform) for s in solids]
        if not solids:
            stats["problems"].append("{} {} ({}): no geometry".format(element.Category.Name,
                                                                       id_value(element.Id), model))
            continue
        # strati separati di una struttura composita: solo per muri e solai
        layers = neighbours(solids, host_kind) if host_kind != "auto" else []
        labels = dict((layer[0], layer[1]) for layer in layers)
        for mep in mep_near(solids):
            if only is not None and mep.text not in only:
                continue
            line = mep.curve
            try:
                section = mep_section(mep)
            except Exception:
                stats["problems"].append("{} {}: section not readable".format(mep.category_name, mep.text))
                continue
            for solid in solids:
                result = solid.IntersectWithCurve(line, options)
                for i in range(result.SegmentCount):
                    segment = result.GetCurveSegment(i)
                    if mm(segment.Length) < 1.0:
                        continue
                    p0, p1 = segment.GetEndPoint(0), segment.GetEndPoint(1)
                    mid = p0.Add(p1).Divide(2.0)
                    if host_kind == "auto":
                        # strutture: foro a solaio per elementi MEP verticali, a muro per gli altri; oltre 1 m
                        # dentro la struttura l'elemento ci corre dentro (es. montante in un pilastro)
                        if segment.Length > MAX_LAYER:
                            stats["inside"] += 1
                            continue
                        kind = "floor" if abs(line.Direction.Z) > VERTICAL_SIN else "wall"
                    else:
                        kind = host_kind
                    w = crossing_normal(element, transform, mid, kind, line.Direction)
                    if abs(line.Direction.DotProduct(w)) < ctx.min_cross_sin:
                        stats["inside"] += 1  # corre dentro il muro o il solaio: nessun foro
                        continue
                    p0, p1, used = extend_through(p0, p1, layers, options, kind, w)
                    if area is not None and not crossing_in_area(area, p0, p1):
                        stats["outside_area"] = stats.get("outside_area", 0) + 1
                        continue
                    mid = p0.Add(p1).Divide(2.0)
                    opening = inside_existing(mid, boxes)
                    if opening is not None:
                        stats["existing"] += 1
                        if not keep_existing:
                            continue
                    item = crossing_item(mep, section, p0, p1, make_frame(kind, w, section[0]), kind)
                    item["opening"] = opening
                    # chiave dell'intera pila di strati: scegliendo piu' strati della stessa struttura, i loro
                    # attraversamenti finiscono nello stesso foro invece di dare fori sovrapposti
                    stack = tuple(sorted(used | set([key])))
                    others = [labels[k] for k in sorted(used) if k != key and k in labels]
                    if others:
                        stats["extended"] += 1
                    item["host"] = (stack, element, model + (" + " + ", ".join(others) if others else ""))
                    item["host_kind"] = host_kind
                    items.append(item)
    return items


def merge_overlapping(items):
    """Attraversamenti dello stesso elemento MEP attraverso elementi diversi che si sovrappongono o si toccano lungo
    l'elemento (es. trave annegata nel muro, pilastro nel muro, strato di un'altra categoria): diventano uno solo, con
    l'intervallo complessivo e tutti gli elementi attraversati. Senza questo, ogni elemento darebbe un suo foro, quasi
    coincidente con l'altro. Vale per stesso tipo di foro (muro / solaio) e assi di attraversamento paralleli."""
    gap = max(ctx.settings["layer_gap_mm"], 5.0) / MM_PER_FT
    merged = []
    for item in items:
        direction = item["mep"].direction
        target = None
        for other in merged:
            if other["mep"] != item["mep"] or other["kind"] != item["kind"]:
                continue
            if abs(other["frame"][2].DotProduct(item["frame"][2])) < PARALLEL_COS:
                continue
            a = sorted((other["p0"].DotProduct(direction), other["p1"].DotProduct(direction)))
            b = sorted((item["p0"].DotProduct(direction), item["p1"].DotProduct(direction)))
            if b[0] <= a[1] + gap and a[0] <= b[1] + gap:
                target = other
                break
        if target is None:
            merged.append(dict(item))
            continue
        ends = [target["p0"], target["p1"], item["p0"], item["p1"]]
        p0 = min(ends, key=lambda p: p.DotProduct(direction))
        p1 = max(ends, key=lambda p: p.DotProduct(direction))
        # il sistema del foro viene da un muro o solaio, se c'e' (le strutture lo hanno solo approssimato)
        keep = target if target.get("host_kind") != "auto" or item.get("host_kind") == "auto" else item
        frame, kind = keep["frame"], keep["kind"]
        stack = tuple(sorted(set(target["host"][0]) | set(item["host"][0])))
        label = target["host"][2] + " + " + "{} {} ({})".format(item["host"][1].Category.Name,
                                                                id_value(item["host"][1].Id), item["host"][2])
        target.update(crossing_item(item["mep"], item["section"], p0, p1, frame, kind))
        target["host"] = (stack, keep["host"][1], label)
        target["host_kind"] = keep.get("host_kind")
    return merged


def clusters(items):
    """Gruppi di attraversamenti dello stesso elemento, paralleli, a distanza <= merge nel piano del foro (e della
    stessa disciplina, se MEC ed ELE sono separati)."""
    # distanza tra i fori finiti (franco compreso), non tra le sezioni nude: due fori che si toccherebbero o
    # sovrapporrebbero si uniscono sempre
    merge = (ctx.settings["merge_mm"] + 2.0 * ctx.settings["clearance_mm"]) / MM_PER_FT
    parent = list(range(len(items)))

    def root(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def rect(item, frame):
        u, v, _ = frame
        pu = [p.DotProduct(u) for p in item["points"]]
        pv = [p.DotProduct(v) for p in item["points"]]
        return min(pu), max(pu), min(pv), max(pv)

    for i in range(len(items)):
        for j in range(i + 1, len(items)):
            a, b = items[i], items[j]
            # stessa struttura: almeno un elemento attraversato in comune (pila di strati, muro + trave...)
            if not set(a["host"][0]) & set(b["host"][0]) or a["kind"] != b["kind"]:
                continue
            if ctx.settings["separate"] and a["discipline"] != b["discipline"]:
                continue
            if abs(a["frame"][2].DotProduct(b["frame"][2])) < 0.999:
                continue
            ra, rb = rect(a, a["frame"]), rect(b, a["frame"])
            gap = max(rb[0] - ra[1], ra[0] - rb[1], rb[2] - ra[3], ra[2] - rb[3])
            if gap <= merge:
                parent[root(j)] = root(i)
    groups = {}
    for i, item in enumerate(items):
        groups.setdefault(root(i), []).append(item)
    return list(groups.values())


# ----------------------------------------------------------------------------------------------------------------
# CHECK & REPORT
# ----------------------------------------------------------------------------------------------------------------
def read_length(instance, name):
    param = instance.LookupParameter(name)
    return mm(param.AsDouble()) if param is not None and param.HasValue else None


def opening_kind(instance, meps):
    """(muro|solaio, rect|round): dalla famiglia, se e' una di quelle di Settings; altrimenti dai parametri (tondo
    se ha il diametro) e dalla direzione dei suoi elementi MEP (verticali: solaio)."""
    name = instance.Symbol.Family.Name
    for key, setting in FAMILY_KEYS.items():
        if ctx.settings[setting] == name:
            return key
    shape = "round" if read_length(instance, PARAM_DIAMETER) else "rect"
    vertical = any(abs(m.direction.Z) > VERTICAL_SIN for m in meps)
    return ("floor" if vertical else "wall"), shape


def opening_frame(instance, kind, shape):
    x = instance.GetTransform().BasisX
    if kind == "wall" and shape == "round":
        return Z.CrossProduct(x), Z, x  # cilindro lungo X: X e' l'asse di attraversamento
    return (x, Z, Z.CrossProduct(x)) if kind == "wall" else (x, Z.CrossProduct(x), Z)


def opening_center(instance, kind, shape):
    """Centro del solido: il punto di inserimento, alzato di meta' altezza per le famiglie che poggiano sul livello
    di riferimento (BASE_PARAMS)."""
    point = instance.Location.Point
    base = BASE_PARAMS.get((kind, shape))
    if base:
        point = point.Add(Z.Multiply((read_length(instance, base) or 0.0) / 2.0 / MM_PER_FT))
    return point


def current_data(instance, kind, shape):
    data = {"center": opening_center(instance, kind, shape), "shape": shape,
            "depth": read_length(instance, LENGTH_PARAMS[(kind, shape)]) or 0.0}
    if shape == "round":
        data["diameter"] = read_length(instance, PARAM_DIAMETER) or 0.0
    else:
        data["width"] = read_length(instance, PARAM_WIDTH) or 0.0
        data["height"] = read_length(instance, PARAM_HEIGHT) or 0.0
    return data


def half_sizes(data):
    if data["shape"] == "round":
        return data["diameter"] / 2.0, data["diameter"] / 2.0
    return data["width"] / 2.0, data["height"] / 2.0


def recross(mep, frame, data, limit_reach=True):
    """Attraversamento dell'elemento MEP con lo spessore del foro (il foro fa da muro: piano per il centro,
    perpendicolare a w, spessore = profondita'). None se non lo attraversa piu' o (limit_reach) se e' uscito dal foro
    oltre la distanza di unione. mep: MepRef (curva in coordinate host)."""
    line = mep.curve
    if not isinstance(line, DB.Line):
        return None
    u, v, w = frame
    direction = line.Direction
    cos = direction.DotProduct(w)
    if abs(cos) < ctx.min_cross_sin:
        return None
    start = line.GetEndPoint(0)
    half = data["depth"] / MM_PER_FT / 2.0
    base = data["center"].Subtract(start).DotProduct(w)
    t0, t1 = sorted(((base - half) / cos, (base + half) / cos))
    if t1 < -1e-6 or t0 > line.Length + 1e-6:
        return None  # il tratto non arriva piu' al foro
    p0, p1 = start.Add(direction.Multiply(t0)), start.Add(direction.Multiply(t1))
    mid = p0.Add(p1).Divide(2.0).Subtract(data["center"])
    reach_u, reach_v = [(h + ctx.settings["merge_mm"]) / MM_PER_FT for h in half_sizes(data)]
    if limit_reach and (abs(mid.DotProduct(u)) > reach_u or abs(mid.DotProduct(v)) > reach_v):
        return None  # spostato fuori dal foro
    return p0, p1


def inside_opening(item, frame, data):
    """La sezione dell'elemento (isolamento compreso) sta tutta dentro il foro attuale."""
    u, v, _ = frame
    if data["shape"] == "round":
        radius = data["diameter"] / 2.0 / MM_PER_FT
        return all(math.hypot(p.Subtract(data["center"]).DotProduct(u), p.Subtract(data["center"]).DotProduct(v))
                   <= radius + 1e-6 for p in item["points"])
    hu, hv = [h / MM_PER_FT for h in half_sizes(data)]
    return all(abs(p.Subtract(data["center"]).DotProduct(u)) <= hu + 1e-6 and
               abs(p.Subtract(data["center"]).DotProduct(v)) <= hv + 1e-6 for p in item["points"])


def overlaps_opening(item, frame, data):
    """La sezione dell'elemento tocca almeno in parte il foro attuale (altrimenti e' uscita del tutto, anche se e'
    ancora entro la distanza di unione)."""
    u, v, _ = frame
    center = data["center"]
    pu = [p.Subtract(center).DotProduct(u) for p in item["points"]]
    pv = [p.Subtract(center).DotProduct(v) for p in item["points"]]
    if data["shape"] == "round":
        # distanza dal centro del rettangolo che contiene la sezione
        du = 0.0 if min(pu) <= 0.0 <= max(pu) else min(abs(min(pu)), abs(max(pu)))
        dv = 0.0 if min(pv) <= 0.0 <= max(pv) else min(abs(min(pv)), abs(max(pv)))
        return math.hypot(du, dv) < data["diameter"] / 2.0 / MM_PER_FT
    hu, hv = [h / MM_PER_FT for h in half_sizes(data)]
    return min(pu) < hu and max(pu) > -hu and min(pv) < hv and max(pv) > -hv


def differences(old, new):
    notes = []
    if old["shape"] != new["shape"]:
        notes.append("shape {} -> {}".format(old["shape"], new["shape"]))
    moved = mm(old["center"].DistanceTo(new["center"]))
    if moved > MOVE_TOL_MM:
        notes.append("moved {:.0f} mm".format(moved))
    if old["shape"] == new["shape"]:
        keys = ("diameter",) if new["shape"] == "round" else ("width", "height")
        for key in keys:
            if abs(old[key] - new[key]) > SIZE_TOL_MM:
                notes.append("{} {:.0f} -> {:.0f}".format(key, old[key], new[key]))
    if abs(old["depth"] - new["depth"]) > SIZE_TOL_MM:
        notes.append("depth {:.0f} -> {:.0f} (walls / floors crossed changed)".format(old["depth"], new["depth"]))
    return notes


def scanned_plane(meps, frame, old, scanned):
    """Piano del foro (centro, profondita' in mm) dagli attraversamenti della scansione del check: per i suoi
    elementi, quelli paralleli al foro che ne toccano lo spessore, gia' allungati da find_crossings attraverso gli
    strati entro l'intercapedine di Settings (es. un muro nuovo accanto). Senza dati: il piano attuale."""
    if not scanned:
        return old["center"], old["depth"]
    w = frame[2]
    half = old["depth"] / MM_PER_FT / 2.0
    tol = MOVE_TOL_MM / MM_PER_FT
    lows, highs = [], []
    for mep in meps:
        for item in scanned.get(mep.text, []):
            if abs(item["frame"][2].DotProduct(w)) < 0.999:
                continue
            a, b = sorted((item["p0"].Subtract(old["center"]).DotProduct(w),
                           item["p1"].Subtract(old["center"]).DotProduct(w)))
            if b < -half - tol or a > half + tol:
                continue  # un altro muro o solaio: e' un foro mancante, non questo
            lows.append(a)
            highs.append(b)
    if not lows:
        return old["center"], old["depth"]
    extra = ctx.settings["extra_depth_mm"] / MM_PER_FT
    low, high = min(lows) - extra, max(highs) + extra
    return old["center"].Add(w.Multiply((low + high) / 2.0)), mm(high - low)


def rebuilt(items, frame, center, depth):
    """Foro ricalcolato dagli attraversamenti: misure nel piano dagli elementi, centro sul piano per center
    (perpendicolare a w), profondita' data (quella attuale, o quella della scansione: scanned_plane)."""
    new = sizes(items, frame)
    new["center"] = new["center"].Subtract(frame[2].Multiply(new["center"].Subtract(center).DotProduct(frame[2])))
    new["depth"] = depth
    return new


def opening_on_hosts(instance, host_solids):
    """Il foro sta su uno dei muri/solai scelti: una linea per il suo centro, lungo X, Y e Z della famiglia (gli
    assi di attraversamento: Y muro rettangolare, X muro circolare, Z solaio), entra nel solido dell'elemento."""
    transform = instance.GetTransform()
    # centro del bounding box = centro del solido, anche per le famiglie che poggiano sul livello di riferimento
    box = instance.get_BoundingBox(None)
    center = box.Min.Add(box.Max).Divide(2.0) if box is not None else instance.Location.Point
    lengths = [read_length(instance, name) or 0.0 for name in set(LENGTH_PARAMS.values())]
    reach = max(max(lengths) / MM_PER_FT, 1.0)
    options = DB.SolidCurveIntersectionOptions()
    for axis in (transform.BasisX, transform.BasisY, transform.BasisZ):
        line = DB.Line.CreateBound(center.Subtract(axis.Multiply(reach)), center.Add(axis.Multiply(reach)))
        for solid in host_solids:
            try:
                if solid.IntersectWithCurve(line, options).SegmentCount > 0:
                    return True
            except Exception:
                continue
    return False


def openings_on_hosts(openings, hosts):
    solids = []
    for key, element, transform, model in hosts:
        for solid in solids_of(element):
            solids.append(DB.SolidUtils.CreateTransformed(solid, transform) if transform is not None else solid)
    return [o for o in openings if opening_on_hosts(o, solids)]


def opening_shift(items, frame, data, single):
    """Spostamento del foro (nel suo piano, senza cambiarne le misure) che lo riporta sui suoi elementi: centrato su
    di essi se il foro e' di un solo elemento, tondo, o se gli elementi non ci stanno comunque; altrimenti il minimo
    perche' tutti rientrino con il franco."""
    u, v, _ = frame
    center = data["center"]
    points = [p for item in items for p in item["points"]]
    pu = [p.Subtract(center).DotProduct(u) for p in points]
    pv = [p.Subtract(center).DotProduct(v) for p in points]
    if single or data["shape"] == "round":
        return u.Multiply((max(pu) + min(pu)) / 2.0).Add(v.Multiply((max(pv) + min(pv)) / 2.0))
    clear = ctx.settings["clearance_mm"] / MM_PER_FT
    shifts = []
    for values, half in ((pu, data["width"] / 2.0 / MM_PER_FT), (pv, data["height"] / 2.0 / MM_PER_FT)):
        room = max(half - clear, 0.0)
        low, high = min(values), max(values)
        if high - low > 2.0 * room:
            shifts.append((high + low) / 2.0)
        elif high > room:
            shifts.append(high - room)
        elif low < -room:
            shifts.append(low + room)
        else:
            shifts.append(0.0)
    return u.Multiply(shifts[0]).Add(v.Multiply(shifts[1]))
