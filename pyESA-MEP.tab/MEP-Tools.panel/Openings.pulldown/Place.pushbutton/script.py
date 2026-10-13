# -*- coding: utf-8 -*-
__title__ = "Place"
__doc__ = """Version = 3.0
Date    = 13.10.2026
________________________________________________________________
Description:
Openings (provision for void, PFV). Place and Refresh share the same
script; only MODE changes. Settings (families, clearance, rounding,
merge distance, extra depth, MEC/ELE, Misaligned) come from the
Settings button of this menu.

PLACE: select walls and floors, in the host model or in RVT links, then
Finish. Every straight pipe, duct, cable tray or conduit of the host
model that crosses them gets a Generic Model opening at the centre of
the crossing, sized with the outer section (insulation included) plus
the clearance, rounded up; the depth is the thickness plus the extra
depth on each side. Crossings of the same wall or floor closer than the
merge distance share one rectangular opening (only within the same
discipline, if MEC and ELE are kept separate).
- Opening: e_GEN_PFV_ID = Ids of its MEP elements (111,222);
  e_GEN_PFV_Status = <discipline> - New (MEC, ELE or MEC+ELE)
- MEP elements: e_GEN_PFV_ID = codes of their openings, PFV + Id of the
  opening (PFV123456, comma separated); e_GEN_PFV_Status = <MEC|ELE> - New
Comments and Mark are not touched. Crossings that already have an
opening are skipped.

REFRESH: quick check of the openings and MEP elements with
e_GEN_PFV_Status filled: are the MEP elements still there, still
crossing, in the same place and size? Statuses: OK, Misaligned (the
elements moved but are still inside the opening; optional), Changed,
Obsolete (no element left: the opening is kept). Then choose whether to
recreate the changed (and misaligned) openings.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import math

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Collections.Generic import List

from pyrevit import revit, script, DB, forms

MODE = "place"  # "place" | "refresh": unica differenza tra Place e Refresh, tenere i due script allineati

# =================================================================================================================
# Blocco comune a Place, Refresh e Settings (nomi, parametri, impostazioni predefinite): tenerlo allineato
# =================================================================================================================
CONFIG_NAME = "ESA_Openings"
DEFAULTS = {"fam_wall_rect": "e_GM.PH.PFV_Wall_Rect", "fam_wall_round": "e_GM.PH.PFV_Wall_Round",
            "fam_floor_rect": "e_GM.PH.PFV_Floor_Rect", "fam_floor_round": "e_GM.PH.PFV_Floor_Round",
            "clearance_mm": 25.0, "rounding_mm": 50.0, "merge_mm": 100.0, "extra_depth_mm": 0.0, "layer_gap_mm": 200.0,
            "min_angle_deg": 15.0,
            "separate": False, "misaligned": True}
NUMBER_KEYS = ("clearance_mm", "rounding_mm", "merge_mm", "extra_depth_mm", "layer_gap_mm", "min_angle_deg")
BOOL_KEYS = ("separate", "misaligned")
FAMILY_KEYS = {("wall", "rect"): "fam_wall_rect", ("wall", "round"): "fam_wall_round",
               ("floor", "rect"): "fam_floor_rect", ("floor", "round"): "fam_floor_round"}
# famiglie Generic Model NON ospitate, origine al centro. Muro: X larghezza, Y profondita', Z altezza.
# Solaio: X e Y in pianta, Z profondita'. Parametri di istanza (condivisi), lunghezza
PARAM_WIDTH = "e_DIM_Width_1"
PARAM_HEIGHT = "e_DIM_Height_1"
PARAM_DEPTH = "e_DIM_Depth_1"
PARAM_DIAMETER = "e_DIM_Diameter_1"
STATUS_PARAM = "e_GEN_PFV_Status"
ID_PARAM = "e_GEN_PFV_ID"
CODE_PREFIX = "PFV"
NEW, OK, MISALIGNED, CHANGED, OBSOLETE = "New", "OK", "Misaligned", "Changed", "Obsolete"
SPLIT, JOINED = "Split", "Joined"  # fori ricreati dal Refresh: divisi per disciplina, o uniti in uno
# =================================================================================================================

MOVE_TOL_MM = 5.0
SIZE_TOL_MM = 1.0
MM_PER_FT = 304.8
VERTICAL_SIN = math.sin(math.radians(45.0))   # oltre: elemento MEP verticale, foro a solaio
ROUND_COS = 0.995            # tondo solo se l'attraversamento e' (quasi) perpendicolare
Z = DB.XYZ.BasisZ
MEP_CLASSES = (Pipe, Duct, CableTray, Conduit)
MEP_CATEGORIES = (DB.BuiltInCategory.OST_PipeCurves, DB.BuiltInCategory.OST_DuctCurves,
                  DB.BuiltInCategory.OST_CableTray, DB.BuiltInCategory.OST_Conduit)
# muri e solai: foro a muro / a solaio; strutture (travi, pilastri, fondazioni): secondo la direzione dell'elemento MEP
HOST_KINDS = {int(DB.BuiltInCategory.OST_Walls): "wall", int(DB.BuiltInCategory.OST_Floors): "floor",
              int(DB.BuiltInCategory.OST_StructuralFraming): "auto", int(DB.BuiltInCategory.OST_StructuralColumns): "auto",
              int(DB.BuiltInCategory.OST_StructuralFoundation): "auto"}
STATE_RANK = {OK: 0, NEW: 1, SPLIT: 1, JOINED: 1, MISALIGNED: 2, CHANGED: 3, OBSOLETE: 4}
TITLE = "Openings Place" if MODE == "place" else "Openings Refresh"

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


def load_settings():
    settings = dict(DEFAULTS)
    try:
        config = script.get_config(CONFIG_NAME)
        for key in DEFAULTS:
            value = config.get_option(key, DEFAULTS[key])
            if key in NUMBER_KEYS:
                value = float(value)
            elif key in BOOL_KEYS and not isinstance(value, bool):
                value = str(value).strip().lower() in ("true", "1", "yes")
            settings[key] = value
    except Exception:
        pass
    return settings


SETTINGS = load_settings()
# sotto quest'angolo con il piano di muro o solaio l'elemento MEP corre dentro (scarichi a pavimento, montanti nel
# muro): nessun foro
MIN_CROSS_SIN = math.sin(math.radians(min(max(SETTINGS["min_angle_deg"], 0.0), 89.0)))


def id_value(eid):
    return eid.Value if hasattr(eid, "Value") else eid.IntegerValue  # Revit 2026+ / <= 2025


def mm(value_ft):
    return value_ft * MM_PER_FT


def round_up(value_mm):
    step = SETTINGS["rounding_mm"]
    if step <= 0:
        return value_mm
    return math.ceil(value_mm / step - 1e-9) * step


def level_elevation(level):
    # ProjectElevation e' coerente con la geometria (coordinate interne), Elevation puo' riferirsi al Survey Point
    value = getattr(level, "ProjectElevation", None)
    return level.Elevation if value is None else value


# ----------------------------------------------------------------------------------------------------------------
# Parametri PFV: stato "<disciplina> - <stato>", codici
# ----------------------------------------------------------------------------------------------------------------
def get_text(element, name):
    """Valore di testo del parametro, "" se vuoto, None se il parametro non c'e'."""
    param = element.LookupParameter(name)
    if param is None:
        return None
    return param.AsString() or ""


def set_text(element, name, value, missing):
    param = element.LookupParameter(name)
    if param is None or param.IsReadOnly:
        missing.add("{} on {}".format(name, element.Category.Name if element.Category else "element"))
        return
    param.Set(value)


def discipline(mep):
    return "MEC" if isinstance(mep, (Pipe, Duct)) else "ELE"


def disciplines_text(meps):
    found = set(discipline(m) for m in meps)
    return "+".join(d for d in ("MEC", "ELE") if d in found)


def status_text(disciplines, state):
    return "{} - {}".format(disciplines, state) if disciplines else state


def parse_status(text):
    """(disciplina, stato) da "MEC - New"."""
    if " - " in text:
        left, right = text.split(" - ", 1)
        return left.strip(), right.strip()
    return "", text.strip()


def code_of(opening):
    return "{}{}".format(CODE_PREFIX, id_value(opening.Id))


def split_list(text):
    return [part.strip() for part in (text or "").split(",") if part.strip()]


def element_by_id(text):
    try:
        return doc.GetElement(DB.ElementId(int(text)))
    except Exception:
        return None


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


def mep_section(mep):
    """Sezione esterna, isolamento compreso, letta dalla geometria: (asse X, asse Y, semi-larghezza, semi-altezza,
    tonda). Gli assi sono quelli del connettore, perpendicolari al tratto; le misure vengono dai solidi di elemento
    e isolamento, cosi' valgono anche per canali ruotati e passerelle di costa."""
    conn = [c for c in mep.ConnectorManager.Connectors if c.ConnectorType == DB.ConnectorType.End][0]
    bx, by = conn.CoordinateSystem.BasisX, conn.CoordinateSystem.BasisY
    solids = solids_of(mep)
    if isinstance(mep, (Pipe, Duct)):
        for insulation_id in DB.InsulationLiningBase.GetInsulationIds(doc, mep.Id):
            solids.extend(solids_of(doc.GetElement(insulation_id)))
    origin = mep.Location.Curve.GetEndPoint(0)
    points = solid_points(solids)
    if not points:
        raise ValueError("no geometry")
    xs = [p.Subtract(origin).DotProduct(bx) for p in points]
    ys = [p.Subtract(origin).DotProduct(by) for p in points]
    return bx, by, (max(xs) - min(xs)) / 2.0, (max(ys) - min(ys)) / 2.0, conn.Shape == DB.ConnectorProfileType.Round


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
    clear = 2.0 * SETTINGS["clearance_mm"]
    depth = mm(thickness) + 2.0 * SETTINGS["extra_depth_mm"]
    if len(items) == 1 and items[0]["round"]:
        return {"center": center, "shape": "round", "diameter": round_up(mm(max(ext_u, ext_v)) + clear),
                "depth": depth}
    return {"center": center, "shape": "rect", "width": round_up(mm(ext_u) + clear),
            "height": round_up(mm(ext_v) + clear), "depth": depth}


def crossing_item(mep, section, p0, p1, frame, kind):
    direction = mep.Location.Curve.Direction
    return {"mep": mep, "p0": p0, "p1": p1, "mid": p0.Add(p1).Divide(2.0), "frame": frame, "kind": kind,
            "points": section_points(p0, section) + section_points(p1, section), "discipline": discipline(mep),
            "round": section[4] and abs(direction.DotProduct(frame[2])) > ROUND_COS, "section": section}


# ----------------------------------------------------------------------------------------------------------------
# Famiglie, fori esistenti, inserimento
# ----------------------------------------------------------------------------------------------------------------
def family_symbols():
    """{(muro|solaio, rect|round): primo tipo della famiglia scelta in Settings}; mancanti: elenco dei nomi."""
    by_name = {}
    for symbol in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(
            DB.BuiltInCategory.OST_GenericModel):
        by_name.setdefault(symbol.Family.Name, symbol)
    symbols, missing = {}, []
    for key, setting in FAMILY_KEYS.items():
        name = SETTINGS[setting]
        if name in by_name:
            symbols[key] = by_name[name]
        else:
            missing.append(name)
    return symbols, sorted(set(missing))


def existing_openings():
    """Fori gia' gestiti: Generic Model con e_GEN_PFV_Status compilato (qualunque famiglia)."""
    return [i for i in DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).OfCategory(
        DB.BuiltInCategory.OST_GenericModel) if get_text(i, STATUS_PARAM)]


def pick_level(z):
    levels = sorted(DB.FilteredElementCollector(doc).OfClass(DB.Level), key=level_elevation)
    below = [l for l in levels if level_elevation(l) <= z + 1e-6]
    return below[-1] if below else levels[0]


def create_opening(data, frame, kind, symbols, meps, state, missing):
    """Inserisce il foro al centro, con l'asse X della famiglia lungo u; scrive misure, stato e Id dei MEP."""
    symbol = symbols[(kind, data["shape"])]
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()
    center = data["center"]
    instance = doc.Create.NewFamilyInstance(center, symbol, pick_level(center.Z), StructuralType.NonStructural)
    doc.Regenerate()
    current = instance.GetTransform().BasisX
    u = frame[0]
    angle = math.atan2(current.CrossProduct(u).Z, current.DotProduct(u))
    if abs(angle) > 1e-9:
        axis = DB.Line.CreateBound(instance.Location.Point, instance.Location.Point.Add(Z))
        DB.ElementTransformUtils.RotateElement(doc, instance.Id, axis, angle)
    delta = center.Subtract(instance.Location.Point)
    if delta.GetLength() > 1e-9:
        DB.ElementTransformUtils.MoveElement(doc, instance.Id, delta)
    values = ((PARAM_DIAMETER, data["diameter"]), (PARAM_DEPTH, data["depth"])) if data["shape"] == "round" else \
        ((PARAM_WIDTH, data["width"]), (PARAM_HEIGHT, data["height"]), (PARAM_DEPTH, data["depth"]))
    for name, value in values:
        param = instance.LookupParameter(name)
        if param is None or param.IsReadOnly:
            missing.add("{} on family {}".format(name, symbol.Family.Name))
        else:
            param.Set(value / MM_PER_FT)
    set_text(instance, STATUS_PARAM, status_text(disciplines_text(meps), state), missing)
    set_text(instance, ID_PARAM, ",".join(str(id_value(m.Id)) for m in meps), missing)
    return instance


def size_text(data):
    if data["shape"] == "round":
        return "D{:.0f} x {:.0f}".format(data["diameter"], data["depth"])
    return "{:.0f} x {:.0f} x {:.0f}".format(data["width"], data["height"], data["depth"])


def write_mep(mep, codes, state, missing):
    """Codici dei fori del MEP e stato; senza codici i campi PFV si svuotano."""
    set_text(mep, ID_PARAM, ",".join(codes), missing)
    set_text(mep, STATUS_PARAM, status_text(discipline(mep), state) if codes else "", missing)


# ----------------------------------------------------------------------------------------------------------------
# PLACE
# ----------------------------------------------------------------------------------------------------------------
class HostFilter(ISelectionFilter):
    """Muri, solai e strutture (travi, pilastri, fondazioni), nel modello host o in un link RVT."""

    def AllowElement(self, element):
        if isinstance(element, DB.RevitLinkInstance):
            return True
        return element.Category is not None and id_value(element.Category.Id) in HOST_KINDS

    def AllowReference(self, reference, point):
        host = doc.GetElement(reference.ElementId)
        if isinstance(host, DB.RevitLinkInstance):
            link_doc = host.GetLinkDocument()
            element = link_doc.GetElement(reference.LinkedElementId) if link_doc is not None else None
        else:
            element = host
        return element is not None and element.Category is not None and \
            id_value(element.Category.Id) in HOST_KINDS


def pick_hosts():
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.PointOnElement, HostFilter(),
                                           "Openings: select walls, floors and structural elements, also in linked "
                                           "models, then Finish")
    except OperationCanceledException:
        script.exit()
    hosts, seen = [], set()
    for ref in refs:
        host = doc.GetElement(ref.ElementId)
        if isinstance(host, DB.RevitLinkInstance):
            element = host.GetLinkDocument().GetElement(ref.LinkedElementId)
            key = (id_value(host.Id), id_value(element.Id))
            record = (element, host.GetTotalTransform(), host.GetLinkDocument().Title)
        else:
            key = (0, id_value(host.Id))
            record = (host, None, "Host model")
        if key not in seen:
            seen.add(key)
            hosts.append((key,) + record)
    return hosts


class MepPickFilter(ISelectionFilter):
    """Tubi, canali, passerelle e conduit rettilinei del modello host."""

    def AllowElement(self, element):
        return is_straight_mep(element)

    def AllowReference(self, reference, point):
        return False


def is_straight_mep(element):
    return isinstance(element, MEP_CLASSES) and isinstance(element.Location.Curve, DB.Line)


def pick_meps():
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, MepPickFilter(),
                                           "Openings: select pipes, ducts, cable trays and conduits, then Finish")
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def hosts_from_meps(meps):
    """Modalita' inversa: muri, solai e strutture del modello host e dei link caricati che toccano il bounding box
    di almeno un elemento MEP scelto. Stesso formato di pick_hosts."""
    categories = List[DB.BuiltInCategory]([DB.BuiltInCategory.OST_Walls, DB.BuiltInCategory.OST_Floors,
                                           DB.BuiltInCategory.OST_StructuralFraming,
                                           DB.BuiltInCategory.OST_StructuralColumns,
                                           DB.BuiltInCategory.OST_StructuralFoundation])
    boxes = [m.get_BoundingBox(None) for m in meps]
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
    """Elementi MEP rettilinei del modello host vicini ai solidi (filtro sul bounding box)."""
    points = solid_points(solids)
    low = DB.XYZ(min(p.X for p in points), min(p.Y for p in points), min(p.Z for p in points))
    high = DB.XYZ(max(p.X for p in points), max(p.Y for p in points), max(p.Z for p in points))
    box_filter = DB.BoundingBoxIntersectsFilter(DB.Outline(low, high))
    categories = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory](MEP_CATEGORIES))
    return [e for e in DB.FilteredElementCollector(doc).WherePasses(categories).WherePasses(box_filter)
            .WhereElementIsNotElementType() if isinstance(e, MEP_CLASSES) and isinstance(e.Location.Curve, DB.Line)]


def inside_existing(point, boxes):
    for low, high in boxes:
        if low.X <= point.X <= high.X and low.Y <= point.Y <= high.Y and low.Z <= point.Z <= high.Z:
            return True
    return False


ADJACENT_SEARCH = 1000.0 / MM_PER_FT  # margine attorno all'elemento scelto in cui si cercano gli strati vicini
MAX_LAYER = 1000.0 / MM_PER_FT  # strato attraversato per piu' di cosi': l'elemento MEP ci corre dentro, non conta
PARALLEL_COS = 0.98             # strati di muro paralleli all'elemento scelto entro circa 11 gradi
NEIGHBOUR_CACHE = {}  # (id del link o 0, id dell'elemento) -> solidi in coordinate host


def model_sources():
    """(documento, transform o None, id del link o 0) del modello host e dei link RVT caricati."""
    sources = [(doc, None, 0)]
    for link in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance):
        link_doc = link.GetLinkDocument()
        if link_doc is not None:
            sources.append((link_doc, link.GetTotalTransform(), id_value(link.Id)))
    return sources


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
            if key not in NEIGHBOUR_CACHE:
                own = solids_of(element)
                NEIGHBOUR_CACHE[key] = [DB.SolidUtils.CreateTransformed(s, transform) for s in own] \
                    if transform is not None else own
            label = "{} {} ({})".format(element.Category.Name, id_value(element.Id),
                                        source_doc.Title if link_id else "Host model")
            found.append((key, label, NEIGHBOUR_CACHE[key], element, transform))
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
    gap = max(SETTINGS["layer_gap_mm"], 5.0) / MM_PER_FT
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



def find_crossings(hosts, stats, only=None):
    """Attraversamenti degli elementi MEP con muri, solai e strutture, allungati attraverso gli strati a contatto
    (muri o solai separati di una struttura composita). Salta quelli che hanno gia' un foro. only: id degli elementi
    MEP da considerare (modalita' inversa), None = tutti quelli vicini."""
    boxes = []
    for opening in existing_openings():
        box = opening.get_BoundingBox(None)
        if box is not None:
            boxes.append((box.Min, box.Max))
    items = []
    options = DB.SolidCurveIntersectionOptions()
    for key, element, transform, model in hosts:
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
            if only is not None and id_value(mep.Id) not in only:
                continue
            line = mep.Location.Curve
            try:
                section = mep_section(mep)
            except Exception:
                stats["problems"].append("{} {}: section not readable".format(mep.Category.Name,
                                                                               id_value(mep.Id)))
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
                    if abs(line.Direction.DotProduct(w)) < MIN_CROSS_SIN:
                        stats["inside"] += 1  # corre dentro il muro o il solaio: nessun foro
                        continue
                    p0, p1, used = extend_through(p0, p1, layers, options, kind, w)
                    mid = p0.Add(p1).Divide(2.0)
                    if inside_existing(mid, boxes):
                        stats["existing"] += 1
                        continue
                    item = crossing_item(mep, section, p0, p1, make_frame(kind, w, section[0]), kind)
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
    gap = max(SETTINGS["layer_gap_mm"], 5.0) / MM_PER_FT
    merged = []
    for item in items:
        direction = item["mep"].Location.Curve.Direction
        target = None
        for other in merged:
            if other["mep"].Id != item["mep"].Id or other["kind"] != item["kind"]:
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
    merge = (SETTINGS["merge_mm"] + 2.0 * SETTINGS["clearance_mm"]) / MM_PER_FT
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
            if SETTINGS["separate"] and a["discipline"] != b["discipline"]:
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


def settings_line():
    return ("clearance {clearance_mm:g} mm per side, rounding {rounding_mm:g} mm, merge within {merge_mm:g} mm, "
            "extra depth {extra_depth_mm:g} mm per side, layers joined across gaps up to {layer_gap_mm:g} mm, "
            "elements under {min_angle_deg:g} degrees to the wall/floor skipped, MEC/ELE {sep}").format(
        sep="separate" if SETTINGS["separate"] else "together", **SETTINGS)


def place_main():
    symbols, missing_families = family_symbols()
    if missing_families:
        forms.alert("Load the opening families first (or choose others in Settings):\n{}".format(
            "\n".join(missing_families)), title=TITLE, exitscript=True)
    # modalita' dalla selezione: elementi MEP gia' selezionati = inversa; nessuna selezione = si chiede
    selected = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    meps = [e for e in selected if e is not None and is_straight_mep(e)]
    if not meps:
        choice = forms.CommandSwitchWindow.show(["Select walls, floors and structure", "Select MEP elements"],
                                                message="Openings: what do you select?")
        if choice is None:
            script.exit()
        if choice == "Select MEP elements":
            meps = pick_meps()
    stats = {"inside": 0, "existing": 0, "extended": 0, "problems": []}
    if meps:
        hosts = hosts_from_meps(meps)
        items = find_crossings(hosts, stats, only=set(id_value(m.Id) for m in meps))
        mode_text = "from **{}** MEP elements: walls, floors and structural elements found **{}**".format(
            len(meps), len(hosts))
    else:
        hosts = pick_hosts()
        items = find_crossings(hosts, stats)
        mode_text = "walls, floors and structural elements selected: **{}**".format(len(hosts))
    items = merge_overlapping(items)
    groups = clusters(items)
    output.print_md("# Openings Place")
    output.print_md("Settings: {}".format(settings_line()))
    output.print_md("Mode: {} | crossings: **{}** | openings to create: **{}** | already with an opening: **{}** | "
                    "MEP running inside (no opening): **{}** | extended through adjacent layers: **{}**".format(
                        mode_text, len(items), len(groups), stats["existing"], stats["inside"], stats["extended"]))
    for problem in stats["problems"]:
        output.print_md("- {}".format(problem))
    if not groups:
        output.print_md("Nothing to create.")
        return
    if not forms.alert("Create {} openings for {} crossings?".format(len(groups), len(items)), title=TITLE,
                       yes=True, no=True):
        script.exit()

    missing = set()
    rows = []
    mep_codes = {}  # id elemento MEP -> (elemento, codici dei fori nuovi)
    transaction = DB.Transaction(doc, TITLE)
    transaction.Start()
    try:
        for group in groups:
            frame, kind = group[0]["frame"], group[0]["kind"]
            data = sizes(group, frame)
            meps = []
            for item in group:
                if all(m.Id != item["mep"].Id for m in meps):
                    meps.append(item["mep"])
            instance = create_opening(data, frame, kind, symbols, meps, NEW, missing)
            code = code_of(instance)
            for mep in meps:
                mep_codes.setdefault(id_value(mep.Id), (mep, []))[1].append(code)
            key, host, model = group[0]["host"]
            rows.append([output.linkify(instance.Id), code, disciplines_text(meps),
                         "{} {} ({})".format(host.Category.Name, id_value(host.Id), model), data["shape"],
                         size_text(data), ", ".join(str(id_value(m.Id)) for m in meps)])
        for mep, new_codes in mep_codes.values():
            old = [c for c in split_list(get_text(mep, ID_PARAM)) if c.startswith(CODE_PREFIX)]
            write_mep(mep, old + [c for c in new_codes if c not in old], NEW, missing)
        transaction.Commit()
    except Exception:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    output.print_table(table_data=rows, columns=["Opening", "Code", "Discipline", "Crossed element", "Shape",
                                                 "Size [mm]", "MEP elements"])
    if missing:
        output.print_md("**Parameters not found (not written):** {}".format("; ".join(sorted(missing))))


# ----------------------------------------------------------------------------------------------------------------
# REFRESH
# ----------------------------------------------------------------------------------------------------------------
def read_length(instance, name):
    param = instance.LookupParameter(name)
    return mm(param.AsDouble()) if param is not None and param.HasValue else None


def opening_kind(instance, meps):
    """(muro|solaio, rect|round): dalla famiglia, se e' una di quelle di Settings; altrimenti dai parametri (tondo
    se ha il diametro) e dalla direzione dei suoi elementi MEP (verticali: solaio)."""
    name = instance.Symbol.Family.Name
    for key, setting in FAMILY_KEYS.items():
        if SETTINGS[setting] == name:
            return key
    shape = "round" if read_length(instance, PARAM_DIAMETER) else "rect"
    vertical = any(abs(m.Location.Curve.Direction.Z) > VERTICAL_SIN for m in meps)
    return ("floor" if vertical else "wall"), shape


def opening_frame(instance, kind):
    u = instance.GetTransform().BasisX
    return (u, Z, Z.CrossProduct(u)) if kind == "wall" else (u, Z.CrossProduct(u), Z)


def current_data(instance, shape):
    data = {"center": instance.Location.Point, "shape": shape, "depth": read_length(instance, PARAM_DEPTH) or 0.0}
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
    oltre la distanza di unione."""
    line = mep.Location.Curve
    if not isinstance(line, DB.Line):
        return None
    u, v, w = frame
    direction = line.Direction
    cos = direction.DotProduct(w)
    if abs(cos) < MIN_CROSS_SIN:
        return None
    start = line.GetEndPoint(0)
    half = data["depth"] / MM_PER_FT / 2.0
    base = data["center"].Subtract(start).DotProduct(w)
    t0, t1 = sorted(((base - half) / cos, (base + half) / cos))
    if t1 < -1e-6 or t0 > line.Length + 1e-6:
        return None  # il tratto non arriva piu' al foro
    p0, p1 = start.Add(direction.Multiply(t0)), start.Add(direction.Multiply(t1))
    mid = p0.Add(p1).Divide(2.0).Subtract(data["center"])
    reach_u, reach_v = [(h + SETTINGS["merge_mm"]) / MM_PER_FT for h in half_sizes(data)]
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
    return notes


def rebuilt(items, frame, center, depth):
    """Foro ricalcolato dagli attraversamenti. Lo spessore attraversato non si ricalcola (il muro non e' noto): il
    centro resta sul piano del foro e la profondita' e' quella attuale."""
    new = sizes(items, frame)
    new["center"] = new["center"].Subtract(frame[2].Multiply(new["center"].Subtract(center).DotProduct(frame[2])))
    new["depth"] = depth
    return new


def check_opening(instance):
    """Confronto di un foro con i suoi elementi MEP. Ritorna il dict del risultato."""
    listed = split_list(get_text(instance, ID_PARAM))
    found = []
    result = {"instance": instance, "code": code_of(instance), "keep": [], "notes": [], "state": OK, "new": None,
              "old_status": get_text(instance, STATUS_PARAM)}
    for text in listed:
        element = element_by_id(text)
        if element is None or not isinstance(element, MEP_CLASSES):
            result["notes"].append("{} no longer exists".format(text))
        else:
            found.append(element)
    kind, shape = opening_kind(instance, found)
    frame = opening_frame(instance, kind)
    old = current_data(instance, shape)
    result.update({"kind": kind, "frame": frame, "old": old})
    items, all_inside = [], True
    result["misplaced"] = []  # (elemento, attraversamento): fuori dal foro, il foro si puo' spostare su di lui
    result["found"] = len(found)  # elementi del foro ancora esistenti (per spostare il foro: foro di un solo elemento)
    for mep in found:
        try:
            section = mep_section(mep)
        except Exception:
            result["notes"].append("{}: section not readable".format(id_value(mep.Id)))
            continue
        crossing = recross(mep, frame, old)
        if crossing is None:
            far = recross(mep, frame, old, limit_reach=False)
            if far is None:
                result["notes"].append("{} no longer crosses the opening".format(id_value(mep.Id)))
            else:
                # ancora parallelo all'attraversamento ma spostato lontano: il foro si puo' spostare su di lui
                result["notes"].append("{} moved out of the opening".format(id_value(mep.Id)))
                result["misplaced"].append((mep, crossing_item(mep, section, far[0], far[1], frame, kind)))
            continue
        item = crossing_item(mep, section, crossing[0], crossing[1], frame, kind)
        if not inside_opening(item, frame, old):
            all_inside = False
            result["notes"].append("{} partly outside the opening".format(id_value(mep.Id)))
            result["misplaced"].append((mep, item))
        items.append(item)
        result["keep"].append(mep)
    result["items"] = items
    result["groups"] = [items]  # fori che ne prendono il posto, se ricreato (due se MEC ed ELE vanno separati)
    if not items:
        result["state"] = OBSOLETE
        return result
    new = rebuilt(items, frame, old["center"], old["depth"])
    result["new"] = new
    lost = len(result["keep"]) < len(listed)  # elementi spariti, non piu' attraversanti o illeggibili
    diffs = differences(old, new)
    result["notes"].extend(diffs)
    if lost or not all_inside:
        result["state"] = CHANGED
    elif diffs:
        result["state"] = MISALIGNED if SETTINGS["misaligned"] else CHANGED
    return result


def opening_on_hosts(instance, host_solids):
    """Il foro sta su uno dei muri/solai scelti: una linea per il suo centro, lungo Y e lungo Z della famiglia (gli
    assi di attraversamento di muro e solaio), entra nel solido dell'elemento."""
    transform = instance.GetTransform()
    center = instance.Location.Point
    reach = max((read_length(instance, PARAM_DEPTH) or 0.0) / MM_PER_FT, 1.0)
    options = DB.SolidCurveIntersectionOptions()
    for axis in (transform.BasisY, transform.BasisZ):
        line = DB.Line.CreateBound(center.Subtract(axis.Multiply(reach)), center.Add(axis.Multiply(reach)))
        for solid in host_solids:
            try:
                if solid.IntersectWithCurve(line, options).SegmentCount > 0:
                    return True
            except Exception:
                continue
    return False


def ask_refresh_scope():
    """Fori da controllare. Con elementi gia' selezionati (MEP o fori): solo i fori legati a quelli, tramite
    e_GEN_PFV_ID. Altrimenti si chiede: tutti, o quelli di muri, solai e strutture da selezionare. Ritorna
    (fori, descrizione)."""
    openings = existing_openings()
    selected = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    selected = [e for e in selected if e is not None]
    opening_ids = set(id_value(o.Id) for o in openings)
    sel_meps = [e for e in selected if isinstance(e, MEP_CLASSES)]
    sel_openings = [e for e in selected if id_value(e.Id) in opening_ids]
    if sel_meps or sel_openings:
        mep_ids = set(str(id_value(m.Id)) for m in sel_meps)
        codes = set(c for m in sel_meps for c in split_list(get_text(m, ID_PARAM)))
        wanted = set(id_value(o.Id) for o in sel_openings)
        chosen = [o for o in openings if id_value(o.Id) in wanted or code_of(o) in codes or
                  mep_ids & set(split_list(get_text(o, ID_PARAM)))]
        return chosen, "selection ({} MEP elements, {} openings)".format(len(sel_meps), len(sel_openings))
    choice = forms.CommandSwitchWindow.show(["All openings", "Openings of selected walls, floors and structure"],
                                            message="Refresh which openings?")
    if choice is None:
        script.exit()
    if choice == "All openings":
        return openings, "all openings"
    hosts = pick_hosts()
    solids = []
    for key, element, transform, model in hosts:
        for solid in solids_of(element):
            solids.append(DB.SolidUtils.CreateTransformed(solid, transform) if transform is not None else solid)
    return [o for o in openings if opening_on_hosts(o, solids)], "{} selected walls, floors and structural elements".format(len(hosts))


def discipline_checks(results):
    """Divisioni e unioni dei fori secondo le impostazioni attuali di Settings.
    - MEC ed ELE separati: un foro con entrambe le discipline e' Changed; ricreato, diventa due fori (Split).
    - Unione: due o piu' fori sullo stesso piano i cui fori ricalcolati (franco compreso) distano al massimo la
      distanza di unione sono Changed; ricreati, diventano un foro solo (Joined). Vale per elementi avvicinati e
      per la spunta MEC/ELE tolta; con MEC ed ELE separati si uniscono solo fori della stessa disciplina."""
    live = [r for r in results if r["state"] != OBSOLETE and r["items"]]
    if SETTINGS["separate"]:
        for r in live:
            split = [[i for i in r["items"] if i["discipline"] == d] for d in ("MEC", "ELE")]
            split = [g for g in split if g]
            if len(split) > 1:
                r["groups"] = split
                r["state"] = CHANGED
                r["notes"].append("MEC and ELE must be separate (Settings): it becomes two openings")
    merge = SETTINGS["merge_mm"] / MM_PER_FT
    candidates = [r for r in live if len(r["groups"]) == 1]
    parent = dict((id(r), id(r)) for r in candidates)

    def root(key):
        while parent[key] != key:
            key = parent[key]
        return key

    def disciplines(r):
        return set(i["discipline"] for i in r["items"])

    for i in range(len(candidates)):
        for j in range(i + 1, len(candidates)):
            a, b = candidates[i], candidates[j]
            if SETTINGS["separate"] and (len(disciplines(a)) > 1 or disciplines(a) != disciplines(b)):
                continue
            u, v, w = a["frame"]
            if abs(w.DotProduct(b["frame"][2])) < 0.999:
                continue
            if mm(abs(b["old"]["center"].Subtract(a["old"]["center"]).DotProduct(w))) > MOVE_TOL_MM:
                continue  # non sullo stesso piano: muri o solai diversi
            # fori come andrebbero ricalcolati (misure con il franco): si uniscono se distano <= unione
            d = b["new"]["center"].Subtract(a["new"]["center"])
            ha, hb = [[h / MM_PER_FT for h in half_sizes(r["new"])] for r in (a, b)]
            gap = max(abs(d.DotProduct(u)) - ha[0] - hb[0], abs(d.DotProduct(v)) - ha[1] - hb[1])
            if gap <= merge:
                parent[root(id(b))] = root(id(a))
    groups = {}
    for r in candidates:
        groups.setdefault(root(id(r)), []).append(r)
    for members in groups.values():
        if len(members) < 2:
            continue
        codes = ", ".join(m["code"] for m in members)
        for m in members:
            m["merge_group"] = members
            m["state"] = CHANGED
            m["notes"].append("close enough to be one opening: {} become one".format(codes))


class MisplacedOpening(object):
    """Riga della lista dei fori da spostare sui loro elementi."""

    def __init__(self, instance, vector, label):
        self.instance, self.vector, self.label = instance, vector, label

    def __str__(self):
        return self.label


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
    clear = SETTINGS["clearance_mm"] / MM_PER_FT
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


def ask_alignment(results):
    """Fori i cui elementi (trovati tramite e_GEN_PFV_ID) ne sono usciti: si sceglie quali fori spostare sui loro
    elementi. Ritorna [MisplacedOpening] scelti (vuoto: nessuno)."""
    candidates = []
    for r in results:
        if not r.get("misplaced"):
            continue
        items = list(r["items"]) + [item for _, item in r["misplaced"] if item not in r["items"]]
        vector = opening_shift(items, r["frame"], r["old"], r["found"] <= 1)
        if vector.GetLength() * MM_PER_FT < 1.0:
            continue
        names = ", ".join("{} {}".format(m.Category.Name, id_value(m.Id)) for m, _ in r["misplaced"])
        label = "Opening {}: move {:.0f} mm onto {}".format(r["code"], vector.GetLength() * MM_PER_FT, names)
        candidates.append(MisplacedOpening(r["instance"], vector, label))
    if not candidates:
        return []
    picked = forms.SelectFromList.show(
        [forms.TemplateListItem(c, checked=True) for c in candidates], multiselect=True,
        title="Elements outside their opening: tick the openings to move onto them", button_name="Move the openings")
    return picked or []


def refresh_main():
    symbols, missing_families = family_symbols()
    all_openings = existing_openings()
    all_codes = set(code_of(o) for o in all_openings)
    openings, scope = ask_refresh_scope()
    results = [check_opening(o) for o in openings]

    # elementi usciti dal loro foro: si chiede se spostare il foro su di loro, poi si ricontrolla
    aligned = ask_alignment(results)
    if aligned:
        transaction = DB.Transaction(doc, "Openings: move openings onto their elements")
        transaction.Start()
        try:
            for entry in aligned:
                DB.ElementTransformUtils.MoveElement(doc, entry.instance.Id, entry.vector)
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        results = [check_opening(o) for o in openings]
    discipline_checks(results)
    in_scope = set(r["code"] for r in results)
    categories = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory](MEP_CATEGORIES))
    meps = [e for e in DB.FilteredElementCollector(doc).WherePasses(categories).WhereElementIsNotElementType()
            if isinstance(e, MEP_CLASSES) and get_text(e, STATUS_PARAM)]
    if scope != "all openings":
        # fuori dall'ambito si toccano solo i MEP che rimandano a uno dei fori controllati
        meps = [m for m in meps if any(c in in_scope for c in split_list(get_text(m, ID_PARAM)))]
    count = dict((s, len([r for r in results if r["state"] == s])) for s in STATE_RANK)

    output.print_md("# Openings Refresh")
    output.print_md("Scope: **{}** | MEC/ELE: **{}** | Misaligned: **{}** | minimum angle: **{:g} deg**".format(
        scope, "separate" if SETTINGS["separate"] else "together", "on" if SETTINGS["misaligned"] else "off",
        SETTINGS["min_angle_deg"]))
    output.print_md("Openings checked: **{}** | OK: **{}** | Misaligned: **{}** | Changed: **{}** | Obsolete: **{}**"
                    " | MEP elements involved: **{}**".format(len(results), count[OK], count[MISALIGNED],
                                                              count[CHANGED], count[OBSOLETE], len(meps)))
    if aligned:
        output.print_md("Openings moved onto their elements: **{}** ({}); they were checked again after the move."
                        .format(len(aligned), ", ".join(output.linkify(a.instance.Id) for a in aligned)))
    rows = [[output.linkify(r["instance"].Id), r["code"], r["state"], "; ".join(r["notes"])]
            for r in results if r["state"] != OK]
    if rows:
        output.print_table(table_data=rows, columns=["Opening", "Code", "Status", "Detail"])
    if not results:
        output.print_md("No opening with {} filled in this scope.".format(STATUS_PARAM))
        return

    recreate = ()
    choices = []
    if count[CHANGED]:
        choices.append("Recreate the changed openings")
    if count[MISALIGNED]:
        choices.append("Recreate the changed and misaligned openings")
    if choices:
        choices.append("Only update the statuses")
        choice = forms.CommandSwitchWindow.show(choices, message="{} changed, {} misaligned, {} obsolete. "
                                                                 "Recreate?".format(count[CHANGED], count[MISALIGNED],
                                                                                    count[OBSOLETE]))
        if choice is None:
            script.exit()
        if choice.startswith("Recreate the changed and"):
            recreate = (CHANGED, MISALIGNED)
        elif choice.startswith("Recreate"):
            recreate = (CHANGED,)
    if recreate and missing_families:
        forms.alert("Recreating needs the families chosen in Settings:\n{}".format("\n".join(missing_families)),
                    title=TITLE, exitscript=True)

    missing = set()
    holders = {}  # id MEP -> (elemento, [(codice del foro, stato)])
    recreated, removed = [], 0
    done_groups = set()

    def hold(meps_of, code, state):
        for mep in meps_of:
            holders.setdefault(id_value(mep.Id), (mep, []))[1].append((code, state))

    transaction = DB.Transaction(doc, TITLE)
    transaction.Start()
    try:
        for r in results:
            instance = r["instance"]
            if r["state"] in recreate and r.get("merge_group"):
                # MEC ed ELE di nuovo insieme: i fori del gruppo diventano uno
                key = id(r["merge_group"])
                if key in done_groups:
                    continue
                done_groups.add(key)
                members = r["merge_group"]
                items = [i for m in members for i in m["items"]]
                data = rebuilt(items, r["frame"], r["old"]["center"], max(m["old"]["depth"] for m in members))
                meps_of = [m_ for m in members for m_ in m["keep"]]
                for m in members:
                    doc.Delete(m["instance"].Id)
                removed += len(members)
                new = create_opening(data, r["frame"], r["kind"], symbols, meps_of, JOINED, missing)
                recreated.append(code_of(new))
                hold(meps_of, code_of(new), JOINED)
            elif r["state"] in recreate:
                # uno per uno, o due se MEC ed ELE vanno separati
                doc.Delete(instance.Id)
                removed += 1
                state = SPLIT if len(r["groups"]) > 1 else NEW
                for group in r["groups"]:
                    data = rebuilt(group, r["frame"], r["old"]["center"], r["old"]["depth"])
                    meps_of = []
                    for item in group:
                        if all(m.Id != item["mep"].Id for m in meps_of):
                            meps_of.append(item["mep"])
                    new = create_opening(data, r["frame"], r["kind"], symbols, meps_of, state, missing)
                    recreated.append(code_of(new))
                    hold(meps_of, code_of(new), state)
            elif r["state"] == OBSOLETE:
                old_discipline = parse_status(r["old_status"])[0]
                set_text(instance, STATUS_PARAM, status_text(old_discipline, OBSOLETE), missing)
            else:
                set_text(instance, STATUS_PARAM, status_text(disciplines_text(r["keep"]), r["state"]), missing)
                set_text(instance, ID_PARAM, ",".join(str(id_value(m.Id)) for m in r["keep"]), missing)
                hold(r["keep"], r["code"], r["state"])
        # MEP: codici dei fori fuori ambito (se esistono ancora) + quelli dei fori controllati che li contengono
        cleared = 0
        for mep in meps:
            old_codes = split_list(get_text(mep, ID_PARAM))
            outside = [c for c in old_codes if c not in in_scope and c in all_codes]
            held = holders.get(id_value(mep.Id), (mep, []))[1]
            codes = outside + [c for c, _ in held if c not in outside]
            states = [s for _, s in held]
            if outside:
                states.append(parse_status(get_text(mep, STATUS_PARAM))[1] or OK)
            state = max(states, key=lambda s: STATE_RANK.get(s, 0)) if states else OK
            if not codes:
                cleared += 1
            write_mep(mep, codes, state, missing)
        transaction.Commit()
    except Exception:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    if recreated:
        output.print_md("Openings removed: **{}** | created in their place: **{}** ({}), status New, Split or Joined; the codes of "
                        "their MEP elements are updated.".format(removed, len(recreated), ", ".join(recreated)))
    if cleared:
        output.print_md("MEP elements without an opening any more (PFV fields cleared): **{}**".format(cleared))
    if missing:
        output.print_md("**Parameters not found (not written):** {}".format("; ".join(sorted(missing))))
    if not rows:
        output.print_md("Everything matches: statuses set to OK.")


if MODE == "place":
    place_main()
else:
    refresh_main()
