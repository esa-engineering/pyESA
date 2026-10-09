# -*- coding: utf-8 -*-
__title__ = "Turn"
__doc__ = """Version = 0.4
Date    = 09.10.2026
________________________________________________________________
Description:
Adds a run that turns left or right from the free end of a pipe, duct,
cable tray or conduit, with an elbow. Replaces Turn Left and Turn Right.

Instruction:
- CLICK: one window with all the options, the last used are proposed
  and saved; they hold for all the clicks of this run of the command:
  - turn angle (common angles or a custom one)
  - length of the new run (default 1000 mm) and slope (keep the slope
    of the original run, or make the new run horizontal)
  - clashes: if the new run hits a wall (columns, beams and curtain
    walls too, also in linked RVT and IFC models) the angle is adjusted
    to run parallel to it. A wall, beam or column with a straight axis
    is followed along its axis; anything else (curtain wall panel, IFC
    element, column) along the face that was hit
- Click each element near the free end you want to extend, then click
  on the side the new run must turn to (ESC on this click skips the
  element)
- Press ESC to finish: a report lists the new elements

The side is taken from the second click, as seen from above, relative
to the direction of travel (from the element towards the clicked end).
Vertical ends are skipped.
The new parts keep the parameters and the insulation of the starting element.
Cable trays keep the orientation of their section (a tray on edge stays
on edge after the turn).
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import math

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB.Plumbing import Pipe, PipeInsulation
from Autodesk.Revit.DB.Mechanical import Duct, DuctInsulation, DuctLining
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Collections.Generic import List

from pyrevit import revit, script, forms, DB

import Turn_ui

CONFIG_NAME = "ESA_MEPTurn"  # stessa configurazione dei vecchi Turn Right e Turn Left: le impostazioni restano
DEFAULTS = {"length_mm": 1000.0, "angle_deg": 90.0, "keep_slope": True, "avoid_clash": False}
ANGLE_XAML_FILE_NAME = "Angle_form.xaml"
QUICK_ANGLES = (10.0, 15.0, 22.5, 30.0, 45.0, 60.0, 90.0)
MIN_ANGLE, MAX_ANGLE = 0.1, 90.0

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


REPORT_CONFIG = "ESA_Report"  # impostazione condivisa da tutti i pulsanti, cambiata col pulsante Report


def want_report(total, problems, noun):
    """True se va stampato il report completo. Modalita' "problems" (predefinita): solo se qualcosa non e' andato a
    buon fine; "always": sempre; "never": mai, resta la notifica. Altrimenti notifica breve (e nessuna finestra)."""
    try:
        mode = script.get_config(REPORT_CONFIG).get_option("mode", "problems")
    except Exception:
        mode = "problems"
    if mode == "always" or (mode == "problems" and problems):
        return True
    if not total:
        return False
    text = "{} {}".format(total - problems, noun) + (", {} not done".format(problems) if problems else "")
    try:
        script.toast(text, title="pyESA")
    except Exception:
        return problems > 0  # notifica non disponibile: con dei problemi meglio il report che niente
    return False

MM_PER_FT = 304.8
VERTICAL_TOL = 0.999
MIN_REMAINING_MM = 20.0  # lunghezza minima dei tratti dopo l'inserimento del gomito

CATEGORY_NAMES = (
    (Pipe, "Pipe"),
    (Duct, "Duct"),
    (CableTray, "Cable Tray"),
    (Conduit, "Conduit"),
)


class MepCurveFilter(ISelectionFilter):
    """Accetta solo tubi, condotti, passerelle e conduit."""

    def AllowElement(self, element):
        return isinstance(element, (Pipe, Duct, CableTray, Conduit))

    def AllowReference(self, reference, point):
        return False


class TurnError(Exception):
    """Errore atteso (es. spazio insufficiente): il messaggio va nel report."""


def category_name(element):
    for cls, name in CATEGORY_NAMES:
        if isinstance(element, cls):
            return name
    return "?"


def level_elevation(level):
    # ProjectElevation e' coerente con la geometria (coordinate interne), Elevation puo' riferirsi al Survey Point
    value = getattr(level, "ProjectElevation", None)
    return level.Elevation if value is None else value


def get_level_id(element):
    param = element.get_Parameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM)
    if param is not None and param.AsElementId() != DB.ElementId.InvalidElementId:
        return param.AsElementId()
    ref_level = getattr(element, "ReferenceLevel", None)
    return ref_level.Id if ref_level is not None else None


def nearest_free_connector(element, point):
    # in pianta la Z di GlobalPoint non e' significativa: si confronta solo XY
    planar = isinstance(doc.ActiveView, DB.ViewPlan)
    best, best_dist = None, None
    for conn in element.ConnectorManager.UnusedConnectors:
        origin = conn.Origin
        if planar:
            dist = ((origin.X - point.X) ** 2 + (origin.Y - point.Y) ** 2) ** 0.5
        else:
            dist = origin.DistanceTo(point)
        if best is None or dist < best_dist:
            best, best_dist = conn, dist
    return best


def nearest_connector(element, point):
    best, best_dist = None, None
    for c in element.ConnectorManager.Connectors:
        dist = c.Origin.DistanceTo(point)
        if best is None or dist < best_dist:
            best, best_dist = c, dist
    return best


def find_fitting(conn, element):
    """Raccordo (FamilyInstance) collegato al connettore, se esiste."""
    for ref in conn.AllRefs:
        owner = ref.Owner
        if owner.Id != element.Id and isinstance(owner, DB.FamilyInstance):
            return owner
    return None


def fitting_label(fitting):
    try:
        return "{}: {}".format(fitting.Symbol.FamilyName, fitting.Name)
    except Exception:
        return "elbow"


def copy_tray_dimensions(source, target):
    for bip in (DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM,
                DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM):
        src = source.get_Parameter(bip)
        tgt = target.get_Parameter(bip)
        if src and tgt and not tgt.IsReadOnly:
            tgt.Set(src.AsDouble())


def copy_conduit_diameter(source, target):
    for bip in (DB.BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM,
                DB.BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM):
        src = source.get_Parameter(bip)
        tgt = target.get_Parameter(bip)
        if src and tgt and not tgt.IsReadOnly:
            tgt.Set(src.AsDouble())
            return


def run_direction_at(element, conn):
    """Direzione di percorrenza: dall'altra estremita' verso quella del connettore."""
    curve = element.Location.Curve
    a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
    if a.DistanceTo(conn.Origin) < b.DistanceTo(conn.Origin):
        return a.Subtract(b).Normalize()
    return b.Subtract(a).Normalize()


def signed_turn_deg(angle_deg, side):
    """Angolo di curva con segno: positivo = antiorario visto dall'alto (sinistra)."""
    return angle_deg if side == "left" else -angle_deg


def side_of(travel, origin, click):
    """Lato del clic rispetto alla direzione di percorrenza, visto dall'alto: "left", "right" o None se il clic e'
    in linea con il tratto (in pianta conta solo XY)."""
    cross = travel.X * (click.Y - origin.Y) - travel.Y * (click.X - origin.X)
    horizontal_len = math.sqrt(travel.X ** 2 + travel.Y ** 2)
    if abs(cross) / horizontal_len < 1.0 / MM_PER_FT:  # meno di 1 mm dalla linea del tratto
        return None
    return "left" if cross > 0 else "right"


def turn_direction(travel, signed_deg, keep_slope):
    """Direzione del nuovo tratto, ruotando la direzione di percorrenza di signed_deg attorno a Z (positivo =
    sinistra). None se il tratto e' verticale (destra/sinistra non definite)."""
    if abs(travel.Z) > VERTICAL_TOL:
        return None
    horizontal_len = math.sqrt(travel.X ** 2 + travel.Y ** 2)
    forward = DB.XYZ(travel.X / horizontal_len, travel.Y / horizontal_len, 0)
    left = DB.XYZ(-forward.Y, forward.X, 0)  # asse Z x forward
    angle = math.radians(signed_deg)
    new_dir = forward.Multiply(math.cos(angle)).Add(left.Multiply(math.sin(angle)))
    if keep_slope:
        # stessa pendenza: stessa salita per unita' di percorso orizzontale
        new_dir = DB.XYZ(new_dir.X, new_dir.Y, travel.Z / horizontal_len).Normalize()
    return new_dir


# ----------------------------------------------------------------------------------------------------------------
# Interferenze: il nuovo tratto che colpisce un muro (anche in un link) viene girato parallelo al muro
# ----------------------------------------------------------------------------------------------------------------
CLASH_CATEGORIES = (DB.BuiltInCategory.OST_Walls, DB.BuiltInCategory.OST_Columns,
                    DB.BuiltInCategory.OST_StructuralColumns, DB.BuiltInCategory.OST_StructuralFraming,
                    DB.BuiltInCategory.OST_CurtainWallPanels, DB.BuiltInCategory.OST_CurtainWallMullions)
VERTICAL_FACE_TOL = 0.1  # sotto questa componente orizzontale la normale della faccia colpita non da' una direzione
MIN_ADJUSTED_DEG = 5.0  # sotto questo angolo la curva parallela al muro non e' piu' una curva


def clash_view():
    """Vista 3D per il ReferenceIntersector: quella attiva se 3D, altrimenti la prima 3D non modello."""
    view = doc.ActiveView
    if isinstance(view, DB.View3D) and not view.IsTemplate:
        return view
    views = [v for v in DB.FilteredElementCollector(doc).OfClass(DB.View3D) if not v.IsTemplate]
    default = [v for v in views if v.Name == "{3D}"]
    return (default or views or [None])[0]


def hit_face_normal(reference, element, link_instance):
    """Normale (coordinate host) della faccia piana colpita, o None se non e' leggibile. Per gli elementi di un link
    la faccia si legge dal documento del link e la normale si porta nel modello host."""
    try:
        if link_instance is not None:
            face = element.GetGeometryObjectFromReference(reference.CreateReferenceInLink())
        else:
            face = element.GetGeometryObjectFromReference(reference)
        if not isinstance(face, DB.PlanarFace):
            return None
        normal = face.FaceNormal
        return link_instance.GetTotalTransform().OfVector(normal) if link_instance is not None else normal
    except Exception:
        return None


def category_value(element):
    cat_id = element.Category.Id
    return cat_id.Value if hasattr(cat_id, "Value") else cat_id.IntegerValue  # Revit 2026+ / <= 2025


def find_clash(view3d, origin, direction, length):
    """Primo muro / pilastro / trave / curtain wall lungo il nuovo tratto (asse), anche nei link RVT e IFC.
    Ritorna (elemento, transform del link o None, distanza dall'origine, nome del modello, punto colpito, normale
    della faccia colpita o None) o None."""
    categories = List[DB.BuiltInCategory]()
    for cat in CLASH_CATEGORIES:
        categories.Add(cat)
    # con FindReferencesInRevitLinks il filtro vale anche per l'istanza di link: senza la categoria dei link
    # l'istanza viene scartata e nessun elemento dei link viene trovato. La categoria degli elementi del link si
    # controlla qui sotto
    categories.Add(DB.BuiltInCategory.OST_RvtLinks)
    wanted = set(int(c) for c in CLASH_CATEGORIES)
    intersector = DB.ReferenceIntersector(DB.ElementMulticategoryFilter(categories),
                                          DB.FindReferenceTarget.Face, view3d)
    intersector.FindReferencesInRevitLinks = True
    best = None
    for context in intersector.Find(origin, direction):
        distance = context.Proximity
        if distance > length or (best is not None and distance >= best[2]):
            continue
        reference = context.GetReference()
        hit = origin.Add(direction.Multiply(distance))
        host = doc.GetElement(reference.ElementId)
        if isinstance(host, DB.RevitLinkInstance):
            link_doc = host.GetLinkDocument()
            if link_doc is None or reference.LinkedElementId == DB.ElementId.InvalidElementId:
                continue
            element = link_doc.GetElement(reference.LinkedElementId)
            if element is None or element.Category is None or category_value(element) not in wanted:
                continue
            best = (element, host.GetTotalTransform(), distance, link_doc.Title, hit,
                    hit_face_normal(reference, element, host))
        else:
            best = (host, None, distance, "host model", hit, hit_face_normal(reference, host, None))
    return best


def clash_label(clash):
    element, _, distance, model, _, _ = clash
    category = element.Category.Name if element.Category is not None else "element"
    return "{} {} ({}) at {:.0f} mm".format(category, element.Id, model, distance * MM_PER_FT)


def parallel_turn(clash, origin, travel, signed_deg):
    """Angolo con segno per correre parallelo all'elemento colpito, e distanza dell'asse del tratto dalla faccia.
    Ritorna (angolo, distanza mm o None) oppure solleva TurnError se l'elemento non e' lineare."""
    element, transform, _, _, hit, face_normal = clash
    location = getattr(element, "Location", None)
    tangent, on_line, from_axis = None, hit, False
    if isinstance(location, DB.LocationCurve):
        curve = location.Curve if transform is None else location.Curve.CreateTransformed(transform)
        projection = curve.Project(hit)
        direction_3d = curve.ComputeDerivatives(projection.Parameter, False).BasisX
        flat = DB.XYZ(direction_3d.X, direction_3d.Y, 0)
        if flat.GetLength() > 1e-9:
            tangent, on_line, from_axis = flat.Normalize(), projection.XYZPoint, True
    if tangent is None:
        # senza asse orizzontale utilizzabile (DirectShape IFC, pilastro, pannello): lungo la faccia colpita
        if face_normal is None:
            raise TurnError("The new run hits {}: it has no horizontal axis and the face that was hit could not be "
                            "read, so the turn cannot follow it".format(clash_label(clash)))
        flat = DB.XYZ(-face_normal.Y, face_normal.X, 0)
        if flat.GetLength() < VERTICAL_FACE_TOL:
            raise TurnError("The new run hits {} on a horizontal face (floor, soffit): no direction to "
                            "follow".format(clash_label(clash)))
        tangent = flat.Normalize()
    horizontal_len = math.sqrt(travel.X ** 2 + travel.Y ** 2)
    forward = DB.XYZ(travel.X / horizontal_len, travel.Y / horizontal_len, 0)
    wanted = turn_direction(DB.XYZ(forward.X, forward.Y, 0), signed_deg, False)
    # dei due versi lungo il muro si tiene quello piu' vicino alla curva richiesta
    if tangent.DotProduct(wanted) < 0:
        tangent = tangent.Negate()
    angle = math.degrees(math.atan2(forward.CrossProduct(tangent).Z, forward.DotProduct(tangent)))
    if abs(angle) < MIN_ADJUSTED_DEG:
        raise TurnError("The new run hits {}, which is almost parallel to the existing run: following it "
                        "means no turn ({:.1f} degrees)".format(clash_label(clash), angle))
    # distanza in pianta dell'estremo dalla linea del muro (retta per il punto colpito, lungo la tangente)
    axis_distance = abs((origin.X - on_line.X) * tangent.Y - (origin.Y - on_line.Y) * tangent.X)
    width = getattr(element, "Width", None)  # solo i muri hanno lo spessore
    if not from_axis:
        face_mm = axis_distance * MM_PER_FT  # la linea passa gia' per la faccia colpita
    else:
        face_mm = (axis_distance - width / 2.0) * MM_PER_FT if isinstance(width, float) else None
    return angle, face_mm


CLASH_VIEW = []  # vista 3D trovata al primo uso


def avoid_clash(origin, travel, signed_deg, direction, length, keep_slope):
    """Se il nuovo tratto colpisce un muro, gira parallelo al muro. Ritorna (angolo, direzione, nota per il report);
    solleva TurnError se l'interferenza non si puo' evitare."""
    if not CLASH_VIEW:
        CLASH_VIEW.append(clash_view())
    view3d = CLASH_VIEW[0]
    if view3d is None:
        return signed_deg, direction, "Not checked: no 3D view in the model"
    clash = find_clash(view3d, origin, direction, length)
    if clash is None:
        return signed_deg, direction, "None"
    angle, face_mm = parallel_turn(clash, origin, travel, signed_deg)
    new_direction = turn_direction(travel, angle, keep_slope)
    note = "Hit {}: turned {:.1f} degrees {} to run parallel".format(
        clash_label(clash), abs(angle), "left" if angle > 0 else "right")
    if face_mm is not None:
        note += ", run axis {:.0f} mm from the wall face".format(face_mm)
    again = find_clash(view3d, origin, new_direction, length)
    if again is not None:
        note += ". WARNING: the new run still hits {}".format(clash_label(again))
    return angle, new_direction, note


def tray_turned_copy(element, conn, settings, signed_deg):
    """Passerella nuova con lo stesso orientamento della sezione (es. di costa): CableTray.Create la farebbe sempre
    piatta e Revit non ruota una passerella orizzontale attorno al proprio asse. Si copia il tratto di partenza, lo si
    mette in prosecuzione dall'estremo libero e lo si ruota attorno alla verticale passante per l'estremo: la
    rotazione attorno a Z conserva sia l'orientamento della sezione sia la pendenza."""
    travel = run_direction_at(element, conn)
    start = conn.Origin
    copy_id = list(DB.ElementTransformUtils.CopyElement(doc, element.Id, DB.XYZ.Zero))[0]
    new_element = doc.GetElement(copy_id)
    new_element.Location.Curve = DB.Line.CreateBound(
        start, start.Add(travel.Multiply(settings["length_mm"] / MM_PER_FT)))
    doc.Regenerate()
    # signed_deg positivo = antiorario visto dall'alto (sinistra), eventualmente corretto per seguire un muro
    axis = DB.Line.CreateBound(start, start.Add(DB.XYZ.BasisZ))
    DB.ElementTransformUtils.RotateElement(doc, copy_id, axis, math.radians(signed_deg))
    mark = new_element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
    if mark is not None and not mark.IsReadOnly:
        mark.Set("")  # la copia eredita il Mark: duplicato
    doc.Regenerate()
    return new_element


def create_segment(element, conn, end_point, level_id, settings, signed_deg):
    """Crea il nuovo tratto (senza garantire il raccordo). Ritorna l'elemento creato."""
    type_id = element.GetTypeId()
    if isinstance(element, Pipe):
        return Pipe.Create(doc, type_id, level_id, conn, end_point)
    if isinstance(element, Duct):
        return Duct.Create(doc, type_id, level_id, conn, end_point)
    sloped = abs(run_direction_at(element, conn).Z) > 1e-6
    if isinstance(element, CableTray) and (settings["keep_slope"] or not sloped):
        return tray_turned_copy(element, conn, settings, signed_deg)
    if isinstance(element, CableTray):
        new_element = CableTray.Create(doc, type_id, conn.Origin, end_point, level_id)
        copy_tray_dimensions(element, new_element)
    else:
        new_element = Conduit.Create(doc, type_id, conn.Origin, end_point, level_id)
        copy_conduit_diameter(element, new_element)
    return new_element


def ensure_elbow(element, conn, new_element):
    """Garantisce il gomito tra il tratto esistente e quello nuovo. Ritorna il gomito."""
    fitting = find_fitting(conn, element)
    if fitting is not None:
        return fitting
    new_conn = nearest_connector(new_element, conn.Origin)
    if new_conn.IsConnectedTo(conn):
        conn.DisconnectFrom(new_conn)
    return doc.Create.NewElbowFitting(conn, new_conn)


# parametri istanza copiati dal tratto di partenza: quelli di progetto/condivisi e pochi built-in sicuri.
# Esclusi Mark (i duplicati generano avvisi), quote, offset, lunghezze e sistema: si ricalcolano da soli.
COPIED_BUILTIN = (DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, DB.BuiltInParameter.ELEM_PARTITION_PARAM,
                  DB.BuiltInParameter.PHASE_CREATED, DB.BuiltInParameter.PHASE_DEMOLISHED)
COPYABLE_STORAGE = (DB.StorageType.String, DB.StorageType.Integer, DB.StorageType.Double,
                    DB.StorageType.ElementId)


def copy_parameters(source, target):
    """Copia i parametri istanza dal tratto di partenza al nuovo. Ritorna il numero di valori copiati."""
    copied = 0
    for param in source.Parameters:
        try:
            if param.IsReadOnly or not param.HasValue or param.StorageType not in COPYABLE_STORAGE:
                continue
            bip = getattr(param.Definition, "BuiltInParameter", None)
            if bip is not None and bip != DB.BuiltInParameter.INVALID:
                if bip not in COPIED_BUILTIN:
                    continue
                dest = target.get_Parameter(bip)
            else:
                dest = target.LookupParameter(param.Definition.Name)
            if dest is None or dest.IsReadOnly or dest.StorageType != param.StorageType:
                continue
            if param.StorageType == DB.StorageType.String:
                value = param.AsString()
                if value is None:
                    continue
                dest.Set(value)
            elif param.StorageType == DB.StorageType.Integer:
                dest.Set(param.AsInteger())
            elif param.StorageType == DB.StorageType.Double:
                dest.Set(param.AsDouble())
            else:
                dest.Set(param.AsElementId())
            copied += 1
        except Exception:
            continue
    return copied


def insulation_specs(source):
    """Isolamento (e rivestimento interno dei condotti) del tratto di partenza: [(classe, id tipo, spessore)]."""
    if isinstance(source, Pipe):
        kinds = ((PipeInsulation, DB.InsulationLiningBase.GetInsulationIds),)
    elif isinstance(source, Duct):
        kinds = ((DuctInsulation, DB.InsulationLiningBase.GetInsulationIds),
                 (DuctLining, DB.InsulationLiningBase.GetLiningIds))
    else:
        return []
    specs = []
    for cls, getter in kinds:
        for insulation_id in getter(doc, source.Id):
            insulation = doc.GetElement(insulation_id)
            specs.append((cls, getter, insulation.GetTypeId(), insulation.Thickness))
    return specs


def apply_insulation(specs, hosts):
    """Crea lo stesso isolamento sui nuovi tratti e sui raccordi. Ritorna il numero di elementi isolati."""
    made = 0
    for host in hosts:
        if host is None:
            continue
        for cls, getter, type_id, thickness in specs:
            try:
                if len(list(getter(doc, host.Id))) > 0:
                    continue
                cls.Create(doc, host.Id, type_id, thickness)
                made += 1
            except Exception:
                continue
    return made


def inherit(source, new_elements, fittings):
    """Parametri e isolamento del tratto di partenza sui nuovi elementi. Ritorna una nota per il report."""
    try:
        count = sum(copy_parameters(source, e) for e in new_elements)
        specs = insulation_specs(source)
        text = "{} parameter values copied".format(count)
        if specs:
            text += ", insulation on {} elements".format(
                apply_insulation(specs, list(new_elements) + list(fittings)))
        return text
    except Exception as ex:
        return "parameters and insulation not copied: {}".format(ex)

def check_remaining_space(element, new_element):
    """Dopo il gomito, i due tratti devono avere ancora una lunghezza utile."""
    doc.Regenerate()
    for target, name in ((element, "existing element"), (new_element, "new run")):
        remaining = target.Location.Curve.Length * MM_PER_FT
        if remaining < MIN_REMAINING_MM:
            raise TurnError("Not enough space for the elbow: the {} would be only {:.0f} mm long "
                            "(minimum {:.0f} mm)".format(name, remaining, MIN_REMAINING_MM))


def process_pick(reference, settings):
    """Elabora un elemento. Ritorna un dict con l'esito (usato per il report)."""
    element = doc.GetElement(reference.ElementId)
    result = {"element": element, "category": category_name(element), "status": "", "ok": False,
              "new": None, "fitting": "", "added_mm": None, "z_mm": None, "level": "", "rel_mm": None,
              "clash": "", "angle": None, "side": ""}

    click = reference.GlobalPoint
    if click is None:
        click = element.Location.Curve.Evaluate(0.5, True)

    conn = nearest_free_connector(element, click)
    if conn is None:
        result["status"] = "No free end (both ends are connected)"
        return result

    travel = run_direction_at(element, conn)
    if abs(travel.Z) > VERTICAL_TOL:
        result["status"] = "Vertical end: right and left are not defined"
        return result

    # il lato si sceglie con un secondo clic, per ogni tratto
    try:
        side_click = uidoc.Selection.PickPoint("Turn: click on the side the new run must turn to "
                                               "(ESC to skip this element)")
    except OperationCanceledException:
        return None
    side = side_of(travel, conn.Origin, side_click)
    if side is None:
        result["status"] = "Not created. The side click is in line with the run: no side"
        return result
    result["side"] = side
    signed_deg = signed_turn_deg(settings["angle_deg"], side)
    direction = turn_direction(travel, signed_deg, settings["keep_slope"])

    level_id = get_level_id(element)
    if level_id is None:
        result["status"] = "No reference level on the element"
        return result
    ref_level = doc.GetElement(level_id)
    result["level"] = ref_level.Name

    length = settings["length_mm"] / MM_PER_FT
    if settings.get("avoid_clash"):
        try:
            signed_deg, direction, result["clash"] = avoid_clash(conn.Origin, travel, signed_deg, direction,
                                                                length, settings["keep_slope"])
        except TurnError as ex:
            result["status"] = "Not created. {}".format(ex)
            return result
    end_point = conn.Origin.Add(direction.Multiply(length))
    result["angle"] = abs(signed_deg)  # angolo effettivo: scelto al clic o corretto per seguire un muro

    transaction = DB.Transaction(doc, "Turn {}".format(side))
    transaction.Start()
    try:
        new_element = create_segment(element, conn, end_point, level_id, settings, signed_deg)
        fitting = ensure_elbow(element, conn, new_element)
        check_remaining_space(element, new_element)
        inherited = inherit(element, [new_element], [fitting])
        transaction.Commit()
    except TurnError as ex:
        transaction.RollBack()
        result["status"] = "Not created. {}".format(ex)
        return result
    except Exception as ex:
        transaction.RollBack()
        result["status"] = ("Not created. The elbow could not be placed (not enough space, or no valid "
                            "fitting in the routing preferences): {}".format(ex))
        return result

    result.update({"ok": True, "status": "Created ({})".format(inherited), "new": new_element,
                   "fitting": fitting_label(fitting), "added_mm": settings["length_mm"],
                   "z_mm": end_point.Z * MM_PER_FT,
                   "rel_mm": (end_point.Z - level_elevation(ref_level)) * MM_PER_FT})
    return result


def fmt(value):
    return "" if value is None else "{:.1f}".format(value)


def settings_text(settings):
    return "{:g} mm, {:g} degrees, {}{}".format(
        settings["length_mm"], settings["angle_deg"],
        "keep slope" if settings["keep_slope"] else "horizontal",
        ", parallel to walls on clash" if settings.get("avoid_clash") else "")


def print_report(results, settings):
    output.print_md("# Turn report")
    output.print_md("Settings: **{}**".format(settings_text(settings)))
    done = [r for r in results if r["ok"]]
    output.print_md("Elements picked: **{}** | runs created: **{}** | not processed: **{}**".format(
        len(results), len(done), len(results) - len(done)))
    if not results:
        output.print_md("No element was picked.")
        return

    rows = []
    for i, r in enumerate(results, 1):
        rows.append([
            i, output.linkify(r["element"].Id), r["category"], r["status"], r["side"],
            fmt(r["angle"]), fmt(r["added_mm"]),
            output.linkify(r["new"].Id) if r["new"] is not None else "", r["fitting"],
            fmt(r["z_mm"]), r["level"], fmt(r["rel_mm"]), r["clash"],
        ])
    output.print_table(
        table_data=rows,
        columns=["#", "Element", "Category", "Status", "Side", "Angle [deg]", "Added length [mm]", "New element", "Fitting",
                 "End Z, internal coordinates [mm]", "Reference level", "End Z vs reference level [mm]",
                 "Clash"])


def load_settings(config):
    settings = dict(DEFAULTS)
    try:
        settings["length_mm"] = float(config.get_option("length_mm", DEFAULTS["length_mm"]))
        settings["angle_deg"] = float(config.get_option("angle_deg", DEFAULTS["angle_deg"]))
        keep = config.get_option("keep_slope", DEFAULTS["keep_slope"])
        if not isinstance(keep, bool):
            keep = str(keep).strip().lower() in ("true", "1", "yes")
        settings["keep_slope"] = keep
        avoid = config.get_option("avoid_clash", DEFAULTS["avoid_clash"])
        if not isinstance(avoid, bool):
            avoid = str(avoid).strip().lower() in ("true", "1", "yes")
        settings["avoid_clash"] = avoid
    except Exception:
        pass
    return settings


def save_settings(config, settings):
    try:
        config.length_mm = settings["length_mm"]
        config.angle_deg = settings["angle_deg"]
        config.keep_slope = settings["keep_slope"]
        config.avoid_clash = settings["avoid_clash"]
        script.save_config()
    except Exception:
        pass


def ask_settings(settings):
    """Unica finestra del comando: angolo, lunghezza, pendenza, interferenze. Ritorna le impostazioni o None se
    annullata (ESC). Non forms.CommandSwitchWindow: si chiude appena perde il focus, cioe' subito dopo il clic
    nel modello."""
    return Turn_ui.AngleWindow(script.get_bundle_file(ANGLE_XAML_FILE_NAME), settings, QUICK_ANGLES,
                               MIN_ANGLE, MAX_ANGLE).show()


def pick_loop(settings):
    results = []
    selection_filter = MepCurveFilter()
    prompt = "Turn ({}): click near the free end (ESC to finish)".format(settings_text(settings))
    while True:
        try:
            reference = uidoc.Selection.PickObject(ObjectType.Element, selection_filter, prompt)
        except OperationCanceledException:
            break
        result = process_pick(reference, settings)
        if result is not None:  # None: lato annullato con ESC, l'elemento viene saltato
            results.append(result)
    return results


def main():
    config = script.get_config(CONFIG_NAME)
    settings = load_settings(config)

    # tutte le opzioni (angolo, lunghezza, pendenza, interferenze) in un'unica finestra a ogni avvio, una volta sola:
    # valgono per tutti i clic di questa sessione e si salvano come valori proposti la volta successiva
    new_settings = ask_settings(settings)
    if new_settings is None:
        script.exit()
    settings = new_settings
    save_settings(config, settings)

    warnings = []
    if not settings["keep_slope"]:
        warnings.append("new runs forced horizontal")
    if settings.get("avoid_clash"):
        warnings.append("turns parallel to walls on clash")

    if warnings:
        message = "Turn: {} (change it in the window at the next start)".format(", ".join(warnings))
        with forms.WarningBar(title=message):
            results = pick_loop(settings)
    else:
        results = pick_loop(settings)

    problems = len([r for r in results if not r["ok"]])
    if want_report(len(results), problems, "runs created"):
        print_report(results, settings)


main()
