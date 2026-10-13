# -*- coding: utf-8 -*-
__title__ = "Clash\nOverpass"
__doc__ = """Version = 0.4
Date    = 12.10.2026
________________________________________________________________
Route with clash check, solved automatically.
1. Select one or more pipes, ducts or cable trays (or select them first)
2. A window (only at the start) shows how many clashes were found and
   the settings: pass over, pass under or report only, the clearance
   from the obstacle, the merge distance and the angle of the rises and
   drops, plus the clash detection. On OK all the clashes are solved
   with these settings (report only: nothing is changed)
3. Click a point: the runs are extended to it from their free end, the
   check runs again and new clashes are solved in the same way. ESC to
   finish. The report shows the number of clashes found
Clash detection (kept for the next time): distribution elements only
(pipes, ducts, cable trays, conduits, fittings, accessories, insulation)
or everything (all model elements of this file; linked models are not
checked). Changing it in the window searches again.
Merge: obstacles on the same run get one single overpass when the run
left at its elevation between two overpasses would be shorter than the
merge distance (default 2000 mm).
If overpasses keep creating new clashes, after 10 automatic rounds in a
row the remaining ones are left as they are and listed in the report.
A manual overpass between two clicked points is done by Quick Overpass.
The new parts keep the parameters and the insulation of the starting element.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import math

try:
    import traceback
except ImportError:
    traceback = None

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB.Plumbing import Pipe, PipeInsulation
from Autodesk.Revit.DB.Mechanical import Duct, DuctInsulation, DuctLining
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Collections.Generic import List

from pyrevit import revit, script, DB

import ClashOverpass_ui as Overpass_ui  # nome unico: con Rocket mode i moduli restano in memoria tra i comandi

XAML_FILE_NAME = "Overpass_form.xaml"
CONFIG_NAME = "ESA_Overpass"

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
STRAIGHT_TOL = 0.999
HORIZONTAL_TOL = 1e-4
MIN_REMAINING_MM = 20.0  # lunghezza minima di ogni tratto dopo l'inserimento dei gomiti
TAIL_MIN_MM = 300.0  # tratto orizzontale minimo oltre lo scavalco, dopo il gomito: sotto Revit inverte il tratto

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


class SameElementFilter(ISelectionFilter):
    """Accetta solo l'elemento indicato (secondo click sullo stesso tratto)."""

    def __init__(self, element_id):
        self.element_id = element_id

    def AllowElement(self, element):
        return element.Id == self.element_id

    def AllowReference(self, reference, point):
        return False


STAGE = [""]  # passo in corso, per i messaggi di errore
PHYSICAL_TYPES = (DB.ConnectorType.End, DB.ConnectorType.Curve, DB.ConnectorType.Physical)


class OverpassError(Exception):
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


def nearest_connector(element, point):
    best, best_dist = None, None
    for c in element.ConnectorManager.Connectors:
        dist = c.Origin.DistanceTo(point)
        if best is None or dist < best_dist:
            best, best_dist = c, dist
    return best


def connector_manager(element):
    """ConnectorManager di un tratto MEP o di un raccordo (FamilyInstance)."""
    manager = getattr(element, "ConnectorManager", None)
    if manager is None:
        model = getattr(element, "MEPModel", None)
        manager = model.ConnectorManager if model is not None else None
    return manager


def free_connector_near(element, point):
    """Connettore non collegato, il piu' vicino al punto."""
    manager = connector_manager(element)
    if manager is None:
        return None
    best, best_dist = None, None
    for c in manager.Connectors:
        if c.IsConnected:
            continue
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


def copy_tray_dimensions(source, target, swap):
    """Copia larghezza/altezza della passerella; sui tratti verticali si scambiano (come in MEPConnect)."""
    src_w = source.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
    src_h = source.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
    tgt_w = target.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
    tgt_h = target.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
    if not (src_w and src_h and tgt_w and tgt_h):
        return
    width, height = src_w.AsDouble(), src_h.AsDouble()
    if swap:
        width, height = height, width
    if not tgt_w.IsReadOnly:
        tgt_w.Set(width)
    if not tgt_h.IsReadOnly:
        tgt_h.Set(height)


def copy_conduit_diameter(source, target):
    for bip in (DB.BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM,
                DB.BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM):
        src = source.get_Parameter(bip)
        tgt = target.get_Parameter(bip)
        if src and tgt and not tgt.IsReadOnly:
            tgt.Set(src.AsDouble())
            return


def create_segment(dims_source, conn, end_point, level_id, vertical):
    """Crea un tratto da conn a end_point con tipo e dimensioni di dims_source."""
    type_id = dims_source.GetTypeId()
    if isinstance(dims_source, Pipe):
        # Pipe.Create da connettore puo' inserire da solo il gomito: si verifica in ensure_connection
        return Pipe.Create(doc, type_id, level_id, conn, end_point)
    if isinstance(dims_source, Duct):
        return Duct.Create(doc, type_id, level_id, conn, end_point)
    if isinstance(dims_source, CableTray):
        new_element = CableTray.Create(doc, type_id, conn.Origin, end_point, level_id)
        copy_tray_dimensions(dims_source, new_element, swap=vertical)
    else:
        new_element = Conduit.Create(doc, type_id, conn.Origin, end_point, level_id)
        copy_conduit_diameter(dims_source, new_element)
    return new_element


def ensure_connection(owner, conn, new_element, straight):
    """Garantisce il collegamento col nuovo tratto. Ritorna il gomito, o None se continuita' diretta."""
    new_conn = nearest_connector(new_element, conn.Origin)
    if straight:
        if not new_conn.IsConnectedTo(conn):
            new_conn.ConnectTo(conn)
        return None
    fitting = find_fitting(conn, owner)
    if fitting is None:
        # collegamento diretto senza raccordo: si scollega e si posiziona il gomito esplicitamente
        if new_conn.IsConnectedTo(conn):
            conn.DisconnectFrom(new_conn)
        fitting = doc.Create.NewElbowFitting(conn, new_conn)
    return fitting


def extend(owner, conn, end_point, dims_source, level_id, vertical):
    """Aggiunge un tratto da conn a end_point. Ritorna (nuovo elemento, connettore libero all'estremo, raccordo)."""
    direction = end_point.Subtract(conn.Origin).Normalize()
    dot = conn.CoordinateSystem.BasisZ.DotProduct(direction)
    if dot < -STRAIGHT_TOL:
        raise OverpassError("A new run would fold back onto the previous element")
    straight = dot > STRAIGHT_TOL
    new_element = create_segment(dims_source, conn, end_point, level_id, vertical)
    fitting = ensure_connection(owner, conn, new_element, straight)
    return new_element, nearest_connector(new_element, end_point), fitting


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

def check_remaining_space(parts):
    """parts: [(nome, elemento)]. Dopo i gomiti ogni tratto deve avere ancora una lunghezza utile."""
    doc.Regenerate()
    for name, element in parts:
        remaining = element.Location.Curve.Length * MM_PER_FT
        if remaining < MIN_REMAINING_MM:
            raise OverpassError("Not enough space for the elbows: the {} would be only {:.0f} mm long "
                                "(minimum {:.0f} mm)".format(name, remaining, MIN_REMAINING_MM))


def project_on_run(reference, element, curve):
    """Proietta il punto cliccato sull'asse del tratto."""
    point = reference.GlobalPoint
    if point is None:
        point = curve.Evaluate(0.5, True)
    return curve.Project(point).XYZPoint


def build_overpass(element, curve, p1, p2, offset_ft, inset_ft, vertical_legs, level_id):
    """Costruisce lo scavalco. Ritorna (tratto sollevato, elementi nuovi, numero di gomiti).

    inset_ft: ingombro orizzontale di ogni salita, verso l'interno (0 se verticale).
    """
    start, end = curve.GetEndPoint(0), curve.GetEndPoint(1)
    shift = DB.XYZ(0, 0, offset_ft)
    inset = curve.Direction.Multiply(inset_ft)
    top_start = p1.Add(inset).Add(shift)
    top_end = p2.Subtract(inset).Add(shift)

    STAGE[0] = "reading the connections at the far end of the run"
    far_conn = nearest_connector(element, end)
    # solo connettori fisici: AllRefs include anche il connettore logico del sistema MEP, senza posizione
    far_refs = [r for r in far_conn.AllRefs
                if r.Owner.Id != element.Id and r.ConnectorType in PHYSICAL_TYPES]
    # i connettori degli altri elementi diventano non validi dopo le modifiche: si salvano id e posizione
    far_info = [(r.Owner.Id, r.Origin) for r in far_refs]
    STAGE[0] = "disconnecting the far end of the run"
    for ref in far_refs:
        far_conn.DisconnectFrom(ref)

    STAGE[0] = "shortening the original run"
    element.Location.Curve = DB.Line.CreateBound(start, p1)
    doc.Regenerate()

    conn = nearest_connector(element, p1)
    STAGE[0] = "first leg"
    leg1, conn, f1 = extend(element, conn, top_start, element, level_id, vertical_legs)
    STAGE[0] = "raised run"
    top, conn, f2 = extend(leg1, conn, top_end, element, level_id, False)
    STAGE[0] = "second leg"
    leg2, conn, f3 = extend(top, conn, p2, element, level_id, vertical_legs)
    STAGE[0] = "last part of the run"
    tail, conn, f4 = extend(leg2, conn, end, element, level_id, False)

    STAGE[0] = "reconnecting the far end of the run"
    conn = nearest_connector(tail, end)
    for owner_id, origin in far_info:
        other = doc.GetElement(owner_id)
        if other is None:
            raise OverpassError("The element connected to the far end of the run no longer exists")
        target = free_connector_near(other, origin)
        if target is None:
            raise OverpassError("No free connector found on the element connected to the far end of the run")
        conn.ConnectTo(target)

    STAGE[0] = "checking the remaining space"
    check_remaining_space([("first part of the run", element), ("first leg", leg1),
                           ("raised run", top), ("second leg", leg2),
                           ("last part of the run", tail)])
    STAGE[0] = "copying parameters and insulation"
    inherited = inherit(element, [leg1, top, leg2, tail], [f1, f2, f3, f4])
    fittings = [f for f in (f1, f2, f3, f4) if f is not None]
    return top, [leg1, top, leg2, tail], len(fittings), inherited, fittings


def leg_inset_ft(offset_ft, angle_deg):
    """Ingombro orizzontale di una salita inclinata: |offset| / tan(angolo)."""
    if angle_deg >= 90 - 1e-6:
        return 0.0
    return abs(offset_ft) / math.tan(math.radians(angle_deg))


def process_run(element, ref1, ref2, offset_mm, angle_deg):
    """Elabora un tratto. Ritorna un dict con l'esito (usato per il report)."""
    result = {"element": element, "category": category_name(element), "status": "", "ok": False,
              "top": None, "new_count": 0, "elbows": 0, "z_mm": None, "level": "", "rel_mm": None}

    curve = element.Location.Curve
    if not isinstance(curve, DB.Line):
        result["status"] = "Only straight runs are supported"
        return result
    if abs(curve.Direction.Z) > HORIZONTAL_TOL:
        result["status"] = "Only horizontal runs are supported"
        return result

    level_id = get_level_id(element)
    if level_id is None:
        result["status"] = "No reference level on the element"
        return result
    ref_level = doc.GetElement(level_id)
    result["level"] = ref_level.Name

    start = curve.GetEndPoint(0)
    direction = curve.Direction
    points = sorted((project_on_run(ref1, element, curve), project_on_run(ref2, element, curve)),
                    key=lambda p: p.Subtract(start).DotProduct(direction))
    p1, p2 = points
    min_ft = MIN_REMAINING_MM / MM_PER_FT
    t1 = p1.Subtract(start).DotProduct(direction)
    t2 = p2.Subtract(start).DotProduct(direction)
    if t1 < min_ft or curve.Length - t2 < min_ft or t2 - t1 < min_ft:
        result["status"] = ("Points too close to each other or to the ends of the run "
                            "(minimum {:.0f} mm)".format(MIN_REMAINING_MM))
        return result

    offset_ft = offset_mm / MM_PER_FT
    inset_ft = leg_inset_ft(offset_ft, angle_deg)
    raised_ft = (t2 - t1) - 2 * inset_ft
    if raised_ft < min_ft:
        result["status"] = ("Points too close for {:g} degree legs: each leg needs {:.0f} mm horizontally, "
                            "the points are {:.0f} mm apart".format(
                                angle_deg, inset_ft * MM_PER_FT, (t2 - t1) * MM_PER_FT))
        return result

    vertical_legs = inset_ft == 0.0
    transaction = DB.Transaction(doc, "Overpass")
    transaction.Start()
    try:
        top, new_elements, elbows, inherited, _ = build_overpass(element, curve, p1, p2, offset_ft, inset_ft,
                                                                 vertical_legs, level_id)
        transaction.Commit()
    except OverpassError as ex:
        transaction.RollBack()
        result["status"] = "Unchanged. {}".format(ex)
        return result
    except Exception as ex:
        transaction.RollBack()
        result["status"] = ("Unchanged. The overpass could not be built (step: {}): {}".format(STAGE[0], ex))
        result["detail"] = traceback.format_exc() if traceback else ""
        return result

    z = p1.Z + offset_ft
    result.update({"ok": True, "status": "Created ({})".format(inherited), "top": top, "new_count": len(new_elements),
                   "elbows": elbows, "z_mm": z * MM_PER_FT,
                   "rel_mm": (z - level_elevation(ref_level)) * MM_PER_FT})
    return result


def fmt(value):
    return "" if value is None else "{:.1f}".format(value)


def print_report(results, offset_mm, angle_deg):
    output.print_md("# Overpass report")
    output.print_md("Offset: **{:+g} mm** | leg angle: **{:g} degrees** from the horizontal".format(
        offset_mm, angle_deg))
    done = [r for r in results if r["ok"]]
    output.print_md("Runs picked: **{}** | overpasses created: **{}** | left unchanged: **{}**".format(
        len(results), len(done), len(results) - len(done)))
    output.print_md("New elements created: **{}** | elbows placed: **{}**".format(
        sum(r["new_count"] for r in done), sum(r["elbows"] for r in done)))
    if not results:
        output.print_md("No run was picked.")
        return

    rows = []
    for i, r in enumerate(results, 1):
        rows.append([
            i, output.linkify(r["element"].Id), r["category"], r["status"],
            r["new_count"] or "", r["elbows"] or "",
            output.linkify(r["top"].Id) if r["top"] is not None else "",
            fmt(r["z_mm"]), r["level"], fmt(r["rel_mm"]),
        ])
    output.print_table(
        table_data=rows,
        columns=["#", "Run", "Category", "Status", "New elements", "Elbows", "Raised run",
                 "Raised run Z, internal coordinates [mm]", "Reference level",
                 "Raised run Z vs reference level [mm]"])

    for i, r in enumerate(results, 1):
        if r.get("detail"):
            output.print_md("**Technical detail for run #{}**".format(i))
            output.print_code(r["detail"])


# ---------------------------------------------------------------------------
# CLICK: percorso con controllo delle interferenze
# ---------------------------------------------------------------------------
CLASH_XAML_FILE_NAME = "Clash_form.xaml"
# scope: "distribution" = solo elementi distributivi, "all" = tutti gli elementi di modello
# merge_mm: scavalchi vicini sullo stesso tratto diventano uno solo se il tratto a quota originale fra i due
# sarebbe piu' corto di questa distanza
ROUTE_DEFAULTS = {"action": "over", "clearance_mm": 50.0, "angle_deg": 90.0, "scope": "distribution",
                  "merge_mm": 2000.0}
SCOPE_LABELS = {"distribution": "distribution elements", "all": "all model elements"}
CORRIDOR_HALF_HEIGHT_FT = 30.0  # quota cercata sopra e sotto il tratto per misurare l'ostacolo (circa 9 m)
BBOX_TOL_FT = 0.01
ROUTE_CLASSES = (Pipe, Duct, CableTray)
DISTRIBUTION_CATEGORIES = (
    DB.BuiltInCategory.OST_PipeCurves, DB.BuiltInCategory.OST_DuctCurves, DB.BuiltInCategory.OST_CableTray,
    DB.BuiltInCategory.OST_Conduit, DB.BuiltInCategory.OST_FlexPipeCurves, DB.BuiltInCategory.OST_FlexDuctCurves,
    DB.BuiltInCategory.OST_PipeFitting, DB.BuiltInCategory.OST_DuctFitting,
    DB.BuiltInCategory.OST_CableTrayFitting, DB.BuiltInCategory.OST_ConduitFitting,
    DB.BuiltInCategory.OST_PipeAccessory, DB.BuiltInCategory.OST_DuctAccessory,
    DB.BuiltInCategory.OST_PipeInsulations, DB.BuiltInCategory.OST_DuctInsulations,
    DB.BuiltInCategory.OST_DuctLinings,
)
# controllo "tutto": categorie di modello senza un ingombro fisico da scavalcare
NON_PHYSICAL_CATEGORIES = set(int(c) for c in (
    DB.BuiltInCategory.OST_Rooms, DB.BuiltInCategory.OST_MEPSpaces, DB.BuiltInCategory.OST_Areas,
    DB.BuiltInCategory.OST_HVAC_Zones, DB.BuiltInCategory.OST_RvtLinks, DB.BuiltInCategory.OST_Cameras,
    DB.BuiltInCategory.OST_VolumeOfInterest, DB.BuiltInCategory.OST_SectionBox, DB.BuiltInCategory.OST_Lines,
    DB.BuiltInCategory.OST_SketchLines,
))


class RouteFilter(ISelectionFilter):
    """Tubi, canali e passerelle."""

    def AllowElement(self, element):
        return isinstance(element, ROUTE_CLASSES)

    def AllowReference(self, reference, point):
        return False


def id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


def is_horizontal_line(element):
    curve = element.Location.Curve
    return isinstance(curve, DB.Line) and abs(curve.Direction.Z) < HORIZONTAL_TOL


def half_sizes(element):
    """(meta' larghezza, meta' altezza, sezione tonda) della sezione esterna, in piedi."""
    def value(bip):
        param = element.get_Parameter(bip)
        return param.AsDouble() if param is not None and param.HasValue and param.AsDouble() > 0 else None

    round_section = False
    if isinstance(element, Pipe):
        width = height = value(DB.BuiltInParameter.RBS_PIPE_OUTER_DIAMETER) or \
            value(DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        round_section = True
    elif isinstance(element, CableTray):
        width = value(DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
        height = value(DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
    else:
        width = value(DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        height = value(DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if not (width and height):
            width = height = value(DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
            round_section = True
    if not (width and height):
        raise OverpassError("Cannot read the size of element {}".format(element.Id))
    return width / 2.0, height / 2.0, round_section


def connected_ids(element):
    """Elementi collegati fisicamente (raccordi alle estremita'): non sono interferenze."""
    ids = set()
    manager = connector_manager(element)
    if manager is None:
        return ids
    for conn in manager.Connectors:
        try:
            for ref in conn.AllRefs:
                if ref.Owner.Id != element.Id and ref.ConnectorType in PHYSICAL_TYPES:
                    ids.add(id_value(ref.Owner.Id))
        except Exception:
            continue
    return ids


def run_pieces(run):
    """Tratti (tubi, canali, passerelle) che compongono il percorso, esclusi i raccordi."""
    pieces = []
    for eid in run["chain"].values():
        element = doc.GetElement(eid)
        if element is not None and isinstance(element, ROUTE_CLASSES):
            pieces.append(element)
    return pieces


def is_physical_obstacle(element):
    """Controllo "tutto": elementi di modello con un ingombro fisico (no locali, vani, link, telecamere...)."""
    category = element.Category
    if category is None or category.CategoryType != DB.CategoryType.Model:
        return False
    return id_value(category.Id) not in NON_PHYSICAL_CATEGORIES


def find_clashes(runs, ignored, scope):
    """Interferenze tra i tratti dei percorsi e gli altri elementi: solo distributivi o tutti (scope).
    ignored: {(n. percorso, id ostacolo)} gia' decise, da non riproporre."""
    clashes, seen = [], set()
    category_filter = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory](DISTRIBUTION_CATEGORIES))
    tol = DB.XYZ(BBOX_TOL_FT, BBOX_TOL_FT, BBOX_TOL_FT)
    for run in runs:
        for piece in run_pieces(run):
            box = piece.get_BoundingBox(None)
            if box is None:
                continue
            skip = set(run["chain"].keys()) | connected_ids(piece)
            piece_id = id_value(piece.Id)
            try:
                collector = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()
                if scope != "all":
                    collector = collector.WherePasses(category_filter)
                found = list(collector
                             .WherePasses(DB.BoundingBoxIntersectsFilter(DB.Outline(box.Min.Subtract(tol),
                                                                                    box.Max.Add(tol))))
                             .WherePasses(DB.ElementIntersectsElementFilter(piece)))
            except Exception:
                continue
            if scope == "all":
                found = [e for e in found if is_physical_obstacle(e)]
            for obstacle in found:
                obstacle_id = id_value(obstacle.Id)
                if obstacle_id in skip or (run["index"], obstacle_id) in ignored:
                    continue
                # isolamento dei tratti stessi o dei raccordi collegati: non e' un ostacolo
                if isinstance(obstacle, DB.InsulationLiningBase) and id_value(obstacle.HostElementId) in skip:
                    continue
                pair = (min(piece_id, obstacle_id), max(piece_id, obstacle_id))
                if pair in seen:
                    continue
                seen.add(pair)
                clashes.append({"run": run, "piece": piece, "obstacle": obstacle})
    return clashes


def element_label(element):
    category = element.Category.Name if element.Category is not None else "?"
    try:
        name = element.Name
    except Exception:
        name = ""
    return "{} {}{}".format(category, element.Id, " ({})".format(name) if name else "")


def clash_label(clash):
    piece = clash["piece"]
    return "Run #{}: {} {}  x  {}".format(clash["run"]["index"], category_name(piece), piece.Id,
                                         element_label(clash["obstacle"]))


def element_solids(element):
    options = DB.Options()
    options.DetailLevel = DB.ViewDetailLevel.Fine
    solids = []
    geometry = element.get_Geometry(options)
    stack = list(geometry) if geometry is not None else []
    while stack:
        item = stack.pop()
        if isinstance(item, DB.Solid):
            if item.Volume > 1e-9:
                solids.append(item)
        elif isinstance(item, DB.GeometryInstance):
            stack.extend(list(item.GetInstanceGeometry()))
    return solids


def corridor_solid(piece, half_lateral):
    """Prisma verticale attorno al tratto, largo quanto il tratto piu' la distanza richiesta."""
    curve = piece.Location.Curve
    a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
    d = curve.Direction
    n = DB.XYZ(-d.Y, d.X, 0)
    z0 = a.Z - CORRIDOR_HALF_HEIGHT_FT
    points = []
    for p, side in ((a, -1), (b, -1), (b, 1), (a, 1)):  # antiorario visto dall'alto
        q = p.Add(n.Multiply(side * half_lateral))
        points.append(DB.XYZ(q.X, q.Y, z0))
    loop = DB.CurveLoop()
    for i in range(4):
        loop.Append(DB.Line.CreateBound(points[i], points[(i + 1) % 4]))
    return DB.GeometryCreationUtilities.CreateExtrusionGeometry(List[DB.CurveLoop]([loop]), DB.XYZ.BasisZ,
                                                                2 * CORRIDOR_HALF_HEIGHT_FT)


def box_corners(box_min, box_max, transform=None):
    corners = [DB.XYZ(x, y, z) for x in (box_min.X, box_max.X) for y in (box_min.Y, box_max.Y)
               for z in (box_min.Z, box_max.Z)]
    return [transform.OfPoint(c) for c in corners] if transform is not None else corners


def obstacle_extent(piece, obstacle, half_lateral):
    """Ingombro dell'ostacolo lungo il tratto: (t min, t max, z min, z max). Solo la parte vicina al tratto."""
    curve = piece.Location.Curve
    start, d = curve.GetEndPoint(0), curve.Direction
    points = []
    try:
        corridor = corridor_solid(piece, half_lateral)
        for solid in element_solids(obstacle):
            try:
                part = DB.BooleanOperationsUtils.ExecuteBooleanOperation(solid, corridor,
                                                                         DB.BooleanOperationsType.Intersect)
            except Exception:
                part = None
            if part is not None and part.Volume > 1e-9:
                box = part.GetBoundingBox()
                points += box_corners(box.Min, box.Max, box.Transform)
    except Exception:
        points = []
    if not points:
        # ripiego: ingombro dell'intero ostacolo (prudente)
        box = obstacle.get_BoundingBox(None)
        points = box_corners(box.Min, box.Max)
    ts = [p.Subtract(start).DotProduct(d) for p in points]
    zs = [p.Z for p in points]
    return min(ts), max(ts), min(zs), max(zs)


def plan_groups(piece, clashes, action, clearance_ft, angle_deg, merge_ft):
    """Raggruppa gli ostacoli vicini sullo stesso tratto: un solo scavalco per gruppo. Due scavalchi si uniscono se
    il tratto che resterebbe alla quota originale fra i due e' piu' corto di merge_ft (minimo MIN_REMAINING_MM);
    lo scavalco unito sta alla quota piu' sfavorevole dei due."""
    half_w, half_h, _ = half_sizes(piece)
    z_run = piece.Location.Curve.GetEndPoint(0).Z
    # la salita inizia prima dell'ostacolo: distanza + spessore della salita + ingombro del gomito
    margin = clearance_ft + 2.0 * half_h / math.sin(math.radians(angle_deg))
    items = []
    for clash in clashes:
        t0, t1, z0, z1 = obstacle_extent(piece, clash["obstacle"], half_w + clearance_ft)
        z = z1 + clearance_ft + half_h if action == "over" else z0 - clearance_ft - half_h
        items.append({"lo": t0 - margin, "hi": t1 + margin, "z": z, "clashes": [clash]})
    items.sort(key=lambda i: i["lo"])
    min_low_ft = max(MIN_REMAINING_MM / MM_PER_FT, merge_ft)
    groups = []
    for item in items:
        if groups:
            last = groups[-1]
            offset = max(abs(last["z"] - z_run), abs(item["z"] - z_run))
            # tratto a quota originale fra i due scavalchi = distanza fra i gruppi meno le due salite
            if item["lo"] - last["hi"] < 2 * leg_inset_ft(offset, angle_deg) + min_low_ft:
                last["hi"] = max(last["hi"], item["hi"])
                last["z"] = max(last["z"], item["z"]) if action == "over" else min(last["z"], item["z"])
                last["clashes"] += item["clashes"]
                continue
        groups.append(item)
    for group in groups:
        group["offset"] = group["z"] - z_run
    return groups


CLEARANCE_TOL_FT = 1.0 / MM_PER_FT  # tolleranza sulla distanza verificata (1 mm)
MAX_CLEARANCE_ATTEMPTS = 4


def profile_loop(origin, axis_w, axis_h, half_w, half_h, round_section):
    """Sezione (cerchio o rettangolo) centrata in origin: larghezza lungo axis_w, altezza lungo axis_h."""
    loop = DB.CurveLoop()
    if round_section:
        plane = DB.Plane.CreateByOriginAndBasis(origin, axis_w, axis_h)
        loop.Append(DB.Arc.Create(plane, half_w, 0.0, math.pi))
        loop.Append(DB.Arc.Create(plane, half_w, math.pi, 2 * math.pi))
        return loop
    corners = [origin.Add(axis_w.Multiply(sw * half_w)).Add(axis_h.Multiply(sh * half_h))
               for sw, sh in ((-1, -1), (1, -1), (1, 1), (-1, 1))]
    for i in range(4):
        loop.Append(DB.Line.CreateBound(corners[i], corners[(i + 1) % 4]))
    return loop


def inflated_segment(curve, lateral, half_w, half_h, round_section):
    """Solido di un tratto rettilineo, con la sezione allargata della distanza richiesta."""
    a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
    d = b.Subtract(a).Normalize()
    up = d.CrossProduct(lateral).Normalize()
    loop = profile_loop(a, lateral, up, half_w, half_h, round_section)
    return DB.GeometryCreationUtilities.CreateExtrusionGeometry(List[DB.CurveLoop]([loop]), d, a.DistanceTo(b))


def inflated_elbow(fitting, half_w, half_h, round_section):
    """Solido di rivoluzione del gomito, ricostruito dai due connettori, con la sezione allargata.
    Ritorna None se il gomito non si puo' ricostruire (il controllo resta non verificato)."""
    manager = connector_manager(fitting)
    conns = [c for c in manager.Connectors if c.ConnectorType in PHYSICAL_TYPES] if manager else []
    if len(conns) != 2:
        return None
    a, b = conns[0].Origin, conns[1].Origin
    da, db = conns[0].CoordinateSystem.BasisZ, conns[1].CoordinateSystem.BasisZ  # verso l'esterno del raccordo
    phi = math.acos(max(-1.0, min(1.0, da.Negate().DotProduct(db))))  # angolo di deviazione
    if phi < 1e-3:
        return None
    chord = b.Subtract(a)
    radius = chord.GetLength() / (2 * math.sin(phi / 2.0))
    if radius <= half_h:
        return None  # la sezione allargata attraverserebbe l'asse di rivoluzione
    toward_b = chord.Subtract(da.Multiply(chord.DotProduct(da))).Normalize()
    center = a.Add(toward_b.Multiply(radius))
    x = a.Subtract(center).Normalize()
    z = x.CrossProduct(b.Subtract(center).Normalize()).Normalize()  # asse: rotazione da a verso b
    frame = DB.Frame(center, x, z.CrossProduct(x), z)
    # larghezza lungo l'asse (orizzontale), altezza nel piano della curva
    loop = profile_loop(a, z, x, half_w, half_h, round_section)
    return DB.GeometryCreationUtilities.CreateRevolvedGeometry(frame, List[DB.CurveLoop]([loop]), 0.0, phi)


def clearance_check(parts, piece, obstacles, clearance_ft):
    """Le parti nuove, allargate della distanza, toccano gli ostacoli? Ritorna (violata, verificata del tutto)."""
    half_w, half_h, round_section = half_sizes(piece)
    grow = max(clearance_ft - CLEARANCE_TOL_FT, 0.0)
    half_w, half_h = half_w + grow, half_h + grow
    d = piece.Location.Curve.Direction
    lateral = DB.XYZ(-d.Y, d.X, 0).Normalize()
    obstacle_solids = []
    for obstacle in obstacles:
        obstacle_solids += element_solids(obstacle)
    verified = True
    for part in parts:
        try:
            if isinstance(part, ROUTE_CLASSES):
                solid = inflated_segment(part.Location.Curve, lateral, half_w, half_h, round_section)
            else:
                solid = inflated_elbow(part, half_w, half_h, round_section)
        except Exception:
            solid = None
        if solid is None:
            verified = False
            continue
        for obstacle_solid in obstacle_solids:
            try:
                common = DB.BooleanOperationsUtils.ExecuteBooleanOperation(solid, obstacle_solid,
                                                                           DB.BooleanOperationsType.Intersect)
            except Exception:
                verified = False
                continue
            if common is not None and common.Volume > 1e-6:
                return True, verified
    return False, verified


def grow_free_end(piece, idx, amount_ft):
    """Allunga di amount_ft l'estremo idx del tratto, nella sua direzione, se e' libero. Ritorna True se fatto.
    Usato quando dopo l'ostacolo manca tratto: il punto di arrivo cliccato era troppo corto per lo scavalco."""
    curve = piece.Location.Curve
    end = curve.GetEndPoint(idx)
    if nearest_connector(piece, end).IsConnected:
        return False
    other = curve.GetEndPoint(1 - idx)
    d = end.Subtract(other).Normalize()
    new_end = end.Add(d.Multiply(amount_ft))
    transaction = DB.Transaction(doc, "Overpass extend to clear")
    transaction.Start()
    try:
        piece.Location.Curve = DB.Line.CreateBound(other, new_end) if idx == 1 else \
            DB.Line.CreateBound(new_end, other)
        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise
    return True


def apply_group(piece, group, angle_deg, clearance_ft):
    """Scavalco di un gruppo di ostacoli, con verifica della distanza dopo la costruzione: se salite o gomiti
    stanno troppo vicini all'ostacolo si annulla, si allarga lo scavalco e si riprova. Se dopo l'ostacolo manca
    tratto e l'estremo e' libero, lo si allunga della quantita' mancante (grown).
    Ritorna (nuovi tratti, raccordi, allargamento in mm, distanza verificata del tutto, allungamento in mm)."""
    obstacles = [c["obstacle"] for c in group["clashes"]]
    min_ft = MIN_REMAINING_MM / MM_PER_FT
    # il gomito finale ha bisogno di spazio: il tratto dopo lo scavalco e' almeno TAIL_MIN_MM (o 3 diametri)
    tail_ft = max(TAIL_MIN_MM / MM_PER_FT, 3.0 * 2.0 * half_sizes(piece)[1])
    offset_ft = group["offset"]
    if abs(offset_ft) < min_ft:
        raise OverpassError("No change of elevation is needed (offset {:.0f} mm)".format(offset_ft * MM_PER_FT))
    level_id = get_level_id(piece)
    if level_id is None:
        raise OverpassError("No reference level on the element")
    inset_ft = leg_inset_ft(offset_ft, angle_deg)
    step = max(half_sizes(piece)[1], 25.0 / MM_PER_FT)
    extra = 0.0
    grown = 0.0
    for _ in range(MAX_CLEARANCE_ATTEMPTS + 2):  # due passi in piu' per gli allungamenti
        curve = piece.Location.Curve
        start, d = curve.GetEndPoint(0), curve.Direction
        t1, t2 = group["lo"] - extra - inset_ft, group["hi"] + extra + inset_ft
        why = " to keep the clearance near the elbows" if extra > 0 else ""
        if t1 < tail_ft:
            raise OverpassError("Not enough run before the obstacle{}: {:.0f} mm more are needed".format(
                why, (tail_ft - t1) * MM_PER_FT))
        if curve.Length - t2 < tail_ft:
            missing = tail_ft - (curve.Length - t2) + 1.0 / MM_PER_FT
            if grow_free_end(piece, 1, missing):
                grown += missing
                continue
            raise OverpassError("Not enough run after the obstacle{}: {:.0f} mm more are needed, and the free end "
                                "is connected".format(why, (missing - 1.0 / MM_PER_FT) * MM_PER_FT))
        transaction = DB.Transaction(doc, "Overpass")
        transaction.Start()
        try:
            _, new_elements, _, _, fittings = build_overpass(piece, curve, start.Add(d.Multiply(t1)),
                                                             start.Add(d.Multiply(t2)), offset_ft, inset_ft,
                                                             inset_ft == 0.0, level_id)
            STAGE[0] = "checking the clearance from the obstacle"
            # salite, tratto rialzato e gomiti: il primo e l'ultimo tratto restano alla quota originale
            violated, verified = clearance_check(new_elements[:3] + fittings, piece, obstacles, clearance_ft)
        except Exception:
            transaction.RollBack()
            raise
        if not violated:
            transaction.Commit()
            return new_elements, fittings, extra * MM_PER_FT, verified, grown * MM_PER_FT
        transaction.RollBack()
        extra += step
    raise OverpassError("The clearance of {:.0f} mm is not reached near the elbows, even starting the legs "
                        "{:.0f} mm farther from the obstacle".format(clearance_ft * MM_PER_FT,
                                                                    (extra - step) * MM_PER_FT))


def make_row(run_index, piece, obstacle, action, result):
    return {"run": run_index, "piece": piece.Id, "obstacle": obstacle.Id if obstacle is not None else None,
            "obstacle_cat": obstacle.Category.Name if obstacle is not None and obstacle.Category else "",
            "action": action, "result": result}


def give_up(clashes, message, action, ignored, rows):
    """Interferenze non risolte: nel report e non piu' riproposte."""
    for clash in clashes:
        ignored.add((clash["run"]["index"], id_value(clash["obstacle"].Id)))
        rows.append(make_row(clash["run"]["index"], clash["piece"], clash["obstacle"], action,
                             "Unchanged. {}".format(message)))


def resolve(todo, answer, ignored):
    """Scavalca sopra o sotto le interferenze spuntate. Ritorna le righe del report."""
    rows = []
    action = "Over" if answer["action"] == "over" else "Under"
    clearance_ft = answer["clearance_mm"] / MM_PER_FT
    by_piece = {}
    for clash in todo:
        by_piece.setdefault(id_value(clash["piece"].Id), []).append(clash)
    for clashes in by_piece.values():
        piece, run = clashes[0]["piece"], clashes[0]["run"]
        if not is_horizontal_line(piece):
            give_up(clashes, "Only straight horizontal runs can pass over or under an obstacle", action, ignored, rows)
            continue
        try:
            groups = plan_groups(piece, clashes, answer["action"], clearance_ft, answer["angle_deg"],
                                 answer["merge_mm"] / MM_PER_FT)
        except Exception as ex:
            give_up(clashes, "Cannot measure the obstacle: {}".format(ex), action, ignored, rows)
            continue
        # dal gruppo piu' lontano: il tratto originale resta quello iniziale, gli altri gruppi restano su di lui
        for group in sorted(groups, key=lambda g: -g["lo"]):
            try:
                new_elements, fittings, extra_mm, verified, grown_mm = apply_group(
                    piece, group, answer["angle_deg"], clearance_ft)
            except OverpassError as ex:
                give_up(group["clashes"], str(ex), action, ignored, rows)
                continue
            except Exception as ex:
                give_up(group["clashes"], "The overpass could not be built (step: {}): {}".format(STAGE[0], ex),
                        action, ignored, rows)
                continue
            for element in list(new_elements) + list(fittings):
                run["chain"][id_value(element.Id)] = element.Id
            note = "Created: offset {:+.0f} mm, {} new elements".format(group["offset"] * MM_PER_FT, len(new_elements))
            if len(group["clashes"]) > 1:
                note += ", one overpass for {} obstacles".format(len(group["clashes"]))
            if grown_mm > 0:
                note += ", run extended {:.0f} mm beyond the obstacle to make room for the overpass".format(grown_mm)
            if extra_mm > 0:
                note += ", legs moved {:.0f} mm farther from the obstacle to keep the clearance".format(extra_mm)
            note += "" if verified else ", clearance near the elbows not fully verified"
            for clash in group["clashes"]:
                rows.append(make_row(run["index"], piece, clash["obstacle"], action, note))
    return rows


def extend_run(run, click):
    """Allunga l'estremo libero orizzontale piu' vicino al punto, fino alla proiezione del punto. Ritorna (tratto, mm)."""
    best = None
    for piece in run_pieces(run):
        if not is_horizontal_line(piece):
            continue
        curve = piece.Location.Curve
        for idx in (0, 1):
            end = curve.GetEndPoint(idx)
            if nearest_connector(piece, end).IsConnected:
                continue
            dist = DB.XYZ(click.X - end.X, click.Y - end.Y, 0).GetLength()
            if best is None or dist < best[0]:
                best = (dist, piece, idx, end, curve.GetEndPoint(1 - idx))
    if best is None:
        raise OverpassError("No free horizontal end to extend")
    _, piece, idx, end, other = best
    d = end.Subtract(other).Normalize()
    t = DB.XYZ(click.X - end.X, click.Y - end.Y, 0).DotProduct(d)
    if t < MIN_REMAINING_MM / MM_PER_FT:
        raise OverpassError("The point is behind the free end of the run, or too close to it")
    new_end = end.Add(d.Multiply(t))
    transaction = DB.Transaction(doc, "Overpass extend")
    transaction.Start()
    try:
        piece.Location.Curve = DB.Line.CreateBound(other, new_end) if idx == 1 else \
            DB.Line.CreateBound(new_end, other)
        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise
    return piece, t * MM_PER_FT


def extend_runs(runs, click):
    rows = []
    for run in runs:
        try:
            piece, length_mm = extend_run(run, click)
            rows.append(make_row(run["index"], piece, None, "Extend", "Extended by {:.0f} mm".format(length_mm)))
        except Exception as ex:
            rows.append(make_row(run["index"], run["source"], None, "Extend", "Unchanged. {}".format(ex)))
    return rows


def get_route_selection():
    selected = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    runs = [e for e in selected if isinstance(e, ROUTE_CLASSES)]
    if runs:
        return runs
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, RouteFilter(),
                                           "Select the pipes, ducts or cable trays to route, then press Finish")
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def load_route_settings(config):
    settings = dict(ROUTE_DEFAULTS)
    try:
        settings["action"] = config.get_option("route_action", ROUTE_DEFAULTS["action"])
        settings["clearance_mm"] = float(config.get_option("route_clearance_mm", ROUTE_DEFAULTS["clearance_mm"]))
        settings["angle_deg"] = float(config.get_option("route_angle_deg", ROUTE_DEFAULTS["angle_deg"]))
        settings["merge_mm"] = float(config.get_option("route_merge_mm", ROUTE_DEFAULTS["merge_mm"]))
        scope = config.get_option("route_scope", ROUTE_DEFAULTS["scope"])
        settings["scope"] = scope if scope in SCOPE_LABELS else ROUTE_DEFAULTS["scope"]
    except Exception:
        pass
    return settings


def save_route_settings(config, settings):
    # le categorie accettate non si salvano: valgono solo per l'esecuzione in corso
    try:
        config.route_action = settings["action"]
        config.route_clearance_mm = settings["clearance_mm"]
        config.route_angle_deg = settings["angle_deg"]
        config.route_merge_mm = settings["merge_mm"]
        config.route_scope = settings["scope"]
        script.save_config()
    except Exception:
        pass


ACTION_LABELS = {"over": "pass over", "under": "pass under", "none": "report only"}
MAX_AUTO_ROUNDS = 10  # ricerche consecutive senza clic: uno scavalco puo' creare nuove interferenze, non all'infinito


def print_route_report(runs, rows, settings, clashes_found):
    output.print_md("# Clash Overpass report")
    if clashes_found:
        output.print_md("**⚠ Clashes found: {}** (with {})".format(clashes_found, SCOPE_LABELS[settings["scope"]]))
    else:
        output.print_md("No clash found with {}.".format(SCOPE_LABELS[settings["scope"]]))
    output.print_md("Runs: **{}** | action: **{}** | clearance **{:g} mm** | leg angle **{:g} degrees** from the "
                    "horizontal | overpasses merged when closer than **{:g} mm**".format(
                        len(runs), ACTION_LABELS.get(settings["action"], settings["action"]),
                        settings["clearance_mm"], settings["angle_deg"], settings["merge_mm"]))
    if not rows:
        output.print_md("No extension was made: nothing was changed.")
        return
    created = len([r for r in rows if r["result"].startswith("Created")])
    extended = len([r for r in rows if r["result"].startswith("Extended")])
    unchanged = len([r for r in rows if r["result"].startswith("Unchanged")])
    listed = len([r for r in rows if r["result"].startswith("Left")])
    output.print_md("Clashes solved: **{}** | only listed: **{}** | extensions: **{}** | not done: **{}**".format(
        created, listed, extended, unchanged))
    table = [[r["run"], output.linkify(r["piece"]), output.linkify(r["obstacle"]) if r["obstacle"] else "",
              r["obstacle_cat"], r["action"], r["result"]] for r in rows]
    output.print_table(table_data=table,
                       columns=["Run #", "Element", "Obstacle", "Obstacle category", "Action", "Result"])


def route_main():
    config = script.get_config(CONFIG_NAME)
    settings = load_route_settings(config)
    selected = get_route_selection()
    if not selected:
        script.exit()
    runs = [{"index": i + 1, "source": e, "chain": {id_value(e.Id): e.Id}} for i, e in enumerate(selected)]
    xaml_path = script.get_bundle_file(CLASH_XAML_FILE_NAME)
    ignored, rows = set(), []

    # finestra solo all'avvio: impostazioni e avviso con il numero di interferenze; cambiando il tipo di controllo
    # la ricerca si ripete e la finestra si riapre
    while True:
        clashes = find_clashes(runs, ignored, settings["scope"])
        answer = Overpass_ui.ClashWindow(xaml_path, len(clashes), settings).show()
        if answer is None:
            script.exit()  # Stop all'avvio: nulla e' stato modificato
        if answer.get("rescan"):
            settings["scope"] = answer["scope"]
            save_route_settings(config, settings)
            continue
        settings.update({k: answer[k] for k in ROUTE_DEFAULTS})
        save_route_settings(config, settings)
        break

    # poi tutto in automatico con quelle impostazioni: interferenze risolte (o solo elencate), clic per estendere,
    # nuova ricerca dopo ogni estensione
    clashes_found = 0
    rounds = 0
    while True:
        if clashes and rounds < MAX_AUTO_ROUNDS:
            rounds += 1
            clashes_found += len(clashes)
            if settings["action"] == "none":
                for clash in clashes:
                    ignored.add((clash["run"]["index"], id_value(clash["obstacle"].Id)))
                    rows.append(make_row(clash["run"]["index"], clash["piece"], clash["obstacle"],
                                         "Report only", "Left as it is"))
            else:
                rows += resolve(clashes, settings, ignored)
            clashes = find_clashes(runs, ignored, settings["scope"])
            continue
        if clashes:
            # troppe ricerche di fila: gli scavalchi continuano a creare interferenze, si lasciano come sono
            clashes_found += len(clashes)
            give_up(clashes, "Still clashing after {} automatic rounds: solve it by hand".format(MAX_AUTO_ROUNDS),
                    ACTION_LABELS[settings["action"]].capitalize(), ignored, rows)
            clashes = []
        try:
            click = uidoc.Selection.PickPoint("Clash Overpass ({} clashes found so far): click the point to extend "
                                              "the runs to (ESC to finish)".format(clashes_found))
        except OperationCanceledException:
            break
        except Exception as ex:
            output.print_md("**Cannot pick a point in this view** (set a work plane or use a plan view): {}".format(ex))
            break
        rows += extend_runs(runs, click)
        rounds = 0
        clashes = find_clashes(runs, ignored, settings["scope"])
    problems = len([r for r in rows if r["result"].startswith("Unchanged")])
    # solo report: il report e' lo scopo, si apre sempre se ci sono interferenze
    if (settings["action"] == "none" and clashes_found) or \
            want_report(len(rows), problems, "actions done, {} clashes found".format(clashes_found)):
        print_route_report(runs, rows, settings, clashes_found)


def overpass_main():
    config = script.get_config(CONFIG_NAME)
    default_offset, default_angle = 500.0, 90.0
    try:
        default_offset = float(config.get_option("offset_mm", 500.0))
        default_angle = float(config.get_option("angle_deg", 90.0))
    except Exception:
        pass

    xaml_path = script.get_bundle_file(XAML_FILE_NAME)
    settings = Overpass_ui.OverpassWindow(xaml_path, default_offset, default_angle).show()
    if settings is None:
        script.exit()
    offset_mm, angle_deg = settings["offset_mm"], settings["angle_deg"]

    try:
        config.offset_mm = offset_mm
        config.angle_deg = angle_deg
        script.save_config()
    except Exception:
        pass

    results = []
    first_filter = MepCurveFilter()
    while True:
        try:
            ref1 = uidoc.Selection.PickObject(
                ObjectType.Element, first_filter,
                "Click the FIRST point on a straight run (ESC to finish)")
        except OperationCanceledException:
            break
        element = doc.GetElement(ref1.ElementId)
        try:
            ref2 = uidoc.Selection.PickObject(
                ObjectType.Element, SameElementFilter(element.Id),
                "Click the SECOND point on the same run (ESC to skip this run)")
        except OperationCanceledException:
            results.append({"element": element, "category": category_name(element),
                            "status": "Skipped: second point not picked", "ok": False, "top": None,
                            "new_count": 0, "elbows": 0, "z_mm": None, "level": "", "rel_mm": None})
            continue
        results.append(process_run(element, ref1, ref2, offset_mm, angle_deg))

    problems = len([r for r in results if not r["ok"]])
    if want_report(len(results), problems, "overpasses created"):
        print_report(results, offset_mm, angle_deg)


route_main()
