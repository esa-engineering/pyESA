# -*- coding: utf-8 -*-
__title__ = "Overpass"
__doc__ = """Version = 0.1
Date    = 07.10.2026
________________________________________________________________
Description:
Makes a straight horizontal run of pipe, duct, cable tray or conduit
go up (or down) between two points, as if it had to pass over an
obstacle.

Instruction:
- Set the offset (mm): positive = up, negative = down
- Set the leg angle from the horizontal (90 = vertical)
- Click the first point on the run, then the second point on the
  same run
- The part between the two points is moved by the offset and joined
  to the rest of the run with two legs and four elbows; inclined legs
  start at the clicked points and go inwards
- If an elbow does not fit, the run is left unchanged
- Press ESC (at the first click) to finish and get the report
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

from pyrevit import revit, script, DB

import Overpass_ui

XAML_FILE_NAME = "Overpass_form.xaml"
CONFIG_NAME = "ESA_Overpass"

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

MM_PER_FT = 304.8
STRAIGHT_TOL = 0.999
HORIZONTAL_TOL = 1e-4
MIN_REMAINING_MM = 20.0  # lunghezza minima di ogni tratto dopo l'inserimento dei gomiti

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
    return top, [leg1, top, leg2, tail], len([f for f in (f1, f2, f3, f4) if f is not None]), inherited


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
        top, new_elements, elbows, inherited = build_overpass(element, curve, p1, p2, offset_ft, inset_ft,
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


def main():
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

    print_report(results, offset_mm, angle_deg)


main()
