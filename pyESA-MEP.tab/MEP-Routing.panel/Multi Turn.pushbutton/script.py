# -*- coding: utf-8 -*-
__title__ = "Bone"
__doc__ = """Version = 0.7
Date    = 12.10.2026
________________________________________________________________
Bends the selected parallel horizontal pipes, ducts or cable trays to the left or
right, keeping their spacing.
CLICK opens the settings window first, before any selection or click:
- angle (common angles or a custom one), or a direction of the view (arrows up,
  down, left, right: the angle is computed from the direction of the row)
- length after the bend (all ends aligned)
- align to an element: the row bends to run parallel to a clicked wall, beam,
  column, pipe, duct or vertical face, also in linked RVT and IFC models
Then select the elements and click where the bend starts (on one of them, or any
point, also on a linked model: the closest element is stretched to that point and
the others bend in line with it). With an angle, click to the left or right of the
row on the side to turn to; with the alignment, click the element to follow (if it
is perpendicular to the row, click it on the side to turn to).
SHIFT + CLICK skips the window and uses the last settings.
The part of each element beyond the start point is replaced. The new parts keep the
parameters and the insulation of the starting element. Cable trays keep the
orientation of their section (a tray on edge stays on edge after the bend).
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
from Autodesk.Revit.DB.Electrical import CableTray
from Autodesk.Revit.UI.Selection import ObjectType, ObjectSnapTypes, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, DB

import Comb_ui
import Turn_ui

XAML_FILE_NAME = "Comb_form.xaml"
CONFIG_NAME = "ESA_Comb"
BEND_DEFAULTS = {"angle_deg": 90.0, "length_mm": 1500.0, "mode": "angle", "direction": "up"}
ANGLE_XAML_FILE_NAME = "Angle_form.xaml"
QUICK_ANGLES = (10.0, 15.0, 22.5, 30.0, 45.0, 60.0, 90.0)
MIN_ANGLE, MAX_ANGLE = 5.0, 90.0  # limiti dell'angolo di curva, anche quando si calcola da freccia o elemento
DEFAULTS = {"mode": "below", "side": "Up", "factor": 3.0, "standoff_mm": 100.0, "length_mm": 1500.0}

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
HORIZONTAL_TOL = 1e-3
PARALLEL_TOL = 1e-3
SAME_Z_TOL_FT = 0.03     # tolleranza sulla quota comune dei tubi (circa 9 mm)
HEADING_TOL = 0.3        # componente minima del tubo verso la parete
CLICK_TOL_FT = 0.05      # il click sulla parete deve stare a lato della fila
MIN_REMAINING_MM = 20.0  # lunghezza minima di ogni tratto dopo l'inserimento dei gomiti
STAGE = [""]             # passo in corso, per i messaggi di errore
PHYSICAL_TYPES = (DB.ConnectorType.End, DB.ConnectorType.Curve, DB.ConnectorType.Physical)


class CombError(Exception):
    """Errore atteso: il messaggio va nel report."""


class PipeFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, Pipe)

    def AllowReference(self, reference, point):
        return False


class CurveFilter(ISelectionFilter):
    """Tubi, canali e passerelle (solo per la curva in orizzontale)."""

    def AllowElement(self, element):
        return isinstance(element, (Pipe, Duct, CableTray))

    def AllowReference(self, reference, point):
        return False


class WallFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, DB.Wall)

    def AllowReference(self, reference, point):
        return False


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


def outer_diameter(pipe):
    for bip in (DB.BuiltInParameter.RBS_PIPE_OUTER_DIAMETER, DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM):
        param = pipe.get_Parameter(bip)
        if param is not None and param.HasValue and param.AsDouble() > 0:
            return param.AsDouble()
    raise CombError("Cannot read the diameter of pipe {}".format(pipe.Id))


def nearest_connector(element, point):
    best, best_dist = None, None
    for c in element.ConnectorManager.Connectors:
        dist = c.Origin.DistanceTo(point)
        if best is None or dist < best_dist:
            best, best_dist = c, dist
    return best


def physical_refs(conn, element):
    """Connettori fisici di altri elementi collegati a conn (esclude il connettore logico del sistema)."""
    return [r for r in conn.AllRefs if r.Owner.Id != element.Id and r.ConnectorType in PHYSICAL_TYPES]


def find_fitting(conn, element):
    for ref in conn.AllRefs:
        owner = ref.Owner
        if owner.Id != element.Id and isinstance(owner, DB.FamilyInstance):
            return owner
    return None


SIZE_PARAMS = (DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM, DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM,
               DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM, DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM,
               DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)


def copy_size(source, target):
    """Sezione di canali e passerelle: il nuovo tratto deve avere le stesse dimensioni del tratto di partenza."""
    for bip in SIZE_PARAMS:
        src, dest = source.get_Parameter(bip), target.get_Parameter(bip)
        if src is None or dest is None or dest.IsReadOnly or not src.HasValue:
            continue
        if abs(dest.AsDouble() - src.AsDouble()) > 1e-9:
            dest.Set(src.AsDouble())


def tray_turned_copy(source, conn, end_point):
    """Passerella nuova con lo stesso orientamento della sezione (es. di costa): CableTray.Create la farebbe sempre
    piatta e Revit non ruota una passerella orizzontale attorno al proprio asse. Si copia il tratto di partenza, lo si
    mette in prosecuzione dall'estremo e lo si ruota attorno alla verticale passante per l'estremo fino alla
    direzione voluta: la rotazione attorno a Z conserva l'orientamento della sezione."""
    origin = conn.Origin
    travel = conn.CoordinateSystem.BasisZ  # verso uscente dall'estremo
    wanted = end_point.Subtract(origin)
    copy_id = list(DB.ElementTransformUtils.CopyElement(doc, source.Id, DB.XYZ.Zero))[0]
    new = doc.GetElement(copy_id)
    new.Location.Curve = DB.Line.CreateBound(origin, origin.Add(travel.Multiply(wanted.GetLength())))
    doc.Regenerate()
    flat_travel, flat_wanted = DB.XYZ(travel.X, travel.Y, 0), DB.XYZ(wanted.X, wanted.Y, 0)
    angle = math.atan2(flat_travel.CrossProduct(flat_wanted).Z, flat_travel.DotProduct(flat_wanted))
    if abs(angle) > 1e-6:
        axis = DB.Line.CreateBound(origin, origin.Add(DB.XYZ.BasisZ))
        DB.ElementTransformUtils.RotateElement(doc, copy_id, axis, angle)
    mark = new.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
    if mark is not None and not mark.IsReadOnly:
        mark.Set("")  # la copia eredita il Mark: duplicato
    doc.Regenerate()
    return new


def create_curve(source, conn, end_point, type_id, level_id):
    """Crea un tratto dello stesso tipo di source (tubo, canale o passerella) da conn a end_point."""
    if isinstance(source, CableTray):
        return tray_turned_copy(source, conn, end_point)
    elif isinstance(source, Duct):
        new = Duct.Create(doc, type_id, level_id, conn, end_point)
    else:
        return Pipe.Create(doc, type_id, level_id, conn, end_point)
    doc.Regenerate()
    copy_size(source, new)
    return new


def extend(owner, conn, end_point, type_id, level_id):
    """Aggiunge un tubo da conn a end_point con il gomito. Ritorna (tubo, connettore libero all'estremo, gomito)."""
    direction = end_point.Subtract(conn.Origin).Normalize()
    dot = conn.CoordinateSystem.BasisZ.DotProduct(direction)
    if dot < -STRAIGHT_TOL:
        raise CombError("A new run would fold back onto the previous pipe")
    new_pipe = create_curve(owner, conn, end_point, type_id, level_id)
    new_conn = nearest_connector(new_pipe, conn.Origin)
    fitting = None
    if dot > STRAIGHT_TOL:
        if not new_conn.IsConnectedTo(conn):
            new_conn.ConnectTo(conn)
    else:
        fitting = find_fitting(conn, owner)
        if fitting is None:
            # collegamento diretto senza raccordo: si scollega e si posiziona il gomito esplicitamente
            if new_conn.IsConnectedTo(conn):
                conn.DisconnectFrom(new_conn)
            fitting = doc.Create.NewElbowFitting(conn, new_conn)
    return new_pipe, nearest_connector(new_pipe, end_point), fitting


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
    """parts: [(nome, tubo)]. Dopo i gomiti ogni tratto deve avere ancora una lunghezza utile."""
    doc.Regenerate()
    for name, element in parts:
        remaining = element.Location.Curve.Length * MM_PER_FT
        if remaining < MIN_REMAINING_MM:
            raise CombError("Not enough space for the elbows: the {} would be only {:.0f} mm long "
                            "(minimum {:.0f} mm)".format(name, remaining, MIN_REMAINING_MM))


def pick_pipes(prompt, selection_filter=None):
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, selection_filter or PipeFilter(), prompt)
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def pick_one(selection_filter, prompt):
    try:
        return uidoc.Selection.PickObject(ObjectType.Element, selection_filter, prompt)
    except OperationCanceledException:
        script.exit()


def ends_of(pipe):
    curve = pipe.Location.Curve
    return curve.GetEndPoint(0), curve.GetEndPoint(1)


def check_straight_parallel(pipes):
    """Tubi rettilinei, orizzontali e paralleli. Ritorna la direzione del primo."""
    first_dir = None
    for pipe in pipes:
        curve = pipe.Location.Curve
        if not isinstance(curve, DB.Line):
            raise CombError("Element {} is not straight".format(pipe.Id))
        if abs(curve.Direction.Z) > HORIZONTAL_TOL:
            raise CombError("Element {} is not horizontal".format(pipe.Id))
        if first_dir is None:
            first_dir = curve.Direction
        elif first_dir.CrossProduct(curve.Direction).GetLength() > PARALLEL_TOL:
            raise CombError("Element {} is not parallel to the others".format(pipe.Id))
    return first_dir


def analyse(pipes, wall, wall_click):
    """Geometria di partenza, senza toccare il modello. Ritorna (piani dei tubi in ordine di corsa, dati comuni)."""
    check_straight_parallel(pipes)
    zs = [ends_of(p)[0].Z for p in pipes]
    if max(zs) - min(zs) > SAME_Z_TOL_FT:
        raise CombError("The pipes of the row must be at the same elevation")

    wall_curve = wall.Location.Curve
    if not isinstance(wall_curve, DB.Line):
        raise CombError("Only straight walls are supported")
    tau = DB.XYZ(wall_curve.Direction.X, wall_curve.Direction.Y, 0).Normalize()
    normal = DB.XYZ(-tau.Y, tau.X, 0)
    pw = wall_curve.GetEndPoint(0)
    side = sum(pipe.Location.Curve.Evaluate(0.5, True).Subtract(pw).DotProduct(normal) for pipe in pipes)
    n_away = normal.Multiply(1 if side >= 0 else -1)  # normale della parete verso i tubi

    items = []
    for pipe in pipes:
        a, b = ends_of(pipe)
        dist_a, dist_b = a.Subtract(pw).DotProduct(n_away), b.Subtract(pw).DotProduct(n_away)
        far, near = (a, b) if dist_a > dist_b else (b, a)
        direction = near.Subtract(far).Normalize()
        heading = direction.DotProduct(n_away)
        if heading > -HEADING_TOL:
            raise CombError("Pipe {} does not head to the wall".format(pipe.Id))
        t_hit = -far.Subtract(pw).DotProduct(n_away) / heading
        hit = far.Add(direction.Multiply(t_hit))
        items.append({"element": pipe, "far_end": far, "near_end": near, "direction": direction,
                      "heading": heading, "z": far.Z, "diameter": outer_diameter(pipe),
                      "x": hit.Subtract(pw).DotProduct(tau)})

    mean_x = sum(i["x"] for i in items) / len(items)
    along = wall_click.Subtract(pw).DotProduct(tau) - mean_x
    if abs(along) < CLICK_TOL_FT:
        raise CombError("Click the wall clearly to the left or to the right of the row of pipes")
    p_sign = 1 if along > 0 else -1
    items.sort(key=lambda i: p_sign * i["x"])  # #1 = corsa piu' lunga
    geo = {"tau": tau, "n_away": n_away, "pw": pw, "half": wall.Width / 2.0, "p_sign": p_sign,
           "p_dir": tau.Multiply(p_sign), "zc": sum(zs) / len(zs), "wall": wall}
    return items, geo


def assign_levels(plans, geo, settings):
    """Quota di ogni tubo della colonna. Chi ha la corsa piu' lunga prende la quota piu' lontana dalla fila."""
    n = len(plans)
    zc = geo["zc"]
    step = [settings["factor"] * p["diameter"] for p in plans]
    levels = [zc] * n
    if settings["mode"] == "reference":
        if settings["reference_index"] != n - 1:
            raise CombError("The pipe kept at its elevation must be the last one in the direction of the run "
                            "(#{}). Choose it, or click the wall on the other side of the row.".format(n))
        # il tubo tenuto e' l'ultimo: tutti gli altri si impilano dallo stesso lato, tutti su o tutti giu'
        side = 1.0 if settings["side"] == "Up" else -1.0
        for idx in range(n - 2, -1, -1):
            levels[idx] = levels[idx + 1] + side * step[idx + 1]
    else:
        sign = 1.0 if settings["mode"] == "above" else -1.0
        levels[n - 1] = zc + sign * step[n - 1]
        for idx in range(n - 2, -1, -1):
            levels[idx] = levels[idx + 1] + sign * step[idx + 1]
    for plan, level in zip(plans, levels):
        plan["level_z"] = level
        plan["rise"] = level - plan["z"]


def compute_corners(plans, geo, settings):
    """Punto in cui ogni tubo raggiunge la linea della colonna e fine delle corse lungo la parete."""
    target = geo["half"] + settings["standoff_mm"] / MM_PER_FT
    min_ft = MIN_REMAINING_MM / MM_PER_FT
    for plan in plans:
        gap = target - plan["far_end"].Subtract(geo["pw"]).DotProduct(geo["n_away"])
        if gap >= 0:
            raise CombError("Pipe {} starts nearer to the wall than the line of the stack".format(plan["element"].Id))
        t = gap / plan["heading"]
        if t < min_ft:
            raise CombError("Pipe {} is too short to reach the line of the stack".format(plan["element"].Id))
        corner = plan["far_end"].Add(plan["direction"].Multiply(t))
        plan["corner"] = corner
        plan["q"] = corner.Subtract(geo["pw"]).DotProduct(geo["p_dir"])
    q_end = max(p["q"] for p in plans) + settings["length_mm"] / MM_PER_FT
    for plan in plans:
        plan["run_length"] = q_end - plan["q"]


def build_pipe(plan, geo, level_id):
    """Costruisce salita e corsa di un tubo. Ritorna (nuovi tubi, tubi da controllare, raccordi creati)."""
    element = plan["element"]
    type_id = element.GetTypeId()
    corner = plan["corner"]

    STAGE[0] = "disconnecting the wall end of pipe {}".format(element.Id)
    near_conn = nearest_connector(element, plan["near_end"])
    for ref in physical_refs(near_conn, element):
        near_conn.DisconnectFrom(ref)

    STAGE[0] = "adjusting pipe {} to the line of the stack".format(element.Id)
    element.Location.Curve = DB.Line.CreateBound(plan["far_end"], corner)
    doc.Regenerate()
    conn = nearest_connector(element, corner)

    created, check, fittings = [], [("pipe of the row", element)], []
    owner = element
    top = corner
    if abs(plan["rise"]) > 1e-6:
        STAGE[0] = "vertical run of pipe {}".format(element.Id)
        top = corner.Add(DB.XYZ(0, 0, plan["rise"]))
        riser, conn, fitting = extend(owner, conn, top, type_id, level_id)
        created.append(riser)
        check.append(("vertical run", riser))
        fittings.append(fitting)
        owner = riser
    STAGE[0] = "run along the wall of pipe {}".format(element.Id)
    run, conn, fitting = extend(owner, conn, top.Add(geo["p_dir"].Multiply(plan["run_length"])), type_id, level_id)
    created.append(run)
    check.append(("run along the wall", run))
    fittings.append(fitting)
    return created, check, [f for f in fittings if f is not None]


def build_comb(plans, geo):
    results = []
    for i, plan in enumerate(plans):
        results.append({"element": plan["element"], "diameter": plan["diameter"], "z_from": plan["z"],
                        "z_to": None, "rise": plan["rise"], "rank": i, "new": [], "elbows": 0,
                        "status": "", "level": "", "rel_mm": None})
    transaction = DB.Transaction(doc, "Comb")
    transaction.Start()
    all_checks = []
    try:
        for plan, res in zip(plans, results):
            level_id = get_level_id(plan["element"])
            if level_id is None:
                raise CombError("No reference level on pipe {}".format(plan["element"].Id))
            level = doc.GetElement(level_id)
            created, check, fittings = build_pipe(plan, geo, level_id)
            all_checks += check
            plan["created"], plan["fittings"] = created, fittings
            res.update({"z_to": plan["level_z"], "new": created, "elbows": len(fittings), "level": level.Name,
                        "rel_mm": (plan["level_z"] - level_elevation(level)) * MM_PER_FT, "status": "Created"})
        STAGE[0] = "checking the remaining space"
        check_remaining_space(all_checks)
        STAGE[0] = "copying parameters and insulation"
        for plan, res in zip(plans, results):
            res["status"] = "Created ({})".format(inherit(plan["element"], plan["created"], plan["fittings"]))
        transaction.Commit()
    except Exception as ex:
        transaction.RollBack()
        detail = ""
        if isinstance(ex, CombError):
            message = "Unchanged. {}".format(ex)
        else:
            message = ("Unchanged. The comb could not be built (step: {}). If an elbow does not fit, "
                       "increase the spacing or the distance from the wall: {}".format(STAGE[0], ex))
            detail = traceback.format_exc() if traceback else ""
        for res in results:
            res.update({"z_to": None, "new": [], "elbows": 0, "status": message, "level": "", "rel_mm": None})
        return results, detail
    return results, ""


def fmt(value):
    return "" if value is None else "{:.0f}".format(value * MM_PER_FT)


def print_report(results, geo, settings, detail):
    output.print_md("# Comb report")
    how = {"below": "stack below the row", "above": "stack above the row",
           "reference": "pipe #{} kept at its elevation, the others {}".format(
               settings["reference_index"] + 1, "up" if settings["side"] == "Up" else "down")}[settings["mode"]]
    output.print_md("Wall: {} | **{}** | spacing **{:g} x diameter** | distance from the wall **{:g} mm** | "
                    "length along the wall **{:g} mm**".format(
                        output.linkify(geo["wall"].Id), how, settings["factor"],
                        settings["standoff_mm"], settings["length_mm"]))
    done = [r for r in results if r["status"].startswith("Created")]
    output.print_md("Pipes: **{}** | combed: **{}** | new pipes: **{}** | elbows: **{}**".format(
        len(results), len(done), sum(len(r["new"]) for r in done), sum(r["elbows"] for r in done)))
    rows = []
    for r in results:
        rows.append([r["rank"] + 1, output.linkify(r["element"].Id), fmt(r["diameter"]), fmt(r["z_from"]),
                     fmt(r["z_to"]), fmt(r["rise"]), r["level"],
                     "" if r["rel_mm"] is None else "{:.0f}".format(r["rel_mm"]),
                     ", ".join(output.linkify(p.Id) for p in r["new"]), r["status"]])
    output.print_table(table_data=rows,
                       columns=["# (longest run first)", "Pipe", "Diameter [mm]", "Z row [mm]", "Z stack [mm]",
                                "Rise [mm]", "Reference level", "Z stack vs level [mm]", "New pipes", "Status"])
    if detail:
        output.print_md("**Technical detail**")
        output.print_code(detail)


def load_settings(config):
    settings = dict(DEFAULTS)
    try:
        settings["mode"] = config.get_option("mode", DEFAULTS["mode"])
        settings["side"] = config.get_option("side", DEFAULTS["side"])
        settings["factor"] = float(config.get_option("factor", DEFAULTS["factor"]))
        settings["standoff_mm"] = float(config.get_option("standoff_mm", DEFAULTS["standoff_mm"]))
        settings["length_mm"] = float(config.get_option("length_mm", DEFAULTS["length_mm"]))
    except Exception:
        pass
    return settings


def save_settings(config, settings):
    try:
        for key in DEFAULTS:
            setattr(config, key, settings[key])
        script.save_config()
    except Exception:
        pass


def stop(message):
    output.print_md("# Multi Turn report")
    output.print_md("**Nothing was changed.** {}".format(message))
    script.exit()


# ----------------------------------------------------------------------------------------------------------------
# Direzione dopo la curva: angolo e clic sul lato, freccia nella vista, oppure parallela a un elemento cliccato
# ----------------------------------------------------------------------------------------------------------------
VERTICAL_FACE_TOL = 0.1  # sotto questa componente orizzontale la normale della faccia cliccata non da' una direzione
PERPENDICULAR_SIN = math.sin(math.radians(1.0))  # entro 1 grado dalla perpendicolare il verso lo decide il clic


def signed_angle_to(forward, target):
    """Angolo con segno in gradi da forward a target (orizzontali): positivo = sinistra visto dall'alto."""
    return math.degrees(math.atan2(forward.CrossProduct(target).Z, forward.DotProduct(target)))


def view_direction(name):
    """Direzione orizzontale della freccia scelta, come si vede nella vista attiva: destra = RightDirection della
    vista, su = 90 gradi a sinistra della destra (in pianta: verso l'alto dello schermo, anche con la vista ruotata)."""
    right = doc.ActiveView.RightDirection
    right = DB.XYZ(right.X, right.Y, 0)
    right = right.Normalize() if right.GetLength() > 1e-6 else DB.XYZ.BasisX
    up = DB.XYZ(-right.Y, right.X, 0)
    return {"right": right, "left": right.Negate(), "up": up, "down": up.Negate()}[name]


def hit_face_normal(reference, element, link_instance):
    """Normale (coordinate host) della faccia piana cliccata, o None se non e' leggibile. Per gli elementi di un link
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


def pick_alignment(prompt):
    """Clic su un elemento a cui allinearsi, anche in un link RVT o IFC. Ritorna (elemento, transform del link o None,
    nome del modello, punto cliccato, normale della faccia cliccata o None), o None con ESC."""
    try:
        reference = uidoc.Selection.PickObject(ObjectType.PointOnElement, prompt)
    except OperationCanceledException:
        return None
    host = doc.GetElement(reference.ElementId)
    hit = reference.GlobalPoint
    if isinstance(host, DB.RevitLinkInstance):
        link_doc = host.GetLinkDocument()
        if link_doc is None or reference.LinkedElementId == DB.ElementId.InvalidElementId:
            raise CombError("The element clicked in the linked model could not be read")
        element = link_doc.GetElement(reference.LinkedElementId)
        return element, host.GetTotalTransform(), link_doc.Title, hit, hit_face_normal(reference, element, host)
    return host, None, "host model", hit, hit_face_normal(reference, host, None)


def element_axis(element, transform, hit):
    """Asse orizzontale dell'elemento vicino al punto cliccato (muri, travi, tubi, canali, griglie, linee): ritorna
    la tangente orizzontale unitaria in coordinate host, o None se non c'e'."""
    location = getattr(element, "Location", None)
    if isinstance(location, DB.LocationCurve):
        curve = location.Curve
    elif isinstance(element, DB.Grid):
        curve = element.Curve
    elif isinstance(element, DB.CurveElement):
        curve = element.GeometryCurve
    else:
        return None
    if transform is not None:
        curve = curve.CreateTransformed(transform)
    projection = curve.Project(hit)
    if projection is None:
        return None
    direction = curve.ComputeDerivatives(projection.Parameter, False).BasisX
    flat = DB.XYZ(direction.X, direction.Y, 0)
    return flat.Normalize() if flat.GetLength() > 1e-9 else None


def align_target(pick, d, side_of_point):
    """Direzione dopo la curva parallela all'elemento cliccato, e nota per il report. Dei due versi lungo l'elemento
    si tiene quello che curva al massimo 90 gradi; se l'elemento e' perpendicolare ai tubi, il lato lo decide il punto
    cliccato (side_of_point: True = sinistra della fila, False = destra, None = sulla fila)."""
    element, transform, model, hit, face_normal = pick
    category = element.Category.Name if element.Category is not None else "Element"
    label = "{} {} ({})".format(category, element.Id, model)
    tangent = element_axis(element, transform, hit)
    if tangent is None:
        # senza asse (pilastro, pannello, DirectShape IFC): lungo la faccia cliccata
        if face_normal is None:
            raise CombError("{} has no axis and the clicked face could not be read: click a wall, beam, pipe or "
                            "duct, or a vertical face".format(label))
        flat = DB.XYZ(-face_normal.Y, face_normal.X, 0)
        if flat.GetLength() < VERTICAL_FACE_TOL:
            raise CombError("The clicked face of {} is horizontal (floor, soffit, top of an element): it gives no "
                            "direction. Click a vertical face".format(label))
        tangent = flat.Normalize()
    dot = d.DotProduct(tangent)
    if abs(dot) < PERPENDICULAR_SIN:
        side = side_of_point(hit)
        if side is None:
            raise CombError("{} is perpendicular to the pipes: click it to the left or to the right of the row, on "
                            "the side to turn to".format(label))
        left = DB.XYZ(-d.Y, d.X, 0)
        if (tangent.DotProduct(left) > 0) != side:
            tangent = tangent.Negate()
    elif dot < 0:
        tangent = tangent.Negate()
    return tangent, "Parallel to {}".format(label)


def check_angle(signed_deg):
    if abs(signed_deg) < MIN_ANGLE:
        raise CombError("The new direction is only {:.1f} degrees from the pipes: the minimum is {:g} degrees".format(
            abs(signed_deg), MIN_ANGLE))
    if abs(signed_deg) > MAX_ANGLE + 0.01:
        raise CombError("The new direction is {:.1f} degrees from the pipes: more than {:g} degrees needs two "
                        "bends".format(abs(signed_deg), MAX_ANGLE))


def click_side(pipes, left, click):
    """Lato della fila su cui sta il punto: True = sinistra, False = destra, None = sulla fila."""
    lats = [ends_of(p)[0].Subtract(click).DotProduct(left) for p in pipes]
    click_left = -sum(lats) / len(lats)
    if abs(click_left) < CLICK_TOL_FT:
        return None
    return click_left > 0


def resolve_turn(pipes, d, settings, side_click, pick):
    """Angolo (positivo), verso (True = sinistra) e nota per il report, secondo la modalita' della finestra."""
    left = DB.XYZ(-d.Y, d.X, 0)
    mode = settings["mode"]
    if mode == "angle":
        side = click_side(pipes, left, side_click)
        if side is None:
            raise CombError("Click clearly to the left or to the right of the row of pipes")
        return settings["angle_deg"], side, ""
    if mode == "direction":
        target, note = view_direction(settings["direction"]), Turn_ui.describe(settings).capitalize()
    else:
        target, note = align_target(pick, d, lambda point: click_side(pipes, left, point))
    signed_deg = signed_angle_to(d, target)
    check_angle(signed_deg)
    return abs(signed_deg), signed_deg > 0, note


def analyse_bend(pipes, start_click, settings, side_click=None, pick=None):
    """Geometria di partenza della curva, senza toccare il modello. Ritorna (piani dei tubi, dati comuni).
    start_click: punto qualsiasi (anche fuori dai tubi o su un link): il tubo piu' vicino fa da riferimento e la
    curva parte dalla proiezione del punto sulla sua linea (i tubi vengono prolungati fino li').
    side_click (modalita' angolo): punto a lato della fila che sceglie il verso della piega.
    pick (modalita' allineamento): elemento cliccato a cui la fila diventa parallela dopo la curva."""
    d0 = check_straight_parallel(pipes)
    d0 = DB.XYZ(d0.X, d0.Y, 0).Normalize()
    perp = DB.XYZ(-d0.Y, d0.X, 0)
    anchor = min(pipes, key=lambda p: abs(ends_of(p)[0].Subtract(start_click).DotProduct(perp)))
    a0, a1 = ends_of(anchor)
    along = start_click.Subtract(a0).DotProduct(d0)
    start_point = DB.XYZ(a0.X + d0.X * along, a0.Y + d0.Y * along, a0.Z)
    mid = a0.Add(a1).Multiply(0.5)
    # verso della curva: la parte del tubo oltre il punto di partenza viene sostituita
    d = d0 if start_point.Subtract(mid).DotProduct(d0) >= 0 else d0.Negate()
    angle_deg, turn_left, note = resolve_turn(pipes, d, settings, side_click, pick)
    left = DB.XYZ(-d.Y, d.X, 0)
    n = left if turn_left else left.Negate()
    base = min(ends_of(anchor), key=lambda e: e.DotProduct(d))
    items = []
    for pipe in pipes:
        back, forward = sorted(ends_of(pipe), key=lambda e: e.DotProduct(d))
        items.append({"element": pipe, "back": back, "forward": forward,
                      "w": back.Subtract(base).DotProduct(n),
                      "u0": start_point.Subtract(back).DotProduct(d)})
    geo = {"d": d, "n": n, "left": turn_left, "angle_deg": angle_deg, "note": note}
    return items, geo


def compute_bend(plans, geo, settings):
    """Angolo di curva di ogni tubo: gli angoli stanno sulla bisettrice, cosi' l'interasse si conserva."""
    theta = math.radians(geo["angle_deg"])
    d, n = geo["d"], geo["n"]
    d_new = d.Multiply(math.cos(theta)).Add(n.Multiply(math.sin(theta)))
    geo["d_new"] = d_new
    min_ft = MIN_REMAINING_MM / MM_PER_FT
    for plan in plans:
        u = plan["u0"] - plan["w"] * math.tan(theta / 2.0)
        if u < min_ft:
            raise CombError("The bend of pipe {} falls outside the pipe, too close to its start. "
                            "Click farther along the pipes".format(plan["element"].Id))
        plan["corner"] = plan["back"].Add(d.Multiply(u))
        plan["q"] = plan["corner"].DotProduct(d_new)
    q_end = max(p["q"] for p in plans) + settings["length_mm"] / MM_PER_FT
    for plan in plans:
        plan["leg"] = q_end - plan["q"]
        plan["leg_end"] = plan["corner"].Add(d_new.Multiply(plan["leg"]))


def build_bend_pipe(plan):
    """Taglia il tubo alla curva e aggiunge il tratto dopo il gomito. Ritorna (nuovo tubo, raccordo, scollegati)."""
    element = plan["element"]
    type_id = element.GetTypeId()
    level_id = get_level_id(element)
    if level_id is None:
        raise CombError("No reference level on pipe {}".format(element.Id))

    STAGE[0] = "disconnecting the end of pipe {}".format(element.Id)
    end_conn = nearest_connector(element, plan["forward"])
    refs = physical_refs(end_conn, element)
    for ref in refs:
        end_conn.DisconnectFrom(ref)

    STAGE[0] = "cutting pipe {} at the bend".format(element.Id)
    # si mantiene il verso di disegno originale (inizio/fine): invertirlo fa rifiutare il cambio a passerelle con raccordi
    if element.Location.Curve.GetEndPoint(0).DistanceTo(plan["back"]) <= element.Location.Curve.GetEndPoint(1).DistanceTo(plan["back"]):
        new_line = DB.Line.CreateBound(plan["back"], plan["corner"])
    else:
        new_line = DB.Line.CreateBound(plan["corner"], plan["back"])
    element.Location.Curve = new_line
    doc.Regenerate()
    conn = nearest_connector(element, plan["corner"])

    STAGE[0] = "new leg of pipe {}".format(element.Id)
    leg, conn, fitting = extend(element, conn, plan["leg_end"], type_id, level_id)
    return leg, fitting, len(refs)


def build_bend(plans, geo):
    results = [{"element": p["element"], "new": [], "elbows": 0, "freed": 0, "status": ""} for p in plans]
    transaction = DB.Transaction(doc, "Comb bend")
    transaction.Start()
    try:
        checks = []
        for plan, res in zip(plans, results):
            leg, fitting, freed = build_bend_pipe(plan)
            plan["created"], plan["fittings"] = [leg], [f for f in [fitting] if f is not None]
            checks += [("pipe before the bend", plan["element"]), ("new leg", leg)]
            res.update({"new": [leg], "elbows": len(plan["fittings"]), "freed": freed, "status": "Created"})
        STAGE[0] = "checking the remaining space"
        check_remaining_space(checks)
        STAGE[0] = "copying parameters and insulation"
        for plan, res in zip(plans, results):
            res["status"] = "Created ({})".format(inherit(plan["element"], plan["created"], plan["fittings"]))
        STAGE[0] = "final check of the new elements"
        doc.Regenerate()
        for plan in plans:
            if not all(e.IsValidObject for e in [plan["element"]] + plan["created"]):
                raise CombError("Revit removed the new pipe of element {} while rebuilding the bend. "
                                "Try a smaller angle, a longer length after the bend, or a start point farther "
                                "from the ends of the pipe".format(plan["element"].Id))
        transaction.Commit()
    except Exception as ex:
        transaction.RollBack()
        detail = ""
        if isinstance(ex, CombError):
            message = "Unchanged. {}".format(ex)
        else:
            message = ("Unchanged. The bend could not be built (step: {}). If an elbow does not fit, "
                       "increase the length after the bend or reduce the angle: {}".format(STAGE[0], ex))
            detail = traceback.format_exc() if traceback else ""
        for res in results:
            res.update({"new": [], "elbows": 0, "freed": 0, "status": message})
        return results, detail
    return results, ""


def print_bend_report(results, geo, settings, detail):
    output.print_md("# Multi Turn report")
    output.print_md("Bend to the **{}** | angle **{:.1f} deg** | length after the bend **{:g} mm**".format(
        "left" if geo["left"] else "right", geo["angle_deg"], settings["length_mm"]))
    if geo.get("note"):
        output.print_md("Direction: **{}**".format(geo["note"]))
    done = [r for r in results if r["status"].startswith("Created")]
    output.print_md("Pipes: **{}** | bent: **{}** | new pipes: **{}** | elbows: **{}**".format(
        len(results), len(done), sum(len(r["new"]) for r in done), sum(r["elbows"] for r in done)))
    freed = sum(r["freed"] for r in done)
    if freed:
        output.print_md("**Note:** {} connections at the far end of the pipes were removed, "
                        "because the part beyond the bend was replaced.".format(freed))
    rows = [[output.linkify(r["element"].Id) if r["element"].IsValidObject else "(removed)",
             ", ".join(output.linkify(p.Id) for p in r["new"] if p.IsValidObject), r["elbows"], r["status"]]
            for r in results]
    output.print_table(table_data=rows, columns=["Pipe", "New pipes", "Elbows", "Status"])
    if detail:
        output.print_md("**Technical detail**")
        output.print_code(detail)


def load_bend_settings(config):
    settings = dict(BEND_DEFAULTS)
    try:
        settings["angle_deg"] = float(config.get_option("angle_deg", BEND_DEFAULTS["angle_deg"]))
        # nella configurazione la lunghezza resta "leg_mm", come nelle versioni precedenti
        settings["length_mm"] = float(config.get_option("leg_mm", BEND_DEFAULTS["length_mm"]))
        mode = config.get_option("mode", BEND_DEFAULTS["mode"])
        settings["mode"] = mode if mode in ("angle", "direction", "align") else BEND_DEFAULTS["mode"]
        direction = config.get_option("direction", BEND_DEFAULTS["direction"])
        settings["direction"] = direction if direction in Turn_ui.DIRECTIONS else BEND_DEFAULTS["direction"]
    except Exception:
        pass
    return settings


def save_bend_settings(config, settings):
    try:
        config.angle_deg = settings["angle_deg"]
        config.leg_mm = settings["length_mm"]
        config.mode = settings["mode"]
        config.direction = settings["direction"]
        script.save_config()
    except Exception:
        pass


WINDOW_TEXTS = {
    "caption": "Multi Turn",
    "title": u"↪ MULTI TURN",
    "length_header": u"\U0001F4CF LENGTH",
    "length_unit": "mm after the farthest bend (all ends aligned)",
    "align_hint": "After selecting the elements and clicking where the bend starts, click the element to follow: a "
                  "wall, beam, column, pipe, duct or vertical face, also in linked RVT and IFC models. The row bends "
                  "to run parallel to it (at most 90 degrees). If the element is perpendicular to the row, click it "
                  "on the side to turn to.",
    "info": "Click an angle button or an arrow to start, or type a custom angle and press Use custom. Then select the "
            "elements, click where the bend starts and, with an angle, click on the side to turn to; with an arrow "
            "the row turns towards that direction of the view. The elements keep their spacing through the bend and "
            "the part beyond the bend is replaced. SHIFT + CLICK skips this window and uses the last settings.",
}


def bend_main():
    config = script.get_config(CONFIG_NAME + "_Bend")
    settings = load_bend_settings(config)
    # CLICK: la finestra si apre sempre, prima di selezione e clic; SHIFT + CLICK: ultime impostazioni, senza finestra
    if not __shiftclick__:  # noqa: F821
        settings = Turn_ui.TurnWindow(script.get_bundle_file(ANGLE_XAML_FILE_NAME), settings, WINDOW_TEXTS,
                                      QUICK_ANGLES, MIN_ANGLE, MAX_ANGLE, show_slope=False).show()
        if settings is None:
            script.exit()
        save_bend_settings(config, settings)

    steps = 3 if settings["mode"] in ("angle", "align") else 2
    pipes = pick_pipes("Step 1 of {}: select the pipes, ducts or cable trays to bend, then press Finish".format(steps),
                       CurveFilter())
    if not pipes:
        script.exit()

    snaps = (ObjectSnapTypes.Endpoints | ObjectSnapTypes.Midpoints | ObjectSnapTypes.Nearest |
             ObjectSnapTypes.Intersections | ObjectSnapTypes.Perpendicular | ObjectSnapTypes.Centers)
    try:
        start_click = uidoc.Selection.PickPoint(snaps, "Step 2 of {}: click where the bend starts (on a pipe, or "
                                                       "any point, also on a linked model: the closest pipe is "
                                                       "stretched to it)".format(steps))
    except OperationCanceledException:
        script.exit()

    side_click, pick = None, None
    try:
        if settings["mode"] == "angle":
            side_click = uidoc.Selection.PickPoint("Step 3 of 3: click to the left or right of the row, "
                                                   "on the side the pipes must turn to")
        elif settings["mode"] == "align":
            pick = pick_alignment("Step 3 of 3: click the element the pipes must be parallel to after the bend, "
                                  "also in a linked model")
            if pick is None:
                script.exit()
    except OperationCanceledException:
        script.exit()
    except CombError as ex:
        stop(ex)

    try:
        plans, geo = analyse_bend(pipes, start_click, settings, side_click, pick)
        compute_bend(plans, geo, settings)
    except CombError as ex:
        stop(ex)

    results, detail = build_bend(plans, geo)
    problems = len([r for r in results if not r["status"].startswith("Created")])
    if want_report(len(results), problems, "elements bent"):
        print_bend_report(results, geo, settings, detail)


def comb_main():
    pipes = pick_pipes("Step 1 of 2: select the pipes of the row, heading to the wall, then press Finish")
    if not pipes:
        script.exit()
    wall_ref = pick_one(WallFilter(), "Step 2 of 2: click the wall, to the left or right of the row "
                                      "(the side sets the direction of the runs)")
    wall = doc.GetElement(wall_ref.ElementId)
    wall_click = wall_ref.GlobalPoint or wall.Location.Curve.Evaluate(0.5, True)

    try:
        plans, geo = analyse(pipes, wall, wall_click)
    except CombError as ex:
        stop(ex)

    config = script.get_config(CONFIG_NAME)
    labels = ["#{} | {:.0f} mm{}".format(i + 1, p["diameter"] * MM_PER_FT,
                                         " (longest run)" if i == 0 else (" (shortest run)" if i == len(plans) - 1 else ""))
              for i, p in enumerate(plans)]
    settings = Comb_ui.CombWindow(script.get_bundle_file(XAML_FILE_NAME), labels, load_settings(config)).show()
    if settings is None:
        script.exit()
    save_settings(config, settings)

    try:
        assign_levels(plans, geo, settings)
        compute_corners(plans, geo, settings)
    except CombError as ex:
        stop(ex)

    results, detail = build_comb(plans, geo)
    problems = len([r for r in results if not r["status"].startswith("Created")])
    if want_report(len(results), problems, "pipes combed"):
        print_report(results, geo, settings, detail)


bend_main()
