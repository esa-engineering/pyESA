# -*- coding: utf-8 -*-
__title__ = "Comb"
__doc__ = """Version = 0.4
Date    = 08.10.2026
________________________________________________________________
CLICK: bends the selected parallel horizontal pipes, ducts or cable trays to the left or
right by a chosen angle, keeping their spacing. Select the pipes,
then click the point where the bend starts, on the side the pipes
must turn to. The part of each pipe beyond the bend is replaced.

SHIFT + CLICK (comb, pipes only):
Turns a row of horizontal pipes heading to a wall into a comb at the
wall: each pipe reaches the wall, goes up or down with an elbow to
its own elevation, then turns 90 degrees and runs along the wall.
The pipes end up stacked one above the other, with aligned ends.

Instruction:
1. Select the pipes of the row (horizontal, same elevation, side by
   side, heading to the wall) and press Finish
2. Click the wall, to the left or right of the row: the side sets the
   direction of the runs along the wall
3. Choose the elevations of the stack: below the row, above the row,
   or keeping the shortest-run pipe at its elevation, with all the
   others going up or all down

The pipe with the longest run takes the farthest elevation, so the
risers never cross the runs. If anything does not fit, nothing is
changed. Pipes only, straight, with a straight wall.
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
from Autodesk.Revit.DB.Electrical import CableTray
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, DB

import Comb_ui

XAML_FILE_NAME = "Comb_form.xaml"
BEND_XAML_FILE_NAME = "Bend_form.xaml"
CONFIG_NAME = "ESA_Comb"
BEND_DEFAULTS = {"angle_deg": 90.0, "leg_mm": 1500.0}
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


def create_curve(source, conn, end_point, type_id, level_id):
    """Crea un tratto dello stesso tipo di source (tubo, canale o passerella) da conn a end_point."""
    if isinstance(source, CableTray):
        new = CableTray.Create(doc, type_id, conn.Origin, end_point, level_id)
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
    output.print_md("# Comb report")
    output.print_md("**Nothing was changed.** {}".format(message))
    script.exit()


def analyse_bend(pipes, click):
    """Geometria di partenza della curva, senza toccare il modello. Ritorna (piani dei tubi, dati comuni)."""
    d0 = check_straight_parallel(pipes)
    d0 = DB.XYZ(d0.X, d0.Y, 0).Normalize()
    anchor = min(pipes, key=lambda p: abs(DB.XYZ(-d0.Y, d0.X, 0).DotProduct(ends_of(p)[0].Subtract(click))))
    a0, a1 = ends_of(anchor)
    mid = a0.Add(a1).Multiply(0.5)
    # verso della curva: la parte del tubo oltre il click viene sostituita
    d = d0 if click.Subtract(mid).DotProduct(d0) >= 0 else d0.Negate()
    left = DB.XYZ(-d.Y, d.X, 0)
    lats = [ends_of(p)[0].Subtract(click).DotProduct(left) for p in pipes]
    click_left = -sum(lats) / len(lats)
    if abs(click_left) < CLICK_TOL_FT:
        raise CombError("Click clearly to the left or to the right of the row of pipes")
    n = left.Multiply(1 if click_left > 0 else -1)
    base = min(ends_of(anchor), key=lambda e: e.DotProduct(d))
    items = []
    for pipe in pipes:
        back, forward = sorted(ends_of(pipe), key=lambda e: e.DotProduct(d))
        items.append({"element": pipe, "back": back, "forward": forward,
                      "w": back.Subtract(base).DotProduct(n),
                      "u0": click.Subtract(back).DotProduct(d)})
    geo = {"d": d, "n": n, "left": click_left > 0}
    return items, geo


def compute_bend(plans, geo, settings):
    """Angolo di curva di ogni tubo: gli angoli stanno sulla bisettrice, cosi' l'interasse si conserva."""
    theta = math.radians(settings["angle_deg"])
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
    q_end = max(p["q"] for p in plans) + settings["leg_mm"] / MM_PER_FT
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
    element.Location.Curve = DB.Line.CreateBound(plan["back"], plan["corner"])
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
    output.print_md("# Comb bend report")
    output.print_md("Bend to the **{}** | angle **{:g} deg** | length after the bend **{:g} mm**".format(
        "left" if geo["left"] else "right", settings["angle_deg"], settings["leg_mm"]))
    done = [r for r in results if r["status"].startswith("Created")]
    output.print_md("Pipes: **{}** | bent: **{}** | new pipes: **{}** | elbows: **{}**".format(
        len(results), len(done), sum(len(r["new"]) for r in done), sum(r["elbows"] for r in done)))
    freed = sum(r["freed"] for r in done)
    if freed:
        output.print_md("**Note:** {} connections at the far end of the pipes were removed, "
                        "because the part beyond the bend was replaced.".format(freed))
    rows = [[output.linkify(r["element"].Id), ", ".join(output.linkify(p.Id) for p in r["new"]),
             r["elbows"], r["status"]] for r in results]
    output.print_table(table_data=rows, columns=["Pipe", "New pipes", "Elbows", "Status"])
    if detail:
        output.print_md("**Technical detail**")
        output.print_code(detail)


def load_bend_settings(config):
    settings = dict(BEND_DEFAULTS)
    try:
        for key in BEND_DEFAULTS:
            settings[key] = float(config.get_option(key, BEND_DEFAULTS[key]))
    except Exception:
        pass
    return settings


def save_bend_settings(config, settings):
    try:
        for key in BEND_DEFAULTS:
            setattr(config, key, settings[key])
        script.save_config()
    except Exception:
        pass


def bend_main():
    pipes = pick_pipes("Step 1 of 2: select the pipes, ducts or cable trays to bend, then press Finish", CurveFilter())
    if not pipes:
        script.exit()
    try:
        click = uidoc.Selection.PickPoint("Step 2 of 2: click where the bend starts, on the side the pipes "
                                          "must turn to (left or right of the row)")
    except OperationCanceledException:
        script.exit()
    try:
        plans, geo = analyse_bend(pipes, click)
    except CombError as ex:
        stop(ex)

    config = script.get_config(CONFIG_NAME + "_Bend")
    settings = Comb_ui.BendWindow(script.get_bundle_file(BEND_XAML_FILE_NAME),
                                  (len(plans), "left" if geo["left"] else "right"),
                                  load_bend_settings(config)).show()
    if settings is None:
        script.exit()
    save_bend_settings(config, settings)

    try:
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


if __shiftclick__:  # noqa: F821
    comb_main()
else:
    bend_main()
