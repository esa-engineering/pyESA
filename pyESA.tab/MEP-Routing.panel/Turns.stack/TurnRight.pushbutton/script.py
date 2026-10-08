# -*- coding: utf-8 -*-
__title__ = "Turn Right"
__doc__ = """Version = 0.1
Date    = 07.10.2026
________________________________________________________________
Description:
Adds a run that turns right from the free end of a pipe, duct,
cable tray or conduit, with an elbow.

Instruction:
- CLICK: uses the saved settings (default: 1000 mm, 90 degrees,
  keep the slope of the original run)
- SHIFT + CLICK: change length, turn angle and slope option
- Click each element near the free end you want to extend
- Press ESC to finish: a report lists the new elements

Right and left are relative to the direction of travel, looking
from the element towards the clicked end. Vertical ends are skipped.
The new parts keep the parameters and the insulation of the starting element.
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

from pyrevit import revit, script, forms, DB

import Turn_ui

# unica differenza tra TurnRight e TurnLeft: tenere i due script allineati
TURN_SIDE = "right"

XAML_FILE_NAME = "Turn_form.xaml"
CONFIG_NAME = "ESA_MEPTurn"  # condiviso tra Turn Right e Turn Left
DEFAULTS = {"length_mm": 1000.0, "angle_deg": 90.0, "keep_slope": True}

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


def turn_direction(travel, angle_deg, keep_slope):
    """Direzione del nuovo tratto. None se il tratto e' verticale (destra/sinistra non definite)."""
    if abs(travel.Z) > VERTICAL_TOL:
        return None
    horizontal_len = math.sqrt(travel.X ** 2 + travel.Y ** 2)
    forward = DB.XYZ(travel.X / horizontal_len, travel.Y / horizontal_len, 0)
    right = DB.XYZ(forward.Y, -forward.X, 0)  # forward x asse Z
    side = right if TURN_SIDE == "right" else right.Negate()
    angle = math.radians(angle_deg)
    new_dir = forward.Multiply(math.cos(angle)).Add(side.Multiply(math.sin(angle)))
    if keep_slope:
        # stessa pendenza: stessa salita per unita' di percorso orizzontale
        new_dir = DB.XYZ(new_dir.X, new_dir.Y, travel.Z / horizontal_len).Normalize()
    return new_dir


def create_segment(element, conn, end_point, level_id):
    """Crea il nuovo tratto (senza garantire il raccordo). Ritorna l'elemento creato."""
    type_id = element.GetTypeId()
    if isinstance(element, Pipe):
        return Pipe.Create(doc, type_id, level_id, conn, end_point)
    if isinstance(element, Duct):
        return Duct.Create(doc, type_id, level_id, conn, end_point)
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
              "new": None, "fitting": "", "added_mm": None, "z_mm": None, "level": "", "rel_mm": None}

    click = reference.GlobalPoint
    if click is None:
        click = element.Location.Curve.Evaluate(0.5, True)

    conn = nearest_free_connector(element, click)
    if conn is None:
        result["status"] = "No free end (both ends are connected)"
        return result

    direction = turn_direction(run_direction_at(element, conn),
                               settings["angle_deg"], settings["keep_slope"])
    if direction is None:
        result["status"] = "Vertical end: right and left are not defined"
        return result

    level_id = get_level_id(element)
    if level_id is None:
        result["status"] = "No reference level on the element"
        return result
    ref_level = doc.GetElement(level_id)
    result["level"] = ref_level.Name

    length = settings["length_mm"] / MM_PER_FT
    end_point = conn.Origin.Add(direction.Multiply(length))

    transaction = DB.Transaction(doc, "Turn {}".format(TURN_SIDE.capitalize()))
    transaction.Start()
    try:
        new_element = create_segment(element, conn, end_point, level_id)
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
    return "{:g} mm, {:g} degrees, {}".format(
        settings["length_mm"], settings["angle_deg"],
        "keep slope" if settings["keep_slope"] else "horizontal")


def print_report(results, settings):
    output.print_md("# Turn {} report".format(TURN_SIDE.capitalize()))
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
            i, output.linkify(r["element"].Id), r["category"], r["status"],
            fmt(r["added_mm"]),
            output.linkify(r["new"].Id) if r["new"] is not None else "", r["fitting"],
            fmt(r["z_mm"]), r["level"], fmt(r["rel_mm"]),
        ])
    output.print_table(
        table_data=rows,
        columns=["#", "Element", "Category", "Status", "Added length [mm]", "New element", "Fitting",
                 "End Z, internal coordinates [mm]", "Reference level", "End Z vs reference level [mm]"])


def load_settings(config):
    settings = dict(DEFAULTS)
    try:
        settings["length_mm"] = float(config.get_option("length_mm", DEFAULTS["length_mm"]))
        settings["angle_deg"] = float(config.get_option("angle_deg", DEFAULTS["angle_deg"]))
        keep = config.get_option("keep_slope", DEFAULTS["keep_slope"])
        if not isinstance(keep, bool):
            keep = str(keep).strip().lower() in ("true", "1", "yes")
        settings["keep_slope"] = keep
    except Exception:
        pass
    return settings


def save_settings(config, settings):
    try:
        config.length_mm = settings["length_mm"]
        config.angle_deg = settings["angle_deg"]
        config.keep_slope = settings["keep_slope"]
        script.save_config()
    except Exception:
        pass


def pick_loop(settings):
    results = []
    selection_filter = MepCurveFilter()
    prompt = "Turn {} ({}): click near the free end (ESC to finish)".format(TURN_SIDE, settings_text(settings))
    while True:
        try:
            reference = uidoc.Selection.PickObject(ObjectType.Element, selection_filter, prompt)
        except OperationCanceledException:
            break
        results.append(process_pick(reference, settings))
    return results


def main():
    config = script.get_config(CONFIG_NAME)
    settings = load_settings(config)

    if __shiftclick__:  # noqa: F821
        xaml_path = script.get_bundle_file(XAML_FILE_NAME)
        new_settings = Turn_ui.TurnWindow(xaml_path, settings).show()
        if new_settings is None:
            script.exit()
        settings = new_settings
        save_settings(config, settings)

    warnings = []
    if abs(settings["angle_deg"] - 90.0) > 1e-6:
        warnings.append("turn angle {:g} degrees".format(settings["angle_deg"]))
    if not settings["keep_slope"]:
        warnings.append("new runs forced horizontal")

    if warnings:
        message = "Turn {}: {} (SHIFT + CLICK to change)".format(TURN_SIDE, ", ".join(warnings))
        with forms.WarningBar(title=message):
            results = pick_loop(settings)
    else:
        results = pick_loop(settings)

    problems = len([r for r in results if not r["ok"]])
    if want_report(len(results), problems, "runs created"):
        print_report(results, settings)


main()
