# -*- coding: utf-8 -*-
__title__ = "Riser Up"
__doc__ = """Version = 0.2
Date    = 07.10.2026
________________________________________________________________
Description:
Adds a vertical run upwards to pipes, ducts, cable trays and conduits.

Instruction:
- CLICK: uses the last settings (default: 1000 mm)
- SHIFT + CLICK: set the height, either a length from the selected
  end or a level plus an offset
- Click each element near the free end you want to extend
  (the nearest free end is used)
- Press ESC to finish: a report lists the new elements and their
  final elevation
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
from Autodesk.Revit.DB.Mechanical import Duct, DuctInsulation, DuctLining, MechanicalSystemType
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, forms, DB

import Riser_ui

# unica differenza tra RiserUp e RiserDown: tenere i due script allineati
RISER_DIRECTION = "up"

SIGN = 1 if RISER_DIRECTION == "up" else -1
WORDS = {"up": {"side": "above", "towards": "upwards", "opposite": "downwards"},
         "down": {"side": "below", "towards": "downwards", "opposite": "upwards"}}[RISER_DIRECTION]

XAML_FILE_NAME = "Riser_form.xaml"
CONFIG_NAME = "ESA_Riser{}".format(RISER_DIRECTION.capitalize())
DEFAULTS = {"mode": "length", "length_mm": 1000.0, "offset_mm": 0.0, "level_name": ""}

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

MM_PER_FT = 304.8
MIN_LENGTH_FT = 0.01  # l'API rifiuta tratti piu' corti di 1/10 pollice
STRAIGHT_TOL = 0.999
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


def category_name(element):
    for cls, name in CATEGORY_NAMES:
        if isinstance(element, cls):
            return name
    return "?"


def level_elevation(level):
    # ProjectElevation e' coerente con la geometria (coordinate interne), Elevation puo' riferirsi al Survey Point
    value = getattr(level, "ProjectElevation", None)
    return level.Elevation if value is None else value


def get_levels():
    levels = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
    return sorted(levels, key=level_elevation)


def levels_from_active_view(levels):
    """Livello della vista attiva e quelli nella direzione del riser (sopra o sotto).
    Se la vista non ha un livello (3D, sezione) li lascia tutti."""
    view_level = getattr(doc.ActiveView, "GenLevel", None)
    if view_level is None:
        return levels
    base = level_elevation(view_level)
    return [lvl for lvl in levels if SIGN * (level_elevation(lvl) - base) >= -1e-6]


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


def element_midpoint(element):
    curve = element.Location.Curve
    return curve.Evaluate(0.5, True)


def copy_size(source, target, bips):
    """Copia le dimensioni indicate. Nessuno scambio: l'orientamento si sistema con align_section."""
    for bip in bips:
        src = source.get_Parameter(bip)
        tgt = target.get_Parameter(bip)
        if src and tgt and src.HasValue and not tgt.IsReadOnly:
            tgt.Set(src.AsDouble())


TRAY_SIZE = (DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
DUCT_SIZE = (DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM, DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM,
             DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)


def horizontal_axis(cs):
    """(nome, vettore) dell'asse orizzontale nel piano del connettore."""
    if abs(cs.BasisX.Z) < 0.01:
        return "X", cs.BasisX
    return "Y", cs.BasisY


def align_section(source_conn, new_element, start_point):
    """Ruota il tratto verticale sul proprio asse: la sezione deve essere orientata come quella di partenza.

    Senza rotazione Revit orienta la sezione sugli assi del progetto e, con tratti non paralleli agli assi,
    il gomito viene storto o compaiono raccordi di transizione.
    """
    if source_conn.Shape == DB.ConnectorProfileType.Round:
        return
    doc.Regenerate()
    new_conn = nearest_connector(new_element, start_point)
    name, src_axis = horizontal_axis(source_conn.CoordinateSystem)
    new_cs = new_conn.CoordinateSystem
    new_axis = new_cs.BasisX if name == "X" else new_cs.BasisY
    angle = math.atan2(src_axis.Y, src_axis.X) - math.atan2(new_axis.Y, new_axis.X)
    # sezione simmetrica: basta allineare la direzione, a meno di 180 gradi
    while angle > math.pi / 2:
        angle -= math.pi
    while angle <= -math.pi / 2:
        angle += math.pi
    if abs(angle) > 1e-6:
        axis = DB.Line.CreateBound(start_point, start_point.Add(DB.XYZ.BasisZ))
        DB.ElementTransformUtils.RotateElement(doc, new_element.Id, axis, angle)
        doc.Regenerate()


def copy_conduit_diameter(source, target):
    for bip in (DB.BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM,
                DB.BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM):
        src = source.get_Parameter(bip)
        tgt = target.get_Parameter(bip)
        if src and tgt and not tgt.IsReadOnly:
            tgt.Set(src.AsDouble())
            return


class RiserError(Exception):
    """Errore atteso (es. spazio insufficiente): il messaggio va nel report."""


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
    if fitting is None:
        return "none (straight run)"
    try:
        return "{}: {}".format(fitting.Symbol.FamilyName, fitting.Name)
    except Exception:
        return "elbow"


def type_name(element_type):
    try:
        return element_type.Name
    except Exception:
        param = element_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
        return param.AsString() if param else "?"


def duct_system_type(element):
    """(id tipo di sistema, nota). Se il condotto non ne ha uno, si usa il primo tipo di sistema del modello."""
    param = element.get_Parameter(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)
    system_id = param.AsElementId() if param else DB.ElementId.InvalidElementId
    if system_id != DB.ElementId.InvalidElementId:
        return system_id, ""
    types = list(DB.FilteredElementCollector(doc).OfClass(MechanicalSystemType))
    if not types:
        raise RiserError("The duct has no system type and the model has no duct system types")
    chosen = sorted(types, key=type_name)[0]
    return chosen.Id, "duct had no system: system type '{}' assigned to the new run".format(type_name(chosen))


def create_segment(element, conn, end_point, level_id):
    """Crea il tratto verticale (senza garantire il raccordo). Ritorna (elemento creato, nota per il report)."""
    type_id = element.GetTypeId()
    note = ""
    if isinstance(element, Pipe):
        # sezione tonda: puo' inserire da solo il gomito, lo si verifica in ensure_connection
        return Pipe.Create(doc, type_id, level_id, conn, end_point), note
    if isinstance(element, Duct):
        # due punti, non da connettore: la sezione va orientata prima di inserire il gomito
        system_id, note = duct_system_type(element)
        new_element = Duct.Create(doc, system_id, type_id, level_id, conn.Origin, end_point)
        copy_size(element, new_element, DUCT_SIZE)
    elif isinstance(element, CableTray):
        new_element = CableTray.Create(doc, type_id, conn.Origin, end_point, level_id)
        copy_size(element, new_element, TRAY_SIZE)
    else:
        new_element = Conduit.Create(doc, type_id, conn.Origin, end_point, level_id)
        copy_conduit_diameter(element, new_element)
    align_section(conn, new_element, conn.Origin)
    return new_element, note


def ensure_connection(element, conn, new_element, straight):
    """Garantisce il collegamento col nuovo tratto. Ritorna il gomito, o None se continuita' diretta."""
    new_conn = nearest_connector(new_element, conn.Origin)
    if straight:
        if not new_conn.IsConnectedTo(conn):
            new_conn.ConnectTo(conn)
        return None
    fitting = find_fitting(conn, element)
    if fitting is None:
        # collegamento diretto senza raccordo: si scollega e si posiziona il gomito esplicitamente
        if new_conn.IsConnectedTo(conn):
            conn.DisconnectFrom(new_conn)
        fitting = doc.Create.NewElbowFitting(conn, new_conn)
    return fitting


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

def check_remaining_space(element, new_element, fitting):
    """Dopo il gomito, i due tratti devono avere ancora una lunghezza utile."""
    if fitting is None:
        return
    doc.Regenerate()
    for target, name in ((element, "existing element"), (new_element, "riser")):
        remaining = target.Location.Curve.Length * MM_PER_FT
        if remaining < MIN_REMAINING_MM:
            raise RiserError("Not enough space for the elbow: the {} would be only {:.0f} mm long "
                             "(minimum {:.0f} mm)".format(name, remaining, MIN_REMAINING_MM))


def process_pick(reference, settings):
    """Elabora un elemento. Ritorna un dict con l'esito (usato per il report)."""
    element = doc.GetElement(reference.ElementId)
    result = {"element": element, "category": category_name(element), "status": "",
              "ok": False, "new": None, "fitting": "", "added_mm": None, "z_mm": None,
              "level": "", "rel_mm": None}

    click = reference.GlobalPoint
    if click is None:
        click = element_midpoint(element)

    conn = nearest_free_connector(element, click)
    if conn is None:
        result["status"] = "No free end (both ends are connected)"
        return result

    facing = SIGN * conn.CoordinateSystem.BasisZ.DotProduct(DB.XYZ.BasisZ)
    if facing < -STRAIGHT_TOL:
        result["status"] = "The end faces {}: the riser would overlap the element".format(WORDS["opposite"])
        return result
    straight = facing > STRAIGHT_TOL

    level_id = get_level_id(element)
    if level_id is None:
        result["status"] = "No reference level on the element"
        return result
    ref_level = doc.GetElement(level_id)
    result["level"] = ref_level.Name

    origin = conn.Origin
    if settings["mode"] == "level":
        target_z = level_elevation(settings["level"]) + settings["offset_mm"] / MM_PER_FT
        length = SIGN * (target_z - origin.Z)
    else:
        length = settings["length_mm"] / MM_PER_FT
    if length < MIN_LENGTH_FT:
        result["status"] = "Target is not {} the element end ({:.0f} mm)".format(
            WORDS["side"], length * MM_PER_FT)
        return result

    end_point = DB.XYZ(origin.X, origin.Y, origin.Z + SIGN * length)

    transaction = DB.Transaction(doc, "Riser {}".format(RISER_DIRECTION.capitalize()))
    transaction.Start()
    try:
        new_element, note = create_segment(element, conn, end_point, level_id)
        fitting = ensure_connection(element, conn, new_element, straight)
        check_remaining_space(element, new_element, fitting)
        inherited = inherit(element, [new_element], [fitting])
        transaction.Commit()
    except RiserError as ex:
        transaction.RollBack()
        result["status"] = "Not created. {}".format(ex)
        return result
    except Exception as ex:
        transaction.RollBack()
        result["status"] = ("Not created. The elbow could not be placed (not enough space, or no valid "
                            "fitting in the routing preferences): {}".format(ex))
        return result

    result.update({"ok": True, "status": "Created ({})".format("; ".join(t for t in (note, inherited) if t)),
                   "new": new_element,
                   "fitting": fitting_label(fitting),
                   "added_mm": length * MM_PER_FT, "z_mm": end_point.Z * MM_PER_FT,
                   "rel_mm": (end_point.Z - level_elevation(ref_level)) * MM_PER_FT})
    return result


def fmt(value):
    return "" if value is None else "{:.1f}".format(value)


def settings_text(settings):
    if settings["mode"] == "level":
        return "to {} {:+.0f} mm".format(settings["level"].Name, settings["offset_mm"])
    return "{:g} mm {} the selected end".format(settings["length_mm"], WORDS["side"])


def print_report(results, settings):
    output.print_md("# Riser {} report".format(RISER_DIRECTION.capitalize()))
    output.print_md("Height: **{}**".format(settings_text(settings)))

    done = [r for r in results if r["ok"]]
    output.print_md("Elements picked: **{}** | risers created: **{}** | not processed: **{}**".format(
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
                 "Final Z, internal coordinates [mm]", "Reference level", "Final Z vs reference level [mm]"])


def load_saved(config):
    saved = dict(DEFAULTS)
    try:
        saved["mode"] = config.get_option("mode", DEFAULTS["mode"])
        saved["length_mm"] = float(config.get_option("length_mm", DEFAULTS["length_mm"]))
        saved["offset_mm"] = float(config.get_option("offset_mm", DEFAULTS["offset_mm"]))
        saved["level_name"] = config.get_option("level_name", DEFAULTS["level_name"])
    except Exception:
        pass
    return saved


def resolve_settings(saved, levels):
    """Impostazioni pronte all'uso. None se il livello salvato non e' disponibile in questa vista."""
    settings = dict(saved)
    if saved["mode"] == "level":
        matches = [lvl for lvl in levels if lvl.Name == saved["level_name"]]
        if not matches:
            return None
        settings["level"] = matches[0]
    return settings


def ask_settings(saved, levels):
    """SHIFT+CLICK (o livello salvato non disponibile): finestra delle impostazioni."""
    xaml_path = script.get_bundle_file(XAML_FILE_NAME)
    level_info = [(lvl.Name, level_elevation(lvl) * MM_PER_FT) for lvl in levels]
    values = Riser_ui.RiserWindow(xaml_path, level_info, saved, RISER_DIRECTION).show()
    if values is None:
        script.exit()
    settings = dict(saved)
    settings["mode"] = values["mode"]
    if values["mode"] == "level":
        settings["offset_mm"] = values["offset_mm"]
        settings["level"] = levels[values["level_index"]]
        settings["level_name"] = settings["level"].Name
    else:
        settings["length_mm"] = values["length_mm"]
    return settings


def save_settings(config, settings):
    try:
        config.mode = settings["mode"]
        config.length_mm = settings["length_mm"]
        config.offset_mm = settings["offset_mm"]
        config.level_name = settings["level_name"]
        script.save_config()
    except Exception:
        pass


def main():
    levels = levels_from_active_view(get_levels())
    config = script.get_config(CONFIG_NAME)
    saved = load_saved(config)

    settings = None if __shiftclick__ else resolve_settings(saved, levels)  # noqa: F821
    if settings is None:
        settings = ask_settings(saved, levels)
        save_settings(config, settings)

    results = []
    selection_filter = MepCurveFilter()
    prompt = "Riser {} ({}): click near the free end (ESC to finish)".format(
        RISER_DIRECTION, settings_text(settings))
    while True:
        try:
            reference = uidoc.Selection.PickObject(ObjectType.Element, selection_filter, prompt)
        except OperationCanceledException:
            break
        results.append(process_pick(reference, settings))

    print_report(results, settings)


main()
