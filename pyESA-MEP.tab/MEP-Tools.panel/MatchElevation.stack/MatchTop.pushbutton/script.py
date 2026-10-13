# -*- coding: utf-8 -*-
__title__ = "Match\nTop"
__doc__ = """Version = 0.1
Date    = 09.10.2026
________________________________________________________________
Description:
Aligns pipes, ducts, cable trays and conduits to the upper end top
elevation of a reference element.

CLICK:
1. Click the reference element
2. Select the elements to align (click or window) and press Finish
3. Each element is moved vertically so that its upper end top
   elevation matches the one of the reference. Plan position and
   length are kept
4. The report lists the elements, how much they moved and what
   could not be changed

SHIFT + CLICK: settings shared by Match Top, Match Middle and Match
Bottom:
- Overwrite the reference level: the elements get the reference level
  of the first element (on by default)
- Remove the slope: sloped elements become horizontal at the matched
  elevation (off by default; otherwise they keep their slope and are
  moved as a whole)

Vertical elements can be the reference (their top end is used) or be
aligned: they are moved as a whole, never made horizontal.
Elements connected to the moved ones follow as when the elevation is
changed in the Properties palette; if Revit refuses the result, the
element is left unchanged and the reason goes to the report.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import clr
import math
import os.path as op
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, DB, forms

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

MATCH_MODE = "top"  # "top" | "middle" | "bottom": unica differenza tra i tre pulsanti dello stack

MODE_LABELS = {"top": "upper end top elevation", "middle": "middle elevation",
               "bottom": "lower end bottom elevation"}
TOOL_TITLES = {"top": "Match Top", "middle": "Match Middle", "bottom": "Match Bottom"}
TOOL_TITLE = TOOL_TITLES[MATCH_MODE]

SETTINGS_CONFIG = "ESA_MatchElevation"  # condivisa dai tre pulsanti, si cambia con SHIFT + CLICK
DEFAULT_SETTINGS = {"override_level": True, "remove_slope": False}
XAML_FILE_NAME = "MatchElevation_form.xaml"

REPORT_CONFIG = "ESA_Report"  # impostazione condivisa da tutti i pulsanti, cambiata col pulsante Report

# Match Middle con tratti inclinati: "upper" / "lower" = Upper / Lower End Centerline Elevation, scelto dall'utente
# nel popup. None: nessun tratto inclinato, si usa il centro dell'asse (per un tratto orizzontale e' lo stesso)
MIDDLE_END = [None]
MIDDLE_END_LABELS = {"upper": "upper end centerline elevation", "lower": "lower end centerline elevation"}


def mode_label():
    if MATCH_MODE == "middle" and MIDDLE_END[0] is not None:
        return MIDDLE_END_LABELS[MIDDLE_END[0]]
    return MODE_LABELS[MATCH_MODE]


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
TOL_FT = 0.1 / MM_PER_FT  # sotto 0,1 mm una differenza di quota non si considera
VERTICAL_TOL_DEG = 5.0  # scostamento massimo dall'asse Z per considerare verticale un elemento
CLASSES = (Pipe, Duct, CableTray, Conduit)


class MatchError(Exception):
    """Errore atteso: il messaggio va nel report."""


class RevitErrors(DB.IFailuresPreprocessor):
    """Errori di Revit al commit (es. elementi collegati che non si possono adattare): invece della finestra
    bloccante, la transazione viene annullata e i messaggi vanno nel report. Gli avvisi passano come sempre."""

    def __init__(self):
        self.messages = []

    def PreprocessFailures(self, accessor):
        for failure in accessor.GetFailureMessages():
            if failure.GetSeverity() != DB.FailureSeverity.Warning:
                self.messages.append(failure.GetDescriptionText())
        if self.messages:
            return DB.FailureProcessingResult.ProceedWithRollBack
        return DB.FailureProcessingResult.Continue


# --- impostazioni condivise -------------------------------------------------------------------------------------

def as_bool(value, default):
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return str(value).strip().lower() in ("true", "1", "yes")


def load_settings():
    settings = dict(DEFAULT_SETTINGS)
    try:
        config = script.get_config(SETTINGS_CONFIG)
        for key, default in DEFAULT_SETTINGS.items():
            settings[key] = as_bool(config.get_option(key, default), default)
    except Exception:
        pass
    return settings


def save_settings(settings):
    try:
        config = script.get_config(SETTINGS_CONFIG)
        for key, value in settings.items():
            setattr(config, key, bool(value))
        script.save_config()
    except Exception:
        pass


def settings_main():
    """SHIFT + CLICK: finestra delle impostazioni, condivisa dai tre pulsanti. Salva ed esce, non modifica il modello."""
    from MatchElevation_ui import SettingsWindow
    xaml_path = script.get_bundle_file(XAML_FILE_NAME) or op.join(op.dirname(__file__), XAML_FILE_NAME)
    result = SettingsWindow(xaml_path, load_settings()).show()
    if result is None:
        script.exit()
    save_settings(result)
    try:
        script.toast("Match elevation settings saved: level {}, slope {}".format(
            "overwritten" if result["override_level"] else "kept",
            "removed" if result["remove_slope"] else "kept"), title="pyESA")
    except Exception:
        pass


# --- geometria --------------------------------------------------------------------------------------------------

class CurveFilter(ISelectionFilter):
    def __init__(self, exclude_id=None):
        self.exclude_id = exclude_id

    def AllowElement(self, element):
        if self.exclude_id is not None and element.Id == self.exclude_id:
            return False
        return isinstance(element, CLASSES) and isinstance(getattr(element.Location, "Curve", None), DB.Line)

    def AllowReference(self, reference, point):
        return False


def category_name(element):
    for cls, name in ((Pipe, "Pipe"), (Duct, "Duct"), (CableTray, "Cable Tray"), (Conduit, "Conduit")):
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


def level_name(level_id):
    level = doc.GetElement(level_id) if level_id is not None else None
    return level.Name if level is not None else "no level"


def is_vertical(direction):
    return abs(direction.Z) >= math.cos(math.radians(VERTICAL_TOL_DEG))


def end_connector(element):
    for conn in element.ConnectorManager.Connectors:
        if conn.ConnectorType == DB.ConnectorType.End:
            return conn
    return None


def param_double(element, bip):
    param = element.get_Parameter(bip)
    if param is not None and param.HasValue and param.AsDouble() > 0:
        return param.AsDouble()
    return None


def bbox_half_height(element):
    """Stima della meta' dell'ingombro verticale della sezione dal bounding box: altezza del box meno il dislivello
    dell'asse. None se il box non c'e'."""
    box = element.get_BoundingBox(None)
    if box is None:
        return None
    curve = element.Location.Curve
    rise = abs(curve.GetEndPoint(1).Z - curve.GetEndPoint(0).Z)
    return max(0.0, (box.Max.Z - box.Min.Z - rise) / 2.0)


def round_radius(element, conn):
    """Raggio esterno per le sezioni circolari (tubi, conduit, canali tondi); None per le altre."""
    if isinstance(element, Pipe):
        diameter = param_double(element, DB.BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
    elif isinstance(element, Conduit):
        diameter = param_double(element, DB.BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM)
    elif isinstance(element, Duct) and conn is not None and conn.Shape == DB.ConnectorProfileType.Round:
        diameter = param_double(element, DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if diameter is None:
            return conn.Radius
    else:
        return None
    if diameter is None and conn is not None and conn.Shape == DB.ConnectorProfileType.Round:
        return conn.Radius
    return diameter / 2.0 if diameter is not None else None


def half_height(element, horizontal=False):
    """Meta' dell'ingombro verticale della sezione (piedi), misurato sull'asse. Vale 0 per un elemento verticale.
    horizontal=True: come se l'elemento fosse orizzontale (serve quando si toglie la pendenza)."""
    direction = element.Location.Curve.Direction
    factor = 1.0 if horizontal else math.sqrt(max(0.0, 1.0 - direction.Z ** 2))  # coseno della pendenza
    conn = end_connector(element)
    radius = round_radius(element, conn)
    if radius is not None:
        return radius * factor
    if conn is not None:
        try:
            # sezione rettangolare / ovale, anche ruotata: larghezza e altezza proiettate sulla verticale. Quale
            # delle due stia su BasisX non e' documentato con certezza: tra le due combinazioni si sceglie quella
            # piu' vicina all'ingombro letto dal bounding box
            basis = conn.CoordinateSystem
            half_w, half_h = conn.Width / 2.0, conn.Height / 2.0
            option_a = abs(basis.BasisX.Z) * half_w + abs(basis.BasisY.Z) * half_h
            option_b = abs(basis.BasisX.Z) * half_h + abs(basis.BasisY.Z) * half_w
            if abs(option_a - option_b) < TOL_FT:
                return option_a
            estimate = bbox_half_height(element)
            if estimate is None:
                return option_a
            return option_a if abs(option_a - estimate) <= abs(option_b - estimate) else option_b
        except Exception:
            pass
    estimate = bbox_half_height(element)
    if estimate is not None:
        return estimate
    for bip in (DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM):
        value = param_double(element, bip)
        if value is not None:
            return value / 2.0 * factor
    return 0.0


def mode_elevation(element, mode):
    """Quota assoluta (coordinate interne, piedi) del punto che il pulsante allinea: sommita' dell'estremo piu' alto,
    centro dell'asse, fondo dell'estremo piu' basso. Per un elemento verticale: estremo superiore / centro /
    estremo inferiore. Middle su un tratto inclinato: quota d'asse dell'estremo scelto nel popup (MIDDLE_END)."""
    curve = element.Location.Curve
    z0, z1 = curve.GetEndPoint(0).Z, curve.GetEndPoint(1).Z
    if mode == "middle":
        if MIDDLE_END[0] is not None and not is_vertical(curve.Direction):
            return max(z0, z1) if MIDDLE_END[0] == "upper" else min(z0, z1)
        return (z0 + z1) / 2.0
    if mode == "top":
        return max(z0, z1) + half_height(element)
    return min(z0, z1) - half_height(element)


def axis_elevation_for(element, target, mode):
    """Quota dell'asse orizzontale che porta l'elemento, senza pendenza, alla quota target."""
    if mode == "middle":
        return target
    h = half_height(element, horizontal=True)
    return target - h if mode == "top" else target + h


def mm(value_ft):
    return value_ft * MM_PER_FT


def signed_mm(value_ft):
    return "{:+.0f} mm".format(mm(value_ft))




# --- allineamento -----------------------------------------------------------------------------------------------

def set_level(element, level_id):
    param = element.get_Parameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM)
    if param is not None and not param.IsReadOnly:
        param.Set(level_id)
        return
    try:
        element.ReferenceLevel = doc.GetElement(level_id)
    except Exception as ex:
        raise MatchError("The reference level of {} {} could not be changed: {}".format(
            category_name(element), element.Id, ex))


def make_plan(element, target, ref_level_id, settings):
    """Cosa va fatto sull'elemento, calcolato sulla geometria di partenza senza toccare il modello. La nuova
    posizione e' assoluta (estremi dell'asse), cosi' non dipende da come Revit tratta il cambio di livello."""
    curve = element.Location.Curve
    a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
    vertical = is_vertical(curve.Direction)
    plan = {"element": element, "notes": [], "change_level": False, "flatten": False, "line": None,
            "old_level": level_name(get_level_id(element)), "status": "", "problem": False}
    level_id = get_level_id(element)
    if ref_level_id is not None and level_id != ref_level_id:
        if settings["override_level"]:
            plan["change_level"] = True
        else:
            plan["notes"].append("level kept ({})".format(level_name(level_id)))
    sloped = abs(b.Z - a.Z) > TOL_FT
    plan["flatten"] = settings["remove_slope"] and sloped and not vertical
    if vertical:
        plan["notes"].append("vertical: moved as a whole")
    elif sloped and not plan["flatten"]:
        plan["notes"].append("slope kept")
    plan["shift"] = target - mode_elevation(element, MATCH_MODE)
    if plan["flatten"]:
        axis_z = axis_elevation_for(element, target, MATCH_MODE)
        plan["line"] = DB.Line.CreateBound(DB.XYZ(a.X, a.Y, axis_z), DB.XYZ(b.X, b.Y, axis_z))
    elif abs(plan["shift"]) > TOL_FT:
        up = DB.XYZ(0.0, 0.0, plan["shift"])
        plan["line"] = DB.Line.CreateBound(a.Add(up), b.Add(up))
    elif plan["change_level"]:
        plan["line"] = DB.Line.CreateBound(a, b)  # stessa posizione: la blocca anche se il cambio di livello la sposta
    plan["todo"] = plan["change_level"] or plan["line"] is not None
    return plan


def apply_plans(plans, ref_level_id):
    """Dentro la transazione. Prima tutti i livelli, poi tutte le geometrie, senza rigenerare tra un elemento e
    l'altro: gli elementi selezionati e collegati tra loro (es. tratti uniti da un manicotto) si spostano insieme,
    come quando in Revit si cambia la quota di piu' elementi selezionati."""
    level_plans = [p for p in plans if p["change_level"]]
    for plan in level_plans:
        set_level(plan["element"], ref_level_id)
    if level_plans:
        doc.Regenerate()
    for plan in plans:
        if plan["line"] is not None:
            plan["element"].Location.Curve = plan["line"]
    doc.Regenerate()


def run_transaction(plans, ref_level_id):
    """Applica i piani in una transazione. Ritorna None se confermata, altrimenti il motivo (modello invariato)."""
    transaction = DB.Transaction(doc, TOOL_TITLE)
    errors = RevitErrors()
    options = transaction.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(errors)
    options.SetClearAfterRollback(True)
    transaction.SetFailureHandlingOptions(options)
    transaction.Start()
    try:
        apply_plans(plans, ref_level_id)
        if transaction.Commit() != DB.TransactionStatus.Committed:
            return "Revit refused the result: {}".format("; ".join(errors.messages) or "the transaction was rolled back")
    except MatchError as ex:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        return str(ex)
    except Exception as ex:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        return "The element could not be moved: {}".format(ex)
    return None


def describe_done(plan, target, ref_level_id):
    """Riga del report per un elemento modificato, con il controllo della quota finale: un elemento collegato a un
    altro spostato dopo puo' essere stato trascinato via."""
    parts = []
    if abs(plan["shift"]) > TOL_FT:
        parts.append("Moved {}".format(signed_mm(plan["shift"])))
    if plan["flatten"]:
        parts.append("slope removed")
    if plan["change_level"]:
        parts.append("level {} -> {}".format(plan["old_level"], level_name(ref_level_id)))
    parts = [parts[0][0].upper() + parts[0][1:]] + parts[1:]
    status = ", ".join(parts + plan["notes"])
    try:
        gap = mode_elevation(plan["element"], MATCH_MODE) - target
    except Exception:
        gap = 0.0
    if abs(gap) > 1.0 / MM_PER_FT:
        plan["problem"] = True
        status += (". CHECK: the {} ended {} from the reference, probably moved by a connected element changed "
                   "later".format(mode_label(), signed_mm(gap)))
    return status


def process_all(elements, target, ref_level_id, settings):
    """Ritorna (piani con lo stato, nota sul metodo usato)."""
    plans = []
    for element in elements:
        try:
            plan = make_plan(element, target, ref_level_id, settings)
        except Exception as ex:
            plan = {"element": element, "todo": False, "problem": True,
                    "status": "Unchanged. The geometry could not be read: {}".format(ex)}
            plans.append(plan)
            continue
        if not plan["todo"]:
            plan["status"] = ", ".join(["Already matching"] + plan["notes"])
        plans.append(plan)
    todo = [p for p in plans if p["todo"]]
    if not todo:
        return plans, ""
    reason = run_transaction(todo, ref_level_id)
    if reason is None:
        for plan in todo:
            plan["status"] = describe_done(plan, target, ref_level_id)
        return plans, ""
    if len(todo) == 1:
        todo[0].update({"status": "Unchanged. {}".format(reason), "problem": True})
        return plans, ""
    # Revit rifiuta l'insieme: un elemento alla volta, cosi' si salva quello che si puo' e si sa chi da' problemi
    for plan in todo:
        try:
            plan.update(make_plan(plan["element"], target, ref_level_id, settings))
        except Exception as ex:
            plan.update({"status": "Unchanged. The geometry could not be read: {}".format(ex), "problem": True})
            continue
        if not plan["todo"]:
            plan["status"] = ", ".join(["Already matching"] + plan["notes"])
            continue
        single = run_transaction([plan], ref_level_id)
        if single is None:
            plan["status"] = describe_done(plan, target, ref_level_id)
        else:
            plan.update({"status": "Unchanged. {}".format(single), "problem": True})
    return plans, ("Revit refused to change all the elements together ({}), so they were changed one at a time."
                   .format(reason))


def pick_reference():
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, CurveFilter(),
                                         "{}: click the reference pipe, duct, cable tray or conduit".format(TOOL_TITLE))
    except OperationCanceledException:
        script.exit()
    return doc.GetElement(ref.ElementId)


def pick_targets(reference, value_text):
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, CurveFilter(reference.Id),
            "{}: select the elements to align to {} {} ({}), then press Finish".format(
                TOOL_TITLE, category_name(reference), reference.Id, value_text))
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def is_sloped(element):
    """Tratto inclinato: estremi a quote diverse, ma non verticale."""
    curve = element.Location.Curve
    return not is_vertical(curve.Direction) and abs(curve.GetEndPoint(1).Z - curve.GetEndPoint(0).Z) > TOL_FT


def ask_middle_end():
    """Match Middle con tratti inclinati: la quota d'asse non e' unica, si sceglie quale estremo allineare."""
    options = {"Upper end centerline elevation": "upper", "Lower end centerline elevation": "lower"}
    choice = forms.CommandSwitchWindow.show(
        sorted(options.keys(), reverse=True),
        message="Some elements are sloped. Which centerline elevation should be aligned?")
    if not choice:
        script.exit()
    MIDDLE_END[0] = options[choice]


def value_text(reference, ref_level_id):
    """Quota allineata del riferimento, dal suo livello, per il prompt e il report."""
    value = mode_elevation(reference, MATCH_MODE)
    level = doc.GetElement(ref_level_id) if ref_level_id is not None else None
    if level is not None:
        return "{} {:.0f} mm from {}".format(mode_label(), mm(value - level_elevation(level)), level.Name)
    return "{} {:.0f} mm from the internal origin".format(mode_label(), mm(value))


def main():
    settings = load_settings()
    reference = pick_reference()
    ref_level_id = get_level_id(reference)
    # Match Middle: se il riferimento e' inclinato si chiede subito l'estremo, cosi' il prompt mostra la quota giusta
    if MATCH_MODE == "middle" and is_sloped(reference):
        ask_middle_end()
    elements = pick_targets(reference, value_text(reference, ref_level_id))
    if not elements:
        script.exit()
    if MATCH_MODE == "middle" and MIDDLE_END[0] is None and any(is_sloped(e) for e in elements):
        ask_middle_end()
    target = mode_elevation(reference, MATCH_MODE)
    ref_text = value_text(reference, ref_level_id)

    plans, method_note = process_all(elements, target, ref_level_id, settings)
    changed = len([p for p in plans if p["todo"] and not p["status"].startswith("Unchanged")])
    problems = len([p for p in plans if p["problem"]])
    if not want_report(len(elements), problems, "elements aligned"):
        return

    output.print_md("# {} report".format(TOOL_TITLE))
    curve = reference.Location.Curve
    ref_shape = ""
    if is_vertical(curve.Direction):
        ref_shape = " (vertical)"
    elif abs(curve.GetEndPoint(1).Z - curve.GetEndPoint(0).Z) > TOL_FT:
        ref_shape = " (sloped)"
    output.print_md("Reference: {} {}{} | level: {} | {}".format(
        category_name(reference), output.linkify(reference.Id), ref_shape, level_name(ref_level_id), ref_text))
    output.print_md("Settings: reference level **{}** | slope **{}** (SHIFT + CLICK to change)".format(
        "overwritten" if settings["override_level"] else "kept",
        "removed" if settings["remove_slope"] else "kept"))
    output.print_md("Elements: **{}** | changed: **{}** | already matching: **{}** | with problems: **{}**".format(
        len(elements), changed, len([p for p in plans if not p["todo"] and not p["problem"]]), problems))
    if method_note:
        output.print_md("_{}_".format(method_note))
    # righe in testo, non in tabella: si vedono sempre, anche se la tabella non le mostra
    for i, plan in enumerate(plans, 1):
        element = plan["element"]
        output.print_md("**#{}** {} {} | {}".format(i, category_name(element), output.linkify(element.Id),
                                                   plan["status"]))


if __shiftclick__:  # noqa: F821
    settings_main()
else:
    main()
