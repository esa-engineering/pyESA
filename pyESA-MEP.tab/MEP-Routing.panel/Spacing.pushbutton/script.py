# -*- coding: utf-8 -*-
__title__ = "Spacing"
__doc__ = """Version = 0.1
Date    = 09.10.2026
________________________________________________________________
Description:
Sets the same distance between parallel pipes, ducts, cable trays and
conduits, insulation included, and checks that the accessories on
them do not clash.

CLICK:
1. Select the elements (or select them first) and press Finish
2. In the window set the distance and which element stays fixed
3. The elements are moved sideways, in the plane of the active view,
   so that the clear distance between neighbours is the one you set
4. The accessories connected to the elements (valves, dampers) are
   checked against the neighbouring elements and their insulation;
   clashes are listed in the report

SHIFT + CLICK: copies the spacing from one series of elements to
another one with the same number of lines:
1. Select the reference series and press Finish
2. Select the series to change and press Finish
The lines of the second series are moved so that the distances between
neighbours are the same as in the first one, matching them in order
(left to right or bottom to top in the view). Measure (faces with
insulation or centerlines) and fixed line are the ones last used with
CLICK. If the two series do not have the same number of lines nothing
is changed.

Run it in a plan view (elements side by side) or in a section or
elevation view (elements one above the other, or side by side).
Elements must be parallel within 10 degrees of their mean direction;
the others are left out and listed in the report. Straight segments of
the same line (e.g. on both sides of a valve) move together.
The distance is measured between the outer faces, so the insulation of
pipes and ducts counts; as an option it can be measured between the
centerlines (insulation ignored).
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
from System.Collections.Generic import List

from pyrevit import revit, script, DB, forms

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

SETTINGS_CONFIG = "ESA_Spacing"
DEFAULT_SETTINGS = {"gap_mm": 100.0, "measure": "faces", "anchor": "first"}
MEASURES = ("faces", "centerlines")
ANCHORS = ("first", "last", "centre")
ANCHOR_LABELS = {"first": "leftmost / lowest stays", "last": "rightmost / highest stays",
                 "centre": "centre of the group stays"}
XAML_FILE_NAME = "Spacing_form.xaml"

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
TOL_FT = 0.1 / MM_PER_FT          # sotto 0,1 mm uno spostamento non conta
SAME_LINE_FT = 1.0 / MM_PER_FT    # centri a meno di 1 mm: tratti della stessa linea, si muovono insieme
ACHIEVED_TOL_FT = 1.0 / MM_PER_FT  # scarto massimo accettato tra distanza voluta e ottenuta
MAX_ANGLE_DEG = 10.0
CLASSES = (Pipe, Duct, CableTray, Conduit)
ACCESSORY_CATEGORIES = (DB.BuiltInCategory.OST_PipeAccessory, DB.BuiltInCategory.OST_DuctAccessory)


class SpacingError(Exception):
    """Errore atteso: il messaggio va nel report."""


class RevitErrors(DB.IFailuresPreprocessor):
    """Errori di Revit al commit: invece della finestra bloccante, la transazione viene annullata e i messaggi
    vanno nel report. Gli avvisi passano come sempre."""

    def __init__(self):
        self.messages = []

    def PreprocessFailures(self, accessor):
        for failure in accessor.GetFailureMessages():
            if failure.GetSeverity() != DB.FailureSeverity.Warning:
                self.messages.append(failure.GetDescriptionText())
        if self.messages:
            return DB.FailureProcessingResult.ProceedWithRollBack
        return DB.FailureProcessingResult.Continue


# --- impostazioni -----------------------------------------------------------------------------------------------

def load_settings():
    settings = dict(DEFAULT_SETTINGS)
    try:
        config = script.get_config(SETTINGS_CONFIG)
        settings["gap_mm"] = float(config.get_option("gap_mm", DEFAULT_SETTINGS["gap_mm"]))
        measure = config.get_option("measure", DEFAULT_SETTINGS["measure"])
        anchor = config.get_option("anchor", DEFAULT_SETTINGS["anchor"])
        settings["measure"] = measure if measure in MEASURES else DEFAULT_SETTINGS["measure"]
        settings["anchor"] = anchor if anchor in ANCHORS else DEFAULT_SETTINGS["anchor"]
    except Exception:
        pass
    return settings


def save_settings(settings):
    try:
        config = script.get_config(SETTINGS_CONFIG)
        for key, value in settings.items():
            setattr(config, key, value)
        script.save_config()
    except Exception:
        pass


# --- selezione --------------------------------------------------------------------------------------------------

class CurveFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, CLASSES) and isinstance(getattr(element.Location, "Curve", None), DB.Line)

    def AllowReference(self, reference, point):
        return False


def get_selection():
    selected = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    elements = [e for e in selected if CurveFilter().AllowElement(e)]
    if len(elements) >= 2:
        return elements
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, CurveFilter(),
                                           "Select the pipes, ducts, cable trays or conduits to space, then press "
                                           "Finish")
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def category_name(element):
    for cls, name in ((Pipe, "Pipe"), (Duct, "Duct"), (CableTray, "Cable Tray"), (Conduit, "Conduit")):
        if isinstance(element, cls):
            return name
    name = element.Category.Name if element.Category is not None else "Element"
    return name


# --- geometria --------------------------------------------------------------------------------------------------

def mm(value_ft):
    return value_ft * MM_PER_FT


def unit(vector):
    length = vector.GetLength()
    return vector.Multiply(1.0 / length) if length > 1e-12 else vector


def midpoint(element):
    return element.Location.Curve.Evaluate(0.5, True)


def param_double(element, bip):
    param = element.get_Parameter(bip)
    if param is not None and param.HasValue and param.AsDouble() > 0:
        return param.AsDouble()
    return None


def end_connector(element):
    for conn in element.ConnectorManager.Connectors:
        if conn.ConnectorType == DB.ConnectorType.End:
            return conn
    return None


def round_radius(element):
    """Raggio esterno (senza isolamento) per le sezioni circolari; None per le altre."""
    conn = end_connector(element)
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


def insulation_of(element):
    """Elementi isolamento del tratto (tubi e canali; per le altre categorie la lista e' vuota)."""
    if not isinstance(element, (Pipe, Duct)):
        return []
    try:
        ids = DB.InsulationLiningBase.GetInsulationIds(doc, element.Id)
    except Exception:
        return []
    return [e for e in (doc.GetElement(i) for i in ids) if e is not None]


def insulation_thickness(element):
    values = [getattr(ins, "Thickness", 0.0) or 0.0 for ins in insulation_of(element)]
    return max(values) if values else 0.0


def solid_extent(element, u):
    """(min, max) della proiezione su u dei vertici del solido: esatto per le sezioni rettangolari. None se la
    geometria non e' leggibile."""
    options = DB.Options()
    options.DetailLevel = DB.ViewDetailLevel.Fine
    geometry = element.get_Geometry(options)
    if geometry is None:
        return None
    low, high = None, None
    for item in geometry:
        solids = [item] if isinstance(item, DB.Solid) else []
        if isinstance(item, DB.GeometryInstance):
            solids = [s for s in item.GetInstanceGeometry() if isinstance(s, DB.Solid)]
        for solid in solids:
            if solid.Volume <= 0:
                continue
            for edge in solid.Edges:
                for point in edge.Tessellate():
                    t = point.DotProduct(u)
                    low = t if low is None else min(low, t)
                    high = t if high is None else max(high, t)
    return (low, high) if low is not None else None


def extent_along(element, u):
    """(centro, minimo, massimo, minimo senza isolamento, massimo senza isolamento) lungo u, in piedi. Il centro e'
    la proiezione del punto medio dell'asse. Isolamento di tubi e canali compreso."""
    center = midpoint(element).DotProduct(u)
    thickness = insulation_thickness(element)
    radius = round_radius(element)
    if radius is not None:
        bare = (center - radius, center + radius)
    else:
        bare = solid_extent(element, u)
        if bare is None:
            # ripiego: la dimensione maggiore della sezione, prudente
            half = max(param_double(element, bip) or 0.0 for bip in (
                DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM, DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM,
                DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)) / 2.0
            if half <= 0:
                raise SpacingError("The size of the section could not be read")
            bare = (center - half, center + half)
    return center, bare[0] - thickness, bare[1] + thickness, bare[0], bare[1]


def offset_axis(view, directions, mids):
    """Direzione u lungo cui si affiancano gli elementi, nel piano della vista e perpendicolare agli elementi.
    Orientata verso destra della vista (o verso l'alto se e' verticale)."""
    if not isinstance(view, (DB.ViewPlan, DB.ViewSection)):
        raise SpacingError("Run the command in a plan, section or elevation view")
    normal = view.ViewDirection
    mean = DB.XYZ.Zero
    for d in directions:
        mean = mean.Add(d)
    mean = unit(mean)
    u = unit(normal.CrossProduct(mean))
    if normal.CrossProduct(mean).GetLength() < 0.2:
        # elementi quasi perpendicolari alla vista (visti come punti): u e' la direzione dei punti piu' lontani
        flat = [m.Subtract(normal.Multiply(m.DotProduct(normal))) for m in mids]
        best, best_pair = 0.0, None
        for i in range(len(flat)):
            for j in range(i + 1, len(flat)):
                distance = flat[i].DistanceTo(flat[j])
                if distance > best:
                    best, best_pair = distance, (i, j)
        if best_pair is None or best < TOL_FT:
            raise SpacingError("The elements are seen as points and are all at the same place: there is no "
                               "direction to space them along")
        u = unit(flat[best_pair[1]].Subtract(flat[best_pair[0]]))
    if abs(u.DotProduct(view.RightDirection)) > 0.2:
        return u if u.DotProduct(view.RightDirection) > 0 else u.Negate()
    return u if u.DotProduct(view.UpDirection) >= 0 else u.Negate()


def check_parallel(elements):
    """Divide gli elementi in quelli entro MAX_ANGLE_DEG dalla direzione media e quelli fuori."""
    reference = elements[0].Location.Curve.Direction
    aligned = []
    for e in elements:
        d = e.Location.Curve.Direction
        aligned.append(d if d.DotProduct(reference) >= 0 else d.Negate())
    mean = DB.XYZ.Zero
    for d in aligned:
        mean = mean.Add(d)
    mean = unit(mean)
    limit = math.cos(math.radians(MAX_ANGLE_DEG))
    ok, out = [], []
    for e, d in zip(elements, aligned):
        angle = math.degrees(math.acos(max(-1.0, min(1.0, d.DotProduct(mean)))))
        (ok if d.DotProduct(mean) >= limit else out).append((e, d, angle))
    return ok, out


def build_groups(items, u, measure):
    """Raggruppa i tratti della stessa linea (centri vicini) e ne ordina le righe lungo u. Ritorna una lista di
    dict con elementi, centro, minimo e massimo (con isolamento se measure e' "faces")."""
    rows = []
    for element, direction in items:
        center, low, high, bare_low, bare_high = extent_along(element, u)
        if measure == "centerlines":
            low, high = center, center
        rows.append({"element": element, "center": center, "low": low, "high": high,
                     "bare_low": bare_low, "bare_high": bare_high, "direction": direction})
    rows.sort(key=lambda r: r["center"])
    groups = []
    for row in rows:
        if groups and abs(row["center"] - groups[-1]["center"]) < SAME_LINE_FT:
            group = groups[-1]
            group["rows"].append(row)
            group["low"] = min(group["low"], row["low"])
            group["high"] = max(group["high"], row["high"])
            group["bare_low"] = min(group["bare_low"], row["bare_low"])
            group["bare_high"] = max(group["bare_high"], row["bare_high"])
        else:
            groups.append({"rows": [row], "center": row["center"], "low": row["low"], "high": row["high"],
                           "bare_low": row["bare_low"], "bare_high": row["bare_high"]})
    return groups


def plan_shifts(groups, gaps, anchor):
    """Spostamento lungo u di ogni gruppo (piedi), con il gruppo fisso scelto in anchor. gaps[i] e' la distanza
    voluta tra il gruppo i e il successivo."""
    n = len(groups)
    new_low = [g["low"] for g in groups]
    new_high = [g["high"] for g in groups]
    shifts = [0.0] * n

    def place(i, previous):
        """Posiziona il gruppo i subito dopo il precedente (previous) o prima (previous > i)."""
        if previous < i:
            target_low = new_high[previous] + gaps[previous]
            shifts[i] = target_low - groups[i]["low"]
        else:
            target_high = new_low[previous] - gaps[i]
            shifts[i] = target_high - groups[i]["high"]
        new_low[i] = groups[i]["low"] + shifts[i]
        new_high[i] = groups[i]["high"] + shifts[i]

    if anchor == "last":
        for i in range(n - 2, -1, -1):
            place(i, i + 1)
    else:
        for i in range(1, n):
            place(i, i - 1)
        if anchor == "centre":
            before = (groups[0]["low"] + groups[-1]["high"]) / 2.0
            after = (new_low[0] + new_high[-1]) / 2.0
            shifts = [s + (before - after) for s in shifts]
    return shifts


# --- accessori --------------------------------------------------------------------------------------------------

def id_value(element_id):
    if hasattr(element_id, "Value"):
        return element_id.Value          # Revit 2026+
    return element_id.IntegerValue       # Revit <= 2025


ACCESSORY_CATEGORY_IDS = [int(c) for c in ACCESSORY_CATEGORIES]


def connected_accessories(element):
    """Accessori (valvole, serrande) collegati direttamente all'elemento."""
    found = {}
    for conn in element.ConnectorManager.Connectors:
        for ref in conn.AllRefs:
            owner = ref.Owner
            if owner.Id == element.Id or owner.Category is None:
                continue
            if id_value(owner.Category.Id) in ACCESSORY_CATEGORY_IDS:
                found[id_value(owner.Id)] = owner
    return list(found.values())


def accessory_position(accessory):
    location = accessory.Location
    if isinstance(location, DB.LocationPoint):
        return location.Point
    box = accessory.get_BoundingBox(None)
    return (box.Min.Add(box.Max)).Multiply(0.5) if box is not None else None


def clash_check(accessories, host_by_accessory, members, insulation_ids_of):
    """Dentro la transazione, a spostamento fatto. Per ogni accessorio cerca le intersezioni con gli altri elementi
    spostati, il loro isolamento e gli altri accessori. Ritorna [(accessorio, elemento in conflitto)]."""
    clashes = []
    all_accessory_ids = set(id_value(a.Id) for a in accessories)
    for accessory in accessories:
        own_hosts = host_by_accessory[id_value(accessory.Id)]
        excluded = set(id_value(h.Id) for h in own_hosts)
        for h in own_hosts:
            excluded.update(id_value(i) for i in insulation_ids_of[id_value(h.Id)])
        candidates = [m.Id for m in members if id_value(m.Id) not in excluded]
        for m in members:
            if id_value(m.Id) not in excluded:
                candidates.extend(insulation_ids_of[id_value(m.Id)])
        candidates.extend(a.Id for a in accessories if a.Id != accessory.Id)  # filtrati sotto per linea
        if not candidates:
            continue
        try:
            collector = DB.FilteredElementCollector(doc, List[DB.ElementId](candidates))
            hits = list(collector.WherePasses(DB.ElementIntersectsElementFilter(accessory)))
        except Exception:
            continue
        for hit in hits:
            if id_value(hit.Id) in all_accessory_ids and hit_is_same_line(accessory, hit, host_by_accessory):
                continue
            clashes.append((accessory, hit))
    return clashes


def hit_is_same_line(accessory, other, host_by_accessory):
    """Due accessori attaccati agli stessi tratti non vanno confrontati tra loro."""
    a = set(id_value(h.Id) for h in host_by_accessory[id_value(accessory.Id)])
    b = set(id_value(h.Id) for h in host_by_accessory.get(id_value(other.Id), []))
    return bool(a & b)


# --- comando ----------------------------------------------------------------------------------------------------

def analyse(elements, measure, label=""):
    """Elementi -> (linee ordinate lungo u, u, elementi esclusi perche' non paralleli). label ("Reference series")
    serve solo ai messaggi."""
    prefix = label + ": " if label else ""
    ok, out = check_parallel(elements)
    if len(ok) < 2:
        raise SpacingError("{}fewer than two elements are parallel within {:g} degrees".format(prefix, MAX_ANGLE_DEG))
    items = [(e, d) for e, d, _ in ok]
    u = offset_axis(doc.ActiveView, [d for _, d in items], [midpoint(e) for e, _ in items])
    groups = build_groups(items, u, measure)
    if len(groups) < 2:
        raise SpacingError("{}all the elements are on the same line: there is nothing to space".format(prefix))
    return groups, u, out


def copied_gaps(reference, groups, measure):
    """Distanze tra le linee consecutive della serie di riferimento (piedi). La serie da modificare deve avere lo
    stesso numero di linee."""
    ref_groups, _, ref_excluded = analyse(reference, measure, "Reference series")
    if len(ref_groups) != len(groups):
        raise SpacingError("The two series do not match: the reference has {} lines, the other one {}".format(
            len(ref_groups), len(groups)))
    return [ref_groups[i + 1]["low"] - ref_groups[i]["high"] for i in range(len(ref_groups) - 1)], ref_excluded


def run(elements, settings, reference=None):
    """Calcola, sposta, controlla. Con reference (altra serie) le distanze sono quelle della serie di riferimento,
    altrimenti quella delle impostazioni. Ritorna il dizionario del risultato per il report."""
    result = {"excluded": [], "ref_excluded": [], "groups": [], "clashes": [], "accessory_count": 0, "notes": [],
              "error": None, "achieved": [], "moved_accessories": 0}
    groups, u, result["excluded"] = analyse(elements, settings["measure"], "Series to change" if reference else "")
    if reference:
        gaps, result["ref_excluded"] = copied_gaps(reference, groups, settings["measure"])
    else:
        gaps = [settings["gap_mm"] / MM_PER_FT] * (len(groups) - 1)
    shifts = plan_shifts(groups, gaps, settings["anchor"])

    members = [row["element"] for g in groups for row in g["rows"]]
    insulation_ids = {id_value(m.Id): [i.Id for i in insulation_of(m)] for m in members}
    accessories, host_by_accessory, group_of_accessory = [], {}, {}
    for gi, g in enumerate(groups):
        for row in g["rows"]:
            for accessory in connected_accessories(row["element"]):
                key = id_value(accessory.Id)
                if key not in host_by_accessory:
                    accessories.append(accessory)
                    host_by_accessory[key] = []
                    group_of_accessory[key] = gi
                host_by_accessory[key].append(row["element"])
    result["accessory_count"] = len(accessories)
    before_positions = {id_value(a.Id): accessory_position(a) for a in accessories}

    transaction = DB.Transaction(doc, "Spacing")
    errors = RevitErrors()
    options = transaction.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(errors)
    options.SetClearAfterRollback(True)
    transaction.SetFailureHandlingOptions(options)
    transaction.Start()
    try:
        for g, shift in zip(groups, shifts):
            if abs(shift) <= TOL_FT:
                continue
            vector = u.Multiply(shift)
            for row in g["rows"]:
                row["element"].Location.Move(vector)
        doc.Regenerate()
        # gli accessori seguono il tratto a cui sono collegati? Se Revit non li ha spostati, si spostano qui
        for accessory in accessories:
            key = id_value(accessory.Id)
            shift = shifts[group_of_accessory[key]]
            before, after = before_positions[key], accessory_position(accessory)
            if before is None or after is None or abs(shift) <= TOL_FT:
                continue
            followed = after.Subtract(before).DotProduct(u)
            if abs(followed) < abs(shift) / 2.0:
                DB.ElementTransformUtils.MoveElement(doc, accessory.Id, u.Multiply(shift - followed))
                result["moved_accessories"] += 1
        doc.Regenerate()
        result["clashes"] = clash_check(accessories, host_by_accessory, members, insulation_ids)
        if transaction.Commit() != DB.TransactionStatus.Committed:
            raise SpacingError("Revit refused the result: {}".format(
                "; ".join(errors.messages) or "the transaction was rolled back"))
    except SpacingError:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    except Exception as ex:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise SpacingError("The elements could not be moved: {}".format(
            "; ".join(errors.messages) or ex))

    # distanze ottenute, rilette sulla geometria finale
    after_groups = build_groups([(row["element"], row["direction"]) for g in groups for row in g["rows"]], u,
                                settings["measure"])
    for i in range(len(after_groups) - 1):
        achieved = after_groups[i + 1]["low"] - after_groups[i]["high"]
        before = groups[i + 1]["low"] - groups[i]["high"]
        result["achieved"].append((before, achieved, gaps[i], abs(achieved - gaps[i]) > ACHIEVED_TOL_FT))
    result["groups"] = [(g, s) for g, s in zip(groups, shifts)]
    return result


def pick_series(prompt):
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, CurveFilter(), prompt)
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def main():
    copy_mode = __shiftclick__  # noqa: F821
    settings = load_settings()
    reference = None
    if copy_mode:
        # SHIFT + CLICK: copia le distanze da una serie di riferimento a un'altra; misura e linea ferma sono quelle
        # salvate con il CLICK
        reference = pick_series("Copy spacing, step 1 of 2: select the REFERENCE series (the one to copy the "
                                "distances from), then press Finish")
        elements = pick_series("Copy spacing, step 2 of 2: select the series to CHANGE, then press Finish")
        if set(e.Id for e in elements) & set(e.Id for e in reference):
            forms.alert("The two series share some elements. Select two separate series.", title="Spacing")
            script.exit()
    else:
        elements = get_selection()
    if len(elements) < 2:
        forms.alert("Select at least two elements.", title="Spacing")
        script.exit()
    if not copy_mode:
        import Spacing_ui
        xaml_path = script.get_bundle_file(XAML_FILE_NAME) or op.join(op.dirname(__file__), XAML_FILE_NAME)
        picked = Spacing_ui.SpacingWindow(xaml_path, settings).show()
        if picked is None:
            script.exit()
        settings = picked
        save_settings(settings)
    try:
        result = run(elements, settings, reference)
    except SpacingError as ex:
        if want_report(1, 1, "groups spaced"):
            output.print_md("# Spacing report")
            output.print_md("**Unchanged.** {}".format(ex))
        else:
            forms.alert("Unchanged. {}".format(ex), title="Spacing")
        return

    clashes = result["clashes"]
    wrong = len([a for a in result["achieved"] if a[3]])
    problems = len(clashes) + wrong + len(result["excluded"]) + len(result["ref_excluded"])
    if clashes:
        try:
            uidoc.Selection.SetElementIds(List[DB.ElementId](list(set(a.Id for a, _ in clashes))))
        except Exception:
            pass
    if not want_report(len(result["groups"]), problems, "lines spaced"):
        return
    output.print_md("# Spacing report")
    measure_text = ("between outer faces, insulation included" if settings["measure"] == "faces"
                    else "between centerlines")
    distance_text = "copied from the reference series" if reference else "**{:g} mm**".format(settings["gap_mm"])
    output.print_md("Distance: {} ({}) | {} | lines: **{}** | elements: **{}**".format(
        distance_text, measure_text, ANCHOR_LABELS[settings["anchor"]], len(result["groups"]),
        sum(len(g["rows"]) for g, _ in result["groups"])))
    for i, (group, shift) in enumerate(result["groups"], 1):
        names = ", ".join("{} {}".format(category_name(r["element"]), output.linkify(r["element"].Id))
                          for r in group["rows"])
        output.print_md("**Line {}** {} | moved **{:+.0f} mm**".format(i, names, mm(shift)))
    for i, (before, achieved, wanted, bad) in enumerate(result["achieved"], 1):
        output.print_md("Between lines {} and {}: {:.0f} mm -> **{:.0f} mm**{}".format(
            i, i + 1, mm(before), mm(achieved),
            " CHECK: not the requested {:.0f} mm, probably moved by a connected element".format(mm(wanted))
            if bad else ""))
    if result["accessory_count"]:
        output.print_md("Accessories checked: **{}** | clashes: **{}**{}".format(
            result["accessory_count"], len(clashes),
            " | moved with their run by the tool: {}".format(result["moved_accessories"])
            if result["moved_accessories"] else ""))
    for accessory, hit in clashes:
        output.print_md("**CLASH** {} {} with {} {}".format(
            category_name(accessory), output.linkify(accessory.Id), category_name(hit), output.linkify(hit.Id)))
    if clashes:
        output.print_md("The clashing accessories are selected. The elements were moved anyway: change the "
                        "distance or the accessories and run it again (Undo restores everything).")
    for element, _, angle in result["excluded"] + result["ref_excluded"]:
        output.print_md("Left out: {} {} is {:.0f} degrees from the mean direction (limit {:g})".format(
            category_name(element), output.linkify(element.Id), angle, MAX_ANGLE_DEG))


main()
