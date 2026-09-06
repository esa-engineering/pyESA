# -*- coding: utf-8 -*-
__title__ = "Elements\nin Room"

__doc__ = """Version = 3.0
Date    = 06.09.2026
_____________________________________________________________________
Writes on a text parameter of the elements the value of a parameter
read from the room that contains them, so that objects can be grouped
by their actual spatial location.

Membership is computed neither with bounding boxes nor with solid
intersections: the Revit API answers, on the computed volume. The
first method that answers wins:

  a)  doors and windows -> From Room + To Room (up to two rooms)
  a*) doors and windows -> the rooms on both sides of the opening
  b)  family instance   -> Room (uses the Room Calculation Point when
                           the family has one, otherwise the Location)
  c)  walls, separators -> the rooms the element bounds
  c*) walls, separators -> the rooms on both sides of the element
  d)  everything else   -> the room containing the element point
  e)  recovery          -> the point is probed again offset by the
                           given tolerances (X, Y, Z+, Z-), along the
                           global axes or the object's own axes
  d+) recovery          -> the point is probed again at the elevation
                           of the element's level

Rooms can come from the current document (CLICK) or from a loaded
Revit link (SHIFT + CLICK, the link instance is chosen first). With a
link the transform of the instance is applied to every probe, methods
a) and c) are unavailable (the API only knows the rooms of the current
document) and a*) / c*) take their place.

Rooms can be taken from the whole project or from the active view
only; elements too. Only elements that exist in the chosen element
phase are processed (demolished and not yet created ones are skipped).

When an element belongs to more than one room the values are joined
with the given separator.

CLICK: rooms from the current document.
SHIFT + CLICK: rooms from a linked model.
_____________________________________________________________________
Author(s): Claude + Antonio Miano
"""

__author__ = "Claude + Antonio Miano"

from collections import OrderedDict

from System.Collections.Generic import List

from pyrevit import revit, script, DB, forms

from elementsinroom_ui import show_config_form

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

BIC = DB.BuiltInCategory
BIP = DB.BuiltInParameter
ST = DB.StorageType

MAX_SKIPPED_ROWS = 200
# Rialzo usato nel secondo tentativo del metodo (d): un metro sopra il livello.
LEVEL_RETRY_OFFSET_FT = 3.28084
# Sonde bilaterali (a*, c*): margine oltre la faccia dell'elemento e sollevamento
# del punto, che per porte e muri sta esattamente sulla faccia inferiore della room.
PROBE_MARGIN_FT = 0.164042   # 5 cm
PROBE_LIFT_FT = 0.328084     # 10 cm

# Sonde del metodo (e): asse della tolleranza e verso lungo quell'asse.
TOLERANCE_PROBES = (("x", 1), ("x", -1), ("y", 1), ("y", -1),
                    ("zup", 1), ("zdown", -1))

# Stati del piano. TO_WRITE e' solo il marcatore interno prima della scrittura;
# gli altri finiscono nelle tabelle del report e sono in inglese.
TO_WRITE = "TO WRITE"
WRITTEN = "WRITTEN"
ALREADY_OK = "ALREADY CORRECT"
ALREADY_FILLED = "ALREADY FILLED"
NO_ROOM = "NO ROOM"
NO_PARAM = "PARAM NOT WRITABLE"
FAILED = "ERROR"

METHOD_LABELS = {
    "a": "a - From/To Room",
    "a*": "a* - two-sided probe (door/window)",
    "b": "b - Room (calc. point)",
    "c": "c - bounding element",
    "c*": "c* - two-sided probe (wall/separator)",
    "d": "d - point in volume",
    "e": "e - tolerance",
    "d+": "d+ - point at level elevation",
}

DOOR_WINDOW_IDS = (int(BIC.OST_Doors), int(BIC.OST_Windows))
SEPARATOR_ID = int(BIC.OST_RoomSeparationLines)


# ---------------------------------------------------------------- helpers

def element_id_value(eid):
    """ElementId.IntegerValue e' stato rimosso in Revit 2026 (sostituito da .Value)."""
    if eid is None:
        return -1
    if hasattr(eid, "Value"):
        return eid.Value
    return eid.IntegerValue


def level_internal_elevation(level):
    """Quota del livello nel sistema di coordinate INTERNO.

    Level.Elevation puo' essere riferita al Survey Point mentre tutta la geometria
    (punti di inserimento, bounding box) sta sempre in coordinate interne:
    confrontare i due sistemi esclude silenziosamente tutto.
    """
    if level is None:
        return None
    try:
        return level.ProjectElevation
    except Exception:
        pass
    try:
        return level.Elevation
    except Exception:
        return None


def as_text(param, document):
    """Valore del parametro come testo, qualunque sia lo StorageType.

    document: quello a cui appartiene l'elemento del parametro (il documento del
    link per le room di un link), serve a risolvere i valori ElementId.
    """
    if param is None:
        return None
    try:
        storage = param.StorageType
        if storage == ST.String:
            return param.AsString()
        value = param.AsValueString()
        if value:
            return value
        if storage == ST.Integer:
            return str(param.AsInteger())
        if storage == ST.Double:
            return str(param.AsDouble())
        if storage == ST.ElementId:
            referenced = document.GetElement(param.AsElementId())
            if referenced is not None:
                return referenced.Name
    except Exception:
        return None
    return None


def join_values(values, separator):
    """Concatena i valori togliendo vuoti e duplicati, mantenendo l'ordine."""
    kept = []
    for value in values:
        if value is None:
            continue
        text = value.strip() if hasattr(value, "strip") else str(value)
        if text and text not in kept:
            kept.append(text)
    return separator.join(kept)


def category_name(element):
    try:
        if element.Category is not None:
            return element.Category.Name
    except Exception:
        pass
    return "(no category)"


def category_id_value(element):
    try:
        if element.Category is not None:
            return element_id_value(element.Category.Id)
    except Exception:
        pass
    return None


def element_point(element):
    """Punto rappresentativo: Location, poi punto medio della curva, poi centro bbox."""
    try:
        location = element.Location
        if isinstance(location, DB.LocationPoint):
            return location.Point
        if isinstance(location, DB.LocationCurve):
            return location.Curve.Evaluate(0.5, True)
    except Exception:
        pass
    try:
        box = element.get_BoundingBox(None)
        if box is not None:
            return DB.XYZ((box.Min.X + box.Max.X) / 2.0,
                          (box.Min.Y + box.Max.Y) / 2.0,
                          (box.Min.Z + box.Max.Z) / 2.0)
    except Exception:
        pass
    return None


GLOBAL_AXES = (DB.XYZ.BasisX, DB.XYZ.BasisY, DB.XYZ.BasisZ)


def element_axes(element):
    """(BasisX, BasisY, BasisZ) propri dell'elemento, assi globali se non ne ha.

    Costo trascurabile: e' una lettura di Transform o della tangente della curva
    di posizione, non rigenera geometria.
    """
    try:
        if isinstance(element, DB.FamilyInstance):
            transform = element.GetTotalTransform()
            return transform.BasisX, transform.BasisY, transform.BasisZ
        location = element.Location
        if isinstance(location, DB.LocationCurve):
            tangent = location.Curve.ComputeDerivatives(0.5, True).BasisX.Normalize()
            side = tangent.CrossProduct(DB.XYZ.BasisZ)
            if side.GetLength() > 1e-9:
                return tangent, side.Normalize(), DB.XYZ.BasisZ
    except Exception:
        pass
    return GLOBAL_AXES


LEVEL_BIPS = ("FAMILY_LEVEL_PARAM", "SCHEDULE_LEVEL_PARAM",
              "INSTANCE_REFERENCE_LEVEL_PARAM", "RBS_START_LEVEL_PARAM",
              "LEVEL_PARAM", "WALL_BASE_CONSTRAINT")


def element_level(element):
    """Livello dell'elemento, None se non ne ha uno riconoscibile."""
    try:
        level_id = element.LevelId
    except Exception:
        level_id = None
    if level_id is None or element_id_value(level_id) < 0:
        level_id = None
        for bip_name in LEVEL_BIPS:
            member = getattr(BIP, bip_name, None)
            if member is None:
                continue
            try:
                param = element.get_Parameter(member)
            except Exception:
                param = None
            if param is not None and param.StorageType == ST.ElementId:
                candidate = param.AsElementId()
                if element_id_value(candidate) > 0:
                    level_id = candidate
                    break
    if level_id is None:
        return None
    found = doc.GetElement(level_id)
    return found if isinstance(found, DB.Level) else None


# ---------------------------------------------------------------- guard iniziali

if doc.IsFamilyDocument:
    forms.alert("This command is not available in a family document.", exitscript=True)


# ---------------------------------------------------------------- sorgente room

# SHIFT + CLICK: le room vengono da un link. Il punto di ogni sonda va portato nel
# sistema di coordinate del link con l'inversa della trasformazione dell'istanza;
# in modalita' normale la trasformazione e' l'identita' e il codice non si biforca.
link_mode = bool(__shiftclick__)  # noqa: F821
link_instance = None

if link_mode:
    links = [lk for lk in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)
             if lk.GetLinkDocument() is not None]
    if not links:
        forms.alert("No loaded Revit link in this model.",
                    title="ElementsInRoom", exitscript=True)
    links.sort(key=lambda lk: lk.Name)
    link_instance = forms.SelectFromList.show(
        links, name_attr="Name", multiselect=False,
        title="Rooms from link", button_name="Select")
    if link_instance is None:
        script.exit()

if link_instance is not None:
    room_doc = link_instance.GetLinkDocument()
    to_room_doc = link_instance.GetTotalTransform().Inverse
    room_source_label = u"link '{}'".format(link_instance.Name)
else:
    room_doc = doc
    to_room_doc = DB.Transform.Identity
    room_source_label = u"current document"

try:
    volumes_on = DB.AreaVolumeSettings.GetAreaVolumeSettings(room_doc).ComputeVolumes
except Exception:
    volumes_on = False

if not volumes_on:
    forms.alert(
        "Room volumes are not computed in the {}.\n\n"
        "Without volumes, vertical membership is not reliable.\n"
        "Turn on Area and Volume Computations -> Areas and Volumes in that model, "
        "then run again.".format(room_source_label),
        title="ElementsInRoom", exitscript=True)


def room_at_point(point, run_phase):
    """Room del documento delle room che contiene il punto (coordinate dell'host)."""
    if point is None:
        return None
    try:
        return room_doc.GetRoomAtPoint(to_room_doc.OfPoint(point), run_phase)
    except Exception:
        return None


# ---------------------------------------------------------------- configurazione

config = show_config_form(doc, room_doc, link_instance)
if config is None:
    script.exit()

phase = config.phase
element_phase = config.element_phase
rooms = config.rooms
selected_categories = config.categories
source_name = config.source_name
target_name = config.target_name
separator = config.separator or ";"
overwrite = config.overwrite
retry_at_level = config.retry_at_level
local_axes = config.local_axes
tolerances = config.tol_ft

category_ids = List[DB.ElementId]()
for category in selected_categories:
    category_ids.Add(category.Id)

# Solo gli elementi che esistono nella fase scelta: demoliti e non ancora creati
# non hanno una room a cui appartenere.
phase_statuses = List[DB.ElementOnPhaseStatus]()
phase_statuses.Add(DB.ElementOnPhaseStatus.New)
phase_statuses.Add(DB.ElementOnPhaseStatus.Existing)
phase_statuses.Add(DB.ElementOnPhaseStatus.Temporary)

if config.only_view_elements:
    collector = DB.FilteredElementCollector(doc, doc.ActiveView.Id)
else:
    collector = DB.FilteredElementCollector(doc)

elements = collector\
    .WherePasses(DB.ElementMulticategoryFilter(category_ids))\
    .WherePasses(DB.ElementPhaseStatusFilter(element_phase.Id, phase_statuses))\
    .WhereElementIsNotElementType()\
    .ToElements()

if not elements:
    forms.alert(
        "No element of the selected categories exists in phase '{}'{}.".format(
            element_phase.Name,
            " in the active view" if config.only_view_elements else ""),
        title="ElementsInRoom", exitscript=True)


# ---------------------------------------------------------------- indice room

room_value = {}
bound_rooms = {}

boundary_options = DB.SpatialElementBoundaryOptions()

for room in rooms:
    key = element_id_value(room.Id)
    room_value[key] = as_text(room.LookupParameter(source_name), room_doc)
    if link_mode:
        # I segmenti di contorno delle room di un link sono ElementId del documento
        # del link: confrontarli con gli Id dell'host darebbe falsi positivi.
        continue
    try:
        loops = room.GetBoundarySegments(boundary_options)
    except Exception:
        loops = None
    if not loops:
        continue
    for loop in loops:
        for segment in loop:
            bounding_id = element_id_value(segment.ElementId)
            if bounding_id < 0:
                continue
            hosts = bound_rooms.setdefault(bounding_id, [])
            if key not in hosts:
                hosts.append(key)

rooms_without_value = len([key for key, value in room_value.items() if not value])


# ---------------------------------------------------------------- risoluzione

def known_room_key(room, found):
    """Aggiunge a found la chiave della room se fa parte dell'indice."""
    if room is None:
        return
    key = element_id_value(room.Id)
    if key in room_value and key not in found:
        found.append(key)


def two_sided_probe(point, direction, offset):
    """Room ai due lati del punto, lungo direction, a distanza offset (piedi).

    Il punto viene sollevato di PROBE_LIFT_FT: porte e muri hanno il punto di
    posizione esattamente a quota livello, sulla faccia inferiore della room.
    """
    if point is None or direction is None:
        return []
    try:
        if direction.GetLength() < 1e-9:
            return []
        unit = direction.Normalize()
    except Exception:
        return []
    base = DB.XYZ(point.X, point.Y, point.Z + PROBE_LIFT_FT)
    found = []
    for sign in (1, -1):
        known_room_key(room_at_point(base + unit.Multiply(sign * offset), phase), found)
    return found


def door_window_probe(element):
    """(a*) sonde ai due lati dell'apertura, oltre le facce del muro ospite."""
    half_width = 0.0
    try:
        host = element.Host
        if isinstance(host, DB.Wall):
            half_width = host.Width / 2.0
    except Exception:
        pass
    try:
        direction = element.FacingOrientation
    except Exception:
        direction = None
    return two_sided_probe(element_point(element), direction, half_width + PROBE_MARGIN_FT)


def wall_probe(element):
    """(c*) sonde ai due lati del muro o del separatore, oltre le facce."""
    try:
        if not isinstance(element.Location, DB.LocationCurve):
            return []
    except Exception:
        return []
    half_width = 0.0
    if isinstance(element, DB.Wall):
        try:
            half_width = element.Width / 2.0
        except Exception:
            pass
    _, side, _ = element_axes(element)
    return two_sided_probe(element_point(element), side, half_width + PROBE_MARGIN_FT)


def resolve_rooms(element):
    """(metodo, [room_id]) - il primo metodo che risponde vince.

    Catena: a, a*, b, c, c*, d, e, d+. In modalita' link a) e c) non possono
    rispondere (l'API conosce solo le room del documento corrente) e si saltano.
    """
    if isinstance(element, DB.FamilyInstance):
        is_opening = category_id_value(element) in DOOR_WINDOW_IDS

        if not link_mode and is_opening:
            found = []
            for getter_name in ("get_FromRoom", "get_ToRoom"):
                getter = getattr(element, getter_name, None)
                if getter is None:
                    continue
                try:
                    room = getter(phase)
                except Exception:
                    room = None
                known_room_key(room, found)
            if found:
                return "a", found

        if is_opening:
            found = door_window_probe(element)
            if found:
                return "a*", found

        found = []
        if not link_mode:
            try:
                known_room_key(element.get_Room(phase), found)
            except Exception:
                pass
        else:
            # Equivalente manuale di get_Room: il Room Calculation Point della
            # famiglia, se c'e'. Senza, resta il punto di posizione del metodo (d).
            try:
                if element.HasSpatialElementCalculationPoint:
                    calc_point = element.GetSpatialElementCalculationPoint()
                    known_room_key(room_at_point(calc_point, phase), found)
            except Exception:
                pass
        if found:
            return "b", found

    if not link_mode:
        hosted = bound_rooms.get(element_id_value(element.Id))
        if hosted:
            return "c", list(hosted)

    if isinstance(element, DB.Wall) or category_id_value(element) == SEPARATOR_ID:
        found = wall_probe(element)
        if found:
            return "c*", found

    point = element_point(element)
    room = room_at_point(point, phase)
    if room is not None:
        key = element_id_value(room.Id)
        if key in room_value:
            return "d", [key]

    # (e) recupero: il punto e' fuori dal solido per pochi centimetri. Si sonda
    # spostato delle tolleranze indicate; tutte le room raggiunte si concatenano,
    # come gia' fanno (a) per From/To Room e (c) per i muri di delimitazione.
    if point is not None and any(value > 0 for value in tolerances.values()):
        if local_axes:
            basis_x, basis_y, basis_z = element_axes(element)
        else:
            basis_x, basis_y, basis_z = GLOBAL_AXES
        vectors = {"x": basis_x, "y": basis_y, "zup": basis_z, "zdown": basis_z}
        found = []
        for axis, sign in TOLERANCE_PROBES:
            distance = tolerances[axis]
            if distance <= 0:
                continue
            probe = point + vectors[axis].Multiply(sign * distance)
            known_room_key(room_at_point(probe, phase), found)
        if found:
            return "e", found

    # ponytail: un solo punto per elemento. Se servisse coprire tubi e canali che
    # attraversano piu' locali, qui si testano anche i due estremi della curva.
    if retry_at_level and point is not None:
        elevation = level_internal_elevation(element_level(element))
        if elevation is not None:
            raised = DB.XYZ(point.X, point.Y, elevation + LEVEL_RETRY_OFFSET_FT)
            room = room_at_point(raised, phase)
            if room is not None:
                key = element_id_value(room.Id)
                if key in room_value:
                    return "d+", [key]

    return None, []


# ---------------------------------------------------------------- piano (sola lettura)

plan = []

with forms.ProgressBar(title="Analysing ({value} of {max_value})", cancellable=True) as pb:
    total = len(elements)
    for index, element in enumerate(elements):
        if pb.cancelled:
            script.exit()
        pb.update_progress(index + 1, total)

        item = {"element": element, "category": category_name(element),
                "method": None, "rooms": [], "value": "",
                "status": None, "reason": u""}
        plan.append(item)

        try:
            method, room_keys = resolve_rooms(element)
        except Exception as error:
            item["status"] = FAILED
            item["reason"] = u"analysis error: {}".format(error)
            continue

        if not room_keys:
            item["status"] = NO_ROOM
            item["reason"] = u"no room found with methods a/a*/b/c/c*/d/e/d+"
            continue

        item["method"] = method
        item["rooms"] = room_keys
        new_value = join_values([room_value.get(key) for key in room_keys], separator)
        item["value"] = new_value

        if not new_value:
            item["status"] = NO_ROOM
            item["reason"] = u"room found but '{}' is empty".format(source_name)
            continue

        target = element.LookupParameter(target_name)
        if target is None:
            item["status"] = NO_PARAM
            item["reason"] = u"parameter '{}' not present".format(target_name)
            continue
        if target.IsReadOnly:
            item["status"] = NO_PARAM
            item["reason"] = u"parameter '{}' is read-only".format(target_name)
            continue
        if target.StorageType != ST.String:
            item["status"] = NO_PARAM
            item["reason"] = u"parameter '{}' is not of type Text".format(target_name)
            continue

        current = target.AsString()
        if current == new_value:
            item["status"] = ALREADY_OK
            continue
        if current and current.strip() and not overwrite:
            item["status"] = ALREADY_FILLED
            item["reason"] = u"existing value '{}' kept".format(current)
            continue

        item["status"] = TO_WRITE


# ---------------------------------------------------------------- scrittura

with revit.Transaction("ElementsInRoom"):
    for item in plan:
        if item["status"] != TO_WRITE:
            continue
        try:
            target = item["element"].LookupParameter(target_name)
            if target is None or target.IsReadOnly:
                item["status"] = NO_PARAM
                item["reason"] = u"parameter no longer writable"
                continue
            target.Set(item["value"])
            item["status"] = WRITTEN
        except Exception as error:
            item["status"] = FAILED
            item["reason"] = u"Revit error: {}".format(error)


# ---------------------------------------------------------------- report

output.close_others()
output.print_md("# Elements in Room")
tolerance_text = "off"
if config.has_tolerance:
    tolerance_text = "X {} | Y {} | Z+ {} | Z- {} cm, {} axes".format(
        config.tol_cm["x"], config.tol_cm["y"],
        config.tol_cm["zup"], config.tol_cm["zdown"],
        "object" if local_axes else "global")

output.print_md(
    u"- Rooms from: **{}**\n"
    u"- Room phase: **{}** | element phase: **{}**\n"
    u"- Categories: **{}**\n"
    u"- Rooms ({}): **{}** (with no value in '{}': {})\n"
    u"- Elements ({}): **{}**\n"
    u"- Room parameter -> elements: **{}** -> **{}**\n"
    u"- Separator: `{}` | overwrite: **{}** | retry at level elevation: **{}**\n"
    u"- Tolerance: **{}**".format(
        room_source_label,
        phase.Name, element_phase.Name,
        ", ".join(sorted(set(c.Name for c in selected_categories))),
        "active view only" if config.only_active_view else "whole project",
        len(rooms), source_name, rooms_without_value,
        "active view only" if config.only_view_elements else "whole project",
        len(elements),
        source_name, target_name,
        separator, "yes" if overwrite else "no", "yes" if retry_at_level else "no",
        tolerance_text))

STATUS_ORDER = [WRITTEN, ALREADY_OK, ALREADY_FILLED, NO_ROOM, NO_PARAM, FAILED]

per_category = OrderedDict()
multi_room = []
skipped = []

for item in plan:
    counts = per_category.setdefault(item["category"], OrderedDict())
    counts[item["status"]] = counts.get(item["status"], 0) + 1
    if len(item["rooms"]) > 1:
        multi_room.append(item)
    if item["status"] in (NO_ROOM, NO_PARAM, FAILED, ALREADY_FILLED):
        skipped.append(item)

rows = []
for name in sorted(per_category):
    counts = per_category[name]
    rows.append([name] + [counts.get(status, 0) for status in STATUS_ORDER])
totals = ["**Total**"]
for status in STATUS_ORDER:
    totals.append(sum(counts.get(status, 0) for counts in per_category.values()))
rows.append(totals)

output.print_md("## Result by category")
output.print_table(table_data=rows, title="", columns=["Category"] + STATUS_ORDER)

method_counts = OrderedDict()
for item in plan:
    if item["method"]:
        method_counts[item["method"]] = method_counts.get(item["method"], 0) + 1
if method_counts:
    output.print_md("## Resolution method")
    output.print_table(
        table_data=[[METHOD_LABELS.get(key, key), value]
                    for key, value in method_counts.items()],
        title="", columns=["Method", "Elements"])

if multi_room:
    output.print_md("## Elements in more than one room ({})".format(len(multi_room)))
    table = [[output.linkify(item["element"].Id), item["category"],
              METHOD_LABELS.get(item["method"], item["method"]),
              len(item["rooms"]), item["value"]]
             for item in multi_room[:MAX_SKIPPED_ROWS]]
    output.print_table(table_data=table, title="",
                       columns=["Element", "Category", "Method", "Rooms", "Value"])
    if len(multi_room) > MAX_SKIPPED_ROWS:
        output.print_md("_...and {} more._".format(len(multi_room) - MAX_SKIPPED_ROWS))

if skipped:
    output.print_md("## Not written ({})".format(len(skipped)))
    table = [[output.linkify(item["element"].Id), item["category"],
              item["status"], item["reason"]]
             for item in skipped[:MAX_SKIPPED_ROWS]]
    output.print_table(table_data=table, title="",
                       columns=["Element", "Category", "Status", "Reason"])
    if len(skipped) > MAX_SKIPPED_ROWS:
        output.print_md("_...and {} more._".format(len(skipped) - MAX_SKIPPED_ROWS))
