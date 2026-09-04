# -*- coding: utf-8 -*-
__title__ = "Elements\nin Room"

__doc__ = """Version = 1.0
Date    = 04.09.2026
_____________________________________________________________________
Scrive su un parametro testuale degli elementi il valore di un
parametro letto dalla room che li contiene, per ottenere un
raggruppamento basato sulla collocazione spaziale reale.

L'appartenenza non e' calcolata con bounding box ne' con intersezioni
fra solidi: sono le API di Revit a rispondere, sul volume computato.

  a) porte e finestre  -> From Room + To Room (fino a due room)
  b) family instance   -> Room (usa il Room Calculation Point se la
                          famiglia ce l'ha, altrimenti la Location)
  c) muri, separatori  -> le room di cui l'elemento e' delimitazione
  d) tutto il resto    -> la room che contiene il punto dell'elemento

Se un elemento appartiene a piu' room, i valori vengono concatenati
con il separatore indicato.

CLICK: analizza e scrive.
SHIFT + CLICK: solo anteprima, il modello non viene modificato.
_____________________________________________________________________
Author(s): Claude + Antonio Miano
"""

__author__ = "Claude + Antonio Miano"

from collections import OrderedDict

from System.Collections.Generic import List

from pyrevit import revit, script, DB, forms

from rpw.ui.forms import CheckBox, FlexForm, Label, Separator, Button, ComboBox, TextBox

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

BIC = DB.BuiltInCategory
BIP = DB.BuiltInParameter
ST = DB.StorageType

MAX_SKIPPED_ROWS = 200
SAMPLE_PER_CATEGORY = 10
FORM_WIDTH = 340
# Rialzo usato nel secondo tentativo del metodo (d): un metro sopra il livello.
LEVEL_RETRY_OFFSET_FT = 3.28084

# Stati del piano.
DA_SCRIVERE = "DA SCRIVERE"
SCRITTO = "SCRITTO"
GIA_UGUALE = "GIA' CORRETTO"
GIA_COMPILATO = "GIA' COMPILATO"
SENZA_ROOM = "NESSUNA ROOM"
NO_PARAM = "PARAM NON SCRIVIBILE"
ERRORE = "ERRORE"

METHOD_LABELS = {
    "a": "a - From/To Room",
    "b": "b - Room (calc. point)",
    "c": "c - delimitazione",
    "d": "d - punto nel volume",
    "d+": "d - punto a quota livello",
}


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


def as_text(param):
    """Valore del parametro come testo, qualunque sia lo StorageType."""
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
            referenced = doc.GetElement(param.AsElementId())
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
    return "(senza categoria)"


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


def room_at_point(point, run_phase):
    if point is None:
        return None
    try:
        return doc.GetRoomAtPoint(point, run_phase)
    except Exception:
        return None


# ---------------------------------------------------------------- guard iniziali

if doc.IsFamilyDocument:
    forms.alert("Comando non disponibile in un documento di famiglia.", exitscript=True)

try:
    volumes_on = DB.AreaVolumeSettings.GetAreaVolumeSettings(doc).ComputeVolumes
except Exception:
    volumes_on = False

if not volumes_on:
    forms.alert(
        "I volumi delle room non sono calcolati in questo modello.\n\n"
        "Senza volumi l'appartenenza verticale non e' attendibile.\n"
        "Attiva Area and Volume Computations -> Areas and Volumes, poi rilancia.",
        title="ElementsInRoom", exitscript=True)


# ---------------------------------------------------------------- fase

phases = [item for item in doc.Phases]
if not phases:
    forms.alert("Il modello non ha fasi.", exitscript=True)

default_phase = phases[-1]
try:
    view_phase_param = doc.ActiveView.get_Parameter(BIP.VIEW_PHASE)
    if view_phase_param is not None:
        view_phase = doc.GetElement(view_phase_param.AsElementId())
        if view_phase is not None:
            default_phase = view_phase
except Exception:
    pass

phase_dict = OrderedDict()
for item in phases:
    phase_dict[item.Name] = item

phase_form = FlexForm("Elements in Room  -  fase", [
    Label("Fase delle room da usare"),
    ComboBox("phase", phase_dict, default=default_phase.Name, sort=False,
             Width=FORM_WIDTH),
    Separator(),
    Button("OK"),
])
phase_form.show()
if not phase_form.values:
    script.exit()
phase = phase_form.values["phase"]


# ---------------------------------------------------------------- room della fase

phase_id = element_id_value(phase.Id)

all_rooms = DB.FilteredElementCollector(doc)\
    .OfCategory(BIC.OST_Rooms)\
    .WhereElementIsNotElementType()\
    .ToElements()

rooms = []
for room in all_rooms:
    try:
        if room.Area <= 0 or room.Location is None:
            continue
        room_phase = room.get_Parameter(BIP.ROOM_PHASE)
        if room_phase is not None and element_id_value(room_phase.AsElementId()) != phase_id:
            continue
    except Exception:
        continue
    rooms.append(room)

if not rooms:
    forms.alert("Nessuna room posizionata nella fase '{}'.".format(phase.Name),
                title="ElementsInRoom", exitscript=True)


# ---------------------------------------------------------------- categorie

model_categories = []
for category in doc.Settings.Categories:
    try:
        if category.CategoryType != DB.CategoryType.Model:
            continue
        if element_id_value(category.Id) == int(BIC.OST_Rooms):
            continue
        if not category.AllowsBoundParameters:
            continue
    except Exception:
        continue
    model_categories.append(category)
model_categories.sort(key=lambda c: c.Name)

selected_categories = forms.SelectFromList.show(
    model_categories, multiselect=True, name_attr="Name",
    title="Categorie da elaborare", button_name="Continua")
if not selected_categories:
    script.exit()

category_ids = List[DB.ElementId]()
for category in selected_categories:
    category_ids.Add(category.Id)

elements = DB.FilteredElementCollector(doc)\
    .WherePasses(DB.ElementMulticategoryFilter(category_ids))\
    .WhereElementIsNotElementType()\
    .ToElements()

if not elements:
    forms.alert("Nessun elemento nelle categorie selezionate.",
                title="ElementsInRoom", exitscript=True)


# ---------------------------------------------------------------- form parametri

def room_param_names(sample):
    names = set()
    for room in sample:
        for param in room.Parameters:
            try:
                names.add(param.Definition.Name)
            except Exception:
                pass
    return sorted(names)


def writable_text_param_names(sample_by_category):
    """Parametri istanza di testo scrivibili, unione su un campione per categoria.

    L'intersezione svuoterebbe la lista appena le categorie sono eterogenee: si fa
    l'unione e gli elementi privi del parametro finiscono nella tabella dei non
    scritti con il motivo esplicito.
    """
    names = set()
    for sample in sample_by_category.values():
        for element in sample:
            for param in element.Parameters:
                try:
                    if param.IsReadOnly or param.StorageType != ST.String:
                        continue
                    names.add(param.Definition.Name)
                except Exception:
                    pass
    return sorted(names)


samples = OrderedDict()
for element in elements:
    bucket = samples.setdefault(category_name(element), [])
    if len(bucket) < SAMPLE_PER_CATEGORY:
        bucket.append(element)

source_names = room_param_names(rooms[:SAMPLE_PER_CATEGORY])
target_names = writable_text_param_names(samples)

if not source_names:
    forms.alert("Nessun parametro leggibile sulle room.", exitscript=True)
if not target_names:
    forms.alert("Nessun parametro istanza di testo scrivibile sulle categorie scelte.\n"
                "Serve un parametro di progetto o condiviso di tipo Testo.",
                title="ElementsInRoom", exitscript=True)

default_source = "Name" if "Name" in source_names else source_names[0]

main_form = FlexForm("Elements in Room  -  {} elementi".format(len(elements)), [
    Label("Parametro della room da leggere"),
    ComboBox("source", source_names, default=default_source, Width=FORM_WIDTH),
    Separator(),
    Label("Parametro degli oggetti su cui scrivere"),
    ComboBox("target", target_names, default=target_names[0], Width=FORM_WIDTH),
    Separator(),
    Label("Separatore per gli elementi appartenenti a piu' room"),
    TextBox("sep", Text=";", Width=FORM_WIDTH),
    Separator(),
    CheckBox("overwrite", "Sovrascrivi i valori gia' compilati", default=True),
    CheckBox("retry", "Se il punto non cade in nessuna room, riprova a quota livello",
             default=True),
    Separator(),
    Button("OK"),
])
main_form.show()
if not main_form.values:
    script.exit()

source_name = main_form.values["source"]
target_name = main_form.values["target"]
separator = main_form.values["sep"] or ";"
overwrite = main_form.values["overwrite"]
retry_at_level = main_form.values["retry"]


# ---------------------------------------------------------------- indice room

room_value = {}
bound_rooms = {}

boundary_options = DB.SpatialElementBoundaryOptions()

for room in rooms:
    key = element_id_value(room.Id)
    room_value[key] = as_text(room.LookupParameter(source_name))
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

def resolve_rooms(element):
    """(metodo, [room_id]) - il primo metodo che risponde vince."""
    if isinstance(element, DB.FamilyInstance):
        found = []
        for getter_name in ("get_FromRoom", "get_ToRoom"):
            getter = getattr(element, getter_name, None)
            if getter is None:
                continue
            try:
                room = getter(phase)
            except Exception:
                room = None
            if room is not None:
                key = element_id_value(room.Id)
                if key in room_value and key not in found:
                    found.append(key)
        if found:
            return "a", found
        try:
            room = element.get_Room(phase)
        except Exception:
            room = None
        if room is not None:
            key = element_id_value(room.Id)
            if key in room_value:
                return "b", [key]

    hosted = bound_rooms.get(element_id_value(element.Id))
    if hosted:
        return "c", list(hosted)

    point = element_point(element)
    room = room_at_point(point, phase)
    if room is not None:
        key = element_id_value(room.Id)
        if key in room_value:
            return "d", [key]

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

with forms.ProgressBar(title="Analisi ({value} di {max_value})", cancellable=True) as pb:
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
            item["status"] = ERRORE
            item["reason"] = u"errore in analisi: {}".format(error)
            continue

        if not room_keys:
            item["status"] = SENZA_ROOM
            item["reason"] = u"nessuna room trovata con i metodi a/b/c/d"
            continue

        item["method"] = method
        item["rooms"] = room_keys
        new_value = join_values([room_value.get(key) for key in room_keys], separator)
        item["value"] = new_value

        if not new_value:
            item["status"] = SENZA_ROOM
            item["reason"] = u"room trovata ma '{}' e' vuoto".format(source_name)
            continue

        target = element.LookupParameter(target_name)
        if target is None:
            item["status"] = NO_PARAM
            item["reason"] = u"parametro '{}' non presente".format(target_name)
            continue
        if target.IsReadOnly:
            item["status"] = NO_PARAM
            item["reason"] = u"parametro '{}' in sola lettura".format(target_name)
            continue
        if target.StorageType != ST.String:
            item["status"] = NO_PARAM
            item["reason"] = u"parametro '{}' non e' di tipo Testo".format(target_name)
            continue

        current = target.AsString()
        if current == new_value:
            item["status"] = GIA_UGUALE
            continue
        if current and current.strip() and not overwrite:
            item["status"] = GIA_COMPILATO
            item["reason"] = u"valore esistente '{}' mantenuto".format(current)
            continue

        item["status"] = DA_SCRIVERE


# ---------------------------------------------------------------- scrittura

preview_only = bool(__shiftclick__)  # noqa: F821

if not preview_only:
    with revit.Transaction("ElementsInRoom"):
        for item in plan:
            if item["status"] != DA_SCRIVERE:
                continue
            try:
                target = item["element"].LookupParameter(target_name)
                if target is None or target.IsReadOnly:
                    item["status"] = NO_PARAM
                    item["reason"] = u"parametro non piu' scrivibile"
                    continue
                target.Set(item["value"])
                item["status"] = SCRITTO
            except Exception as error:
                item["status"] = ERRORE
                item["reason"] = u"errore Revit: {}".format(error)


# ---------------------------------------------------------------- report

output.close_others()
output.print_md("# Elements in Room")
output.print_md(
    "- Modalita': **{}**\n"
    "- Fase: **{}**\n"
    "- Categorie: **{}**\n"
    "- Room nella fase: **{}** (senza valore in '{}': {})\n"
    "- Parametro room -> oggetti: **{}** -> **{}**\n"
    "- Separatore: `{}` | sovrascrivi: **{}** | riprova a quota livello: **{}**".format(
        "ANTEPRIMA (nessuna modifica)" if preview_only else "scrittura",
        phase.Name,
        ", ".join(sorted(set(c.Name for c in selected_categories))),
        len(rooms), source_name, rooms_without_value,
        source_name, target_name,
        separator, "si" if overwrite else "no", "si" if retry_at_level else "no"))

STATUS_ORDER = [DA_SCRIVERE, SCRITTO, GIA_UGUALE, GIA_COMPILATO,
                SENZA_ROOM, NO_PARAM, ERRORE]

per_category = OrderedDict()
multi_room = []
skipped = []

for item in plan:
    counts = per_category.setdefault(item["category"], OrderedDict())
    counts[item["status"]] = counts.get(item["status"], 0) + 1
    if len(item["rooms"]) > 1:
        multi_room.append(item)
    if item["status"] in (SENZA_ROOM, NO_PARAM, ERRORE, GIA_COMPILATO):
        skipped.append(item)

rows = []
for name in sorted(per_category):
    counts = per_category[name]
    rows.append([name] + [counts.get(status, 0) for status in STATUS_ORDER])
totals = ["**Totale**"]
for status in STATUS_ORDER:
    totals.append(sum(counts.get(status, 0) for counts in per_category.values()))
rows.append(totals)

output.print_md("## Esito per categoria")
output.print_table(table_data=rows, title="", columns=["Categoria"] + STATUS_ORDER)

method_counts = OrderedDict()
for item in plan:
    if item["method"]:
        method_counts[item["method"]] = method_counts.get(item["method"], 0) + 1
if method_counts:
    output.print_md("## Metodo di risoluzione")
    output.print_table(
        table_data=[[METHOD_LABELS.get(key, key), value]
                    for key, value in method_counts.items()],
        title="", columns=["Metodo", "Elementi"])

if multi_room:
    output.print_md("## Elementi in piu' room ({})".format(len(multi_room)))
    table = [[output.linkify(item["element"].Id), item["category"],
              METHOD_LABELS.get(item["method"], item["method"]),
              len(item["rooms"]), item["value"]]
             for item in multi_room[:MAX_SKIPPED_ROWS]]
    output.print_table(table_data=table, title="",
                       columns=["Elemento", "Categoria", "Metodo", "Room", "Valore"])
    if len(multi_room) > MAX_SKIPPED_ROWS:
        output.print_md("_...e altri {}._".format(len(multi_room) - MAX_SKIPPED_ROWS))

if skipped:
    output.print_md("## Non scritti ({})".format(len(skipped)))
    table = [[output.linkify(item["element"].Id), item["category"],
              item["status"], item["reason"]]
             for item in skipped[:MAX_SKIPPED_ROWS]]
    output.print_table(table_data=table, title="",
                       columns=["Elemento", "Categoria", "Stato", "Motivo"])
    if len(skipped) > MAX_SKIPPED_ROWS:
        output.print_md("_...e altri {}._".format(len(skipped) - MAX_SKIPPED_ROWS))

if preview_only:
    output.print_md("---\n**Anteprima**: nessun parametro e' stato modificato. "
                    "Rilancia con un CLICK normale per scrivere.")
