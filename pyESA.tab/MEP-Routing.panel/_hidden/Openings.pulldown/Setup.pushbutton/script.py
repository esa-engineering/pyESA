# -*- coding: utf-8 -*-
__title__ = "Setup"
__doc__ = """Version = 0.1
Date    = 09.10.2026
________________________________________________________________
Description:
Settings of the openings for this project, saved in the model (the same
for everyone who opens it):
- families: simple mode (one rectangular and one round family, for
  walls and for floors) or matrix (a family for each MEP type and case:
  wall / floor, normal / REI / structural, plus merged openings)
- names of the size parameters (height, width, depth, diameter) and a
  check of which ones each family has as instance parameters
- REI rules: one or more parameters to read on the crossed elements,
  with a test that counts the matching elements in each model
- sizing rules by class: clearance, rounding, merge distance, minimum
  size, extra depth
- MEP categories to check, parameters to fill, what to do with the
  openings no longer needed, workset, schedule, view filters

Export / Import save and load the setup as a file (without the list of
Links), to reuse it in another project.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import clr
import io
import json
clr.AddReference("RevitAPI")
import System
from System.Collections.Generic import List
from Autodesk.Revit.DB.ExtensibleStorage import Schema, SchemaBuilder, Entity, AccessLevel

from pyrevit import revit, script, DB, forms

import Setup_ui

doc = revit.doc
output = script.get_output()

XAML_FILE_NAME = "Setup_form.xaml"

# =================================================================================================================
# Configurazione di progetto: blocco identico in Links, Setup, Place e Refresh (correggere in tutti e quattro)
# =================================================================================================================
CONFIG_SCHEMA_GUID = "5d0a2f6e-8c41-4b7a-9e3d-2a6f1c8b4e01"
OPENING_SCHEMA_GUID = "a3e7c9d2-4f18-4c6b-b5e0-7d2f9a1c3e52"
JSON_FIELD = "Json"
HOST_MODEL = "Host model"

MEP_KINDS = ("pipe", "duct_round", "duct_rect", "tray", "conduit", "merged")
MEP_KIND_LABELS = {"pipe": "Pipe", "duct_round": "Round duct", "duct_rect": "Rectangular duct",
                   "tray": "Cable tray", "conduit": "Conduit", "merged": "Merged opening"}
CASES = ("h_normal", "h_rei", "h_struct", "v_normal", "v_rei", "v_struct")
CASE_LABELS = {"h_normal": "Wall normal", "h_rei": "Wall REI", "h_struct": "Wall structural",
               "v_normal": "Floor normal", "v_rei": "Floor REI", "v_struct": "Floor structural"}
CLASS_LABELS = {"normal": "Normal", "rei": "REI", "struct": "Structural"}
PARAM_KEYS = ("level", "systems", "mep_categories", "mep_count", "host_model", "host_category", "host_type",
              "crossing_class", "status", "updated")
PARAM_LABELS = {"level": "Level", "systems": "Systems", "mep_categories": "MEP categories",
                "mep_count": "Number of MEP elements", "host_model": "Crossed model",
                "host_category": "Crossed category", "host_type": "Crossed type",
                "crossing_class": "Class (Normal / REI / Structural)", "status": "Status",
                "updated": "Last update (date)"}

DEFAULT_CONFIG = {
    "version": 1,
    "rev": 0,
    "mep": {"pipe": True, "duct": True, "tray": True, "conduit": True},
    "mode": "simple",
    "simple": {"h_rect": "", "h_round": "", "v_rect": "", "v_round": ""},
    "matrix": {},
    "dims": {"width": "", "height": "", "depth": "", "diameter": ""},
    "rei_rules": [],
    "rules": {
        "normal": {"clearance_mm": 25.0, "rounding_mm": 50.0, "merge_mm": 100.0, "min_mm": 0.0, "extra_depth_mm": 0.0},
        "rei": {"clearance_mm": 25.0, "rounding_mm": 10.0, "merge_mm": 0.0, "min_mm": 0.0, "extra_depth_mm": 0.0},
        "struct": {"clearance_mm": 50.0, "rounding_mm": 50.0, "merge_mm": 0.0, "min_mm": 0.0, "extra_depth_mm": 0.0},
    },
    "params": dict((k, "") for k in PARAM_KEYS),
    "obsolete": "mark",
    "workset": "MEP_Openings",
    "schedule": "MEP Openings",
    "filters": True,
    "inventory": {},
}

HOST_CATEGORIES = (DB.BuiltInCategory.OST_Walls, DB.BuiltInCategory.OST_Floors, DB.BuiltInCategory.OST_Roofs,
                   DB.BuiltInCategory.OST_StructuralFraming, DB.BuiltInCategory.OST_StructuralColumns,
                   DB.BuiltInCategory.OST_StructuralFoundation)
LINK_ONLY_CATEGORIES = (DB.BuiltInCategory.OST_GenericModel,)  # elementi IFC arrivati come modelli generici


def id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


def deep_merge(base, data):
    """Copia di base aggiornata con data (i dizionari annidati si fondono, il resto si sostituisce)."""
    result = json.loads(json.dumps(base))
    for key, value in (data or {}).items():
        if isinstance(value, dict) and isinstance(result.get(key), dict) and key not in ("matrix", "inventory"):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def get_schema(guid_text, name):
    guid = System.Guid(guid_text)
    schema = Schema.Lookup(guid)
    if schema is not None:
        return schema
    builder = SchemaBuilder(guid)
    builder.SetReadAccessLevel(AccessLevel.Public)
    builder.SetWriteAccessLevel(AccessLevel.Public)
    builder.SetSchemaName(name)
    builder.AddSimpleField(JSON_FIELD, clr.GetClrType(System.String))
    return builder.Finish()


def config_schema():
    return get_schema(CONFIG_SCHEMA_GUID, "ESA_OpeningsConfig")


def opening_schema():
    return get_schema(OPENING_SCHEMA_GUID, "ESA_OpeningData")


def read_json(element, schema):
    try:
        entity = element.GetEntity(schema)
        if entity is None or not entity.IsValid():
            return None
        text = entity.Get[System.String](JSON_FIELD)
        return json.loads(text) if text else None
    except Exception:
        return None


def write_json(element, schema, data):
    """Dentro una transazione."""
    entity = Entity(schema)
    entity.Set[System.String](JSON_FIELD, json.dumps(data))
    element.SetEntity(entity)


def load_config():
    return deep_merge(DEFAULT_CONFIG, read_json(doc.ProjectInformation, config_schema()))


def save_config(config):
    """Dentro una transazione. Ogni salvataggio alza la revisione del setup."""
    config["rev"] = int(config.get("rev", 0)) + 1
    write_json(doc.ProjectInformation, config_schema(), config)


def structural_category_ids():
    """Categorie strutturali: quelle il cui nome interno inizia con OST_Structural (indipendente dalla lingua)."""
    ids = set()
    for value in System.Enum.GetValues(clr.GetClrType(DB.BuiltInCategory)):
        if str(value).startswith("OST_Structural"):
            try:
                ids.add(int(value))
            except Exception:
                pass
    return ids


STRUCTURAL_IDS = structural_category_ids()


def category_value(element):
    return id_value(element.Category.Id) if element.Category is not None else None


def link_name(instance):
    link_type = doc.GetElement(instance.GetTypeId())
    return DB.Element.Name.GetValue(link_type) if link_type is not None else instance.Name


def model_sources(unique_models=False):
    """Modello host e link caricati: [{doc, T, model, inst, label}], [link non caricati]. unique_models: un solo
    elemento per link (inventario), altrimenti una voce per ogni istanza (geometria)."""
    sources = [{"doc": doc, "T": DB.Transform.Identity, "model": HOST_MODEL, "inst": "", "label": HOST_MODEL}]
    unloaded, seen = [], set()
    for instance in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance):
        name = link_name(instance)
        link_doc = instance.GetLinkDocument()
        if link_doc is None:
            if name not in unloaded:
                unloaded.append(name)
            continue
        if unique_models and name in seen:
            continue
        seen.add(name)
        sources.append({"doc": link_doc, "T": instance.GetTotalTransform(), "model": name,
                        "inst": instance.UniqueId, "label": name})
    return sources, unloaded


def source_categories(source):
    return HOST_CATEGORIES + (LINK_ONLY_CATEGORIES if source["inst"] else ())


def source_elements(source):
    categories = List[DB.BuiltInCategory]()
    for category in source_categories(source):
        categories.Add(category)
    return (DB.FilteredElementCollector(source["doc"]).WherePasses(DB.ElementMulticategoryFilter(categories))
            .WhereElementIsNotElementType())


def type_label(element, source_doc):
    type_element = source_doc.GetElement(element.GetTypeId())
    if type_element is None:
        return element.Name
    family = getattr(type_element, "FamilyName", "") or ""
    name = DB.Element.Name.GetValue(type_element)
    return "{}: {}".format(family, name) if family else name


def inventory_key(model, category, label):
    return u"{}|{}|{}".format(model, category, label)


def param_text(param):
    if param is None or not param.HasValue:
        return None
    storage = param.StorageType
    if storage == DB.StorageType.String:
        return param.AsString()
    text = param.AsValueString()
    if text is None and storage == DB.StorageType.Integer:
        text = str(param.AsInteger())
    return text


def rule_matches(rule, element, type_element):
    name = (rule.get("param") or "").strip()
    if not name:
        return False
    where = rule.get("where", "both")
    targets = []
    if where in ("instance", "both"):
        targets.append(element)
    if where in ("type", "both") and type_element is not None:
        targets.append(type_element)
    wanted = (rule.get("value") or "").strip().lower()
    for target in targets:
        for param in target.GetParameters(name):
            text = param_text(param)
            if text is None or not text.strip():
                continue
            text = text.strip().lower()
            op = rule.get("op", "nonempty")
            if op == "nonempty" or (op == "equals" and text == wanted) or (op == "contains" and wanted and wanted in text):
                return True
    return False


def rei_by_rules(rules, element, source_doc):
    """Vero se almeno una regola REI del progetto e' soddisfatta dall'elemento o dal suo tipo."""
    type_element = source_doc.GetElement(element.GetTypeId())
    return any(rule_matches(rule, element, type_element) for rule in rules or [])

# =================================================================================================================
# fine del blocco comune di configurazione
# =================================================================================================================


def symbol_keys():
    keys = []
    for symbol in (DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
                   .OfCategory(DB.BuiltInCategory.OST_GenericModel)):
        keys.append(u"{} : {}".format(symbol.FamilyName, DB.Element.Name.GetValue(symbol)))
    return sorted(keys)


def symbol_by_key(key):
    for symbol in (DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
                   .OfCategory(DB.BuiltInCategory.OST_GenericModel)):
        if u"{} : {}".format(symbol.FamilyName, DB.Element.Name.GetValue(symbol)) == key:
            return symbol
    return None


# --- pulsanti della finestra ------------------------------------------------------------------------------------

def test_rei(rules):
    """Conta, per ogni modello e categoria, gli elementi che soddisfano le regole REI."""
    sources, unloaded = model_sources(unique_models=True)
    lines = []
    for source in sources:
        counts = {}
        for element in source_elements(source):
            if element.Category is None:
                continue
            name = element.Category.Name
            total, matching = counts.get(name, (0, 0))
            try:
                hit = rei_by_rules(rules, element, source["doc"])
            except Exception:
                hit = False
            counts[name] = (total + 1, matching + (1 if hit else 0))
        if counts:
            lines.append(u"{}: {}".format(source["label"], "; ".join(
                u"{} {} of {}".format(name, matching, total) for name, (total, matching) in sorted(counts.items()))))
    if unloaded:
        lines.append("Not loaded: {}".format(", ".join(unloaded)))
    return u"Elements REI by the rules:\n\n" + u"\n".join(lines or ["No element found."])


def check_families(keys, dims):
    """Apre ogni famiglia scelta e dice quali parametri dimensionali ha (di istanza / di tipo / assenti)."""
    lines, seen = [], set()
    for key in keys:
        symbol = symbol_by_key(key)
        if symbol is None:
            lines.append(u"{}: not loaded in the project".format(key))
            continue
        family = symbol.Family
        if family.Id in seen:
            continue
        seen.add(family.Id)
        placement = family.FamilyPlacementType
        notes = []
        if placement != DB.FamilyPlacementType.OneLevelBased:
            notes.append("placement {}: it must be non-hosted, level based".format(placement))
        try:
            family_doc = doc.EditFamily(family)
            try:
                found = dict((p.Definition.Name, p.IsInstance) for p in family_doc.FamilyManager.Parameters)
            finally:
                family_doc.Close(False)
        except Exception as ex:
            lines.append(u"{}: could not be opened ({})".format(family.Name, ex))
            continue
        parts = []
        for label, name in (("height", dims.get("height")), ("width", dims.get("width")),
                            ("depth", dims.get("depth")), ("diameter", dims.get("diameter"))):
            if not name:
                continue
            if name not in found:
                parts.append(u"{} '{}' missing".format(label, name))
            elif not found[name]:
                parts.append(u"{} '{}' is a TYPE parameter (not used)".format(label, name))
            else:
                parts.append(u"{} OK".format(label))
        lines.append(u"{}: {}{}".format(family.Name, ", ".join(parts) or "no size parameter set",
                                        (" | " + "; ".join(notes)) if notes else ""))
    return u"\n".join(lines)


def export_setup(values):
    path = forms.save_file(file_ext="json", default_name="ESA_Openings_Setup")
    if not path:
        return None
    data = dict(values)
    data["version"] = DEFAULT_CONFIG["version"]
    with io.open(path, "w", encoding="utf-8") as handle:
        handle.write(json.dumps(data, indent=2, ensure_ascii=False))
    return u"Setup saved to:\n{}".format(path)


def import_setup():
    path = forms.pick_file(file_ext="json")
    if not path:
        return None
    try:
        with io.open(path, "r", encoding="utf-8") as handle:
            data = json.loads(handle.read())
    except Exception as ex:
        forms.alert("The file could not be read: {}".format(ex), title="Openings Setup")
        return None
    data.pop("inventory", None)
    data.pop("rev", None)
    return deep_merge(DEFAULT_CONFIG, data)


# --- workset, abaco, filtri -------------------------------------------------------------------------------------

def run_step(name, action, notes):
    try:
        with revit.Transaction(name):
            notes.extend(action())
    except Exception as ex:
        notes.append("{}: failed ({})".format(name, ex))


def make_workset(config):
    name = config.get("workset", "")
    if not name:
        return ["Workset: none set"]
    if not doc.IsWorkshared:
        return ["Workset: the model is not workshared, none created"]
    if not DB.WorksetTable.IsWorksetNameUnique(doc, name):
        return [u"Workset '{}': already exists".format(name)]
    DB.Workset.Create(doc, name)
    return [u"Workset '{}': created".format(name)]


def make_schedule(config):
    name = config.get("schedule", "")
    if not name:
        return ["Schedule: none set"]
    for view in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule):
        if view.Name == name:
            return [u"Schedule '{}': already exists, not changed".format(name)]
    schedule = DB.ViewSchedule.CreateSchedule(doc, DB.ElementId(DB.BuiltInCategory.OST_GenericModel))
    schedule.Name = name
    definition = schedule.Definition
    fields = list(definition.GetSchedulableFields())
    family_type = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
    for field in fields:
        if field.ParameterId == family_type:
            definition.AddField(field)
            break
    wanted = []
    for key in PARAM_KEYS:
        if config["params"].get(key):
            wanted.append(config["params"][key])
    for key in ("width", "height", "depth", "diameter"):
        if config["dims"].get(key):
            wanted.append(config["dims"][key])
    added, missing, status_field = [], [], None
    for name_wanted in wanted:
        if name_wanted in added:
            continue
        match = next((f for f in fields if f.GetName(doc) == name_wanted), None)
        if match is None:
            missing.append(name_wanted)
            continue
        schedule_field = definition.AddField(match)
        added.append(name_wanted)
        if name_wanted == config["params"].get("status"):
            status_field = schedule_field
    count = next((f for f in fields if f.FieldType == DB.ScheduleFieldType.Count), None)
    if count is not None:
        definition.AddField(count)
    notes = [u"Schedule '{}': created with {} parameters".format(name, len(added))]
    if status_field is not None:
        try:
            definition.AddFilter(DB.ScheduleFilter(status_field.FieldId, DB.ScheduleFilterType.HasValue))
            notes.append("Schedule filtered on the Status parameter (only the openings)")
        except Exception as ex:
            notes.append("Schedule filter not added: {}".format(ex))
    else:
        notes.append("Schedule not filtered: it lists all the generic models (set a Status parameter to filter)")
    if missing:
        notes.append(u"Not schedulable (not shared / project parameters): {}".format(", ".join(missing)))
    return notes


def parameter_element_id(name):
    for cls in (DB.SharedParameterElement, DB.ParameterElement):
        for element in DB.FilteredElementCollector(doc).OfClass(cls):
            try:
                if element.GetDefinition().Name == name:
                    return element.Id
            except Exception:
                continue
    return None


def equals_rule(parameter_id, value):
    try:
        return DB.ParameterFilterRuleFactory.CreateEqualsRule(parameter_id, value)
    except Exception:
        return DB.ParameterFilterRuleFactory.CreateEqualsRule(parameter_id, value, False)


def make_filters(config):
    if not config.get("filters"):
        return ["View filters: not requested"]
    status = config["params"].get("status", "")
    if not status:
        return ["View filters: no Status parameter set, not created"]
    parameter_id = parameter_element_id(status)
    if parameter_id is None:
        return [u"View filters: '{}' is not a shared / project parameter, not created".format(status)]
    existing = set(f.Name for f in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement))
    categories = List[DB.ElementId]()
    categories.Add(DB.ElementId(DB.BuiltInCategory.OST_GenericModel))
    notes = []
    for value in ("New", "Modified", "Obsolete"):
        name = "Openings - {}".format(value)
        if name in existing:
            notes.append(u"Filter '{}': already exists".format(name))
            continue
        DB.ParameterFilterElement.Create(doc, name, categories, DB.ElementParameterFilter(equals_rule(parameter_id,
                                                                                                    value)))
        notes.append(u"Filter '{}': created (add it to the views or view templates)".format(name))
    return notes


# --- comando ----------------------------------------------------------------------------------------------------

def main():
    config = load_config()
    window = Setup_ui.SetupWindow(
        script.get_bundle_file(XAML_FILE_NAME), config, symbol_keys(),
        [(k, MEP_KIND_LABELS[k]) for k in MEP_KINDS], [(c, CASE_LABELS[c]) for c in CASES],
        [(k, PARAM_LABELS[k]) for k in PARAM_KEYS],
        {"test_rei": test_rei, "check_families": check_families, "export": export_setup,
         "import_setup": import_setup})
    result = window.show()
    if result is None:
        script.exit()
    values, apply_now = result
    config.update(values)
    try:
        with revit.Transaction("Openings Setup"):
            save_config(config)
    except Exception as ex:
        forms.alert("The setup could not be saved in the model: {}".format(ex), title="Openings Setup")
        script.exit()

    output.print_md("# Openings Setup")
    output.print_md("Saved in the model (revision **{}**) | mode: **{}** | REI rules: **{}** | types in Links: "
                    "**{}**".format(config["rev"], config["mode"], len(config["rei_rules"]),
                                    len(config["inventory"])))
    if config["mode"] == "simple":
        for key, label in (("h_rect", "Wall, rectangular"), ("h_round", "Wall, round"),
                           ("v_rect", "Floor, rectangular"), ("v_round", "Floor, round")):
            output.print_md(u"{}: {}".format(label, config["simple"].get(key) or "-"))
    else:
        output.print_md("Matrix: **{}** of {} cases with a family".format(len(config["matrix"]),
                                                                       len(MEP_KINDS) * len(CASES)))
    if not config["inventory"]:
        output.print_md("**Run Links** to choose the types of walls, floors and structural elements to check.")
    if apply_now:
        notes = []
        run_step("Openings workset", lambda: make_workset(config), notes)
        run_step("Openings schedule", lambda: make_schedule(config), notes)
        run_step("Openings view filters", lambda: make_filters(config), notes)
        output.print_md("## Workset, schedule, filters")
        for note in notes:
            output.print_md(note)


main()
