# -*- coding: utf-8 -*-
__title__ = "Links"
__doc__ = """Version = 0.1
Date    = 09.10.2026
________________________________________________________________
Description:
First command of the openings workflow, once per project and again
whenever links are added or reloaded.

Finds this model and the loaded links (RVT and IFC) and lists the types
of walls, floors, roofs and structural elements (categories starting
with Structural), plus the generic models of the links (IFC elements
often arrive as generic models). For each type:
- tick it to check it for openings (generic models start unticked)
- REI: Auto uses the REI rules of Setup (the count of elements that
  match is shown), or force REI / Not REI for the whole type
- Structural is automatic from the category

The list is saved in the model and used by Place and Refresh. New
types (after a link reload) and types no longer found are reported;
the settings of types no longer found are kept, in case a link is only
unloaded.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import clr
import json
clr.AddReference("RevitAPI")
import System
from System.Collections.Generic import List
from Autodesk.Revit.DB.ExtensibleStorage import Schema, SchemaBuilder, Entity, AccessLevel

from pyrevit import revit, script, DB, forms

import Links_ui

doc = revit.doc
output = script.get_output()

XAML_FILE_NAME = "Links_form.xaml"

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


def scan(config):
    """Tipi presenti nel modello e nei link, uniti alle scelte gia' salvate."""
    sources, unloaded = model_sources(unique_models=True)
    entries = {}
    for source in sources:
        for element in source_elements(source):
            if element.Category is None:
                continue
            category = category_value(element)
            label = type_label(element, source["doc"])
            key = inventory_key(source["model"], category, label)
            entry = entries.get(key)
            if entry is None:
                entry = {"key": key, "model": source["model"], "category": element.Category.Name, "cat": category,
                         "type": label, "count": 0, "rei_count": 0, "structural": category in STRUCTURAL_IDS}
                entries[key] = entry
            entry["count"] += 1
            try:
                if rei_by_rules(config["rei_rules"], element, source["doc"]):
                    entry["rei_count"] += 1
            except Exception:
                pass
    saved = config["inventory"]
    generic = int(DB.BuiltInCategory.OST_GenericModel)
    for key, entry in entries.items():
        old = saved.get(key)
        entry["new"] = old is None
        entry["include"] = bool(old.get("include", True)) if old else entry["cat"] != generic
        entry["rei"] = old.get("rei", "auto") if old else "auto"
    missing = []
    for key, old in saved.items():
        if key not in entries:
            item = dict(old)
            item["key"] = key
            missing.append(item)
    ordered = sorted(entries.values(), key=lambda e: (e["model"] != HOST_MODEL, e["model"], e["category"], e["type"]))
    return ordered, missing, unloaded


def main():
    config = load_config()
    entries, missing, unloaded = scan(config)
    if not entries:
        forms.alert("No wall, floor, roof or structural element found in this model or in the loaded links.",
                    title="Openings Links")
        script.exit()
    xaml_path = script.get_bundle_file(XAML_FILE_NAME)
    choices = Links_ui.LinksWindow(xaml_path, entries).show()
    if choices is None:
        script.exit()

    inventory = {}
    for item in missing:  # tipi non trovati (link scaricato o tipo rinominato): le scelte restano
        key = item.pop("key")
        inventory[key] = item
    for entry in entries:
        choice = choices[entry["key"]]
        inventory[entry["key"]] = {"include": choice["include"], "rei": choice["rei"], "model": entry["model"],
                                   "category": entry["category"], "cat": entry["cat"], "type": entry["type"],
                                   "count": entry["count"], "structural": entry["structural"]}
    config["inventory"] = inventory
    try:
        with revit.Transaction("Openings Links"):
            save_config(config)
    except Exception as ex:
        forms.alert("The list could not be saved in the model: {}".format(ex), title="Openings Links")
        script.exit()

    output.print_md("# Openings Links")
    models = []
    for entry in entries:
        if entry["model"] not in models:
            models.append(entry["model"])
    for model in models:
        rows = [e for e in entries if e["model"] == model]
        checked = [e for e in rows if choices[e["key"]]["include"]]
        output.print_md("**{}** | types: **{}** | checked: **{}** | elements checked: **{}**".format(
            model, len(rows), len(checked), sum(e["count"] for e in checked)))
    new = [e for e in entries if e["new"]]
    if new:
        output.print_md("## New types ({})".format(len(new)))
        for e in new:
            output.print_md("{} | {} | {} | {} elements | {}".format(
                e["model"], e["category"], e["type"], e["count"],
                "checked" if choices[e["key"]]["include"] else "not checked"))
    if missing:
        output.print_md("## Types no longer found ({}), settings kept".format(len(missing)))
        for item in missing:
            output.print_md("{} | {} | {}".format(item.get("model", ""), item.get("category", ""),
                                                  item.get("type", "")))
    if unloaded:
        output.print_md("Links not loaded (not listed): {}".format(", ".join(unloaded)))
    rei = [e for e in entries if choices[e["key"]]["rei"] == "rei" or
           (choices[e["key"]]["rei"] == "auto" and e["rei_count"])]
    output.print_md("REI types (forced or by the rules of Setup): **{}**. Structural types: **{}**.".format(
        len(rei), len([e for e in entries if e["structural"]])))


main()
