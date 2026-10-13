# -*- coding: utf-8 -*-
__title__ = "Param Copy\nType"
__doc__ = """Version = 0.3
Date    = 12.10.2026
________________________________________________________________
Copy type parameters from the type of a source element to the types
of the target elements.
Step 1: pick source element (its type is the source).
Step 2: choose the type parameters to copy, and whether to overwrite
        values already present on the target types.
Step 3: pick target elements: each different type among them is
        written once. Changing a type changes all its instances.
The type of the source element itself is never changed.
Instance parameters are copied by the Param Copy Instance button.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import StorageType, ElementId, Element
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script
import os.path as op

import ParamCopyType_ui

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

XAML_FILE_NAME = "ParamCopyType_form.xaml"
CONFIG_NAME = "ESA_ParamCopyType"
REPORT_CONFIG = "ESA_Report"

ALLOWED_STORAGE = (StorageType.String, StorageType.Integer,
                   StorageType.Double, StorageType.ElementId)


def get_element_id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


def want_report(total, problems, noun):
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
        return problems > 0
    return False


def pick_source():
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, "Select the source element: its type parameters "
                                                             "will be copied")
    except OperationCanceledException:
        script.exit()
    return doc.GetElement(ref.ElementId)


def pick_targets():
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select elements whose types receive the values, "
                                                               "then press Finish")
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def get_type_element(element):
    try:
        tid = element.GetTypeId()
        if tid is not None and get_element_id_value(tid) != -1:
            return doc.GetElement(tid)
    except Exception:
        pass
    return None


def type_label(element_type):
    """'Famiglia : Tipo'. Element.Name sui tipi in IronPython va letto con Element.Name.GetValue."""
    try:
        name = Element.Name.GetValue(element_type)
    except Exception:
        name = getattr(element_type, "Name", "Type")
    family = getattr(element_type, "FamilyName", "") or ""
    return "{} : {}".format(family, name) if family else name


def element_label(element):
    try:
        name = element.Name
    except Exception:
        name = "Element"
    return "{} (Id {})".format(name, get_element_id_value(element.Id))


def get_param_list(element_type):
    """Parametri di tipo copiabili, deduplicati per nome."""
    result = {}
    for p in element_type.Parameters:
        if p.StorageType not in ALLOWED_STORAGE or p.IsReadOnly:
            continue
        name = p.Definition.Name
        if name not in result:
            result[name] = {"name": name, "storage": p.StorageType, "label": name}
    return sorted(result.values(), key=lambda d: d["name"].lower())


def read_value(param):
    st = param.StorageType
    if st == StorageType.String:
        return param.AsString()
    if st == StorageType.Integer:
        return param.AsInteger()
    if st == StorageType.Double:
        return param.AsDouble()
    return param.AsElementId()


def is_empty(param):
    """True se il parametro non ha ancora un valore significativo."""
    if not param.HasValue:
        return True
    st = param.StorageType
    if st == StorageType.String:
        return not param.AsString()
    if st == StorageType.Integer:
        return param.AsInteger() == 0
    if st == StorageType.Double:
        return param.AsDouble() == 0.0
    if st == StorageType.ElementId:
        return get_element_id_value(param.AsElementId()) == -1
    return True


def empty_value(storage):
    if storage == StorageType.String:
        return ""
    if storage == StorageType.Integer:
        return 0
    if storage == StorageType.Double:
        return 0.0
    return ElementId.InvalidElementId


def group_by_type(targets, source, source_type):
    """Tipi distinti dei target, in ordine di selezione. Ritorna (tipi, n. elementi senza tipo, n. dello stesso tipo
    della sorgente)."""
    types, order = {}, []
    no_type = same_type = 0
    for tgt in targets:
        if tgt.Id == source.Id:
            continue
        t = get_type_element(tgt)
        if t is None:
            no_type += 1
            continue
        if t.Id == source_type.Id:
            same_type += 1
            continue
        key = get_element_id_value(t.Id)
        if key not in types:
            types[key] = {"type": t, "instances": 0}
            order.append(key)
        types[key]["instances"] += 1
    return [types[k] for k in order], no_type, same_type


def main():
    source = pick_source()
    source_type = get_type_element(source)
    if source_type is None:
        output.print_md("# Param Copy Type report")
        output.print_md("The source element {} has no type, so there are no type parameters to copy. "
                        "Use Param Copy Instance.".format(output.linkify(source.Id)))
        script.exit()
    params = get_param_list(source_type)
    if not params:
        output.print_md("# Param Copy Type report")
        output.print_md("The type {} has no writable type parameter that can be copied "
                        "(text, integer, number or element).".format(output.linkify(source_type.Id)))
        script.exit()

    xaml_path = script.get_bundle_file(XAML_FILE_NAME) or op.join(op.dirname(__file__), XAML_FILE_NAME)
    choice = ParamCopyType_ui.ParamCopyWindow(xaml_path, CONFIG_NAME, element_label(source),
                                              type_label(source_type), params).show()
    if not choice:
        script.exit()
    chosen = choice["params"]
    overwrite = choice["overwrite"]

    target_types, no_type, same_type = group_by_type(pick_targets(), source, source_type)
    if not target_types:
        output.print_md("# Param Copy Type report")
        output.print_md("No target type to change: elements of the same type as the source: **{}**, "
                        "elements without a type: **{}**.".format(same_type, no_type))
        script.exit()

    # valori sorgente letti una volta sola; sorgente vuota: con overwrite il valore di destinazione si azzera
    source_values = {}
    source_empty = set()
    for p in chosen:
        sp = source_type.LookupParameter(p["name"])
        if sp is None or sp.StorageType != p["storage"]:
            continue
        if is_empty(sp):
            source_values[p["name"]] = empty_value(p["storage"])
            source_empty.add(p["name"])
        else:
            source_values[p["name"]] = read_value(sp)

    results = []
    with revit.Transaction("Param Copy Type"):
        for item in target_types:
            holder = item["type"]
            copied = skipped = failed = 0
            for p in chosen:
                name = p["name"]
                if name not in source_values:
                    continue
                tp = holder.LookupParameter(name)
                if tp is None or tp.IsReadOnly or tp.StorageType != p["storage"]:
                    failed += 1
                    continue
                if not overwrite and (name in source_empty or not is_empty(tp)):
                    skipped += 1
                    continue
                try:
                    if tp.Set(source_values[name]):
                        copied += 1
                    else:
                        failed += 1
                except Exception:
                    failed += 1
            results.append({"type": holder, "instances": item["instances"], "copied": copied,
                            "skipped": skipped, "failed": failed})

    total = len(results)
    problems = sum(1 for r in results if r["failed"] or (r["copied"] == 0 and r["skipped"] == 0))
    if want_report(total, problems, "types updated"):
        output.print_md("# Param Copy Type report")
        output.print_md("Source type: {} {} | parameters chosen: **{}** | overwrite: **{}**".format(
            type_label(source_type), output.linkify(source_type.Id), len(chosen), "on" if overwrite else "off"))
        rows = [[output.linkify(r["type"].Id), type_label(r["type"]), str(r["instances"]), str(r["copied"]),
                 str(r["skipped"]), str(r["failed"])] for r in results]
        output.print_table(table_data=rows,
                           columns=["Type", "Family : Type", "Selected elements", "Parameters copied",
                                    "Skipped (already filled)", "Failed (missing, read-only or other type)"])
        if same_type or no_type:
            output.print_md("Not changed: **{}** selected elements of the same type as the source, **{}** "
                            "elements without a type.".format(same_type, no_type))


if __name__ == "__main__":
    main()
