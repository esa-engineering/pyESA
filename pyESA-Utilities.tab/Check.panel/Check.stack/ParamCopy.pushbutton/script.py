# -*- coding: utf-8 -*-
__title__ = "Param Copy"
__doc__ = """Version = 0.2
Date    = 08.10.2026
________________________________________________________________
Copy parameters from a source element to multiple targets.
Step 1: pick source element.
Step 2: choose instance and type parameters to copy, and whether
        to overwrite values already present on the targets.
Step 3: pick target elements.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import StorageType, ElementId
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script
import os.path as op

import ParamCopy_ui

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

XAML_FILE_NAME = "ParamCopy_form.xaml"
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
        ref = uidoc.Selection.PickObject(ObjectType.Element, "Select source element to copy from")
    except OperationCanceledException:
        script.exit()
    return doc.GetElement(ref.ElementId)


def pick_targets():
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select target elements to copy to")
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


def get_param_list(source):
    """Parametri copiabili: istanza ([inst]) e tipo ([type]), deduplicati per scope+nome."""
    result = {}
    sources = [("inst", source), ("type", get_type_element(source))]
    for scope, elem in sources:
        if elem is None:
            continue
        for p in elem.Parameters:
            if p.StorageType not in ALLOWED_STORAGE or p.IsReadOnly:
                continue
            name = p.Definition.Name
            key = (scope, name)
            if key not in result:
                result[key] = {
                    "name": name,
                    "scope": scope,
                    "storage": p.StorageType,
                    "label": "[{}] {}".format(scope, name),
                }
    return sorted(result.values(), key=lambda d: (d["scope"], d["name"].lower()))


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
    """True se il parametro di destinazione non ha ancora un valore significativo."""
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


def main():
    source = pick_source()
    params = get_param_list(source)
    if not params:
        output.print_md("No copyable parameters found on the source element.")
        script.exit()

    xaml_path = script.get_bundle_file(XAML_FILE_NAME) or op.join(op.dirname(__file__), XAML_FILE_NAME)
    label = "{} (Id {})".format(source.Name if hasattr(source, "Name") else "Element",
                                get_element_id_value(source.Id))
    choice = ParamCopy_ui.ParamCopyWindow(xaml_path, label, params).show()
    if not choice:
        script.exit()
    chosen = choice["params"]
    overwrite = choice["overwrite"]

    targets = pick_targets()
    if not targets:
        output.print_md("No targets selected.")
        script.exit()

    # Valori sorgente letti una volta sola
    source_type = get_type_element(source)
    source_values = {}
    source_empty = set()
    for p in chosen:
        elem = source if p["scope"] == "inst" else source_type
        if elem is None:
            continue
        sp = elem.LookupParameter(p["name"])
        if sp is None or sp.StorageType != p["storage"]:
            continue
        # sorgente vuota: con overwrite attivo il valore di destinazione viene azzerato
        if is_empty(sp):
            source_values[(p["scope"], p["name"])] = empty_value(p["storage"])
            source_empty.add((p["scope"], p["name"]))
        else:
            source_values[(p["scope"], p["name"])] = read_value(sp)

    results = []
    done_types = set()  # un parametro di tipo si scrive una sola volta per tipo
    with revit.Transaction("Param Copy"):
        for tgt in targets:
            if tgt.Id == source.Id:
                continue
            copied = skipped = failed = 0
            for p in chosen:
                key = (p["scope"], p["name"])
                if key not in source_values:
                    continue
                if p["scope"] == "inst":
                    holder = tgt
                else:
                    holder = get_type_element(tgt)
                    if holder is None or (holder.Id.ToString(), p["name"]) in done_types:
                        continue
                    # se il tipo e' lo stesso della sorgente non c'e' nulla da fare
                    if source_type is not None and holder.Id == source_type.Id:
                        continue
                tp = holder.LookupParameter(p["name"])
                if tp is None or tp.IsReadOnly or tp.StorageType != p["storage"]:
                    failed += 1
                    continue
                if not overwrite and (key in source_empty or not is_empty(tp)):
                    skipped += 1
                    continue
                try:
                    if tp.Set(source_values[key]):
                        copied += 1
                        if p["scope"] == "type":
                            done_types.add((holder.Id.ToString(), p["name"]))
                    else:
                        failed += 1
                except Exception:
                    failed += 1
            results.append({"target": tgt, "copied": copied, "skipped": skipped, "failed": failed})

    total = len(results)
    problems = sum(1 for r in results if r["copied"] == 0 and r["skipped"] == 0)
    if want_report(total, problems, "elements updated"):
        rows = [[output.linkify(r["target"].Id), str(r["copied"]), str(r["skipped"]), str(r["failed"])]
                for r in results]
        output.print_table(table_data=rows,
                           columns=["Element", "Parameters copied", "Skipped (already filled)", "Failed"])
    else:
        output.print_md("Copied parameters to {} elements.".format(total))


if __name__ == "__main__":
    main()
