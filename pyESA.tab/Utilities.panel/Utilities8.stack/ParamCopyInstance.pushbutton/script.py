# -*- coding: utf-8 -*-
__title__ = "Param Copy\nInstance"
__doc__ = """Version = 0.3
Date    = 12.10.2026
________________________________________________________________
Copy instance parameters from a source element to multiple targets.
Targets: the elements already selected when the command starts (any
Revit selection: Tab for a connected chain, window, filters). With
nothing selected, the targets are picked after the window.
Step 1: pick source element (excluded from the targets).
Step 2: choose the instance parameters to copy, and whether
        to overwrite values already present on the targets.
Step 3: only with nothing preselected, pick the target elements.
Parameters that a target does not have (e.g. fittings in a chain)
are counted apart, not as errors.
Type parameters are copied by the Param Copy Type button.
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

import ParamCopyInstance_ui

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

XAML_FILE_NAME = "ParamCopyInstance_form.xaml"
CONFIG_NAME = "ESA_ParamCopyInstance"
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


def get_preselection():
    """Elementi gia' selezionati all'avvio: sono i target (selezione nativa di Revit, Tab compreso)."""
    elements = []
    for eid in uidoc.Selection.GetElementIds():
        element = doc.GetElement(eid)
        if element is not None:
            elements.append(element)
    return elements


def pick_source():
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, "Select the source element to copy instance "
                                                             "parameters from")
    except OperationCanceledException:
        script.exit()
    return doc.GetElement(ref.ElementId)


def pick_targets():
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select the target elements, then press Finish")
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def category_name(element):
    try:
        return element.Category.Name
    except Exception:
        return ""


def element_label(element):
    try:
        name = element.Name
    except Exception:
        name = "Element"
    return "{} (Id {})".format(name, get_element_id_value(element.Id))


def get_param_list(element):
    """Parametri d'istanza copiabili dell'elemento sorgente, deduplicati per nome."""
    result = {}
    for p in element.Parameters:
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


def main():
    # la selezione va letta prima del clic sulla sorgente
    preselected = get_preselection()
    source = pick_source()
    targets = [t for t in preselected if t.Id != source.Id]
    if targets:
        targets_text = "{} element(s) already selected in Revit".format(len(targets))
        if len(targets) < len(preselected):
            targets_text += " (the source element is excluded)"
    else:
        targets_text = "None selected: you will pick them after OK"
    params = get_param_list(source)
    if not params:
        output.print_md("# Param Copy Instance report")
        output.print_md("The source element {} has no writable instance parameter that can be copied "
                        "(text, integer, number or element).".format(output.linkify(source.Id)))
        script.exit()

    xaml_path = script.get_bundle_file(XAML_FILE_NAME) or op.join(op.dirname(__file__), XAML_FILE_NAME)
    choice = ParamCopyInstance_ui.ParamCopyWindow(xaml_path, CONFIG_NAME, element_label(source), targets_text,
                                                  params).show()
    if not choice:
        script.exit()
    chosen = choice["params"]
    overwrite = choice["overwrite"]

    preselection_used = bool(targets)
    if not preselection_used:
        targets = [t for t in pick_targets() if t.Id != source.Id]
    if not targets:
        output.print_md("# Param Copy Instance report")
        output.print_md("No target element was selected (the source element itself is ignored).")
        script.exit()

    # valori sorgente letti una volta sola; sorgente vuota: con overwrite il valore di destinazione si azzera
    source_values = {}
    source_empty = set()
    for p in chosen:
        sp = source.LookupParameter(p["name"])
        if sp is None or sp.StorageType != p["storage"]:
            continue
        if is_empty(sp):
            source_values[p["name"]] = empty_value(p["storage"])
            source_empty.add(p["name"])
        else:
            source_values[p["name"]] = read_value(sp)

    results = []
    with revit.Transaction("Param Copy Instance"):
        for tgt in targets:
            copied = skipped = absent = failed = 0
            for p in chosen:
                name = p["name"]
                if name not in source_values:
                    continue
                tp = tgt.LookupParameter(name)
                if tp is None:
                    # parametro assente sul target (es. raccordi di una catena): non e' un errore
                    absent += 1
                    continue
                if tp.IsReadOnly or tp.StorageType != p["storage"]:
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
            results.append({"target": tgt, "copied": copied, "skipped": skipped, "absent": absent,
                            "failed": failed})

    total = len(results)
    problems = sum(1 for r in results if r["failed"])
    if want_report(total, problems, "elements checked"):
        output.print_md("# Param Copy Instance report")
        output.print_md("Source: {} | targets: **{}** ({}) | parameters chosen: **{}** | overwrite: **{}**".format(
            output.linkify(source.Id), total, "preselected" if preselection_used else "picked", len(chosen),
            "on" if overwrite else "off"))
        updated = sum(1 for r in results if r["copied"])
        not_applicable = sum(1 for r in results if not r["copied"] and not r["skipped"] and not r["failed"])
        output.print_md("Elements updated: **{}** | without any of the chosen parameters: **{}** | "
                        "with errors: **{}**".format(updated, not_applicable, problems))
        rows = [[output.linkify(r["target"].Id), category_name(r["target"]), str(r["copied"]), str(r["skipped"]),
                 str(r["absent"]), str(r["failed"])] for r in results]
        output.print_table(table_data=rows,
                           columns=["Element", "Category", "Parameters copied", "Skipped (already filled)",
                                    "Not on element", "Failed (read-only or other type)"])


if __name__ == "__main__":
    main()
