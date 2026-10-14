# -*- coding: utf-8 -*-
__title__ = "Select\nRun"
__doc__ = """Version = 0.1
Date    = 12.10.2026
________________________________________________________________
Description:
Selects the whole run of the selected elements (select them first,
or pick them after clicking): from each one it follows straight runs,
flex runs, fittings with two connections (elbows, transitions,
unions) and inline accessories (valves, dampers), and stops at tees,
crosses, taps, equipment and terminals, which are not selected. From
a run only the end connections are followed, so a tap on a main does
not lead into the branch. The insulation and lining of the found
elements are selected too.

CLICK: selects, leaving out the categories excluded last time.
SHIFT + CLICK: after the search, a window lists the categories found
(with the number of elements), all checked: uncheck the ones you do not
want selected. The search still goes through them. The choice is saved
for the normal click; a category never seen before is selected.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import os

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Collections.Generic import List

from pyrevit import revit, script, DB

import Select_ui

MODE = "run"  # "run": fino alle diramazioni; "connected": tutto il collegato, anche attraverso le apparecchiature
TITLE = {"run": "Select Run", "connected": "Select Connected"}[MODE]
CONFIG_NAME = {"run": "ESA_SelectRun", "connected": "ESA_SelectConnected"}[MODE]
XAML_FILE_NAME = "Select_form.xaml"

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


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


PHYSICAL_TYPES = (DB.ConnectorType.End, DB.ConnectorType.Curve, DB.ConnectorType.Physical)
# raccordi e accessori in linea: con due connessioni la run li attraversa, con piu' di due (tee, croci) si ferma
INLINE_CATEGORIES = set(int(c) for c in (
    DB.BuiltInCategory.OST_PipeFitting, DB.BuiltInCategory.OST_DuctFitting,
    DB.BuiltInCategory.OST_CableTrayFitting, DB.BuiltInCategory.OST_ConduitFitting,
    DB.BuiltInCategory.OST_PipeAccessory, DB.BuiltInCategory.OST_DuctAccessory))


# isolamento solo dove Revit lo ammette: chiamare GetInsulationIds / GetLiningIds su passerelle e loro raccordi
# e' il sospetto principale del crash di Revit durante Select Run
INSULATION_CATEGORIES = set(int(c) for c in (
    DB.BuiltInCategory.OST_PipeCurves, DB.BuiltInCategory.OST_FlexPipeCurves, DB.BuiltInCategory.OST_PipeFitting,
    DB.BuiltInCategory.OST_PipeAccessory, DB.BuiltInCategory.OST_DuctCurves, DB.BuiltInCategory.OST_FlexDuctCurves,
    DB.BuiltInCategory.OST_DuctFitting, DB.BuiltInCategory.OST_DuctAccessory))
LINING_CATEGORIES = set(int(c) for c in (
    DB.BuiltInCategory.OST_DuctCurves, DB.BuiltInCategory.OST_FlexDuctCurves, DB.BuiltInCategory.OST_DuctFitting,
    DB.BuiltInCategory.OST_DuctAccessory))

# traccia su file, scritta e chiusa a ogni passo: se Revit si chiude, l'ultima riga dice dove (diagnosi del crash)
TRACE_PATH = os.path.join(os.environ.get("TEMP", "."), "ESA_Select_trace.txt")


def trace(text, reset=False):
    try:
        with open(TRACE_PATH, "w" if reset else "a") as handle:
            handle.write(text + "\n")
    except Exception:
        pass


def get_id(element_id):
    return element_id.Value if hasattr(element_id, "Value") else element_id.IntegerValue  # Revit 2026+ / <= 2025


def connector_manager(element):
    """ConnectorManager di un tratto MEP o di una famiglia (raccordo, accessorio, apparecchiatura); None se non c'e'."""
    if isinstance(element, DB.MEPCurve):
        return element.ConnectorManager
    model = getattr(element, "MEPModel", None)
    try:
        return model.ConnectorManager if model is not None else None
    except Exception:
        return None


def physical_connectors(element):
    manager = connector_manager(element)
    if manager is None:
        return []
    return [c for c in manager.Connectors if c.ConnectorType in PHYSICAL_TYPES]


def is_mep(element):
    return element is not None and bool(physical_connectors(element))


class MepFilter(ISelectionFilter):
    def AllowElement(self, element):
        try:
            return is_mep(element)
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return False


def neighbours(element, run_only):
    """Elementi collegati fisicamente a element. run_only: da un tratto si seguono solo i connettori di estremita'
    e non si entra in un tratto dal suo connettore laterale (uno stacco sul principale e' una diramazione)."""
    found = []
    trace("neighbours of {} ({})".format(get_id(element.Id), category_key(element)[1]))
    for conn in physical_connectors(element):
        if not conn.IsConnected:
            continue
        if run_only and isinstance(element, DB.MEPCurve) and conn.ConnectorType == DB.ConnectorType.Curve:
            continue
        trace("  connector {} type {}: AllRefs".format(conn.Id, conn.ConnectorType))
        for ref in conn.AllRefs:
            owner = ref.Owner
            if owner is None or owner.Id == element.Id or ref.ConnectorType not in PHYSICAL_TYPES:
                continue
            if isinstance(owner, DB.MEPSystem):
                continue
            if run_only and isinstance(owner, DB.MEPCurve) and ref.ConnectorType == DB.ConnectorType.Curve:
                continue
            found.append(owner)
    return found


def passes_through(element):
    """La run continua oltre l'elemento: tratti (anche flessibili), raccordi e accessori con due connessioni."""
    if isinstance(element, DB.MEPCurve):
        return True
    category = element.Category
    if category is None or get_id(category.Id) not in INLINE_CATEGORIES:
        return False
    return len(physical_connectors(element)) <= 2


def search(seeds):
    """Ricerca a partire dagli elementi selezionati. Ritorna (elementi trovati, elementi dove la run si ferma)."""
    run_only = MODE == "run"
    found, stops = {}, {}
    queue = list(seeds)
    seed_ids = set(get_id(e.Id) for e in seeds)
    while queue:
        element = queue.pop()
        key = get_id(element.Id)
        if key in found or key in stops:
            continue
        if run_only and key not in seed_ids and not passes_through(element):
            stops[key] = element  # tee, apparecchiatura, terminale: fine della run, non selezionato
            continue
        found[key] = element
        if run_only and key in seed_ids and not passes_through(element):
            continue  # selezionato direttamente un tee o un'apparecchiatura: resta solo lui
        queue.extend(neighbours(element, run_only))
    return list(found.values()), list(stops.values())


def insulation_of(elements):
    """Isolamenti e rivestimenti interni degli elementi trovati."""
    extra = []
    for element in elements:
        category = category_key(element)[0]
        for getter, allowed in ((DB.InsulationLiningBase.GetInsulationIds, INSULATION_CATEGORIES),
                                (DB.InsulationLiningBase.GetLiningIds, LINING_CATEGORIES)):
            if category not in allowed:
                continue
            trace("insulation of {}".format(get_id(element.Id)))
            try:
                extra.extend(doc.GetElement(i) for i in getter(doc, element.Id))
            except Exception:
                continue
    return [e for e in extra if e is not None]


def category_key(element):
    category = element.Category
    if category is None:
        return None, "No category"
    return get_id(category.Id), category.Name


def get_seeds():
    selected = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    seeds = [e for e in selected if is_mep(e)]
    if seeds:
        return seeds, len(selected) - len(seeds)
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, MepFilter(),
                                           "{}: select the starting pipes, ducts, trays or fittings, then press "
                                           "Finish".format(TITLE))
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs], 0


def load_excluded(config):
    try:
        # chiave nuova: la vecchia "excluded" era stata salvata con la finestra a logica invertita (spunta = escludi)
        text = config.get_option("excluded_categories", "")
        return set(int(v) for v in str(text).split(",") if v.strip())
    except Exception:
        return set()


def save_excluded(config, excluded):
    try:
        config.excluded_categories = ",".join(str(v) for v in sorted(excluded))
        script.save_config()
    except Exception:
        pass


def print_report(seeds, ignored, elements, stops, excluded_names, counts):
    output.print_md("# {} report".format(TITLE))
    output.print_md("Starting elements: **{}**{} | selected: **{}**".format(
        len(seeds), " ({} selected elements without connectors ignored)".format(ignored) if ignored else "",
        sum(c[2] for c in counts if not c[3])))
    if MODE == "run" and stops:
        output.print_md("The run stops at: {}".format(", ".join(
            "{} {}".format(category_key(e)[1], output.linkify(e.Id)) for e in stops)))
    if excluded_names:
        output.print_md("Excluded categories: **{}**".format(", ".join(sorted(excluded_names))))
    rows = [[name, count, "excluded" if excluded else "selected"] for _, name, count, excluded in
            sorted(counts, key=lambda c: c[1].lower())]
    output.print_table(table_data=rows, columns=["Category", "Elements", "Status"])
    if not elements:
        output.print_md("Nothing was selected: every category found is excluded (change it with SHIFT + CLICK).")


def main():
    seeds, ignored = get_seeds()
    if not seeds:
        script.exit()
    trace("{}: {} starting elements: {}".format(TITLE, len(seeds), ", ".join(
        "{} ({})".format(get_id(e.Id), category_key(e)[1]) for e in seeds)), reset=True)
    found, stops = search(seeds)
    trace("search done: {} found, {} stops".format(len(found), len(stops)))
    found += insulation_of(found)
    trace("insulation done")

    counts = {}
    for element in found:
        key, name = category_key(element)
        counts.setdefault(key, [key, name, 0])[2] += 1

    config = script.get_config(CONFIG_NAME)
    excluded = load_excluded(config)
    if __shiftclick__:  # noqa: F821
        summary = "{} elements found from {} starting elements{}.".format(
            len(found), len(seeds), ", the run stops at {} branches or equipment".format(len(stops))
            if MODE == "run" else "")
        chosen = Select_ui.ExcludeWindow(script.get_bundle_file(XAML_FILE_NAME), TITLE, summary,
                                         [tuple(c) for c in counts.values()], excluded).show()
        if chosen is None:
            script.exit()
        # si tengono anche le esclusioni di categorie non presenti questa volta
        excluded = (excluded - set(counts)) | chosen
        save_excluded(config, excluded)

    elements = [e for e in found if category_key(e)[0] not in excluded]
    trace("selecting {} elements".format(len(elements)))
    uidoc.Selection.SetElementIds(List[DB.ElementId]([e.Id for e in elements]))
    trace("finished")

    table = [(key, name, count, key in excluded) for key, name, count in counts.values()]
    excluded_names = [name for key, name, _, is_excluded in table if is_excluded]
    if want_report(len(elements) or 1, 0 if elements else 1, "elements selected"):
        print_report(seeds, ignored, elements, stops, excluded_names, table)


main()
