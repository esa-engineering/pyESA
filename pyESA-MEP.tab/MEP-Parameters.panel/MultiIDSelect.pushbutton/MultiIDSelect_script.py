# -*- coding: utf-8 -*-
__title__ = "Multi ID\nSelect"
__doc__ = """Version = 0.2
Date    = 20.10.2026
________________________________________________________________
Select elements by Element Id, also in linked models, and list the
selected elements (links included) in a table.
Select by Id: paste the Ids and press Select. Each Id is searched in
this model and in every loaded link; all matches are selected and
listed in the table.
Load selection: reads what is selected in Revit now. The window stays
open, so you can select in Revit and load again.
The table shows model, Id, category, family and type, plus the
parameters ticked in the list (instance first, then type).
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"
__persistentengine__ = True     # finestra non modale con ExternalEvent

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
from System.Collections.Generic import List
from System.Windows import MessageBox
from Autodesk.Revit.DB import (ElementId, Element, Reference, RevitLinkInstance,
                               FilteredElementCollector, StorageType)
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

from pyrevit import revit, script

import MultiIDSelect_ui

XAML_FILE_NAME = "MultiIDSelect_form.xaml"
CONFIG_NAME = "ESA_MultiIDSelect"
TITLE = "Multi ID Select"


def get_element_id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


def element_name(element):
    try:
        return Element.Name.GetValue(element)
    except Exception:
        return getattr(element, "Name", "") or ""


def param_text(edoc, param):
    """Valore leggibile del parametro (come in Revit, con unita')."""
    try:
        if not param.HasValue:
            return ""
        st = param.StorageType
        if st == StorageType.String:
            return param.AsString() or ""
        text = param.AsValueString()
        if text:
            return text
        if st == StorageType.ElementId:
            target = edoc.GetElement(param.AsElementId())
            return element_name(target) if target is not None else ""
        if st == StorageType.Integer:
            return str(param.AsInteger())
        if st == StorageType.Double:
            return str(param.AsDouble())
    except Exception:
        pass
    return ""


def make_row(model, edoc, element):
    """Riga della tabella: dati fissi + tutti i parametri (istanza, poi tipo)."""
    row = {"model": model, "id": get_element_id_value(element.Id),
           "category": "", "family": "", "type": "", "params": {}}
    try:
        if element.Category is not None:
            row["category"] = element.Category.Name
    except Exception:
        pass
    element_type = None
    try:
        tid = element.GetTypeId()
        if tid is not None and get_element_id_value(tid) != -1:
            element_type = edoc.GetElement(tid)
    except Exception:
        pass
    if element_type is not None:
        row["family"] = getattr(element_type, "FamilyName", "") or ""
        row["type"] = element_name(element_type)
    else:
        row["type"] = element_name(element)
    for owner in (element, element_type):
        if owner is None:
            continue
        try:
            for p in owner.Parameters:
                name = p.Definition.Name
                if name not in row["params"]:
                    row["params"][name] = param_text(edoc, p)
        except Exception:
            pass
    return row


def host_label(doc):
    return "This model: {}".format(doc.Title)


def link_label(instance):
    return "Link: {}".format(element_name(instance))


def collect_sources(doc):
    """[(etichetta, documento, istanza link o None)]: modello aperto + link caricati."""
    sources = [(host_label(doc), doc, None)]
    for inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = inst.GetLinkDocument()
        if link_doc is not None:        # link non caricati esclusi
            sources.append((link_label(inst), link_doc, inst))
    return sources


def read_selection(uidoc):
    """Selezione corrente come righe della tabella, link inclusi."""
    doc = uidoc.Document
    rows = []
    try:
        refs = list(uidoc.Selection.GetReferences())
    except Exception:
        refs = None
    if refs is None:
        for eid in uidoc.Selection.GetElementIds():
            el = doc.GetElement(eid)
            if el is not None:
                rows.append(make_row(host_label(doc), doc, el))
        return rows
    for r in refs:
        linked_id = r.LinkedElementId
        if get_element_id_value(linked_id) == -1:
            el = doc.GetElement(r.ElementId)
            if el is not None:
                rows.append(make_row(host_label(doc), doc, el))
            continue
        inst = doc.GetElement(r.ElementId)
        link_doc = inst.GetLinkDocument() if inst is not None else None
        if link_doc is None:
            continue
        el = link_doc.GetElement(linked_id)
        if el is not None:
            rows.append(make_row(link_label(inst), link_doc, el))
    return rows


def find_element(edoc, n):
    try:
        return edoc.GetElement(ElementId(n))
    except Exception:
        return None


def select_everywhere(uidoc, ids):
    """Cerca ogni Id nel modello e in tutti i link, seleziona tutte le corrispondenze.
    Ritorna (righe, testo di stato)."""
    sources = collect_sources(uidoc.Document)
    found = []                          # (label, doc, inst, element, n)
    hits = {}
    for n in ids:
        for label, edoc, inst in sources:
            el = find_element(edoc, n)
            if el is not None:
                found.append((label, edoc, inst, el, n))
                hits.setdefault(n, []).append(label)
    missing = [n for n in ids if n not in hits]
    multiple = [n for n in ids if len(hits.get(n, [])) > 1]

    not_selectable = []
    selected = []
    if any(inst is not None for label, edoc, inst, el, n in found):
        if not hasattr(uidoc.Selection, "SetReferences"):
            return [], "This Revit version cannot select elements of linked models."
        refs = List[Reference]()
        for item in found:
            label, edoc, inst, el, n = item
            try:
                ref = Reference(el)
                refs.Add(ref.CreateLinkReference(inst) if inst is not None else ref)
                selected.append(item)
            except Exception:
                not_selectable.append(n)
        uidoc.Selection.SetReferences(refs)
    else:
        id_list = List[ElementId]()
        for item in found:
            id_list.Add(item[3].Id)
            selected.append(item)
        uidoc.Selection.SetElementIds(id_list)

    rows = [make_row(label, edoc, el) for label, edoc, inst, el, n in selected]
    parts = ["{} element(s) selected.".format(len(selected))]
    if missing:
        parts.append("Not found in any model: {}.".format(", ".join(str(i) for i in missing)))
    if multiple:
        parts.append("Found in more than one model (all selected): {}."
                     .format(", ".join(str(i) for i in multiple)))
    if not_selectable:
        parts.append("Not selectable in a link (no geometry, e.g. types or views): {}."
                     .format(", ".join(str(i) for i in sorted(set(not_selectable)))))
    return rows, "  ".join(parts)


class ApiHandler(IExternalEventHandler):
    """Esegue nel contesto API di Revit l'azione chiesta dalla finestra."""

    def __init__(self):
        self.action = None

    def Execute(self, uiapp):
        action, self.action = self.action, None
        if action is None:
            return
        try:
            uidoc = uiapp.ActiveUIDocument
            if uidoc is None:
                MessageBox.Show("No model is open.", TITLE)
                return
            action(uidoc)
        except Exception as ex:
            MessageBox.Show("Error: {}".format(ex), TITLE)

    def GetName(self):
        return TITLE


handler = ApiHandler()
api_event = ExternalEvent.Create(handler)
window = None


def run(action):
    handler.action = action
    api_event.Raise()


def do_load(uidoc):
    rows = read_selection(uidoc)
    window.set_rows(rows)
    window.set_status("Selection loaded: {} element(s).".format(len(rows)))


def do_select(ids):
    def action(uidoc):
        rows, status = select_everywhere(uidoc, ids)
        window.set_rows(rows)
        window.set_status(status)
    run(action)


xaml_path = script.get_bundle_file(XAML_FILE_NAME)
window = MultiIDSelect_ui.MultiIDSelectWindow(xaml_path, CONFIG_NAME, lambda: run(do_load), do_select)
window.set_rows(read_selection(revit.uidoc))       # siamo gia' nel contesto API
window.show()
