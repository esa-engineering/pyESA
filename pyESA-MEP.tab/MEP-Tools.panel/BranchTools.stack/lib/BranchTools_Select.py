# -*- coding: utf-8 -*-
"""Selezione e impostazioni comuni ai comandi di Branch Tools.
Nessun documento a livello di modulo (Rocket mode): doc e uidoc arrivano come argomenti."""
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

import BranchTools_Rotate as fr


class PredicateFilter(ISelectionFilter):
    """Accetta gli elementi per cui accept(elemento) e' vero, escluso exclude_id."""

    def __init__(self, accept, exclude_id=None):
        self.accept = accept
        self.exclude = fr.get_element_id_value(exclude_id) if exclude_id is not None else None

    def AllowElement(self, element):
        if self.exclude is not None and fr.get_element_id_value(element.Id) == self.exclude:
            return False
        return self.accept(element)

    def AllowReference(self, reference, point):
        return False


def pick(uidoc, prompt, accept, exclude_id=None):
    """(elemento, punto cliccato o None); (None, None) se Esc."""
    try:
        reference = uidoc.Selection.PickObject(ObjectType.Element, PredicateFilter(accept, exclude_id), prompt)
    except OperationCanceledException:
        return None, None
    return uidoc.Document.GetElement(reference.ElementId), reference.GlobalPoint


def get_element(uidoc, prompt, accept):
    """Un solo elemento: quello preselezionato se e' uno solo e va bene, altrimenti uno cliccato.
    Ritorna (elemento, punto cliccato o None), o (None, None) se Esc."""
    doc = uidoc.Document
    ids = list(uidoc.Selection.GetElementIds())
    if len(ids) == 1:
        element = doc.GetElement(ids[0])
        if accept(element):
            return element, None
    return pick(uidoc, prompt, accept)


def load_settings(config, numbers, texts=None):
    """numbers: {chiave: (opzione, default numerico)}; texts: {chiave: opzione}. Valori mancanti o illeggibili
    tornano al default; i numeri accettano anche la virgola decimale."""
    settings = {}
    for key, (option, default) in numbers.items():
        try:
            settings[key] = float(str(config.get_option(option, default)).replace(",", "."))
        except (TypeError, ValueError):
            settings[key] = default
    for key, option in (texts or {}).items():
        try:
            settings[key] = config.get_option(option, "") or ""
        except Exception:
            settings[key] = ""
    return settings


def save_settings(script, config, settings, numbers, texts=None):
    """Salva le chiavi presenti in settings; in silenzio se il file di configurazione non e' scrivibile."""
    try:
        options = [(key, option) for key, (option, _) in numbers.items()] + list((texts or {}).items())
        for key, option in options:
            if key in settings:
                config.set_option(option, settings[key])
        script.save_config()
    except Exception:
        pass
