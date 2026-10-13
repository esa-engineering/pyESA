# -*- coding: utf-8 -*-
__title__ = "Create\nSimilar"
__doc__ = """Version = 0.1
Date    = 08.10.2026
________________________________________________________________
Description:
Creates new instances that carry the parameters of an existing one.

CLICK: new instances of the same type as the source element.
SHIFT + CLICK: choose another type or family of the same category,
the new instances take it and keep the parameters of the source.

Instruction:
1. Select one element beforehand, or run the tool and click it
2. Click in the view to place each new instance, press ESC to finish
   (point-based elements: one click each; straight pipes, ducts, cable
   trays and conduits: two clicks, start and end)

The new instance is a copy of the source, so all its parameters come along.
With SHIFT the type is then changed: parameters that do not exist in the
new family are lost. Use plan views: only X and Y of the click are used,
the elevation of the source is kept.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Collections.Generic import List

from pyrevit import revit, script, DB, forms

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

MM_PER_FT = 304.8
MIN_LENGTH_FT = 1.0 / MM_PER_FT   # lunghezza minima di un elemento lineare (1 mm)
HORIZONTAL_TOL = 1e-3
POINT, LINEAR = "point", "linear"


class SimilarError(Exception):
    """Errore atteso: il messaggio va nel report."""


class SourceFilter(ISelectionFilter):
    def AllowElement(self, element):
        return element.Category is not None and kind_of(element) is not None

    def AllowReference(self, reference, point):
        return False


def kind_of(element):
    location = element.Location
    if isinstance(location, DB.LocationPoint):
        return POINT
    if isinstance(location, DB.LocationCurve):
        return LINEAR
    return None


def type_name(element_type):
    family = getattr(element_type, "FamilyName", "")
    name = DB.Element.Name.GetValue(element_type)
    return "{} : {}".format(family, name) if family else name


def get_source():
    """Elemento di partenza: selezione corrente se presente, altrimenti click."""
    ids = list(uidoc.Selection.GetElementIds())
    if len(ids) > 1:
        forms.alert("Select a single element, or none and click it.", title="Create Similar")
        script.exit()
    if ids:
        element = doc.GetElement(ids[0])
        if element.Category is None or kind_of(element) is None:
            raise SimilarError("The selected element has no location that can be copied")
        return element
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, SourceFilter(),
                                         "Click the element to create similar instances from")
    except OperationCanceledException:
        script.exit()
    return doc.GetElement(ref.ElementId)


def choose_type(source):
    """SHIFT: tipo scelto tra quelli validi per l'elemento, stessa categoria."""
    current = source.GetTypeId()
    candidates = []
    for element_type in DB.FilteredElementCollector(doc).WhereElementIsElementType() \
            .OfCategoryId(source.Category.Id):
        if source.IsValidType(element_type.Id):
            candidates.append(element_type)
    labels = {}
    for element_type in candidates:
        label = type_name(element_type)
        if element_type.Id == current:
            label += "  (current)"
        labels[label] = element_type
    if len(labels) < 2:
        raise SimilarError("No other type of category {} can be used for this element".format(source.Category.Name))
    picked = forms.SelectFromList.show(sorted(labels.keys()), title="New type ({})".format(source.Category.Name),
                                       button_name="Use this type", multiselect=False)
    if not picked:
        script.exit()
    return labels[picked]


def pick_point(prompt):
    """Punto cliccato, None se l'utente preme ESC."""
    try:
        return uidoc.Selection.PickPoint(prompt)
    except OperationCanceledException:
        return None


def source_geometry(source, kind):
    """Punto di riferimento della sorgente e, per i lineari, se si traccia con due click."""
    if kind == POINT:
        return source.Location.Point, False
    curve = source.Location.Curve
    if not isinstance(curve, DB.Line):
        raise SimilarError("Only straight linear elements are supported")
    horizontal = abs(curve.Direction.Z) < HORIZONTAL_TOL
    return curve.GetEndPoint(0), horizontal


def ask_placement(source, kind, origin, two_clicks):
    """Click dell'utente. Ritorna (vettore di copia, nuova curva o None) oppure None per terminare."""
    first = pick_point("Click the location of the new instance (ESC to finish)" if kind == POINT else
                       "Click the start of the new element (ESC to finish)")
    if first is None:
        return None
    start = DB.XYZ(first.X, first.Y, origin.Z)
    move = start.Subtract(origin)
    if kind == POINT:
        return move, None
    curve = source.Location.Curve
    if two_clicks:
        second = pick_point("Click the end of the new element (ESC to finish)")
        if second is None:
            return None
        end = DB.XYZ(second.X, second.Y, origin.Z)
        if start.DistanceTo(end) < MIN_LENGTH_FT:
            raise SimilarError("Start and end are too close")
    else:
        # sorgente non orizzontale: stesso vettore (lunghezza e inclinazione) della sorgente
        end = start.Add(curve.GetEndPoint(1).Subtract(curve.GetEndPoint(0)))
    return move, DB.Line.CreateBound(start, end)


def create_instance(source, move, new_curve, new_type):
    """Copia la sorgente, la sposta e se serve cambia il tipo. Ritorna il nuovo elemento."""
    transaction = DB.Transaction(doc, "Create Similar")
    transaction.Start()
    try:
        if new_type is not None and isinstance(new_type, DB.FamilySymbol) and not new_type.IsActive:
            new_type.Activate()
            doc.Regenerate()
        ids = DB.ElementTransformUtils.CopyElement(doc, source.Id, move)
        copies = [doc.GetElement(i) for i in ids]
        new = next((e for e in copies if e.Category is not None and e.Category.Id == source.Category.Id), copies[0])
        if new_curve is not None:
            new.Location.Curve = new_curve
        if new_type is not None:
            changed = new.ChangeTypeId(new_type.Id)
            if changed != DB.ElementId.InvalidElementId:
                new = doc.GetElement(changed)
        transaction.Commit()
        return new
    except Exception:
        transaction.RollBack()
        raise


def print_report(source, new_type, created, errors):
    output.print_md("# Create Similar report")
    output.print_md("Source: {} | category **{}** | type **{}**".format(
        output.linkify(source.Id), source.Category.Name, type_name(doc.GetElement(source.GetTypeId()))))
    if new_type is not None:
        output.print_md("New type: **{}**".format(type_name(new_type)))
    output.print_md("New instances: **{}** | not created: **{}**".format(len(created), len(errors)))
    if created:
        output.print_table(table_data=[[i + 1, output.linkify(e.Id)] for i, e in enumerate(created)],
                           columns=["#", "New instance"])
    if errors:
        output.print_md("**Not created**")
        output.print_table(table_data=[[i + 1, msg] for i, msg in enumerate(errors)], columns=["#", "Reason"])
    if not created and not errors:
        output.print_md("No instance was placed: the command was cancelled before any click.")


def main():
    output.close_others()
    try:
        source = get_source()
        kind = kind_of(source)
        origin, two_clicks = source_geometry(source, kind)
        new_type = choose_type(source) if __shiftclick__ else None  # noqa: F821
    except SimilarError as ex:
        forms.alert(str(ex), title="Create Similar")
        script.exit()

    created, errors = [], []
    while True:
        try:
            placement = ask_placement(source, kind, origin, two_clicks)
        except SimilarError as ex:
            errors.append(str(ex))
            continue
        if placement is None:
            break
        try:
            created.append(create_instance(source, placement[0], placement[1], new_type))
        except Exception as ex:
            errors.append("{}: {}".format(type(ex).__name__, ex))
    if created:
        try:
            uidoc.Selection.SetElementIds(List[DB.ElementId]([e.Id for e in created]))
        except Exception:
            pass
    if want_report(len(created) + len(errors), len(errors), "instances created"):
        print_report(source, new_type, created, errors)


main()
