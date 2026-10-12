# -*- coding: utf-8 -*-
__title__ = "Quick\nExtend"
__doc__ = """Version = 0.1
Date    = 12.10.2026
________________________________________________________________
Description:
Extends a straight pipe, duct, cable tray or conduit up to a picked
point, one element at a time until ESC. No selection and no Finish.

CLICK:
1. Click the element near the end to extend
2. Click the target point: the end moves along the element axis to
   the point (the point is projected onto the axis). A point inside
   the element shortens it instead
3. Repeat on the next element; ESC on the first click finishes

SHIFT + CLICK:
Select several elements (or pre-select them), then click one point:
each element is extended along its own axis, from the end closer to
the point, so that parallel elements end aligned on the same line.

In plans, sections and elevations the target point uses Revit snaps
(endpoints, midpoints, intersections, perpendicular, centers,
nearest); in plans only XY counts and the slope is kept. In 3D views
the target point is clicked on the face of any element.
Only a free end can be moved: an end connected to a fitting or to
another element is left unchanged. Insulation and lining follow the
element. The report lists each element and the length change.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import clr
import math
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
from Autodesk.Revit.UI.Selection import ObjectType, ObjectSnapTypes, ISelectionFilter
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
MIN_LENGTH_MM = 20.0  # lunghezza minima dell'elemento dopo la modifica
MIN_CHANGE_MM = 0.5  # sotto questa variazione l'elemento si considera gia' al punto
VERTICAL_TOL_DEG = 5.0  # scostamento massimo dall'asse Z per considerare verticale un elemento
PHYSICAL_TYPES = (DB.ConnectorType.End, DB.ConnectorType.Curve, DB.ConnectorType.Physical)
CLASSES = (Pipe, Duct, CableTray, Conduit)
SNAPS = (ObjectSnapTypes.Endpoints | ObjectSnapTypes.Midpoints | ObjectSnapTypes.Nearest |
         ObjectSnapTypes.Intersections | ObjectSnapTypes.Perpendicular | ObjectSnapTypes.Centers)


class ExtendError(Exception):
    """Errore atteso: il messaggio va nel report."""


class CurveFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, CLASSES) and isinstance(getattr(element.Location, "Curve", None), DB.Line)

    def AllowReference(self, reference, point):
        return False


def category_name(element):
    for cls, name in ((Pipe, "Pipe"), (Duct, "Duct"), (CableTray, "Cable Tray"), (Conduit, "Conduit")):
        if isinstance(element, cls):
            return name
    return "?"


def nearest_connector(element, point):
    best, best_dist = None, None
    for c in element.ConnectorManager.Connectors:
        if c.ConnectorType not in PHYSICAL_TYPES:
            continue
        dist = c.Origin.DistanceTo(point)
        if best is None or dist < best_dist:
            best, best_dist = c, dist
    return best


def connected_to(conn, element):
    """Descrizione degli elementi collegati fisicamente al connettore (vuota se il connettore e' libero)."""
    if conn is None or not conn.IsConnected:
        return []
    return ["{} {}".format(r.Owner.Category.Name if r.Owner.Category else "element", r.Owner.Id)
            for r in conn.AllRefs if r.Owner.Id != element.Id and r.ConnectorType in PHYSICAL_TYPES]


def is_vertical(element):
    curve = element.Location.Curve
    return abs(curve.Direction.Z) >= math.cos(math.radians(VERTICAL_TOL_DEG))


def axis_param(element, point):
    """Posizione del punto lungo l'asse dell'elemento (piedi dall'estremo 0). In pianta, per un elemento non
    verticale, la Z del clic non conta: si proietta in XY e la pendenza resta quella dell'elemento."""
    curve = element.Location.Curve
    a, d = curve.GetEndPoint(0), curve.Direction
    if isinstance(doc.ActiveView, DB.ViewPlan) and not is_vertical(element):
        dxy2 = d.X ** 2 + d.Y ** 2
        return ((point.X - a.X) * d.X + (point.Y - a.Y) * d.Y) / dxy2
    return point.Subtract(a).DotProduct(d)


def end_toward(element, point):
    """Indice (0/1) dell'estremo dalla parte del punto lungo l'asse."""
    return 0 if axis_param(element, point) < element.Location.Curve.Length / 2.0 else 1


def extend(element, end_index, target):
    """Sposta l'estremo end_index lungo l'asse fino alla proiezione del punto. Ritorna la variazione di lunghezza in
    mm (positiva = allungato)."""
    curve = element.Location.Curve
    if not isinstance(curve, DB.Line):
        raise ExtendError("Only straight elements are supported")
    if isinstance(doc.ActiveView, DB.ViewPlan) and is_vertical(element):
        raise ExtendError("A vertical element cannot be extended in a plan view: use a section, elevation or 3D view")
    a, b, d = curve.GetEndPoint(0), curve.GetEndPoint(1), curve.Direction
    length = curve.Length
    t = axis_param(element, target)
    new_length = t if end_index == 1 else length - t
    if new_length < MIN_LENGTH_MM / MM_PER_FT:
        raise ExtendError("The point is beyond the other end, or the element would be shorter than {:.0f} mm".format(
            MIN_LENGTH_MM))
    change = (new_length - length) * MM_PER_FT
    if abs(change) < MIN_CHANGE_MM:
        raise ExtendError("The end is already at the point")
    moving_end = b if end_index == 1 else a
    others = connected_to(nearest_connector(element, moving_end), element)
    if others:
        raise ExtendError("The end is connected to {}: only a free end can be moved".format(", ".join(others)))
    new_end = a.Add(d.Multiply(t))
    element.Location.Curve = DB.Line.CreateBound(a, new_end) if end_index == 1 else DB.Line.CreateBound(new_end, b)
    doc.Regenerate()
    return change


def pick_target(prompt):
    """Punto di arrivo. Pianta/sezione/prospetto: PickPoint con snap. Vista 3D (niente piano di lavoro): punto sulla
    faccia di un elemento qualsiasi. OperationCanceledException passa al chiamante."""
    if isinstance(doc.ActiveView, DB.View3D):
        ref = uidoc.Selection.PickObject(ObjectType.PointOnElement, prompt + " (on the face of an element)")
        return ref.GlobalPoint
    return uidoc.Selection.PickPoint(SNAPS, prompt)


def apply(element, end_index, target, rows):
    """Una transazione per elemento: quello che non riesce resta com'era. Ritorna True se fatto."""
    transaction = DB.Transaction(doc, "Quick Extend")
    transaction.Start()
    try:
        if target is None:
            raise ExtendError("The clicked point could not be read in this view")
        change = extend(element, end_index, target)
        transaction.Commit()
        rows.append((element, "{:+.0f}".format(change), "Extended" if change > 0 else "Shortened"))
        return True
    except ExtendError as ex:
        transaction.RollBack()
        rows.append((element, "", "Unchanged. {}".format(ex)))
    except Exception as ex:
        transaction.RollBack()
        rows.append((element, "", "Unchanged. The element could not be extended: {}".format(ex)))
    return False


def report(title, rows, done):
    if not want_report(len(rows), len(rows) - done, "elements extended"):
        return
    output.print_md("# {}".format(title))
    output.print_md("Elements: **{}** | changed: **{}** | left unchanged: **{}**".format(
        len(rows), done, len(rows) - done))
    # righe in testo, non in tabella: si vedono sempre
    for i, (element, change, status) in enumerate(rows, 1):
        output.print_md("**#{}** {} {} | length change [mm]: {} | **{}**".format(
            i, category_name(element), output.linkify(element.Id), change or "-", status))


def quick_main():
    """CLICK: un elemento alla volta, fino a ESC. Primo clic sull'elemento vicino all'estremo da estendere, secondo
    clic sul punto di arrivo."""
    rows, done = [], 0
    while True:
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, CurveFilter(),
                                             "Quick Extend: click the element near the end to extend (ESC to finish)")
        except OperationCanceledException:
            break
        element = doc.GetElement(ref.ElementId)
        end_index = end_toward(element, ref.GlobalPoint) if ref.GlobalPoint is not None else 1
        uidoc.Selection.SetElementIds(List[DB.ElementId]([element.Id]))  # si vede quale elemento e' stato preso
        try:
            target = pick_target("Quick Extend: click the point to extend the highlighted element to (ESC to skip it)")
        except OperationCanceledException:
            uidoc.Selection.SetElementIds(List[DB.ElementId]())
            continue
        except Exception as ex:
            uidoc.Selection.SetElementIds(List[DB.ElementId]())
            forms.alert("The point could not be picked in this view: {}".format(ex), title="Quick Extend")
            break
        uidoc.Selection.SetElementIds(List[DB.ElementId]())
        if apply(element, end_index, target, rows):
            done += 1
    report("Quick Extend report", rows, done)


def get_selection():
    selected = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    elements = [e for e in selected if CurveFilter().AllowElement(e)]
    if elements:
        return elements
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, CurveFilter(),
                                           "Quick Extend: select the elements to extend, then press Finish")
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def multi_main():
    """SHIFT + CLICK: piu' elementi, un solo punto. Ogni elemento si estende lungo il proprio asse dall'estremo piu'
    vicino al punto."""
    elements = get_selection()
    if not elements:
        script.exit()
    try:
        target = pick_target("Quick Extend: click the point to extend the {} elements to".format(len(elements)))
    except OperationCanceledException:
        script.exit()
    except Exception as ex:
        forms.alert("The point could not be picked in this view: {}".format(ex), title="Quick Extend")
        script.exit()
    rows, done = [], 0
    for element in elements:
        end_index = end_toward(element, target) if target is not None else 1
        if apply(element, end_index, target, rows):
            done += 1
    report("Quick Extend report (several elements, one point)", rows, done)


if __shiftclick__:  # noqa: F821
    multi_main()
else:
    quick_main()
