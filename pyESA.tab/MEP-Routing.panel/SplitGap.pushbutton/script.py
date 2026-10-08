# -*- coding: utf-8 -*-
__title__ = "Split\nGap"
__doc__ = """Version = 0.1
Date    = 08.10.2026
________________________________________________________________
Description:
Opens a gap in straight pipes, ducts, cable trays and conduits.

Instruction:
1. Select the elements (or select them first) and press Finish
2. Click two points: between them, each element is cut and the
   piece in between is removed. Each element is cut where the two
   points project on its axis
3. The report lists the elements, the gap length and what could not
   be cut. An element that cannot be cut is left unchanged

SHIFT + CLICK: vertical mode. Only vertical straight elements are processed
and the two points define the bottom and the top of the gap. Run it in a
section, elevation or 3D view (in 3D a temporary vertical work plane facing
the camera is used; the view must not look straight up or down).

Pipes and ducts are split with BreakCurve, so their insulation is
kept. Cable trays and conduits are split by copy.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import clr
import math
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB.Plumbing import Pipe, PlumbingUtils
from Autodesk.Revit.DB.Mechanical import Duct, MechanicalUtils
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

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
MIN_REMAINING_MM = 20.0  # lunghezza minima dei due tratti che restano
MIN_GAP_MM = 1.0
PHYSICAL_TYPES = (DB.ConnectorType.End, DB.ConnectorType.Curve, DB.ConnectorType.Physical)
CLASSES = (Pipe, Duct, CableTray, Conduit)


class GapError(Exception):
    """Errore atteso: il messaggio va nel report."""


class CurveFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, CLASSES)

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


def physical_refs(conn, element):
    return [r for r in conn.AllRefs if r.Owner.Id != element.Id and r.ConnectorType in PHYSICAL_TYPES]


def free_connector_near(element, point):
    manager = getattr(element, "ConnectorManager", None)
    if manager is None:
        model = getattr(element, "MEPModel", None)
        manager = model.ConnectorManager if model is not None else None
    best, best_dist = None, None
    for c in (manager.Connectors if manager is not None else []):
        if c.IsConnected:
            continue
        dist = c.Origin.DistanceTo(point)
        if best is None or dist < best_dist:
            best, best_dist = c, dist
    return best


def get_selection():
    selected = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    elements = [e for e in selected if isinstance(e, CLASSES)]
    if elements:
        return elements
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, CurveFilter(),
                                           "Select the pipes, ducts, cable trays or conduits, then press Finish")
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


VERTICAL_TOL_DEG = 5.0  # scostamento massimo dall'asse Z per considerare verticale un elemento


def get_vertical_elements(elements):
    """Tiene solo i tratti rettilinei verticali (asse entro VERTICAL_TOL_DEG dalla verticale)."""
    min_dz = math.cos(math.radians(VERTICAL_TOL_DEG))
    vertical = []
    for element in elements:
        curve = getattr(element.Location, "Curve", None)
        if isinstance(curve, DB.Line) and abs(curve.Direction.Z) >= min_dz:
            vertical.append(element)
    return vertical


def pick(prompt):
    try:
        return uidoc.Selection.PickPoint(prompt)
    except OperationCanceledException:
        script.exit()
    except Exception as ex:
        forms.alert("The point could not be picked in this view: {}".format(ex), title="Split Gap")
        script.exit()


def set_vertical_work_plane(view, element):
    """Vista 3D: imposta un piano di lavoro verticale per l'asse dell'elemento, rivolto verso la camera, cosi' i click
    hanno una quota Z corretta. La transazione resta aperta e va annullata (RollBack) dopo i click. Ritorna la
    transazione, o None se non serve."""
    if not isinstance(view, DB.View3D):
        return None
    view_dir = view.ViewDirection
    horizontal = DB.XYZ(view_dir.X, view_dir.Y, 0.0)
    if horizontal.GetLength() < 0.05:
        forms.alert("The view looks straight down or up: rotate the 3D view so that the vertical elements are seen "
                    "from the side.", title="Split Gap")
        script.exit()
    origin = element.Location.Curve.GetEndPoint(0)
    plane = DB.Plane.CreateByNormalAndOrigin(horizontal.Normalize(), origin)
    transaction = DB.Transaction(doc, "Split gap work plane")
    transaction.Start()
    try:
        view.SketchPlane = DB.SketchPlane.Create(doc, plane)
    except Exception:
        transaction.RollBack()
        raise
    return transaction


def midpoint(element):
    return element.Location.Curve.Evaluate(0.5, True)


def cut_by_copy(element, a, b, p1, p2):
    """Passerelle e conduit: l'originale diventa a-p1, una copia p2-b; il vuoto resta tra i due."""
    far_conn = nearest_connector(element, b)
    far_refs = physical_refs(far_conn, element)
    far_info = [(r.Owner.Id, r.Origin) for r in far_refs]
    for ref in far_refs:
        far_conn.DisconnectFrom(ref)
    tail_id = list(DB.ElementTransformUtils.CopyElement(doc, element.Id, DB.XYZ.Zero))[0]
    tail = doc.GetElement(tail_id)
    element.Location.Curve = DB.Line.CreateBound(a, p1)
    tail.Location.Curve = DB.Line.CreateBound(p2, b)
    doc.Regenerate()
    conn = nearest_connector(tail, b)
    for owner_id, origin in far_info:
        target = free_connector_near(doc.GetElement(owner_id), origin)
        if target is None:
            raise GapError("No free connector found on the element connected to the far end")
        conn.ConnectTo(target)


ON_CURVE_TOL_FT = 1e-4  # circa 0,03 mm


def on_curve(element, point):
    """Il punto sta sull'asse dell'elemento (entro tolleranza, estremi compresi)."""
    return element.Location.Curve.Distance(point) < ON_CURVE_TOL_FT


def cut_by_break(element, p1, p2):
    """Tubi e canali: BreakCurve nei due punti, il pezzo di mezzo si elimina (isolamento compreso)."""
    utils = PlumbingUtils if isinstance(element, Pipe) else MechanicalUtils
    pieces = [element]
    pieces.append(doc.GetElement(utils.BreakCurve(doc, element.Id, p2)))
    doc.Regenerate()
    # dopo la prima rottura l'originale puo' essere la parte prima o dopo p2: si rompe il pezzo che contiene p1
    holder = next((e for e in pieces if on_curve(e, p1)), None)
    if holder is None:
        raise GapError("After the first cut, no piece contains the first point of the gap")
    pieces.append(doc.GetElement(utils.BreakCurve(doc, holder.Id, p1)))
    doc.Regenerate()
    mid = p1.Add(p2).Multiply(0.5)
    middle = min(pieces, key=lambda e: midpoint(e).DistanceTo(mid))
    if middle.Id == element.Id or midpoint(middle).DistanceTo(mid) > p1.DistanceTo(p2) / 2.0 + 1e-4:
        raise GapError("The piece between the two points could not be identified")
    doc.Delete(middle.Id)


def diagnostics(element, click1, click2):
    """Dati geometrici per capire perche' un taglio non riesce (distanze in mm lungo l'asse dell'elemento)."""
    try:
        curve = element.Location.Curve
        a, d = curve.GetEndPoint(0), curve.Direction
        t = [click.Subtract(a).DotProduct(d) * MM_PER_FT for click in (click1, click2)]
        return ("length {:.0f} mm | direction ({:.2f}, {:.2f}, {:.2f}) | clicks at {:.0f} and {:.0f} mm along the "
                "axis | clicks Z {:.0f} / {:.0f} mm, element Z {:.0f} mm".format(
                    curve.Length * MM_PER_FT, d.X, d.Y, d.Z, t[0], t[1], click1.Z * MM_PER_FT,
                    click2.Z * MM_PER_FT, a.Z * MM_PER_FT))
    except Exception as ex:
        return "geometry not readable: {}".format(ex)


def make_gap(element, click1, click2):
    """Taglia un elemento. Ritorna la lunghezza del vuoto in mm."""
    curve = element.Location.Curve
    if not isinstance(curve, DB.Line):
        raise GapError("Only straight elements are supported")
    a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
    d = curve.Direction
    t = sorted((click1.Subtract(a).DotProduct(d), click2.Subtract(a).DotProduct(d)))
    min_ft = MIN_REMAINING_MM / MM_PER_FT
    if t[1] - t[0] < MIN_GAP_MM / MM_PER_FT:
        raise GapError("The gap would be shorter than {:g} mm".format(MIN_GAP_MM))
    if t[0] < min_ft or curve.Length - t[1] < min_ft:
        raise GapError("The points are outside the element, or closer than {:.0f} mm to its ends".format(
            MIN_REMAINING_MM))
    p1, p2 = a.Add(d.Multiply(t[0])), a.Add(d.Multiply(t[1]))
    if isinstance(element, (Pipe, Duct)):
        cut_by_break(element, p1, p2)
    else:
        cut_by_copy(element, a, b, p1, p2)
    return (t[1] - t[0]) * MM_PER_FT


def main():
    elements = get_selection()
    if not elements:
        script.exit()
    
    vertical_mode = __shiftclick__  # noqa: F821
    if vertical_mode:
        elements = get_vertical_elements(elements)
        if not elements:
            forms.alert("No vertical straight element in the selection.", title="Split Gap")
            script.exit()

    if vertical_mode:
        prompt_1, prompt_2 = "Click the bottom of the gap", "Click the top of the gap"
    else:
        prompt_1, prompt_2 = "Click the first point of the gap", "Click the second point of the gap"
    plane_transaction = None
    try:
        if vertical_mode:
            plane_transaction = set_vertical_work_plane(uidoc.ActiveView, elements[0])
        click1 = pick(prompt_1)
        click2 = pick(prompt_2)
    finally:
        if plane_transaction is not None:
            plane_transaction.RollBack()  # il piano di lavoro serve solo per i click, nel modello non resta nulla
    if vertical_mode and abs(click1.Z - click2.Z) < MIN_GAP_MM / MM_PER_FT:
        forms.alert("The two points have the same height. For vertical elements run the command in a section, "
                    "elevation or 3D view and click the two heights on the element.", title="Split Gap")
        script.exit()

    rows, done = [], 0
    for element in elements:
        info = diagnostics(element, click1, click2)
        transaction = DB.Transaction(doc, "Split gap")
        transaction.Start()
        try:
            gap = make_gap(element, click1, click2)
            transaction.Commit()
            rows.append((element, "{:.0f}".format(gap), "Gap created", info))
            done += 1
        except GapError as ex:
            transaction.RollBack()
            rows.append((element, "", "Unchanged. {}".format(ex), info))
        except Exception as ex:
            transaction.RollBack()
            rows.append((element, "", "Unchanged. The gap could not be created: {}".format(ex), info))
    if not want_report(len(elements), len(elements) - done, "gaps created"):
        return
    output.print_md("# Split Gap report")
    output.print_md("Elements: **{}** | gaps created: **{}** | left unchanged: **{}**".format(
        len(elements), done, len(elements) - done))
    # righe in testo, non in tabella: si vedono sempre, anche se la tabella non le mostra
    for i, (element, gap, status, info) in enumerate(rows, 1):
        output.print_md("**#{}** {} {} | gap [mm]: {} | **{}**".format(
            i, category_name(element), output.linkify(element.Id), gap or "-", status))
        output.print_md("&nbsp;&nbsp;&nbsp;{}".format(info))


main()
