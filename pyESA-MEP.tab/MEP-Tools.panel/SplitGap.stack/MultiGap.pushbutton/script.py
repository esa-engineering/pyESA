# -*- coding: utf-8 -*-
__title__ = "Multi\nGap"
__doc__ = """Version = 0.3
Date    = 16.10.2026
________________________________________________________________
Description:
Opens a gap in straight pipes, ducts, cable trays and conduits.

1. Select the elements (or select them first) and press Finish
2. Click two points: between them, each element is cut and the
   piece in between is removed. All the cuts lie on the same two
   planes, perpendicular to the mean direction of the elements in
   plan, so the cut is clean even if the elements are sloped or not
   exactly parallel (the height of the clicks does not matter). Where
   a point is beyond the end of an element, the whole rest of it is
   removed, up to that end
3. The report lists the elements, the gap length and what could not
   be cut. An element that cannot be cut is left unchanged
Vertical or not is detected from the selection. Vertical elements: the
two points are the bottom and the top of the gap; run it in a section,
elevation or 3D view (in 3D a temporary vertical work plane facing the
camera is used; the view must not look straight up or down). A
selection mixing vertical and non-vertical elements is refused.

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
        forms.alert("The point could not be picked in this view: {}".format(ex), title="Multi Gap")
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
                    "from the side.", title="Multi Gap")
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


def cut_reference(elements, vertical_mode):
    """Riferimento dei piani di taglio, uguale per tutti gli elementi: None per i verticali (piani orizzontali, conta
    solo la quota dei click), altrimenti la direzione media in pianta degli elementi (unitaria, concorde con il
    primo). I tagli di tutti gli elementi stanno cosi' sullo stesso piano, anche se non sono esattamente paralleli."""
    if vertical_mode:
        return None
    total, first = DB.XYZ.Zero, None
    for element in elements:
        d = element.Location.Curve.Direction
        flat = DB.XYZ(d.X, d.Y, 0)
        if flat.GetLength() < 1e-6:
            continue
        flat = flat.Normalize()
        if first is None:
            first = flat
        elif flat.DotProduct(first) < 0:
            flat = flat.Negate()
        total = total.Add(flat)
    if first is None or total.GetLength() < 1e-6:
        raise GapError("The elements have no direction in plan")
    return total.Normalize()


def axis_positions(element, click1, click2, ref):
    """Posizione dei due tagli lungo l'asse dell'elemento (piedi dall'estremo 0, ordinate). Ogni taglio e' il punto
    dell'asse che sta sul piano passante per il click: perpendicolare a ref in pianta (Z del click ignorata, la quota
    della vista non c'entra), oppure orizzontale per gli elementi verticali."""
    curve = element.Location.Curve
    a, d = curve.GetEndPoint(0), curve.Direction
    positions = []
    for click in (click1, click2):
        if ref is None:
            if abs(d.Z) < 1e-6:
                raise GapError("The element is horizontal: its position along the axis cannot be read from a height")
            positions.append((click.Z - a.Z) / d.Z)
        else:
            slope = d.X * ref.X + d.Y * ref.Y
            if abs(slope) < 1e-6:
                raise GapError("The element is perpendicular to the direction of the cut")
            positions.append(((click.X - a.X) * ref.X + (click.Y - a.Y) * ref.Y) / slope)
    return sorted(positions)


def diagnostics(element, click1, click2, t):
    """Dati geometrici per capire perche' un taglio non riesce (distanze in mm lungo l'asse dell'elemento)."""
    try:
        curve = element.Location.Curve
        a, d = curve.GetEndPoint(0), curve.Direction
        return ("length {:.0f} mm | direction ({:.2f}, {:.2f}, {:.2f}) | cuts at {:.0f} and {:.0f} mm along the "
                "axis | clicks Z {:.0f} / {:.0f} mm, element Z {:.0f} mm".format(
                    curve.Length * MM_PER_FT, d.X, d.Y, d.Z, t[0] * MM_PER_FT, t[1] * MM_PER_FT,
                    click1.Z * MM_PER_FT, click2.Z * MM_PER_FT, a.Z * MM_PER_FT))
    except Exception as ex:
        return "geometry not readable: {}".format(ex)


def make_gap(element, t):
    """Taglia un elemento nelle posizioni t (due valori ordinati, vedi axis_positions). Ritorna (lunghezza eliminata
    in mm, stato per il report). Se un punto cade oltre un estremo, o a meno di MIN_REMAINING_MM da esso, il vuoto
    arriva fino a quell'estremo: si elimina tutto il tratto restante e quanto era collegato a quell'estremo resta,
    scollegato."""
    curve = element.Location.Curve
    if not isinstance(curve, DB.Line):
        raise GapError("Only straight elements are supported")
    a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
    d = curve.Direction
    length = curve.Length
    min_ft = MIN_REMAINING_MM / MM_PER_FT
    if t[1] - t[0] < MIN_GAP_MM / MM_PER_FT:
        raise GapError("The gap would be shorter than {:g} mm".format(MIN_GAP_MM))
    if t[1] <= 0 or t[0] >= length:
        raise GapError("The gap is outside the element")
    to_start, to_end = t[0] < min_ft, length - t[1] < min_ft
    if to_start and to_end:
        raise GapError("The gap covers the whole element (or leaves less than {:.0f} mm): delete the element "
                       "instead".format(MIN_REMAINING_MM))
    if to_start or to_end:
        # t_side oltre l'estremo: indica solo da che parte eliminare
        removed, disconnected = trim(element, t[0], length + 1.0) if to_end else trim(element, t[1], -1.0)
        status = "Rest of the element removed up to its {} end".format("second" if to_end else "first")
        if disconnected:
            status += ". Disconnected from: {}".format(", ".join(disconnected))
        return removed, status
    p1, p2 = a.Add(d.Multiply(t[0])), a.Add(d.Multiply(t[1]))
    if isinstance(element, (Pipe, Duct)):
        cut_by_break(element, p1, p2)
    else:
        cut_by_copy(element, a, b, p1, p2)
    return (t[1] - t[0]) * MM_PER_FT, "Gap created"


def axis_param(element, point):
    """Posizione del punto lungo l'asse dell'elemento (piedi dall'estremo 0). In pianta, per un elemento non
    verticale, la Z del clic non conta: si proietta in XY."""
    curve = element.Location.Curve
    a, d = curve.GetEndPoint(0), curve.Direction
    if isinstance(doc.ActiveView, DB.ViewPlan) and not get_vertical_elements([element]):
        dxy = math.sqrt(d.X ** 2 + d.Y ** 2)
        return ((point.X - a.X) * d.X + (point.Y - a.Y) * d.Y) / (dxy ** 2)  # parametro lungo l'asse 3D
    return point.Subtract(a).DotProduct(d)


def trim(element, t_cut, t_side):
    """Taglia l'elemento in t_cut ed elimina la parte dal taglio all'estremo dalla parte di t_side. Ritorna
    (lunghezza eliminata in mm, elementi scollegati). Quanto era collegato a quell'estremo resta, scollegato."""
    curve = element.Location.Curve
    if not isinstance(curve, DB.Line):
        raise GapError("Only straight elements are supported")
    a, b, d = curve.GetEndPoint(0), curve.GetEndPoint(1), curve.Direction
    length = curve.Length
    if abs(t_side - t_cut) < MIN_GAP_MM / MM_PER_FT:
        raise GapError("The second click is on the cut point: no side to delete")
    if t_cut <= 0 or t_cut >= length:
        raise GapError("The cut point is outside the element")
    delete_after = t_side > t_cut  # True: si elimina dal taglio verso l'estremo 1
    kept = t_cut if delete_after else length - t_cut
    if kept < MIN_REMAINING_MM / MM_PER_FT:
        raise GapError("The part that stays would be shorter than {:.0f} mm".format(MIN_REMAINING_MM))
    lost_end = b if delete_after else a
    conn = nearest_connector(element, lost_end)
    refs = physical_refs(conn, element)
    disconnected = ["{} {}".format(r.Owner.Category.Name if r.Owner.Category else "element", r.Owner.Id)
                    for r in refs]
    for ref in refs:
        conn.DisconnectFrom(ref)
    cut = a.Add(d.Multiply(t_cut))
    element.Location.Curve = DB.Line.CreateBound(a, cut) if delete_after else DB.Line.CreateBound(cut, b)
    doc.Regenerate()
    return (length - kept) * MM_PER_FT, disconnected


SNAPS = (ObjectSnapTypes.Endpoints | ObjectSnapTypes.Midpoints | ObjectSnapTypes.Nearest |
         ObjectSnapTypes.Intersections | ObjectSnapTypes.Perpendicular | ObjectSnapTypes.Centers)
NEAR_MARGIN_MM = 50.0  # tolleranza oltre la meta' della sezione per riconoscere l'elemento cliccato


def half_section(element):
    """Meta' dell'ingombro maggiore della sezione (piedi)."""
    values = []
    for bip in (DB.BuiltInParameter.RBS_PIPE_OUTER_DIAMETER, DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM,
                DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM, DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM,
                DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM,
                DB.BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM):
        param = element.get_Parameter(bip)
        if param is not None and param.HasValue and param.AsDouble() > 0:
            values.append(param.AsDouble())
    return max(values) / 2.0 if values else 0.0


def element_near(point):
    """Tratto rettilineo visibile nella vista il cui asse passa piu' vicino al punto, misurando nel piano della
    vista (in pianta conta solo XY). None se nessuno e' entro meta' sezione + NEAR_MARGIN_MM."""
    view = doc.ActiveView
    normal = view.ViewDirection

    def on_view_plane(p):
        return p.Subtract(normal.Multiply(p.DotProduct(normal)))

    p = on_view_plane(point)
    best, best_score = None, None
    for element in DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType():
        if not isinstance(element, CLASSES):
            continue
        curve = getattr(element.Location, "Curve", None)
        if not isinstance(curve, DB.Line):
            continue
        a, b = on_view_plane(curve.GetEndPoint(0)), on_view_plane(curve.GetEndPoint(1))
        ab = b.Subtract(a)
        if ab.GetLength() < 1e-9:
            distance = p.DistanceTo(a)  # elemento perpendicolare alla vista: visto come un punto
        else:
            t = max(0.0, min(1.0, p.Subtract(a).DotProduct(ab) / ab.DotProduct(ab)))
            distance = p.DistanceTo(a.Add(ab.Multiply(t)))
        limit = half_section(element) + NEAR_MARGIN_MM / MM_PER_FT
        if distance <= limit and (best is None or distance < best_score):
            best, best_score = element, distance
    return best


def quick_main():
    """SHIFT + CLICK: un elemento alla volta, fino a ESC. Primo clic sull'elemento = punto di taglio, secondo clic
    sulla parte dello stesso elemento da eliminare."""
    rows, done = [], 0
    use_snaps = not isinstance(doc.ActiveView, DB.View3D)
    while True:
        if use_snaps:
            # primo clic con gli snap di Revit; l'elemento e' quello con l'asse piu' vicino al punto agganciato
            try:
                cut_point = uidoc.Selection.PickPoint(SNAPS, "Split: click where the element must be cut, snaps "
                                                             "active (ESC to finish)")
            except OperationCanceledException:
                break
            except Exception as ex:
                forms.alert("Points with snaps cannot be picked in this view ({}). The command continues "
                            "without snaps: click directly on the elements.".format(ex), title="Multi Gap")
                use_snaps = False
                continue
            element = element_near(cut_point)
            if element is None:
                rows.append((None, "", "Unchanged. No straight pipe, duct, cable tray or conduit near the clicked "
                                       "point", ""))
                continue
        else:
            # vista 3D: PickPoint richiede un piano di lavoro, si clicca direttamente sull'elemento (niente snap)
            try:
                first = uidoc.Selection.PickObject(ObjectType.Element, CurveFilter(),
                                                   "Split: click the element where it must be cut (ESC to finish)")
            except OperationCanceledException:
                break
            element = doc.GetElement(first.ElementId)
            cut_point = first.GlobalPoint
        uidoc.Selection.SetElementIds(List[DB.ElementId]([element.Id]))  # si vede quale elemento e' stato preso
        try:
            second = uidoc.Selection.PickObject(ObjectType.Element, SameElementFilter(element.Id),
                                                "Split: click the part of the highlighted element to delete "
                                                "(ESC to skip)")
        except OperationCanceledException:
            uidoc.Selection.SetElementIds(List[DB.ElementId]())
            continue
        uidoc.Selection.SetElementIds(List[DB.ElementId]())
        transaction = DB.Transaction(doc, "Split")
        transaction.Start()
        try:
            if cut_point is None or second.GlobalPoint is None:
                raise GapError("The clicked points could not be read in this view")
            if isinstance(doc.ActiveView, DB.ViewPlan) and get_vertical_elements([element]):
                raise GapError("A vertical element cannot be cut in a plan view: use a section, elevation or "
                               "3D view")
            removed, disconnected = trim(element, axis_param(element, cut_point),
                                         axis_param(element, second.GlobalPoint))
            transaction.Commit()
            status = "Cut, {:.0f} mm removed".format(removed)
            if disconnected:
                status += ". Disconnected from: {}".format(", ".join(disconnected))
            rows.append((element, "{:.0f}".format(removed), status, ""))
            done += 1
        except GapError as ex:
            transaction.RollBack()
            rows.append((element, "", "Unchanged. {}".format(ex), ""))
        except Exception as ex:
            transaction.RollBack()
            rows.append((element, "", "Unchanged. The element could not be cut: {}".format(ex), ""))
    if not want_report(len(rows), len(rows) - done, "elements cut"):
        return
    output.print_md("# Split report (cut and delete one side)")
    output.print_md("Elements: **{}** | cut: **{}** | left unchanged: **{}**".format(
        len(rows), done, len(rows) - done))
    for i, (element, removed, status, _) in enumerate(rows, 1):
        label = "{} {}".format(category_name(element), output.linkify(element.Id)) if element is not None \
            else "No element"
        output.print_md("**{}.** {} | {}".format(i, label, status))


class SameElementFilter(ISelectionFilter):
    """Solo l'elemento appena cliccato (secondo clic dello SHIFT + CLICK)."""

    def __init__(self, element_id):
        self.element_id = element_id

    def AllowElement(self, element):
        return element.Id == self.element_id

    def AllowReference(self, reference, point):
        return False


def main():
    elements = get_selection()
    if not elements:
        script.exit()

    # verticale o orizzontale si capisce dagli elementi; selezione mista: avviso e stop
    vertical = get_vertical_elements(elements)
    if vertical and len(vertical) < len(elements):
        forms.alert("The selection mixes {} vertical and {} non-vertical elements. Select only vertical or only "
                    "non-vertical elements, then run the command again.".format(
                        len(vertical), len(elements) - len(vertical)), title="Multi Gap")
        script.exit()
    vertical_mode = bool(vertical)

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
                    "elevation or 3D view and click the two heights on the element.", title="Multi Gap")
        script.exit()

    try:
        ref = cut_reference(elements, vertical_mode)
    except GapError as ex:
        forms.alert(str(ex), title="Multi Gap")
        script.exit()

    rows, done = [], 0
    for element in elements:
        transaction = DB.Transaction(doc, "Split gap")
        transaction.Start()
        info = ""
        try:
            t = axis_positions(element, click1, click2, ref)
            info = diagnostics(element, click1, click2, t)
            gap, status = make_gap(element, t)
            transaction.Commit()
            rows.append((element, "{:.0f}".format(gap), status, info))
            done += 1
        except GapError as ex:
            transaction.RollBack()
            rows.append((element, "", "Unchanged. {}".format(ex), info))
        except Exception as ex:
            transaction.RollBack()
            rows.append((element, "", "Unchanged. The gap could not be created: {}".format(ex), info))
    if not want_report(len(elements), len(elements) - done, "gaps created"):
        return
    output.print_md("# Multi Gap report")
    output.print_md("Elements: **{}** | gaps created: **{}** | left unchanged: **{}**".format(
        len(elements), done, len(elements) - done))
    # righe in testo, non in tabella: si vedono sempre, anche se la tabella non le mostra
    for i, (element, gap, status, info) in enumerate(rows, 1):
        output.print_md("**#{}** {} {} | gap [mm]: {} | **{}**".format(
            i, category_name(element), output.linkify(element.Id), gap or "-", status))
        output.print_md("&nbsp;&nbsp;&nbsp;{}".format(info))


main()
