# -*- coding: utf-8 -*-
"""Openings: selezioni nel modello (muri e solai anche nei link, elementi MEP, area dalla vista attiva) e
smistamento degli elementi gia' selezionati."""
from openings_geometry import *  # noqa: F401,F403


# ----------------------------------------------------------------------------------------------------------------
# PLACE
# ----------------------------------------------------------------------------------------------------------------
class HostFilter(ISelectionFilter):
    """Muri, solai e strutture (travi, pilastri, fondazioni), nel modello host o in un link RVT."""

    def AllowElement(self, element):
        if isinstance(element, DB.RevitLinkInstance):
            return True
        return element.Category is not None and id_value(element.Category.Id) in HOST_KINDS

    def AllowReference(self, reference, point):
        host = ctx.doc.GetElement(reference.ElementId)
        if isinstance(host, DB.RevitLinkInstance):
            link_doc = host.GetLinkDocument()
            element = link_doc.GetElement(reference.LinkedElementId) if link_doc is not None else None
        else:
            element = host
        return element is not None and element.Category is not None and \
            id_value(element.Category.Id) in HOST_KINDS


def pick_hosts():
    try:
        refs = ctx.uidoc.Selection.PickObjects(ObjectType.PointOnElement, HostFilter(),
                                           "Openings: select walls, floors and structural elements, also in linked "
                                           "models, then Finish")
    except OperationCanceledException:
        script.exit()
    hosts, seen = [], set()
    for ref in refs:
        host = ctx.doc.GetElement(ref.ElementId)
        if isinstance(host, DB.RevitLinkInstance):
            element = host.GetLinkDocument().GetElement(ref.LinkedElementId)
            key = (id_value(host.Id), id_value(element.Id))
            record = (element, host.GetTotalTransform(), host.GetLinkDocument().Title)
        else:
            key = (0, id_value(host.Id))
            record = (host, None, "Host model")
        if key not in seen:
            seen.add(key)
            hosts.append((key,) + record)
    return hosts


class MepPickFilter(ISelectionFilter):
    """Tubi, canali, passerelle e conduit rettilinei del modello host."""

    def AllowElement(self, element):
        return is_straight_mep(element)

    def AllowReference(self, reference, point):
        return False


def pick_meps():
    try:
        refs = ctx.uidoc.Selection.PickObjects(ObjectType.Element, MepPickFilter(),
                                           "Openings: select pipes, ducts, cable trays and conduits, then Finish")
    except OperationCanceledException:
        script.exit()
    return [ctx.doc.GetElement(r.ElementId) for r in refs]


def plan_right(view):
    """Destra dello schermo nella pianta; con il riquadro di taglio ruotato (scope box) quella del riquadro."""
    right = flat_unit(view.RightDirection) or DB.XYZ.BasisX
    try:
        if view.CropBoxActive:
            right = flat_unit(view.CropBox.Transform.BasisX) or right
    except Exception:
        pass
    return right


def view_range_z(view):
    """(z min, z max, nota): dal fondo alla cima del view range della pianta, coordinate interne."""
    def plane_z(view_range, plane):
        level = ctx.doc.GetElement(view_range.GetLevelId(plane))
        if not isinstance(level, DB.Level):
            return None  # livello illimitato o relativo
        return level_elevation(level) + view_range.GetOffset(plane)

    bottom = top = None
    try:
        view_range = view.GetViewRange()
        bottom = plane_z(view_range, DB.PlanViewPlane.BottomClipPlane)
        top = plane_z(view_range, DB.PlanViewPlane.TopClipPlane)
    except Exception:
        pass
    note = "view range"
    if bottom is None or top is None:
        base = level_elevation(view.GenLevel) if view.GenLevel is not None else 0.0
        note = "view range not readable, level + {:g} mm".format(FALLBACK_HEIGHT_MM)
        bottom = base if bottom is None else bottom
        top = base + FALLBACK_HEIGHT_MM / MM_PER_FT if top is None else top
    return min(bottom, top), max(bottom, top), note  # nei controsoffitti l'ordine dei piani e' rovesciato


def pick_area():
    """Area dalla vista attiva: riquadro trascinato in pianta, o section box della vista 3D (senza trascinare)."""
    view = ctx.doc.ActiveView
    if isinstance(view, DB.View3D) and not view.IsTemplate:
        if not view.IsSectionBoxActive:
            forms.alert("This 3D view has no section box. Turn it on (or use Quick SectionBox), or drag a rectangle "
                        "in a plan view.", title=ctx.title, exitscript=True)
        box = view.GetSectionBox()
        corners = [box.Transform.OfPoint(DB.XYZ(x, y, z)) for x in (box.Min.X, box.Max.X)
                   for y in (box.Min.Y, box.Max.Y) for z in (box.Min.Z, box.Max.Z)]
        right = flat_unit(box.Transform.BasisX) or DB.XYZ.BasisX
        return make_area(right, corners, min(c.Z for c in corners), max(c.Z for c in corners),
                         "section box of the 3D view '{}'".format(view.Name))
    if isinstance(view, DB.ViewPlan) and not view.IsTemplate and view.ViewType in PLAN_TYPES:
        try:
            picked = ctx.uidoc.Selection.PickBox(PickBoxStyle.Enclosing,
                                             "Openings: drag a rectangle from corner to corner (ESC to cancel)")
        except OperationCanceledException:
            script.exit()
        z_min, z_max, note = view_range_z(view)
        area = make_area(plan_right(view), [picked.Min, picked.Max], z_min, z_max,
                         "rectangle in the plan '{}' ({})".format(view.Name, note))
        if min(area["max"][0] - area["min"][0], area["max"][1] - area["min"][1]) * MM_PER_FT < MIN_AREA_MM:
            forms.alert("The rectangle is less than {:g} mm wide or high: drag from corner to corner.".format(
                MIN_AREA_MM), title=ctx.title, exitscript=True)
        return area
    forms.alert("Area: run it from a floor or ceiling plan (drag a rectangle) or from a 3D view with a section box.",
                title=ctx.title, exitscript=True)


def split_selection():
    """Elementi gia' selezionati, divisi per destinazione:
    - fori gestiti, ed elementi MEP che rimandano ad almeno un foro esistente: check;
    - elementi MEP rettilinei senza foro (anche Not Associated), muri/solai/strutture del modello host: Place.
    Ritorna (fori, MEP associati, MEP da forare, elementi attraversati, numero di elementi non usati)."""
    managed = set(code_of(o) for o in existing_openings())
    openings, linked, meps, hosts, ignored = [], [], [], [], 0
    for element_id in ctx.uidoc.Selection.GetElementIds():
        element = ctx.doc.GetElement(element_id)
        if element is None:
            continue
        if code_of(element) in managed:
            openings.append(element)
        elif isinstance(element, MEP_CLASSES):
            if any(c in managed for c in split_list(get_text(element, ctx.id_param))):
                linked.append(element)
            elif is_straight_mep(element):
                meps.append(element)
            else:
                ignored += 1
        elif element.Category is not None and id_value(element.Category.Id) in HOST_KINDS:
            hosts.append(((0, id_value(element.Id)), element, None, "Host model"))
        else:
            ignored += 1
    return openings, linked, meps, hosts, ignored


def selection_for_check():
    """Elementi gia' selezionati: (elementi MEP, fori gestiti, muri/solai/strutture del modello host)."""
    managed = set(code_of(o) for o in existing_openings())
    meps, openings, hosts = [], [], []
    for element_id in ctx.uidoc.Selection.GetElementIds():
        element = ctx.doc.GetElement(element_id)
        if element is None:
            continue
        if code_of(element) in managed:
            openings.append(element)
        elif isinstance(element, MEP_CLASSES):
            meps.append(element)
        elif element.Category is not None and id_value(element.Category.Id) in HOST_KINDS:
            hosts.append(((0, id_value(element.Id)), element, None, "Host model"))
    return meps, openings, hosts
