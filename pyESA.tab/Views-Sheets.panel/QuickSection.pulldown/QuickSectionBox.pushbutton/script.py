# -*- coding: utf-8 -*-
__title__ = "Quick\nSectionBox"
__doc__ = """Version = 0.1
Date    = 16.10.2026
________________________________________________________________
Description:
Opens a 3D view boxed around a part of the model, from a plan view,
always reusing the same 3D view (one per user, "ESA Quick SectionBox -
<user name>"): only its section box changes, it is never duplicated.

CLICK:
- Click an element (also in linked RVT and IFC models): the section box
  wraps the element, insulation included, plus the margin on every side
  (also below and above)
- OR press ESC and drag a rectangle in the plan: the section box covers
  the rectangle plus the margin; its height is the view range of the
  plan, plus the margins below and above
The box follows the axes of the plan on screen (also with a rotated
crop region). The 3D view opens and is zoomed on the box.

Once the 3D view exists, only its section box changes: template,
visibility, orientation and the other settings you give it are kept.

SHIFT + CLICK: settings first (3D view name, margin, margins below /
above the view range, view template), then the same clicks. A new name
renames the existing 3D view, unless a view with that name already
exists. The template goes to an existing view only if changed in the
window.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

try:
    import traceback
except ImportError:
    traceback = None

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI.Selection import ObjectType, PickBoxStyle
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, DB, forms

import QuickSectionBox_ui

XAML_FILE_NAME = "QuickSectionBox_form.xaml"
CONFIG_NAME = "ESA_QuickSectionBox"
DEFAULTS = {"margin_mm": 300.0, "below_mm": 0.0, "above_mm": 0.0, "view_name": "",
            "template_name": QuickSectionBox_ui.NO_TEMPLATE}

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
MIN_SIZE_MM = 10.0           # un riquadro piu' piccolo di cosi' (senza margine) e' un clic, non un trascinamento
MIN_HEIGHT_MM = 100.0
FALLBACK_HEIGHT_MM = 4000.0  # altezza se il view range non e' leggibile (livelli illimitati)


class BoxError(Exception):
    """Errore atteso: il messaggio va nel report."""


def get_id(element_id):
    return element_id.Value if hasattr(element_id, "Value") else element_id.IntegerValue  # Revit 2026+ / <= 2025


def level_elevation(level):
    # ProjectElevation e' coerente con la geometria (coordinate interne), Elevation puo' riferirsi al Survey Point
    value = getattr(level, "ProjectElevation", None)
    return level.Elevation if value is None else value


def flat(vector):
    return DB.XYZ(vector.X, vector.Y, 0)


def unit_flat(vector):
    """Componente orizzontale unitaria, o None se il vettore e' (quasi) verticale."""
    f = flat(vector)
    return f.Normalize() if f.GetLength() > 1e-9 else None


# ----------------------------------------------------------------------------------------------------------------
# Vista di partenza: assi dello schermo e quote del view range (come in Quick Section)
# ----------------------------------------------------------------------------------------------------------------
def view_axes(view):
    """Destra e alto dello schermo nella pianta. Una pianta con il riquadro di taglio ruotato (scope box) si vede
    ruotata: l'orientamento vero e' quello del riquadro, se attivo."""
    right = unit_flat(view.RightDirection) or DB.XYZ.BasisX
    try:
        if view.CropBoxActive:
            right = unit_flat(view.CropBox.Transform.BasisX) or right
    except Exception:
        pass
    return right, DB.XYZ.BasisZ.CrossProduct(right)  # pianta vista dall'alto: alto = Z x destra


def plane_elevation(view_range, plane):
    """Quota (coordinate interne) di un piano del view range, o None se legato a un livello illimitato/relativo."""
    level = doc.GetElement(view_range.GetLevelId(plane))
    if not isinstance(level, DB.Level):
        return None
    return level_elevation(level) + view_range.GetOffset(plane)


def vertical_range(view, settings):
    """Quote minima e massima del riquadro trascinato: dal fondo alla cima del view range della pianta, piu' i
    margini. Ritorna (z min, z max, nota per il report)."""
    bottom = top = None
    try:
        view_range = view.GetViewRange()
        bottom = plane_elevation(view_range, DB.PlanViewPlane.BottomClipPlane)
        top = plane_elevation(view_range, DB.PlanViewPlane.TopClipPlane)
    except Exception:
        pass
    note = "view range of the plan"
    base = level_elevation(view.GenLevel) if view.GenLevel is not None else 0.0
    if bottom is None or top is None:
        note = "view range not readable (unlimited or relative levels): level of the plan + {:g} mm".format(
            FALLBACK_HEIGHT_MM)
        bottom = base if bottom is None else bottom
        top = base + FALLBACK_HEIGHT_MM / MM_PER_FT if top is None else top
    low, high = min(bottom, top), max(bottom, top)  # nei controsoffitti l'ordine dei piani e' rovesciato
    low -= settings["below_mm"] / MM_PER_FT
    high += settings["above_mm"] / MM_PER_FT
    if (high - low) * MM_PER_FT < MIN_HEIGHT_MM:
        raise BoxError("The height of the box would be less than {:g} mm: check the margins below and above the "
                       "view range (SHIFT + CLICK)".format(MIN_HEIGHT_MM))
    return low, high, note


# ----------------------------------------------------------------------------------------------------------------
# Elemento cliccato (anche in un link): ingombro
# ----------------------------------------------------------------------------------------------------------------
def picked_element(reference):
    """(elemento, transform del link o None, nome del modello)."""
    host = doc.GetElement(reference.ElementId)
    if isinstance(host, DB.RevitLinkInstance):
        link_doc = host.GetLinkDocument()
        if link_doc is None or reference.LinkedElementId == DB.ElementId.InvalidElementId:
            raise BoxError("The element clicked in the linked model could not be read")
        return link_doc.GetElement(reference.LinkedElementId), host.GetTotalTransform(), link_doc.Title
    return host, None, "host model"


def element_label(element, model):
    category = element.Category.Name if element.Category is not None else "Element"
    return "{} {} ({})".format(category, get_id(element.Id), model)


def bbox_points(element, transform):
    """Gli otto vertici del bounding box dell'elemento, in coordinate host."""
    box = element.get_BoundingBox(None)
    if box is None:
        return []
    points = []
    for x in (box.Min.X, box.Max.X):
        for y in (box.Min.Y, box.Max.Y):
            for z in (box.Min.Z, box.Max.Z):
                p = DB.XYZ(x, y, z)
                points.append(transform.OfPoint(p) if transform is not None else p)
    return points


def solid_points(element, transform):
    """Estremi degli spigoli dei solidi dell'elemento, in coordinate host: l'ingombro vero (il bounding box, allineato
    agli assi del progetto, e' troppo grande per un elemento obliquo). Lista vuota se l'elemento non ha solidi."""
    points = []
    try:
        pending = list(element.get_Geometry(DB.Options()) or [])
    except Exception:
        return points
    while pending:
        item = pending.pop()
        if isinstance(item, DB.GeometryInstance):
            pending.extend(item.GetInstanceGeometry())
        elif isinstance(item, DB.Solid) and item.Volume > 0:
            for edge in item.Edges:
                for end in (0, 1):
                    p = edge.AsCurve().GetEndPoint(end)
                    points.append(transform.OfPoint(p) if transform is not None else p)
    return points


def body_points(element, transform):
    """Ingombro dell'elemento (con l'isolamento, per tubi e canali del modello host). Ripiego sul bounding box se non
    ci sono solidi leggibili."""
    points = solid_points(element, transform) or bbox_points(element, transform)
    if transform is None and isinstance(element, DB.MEPCurve):
        try:
            for insulation_id in DB.InsulationLiningBase.GetInsulationIds(doc, element.Id):
                points.extend(solid_points(doc.GetElement(insulation_id), None))
        except Exception:
            pass
    return points


# ----------------------------------------------------------------------------------------------------------------
# Section box: riquadro allineato agli assi della pianta
# ----------------------------------------------------------------------------------------------------------------
def extent(points, axes):
    """(min destra, max destra, min alto, max alto) dei punti proiettati sugli assi della pianta."""
    right, up = axes
    xs = [flat(p).DotProduct(right) for p in points]
    ys = [flat(p).DotProduct(up) for p in points]
    return min(xs), max(xs), min(ys), max(ys)


def make_box(axes, x_range, y_range, z_range):
    """BoundingBoxXYZ con X = destra e Y = alto dello schermo della pianta, Z verticale: la section box resta
    allineata alla pianta anche se questa e' ruotata. I range sono gia' comprensivi dei margini."""
    right, up = axes
    transform = DB.Transform.Identity
    transform.BasisX = right
    transform.BasisY = up
    transform.BasisZ = DB.XYZ.BasisZ
    box = DB.BoundingBoxXYZ()
    box.Transform = transform  # origine nell'origine del progetto: le coordinate locali sono le proiezioni
    box.Min = DB.XYZ(x_range[0], y_range[0], z_range[0])
    box.Max = DB.XYZ(x_range[1], y_range[1], z_range[1])
    return box


def element_box(points, axes, margin):
    if not points:
        raise BoxError("The element has no readable geometry or bounding box")
    x0, x1, y0, y1 = extent(points, axes)
    z0, z1 = min(p.Z for p in points), max(p.Z for p in points)
    return make_box(axes, (x0 - margin, x1 + margin), (y0 - margin, y1 + margin), (z0 - margin, z1 + margin))


def rectangle_box(corners, axes, margin, z_min, z_max):
    x0, x1, y0, y1 = extent(corners, axes)
    if min(x1 - x0, y1 - y0) * MM_PER_FT < MIN_SIZE_MM:
        raise BoxError("The rectangle is less than {:g} mm wide or high: drag from corner to corner".format(
            MIN_SIZE_MM))
    return make_box(axes, (x0 - margin, x1 + margin), (y0 - margin, y1 + margin), (z_min, z_max))


def box_size_mm(box):
    size = box.Max.Subtract(box.Min)
    return "{:.0f} x {:.0f} x {:.0f} (right x up x height on the plan)".format(
        size.X * MM_PER_FT, size.Y * MM_PER_FT, size.Z * MM_PER_FT)


# ----------------------------------------------------------------------------------------------------------------
# Vista 3D: creazione, template, section box
# ----------------------------------------------------------------------------------------------------------------
def default_view_name():
    return "ESA Quick SectionBox - {}".format(doc.Application.Username)


def view_name(settings):
    """Nome scelto con SHIFT + CLICK; vuoto = nome automatico, uno per utente."""
    return (settings.get("view_name") or "").strip() or default_view_name()


def find_view(name):
    for view in DB.FilteredElementCollector(doc).OfClass(DB.View3D):
        if not view.IsTemplate and view.Name == name:
            return view
    return None


def view_templates():
    return [v for v in DB.FilteredElementCollector(doc).OfClass(DB.View3D) if v.IsTemplate]


def three_d_type():
    for view_type in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
        if view_type.ViewFamily == DB.ViewFamily.ThreeDimensional:
            return view_type
    return None


def apply_template(view, settings, notes):
    templates = dict((v.Name, v) for v in view_templates())
    template = templates.get(settings["template_name"])
    if template is not None:
        view.ViewTemplateId = template.Id
        return
    if settings["template_name"] != QuickSectionBox_ui.NO_TEMPLATE:
        notes.append("view template '{}' not found: none applied".format(settings["template_name"]))
    view.ViewTemplateId = DB.ElementId.InvalidElementId


def place_box(box, settings, template_changed, previous_name):
    """Applica la section box alla vista 3D dell'utente (o la crea). Una vista esistente cambia solo la section box;
    il template solo se cambiato ora nella finestra. Se il nome e' stato cambiato e non esiste gia' una vista con il
    nuovo nome, quella di prima viene rinominata. Ritorna (vista, cosa e' successo, note)."""
    name = view_name(settings)
    other = [v for v in DB.FilteredElementCollector(doc).OfClass(DB.View)
             if v.Name == name and not (isinstance(v, DB.View3D) and not v.IsTemplate)]
    if other:
        raise BoxError("A view named '{}' already exists and is not a 3D view: choose another name with "
                       "SHIFT + CLICK".format(name))
    notes = []
    with revit.Transaction("Quick SectionBox"):
        view = find_view(name)
        if view is None and previous_name and previous_name != name:
            view = find_view(previous_name)
            if view is not None:
                view.Name = name
                notes.append("renamed from '{}'".format(previous_name))
        action = "updated"
        if view is None:
            view_type = three_d_type()
            if view_type is None:
                raise BoxError("The model has no 3D view type")
            view = DB.View3D.CreateIsometric(doc, view_type.Id)
            view.Name = name
            apply_template(view, settings, notes)
            action = "created"
        elif template_changed:
            apply_template(view, settings, notes)
        view.IsSectionBoxActive = True
        view.SetSectionBox(box)
    return view, action, notes


def show_view(view, box):
    """Rende attiva la vista 3D e la inquadra sulla section box. Se e' gia' aperta si aggiorna nella sua finestra."""
    open_ids = [get_id(ui_view.ViewId) for ui_view in uidoc.GetOpenUIViews()]
    already = get_id(view.Id) in open_ids
    uidoc.ActiveView = view
    zoom = ""
    try:
        for ui_view in uidoc.GetOpenUIViews():
            if get_id(ui_view.ViewId) == get_id(view.Id):
                ui_view.ZoomAndCenterRectangle(box.Transform.OfPoint(box.Min), box.Transform.OfPoint(box.Max))
                break
    except Exception as ex:
        zoom = " (not zoomed on the box: {})".format(ex)
    return ("already open: updated and shown" if already else "opened") + zoom


# ----------------------------------------------------------------------------------------------------------------
# Impostazioni e report
# ----------------------------------------------------------------------------------------------------------------
def load_settings(config):
    settings = dict(DEFAULTS)
    try:
        for key in ("margin_mm", "below_mm", "above_mm"):
            settings[key] = float(config.get_option(key, DEFAULTS[key]))
        for key in ("view_name", "template_name"):
            settings[key] = config.get_option(key, DEFAULTS[key])
    except Exception:
        pass
    return settings


def save_settings(config, settings):
    try:
        for key in DEFAULTS:
            setattr(config, key, settings[key])
        script.save_config()
    except Exception:
        pass


def print_report(info):
    output.print_md("# Quick SectionBox report")
    if info.get("error"):
        output.print_md("**Nothing was changed.** {}".format(info["error"]))
    for label, key in (("Mode", "mode"), ("3D view", "view"), ("Result", "action"), ("Window", "window"),
                       ("Box [mm]", "size"), ("Height", "height"), ("Notes", "notes")):
        if info.get(key):
            output.print_md("**{}:** {}".format(label, info[key]))
    if info.get("detail"):
        output.print_md("**Technical detail**")
        output.print_code(info["detail"])


def finish(info, ok):
    if want_report(1, 0 if ok else 1, "section box placed"):
        print_report(info)


def main():
    plan = doc.ActiveView
    if not isinstance(plan, DB.ViewPlan) or plan.IsTemplate:
        forms.alert("Run Quick SectionBox from a plan view (floor or ceiling plan).", title="Quick SectionBox",
                    exitscript=True)

    config = script.get_config(CONFIG_NAME)
    settings = load_settings(config)
    template_changed = False  # il template della vista esistente cambia solo se cambiato ora nella finestra
    previous_name = view_name(settings)
    if __shiftclick__:  # noqa: F821
        previous = settings
        settings = QuickSectionBox_ui.SettingsWindow(
            script.get_bundle_file(XAML_FILE_NAME), settings, sorted(v.Name for v in view_templates()),
            default_view_name()).show()
        if settings is None:
            script.exit()
        save_settings(config, settings)
        template_changed = settings["template_name"] != previous["template_name"]

    info = {}
    try:
        axes = view_axes(plan)
        margin = settings["margin_mm"] / MM_PER_FT
        try:
            reference = uidoc.Selection.PickObject(
                ObjectType.PointOnElement, "Quick SectionBox: click an element, also in a linked model (ESC: drag a "
                                           "rectangle instead)")
        except OperationCanceledException:
            reference = None
        if reference is not None:
            element, transform, model = picked_element(reference)
            info["mode"] = "Around {}, margin {:g} mm on every side".format(element_label(element, model),
                                                                           settings["margin_mm"])
            box = element_box(body_points(element, transform), axes, margin)
            info["height"] = "element height plus the margin"
        else:
            try:
                picked = uidoc.Selection.PickBox(PickBoxStyle.Enclosing,
                                                 "Quick SectionBox: drag a rectangle from corner to corner (ESC to "
                                                 "cancel)")
            except OperationCanceledException:
                script.exit()
            z_min, z_max, note = vertical_range(plan, settings)
            info["mode"] = "Rectangle in the plan, margin {:g} mm on each side".format(settings["margin_mm"])
            box = rectangle_box([picked.Min, picked.Max], axes, margin, z_min, z_max)
            info["height"] = "{:.0f} to {:.0f} mm, internal coordinates ({})".format(
                z_min * MM_PER_FT, z_max * MM_PER_FT, note)
        info["size"] = box_size_mm(box)
        view, info["action"], notes = place_box(box, settings, template_changed, previous_name)
        info["view"] = "{} {}".format(view.Name, output.linkify(view.Id))
        info["notes"] = "; ".join(notes)
    except BoxError as ex:
        info["error"] = str(ex)
        finish(info, False)
        return
    except Exception as ex:
        info["error"] = "The section box could not be placed: {}".format(ex)
        info["detail"] = traceback.format_exc() if traceback else ""
        finish(info, False)
        return

    try:
        info["window"] = show_view(view, box)
    except Exception as ex:
        info["window"] = "not opened: {}".format(ex)
    finish(info, "not zoomed" not in info.get("window", "") and "not found" not in info["notes"])


main()
