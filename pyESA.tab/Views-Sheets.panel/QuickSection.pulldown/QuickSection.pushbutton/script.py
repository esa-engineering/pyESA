# -*- coding: utf-8 -*-
__title__ = "Quick\nSection"
__doc__ = """Version = 0.2
Date    = 16.10.2026
________________________________________________________________
Description:
Places a section quickly from a plan view, always reusing the same
section (one per user, "ESA Quick Section - <user name>"): it is moved
and turned to the new place, never duplicated.

CLICK:
1. Click an element to cut parallel to it: pipes, ducts, cable trays,
   conduits, walls, beams, grids, lines, also in linked RVT and IFC
   models. The section line runs along its axis. An element without an
   axis is followed along the clicked face. A vertical element (riser)
   is cut through its axis, along the nearest axis of the view
   OR press ESC and click two points: the section line snaps to the
   nearest axis of the view (horizontal or vertical on screen)
2. Click on the side to look at: the distance of this click from the
   section line is the depth (far clip)
The section line is then moved back, opposite to where it looks, just
behind the clicked element (50 mm clear, insulation included), so the
element is seen whole and not cut; the depth grows by the same amount,
so the far clip stays on the click. Grids, lines and the two-point mode
keep the line where it is.
Width: the element (or the two points) plus the side margin. Height:
the view range of the plan, plus the margins below and above.
The first time the section opens beside the plan (windows tiled); if it
is already open it simply updates.

Once the section exists, only its position, crop and depth change:
its other settings (template, visibility, scale, detail level, far clip
mode) are kept. If it has to be recreated, they are copied from the old
one.

SHIFT + CLICK: settings first (section name, side margin, margins below /
above the view range, section type and view template), then the same
clicks. A new name renames the existing section (it stays the same
section), unless a section with that name already exists. Type
and template go to an existing section only if changed in the window.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import math

try:
    import traceback
except ImportError:
    traceback = None

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import RevitCommandId, PostableCommand
from Autodesk.Revit.UI.Selection import ObjectType, ObjectSnapTypes
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, DB, forms

import QuickSection_ui

XAML_FILE_NAME = "QuickSection_form.xaml"
CONFIG_NAME = "ESA_QuickSection"
DEFAULTS = {"margin_mm": 300.0, "below_mm": 0.0, "above_mm": 0.0, "type_name": "", "section_name": "",
            "template_name": QuickSection_ui.NO_TEMPLATE}

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
MIN_DEPTH_MM = 10.0       # il clic finale deve stare almeno a questa distanza dalla linea di sezione
MIN_WIDTH_MM = 10.0       # due punti piu' vicini di cosi' non danno una linea
MIN_HEIGHT_MM = 100.0
BACK_CLEARANCE_MM = 50.0  # distanza libera tra la linea di sezione arretrata e l'elemento cliccato
FALLBACK_HEIGHT_MM = 4000.0  # altezza se il view range non e' leggibile (livelli illimitati)
VERTICAL_TOL = 0.999
VERTICAL_FACE_TOL = 0.1   # sotto questa componente orizzontale la normale della faccia non da' una direzione
SNAPS = (ObjectSnapTypes.Endpoints | ObjectSnapTypes.Midpoints | ObjectSnapTypes.Nearest |
         ObjectSnapTypes.Intersections | ObjectSnapTypes.Perpendicular | ObjectSnapTypes.Centers)


class SectionError(Exception):
    """Errore atteso: il messaggio va nel report."""


def get_id(element_id):
    return element_id.Value if hasattr(element_id, "Value") else element_id.IntegerValue  # Revit 2026+ / <= 2025


def level_elevation(level):
    # ProjectElevation e' coerente con la geometria (coordinate interne), Elevation puo' riferirsi al Survey Point
    value = getattr(level, "ProjectElevation", None)
    return level.Elevation if value is None else value


def type_name(element_type):
    try:
        return element_type.Name
    except Exception:
        return DB.Element.Name.GetValue(element_type)


def flat(vector):
    return DB.XYZ(vector.X, vector.Y, 0)


def unit_flat(vector):
    """Componente orizzontale unitaria, o None se il vettore e' (quasi) verticale."""
    f = flat(vector)
    return f.Normalize() if f.GetLength() > 1e-9 else None


# ----------------------------------------------------------------------------------------------------------------
# Vista di partenza: assi dello schermo e quote del view range
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
    """Quote minima e massima della sezione: dal fondo alla cima del view range della pianta, piu' i margini.
    Ritorna (z min, z max, nota per il report)."""
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
        raise SectionError("The height of the section would be less than {:g} mm: check the margins below and "
                           "above the view range (SHIFT + CLICK)".format(MIN_HEIGHT_MM))
    return low, high, note


# ----------------------------------------------------------------------------------------------------------------
# Elemento cliccato (anche in un link): asse, faccia o montante
# ----------------------------------------------------------------------------------------------------------------
def hit_face_normal(reference, element, link_instance):
    """Normale (coordinate host) della faccia piana cliccata, o None se non e' leggibile. Per gli elementi di un link
    la faccia si legge dal documento del link e la normale si porta nel modello host."""
    try:
        if link_instance is not None:
            face = element.GetGeometryObjectFromReference(reference.CreateReferenceInLink())
        else:
            face = element.GetGeometryObjectFromReference(reference)
        if not isinstance(face, DB.PlanarFace):
            return None
        normal = face.FaceNormal
        return link_instance.GetTotalTransform().OfVector(normal) if link_instance is not None else normal
    except Exception:
        return None


def picked_element(reference):
    """(elemento, transform del link o None, nome del modello, normale della faccia cliccata o None)."""
    host = doc.GetElement(reference.ElementId)
    if isinstance(host, DB.RevitLinkInstance):
        link_doc = host.GetLinkDocument()
        if link_doc is None or reference.LinkedElementId == DB.ElementId.InvalidElementId:
            raise SectionError("The element clicked in the linked model could not be read")
        element = link_doc.GetElement(reference.LinkedElementId)
        return element, host.GetTotalTransform(), link_doc.Title, hit_face_normal(reference, element, host)
    return host, None, "host model", hit_face_normal(reference, host, None)


def element_label(element, model):
    category = element.Category.Name if element.Category is not None else "Element"
    return "{} {} ({})".format(category, get_id(element.Id), model)


def curve_of(element):
    location = getattr(element, "Location", None)
    if isinstance(location, DB.LocationCurve):
        return location.Curve
    if isinstance(element, DB.Grid):
        return element.Curve
    if isinstance(element, DB.CurveElement):
        return element.GeometryCurve
    return None


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
    """Ingombro dell'elemento (con l'isolamento, per tubi e canali del modello host), per tenere la linea di sezione
    fuori dall'elemento. Ripiego sul bounding box se non ci sono solidi leggibili."""
    points = solid_points(element, transform) or bbox_points(element, transform)
    if transform is None and isinstance(element, DB.MEPCurve):
        try:
            for insulation_id in DB.InsulationLiningBase.GetInsulationIds(doc, element.Id):
                points.extend(solid_points(doc.GetElement(insulation_id), None))
        except Exception:
            pass
    return points


def element_line(element, transform, hit, normal):
    """Linea di sezione suggerita dall'elemento. Ritorna dict con kind ("axis", "face" o "vertical"), tangent
    (orizzontale unitaria, None per un montante), origin (punto sulla linea) e points (ingombro, per la larghezza)."""
    curve = curve_of(element)
    if curve is not None:
        if transform is not None:
            curve = curve.CreateTransformed(transform)
        if isinstance(curve, DB.Line) and abs(curve.Direction.Z) > VERTICAL_TOL:
            ends = [curve.GetEndPoint(0), curve.GetEndPoint(1)]
            return {"kind": "vertical", "tangent": None, "origin": ends[0],
                    "points": bbox_points(element, transform) or ends}
        projection = curve.Project(hit)
        if projection is not None:
            tangent = unit_flat(curve.ComputeDerivatives(projection.Parameter, False).BasisX)
            if tangent is not None:
                points = list(curve.Tessellate()) if curve.IsBound else bbox_points(element, transform)
                return {"kind": "axis", "tangent": tangent, "origin": projection.XYZPoint, "points": points}
    # senza asse orizzontale (pilastro, pannello, DirectShape IFC): lungo la faccia cliccata
    if normal is None:
        raise SectionError("The element has no axis and the clicked face could not be read: click a pipe, duct, "
                           "wall or beam, or a vertical face, or press ESC at the first click and use two points")
    along = DB.XYZ(-normal.Y, normal.X, 0)
    if along.GetLength() < VERTICAL_FACE_TOL:
        raise SectionError("The clicked face is horizontal (floor, soffit, top of an element): it gives no "
                           "direction. Click a vertical face, or use two points")
    return {"kind": "face", "tangent": along.Normalize(), "origin": hit, "points": bbox_points(element, transform)}


# ----------------------------------------------------------------------------------------------------------------
# Clic e geometria della sezione (tutto in pianta)
# ----------------------------------------------------------------------------------------------------------------
def pick_point(prompt):
    try:
        return uidoc.Selection.PickPoint(SNAPS, prompt)
    except OperationCanceledException:
        script.exit()


def extent_along(center, direction, points):
    """(min, max) delle proiezioni dei punti sulla direzione orizzontale, misurate da center."""
    values = [flat(p.Subtract(center)).DotProduct(direction) for p in points] or [0.0]
    return min(values), max(values)


def frame(origin, tangent, points, click, margin, body=None):
    """Sezione con la linea per origin lungo tangent, che guarda verso il lato del clic. Ritorna dict con center
    (centro della linea), look (verso in cui guarda), half (meta' larghezza), depth (profondita') e back
    (arretramento della linea). body: ingombro dell'elemento cliccato; se la linea lo taglia, la linea arretra
    (verso opposto a quello di vista) fin dietro l'elemento piu' BACK_CLEARANCE_MM, e la profondita' cresce
    altrettanto, cosi' il limite lontano resta sul clic."""
    o = DB.XYZ(origin.X, origin.Y, 0)
    left = DB.XYZ(-tangent.Y, tangent.X, 0)
    offset = flat(click.Subtract(o)).DotProduct(left)
    if abs(offset) * MM_PER_FT < MIN_DEPTH_MM:
        raise SectionError("The last click is on the section line: click on the side to look at, as far as the "
                           "depth you want")
    low, high = extent_along(o, tangent, points)
    if (high - low) * MM_PER_FT < MIN_WIDTH_MM:
        raise SectionError("The element or the two points are less than {:g} mm long along the section "
                           "line".format(MIN_WIDTH_MM))
    look = left if offset > 0 else left.Negate()
    back = 0.0
    if body:
        behind = extent_along(o, look, body)[0]  # punto dell'elemento piu' dietro la linea (negativo = tagliato)
        if behind < 0:
            back = -behind + BACK_CLEARANCE_MM / MM_PER_FT
    center = o.Add(tangent.Multiply((low + high) / 2.0)).Subtract(look.Multiply(back))
    return {"center": center, "look": look, "half": (high - low) / 2.0 + margin, "depth": abs(offset) + back,
            "back": back}


def riser_frame(origin, points, click, axes, margin, body=None):
    """Montante: sezione per il suo asse, lungo l'asse della vista che lascia il clic di lato. Clic a destra o a
    sinistra: guarda da quella parte (linea verticale sullo schermo); clic sopra o sotto: linea orizzontale."""
    right, up = axes
    c = DB.XYZ(origin.X, origin.Y, 0)
    v = flat(click.Subtract(c))
    tangent = up if abs(v.DotProduct(right)) >= abs(v.DotProduct(up)) else right
    return frame(c, tangent, points, click, margin, body)


def snapped_axis(p1, p2, axes):
    """Asse della vista (destra o alto) piu' vicino alla direzione p1-p2."""
    d = flat(p2.Subtract(p1))
    if d.GetLength() * MM_PER_FT < MIN_WIDTH_MM:
        raise SectionError("The two points are less than {:g} mm apart".format(MIN_WIDTH_MM))
    right, up = axes
    return right if abs(d.DotProduct(right)) >= abs(d.DotProduct(up)) else up


# ----------------------------------------------------------------------------------------------------------------
# Sezione: creazione, spostamento, riquadro
# ----------------------------------------------------------------------------------------------------------------
def desired_transform(geo, z_mid):
    """Sistema della sezione: X = destra di chi guarda, Y = alto, Z = verso chi guarda (la vista guarda verso -Z)."""
    look = geo["look"]
    transform = DB.Transform.Identity
    transform.Origin = DB.XYZ(geo["center"].X, geo["center"].Y, z_mid)
    transform.BasisX = DB.XYZ(look.Y, -look.X, 0)  # X x Y = Z = -look
    transform.BasisY = DB.XYZ.BasisZ
    transform.BasisZ = look.Negate()
    return transform


def section_box(transform, geo, z_min, z_max):
    """Riquadro della sezione nelle coordinate locali di transform: larghezza attorno al centro, quote, profondita'
    dalla linea di sezione (Z = 0) verso il lato guardato (Min.Z). Revit disegna la linea e misura il far clip dal
    piano Z = 0 del sistema: un Max.Z diverso da 0 allarga il riquadro ma non sposta la linea. Per questo la linea
    si porta sul bersaglio spostando il simbolo (align_line), non il riquadro."""
    rel = geo["center"].Subtract(transform.Origin)
    x0 = flat(rel).DotProduct(transform.BasisX)
    box = DB.BoundingBoxXYZ()
    box.Transform = transform
    box.Min = DB.XYZ(x0 - geo["half"], z_min - transform.Origin.Z, -geo["depth"])
    box.Max = DB.XYZ(x0 + geo["half"], z_max - transform.Origin.Z, 0)
    return box


def default_section_name():
    return "ESA Quick Section - {}".format(doc.Application.Username)


def section_name(settings):
    """Nome scelto con SHIFT + CLICK; vuoto = nome automatico, uno per utente."""
    return (settings.get("section_name") or "").strip() or default_section_name()


def find_section(name):
    for view in DB.FilteredElementCollector(doc).OfClass(DB.ViewSection):
        if not view.IsTemplate and view.Name == name:
            return view
    return None


def section_types():
    return [t for t in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
            if t.ViewFamily == DB.ViewFamily.Section]


def section_templates():
    return [v for v in DB.FilteredElementCollector(doc).OfClass(DB.View)
            if v.IsTemplate and v.ViewType == DB.ViewType.Section]


def marker_ids(section):
    """Elementi da ruotare e spostare: prima il simbolo della sezione (categoria Viewers, l'elemento che l'utente
    ruota nella pianta), poi la vista stessa come ultima risorsa."""
    ids = []
    try:
        ids = list(section.GetDependentElements(DB.ElementCategoryFilter(DB.BuiltInCategory.OST_Viewers)))
    except Exception:
        pass
    if not ids:
        ids = [e.Id for e in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Viewers)
               .WhereElementIsNotElementType() if e.Name == section.Name]
    return ids + [section.Id]


def actual_look(section):
    """Direzione in cui guarda davvero la sezione (orizzontale): ViewDirection punta verso chi guarda."""
    return unit_flat(section.ViewDirection.Negate())


def line_point(section):
    """Centro della linea di sezione: meta' della larghezza sul piano Z = 0 del sistema del ritaglio, dove Revit
    disegna la linea. NON su Max.Z: con un Max.Z diverso da 0 la linea restava a meta' del riquadro. L'origine del
    sistema sta sulla linea ma non al suo centro (lungo la linea puo' stare anche a metri), quindi si prende X medio."""
    box = section.CropBox
    return box.Transform.OfPoint(DB.XYZ((box.Min.X + box.Max.X) / 2.0, 0, 0))


def line_offset(section, geo):
    """Distanza in pianta, lungo il verso di vista, dalla linea di sezione a quella voluta (positiva = la linea
    deve avanzare verso il lato guardato)."""
    return flat(geo["center"].Subtract(line_point(section))).DotProduct(geo["look"])


def transform_marker(section, action):
    """Applica action(id) al primo elemento della sezione che Revit accetta. Ritorna l'id usato."""
    errors = []
    for element_id in marker_ids(section):
        try:
            action(element_id)
            doc.Regenerate()
            return element_id
        except Exception as ex:
            errors.append(str(ex))
    raise SectionError("Revit refused to transform the section: {}".format("; ".join(errors)))


def turn_section(section, look, pivot):
    """Ruota la sezione attorno alla verticale per pivot finche' guarda verso look. Ritorna l'angolo in gradi."""
    current = actual_look(section)
    angle = math.atan2(current.CrossProduct(look).Z, current.DotProduct(look))
    if abs(angle) < 1e-9:
        return 0.0
    axis = DB.Line.CreateBound(pivot, pivot.Add(DB.XYZ.BasisZ))
    used = transform_marker(section, lambda i: DB.ElementTransformUtils.RotateElement(doc, i, axis, angle))
    # se l'elemento accettato non ha girato la vista (es. la vista ignora la rotazione), si prova con gli altri
    if actual_look(section).DotProduct(look) < math.cos(math.radians(0.1)):
        for element_id in marker_ids(section):
            if element_id == used:
                continue
            try:
                DB.ElementTransformUtils.RotateElement(doc, element_id, axis, angle)
                doc.Regenerate()
            except Exception:
                continue
            if actual_look(section).DotProduct(look) > math.cos(math.radians(0.1)):
                break
    return math.degrees(angle)


def move_section(section, wanted, look):
    """Ruota (attorno alla verticale) e sposta la sezione esistente sulla nuova linea, verificando sulla direzione
    reale della vista. SectionError se Revit non porta la sezione dove serve."""
    turn_section(section, look, line_point(section))
    delta = flat(wanted.Origin.Subtract(line_point(section)))
    if delta.GetLength() > 1e-9:
        transform_marker(section, lambda i: DB.ElementTransformUtils.MoveElement(doc, i, delta))
    facing = actual_look(section).DotProduct(look)
    distance = abs(flat(wanted.Origin.Subtract(line_point(section))).DotProduct(look)) * MM_PER_FT
    if facing < math.cos(math.radians(0.1)) or distance > 1.0:
        raise SectionError("the section did not reach the expected position (direction {:.1f} degrees off, section "
                           "line {:.0f} mm away)".format(math.degrees(math.acos(max(-1.0, min(1.0, facing)))),
                                                         distance))


VIEW_KEYS = ("type_name", "template_name")  # impostazioni della finestra che toccano la vista


def apply_view_settings(section, plan, settings, notes, keys):
    """Tipo e/o template come da impostazioni, solo per le chiavi in keys. Una sezione esistente non si tocca: keys
    contiene tutto alla creazione, e solo quanto cambiato nella finestra (SHIFT + CLICK) negli spostamenti."""
    if "type_name" in keys:
        types = dict((type_name(t), t) for t in section_types())
        wanted_type = types.get(settings["type_name"])
        if wanted_type is not None and section.GetTypeId() != wanted_type.Id:
            try:
                section.ChangeTypeId(wanted_type.Id)
            except Exception as ex:
                notes.append("section type not changed: {}".format(ex))
    if "template_name" in keys:
        templates = dict((v.Name, v) for v in section_templates())
        template = templates.get(settings["template_name"])
        if template is not None:
            section.ViewTemplateId = template.Id
        else:
            if settings["template_name"] != QuickSection_ui.NO_TEMPLATE:
                notes.append("view template '{}' not found: none applied".format(settings["template_name"]))
            section.ViewTemplateId = DB.ElementId.InvalidElementId


def copy_view_state(old, new, notes):
    """Sezione ricreata: le impostazioni della vecchia (visibilita', scala, dettaglio, far clip...) passano alla nuova
    tramite un template temporaneo creato dalla vecchia e applicato una volta sola, poi eliminato."""
    try:
        if new.GetTypeId() != old.GetTypeId():
            new.ChangeTypeId(old.GetTypeId())
    except Exception:
        pass
    template_id = old.ViewTemplateId
    temporary = None
    try:
        temporary = old.CreateViewTemplate()
        new.ApplyViewTemplateParameters(temporary)
    except Exception as ex:
        notes.append("the settings of the old section could not all be copied: {}".format(ex))
        try:
            new.Scale = old.Scale
            new.DetailLevel = old.DetailLevel
        except Exception:
            pass
    finally:
        if temporary is not None:
            doc.Delete(temporary.Id)
    new.ViewTemplateId = template_id


def set_extent(section, geo, z_min, z_max, notes, created):
    """Riquadro e profondita'. La modalita' di far clip si imposta solo su una sezione nuova: su quella esistente
    resta la scelta dell'utente."""
    box = section_box(section.CropBox.Transform, geo, z_min, z_max)
    section.CropBoxActive = True
    section.CropBox = box
    far = section.get_Parameter(DB.BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)
    if far is not None and not far.IsReadOnly:
        far.Set(geo["depth"])
    else:
        notes.append("far clip controlled by the view template: the depth of the click is not applied")
    clip = section.get_Parameter(DB.BuiltInParameter.VIEWER_BOUND_FAR_CLIPPING)
    if clip is None or clip.AsInteger() != 0:
        return
    if created and not clip.IsReadOnly:
        clip.Set(1)  # 0 = nessun taglio: senza, la profondita' non avrebbe effetto
    else:
        notes.append("far clipping is off in this section: the depth is set but has no effect")


def align_line(section, geo, z_min, z_max, created):
    """Porta la linea di sezione sulla linea voluta spostando il simbolo lungo il verso di vista, poi ricalcola il
    riquadro. Ritorna la distanza che resta (piedi)."""
    for _ in range(3):
        off = line_offset(section, geo)
        if abs(off) * MM_PER_FT < 1.0:
            return off
        delta = geo["look"].Multiply(off)
        transform_marker(section, lambda i: DB.ElementTransformUtils.MoveElement(doc, i, delta))
        set_extent(section, geo, z_min, z_max, [], created)
    return line_offset(section, geo)


def place_section(plan, geo, z_min, z_max, settings, changed_keys, previous_name):
    """Sposta la sezione dell'utente (o la crea). Una sezione esistente cambia solo posizione, riquadro e
    profondita'; tipo e template solo se cambiati ora nella finestra (changed_keys). Se il nome e' stato cambiato
    (previous_name) e non esiste gia' una sezione con il nuovo nome, quella di prima viene rinominata: resta la
    stessa sezione. Ritorna (sezione, cosa e' successo, note, diagnosi della linea)."""
    name = section_name(settings)
    wanted = desired_transform(geo, (z_min + z_max) / 2.0)
    notes = []
    other = [v for v in DB.FilteredElementCollector(doc).OfClass(DB.View)
             if v.Name == name and not isinstance(v, DB.ViewSection)]
    if other:
        raise SectionError("A view named '{}' already exists and is not a section: choose another name with "
                           "SHIFT + CLICK".format(name))
    transaction = DB.Transaction(doc, "Quick Section")
    transaction.Start()
    try:
        section = find_section(name)
        if section is None and previous_name and previous_name != name:
            section = find_section(previous_name)
            if section is not None:
                section.Name = name
                notes.append("renamed from '{}'".format(previous_name))
        old = None
        action = "moved"
        if section is not None:
            attempt = DB.SubTransaction(doc)
            attempt.Start()
            try:
                move_section(section, wanted, geo["look"])
                attempt.Commit()
            except Exception as ex:
                attempt.RollBack()
                # ultima risorsa: si ricrea con lo stesso nome (se era aperta, la sua finestra si chiude). La
                # vecchia si rinomina e si elimina dopo averne copiato le impostazioni
                notes.append("the existing section could not be moved ({}): it was recreated with its "
                             "settings".format(ex))
                old = section
                old.Name = name + " (old)"
                section = None
                action = "recreated"
        if section is None:
            types = dict((type_name(t), t) for t in section_types())
            section_type = types.get(settings["type_name"]) or section_types()[0]
            box = section_box(wanted, geo, z_min, z_max)
            section = DB.ViewSection.CreateSection(doc, section_type.Id, box)
            doc.Regenerate()
            if actual_look(section).DotProduct(geo["look"]) < math.cos(math.radians(0.1)):
                # Revit l'ha creata girata: si porta sul lato cliccato attorno al centro della linea
                turned = turn_section(section, geo["look"], line_point(section))
                notes.append("created by Revit looking the other way, turned by {:.0f} degrees".format(turned))
            if old is not None:
                copy_view_state(old, section, notes)
                doc.Delete(old.Id)
                apply_view_settings(section, plan, settings, notes, changed_keys)
            else:
                action = "created"
                apply_view_settings(section, plan, settings, notes, VIEW_KEYS)
                if section.ViewTemplateId == DB.ElementId.InvalidElementId:
                    try:
                        section.Scale = plan.Scale
                    except Exception:
                        pass
            section.Name = name
        else:
            apply_view_settings(section, plan, settings, notes, changed_keys)
        doc.Regenerate()
        set_extent(section, geo, z_min, z_max, notes, created=(action == "created"))
        off = align_line(section, geo, z_min, z_max, created=(action == "created"))
        # linea misurata sul piano Z = 0 del ritaglio (dove Revit la disegna); il riquadro parte da li'
        box = section.CropBox
        diag = "section line {:.0f} mm from the target; crop near side {:.0f} mm from the line".format(
            abs(off) * MM_PER_FT, abs(box.Max.Z) * MM_PER_FT)
        if abs(off) * MM_PER_FT > 1.0:
            notes.append("WARNING: the section line is {:.0f} mm away from the target line".format(
                abs(off) * MM_PER_FT))
        facing = actual_look(section).DotProduct(geo["look"])
        if facing < math.cos(math.radians(0.1)):
            notes.append("WARNING: the section does not look towards the clicked side ({:.0f} degrees off)".format(
                math.degrees(math.acos(max(-1.0, min(1.0, facing))))))
        transaction.Commit()
    except Exception:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    return section, action, notes, diag


def show_section(section, plan):
    """Mostra solo la sezione: la rende la vista attiva, senza affiancare le altre schede (l'affiancamento mostrava
    tutte le viste aperte). La pianta resta aperta nella sua scheda. Se la sezione e' gia' aperta si aggiorna nella
    sua finestra, senza aprirne una nuova."""
    open_ids = [get_id(ui_view.ViewId) for ui_view in uidoc.GetOpenUIViews()]
    already = get_id(section.Id) in open_ids
    uidoc.ActiveView = section
    return "already open: updated and shown" if already else "opened (the plan stays open in its own tab)"


# ----------------------------------------------------------------------------------------------------------------
# Impostazioni e report
# ----------------------------------------------------------------------------------------------------------------
def load_settings(config):
    settings = dict(DEFAULTS)
    try:
        for key in ("margin_mm", "below_mm", "above_mm"):
            settings[key] = float(config.get_option(key, DEFAULTS[key]))
        for key in ("type_name", "template_name", "section_name"):
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
    output.print_md("# Quick Section report")
    if info.get("error"):
        output.print_md("**Nothing was changed.** {}".format(info["error"]))
    for label, key in (("Mode", "mode"), ("Section", "section"), ("Result", "action"), ("Window", "window"),
                       ("Direction", "direction"), ("Section line", "line"), ("Offset from the element", "offset"),
                       ("Depth (far clip) [mm]", "depth"), ("Width [mm]", "width"), ("Height", "height"),
                       ("Notes", "notes")):
        if info.get(key):
            output.print_md("**{}:** {}".format(label, info[key]))
    if info.get("detail"):
        output.print_md("**Technical detail**")
        output.print_code(info["detail"])


def screen_angle(vector, axes):
    """Angolo sullo schermo della pianta, in gradi: 0 = destra, 90 = alto, 180 = sinistra, -90 = basso."""
    return math.degrees(math.atan2(vector.DotProduct(axes[1]), vector.DotProduct(axes[0])))


def finish(info, ok):
    """ok False: errore (nulla e' cambiato) oppure sezione spostata con un avviso da leggere."""
    if want_report(1, 0 if ok else 1, "section placed"):
        print_report(info)


def main():
    plan = doc.ActiveView
    if not isinstance(plan, DB.ViewPlan) or plan.IsTemplate:
        forms.alert("Run Quick Section from a plan view (floor or ceiling plan).", title="Quick Section",
                    exitscript=True)
    if not section_types():
        forms.alert("The model has no section view type.", title="Quick Section", exitscript=True)

    config = script.get_config(CONFIG_NAME)
    settings = load_settings(config)
    changed_keys = ()  # tipo e template della sezione esistente cambiano solo se cambiati ora nella finestra
    previous_name = section_name(settings)
    if __shiftclick__:  # noqa: F821
        previous = settings
        settings = QuickSection_ui.SettingsWindow(
            script.get_bundle_file(XAML_FILE_NAME), settings, [type_name(t) for t in section_types()],
            sorted(v.Name for v in section_templates()), default_section_name()).show()
        if settings is None:
            script.exit()
        save_settings(config, settings)
        changed_keys = tuple(k for k in VIEW_KEYS if settings[k] != previous[k])

    info = {}
    try:
        axes = view_axes(plan)
        z_min, z_max, info["height"] = vertical_range(plan, settings)
        margin = settings["margin_mm"] / MM_PER_FT
        side_prompt = "Quick Section: click on the side to look at; the distance from the section line is the depth"
        try:
            reference = uidoc.Selection.PickObject(
                ObjectType.PointOnElement, "Quick Section: click an element to cut parallel to it, also in a linked "
                                           "model (ESC: click two points instead)")
        except OperationCanceledException:
            reference = None
        if reference is not None:
            element, transform, model, normal = picked_element(reference)
            line = element_line(element, transform, reference.GlobalPoint, normal)
            info["mode"] = "{} of {}".format({"axis": "Along the axis", "face": "Along the clicked face",
                                              "vertical": "Through the riser, along a view axis"}[line["kind"]],
                                             element_label(element, model))
            click = pick_point(side_prompt)
            # griglie e linee non hanno ingombro: la linea di sezione resta sul loro asse
            body = [] if isinstance(element, (DB.Grid, DB.CurveElement, DB.ReferencePlane)) else \
                body_points(element, transform)
            if line["kind"] == "vertical":
                geo = riser_frame(line["origin"], line["points"], click, axes, margin, body)
            else:
                geo = frame(line["origin"], line["tangent"], line["points"], click, margin, body)
            if geo["back"]:
                info["offset"] = ("section line moved back {:.0f} mm, behind the element ({:g} mm clear), so that "
                                  "it is seen whole and not cut; the depth grows by the same amount").format(
                                      geo["back"] * MM_PER_FT, BACK_CLEARANCE_MM)
        else:
            p1 = pick_point("Quick Section: click the first point of the section line (ESC to cancel)")
            p2 = pick_point("Quick Section: click the second point (the line snaps to the nearest axis of the view)")
            tangent = snapped_axis(p1, p2, axes)
            info["mode"] = "Two points, {} on screen".format(
                "horizontal" if tangent.IsAlmostEqualTo(axes[0]) else "vertical")
            geo = frame(p1, tangent, [p1, p2], click=pick_point(side_prompt), margin=margin)
        info["depth"] = "{:.0f}".format(geo["depth"] * MM_PER_FT)
        info["width"] = "{:.0f} (side margin {:g} mm included)".format(geo["half"] * 2 * MM_PER_FT,
                                                                       settings["margin_mm"])
        info["height"] = "{:.0f} to {:.0f} mm, internal coordinates ({})".format(
            z_min * MM_PER_FT, z_max * MM_PER_FT, info["height"])
        section, info["action"], notes, info["line"] = place_section(plan, geo, z_min, z_max, settings,
                                                                     changed_keys, previous_name)
        info["section"] = "{} {}".format(section.Name, output.linkify(section.Id))
        info["notes"] = "; ".join(notes)
        # diagnosi del verso: lato cliccato e direzione reale della vista, come angoli sullo schermo
        info["direction"] = "clicked side {:.0f} deg, section looks {:.0f} deg (on screen: 0 = right, 90 = up)".format(
            screen_angle(geo["look"], axes), screen_angle(actual_look(section), axes))
    except SectionError as ex:
        info["error"] = str(ex)
        finish(info, False)
        return
    except Exception as ex:
        info["error"] = "The section could not be placed: {}".format(ex)
        info["detail"] = traceback.format_exc() if traceback else ""
        finish(info, False)
        return

    try:
        info["window"] = show_section(section, plan)
    except Exception as ex:
        info["window"] = "not opened: {}".format(ex)
    finish(info, "WARNING" not in info["notes"])


main()
