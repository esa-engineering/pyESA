# -*- coding: utf-8 -*-
# Intestazione dello script con metadati
__title__   = "Linked Room Tag\nin Multiple Views"
__doc__     = """Version = 5.1
Date    = 27.08.2026
________________________________________________________________
Automatically tag the rooms visible in the selected views
with the chosen tag, taking the rooms from the current model or
from a linked model selected by the user.
________________________________________________________________
Author(s): Andrea Patti
"""

# -------------------------------
# SEZIONE IMPORT MODULI
# -------------------------------
import os.path as op

import clr
# Gli assembly WPF servono per costruire in codice le checkbox delle viste.
# Referenziati esplicitamente per non dipendere dall'ordine di import ne da
# cosa il motore IronPython di pyRevit ha gia caricato.
for _asm in ("WindowsBase", "PresentationCore", "PresentationFramework"):
    try:
        clr.AddReference(_asm)
    except Exception:
        pass

from System.Collections.Generic import List
from System.Windows import Thickness, Visibility
from System.Windows import Controls

from pyrevit import revit, script, DB, forms
from Autodesk.Revit.DB import (
    Transaction,
    RevitLinkInstance,
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    LinkElementId,
    ElementId,
    Line,
    UV,
    XYZ,
)

# -------------------------------
# SEZIONE INIZIALIZZAZIONE
# -------------------------------
doc = revit.doc
uidoc = revit.uidoc
app = doc.Application
output = script.get_output()

try:
    RVT_VER = int(app.VersionNumber)
except Exception:
    # Non blocchiamo lo script: l'helper di compatibilita degradera
    # sul ramo con doppio tentativo.
    RVT_VER = 0

XY_TOL = 1.0 / 304.8 * 10.0    # ~10 mm di tolleranza in piedi sul test XY
Z_TOL  = 1.0 / 304.8 * 10.0    # ~10 mm di tolleranza in piedi sul test Z
MIN_SEG = 0.01                 # piedi, ~3 mm: sopra la short curve tolerance di Revit
DEFAULT_ROOM_HEIGHT = 8.0      # piedi, fallback se l'estensione verticale non e' calcolabile
HUGE_Z = 3280.0                # piedi, ~1 km: sostituto finito di un view range illimitato
MAX_SEGMENT_TESTS = 40000      # valvola di sicurezza sul test segmento-segmento
PREVIEW_ROW_LIMIT = 250        # righe di dettaglio mostrate in anteprima

# Altezza della fascia considerata sopra il livello della vista, quando
# l'opzione relativa e' attiva. Cambiare qui per modificarla per tutte le viste.
LEVEL_BAND_METERS = 2.0
LEVEL_BAND_FT = LEVEL_BAND_METERS / 0.3048

HOST_LABEL = "<< Modello corrente >>"

# Grafica della finestra di setup, nella stessa cartella dello script.
XAML_FILE_NAME = "TagLinkedRoomsUI.xaml"

# -------------------------------------------------------------------------
# COMPORTAMENTI FISSI
# -------------------------------------------------------------------------
# Queste funzioni non sono piu opzionali: sono sempre attive. Restano come
# costanti nominate perche il codice a valle resti leggibile e perche
# riportarle a opzione, se un giorno servisse, sia un intervento localizzato.
#
# opt_crop      limita le rooms alla crop region / scope box della vista
# opt_skip_dup  salta le rooms che hanno gia un tag nella vista
# opt_no_leader crea i tag senza leader
# opt_move_pt   riposiziona il tag nella porzione visibile della room
#
# La fascia verticale e' sempre attiva, ma la sua altezza arriva dall'input
# della finestra: 0 significa nessuna fascia, si usa l'intero view range.
opt_crop = True
opt_skip_dup = True
opt_no_leader = True
opt_move_pt = True

# Avvisi raccolti durante l'analisi e stampati in coda al report
WARNINGS = []


def warn(message):
    """Registra un avviso: niente degradazione silenziosa."""
    if message not in WARNINGS:
        WARNINGS.append(message)


# -------------------------------
# SEZIONE HELPER - COMPATIBILITA API
# -------------------------------
def get_element_id_value(element_id):
    """ElementId -> int.
    Compatibile con Revit <= 2025 (.IntegerValue) e >= 2026 (.Value)."""
    if element_id is None:
        return -1
    if RVT_VER >= 2026:
        return element_id.Value
    if RVT_VER > 0:
        return element_id.IntegerValue
    # Versione non determinata: proviamo entrambe le proprieta.
    try:
        return element_id.Value
    except Exception:
        return element_id.IntegerValue


def get_param_double(element, bip, default=0.0):
    """Legge un parametro double, restituendo default se assente o vuoto."""
    try:
        p = element.get_Parameter(bip)
        if p and p.HasValue:
            return p.AsDouble()
    except Exception:
        pass
    return default


def get_param_elem_id(element, bip):
    """Legge un parametro ElementId, restituendo None se assente o invalido."""
    try:
        p = element.get_Parameter(bip)
        if p:
            eid = p.AsElementId()
            if eid and eid != ElementId.InvalidElementId:
                return eid
    except Exception:
        pass
    return None


def get_rooms_category_id():
    """ElementId della categoria Rooms nel documento host."""
    try:
        cat = DB.Category.GetCategory(doc, BuiltInCategory.OST_Rooms)
        if cat is not None:
            return cat.Id
    except Exception:
        pass
    return None


ROOMS_CAT_ID = get_rooms_category_id()


# -------------------------------
# SEZIONE HELPER - GEOMETRIA 2D
# -------------------------------
def point_in_polygon(x, y, polygon):
    """Ray casting. polygon = lista di tuple (x, y), chiusa implicitamente."""
    n = len(polygon)
    if n < 3:
        return True
    inside = False
    j = n - 1
    for i in range(n):
        xi, yi = polygon[i]
        xj, yj = polygon[j]
        if (yi > y) != (yj > y):
            denom = yj - yi
            if abs(denom) > 1e-12:
                x_int = xi + (y - yi) * (xj - xi) / denom
                if x < x_int:
                    inside = not inside
        j = i
    return inside


def polygon_bounds(polygon):
    """(min_x, min_y, max_x, max_y) di un poligono."""
    xs = [p[0] for p in polygon]
    ys = [p[1] for p in polygon]
    return (min(xs), min(ys), max(xs), max(ys))


def polygon_signed_area(polygon):
    """Area con segno (shoelace). Positiva se antiorario."""
    total = 0.0
    n = len(polygon)
    for i in range(n):
        x0, y0 = polygon[i]
        x1, y1 = polygon[(i + 1) % n]
        total += x0 * y1 - x1 * y0
    return total / 2.0


def ensure_ccw(polygon):
    """Restituisce il poligono orientato in senso antiorario (normale +Z)."""
    if polygon_signed_area(polygon) < 0.0:
        return list(reversed(polygon))
    return polygon


def polygon_centroid(polygon):
    """Baricentro dell'area del poligono, con fallback sulla media dei vertici."""
    n = len(polygon)
    if n == 0:
        return None
    a2 = 0.0
    cx = 0.0
    cy = 0.0
    for i in range(n):
        x0, y0 = polygon[i]
        x1, y1 = polygon[(i + 1) % n]
        cross = x0 * y1 - x1 * y0
        a2 += cross
        cx += (x0 + x1) * cross
        cy += (y0 + y1) * cross
    if abs(a2) < 1e-12:
        return (sum(p[0] for p in polygon) / float(n),
                sum(p[1] for p in polygon) / float(n))
    return (cx / (3.0 * a2), cy / (3.0 * a2))


def _orient(ax, ay, bx, by, cx, cy):
    return (bx - ax) * (cy - ay) - (by - ay) * (cx - ax)


def segments_intersect(a1, a2, b1, b2):
    """Intersezione propria fra due segmenti. Il contatto collineare
    non e' considerato sovrapposizione utile."""
    d1 = _orient(b1[0], b1[1], b2[0], b2[1], a1[0], a1[1])
    d2 = _orient(b1[0], b1[1], b2[0], b2[1], a2[0], a2[1])
    d3 = _orient(a1[0], a1[1], a2[0], a2[1], b1[0], b1[1])
    d4 = _orient(a1[0], a1[1], a2[0], a2[1], b2[0], b2[1])
    return ((d1 > 0) != (d2 > 0)) and ((d3 > 0) != (d4 > 0))


def polygons_overlap(poly_a, poly_b):
    """Sovrapposizione area contro area fra due poligoni.

    Tre test in cascata, dal piu economico al piu costoso:
    1. AABB disgiunti -> nessuna sovrapposizione
    2. un vertice di A dentro B, oppure un vertice di B dentro A
       (il doppio verso copre il caso 'B interamente dentro A')
    3. intersezione segmento-segmento, che copre il caso a croce in cui
       nessun vertice dell'uno cade dentro l'altro
    """
    if not poly_a or not poly_b:
        return False

    ax0, ay0, ax1, ay1 = polygon_bounds(poly_a)
    bx0, by0, bx1, by1 = polygon_bounds(poly_b)
    if ax1 < bx0 - XY_TOL or bx1 < ax0 - XY_TOL:
        return False
    if ay1 < by0 - XY_TOL or by1 < ay0 - XY_TOL:
        return False

    for x, y in poly_a:
        if point_in_polygon(x, y, poly_b):
            return True
    for x, y in poly_b:
        if point_in_polygon(x, y, poly_a):
            return True

    n = len(poly_a)
    m = len(poly_b)
    if n * m > MAX_SEGMENT_TESTS:
        warn("Test segmento-segmento saltato su un poligono con {}x{} lati: "
             "possibili falsi negativi su sovrapposizioni a croce.".format(n, m))
        return False

    for i in range(n):
        a1 = poly_a[i]
        a2 = poly_a[(i + 1) % n]
        for j in range(m):
            b1 = poly_b[j]
            b2 = poly_b[(j + 1) % m]
            if segments_intersect(a1, a2, b1, b2):
                return True
    return False


# -------------------------------
# SEZIONE HELPER - LIVELLI E VIEW RANGE
# -------------------------------
def level_internal_elevation(level):
    """Quota del livello nel sistema di coordinate INTERNO del suo documento.

    ATTENZIONE, e' il punto piu insidioso di tutto lo script.
    Level.Elevation restituisce la quota riferita al Survey Point (quindi
    l'altitudine sul livello del mare, centinaia di piedi) quando il parametro
    'Elevation Base' del livello e' impostato su Survey Point. La geometria
    invece e' SEMPRE in coordinate interne: bounding box, punti di
    inserimento, Transform dei link. Confrontare le due scale senza
    correzione produce uno scostamento pari all'altitudine del sito e fa
    escludere tutte le rooms dal test verticale.

    ProjectElevation e' la variante riferita al Project Base Point e coincide
    con la quota interna nel caso normale (PBP a quota 0). La coerenza viene
    verificata contro la geometria da check_level_geometry_consistency: se il
    controllo fallisce, il test verticale non e' attendibile e viene detto.
    """
    if level is None:
        return None
    try:
        return level.ProjectElevation
    except Exception:
        pass
    try:
        return level.Elevation
    except Exception:
        return None


def get_sorted_levels(document):
    """Livelli del documento ordinati per quota interna."""
    levels = FilteredElementCollector(document).OfClass(DB.Level)\
                                               .WhereElementIsNotElementType()\
                                               .ToElements()
    return sorted(levels, key=lambda l: (level_internal_elevation(l) or 0.0))


def check_level_geometry_consistency(rooms, limit=40):
    """Verifica che le quote dei livelli siano coerenti con la geometria.

    Confronta la quota interna del livello di ogni room con la Z del suo punto
    di inserimento, che per una room sta sul piano del suo livello. Uno scarto
    sistematico significa che le quote dei livelli non sono confrontabili con
    la geometria, ed e' esattamente il difetto che rendeva inutilizzabile il
    test verticale.

    Restituisce (rooms controllate, scarto massimo in piedi).
    """
    worst = 0.0
    checked = 0
    for room in rooms[:limit]:
        try:
            lvl_z = level_internal_elevation(room.Level)
            loc = room.Location
            if lvl_z is None or not isinstance(loc, DB.LocationPoint):
                continue
            checked += 1
            delta = abs(loc.Point.Z - lvl_z)
            if delta > worst:
                worst = delta
        except Exception:
            continue
    return checked, worst


def _special_plan_level_ids():
    """ElementId sentinella usati da PlanViewRange (Unlimited, Current, ...).
    Restituisce (dict nome -> ElementId, tutti_risolti)."""
    special = {}
    names = ("Unlimited", "Current", "LevelAbove", "LevelBelow")
    for name in names:
        try:
            special[name] = getattr(DB.PlanViewRange, name)
        except Exception:
            pass
    return special, len(special) == len(names)


SPECIAL_PLAN_IDS, SPECIAL_PLAN_OK = _special_plan_level_ids()
if not SPECIAL_PLAN_OK:
    warn("Non tutti i valori sentinella di PlanViewRange sono accessibili: "
         "i view range 'Unlimited' / 'Level Above' / 'Level Below' potrebbero "
         "essere interpretati come limitati al livello della vista.")


def _resolve_plane_elevation(view, view_range, plane, host_levels):
    """Elevazione assoluta (piedi) di un piano del view range.
    Restituisce None se il piano e' 'Unlimited'."""
    level_id = view_range.GetLevelId(plane)
    offset = view_range.GetOffset(plane)

    gen_level = view.GenLevel
    base_elev = level_internal_elevation(gen_level)
    if base_elev is None:
        base_elev = 0.0

    unlimited = SPECIAL_PLAN_IDS.get("Unlimited")
    if unlimited is not None and level_id == unlimited:
        return None

    current = SPECIAL_PLAN_IDS.get("Current")
    if level_id == ElementId.InvalidElementId or (current is not None and level_id == current):
        return base_elev + offset

    above = SPECIAL_PLAN_IDS.get("LevelAbove")
    if above is not None and level_id == above:
        nxt = None
        for lvl in host_levels:
            lvl_z = level_internal_elevation(lvl)
            if lvl_z is not None and lvl_z > base_elev + 1e-6:
                nxt = lvl_z
                break
        return (nxt if nxt is not None else base_elev) + offset

    below = SPECIAL_PLAN_IDS.get("LevelBelow")
    if below is not None and level_id == below:
        prev = None
        for lvl in host_levels:
            lvl_z = level_internal_elevation(lvl)
            if lvl_z is not None and lvl_z < base_elev - 1e-6:
                prev = lvl_z
        return (prev if prev is not None else base_elev) + offset

    lvl = doc.GetElement(level_id)
    if lvl is None and get_element_id_value(level_id) < 0:
        # Sentinella non riconosciuta: lo dichiariamo invece di ignorarlo.
        warn("Vista '{}': valore sentinella {} del view range non riconosciuto, "
             "trattato come livello della vista.".format(
                 view.Name, get_element_id_value(level_id)))
    lvl_z = level_internal_elevation(lvl) if lvl is not None else None
    return ((lvl_z if lvl_z is not None else base_elev) + offset)


def get_view_range_info(view, host_levels):
    """Estensione verticale della vista in coordinate host.
    top / low possono essere None = illimitato."""
    try:
        view_range = view.GetViewRange()
    except Exception:
        return None
    if not view_range:
        return None

    try:
        top = _resolve_plane_elevation(view, view_range, DB.PlanViewPlane.TopClipPlane, host_levels)
        cut = _resolve_plane_elevation(view, view_range, DB.PlanViewPlane.CutPlane, host_levels)
        bottom = _resolve_plane_elevation(view, view_range, DB.PlanViewPlane.BottomClipPlane, host_levels)
        depth = _resolve_plane_elevation(view, view_range, DB.PlanViewPlane.ViewDepthPlane, host_levels)
    except Exception as err:
        warn("Vista '{}': view range non leggibile ({}).".format(view.Name, err))
        return None

    # Il limite inferiore utile e' il piu basso fra bottom e view depth
    if bottom is None or depth is None:
        low = None
    else:
        low = min(bottom, depth)

    gen_level = view.GenLevel
    base_elev = level_internal_elevation(gen_level)
    if base_elev is None:
        base_elev = 0.0

    return {"top": top, "cut": cut, "bottom": bottom, "depth": depth,
            "low": low, "base": base_elev, "gen_level": gen_level}


def format_range(vr):
    def fmt(v):
        return "illimitato" if v is None else "{:.2f}".format(v)
    return "Top: {} | Cut: {} | Bottom: {} | Depth: {}".format(
        fmt(vr["top"]), fmt(vr["cut"]), fmt(vr["bottom"]), fmt(vr["depth"]))


def get_vertical_limits(view_range, use_band):
    """Estremi verticali effettivi del test, in coordinate interne.
    None su un lato significa illimitato.

    Senza fascia sono gli estremi del view range.
    Con fascia sono [quota del livello, quota + LEVEL_BAND_FT], intersecati
    col view range. La fascia non puo' allargare il view range, solo
    restringerlo. Serve a non intercettare le rooms del piano superiore, che
    un view range con Top generoso o illimitato includerebbe: con la fascia
    attiva entrambi i limiti sono sempre finiti.
    """
    low = view_range["low"]
    top = view_range["top"]
    if not use_band:
        return low, top

    base = view_range["base"]
    band_low = base
    band_top = base + LEVEL_BAND_FT

    if low is not None and low > band_low:
        band_low = low
    if top is not None and top < band_top:
        band_top = top
    if band_top < band_low:
        band_top = band_low
    return band_low, band_top


def z_overlaps(z_min, z_max, low, top):
    """Sovrapposizione fra l'estensione verticale della room e i limiti della
    vista. Sovrapposizione, non contenimento: basta che si intersechino.
    None su un limite = illimitato su quel lato."""
    if top is not None and z_min > top + Z_TOL:
        return False
    if low is not None and z_max < low - Z_TOL:
        return False
    return True


def view_solid_z_extent(view_range, use_band=False):
    """Estensione verticale finita per costruire il solido di vista.
    Sostituisce i limiti illimitati con +/- HUGE_Z dal livello della vista."""
    base = view_range["base"]
    z_low, z_top = get_vertical_limits(view_range, use_band)
    if z_top is None:
        z_top = base + HUGE_Z
    if z_low is None:
        z_low = base - HUGE_Z
    if z_top - z_low < MIN_SEG:
        z_top = z_low + MIN_SEG
    return z_low, z_top


# -------------------------------
# SEZIONE HELPER - CROP REGION E SCOPE BOX
# -------------------------------
def get_scope_box(view):
    """Elemento scope box assegnato alla vista, o None."""
    sb_id = get_param_elem_id(view, BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
    if sb_id is not None:
        return doc.GetElement(sb_id)
    return None


def _polygon_from_crop_shape(view):
    """Poligono XY (coordinate modello host) della crop region reale.
    Le curve sono gia in coordinate mondo: la rotazione da scope box e'
    quindi gestita implicitamente, senza conoscerne l'angolo."""
    try:
        manager = view.GetCropRegionShapeManager()
        loops = list(manager.GetCropShape())
    except Exception:
        return None
    if not loops:
        return None

    polygon = []
    try:
        for curve in loops[0]:
            for p in curve.Tessellate():
                if polygon:
                    lx, ly = polygon[-1]
                    if abs(lx - p.X) < 1e-9 and abs(ly - p.Y) < 1e-9:
                        continue
                polygon.append((p.X, p.Y))
    except Exception:
        return None

    if len(polygon) >= 2:
        fx, fy = polygon[0]
        lx, ly = polygon[-1]
        if abs(fx - lx) < 1e-9 and abs(fy - ly) < 1e-9:
            polygon.pop()

    if len(polygon) < 3:
        return None
    return polygon


def _polygon_from_bbox(bbox):
    """Poligono XY in coordinate mondo dei 4 spigoli di un BoundingBoxXYZ,
    applicandone la Transform. Indispensabile per box RUOTATI: Min e Max
    sono in coordinate locali, non mondo."""
    try:
        transform = bbox.Transform
        mn = bbox.Min
        mx = bbox.Max
        corners = [
            XYZ(mn.X, mn.Y, mn.Z),
            XYZ(mx.X, mn.Y, mn.Z),
            XYZ(mx.X, mx.Y, mn.Z),
            XYZ(mn.X, mx.Y, mn.Z),
        ]
        polygon = []
        for c in corners:
            w = transform.OfPoint(c)
            polygon.append((w.X, w.Y))
        return polygon
    except Exception:
        return None


def build_view_xy_polygon(view):
    """Poligono XY (coordinate host) che delimita la vista in pianta.
    Restituisce (polygon | None, descrizione). None = nessun limite XY."""
    if view.CropBoxActive:
        polygon = _polygon_from_crop_shape(view)
        if polygon:
            return polygon, "crop shape ({} vertici)".format(len(polygon))
        try:
            crop_box = view.CropBox
            if crop_box:
                polygon = _polygon_from_bbox(crop_box)
                if polygon:
                    warn("Vista '{}': crop shape non leggibile, uso il crop box "
                         "rettangolare con la sua Transform.".format(view.Name))
                    return polygon, "crop box (con transform)"
        except Exception:
            pass

    scope_box = get_scope_box(view)
    if scope_box is not None:
        try:
            bbox = scope_box.get_BoundingBox(None)
            if bbox:
                polygon = _polygon_from_bbox(bbox)
                if polygon:
                    return polygon, "scope box '{}'".format(scope_box.Name)
        except Exception:
            pass

    return None, "nessun limite XY"


# -------------------------------
# SEZIONE HELPER - SOLIDO DI VISTA
# -------------------------------
def curveloop_from_polygon(polygon, z):
    """CurveLoop piano e chiuso a quota z da un poligono XY."""
    pts = []
    for x, y in polygon:
        p = XYZ(x, y, z)
        if pts and pts[-1].DistanceTo(p) < MIN_SEG:
            continue
        pts.append(p)
    if len(pts) >= 2 and pts[0].DistanceTo(pts[-1]) < MIN_SEG:
        pts.pop()
    if len(pts) < 3:
        return None

    loop = DB.CurveLoop()
    try:
        for i in range(len(pts)):
            a = pts[i]
            b = pts[(i + 1) % len(pts)]
            if a.DistanceTo(b) < MIN_SEG:
                return None
            loop.Append(Line.CreateBound(a, b))
    except Exception:
        return None
    return loop


def build_view_solid(polygon, z_low, z_top):
    """Solido che rappresenta il volume visibile della vista, in coordinate host.
    Restituisce None se la geometria non e' estrudibile."""
    if not polygon:
        return None
    loop = curveloop_from_polygon(ensure_ccw(polygon), z_low)
    if loop is None:
        return None
    try:
        loops = List[DB.CurveLoop]()
        loops.Add(loop)
        return DB.GeometryCreationUtilities.CreateExtrusionGeometry(
            loops, XYZ.BasisZ, z_top - z_low)
    except Exception as err:
        warn("Estrusione del volume di vista non riuscita ({}): "
             "uso il modo geometrico.".format(err))
        return None


def build_outline_in_link(polygon, z_low, z_top, inverse_transform):
    """Outline (AABB) del volume di vista espresso in coordinate del link.
    Serve come quick filter davanti al filtro a solido, che e' uno slow filter.
    L'AABB di una regione trasformata e' conservativo: non perde elementi."""
    if not polygon:
        return None
    x0, y0, x1, y1 = polygon_bounds(polygon)
    corners = []
    for x in (x0, x1):
        for y in (y0, y1):
            for z in (z_low, z_top):
                corners.append(inverse_transform.OfPoint(XYZ(x, y, z)))
    xs = [c.X for c in corners]
    ys = [c.Y for c in corners]
    zs = [c.Z for c in corners]
    try:
        return DB.Outline(XYZ(min(xs), min(ys), min(zs)),
                          XYZ(max(xs), max(ys), max(zs)))
    except Exception:
        return None


# -------------------------------
# SEZIONE HELPER - GEOMETRIA ROOM
# -------------------------------
BOUNDARY_OPTIONS = DB.SpatialElementBoundaryOptions()
try:
    BOUNDARY_OPTIONS.SpatialElementBoundaryLocation = \
        DB.SpatialElementBoundaryLocation.Finish
except Exception:
    pass


def get_room_location_point(room):
    """Punto di inserimento della room (coordinate del suo documento)."""
    try:
        location = room.Location
        if isinstance(location, DB.LocationPoint):
            return location.Point
    except Exception:
        pass
    try:
        bbox = room.get_BoundingBox(None)
        if bbox:
            return (bbox.Min + bbox.Max) / 2.0
    except Exception:
        pass
    return None


def get_room_boundary_points(room):
    """Punti del contorno esterno della room, in coordinate del suo documento.
    Fra i loop restituiti scegliamo quello di perimetro maggiore, che e'
    il contorno esterno; i loop interni (isole) non servono al test XY."""
    try:
        loops = room.GetBoundarySegments(BOUNDARY_OPTIONS)
    except Exception:
        return None
    if not loops:
        return None

    best_pts = None
    best_len = -1.0
    for loop in loops:
        pts = []
        length = 0.0
        try:
            for seg in loop:
                crv = seg.GetCurve()
                for p in crv.Tessellate():
                    if pts and pts[-1].DistanceTo(p) < 1e-9:
                        continue
                    if pts:
                        length += pts[-1].DistanceTo(p)
                    pts.append(p)
        except Exception:
            continue
        if len(pts) >= 3 and length > best_len:
            best_len = length
            best_pts = pts

    if best_pts is None:
        return None
    if len(best_pts) >= 2 and best_pts[0].DistanceTo(best_pts[-1]) < 1e-9:
        best_pts.pop()
    if len(best_pts) < 3:
        return None
    return best_pts


# Conteggio dei metodi con cui e' stata risolta la quota di base delle rooms.
# Se qui compare 'LocationPoint.Z' o 'fallback 0.0' significa che il livello
# della room non e' risolvibile: e' la causa piu probabile di un test verticale
# sbagliato in modo sistematico.
BASE_ELEV_METHODS = {}


def get_room_base_elevation(room, room_doc):
    """Quota del livello della room, nel suo documento, con fallback in cascata.

    L'ultimo fallback usa la Z del punto di inserimento: per una room quel
    punto sta sul piano del suo livello, quindi e' una stima corretta e non
    puo fallire. Meglio di un 0.0 silenzioso, che sposterebbe la room
    all'origine del modello e la farebbe scartare da ogni view range.
    """
    method = None
    elevation = None

    try:
        elevation = level_internal_elevation(room.Level)
        if elevation is not None:
            method = "Room.Level"
    except Exception:
        pass

    if elevation is None:
        try:
            lid = room.LevelId
            if lid is not None and lid != ElementId.InvalidElementId:
                elevation = level_internal_elevation(room_doc.GetElement(lid))
                if elevation is not None:
                    method = "LevelId"
        except Exception:
            pass

    if elevation is None:
        bip = getattr(BuiltInParameter, "ROOM_LEVEL_ID", None)
        if bip is not None:
            lid = get_param_elem_id(room, bip)
            if lid is not None:
                elevation = level_internal_elevation(room_doc.GetElement(lid))
                if elevation is not None:
                    method = "ROOM_LEVEL_ID"

    if elevation is None:
        try:
            loc = room.Location
            if isinstance(loc, DB.LocationPoint):
                elevation = loc.Point.Z
                method = "LocationPoint.Z"
        except Exception:
            pass

    if elevation is None:
        elevation = 0.0
        method = "fallback 0.0"

    BASE_ELEV_METHODS[method] = BASE_ELEV_METHODS.get(method, 0) + 1
    return elevation


def get_room_z_span(room, room_doc):
    """Estensione verticale (z_min, z_max, base_elev) della room nel suo
    documento, RICOSTRUITA da Level / Base Offset / Upper Limit / Limit Offset.

    E' una ricostruzione nominale: dipende da come sono compilati i parametri
    della room. Dove disponibile va preferito il bounding box reale, vedi
    get_room_z_span_host_from_bbox.
    """
    base_elev = get_room_base_elevation(room, room_doc)

    lower_offset = get_param_double(room, BuiltInParameter.ROOM_LOWER_OFFSET, 0.0)
    z_min = base_elev + lower_offset

    upper_offset = get_param_double(room, BuiltInParameter.ROOM_UPPER_OFFSET, 0.0)

    upper_level_elev = None
    up_id = get_param_elem_id(room, BuiltInParameter.ROOM_UPPER_LEVEL)
    if up_id is not None:
        upper_level_elev = level_internal_elevation(room_doc.GetElement(up_id))
    if upper_level_elev is None:
        upper_level_elev = base_elev

    z_max = upper_level_elev + upper_offset

    if z_max - z_min < 1e-6:
        # Room senza altezza utile: fallback sull'altezza reale o su un default
        height = get_param_double(room, BuiltInParameter.ROOM_HEIGHT, 0.0)
        if height <= 1e-6:
            height = DEFAULT_ROOM_HEIGHT
        z_max = z_min + height

    return z_min, z_max, base_elev


# Conteggio della fonte usata per l'estensione verticale, per diagnostica.
Z_SOURCE_COUNTS = {}


def get_room_z_span_host_from_bbox(room, transform):
    """Estensione verticale (z_min, z_max) in coordinate HOST ricavata dal
    bounding box reale della room, trasformato.

    Non ricostruisce niente dai parametri: se i volumi delle rooms sono
    calcolati, questa e' la fonte piu affidabile, perche riflette la geometria
    che Revit usa davvero. Restituisce None se il bbox non e' disponibile.
    """
    try:
        bbox = room.get_BoundingBox(None)
    except Exception:
        return None
    if bbox is None:
        return None

    try:
        mn = bbox.Min
        mx = bbox.Max
        box_transform = bbox.Transform
    except Exception:
        return None

    zs = []
    for x in (mn.X, mx.X):
        for y in (mn.Y, mx.Y):
            for z in (mn.Z, mx.Z):
                p = XYZ(x, y, z)
                try:
                    if box_transform is not None:
                        p = box_transform.OfPoint(p)
                    zs.append(transform.OfPoint(p).Z)
                except Exception:
                    return None
    if not zs:
        return None
    return min(zs), max(zs)


def room_phase_ok(room, view):
    """Confronto fra la fase della room e quella della vista.
    Applicato solo alle rooms dell'host: nel modello linkato la fase
    dipende dalla mappatura di fase del link, che non replichiamo."""
    room_phase = get_param_elem_id(room, BuiltInParameter.ROOM_PHASE)
    view_phase = get_param_elem_id(view, BuiltInParameter.VIEW_PHASE)
    if room_phase is None or view_phase is None:
        return True
    return get_element_id_value(room_phase) == get_element_id_value(view_phase)


# -------------------------------
# SEZIONE CACHE GEOMETRIA ROOM
# -------------------------------
# Le geometrie delle rooms sono in coordinate HOST e vengono calcolate una
# volta sola, alla prima vista che ne ha bisogno, poi riusate.
ROOM_GEO_CACHE = {}


def get_room_geo(room, ctx):
    """Geometria della room in coordinate host, con cache per ElementId.
    Chiavi: point (XYZ), polygon (lista di tuple XY), z_min, z_max."""
    key = get_element_id_value(room.Id)
    geo = ROOM_GEO_CACHE.get(key)
    if geo is not None:
        return geo

    transform = ctx["link_transform"]

    loc = get_room_location_point(room)
    point_host = transform.OfPoint(loc) if loc is not None else None

    pts = get_room_boundary_points(room)
    polygon = None
    if pts:
        polygon = []
        for p in pts:
            ph = transform.OfPoint(p)
            polygon.append((ph.X, ph.Y))
        if len(polygon) < 3:
            polygon = None

    # --- Estensione verticale, due fonti indipendenti ---
    # A) ricostruita dai parametri della room (nominale)
    z_min_p, z_max_p, base_elev_link = get_room_z_span(room, ctx["room_doc"])
    ref_x = loc.X if loc is not None else 0.0
    ref_y = loc.Y if loc is not None else 0.0
    z_min_param = transform.OfPoint(XYZ(ref_x, ref_y, z_min_p)).Z
    z_max_param = transform.OfPoint(XYZ(ref_x, ref_y, z_max_p)).Z
    if z_max_param < z_min_param:
        z_min_param, z_max_param = z_max_param, z_min_param

    # B) dal bounding box reale della room (preferita quando disponibile)
    bbox_span = get_room_z_span_host_from_bbox(room, transform)
    if bbox_span is not None:
        z_min_host, z_max_host = bbox_span
        z_source = "bounding box"
    else:
        z_min_host, z_max_host = z_min_param, z_max_param
        z_source = "parametri"
    Z_SOURCE_COUNTS[z_source] = Z_SOURCE_COUNTS.get(z_source, 0) + 1

    geo = {"point": point_host, "polygon": polygon,
           "z_min": z_min_host, "z_max": z_max_host,
           # campi di sola diagnostica
           "z_source": z_source,
           "z_min_param": z_min_param, "z_max_param": z_max_param,
           "base_elev_link": base_elev_link,
           "loc_z_link": loc.Z if loc is not None else None,
           "loc_z_host": point_host.Z if point_host is not None else None}
    ROOM_GEO_CACHE[key] = geo
    return geo


# -------------------------------
# SEZIONE POSIZIONAMENTO DEL TAG
# -------------------------------
def _nudge(pt, toward, factor=0.05):
    """Sposta pt verso 'toward' di una frazione della distanza,
    per rientrare dentro il poligono quando pt e' un vertice."""
    return (pt[0] + (toward[0] - pt[0]) * factor,
            pt[1] + (toward[1] - pt[1]) * factor)


def pick_insertion_point(geo, xy_polygon, allow_move):
    """Punto (x, y) dove creare il tag, in coordinate host.

    Vincolo forte: il punto deve stare DENTRO la room, altrimenti
    NewRoomTag non riesce ad associare il tag.
    Vincolo debole: dovrebbe stare anche dentro la crop, altrimenti il
    tag esiste ma non e' visibile nella vista.

    Restituisce (punto, dentro_crop). punto None se non determinabile.
    """
    loc = geo["point"]
    room_poly = geo["polygon"]

    loc_xy = (loc.X, loc.Y) if loc is not None else None

    def in_crop(p):
        if xy_polygon is None:
            return True
        return point_in_polygon(p[0], p[1], xy_polygon)

    # Caso normale: il punto di inserimento e' gia buono.
    if loc_xy is not None and in_crop(loc_xy):
        return loc_xy, True

    if not allow_move or room_poly is None:
        if loc_xy is not None:
            return loc_xy, in_crop(loc_xy)
        return None, False

    # La room e' visibile solo in parte e il suo punto di inserimento cade
    # fuori dalla crop: cerchiamo un punto interno alla room e alla crop.
    centroid = polygon_centroid(room_poly)
    candidates = []
    if centroid is not None:
        candidates.append(centroid)
        # midpoint dei lati e vertici, tirati verso il centroide
        n = len(room_poly)
        for i in range(n):
            x0, y0 = room_poly[i]
            x1, y1 = room_poly[(i + 1) % n]
            candidates.append(_nudge(((x0 + x1) / 2.0, (y0 + y1) / 2.0), centroid))
        for v in room_poly:
            candidates.append(_nudge(v, centroid))
        # vertici della crop tirati verso il centroide della room:
        # copre il caso 'crop piccola interamente dentro la room'
        if xy_polygon is not None:
            for v in xy_polygon:
                candidates.append(_nudge(v, centroid))

    for cand in candidates:
        if point_in_polygon(cand[0], cand[1], room_poly) and in_crop(cand):
            return cand, True

    # Ultima risorsa: griglia grossolana sull'AABB comune.
    if xy_polygon is not None:
        rx0, ry0, rx1, ry1 = polygon_bounds(room_poly)
        cx0, cy0, cx1, cy1 = polygon_bounds(xy_polygon)
        gx0 = max(rx0, cx0)
        gy0 = max(ry0, cy0)
        gx1 = min(rx1, cx1)
        gy1 = min(ry1, cy1)
        if gx1 > gx0 and gy1 > gy0:
            steps = 8
            for i in range(1, steps):
                for j in range(1, steps):
                    cand = (gx0 + (gx1 - gx0) * i / float(steps),
                            gy0 + (gy1 - gy0) * j / float(steps))
                    if point_in_polygon(cand[0], cand[1], room_poly) and in_crop(cand):
                        return cand, True

    # Nessun punto sta in entrambi: creiamo il tag dentro la room e lo
    # dichiariamo fuori crop nel report.
    if loc_xy is not None:
        return loc_xy, False
    if centroid is not None:
        return centroid, False
    return None, False


# -------------------------------
# SEZIONE TAG ESISTENTI
# -------------------------------
def collect_existing_tag_keys(view):
    """Insieme delle chiavi (link_id, room_id) delle rooms gia taggate
    nella vista. Per le rooms del modello host link_id = -1."""
    keys = set()
    try:
        tags = FilteredElementCollector(doc, view.Id)\
            .OfCategory(BuiltInCategory.OST_RoomTags)\
            .WhereElementIsNotElementType()\
            .ToElements()
    except Exception:
        return keys

    for tag in tags:
        key = None
        try:
            tagged = tag.TaggedRoomId
            if tagged is not None:
                linked_id = tagged.LinkedElementId
                if linked_id and linked_id != ElementId.InvalidElementId:
                    key = (get_element_id_value(tagged.LinkInstanceId),
                           get_element_id_value(linked_id))
                else:
                    key = (-1, get_element_id_value(tagged.HostElementId))
        except Exception:
            key = None

        if key is None:
            try:
                room = tag.Room
                if room is not None:
                    key = (-1, get_element_id_value(room.Id))
            except Exception:
                key = None

        if key is not None:
            keys.add(key)

    return keys


# -------------------------------
# SEZIONE STRATEGIE DI IDENTIFICAZIONE
# -------------------------------
def filter_rooms_by_band(rooms, view, ctx):
    """Post-filtro verticale sulla fascia sopra il livello della vista.

    Serve alla strategia L0, dove la selezione la fa Revit col suo view range
    e non e' possibile passargli un limite piu restrittivo. Le altre strategie
    applicano la fascia direttamente nel proprio test.

    Restituisce (rooms tenute, numero di rooms escluse).
    """
    if not ctx["opt_band"]:
        return rooms, 0

    view_range = ctx["view_info"][get_element_id_value(view.Id)]["view_range"]
    if view_range is None:
        return rooms, 0

    low, top = get_vertical_limits(view_range, True)
    kept = []
    removed = 0
    for room in rooms:
        geo = get_room_geo(room, ctx)
        if z_overlaps(geo["z_min"], geo["z_max"], low, top):
            kept.append(room)
        else:
            removed += 1
    return kept, removed



def strategy_revit_visibility(view, ctx):
    """L0 - Rooms dell'host: chiediamo a Revit quali sono visibili.

    Un collector costruito con un view.Id restituisce solo gli elementi
    effettivamente visibili in quella vista. Revit applica da se crop
    region (anche ruotata o non rettangolare), view range con tutti i suoi
    casi speciali, filtro di fase, view filter, design option attiva ed
    elementi nascosti singolarmente.

    Prerequisito: la categoria Rooms deve essere visibile nella vista. Nelle
    viste dove si taggano le rooms e' quasi sempre spenta, quindi viene
    riattivata nella transazione di sola analisi (poi annullata).

    Restituisce (lista di rooms, nota) oppure None se la strategia non e'
    applicabile in questa vista.
    """
    probe = ctx["probe"].get(get_element_id_value(view.Id))
    if probe is None or not probe.get("cat_ok", False):
        return None

    try:
        rooms = FilteredElementCollector(doc, view.Id)\
            .OfCategory(BuiltInCategory.OST_Rooms)\
            .WhereElementIsNotElementType()\
            .ToElements()
    except Exception as err:
        warn("Vista '{}': collector view-scoped non utilizzabile ({}).".format(
            view.Name, err))
        return None

    result = []
    for room in rooms:
        try:
            if room.Area > 0 and room.Location is not None:
                result.append(room)
        except Exception:
            continue

    # La selezione l'ha fatta Revit col suo view range: la fascia e' un limite
    # piu restrittivo che non possiamo passare al collector, quindi va
    # applicata come post-filtro.
    note = ", ".join(probe.get("notes", [])) or "nessuna modifica temporanea"
    result, removed = filter_rooms_by_band(result, view, ctx)
    if removed:
        note += " | {} rooms escluse dalla fascia".format(removed)
    return result, note


def strategy_solid_filter(view, ctx):
    """L1 - Rooms linkate: filtro a solido nel documento del link.

    Il volume visibile della vista viene estruso in coordinate host, portato
    in coordinate del link con la Transform inversa e usato come
    ElementIntersectsSolidFilter. Il test e' volume contro volume, quindi
    gestisce sovrapposizioni parziali, rooms concave, crop ruotate e
    l'estensione verticale reale della room.

    Prerequisito: volumi calcolati nel modello linkato.
    Restituisce (lista di rooms, nota) oppure None se non applicabile.
    """
    if not ctx["use_solid_filter"]:
        # L'utente ha scelto il modo geometrico pur avendo L1 disponibile.
        return None
    if not ctx["volumes_ok"]:
        return None

    xy_polygon = ctx["view_info"][get_element_id_value(view.Id)]["xy_polygon"]
    view_range = ctx["view_info"][get_element_id_value(view.Id)]["view_range"]
    if xy_polygon is None or view_range is None:
        return None

    z_low, z_top = view_solid_z_extent(view_range, ctx["opt_band"])
    solid = build_view_solid(xy_polygon, z_low, z_top)
    if solid is None:
        return None

    inverse = ctx["link_transform"].Inverse
    try:
        solid_link = DB.SolidUtils.CreateTransformed(solid, inverse)
    except Exception as err:
        warn("Vista '{}': trasformazione del solido di vista non riuscita ({}).".format(
            view.Name, err))
        return None

    collector = FilteredElementCollector(ctx["room_doc"])\
        .OfCategory(BuiltInCategory.OST_Rooms)\
        .WhereElementIsNotElementType()

    # Quick filter davanti allo slow filter: l'AABB e' conservativo.
    outline = build_outline_in_link(xy_polygon, z_low, z_top, inverse)
    if outline is not None:
        try:
            collector = collector.WherePasses(
                DB.BoundingBoxIntersectsFilter(outline))
        except Exception:
            pass

    try:
        collector = collector.WherePasses(
            DB.ElementIntersectsSolidFilter(solid_link))
        rooms = collector.ToElements()
    except Exception as err:
        warn("Vista '{}': filtro a solido non applicabile ({}).".format(
            view.Name, err))
        return None

    result = []
    for room in rooms:
        try:
            if room.Area > 0 and room.Location is not None:
                result.append(room)
        except Exception:
            continue
    return result, "volume di vista da {:.2f} a {:.2f}".format(z_low, z_top)


def strategy_geometric(view, ctx):
    """L2 - Fallback geometrico, test AREA contro AREA.

    Differenza sostanziale rispetto alla v3: il test XY confronta il
    contorno della room con quello della crop, non il solo punto di
    inserimento. Una room a cavallo del bordo viene riconosciuta.

    Restituisce (lista di rooms, statistiche di esclusione).
    """
    vid = get_element_id_value(view.Id)
    info = ctx["view_info"][vid]
    view_range = info["view_range"]
    xy_polygon = info["xy_polygon"]

    stats = {"out_z": 0, "out_xy": 0, "out_phase": 0, "no_geo": 0,
             "out_band": 0,
             "z_seen_min": None, "z_seen_max": None, "samples": []}
    if view_range is None:
        return None, stats

    # Limiti effettivi (eventualmente ristretti dalla fascia) e limiti grezzi
    # del view range: il confronto fra i due permette di attribuire ogni
    # esclusione alla causa giusta nel report.
    low, top = get_vertical_limits(view_range, ctx["opt_band"])
    raw_low = view_range["low"]
    raw_top = view_range["top"]

    result = []
    for room in ctx["placed_rooms"]:
        if ctx["source_is_host"] and not room_phase_ok(room, view):
            stats["out_phase"] += 1
            continue

        geo = get_room_geo(room, ctx)
        if geo["point"] is None and geo["polygon"] is None:
            stats["no_geo"] += 1
            continue

        # Estensione verticale complessiva delle rooms viste, per diagnostica.
        if stats["z_seen_min"] is None or geo["z_min"] < stats["z_seen_min"]:
            stats["z_seen_min"] = geo["z_min"]
        if stats["z_seen_max"] is None or geo["z_max"] > stats["z_seen_max"]:
            stats["z_seen_max"] = geo["z_max"]

        def _sample_z(why):
            if len(stats["samples"]) < 5:
                stats["samples"].append({
                    "id": get_element_id_value(room.Id),
                    "z_min": geo["z_min"], "z_max": geo["z_max"], "why": why,
                    "z_source": geo.get("z_source"),
                    "z_min_param": geo.get("z_min_param"),
                    "z_max_param": geo.get("z_max_param"),
                    "base_elev_link": geo.get("base_elev_link"),
                    "loc_z_link": geo.get("loc_z_link"),
                    "loc_z_host": geo.get("loc_z_host")})

        # Test verticale: sovrapposizione fra intervalli, non contenimento.
        if not z_overlaps(geo["z_min"], geo["z_max"], low, top):
            # Se la room sarebbe passata col view range grezzo, e' la fascia
            # ad averla esclusa: e' proprio il caso 'room del piano superiore'.
            if ctx["opt_band"] and z_overlaps(geo["z_min"], geo["z_max"],
                                              raw_low, raw_top):
                stats["out_band"] += 1
                _sample_z("fuori dalla fascia {:.2f} .. {:.2f} "
                          "(primi {:.1f} m sopra il livello)".format(
                              low, top, LEVEL_BAND_METERS))
            else:
                stats["out_z"] += 1
                _sample_z("fuori dal view range")
            continue

        # Test orizzontale: area contro area, con fallback sul punto se il
        # contorno della room non e' leggibile.
        if xy_polygon is not None:
            if geo["polygon"] is not None:
                if not polygons_overlap(geo["polygon"], xy_polygon):
                    stats["out_xy"] += 1
                    continue
            else:
                p = geo["point"]
                if p is None or not point_in_polygon(p.X, p.Y, xy_polygon):
                    stats["out_xy"] += 1
                    continue

        result.append(room)

    return result, stats


def resolve_visible_rooms(view, ctx):
    """Dispatcher delle tre strategie, con verifica incrociata.

    Se la strategia primaria non restituisce nessuna room, ricontrolliamo
    col modo geometrico: e' l'unico modo di distinguere 'in questa vista non
    ci sono rooms' da 'la strategia primaria non ha funzionato'. Il costo si
    paga solo nel caso a zero.

    Restituisce (rooms, nome_strategia, nota, stats).
    """
    if ctx["source_is_host"]:
        primary = strategy_revit_visibility(view, ctx)
        primary_name = "visibilita Revit"
    else:
        primary = strategy_solid_filter(view, ctx)
        primary_name = "filtro a solido"

    if primary is not None:
        rooms, note = primary
        if rooms:
            return rooms, primary_name, note, {}
        geo_rooms, geo_stats = strategy_geometric(view, ctx)
        if geo_rooms:
            warn("Vista '{}': la strategia '{}' ha restituito 0 rooms mentre "
                 "il modo geometrico ne trova {}. Uso il modo geometrico e "
                 "segnalo la discrepanza.".format(
                     view.Name, primary_name, len(geo_rooms)))
            return geo_rooms, "geometrico (fallback)", \
                   "discrepanza con '{}'".format(primary_name), geo_stats
        return [], primary_name, note, {}

    geo_rooms, geo_stats = strategy_geometric(view, ctx)
    if geo_rooms is None:
        return [], "nessuna", "view range non leggibile", geo_stats

    if not ctx["source_is_host"] and not ctx["use_solid_filter"]:
        reason = "modo geometrico scelto dall'utente"
    else:
        reason = "strategia primaria non applicabile"
    return geo_rooms, "geometrico", reason, geo_stats


# -------------------------------
# SEZIONE PREPARAZIONE VISTE PER L'ANALISI
# -------------------------------
def prepare_view_for_probe(view):
    """Rende la vista interrogabile dal collector view-scoped.

    Riattiva la categoria Rooms, che nelle viste dove si taggano le rooms e'
    quasi sempre spenta: un collector view-scoped su categoria nascosta
    restituisce zero elementi.

    Va chiamata dentro la transazione di sola analisi, che viene poi
    annullata: nessuna di queste modifiche sopravvive.
    """
    info = {"cat_ok": True, "notes": []}

    if ROOMS_CAT_ID is None:
        info["cat_ok"] = False
        info["notes"].append("categoria Rooms non risolvibile")
        return info

    try:
        if view.GetCategoryHidden(ROOMS_CAT_ID):
            if view.CanCategoryBeHidden(ROOMS_CAT_ID):
                view.SetCategoryHidden(ROOMS_CAT_ID, False)
                info["notes"].append("categoria Rooms riattivata temporaneamente")
            else:
                info["cat_ok"] = False
                info["notes"].append("categoria Rooms non riattivabile")
    except Exception as err:
        info["cat_ok"] = False
        info["notes"].append("V/G non modificabile ({})".format(err))

    return info


# =========================================================================
# FASE 0 - RACCOLTA DATI E FINESTRA DI SETUP
# =========================================================================

# --- Sorgenti disponibili: modello corrente e link caricati ---
link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
loaded_links = [lk for lk in link_instances if lk.GetLinkDocument() is not None]

source_dict = {HOST_LABEL: None}
for link in loaded_links:
    label = link.Name.split(" : ")[0]
    if label in source_dict:
        label = "{} [id {}]".format(label, get_element_id_value(link.Id))
    source_dict[label] = link

source_labels = [HOST_LABEL] + sorted([k for k in source_dict.keys() if k != HOST_LABEL])

# --- Viste candidate ---
plan_types = (DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan)
all_plan_views = FilteredElementCollector(doc).OfClass(DB.ViewPlan).ToElements()
plan_views = [v for v in all_plan_views if not v.IsTemplate and v.ViewType in plan_types]
plan_views.sort(key=lambda v: v.Name)

if not plan_views:
    forms.alert("Non ci sono piante (Floor Plan / Ceiling Plan) nel progetto.",
                exitscript=True)

# --- Tipi di Room Tag caricati ---
room_tag_types = list(FilteredElementCollector(doc)
                      .OfClass(DB.FamilySymbol)
                      .OfCategory(BuiltInCategory.OST_RoomTags))

if not room_tag_types:
    forms.alert("Non ci sono Room Tags caricati nel progetto.", exitscript=True)

tag_dict = {}
for tt in room_tag_types:
    try:
        type_name = tt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    except Exception:
        type_name = "?"
    full_name = "{} - {}".format(tt.FamilyName, type_name)
    if full_name in tag_dict:
        full_name = "{} [id {}]".format(full_name, get_element_id_value(tt.Id))
    tag_dict[full_name] = tt

tag_labels = sorted(tag_dict.keys())


def source_volumes_ok(label):
    """I volumi delle rooms sono calcolati nel documento della sorgente?
    Prerequisito del filtro a solido, va verificato sul documento giusto."""
    link = source_dict.get(label)
    document = doc if link is None else link.GetLinkDocument()
    if document is None:
        return False, "documento del link non disponibile"
    try:
        computed = bool(
            DB.AreaVolumeSettings.GetAreaVolumeSettings(document).ComputeVolumes)
    except Exception as err:
        return False, "impostazione non leggibile ({})".format(err)
    if computed:
        return True, "volumi delle rooms calcolati"
    return False, ("volumi delle rooms NON calcolati "
                   "(Area and Volume Computations su 'Areas only')")


class TagRoomsWindow(forms.WPFWindow):
    """Finestra unica di setup, grafica definita in TagLinkedRoomsUI.xaml.

    Sostituisce i quattro dialoghi pyRevit precedenti. La disponibilita del
    filtro a solido dipende dalla sorgente scelta, quindi i radio button
    della strategia vengono abilitati o disabilitati quando la combo cambia,
    con il motivo scritto accanto invece di essere taciuto.
    """

    # Attributo di classe: gli handler XAML possono scattare durante il
    # popolamento iniziale, prima che __init__ abbia finito.
    _ready = False

    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)

        self.result = None
        self.view_checks = []

        self.header_hint.Text = (
            "Tagga le rooms visibili nelle viste selezionate. "
            "L'analisi non modifica il modello: i tag vengono creati solo "
            "dopo la conferma.")

        for label in source_labels:
            self.source_combo.Items.Add(label)
        self.source_combo.SelectedIndex = 0

        for label in tag_labels:
            self.tag_combo.Items.Add(label)
        self.tag_combo.SelectedIndex = 0

        self.band_input.Text = "{:.2f}".format(LEVEL_BAND_METERS)

        for view in plan_views:
            check = Controls.CheckBox()
            check.Content = "{}  [{}]".format(
                view.Name,
                "Ceiling" if view.ViewType == DB.ViewType.CeilingPlan else "Floor")
            check.Tag = view
            check.Margin = Thickness(2, 2, 2, 2)
            check.Checked += self.view_toggled
            check.Unchecked += self.view_toggled
            self.views_panel.Children.Add(check)
            self.view_checks.append(check)

        self._ready = True
        self.refresh_strategy()
        self.refresh_count()

    # ------------- helper -------------
    def selected_source_label(self):
        item = self.source_combo.SelectedItem
        return str(item) if item is not None else None

    def refresh_strategy(self):
        """Abilita la scelta della strategia solo quando L1 e' applicabile."""
        label = self.selected_source_label()
        if label is None:
            return

        is_host = (source_dict.get(label) is None)
        volumes, reason = source_volumes_ok(label)
        self.source_info.Text = reason

        if is_host:
            self.strategy_group.IsEnabled = False
            self.strategy_solid.IsChecked = False
            self.strategy_geo.IsChecked = False
            self.strategy_info.Text = (
                "Rooms del modello corrente: viene usata la visibilita di "
                "Revit (L0), che applica crop, view range, fase e filtri.")
            return

        if not volumes:
            self.strategy_group.IsEnabled = False
            self.strategy_solid.IsChecked = False
            self.strategy_geo.IsChecked = True
            self.strategy_info.Text = (
                "Filtro a solido non disponibile: senza volumi calcolati le "
                "rooms non hanno geometria 3D. Viene usato il modo geometrico.")
            return

        self.strategy_group.IsEnabled = True
        self.strategy_info.Text = (
            "Se su una vista la strategia scelta non e' applicabile, lo "
            "script ripiega automaticamente e lo dichiara nel report.")
        if not self.strategy_solid.IsChecked and not self.strategy_geo.IsChecked:
            self.strategy_solid.IsChecked = True

    def refresh_count(self):
        n = len([c for c in self.view_checks if c.IsChecked])
        self.views_count.Text = "{} di {} viste selezionate".format(
            n, len(self.view_checks))

    def parse_band(self):
        """Altezza fascia in metri. Accetta virgola o punto come separatore.
        Restituisce (valore, messaggio_errore)."""
        raw = (self.band_input.Text or "").strip().replace(",", ".")
        if not raw:
            return None, "Inserisci l'altezza della fascia verticale."
        try:
            value = float(raw)
        except ValueError:
            return None, "L'altezza della fascia non e' un numero valido."
        if value < 0:
            return None, "L'altezza della fascia non puo essere negativa."
        if value > 100.0:
            return None, "L'altezza della fascia sembra fuori scala (> 100 m)."
        return value, None

    # ------------- handler XAML -------------
    def source_changed(self, sender, args):
        if self._ready:
            self.refresh_strategy()

    def view_toggled(self, sender, args):
        if self._ready:
            self.refresh_count()

    def filter_changed(self, sender, args):
        if not self._ready:
            return
        needle = (self.views_filter.Text or "").strip().lower()
        for check in self.view_checks:
            visible = (not needle) or (needle in str(check.Content).lower())
            check.Visibility = Visibility.Visible if visible else Visibility.Collapsed

    def select_all_click(self, sender, args):
        for check in self.view_checks:
            check.IsChecked = True

    def select_none_click(self, sender, args):
        for check in self.view_checks:
            check.IsChecked = False

    def select_filtered_click(self, sender, args):
        """Seleziona solo le viste che passano il filtro corrente."""
        for check in self.view_checks:
            check.IsChecked = (check.Visibility == Visibility.Visible)

    def annulla_click(self, sender, args):
        self.result = None
        self.Close()

    def procedi_click(self, sender, args):
        views = [c.Tag for c in self.view_checks if c.IsChecked]
        if not views:
            forms.alert("Seleziona almeno una vista.", title="Nessuna vista")
            return

        band, band_error = self.parse_band()
        if band_error:
            forms.alert(band_error, title="Fascia verticale")
            return

        label = self.selected_source_label()
        volumes, _ = source_volumes_ok(label)
        is_host = (source_dict.get(label) is None)

        self.result = {
            "source_label": label,
            "views": views,
            "tag_label": str(self.tag_combo.SelectedItem),
            "band_meters": band,
            "dry_run": bool(self.dryrun_check.IsChecked),
            "verbose": bool(self.diag_check.IsChecked),
            "use_solid_filter": bool(self.strategy_solid.IsChecked),
            "volumes_ok": volumes,
            "source_is_host": is_host,
        }
        self.Close()


xaml_path = None
try:
    xaml_path = script.get_bundle_file(XAML_FILE_NAME)
except Exception:
    pass
if not xaml_path:
    xaml_path = op.join(op.dirname(__file__), XAML_FILE_NAME)

if not op.isfile(xaml_path):
    forms.alert("File grafica non trovato:\n\n{}\n\nDeve stare nella stessa "
                "cartella dello script.".format(xaml_path),
                title="Grafica mancante", exitscript=True)

setup_window = TagRoomsWindow(xaml_path)
setup_window.show_dialog()
setup = setup_window.result

if not setup:
    script.exit()

# --- Traduzione delle scelte nelle variabili usate dal resto dello script ---
selected_source = setup["source_label"]
views_selected = setup["views"]
selected_tag_name = setup["tag_label"]
tag_type = tag_dict[selected_tag_name]
dry_run = setup["dry_run"]
# verbose = diagnostica avanzata. Il report base e quello dettagliato sono
# alternativi: chi chiede il dettaglio non vuole anche il riassunto.
verbose = setup["verbose"]
volumes_ok = setup["volumes_ok"]
source_is_host = setup["source_is_host"]

selected_link = source_dict[selected_source]
if selected_link is None:
    room_doc = doc
    link_transform = DB.Transform.Identity
    link_id_key = -1
else:
    room_doc = selected_link.GetLinkDocument()
    link_transform = selected_link.GetTotalTransform()
    link_id_key = get_element_id_value(selected_link.Id)

# Altezza della fascia verticale, da input. 0 = nessuna fascia, si usa
# l'intero view range della vista.
LEVEL_BAND_METERS = setup["band_meters"]
LEVEL_BAND_FT = LEVEL_BAND_METERS / 0.3048
opt_band = LEVEL_BAND_FT > 1e-9

# Il filtro a solido richiede sorgente linkata, volumi calcolati e la scelta
# dell'utente. Il motivo di un'eventuale indisponibilita viene dichiarato.
l1_available = (not source_is_host) and volumes_ok
if l1_available:
    use_solid_filter = setup["use_solid_filter"]
    l1_choice_note = ("filtro a solido (L1), scelto dall'utente"
                      if use_solid_filter
                      else "modo geometrico (L2), scelto dall'utente")
else:
    use_solid_filter = False
    if source_is_host:
        l1_choice_note = "non applicabile: sorgente = modello corrente"
    elif not volumes_ok:
        l1_choice_note = "L2 obbligato: volumi non calcolati nel modello linkato"
        warn("Nel modello '{}' i volumi delle rooms NON sono calcolati "
             "(Area and Volume Computations su 'Areas only'). Il filtro a "
             "solido non e' utilizzabile: uso il modo geometrico."
             .format(selected_source))
    else:
        l1_choice_note = "L2 obbligato"


# =========================================================================
# FASE 1 - PLAN: identificazione, senza modificare il modello
# =========================================================================
all_rooms = FilteredElementCollector(room_doc)\
    .OfCategory(BuiltInCategory.OST_Rooms)\
    .WhereElementIsNotElementType()\
    .ToElements()

placed_rooms = []
for room in all_rooms:
    try:
        if room.Area > 0 and room.Location is not None:
            placed_rooms.append(room)
    except Exception:
        continue

if not placed_rooms:
    forms.alert("Nessuna room posizionata trovata in '{}'.".format(selected_source),
                exitscript=True)

host_levels = get_sorted_levels(doc)

# Le quote dei livelli sono confrontabili con la geometria? Se il parametro
# 'Elevation Base' dei livelli e' riferito al Survey Point, Level.Elevation
# restituisce l'altitudine sul livello del mare mentre la geometria resta in
# coordinate interne: il test verticale confronterebbe due scale diverse e
# scarterebbe tutte le rooms. Il controllo lo intercetta invece di lasciare
# che si manifesti come 'zero tag creati'.
elev_checked, elev_worst = check_level_geometry_consistency(placed_rooms)
if elev_checked and elev_worst > 1.0:
    warn("Le quote dei livelli di '{}' non sono coerenti con la geometria: "
         "scarto fino a {:.2f} piedi ({:.2f} m) su {} rooms controllate. "
         "Verifica il parametro 'Elevation Base' dei livelli. Il test "
         "verticale su questo modello non e' attendibile."
         .format(selected_source, elev_worst, elev_worst * 0.3048, elev_checked))

# Caratterizzazione delle viste: una volta per vista, non per room.
view_info = {}
for view in views_selected:
    vid = get_element_id_value(view.Id)
    xy_polygon, xy_desc = build_view_xy_polygon(view)
    view_info[vid] = {
        "view": view,
        "xy_polygon": xy_polygon,
        "xy_desc": xy_desc,
        "view_range": get_view_range_info(view, host_levels),
    }

ctx = {
    "room_doc": room_doc,
    "link_transform": link_transform,
    "link_id_key": link_id_key,
    "source_is_host": source_is_host,
    "volumes_ok": volumes_ok,
    "placed_rooms": placed_rooms,
    "view_info": view_info,
    "probe": {},
    "opt_band": opt_band,
    "use_solid_filter": use_solid_filter,
}

if verbose:
    print("=" * 80)
    print("**ANALISI** (nessuna modifica al modello in questa fase)")
    print("=" * 80)
    print("SORGENTE ROOMS : {}".format(selected_source))
    print("ROOMS          : {} totali, {} posizionate".format(
        len(all_rooms), len(placed_rooms)))
    print("VOLUMI ROOMS   : {}".format("calcolati" if volumes_ok else "NON calcolati"))
    print("TAG            : {}".format(selected_tag_name))
    print("VISTE          : {}".format(len(views_selected)))
    print("FISSI          : crop/scope box, salta duplicati, senza leader, "
          "riposiziona nella porzione visibile")
    print("FASCIA         : {}".format(
        "primi {:.2f} m sopra il livello della vista ({:.2f} piedi)".format(
            LEVEL_BAND_METERS, LEVEL_BAND_FT)
        if opt_band else "disattivata (0 m), uso l'intero view range"))
    print("MODO           : {}".format(
        "SOLO ANTEPRIMA, nessun tag verra creato" if dry_run
        else "analisi + creazione dei tag dopo conferma"))

if verbose and not source_is_host:
    if l1_available:
        print("STRATEGIA LINK : scelta dall'utente -> {}".format(l1_choice_note))
    else:
        print("STRATEGIA LINK : {}".format(l1_choice_note or "modo geometrico"))
    try:
        org = link_transform.Origin
        print("TRASF. LINK    : origine ({:.2f}, {:.2f}, {:.2f}) | "
              "traslazione Z = {:.2f} piedi".format(org.X, org.Y, org.Z, org.Z))
    except Exception:
        pass

# Se lo stesso documento e' linkato piu volte, la Transform usata potrebbe
# essere quella dell'istanza sbagliata: e' una causa possibile di uno
# scostamento sistematico di quota. Il controllo gira sempre, perche produce
# un avviso; il dettaglio delle istanze solo in diagnostica avanzata.
if not source_is_host:
    try:
        siblings = []
        for lk in loaded_links:
            d = lk.GetLinkDocument()
            if d is not None and d.PathName == room_doc.PathName:
                z = 0.0
                try:
                    z = lk.GetTotalTransform().Origin.Z
                except Exception:
                    pass
                siblings.append((get_element_id_value(lk.Id), z))
        if verbose:
            print("ISTANZA USATA  : id {} | istanze dello stesso documento: {}".format(
                get_element_id_value(selected_link.Id), len(siblings)))
        if len(siblings) > 1:
            if verbose:
                for sid, sz in siblings:
                    print("                 id {} -> Z = {:.2f}".format(sid, sz))
            warn("Il documento '{}' e' linkato {} volte. Se la vista mostra una "
                 "istanza diversa da quella selezionata, le quote calcolate "
                 "saranno sistematicamente sbagliate: seleziona l'istanza con "
                 "l'id corretto.".format(selected_source, len(siblings)))
    except Exception:
        pass

if verbose:
    print("=" * 80)

plan = []            # elementi da creare
view_reports = []    # una riga per vista
views_skipped = []

# La transazione di analisi serve SOLO alla strategia L0, per riattivare la
# categoria Rooms nelle viste e rendere interrogabile il collector view-scoped.
# Viene sempre annullata: nessuna modifica sopravvive e non entra nell'undo
# stack. Se la sorgente e' un link, l'analisi non apre nessuna transazione:
# e' lettura pura.
probe_t = Transaction(doc, "Analisi visibilita rooms (annullata)") \
    if source_is_host else None

try:
    if probe_t is not None:
        probe_t.Start()
        for view in views_selected:
            ctx["probe"][get_element_id_value(view.Id)] = \
                prepare_view_for_probe(view)
        doc.Regenerate()

    with forms.ProgressBar(title='Analisi viste... ({value} di {max_value})',
                           cancellable=True) as pb:
        for idx, view in enumerate(views_selected):
            if pb.cancelled:
                print("\nAnalisi annullata dall'utente.")
                script.exit()
            pb.update_progress(idx, len(views_selected))

            vid = get_element_id_value(view.Id)
            info = view_info[vid]

            if info["view_range"] is None:
                views_skipped.append(view.Name)
                view_reports.append({
                    "name": view.Name, "view_id": vid, "strategy": "-",
                    "found": 0, "planned": 0, "dup": 0, "excluded": 0,
                    "no_point": 0, "outside": 0,
                    "note": "view range non leggibile, vista saltata"})
                continue

            rooms, strategy, note, stats = resolve_visible_rooms(view, ctx)

            # Diagnostica per vista, solo in modo avanzato. Serve a distinguere
            # 'la room non e' davvero in questa vista' da 'il test verticale
            # sbaglia': senza questi numeri l'esclusione non e' interpretabile.
            if verbose:
                print("")
                print("**VISTA: {}** (strategia: {})".format(view.Name, strategy))
                print("  View range -> {}".format(format_range(info["view_range"])))
                gen_lvl = info["view_range"].get("gen_level")
                if gen_lvl is not None:
                    try:
                        print("  Livello vista -> '{}': Elevation={:.2f} | "
                              "ProjectElevation={:.2f} | uso {:.2f} (interna)".format(
                                  gen_lvl.Name, gen_lvl.Elevation,
                                  gen_lvl.ProjectElevation,
                                  level_internal_elevation(gen_lvl)))
                    except Exception:
                        pass
                if opt_band:
                    b_low, b_top = get_vertical_limits(info["view_range"], True)
                    print("  Fascia     -> da {:.2f} a {:.2f} (primi {:.2f} m "
                          "sopra il livello, intersecati col view range)".format(
                              b_low, b_top, LEVEL_BAND_METERS))
                print("  Limite XY  -> {}".format(info["xy_desc"]))
                if stats and stats.get("z_seen_min") is not None:
                    print("  Rooms      -> estensione verticale da {:.2f} a {:.2f} "
                          "(coordinate host)".format(
                              stats["z_seen_min"], stats["z_seen_max"]))

                def _fmt(v):
                    return "n/d" if v is None else "{:.2f}".format(v)

                for s in (stats.get("samples") if stats else None) or []:
                    print("    [escluso] room {}: z da {:.2f} a {:.2f} ({}) -> {}".format(
                        s["id"], s["z_min"], s["z_max"], s.get("z_source"), s["why"]))
                    # Confronto fra le fonti: e' cio che permette di capire se
                    # sbaglia la ricostruzione, il bounding box o la Transform.
                    print("        bbox host {} .. {} | parametri host {} .. {} | "
                          "livello link {} | punto ins. link {} -> host {}".format(
                              _fmt(s["z_min"]) if s.get("z_source") == "bounding box" else "n/d",
                              _fmt(s["z_max"]) if s.get("z_source") == "bounding box" else "n/d",
                              _fmt(s.get("z_min_param")), _fmt(s.get("z_max_param")),
                              _fmt(s.get("base_elev_link")),
                              _fmt(s.get("loc_z_link")), _fmt(s.get("loc_z_host"))))

            existing_keys = collect_existing_tag_keys(view)
            xy_polygon = info["xy_polygon"]

            planned = 0
            dup = 0
            no_point = 0
            outside = 0

            for room in rooms:
                key = (link_id_key, get_element_id_value(room.Id))
                if key in existing_keys:
                    dup += 1
                    continue

                geo = get_room_geo(room, ctx)
                point, in_crop = pick_insertion_point(geo, xy_polygon, True)
                if point is None:
                    no_point += 1
                    continue
                if not in_crop:
                    outside += 1

                plan.append({
                    "view": view,
                    "view_name": view.Name,
                    "room": room,
                    "room_id": get_element_id_value(room.Id),
                    "point": point,
                    "in_crop": in_crop,
                })
                existing_keys.add(key)
                planned += 1

            excluded = len(placed_rooms) - len(rooms)
            note_parts = []
            if note:
                note_parts.append(note)
            if stats:
                detail = []
                if stats.get("out_z"):
                    detail.append("fuori view range {}".format(stats["out_z"]))
                if stats.get("out_band"):
                    detail.append("fuori fascia {}".format(stats["out_band"]))
                if stats.get("out_xy"):
                    detail.append("fuori crop {}".format(stats["out_xy"]))
                if stats.get("out_phase"):
                    detail.append("fase diversa {}".format(stats["out_phase"]))
                if stats.get("no_geo"):
                    detail.append("senza geometria {}".format(stats["no_geo"]))
                if detail:
                    note_parts.append(", ".join(detail))
            if outside:
                note_parts.append("{} tag fuori crop (non visibili)".format(outside))

            view_reports.append({
                "name": view.Name,
                "view_id": vid,
                "strategy": strategy,
                "found": len(rooms),
                "planned": planned,
                "dup": dup,
                "excluded": excluded,
                "no_point": no_point,
                "outside": outside,
                "note": " | ".join(note_parts),
            })
finally:
    if probe_t is not None and probe_t.HasStarted() and not probe_t.HasEnded():
        probe_t.RollBack()


# =========================================================================
# FASE 2 - ANTEPRIMA
# =========================================================================
total_planned = len(plan)
total_found = sum(r["found"] for r in view_reports)
total_dup = sum(r["dup"] for r in view_reports)
total_no_point = sum(r["no_point"] for r in view_reports)
total_outside = sum(1 for p in plan if not p["in_crop"])

# La quota di base a 0.0 e' un avviso, va emesso in ogni modo: il conteggio
# per fonte invece e' diagnostica.
if "fallback 0.0" in BASE_ELEV_METHODS:
    warn("Per {} rooms non e' stato possibile risolvere ne il livello ne il "
         "punto di inserimento: la quota di base e' stata assunta a 0.0, "
         "quindi il test verticale su quelle rooms non e' attendibile."
         .format(BASE_ELEV_METHODS["fallback 0.0"]))


def print_warnings():
    """Gli avvisi si vedono in entrambi i report: sono la sola cosa che
    distingue 'nessun tag da creare perche non serviva' da 'nessun tag da
    creare perche qualcosa non ha funzionato'."""
    if WARNINGS:
        print("")
        print("### Avvisi")
        print("")
        for w in WARNINGS:
            print("- {}".format(w))


def print_basic_report(created=None, errors=None):
    """Report base: le informazioni necessarie a capire cosa e' stato fatto.
    Non viene stampato quando e' attiva la diagnostica avanzata, che le
    contiene tutte in forma piu estesa."""
    print("=" * 80)
    print("**LINKED ROOM TAG - REPORT**")
    print("=" * 80)
    print("Modello rooms      : {}".format(selected_source))
    print("Tipo di tag        : {}".format(selected_tag_name))
    print("Strategia          : {}".format(", ".join(
        sorted(set(r["strategy"] for r in view_reports if r["strategy"] != "-")))
        or "nessuna"))
    print("Fascia verticale   : {}".format(
        "primi {:.2f} m sopra il livello della vista".format(LEVEL_BAND_METERS)
        if opt_band else "disattivata, intero view range"))
    print("")
    print("Viste selezionate  : {}".format(len(views_selected)))
    print("")
    if created is None:
        print("| Vista | Rooms individuate | Da taggare | Gia taggate |")
        print("|---|---:|---:|---:|")
        for r in view_reports:
            print("| {} | {} | {} | {} |".format(
                r["name"], r["found"], r["planned"], r["dup"]))
    else:
        print("| Vista | Rooms individuate | Tag creati | Gia taggate | Errori |")
        print("|---|---:|---:|---:|---:|")
        for r in view_reports:
            print("| {} | {} | {} | {} | {} |".format(
                r["name"], r["found"],
                created_by_view.get(r["view_id"], 0),
                r["dup"],
                errors_by_view.get(r["view_id"], 0)))
    print("")
    print("Rooms nel modello  : {} ({} posizionate)".format(
        len(all_rooms), len(placed_rooms)))
    print("Rooms individuate  : {}".format(total_found))
    if created is None:
        print("Rooms da taggare   : {}".format(total_planned))
    else:
        print("Rooms taggate      : {} su {} pianificate".format(
            created, total_planned))
    print("Rooms gia taggate  : {}".format(total_dup))
    if total_no_point:
        print("Rooms scartate     : {} senza punto di inserimento valido".format(
            total_no_point))
    if total_outside:
        print("Tag fuori crop     : {} (creati ma non visibili nella vista)".format(
            total_outside))
    if errors is not None:
        print("Errori di creazione: {}".format(errors))
    if views_skipped:
        print("Viste saltate      : {} ({})".format(
            len(views_skipped), ", ".join(views_skipped)))
    print("=" * 80)
    print_warnings()


if verbose:
    print("")
    print("### Riepilogo per vista")
    print("")
    print("| Vista | Strategia | Individuate | Da creare | Gia taggate | Escluse | Note |")
    print("|---|---|---:|---:|---:|---:|---|")
    for r in view_reports:
        print("| {} | {} | {} | {} | {} | {} | {} |".format(
            r["name"], r["strategy"], r["found"], r["planned"], r["dup"],
            r["excluded"], r["note"] or ""))

    print("")
    print("**Tag da creare: {}**".format(total_planned))
    if total_dup:
        print("Rooms gia taggate e saltate: {}".format(total_dup))
    if total_no_point:
        print("Rooms scartate per punto di inserimento non determinabile: {}".format(
            total_no_point))
    if total_outside:
        print("Tag il cui punto cade fuori dalla crop (verranno creati ma non "
              "visibili nella vista): {}".format(total_outside))

    if plan:
        print("")
        print("### Dettaglio")
        print("")
        shown = plan[:PREVIEW_ROW_LIMIT]
        print("| Vista | Room Id | Nome room | X | Y | Visibile |")
        print("|---|---|---|---:|---:|---|")
        for p in shown:
            try:
                room_name = p["room"].get_Parameter(
                    BuiltInParameter.ROOM_NAME).AsString() or ""
            except Exception:
                room_name = ""
            # Le rooms linkate non sono selezionabili dall'host: nessun linkify.
            id_cell = str(p["room_id"])
            if source_is_host:
                try:
                    id_cell = output.linkify(p["room"].Id)
                except Exception:
                    pass
            print("| {} | {} | {} | {:.2f} | {:.2f} | {} |".format(
                p["view_name"], id_cell, room_name,
                p["point"][0], p["point"][1],
                "si" if p["in_crop"] else "NO"))
        if len(plan) > PREVIEW_ROW_LIMIT:
            print("")
            print("_Mostrate {} righe su {}: {} righe non elencate._".format(
                len(shown), len(plan), len(plan) - len(shown)))

    if BASE_ELEV_METHODS or Z_SOURCE_COUNTS:
        print("")
        print("### Diagnostica estensione verticale delle rooms")
        print("")
        for method in sorted(BASE_ELEV_METHODS.keys()):
            print("- quota di base risolta con {}: {} rooms".format(
                method, BASE_ELEV_METHODS[method]))
        for src in sorted(Z_SOURCE_COUNTS.keys()):
            print("- estensione verticale da {}: {} rooms".format(
                src, Z_SOURCE_COUNTS[src]))

    print_warnings()

    if views_skipped:
        print("")
        print("Viste saltate: {}".format(", ".join(views_skipped)))

if not plan:
    # Anche senza niente da fare il report va emesso: un avviso che rimanda a
    # un pannello vuoto e' proprio il caso da evitare.
    if not verbose:
        print_basic_report()
    forms.alert("L'analisi non ha prodotto nessun tag da creare.\n\n"
                "Controlla il report e gli avvisi nel pannello di output.",
                title="Niente da fare", exitscript=True)

if dry_run:
    # Nessuna creazione: il report base va emesso qui, con i dati dell'analisi.
    if not verbose:
        print_basic_report()
    forms.alert("Modo 'solo anteprima' attivo.\n\n"
                "{} tag sarebbero stati creati. Nessuna modifica applicata.\n"
                "Il report e' nel pannello di output.".format(total_planned),
                title="Anteprima - nessuna modifica")
    script.exit()

confirm_msg = "Creo {} tag su {} viste?".format(
    total_planned, len(set(p["view_name"] for p in plan)))
if total_outside:
    confirm_msg += "\n\nAttenzione: {} tag cadranno fuori dalla crop region " \
                   "e non saranno visibili.".format(total_outside)
if WARNINGS:
    confirm_msg += "\n\n{} avvisi da leggere nel pannello di output.".format(
        len(WARNINGS))
if verbose:
    confirm_msg += "\n\nIl dettaglio completo e' nel pannello di output."
else:
    confirm_msg += "\n\nIl report verra scritto nel pannello di output."

if not forms.alert(confirm_msg, title="Conferma creazione", yes=True, no=True):
    print("")
    print("Creazione annullata dall'utente: nessuna modifica applicata.")
    script.exit()


# =========================================================================
# FASE 3 - APPLY: transazione corta, solo scritture
# =========================================================================
created = 0
errors = 0
created_by_view = {}
errors_by_view = {}

# Raggruppiamo per vista: il piano e' costruito vista per vista, quindi le
# righe consecutive appartengono alla stessa vista.
plan_by_view = []
current_view = None
current_items = None
for p in plan:
    if current_view is None or p["view"].Id != current_view.Id:
        current_view = p["view"]
        current_items = []
        plan_by_view.append((current_view, current_items))
    current_items.append(p)

t = Transaction(doc, "Tag Rooms in Multiple Views")
t.Start()
try:
    if not tag_type.IsActive:
        tag_type.Activate()
        doc.Regenerate()

    with forms.ProgressBar(title='Creazione tag... ({value} di {max_value})',
                           cancellable=True) as pb:
        step = 0
        for view, items in plan_by_view:
            view_key = get_element_id_value(view.Id)
            for p in items:
                if pb.cancelled:
                    t.RollBack()
                    print("")
                    print("Operazione annullata dall'utente: nessuna modifica applicata.")
                    script.exit()
                pb.update_progress(step, total_planned)
                step += 1

                room = p["room"]
                try:
                    if selected_link is None:
                        link_elem_id = LinkElementId(room.Id)
                    else:
                        link_elem_id = LinkElementId(selected_link.Id, room.Id)

                    uv_point = UV(p["point"][0], p["point"][1])
                    new_tag = doc.Create.NewRoomTag(link_elem_id, uv_point, view.Id)

                    if new_tag is None:
                        errors += 1
                        errors_by_view[view_key] = errors_by_view.get(view_key, 0) + 1
                        continue

                    if new_tag.GetTypeId() != tag_type.Id:
                        new_tag.ChangeTypeId(tag_type.Id)

                    try:
                        new_tag.HasLeader = False
                    except Exception:
                        pass

                    created += 1
                    created_by_view[view_key] = created_by_view.get(view_key, 0) + 1

                except Exception as err:
                    errors += 1
                    errors_by_view[view_key] = errors_by_view.get(view_key, 0) + 1
                    # Gli errori di creazione si stampano sempre: sono la
                    # ragione per cui il conteggio finale non torna.
                    print("  [!] Vista '{}', room {}: {}".format(
                        view.Name, p["room_id"], err))

    t.Commit()

    print("")
    if verbose:
        print("=" * 80)
        print("RIEPILOGO FINALE")
        print("=" * 80)
        print("Tag pianificati       : {}".format(total_planned))
        print("Tag creati            : {}".format(created))
        print("Errori in creazione   : {}".format(errors))
        print("Rooms gia taggate     : {}".format(total_dup))
        print("Viste elaborate       : {}".format(len(plan_by_view)))
        if views_skipped:
            print("Viste saltate         : {}".format(", ".join(views_skipped)))
    else:
        print_basic_report(created=created, errors=errors)

    forms.alert(
        "Operazione completata.\n\n"
        "Tag creati: {} su {} pianificati\n"
        "Rooms gia taggate: {}\n"
        "Errori: {}\n"
        "Viste elaborate: {}".format(
            created, total_planned, total_dup, errors, len(plan_by_view)),
        title="Linked Room Tag - Completato")

except Exception as err:
    if t.HasStarted() and not t.HasEnded():
        t.RollBack()
    print("")
    print("ERRORE CRITICO: {}".format(err))
    forms.alert("Errore durante la creazione, nessuna modifica applicata:\n\n{}".format(err),
                title="Errore", exitscript=True)
