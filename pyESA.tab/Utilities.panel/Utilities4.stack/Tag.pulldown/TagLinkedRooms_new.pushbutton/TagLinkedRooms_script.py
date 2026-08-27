# -*- coding: utf-8 -*-
# Intestazione dello script con metadati
__title__   = "Linked Room Tag\nin Multiple Views"
__doc__     = """Version = 3.0
Date    = 27.08.2026
________________________________________________________________
Tagga automaticamente le rooms visibili nelle viste selezionate,
con il tag scelto, prendendo le rooms dal modello corrente oppure
da un modello linkato selezionato dall'utente.

NOVITA' v3.0:
- Gestione corretta delle viste RUOTATE da scope box: il test di
  appartenenza alla crop region viene fatto sulla crop shape reale
  (o con la Transform inversa del crop box), non sul bounding box
  in coordinate mondo.
- Tag allineati alla rotazione della vista (opzionale).
- View range risolto correttamente (Unlimited / Level Above /
  Level Below / Current) e confronto con l'estensione verticale
  reale della room (Base Offset / Upper Limit / Limit Offset).
- Ricerca tag esistenti in O(1) per room (no scansione ripetuta).
________________________________________________________________
Author(s): Tommaso Lorenzi, Andrea Patti
"""

# -------------------------------
# SEZIONE IMPORT MODULI
# -------------------------------
import math

from pyrevit import revit, script, DB, forms
from Autodesk.Revit.DB import (
    Transaction,
    RevitLinkInstance,
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    LinkElementId,
    ElementId,
    ElementTransformUtils,
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

RVT_VER = int(app.VersionNumber)

XY_TOL = 1.0 / 304.8 * 10.0   # ~10 mm di tolleranza in piedi sul test XY
Z_TOL  = 1.0 / 304.8 * 10.0   # ~10 mm di tolleranza in piedi sul test Z
DEFAULT_ROOM_HEIGHT = 8.0     # piedi, fallback se l'estensione verticale non e' calcolabile

HOST_LABEL = "<< Modello corrente >>"

# Etichette opzioni (selezionate = attive)
OPT_CROP    = "Considera solo le rooms dentro la crop region / scope box"
OPT_ROTATE  = "Allinea i tag alla rotazione della vista (scope box)"
OPT_SKIPDUP = "Salta le rooms che hanno gia un tag nella vista"
OPT_NOLEAD  = "Crea i tag senza leader"
OPT_CUT     = "Richiedi che il cut plane attraversi la room (piu restrittivo)"
DEFAULT_OPTIONS = [OPT_CROP, OPT_ROTATE, OPT_SKIPDUP, OPT_NOLEAD]


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
    return element_id.IntegerValue


def get_param_double(element, bip, default=0.0):
    """Legge un parametro double, restituendo default se assente/vuoto."""
    try:
        p = element.get_Parameter(bip)
        if p and p.HasValue:
            return p.AsDouble()
    except Exception:
        pass
    return default


# -------------------------------
# SEZIONE HELPER - LIVELLI E VIEW RANGE
# -------------------------------
def get_sorted_levels(document):
    """Livelli del documento ordinati per elevazione."""
    levels = FilteredElementCollector(document).OfClass(DB.Level)\
                                               .WhereElementIsNotElementType()\
                                               .ToElements()
    return sorted(levels, key=lambda l: l.Elevation)


def _special_plan_level_ids():
    """ElementId speciali usati da PlanViewRange (Unlimited, Current, ...).
    Restituisce un dict nome -> ElementId (vuoto se l'API non li espone)."""
    special = {}
    for name in ("Unlimited", "Current", "LevelAbove", "LevelBelow"):
        try:
            special[name] = getattr(DB.PlanViewRange, name)
        except Exception:
            pass
    return special


SPECIAL_PLAN_IDS = _special_plan_level_ids()


def _resolve_plane_elevation(view, view_range, plane, host_levels):
    """Elevazione assoluta (piedi) di un piano del view range.
    Restituisce None se il piano e' 'Unlimited'."""
    level_id = view_range.GetLevelId(plane)
    offset = view_range.GetOffset(plane)

    gen_level = view.GenLevel
    base_elev = gen_level.Elevation if gen_level else 0.0

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
            if lvl.Elevation > base_elev + 1e-6:
                nxt = lvl.Elevation
                break
        return (nxt if nxt is not None else base_elev) + offset

    below = SPECIAL_PLAN_IDS.get("LevelBelow")
    if below is not None and level_id == below:
        prev = None
        for lvl in host_levels:
            if lvl.Elevation < base_elev - 1e-6:
                prev = lvl.Elevation
        return (prev if prev is not None else base_elev) + offset

    lvl = doc.GetElement(level_id)
    return ((lvl.Elevation if lvl else base_elev) + offset)


def get_view_range_info(view, host_levels):
    """Estensione verticale della vista in coordinate host.
    top/bottom possono essere None = illimitato."""
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
        print("  [!] View range non leggibile: {}".format(err))
        return None

    # Il limite inferiore utile e' il piu basso tra bottom e view depth
    if bottom is None or depth is None:
        low = None
    else:
        low = min(bottom, depth)

    return {"top": top, "cut": cut, "bottom": bottom, "depth": depth, "low": low}


def format_range(view_range_info):
    def fmt(v):
        return "illimitato" if v is None else "{:.2f}".format(v)
    return "Top: {} | Cut: {} | Bottom: {} | Depth: {}".format(
        fmt(view_range_info["top"]), fmt(view_range_info["cut"]),
        fmt(view_range_info["bottom"]), fmt(view_range_info["depth"]))


# -------------------------------
# SEZIONE HELPER - GEOMETRIA ROOM
# -------------------------------
def get_room_point(room):
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


def get_room_z_span(room, room_doc):
    """Estensione verticale (z_min, z_max) della room nel suo documento,
    letta da Level / Base Offset / Upper Limit / Limit Offset."""
    base_level = room_doc.GetElement(room.LevelId) if room.LevelId else None
    base_elev = base_level.Elevation if base_level else 0.0

    lower_offset = get_param_double(room, BuiltInParameter.ROOM_LOWER_OFFSET, 0.0)
    z_min = base_elev + lower_offset

    upper_offset = get_param_double(room, BuiltInParameter.ROOM_UPPER_OFFSET, None)
    if upper_offset is None:
        upper_offset = 0.0

    upper_level_elev = None
    try:
        p_up = room.get_Parameter(BuiltInParameter.ROOM_UPPER_LEVEL)
        if p_up:
            up_id = p_up.AsElementId()
            if up_id and up_id != ElementId.InvalidElementId:
                up_lvl = room_doc.GetElement(up_id)
                if up_lvl is not None and hasattr(up_lvl, "Elevation"):
                    upper_level_elev = up_lvl.Elevation
    except Exception:
        pass

    if upper_level_elev is None:
        upper_level_elev = base_elev

    z_max = upper_level_elev + upper_offset

    if z_max - z_min < 1e-6:
        # Room senza altezza utile: fallback sull'altezza reale o su un default
        height = get_param_double(room, BuiltInParameter.ROOM_HEIGHT, 0.0)
        if height <= 1e-6:
            height = DEFAULT_ROOM_HEIGHT
        z_max = z_min + height

    return z_min, z_max


# -------------------------------
# SEZIONE HELPER - CROP / SCOPE BOX (VISTE RUOTATE)
# -------------------------------
def _point_in_polygon(x, y, polygon):
    """Ray casting. polygon = lista di tuple (x, y) chiusa implicitamente."""
    inside = False
    n = len(polygon)
    if n < 3:
        return True
    j = n - 1
    for i in range(n):
        xi, yi = polygon[i]
        xj, yj = polygon[j]
        if ((yi > y) != (yj > y)):
            denom = (yj - yi)
            if abs(denom) > 1e-12:
                x_int = xi + (y - yi) * (xj - xi) / denom
                if x < x_int:
                    inside = not inside
        j = i
    return inside


def _polygon_from_crop_shape(view):
    """Poligono XY (coordinate modello host) della crop region reale.
    Gestisce automaticamente crop ruotate e non rettangolari."""
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
            pts = curve.Tessellate()
            for p in pts:
                if polygon:
                    lx, ly = polygon[-1]
                    if abs(lx - p.X) < 1e-9 and abs(ly - p.Y) < 1e-9:
                        continue
                polygon.append((p.X, p.Y))
    except Exception:
        return None

    if len(polygon) < 3:
        return None
    return polygon


def _box_test_factory(bbox):
    """Test XY su un BoundingBoxXYZ tenendo conto della sua Transform
    (indispensabile per crop box / scope box RUOTATI)."""
    transform = bbox.Transform
    inverse = transform.Inverse
    min_pt = bbox.Min
    max_pt = bbox.Max

    def test(point_host):
        local = inverse.OfPoint(point_host)
        if local.X < min_pt.X - XY_TOL or local.X > max_pt.X + XY_TOL:
            return False
        if local.Y < min_pt.Y - XY_TOL or local.Y > max_pt.Y + XY_TOL:
            return False
        return True

    return test


def get_scope_box(view):
    """Elemento scope box assegnato alla vista, o None."""
    try:
        p = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
        if p:
            sb_id = p.AsElementId()
            if sb_id and sb_id != ElementId.InvalidElementId:
                return doc.GetElement(sb_id)
    except Exception:
        pass
    return None


def build_xy_test(view):
    """Costruisce il predicato di appartenenza XY della vista.
    Restituisce (funzione(point_host) -> bool, descrizione)."""
    # 1) Crop shape reale: gia' in coordinate modello, quindi la rotazione
    #    da scope box e' gestita implicitamente dal test punto-in-poligono.
    if view.CropBoxActive:
        polygon = _polygon_from_crop_shape(view)
        if polygon:
            def test_poly(point_host, _poly=polygon):
                return _point_in_polygon(point_host.X, point_host.Y, _poly)
            return test_poly, "crop shape ({} vertici)".format(len(polygon))

        # 2) Fallback: crop box con la sua Transform (mai Min/Max in coord. mondo)
        try:
            crop_box = view.CropBox
            if crop_box:
                return _box_test_factory(crop_box), "crop box (con transform)"
        except Exception:
            pass

    # 3) Scope box assegnato ma crop non attiva
    scope_box = get_scope_box(view)
    if scope_box is not None:
        try:
            bbox = scope_box.get_BoundingBox(None)
            if bbox:
                return _box_test_factory(bbox), "scope box '{}'".format(scope_box.Name)
        except Exception:
            pass

    return (lambda point_host: True), "nessun limite XY"


def get_view_rotation(view):
    """Angolo (radianti) tra l'asse X del modello e la direzione 'destra'
    della vista, misurato attorno alla direzione di vista.
    Diverso da 0 quando la pianta e' ruotata da uno scope box."""
    try:
        angle = XYZ.BasisX.AngleOnPlaneTo(view.RightDirection, view.ViewDirection)
    except Exception:
        return 0.0
    # normalizza in -pi..pi per applicare la rotazione minima
    while angle > math.pi:
        angle -= 2.0 * math.pi
    while angle < -math.pi:
        angle += 2.0 * math.pi
    return angle


# -------------------------------
# SEZIONE HELPER - TAG ESISTENTI
# -------------------------------
def collect_existing_tag_keys(view):
    """Insieme delle chiavi (link_id, room_id) delle rooms gia' taggate
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
# SEZIONE SELEZIONE MODELLO CONTENENTE LE ROOMS
# -------------------------------
link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
loaded_links = [lk for lk in link_instances if lk.GetLinkDocument() is not None]

source_dict = {HOST_LABEL: None}
for link in loaded_links:
    label = link.Name.split(" : ")[0]
    # piu istanze dello stesso link: distinguiamo con l'ID istanza
    if label in source_dict:
        label = "{} [id {}]".format(label, get_element_id_value(link.Id))
    source_dict[label] = link

source_labels = [HOST_LABEL] + sorted([k for k in source_dict.keys() if k != HOST_LABEL])

selected_source = forms.SelectFromList.show(
    source_labels,
    multiselect=False,
    button_name='Seleziona Modello',
    title='Modello che contiene le Rooms'
)

if not selected_source:
    script.exit()

selected_link = source_dict[selected_source]
if selected_link is None:
    room_doc = doc
    link_transform = DB.Transform.Identity
    link_id_key = -1
else:
    room_doc = selected_link.GetLinkDocument()
    link_transform = selected_link.GetTotalTransform()
    link_id_key = get_element_id_value(selected_link.Id)

# -------------------------------
# SEZIONE SELEZIONE VISTE
# -------------------------------
plan_types = (DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan)
all_views = FilteredElementCollector(doc).OfClass(DB.ViewPlan).ToElements()
plan_views = [v for v in all_views if not v.IsTemplate and v.ViewType in plan_types]
plan_views.sort(key=lambda v: v.Name)

if not plan_views:
    forms.alert("Non ci sono piante (Floor Plan / Ceiling Plan) nel progetto.", exitscript=True)

views_selected = forms.SelectFromList.show(
    plan_views,
    multiselect=True,
    name_attr='Name',
    button_name='Seleziona Viste',
    title='Viste da taggare'
)

if not views_selected:
    script.exit()

# -------------------------------
# SEZIONE SELEZIONE ROOM TAG
# -------------------------------
room_tag_types = list(FilteredElementCollector(doc)
                      .OfClass(DB.FamilySymbol)
                      .OfCategory(BuiltInCategory.OST_RoomTags))

if not room_tag_types:
    forms.alert("Non ci sono Room Tags caricati nel progetto.", exitscript=True)

tag_dict = {}
for tag_type in room_tag_types:
    try:
        type_name = tag_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    except Exception:
        type_name = "?"
    full_name = "{} - {}".format(tag_type.FamilyName, type_name)
    if full_name in tag_dict:
        full_name = "{} [id {}]".format(full_name, get_element_id_value(tag_type.Id))
    tag_dict[full_name] = tag_type

selected_tag_name = forms.SelectFromList.show(
    sorted(tag_dict.keys()),
    multiselect=False,
    button_name='Seleziona Room Tag',
    title='Tipo di Room Tag'
)

if not selected_tag_name:
    script.exit()

tag_type = tag_dict[selected_tag_name]

# -------------------------------
# SEZIONE OPZIONI
# -------------------------------
selected_options = forms.SelectFromList.show(
    [OPT_CROP, OPT_ROTATE, OPT_SKIPDUP, OPT_NOLEAD, OPT_CUT],
    multiselect=True,
    button_name='Conferma Opzioni',
    title='Opzioni (le voci selezionate sono attive)'
)

if selected_options is None:
    selected_options = DEFAULT_OPTIONS

opt_crop = OPT_CROP in selected_options
opt_rotate = OPT_ROTATE in selected_options
opt_skip_dup = OPT_SKIPDUP in selected_options
opt_no_leader = OPT_NOLEAD in selected_options
opt_cut_plane = OPT_CUT in selected_options

# -------------------------------
# SEZIONE RACCOLTA ROOMS
# -------------------------------
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

# Pre-calcolo geometrie delle rooms in coordinate host (una volta sola)
room_data = []
for room in placed_rooms:
    point = get_room_point(room)
    if point is None:
        continue
    z_min, z_max = get_room_z_span(room, room_doc)
    p_host = link_transform.OfPoint(point)
    z_min_host = link_transform.OfPoint(XYZ(point.X, point.Y, z_min)).Z
    z_max_host = link_transform.OfPoint(XYZ(point.X, point.Y, z_max)).Z
    if z_max_host < z_min_host:
        z_min_host, z_max_host = z_max_host, z_min_host
    room_data.append({
        "room": room,
        "id": room.Id,
        "point_host": p_host,
        "z_min": z_min_host,
        "z_max": z_max_host,
    })

host_levels = get_sorted_levels(doc)

print("=" * 80)
print("SORGENTE ROOMS : {}".format(selected_source))
print("ROOMS          : {} totali, {} posizionate, {} elaborabili".format(
    len(all_rooms), len(placed_rooms), len(room_data)))
print("TAG            : {}".format(selected_tag_name))
print("VISTE          : {}".format(len(views_selected)))
print("OPZIONI        : crop={} | rotazione tag={} | salta duplicati={} | senza leader={} | cut plane={}".format(
    opt_crop, opt_rotate, opt_skip_dup, opt_no_leader, opt_cut_plane))
print("=" * 80)

# -------------------------------
# SEZIONE OPERAZIONI PRINCIPALI
# -------------------------------
total_created = 0
total_duplicates = 0
total_out_xy = 0
total_out_z = 0
total_errors = 0
views_skipped = []

t = Transaction(doc, "Tag Rooms in Multiple Views")
t.Start()

try:
    if not tag_type.IsActive:
        tag_type.Activate()
        doc.Regenerate()

    with forms.ProgressBar(title='Tagging Rooms... ({value} di {max_value})',
                           cancellable=True) as pb:
        step = 0
        total_steps = max(1, len(views_selected) * len(room_data))

        for view in views_selected:
            print("")
            print("VISTA: {}".format(view.Name))

            view_range = get_view_range_info(view, host_levels)
            if view_range is None:
                print("  [!] View range non disponibile: vista saltata.")
                views_skipped.append(view.Name)
                step += len(room_data)
                continue
            print("  View range -> {}".format(format_range(view_range)))

            xy_test, xy_desc = build_xy_test(view)
            rotation = get_view_rotation(view)
            print("  Limite XY  -> {}".format(xy_desc if opt_crop else "disattivato da opzioni"))
            if abs(rotation) > 1e-9:
                print("  Vista RUOTATA di {:.2f} gradi (scope box): {}".format(
                    math.degrees(rotation),
                    "i tag verranno allineati" if opt_rotate else "tag NON allineati (opzione off)"))

            existing_keys = collect_existing_tag_keys(view) if opt_skip_dup else set()

            created_in_view = 0
            dup_in_view = 0
            out_xy_in_view = 0
            out_z_in_view = 0
            err_in_view = 0
            tags_to_rotate = []

            for data in room_data:
                if pb.cancelled:
                    t.RollBack()
                    print("\nOperazione annullata dall'utente: nessuna modifica applicata.")
                    script.exit()

                pb.update_progress(step, total_steps)
                step += 1

                room = data["room"]
                p_host = data["point_host"]

                try:
                    # --- Test verticale (view range vs estensione della room) ---
                    top = view_range["top"]
                    low = view_range["low"]
                    if top is not None and data["z_min"] > top + Z_TOL:
                        out_z_in_view += 1
                        continue
                    if low is not None and data["z_max"] < low - Z_TOL:
                        out_z_in_view += 1
                        continue
                    if opt_cut_plane:
                        cut = view_range["cut"]
                        if cut is not None and not (data["z_min"] - Z_TOL <= cut <= data["z_max"] + Z_TOL):
                            out_z_in_view += 1
                            continue

                    # --- Test orizzontale (crop region / scope box, anche ruotati) ---
                    if opt_crop and not xy_test(p_host):
                        out_xy_in_view += 1
                        continue

                    # --- Tag gia' presente? ---
                    key = (link_id_key, get_element_id_value(room.Id))
                    if opt_skip_dup and key in existing_keys:
                        dup_in_view += 1
                        continue

                    # --- Creazione tag ---
                    if selected_link is None:
                        link_elem_id = LinkElementId(room.Id)
                    else:
                        link_elem_id = LinkElementId(selected_link.Id, room.Id)

                    uv_point = UV(p_host.X, p_host.Y)
                    new_tag = doc.Create.NewRoomTag(link_elem_id, uv_point, view.Id)

                    if new_tag is None:
                        err_in_view += 1
                        continue

                    if new_tag.GetTypeId() != tag_type.Id:
                        new_tag.ChangeTypeId(tag_type.Id)

                    if opt_no_leader:
                        try:
                            new_tag.HasLeader = False
                        except Exception:
                            pass

                    if opt_rotate and abs(rotation) > 1e-9:
                        tags_to_rotate.append((new_tag.Id, p_host))

                    existing_keys.add(key)
                    created_in_view += 1

                except Exception as err:
                    err_in_view += 1
                    print("  [!] Room {}: {}".format(get_element_id_value(room.Id), err))

            # --- Rotazione tag allineata alla vista ---
            if tags_to_rotate:
                doc.Regenerate()
                axis_dir = view.ViewDirection
                for tag_id, base_point in tags_to_rotate:
                    try:
                        axis = Line.CreateBound(base_point, base_point + axis_dir)
                        ElementTransformUtils.RotateElement(doc, tag_id, axis, rotation)
                    except Exception as err:
                        print("  [!] Rotazione tag {} non applicata: {}".format(
                            get_element_id_value(tag_id), err))

            total_created += created_in_view
            total_duplicates += dup_in_view
            total_out_xy += out_xy_in_view
            total_out_z += out_z_in_view
            total_errors += err_in_view

            print("  Creati: {} | Fuori crop: {} | Fuori view range: {} | Gia taggate: {} | Errori: {}".format(
                created_in_view, out_xy_in_view, out_z_in_view, dup_in_view, err_in_view))

    t.Commit()

    print("")
    print("=" * 80)
    print("RIEPILOGO")
    print("=" * 80)
    print("Tag creati            : {}".format(total_created))
    print("Rooms gia taggate     : {}".format(total_duplicates))
    print("Fuori crop/scope box  : {}".format(total_out_xy))
    print("Fuori view range      : {}".format(total_out_z))
    print("Errori                : {}".format(total_errors))
    print("Viste elaborate       : {}".format(len(views_selected) - len(views_skipped)))
    if views_skipped:
        print("Viste saltate         : {}".format(", ".join(views_skipped)))

    forms.alert(
        "Operazione completata.\n\n"
        "Tag creati: {}\n"
        "Rooms gia taggate: {}\n"
        "Fuori crop/scope box: {}\n"
        "Fuori view range: {}\n"
        "Errori: {}\n"
        "Viste elaborate: {}".format(
            total_created, total_duplicates, total_out_xy,
            total_out_z, total_errors, len(views_selected) - len(views_skipped)),
        title="Linked Room Tag - Completato")

except Exception as err:
    if t.HasStarted() and not t.HasEnded():
        t.RollBack()
    print("")
    print("ERRORE CRITICO: {}".format(err))
    forms.alert("Errore durante l'operazione, nessuna modifica applicata:\n\n{}".format(err),
                title="Errore", exitscript=True)
