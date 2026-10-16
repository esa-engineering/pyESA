# -*- coding: utf-8 -*-
__title__ = "Last Report"
__doc__ = """Version = 4.0
Date    = 15.10.2026
________________________________________________________________
Description:
Openings (void demand). Place, New Report and Last Report share the same script; only
MODE changes. Settings (families, parameters, clearance, rounding,
merge distance, extra depth, MEC/ELE, To Be Refreshed) come from the
Settings button of this menu.

PLACE: with elements already selected, those already associated with an
opening (openings, MEP elements listing an existing opening) are
checked; the other MEP elements, walls, floors and structural elements
get openings. With nothing selected, a window asks what to select: walls
and floors (host model or RVT links), MEP elements, or an area (rectangle
dragged in a plan view, or the section box of a 3D view). Every straight pipe, duct, cable tray or conduit of the host
model that crosses them gets a Generic Model opening at the centre of
the crossing, sized with the outer section (insulation included) plus
the clearance, rounded up; the depth is the thickness plus the extra
depth on each side. Crossings of the same wall or floor closer than the
merge distance share one rectangular opening (only within the same
discipline, if MEC and ELE are kept separate).
- Opening: e_GEN_RelatedElement_ID = Ids of its MEP elements (111,222);
  e_GEN_RelatedElement_Cat = their categories (Pipes, Ducts);
  e_GEN_Void_Status = Valid
- MEP elements: e_GEN_RelatedElement_ID = Ids of their openings;
  e_GEN_RelatedElement_Cat = Generic Models; e_GEN_Void_Status = Valid
Parameter names can be changed in Settings. Comments and Mark are not
touched. Crossings that already have an opening are skipped.

NEW REPORT: choose the scope (whole model, MEP selection, ARC
selection, area). Every opening is compared with its MEP elements: still
there, still crossing, in the same place and size? Statuses: Valid,
To Be Refreshed (the elements moved but are still inside the opening;
optional), Not Associated (elements removed or moved out, no element
left; MEP elements left without an opening). Statuses and Ids are
written, stale Ids (deleted elements or openings) removed; the MEP
elements of the report are checked for crossings without an opening.
LAST REPORT: no check; the last New Report again, same scope, with the
statuses in the model.
A summary comes first, then a Text Report (export to Excel) or an
Interactive Report to select the elements, create, recreate, move or
delete openings.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"
# la finestra del report e' non modale e lavora tramite ExternalEvent dopo la fine del comando: il motore
# IronPython deve restare vivo
__persistentengine__ = True

import json
import math
import os
import os.path as op
from datetime import datetime

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter, PickBoxStyle
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Collections.Generic import List

from pyrevit import revit, script, DB, forms

MODE = "last"  # "place" | "new" | "last": unica differenza tra Place, New Report e Last Report, tenere allineati

# =================================================================================================================
# Blocco comune a Place, New Report, Last Report e Settings (nomi, parametri, impostazioni predefinite): tenerlo allineato
# =================================================================================================================
CONFIG_NAME = "ESA_Openings"
DEFAULTS = {"fam_wall_rect": "e_GM.PH.VD_Void Demand_Wall_Rect", "fam_wall_round": "e_GM.PH.VD_Void Demand_Wall_Circ",
            "fam_floor_rect": "e_GM.PH.VD_Void Demand_Floor_Rect",
            "fam_floor_round": "e_GM.PH.VD_Void Demand_Floor_Circ",
            "status_param": "e_GEN_Void_Status", "related_id_param": "e_GEN_RelatedElement_ID",
            "related_cat_param": "e_GEN_RelatedElement_Cat",
            "clearance_mm": 25.0, "rounding_mm": 50.0, "merge_mm": 100.0, "extra_depth_mm": 0.0, "layer_gap_mm": 200.0,
            "min_angle_deg": 15.0,
            "separate": False, "misaligned": True}
NUMBER_KEYS = ("clearance_mm", "rounding_mm", "merge_mm", "extra_depth_mm", "layer_gap_mm", "min_angle_deg")
BOOL_KEYS = ("separate", "misaligned")
TEXT_KEYS = ("status_param", "related_id_param", "related_cat_param")  # vuoti: tornano al predefinito
FAMILY_KEYS = {("wall", "rect"): "fam_wall_rect", ("wall", "round"): "fam_wall_round",
               ("floor", "rect"): "fam_floor_rect", ("floor", "round"): "fam_floor_round"}
# famiglie Generic Model NON ospitate, origine al centro, parametri di istanza (condivisi), lunghezza.
# Muro rettangolare: X larghezza, Y profondita', Z altezza. Muro circolare: cilindro lungo X, Width = lunghezza.
# Solaio rettangolare: X e Y in pianta, Z profondita'. Solaio circolare: cilindro lungo Z, Height = lunghezza
PARAM_WIDTH = "e_DIM_Width_1"
PARAM_HEIGHT = "e_DIM_Height_1"
PARAM_DEPTH = "e_DIM_Depth_1"
PARAM_DIAMETER = "e_DIM_Diameter_1"
# lunghezza del foro da parte a parte (lungo l'asse di attraversamento), per famiglia
LENGTH_PARAMS = {("wall", "rect"): PARAM_DEPTH, ("wall", "round"): PARAM_WIDTH,
                 ("floor", "rect"): PARAM_DEPTH, ("floor", "round"): PARAM_HEIGHT}
# famiglie rettangolari: il solido poggia sul livello di riferimento (punto di inserimento = centro della faccia
# inferiore), non e' centrato in Z. Parametro della misura lungo Z: il centro sta a meta' sopra l'inserimento
BASE_PARAMS = {("wall", "rect"): PARAM_HEIGHT, ("floor", "rect"): PARAM_DEPTH}
# stati scritti in e_GEN_Void_Status, dal migliore al peggiore
VALID, TO_REFRESH, NOT_ASSOCIATED = "Valid", "To Be Refreshed", "Not Associated"
# =================================================================================================================

# stati interni del controllo; nel modello si scrive solo lo stato corrispondente di WRITTEN.
# REGROUP: foro da dividere (MEC/ELE separati) o da unire a un vicino secondo Settings, elementi ancora dentro
# PARTIAL: almeno un elemento ancora dentro o in parte fuori, altri tolti o usciti: resta To Be Refreshed
NEW, OK, MISALIGNED, REGROUP, PARTIAL, CHANGED, OBSOLETE = "New", "OK", "Misaligned", "Regroup", "Partial", "Changed", \
    "Obsolete"
WRITTEN = {NEW: VALID, OK: VALID, MISALIGNED: TO_REFRESH, REGROUP: TO_REFRESH, PARTIAL: TO_REFRESH,
           CHANGED: NOT_ASSOCIATED, OBSOLETE: NOT_ASSOCIATED}
WRITTEN_RANK = {VALID: 0, TO_REFRESH: 1, NOT_ASSOCIATED: 2}
OPENING_CATEGORY = "Generic Models"  # e_GEN_RelatedElement_Cat sugli elementi MEP

MOVE_TOL_MM = 5.0
SIZE_TOL_MM = 1.0
MM_PER_FT = 304.8
VERTICAL_SIN = math.sin(math.radians(45.0))   # oltre: elemento MEP verticale, foro a solaio
ROUND_COS = 0.995            # tondo solo se l'attraversamento e' (quasi) perpendicolare
Z = DB.XYZ.BasisZ
MEP_CLASSES = (Pipe, Duct, CableTray, Conduit)
# nomi fissi in inglese per e_GEN_RelatedElement_Cat (Category.Name dipende dalla lingua di Revit)
MEP_LABELS = ((Pipe, "Pipes"), (Duct, "Ducts"), (CableTray, "Cable Trays"), (Conduit, "Conduits"))
MEP_SINGULAR = ((Pipe, "Pipe"), (Duct, "Duct"), (CableTray, "Cable Tray"), (Conduit, "Conduit"))
# stato di ogni elemento elencato in un foro (report del Check)
ELEMENT_INSIDE = "inside"
ELEMENT_PARTLY_OUT = "partly outside"
ELEMENT_MOVED_OUT = "moved out"
ELEMENT_NOT_CROSSING = "no longer crosses"
ELEMENT_UNREADABLE = "section not readable"
ELEMENT_DELETED = "deleted"
ELEMENT_ADDED = "crossing it, now associated"
ELEMENT_ADDED_OUT = "crossing it partly outside, now associated"
ELEMENT_NO_OPENING = "no opening"  # attraversamento senza foro (Missing openings)
MISSING = "Missing openings"  # gruppo del report: attraversamenti senza foro
MEP_CATEGORIES = (DB.BuiltInCategory.OST_PipeCurves, DB.BuiltInCategory.OST_DuctCurves,
                  DB.BuiltInCategory.OST_CableTray, DB.BuiltInCategory.OST_Conduit)
# muri e solai: foro a muro / a solaio; strutture (travi, pilastri, fondazioni): secondo la direzione dell'elemento MEP
HOST_KINDS = {int(DB.BuiltInCategory.OST_Walls): "wall", int(DB.BuiltInCategory.OST_Floors): "floor",
              int(DB.BuiltInCategory.OST_StructuralFraming): "auto", int(DB.BuiltInCategory.OST_StructuralColumns): "auto",
              int(DB.BuiltInCategory.OST_StructuralFoundation): "auto"}
STATE_RANK = {OK: 0, NEW: 1, MISALIGNED: 2, REGROUP: 3, PARTIAL: 4, CHANGED: 5, OBSOLETE: 6}
TITLE = {"place": "Openings Place", "new": "Openings New Report", "last": "Openings Last Report"}[MODE]
# finestre di Place (OpeningsPlace_ui.py e XAML solo nella cartella di Place: importati solo da place_main)
PICK_XAML = "OpeningsPlace_pick.xaml"
CONFIRM_XAML = "OpeningsPlace_confirm.xaml"
# finestre dei report (OpeningsCheck_ui.py, OpeningsReport_xlsx.py e XAML: copie identiche nelle cartelle di Place,
# NewReport e LastReport)
SCOPE_XAML = "OpeningsCheck_scope.xaml"
SUMMARY_XAML = "OpeningsCheck_summary.xaml"
REPORT_XAML = "OpeningsCheck_report.xaml"
TEXT_XAML = "OpeningsCheck_text.xaml"

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


def load_settings():
    settings = dict(DEFAULTS)
    try:
        config = script.get_config(CONFIG_NAME)
        for key in DEFAULTS:
            value = config.get_option(key, DEFAULTS[key])
            if key in NUMBER_KEYS:
                value = float(value)
            elif key in BOOL_KEYS and not isinstance(value, bool):
                value = str(value).strip().lower() in ("true", "1", "yes")
            elif key in TEXT_KEYS:
                value = str(value or "").strip() or DEFAULTS[key]
            settings[key] = value
    except Exception:
        pass
    return settings


SETTINGS = load_settings()
STATUS_PARAM = SETTINGS["status_param"]
ID_PARAM = SETTINGS["related_id_param"]
CAT_PARAM = SETTINGS["related_cat_param"]
# sotto quest'angolo con il piano di muro o solaio l'elemento MEP corre dentro (scarichi a pavimento, montanti nel
# muro): nessun foro
MIN_CROSS_SIN = math.sin(math.radians(min(max(SETTINGS["min_angle_deg"], 0.0), 89.0)))


def id_value(eid):
    return eid.Value if hasattr(eid, "Value") else eid.IntegerValue  # Revit 2026+ / <= 2025


def mm(value_ft):
    return value_ft * MM_PER_FT


def round_up(value_mm):
    step = SETTINGS["rounding_mm"]
    if step <= 0:
        return value_mm
    return math.ceil(value_mm / step - 1e-9) * step


def level_elevation(level):
    # ProjectElevation e' coerente con la geometria (coordinate interne), Elevation puo' riferirsi al Survey Point
    value = getattr(level, "ProjectElevation", None)
    return level.Elevation if value is None else value


# ----------------------------------------------------------------------------------------------------------------
# Parametri e_GEN: stato, Id e categorie degli elementi collegati (foro <-> elementi MEP)
# ----------------------------------------------------------------------------------------------------------------
def get_text(element, name):
    """Valore di testo del parametro, "" se vuoto, None se il parametro non c'e'."""
    param = element.LookupParameter(name)
    if param is None:
        return None
    return param.AsString() or ""


def set_text(element, name, value, missing):
    """Scrive solo se il valore cambia (modelli condivisi: non si prende in prestito un elemento per niente)."""
    param = element.LookupParameter(name)
    if param is None or param.IsReadOnly:
        missing.add("{} on {}".format(name, element.Category.Name if element.Category else "element"))
        return
    if (param.AsString() or "") != (value or ""):
        param.Set(value)


def discipline(mep):
    return "MEC" if isinstance(mep, (Pipe, Duct)) else "ELE"


def disciplines_text(meps):
    found = set(discipline(m) for m in meps)
    return "+".join(d for d in ("MEC", "ELE") if d in found)


def categories_text(meps):
    """Categorie degli elementi MEP del foro, nomi fissi in inglese: "Pipes, Ducts"."""
    return ", ".join(label for cls, label in MEP_LABELS if any(isinstance(m, cls) for m in meps))


def worst(statuses):
    """Stato scritto peggiore (Valid < To Be Refreshed < Not Associated); testi sconosciuti contano come Valid."""
    return max(statuses or [VALID], key=lambda s: WRITTEN_RANK.get(s, 0))


def code_of(opening):
    """Codice del foro scritto sugli elementi MEP: il suo Id."""
    return str(id_value(opening.Id))


def write_opening(instance, meps, state, missing):
    """Stato, Id e categorie degli elementi MEP del foro."""
    set_text(instance, STATUS_PARAM, WRITTEN[state], missing)
    set_text(instance, ID_PARAM, ",".join(str(id_value(m.Id)) for m in meps), missing)
    set_text(instance, CAT_PARAM, categories_text(meps), missing)


def split_list(text):
    return [part.strip() for part in (text or "").split(",") if part.strip()]


def element_by_id(text):
    try:
        return doc.GetElement(DB.ElementId(int(text)))
    except Exception:
        return None


# ----------------------------------------------------------------------------------------------------------------
# Geometria
# ----------------------------------------------------------------------------------------------------------------
def solids_of(element):
    options = DB.Options()
    options.DetailLevel = DB.ViewDetailLevel.Fine
    result = []
    try:
        pending = list(element.get_Geometry(options) or [])
    except Exception:
        return result
    while pending:
        item = pending.pop()
        if isinstance(item, DB.GeometryInstance):
            pending.extend(item.GetInstanceGeometry())
        elif isinstance(item, DB.Solid) and item.Volume > 1e-9:
            result.append(item)
    return result


def solid_points(solids):
    points = []
    for solid in solids:
        for edge in solid.Edges:
            curve = edge.AsCurve()
            points.extend((curve.GetEndPoint(0), curve.GetEndPoint(1)))
    return points


def unit(vector):
    return vector.Normalize() if vector.GetLength() > 1e-9 else None


def mep_section(mep):
    """Sezione esterna, isolamento compreso, letta dalla geometria: (asse X, asse Y, semi-larghezza, semi-altezza,
    tonda). Gli assi sono quelli del connettore, perpendicolari al tratto; le misure vengono dai solidi di elemento
    e isolamento, cosi' valgono anche per canali ruotati e passerelle di costa."""
    conn = [c for c in mep.ConnectorManager.Connectors if c.ConnectorType == DB.ConnectorType.End][0]
    bx, by = conn.CoordinateSystem.BasisX, conn.CoordinateSystem.BasisY
    solids = solids_of(mep)
    if isinstance(mep, (Pipe, Duct)):
        for insulation_id in DB.InsulationLiningBase.GetInsulationIds(doc, mep.Id):
            solids.extend(solids_of(doc.GetElement(insulation_id)))
    origin = mep.Location.Curve.GetEndPoint(0)
    points = solid_points(solids)
    if not points:
        raise ValueError("no geometry")
    xs = [p.Subtract(origin).DotProduct(bx) for p in points]
    ys = [p.Subtract(origin).DotProduct(by) for p in points]
    return bx, by, (max(xs) - min(xs)) / 2.0, (max(ys) - min(ys)) / 2.0, conn.Shape == DB.ConnectorProfileType.Round


def section_points(center, section):
    bx, by, hx, hy, round_shape = section
    if round_shape:
        r = max(hx, hy)
        return [center.Add(bx.Multiply(r * math.cos(a))).Add(by.Multiply(r * math.sin(a)))
                for a in [i * math.pi / 8.0 for i in range(16)]]
    return [center.Add(bx.Multiply(sx * hx)).Add(by.Multiply(sy * hy)) for sx in (-1, 1) for sy in (-1, 1)]


def make_frame(kind, w, hint):
    """(u, v, w): w attraverso l'elemento, u orizzontale (asse X della famiglia), v la verticale o w x u."""
    if kind == "wall":
        w = unit(DB.XYZ(w.X, w.Y, 0))
        return Z.CrossProduct(w), Z, w
    u = unit(DB.XYZ(hint.X, hint.Y, 0)) if hint is not None else None
    u = u or DB.XYZ.BasisX
    return u, Z.CrossProduct(u), Z


def opening_shape(items, frame):
    """Rettangolo che contiene tutti gli attraversamenti nel piano (u, v) e spessore lungo w. Ritorna (centro,
    estensione u, estensione v, spessore) in piedi."""
    u, v, w = frame
    o = items[0]["mid"]
    pu, pv, pw = [], [], []
    for item in items:
        for p in item["points"]:
            d = p.Subtract(o)
            pu.append(d.DotProduct(u))
            pv.append(d.DotProduct(v))
        for p in (item["p0"], item["p1"]):
            pw.append(p.Subtract(o).DotProduct(w))
    center = o.Add(u.Multiply((max(pu) + min(pu)) / 2.0)).Add(v.Multiply((max(pv) + min(pv)) / 2.0)).Add(
        w.Multiply((max(pw) + min(pw)) / 2.0))
    return center, max(pu) - min(pu), max(pv) - min(pv), max(pw) - min(pw)


def sizes(items, frame):
    """Dati del foro: centro, forma e misure in mm (franco, arrotondamento e sporgenza compresi)."""
    center, ext_u, ext_v, thickness = opening_shape(items, frame)
    clear = 2.0 * SETTINGS["clearance_mm"]
    depth = mm(thickness) + 2.0 * SETTINGS["extra_depth_mm"]
    if len(items) == 1 and items[0]["round"]:
        return {"center": center, "shape": "round", "diameter": round_up(mm(max(ext_u, ext_v)) + clear),
                "depth": depth}
    return {"center": center, "shape": "rect", "width": round_up(mm(ext_u) + clear),
            "height": round_up(mm(ext_v) + clear), "depth": depth}


def crossing_item(mep, section, p0, p1, frame, kind):
    direction = mep.Location.Curve.Direction
    return {"mep": mep, "p0": p0, "p1": p1, "mid": p0.Add(p1).Divide(2.0), "frame": frame, "kind": kind,
            "points": section_points(p0, section) + section_points(p1, section), "discipline": discipline(mep),
            "round": section[4] and abs(direction.DotProduct(frame[2])) > ROUND_COS, "section": section}


# ----------------------------------------------------------------------------------------------------------------
# Famiglie, fori esistenti, inserimento
# ----------------------------------------------------------------------------------------------------------------
def family_symbols():
    """{(muro|solaio, rect|round): primo tipo della famiglia scelta in Settings}; mancanti: elenco dei nomi."""
    by_name = {}
    for symbol in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(
            DB.BuiltInCategory.OST_GenericModel):
        by_name.setdefault(symbol.Family.Name, symbol)
    symbols, missing = {}, []
    for key, setting in FAMILY_KEYS.items():
        name = SETTINGS[setting]
        if name in by_name:
            symbols[key] = by_name[name]
        else:
            missing.append(name)
    return symbols, sorted(set(missing))


def existing_openings():
    """Fori gia' gestiti: Generic Model con e_GEN_Void_Status compilato (qualunque famiglia)."""
    return [i for i in DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).OfCategory(
        DB.BuiltInCategory.OST_GenericModel) if get_text(i, STATUS_PARAM)]


def pick_level(z):
    levels = sorted(DB.FilteredElementCollector(doc).OfClass(DB.Level), key=level_elevation)
    below = [l for l in levels if level_elevation(l) <= z + 1e-6]
    return below[-1] if below else levels[0]


def create_opening(data, frame, kind, symbols, meps, state, missing):
    """Inserisce il foro con il solido centrato sull'attraversamento (famiglie rettangolari: inserimento sotto il
    centro di meta' altezza), con l'asse X della famiglia lungo u (muro circolare: lungo w, l'asse del cilindro);
    scrive misure, stato, Id e categorie dei MEP."""
    symbol = symbols[(kind, data["shape"])]
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()
    length = LENGTH_PARAMS[(kind, data["shape"])]
    values = ((PARAM_DIAMETER, data["diameter"]), (length, data["depth"])) if data["shape"] == "round" else \
        ((PARAM_WIDTH, data["width"]), (PARAM_HEIGHT, data["height"]), (length, data["depth"]))
    base = BASE_PARAMS.get((kind, data["shape"]))
    lift = dict(values)[base] / 2.0 / MM_PER_FT if base else 0.0
    point = data["center"].Subtract(Z.Multiply(lift))
    instance = doc.Create.NewFamilyInstance(point, symbol, pick_level(point.Z), StructuralType.NonStructural)
    doc.Regenerate()
    current = instance.GetTransform().BasisX
    u = frame[2] if (kind, data["shape"]) == ("wall", "round") else frame[0]
    angle = math.atan2(current.CrossProduct(u).Z, current.DotProduct(u))
    if abs(angle) > 1e-9:
        axis = DB.Line.CreateBound(instance.Location.Point, instance.Location.Point.Add(Z))
        DB.ElementTransformUtils.RotateElement(doc, instance.Id, axis, angle)
    delta = point.Subtract(instance.Location.Point)
    if delta.GetLength() > 1e-9:
        DB.ElementTransformUtils.MoveElement(doc, instance.Id, delta)
    for name, value in values:
        param = instance.LookupParameter(name)
        if param is None or param.IsReadOnly:
            missing.add("{} on family {}".format(name, symbol.Family.Name))
        else:
            param.Set(value / MM_PER_FT)
    write_opening(instance, meps, state, missing)
    return instance


def size_text(data):
    if data["shape"] == "round":
        return "D{:.0f} x {:.0f}".format(data["diameter"], data["depth"])
    return "{:.0f} x {:.0f} x {:.0f}".format(data["width"], data["height"], data["depth"])


def write_mep(mep, codes, status, missing):
    """Id dei fori del MEP, categoria e stato scritto (il peggiore dei suoi fori). Senza fori: Id e categoria
    vuoti, stato Not Associated."""
    set_text(mep, ID_PARAM, ",".join(codes), missing)
    set_text(mep, CAT_PARAM, OPENING_CATEGORY if codes else "", missing)
    set_text(mep, STATUS_PARAM, status if codes else NOT_ASSOCIATED, missing)


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
        host = doc.GetElement(reference.ElementId)
        if isinstance(host, DB.RevitLinkInstance):
            link_doc = host.GetLinkDocument()
            element = link_doc.GetElement(reference.LinkedElementId) if link_doc is not None else None
        else:
            element = host
        return element is not None and element.Category is not None and \
            id_value(element.Category.Id) in HOST_KINDS


def pick_hosts():
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.PointOnElement, HostFilter(),
                                           "Openings: select walls, floors and structural elements, also in linked "
                                           "models, then Finish")
    except OperationCanceledException:
        script.exit()
    hosts, seen = [], set()
    for ref in refs:
        host = doc.GetElement(ref.ElementId)
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


def is_straight_mep(element):
    return isinstance(element, MEP_CLASSES) and isinstance(element.Location.Curve, DB.Line)


def pick_meps():
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, MepPickFilter(),
                                           "Openings: select pipes, ducts, cable trays and conduits, then Finish")
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def hosts_from_meps(meps):
    """Modalita' inversa: muri, solai e strutture del modello host e dei link caricati che toccano il bounding box
    di almeno un elemento MEP scelto. Stesso formato di pick_hosts."""
    categories = List[DB.BuiltInCategory]([DB.BuiltInCategory.OST_Walls, DB.BuiltInCategory.OST_Floors,
                                           DB.BuiltInCategory.OST_StructuralFraming,
                                           DB.BuiltInCategory.OST_StructuralColumns,
                                           DB.BuiltInCategory.OST_StructuralFoundation])
    boxes = [m.get_BoundingBox(None) for m in meps]
    boxes = [b for b in boxes if b is not None]
    hosts, seen = [], set()
    for source_doc, transform, link_id in model_sources():
        inverse = transform.Inverse if transform is not None else None
        model = source_doc.Title if link_id else "Host model"
        for box in boxes:
            corners = [DB.XYZ(x, y, z) for x in (box.Min.X, box.Max.X) for y in (box.Min.Y, box.Max.Y)
                       for z in (box.Min.Z, box.Max.Z)]
            if inverse is not None:
                corners = [inverse.OfPoint(c) for c in corners]
            outline = DB.Outline(DB.XYZ(min(c.X for c in corners), min(c.Y for c in corners),
                                        min(c.Z for c in corners)),
                                 DB.XYZ(max(c.X for c in corners), max(c.Y for c in corners),
                                        max(c.Z for c in corners)))
            for element in DB.FilteredElementCollector(source_doc).WherePasses(
                    DB.ElementMulticategoryFilter(categories)).WhereElementIsNotElementType()\
                    .WherePasses(DB.BoundingBoxIntersectsFilter(outline)):
                key = (link_id, id_value(element.Id))
                if key not in seen:
                    seen.add(key)
                    hosts.append((key, element, transform, model))
    return hosts


# ----------------------------------------------------------------------------------------------------------------
# Area: riquadro trascinato in pianta (quote dal view range, come Quick SectionBox) o section box della vista 3D
# ----------------------------------------------------------------------------------------------------------------
MIN_AREA_MM = 10.0           # riquadro piu' stretto di cosi': e' un clic, non un trascinamento
FALLBACK_HEIGHT_MM = 4000.0  # altezza se il view range non e' leggibile (livelli illimitati)
PLAN_TYPES = (DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.EngineeringPlan)
HOST_CATEGORIES = (DB.BuiltInCategory.OST_Walls, DB.BuiltInCategory.OST_Floors,
                   DB.BuiltInCategory.OST_StructuralFraming, DB.BuiltInCategory.OST_StructuralColumns,
                   DB.BuiltInCategory.OST_StructuralFoundation)


def flat_unit(vector):
    flat = DB.XYZ(vector.X, vector.Y, 0)
    return flat.Normalize() if flat.GetLength() > 1e-9 else None


def area_axes(area):
    right = DB.XYZ(area["right"][0], area["right"][1], 0)
    return right, Z.CrossProduct(right)


def make_area(right, points, z_min, z_max, source):
    """Area allineata agli assi della vista: min / max = (destra, alto, quota) dei punti proiettati. Solo numeri: si
    salva cosi' com'e' nel file del Last Report e si rilegge per il Refresh."""
    up = Z.CrossProduct(right)
    us = [p.DotProduct(right) for p in points]
    vs = [p.DotProduct(up) for p in points]
    return {"right": [right.X, right.Y], "min": [min(us), min(vs), z_min], "max": [max(us), max(vs), z_max],
            "source": source}


def area_hit(area, point, z_low, z_high):
    """Il punto sta nel rettangolo in pianta e l'intervallo di quote [z_low, z_high] tocca quello dell'area."""
    right, up = area_axes(area)
    tol = MOVE_TOL_MM / MM_PER_FT
    u, v = point.DotProduct(right), point.DotProduct(up)
    return area["min"][0] - tol <= u <= area["max"][0] + tol and area["min"][1] - tol <= v <= area["max"][1] + tol \
        and z_high >= area["min"][2] - tol and z_low <= area["max"][2] + tol


def crossing_in_area(area, p0, p1):
    """Attraversamento nell'area: punto medio nel rettangolo, tratto che tocca le quote. Il tratto e non il solo punto
    medio: con il view range consueto (fondo al livello) il solaio sotto il livello ha il centro fuori quota."""
    return area_hit(area, p0.Add(p1).Divide(2.0), min(p0.Z, p1.Z), max(p0.Z, p1.Z))


def opening_in_area(area, opening):
    box = opening.get_BoundingBox(None)
    if box is None:
        return False
    return area_hit(area, box.Min.Add(box.Max).Divide(2.0), box.Min.Z, box.Max.Z)


def area_outline(area, transform=None):
    """Outline allineato agli assi del documento (del link, con transform) che contiene l'area: filtro grossolano per
    i collector, il test vero e' area_hit."""
    right, up = area_axes(area)
    inverse = transform.Inverse if transform is not None else None
    corners = []
    for u in (area["min"][0], area["max"][0]):
        for v in (area["min"][1], area["max"][1]):
            for z in (area["min"][2], area["max"][2]):
                corner = right.Multiply(u).Add(up.Multiply(v)).Add(Z.Multiply(z))
                corners.append(inverse.OfPoint(corner) if inverse is not None else corner)
    return DB.Outline(DB.XYZ(min(c.X for c in corners), min(c.Y for c in corners), min(c.Z for c in corners)),
                      DB.XYZ(max(c.X for c in corners), max(c.Y for c in corners), max(c.Z for c in corners)))


def area_text(area):
    return "{}, {:.0f} x {:.0f} mm, height {:.0f} to {:.0f} mm".format(
        area["source"], (area["max"][0] - area["min"][0]) * MM_PER_FT, (area["max"][1] - area["min"][1]) * MM_PER_FT,
        area["min"][2] * MM_PER_FT, area["max"][2] * MM_PER_FT)


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
        level = doc.GetElement(view_range.GetLevelId(plane))
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
    view = doc.ActiveView
    if isinstance(view, DB.View3D) and not view.IsTemplate:
        if not view.IsSectionBoxActive:
            forms.alert("This 3D view has no section box. Turn it on (or use Quick SectionBox), or drag a rectangle "
                        "in a plan view.", title=TITLE, exitscript=True)
        box = view.GetSectionBox()
        corners = [box.Transform.OfPoint(DB.XYZ(x, y, z)) for x in (box.Min.X, box.Max.X)
                   for y in (box.Min.Y, box.Max.Y) for z in (box.Min.Z, box.Max.Z)]
        right = flat_unit(box.Transform.BasisX) or DB.XYZ.BasisX
        return make_area(right, corners, min(c.Z for c in corners), max(c.Z for c in corners),
                         "section box of the 3D view '{}'".format(view.Name))
    if isinstance(view, DB.ViewPlan) and not view.IsTemplate and view.ViewType in PLAN_TYPES:
        try:
            picked = uidoc.Selection.PickBox(PickBoxStyle.Enclosing,
                                             "Openings: drag a rectangle from corner to corner (ESC to cancel)")
        except OperationCanceledException:
            script.exit()
        z_min, z_max, note = view_range_z(view)
        area = make_area(plan_right(view), [picked.Min, picked.Max], z_min, z_max,
                         "rectangle in the plan '{}' ({})".format(view.Name, note))
        if min(area["max"][0] - area["min"][0], area["max"][1] - area["min"][1]) * MM_PER_FT < MIN_AREA_MM:
            forms.alert("The rectangle is less than {:g} mm wide or high: drag from corner to corner.".format(
                MIN_AREA_MM), title=TITLE, exitscript=True)
        return area
    forms.alert("Area: run it from a floor or ceiling plan (drag a rectangle) or from a 3D view with a section box.",
                title=TITLE, exitscript=True)


def hosts_in_area(area):
    """Muri, solai e strutture del modello host e dei link caricati che toccano l'area. Stesso formato di
    pick_hosts."""
    categories = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory](HOST_CATEGORIES))
    hosts = []
    for source_doc, transform, link_id in model_sources():
        model = source_doc.Title if link_id else "Host model"
        for element in DB.FilteredElementCollector(source_doc).WherePasses(categories).WhereElementIsNotElementType()\
                .WherePasses(DB.BoundingBoxIntersectsFilter(area_outline(area, transform))):
            hosts.append(((link_id, id_value(element.Id)), element, transform, model))
    return hosts


def meps_in_area(area):
    """Elementi MEP rettilinei del modello host che toccano l'area (anche solo in parte)."""
    categories = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory](MEP_CATEGORIES))
    return [e for e in DB.FilteredElementCollector(doc).WherePasses(categories).WhereElementIsNotElementType()
            .WherePasses(DB.BoundingBoxIntersectsFilter(area_outline(area))) if is_straight_mep(e)]


def host_normal(element, transform, point):
    """Normale del muro (o della trave) nel punto, anche curvo, in coordinate host."""
    curve = element.Location.Curve
    if transform is not None:
        curve = curve.CreateTransformed(transform)
    projection = curve.Project(point)
    tangent = curve.ComputeDerivatives(projection.Parameter, False).BasisX
    return Z.CrossProduct(DB.XYZ(tangent.X, tangent.Y, 0).Normalize())


def crossing_normal(element, transform, point, kind, direction):
    """Asse di attraversamento del foro: solaio, la verticale; muro o trave, la normale nel punto; pilastri e casi
    senza asse orizzontale leggibile, la direzione dell'elemento MEP in pianta."""
    if kind == "floor":
        return Z
    try:
        return host_normal(element, transform, point)
    except Exception:
        return DB.XYZ(direction.X, direction.Y, 0).Normalize()


def mep_near(solids):
    """Elementi MEP rettilinei del modello host vicini ai solidi (filtro sul bounding box)."""
    points = solid_points(solids)
    low = DB.XYZ(min(p.X for p in points), min(p.Y for p in points), min(p.Z for p in points))
    high = DB.XYZ(max(p.X for p in points), max(p.Y for p in points), max(p.Z for p in points))
    box_filter = DB.BoundingBoxIntersectsFilter(DB.Outline(low, high))
    categories = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory](MEP_CATEGORIES))
    return [e for e in DB.FilteredElementCollector(doc).WherePasses(categories).WherePasses(box_filter)
            .WhereElementIsNotElementType() if isinstance(e, MEP_CLASSES) and isinstance(e.Location.Curve, DB.Line)]


def inside_existing(point, boxes):
    """Foro gestito il cui bounding box contiene il punto, o None."""
    for low, high, opening in boxes:
        if low.X <= point.X <= high.X and low.Y <= point.Y <= high.Y and low.Z <= point.Z <= high.Z:
            return opening
    return None


ADJACENT_SEARCH = 1000.0 / MM_PER_FT  # margine attorno all'elemento scelto in cui si cercano gli strati vicini
MAX_LAYER = 1000.0 / MM_PER_FT  # strato attraversato per piu' di cosi': l'elemento MEP ci corre dentro, non conta
PARALLEL_COS = 0.98             # strati di muro paralleli all'elemento scelto entro circa 11 gradi
NEIGHBOUR_CACHE = {}  # (id del link o 0, id dell'elemento) -> solidi in coordinate host


def model_sources():
    """(documento, transform o None, id del link o 0) del modello host e dei link RVT caricati."""
    sources = [(doc, None, 0)]
    for link in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance):
        link_doc = link.GetLinkDocument()
        if link_doc is not None:
            sources.append((link_doc, link.GetTotalTransform(), id_value(link.Id)))
    return sources


def neighbours(solids, kind):
    """Muri (o solai, per i fori a solaio) del modello host e dei link vicini all'elemento scelto: gli strati di una
    struttura composita modellati come elementi separati. Ritorna [(chiave, etichetta, solidi in coordinate host,
    elemento, transform)]."""
    points = solid_points(solids)
    low = DB.XYZ(min(p.X for p in points) - ADJACENT_SEARCH, min(p.Y for p in points) - ADJACENT_SEARCH,
                 min(p.Z for p in points) - ADJACENT_SEARCH)
    high = DB.XYZ(max(p.X for p in points) + ADJACENT_SEARCH, max(p.Y for p in points) + ADJACENT_SEARCH,
                  max(p.Z for p in points) + ADJACENT_SEARCH)
    category = DB.BuiltInCategory.OST_Walls if kind == "wall" else DB.BuiltInCategory.OST_Floors
    found = []
    for source_doc, transform, link_id in model_sources():
        corners = [DB.XYZ(x, y, z) for x in (low.X, high.X) for y in (low.Y, high.Y) for z in (low.Z, high.Z)]
        if transform is not None:
            inverse = transform.Inverse
            corners = [inverse.OfPoint(c) for c in corners]
        outline = DB.Outline(DB.XYZ(min(c.X for c in corners), min(c.Y for c in corners), min(c.Z for c in corners)),
                             DB.XYZ(max(c.X for c in corners), max(c.Y for c in corners), max(c.Z for c in corners)))
        for element in DB.FilteredElementCollector(source_doc).OfCategory(category).WhereElementIsNotElementType()\
                .WherePasses(DB.BoundingBoxIntersectsFilter(outline)):
            key = (link_id, id_value(element.Id))
            if key not in NEIGHBOUR_CACHE:
                own = solids_of(element)
                NEIGHBOUR_CACHE[key] = [DB.SolidUtils.CreateTransformed(s, transform) for s in own] \
                    if transform is not None else own
            label = "{} {} ({})".format(element.Category.Name, id_value(element.Id),
                                        source_doc.Title if link_id else "Host model")
            found.append((key, label, NEIGHBOUR_CACHE[key], element, transform))
    return found


def parallel_layer(element, transform, kind, w, point):
    """Lo strato e' parallelo all'elemento scelto (muri: normale nel punto quasi parallela a w; solai: sempre). Un
    muro a T o d'angolo non e' uno strato della stessa struttura."""
    if kind != "wall":
        return True
    try:
        return abs(host_normal(element, transform, point).DotProduct(w)) > PARALLEL_COS
    except Exception:
        return True  # muro senza linea di posizione leggibile: si giudica solo dalla distanza


def extend_through(p0, p1, layers, options, kind, w):
    """Allunga l'attraversamento attraverso tutti gli strati della struttura composita (elementi separati): ogni
    strato parallelo che, lungo l'asse dell'elemento MEP, sta entro l'intercapedine massima di Settings da quanto
    gia' attraversato viene incluso, intercapedine compresa. Si ripete finche' non se ne trovano altri. Ritorna
    (p0, p1, chiavi degli strati attraversati)."""
    gap = max(SETTINGS["layer_gap_mm"], 5.0) / MM_PER_FT
    direction = p1.Subtract(p0).Normalize()
    t0, t1 = 0.0, p1.DistanceTo(p0)
    used = set()
    changed = True
    while changed:
        changed = False
        probe = DB.Line.CreateBound(p0.Add(direction.Multiply(t0 - gap - MAX_LAYER)),
                                    p0.Add(direction.Multiply(t1 + gap + MAX_LAYER)))
        for key, label, solids, element, transform in layers:
            for solid in solids:
                try:
                    result = solid.IntersectWithCurve(probe, options)
                except Exception:
                    continue
                for i in range(result.SegmentCount):
                    segment = result.GetCurveSegment(i)
                    a = segment.GetEndPoint(0).Subtract(p0).DotProduct(direction)
                    b = segment.GetEndPoint(1).Subtract(p0).DotProduct(direction)
                    a, b = min(a, b), max(a, b)
                    if b - a < 1e-6 or b - a > MAX_LAYER or b < t0 - gap or a > t1 + gap:
                        continue  # troppo lontano da quanto gia' attraversato, o l'elemento MEP ci corre dentro
                    if key not in used:
                        middle = p0.Add(direction.Multiply((a + b) / 2.0))
                        if not parallel_layer(element, transform, kind, w, middle):
                            continue
                    used.add(key)
                    if a < t0 - 1e-6 or b > t1 + 1e-6:
                        t0, t1 = min(t0, a), max(t1, b)
                        changed = True
    return p0.Add(direction.Multiply(t0)), p0.Add(direction.Multiply(t1)), used



def find_crossings(hosts, stats, only=None, keep_existing=False, progress=None, area=None):
    """Attraversamenti degli elementi MEP con muri, solai e strutture, allungati attraverso gli strati a contatto
    (muri o solai separati di una struttura composita). only: id degli elementi MEP da considerare (modalita'
    inversa), None = tutti quelli vicini. Quelli che hanno gia' un foro si saltano (Place) o, con keep_existing (check),
    restano con item["opening"] = il foro che li contiene. progress(i, n): avanzamento sugli elementi attraversati.
    area: solo gli attraversamenti dentro l'area (crossing_in_area), gli altri contati in stats["outside_area"]."""
    boxes = []
    for opening in existing_openings():
        box = opening.get_BoundingBox(None)
        if box is not None:
            boxes.append((box.Min, box.Max, opening))
    items = []
    options = DB.SolidCurveIntersectionOptions()
    for index, (key, element, transform, model) in enumerate(hosts):
        if progress is not None:
            progress(index, len(hosts))
        host_kind = HOST_KINDS[id_value(element.Category.Id)]
        solids = solids_of(element)
        if transform is not None:
            solids = [DB.SolidUtils.CreateTransformed(s, transform) for s in solids]
        if not solids:
            stats["problems"].append("{} {} ({}): no geometry".format(element.Category.Name,
                                                                       id_value(element.Id), model))
            continue
        # strati separati di una struttura composita: solo per muri e solai
        layers = neighbours(solids, host_kind) if host_kind != "auto" else []
        labels = dict((layer[0], layer[1]) for layer in layers)
        for mep in mep_near(solids):
            if only is not None and id_value(mep.Id) not in only:
                continue
            line = mep.Location.Curve
            try:
                section = mep_section(mep)
            except Exception:
                stats["problems"].append("{} {}: section not readable".format(mep.Category.Name,
                                                                               id_value(mep.Id)))
                continue
            for solid in solids:
                result = solid.IntersectWithCurve(line, options)
                for i in range(result.SegmentCount):
                    segment = result.GetCurveSegment(i)
                    if mm(segment.Length) < 1.0:
                        continue
                    p0, p1 = segment.GetEndPoint(0), segment.GetEndPoint(1)
                    mid = p0.Add(p1).Divide(2.0)
                    if host_kind == "auto":
                        # strutture: foro a solaio per elementi MEP verticali, a muro per gli altri; oltre 1 m
                        # dentro la struttura l'elemento ci corre dentro (es. montante in un pilastro)
                        if segment.Length > MAX_LAYER:
                            stats["inside"] += 1
                            continue
                        kind = "floor" if abs(line.Direction.Z) > VERTICAL_SIN else "wall"
                    else:
                        kind = host_kind
                    w = crossing_normal(element, transform, mid, kind, line.Direction)
                    if abs(line.Direction.DotProduct(w)) < MIN_CROSS_SIN:
                        stats["inside"] += 1  # corre dentro il muro o il solaio: nessun foro
                        continue
                    p0, p1, used = extend_through(p0, p1, layers, options, kind, w)
                    if area is not None and not crossing_in_area(area, p0, p1):
                        stats["outside_area"] = stats.get("outside_area", 0) + 1
                        continue
                    mid = p0.Add(p1).Divide(2.0)
                    opening = inside_existing(mid, boxes)
                    if opening is not None:
                        stats["existing"] += 1
                        if not keep_existing:
                            continue
                    item = crossing_item(mep, section, p0, p1, make_frame(kind, w, section[0]), kind)
                    item["opening"] = opening
                    # chiave dell'intera pila di strati: scegliendo piu' strati della stessa struttura, i loro
                    # attraversamenti finiscono nello stesso foro invece di dare fori sovrapposti
                    stack = tuple(sorted(used | set([key])))
                    others = [labels[k] for k in sorted(used) if k != key and k in labels]
                    if others:
                        stats["extended"] += 1
                    item["host"] = (stack, element, model + (" + " + ", ".join(others) if others else ""))
                    item["host_kind"] = host_kind
                    items.append(item)
    return items


def merge_overlapping(items):
    """Attraversamenti dello stesso elemento MEP attraverso elementi diversi che si sovrappongono o si toccano lungo
    l'elemento (es. trave annegata nel muro, pilastro nel muro, strato di un'altra categoria): diventano uno solo, con
    l'intervallo complessivo e tutti gli elementi attraversati. Senza questo, ogni elemento darebbe un suo foro, quasi
    coincidente con l'altro. Vale per stesso tipo di foro (muro / solaio) e assi di attraversamento paralleli."""
    gap = max(SETTINGS["layer_gap_mm"], 5.0) / MM_PER_FT
    merged = []
    for item in items:
        direction = item["mep"].Location.Curve.Direction
        target = None
        for other in merged:
            if other["mep"].Id != item["mep"].Id or other["kind"] != item["kind"]:
                continue
            if abs(other["frame"][2].DotProduct(item["frame"][2])) < PARALLEL_COS:
                continue
            a = sorted((other["p0"].DotProduct(direction), other["p1"].DotProduct(direction)))
            b = sorted((item["p0"].DotProduct(direction), item["p1"].DotProduct(direction)))
            if b[0] <= a[1] + gap and a[0] <= b[1] + gap:
                target = other
                break
        if target is None:
            merged.append(dict(item))
            continue
        ends = [target["p0"], target["p1"], item["p0"], item["p1"]]
        p0 = min(ends, key=lambda p: p.DotProduct(direction))
        p1 = max(ends, key=lambda p: p.DotProduct(direction))
        # il sistema del foro viene da un muro o solaio, se c'e' (le strutture lo hanno solo approssimato)
        keep = target if target.get("host_kind") != "auto" or item.get("host_kind") == "auto" else item
        frame, kind = keep["frame"], keep["kind"]
        stack = tuple(sorted(set(target["host"][0]) | set(item["host"][0])))
        label = target["host"][2] + " + " + "{} {} ({})".format(item["host"][1].Category.Name,
                                                                id_value(item["host"][1].Id), item["host"][2])
        target.update(crossing_item(item["mep"], item["section"], p0, p1, frame, kind))
        target["host"] = (stack, keep["host"][1], label)
        target["host_kind"] = keep.get("host_kind")
    return merged


def clusters(items):
    """Gruppi di attraversamenti dello stesso elemento, paralleli, a distanza <= merge nel piano del foro (e della
    stessa disciplina, se MEC ed ELE sono separati)."""
    # distanza tra i fori finiti (franco compreso), non tra le sezioni nude: due fori che si toccherebbero o
    # sovrapporrebbero si uniscono sempre
    merge = (SETTINGS["merge_mm"] + 2.0 * SETTINGS["clearance_mm"]) / MM_PER_FT
    parent = list(range(len(items)))

    def root(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def rect(item, frame):
        u, v, _ = frame
        pu = [p.DotProduct(u) for p in item["points"]]
        pv = [p.DotProduct(v) for p in item["points"]]
        return min(pu), max(pu), min(pv), max(pv)

    for i in range(len(items)):
        for j in range(i + 1, len(items)):
            a, b = items[i], items[j]
            # stessa struttura: almeno un elemento attraversato in comune (pila di strati, muro + trave...)
            if not set(a["host"][0]) & set(b["host"][0]) or a["kind"] != b["kind"]:
                continue
            if SETTINGS["separate"] and a["discipline"] != b["discipline"]:
                continue
            if abs(a["frame"][2].DotProduct(b["frame"][2])) < 0.999:
                continue
            ra, rb = rect(a, a["frame"]), rect(b, a["frame"])
            gap = max(rb[0] - ra[1], ra[0] - rb[1], rb[2] - ra[3], ra[2] - rb[3])
            if gap <= merge:
                parent[root(j)] = root(i)
    groups = {}
    for i, item in enumerate(items):
        groups.setdefault(root(i), []).append(item)
    return list(groups.values())


def settings_line():
    return ("clearance {clearance_mm:g} mm per side, rounding {rounding_mm:g} mm, merge within {merge_mm:g} mm, "
            "extra depth {extra_depth_mm:g} mm per side, layers joined across gaps up to {layer_gap_mm:g} mm, "
            "elements under {min_angle_deg:g} degrees to the wall/floor skipped, MEC/ELE {sep}").format(
        sep="separate" if SETTINGS["separate"] else "together", **SETTINGS)


def bundle_path(name):
    return script.get_bundle_file(name) or op.join(op.dirname(__file__), name)


def split_selection():
    """Elementi gia' selezionati, divisi per destinazione:
    - fori gestiti, ed elementi MEP che rimandano ad almeno un foro esistente: check;
    - elementi MEP rettilinei senza foro (anche Not Associated), muri/solai/strutture del modello host: Place.
    Ritorna (fori, MEP associati, MEP da forare, elementi attraversati, numero di elementi non usati)."""
    managed = set(code_of(o) for o in existing_openings())
    openings, linked, meps, hosts, ignored = [], [], [], [], 0
    for element_id in uidoc.Selection.GetElementIds():
        element = doc.GetElement(element_id)
        if element is None:
            continue
        if code_of(element) in managed:
            openings.append(element)
        elif isinstance(element, MEP_CLASSES):
            if any(c in managed for c in split_list(get_text(element, ID_PARAM))):
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


def last_pick():
    try:
        return script.get_config(CONFIG_NAME).get_option("place_pick", "hosts")
    except Exception:
        return "hosts"


def remember_pick(pick):
    try:
        config = script.get_config(CONFIG_NAME)
        config.place_pick = pick
        script.save_config()
    except Exception:
        pass


def family_line():
    return "Families: wall {} / {}, floor {} / {}.".format(SETTINGS["fam_wall_rect"], SETTINGS["fam_wall_round"],
                                                           SETTINGS["fam_floor_rect"], SETTINGS["fam_floor_round"])


def missing_families_exit(missing_families):
    forms.alert("Load the opening families first (or choose others in Settings):\n{}".format(
        "\n".join(missing_families)), title=TITLE, exitscript=True)


def place_main():
    import OpeningsPlace_ui  # solo nella cartella di Place
    symbols, missing_families = family_symbols()
    meps, hosts, ignored, area = [], [], 0, None
    if uidoc.Selection.GetElementIds().Count:
        # selezione: gli elementi gia' associati a un foro vanno al check, gli altri ai fori nuovi
        openings, linked, meps, hosts, ignored = split_selection()
        if openings or linked:
            chosen = linked_openings(existing_openings(), linked, openings)
            check_main(CheckScope("mep", "selection, already associated ({} openings, {} MEP elements)".format(
                len(openings), len(linked)), set(code_of(o) for o in chosen), linked))
        if not meps and not hosts:
            if not openings and not linked:
                output.print_md("# Openings Place")
                output.print_md("Nothing to do: the **{}** selected elements are not pipes, ducts, cable trays, "
                                "conduits, openings, walls, floors or structural elements.".format(ignored))
            return
        if missing_families:
            missing_families_exit(missing_families)
    else:
        if missing_families:
            missing_families_exit(missing_families)
        pick = OpeningsPlace_ui.PickWindow(bundle_path(PICK_XAML), last_pick(),
                                           "{}\n{}".format(settings_line(), family_line())).show()
        if pick is None:
            script.exit()
        remember_pick(pick)
        if pick == OpeningsPlace_ui.PICK_MEPS:
            meps = pick_meps()
        elif pick == OpeningsPlace_ui.PICK_AREA:
            area = pick_area()
        else:
            hosts = pick_hosts()

    stats = {"inside": 0, "existing": 0, "extended": 0, "outside_area": 0, "problems": []}
    items, modes = [], []
    if area is not None:
        # ogni elemento MEP dell'area contro ogni muro, solaio e struttura dell'area (anche nei link); fori solo per
        # gli attraversamenti dentro l'area, anche se l'elemento MEP ne esce
        area_meps, area_hosts = meps_in_area(area), hosts_in_area(area)
        if area_meps and area_hosts:
            with forms.ProgressBar(title="Openings: looking for crossings ({value} of {max_value})") as bar:
                items.extend(find_crossings(area_hosts, stats, only=set(id_value(m.Id) for m in area_meps),
                                            progress=lambda i, n: bar.update_progress(i + 1, n), area=area))
        modes.append("area: {}; MEP elements {}, walls, floors and structural elements {}".format(
            area_text(area), len(area_meps), len(area_hosts)))
    if hosts:
        items.extend(find_crossings(hosts, stats))
        modes.append("walls, floors and structural elements selected: {}".format(len(hosts)))
    if meps:
        # gli elementi gia' scelti sopra hanno gia' dato tutti i loro attraversamenti
        chosen = set(h[0] for h in hosts)
        found = [h for h in hosts_from_meps(meps) if h[0] not in chosen]
        items.extend(find_crossings(found, stats, only=set(id_value(m.Id) for m in meps)))
        modes.append("from {} MEP elements: walls, floors and structural elements found {}".format(
            len(meps), len(found)))
    mode_text = "; ".join(modes)
    items = merge_overlapping(items)
    groups = clusters(items)
    output.print_md("# Openings Place")
    output.print_md("Settings: {}".format(settings_line()))
    output.print_md("Mode: **{}** | crossings: **{}** | openings to create: **{}** | already with an opening: **{}** | "
                    "MEP running inside (no opening): **{}** | extended through adjacent layers: **{}**".format(
                        mode_text, len(items), len(groups), stats["existing"], stats["inside"], stats["extended"]))
    if stats["outside_area"]:
        output.print_md("Crossings of these MEP elements outside the area (skipped): **{}**".format(
            stats["outside_area"]))
    if ignored:
        output.print_md("Selected elements not used (not MEP, opening, wall, floor or structure, or not straight): "
                        "**{}**".format(ignored))
    for problem in stats["problems"]:
        output.print_md("- {}".format(problem))
    if not groups:
        output.print_md("Nothing to create.")
        return
    preview = {"mode": mode_text, "crossings": len(items), "openings": len(groups), "existing": stats["existing"],
               "inside": stats["inside"], "extended": stats["extended"]}
    if not OpeningsPlace_ui.ConfirmWindow(bundle_path(CONFIRM_XAML), preview, stats["problems"]).show():
        script.exit()

    missing = set()
    rows = []
    mep_codes = {}  # id elemento MEP -> (elemento, codici dei fori nuovi)
    managed = set(code_of(o) for o in existing_openings())  # sui MEP si tengono solo i codici di fori esistenti
    transaction = DB.Transaction(doc, TITLE)
    transaction.Start()
    try:
        for group in groups:
            frame, kind = group[0]["frame"], group[0]["kind"]
            data = sizes(group, frame)
            meps = []
            for item in group:
                if all(m.Id != item["mep"].Id for m in meps):
                    meps.append(item["mep"])
            instance = create_opening(data, frame, kind, symbols, meps, NEW, missing)
            code = code_of(instance)
            for mep in meps:
                mep_codes.setdefault(id_value(mep.Id), (mep, []))[1].append(code)
            key, host, model = group[0]["host"]
            rows.append([output.linkify(instance.Id), disciplines_text(meps), categories_text(meps),
                         "{} {} ({})".format(host.Category.Name, id_value(host.Id), model), data["shape"],
                         size_text(data), ", ".join(str(id_value(m.Id)) for m in meps)])
        for mep, new_codes in mep_codes.values():
            old = [c for c in split_list(get_text(mep, ID_PARAM)) if c in managed]
            # con altri fori gia' presenti resta il loro stato, se peggiore
            status = worst([VALID] + ([get_text(mep, STATUS_PARAM) or VALID] if old else []))
            write_mep(mep, old + [c for c in new_codes if c not in old], status, missing)
        transaction.Commit()
    except Exception:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    output.print_table(table_data=rows, columns=["Opening", "Discipline", "Categories", "Crossed element", "Shape",
                                                 "Size [mm]", "MEP elements"])
    output.print_md("New openings and their MEP elements: {} = **{}**.".format(STATUS_PARAM, VALID))
    if missing:
        output.print_md("**Parameters not found (not written):** {}".format("; ".join(sorted(missing))))


# ----------------------------------------------------------------------------------------------------------------
# CHECK & REPORT
# ----------------------------------------------------------------------------------------------------------------
def read_length(instance, name):
    param = instance.LookupParameter(name)
    return mm(param.AsDouble()) if param is not None and param.HasValue else None


def opening_kind(instance, meps):
    """(muro|solaio, rect|round): dalla famiglia, se e' una di quelle di Settings; altrimenti dai parametri (tondo
    se ha il diametro) e dalla direzione dei suoi elementi MEP (verticali: solaio)."""
    name = instance.Symbol.Family.Name
    for key, setting in FAMILY_KEYS.items():
        if SETTINGS[setting] == name:
            return key
    shape = "round" if read_length(instance, PARAM_DIAMETER) else "rect"
    vertical = any(abs(m.Location.Curve.Direction.Z) > VERTICAL_SIN for m in meps)
    return ("floor" if vertical else "wall"), shape


def opening_frame(instance, kind, shape):
    x = instance.GetTransform().BasisX
    if kind == "wall" and shape == "round":
        return Z.CrossProduct(x), Z, x  # cilindro lungo X: X e' l'asse di attraversamento
    return (x, Z, Z.CrossProduct(x)) if kind == "wall" else (x, Z.CrossProduct(x), Z)


def opening_center(instance, kind, shape):
    """Centro del solido: il punto di inserimento, alzato di meta' altezza per le famiglie che poggiano sul livello
    di riferimento (BASE_PARAMS)."""
    point = instance.Location.Point
    base = BASE_PARAMS.get((kind, shape))
    if base:
        point = point.Add(Z.Multiply((read_length(instance, base) or 0.0) / 2.0 / MM_PER_FT))
    return point


def current_data(instance, kind, shape):
    data = {"center": opening_center(instance, kind, shape), "shape": shape,
            "depth": read_length(instance, LENGTH_PARAMS[(kind, shape)]) or 0.0}
    if shape == "round":
        data["diameter"] = read_length(instance, PARAM_DIAMETER) or 0.0
    else:
        data["width"] = read_length(instance, PARAM_WIDTH) or 0.0
        data["height"] = read_length(instance, PARAM_HEIGHT) or 0.0
    return data


def half_sizes(data):
    if data["shape"] == "round":
        return data["diameter"] / 2.0, data["diameter"] / 2.0
    return data["width"] / 2.0, data["height"] / 2.0


def recross(mep, frame, data, limit_reach=True):
    """Attraversamento dell'elemento MEP con lo spessore del foro (il foro fa da muro: piano per il centro,
    perpendicolare a w, spessore = profondita'). None se non lo attraversa piu' o (limit_reach) se e' uscito dal foro
    oltre la distanza di unione."""
    line = mep.Location.Curve
    if not isinstance(line, DB.Line):
        return None
    u, v, w = frame
    direction = line.Direction
    cos = direction.DotProduct(w)
    if abs(cos) < MIN_CROSS_SIN:
        return None
    start = line.GetEndPoint(0)
    half = data["depth"] / MM_PER_FT / 2.0
    base = data["center"].Subtract(start).DotProduct(w)
    t0, t1 = sorted(((base - half) / cos, (base + half) / cos))
    if t1 < -1e-6 or t0 > line.Length + 1e-6:
        return None  # il tratto non arriva piu' al foro
    p0, p1 = start.Add(direction.Multiply(t0)), start.Add(direction.Multiply(t1))
    mid = p0.Add(p1).Divide(2.0).Subtract(data["center"])
    reach_u, reach_v = [(h + SETTINGS["merge_mm"]) / MM_PER_FT for h in half_sizes(data)]
    if limit_reach and (abs(mid.DotProduct(u)) > reach_u or abs(mid.DotProduct(v)) > reach_v):
        return None  # spostato fuori dal foro
    return p0, p1


def inside_opening(item, frame, data):
    """La sezione dell'elemento (isolamento compreso) sta tutta dentro il foro attuale."""
    u, v, _ = frame
    if data["shape"] == "round":
        radius = data["diameter"] / 2.0 / MM_PER_FT
        return all(math.hypot(p.Subtract(data["center"]).DotProduct(u), p.Subtract(data["center"]).DotProduct(v))
                   <= radius + 1e-6 for p in item["points"])
    hu, hv = [h / MM_PER_FT for h in half_sizes(data)]
    return all(abs(p.Subtract(data["center"]).DotProduct(u)) <= hu + 1e-6 and
               abs(p.Subtract(data["center"]).DotProduct(v)) <= hv + 1e-6 for p in item["points"])


def overlaps_opening(item, frame, data):
    """La sezione dell'elemento tocca almeno in parte il foro attuale (altrimenti e' uscita del tutto, anche se e'
    ancora entro la distanza di unione)."""
    u, v, _ = frame
    center = data["center"]
    pu = [p.Subtract(center).DotProduct(u) for p in item["points"]]
    pv = [p.Subtract(center).DotProduct(v) for p in item["points"]]
    if data["shape"] == "round":
        # distanza dal centro del rettangolo che contiene la sezione
        du = 0.0 if min(pu) <= 0.0 <= max(pu) else min(abs(min(pu)), abs(max(pu)))
        dv = 0.0 if min(pv) <= 0.0 <= max(pv) else min(abs(min(pv)), abs(max(pv)))
        return math.hypot(du, dv) < data["diameter"] / 2.0 / MM_PER_FT
    hu, hv = [h / MM_PER_FT for h in half_sizes(data)]
    return min(pu) < hu and max(pu) > -hu and min(pv) < hv and max(pv) > -hv


def differences(old, new):
    notes = []
    if old["shape"] != new["shape"]:
        notes.append("shape {} -> {}".format(old["shape"], new["shape"]))
    moved = mm(old["center"].DistanceTo(new["center"]))
    if moved > MOVE_TOL_MM:
        notes.append("moved {:.0f} mm".format(moved))
    if old["shape"] == new["shape"]:
        keys = ("diameter",) if new["shape"] == "round" else ("width", "height")
        for key in keys:
            if abs(old[key] - new[key]) > SIZE_TOL_MM:
                notes.append("{} {:.0f} -> {:.0f}".format(key, old[key], new[key]))
    if abs(old["depth"] - new["depth"]) > SIZE_TOL_MM:
        notes.append("depth {:.0f} -> {:.0f} (walls / floors crossed changed)".format(old["depth"], new["depth"]))
    return notes


def scanned_plane(meps, frame, old, scanned):
    """Piano del foro (centro, profondita' in mm) dagli attraversamenti della scansione del check: per i suoi
    elementi, quelli paralleli al foro che ne toccano lo spessore, gia' allungati da find_crossings attraverso gli
    strati entro l'intercapedine di Settings (es. un muro nuovo accanto). Senza dati: il piano attuale."""
    if not scanned:
        return old["center"], old["depth"]
    w = frame[2]
    half = old["depth"] / MM_PER_FT / 2.0
    tol = MOVE_TOL_MM / MM_PER_FT
    lows, highs = [], []
    for mep in meps:
        for item in scanned.get(id_value(mep.Id), []):
            if abs(item["frame"][2].DotProduct(w)) < 0.999:
                continue
            a, b = sorted((item["p0"].Subtract(old["center"]).DotProduct(w),
                           item["p1"].Subtract(old["center"]).DotProduct(w)))
            if b < -half - tol or a > half + tol:
                continue  # un altro muro o solaio: e' un foro mancante, non questo
            lows.append(a)
            highs.append(b)
    if not lows:
        return old["center"], old["depth"]
    extra = SETTINGS["extra_depth_mm"] / MM_PER_FT
    low, high = min(lows) - extra, max(highs) + extra
    return old["center"].Add(w.Multiply((low + high) / 2.0)), mm(high - low)


def rebuilt(items, frame, center, depth):
    """Foro ricalcolato dagli attraversamenti: misure nel piano dagli elementi, centro sul piano per center
    (perpendicolare a w), profondita' data (quella attuale, o quella della scansione: scanned_plane)."""
    new = sizes(items, frame)
    new["center"] = new["center"].Subtract(frame[2].Multiply(new["center"].Subtract(center).DotProduct(frame[2])))
    new["depth"] = depth
    return new


def mep_name(mep):
    for cls, label in MEP_SINGULAR:
        if isinstance(mep, cls):
            return label
    return "Element"


def check_opening(instance, extra=(), scanned=None):
    """Confronto di un foro con i suoi elementi MEP. Ritorna il dict del risultato; "elements" ha una riga per ogni
    Id elencato nel foro, con il suo stato (ELEMENT_*). extra: elementi MEP che lo attraversano senza esservi
    associati (scansione del check): se lo attraversano davvero entrano nel foro ("added"). scanned: {id elemento:
    attraversamenti della scansione}, per ricalcolare la profondita' (scanned_plane)."""
    listed = split_list(get_text(instance, ID_PARAM))
    found = []
    result = {"instance": instance, "code": code_of(instance), "keep": [], "notes": [], "state": OK, "new": None,
              "elements": [], "stale": []}

    def element_row(text, mep, state):
        result["elements"].append({"id": text, "element": mep, "name": mep_name(mep) if mep is not None else
                                   "Element", "state": state})

    for text in listed:
        element = element_by_id(text)
        if element is None or not isinstance(element, MEP_CLASSES):
            result["notes"].append("{} no longer exists".format(text))
            result["stale"].append(text)  # Id vecchio: tolto dal foro alla scrittura
            element_row(text, None, ELEMENT_DELETED)
        else:
            found.append(element)
    kind, shape = opening_kind(instance, found)
    frame = opening_frame(instance, kind, shape)
    old = current_data(instance, kind, shape)
    result.update({"kind": kind, "frame": frame, "old": old, "plane": (old["center"], old["depth"])})
    items, all_inside = [], True
    result["misplaced"] = []  # (elemento, attraversamento): fuori dal foro, il foro si puo' spostare su di lui
    listed_set = set(listed)
    # elementi che attraversano il foro senza esservi associati (trovati dalla scansione del check): si associano
    extra = [m for m in extra if str(id_value(m.Id)) not in listed_set]
    result["added"] = []
    for mep in found + extra:
        text = str(id_value(mep.Id))
        if text not in listed_set:
            # non era associato: entra solo se attraversa davvero il foro
            try:
                section = mep_section(mep)
            except Exception:
                continue
            crossing = recross(mep, frame, old)
            if crossing is None:
                continue
            item = crossing_item(mep, section, crossing[0], crossing[1], frame, kind)
            inside = inside_opening(item, frame, old)
            if not inside:
                all_inside = False
                result["misplaced"].append((mep, item))
            result["notes"].append("{} crosses the opening{} but was not associated: added".format(
                text, "" if inside else ", partly outside,"))
            element_row(text, mep, ELEMENT_ADDED if inside else ELEMENT_ADDED_OUT)
            result["added"].append(mep)
            items.append(item)
            result["keep"].append(mep)
            continue
        try:
            section = mep_section(mep)
        except Exception:
            result["notes"].append("{}: section not readable".format(text))
            element_row(text, mep, ELEMENT_UNREADABLE)
            continue
        crossing = recross(mep, frame, old)
        if crossing is None:
            far = recross(mep, frame, old, limit_reach=False)
            if far is None:
                result["notes"].append("{} no longer crosses the opening".format(text))
                element_row(text, mep, ELEMENT_NOT_CROSSING)
            else:
                # ancora parallelo all'attraversamento ma spostato lontano: il foro si puo' spostare su di lui
                result["notes"].append("{} moved out of the opening".format(text))
                result["misplaced"].append((mep, crossing_item(mep, section, far[0], far[1], frame, kind)))
                element_row(text, mep, ELEMENT_MOVED_OUT)
            continue
        item = crossing_item(mep, section, crossing[0], crossing[1], frame, kind)
        if not inside_opening(item, frame, old):
            all_inside = False
            # entro la distanza di unione ma fuori del tutto: uscito, non "in parte fuori"
            gone = not overlaps_opening(item, frame, old)
            result["notes"].append("{} {}".format(text, "moved out of the opening" if gone else
                                                  "partly outside the opening"))
            result["misplaced"].append((mep, item))
            element_row(text, mep, ELEMENT_MOVED_OUT if gone else ELEMENT_PARTLY_OUT)
        else:
            element_row(text, mep, ELEMENT_INSIDE)
        items.append(item)
        result["keep"].append(mep)
    result["found_meps"] = found + result["added"]  # restano associati al foro finche' non lo si ricrea
    result["found"] = len(result["found_meps"])  # per spostare il foro: foro di un solo elemento
    result["items"] = items
    result["groups"] = [items]  # fori che ne prendono il posto, se ricreato (due se MEC ed ELE vanno separati)
    if not items:
        result["state"] = OBSOLETE
        return result
    # piano e profondita': dalla scansione del check se c'e' (strati nuovi o tolti), altrimenti quelli attuali
    result["plane"] = scanned_plane(result["keep"], frame, old, scanned)
    new = rebuilt(items, frame, result["plane"][0], result["plane"][1])
    result["new"] = new
    # elementi elencati spariti, non piu' attraversanti o illeggibili (quelli aggiunti non contano)
    lost = len([m for m in result["keep"] if str(id_value(m.Id)) in listed_set]) < len(listed)
    diffs = differences(old, new)
    result["notes"].extend(diffs)
    if lost or not all_inside:
        # elementi ancora dentro o in parte fuori: To Be Refreshed (se attivo in Settings); Not Associated solo se
        # non ne resta nessuno
        present = (ELEMENT_INSIDE, ELEMENT_PARTLY_OUT, ELEMENT_ADDED, ELEMENT_ADDED_OUT)
        kept = [e for e in result["elements"] if e["state"] in present]
        result["state"] = PARTIAL if kept and SETTINGS["misaligned"] else CHANGED
    elif diffs:
        result["state"] = MISALIGNED if SETTINGS["misaligned"] else CHANGED
    return result


def opening_on_hosts(instance, host_solids):
    """Il foro sta su uno dei muri/solai scelti: una linea per il suo centro, lungo X, Y e Z della famiglia (gli
    assi di attraversamento: Y muro rettangolare, X muro circolare, Z solaio), entra nel solido dell'elemento."""
    transform = instance.GetTransform()
    # centro del bounding box = centro del solido, anche per le famiglie che poggiano sul livello di riferimento
    box = instance.get_BoundingBox(None)
    center = box.Min.Add(box.Max).Divide(2.0) if box is not None else instance.Location.Point
    lengths = [read_length(instance, name) or 0.0 for name in set(LENGTH_PARAMS.values())]
    reach = max(max(lengths) / MM_PER_FT, 1.0)
    options = DB.SolidCurveIntersectionOptions()
    for axis in (transform.BasisX, transform.BasisY, transform.BasisZ):
        line = DB.Line.CreateBound(center.Subtract(axis.Multiply(reach)), center.Add(axis.Multiply(reach)))
        for solid in host_solids:
            try:
                if solid.IntersectWithCurve(line, options).SegmentCount > 0:
                    return True
            except Exception:
                continue
    return False


def linked_openings(openings, sel_meps, sel_openings):
    """Fori legati agli elementi scelti: i fori stessi, quelli elencati sui MEP e quelli che elencano i MEP."""
    mep_ids = set(str(id_value(m.Id)) for m in sel_meps)
    codes = set(c for m in sel_meps for c in split_list(get_text(m, ID_PARAM)))
    wanted = set(id_value(o.Id) for o in sel_openings)
    return [o for o in openings if id_value(o.Id) in wanted or code_of(o) in codes or
            mep_ids & set(split_list(get_text(o, ID_PARAM)))]


class CheckScope(object):
    """Fori del check: tutti (ids None) o un insieme di Id di fori, aggiornato dopo ricreazioni. meps: elementi MEP
    scelti, le cui liste si riscrivono anche se non hanno piu' fori. removed: fori cancellati dalle ricreazioni: i MEP
    che li citano ancora si riscrivono anche se il foro non e' piu' nell'ambito. area (make_area): ambito Area, a
    ogni controllo entrano anche i fori e gli elementi MEP comparsi nell'area nel frattempo."""

    def __init__(self, kind, description, ids=None, meps=None, released=None, area=None):
        self.kind, self.description, self.ids, self.meps = kind, description, ids, list(meps or [])
        self.area = area
        self.removed = set()
        # {Id elemento (testo): [codici dei fori cancellati che lo elencavano]}: elementi rimasti senza foro dopo un
        # Delete, tenuti nell'ambito e nel report finche' un foro nuovo non li circonda
        self.released = dict(released or {})

    def release(self, element, codes):
        text = str(id_value(element.Id))
        self.released[text] = sorted(set(self.released.get(text, [])) | set(codes), key=int)
        if all(m.Id != element.Id for m in self.meps):
            self.meps.append(element)

    def add(self, codes):
        """Fori entrati nell'ambito (attraversati da elementi dell'ambito non associati)."""
        if self.ids is not None:
            self.ids |= set(codes)

    def openings(self):
        found = existing_openings()
        if self.ids is None:
            return found
        return [o for o in found if code_of(o) in self.ids or (self.area is not None and opening_in_area(self.area, o))]

    def replace(self, removed, created):
        self.removed |= set(removed)
        if self.ids is not None:
            self.ids = (self.ids - set(removed)) | set(created)


def selection_for_check():
    """Elementi gia' selezionati: (elementi MEP, fori gestiti, muri/solai/strutture del modello host)."""
    managed = set(code_of(o) for o in existing_openings())
    meps, openings, hosts = [], [], []
    for element_id in uidoc.Selection.GetElementIds():
        element = doc.GetElement(element_id)
        if element is None:
            continue
        if code_of(element) in managed:
            openings.append(element)
        elif isinstance(element, MEP_CLASSES):
            meps.append(element)
        elif element.Category is not None and id_value(element.Category.Id) in HOST_KINDS:
            hosts.append(((0, id_value(element.Id)), element, None, "Host model"))
    return meps, openings, hosts


def openings_on_hosts(openings, hosts):
    solids = []
    for key, element, transform, model in hosts:
        for solid in solids_of(element):
            solids.append(DB.SolidUtils.CreateTransformed(solid, transform) if transform is not None else solid)
    return [o for o in openings if opening_on_hosts(o, solids)]


def last_scope_kind(default):
    try:
        return script.get_config(CONFIG_NAME).get_option("check_scope", default)
    except Exception:
        return default


def remember_scope_kind(kind):
    try:
        config = script.get_config(CONFIG_NAME)
        config.check_scope = kind
        script.save_config()
    except Exception:
        pass


def ask_check_scope(ui):
    """Finestra dell'ambito del New Report: tutto il modello, selezione MEP (fori associati agli elementi),
    selezione ARC (fori su muri, solai e strutture). Usa la selezione corrente se c'e', altrimenti si seleziona
    dopo. Ritorna il CheckScope."""
    sel_meps, sel_openings, sel_hosts = selection_for_check()
    kind = ui.ScopeWindow(bundle_path(SCOPE_XAML), last_scope_kind(ui.SCOPE_ALL), len(sel_meps) + len(sel_openings),
                          len(sel_hosts)).show()
    if kind is None:
        script.exit()
    remember_scope_kind(kind)
    openings = existing_openings()
    if kind == ui.SCOPE_MEP:
        if not sel_meps and not sel_openings:
            sel_meps = pick_meps()
        chosen = linked_openings(openings, sel_meps, sel_openings)
        description = "MEP selection ({} MEP elements, {} openings)".format(len(sel_meps), len(sel_openings))
        return CheckScope(kind, description, set(code_of(o) for o in chosen), sel_meps)
    if kind == ui.SCOPE_ARC:
        hosts = sel_hosts or pick_hosts()
        chosen = openings_on_hosts(openings, hosts)
        description = "ARC selection ({} walls, floors and structural elements)".format(len(hosts))
        return CheckScope(kind, description, set(code_of(o) for o in chosen))
    if kind == ui.SCOPE_AREA:
        area = pick_area()
        chosen = [o for o in openings if opening_in_area(area, o)]
        return CheckScope(kind, "Area: {}".format(area_text(area)), set(code_of(o) for o in chosen),
                          meps_in_area(area), area=area)
    return CheckScope(kind, "whole model")


def discipline_checks(results):
    """Divisioni e unioni dei fori secondo le impostazioni attuali di Settings. Gli elementi sono ancora dentro: lo
    stato e' almeno REGROUP (To Be Refreshed), resta Changed (Not Associated) o Partial se lo era gia'.
    - MEC ed ELE separati: un foro con entrambe le discipline, ricreato, diventa due fori.
    - Unione: due o piu' fori sullo stesso piano i cui fori ricalcolati (franco compreso) distano al massimo la
      distanza di unione, ricreati, diventano un foro solo. Vale per elementi avvicinati e per la spunta MEC/ELE
      tolta; con MEC ed ELE separati si uniscono solo fori della stessa disciplina.
    I fori Changed e Partial (elementi tolti o usciti) non entrano nelle unioni: altrimenti un elemento spostato vicino a un
    foro in ordine lo renderebbe Not Associated. Ricreato il foro Changed, il check dopo propone l'unione. Cosi' i
    gruppi hanno solo fori To Be Refreshed e si ricreano tutti insieme o per niente."""
    live = [r for r in results if r["state"] != OBSOLETE and r["items"]]
    if SETTINGS["separate"]:
        for r in live:
            split = [[i for i in r["items"] if i["discipline"] == d] for d in ("MEC", "ELE")]
            split = [g for g in split if g]
            if len(split) > 1:
                r["groups"] = split
                r["state"] = max(r["state"], REGROUP, key=STATE_RANK.get)
                r["notes"].append("MEC and ELE must be separate (Settings): it becomes two openings")
    merge = SETTINGS["merge_mm"] / MM_PER_FT
    candidates = [r for r in live if len(r["groups"]) == 1 and r["state"] not in (CHANGED, PARTIAL)]
    parent = dict((id(r), id(r)) for r in candidates)

    def root(key):
        while parent[key] != key:
            key = parent[key]
        return key

    def disciplines(r):
        return set(i["discipline"] for i in r["items"])

    for i in range(len(candidates)):
        for j in range(i + 1, len(candidates)):
            a, b = candidates[i], candidates[j]
            if SETTINGS["separate"] and (len(disciplines(a)) > 1 or disciplines(a) != disciplines(b)):
                continue
            u, v, w = a["frame"]
            if abs(w.DotProduct(b["frame"][2])) < 0.999:
                continue
            if mm(abs(b["old"]["center"].Subtract(a["old"]["center"]).DotProduct(w))) > MOVE_TOL_MM:
                continue  # non sullo stesso piano: muri o solai diversi
            # fori come andrebbero ricalcolati (misure con il franco): si uniscono se distano <= unione
            d = b["new"]["center"].Subtract(a["new"]["center"])
            ha, hb = [[h / MM_PER_FT for h in half_sizes(r["new"])] for r in (a, b)]
            gap = max(abs(d.DotProduct(u)) - ha[0] - hb[0], abs(d.DotProduct(v)) - ha[1] - hb[1])
            if gap <= merge:
                parent[root(id(b))] = root(id(a))
    groups = {}
    for r in candidates:
        groups.setdefault(root(id(r)), []).append(r)
    for members in groups.values():
        if len(members) < 2:
            continue
        codes = ", ".join(m["code"] for m in members)
        state = max([REGROUP] + [m["state"] for m in members], key=STATE_RANK.get)
        for m in members:
            m["merge_group"] = members
            m["state"] = state
            m["notes"].append("close enough to be one opening: openings {} become one".format(codes))


def opening_shift(items, frame, data, single):
    """Spostamento del foro (nel suo piano, senza cambiarne le misure) che lo riporta sui suoi elementi: centrato su
    di essi se il foro e' di un solo elemento, tondo, o se gli elementi non ci stanno comunque; altrimenti il minimo
    perche' tutti rientrino con il franco."""
    u, v, _ = frame
    center = data["center"]
    points = [p for item in items for p in item["points"]]
    pu = [p.Subtract(center).DotProduct(u) for p in points]
    pv = [p.Subtract(center).DotProduct(v) for p in points]
    if single or data["shape"] == "round":
        return u.Multiply((max(pu) + min(pu)) / 2.0).Add(v.Multiply((max(pv) + min(pv)) / 2.0))
    clear = SETTINGS["clearance_mm"] / MM_PER_FT
    shifts = []
    for values, half in ((pu, data["width"] / 2.0 / MM_PER_FT), (pv, data["height"] / 2.0 / MM_PER_FT)):
        room = max(half - clear, 0.0)
        low, high = min(values), max(values)
        if high - low > 2.0 * room:
            shifts.append((high + low) / 2.0)
        elif high > room:
            shifts.append(high - room)
        elif low < -room:
            shifts.append(low + room)
        else:
            shifts.append(0.0)
    return u.Multiply(shifts[0]).Add(v.Multiply(shifts[1]))


def alignment_vector(r):
    """Spostamento (nel piano del foro, misure invariate) che riporta il foro sui suoi elementi usciti in parte o del
    tutto; None se nessun elemento e' fuori o lo spostamento e' sotto 1 mm."""
    if not r.get("misplaced"):
        return None
    items = list(r["items"]) + [item for _, item in r["misplaced"] if item not in r["items"]]
    vector = opening_shift(items, r["frame"], r["old"], r["found"] <= 1)
    return vector if vector.GetLength() * MM_PER_FT >= 1.0 else None


def unique_meps(items):
    meps = []
    for item in items:
        if all(m.Id != item["mep"].Id for m in meps):
            meps.append(item["mep"])
    return meps


def group_discipline(items):
    return disciplines_text(unique_meps(items)) or "-"


def outside_meps(r):
    """Elementi del foro che esistono ancora ma non attraversano piu' il foro (nemmeno entro la portata): Recreate non
    li include nel foro nuovo e perdono l'associazione, ma restano nell'ambito (CheckScope.release): la scansione li
    mostra in Missing openings finche' un foro nuovo non li circonda."""
    return [m for m in r["found_meps"] if all(m.Id != k.Id for k in r["keep"])]


def recreate_plan(r):
    """Cosa fa Recreate su questo foro, in una riga (dettaglio del report e conferma)."""
    if not r["items"]:
        return "Recreate: not possible, no element crosses it any more."
    lost = len(outside_meps(r))
    lost_text = "; {} element(s) outside it lose the association and stay listed in Missing openings where they " \
                "cross a wall or floor".format(lost) if lost else ""
    if r.get("merge_group"):
        others = [m["code"] for m in r["merge_group"] if m is not r]
        return "Recreate: joined with opening(s) {} into one opening{}.".format(", ".join(others), lost_text)
    if len(r["groups"]) > 1:
        return "Recreate: split into {} openings ({}){}.".format(
            len(r["groups"]), " and ".join(group_discipline(g) for g in r["groups"]), lost_text)
    return "Recreate: one new opening on {} element(s){}.".format(len(unique_meps(r["items"])), lost_text)


def recreate_openings(results, keys, symbols, missing):
    """Ricrea i fori con Id in keys sugli elementi che li attraversano ancora (un gruppo da unire si ricrea tutto
    insieme, un foro misto da dividere diventa due fori). Chiamare dentro una transazione. Ritorna (codici creati,
    codici tolti, codici saltati perche' senza elementi, eventi in testo, {codice nuovo: origine})."""
    created, removed, skipped, events, origins, done = [], [], [], [], {}, set()
    for r in results:
        if r["code"] not in keys:
            continue
        if not r["items"]:
            skipped.append(r["code"])
            continue
        if r.get("merge_group"):
            key = id(r["merge_group"])
            if key in done:
                continue
            done.add(key)
            members = r["merge_group"]
            codes = [m["code"] for m in members]
            items = [i for m in members for i in m["items"]]
            data = rebuilt(items, r["frame"], r["plane"][0], max(m["plane"][1] for m in members))
            for m in members:
                removed.append(m["code"])
                doc.Delete(m["instance"].Id)
            new = code_of(create_opening(data, r["frame"], r["kind"], symbols, unique_meps(items), NEW, missing))
            created.append(new)
            origins[new] = "joined from {}".format(" + ".join(codes))
            ticked = [c for c in codes if c in keys]
            events.append("Openings {} joined into {}{}".format(" + ".join(codes), new, "" if len(ticked) == len(codes)
                                                                else " (the whole group, also the ones not ticked)"))
            continue
        removed.append(r["code"])
        doc.Delete(r["instance"].Id)
        news = []
        for group in r["groups"]:
            data = rebuilt(group, r["frame"], r["plane"][0], r["plane"][1])
            new = code_of(create_opening(data, r["frame"], r["kind"], symbols, unique_meps(group), NEW, missing))
            created.append(new)
            news.append(new)
            origins[new] = "split from {} ({})".format(r["code"], group_discipline(group)) if len(r["groups"]) > 1 \
                else "recreated from {}".format(r["code"])
        if len(news) > 1:
            events.append("Opening {} split into {}".format(r["code"], " and ".join(
                "{} ({})".format(n, group_discipline(g)) for n, g in zip(news, r["groups"]))))
        else:
            events.append("Opening {} recreated as {}".format(r["code"], news[0]))
    return created, removed, skipped, events, origins


def mep_candidates(reverse):
    """Elementi MEP da tenere allineati: con stato o Id compilati, o citati da un foro."""
    categories = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory](MEP_CATEGORIES))
    found = {}
    for element in DB.FilteredElementCollector(doc).WherePasses(categories).WhereElementIsNotElementType():
        if not isinstance(element, MEP_CLASSES):
            continue
        text = str(id_value(element.Id))
        if text in reverse or get_text(element, STATUS_PARAM) or get_text(element, ID_PARAM):
            found[text] = element
    return found


def plan_writes(results, scope):
    """Cosa scrivere dopo il check, e gli Id vecchi trovati.
    - Fori controllati: stato, Id dei loro elementi ancora esistenti (anche se usciti dal foro: restano associati
      finche' il foro non si ricrea), categorie. Gli Id di elementi cancellati (o non piu' MEP) si tolgono.
    - Elementi MEP: Id dei fori che li citano (la lista e' lo specchio dei fori), categoria, stato peggiore tra
      quei fori; nessun foro: Not Associated. Fuori dall'ambito solo i MEP legati ai fori controllati, a quelli
      cancellati dalle ricreazioni, o scelti."""
    checked = dict((r["code"], r) for r in results)
    lists, statuses = {}, {}
    for opening in existing_openings():
        code = code_of(opening)
        if code in checked:
            r = checked[code]
            lists[code] = [str(id_value(m.Id)) for m in r["found_meps"]]
            statuses[code] = WRITTEN[r["state"]]
        else:
            lists[code] = split_list(get_text(opening, ID_PARAM))
            statuses[code] = get_text(opening, STATUS_PARAM) or VALID
    reverse = {}
    for code, ids in lists.items():
        for text in ids:
            reverse.setdefault(text, []).append(code)
    chosen = set(str(id_value(m.Id)) for m in scope.meps)
    relevant = set(checked) | scope.removed
    meps, stale_meps, added, updated, orphaned = [], [], 0, 0, 0
    for text, mep in mep_candidates(reverse).items():
        current = split_list(get_text(mep, ID_PARAM))
        codes = reverse.get(text, [])
        if scope.ids is not None and text not in chosen and not (set(codes) | set(current)) & relevant:
            continue
        status = worst([statuses[c] for c in codes]) if codes else NOT_ASSOCIATED
        removed = [c for c in current if c not in codes]
        added += len([c for c in codes if c not in current])
        if removed:
            stale_meps.append((mep, removed))
        if sorted(codes) != sorted(current):
            updated += 1
            if current and not codes:
                orphaned += 1
        meps.append((mep, codes, status))
    return {"meps": meps, "stale_meps": stale_meps, "added": added, "updated": updated, "orphaned": orphaned,
            "stale_openings": sum(len(r["stale"]) for r in results)}


# ----------------------------------------------------------------------------------------------------------------
# Ultimo check (Last Report): stati dal modello, dettaglio per elemento dall'ultimo check fatto su questo computer
# ----------------------------------------------------------------------------------------------------------------
LAST_CHECK_FILE = "ESA_Openings_LastCheck"
ELEMENT_UNCHECKED = "not checked yet"


def last_check_path():
    try:
        return script.get_document_data_file(LAST_CHECK_FILE, "json")
    except Exception:
        return None


def read_last_file():
    """Contenuto del file dell'ultimo report; vuoto se non c'e' o non si legge."""
    path = last_check_path()
    if not path or not op.isfile(path):
        return {}
    try:
        with open(path) as handle:
            data = json.loads(handle.read().lstrip(unichr(0xFEFF)))  # noqa: F821 (IronPython 2: BOM)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def load_last_check():
    """{codice del foro: dettaglio} dell'ultimo controllo di ogni foro."""
    return read_last_file().get("openings", {})


def load_last_report():
    """Ultimo New Report (data, ambito, fori mancanti), o None se non ce n'e'."""
    report = read_last_file().get("report")
    return report if isinstance(report, dict) and report.get("scope") else None


def save_last_check(rows, removed, scope=None, missing_rows=None):
    """Aggiorna il file dell'ultimo report: i fori controllati si sovrascrivono, gli altri restano, quelli
    cancellati si tolgono; con scope, anche l'ultimo report (ambito e fori mancanti) per Last Report. Errori
    silenziosi: il file serve solo a Last Report."""
    path = last_check_path()
    if not path:
        return
    data = read_last_file()
    openings = data.get("openings", {})
    for code in removed:
        openings.pop(code, None)
    stamp = datetime.now().strftime("%d.%m.%Y %H:%M")
    for row in rows:
        openings[row["code"]] = {"status": row["status"], "no_element": row["no_element"], "notes": row["notes"],
                                 "move": row["can_move"], "origin": row["origin"], "date": stamp,
                                 "elements": [[e["id"], e["state"]] for e in row["elements"]]}
    data.update({"format": "ESA_Openings_LastCheck", "openings": openings})
    if scope is not None:
        data["report"] = {
            "date": stamp,
            "scope": {"kind": scope.kind, "description": scope.description,
                      "ids": sorted(scope.ids) if scope.ids is not None else None,
                      "meps": [id_value(m.Id) for m in scope.meps if m.IsValidObject],
                      "released": scope.released, "area": scope.area},
            "missing": [dict((key, row[key]) for key in MISSING_SAVED) for row in (missing_rows or [])]}
    try:
        temp = path + ".tmp"
        with open(temp, "w") as handle:
            handle.write(json.dumps(data))
        if op.isfile(path):
            os.remove(path)
        os.rename(temp, path)
    except Exception:
        pass


# ----------------------------------------------------------------------------------------------------------------
# Righe del report (comuni a New Report e Last Report)
# ----------------------------------------------------------------------------------------------------------------
def check_row(r, origins):
    instance = r["instance"]
    return {"code": r["code"], "instance": instance, "status": WRITTEN[r["state"]], "no_element": r["state"] == OBSOLETE,
            "family": instance.Symbol.Family.Name, "size": size_text(r["old"]),
            "notes": list(r["notes"]) + ([recreate_plan(r)] if r["state"] not in (OK, NEW) else []),
            "elements": r["elements"], "can_recreate": bool(r["items"]) and r["state"] not in (OK, NEW),
            "can_move": alignment_vector(r) is not None, "origin": origins.get(r["code"], "")}


def stored_row(instance, saved):
    """Riga del Last Report: stato e Id dal modello, dettaglio dall'ultimo check su questo computer (se c'e')."""
    code = code_of(instance)
    status = get_text(instance, STATUS_PARAM) or VALID
    info = saved.get(code)
    states = dict((e[0], e[1]) for e in info.get("elements", [])) if info else {}
    elements, found = [], []
    for text in split_list(get_text(instance, ID_PARAM)):
        element = element_by_id(text)
        if element is None or not isinstance(element, MEP_CLASSES):
            elements.append({"id": text, "element": None, "name": "Element", "state": ELEMENT_DELETED})
        else:
            found.append(element)
            elements.append({"id": text, "element": element, "name": mep_name(element),
                             "state": states.get(text, ELEMENT_UNCHECKED)})
    kind, shape = opening_kind(instance, found)
    if info:
        header = "Last check on this computer: {}.".format(info.get("date", "?"))
        if info.get("status") != status:
            header += " The status in the model ({}) differs from that check ({}): checked again by someone " \
                      "else, or changed since.".format(status, info.get("status"))
        notes = [header] + list(info.get("notes", []))
    else:
        notes = ["Not checked on this computer yet: run New Report for the detail of the elements."]
    return {"code": code, "instance": instance, "status": status,
            "no_element": bool(info.get("no_element")) if info else not found,
            "family": instance.Symbol.Family.Name, "size": size_text(current_data(instance, kind, shape)),
            "notes": notes, "elements": elements, "can_recreate": status != VALID and bool(found),
            "can_move": bool(info.get("move")) if info else False, "origin": info.get("origin", "") if info else ""}


def row_label(row):
    return "Opening {} | {} | {}{}".format(row["code"], row["family"], row["size"],
                                          " | " + row["origin"] if row["origin"] else "")


def missing_text(missing):
    """Parametri non trovati, raggruppati: "e_GEN_Void_Status (Pipes, Ducts); ..."; vuoto se non ne mancano."""
    by_param = {}
    for text in missing:
        name, _, where = text.partition(" on ")
        by_param.setdefault(name, set()).add(where)
    return "; ".join("{} ({})".format(name, ", ".join(sorted(where))) for name, where in sorted(by_param.items()))


def status_totals(report):
    """Righe del riepilogo; None = linea di separazione."""
    rows = report.rows

    def count(status):
        return len([r for r in rows if r["status"] == status])

    totals = [("Scope", report.scope.description), ("Openings", len(rows)), None]
    if report.missing_rows is None:
        totals.append((MISSING, "not searched (New Report only)"))
    else:
        totals.append((MISSING, "{} ({} crossings without an opening)".format(
            len(report.missing_rows), sum(m["crossings"] for m in report.missing_rows))))
    totals += [(NOT_ASSOCIATED, "{} ({} with no element left)".format(count(NOT_ASSOCIATED),
                                                                      len([r for r in rows if r["no_element"]]))),
               (TO_REFRESH, count(TO_REFRESH)), (VALID, count(VALID))]
    if report.added:
        totals.append(("Elements added to openings", "{} (crossing an opening, not associated before)".format(
            report.added)))
    return totals


def owners_of(report):
    """{Id elemento (testo): [(codice, stato dell'elemento, Id del foro)]} dei fori del report che lo elencano."""
    owners = {}
    for row in report.rows:
        for e in row["elements"]:
            owners.setdefault(e["id"], []).append((row["code"], e["state"], id_value(row["instance"].Id)))
    return owners


def missing_status(mep, owners, released):
    """Stato di un elemento in una riga Missing openings: se e' ancora elencato da fori presenti a modello, il suo stato
    in quei fori (es. moved out) e i loro Id; se i fori sono stati cancellati, i loro Id; altrimenti no opening."""
    text = str(id_value(mep.Id))
    links = owners.get(text, [])
    if links:
        states = []
        for _, state, _ in links:
            if state not in states:
                states.append(state)
        return "{} (opening {})".format(" / ".join(states), ", ".join(code for code, _, _ in links))
    if released.get(text):
        return "{} (removed from opening {})".format(ELEMENT_NO_OPENING, ", ".join(released[text]))
    return ELEMENT_NO_OPENING


def report_groups(ui, report):
    groups = []
    owners = owners_of(report)
    if report.missing_rows:
        groups.append((MISSING, [missing_entry(ui, row, owners, report.scope.released)
                                 for row in report.missing_rows]))
    known = (NOT_ASSOCIATED, TO_REFRESH, VALID)
    for status in known:
        groups.append((status, [row_entry(ui, row, owners) for row in report.rows if row["status"] == status]))
    other = [row_entry(ui, row, owners) for row in report.rows if row["status"] not in known]
    if other:
        groups.append(("Other status", other))
    stale = []
    for mep, removed in report.stale_meps:
        text = "{} {} | removed Ids: {}".format(mep_name(mep), id_value(mep.Id), ", ".join(removed))
        stale.append(ui.OpeningEntry("mep" + str(id_value(mep.Id)), text, text + "\n\nThese Ids pointed to openings "
                                     "that are deleted or no longer list this element.", [id_value(mep.Id)], [],
                                     checkable=False))
    if stale:
        groups.append(("Stale Ids removed from MEP elements", stale))
    return groups


def row_entry(ui, row, owners):
    """owners: owners_of(report). Selezionare un elemento seleziona anche i fori a cui e' associato."""
    select = [id_value(row["instance"].Id)]
    elements = []
    for e in row["elements"]:
        ids = [id_value(e["element"].Id)] if e["element"] is not None else []
        select.extend(ids)
        elements.append(("{} {} - {}".format(e["name"], e["id"], e["state"]),
                         ids + [i for _, _, i in owners.get(e["id"], []) if ids]))
    label = row_label(row)
    detail = "{}\nStatus: {}\n\n{}".format(label, row["status"], "\n".join(row["notes"]) or "Elements in place.")
    return ui.OpeningEntry(row["code"], label, detail, select, elements, can_recreate=row["can_recreate"],
                           can_move=row["can_move"], can_delete=row["status"] == NOT_ASSOCIATED)


def link(element, fallback):
    """Link cliccabile all'elemento; solo testo se nel frattempo e' stato cancellato (la finestra del report e'
    non modale: il modello puo' cambiare)."""
    try:
        if element is not None and element.IsValidObject:
            return output.linkify(element.Id)
    except Exception:
        pass
    return fallback


def report_data(report):
    """Dati del Text Report e dell'export Excel (solo testo: la finestra e l'Excel non leggono Revit). Sezioni per
    stato: Missing openings (solo New Report), Not Associated, To Be Refreshed, Valid (compatti), Id vecchi tolti
    dai MEP. Ogni voce: codice, titolo, descrizione, note, righe (elemento, categoria, stato dell'elemento)."""
    sections = []
    owners = owners_of(report)
    if report.missing_rows is not None:
        sections.append({"title": MISSING, "empty": "None: every crossing of the elements in this report has an "
                                                    "opening.",
                         "items": [{"code": "-", "title": row["label"],
                                    "description": "{} | {} | crossing {}".format(row["kind_text"], row["size"],
                                                                                  row["host_text"]),
                                    "notes": ["No opening at this crossing: Create in the Interactive Report."],
                                    "rows": [(str(id_value(m.Id)), mep_name(m),
                                              missing_status(m, owners, report.scope.released))
                                             for m in row["meps"]]}
                                   for row in report.missing_rows]})
    for status in (NOT_ASSOCIATED, TO_REFRESH, VALID):
        items = []
        for row in report.rows:
            if row["status"] != status:
                continue
            items.append({"code": row["code"], "title": row_label(row), "compact": status == VALID,
                          "description": "{} | {}{}".format(row["family"], row["size"],
                                                            " | " + row["origin"] if row["origin"] else ""),
                          "notes": list(row["notes"]),
                          "rows": [(e["id"], e["name"], e["state"]) for e in row["elements"]]})
        sections.append({"title": status, "items": items})
    if report.stale_meps:
        sections.append({"title": "Stale Ids removed from MEP elements",
                         "items": [{"code": "-", "title": "{} {} | removed Ids: {}".format(
                             mep_name(mep), id_value(mep.Id), ", ".join(removed)),
                             "description": "removed Ids: {}".format(", ".join(removed)),
                             "notes": ["These Ids pointed to openings that are deleted or no longer list it."],
                             "rows": []} for mep, removed in report.stale_meps if mep.IsValidObject]})
    settings = "Settings: MEC/ELE {}, To Be Refreshed {}, minimum angle {:g} deg, layers joined across {:g} mm.".format(
        "separate" if SETTINGS["separate"] else "together", "on" if SETTINGS["misaligned"] else "off",
        SETTINGS["min_angle_deg"], SETTINGS["layer_gap_mm"])
    return {"title": report.title, "model": doc.Title, "date": getattr(report, "date", "") or
            datetime.now().strftime("%d.%m.%Y %H:%M"), "summary": report.totals(), "warning": report.warning(),
            "details": [settings] + report.details(), "log": list(report.log), "sections": sections}


# ----------------------------------------------------------------------------------------------------------------
# Scansione degli attraversamenti dell'ambito (New Report): elementi non associati e fori mancanti
# ----------------------------------------------------------------------------------------------------------------
def scan_meps(scope):
    """Elementi MEP del report: quelli elencati nei fori dell'ambito, piu' quelli scelti (selezione MEP) o nell'area
    (ambito Area, riletti a ogni controllo)."""
    meps = {}
    for mep in list(scope.meps) + (meps_in_area(scope.area) if scope.area is not None else []):
        if not mep.IsValidObject:
            continue
        if is_straight_mep(mep):
            meps[id_value(mep.Id)] = mep
    for opening in scope.openings():
        for text in split_list(get_text(opening, ID_PARAM)):
            element = element_by_id(text)
            if element is not None and is_straight_mep(element):
                meps[id_value(element.Id)] = element
    return list(meps.values())


def scan_crossings(scope):
    """Attraversamenti degli elementi MEP del report (scan_meps), come li trova Place: muri, solai e strutture del
    modello e dei link toccati da quegli elementi, con il foro che contiene ogni attraversamento (item["opening"]) o
    None. Solo questi elementi, qualunque sia l'ambito: non si cercano attraversamenti di altri elementi del modello.
    Ambito Area: muri, solai e strutture dell'area (non uno per elemento MEP, che sarebbero migliaia di ricerche) e
    solo gli attraversamenti dentro l'area. Ritorna (attraversamenti, statistiche)."""
    stats = {"inside": 0, "existing": 0, "extended": 0, "outside_area": 0, "problems": []}
    meps = scan_meps(scope)
    if not meps:
        return [], stats
    only = set(id_value(m.Id) for m in meps)
    hosts = hosts_in_area(scope.area) if scope.area is not None else hosts_from_meps(meps)
    if not hosts:
        return [], stats
    with forms.ProgressBar(title="Openings: looking for crossings ({value} of {max_value})") as bar:
        items = find_crossings(hosts, stats, only=only, keep_existing=True,
                               progress=lambda i, n: bar.update_progress(i + 1, n), area=scope.area)
    return items, stats


def unlisted_crossings(items):
    """{codice del foro: {id: elemento MEP}} degli elementi che attraversano un foro senza esservi associati."""
    found = {}
    for item in items:
        opening = item.get("opening")
        if opening is None:
            continue
        mep_id = id_value(item["mep"].Id)
        if str(mep_id) not in split_list(get_text(opening, ID_PARAM)):
            found.setdefault(code_of(opening), {})[mep_id] = item["mep"]
    return found


def handled_by_opening(item, results):
    """L'attraversamento e' gia' coperto dal check di un foro dell'elemento: l'elemento e' entro la portata del foro
    (anche se uscito: Recreate o Move lo sistemano) e l'attraversamento, parallelo al foro, ne tocca lo spessore (se
    e' piu' lungo, per uno strato nuovo, il foro si allunga: scanned_plane). Non e' un foro mancante, altrimenti si
    creerebbero due fori."""
    mep = item["mep"]
    tol = MOVE_TOL_MM / MM_PER_FT
    for r in results:
        if not any(m.Id == mep.Id for m in r["found_meps"]):
            continue
        if recross(mep, r["frame"], r["old"]) is None:
            continue
        w = r["frame"][2]
        if abs(item["frame"][2].DotProduct(w)) < 0.999:
            continue
        half = r["old"]["depth"] / MM_PER_FT / 2.0
        a, b = sorted((item["p0"].Subtract(r["old"]["center"]).DotProduct(w),
                       item["p1"].Subtract(r["old"]["center"]).DotProduct(w)))
        if b >= -half - tol and a <= half + tol:
            return True
    return False


def host_label(item):
    stack, element, model = item["host"]
    return "{} {} ({})".format(element.Category.Name, id_value(element.Id), model)


# campi di una riga Missing openings salvati per il Last Report (solo testo e Id)
MISSING_SAVED = ("key", "label", "host_text", "kind_text", "size", "crossings", "mep_ids")


def missing_openings(items, results):
    """Attraversamenti senza foro, raggruppati come li creerebbe Place. Ritorna le righe del gruppo Missing
    openings. La chiave e' stabile (elementi attraversati + elementi MEP): una spunta data nel Last Report resta
    valida dopo il ricontrollo che precede Create."""
    items = [i for i in items if i.get("opening") is None and not handled_by_opening(i, results)]
    rows = []
    for group in clusters(merge_overlapping(items)):
        frame, kind = group[0]["frame"], group[0]["kind"]
        data = sizes(group, frame)
        stack, host, _ = group[0]["host"]
        meps = unique_meps(group)
        mep_ids = sorted(id_value(m.Id) for m in meps)
        row = {"key": "new:{}:{}".format("+".join("{}_{}".format(a, b) for a, b in stack),
                                         "-".join(str(i) for i in mep_ids)),
               "group": group, "frame": frame, "kind": kind, "data": data, "meps": meps, "mep_ids": mep_ids,
               "host": host if host.Document.Equals(doc) else None, "host_text": host_label(group[0]),
               "kind_text": "{} {}".format(kind, data["shape"]), "size": size_text(data), "crossings": len(group)}
        row["label"] = "New opening | {} | {} | {}".format(row["kind_text"], row["size"], row["host_text"])
        rows.append(row)
    return rows


def saved_missing_rows(saved):
    """Righe Missing openings del Last Report, dal file: elementi MEP ancora nel modello, niente geometria (Create
    passa prima da un New Report)."""
    rows = []
    for entry in saved or []:
        row = dict((key, entry.get(key)) for key in MISSING_SAVED)
        row["meps"] = [m for m in (element_by_id(i) for i in row["mep_ids"] or []) if m is not None]
        row["host"] = None
        rows.append(row)
    return rows


def missing_entry(ui, row, owners, released):
    """owners: owners_of(report); released: scope.released. Gli elementi ancora elencati da fori del modello mostrano
    lo stato in quei fori e i loro Id (selezionati con l'elemento)."""
    select = [id_value(m.Id) for m in row["meps"]]
    if row["host"] is not None:
        select.append(id_value(row["host"].Id))  # muro/solaio del modello host (quelli dei link non si selezionano)
    elements = [("{} {} - {}".format(mep_name(m), id_value(m.Id), missing_status(m, owners, released)),
                 [id_value(m.Id)] + [i for _, _, i in owners.get(str(id_value(m.Id)), [])]) for m in row["meps"]]
    detail = "{}\n\nNo opening at this crossing. Create makes it as Place would (status Valid).".format(row["label"])
    return ui.OpeningEntry(row["key"], row["label"], detail, select, elements, can_recreate=True)


# ----------------------------------------------------------------------------------------------------------------
# New Report e Last Report
# ----------------------------------------------------------------------------------------------------------------
class CheckSession(object):
    """New Report di un ambito: controlla i fori, scrive stati e Id (Id vecchi tolti), salva il dettaglio per il
    Last Report; ricrea e sposta i fori dalla finestra del report."""
    title = "New Report"

    def __init__(self, scope):
        self.scope = scope
        self.symbols, self.missing_families = family_symbols()
        self.missing = set()
        self.log = []  # azioni fatte dalla finestra del report, per il report testuale
        self.origins = {}  # codice di un foro ricreato -> da dove viene (recreated / split / joined from ...)
        self.run([])

    def run(self, removed):
        # prima gli attraversamenti degli elementi dell'ambito: elementi dentro un foro senza esservi associati
        # (entrano nel suo check, e il foro nell'ambito) e attraversamenti senza foro (Missing openings)
        items, self.scan_stats = scan_crossings(self.scope)
        unlisted = unlisted_crossings(items)
        self.scope.add(unlisted.keys())
        scanned = {}  # id elemento -> attraversamenti: profondita' dei fori (strati aggiunti o tolti)
        for item in items:
            scanned.setdefault(id_value(item["mep"].Id), []).append(item)
        self.results = sorted([check_opening(o, list(unlisted.get(code_of(o), {}).values()), scanned)
                               for o in self.scope.openings()], key=lambda r: int(r["code"]))
        discipline_checks(self.results)
        self.missing_rows = missing_openings(items, self.results)
        self.added = sum(len(r["added"]) for r in self.results)
        self.plan = plan_writes(self.results, self.scope)
        self.stale_meps = self.plan["stale_meps"]
        transaction = DB.Transaction(doc, "Openings: check")
        transaction.Start()
        try:
            for r in self.results:
                write_opening(r["instance"], r["found_meps"], r["state"], self.missing)
            for mep, codes, status in self.plan["meps"]:
                write_mep(mep, codes, status, self.missing)
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        self.rows = [check_row(r, self.origins) for r in self.results]
        self.date = datetime.now().strftime("%d.%m.%Y %H:%M")
        save_last_check(self.rows, removed, self.scope, self.missing_rows)

    def totals(self):
        return status_totals(self)

    def details(self):
        """Righe in piu' del report testuale."""
        stats = self.scan_stats
        return ["Statuses and Ids written, stale Ids removed, in one transaction (one Undo).",
                "Crossings of the MEP elements in this report: {} without an opening ({} openings to create), {} inside "
                "an opening ({} elements not associated before, now added), {} running inside a wall or floor (no "
                "opening).".format(sum(m["crossings"] for m in self.missing_rows), len(self.missing_rows),
                                   stats["existing"], self.added, stats["inside"])] + \
               (["Crossings outside the area (not checked): {}.".format(stats["outside_area"])]
                if stats.get("outside_area") else []) + \
               ["- {}".format(p) for p in stats["problems"]] + [
                "Stale Ids removed: {} on openings (elements deleted), {} on {} MEP elements (openings deleted, or not "
                "listing them). Ids added to MEP elements: {}.".format(
                    self.plan["stale_openings"], sum(len(c) for _, c in self.stale_meps), len(self.stale_meps),
                    self.plan["added"])]

    def warning(self):
        """Avviso del riepilogo, solo se qualcosa non si e' potuto scrivere."""
        if not self.missing:
            return ""
        return "Not written, parameter missing: {}. Bind these parameters to those categories in the model, " \
               "then run New Report again.".format(missing_text(self.missing))

    def recreate_preview(self, keys):
        lines = []
        for row in self.missing_rows:
            if row["key"] in keys:
                lines.append("Create a new opening on {} element(s) crossing {} ({}, {}).".format(
                    len(row["meps"]), row["host_text"], row["kind_text"], row["size"]))
        for r in self.results:
            if r["code"] in keys:
                lines.append("{}: {}".format(r["code"], recreate_plan(r)))
        return "\n".join(lines)

    def recreate(self, keys):
        """Crea i fori mancanti spuntati e ricrea i fori spuntati, in una transazione; poi ricontrolla e scrive."""
        if self.missing_families:
            return "Creating openings needs the families chosen in Settings: {}.".format(
                ", ".join(self.missing_families))
        # elementi usciti dai fori da ricreare (anche dai fori di un gruppo da unire non spuntati): perdono
        # l'associazione ma restano nell'ambito, cosi' la scansione li riporta in Missing openings
        for r in self.results:
            group = [m["code"] for m in r.get("merge_group") or []]
            if r["items"] and (r["code"] in keys or set(group) & set(keys)):
                for mep in outside_meps(r):
                    self.scope.release(mep, [r["code"]])
        transaction = DB.Transaction(doc, "Openings: create / recreate")
        transaction.Start()
        try:
            created, removed, skipped, events, origins = recreate_openings(self.results, set(keys), self.symbols,
                                                                           self.missing)
            for row in self.missing_rows:
                if row["key"] not in keys:
                    continue
                new = code_of(create_opening(row["data"], row["frame"], row["kind"], self.symbols, row["meps"],
                                             NEW, self.missing))
                created.append(new)
                origins[new] = "new"
                events.append("New opening {} on {} ({})".format(
                    new, ", ".join(str(i) for i in row["mep_ids"]), row["host_text"]))
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        for code in removed:
            self.origins.pop(code, None)
        self.origins.update(origins)
        self.scope.replace(removed, created)
        self.run(removed)  # ricontrollo e scrittura: Id dei fori nuovi sugli elementi MEP, vecchi tolti
        lines = events + ["Ids updated on {} MEP elements ({} left without an opening: {}).".format(
            self.plan["updated"], self.plan["orphaned"], NOT_ASSOCIATED)]
        if skipped:
            lines.append("Skipped, no element crossing: {}.".format(", ".join(skipped)))
        message = " ".join(line if line.endswith(".") else line + "." for line in lines)
        self.log.append(message)
        return message

    def move(self, keys):
        moves = [(r, alignment_vector(r)) for r in self.results if r["code"] in keys]
        moves = [(r, v) for r, v in moves if v is not None]
        if not moves:
            return "None of the ticked openings has elements outside it."
        transaction = DB.Transaction(doc, "Openings: move onto their elements")
        transaction.Start()
        try:
            for r, vector in moves:
                DB.ElementTransformUtils.MoveElement(doc, r["instance"].Id, vector)
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        self.run([])
        message = "Moved onto their elements: {} ({}).".format(
            len(moves), ", ".join("{} by {:.0f} mm".format(r["code"], v.GetLength() * MM_PER_FT) for r, v in moves))
        self.log.append(message)
        return message


class LastReport(object):
    """Last Report: nessun controllo. L'ultimo New Report (stesso ambito e fori mancanti, dal file), con stati e Id
    letti dal modello (scritti dall'ultimo New Report, di chiunque) e il dettaglio per elemento dell'ultimo
    controllo fatto su questo computer."""
    title = "Last Report"

    def __init__(self, saved):
        scope = saved["scope"]
        ids = scope.get("ids")
        meps = [m for m in (element_by_id(i) for i in scope.get("meps") or []) if m is not None]
        self.scope = CheckScope(scope.get("kind", "all"), scope.get("description", ""),
                                set(ids) if ids is not None else None, meps, scope.get("released"),
                                scope.get("area"))
        self.date = saved.get("date", "")
        self.log = []
        self.stale_meps = []
        self.missing_rows = saved_missing_rows(saved.get("missing"))
        self.added = 0
        details = load_last_check()
        self.rows = sorted([stored_row(o, details) for o in self.scope.openings()], key=lambda r: int(r["code"]))

    def totals(self):
        totals = status_totals(self)
        totals.insert(1, ("Report of", self.date))
        return totals

    def details(self):
        return ["No check was run: the New Report of {} again, with the statuses as written in the model now. "
                "Create, Recreate, Move and Delete run a New Report on the same scope first.".format(self.date)]

    def warning(self):
        return ""


class ReportRunner(IExternalEventHandler):
    """Esegue nel contesto API di Revit le richieste della finestra del report, che e' non modale: dai suoi pulsanti
    non si puo' chiamare l'API direttamente (selezione, transazioni)."""

    def __init__(self):
        self.jobs = []

    def Execute(self, uiapp):
        jobs, self.jobs = self.jobs, []
        for job in jobs:
            job(uiapp)

    def GetName(self):
        return "ESA Openings report"


class ReportController(object):
    """Azioni della finestra del report (OpeningsCheck_ui.ReportWindow), chiamate con call(nome, argomenti, fatto)
    e messe in coda su un ExternalEvent. Il modello puo' cambiare mentre la finestra e' aperta: Create / Recreate,
    Move e Delete rifanno prima il New Report dell'ambito (da un Last Report e' il primo controllo vero)."""

    def __init__(self, report, ui):
        self.report, self.ui = report, ui
        self.runner = ReportRunner()
        self.event = ExternalEvent.Create(self.runner)  # va creato nel contesto del comando

    def call(self, name, args, done):
        """Accoda l'azione name(*args); a fine esecuzione done(risultato, errore). Solo nel modello del report."""
        def job(uiapp):
            active = uiapp.ActiveUIDocument
            if active is None or not active.Document.Equals(doc):
                done(None, "switch back to the model '{}' to use this window.".format(doc.Title))
                return
            try:
                result = getattr(self, name)(*args)
            except Exception as ex:
                done(None, str(ex))
                return
            done(result, None)

        self.runner.jobs.append(job)
        self.event.Raise()

    # -- azioni (girano nel contesto API) -----------------------------------------------------------------------
    def check_again(self):
        """New Report dell'ambito: un Last Report diventa un check vero, un check si ripete."""
        if isinstance(self.report, CheckSession):
            self.report.run([])
        else:
            self.report = CheckSession(self.report.scope)
            self.report.log.append("New Report run on the scope from the Interactive Report.")

    def groups(self):
        return report_groups(self.ui, self.report)

    def refresh(self):
        self.check_again()
        return "New Report done on the scope: {} openings.".format(len(self.report.rows)), self.groups()

    def select(self, ids, zoom):
        element_ids = List[DB.ElementId]([DB.ElementId(i) for i in ids if doc.GetElement(DB.ElementId(i))])
        uidoc.Selection.SetElementIds(element_ids)
        if zoom and element_ids.Count:
            uidoc.ShowElements(element_ids)
        uidoc.RefreshActiveView()
        missing = len(ids) - element_ids.Count
        return "Selected {} elements{}.".format(element_ids.Count, " ({} no longer in the model)".format(missing)
                                                 if missing else "")

    def prepare_recreate(self, keys):
        self.check_again()
        return self.report.recreate_preview(keys), self.groups()

    def recreate(self, keys):
        return self.report.recreate(keys), self.groups()

    def move(self, keys):
        self.check_again()
        return self.report.move(keys), self.groups()

    def delete(self, keys):
        """Cancella i fori spuntati (la finestra offre solo i Not Associated), poi New Report dell'ambito: gli Id dei
        fori cancellati si tolgono dagli elementi MEP (CheckScope.removed)."""
        managed = dict((code_of(o), o) for o in existing_openings())
        codes = [k for k in keys if k in managed]
        if not codes:
            return "None of the ticked openings is still in the model.", self.groups()
        for code in codes:  # gli elementi dei fori cancellati restano nell'ambito e nel report
            for text in split_list(get_text(managed[code], ID_PARAM)):
                element = element_by_id(text)
                if element is not None and isinstance(element, MEP_CLASSES):
                    self.report.scope.release(element, [code])
        transaction = DB.Transaction(doc, "Openings: delete")
        transaction.Start()
        try:
            for code in codes:
                doc.Delete(managed[code].Id)
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        self.report.scope.replace(codes, [])
        if isinstance(self.report, CheckSession):
            self.report.run(codes)
        else:
            self.report = CheckSession(self.report.scope)
        message = "Deleted {} openings ({}); Ids updated on {} MEP elements ({} left without an opening).".format(
            len(codes), ", ".join(codes), self.report.plan["updated"], self.report.plan["orphaned"])
        self.report.log.append(message)
        return message, self.groups()

    def report_data(self):
        return report_data(self.report)


def revit_owner():
    try:
        return __revit__.MainWindowHandle  # noqa: F821
    except Exception:
        return None


def show_report(ui, report):
    """Riepilogo con le due schede, poi Text Report o Interactive Report (non modali: il comando finisce, la finestra
    resta aperta; l'Interactive Report lavora tramite ExternalEvent, vedi __persistentengine__)."""
    choice = ui.SummaryWindow(bundle_path(SUMMARY_XAML), report.title, report.totals(), report.warning()).show()
    if choice == ui.SHOW_TEXT:
        ui.TextReportWindow(bundle_path(TEXT_XAML), report_data(report), revit_owner()).show()
    elif choice == ui.SHOW_WINDOW:
        controller = ReportController(report, ui)
        ui.ReportWindow(bundle_path(REPORT_XAML), controller.groups(), controller, revit_owner(),
                        bundle_path(TEXT_XAML)).show()


def check_main(scope=None):
    """New Report: ambito (finestra, o dato da Place), controllo e scrittura, poi il report."""
    import OpeningsCheck_ui as ui  # copia identica nelle cartelle di Place, NewReport e LastReport
    if scope is None:
        scope = ask_check_scope(ui)
    show_report(ui, CheckSession(scope))


def last_main():
    """Last Report: l'ultimo New Report del modello, senza controllo e senza chiedere l'ambito."""
    import OpeningsCheck_ui as ui
    saved = load_last_report()
    if saved is None:
        ui.MessageBox.Show("No report yet for this model on this computer: run New Report first.", ui.TITLE)
        script.exit()
    show_report(ui, LastReport(saved))


if MODE == "place":
    place_main()
elif MODE == "last":
    last_main()
else:
    check_main()
