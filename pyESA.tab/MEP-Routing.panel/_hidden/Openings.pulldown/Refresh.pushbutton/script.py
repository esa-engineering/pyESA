# -*- coding: utf-8 -*-
__title__ = "Refresh"
__doc__ = """Version = 0.1
Date    = 09.10.2026
________________________________________________________________
Description:
Compares the provision-for-void openings in the model with the current
crossings of pipes, ducts, cable trays and conduits through walls,
floors and structural elements (host model and linked RVT / IFC), and
updates them.

1. Choose the scope: active view or whole model
2. The report lists:
   - new crossings without an opening
   - openings to change: moved, resized, rotated, other family (e.g.
     the wall became REI), other crossing elements, other parameters,
     or needed again after being marked obsolete
   - openings no longer needed (the MEP elements were removed or no
     longer cross the element)
3. Choose what to apply: all, only new, only changes, only the
   openings no longer needed (marked obsolete or deleted, as set in
   Setup), or report only. Everything applied is one Undo

Uses the same Setup (families, sizes, REI rules, parameters) and the
same list of types (Links) as Place.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

# --- corpo comune a Place e Refresh: da qui in giu' i due script sono identici tranne MODE ---------------------

import clr
import math
import json
import datetime
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import System
from System.Collections.Generic import List
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
from Autodesk.Revit.DB.ExtensibleStorage import Schema, SchemaBuilder, Entity, AccessLevel
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, DB, forms

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

MODE = "refresh"  # "place" | "refresh"
TOOL_TITLE = "Openings Place" if MODE == "place" else "Openings Refresh"


# =================================================================================================================
# Configurazione di progetto: blocco identico in Links, Setup, Place e Refresh (correggere in tutti e quattro)
# =================================================================================================================
CONFIG_SCHEMA_GUID = "5d0a2f6e-8c41-4b7a-9e3d-2a6f1c8b4e01"
OPENING_SCHEMA_GUID = "a3e7c9d2-4f18-4c6b-b5e0-7d2f9a1c3e52"
JSON_FIELD = "Json"
HOST_MODEL = "Host model"

MEP_KINDS = ("pipe", "duct_round", "duct_rect", "tray", "conduit", "merged")
MEP_KIND_LABELS = {"pipe": "Pipe", "duct_round": "Round duct", "duct_rect": "Rectangular duct",
                   "tray": "Cable tray", "conduit": "Conduit", "merged": "Merged opening"}
CASES = ("h_normal", "h_rei", "h_struct", "v_normal", "v_rei", "v_struct")
CASE_LABELS = {"h_normal": "Wall normal", "h_rei": "Wall REI", "h_struct": "Wall structural",
               "v_normal": "Floor normal", "v_rei": "Floor REI", "v_struct": "Floor structural"}
CLASS_LABELS = {"normal": "Normal", "rei": "REI", "struct": "Structural"}
PARAM_KEYS = ("level", "systems", "mep_categories", "mep_count", "host_model", "host_category", "host_type",
              "crossing_class", "status", "updated")
PARAM_LABELS = {"level": "Level", "systems": "Systems", "mep_categories": "MEP categories",
                "mep_count": "Number of MEP elements", "host_model": "Crossed model",
                "host_category": "Crossed category", "host_type": "Crossed type",
                "crossing_class": "Class (Normal / REI / Structural)", "status": "Status",
                "updated": "Last update (date)"}

DEFAULT_CONFIG = {
    "version": 1,
    "rev": 0,
    "mep": {"pipe": True, "duct": True, "tray": True, "conduit": True},
    "mode": "simple",
    "simple": {"h_rect": "", "h_round": "", "v_rect": "", "v_round": ""},
    "matrix": {},
    "dims": {"width": "", "height": "", "depth": "", "diameter": ""},
    "rei_rules": [],
    "rules": {
        "normal": {"clearance_mm": 25.0, "rounding_mm": 50.0, "merge_mm": 100.0, "min_mm": 0.0, "extra_depth_mm": 0.0},
        "rei": {"clearance_mm": 25.0, "rounding_mm": 10.0, "merge_mm": 0.0, "min_mm": 0.0, "extra_depth_mm": 0.0},
        "struct": {"clearance_mm": 50.0, "rounding_mm": 50.0, "merge_mm": 0.0, "min_mm": 0.0, "extra_depth_mm": 0.0},
    },
    "params": dict((k, "") for k in PARAM_KEYS),
    "obsolete": "mark",
    "workset": "MEP_Openings",
    "schedule": "MEP Openings",
    "filters": True,
    "inventory": {},
}

HOST_CATEGORIES = (DB.BuiltInCategory.OST_Walls, DB.BuiltInCategory.OST_Floors, DB.BuiltInCategory.OST_Roofs,
                   DB.BuiltInCategory.OST_StructuralFraming, DB.BuiltInCategory.OST_StructuralColumns,
                   DB.BuiltInCategory.OST_StructuralFoundation)
LINK_ONLY_CATEGORIES = (DB.BuiltInCategory.OST_GenericModel,)  # elementi IFC arrivati come modelli generici


def id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


def deep_merge(base, data):
    """Copia di base aggiornata con data (i dizionari annidati si fondono, il resto si sostituisce)."""
    result = json.loads(json.dumps(base))
    for key, value in (data or {}).items():
        if isinstance(value, dict) and isinstance(result.get(key), dict) and key not in ("matrix", "inventory"):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def get_schema(guid_text, name):
    guid = System.Guid(guid_text)
    schema = Schema.Lookup(guid)
    if schema is not None:
        return schema
    builder = SchemaBuilder(guid)
    builder.SetReadAccessLevel(AccessLevel.Public)
    builder.SetWriteAccessLevel(AccessLevel.Public)
    builder.SetSchemaName(name)
    builder.AddSimpleField(JSON_FIELD, clr.GetClrType(System.String))
    return builder.Finish()


def config_schema():
    return get_schema(CONFIG_SCHEMA_GUID, "ESA_OpeningsConfig")


def opening_schema():
    return get_schema(OPENING_SCHEMA_GUID, "ESA_OpeningData")


def read_json(element, schema):
    try:
        entity = element.GetEntity(schema)
        if entity is None or not entity.IsValid():
            return None
        text = entity.Get[System.String](JSON_FIELD)
        return json.loads(text) if text else None
    except Exception:
        return None


def write_json(element, schema, data):
    """Dentro una transazione."""
    entity = Entity(schema)
    entity.Set[System.String](JSON_FIELD, json.dumps(data))
    element.SetEntity(entity)


def load_config():
    return deep_merge(DEFAULT_CONFIG, read_json(doc.ProjectInformation, config_schema()))


def save_config(config):
    """Dentro una transazione. Ogni salvataggio alza la revisione del setup."""
    config["rev"] = int(config.get("rev", 0)) + 1
    write_json(doc.ProjectInformation, config_schema(), config)


def structural_category_ids():
    """Categorie strutturali: quelle il cui nome interno inizia con OST_Structural (indipendente dalla lingua)."""
    ids = set()
    for value in System.Enum.GetValues(clr.GetClrType(DB.BuiltInCategory)):
        if str(value).startswith("OST_Structural"):
            try:
                ids.add(int(value))
            except Exception:
                pass
    return ids


STRUCTURAL_IDS = structural_category_ids()


def category_value(element):
    return id_value(element.Category.Id) if element.Category is not None else None


def link_name(instance):
    link_type = doc.GetElement(instance.GetTypeId())
    return DB.Element.Name.GetValue(link_type) if link_type is not None else instance.Name


def model_sources(unique_models=False):
    """Modello host e link caricati: [{doc, T, model, inst, label}], [link non caricati]. unique_models: un solo
    elemento per link (inventario), altrimenti una voce per ogni istanza (geometria)."""
    sources = [{"doc": doc, "T": DB.Transform.Identity, "model": HOST_MODEL, "inst": "", "label": HOST_MODEL}]
    unloaded, seen = [], set()
    for instance in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance):
        name = link_name(instance)
        link_doc = instance.GetLinkDocument()
        if link_doc is None:
            if name not in unloaded:
                unloaded.append(name)
            continue
        if unique_models and name in seen:
            continue
        seen.add(name)
        sources.append({"doc": link_doc, "T": instance.GetTotalTransform(), "model": name,
                        "inst": instance.UniqueId, "label": name})
    return sources, unloaded


def source_categories(source):
    return HOST_CATEGORIES + (LINK_ONLY_CATEGORIES if source["inst"] else ())


def source_elements(source):
    categories = List[DB.BuiltInCategory]()
    for category in source_categories(source):
        categories.Add(category)
    return (DB.FilteredElementCollector(source["doc"]).WherePasses(DB.ElementMulticategoryFilter(categories))
            .WhereElementIsNotElementType())


def type_label(element, source_doc):
    type_element = source_doc.GetElement(element.GetTypeId())
    if type_element is None:
        return element.Name
    family = getattr(type_element, "FamilyName", "") or ""
    name = DB.Element.Name.GetValue(type_element)
    return "{}: {}".format(family, name) if family else name


def inventory_key(model, category, label):
    return u"{}|{}|{}".format(model, category, label)


def param_text(param):
    if param is None or not param.HasValue:
        return None
    storage = param.StorageType
    if storage == DB.StorageType.String:
        return param.AsString()
    text = param.AsValueString()
    if text is None and storage == DB.StorageType.Integer:
        text = str(param.AsInteger())
    return text


def rule_matches(rule, element, type_element):
    name = (rule.get("param") or "").strip()
    if not name:
        return False
    where = rule.get("where", "both")
    targets = []
    if where in ("instance", "both"):
        targets.append(element)
    if where in ("type", "both") and type_element is not None:
        targets.append(type_element)
    wanted = (rule.get("value") or "").strip().lower()
    for target in targets:
        for param in target.GetParameters(name):
            text = param_text(param)
            if text is None or not text.strip():
                continue
            text = text.strip().lower()
            op = rule.get("op", "nonempty")
            if op == "nonempty" or (op == "equals" and text == wanted) or (op == "contains" and wanted and wanted in text):
                return True
    return False


def rei_by_rules(rules, element, source_doc):
    """Vero se almeno una regola REI del progetto e' soddisfatta dall'elemento o dal suo tipo."""
    type_element = source_doc.GetElement(element.GetTypeId())
    return any(rule_matches(rule, element, type_element) for rule in rules or [])

# =================================================================================================================
# fine del blocco comune di configurazione
# =================================================================================================================


MM_PER_FT = 304.8
EPS_VOL = 1e-6                     # piedi cubi: sotto, l'intersezione si ignora
VERTICAL_SIN = math.sin(math.radians(45.0))  # oltre 45 gradi di inclinazione l'elemento MEP e' "verticale"
MAX_DEPTH_FT = 3000.0 / MM_PER_FT  # oltre 3 m di profondita' l'elemento corre dentro, non attraversa
MOVE_TOL_FT = 5.0 / MM_PER_FT
SIZE_TOL_MM = 1.0
ROTATION_TOL = math.radians(0.5)
Z = DB.XYZ.BasisZ

# compatibilita' tra direzione dell'elemento MEP e categoria attraversata (evita i tratti che corrono dentro)
H_CATEGORY_IDS = set(int(c) for c in (DB.BuiltInCategory.OST_Walls, DB.BuiltInCategory.OST_StructuralColumns,
                                      DB.BuiltInCategory.OST_StructuralFraming, DB.BuiltInCategory.OST_GenericModel))
V_CATEGORY_IDS = set(int(c) for c in (DB.BuiltInCategory.OST_Floors, DB.BuiltInCategory.OST_Roofs,
                                      DB.BuiltInCategory.OST_StructuralFoundation,
                                      DB.BuiltInCategory.OST_StructuralFraming, DB.BuiltInCategory.OST_GenericModel))
MEP_CATEGORIES = {"pipe": DB.BuiltInCategory.OST_PipeCurves, "duct": DB.BuiltInCategory.OST_DuctCurves,
                  "tray": DB.BuiltInCategory.OST_CableTray, "conduit": DB.BuiltInCategory.OST_Conduit}


class OpeningError(Exception):
    """Errore atteso: il messaggio va nel report."""


class RevitErrors(DB.IFailuresPreprocessor):
    """Errori di Revit al commit: la transazione viene annullata e i messaggi vanno nel report."""

    def __init__(self):
        self.messages = []

    def PreprocessFailures(self, accessor):
        for failure in accessor.GetFailureMessages():
            if failure.GetSeverity() != DB.FailureSeverity.Warning:
                self.messages.append(failure.GetDescriptionText())
        if self.messages:
            return DB.FailureProcessingResult.ProceedWithRollBack
        return DB.FailureProcessingResult.Continue


class MepFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, (Pipe, Duct, CableTray, Conduit))

    def AllowReference(self, reference, point):
        return False


# --- geometria di base ------------------------------------------------------------------------------------------

def unit(vector):
    length = vector.GetLength()
    return vector.Multiply(1.0 / length) if length > 1e-12 else vector


def flat(vector):
    return DB.XYZ(vector.X, vector.Y, 0.0)


def mm(value_ft):
    return value_ft * MM_PER_FT


def round_up(value_mm, step_mm):
    if step_mm <= 0:
        return round(value_mm, 1)
    return math.ceil(value_mm / step_mm - 1e-9) * step_mm


def solids_of(element):
    options = DB.Options()
    options.DetailLevel = DB.ViewDetailLevel.Fine
    geometry = element.get_Geometry(options)
    solids = []
    if geometry is None:
        return solids
    for item in geometry:
        if isinstance(item, DB.GeometryInstance):
            solids.extend(s for s in item.GetInstanceGeometry() if isinstance(s, DB.Solid))
        elif isinstance(item, DB.Solid):
            solids.append(item)
    return [s for s in solids if s.Volume > EPS_VOL]


def solid_points(solid):
    points = []
    for edge in solid.Edges:
        points.extend(edge.Tessellate())
    return points


def box_world(box, transform):
    """(min, max) in coordinate host di un bounding box espresso nelle coordinate del suo documento."""
    corners = []
    for x in (box.Min.X, box.Max.X):
        for y in (box.Min.Y, box.Max.Y):
            for z in (box.Min.Z, box.Max.Z):
                corners.append(transform.OfPoint(DB.XYZ(x, y, z)))
    return (DB.XYZ(min(c.X for c in corners), min(c.Y for c in corners), min(c.Z for c in corners)),
            DB.XYZ(max(c.X for c in corners), max(c.Y for c in corners), max(c.Z for c in corners)))


def boxes_overlap(a_min, a_max, b_min, b_max):
    return (a_min.X <= b_max.X and b_min.X <= a_max.X and a_min.Y <= b_max.Y and b_min.Y <= a_max.Y and
            a_min.Z <= b_max.Z and b_min.Z <= a_max.Z)


def level_elevation(level):
    # ProjectElevation e' coerente con la geometria (coordinate interne), Elevation puo' riferirsi al Survey Point
    value = getattr(level, "ProjectElevation", None)
    return level.Elevation if value is None else value


# --- elementi MEP -----------------------------------------------------------------------------------------------

def mep_kind(element):
    if isinstance(element, Pipe):
        return "pipe"
    if isinstance(element, Conduit):
        return "conduit"
    if isinstance(element, CableTray):
        return "tray"
    if isinstance(element, Duct):
        for conn in element.ConnectorManager.Connectors:
            if conn.ConnectorType == DB.ConnectorType.End:
                return "duct_round" if conn.Shape == DB.ConnectorProfileType.Round else "duct_rect"
        return "duct_rect"
    return None


def param_double(element, bip):
    param = element.get_Parameter(bip)
    if param is not None and param.HasValue and param.AsDouble() > 0:
        return param.AsDouble()
    return None


def insulation_thickness(element):
    if not isinstance(element, (Pipe, Duct)):
        return 0.0
    try:
        ids = DB.InsulationLiningBase.GetInsulationIds(doc, element.Id)
    except Exception:
        return 0.0
    values = [getattr(doc.GetElement(i), "Thickness", 0.0) or 0.0 for i in ids]
    return max(values) if values else 0.0


def round_radius(element, kind):
    if kind == "pipe":
        value = param_double(element, DB.BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
    elif kind == "conduit":
        value = param_double(element, DB.BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM)
    elif kind == "duct_round":
        value = param_double(element, DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
    else:
        return None
    return value / 2.0 if value else None


def perpendicular(direction):
    vector = direction.CrossProduct(Z)
    if vector.GetLength() < 1e-6:
        vector = direction.CrossProduct(DB.XYZ.BasisX)
    return unit(vector)


def mep_section(element):
    """Sezione esterna (isolamento compreso) e asse dell'elemento MEP rettilineo. Le sezioni rettangolari si
    leggono dal solido (vale anche per canali ruotati e passerelle di costa)."""
    curve = element.Location.Curve
    p0, d = curve.GetEndPoint(0), curve.Direction
    kind = mep_kind(element)
    ins = insulation_thickness(element)
    section = {"p0": p0, "d": d, "len": curve.Length, "kind": kind, "round": None, "rect": None}
    radius = round_radius(element, kind)
    if radius is not None:
        section["a"] = perpendicular(d)
        section["b"] = unit(d.CrossProduct(section["a"]))
        section["round"] = radius + ins
        return section
    axis_a, points = None, []
    for solid in solids_of(element):
        for face in solid.Faces:
            if axis_a is None and isinstance(face, DB.PlanarFace) and abs(face.FaceNormal.DotProduct(d)) < 1e-3:
                axis_a = unit(face.FaceNormal)
        points.extend(solid_points(solid))
    if axis_a is not None and points:
        axis_b = unit(d.CrossProduct(axis_a))
        along_a = [p.Subtract(p0).DotProduct(axis_a) for p in points]
        along_b = [p.Subtract(p0).DotProduct(axis_b) for p in points]
        extent = (min(along_a), max(along_a), min(along_b), max(along_b))
    else:
        width = (param_double(element, DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM) or
                 param_double(element, DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM))
        height = (param_double(element, DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM) or
                  param_double(element, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM))
        if not (width and height):
            raise OpeningError("The section of the element could not be read")
        axis_a = perpendicular(d)
        axis_b = unit(d.CrossProduct(axis_a))
        extent = (-width / 2.0, width / 2.0, -height / 2.0, height / 2.0)
    section["a"], section["b"] = axis_a, axis_b
    section["rect"] = (extent[0] - ins, extent[1] + ins, extent[2] - ins, extent[3] + ins)
    return section


def section_offsets(section):
    a, b = section["a"], section["b"]
    if section["round"] is not None:
        r = section["round"]
        return [a.Multiply(r * math.cos(k * math.pi / 4.0)).Add(b.Multiply(r * math.sin(k * math.pi / 4.0)))
                for k in range(8)]
    lo_a, hi_a, lo_b, hi_b = section["rect"]
    return [a.Multiply(x).Add(b.Multiply(y)) for x, y in ((lo_a, lo_b), (hi_a, lo_b), (hi_a, hi_b), (lo_a, hi_b))]


def mep_prism(section):
    """Solido della sezione esterna estrusa lungo l'asse: pulito, quindi robusto nelle operazioni booleane."""
    p0, d = section["p0"], section["d"]
    loop = DB.CurveLoop()
    if section["round"] is not None:
        plane = DB.Plane.CreateByOriginAndBasis(p0, section["a"], section["b"])
        loop.Append(DB.Arc.Create(plane, section["round"], 0.0, math.pi))
        loop.Append(DB.Arc.Create(plane, section["round"], math.pi, 2.0 * math.pi))
    else:
        corners = [p0.Add(offset) for offset in section_offsets(section)]
        for i in range(4):
            loop.Append(DB.Line.CreateBound(corners[i], corners[(i + 1) % 4]))
    loops = List[DB.CurveLoop]()
    loops.Add(loop)
    return DB.GeometryCreationUtilities.CreateExtrusionGeometry(loops, d, section["len"])


def collect_mep(scope, config):
    enabled = config["mep"]
    categories = List[DB.BuiltInCategory]()
    for key, category in MEP_CATEGORIES.items():
        if enabled.get(key, True):
            categories.Add(category)
    if categories.Count == 0:
        forms.alert("No MEP category is enabled in Setup.", title=TOOL_TITLE)
        script.exit()
    if scope == "selection":
        elements = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
        elements = [e for e in elements if MepFilter().AllowElement(e)]
        if not elements:
            try:
                refs = uidoc.Selection.PickObjects(ObjectType.Element, MepFilter(),
                                                   "Select the pipes, ducts, cable trays or conduits, then press "
                                                   "Finish")
            except OperationCanceledException:
                script.exit()
            elements = [doc.GetElement(r.ElementId) for r in refs]
    else:
        collector = (DB.FilteredElementCollector(doc, doc.ActiveView.Id) if scope == "view"
                     else DB.FilteredElementCollector(doc))
        elements = list(collector.WherePasses(DB.ElementMulticategoryFilter(categories)).WhereElementIsNotElementType())
    result = []
    for element in elements:
        kind = mep_kind(element)
        if kind is None or not enabled.get("duct" if kind.startswith("duct") else kind, True):
            continue
        if isinstance(getattr(element.Location, "Curve", None), DB.Line):
            result.append(element)
    return result


# --- elementi attraversati --------------------------------------------------------------------------------------

def gather_hosts(config, notes):
    """Elementi attraversabili dei tipi spuntati in Links, con il bounding box in coordinate host."""
    inventory = config["inventory"]
    sources, unloaded = model_sources()
    notes["unloaded"] = unloaded
    records = []
    for source in sources:
        for element in source_elements(source):
            category = category_value(element)
            label = type_label(element, source["doc"])
            key = inventory_key(source["model"], category, label)
            entry = inventory.get(key)
            if entry is None:
                notes["unknown_types"].add(key)
                continue
            if not entry.get("include"):
                continue
            box = element.get_BoundingBox(None)
            if box is None:
                continue
            box_min, box_max = box_world(box, source["T"])
            records.append({"el": element, "doc": source["doc"], "T": source["T"], "model": source["model"],
                            "inst": source["inst"], "label": source["label"], "id": id_value(element.Id),
                            "cat": category, "cat_name": element.Category.Name, "type": label, "key": key,
                            "min": box_min, "max": box_max})
    return records


SOLID_CACHE = {}


def host_solids(record):
    key = (record["inst"], record["id"])
    if key not in SOLID_CACHE:
        solids = []
        for solid in solids_of(record["el"]):
            try:
                solids.append(solid if record["T"].IsIdentity else DB.SolidUtils.CreateTransformed(solid, record["T"]))
            except Exception:
                continue
        SOLID_CACHE[key] = solids
    return SOLID_CACHE[key]


def axis_points(host_solid, section):
    """Ripiego se l'operazione booleana fallisce: tratto dell'asse dentro l'elemento, allargato alla sezione."""
    line = DB.Line.CreateBound(section["p0"], section["p0"].Add(section["d"].Multiply(section["len"])))
    result = host_solid.IntersectWithCurve(line, DB.SolidCurveIntersectionOptions())
    points = []
    for i in range(result.SegmentCount):
        segment = result.GetCurveSegment(i)
        for end in (segment.GetEndPoint(0), segment.GetEndPoint(1)):
            points.extend(end.Add(offset) for offset in section_offsets(section))
    return points


def intersect_points(prism, section, record):
    points, approx = [], False
    for host_solid in host_solids(record):
        try:
            result = DB.BooleanOperationsUtils.ExecuteBooleanOperation(prism, host_solid,
                                                                        DB.BooleanOperationsType.Intersect)
            if result is not None and result.Volume > EPS_VOL:
                points.extend(solid_points(result))
        except Exception:
            try:
                fallback = axis_points(host_solid, section)
            except Exception:
                fallback = []
            if fallback:
                points.extend(fallback)
                approx = True
    return points, approx


CLASS_CACHE = {}


def classify(record, config):
    """struct (categoria Structural...), poi rei (scelta manuale in Links o regole del Setup), poi normal."""
    key = (record["inst"], record["id"])
    if key not in CLASS_CACHE:
        if record["cat"] in STRUCTURAL_IDS:
            value = "struct"
        else:
            override = config["inventory"].get(record["key"], {}).get("rei", "auto")
            if override == "rei":
                value = "rei"
            elif override == "not_rei":
                value = "normal"
            else:
                value = "rei" if rei_by_rules(config["rei_rules"], record["el"], record["doc"]) else "normal"
        CLASS_CACHE[key] = value
    return CLASS_CACHE[key]


def host_ref(record):
    """Riferimento stabile all'elemento attraversato: IfcGUID se c'e' (stabile anche ricaricando un IFC),
    altrimenti UniqueId."""
    param = record["el"].get_Parameter(DB.BuiltInParameter.IFC_GUID)
    if param is not None and param.HasValue and param.AsString():
        return "ifc:" + param.AsString()
    return record["el"].UniqueId


def host_tangent(record, point):
    location = getattr(record["el"], "Location", None)
    if not isinstance(location, DB.LocationCurve):
        return None
    try:
        curve = location.Curve
        if not record["T"].IsIdentity:
            curve = curve.CreateTransformed(record["T"])
        projection = curve.Project(point)
        parameter = projection.Parameter if projection is not None else curve.GetEndParameter(0)
        tangent = flat(curve.ComputeDerivatives(parameter, False).BasisX)
    except Exception:
        return None
    return unit(tangent) if tangent.GetLength() > 1e-6 else None


# --- attraversamenti e fori -------------------------------------------------------------------------------------

def find_crossings(meps, hosts, stats):
    crossings = []
    for mep in meps:
        vertical = abs(mep.Location.Curve.Direction.Z) >= VERTICAL_SIN
        allowed = V_CATEGORY_IDS if vertical else H_CATEGORY_IDS
        box = mep.get_BoundingBox(None)
        if box is None:
            continue
        grow = DB.XYZ(1.0, 1.0, 1.0).Multiply(insulation_thickness(mep) + 0.05)
        mep_min, mep_max = box.Min.Subtract(grow), box.Max.Add(grow)
        candidates = [r for r in hosts
                      if r["cat"] in allowed and boxes_overlap(mep_min, mep_max, r["min"], r["max"])]
        if not candidates:
            continue
        # sezione e solido solo per gli elementi che toccano qualcosa: sul modello intero la maggior parte no
        try:
            section = mep_section(mep)
            prism = mep_prism(section)
        except Exception as ex:
            stats["mep_errors"].append((mep, str(ex)))
            continue
        for record in candidates:
            points, approx = intersect_points(prism, section, record)
            if points:
                crossings.append({"mep": mep, "section": section, "rec": record, "pts": points, "approx": approx,
                                  "orient": "v" if vertical else "h"})
    return crossings


def centroid(points):
    total = DB.XYZ.Zero
    for p in points:
        total = total.Add(p)
    return total.Multiply(1.0 / len(points))


def group_frame(items, orient):
    """Terna (u, v, w) del foro. Orizzontale: u lungo l'elemento attraversato, v verticale, w attraverso.
    Verticale: u, v in pianta (allineati alla sezione rettangolare o alla fila di montanti), w verticale."""
    if orient == "h":
        tangent = host_tangent(items[0]["rec"], centroid(items[0]["pts"]))
        if tangent is None:
            through = unit(flat(items[0]["section"]["d"]))
            tangent = unit(Z.CrossProduct(through))
        u = tangent
        return u, Z, unit(Z.CrossProduct(u))
    u = None
    rects = [i for i in items if i["section"]["rect"] is not None]
    if rects:
        for axis in (rects[0]["section"]["a"], rects[0]["section"]["b"]):
            if flat(axis).GetLength() > 0.5:
                u = unit(flat(axis))
                break
    if u is None and len(items) >= 2:
        centers = [flat(centroid(i["pts"])) for i in items]
        best, pair = 0.0, None
        for i in range(len(centers)):
            for j in range(i + 1, len(centers)):
                distance = centers[i].DistanceTo(centers[j])
                if distance > best:
                    best, pair = distance, (i, j)
        if pair is not None and best > 1e-6:
            u = unit(centers[pair[1]].Subtract(centers[pair[0]]))
    if u is None:
        u = DB.XYZ.BasisX
    return u, unit(Z.CrossProduct(u)), Z


def extents(points, frame):
    u, v, w = frame
    pu = [p.DotProduct(u) for p in points]
    pv = [p.DotProduct(v) for p in points]
    pw = [p.DotProduct(w) for p in points]
    return min(pu), max(pu), min(pv), max(pv), min(pw), max(pw)


def rect_gap(a, b):
    return max(b[0] - a[1], a[0] - b[1], b[2] - a[3], a[2] - b[3])


def merge_clusters(items, merge_ft):
    clusters = [{"items": [item], "rect": item["ext"][:4]} for item in items]
    if merge_ft <= 0:
        return clusters
    changed = True
    while changed:
        changed = False
        for i in range(len(clusters)):
            for j in range(i + 1, len(clusters)):
                a, b = clusters[i]["rect"], clusters[j]["rect"]
                if rect_gap(a, b) <= merge_ft:
                    clusters[i]["items"].extend(clusters[j]["items"])
                    clusters[i]["rect"] = (min(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), max(a[3], b[3]))
                    del clusters[j]
                    changed = True
                    break
            if changed:
                break
    return clusters


def choose_family(config, orient, cls, kind, single_round):
    if config["mode"] == "simple":
        simple = config["simple"]
        if single_round and simple.get(orient + "_round"):
            return simple[orient + "_round"]
        return simple.get(orient + "_rect", "")
    return config["matrix"].get("{}|{}_{}".format(kind, orient, cls), "")


LEVELS = []


def pick_level(z, orient):
    """Orizzontale: livello sotto il centro del foro. Verticale (solai): livello piu' vicino."""
    if not LEVELS:
        LEVELS.extend(sorted(((level_elevation(l), l) for l in DB.FilteredElementCollector(doc).OfClass(DB.Level)),
                             key=lambda pair: pair[0]))
    if orient == "v":
        return min(LEVELS, key=lambda pair: abs(pair[0] - z))[1]
    below = [pair for pair in LEVELS if pair[0] <= z + 1e-3]
    return (below[-1] if below else LEVELS[0])[1]


def make_opening(cluster, frame, record, orient, cls, rules, config):
    points = []
    approx = False
    meps = {}
    for item in cluster["items"]:
        points.extend(item["pts"])
        approx = approx or item["approx"]
        meps[id_value(item["mep"].Id)] = item["mep"]
    u, v, w = frame
    e = extents(points, frame)
    center = u.Multiply((e[0] + e[1]) / 2.0).Add(v.Multiply((e[2] + e[3]) / 2.0)).Add(w.Multiply((e[4] + e[5]) / 2.0))
    raw = (mm(e[1] - e[0]), mm(e[3] - e[2]), mm(e[5] - e[4]))
    clearance, step = float(rules["clearance_mm"]), float(rules["rounding_mm"])
    dims = {"w": round_up(raw[0] + 2 * clearance, step), "h": round_up(raw[1] + 2 * clearance, step),
            "d": round(raw[2] + 2 * float(rules["extra_depth_mm"]), 1),
            "dia": round_up(max(raw[0], raw[1]) + 2 * clearance, step)}
    mep_list = list(meps.values())
    kind = "merged" if len(mep_list) > 1 else mep_kind(mep_list[0])
    single_round = len(mep_list) == 1 and cluster["items"][0]["section"]["round"] is not None
    opening = {"rec": record, "orient": orient, "cls": cls, "kind": kind, "meps": mep_list, "frame": frame,
               "center": center, "raw": raw, "dims": dims, "approx": approx,
               "family": choose_family(config, orient, cls, kind, single_round),
               "angle": math.atan2(u.Y, u.X), "crossings": len(cluster["items"])}
    opening["level"] = pick_level(center.Z, orient)
    opening["values"] = descriptive_values(opening)
    return opening


def plan_openings(crossings, config, stats):
    groups = {}
    for crossing in crossings:
        record = crossing["rec"]
        groups.setdefault((record["inst"], record["id"], crossing["orient"]), []).append(crossing)
    openings = []
    for items in groups.values():
        record, orient = items[0]["rec"], items[0]["orient"]
        cls = classify(record, config)
        rules = config["rules"][cls]
        frame = group_frame(items, orient)
        kept = []
        for item in items:
            e = extents(item["pts"], frame)
            if e[5] - e[4] > MAX_DEPTH_FT:
                stats["inside"].append(item)
                continue
            if max(mm(e[1] - e[0]), mm(e[3] - e[2])) < float(rules["min_mm"]):
                stats["below"].append(item)
                continue
            item["ext"] = e
            kept.append(item)
        for cluster in merge_clusters(kept, float(rules["merge_mm"]) / MM_PER_FT):
            openings.append(make_opening(cluster, frame, record, orient, cls, rules, config))
    return openings


def descriptive_values(opening):
    systems, categories = set(), set()
    for mep in opening["meps"]:
        param = mep.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
        if param is not None and param.AsString():
            systems.add(param.AsString())
        if mep.Category is not None:
            categories.add(mep.Category.Name)
    record = opening["rec"]
    return {"level": opening["level"].Name, "systems": ", ".join(sorted(systems)),
            "mep_categories": ", ".join(sorted(categories)), "mep_count": str(len(opening["meps"])),
            "host_model": record["label"], "host_category": record["cat_name"], "host_type": record["type"],
            "crossing_class": CLASS_LABELS[opening["cls"]]}


# --- fori esistenti ---------------------------------------------------------------------------------------------

def opening_host_key(record_or_data):
    if "rec" in record_or_data:
        record = record_or_data["rec"]
        return (record["model"], record["inst"], host_ref(record))
    return (record_or_data.get("model"), record_or_data.get("inst"), record_or_data.get("host"))


def existing_openings(schema):
    result = []
    for instance in (DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_GenericModel)
                     .WhereElementIsNotElementType()):
        if not isinstance(instance, DB.FamilyInstance):
            continue
        data = read_json(instance, schema)
        if data:
            result.append({"inst": instance, "data": data})
    return result


def match_openings(openings, existing):
    """Abbina ogni foro calcolato a quello esistente sullo stesso elemento con piu' elementi MEP in comune."""
    by_host = {}
    for index, item in enumerate(existing):
        by_host.setdefault(opening_host_key(item["data"]), []).append(index)
    used = set()
    for opening in openings:
        wanted = set(m.UniqueId for m in opening["meps"])
        best, best_score = None, 0
        for index in by_host.get(opening_host_key(opening), []):
            if index in used:
                continue
            score = len(wanted & set(existing[index]["data"].get("meps", [])))
            if score > best_score:
                best, best_score = index, score
        if best is not None:
            used.add(best)
            opening["existing"] = existing[best]
    return [item for index, item in enumerate(existing) if index not in used]


def compare(opening, existing):
    data, instance = existing["data"], existing["inst"]
    changes = []
    if data.get("status") == "Obsolete":
        changes.append("needed again (was obsolete)")
    if data.get("family") != opening["family"]:
        changes.append("family: {} -> {}".format(data.get("family") or "-", opening["family"] or "-"))
    if data.get("cls") != opening["cls"]:
        changes.append("class: {} -> {}".format(CLASS_LABELS.get(data.get("cls"), "-"), CLASS_LABELS[opening["cls"]]))
    location = instance.Location
    if not isinstance(location, DB.LocationPoint):
        changes.append("position not readable")
    else:
        distance = location.Point.DistanceTo(opening["center"])
        if distance > MOVE_TOL_FT:
            changes.append("moved {:.0f} mm".format(mm(distance)))
        turn = abs((location.Rotation - opening["angle"] + math.pi) % (2 * math.pi) - math.pi)
        if turn > ROTATION_TOL and abs(turn - math.pi) > ROTATION_TOL:
            changes.append("rotated {:.1f} degrees".format(math.degrees(turn)))
    old = data.get("dims", {})
    if any(abs(float(old.get(k, -1)) - opening["dims"][k]) > SIZE_TOL_MM for k in ("w", "h", "d", "dia")):
        changes.append("size: {} -> {}".format(size_text(old), size_text(opening["dims"])))
    if set(data.get("meps", [])) != set(m.UniqueId for m in opening["meps"]):
        changes.append("crossing elements changed ({} -> {})".format(len(data.get("meps", [])), len(opening["meps"])))
    old_values = data.get("values", {})
    changed = [PARAM_LABELS[k] for k in opening["values"] if old_values.get(k) != opening["values"][k]]
    if changed:
        changes.append("parameters: {}".format(", ".join(changed)))
    return changes


def size_text(dims):
    if not dims:
        return "-"
    return "{:.0f} x {:.0f} x {:.0f} mm (diameter {:.0f})".format(float(dims.get("w", 0)), float(dims.get("h", 0)),
                                                               float(dims.get("d", 0)), float(dims.get("dia", 0)))


# --- scrittura nel modello --------------------------------------------------------------------------------------

def today():
    return datetime.date.today().strftime("%d.%m.%Y")


def symbol_key(symbol):
    return u"{} : {}".format(symbol.FamilyName, DB.Element.Name.GetValue(symbol))


def symbol_map():
    result = {}
    for symbol in (DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
                   .OfCategory(DB.BuiltInCategory.OST_GenericModel)):
        result[symbol_key(symbol)] = symbol
    return result


def workset_id(name):
    if not doc.IsWorkshared or not name:
        return None
    for workset in DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset):
        if workset.Name == name:
            return workset.Id
    return None


def set_length(instance, name, value_mm):
    """True scritto, False non trovato (o non di istanza / non lunghezza)."""
    param = instance.LookupParameter(name) if name else None
    if param is None or param.IsReadOnly or param.StorageType != DB.StorageType.Double:
        return False
    param.Set(value_mm / MM_PER_FT)
    return True


def set_dims(instance, opening, names):
    """Diametro se la famiglia ha il parametro di diametro (foro tondo), altrimenti larghezza e altezza; sempre la
    profondita'. Ritorna le note per il report."""
    dims = opening["dims"]
    written, missing = [], []
    diameter_name = names.get("diameter", "")
    if diameter_name and instance.LookupParameter(diameter_name) is not None:
        pairs = (("diameter", dims["dia"]), ("depth", dims["d"]))
    else:
        pairs = (("width", dims["w"]), ("height", dims["h"]), ("depth", dims["d"]))
    for key, value in pairs:
        name = names.get(key, "")
        if not name:
            continue
        (written if set_length(instance, name, value) else missing).append(name)
    if not written:
        return ["size not set: placed and oriented only{}".format(
            " (missing instance parameters: {})".format(", ".join(missing)) if missing else
            " (no size parameter in Setup)")]
    if missing:
        return ["missing instance parameters: {}".format(", ".join(missing))]
    return []


def set_text(instance, name, value):
    param = instance.LookupParameter(name) if name else None
    if param is None or param.IsReadOnly:
        return False
    try:
        if param.StorageType == DB.StorageType.String:
            param.Set(u"{}".format(value))
        elif param.StorageType == DB.StorageType.Integer:
            param.Set(int(value))
        elif param.StorageType == DB.StorageType.Double:
            param.Set(float(value))
        else:
            return False
    except Exception:
        return False
    return True


def write_values(instance, values, status, names):
    missing = []
    full = dict(values)
    full["status"], full["updated"] = status, today()
    for key in PARAM_KEYS:
        name = names.get(key, "")
        if name and not set_text(instance, name, full.get(key, "")):
            missing.append(name)
    return ["parameters not written: {}".format(", ".join(missing))] if missing else []


def align(instance, opening):
    location = instance.Location
    delta = opening["center"].Subtract(location.Point)
    if delta.GetLength() > 1e-6:
        location.Move(delta)
    turn = opening["angle"] - location.Rotation
    if abs(turn) > 1e-6:
        center = opening["center"]
        location.Rotate(DB.Line.CreateBound(center, center.Add(Z)), turn)


def identity(opening, status, created, config):
    record = opening["rec"]
    center = opening["center"]
    return {"v": 1, "model": record["model"], "inst": record["inst"], "host": host_ref(record),
            "orient": opening["orient"], "cls": opening["cls"], "kind": opening["kind"],
            "meps": sorted(m.UniqueId for m in opening["meps"]), "family": opening["family"],
            "center": [center.X, center.Y, center.Z], "dims": opening["dims"], "values": opening["values"],
            "status": status, "created": created, "updated": today(), "rev": config.get("rev", 0)}


def activate(symbol):
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()


def create_opening(opening, symbols, config, workset, schema):
    symbol = symbols.get(opening["family"])
    if symbol is None:
        raise OpeningError("Family type not loaded in the project: {}".format(opening["family"]))
    activate(symbol)
    try:
        instance = doc.Create.NewFamilyInstance(opening["center"], symbol, opening["level"],
                                                StructuralType.NonStructural)
    except Exception as ex:
        raise OpeningError("The family could not be placed (it must be non-hosted, level based): {}".format(ex))
    doc.Regenerate()
    align(instance, opening)
    notes = set_dims(instance, opening, config["dims"])
    notes += write_values(instance, opening["values"], "New", config["params"])
    if workset is not None:
        param = instance.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
        if param is not None and not param.IsReadOnly:
            param.Set(workset.IntegerValue)
    write_json(instance, schema, identity(opening, "New", today(), config))
    return instance, notes


def update_opening(opening, existing, symbols, config, workset, schema):
    instance = existing["inst"]
    notes = []
    if existing["data"].get("family") != opening["family"]:
        symbol = symbols.get(opening["family"])
        if symbol is None:
            raise OpeningError("Family type not loaded in the project: {}".format(opening["family"]))
        activate(symbol)
        try:
            new_id = instance.ChangeTypeId(symbol.Id)
            if new_id is not None and new_id != DB.ElementId.InvalidElementId and new_id != instance.Id:
                instance = doc.GetElement(new_id)
        except Exception:
            # tipo non compatibile: si sostituisce il foro
            doc.Delete(instance.Id)
            instance, notes = create_opening(opening, symbols, config, workset, schema)
            notes.append("replaced with a new instance")
            return instance, notes
        doc.Regenerate()
    align(instance, opening)
    notes += set_dims(instance, opening, config["dims"])
    notes += write_values(instance, opening["values"], "Modified", config["params"])
    write_json(instance, schema, identity(opening, "Modified", existing["data"].get("created", today()), config))
    return instance, notes


def retire_opening(existing, config, schema):
    if config.get("obsolete") == "delete":
        doc.Delete(existing["inst"].Id)
        return "deleted"
    instance = existing["inst"]
    set_text(instance, config["params"].get("status", ""), "Obsolete")
    data = dict(existing["data"])
    data["status"], data["updated"] = "Obsolete", today()
    write_json(instance, schema, data)
    return "marked obsolete"


def run_transaction(name, action):
    """Esegue action() in una transazione. Ritorna (valore, None) oppure (None, motivo)."""
    transaction = DB.Transaction(doc, name)
    errors = RevitErrors()
    options = transaction.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(errors)
    options.SetClearAfterRollback(True)
    transaction.SetFailureHandlingOptions(options)
    transaction.Start()
    try:
        value = action()
        if transaction.Commit() != DB.TransactionStatus.Committed:
            return None, "Revit refused the result: {}".format("; ".join(errors.messages) or "rolled back")
        return value, None
    except Exception as ex:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        return None, str(ex)


# --- report -----------------------------------------------------------------------------------------------------

def opening_line(index, opening, extra=""):
    record = opening["rec"]
    meps = ", ".join("{} {}".format(MEP_KIND_LABELS.get(mep_kind(m), "MEP"), output.linkify(m.Id))
                     for m in opening["meps"][:6])
    if len(opening["meps"]) > 6:
        meps += " (+{})".format(len(opening["meps"]) - 6)
    dims = opening["dims"]
    return (u"**#{}** {} | {} {} ({} {}) | **{}** {} | {:.0f} x {:.0f} x {:.0f} mm (diameter {:.0f}) | "
            u"family: {} | {}{}{}".format(
                index, CLASS_LABELS[opening["cls"]], record["cat_name"], record["type"], record["label"],
                record["id"], "wall" if opening["orient"] == "h" else "floor",
                "merged {}".format(opening["crossings"]) if opening["crossings"] > 1 else "",
                dims["w"], dims["h"], dims["d"], dims["dia"], opening["family"] or "**none set**", meps,
                " | approximate geometry" if opening["approx"] else "", extra))


def print_header(scope, meps, crossings, openings, config, notes, stats):
    output.print_md("# {}".format(TOOL_TITLE))
    output.print_md("Scope: **{}** | MEP elements: **{}** | crossings: **{}** | openings: **{}** | setup rev "
                    "{} | mode: **{}**".format(SCOPE_LABELS[scope], len(meps), len(crossings), len(openings),
                                               config.get("rev", 0), config["mode"]))
    if notes["unloaded"]:
        output.print_md("Links not loaded (not checked): {}".format(", ".join(notes["unloaded"])))
    if notes["unknown_types"]:
        output.print_md("**{} types are not in the list of Links** (new or renamed types: not checked). Run Links "
                        "to add them.".format(len(notes["unknown_types"])))
    if stats["below"]:
        output.print_md("Crossings below the minimum size of their class: {}".format(len(stats["below"])))
    if stats["inside"]:
        output.print_md("Runs inside an element rather than through it (no opening): {}".format(len(stats["inside"])))
    for mep, message in stats["mep_errors"]:
        output.print_md("Not checked: {} | {}".format(output.linkify(mep.Id), message))


SCOPE_LABELS = {"selection": "selected elements", "view": "active view", "model": "whole model"}


# --- comando ----------------------------------------------------------------------------------------------------

def check_setup(config):
    if not config["inventory"]:
        forms.alert("Run Links first: it lists the types of walls, floors and structural elements to check.",
                    title=TOOL_TITLE)
        script.exit()
    if config["mode"] == "simple":
        ready = config["simple"].get("h_rect") or config["simple"].get("v_rect")
    else:
        ready = any(config["matrix"].values())
    if not ready:
        forms.alert("No opening family is set. Run Setup first.", title=TOOL_TITLE)
        script.exit()


def ask_scope():
    options = {"Active view": "view", "Whole model": "model"}
    if MODE == "place":
        options["Selected elements"] = "selection"
    choice = forms.CommandSwitchWindow.show(sorted(options.keys()), message="{}: scope".format(TOOL_TITLE))
    if not choice:
        script.exit()
    return options[choice]


def compute(scope, config):
    notes = {"unloaded": [], "unknown_types": set()}
    stats = {"below": [], "inside": [], "mep_errors": []}
    meps = collect_mep(scope, config)
    hosts = gather_hosts(config, notes)
    crossings = find_crossings(meps, hosts, stats)
    openings = plan_openings(crossings, config, stats)
    return meps, crossings, openings, notes, stats


def place_main(scope, config):
    meps, crossings, openings, notes, stats = compute(scope, config)
    schema = opening_schema()
    existing = existing_openings(schema)
    match_openings(openings, existing)
    new = [o for o in openings if "existing" not in o]
    present = [o for o in openings if "existing" in o]
    print_header(scope, meps, crossings, openings, config, notes, stats)
    output.print_md("Already with an opening (use Refresh to update): **{}** | to create: **{}**".format(
        len(present), len(new)))
    output.print_md("## Preview")
    for index, opening in enumerate(new, 1):
        output.print_md(opening_line(index, opening))
    creatable = [o for o in new if o["family"]]
    if not creatable:
        output.print_md("Nothing to create.")
        return
    if not forms.alert("Create {} openings? (details in the report; {} without a family are skipped)".format(
            len(creatable), len(new) - len(creatable)), title=TOOL_TITLE, yes=True, no=True):
        output.print_md("Cancelled: nothing was created.")
        return
    symbols = symbol_map()
    workset = workset_id(config.get("workset", ""))
    results = []
    group = DB.TransactionGroup(doc, TOOL_TITLE)
    group.Start()
    for opening in creatable:
        value, error = run_transaction("Place opening",
                                       lambda o=opening: create_opening(o, symbols, config, workset, schema))
        results.append((opening, value, error))
    if any(value for _, value, _ in results):
        group.Assimilate()
    else:
        group.RollBack()
    output.print_md("## Result")
    if config.get("workset") and doc.IsWorkshared and workset is None:
        output.print_md("Workset **{}** not found: openings left on the active workset (run Setup > Apply)".format(
            config["workset"]))
    for index, (opening, value, error) in enumerate(results, 1):
        if error:
            output.print_md(opening_line(index, opening, " | **NOT CREATED**: {}".format(error)))
        else:
            instance, extra = value
            output.print_md(opening_line(index, opening, " | created {}{}".format(
                output.linkify(instance.Id), (" | " + "; ".join(extra)) if extra else "")))


def refresh_main(scope, config):
    meps, crossings, openings, notes, stats = compute(scope, config)
    schema = opening_schema()
    existing = existing_openings(schema)
    if scope == "view":
        # nella vista: solo i fori che riguardano elementi MEP della vista (o i cui elementi MEP non esistono piu')
        in_view = set(m.UniqueId for m in meps)
        existing = [e for e in existing if (set(e["data"].get("meps", [])) & in_view) or
                    all(doc.GetElement(uid) is None for uid in e["data"].get("meps", []))]
    unmatched = match_openings(openings, existing)
    new = [o for o in openings if "existing" not in o]
    changed, unchanged = [], 0
    for opening in openings:
        if "existing" in opening:
            opening["changes"] = compare(opening, opening["existing"])
            if opening["changes"]:
                changed.append(opening)
            else:
                unchanged += 1
    obsolete = [e for e in unmatched if e["data"].get("status") != "Obsolete" or config.get("obsolete") == "delete"]
    already_obsolete = len(unmatched) - len(obsolete)

    print_header(scope, meps, crossings, openings, config, notes, stats)
    output.print_md("New: **{}** | changed: **{}** | no longer needed: **{}** | unchanged: **{}** | already marked "
                    "obsolete: {}".format(len(new), len(changed), len(obsolete), unchanged, already_obsolete))
    if new:
        output.print_md("## New")
        for index, opening in enumerate(new, 1):
            output.print_md(opening_line(index, opening))
    if changed:
        output.print_md("## Changed")
        for index, opening in enumerate(changed, 1):
            output.print_md(opening_line(index, opening, " | {} {}".format(
                output.linkify(opening["existing"]["inst"].Id), "; ".join(opening["changes"]))))
    if obsolete:
        output.print_md("## No longer needed")
        for index, item in enumerate(obsolete, 1):
            data = item["data"]
            gone = all(doc.GetElement(uid) is None for uid in data.get("meps", []))
            output.print_md("**#{}** {} | {} {} | {}".format(
                index, output.linkify(item["inst"].Id), data.get("values", {}).get("host_category", ""),
                data.get("values", {}).get("host_type", ""),
                "the MEP elements no longer exist" if gone else
                "no longer crosses the element (moved, removed, or type excluded in Links)"))

    options = []
    if new or changed or obsolete:
        options.append("Apply all")
    if new:
        options.append("Only new")
    if changed:
        options.append("Only changes")
    if obsolete:
        options.append("Only no longer needed ({})".format(
            "delete" if config.get("obsolete") == "delete" else "mark obsolete"))
    if not options:
        output.print_md("Everything is up to date.")
        return
    options.append("Report only")
    choice = forms.CommandSwitchWindow.show(options, message="{}: what to apply?".format(TOOL_TITLE))
    if not choice or choice == "Report only":
        output.print_md("Nothing was changed.")
        return
    do_new = choice in ("Apply all", "Only new")
    do_changed = choice in ("Apply all", "Only changes")
    do_obsolete = choice == "Apply all" or choice.startswith("Only no longer")

    symbols = symbol_map()
    workset = workset_id(config.get("workset", ""))
    lines, done = [], 0
    group = DB.TransactionGroup(doc, TOOL_TITLE)
    group.Start()
    if do_new:
        for opening in [o for o in new if o["family"]]:
            value, error = run_transaction("Place opening",
                                           lambda o=opening: create_opening(o, symbols, config, workset, schema))
            done += 0 if error else 1
            lines.append(opening_line(len(lines) + 1, opening, " | **NOT CREATED**: {}".format(error) if error else
                                      " | created {}".format(output.linkify(value[0].Id))))
    if do_changed:
        for opening in changed:
            value, error = run_transaction(
                "Update opening",
                lambda o=opening: update_opening(o, o["existing"], symbols, config, workset, schema))
            done += 0 if error else 1
            lines.append(opening_line(len(lines) + 1, opening, " | **NOT UPDATED**: {}".format(error) if error else
                                      " | updated {}".format(output.linkify(value[0].Id))))
    if do_obsolete:
        for item in obsolete:
            value, error = run_transaction("Retire opening", lambda i=item: retire_opening(i, config, schema))
            done += 0 if error else 1
            lines.append("**#{}** opening {} | {}".format(len(lines) + 1, item["inst"].Id,
                                                         "**NOT CHANGED**: {}".format(error) if error else value))
    if done:
        group.Assimilate()
    else:
        group.RollBack()
    output.print_md("## Applied ({})".format(choice))
    for line in lines:
        output.print_md(line)


def main():
    config = load_config()
    check_setup(config)
    scope = ask_scope()
    if MODE == "place":
        place_main(scope, config)
    else:
        refresh_main(scope, config)


main()
