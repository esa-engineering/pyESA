# -*- coding: utf-8 -*-
"""Quantity Takeoff - Extract geometric quantities by category"""

__title__ = "Quantity\nTakeoff"
__author__ = "PyRevit Script"
__doc__ = "Extracts geometric quantities (length, area, volume, thickness, etc.) " \
          "for all model elements of the active document and/or of the loaded " \
          "Revit links (nested links included), organized by category."

import clr
import csv
import os
from datetime import datetime

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from pyrevit import revit, DB, forms, script

# Import System for BuiltInCategory conversion
import System

# Conversion constants (feet to meters)
FEET_TO_METERS = 0.3048
SQFEET_TO_SQMETERS = 0.09290304
CUBICFEET_TO_CUBICMETERS = 0.028316846592

# Max depth when walking nested Revit links
MAX_LINK_DEPTH = 5


def get_element_id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value       # Revit 2024+
    return eid.IntegerValue    # Revit <= 2023


def safe_category_id(element):
    """Category id value of an element, None when it has no category."""
    try:
        if element.Category is not None:
            return get_element_id_value(element.Category.Id)
    except:
        pass
    return None


def safe_get_builtin(name):
    """Gets a BuiltInParameter safely, returns None if it doesn't exist."""
    try:
        return getattr(BuiltInParameter, name)
    except AttributeError:
        return None


def safe_get_builtin_category(name):
    """Gets a BuiltInCategory safely, returns None if it doesn't exist."""
    try:
        return getattr(BuiltInCategory, name)
    except AttributeError:
        return None


# =============================================================================
# MAPPING GEOMETRIC QUANTITIES -> LIST OF BUILTIN PARAMETERS
# For each quantity, list of possible BuiltInParameters (in priority order)
# Script will try instance first, then type
# =============================================================================

GEOMETRIC_PARAM_NAMES = {
    "Length": ["CURVE_ELEM_LENGTH", "INSTANCE_LENGTH_PARAM",
               "STRUCTURAL_FRAME_CUT_LENGTH", "STAIRS_ACTUAL_RUN_LENGTH",
               "STAIRS_RUN_ACTUAL_RUN_LENGTH", "RAMP_ATTR_LENGTH",
               "REBAR_ELEM_LENGTH", "STRUCTURAL_FOUNDATION_LENGTH",
               "PATH_REIN_LENGTH_1"],
    "Width": ["WALL_ATTR_WIDTH_PARAM", "DOOR_WIDTH", "WINDOW_WIDTH",
              "FAMILY_WIDTH_PARAM", "STAIRS_ATTR_TREAD_WIDTH",
              "STAIRS_RUN_ACTUAL_RUN_WIDTH", "RAMP_ATTR_WIDTH",
              "RBS_CURVE_WIDTH_PARAM", "RBS_CABLETRAY_WIDTH_PARAM",
              "CURTAIN_WALL_PANELS_WIDTH", "STRUCTURAL_FOUNDATION_WIDTH",
              "CASEWORK_WIDTH"],
    "Height": ["WALL_USER_HEIGHT_PARAM", "DOOR_HEIGHT", "WINDOW_HEIGHT",
               "FAMILY_HEIGHT_PARAM", "STAIRS_ACTUAL_RISER_HEIGHT",
               "RBS_CURVE_HEIGHT_PARAM", "RBS_CABLETRAY_HEIGHT_PARAM",
               "CURTAIN_WALL_PANELS_HEIGHT", "CASEWORK_HEIGHT",
               "INSTANCE_HEIGHT_PARAM"],
    "Depth": ["FAMILY_DEPTH_PARAM", "CASEWORK_DEPTH"],
    "Thickness": ["FLOOR_ATTR_THICKNESS_PARAM",
                  "ROOF_ATTR_DEFAULT_THICKNESS_PARAM",
                  "ROOF_ATTR_THICKNESS_VALUE", "CEILING_THICKNESS",
                  "WALL_ATTR_WIDTH_PARAM", "BUILDINGPAD_THICKNESS",
                  "SLAB_EDGE_THICKNESS"],
    "Diameter": ["RBS_PIPE_DIAMETER_PARAM", "RBS_CONDUIT_DIAMETER_PARAM",
                 "RBS_CURVE_DIAMETER_PARAM", "REBAR_BAR_DIAMETER"],
    "Perimeter": ["HOST_PERIMETER_COMPUTED", "ROOM_PERIMETER"],
    "Area": ["HOST_AREA_COMPUTED", "ROOM_AREA", "RBS_CURVE_SURFACE_AREA",
             "MASS_GROSS_SURFACE_AREA", "MASS_SURFACE_AREA"],
    "Volume": ["HOST_VOLUME_COMPUTED", "ROOM_VOLUME", "MASS_GROSS_VOLUME",
               "MASS_VOLUME"],
}


def build_geometric_params_map():
    """
    Resolves GEOMETRIC_PARAM_NAMES into real BuiltInParameters, dropping the
    ones the running Revit version does not know.
    Returns a dictionary: { "Length": [list of valid BuiltInParameters], ... }
    """
    params_map = {}

    for geom_name, names in GEOMETRIC_PARAM_NAMES.items():
        resolved = [safe_get_builtin(name) for name in names]
        params_map[geom_name] = [bp for bp in resolved if bp is not None]

    return params_map


def get_doc_key(document):
    """
    Unique key for a document: file path when available, title otherwise.
    Used to de-duplicate links (several instances of the same link, or the same
    link nested under different parents).
    """
    try:
        path = document.PathName
        if path:
            return path.lower()
    except:
        pass
    try:
        return document.Title.lower()
    except:
        return str(document)


def get_cached_doc_key(document, cache):
    """get_doc_key() memoized per document: avoids a PathName read per element."""
    doc_keys = cache["doc_keys"]
    try:
        key = doc_keys.get(document)
        if key is None:
            key = get_doc_key(document)
            doc_keys[document] = key
        return key
    except:
        return get_doc_key(document)


def is_import_family(element, elem_type):
    """
    True when the element belongs to an "Import Symbol" family.
    Depends on the type only, so it is evaluated once per type (see
    get_type_entry) instead of once per element.
    """
    try:
        if hasattr(element, 'Symbol') and element.Symbol and element.Symbol.Family:
            if "import" in element.Symbol.Family.Name.lower():
                return True
    except:
        pass

    try:
        family_param = elem_type.get_Parameter(
            BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM) if elem_type else None
        if family_param and family_param.HasValue:
            family_name = family_param.AsString()
            return bool(family_name) and "import" in family_name.lower()
    except:
        pass

    return False


def is_element_from_link_or_import(element, source_doc, cache):
    """
    Checks if an element comes from a link (Revit or CAD/DWG) or is an import,
    relative to the document currently being processed.
    Returns True if the element must be skipped, False otherwise.

    The link/import keyword check on the category name is not repeated here:
    those categories are already discarded once per document by
    get_model_categories().
    """
    try:
        # ImportInstance (DWG, DXF, SAT, ...), and RevitLinkInstance: nested
        # links are offered as their own entry in the document selection, so
        # they are never counted here
        if isinstance(element, (ImportInstance, RevitLinkInstance)):
            return True

        # Check the element really belongs to the document being processed
        if hasattr(element, 'Document') and element.Document:
            if get_cached_doc_key(element.Document, cache) != cache["doc_key"]:
                return True

        # "Import Symbol" families: cached per type
        return get_type_entry(element, source_doc, cache)["is_import"]

    except:
        pass

    return False


# Categories to exclude from the export
EXCLUDED_CATEGORY_NAMES = [
    # Annotations and views
    "OST_Lines", "OST_Cameras", "OST_Views", "OST_Viewers", "OST_Sheets",
    "OST_ScheduleGraphics", "OST_Schedules", "OST_TitleBlocks", "OST_Grids",
    "OST_Levels", "OST_ReferencePlanes", "OST_MatchLine", "OST_ScopeBoxes",
    "OST_DetailComponents", "OST_Annotations", "OST_GenericAnnotation",
    "OST_TextNotes", "OST_Dimensions", "OST_Tags",
    # Revit Links
    "OST_RvtLinks",
    # CAD/DWG Imports
    "OST_ImportObjectStyles", "OST_DWGRefPlanes", "OST_IOSSketchGrid",
    "OST_IOSModelGroups",
    # Other import categories
    "OST_IOS_GeoSite", "OST_PointClouds", "OST_Coordination_Model",
    # Project data
    "OST_ProjectInformation",
]

# Names matching a link or an import, checked against category names
LINK_KEYWORDS = [".dwg", ".dxf", ".dgn", ".sat", ".skp", "import", "link",
                 "cad", ".rvt"]


def get_excluded_categories():
    """Builds the list of excluded categories safely."""
    resolved = [safe_get_builtin_category(name)
                for name in EXCLUDED_CATEGORY_NAMES]
    return [cat for cat in resolved if cat is not None]


# Standard output columns (in order)
STANDARD_COLUMNS = [
    "Source_Document",
    "ID",
    "Category",
    "Name",
    "Length_m",
    "Width_m",
    "Height_m",
    "Depth_m",
    "Thickness_m",
    "Diameter_m",
    "Perimeter_m",
    "Area_m2",
    "Volume_m3",
]


def get_param_value_from_element(element, builtin_param):
    """
    Extracts the value of a BuiltInParameter from an element.
    Returns the numeric value (Double) or None if not found/invalid.
    """
    try:
        param = element.get_Parameter(builtin_param)
        if param and param.HasValue:
            if param.StorageType == StorageType.Double:
                return param.AsDouble()
            elif param.StorageType == StorageType.Integer:
                return float(param.AsInteger())
    except:
        pass
    return None


def has_numeric_param(element, builtin_param):
    """True when the BuiltInParameter exists on the element with a number."""
    try:
        param = element.get_Parameter(builtin_param)
    except:
        return False

    if param is None:
        return False

    try:
        return param.StorageType in (StorageType.Double, StorageType.Integer)
    except:
        return True


def probe_type(element, elem_type, params_map):
    """
    Everything that can be resolved once per type, for every quantity:
      - inst_bips: the BuiltInParameters that exist on the instance with a
        numeric storage type. Existence depends on category and type, never on
        the single instance, so all the other instances reuse this list.
      - type_geom: the value read on the type, identical for every instance.
    """
    inst_bips = {}
    type_geom = {}

    for geom_name, builtin_list in params_map.items():
        available = []
        type_value = None
        for bp in builtin_list:
            if has_numeric_param(element, bp):
                available.append(bp)
            if type_value is None and elem_type:
                found = get_param_value_from_element(elem_type, bp)
                if found is not None and found != 0:
                    type_value = found
        inst_bips[geom_name] = available
        type_geom[geom_name] = type_value

    return inst_bips, type_geom


def make_doc_cache(source_doc, params_map):
    """
    Per-document caches. Built inside the run (never as a module global) because
    the extension is rocket mode compatible and the Python engine is reused
    between two launches of the tool.
    """
    return {
        "params_map": params_map,
        "doc_key": get_doc_key(source_doc),
        "doc_keys": {},   # document -> doc key
        "types": {},      # type id (or category+class) -> type entry
    }


def get_type_entry(element, source_doc, cache):
    """
    Returns the cached entry of the element type, building it on first use.
    Everything it holds depends on the type only (or, for type-less elements
    such as rooms, on category + element class), so it is computed once and
    shared by every instance.
    """
    type_id = None
    try:
        type_id = element.GetTypeId()
        if type_id is not None and type_id == ElementId.InvalidElementId:
            type_id = None
    except:
        type_id = None

    if type_id is not None:
        key = get_element_id_value(type_id)
    else:
        # Type-less elements: the parameter set follows category and class
        key = ("notype", safe_category_id(element), element.__class__.__name__)

    entry = cache["types"].get(key)
    if entry is not None:
        return entry

    elem_type = None
    if type_id is not None:
        try:
            elem_type = source_doc.GetElement(type_id)
        except:
            elem_type = None

    inst_bips, type_geom = probe_type(element, elem_type, cache["params_map"])
    entry = {
        "elem_type": elem_type,
        "is_import": is_import_family(element, elem_type),
        "inst_bips": inst_bips,
        "type_geom": type_geom,
        "extra_inst": {},   # parameter name -> does it exist on the instance?
        "extra_type": {},   # parameter name -> (value, unit label)
    }
    cache["types"][key] = entry
    return entry


# Metric unit label reported for each converted extra parameter
EXTRA_PARAM_UNIT_LABELS = {
    "length": "m",
    "area": "m2",
    "volume": "m3",
    "angle": "deg",
}


# Spec key -> (SpecTypeId / ParameterType name, UnitTypeId name,
#              DisplayUnitType name, fallback conversion factor)
SPEC_TABLE = {
    "length": ("Length", "Meters", "DUT_METERS", FEET_TO_METERS),
    "area": ("Area", "SquareMeters", "DUT_SQUARE_METERS", SQFEET_TO_SQMETERS),
    "volume": ("Volume", "CubicMeters", "DUT_CUBIC_METERS",
               CUBICFEET_TO_CUBICMETERS),
    "angle": ("Angle", "Degrees", "DUT_DECIMAL_DEGREES", 1.0),
}

# Lazily resolved lookup tables. They only hold Revit API constants (never
# document data), so keeping them at module level is rocket mode safe.
_API_TABLES = {}


def api_table(name, build):
    """Resolves a table of Revit API constants once, then reuses it."""
    if name not in _API_TABLES:
        try:
            _API_TABLES[name] = build()
        except:
            _API_TABLES[name] = {}
    return _API_TABLES[name]


def spec_key_by_type_id():
    """{ ForgeTypeId string: spec key } resolved once (Revit 2021+)."""
    def build():
        table = {}
        for key, names in SPEC_TABLE.items():
            spec = getattr(SpecTypeId, names[0], None)
            if spec is not None:
                table[spec.TypeId] = key
        return table
    return api_table("spec_type_id", build)


def unit_by_spec_key(api_name, index):
    """
    { spec key: unit constant } for UnitTypeId (Revit 2021+) or for
    DisplayUnitType (Revit <= 2020). Empty when that API does not exist.
    """
    def build():
        holder = globals().get(api_name)
        if holder is None:
            return {}
        table = {}
        for key, names in SPEC_TABLE.items():
            unit = getattr(holder, names[index], None)
            if unit is not None:
                table[key] = unit
        return table
    return api_table(api_name, build)


def get_param_spec_key(param):
    """
    Returns "length"/"area"/"volume"/"angle" for a Double parameter whose data
    type carries a unit, None for the unitless ones (Number, Currency, ...).
    Works both on Revit 2021+ (SpecTypeId/ForgeTypeId) and on the older
    versions (ParameterType).
    """
    try:
        definition = param.Definition
    except:
        return None

    if definition is None:
        return None

    # Revit 2021+ : ForgeTypeId compared through its TypeId string
    try:
        data_type = definition.GetDataType()
        if data_type is not None:
            return spec_key_by_type_id().get(data_type.TypeId)
    except:
        pass

    # Revit <= 2020 : ParameterType
    try:
        param_type = definition.ParameterType
        for key, names in SPEC_TABLE.items():
            expected = getattr(ParameterType, names[0], None)
            if expected is not None and param_type == expected:
                return key
    except:
        pass

    return None


def convert_from_internal_units(value, spec_key):
    """
    Converts a value from Revit internal units to the metric unit matching
    spec_key, through UnitTypeId (Revit 2021+) or DisplayUnitType (Revit <=
    2020). Unitless parameters (spec_key None) are returned untouched.
    """
    if spec_key is None:
        return value

    for api_name, index in (("UnitTypeId", 1), ("DisplayUnitType", 2)):
        unit = unit_by_spec_key(api_name, index).get(spec_key)
        if unit is not None:
            try:
                return UnitUtils.ConvertFromInternalUnits(value, unit)
            except:
                pass

    # Last resort: same constants used for the standard columns
    return value * SPEC_TABLE[spec_key][3]


def read_param_value(param, param_name, units_seen):
    """
    Reads a parameter, converting the unit-bearing Double values from Revit
    internal units to metric ones. Returns None when there is no value.
    units_seen, when given, collects the unit label used per parameter name.
    """
    if param is None:
        return None

    try:
        if param.HasValue:
            storage = param.StorageType
            if storage == StorageType.Double:
                spec_key = get_param_spec_key(param)
                if units_seen is not None and spec_key is not None:
                    units_seen[param_name] = EXTRA_PARAM_UNIT_LABELS[spec_key]
                return convert_from_internal_units(param.AsDouble(), spec_key)
            elif storage == StorageType.Integer:
                return param.AsInteger()
            elif storage == StorageType.String:
                return param.AsString()
            elif storage == StorageType.ElementId:
                return param.AsValueString()
        return param.AsValueString()
    except:
        return None


def lookup_parameter(element, param_name):
    """LookupParameter by name, None when the element or the parameter is missing."""
    try:
        return element.LookupParameter(param_name) if element else None
    except:
        return None


def get_param_value_by_name(element, param_name, entry, units_seen=None):
    """
    Extracts the value of a parameter via LookupParameter: instance, then type.
    LookupParameter is a linear search by name, so the type level value is read
    once per type (entry["extra_type"]) instead of once per element, and the
    instance search is skipped altogether for the parameters that do not exist
    on the instances of this type (entry["extra_inst"]).
    """
    extra_inst = entry["extra_inst"]

    if extra_inst.get(param_name, True):
        param = lookup_parameter(element, param_name)
        if param_name not in extra_inst:
            extra_inst[param_name] = param is not None
        value = read_param_value(param, param_name, units_seen)
        if value is not None:
            return value

    # Type level value, cached (entry["elem_type"] comes from the document the
    # element belongs to, links included)
    extra_type = entry["extra_type"]
    if param_name not in extra_type:
        type_units = {}
        value = read_param_value(
            lookup_parameter(entry["elem_type"], param_name),
            param_name, type_units)
        extra_type[param_name] = (value, type_units.get(param_name))

    value, unit = extra_type[param_name]
    if unit is not None and units_seen is not None:
        units_seen[param_name] = unit

    return value


# Conversion factor of each geometric column, resolved once instead of matching
# the parameter name against a keyword list for every single cell
GEOMETRIC_CONVERSION_FACTORS = {
    "Length": FEET_TO_METERS,
    "Width": FEET_TO_METERS,
    "Height": FEET_TO_METERS,
    "Depth": FEET_TO_METERS,
    "Thickness": FEET_TO_METERS,
    "Diameter": FEET_TO_METERS,
    "Perimeter": FEET_TO_METERS,
    "Area": SQFEET_TO_SQMETERS,
    "Volume": CUBICFEET_TO_CUBICMETERS,
}

# (CSV column, geometric quantity), in output order
GEOMETRIC_COLUMNS = [
    ("Length_m", "Length"),
    ("Width_m", "Width"),
    ("Height_m", "Height"),
    ("Depth_m", "Depth"),
    ("Thickness_m", "Thickness"),
    ("Diameter_m", "Diameter"),
    ("Perimeter_m", "Perimeter"),
    ("Area_m2", "Area"),
    ("Volume_m3", "Volume"),
]


def convert_to_meters(value, geom_name):
    """Converts the value to metric units based on the quantity it belongs to."""
    if value is None:
        return None

    try:
        value = float(value)
    except:
        return value

    return round(value * GEOMETRIC_CONVERSION_FACTORS.get(geom_name, 1.0), 3)


def get_element_name(element, entry=None):
    """Gets the element Name property."""
    try:
        # Try to get the Name property directly from the element
        if hasattr(element, 'Name') and element.Name:
            return element.Name

        # For some elements, Name might be on the type (already cached)
        elem_type = entry["elem_type"] if entry else \
            element.Document.GetElement(element.GetTypeId())
        if elem_type and hasattr(elem_type, 'Name') and elem_type.Name:
            return elem_type.Name

        return ""
    except:
        return ""


def get_model_categories(doc, excluded_categories):
    """
    Gets all valid Model categories from the document, as { category id: category }.
    Whether the category holds elements is decided by the single collector pass
    of collect_elements_by_category(), not by one collector per category.
    """
    categories = {}

    for cat in doc.Settings.Categories:
        try:
            # Only Model categories
            if cat.CategoryType != CategoryType.Model:
                continue

            # Exclude categories in the exclusion list
            try:
                bic = System.Enum.ToObject(BuiltInCategory, get_element_id_value(cat.Id))
                if bic in excluded_categories:
                    continue
            except:
                pass

            # Exclude categories containing link/import keywords in name
            cat_name_lower = cat.Name.lower()
            if any(x in cat_name_lower for x in LINK_KEYWORDS):
                continue

            categories[get_element_id_value(cat.Id)] = cat
        except:
            continue

    return categories


def build_categories_filter(categories):
    """
    ElementMulticategoryFilter on the wanted categories, so the single pass only
    touches the elements we are going to export. Returns None when the filter
    cannot be built (the caller then scans everything and filters in Python).
    """
    try:
        from System.Collections.Generic import List
        category_ids = List[ElementId]()
        for cat in categories.values():
            category_ids.Add(cat.Id)
        if category_ids.Count == 0:
            return None
        return ElementMulticategoryFilter(category_ids)
    except:
        # Some categories cannot be used in a filter: rather than dropping them
        # (and losing their elements) scan everything and filter in Python
        return None


def fill_category_buckets(collector, categories, buckets):
    """Sorts the elements of a collector in one bucket per wanted category."""
    for elem in collector:
        cat_id_value = safe_category_id(elem)
        if cat_id_value is None:
            continue

        if cat_id_value in buckets:
            bucket = buckets[cat_id_value]
        else:
            bucket = [] if cat_id_value in categories else None
            buckets[cat_id_value] = bucket

        if bucket is not None:
            bucket.append(elem)


def collect_elements_by_category(doc, categories):
    """
    Single pass over the model elements, bucketed by category, instead of two
    collectors per category (one to count, one to fetch).
    Returns [(category name, [elements])] sorted by category name, as before.
    """
    buckets = {}

    categories_filter = build_categories_filter(categories)
    if categories_filter is not None:
        try:
            fill_category_buckets(
                FilteredElementCollector(doc).WhereElementIsNotElementType()
                .WherePasses(categories_filter),
                categories, buckets)
        except:
            # The filter failed while iterating: fall back to the full scan
            buckets = {}
            categories_filter = None

    if categories_filter is None:
        try:
            fill_category_buckets(
                FilteredElementCollector(doc).WhereElementIsNotElementType(),
                categories, buckets)
        except:
            return []

    found = [(categories[cat_id_value].Name, elements)
             for cat_id_value, elements in buckets.items() if elements]

    return sorted(found, key=lambda pair: pair[0])


def get_assembly_instances(doc):
    """Gets all Assembly instances from the document."""
    assemblies = []
    try:
        collector = FilteredElementCollector(doc).OfClass(AssemblyInstance)
        assemblies = list(collector)
    except:
        pass
    return assemblies


def extract_geometric_data(element, entry):
    """
    Extracts all geometric data from an element.
    Uses the logic: first instance, then type, otherwise empty.
    Only the BuiltInParameters that exist on this type are probed on the
    instance, and the type level values come from the type cache.
    """
    data = {}
    type_geom = entry["type_geom"]

    for geom_name, builtin_list in entry["inst_bips"].items():
        value = None
        for bp in builtin_list:
            found = get_param_value_from_element(element, bp)
            if found is not None and found != 0:
                value = found
                break
        if value is None:
            value = type_geom.get(geom_name)
        data[geom_name] = value

    return data


def get_extra_params_from_user():
    """Shows a dialog to get extra parameters from user using rpw TextBox."""
    try:
        from rpw.ui.forms import TextInput
        result = TextInput("Extra Parameters",
                          default="",
                          description='Add extra parameters to extract separated by ";"')
        return result if result else ""
    except ImportError:
        # Fallback to pyrevit forms if rpw is not available
        return forms.ask_for_string(
            prompt='Add extra parameters to extract separated by ";"',
            title="Extra Parameters",
            default=""
        )


def strip_rvt(name):
    """Removes the .rvt extension from a file name."""
    return name[:-4] if name.lower().endswith('.rvt') else name


def get_central_model_name(doc):
    """
    Gets the central model filename (without local copy suffix).
    If it's a central file or non-workshared, returns the document title.
    If it's a local copy, returns the central model name.
    """
    try:
        if doc.IsWorkshared:
            central_path = doc.GetWorksharingCentralModelPath()
            central_name = ModelPathUtils.ConvertModelPathToUserVisiblePath(
                central_path) if central_path else None
            if central_name:
                return strip_rvt(os.path.basename(central_name))
    except:
        pass

    # Fallback: the document title
    return strip_rvt(doc.Title)


# =============================================================================
# DOCUMENT SELECTION (active document + loaded Revit links, nested included)
# =============================================================================

def make_unique_label(label, used_labels):
    """Ensures every entry of the selection list has a distinct label."""
    if label not in used_labels:
        used_labels.add(label)
        return label

    index = 2
    while "{} ({})".format(label, index) in used_labels:
        index += 1
    unique = "{} ({})".format(label, index)
    used_labels.add(unique)
    return unique


def walk_links(parent_doc, prefix, docs, seen_keys, used_labels, depth):
    """
    Recursively collects the loaded Revit links of parent_doc, appending
    (label, document) tuples to docs. Unloaded links and documents already
    collected are skipped; recursion is capped at MAX_LINK_DEPTH.
    """
    if depth >= MAX_LINK_DEPTH:
        return

    try:
        link_instances = FilteredElementCollector(parent_doc).OfClass(RevitLinkInstance).ToElements()
    except:
        return

    for link in link_instances:
        try:
            link_doc = link.GetLinkDocument()
            if link_doc is None:
                continue  # unloaded link / not found

            key = get_doc_key(link_doc)
            if key in seen_keys:
                continue  # several instances, or same link under different parents
            seen_keys.add(key)

            link_name = link.Name.split(" : ")[0]
            label = make_unique_label(prefix + link_name, used_labels)
            docs.append((label, link_doc))

            walk_links(link_doc, label + " > ", docs, seen_keys, used_labels, depth + 1)
        except:
            continue


def collect_available_docs(host_doc):
    """
    Returns the documents that can be processed as (label, document) tuples.
    The active document is always the first entry.
    """
    used_labels = set()
    host_label = make_unique_label(
        "{}  [documento attivo]".format(get_central_model_name(host_doc)), used_labels)
    docs = [(host_label, host_doc)]

    seen_keys = set()
    seen_keys.add(get_doc_key(host_doc))

    walk_links(host_doc, "", docs, seen_keys, used_labels, 0)
    return docs


def select_documents(host_doc):
    """
    Asks the user which documents to process. With no loaded link the dialog is
    skipped and the active document is used, as before.
    """
    available = collect_available_docs(host_doc)

    if len(available) == 1:
        return available

    doc_map = dict(available)
    labels = [label for label, _ in available]

    chosen = forms.SelectFromList.show(
        labels,
        title="Seleziona i documenti da computare",
        button_name="Continua",
        multiselect=True,
    )

    if not chosen:
        script.exit()

    return [(label, doc_map[label]) for label in chosen if label in doc_map]


# =============================================================================
# DATA EXTRACTION
# =============================================================================

def build_row(elem, source_doc, doc_label, category_name, extra_params,
              units_seen, cache):
    """Builds the CSV row of a single element of source_doc."""
    # Element type and everything that only depends on it, from the cache
    entry = get_type_entry(elem, source_doc, cache)

    geo_data = extract_geometric_data(elem, entry)

    row = {
        "Source_Document": doc_label,
        "ID": get_element_id_value(elem.Id),
        "Category": category_name,
        "Name": get_element_name(elem, entry),
    }

    for column, geom_name in GEOMETRIC_COLUMNS:
        value = geo_data.get(geom_name)
        row[column] = convert_to_meters(value, geom_name) if value else ""

    # Extra parameters (unit-bearing ones are converted to metric units)
    for param_name in extra_params:
        value = get_param_value_by_name(elem, param_name, entry, units_seen)
        if value is not None:
            try:
                value = round(float(value), 3)
            except:
                pass
        row[param_name] = value if value is not None else ""

    return row


def advance_progress(progress, step_label, done, total):
    """
    Advances the progress bar, when there is one.
    Returns True if the user asked to cancel.
    """
    if progress is None:
        return False

    pb = progress["pb"]
    try:
        pb.title = "{} [{}/{}] - {} ({}/{})".format(
            progress["doc_label"], progress["doc_index"], progress["doc_total"],
            step_label, done, total)
        pb.update_progress(done, total)
        return bool(pb.cancelled)
    except:
        return False


def process_document(doc_label, source_doc, params_map, excluded_categories,
                     extra_params, units_seen, output, progress=None):
    """
    Extracts all rows from a single document.
    Returns (rows, processed_count, skipped_count, cancelled).
    Messages are collected and printed in one block: the pyRevit output window
    renders HTML, and one print per category is slow on big models.
    """
    rows = []
    processed_count = 0
    skipped_count = 0
    messages = []

    cache = make_doc_cache(source_doc, params_map)

    categories = get_model_categories(source_doc, excluded_categories)
    groups = collect_elements_by_category(source_doc, categories)
    messages.append("Found **{}** categories with elements".format(len(groups)))

    # Assemblies need a dedicated collector, then they are one group like the
    # others, processed last
    assemblies = get_assembly_instances(source_doc)
    if assemblies:
        groups.append(("Assemblies", assemblies))

    total = max(sum(len(elements) for _, elements in groups), 1)
    done = 0
    cancelled = advance_progress(progress, "start", 0, total)

    for group_name, elements in groups:
        if cancelled:
            break

        try:
            group_processed = 0
            group_skipped = 0

            for elem in elements:
                try:
                    # Skip elements from links or imports
                    if is_element_from_link_or_import(elem, source_doc, cache):
                        group_skipped += 1
                        skipped_count += 1
                        continue

                    rows.append(build_row(elem, source_doc, doc_label, group_name,
                                          extra_params, units_seen, cache))
                    processed_count += 1
                    group_processed += 1

                except Exception as e:
                    continue

            if group_processed > 0 or group_skipped > 0:
                msg = "- **{}**: {} elements".format(group_name, group_processed)
                if group_skipped > 0:
                    msg += " ({} skipped - links/imports)".format(group_skipped)
                messages.append(msg)

        except Exception as e:
            messages.append("  - Error in category {}: {}".format(group_name, str(e)))

        done += len(elements)
        cancelled = advance_progress(progress, group_name, done, total)

    output.print_md("\n".join(messages))

    return rows, processed_count, skipped_count, cancelled


def csv_cell(value):
    """Formats a value for the CSV: decimal comma, utf-8 bytes for the text."""
    if value is None or value == "":
        return ""
    if isinstance(value, float):
        return str(value).replace('.', ',')
    try:
        return str(value).encode('utf-8')
    except:
        return str(value)


def write_csv(filepath, headers, rows):
    """Writes the rows to the CSV file, ';' separated."""
    # Use binary mode to prevent extra blank rows on Windows (IronPython 2)
    with open(filepath, 'wb') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=headers, delimiter=';',
                                lineterminator='\n')
        writer.writeheader()
        for row in rows:
            writer.writerow(dict((k, csv_cell(v)) for k, v in row.items()))


def print_summary(output, filepath, per_doc_counts, processed_count,
                  skipped_count, extra_params, extra_param_units):
    """Prints the final report in the pyRevit output window."""
    lines = ["## Export completed!"]

    for doc_label, doc_processed in per_doc_counts:
        lines.append("- **{}**: {} elements".format(doc_label, doc_processed))
    lines.append("- **Elements processed (total):** {}".format(processed_count))

    if skipped_count > 0:
        lines.append("- **Elements skipped (links/imports):** {}".format(skipped_count))

    if extra_param_units:
        legend = ", ".join("{} [{}]".format(name, unit)
                           for name, unit in sorted(extra_param_units.items()))
        lines.append("- **Extra parameters converted to metric units:** {}".format(legend))

    unitless = [p for p in extra_params if p not in extra_param_units]
    if unitless:
        lines.append("- **Extra parameters exported as-is (unitless or "
                     "non numeric):** {}".format(", ".join(unitless)))

    lines.append("- **File saved:** {}".format(filepath))

    output.print_md("\n---")
    output.print_md("\n".join(lines))


def ask_export_settings(doc):
    """
    Asks the user what to export and where.
    Returns (selected documents, extra parameter names, csv file path).
    """
    # Which documents to process (active document and/or links)
    selected_docs = select_documents(doc)
    if not selected_docs:
        forms.alert("No document selected. Operation cancelled.", exitscript=True)

    # Extra parameters to extract
    extra_params_input = get_extra_params_from_user()
    if extra_params_input is False or extra_params_input is None:
        script.exit()

    extra_params = [p.strip() for p in extra_params_input.split(";")
                    if p.strip()] if extra_params_input else []

    # Destination folder
    output_folder = forms.pick_folder(title="Select destination folder")
    if not output_folder:
        forms.alert("No folder selected. Operation cancelled.", exitscript=True)

    # File name: YYMMDD_HHMMSS_QTO_NomeFile[_MULTI].csv
    filename = "{}_QTO_{}{}.csv".format(
        datetime.now().strftime("%y%m%d_%H%M%S"),
        get_central_model_name(doc).replace(" ", "_"),
        "_MULTI" if len(selected_docs) > 1 else "")

    return selected_docs, extra_params, os.path.join(output_folder, filename)


def main():
    """Main script function."""
    doc = revit.doc

    # Build quantities -> BuiltInParameter mapping
    GEOMETRIC_PARAMS_MAP = build_geometric_params_map()
    EXCLUDED_CATEGORIES = get_excluded_categories()

    # Steps 1 to 3: documents, extra parameters, destination file
    selected_docs, extra_params, filepath = ask_export_settings(doc)

    # Step 4: Collect data
    output = script.get_output()
    output.print_md("# Quantity Takeoff in progress...")
    output.print_md("Documents selected: **{}**".format(len(selected_docs)))

    # Prepare column headers
    headers = list(STANDARD_COLUMNS)
    for param in extra_params:
        headers.append(param)

    rows = []
    processed_count = 0
    skipped_count = 0
    per_doc_counts = []
    extra_param_units = {}

    cancelled = False

    with forms.ProgressBar(title="Quantity Takeoff", cancellable=True) as pb:
        for doc_index, (doc_label, source_doc) in enumerate(selected_docs, 1):
            output.print_md("\n## Document: **{}**".format(doc_label))
            progress = {
                "pb": pb,
                "doc_label": doc_label,
                "doc_index": doc_index,
                "doc_total": len(selected_docs),
            }
            try:
                doc_rows, doc_processed, doc_skipped, cancelled = process_document(
                    doc_label, source_doc, GEOMETRIC_PARAMS_MAP, EXCLUDED_CATEGORIES,
                    extra_params, extra_param_units, output, progress)
            except Exception as e:
                output.print_md("  - Error processing document {}: {}".format(doc_label, str(e)))
                continue

            rows.extend(doc_rows)
            processed_count += doc_processed
            skipped_count += doc_skipped
            per_doc_counts.append((doc_label, doc_processed))

            if cancelled:
                break

    if cancelled:
        output.print_md("\n---")
        output.print_md("## Operation cancelled by the user - no file written.")
        return

    # Step 5: Write CSV file
    output.print_md("\n## Writing CSV file...")

    try:
        write_csv(filepath, headers, rows)
        print_summary(output, filepath, per_doc_counts, processed_count,
                      skipped_count, extra_params, extra_param_units)
    except Exception as e:
        forms.alert("Error writing file:\n{}".format(str(e)), exitscript=True)


if __name__ == "__main__":
    main()
