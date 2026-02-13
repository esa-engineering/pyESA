# -*- coding: utf-8 -*-
"""Quantity Takeoff - Extract geometric quantities by category"""

__title__ = "Quantity\nTakeoff"
__author__ = "PyRevit Script"
__doc__ = "Extracts geometric quantities (length, area, volume, thickness, etc.) " \
          "for all model elements, organized by category."

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

def build_geometric_params_map():
    """
    Builds the mapping between geometric quantities and possible BuiltInParameters.
    Returns a dictionary: { "Length": [list of valid BuiltInParameters], ... }
    """
    
    params_map = {
        "Length": [],
        "Width": [],
        "Height": [],
        "Depth": [],
        "Thickness": [],
        "Diameter": [],
        "Perimeter": [],
        "Area": [],
        "Volume": [],
    }
    
    # LENGTH - All possible length parameters
    length_params = [
        "CURVE_ELEM_LENGTH",
        "INSTANCE_LENGTH_PARAM",
        "STRUCTURAL_FRAME_CUT_LENGTH",
        "STAIRS_ACTUAL_RUN_LENGTH",
        "STAIRS_RUN_ACTUAL_RUN_LENGTH",
        "RAMP_ATTR_LENGTH",
        "REBAR_ELEM_LENGTH",
        "STRUCTURAL_FOUNDATION_LENGTH",
        "PATH_REIN_LENGTH_1",
    ]
    for p in length_params:
        bp = safe_get_builtin(p)
        if bp is not None:
            params_map["Length"].append(bp)
    
    # WIDTH - All possible width parameters
    width_params = [
        "WALL_ATTR_WIDTH_PARAM",
        "DOOR_WIDTH",
        "WINDOW_WIDTH",
        "FAMILY_WIDTH_PARAM",
        "STAIRS_ATTR_TREAD_WIDTH",
        "STAIRS_RUN_ACTUAL_RUN_WIDTH",
        "RAMP_ATTR_WIDTH",
        "RBS_CURVE_WIDTH_PARAM",
        "RBS_CABLETRAY_WIDTH_PARAM",
        "CURTAIN_WALL_PANELS_WIDTH",
        "STRUCTURAL_FOUNDATION_WIDTH",
        "CASEWORK_WIDTH",
    ]
    for p in width_params:
        bp = safe_get_builtin(p)
        if bp is not None:
            params_map["Width"].append(bp)
    
    # HEIGHT - All possible height parameters
    height_params = [
        "WALL_USER_HEIGHT_PARAM",
        "DOOR_HEIGHT",
        "WINDOW_HEIGHT",
        "FAMILY_HEIGHT_PARAM",
        "STAIRS_ACTUAL_RISER_HEIGHT",
        "RBS_CURVE_HEIGHT_PARAM",
        "RBS_CABLETRAY_HEIGHT_PARAM",
        "CURTAIN_WALL_PANELS_HEIGHT",
        "CASEWORK_HEIGHT",
        "INSTANCE_HEIGHT_PARAM",
    ]
    for p in height_params:
        bp = safe_get_builtin(p)
        if bp is not None:
            params_map["Height"].append(bp)
    
    # DEPTH - All possible depth parameters
    depth_params = [
        "FAMILY_DEPTH_PARAM",
        "CASEWORK_DEPTH",
    ]
    for p in depth_params:
        bp = safe_get_builtin(p)
        if bp is not None:
            params_map["Depth"].append(bp)
    
    # THICKNESS - All possible thickness parameters
    thickness_params = [
        "FLOOR_ATTR_THICKNESS_PARAM",
        "ROOF_ATTR_DEFAULT_THICKNESS_PARAM",
        "ROOF_ATTR_THICKNESS_VALUE",
        "CEILING_THICKNESS",
        "WALL_ATTR_WIDTH_PARAM",
        "BUILDINGPAD_THICKNESS",
        "SLAB_EDGE_THICKNESS",
    ]
    for p in thickness_params:
        bp = safe_get_builtin(p)
        if bp is not None:
            params_map["Thickness"].append(bp)
    
    # DIAMETER - All possible diameter parameters
    diameter_params = [
        "RBS_PIPE_DIAMETER_PARAM",
        "RBS_CONDUIT_DIAMETER_PARAM",
        "RBS_CURVE_DIAMETER_PARAM",
        "REBAR_BAR_DIAMETER",
    ]
    for p in diameter_params:
        bp = safe_get_builtin(p)
        if bp is not None:
            params_map["Diameter"].append(bp)
    
    # PERIMETER - All possible perimeter parameters
    perimeter_params = [
        "HOST_PERIMETER_COMPUTED",
        "ROOM_PERIMETER",
    ]
    for p in perimeter_params:
        bp = safe_get_builtin(p)
        if bp is not None:
            params_map["Perimeter"].append(bp)
    
    # AREA - All possible area parameters
    area_params = [
        "HOST_AREA_COMPUTED",
        "ROOM_AREA",
        "RBS_CURVE_SURFACE_AREA",
        "MASS_GROSS_SURFACE_AREA",
        "MASS_SURFACE_AREA",
    ]
    for p in area_params:
        bp = safe_get_builtin(p)
        if bp is not None:
            params_map["Area"].append(bp)
    
    # VOLUME - All possible volume parameters
    volume_params = [
        "HOST_VOLUME_COMPUTED",
        "ROOM_VOLUME",
        "MASS_GROSS_VOLUME",
        "MASS_VOLUME",
    ]
    for p in volume_params:
        bp = safe_get_builtin(p)
        if bp is not None:
            params_map["Volume"].append(bp)
    
    return params_map


def is_element_from_link_or_import(element):
    """
    Checks if an element comes from a link (Revit or CAD/DWG) or is an import.
    Returns True if the element is from a link/import, False otherwise.
    """
    try:
        # Check if it's an ImportInstance (DWG, DXF, SAT, etc.)
        if isinstance(element, ImportInstance):
            return True
        
        # Check if it's a RevitLinkInstance
        if isinstance(element, RevitLinkInstance):
            return True
        
        # Check if element's document is linked
        if hasattr(element, 'Document') and element.Document:
            if element.Document.IsLinked:
                return True
        
        # Check category for import/link related keywords
        if element.Category:
            cat_name = element.Category.Name.lower()
            # Check for DWG, import, link keywords
            if any(x in cat_name for x in [".dwg", ".dxf", ".dgn", ".sat", ".skp", 
                                            "import", "link", "cad"]):
                return True
            
            # Check if it's an "Import Symbol" family
            try:
                if hasattr(element, 'Symbol') and element.Symbol:
                    family_name = element.Symbol.Family.Name if element.Symbol.Family else ""
                    if "import" in family_name.lower():
                        return True
            except:
                pass
        
        # Check family name for import symbols
        try:
            elem_type = element.Document.GetElement(element.GetTypeId())
            if elem_type:
                type_name = elem_type.Name if hasattr(elem_type, 'Name') else ""
                family_param = elem_type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                if family_param and family_param.HasValue:
                    family_name = family_param.AsString()
                    if family_name and "import" in family_name.lower():
                        return True
        except:
            pass
            
    except:
        pass
    
    return False


# Categories to exclude
def get_excluded_categories():
    """Builds the list of excluded categories safely."""
    excluded = []
    excluded_names = [
        # Annotations and views
        "OST_Lines",
        "OST_Cameras",
        "OST_Views",
        "OST_Viewers",
        "OST_Sheets",
        "OST_ScheduleGraphics",
        "OST_Schedules",
        "OST_TitleBlocks",
        "OST_Grids",
        "OST_Levels",
        "OST_ReferencePlanes",
        "OST_MatchLine",
        "OST_ScopeBoxes",
        "OST_DetailComponents",
        "OST_Annotations",
        "OST_GenericAnnotation",
        "OST_TextNotes",
        "OST_Dimensions",
        "OST_Tags",
        # Revit Links
        "OST_RvtLinks",
        # CAD/DWG Imports
        "OST_ImportObjectStyles",
        "OST_DWGRefPlanes",
        "OST_IOSSketchGrid",
        "OST_IOSModelGroups",
        # Other import categories
        "OST_IOS_GeoSite",
        "OST_PointClouds",
        "OST_Coordination_Model",
        # Project data
        "OST_ProjectInformation",
    ]
    for name in excluded_names:
        cat = safe_get_builtin_category(name)
        if cat is not None:
            excluded.append(cat)
    return excluded


# Standard output columns (in order)
STANDARD_COLUMNS = [
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


def get_geometric_value(element, elem_type, builtin_params_list):
    """
    Searches for a geometric value trying a list of BuiltInParameters.
    First searches on instance, then on type.
    Returns the first value found, or None if none found.
    """
    # First try on instance
    for bp in builtin_params_list:
        value = get_param_value_from_element(element, bp)
        if value is not None and value != 0:
            return value
    
    # Then try on type
    if elem_type:
        for bp in builtin_params_list:
            value = get_param_value_from_element(elem_type, bp)
            if value is not None and value != 0:
                return value
    
    return None


def get_param_value_by_name(element, param_name):
    """Extracts the value of a parameter via LookupParameter."""
    try:
        param = element.LookupParameter(param_name)
        if param and param.HasValue:
            if param.StorageType == StorageType.Double:
                return param.AsDouble()
            elif param.StorageType == StorageType.Integer:
                return param.AsInteger()
            elif param.StorageType == StorageType.String:
                return param.AsString()
            elif param.StorageType == StorageType.ElementId:
                return param.AsValueString()
        elif param:
            return param.AsValueString()
    except:
        pass
    
    # Also try on type
    try:
        elem_type = element.Document.GetElement(element.GetTypeId())
        if elem_type:
            param = elem_type.LookupParameter(param_name)
            if param and param.HasValue:
                if param.StorageType == StorageType.Double:
                    return param.AsDouble()
                elif param.StorageType == StorageType.Integer:
                    return param.AsInteger()
                elif param.StorageType == StorageType.String:
                    return param.AsString()
                elif param.StorageType == StorageType.ElementId:
                    return param.AsValueString()
            elif param:
                return param.AsValueString()
    except:
        pass
    
    return None


def convert_to_meters(value, param_name):
    """Converts the value to meters based on parameter type."""
    if value is None:
        return None
    
    try:
        value = float(value)
    except:
        return value
    
    param_name_lower = param_name.lower()
    
    # Linear length parameters
    if any(x in param_name_lower for x in ["length", "width", "height", "depth", 
                                            "thickness", "diameter", "perimeter"]):
        return round(value * FEET_TO_METERS, 3)
    # Area parameters
    elif "area" in param_name_lower:
        return round(value * SQFEET_TO_SQMETERS, 3)
    # Volume parameters
    elif "volume" in param_name_lower:
        return round(value * CUBICFEET_TO_CUBICMETERS, 3)
    else:
        return round(value, 3)


def get_element_name(element):
    """Gets the element Name property."""
    try:
        # Try to get the Name property directly from the element
        if hasattr(element, 'Name') and element.Name:
            return element.Name
        
        # For some elements, Name might be on the type
        elem_type = element.Document.GetElement(element.GetTypeId())
        if elem_type and hasattr(elem_type, 'Name') and elem_type.Name:
            return elem_type.Name
        
        return ""
    except:
        return ""


def get_model_categories(doc, excluded_categories):
    """Gets all valid Model categories from the document."""
    categories = []
    
    for cat in doc.Settings.Categories:
        try:
            # Only Model categories
            if cat.CategoryType != CategoryType.Model:
                continue
            
            # Exclude categories in the exclusion list
            try:
                bic = System.Enum.ToObject(BuiltInCategory, cat.Id.IntegerValue)
                if bic in excluded_categories:
                    continue
            except:
                pass
            
            # Exclude categories containing link/import keywords in name
            cat_name_lower = cat.Name.lower()
            if any(x in cat_name_lower for x in [".dwg", ".dxf", ".dgn", ".sat", ".skp",
                                                  "import", "link", "cad", ".rvt"]):
                continue
            
            # Verify there are elements in this category
            collector = FilteredElementCollector(doc).OfCategoryId(cat.Id).WhereElementIsNotElementType()
            if collector.GetElementCount() > 0:
                categories.append(cat)
        except:
            continue
    
    return sorted(categories, key=lambda x: x.Name)


def get_assembly_instances(doc):
    """Gets all Assembly instances from the document."""
    assemblies = []
    try:
        collector = FilteredElementCollector(doc).OfClass(AssemblyInstance)
        assemblies = list(collector)
    except:
        pass
    return assemblies


def extract_geometric_data(element, elem_type, params_map):
    """
    Extracts all geometric data from an element.
    Uses the logic: first instance, then type, otherwise empty.
    """
    data = {}
    
    for geom_name, builtin_list in params_map.items():
        value = get_geometric_value(element, elem_type, builtin_list)
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


def get_central_model_name(doc):
    """
    Gets the central model filename (without local copy suffix).
    If it's a central file or non-workshared, returns the document title.
    If it's a local copy, returns the central model name.
    """
    try:
        # Check if worksharing is enabled
        if doc.IsWorkshared:
            # Try to get the central model path
            central_path = doc.GetWorksharingCentralModelPath()
            if central_path:
                # Get the central filename from path
                central_name = ModelPathUtils.ConvertModelPathToUserVisiblePath(central_path)
                if central_name:
                    # Extract just the filename without path and extension
                    import os
                    filename = os.path.basename(central_name)
                    if filename.lower().endswith('.rvt'):
                        filename = filename[:-4]
                    return filename
        
        # Fallback: use document title (remove .rvt if present)
        title = doc.Title
        if title.lower().endswith('.rvt'):
            title = title[:-4]
        return title
        
    except:
        # Final fallback
        title = doc.Title
        if title.lower().endswith('.rvt'):
            title = title[:-4]
        return title


def main():
    """Main script function."""
    doc = revit.doc
    
    # Build quantities -> BuiltInParameter mapping
    GEOMETRIC_PARAMS_MAP = build_geometric_params_map()
    EXCLUDED_CATEGORIES = get_excluded_categories()
    
    # Step 1: Ask for extra parameters to extract
    extra_params_input = get_extra_params_from_user()
    
    if extra_params_input is False or extra_params_input is None:
        script.exit()
    
    extra_params = []
    if extra_params_input:
        extra_params = [p.strip() for p in extra_params_input.split(";") if p.strip()]
    
    # Step 2: Ask where to save the file
    output_folder = forms.pick_folder(title="Select destination folder")
    
    if not output_folder:
        forms.alert("No folder selected. Operation cancelled.", exitscript=True)
    
    # Generate filename with new format: YYMMDD_HHMMSS_QTO_NomeFile.csv
    timestamp = datetime.now().strftime("%y%m%d_%H%M%S")
    project_name = get_central_model_name(doc).replace(" ", "_")
    filename = "{}_QTO_{}.csv".format(timestamp, project_name)
    filepath = os.path.join(output_folder, filename)
    
    # Step 3: Collect data
    output = script.get_output()
    output.print_md("# Quantity Takeoff in progress...")
    
    # Prepare column headers
    headers = list(STANDARD_COLUMNS)
    for param in extra_params:
        headers.append(param)
    
    rows = []
    processed_count = 0
    skipped_count = 0
    categories = get_model_categories(doc, EXCLUDED_CATEGORIES)
    
    output.print_md("Found **{}** categories with elements".format(len(categories)))
    
    for cat in categories:
        try:
            # Collect elements from category
            collector = FilteredElementCollector(doc).OfCategoryId(cat.Id).WhereElementIsNotElementType()
            elements = list(collector)
            
            if not elements:
                continue
            
            category_processed = 0
            category_skipped = 0
            
            for elem in elements:
                try:
                    # Skip elements from links or imports
                    if is_element_from_link_or_import(elem):
                        category_skipped += 1
                        skipped_count += 1
                        continue
                    
                    # Get element type
                    elem_type = None
                    try:
                        type_id = elem.GetTypeId()
                        if type_id and type_id != ElementId.InvalidElementId:
                            elem_type = doc.GetElement(type_id)
                    except:
                        pass
                    
                    # Basic data
                    elem_id = elem.Id.IntegerValue
                    category_name = cat.Name
                    elem_name = get_element_name(elem)
                    
                    # Geometric data
                    geo_data = extract_geometric_data(elem, elem_type, GEOMETRIC_PARAMS_MAP)
                    
                    # Prepare row
                    row = {
                        "ID": elem_id,
                        "Category": category_name,
                        "Name": elem_name,
                        "Length_m": convert_to_meters(geo_data.get("Length"), "length") if geo_data.get("Length") else "",
                        "Width_m": convert_to_meters(geo_data.get("Width"), "width") if geo_data.get("Width") else "",
                        "Height_m": convert_to_meters(geo_data.get("Height"), "height") if geo_data.get("Height") else "",
                        "Depth_m": convert_to_meters(geo_data.get("Depth"), "depth") if geo_data.get("Depth") else "",
                        "Thickness_m": convert_to_meters(geo_data.get("Thickness"), "thickness") if geo_data.get("Thickness") else "",
                        "Diameter_m": convert_to_meters(geo_data.get("Diameter"), "diameter") if geo_data.get("Diameter") else "",
                        "Perimeter_m": convert_to_meters(geo_data.get("Perimeter"), "perimeter") if geo_data.get("Perimeter") else "",
                        "Area_m2": convert_to_meters(geo_data.get("Area"), "area") if geo_data.get("Area") else "",
                        "Volume_m3": convert_to_meters(geo_data.get("Volume"), "volume") if geo_data.get("Volume") else "",
                    }
                    
                    # Extra parameters
                    for param_name in extra_params:
                        value = get_param_value_by_name(elem, param_name)
                        if value is not None:
                            try:
                                value = round(float(value), 3)
                            except:
                                pass
                        row[param_name] = value if value is not None else ""
                    
                    rows.append(row)
                    processed_count += 1
                    category_processed += 1
                    
                except Exception as e:
                    continue
            
            if category_processed > 0 or category_skipped > 0:
                msg = "- **{}**: {} elements".format(cat.Name, category_processed)
                if category_skipped > 0:
                    msg += " ({} skipped - links/imports)".format(category_skipped)
                output.print_md(msg)
                    
        except Exception as e:
            output.print_md("  - Error in category {}: {}".format(cat.Name, str(e)))
            continue
    
    # Process Assemblies separately (they need a dedicated collector)
    try:
        assemblies = get_assembly_instances(doc)
        if assemblies:
            assembly_processed = 0
            assembly_skipped = 0
            
            for elem in assemblies:
                try:
                    # Skip elements from links or imports
                    if is_element_from_link_or_import(elem):
                        assembly_skipped += 1
                        skipped_count += 1
                        continue
                    
                    # Get element type
                    elem_type = None
                    try:
                        type_id = elem.GetTypeId()
                        if type_id and type_id != ElementId.InvalidElementId:
                            elem_type = doc.GetElement(type_id)
                    except:
                        pass
                    
                    # Basic data
                    elem_id = elem.Id.IntegerValue
                    category_name = "Assemblies"
                    elem_name = get_element_name(elem)
                    
                    # Geometric data
                    geo_data = extract_geometric_data(elem, elem_type, GEOMETRIC_PARAMS_MAP)
                    
                    # Prepare row
                    row = {
                        "ID": elem_id,
                        "Category": category_name,
                        "Name": elem_name,
                        "Length_m": convert_to_meters(geo_data.get("Length"), "length") if geo_data.get("Length") else "",
                        "Width_m": convert_to_meters(geo_data.get("Width"), "width") if geo_data.get("Width") else "",
                        "Height_m": convert_to_meters(geo_data.get("Height"), "height") if geo_data.get("Height") else "",
                        "Depth_m": convert_to_meters(geo_data.get("Depth"), "depth") if geo_data.get("Depth") else "",
                        "Thickness_m": convert_to_meters(geo_data.get("Thickness"), "thickness") if geo_data.get("Thickness") else "",
                        "Diameter_m": convert_to_meters(geo_data.get("Diameter"), "diameter") if geo_data.get("Diameter") else "",
                        "Perimeter_m": convert_to_meters(geo_data.get("Perimeter"), "perimeter") if geo_data.get("Perimeter") else "",
                        "Area_m2": convert_to_meters(geo_data.get("Area"), "area") if geo_data.get("Area") else "",
                        "Volume_m3": convert_to_meters(geo_data.get("Volume"), "volume") if geo_data.get("Volume") else "",
                    }
                    
                    # Extra parameters
                    for param_name in extra_params:
                        value = get_param_value_by_name(elem, param_name)
                        if value is not None:
                            try:
                                value = round(float(value), 3)
                            except:
                                pass
                        row[param_name] = value if value is not None else ""
                    
                    rows.append(row)
                    processed_count += 1
                    assembly_processed += 1
                    
                except Exception as e:
                    continue
            
            if assembly_processed > 0 or assembly_skipped > 0:
                msg = "- **Assemblies**: {} elements".format(assembly_processed)
                if assembly_skipped > 0:
                    msg += " ({} skipped - links/imports)".format(assembly_skipped)
                output.print_md(msg)
    except Exception as e:
        output.print_md("  - Error processing Assemblies: {}".format(str(e)))
    
    # Step 4: Write CSV file
    output.print_md("\n## Writing CSV file...")
    
    try:
        # Use newline='' to prevent extra blank rows on Windows
        with open(filepath, 'wb') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=headers, delimiter=';', 
                                   lineterminator='\n')
            writer.writeheader()
            for row in rows:
                row_str = {}
                for k, v in row.items():
                    if v is None or v == "":
                        row_str[k] = ""
                    elif isinstance(v, float):
                        row_str[k] = str(v).replace('.', ',')
                    else:
                        try:
                            row_str[k] = str(v).encode('utf-8')
                        except:
                            row_str[k] = str(v)
                writer.writerow(row_str)
        
        output.print_md("\n---")
        output.print_md("## Export completed!")
        output.print_md("- **Elements processed:** {}".format(processed_count))
        if skipped_count > 0:
            output.print_md("- **Elements skipped (links/imports):** {}".format(skipped_count))
        output.print_md("- **File saved:** {}".format(filepath))
            
    except Exception as e:
        forms.alert("Error writing file:\n{}".format(str(e)), exitscript=True)


if __name__ == "__main__":
    main()
