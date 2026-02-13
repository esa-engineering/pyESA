# -*- coding: utf-8 -*-
"""
esa_legend.py - Utilità per la gestione automatica di legende in Revit
Author: Claude + Antonio Miano
Refactored: Supporto categorie caricabili, rimosso variabile globale doc
"""

from pyrevit import DB


# =============================================================================
# CATEGORY DEFINITIONS
# =============================================================================

# Compound categories (with layers structure)
COMPOUND_CATEGORIES = [
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_Roofs,
    DB.BuiltInCategory.OST_Ceilings
]

# Loadable family categories
LOADABLE_CATEGORIES = [
    DB.BuiltInCategory.OST_Furniture,
    DB.BuiltInCategory.OST_Casework,
    DB.BuiltInCategory.OST_GenericModel,
    DB.BuiltInCategory.OST_Doors,
    DB.BuiltInCategory.OST_Windows,
    DB.BuiltInCategory.OST_PlumbingFixtures,
    DB.BuiltInCategory.OST_LightingFixtures,
    DB.BuiltInCategory.OST_ElectricalFixtures,
    DB.BuiltInCategory.OST_ElectricalEquipment,
    DB.BuiltInCategory.OST_MechanicalEquipment,
    DB.BuiltInCategory.OST_SpecialityEquipment,
    DB.BuiltInCategory.OST_Entourage,
    DB.BuiltInCategory.OST_Planting,
    DB.BuiltInCategory.OST_Site
]


def get_compound_category_ids(doc):
    """
    Restituisce gli ID delle categorie che supportano strutture compound.
    
    Args:
        doc (DB.Document): Il documento Revit
        
    Returns:
        list: Lista di ElementId delle categorie compound
    """
    ids = []
    for cat in COMPOUND_CATEGORIES:
        try:
            cat_obj = DB.Category.GetCategory(doc, cat)
            if cat_obj:
                ids.append(cat_obj.Id)
        except:
            pass
    return ids


def is_compound_category(category_enum):
    """
    Verifica se una categoria è di tipo compound.
    
    Args:
        category_enum: DB.BuiltInCategory
        
    Returns:
        bool: True se la categoria supporta strutture compound
    """
    return category_enum in COMPOUND_CATEGORIES


# =============================================================================
# LEGEND COMPONENT INFO
# =============================================================================

def get_info(legend, doc):
    """
    Determina se il componente legenda rappresenta un oggetto compound
    e il suo orientamento (orizzontale o verticale).
    
    Args:
        legend (DB.Element): Il componente legenda da ispezionare
        doc (DB.Document): Il documento Revit
        
    Returns:
        tuple: (is_compound: bool, is_horizontal: bool)
    """
    compound_categories = get_compound_category_ids(doc)
    
    # Ottieni l'ID dell'elemento tipo dal parametro LEGEND_COMPONENT
    legend_param = legend.get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT)
    if not legend_param:
        return (False, False)
        
    type_id = legend_param.AsElementId()
    if type_id == DB.ElementId.InvalidElementId:
        return (False, False)
        
    element_type = doc.GetElement(type_id)
    if not element_type:
        return (False, False)
        
    # Verifica se ha una categoria valida
    if not element_type.Category:
        return (False, False)
        
    cat_id = element_type.Category.Id
    
    # Verifica se è una categoria compound
    if cat_id in compound_categories:
        # Walls ID
        walls_id = compound_categories[0]
        if cat_id == walls_id:
            # Per i muri, controlla il parametro vista
            view_param = legend.get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT_VIEW)
            if view_param:
                # -8 indica sezione orizzontale (vista dall'alto)
                is_horizontal = (view_param.AsInteger() == -8)
                return (True, is_horizontal)
            return (True, True)
        else:
            # Floors, Roofs, Ceilings sono sempre orizzontali
            return (True, True)
    else:
        return (False, False)


# =============================================================================
# LAYERS EXTRACTION
# =============================================================================

def _tag_distribution(layer_widths):
    """
    Calcola le posizioni centrali per ogni strato.
    
    Args:
        layer_widths (list): Lista degli spessori degli strati
        
    Returns:
        list: Lista delle posizioni centrali
    """
    positions = []
    for n in range(len(layer_widths)):
        pos = sum(layer_widths[:n]) + layer_widths[n] / 2.0
        positions.append(pos)
    return positions


def get_layers(legend_component, doc):
    """
    Estrae le informazioni sugli strati da un componente legenda compound.
    
    Args:
        legend_component (DB.Element): Il componente legenda
        doc (DB.Document): Il documento Revit
        
    Returns:
        tuple: (all_widths: list, material_ids: list)
    """
    type_id = legend_component.get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT).AsElementId()
    element_type = doc.GetElement(type_id)
    
    all_width = []
    materials_id = []
    
    if element_type and element_type.GetCompoundStructure():
        compound = element_type.GetCompoundStructure()
        for layer in reversed(compound.GetLayers()):
            if layer.Width:
                all_width.append(layer.Width)
                materials_id.append(layer.MaterialId)
                
    return all_width, materials_id


def get_layer_count(legend_component, doc):
    """
    Restituisce il numero di strati di un componente legenda.
    
    Args:
        legend_component (DB.Element): Il componente legenda
        doc (DB.Document): Il documento Revit
        
    Returns:
        int: Numero di strati
    """
    widths, _ = get_layers(legend_component, doc)
    return len(widths)


# =============================================================================
# MATERIAL TAGS
# =============================================================================

def create_material_tags(legend_component, tag_type_id, view, doc, is_horizontal=True,
                         tag_offset=0.1, tag_spacing=0.065):
    """
    Place one material tag per layer on the compound legend component.
    Tags are created from TOP to BOTTOM.
    """
    if not tag_type_id or tag_type_id == DB.ElementId.InvalidElementId:
        return []
        
    ref = DB.Reference(legend_component)
    bb = legend_component.get_BoundingBox(view)
    
    if not bb:
        return []
        
    layer_widths, _ = get_layers(legend_component, doc)
    if not layer_widths:
        return []
    
    pts = []
    pt = bb.Max
    
    for n, y in enumerate(_tag_distribution(layer_widths)):
        if is_horizontal:
            new_pt = pt.Add(DB.XYZ(-(bb.Max.X - bb.Min.X) + ((n + 1) * tag_spacing), -(bb.Max.Y - bb.Min.Y) + y, bb.Min.Z))
            pts.append(new_pt)
        else:
            new_pt = pt.Add(DB.XYZ(-(bb.Max.X - bb.Min.X) + y, -(n + 1) * tag_spacing, bb.Min.Z))
            pts.append(new_pt)
    
    if is_horizontal:
        Xs = [p.X for p in pts]
        for i, x in enumerate(Xs[::-1]):
            pts[i] = DB.XYZ(x, pts[i].Y, pts[i].Z)
    
    created_tags = []
    for n, pt in enumerate(pts):
        try:
            new_tag = DB.IndependentTag.Create(
                doc, tag_type_id, view.Id, ref,
                True, DB.TagOrientation.Horizontal, pt
            )
            
            new_tag.LeaderEndCondition = DB.LeaderEndCondition.Free
            new_tag.SetLeaderEnd(ref, pt)
            
            if is_horizontal:
                new_tag.SetLeaderElbow(ref, DB.XYZ(pt.X, bb.Max.Y + (n + 1) * tag_spacing, pt.Z))
                new_tag.TagHeadPosition = DB.XYZ(bb.Max.X + tag_offset, bb.Max.Y + (n + 1) * tag_spacing, pt.Z)
            else:
                new_tag.SetLeaderElbow(ref, DB.XYZ(bb.Max.X, pt.Y, pt.Z))
                new_tag.TagHeadPosition = DB.XYZ(bb.Max.X + tag_offset, pt.Y, pt.Z)
                
            created_tags.append(new_tag)
        except:
            pass
            
    return created_tags


# =============================================================================
# DIMENSIONS
# =============================================================================

def create_dimensions(legend_component, view, doc, is_horizontal=True, dimension_type_id=None,
                      dim_offset=0.164, position_above_right=True):
    """
    Create dimensions for a compound legend component.
    Creates only individual layer dimensions (no total dimension).
    
    Args:
        legend_component (DB.Element): The legend component
        view (DB.View): The legend view
        doc (DB.Document): The Revit document
        is_horizontal (bool): Component orientation
        dimension_type_id (DB.ElementId): Dimension style ID (optional)
        dim_offset (float): Offset from detail reference lines (internal units)
        position_above_right (bool): True = above/right, False = below/left
        
    Returns:
        list: List of created dimensions
    """
    layer_widths, _ = get_layers(legend_component, doc)
    
    if not layer_widths:
        return []
        
    bb = legend_component.get_BoundingBox(view)
    if not bb:
        return []
        
    created_dims = []
    
    # Apply negative offset for Below/Left position
    effective_offset = dim_offset if position_above_right else -dim_offset
    
    # Configure based on orientation
    if is_horizontal:
        base = bb.Min
        vec = DB.XYZ(0, 1, 0)  # Direction along layers
        ln_base = DB.Line.CreateBound(base, base.Add(DB.XYZ(0.1, 0, 0)))
        cross_vec = DB.XYZ(-1, 0, 0)
    else:
        base = bb.Min
        vec = DB.XYZ(1, 0, 0)  # Direction along layers
        ln_base = DB.Line.CreateBound(base, base.Add(DB.XYZ(0, 0.1, 0)))
        cross_vec = DB.XYZ(0, 1, 0)
        
    # Create base reference curve
    dt_crv = doc.Create.NewDetailCurve(view, ln_base)
    
    # Array for dimensions (includes all layer boundaries)
    ref_array = DB.ReferenceArray()
    ref_array.Append(DB.Reference(dt_crv))
    
    # Create reference curves for each layer boundary
    progression = 0
    for w in layer_widths:
        progression += w
        new_crv = DB.ElementTransformUtils.CopyElement(doc, dt_crv.Id, vec.Multiply(progression))
        ref_array.Append(DB.Reference(doc.GetElement(new_crv[0])))
    
    # Get dimension style
    dim_type = None
    if dimension_type_id and dimension_type_id != DB.ElementId.InvalidElementId:
        dim_type = doc.GetElement(dimension_type_id)
        
    try:
        # Create dimension line for individual layers only
        dim_ln = DB.Line.CreateUnbound(
            base.Add(cross_vec.Multiply(effective_offset)),
            vec
        )
        dim = doc.Create.NewDimension(view, dim_ln, ref_array)
        if dim and dim_type:
            try:
                dim.DimensionType = dim_type
            except:
                pass
        if dim:
            created_dims.append(dim)
            
    except:
        pass
        
    return created_dims


# Keep legacy function for compatibility
def create_dimension_horizontal(legend_component, view, doc, horizontal=True, dimension_type_id=None,
                                dim_offset=0.164, position_above_right=True):
    """
    Alias for create_dimensions for backward compatibility.
    """
    return create_dimensions(legend_component, view, doc, horizontal, dimension_type_id,
                             dim_offset, position_above_right)


# =============================================================================
# TEXT NOTES
# =============================================================================

def _get_value(param):
    """
    Estrae il valore da un parametro Revit.
    
    Args:
        param: Il parametro Revit
        
    Returns:
        str: Il valore come stringa
    """
    if param and param.StorageType == DB.StorageType.String:
        return param.AsString() if param.AsString() else ''
    elif param:
        return param.AsValueString() if param.AsValueString() else ''
    else:
        return '---'


def set_TextNote(text_note, type_id, doc):
    """
    Update a TextNote replacing {parameter} placeholders with actual values.
    
    The TextNote should contain placeholders in the format {ParameterName}.
    For example: "{Type Mark}" will be replaced with the actual Type Mark value.
    If a parameter is empty or not found, the placeholder is kept unchanged.
    
    Args:
        text_note (DB.TextNote or DB.Element): The TextNote to modify
        type_id (DB.ElementId): The ID of the type from which to extract parameters
        doc (DB.Document): The Revit document
    """
    if not text_note or not type_id:
        return
    
    # Ensure type_id is an ElementId
    if not isinstance(type_id, DB.ElementId):
        return
        
    element_type = doc.GetElement(type_id)
    if not element_type:
        return
    
    # Get the current text
    try:
        current_text = text_note.Text
    except:
        return
    
    if not current_text:
        return
        
    # Extract parameter names from placeholders {name}
    param_names = []
    for part in current_text.split('{'):
        if '}' in part:
            param_name = part[:part.index('}')]
            if param_name:
                param_names.append(param_name)
    
    if not param_names:
        return
            
    # Replace each placeholder with the parameter value
    # Only replace if value is not empty, otherwise keep the placeholder
    new_text = current_text
    for p_name in param_names:
        param = element_type.LookupParameter(p_name)
        value = _get_value(param)
        
        # Only replace if value is not empty and not the default '---'
        if value and value != '---' and value.strip():
            new_text = new_text.replace('{' + p_name + '}', value)
    
    # Set the new text
    try:
        text_note.Text = new_text
    except:
        pass


# =============================================================================
# GRID POSITIONING
# =============================================================================

def calculate_grid_position(index, columns, offset_x, offset_y):
    """
    Calculate the grid position for a given index.
    
    Args:
        index (int): Element index (0-based)
        columns (int): Number of columns
        offset_x (float): Horizontal offset between columns (internal units)
        offset_y (float): Vertical offset between rows (internal units)
        
    Returns:
        DB.XYZ: Translation vector
    """
    col = index % columns
    row = index // columns
    
    # Y is negative because rows go downward in Revit
    return DB.XYZ(col * offset_x, -row * offset_y, 0)


def copy_element_to_position(doc, source_id, position_vector):
    """
    Copia un elemento in una nuova posizione.
    
    Args:
        doc (DB.Document): Il documento Revit
        source_id (DB.ElementId): ID dell'elemento sorgente
        position_vector (DB.XYZ): Vettore di traslazione
        
    Returns:
        DB.Element: L'elemento copiato
    """
    new_ids = DB.ElementTransformUtils.CopyElement(doc, source_id, position_vector)
    if new_ids and len(new_ids) > 0:
        return doc.GetElement(new_ids[0])
    return None
