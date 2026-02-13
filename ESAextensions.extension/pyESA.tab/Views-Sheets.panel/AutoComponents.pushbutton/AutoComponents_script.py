# -*- coding: utf-8 -*-
"""
AutoLegend - Creazione automatica di legende in Revit
Refactored: Aggiunta UI XAML, griglia personalizzabile, supporto categorie caricabili

Questo script crea automaticamente una legenda duplicando un componente
di legenda di riferimento per tutti i tipi selezionati, posizionandoli
in una griglia definita dall'utente.
"""

import os
import sys

sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from pyrevit import forms, revit, DB

import esa_legend
from legend_ui import show_config_form, LegendConfig

doc = revit.doc


def select_legend_component(category):
    with forms.WarningBar(title='SELECT A LEGEND COMPONENT (category: {})'.format(
            DB.Category.GetCategory(doc, category).Name)):
        
        legend = revit.pick_element()
        
        if not legend:
            return None
            
        legend_param = legend.get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT)
        if not legend_param:
            forms.alert('You must select a Legend Component.', exitscript=True)
            return None
            
        type_id = legend_param.AsElementId()
        if type_id == DB.ElementId.InvalidElementId:
            forms.alert('The Legend Component does not have a valid type.', exitscript=True)
            return None
            
        element_type = doc.GetElement(type_id)
        if not element_type or not element_type.Category:
            forms.alert('Unable to determine the Legend Component category.', exitscript=True)
            return None
            
        expected_cat_id = DB.Category.GetCategory(doc, category).Id
        if element_type.Category.Id != expected_cat_id:
            forms.alert(
                'The selected Legend Component is category "{}", '
                'but must be category "{}".'.format(
                    element_type.Category.Name,
                    DB.Category.GetCategory(doc, category).Name
                ),
                exitscript=True
            )
            return None
            
        return legend


def select_text_notes():
    with forms.WarningBar(title='SELECT TEXT NOTES (optional) - Press "Finish" or "Esc" to skip'):
        try:
            selected = revit.pick_elements_by_category(
                DB.BuiltInCategory.OST_TextNotes,
                'Select TextNotes to copy'
            )
            return selected if selected else []
        except:
            return []


def select_detail_items():
    with forms.WarningBar(title='SELECT DETAIL ITEMS (optional) - Press "Finish" or "Esc" to skip'):
        try:
            selected = revit.pick_elements_by_category(
                DB.BuiltInCategory.OST_DetailComponents,
                'Select Detail Items to copy'
            )
            return selected if selected else []
        except:
            return []


def get_parameter_value_for_sorting(type_id, param_name):
    element_type = doc.GetElement(type_id)
    if not element_type:
        return ""
        
    param = element_type.LookupParameter(param_name)
    if not param:
        return ""
        
    if param.StorageType == DB.StorageType.String:
        return param.AsString() or ""
    elif param.StorageType == DB.StorageType.Integer:
        return str(param.AsInteger())
    elif param.StorageType == DB.StorageType.Double:
        return param.AsValueString() or str(param.AsDouble())
    else:
        return param.AsValueString() or ""


def sort_types_by_parameter(type_ids, param_name):
    if not param_name:
        return type_ids
        
    type_values = []
    for type_id in type_ids:
        value = get_parameter_value_for_sorting(type_id, param_name)
        type_values.append((type_id, value))
    
    type_values.sort(key=lambda x: x[1].lower() if x[1] else "")
    
    return [tv[0] for tv in type_values]


def create_legend_entry(legend_source, type_id, position_vector, config, view, 
                        text_notes=None, detail_items=None, pt_base=None):
    new_legend = esa_legend.copy_element_to_position(doc, legend_source.Id, position_vector)
    if not new_legend:
        return None
        
    new_legend.get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT).Set(type_id)
    
    new_bb = new_legend.get_BoundingBox(view)
    if pt_base and new_bb:
        accessory_vec = new_bb.Min.Subtract(pt_base)
    else:
        accessory_vec = position_vector
        
    if text_notes:
        for txt in text_notes:
            new_txt = esa_legend.copy_element_to_position(doc, txt.Id, accessory_vec)
            if new_txt:
                esa_legend.set_TextNote(new_txt, type_id, doc)
                
    if detail_items:
        for detail in detail_items:
            esa_legend.copy_element_to_position(doc, detail.Id, accessory_vec)
            
    return new_legend


def cleanup_source_elements(legend, text_notes, detail_items):
    with revit.Transaction('Delete source elements'):
        doc.Delete(legend.Id)
        
        if text_notes:
            for txt in text_notes:
                try:
                    doc.Delete(txt.Id)
                except:
                    pass
                    
        if detail_items:
            for detail in detail_items:
                try:
                    doc.Delete(detail.Id)
                except:
                    pass


def main():
    view = doc.ActiveView
    
    config = show_config_form(doc)
    if not config:
        return
    
    legend = select_legend_component(config.category)
    if not legend:
        return
    
    bb = legend.get_BoundingBox(view)
    pt_base = bb.Min if bb else None
    
    text_notes = select_text_notes()
    detail_items = select_detail_items()
    
    types_to_process = config.selected_types[:]
    
    if config.sort_parameter:
        types_to_process = sort_types_by_parameter(types_to_process, config.sort_parameter)
    
    created_components = []
    
    with revit.TransactionGroup('AutoLegend - Create Legend'):
        
        with forms.ProgressBar(title='Creating legend ({value} of {max_value})',
                               cancellable=True) as pb:
            
            for index, type_id in enumerate(types_to_process):
                
                if pb.cancelled:
                    break
                    
                pb.update_progress(index + 1, len(types_to_process))
                
                position = esa_legend.calculate_grid_position(
                    index,
                    config.columns,
                    config.offset_x_internal,
                    config.offset_y_internal
                )
                
                with revit.Transaction('Create legend entry'):
                    new_legend = create_legend_entry(
                        legend, type_id, position, config, view,
                        text_notes, detail_items, pt_base
                    )
                    
                    if new_legend:
                        created_components.append(new_legend)
                        
                # Add dimensions and tags (skip index=0, will be added after source deletion)
                if new_legend and index > 0 and config.is_compound_category:
                    is_compound, _ = esa_legend.get_info(new_legend, doc)
                    
                    if is_compound and config.insert_dimensions:
                        with revit.Transaction('Create Dimensions'):
                            esa_legend.create_dimensions(
                                new_legend, view, doc,
                                config.is_horizontal,
                                config.dimension_type_id,
                                config.dim_offset_internal,
                                config.dim_position_above_right
                            )
                            
                    if is_compound and config.insert_tags and config.tag_type_id:
                        with revit.Transaction('Create Material Tags'):
                            esa_legend.create_material_tags(
                                new_legend, config.tag_type_id, view, doc,
                                config.is_horizontal,
                                config.tag_offset_internal,
                                config.tag_spacing_internal
                            )
    
    first_element = created_components[0] if created_components else None
    
    cleanup_source_elements(legend, text_notes, detail_items)
    
    # Add dimensions and tags to first element (after source deletion)
    if first_element and config.is_compound_category:
        is_compound, _ = esa_legend.get_info(first_element, doc)
        
        if is_compound and config.insert_dimensions:
            with revit.Transaction('Create Dimensions - First Element'):
                esa_legend.create_dimensions(
                    first_element, view, doc,
                    config.is_horizontal,
                    config.dimension_type_id,
                    config.dim_offset_internal,
                    config.dim_position_above_right
                )
                
        if is_compound and config.insert_tags and config.tag_type_id:
            with revit.Transaction('Create Material Tags - First Element'):
                esa_legend.create_material_tags(
                    first_element, config.tag_type_id, view, doc,
                    config.is_horizontal,
                    config.tag_offset_internal,
                    config.tag_spacing_internal
                )


if __name__ == '__main__':
    try:
        main()
    except Exception as ex:
        forms.alert('Error during execution:\n{}'.format(str(ex)))
        import traceback
        print(traceback.format_exc())
