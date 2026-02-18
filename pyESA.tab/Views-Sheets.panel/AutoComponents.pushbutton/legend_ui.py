# -*- coding: utf-8 -*-
"""
legend_ui.py - Gestione Form XAML per AutoLegend
Author: Giuseppe Dotto - ESA Engineering
Refactored with XAML UI, support for loadable categories
"""

import os
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window, MessageBox, Visibility
from System.Windows.Controls import CheckBox, StackPanel as WpfStackPanel
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode

from pyrevit import DB, revit
import math


# =============================================================================
# UTILITY FUNCTIONS - UNIT CONVERSION
# =============================================================================

def mm_to_internal(mm_value, doc):
    """
    Converte millimetri in unità interne Revit (feet).
    Gestisce la differenza API tra Revit < 2022 e >= 2022.
    """
    if int(doc.Application.VersionNumber) < 2022:
        return DB.UnitUtils.ConvertToInternalUnits(mm_value, DB.DisplayUnitType.DUT_MILLIMETERS)
    else:
        return DB.UnitUtils.ConvertToInternalUnits(mm_value, DB.UnitTypeId.Millimeters)


def internal_to_mm(internal_value, doc):
    """
    Converte unità interne Revit (feet) in millimetri.
    """
    if int(doc.Application.VersionNumber) < 2022:
        return DB.UnitUtils.ConvertFromInternalUnits(internal_value, DB.DisplayUnitType.DUT_MILLIMETERS)
    else:
        return DB.UnitUtils.ConvertFromInternalUnits(internal_value, DB.UnitTypeId.Millimeters)


def get_symbol_name(symbol):
    """
    Restituisce il nome completo di un FamilySymbol/ElementType.
    """
    try:
        if hasattr(symbol, 'FamilyName'):
            param_name = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if param_name:
                return '{}: {}'.format(symbol.FamilyName, param_name.AsString())
        return symbol.Name if hasattr(symbol, 'Name') else str(symbol.Id.IntegerValue)
    except:
        return str(symbol.Id.IntegerValue)


# =============================================================================
# DATA CLASSES
# =============================================================================

class TypeItem(object):
    """Wrapper class for type items with checkbox support."""
    
    def __init__(self, name, type_id, is_selected=False):
        self._name = name
        self._type_id = type_id
        self._is_selected = is_selected
        
    @property
    def Name(self):
        return self._name
        
    @property
    def TypeId(self):
        return self._type_id
        
    @property
    def IsSelected(self):
        return self._is_selected
        
    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value


class LegendConfig(object):
    """Class containing all legend configuration parameters."""
    
    def __init__(self):
        # Category and types
        self.category = None              # DB.BuiltInCategory
        self.is_compound_category = True  # True for Walls/Floors/etc, False for Furniture/etc
        self.selected_types = []          # List of ElementId
        self.sort_parameter = None        # Parameter name for sorting
        
        # Grid layout
        self.columns = 3
        self.rows = 1
        self.offset_x_mm = 300.0
        self.offset_y_mm = 150.0
        self.offset_x_internal = 0.0
        self.offset_y_internal = 0.0
        self.is_horizontal = True
        
        # Dimensions (only for compound categories)
        self.insert_dimensions = True
        self.dimension_type_id = None
        self.dim_offset_mm = 50.0
        self.dim_offset_internal = 0.0
        self.dim_position_above_right = True
        
        # Material Tags (only for compound categories)
        self.insert_tags = True
        self.tag_type_id = None
        self.tag_offset_mm = 30.0
        self.tag_offset_internal = 0.0
        self.tag_spacing_mm = 20.0
        self.tag_spacing_internal = 0.0
        
    def calculate_internal_units(self, doc):
        """Convert all mm values to internal units."""
        self.offset_x_internal = mm_to_internal(self.offset_x_mm, doc)
        self.offset_y_internal = mm_to_internal(self.offset_y_mm, doc)
        self.dim_offset_internal = mm_to_internal(self.dim_offset_mm, doc)
        self.tag_offset_internal = mm_to_internal(self.tag_offset_mm, doc)
        self.tag_spacing_internal = mm_to_internal(self.tag_spacing_mm, doc)
        
    def calculate_rows(self, num_types):
        """Calculate the number of rows needed."""
        if self.columns > 0:
            self.rows = int(math.ceil(float(num_types) / float(self.columns)))
        else:
            self.rows = num_types
        return self.rows


# =============================================================================
# CATEGORY DEFINITIONS
# =============================================================================

# Compound categories (with layers structure) - support dimensions and material tags
COMPOUND_CATEGORIES = [
    (DB.BuiltInCategory.OST_Walls, "Walls"),
    (DB.BuiltInCategory.OST_Floors, "Floors"),
    (DB.BuiltInCategory.OST_Roofs, "Roofs"),
    (DB.BuiltInCategory.OST_Ceilings, "Ceilings")
]

# Loadable family categories - NO dimensions or material tags
LOADABLE_CATEGORIES = [
    (DB.BuiltInCategory.OST_Furniture, "Furniture"),
    (DB.BuiltInCategory.OST_Casework, "Casework"),
    (DB.BuiltInCategory.OST_GenericModel, "Generic Models"),
    (DB.BuiltInCategory.OST_Doors, "Doors"),
    (DB.BuiltInCategory.OST_Windows, "Windows"),
    (DB.BuiltInCategory.OST_PlumbingFixtures, "Plumbing Fixtures"),
    (DB.BuiltInCategory.OST_LightingFixtures, "Lighting Fixtures"),
    (DB.BuiltInCategory.OST_ElectricalFixtures, "Electrical Fixtures"),
    (DB.BuiltInCategory.OST_ElectricalEquipment, "Electrical Equipment"),
    (DB.BuiltInCategory.OST_MechanicalEquipment, "Mechanical Equipment"),
    (DB.BuiltInCategory.OST_SpecialityEquipment, "Specialty Equipment"),
    (DB.BuiltInCategory.OST_Entourage, "Entourage"),
    (DB.BuiltInCategory.OST_Planting, "Planting"),
    (DB.BuiltInCategory.OST_Site, "Site")
]

# Combined list for UI
ALL_CATEGORIES = COMPOUND_CATEGORIES + LOADABLE_CATEGORIES

# Set of compound category enums for quick lookup
COMPOUND_CATEGORY_ENUMS = set([cat[0] for cat in COMPOUND_CATEGORIES])


# =============================================================================
# FORM CLASS
# =============================================================================

class LegendConfigForm(Window):
    """Form XAML per la configurazione della legenda automatica."""
    
    def __init__(self, doc):
        """
        Inizializza la form.
        
        Args:
            doc: Il documento Revit corrente
        """
        self.doc = doc
        self.config = LegendConfig()
        self.result = False
        
        # Data for ComboBoxes
        self._types_data = {}
        self._current_types_map = {}
        self._sort_params = {}
        self._dim_types = {}
        self._tag_types = {}
        
        # Data for filtered list with checkboxes
        self._all_type_items = []      # All TypeItem objects for current category
        self._filtered_type_items = [] # Filtered TypeItem objects (after search)
        
        # Carica XAML
        self._load_xaml()
        
        # Inizializza dati
        self._init_categories()
        self._init_dimension_types()
        self._init_tag_types()
        
        # Aggiorna UI iniziale
        self._update_types_list()
        self._update_rows_label()
        self._update_compound_options()
        
    def _load_xaml(self):
        """Carica e parsa il file XAML."""
        script_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(script_dir, 'legend_form.xaml')
        
        Window.__init__(self)
        
        stream = FileStream(xaml_path, FileMode.Open)
        try:
            root = XamlReader.Load(stream)
        finally:
            stream.Close()
        
        self.Content = root.Content
        self.Title = root.Title
        self.Height = root.Height
        self.Width = root.Width
        self.WindowStartupLocation = root.WindowStartupLocation
        self.ResizeMode = root.ResizeMode
        
        self._find_controls(root)
        self._wire_events()
        
    def _find_controls(self, root):
        """Find all necessary controls in the XAML."""
        self.cbo_category = root.FindName('cbo_category')
        self.lst_types = root.FindName('lst_types')
        self.txt_search = root.FindName('txt_search')
        self.btn_clear_search = root.FindName('btn_clear_search')
        self.thumb_resize = root.FindName('thumb_resize')
        self.btn_select_all = root.FindName('btn_select_all')
        self.btn_deselect_all = root.FindName('btn_deselect_all')
        self.cbo_sort_param = root.FindName('cbo_sort_param')
        self.txt_columns = root.FindName('txt_columns')
        self.lbl_rows = root.FindName('lbl_rows')
        self.txt_offset_x = root.FindName('txt_offset_x')
        self.txt_offset_y = root.FindName('txt_offset_y')
        self.cbo_orientation = root.FindName('cbo_orientation')
        # Dimensions
        self.border_dimensions = root.FindName('border_dimensions')
        self.chk_dimensions = root.FindName('chk_dimensions')
        self.cbo_dim_style = root.FindName('cbo_dim_style')
        self.txt_dim_offset = root.FindName('txt_dim_offset')
        self.cbo_dim_position = root.FindName('cbo_dim_position')
        # Tags
        self.border_tags = root.FindName('border_tags')
        self.chk_tags = root.FindName('chk_tags')
        self.cbo_tag_type = root.FindName('cbo_tag_type')
        self.txt_tag_offset = root.FindName('txt_tag_offset')
        self.txt_tag_spacing = root.FindName('txt_tag_spacing')
        # Buttons
        self.btn_ok = root.FindName('btn_ok')
        self.btn_cancel = root.FindName('btn_cancel')
        
    def _wire_events(self):
        """Collega gli eventi ai metodi handler."""
        self.cbo_category.SelectionChanged += self.OnCategoryChanged
        self.txt_search.TextChanged += self.OnSearchTextChanged
        self.btn_clear_search.Click += self.OnClearSearch
        self.thumb_resize.DragDelta += self.OnResizeList
        self.btn_select_all.Click += self.OnSelectAll
        self.btn_deselect_all.Click += self.OnDeselectAll
        self.txt_columns.TextChanged += self.OnGridParameterChanged
        self.chk_dimensions.Checked += self.OnDimensionsChecked
        self.chk_dimensions.Unchecked += self.OnDimensionsUnchecked
        self.chk_tags.Checked += self.OnTagsChecked
        self.chk_tags.Unchecked += self.OnTagsUnchecked
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel
        self.lst_types.SelectionChanged += self.OnTypesSelectionChanged
        
    def _init_categories(self):
        """Inizializza la ComboBox delle categorie."""
        self.cbo_category.Items.Clear()
        
        # Add separator between compound and loadable
        for cat_enum, cat_name in COMPOUND_CATEGORIES:
            self.cbo_category.Items.Add(cat_name)
            self._load_types_for_category(cat_enum)
        
        # Add visual separator
        self.cbo_category.Items.Add("─── Loadable Families ───")
        
        for cat_enum, cat_name in LOADABLE_CATEGORIES:
            self.cbo_category.Items.Add(cat_name)
            self._load_types_for_category(cat_enum)
            
        if self.cbo_category.Items.Count > 0:
            self.cbo_category.SelectedIndex = 0
            
    def _load_types_for_category(self, category_enum):
        """Carica tutti i tipi disponibili per una categoria."""
        types_dict = {}
        
        try:
            collector = DB.FilteredElementCollector(self.doc)\
                         .OfCategory(category_enum)\
                         .WhereElementIsElementType()
            
            for element_type in collector:
                if element_type and element_type.Id != DB.ElementId.InvalidElementId:
                    # Per i muri, includi solo Basic (no Curtain)
                    if category_enum == DB.BuiltInCategory.OST_Walls:
                        if hasattr(element_type, 'Kind'):
                            if element_type.Kind != DB.WallKind.Basic:
                                continue
                    
                    # Get name - try FamilyName: TypeName format for loadable families
                    if hasattr(element_type, 'FamilyName') and element_type.FamilyName:
                        param_name = element_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                        if param_name and param_name.AsString():
                            type_name = '{}: {}'.format(element_type.FamilyName, param_name.AsString())
                        else:
                            type_name = element_type.FamilyName
                    else:
                        name_param = element_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                        if name_param:
                            type_name = name_param.AsString()
                        else:
                            type_name = None
                    
                    if type_name:
                        types_dict[type_name] = element_type.Id
                            
        except Exception as ex:
            pass
            
        self._types_data[category_enum] = types_dict
        
    def _init_dimension_types(self):
        """Carica tutti gli stili di quota disponibili."""
        self._dim_types = {}
        
        try:
            collector = DB.FilteredElementCollector(self.doc)\
                         .OfClass(DB.DimensionType)
            
            for dim_type in collector:
                if dim_type and dim_type.Id != DB.ElementId.InvalidElementId:
                    name = dim_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                    if name:
                        type_name = name.AsString()
                        if type_name:
                            self._dim_types[type_name] = dim_type.Id
                            
        except:
            pass
            
        self.cbo_dim_style.Items.Clear()
        for name in sorted(self._dim_types.keys()):
            self.cbo_dim_style.Items.Add(name)
            
        if self.cbo_dim_style.Items.Count > 0:
            self.cbo_dim_style.SelectedIndex = 0
            
    def _init_tag_types(self):
        """Carica tutti i tipi di Material Tag disponibili."""
        self._tag_types = {}
        self._tag_types['(None)'] = None
        
        try:
            collector = DB.FilteredElementCollector(self.doc)\
                         .OfCategory(DB.BuiltInCategory.OST_MaterialTags)\
                         .WhereElementIsElementType()
            
            for tag_type in collector:
                if tag_type and tag_type.Id != DB.ElementId.InvalidElementId:
                    name = get_symbol_name(tag_type)
                    self._tag_types[name] = tag_type.Id
                    
        except:
            pass
            
        self.cbo_tag_type.Items.Clear()
        for name in sorted(self._tag_types.keys()):
            self.cbo_tag_type.Items.Add(name)
            
        if self.cbo_tag_type.Items.Count > 0:
            self.cbo_tag_type.SelectedIndex = 0
    
    def _get_selected_category(self):
        """
        Get the currently selected category enum.
        
        Returns:
            tuple: (category_enum, is_compound) or (None, False)
        """
        idx = self.cbo_category.SelectedIndex
        if idx < 0:
            return None, False
        
        # Check if it's a separator
        selected_text = self.cbo_category.SelectedItem
        if selected_text and "───" in str(selected_text):
            return None, False
        
        # Count through categories
        if idx < len(COMPOUND_CATEGORIES):
            cat_enum, _ = COMPOUND_CATEGORIES[idx]
            return cat_enum, True
        else:
            # Adjust for separator
            adjusted_idx = idx - len(COMPOUND_CATEGORIES) - 1
            if 0 <= adjusted_idx < len(LOADABLE_CATEGORIES):
                cat_enum, _ = LOADABLE_CATEGORIES[adjusted_idx]
                return cat_enum, False
        
        return None, False
    
    def _update_compound_options(self):
        """Enable or disable compound-specific options based on selected category."""
        cat_enum, is_compound = self._get_selected_category()
        
        if is_compound:
            # Enable dimensions and tags sections
            self.border_dimensions.Visibility = Visibility.Visible
            self.border_tags.Visibility = Visibility.Visible
            self.cbo_orientation.IsEnabled = True
        else:
            # Disable and hide dimensions and tags sections
            self.border_dimensions.Visibility = Visibility.Collapsed
            self.border_tags.Visibility = Visibility.Collapsed
            self.cbo_orientation.IsEnabled = False
            # Reset checkboxes
            self.chk_dimensions.IsChecked = False
            self.chk_tags.IsChecked = False
        
    def _update_types_list(self):
        """Update the types list based on selected category."""
        self.lst_types.Items.Clear()
        self._current_types_map = {}
        self._all_type_items = []
        self._filtered_type_items = []
        
        # Clear search
        self.txt_search.Text = ""
        
        cat_enum, is_compound = self._get_selected_category()
        if not cat_enum:
            return
            
        types_dict = self._types_data.get(cat_enum, {})
        
        # Create TypeItem objects
        for name in sorted(types_dict.keys()):
            type_id = types_dict[name]
            item = TypeItem(name, type_id, False)
            self._all_type_items.append(item)
            self._current_types_map[name] = type_id
        
        # Copy to filtered list and populate ListBox
        self._filtered_type_items = self._all_type_items[:]
        self._populate_listbox()
            
        self._update_sort_parameters(cat_enum, types_dict)
        self._update_rows_label()
    
    def _populate_listbox(self):
        """Populate the ListBox with filtered items as checkboxes."""
        self.lst_types.Items.Clear()
        for item in self._filtered_type_items:
            cb = CheckBox()
            cb.Content = item.Name
            cb.IsChecked = item.IsSelected
            cb.Tag = item  # Store reference to TypeItem
            cb.Checked += self.OnCheckboxChanged
            cb.Unchecked += self.OnCheckboxChanged
            self.lst_types.Items.Add(cb)
    
    def _filter_types(self, search_text):
        """Filter types based on search text."""
        if not search_text:
            self._filtered_type_items = self._all_type_items[:]
        else:
            search_lower = search_text.lower()
            self._filtered_type_items = [
                item for item in self._all_type_items
                if search_lower in item.Name.lower()
            ]
        self._populate_listbox()
        
    def _update_sort_parameters(self, category_enum, types_dict):
        """Update the sort parameter combobox based on category."""
        self.cbo_sort_param.Items.Clear()
        self._sort_params = {}
        
        self.cbo_sort_param.Items.Add("(None)")
        self._sort_params["(None)"] = None
        
        if not types_dict:
            self.cbo_sort_param.SelectedIndex = 0
            return
            
        first_type_id = list(types_dict.values())[0]
        first_type = self.doc.GetElement(first_type_id)
        
        if not first_type:
            self.cbo_sort_param.SelectedIndex = 0
            return
            
        param_names = set()
        for param in first_type.Parameters:
            if param.Definition and param.Definition.Name:
                param_names.add(param.Definition.Name)
                
        for name in sorted(param_names):
            self.cbo_sort_param.Items.Add(name)
            self._sort_params[name] = name
            
        self.cbo_sort_param.SelectedIndex = 0
        
    def _update_rows_label(self):
        """Aggiorna la label delle righe calcolate."""
        # Count selected items (IsSelected = True)
        selected_count = sum(1 for item in self._all_type_items if item.IsSelected)
        
        if selected_count == 0:
            self.lbl_rows.Text = "Auto (0)"
            return
            
        try:
            columns = int(self.txt_columns.Text) if self.txt_columns.Text else 1
            if columns < 1:
                columns = 1
            rows = int(math.ceil(float(selected_count) / float(columns)))
            self.lbl_rows.Text = str(rows)
        except:
            self.lbl_rows.Text = "Auto"
            
    def _get_selected_types(self):
        """Restituisce la lista degli ElementId dei tipi selezionati."""
        selected_ids = []
        for item in self._all_type_items:
            if item.IsSelected:
                selected_ids.append(item.TypeId)
        return selected_ids
        
    def _validate_input(self):
        """Validate all form inputs."""
        # Check for separator selection
        cat_enum, _ = self._get_selected_category()
        if not cat_enum:
            return False, "Please select a valid category."
        
        selected = self._get_selected_types()
        if not selected:
            return False, "Select at least one type."
            
        try:
            columns = int(self.txt_columns.Text)
            if columns < 1:
                return False, "Number of columns must be at least 1."
        except:
            return False, "Enter a valid number for columns."
            
        try:
            offset_x = float(self.txt_offset_x.Text)
            if offset_x <= 0:
                return False, "Offset X must be greater than 0."
        except:
            return False, "Enter a valid number for Offset X."
            
        try:
            offset_y = float(self.txt_offset_y.Text)
            if offset_y <= 0:
                return False, "Offset Y must be greater than 0."
        except:
            return False, "Enter a valid number for Offset Y."
            
        # Verify dimension settings if enabled (only for compound)
        if self.chk_dimensions.IsChecked:
            if self.cbo_dim_style.SelectedIndex < 0:
                return False, "Select a dimension style."
            try:
                dim_offset = float(self.txt_dim_offset.Text)
                if dim_offset < 0:
                    return False, "Dimension offset must be 0 or greater."
            except:
                return False, "Enter a valid number for dimension offset."
                
        # Verify tag settings if enabled (only for compound)
        if self.chk_tags.IsChecked:
            if self.cbo_tag_type.SelectedIndex < 0:
                return False, "Select a material tag type."
            try:
                tag_offset = float(self.txt_tag_offset.Text)
                if tag_offset < 0:
                    return False, "Tag offset must be 0 or greater."
            except:
                return False, "Enter a valid number for tag offset."
            try:
                tag_spacing = float(self.txt_tag_spacing.Text)
                if tag_spacing <= 0:
                    return False, "Tag spacing must be greater than 0."
            except:
                return False, "Enter a valid number for tag spacing."
                
        return True, ""
        
    def _build_config(self):
        """Build the LegendConfig object from form values."""
        cat_enum, is_compound = self._get_selected_category()
        self.config.category = cat_enum
        self.config.is_compound_category = is_compound
        
        self.config.selected_types = self._get_selected_types()
        
        if self.cbo_sort_param.SelectedIndex > 0:
            sort_name = self.cbo_sort_param.SelectedItem
            self.config.sort_parameter = self._sort_params.get(sort_name)
        else:
            self.config.sort_parameter = None
        
        self.config.columns = int(self.txt_columns.Text)
        self.config.offset_x_mm = float(self.txt_offset_x.Text)
        self.config.offset_y_mm = float(self.txt_offset_y.Text)
        
        self.config.is_horizontal = (self.cbo_orientation.SelectedIndex == 0)
        
        # Dimensions (only if compound and enabled)
        if is_compound:
            self.config.insert_dimensions = self.chk_dimensions.IsChecked
            if self.config.insert_dimensions and self.cbo_dim_style.SelectedIndex >= 0:
                dim_name = self.cbo_dim_style.SelectedItem
                self.config.dimension_type_id = self._dim_types.get(dim_name)
                self.config.dim_offset_mm = float(self.txt_dim_offset.Text)
                self.config.dim_position_above_right = (self.cbo_dim_position.SelectedIndex == 0)
                
            # Material Tags
            self.config.insert_tags = self.chk_tags.IsChecked
            if self.config.insert_tags and self.cbo_tag_type.SelectedIndex >= 0:
                tag_name = self.cbo_tag_type.SelectedItem
                self.config.tag_type_id = self._tag_types.get(tag_name)
                self.config.tag_offset_mm = float(self.txt_tag_offset.Text)
                self.config.tag_spacing_mm = float(self.txt_tag_spacing.Text)
        else:
            # Disable dimensions and tags for non-compound
            self.config.insert_dimensions = False
            self.config.insert_tags = False
            
        self.config.calculate_internal_units(self.doc)
        self.config.calculate_rows(len(self.config.selected_types))
            
    # =========================================================================
    # EVENT HANDLERS
    # =========================================================================
    
    def OnCategoryChanged(self, sender, args):
        """Handler per cambio categoria."""
        # Skip if separator selected
        selected_text = self.cbo_category.SelectedItem
        if selected_text and "───" in str(selected_text):
            # Move to next valid item
            if self.cbo_category.SelectedIndex < len(COMPOUND_CATEGORIES):
                self.cbo_category.SelectedIndex = len(COMPOUND_CATEGORIES) + 1
            return
            
        self._update_types_list()
        self._update_compound_options()
        
    def OnTypesSelectionChanged(self, sender, args):
        """Handler per cambio selezione tipi."""
        self._update_rows_label()
    
    def OnSearchTextChanged(self, sender, args):
        """Handler per cambio testo di ricerca."""
        self._filter_types(self.txt_search.Text)
    
    def OnClearSearch(self, sender, args):
        """Handler per cancella ricerca."""
        self.txt_search.Text = ""
        self._filter_types("")
    
    def OnResizeList(self, sender, args):
        """Handler per ridimensionamento lista."""
        new_height = self.lst_types.Height + args.VerticalChange
        if new_height >= 80 and new_height <= 500:
            self.lst_types.Height = new_height
    
    def OnCheckboxChanged(self, sender, args):
        """Handler per cambio stato checkbox."""
        cb = sender
        if cb and cb.Tag:
            item = cb.Tag
            item.IsSelected = cb.IsChecked
            self._update_rows_label()
        
    def OnSelectAll(self, sender, args):
        """Handler per seleziona tutti (solo filtrati visibili)."""
        for item in self._filtered_type_items:
            item.IsSelected = True
        # Update checkboxes in UI
        for cb in self.lst_types.Items:
            if hasattr(cb, 'IsChecked'):
                cb.IsChecked = True
        self._update_rows_label()
        
    def OnDeselectAll(self, sender, args):
        """Handler per deseleziona tutti (solo filtrati visibili)."""
        for item in self._filtered_type_items:
            item.IsSelected = False
        # Update checkboxes in UI
        for cb in self.lst_types.Items:
            if hasattr(cb, 'IsChecked'):
                cb.IsChecked = False
        self._update_rows_label()
        
    def OnGridParameterChanged(self, sender, args):
        """Handler per cambio parametri griglia."""
        self._update_rows_label()
        
    def OnDimensionsChecked(self, sender, args):
        """Handler per checkbox quote attivato."""
        self.cbo_dim_style.IsEnabled = True
        
    def OnDimensionsUnchecked(self, sender, args):
        """Handler per checkbox quote disattivato."""
        self.cbo_dim_style.IsEnabled = False
        
    def OnTagsChecked(self, sender, args):
        """Handler per checkbox tags attivato."""
        self.cbo_tag_type.IsEnabled = True
        
    def OnTagsUnchecked(self, sender, args):
        """Handler per checkbox tags disattivato."""
        self.cbo_tag_type.IsEnabled = False
        
    def OnOK(self, sender, args):
        """Handler per pulsante OK."""
        is_valid, error_msg = self._validate_input()
        if not is_valid:
            MessageBox.Show(error_msg, "Validation Error")
            return
            
        self._build_config()
        
        self.result = True
        self.Close()
        
    def OnCancel(self, sender, args):
        """Handler per pulsante Annulla."""
        self.result = False
        self.Close()


# =============================================================================
# PUBLIC FUNCTION
# =============================================================================

def show_config_form(doc):
    """
    Mostra la form di configurazione e restituisce i parametri.
    
    Args:
        doc: Il documento Revit corrente
        
    Returns:
        LegendConfig or None: La configurazione se OK, None se annullato
    """
    form = LegendConfigForm(doc)
    form.ShowDialog()
    
    if form.result:
        return form.config
    return None
