# -*- coding: utf-8 -*-
"""
BIM Data Extractor per pyRevit
Estrae dati da modelli Revit e genera CSV per Power BI

Autore: Claude AI Assistant
Versione: 1.0
"""

# ==============================================================================
# IMPORTS
# ==============================================================================

import csv
import os
import sys
import io
import json
from datetime import datetime
from collections import defaultdict

# pyRevit imports
from pyrevit import script, forms, revit, DB, HOST_APP

# Revit API imports
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    WorksetKind,
    RevitLinkInstance,
    RevitLinkType,
    ImportInstance,
    CADLinkType,
    View,
    ViewPlan,
    ViewSection,
    View3D,
    ViewDrafting,
    ViewSchedule,
    ViewSheet,
    ParameterFilterElement,
    DesignOption,
    Transaction,
    ModelPathUtils,
    OpenOptions,
    DetachFromCentralOption,
    WorksetConfiguration,
    WorksetConfigurationOption,
    WorksharingUtils,
    Element,
    Category,
    FamilyInstance,
    Document,
    StartingViewSettings,
    Level,
    Grid,
    BasePoint,
    InstanceBinding,
    TypeBinding
)

# Per i Purgeable elements
from Autodesk.Revit.DB import (
    Family,
    ElementType,
    Material,
    GraphicsStyle,
    FillPatternElement,
    LinePatternElement,
    GroupType,
    Group
)

# WPF imports per interfaccia XAML
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')

from System.Windows import Window, Visibility
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode
from System.Windows.Forms import OpenFileDialog, FolderBrowserDialog, DialogResult

# ==============================================================================
# CONFIGURAZIONE
# ==============================================================================

OUTPUT = script.get_output()
LOGGER = script.get_logger()

# Delimitatore CSV (punto e virgola per compatibilità Excel italiano)
CSV_DELIMITER = ";"
CSV_ENCODING = "utf-8"

# Data estrazione
EXTRACTION_DATE = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ==============================================================================
# CLASSI HELPER
# ==============================================================================

class FileProcessor:
    """Gestisce l'elaborazione di un singolo file Revit."""
    
    def __init__(self, file_path, app):
        self.file_path = file_path
        self.app = app
        self.doc = None
        self.file_name = os.path.basename(file_path).strip()
        self.is_workshared = False
        self._opened_by_script = False
        self.opening_time = ""  # Tempo di apertura in formato mm:ss
        
        # Leggi file size PRIMA di aprire il file (per evitare lock)
        try:
            # Formatta con massimo 2 decimali
            self.file_size_mb = round(os.path.getsize(file_path) / (1024.0 * 1024.0), 2)
        except:
            self.file_size_mb = 0.00
    
    def open_document(self):
        """Apre il documento Revit."""
        # Verifica se il file è già aperto
        for doc in self.app.Documents:
            if doc.PathName and os.path.normpath(doc.PathName).lower() == os.path.normpath(self.file_path).lower():
                OUTPUT.print_md("⚠️ **ATTENZIONE**: Il file **{}** è già aperto in Revit. Saltato.".format(self.file_name))
                return False
        
        try:
            model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(self.file_path)
            
            # Configura le opzioni di apertura (apri sempre con detach per sicurezza)
            open_options = OpenOptions()
            open_options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
            
            # Configura workset: chiudi tutti i workset user-defined per velocizzare l'apertura.
            # Tutti i metadati (warning, viste, link, livelli, griglia, ecc.) rimangono leggibili.
            workset_config = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
            open_options.SetOpenWorksetsConfiguration(workset_config)
            
            # Apri il documento con misurazione del tempo
            import time
            start_time = time.time()
            self.doc = self.app.OpenDocumentFile(model_path, open_options)
            end_time = time.time()
            self._opened_by_script = True
            
            # Calcola il tempo di apertura in formato mm:ss
            elapsed_seconds = int(end_time - start_time)
            minutes = elapsed_seconds // 60
            seconds = elapsed_seconds % 60
            self.opening_time = "{:02d}:{:02d}".format(minutes, seconds)
            
            # CORRETTO: Determina se è workshared DOPO aver aperto il documento
            try:
                self.is_workshared = self.doc.IsWorkshared
            except:
                self.is_workshared = False
            
            return True
            
        except Exception as e:
            OUTPUT.print_md("❌ **ERRORE** apertura file **{}**: {}".format(self.file_name, str(e)))
            return False
    
    def close_document(self):
        """Chiude il documento se aperto dallo script."""
        if self.doc and self._opened_by_script:
            try:
                self.doc.Close(False)  # False = non salvare
            except:
                pass


class CSVWriter:
    """Gestisce la scrittura dei file CSV."""
    
    # Definizione intestazioni per tutte le tabelle
    # NOTA: Queste intestazioni DEVONO corrispondere esattamente ai dizionari restituiti dalle funzioni extract_*
    TABLE_HEADERS = {
        'TAB_Files': ['FileName', 'FileDiscipline', 'FilePath', 'FileSize_MB', 'OpeningTime', 'StartingPage_Name', 
                      'ExtractionDate', 'IsWorkshared', 'ProjectBasePoint', 'SurveyPoint', 'AngleTrueNorth',
                      'Elements', 'Families', 'Types', 'PurgeableElements', 'Warnings',
                      'Views_HasTemplate(N)', 'Views_OnSheet', 'Sheets',
                      'Links_RVT_Pinned(N)', 'Links_DWG_Pinned(N)', 'Levels_Pinned(N)',
                      'Levels_Monitored(N)', 'Grids_Pinned(N)', 'Grids_Monitored(N)',
                      'Views_VC(N)', 'Views_OnSheet(N)', 'Sheets_VC1(N)',
                      'Tags_Host(N)', 'RoomsAreasSpaces_ConfinedPlaced(N)',
                      'ViewTemplates_CompliantName(N)', 'Filters_CompliantName(N)', 'StartingView_Correct',
                      'Links_RVT_LinkedBySharedCoordinates(N)', 'Links_DWG_IsViewSpecific', 'Warnings_DuplicateInstance',
                      'ModelInPlace', 'HideInView', 'KPI_HealthScore'],
        'TAB_Links': ['LinkKey', 'LinkID', 'FileName', 'LinkName', 'LinkDiscipline', 'LinkPath', 'LinkFileName', 'LastSavedDate', 
                      'SharedSite', 'ProjectBasePoint', 'SurveyPoint', 'Workset_(i)', 'Workset_(t)', 
                      'IsPinned', 'LinkType'],
        'TAB_Views': ['ViewKey', 'ViewID', 'FileName', 'ViewName', 'ViewType', 'ViewTemplateID', 'ViewTemplateName', 
                      'IsTemplate', 'HasScopeBox', 'IsDependent', 'PhaseFilter', 'Phase', 
                      'ReferencingSheet', 'TitleOnSheet', 'ViewChapter1', 'ViewChapter2', 'HiddenElements'],
        'TAB_Warnings': ['WarningKey', 'WarningID', 'FileName', 'WarningDescription', 'WarningSeverity', 
                         'WarningDescValidation', 'WarningFailureGUID', 'WarningType', 'ElementID'],
        'TAB_Worksets_UserDefined': ['WorksetKey', 'WorksetID', 'FileName', 'WorksetName', 'IsVisibleInAllViews', 
                                      'Owner', 'IsOpen'],
        'TAB_Sheets': ['SheetKey', 'SheetID', 'FileName', 'SheetNumber', 'SheetName', 'CurrentRevisionNumber', 
                       'CurrentRevisionName', 'ViewChapter1', 'ViewChapter2'],
        'TAB_ViewTemplates': ['ViewTemplateKey', 'ViewTemplateID', 'FileName', 'TemplateName', 'ViewType', 'IsUsed'],
        'TAB_Materials': ['MaterialKey', 'MaterialID', 'FileName', 'MaterialName', 'MaterialClass', 'MaterialCategory', 
                          'MaterialDescription', 'MaterialComments', 'MaterialKeywords', 
                          'MaterialManufacturer', 'MaterialModel', 'MaterialAssetName', 'IsUsed'],
        'TAB_Levels': ['LevelKey', 'LevelID', 'FileName', 'LevelName', 'LevelType', 'LevelOffset', 'IsMonitor', 
                       'MonitorFileName', 'MonitorLevel', 'IsPinned', 'ScopeBox', 'Workset'],
        'TAB_ScopeBoxes': ['ScopeBoxKey', 'ScopeBoxID', 'FileName', 'ScopeBoxName', 'Workset', 'IsPinned'],
        'TAB_Grids': ['GridKey', 'GridID', 'FileName', 'GridName', 'GridType', 'IsPinned', 'IsMonitored', 
                      'MonitorFileName', 'MonitorGrid', 'ScopeBox', 'Workset'],
        'TAB_Filters': ['FilterKey', 'FilterID', 'FileName', 'FilterName', 'IsUsed'],
        'TAB_Families': [
                        'FamilyKey', 'FamilyID', 'FileName', 'FamilyName', 'Category',
                        'IsInPlace', 'IsSystemFamily', 'IsUsed',
                        ],
        'TAB_Types': [
                        'TypeKey', 'TypeID', 'FileName', 'FamilyKey', 'TypeName', 'Family&Type',
                        'Description', 'TypeMark', 'IsUsed',
                        'Export Type to IFC As', 'Type IFC Predefined Type',
                        'Classification.Uniclass.Pr.Number', 'Classification.Uniclass.Pr.Description',
                        'Classification.Uniclass.Ss.Number', 'Classification.Uniclass.Ss.Description',
                        'ClassificationCode[Type]', 'ClassificationCode(2)[Type]', 'ClassificationCode(3)[Type]',
                        ],
        'TAB_Instances': [
                        'ElementKey', 'ElementID', 'FileName', 'TypeKey',
                        'WorksetKey', 'WorksetName', 'PhaseCreation', 'PhaseDemolished',
                        'Export to IFC As', 'IFC Predefined Type',
                        'ClassificationCode', 'ClassificationCode(2)', 'ClassificationCode(3)',
                        ],
        'TAB_PurgeableElements': ['PurgeableElementKey', 'PurgeableElementID', 'FileName', 'Category', 'PurgeableElementName', 'RevitCategory'],
        'TAB_Parameters': ['FamilyKey', 'FamilyID', 'FileName', 'InstanceID', 'ParameterName', 'IsShared', 'TypeOrInstance', 'Param_GUID'],
        'TAB_ObjectStyle': ['FileName', 'FamilyKey', 'ObjectStyle'],
        'TAB_Rooms': ['RoomKey', 'RoomID', 'FileName', 'RoomName', 'RoomNumber', 'Level', 'Area_sqm', 
                      'Perimeter_m', 'Height_m', 'Volume_mc', 'AreaString', 'IsPlaced', 'IsEnclosed', 'IsRedundant', 'Phase', 'Workset'],
        'TAB_Spaces': ['SpaceKey', 'SpaceID', 'FileName', 'SpaceName', 'SpaceNumber', 'Level', 'Area_sqm', 
                       'Height_m', 'Volume_mc', 'Zone', 'AreaString', 'IsPlaced', 'IsEnclosed', 'IsRedundant', 'Phase', 'Workset'],
        'TAB_Areas': ['AreaKey', 'AreaID', 'FileName', 'AreaName', 'AreaNumber', 'AreaType', 'Area_sqm', 
                      'Perimeter_m', 'Level', 'IsPlaced', 'Workset'],
        'TAB_Tags': ['TagKey', 'TagID', 'FileName', 'ViewID', 'ViewKey', 'FamilyName', 'TypeName', 'TagCategory', 'HasHost'],
        'TAB_HealthChecks': ['CheckID', 'CheckDescription', 'FileName', 'CheckPassed', 'Value', 'Score', 'MaxScore'],
        'TAB_DataValidation_Families': ['FamilyKey', 'FieldName', 'FieldValue', 'Status'],
        'TAB_DataValidation_Types': ['TypeKey', 'FieldName', 'FieldValue', 'Status'],
        'TAB_DataValidation_Instances': ['ElementKey', 'FieldName', 'FieldValue', 'Status'],
    }

    
    def __init__(self, output_folder):
        self.output_folder = output_folder
        self.data = defaultdict(list)
        self.blocked_files = []  # Lista dei file che non sono stati scritti
    
    def add_row(self, table_name, row_dict):
        """Aggiunge una riga a una tabella."""
        self.data[table_name].append(row_dict)
    
    def add_rows(self, table_name, rows_list):
        """Aggiunge più righe a una tabella."""
        self.data[table_name].extend(rows_list)
    
    def write_all(self, custom_instance_params=None, custom_type_params=None):
        """Scrive tutti i CSV, anche quelli vuoti con sole intestazioni.
        
        Args:
            custom_instance_params: Lista di nomi di parametri custom di istanza
                                    da aggiungere alle intestazioni di TAB_Instances.
            custom_type_params: Lista di nomi di parametri custom di tipo
                                da aggiungere alle intestazioni di TAB_Types.
        """
        # Estendi intestazioni TAB_Instances con parametri custom di istanza
        original_instance_headers = None
        if custom_instance_params:
            base_headers = list(self.TABLE_HEADERS['TAB_Instances'])
            for p in custom_instance_params:
                if p not in base_headers:
                    base_headers.append(p)
            original_instance_headers = self.TABLE_HEADERS['TAB_Instances']
            self.TABLE_HEADERS['TAB_Instances'] = base_headers
        
        # Estendi intestazioni TAB_Types con parametri custom di tipo
        original_type_headers = None
        if custom_type_params:
            base_headers = list(self.TABLE_HEADERS['TAB_Types'])
            for p in custom_type_params:
                if p not in base_headers:
                    base_headers.append(p)
            original_type_headers = self.TABLE_HEADERS['TAB_Types']
            self.TABLE_HEADERS['TAB_Types'] = base_headers
        
        # Scrivi tutti i CSV definiti in TABLE_HEADERS
        for table_name in self.TABLE_HEADERS.keys():
            rows = self.data.get(table_name, [])
            self._write_csv(table_name, rows)
        
        # Ripristina le intestazioni originali
        if original_instance_headers is not None:
            self.TABLE_HEADERS['TAB_Instances'] = original_instance_headers
        if original_type_headers is not None:
            self.TABLE_HEADERS['TAB_Types'] = original_type_headers
    
    def _write_csv(self, table_name, rows):
        """Scrive un singolo CSV (compatibile IronPython) con gestione errori."""
        file_path = os.path.join(self.output_folder, "{}.csv".format(table_name))
        csv_filename = "{}.csv".format(table_name)
        
        # Usa le intestazioni predefinite se disponibili, altrimenti ricavale dai dati
        if table_name in self.TABLE_HEADERS:
            fieldnames = self.TABLE_HEADERS[table_name]
        else:
            # Fallback: ottieni i campi dai dati (comportamento originale)
            fieldnames = []
            for row in rows:
                for key in row.keys():
                    if key not in fieldnames:
                        fieldnames.append(key)
            
            if not fieldnames:
                # Nessuna intestazione disponibile, salta
                return
        
        try:
            # IronPython compatibile: usa io.open con encoding e lineterminator per evitare righe vuote
            with io.open(file_path, 'w', encoding=CSV_ENCODING, newline='') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=CSV_DELIMITER, 
                                       extrasaction='ignore', lineterminator='\n')
                writer.writeheader()
                
                if rows:
                    writer.writerows(rows)
            
            if rows:
                OUTPUT.print_md("✅ Scritto: **{}** ({} righe)".format(table_name, len(rows)))
            else:
                OUTPUT.print_md("✅ Scritto: **{}** (solo intestazioni)".format(table_name))
            
        except IOError as e:
            if "being used by another process" in str(e) or "cannot access the file" in str(e):
                # File bloccato - aggiungi alla lista
                self.blocked_files.append(csv_filename)
                OUTPUT.print_md("⚠️ **{}** saltato (file aperto)".format(csv_filename))
                LOGGER.warning("File bloccato: {}".format(file_path))
            else:
                # Altro tipo di IOError
                OUTPUT.print_md("❌ **ERRORE** scrittura **{}**: {}".format(table_name, str(e)))
                LOGGER.error("Errore scrittura CSV: {}".format(str(e)))
        except Exception as e:
            OUTPUT.print_md("❌ **ERRORE** scrittura **{}**: {}".format(table_name, str(e)))
            LOGGER.error("Errore scrittura CSV: {}".format(str(e)))


# ==============================================================================
# FORM XAML - INTERFACCIA UTENTE
# ==============================================================================

JSON_SETUP_FILENAME = "ModelReport_ExportSetup.json"


class ModelReportForm(Window):
    """Form XAML per la configurazione del BIM Data Extractor."""
    
    def __init__(self):
        self.result = False
        self.selected_files = []
        self.output_folder = ""
        self.custom_params = []
        self.validation_rules = {}  # dict {param: {allowed_values}}
        self.discipline_rules = []  # Lista di tuple (code, desc)

        # Stato JSON
        self._json_params = []
        self._json_loaded = False
        self._modifying_setup = False
        self._json_validation_rules = {}  # Regole validazione dal JSON
        self._json_validation_loaded = False
        self._modifying_validation = False
        self._json_discipline_rules = []  # Regole discipline dal JSON
        self._discipline_rows = []  # Lista di tuple (TextBox_code, TextBox_desc)
        self._validation_rows = []  # Lista di tuple (TextBox_param, get_values_fn, Button_del)
        self._param_rows = []  # Lista di tuple (TextBox_param, Button_del)
        
        # Carica XAML
        self._load_xaml()
    
    def _load_xaml(self):
        """Carica e parsa il file XAML."""
        script_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(script_dir, 'ModelReportForm.xaml')
        
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
        
        # Stato iniziale: pannelli nascosti
        self.pnl_params_section.Visibility = Visibility.Collapsed
        self.pnl_json_setup.Visibility = Visibility.Collapsed
        self.pnl_validation_edit.Visibility = Visibility.Collapsed
        self.pnl_validation_json.Visibility = Visibility.Collapsed
        
        # Discipline: sezione nascosta finche' non si seleziona la cartella output
        self.pnl_discipline_section.Visibility = Visibility.Collapsed
        self.pnl_discipline_loaded.Visibility = Visibility.Collapsed
        self.pnl_discipline_edit.Visibility = Visibility.Collapsed
    
    def _find_controls(self, root):
        """Trova tutti i controlli nel XAML."""
        # File selection
        self.btn_select_files = root.FindName('btn_select_files')
        self.lst_files = root.FindName('lst_files')
        self.lbl_file_count = root.FindName('lbl_file_count')
        self.btn_remove_file = root.FindName('btn_remove_file')
        self.btn_clear_files = root.FindName('btn_clear_files')
        # Output folder
        self.txt_output_folder = root.FindName('txt_output_folder')
        self.btn_browse_folder = root.FindName('btn_browse_folder')
        # Discipline
        self.pnl_discipline_section = root.FindName('pnl_discipline_section')
        self.pnl_discipline_loaded = root.FindName('pnl_discipline_loaded')
        self.lst_discipline_loaded = root.FindName('lst_discipline_loaded')
        self.chk_modify_discipline = root.FindName('chk_modify_discipline')
        self.pnl_discipline_edit = root.FindName('pnl_discipline_edit')
        self.pnl_discipline_rows = root.FindName('pnl_discipline_rows')
        self.btn_discipline_add = root.FindName('btn_discipline_add')
        self.btn_discipline_remove = root.FindName('btn_discipline_remove')
        # JSON setup panel
        self.pnl_json_setup = root.FindName('pnl_json_setup')
        self.lst_json_params = root.FindName('lst_json_params')
        self.chk_modify_setup = root.FindName('chk_modify_setup')
        # Custom params (tabular)
        self.pnl_params_section = root.FindName('pnl_params_section')
        self.pnl_param_rows = root.FindName('pnl_param_rows')
        self.btn_param_add = root.FindName('btn_param_add')
        # Validation
        self.chk_validation = root.FindName('chk_validation')
        self.pnl_validation_edit = root.FindName('pnl_validation_edit')
        self.pnl_validation_rows = root.FindName('pnl_validation_rows')
        self.btn_validation_add = root.FindName('btn_validation_add')
        self.pnl_validation_json = root.FindName('pnl_validation_json')
        self.lst_validation_loaded = root.FindName('lst_validation_loaded')
        self.chk_modify_validation = root.FindName('chk_modify_validation')
        # Buttons
        self.btn_run = root.FindName('btn_run')
        self.btn_cancel = root.FindName('btn_cancel')
    
    def _wire_events(self):
        """Collega gli eventi ai metodi handler."""
        self.btn_select_files.Click += self.OnSelectFiles
        self.btn_remove_file.Click += self.OnRemoveFile
        self.btn_clear_files.Click += self.OnClearFiles
        self.btn_browse_folder.Click += self.OnBrowseFolder
        self.btn_discipline_add.Click += self.OnDisciplineAdd
        self.btn_discipline_remove.Click += self.OnDisciplineRemove
        self.chk_modify_discipline.Checked += self.OnModifyDisciplineChecked
        self.chk_modify_discipline.Unchecked += self.OnModifyDisciplineUnchecked
        self.chk_modify_setup.Checked += self.OnModifySetupChecked
        self.chk_modify_setup.Unchecked += self.OnModifySetupUnchecked
        self.btn_param_add.Click += self.OnParamAdd
        self.chk_validation.Checked += self.OnValidationChecked
        self.chk_validation.Unchecked += self.OnValidationUnchecked
        self.btn_validation_add.Click += self.OnValidationAdd
        self.chk_modify_validation.Checked += self.OnModifyValidationChecked
        self.chk_modify_validation.Unchecked += self.OnModifyValidationUnchecked
        self.btn_run.Click += self.OnRun
        self.btn_cancel.Click += self.OnCancel
    
    def _add_discipline_row(self, code="", desc=""):
        """Aggiunge una riga alla tabella discipline con due TextBox."""
        from System.Windows.Controls import TextBox, Grid, ColumnDefinition
        from System.Windows import GridLength, GridUnitType, Thickness
        
        row_grid = Grid()
        row_grid.Margin = Thickness(0, 1, 0, 1)
        
        col1 = ColumnDefinition()
        col1.Width = GridLength(1, GridUnitType.Star)
        col2 = ColumnDefinition()
        col2.Width = GridLength(10, GridUnitType.Pixel)
        col3 = ColumnDefinition()
        col3.Width = GridLength(1.5, GridUnitType.Star)
        row_grid.ColumnDefinitions.Add(col1)
        row_grid.ColumnDefinitions.Add(col2)
        row_grid.ColumnDefinitions.Add(col3)
        
        txt_code = TextBox()
        txt_code.Text = code
        txt_code.Padding = Thickness(3, 2, 3, 2)
        txt_code.SetValue(Grid.ColumnProperty, 0)
        
        txt_desc = TextBox()
        txt_desc.Text = desc
        txt_desc.Padding = Thickness(3, 2, 3, 2)
        txt_desc.SetValue(Grid.ColumnProperty, 2)
        
        row_grid.Children.Add(txt_code)
        row_grid.Children.Add(txt_desc)
        
        self.pnl_discipline_rows.Children.Add(row_grid)
        self._discipline_rows.append((txt_code, txt_desc))
    
    def _get_discipline_rules(self):
        """Legge tutte le righe discipline dal form e restituisce una lista di tuple."""
        rules = []
        for txt_code, txt_desc in self._discipline_rows:
            code = txt_code.Text.strip()
            desc = txt_desc.Text.strip()
            if code and desc:
                rules.append((code, desc))
        return rules
    
    def _populate_discipline_rows(self, rules):
        """Popola le righe discipline da una lista di tuple (code, desc)."""
        self.pnl_discipline_rows.Children.Clear()
        self._discipline_rows = []
        
        for code, desc in rules:
            self._add_discipline_row(code, desc)
        
        if not rules:
            self._add_discipline_row()
    
    def _add_param_row(self, param_name=""):
        """Aggiunge una riga alla tabella parametri: TextBox + pulsante cancella."""
        from System.Windows.Controls import Grid, ColumnDefinition, TextBox, Button
        from System.Windows import GridLength, GridUnitType, Thickness

        row_grid = Grid()
        row_grid.Margin = Thickness(0, 1, 0, 1)

        col1 = ColumnDefinition()
        col1.Width = GridLength(1, GridUnitType.Star)
        col2 = ColumnDefinition()
        col2.Width = GridLength(30, GridUnitType.Pixel)
        row_grid.ColumnDefinitions.Add(col1)
        row_grid.ColumnDefinitions.Add(col2)

        txt_param = TextBox()
        txt_param.FontSize = 11
        txt_param.Padding = Thickness(3, 2, 3, 2)
        txt_param.Text = param_name
        txt_param.SetValue(Grid.ColumnProperty, 0)

        btn_del = Button()
        btn_del.Content = "✕"
        btn_del.Width = 24
        btn_del.Height = 24
        btn_del.FontSize = 11
        btn_del.ToolTip = "Rimuovi parametro"
        btn_del.SetValue(Grid.ColumnProperty, 1)

        row_grid.Children.Add(txt_param)
        row_grid.Children.Add(btn_del)

        self.pnl_param_rows.Children.Add(row_grid)
        row_tuple = (txt_param, btn_del)
        self._param_rows.append(row_tuple)

        # Handler per il pulsante delete
        def on_delete(sender, args, grid=row_grid, tup=row_tuple):
            self.pnl_param_rows.Children.Remove(grid)
            if tup in self._param_rows:
                self._param_rows.remove(tup)
        btn_del.Click += on_delete
    
    def _get_custom_params(self):
        """Legge i nomi parametri selezionati dalla tabella.
        
        Returns:
            list: lista di nomi parametri (stringhe)
        """
        params = []
        for txt_param, btn_del in self._param_rows:
            name = txt_param.Text.strip() if txt_param.Text else ''
            if name:
                params.append(name)
        return params
    
    def _populate_param_rows(self, param_list):
        """Popola le righe parametri da una lista di nomi."""
        self.pnl_param_rows.Children.Clear()
        self._param_rows = []
        
        for name in param_list:
            self._add_param_row(name)
        
        if not param_list:
            self._add_param_row()
    
    def _show_values_editor(self, current_values):
        """Apre una finestra popup per modificare la lista dei valori ammessi.

        Returns:
            list oppure None se l'utente ha annullato.
        """
        from System.Windows.Controls import (
            TextBox, Button, StackPanel, ScrollViewer,
            Grid, ColumnDefinition, TextBlock
        )
        from System.Windows import (
            Window, Thickness, WindowStartupLocation,
            SizeToContent, HorizontalAlignment, VerticalAlignment,
            GridLength, GridUnitType
        )
        from System.Windows.Controls import Orientation

        editor = Window()
        editor.Title = "Valori ammessi"
        editor.Width = 320
        editor.SizeToContent = SizeToContent.Height
        editor.MaxHeight = 500
        editor.WindowStartupLocation = WindowStartupLocation.CenterOwner
        editor.Owner = self
        editor.ResizeMode = editor.ResizeMode.CanResizeWithGrip

        result = [None]
        rows = []

        main_panel = StackPanel()
        main_panel.Margin = Thickness(12)

        lbl = TextBlock()
        lbl.Text = "Un valore per riga:"
        lbl.Margin = Thickness(0, 0, 0, 6)
        main_panel.Children.Add(lbl)

        scroll = ScrollViewer()
        scroll.MaxHeight = 300
        scroll.VerticalScrollBarVisibility = scroll.VerticalScrollBarVisibility.Auto
        pnl_rows = StackPanel()
        scroll.Content = pnl_rows
        main_panel.Children.Add(scroll)

        def add_value_row(value=""):
            txt = TextBox()
            txt.Text = value
            txt.Margin = Thickness(0, 2, 0, 2)
            txt.Padding = Thickness(3, 2, 3, 2)
            pnl_rows.Children.Add(txt)
            rows.append(txt)

        for v in (current_values or []):
            add_value_row(v)
        if not current_values:
            add_value_row()

        btn_row = StackPanel()
        btn_row.Orientation = Orientation.Horizontal
        btn_row.Margin = Thickness(0, 6, 0, 0)

        btn_add = Button()
        btn_add.Content = "＋ Aggiungi"
        btn_add.Padding = Thickness(8, 3, 8, 3)
        btn_add.Margin = Thickness(0, 0, 5, 0)

        def on_add(s, a):
            add_value_row()
            scroll.ScrollToBottom()
        btn_add.Click += on_add

        btn_remove = Button()
        btn_remove.Content = "－ Rimuovi ultimo"
        btn_remove.Padding = Thickness(8, 3, 8, 3)

        def on_remove(s, a):
            if rows:
                pnl_rows.Children.Remove(rows[-1])
                rows.pop()
        btn_remove.Click += on_remove

        btn_row.Children.Add(btn_add)
        btn_row.Children.Add(btn_remove)
        main_panel.Children.Add(btn_row)

        ok_row = StackPanel()
        ok_row.Orientation = Orientation.Horizontal
        ok_row.HorizontalAlignment = HorizontalAlignment.Right
        ok_row.Margin = Thickness(0, 12, 0, 0)

        btn_ok = Button()
        btn_ok.Content = "OK"
        btn_ok.Width = 80
        btn_ok.Height = 28
        btn_ok.Margin = Thickness(0, 0, 8, 0)
        btn_ok.IsDefault = True

        def on_ok(s, a):
            result[0] = [txt.Text.strip() for txt in rows if txt.Text.strip()]
            editor.Close()
        btn_ok.Click += on_ok

        btn_cancel = Button()
        btn_cancel.Content = "Annulla"
        btn_cancel.Width = 80
        btn_cancel.Height = 28
        btn_cancel.IsCancel = True

        def on_cancel(s, a):
            editor.Close()
        btn_cancel.Click += on_cancel

        ok_row.Children.Add(btn_ok)
        ok_row.Children.Add(btn_cancel)
        main_panel.Children.Add(ok_row)

        editor.Content = main_panel
        editor.ShowDialog()
        return result[0]

    def _add_validation_row(self, param_name="", values_list=None):
        """Aggiunge una riga alla tabella validazione: TextBox param, label valori (read-only), bottone Edit, bottone cancella."""
        from System.Windows.Controls import TextBox, TextBlock, Grid, ColumnDefinition, Button
        from System.Windows import GridLength, GridUnitType, Thickness
        from System.Windows.Media import Brushes

        if values_list is None:
            values_list = []

        # Contenitore mutabile per i valori (closure-safe)
        mutable_values = [list(values_list)]

        row_grid = Grid()
        row_grid.Margin = Thickness(0, 2, 0, 2)

        col1 = ColumnDefinition()
        col1.Width = GridLength(1, GridUnitType.Star)
        col2 = ColumnDefinition()
        col2.Width = GridLength(8, GridUnitType.Pixel)
        col3 = ColumnDefinition()
        col3.Width = GridLength(1.4, GridUnitType.Star)
        col4 = ColumnDefinition()
        col4.Width = GridLength(30, GridUnitType.Pixel)
        col5 = ColumnDefinition()
        col5.Width = GridLength(30, GridUnitType.Pixel)
        row_grid.ColumnDefinitions.Add(col1)
        row_grid.ColumnDefinitions.Add(col2)
        row_grid.ColumnDefinitions.Add(col3)
        row_grid.ColumnDefinitions.Add(col4)
        row_grid.ColumnDefinitions.Add(col5)

        txt_param = TextBox()
        txt_param.FontSize = 11
        txt_param.Padding = Thickness(3, 2, 3, 2)
        txt_param.Text = param_name
        txt_param.SetValue(Grid.ColumnProperty, 0)

        lbl_values = TextBlock()
        lbl_values.FontSize = 11
        lbl_values.VerticalAlignment = lbl_values.VerticalAlignment.Center
        lbl_values.Foreground = Brushes.DimGray
        lbl_values.TextTrimming = lbl_values.TextTrimming.CharacterEllipsis
        lbl_values.ToolTip = "Clicca ✎ per modificare i valori ammessi"
        lbl_values.Text = ' - '.join(mutable_values[0]) if mutable_values[0] else '(nessuno)'
        lbl_values.SetValue(Grid.ColumnProperty, 2)

        btn_edit = Button()
        btn_edit.Content = "✎"
        btn_edit.Width = 24
        btn_edit.Height = 24
        btn_edit.FontSize = 12
        btn_edit.ToolTip = "Modifica valori ammessi"
        btn_edit.SetValue(Grid.ColumnProperty, 3)

        btn_del = Button()
        btn_del.Content = "✕"
        btn_del.Width = 24
        btn_del.Height = 24
        btn_del.FontSize = 11
        btn_del.ToolTip = "Rimuovi regola"
        btn_del.SetValue(Grid.ColumnProperty, 4)

        row_grid.Children.Add(txt_param)
        row_grid.Children.Add(lbl_values)
        row_grid.Children.Add(btn_edit)
        row_grid.Children.Add(btn_del)

        self.pnl_validation_rows.Children.Add(row_grid)

        def get_values(mv=mutable_values):
            return list(mv[0])

        row_tuple = (txt_param, get_values, btn_del)
        self._validation_rows.append(row_tuple)

        def on_edit(s, a, lbl=lbl_values, mv=mutable_values):
            new_vals = self._show_values_editor(mv[0])
            if new_vals is not None:
                mv[0] = new_vals
                lbl.Text = ' - '.join(new_vals) if new_vals else '(nessuno)'
        btn_edit.Click += on_edit

        def on_delete(sender, args, grid=row_grid, tup=row_tuple):
            self.pnl_validation_rows.Children.Remove(grid)
            if tup in self._validation_rows:
                self._validation_rows.remove(tup)
        btn_del.Click += on_delete

    def _get_validation_rules(self):
        """Legge tutte le righe validazione e restituisce un dict di regole.

        Returns:
            dict: {param_name: {'allowed_values': [...]}}
        """
        rules = {}
        for txt_param, get_values, btn_del in self._validation_rows:
            param_name = txt_param.Text.strip() if txt_param.Text else ''
            if not param_name:
                continue
            rules[param_name] = {'allowed_values': get_values()}
        return rules

    def _populate_validation_rows(self, rules_dict):
        """Popola le righe validazione da un dict di regole.

        Args:
            rules_dict: dict {param_name: {'allowed_values': [...]}}
        """
        self.pnl_validation_rows.Children.Clear()
        self._validation_rows = []

        for param_name, rule in rules_dict.items():
            values = rule.get('allowed_values', [])
            self._add_validation_row(param_name, values)

        if not rules_dict:
            self._add_validation_row()
    
    def _update_file_count(self):
        """Aggiorna l'etichetta conteggio file."""
        count = self.lst_files.Items.Count
        if count == 0:
            self.lbl_file_count.Text = "Nessun file selezionato"
        elif count == 1:
            self.lbl_file_count.Text = "1 file selezionato"
        else:
            self.lbl_file_count.Text = "{} file selezionati".format(count)
    
    def _check_json_setup(self, folder_path):
        """Controlla se esiste il file JSON di setup nella cartella selezionata."""
        self._json_params = []
        self._json_loaded = False
        self._modifying_setup = False
        self._json_validation_rules = {}
        self._json_validation_loaded = False
        self._modifying_validation = False
        self._json_discipline_rules = []
        self._modifying_discipline = False
        self.chk_modify_setup.IsChecked = False
        self.chk_modify_discipline.IsChecked = False
        self.chk_modify_validation.IsChecked = False
        
        json_path = os.path.join(folder_path, JSON_SETUP_FILENAME)
        
        disc_from_json = False  # True se le regole discipline sono nel JSON
        
        if os.path.isfile(json_path):
            try:
                with io.open(json_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                # Parametri custom
                params = data.get('custom_parameters', [])
                if params and isinstance(params, list):
                    self._json_params = [str(p) for p in params if p]
                    self._json_loaded = True
                    
                    self.lst_json_params.Items.Clear()
                    for p in self._json_params:
                        self.lst_json_params.Items.Add(p)
                    
                    self.pnl_json_setup.Visibility = Visibility.Visible
                    self.pnl_params_section.Visibility = Visibility.Collapsed
                
                # Regole di validazione salvate
                val_rules = data.get('validation_rules', {})
                if val_rules and isinstance(val_rules, dict):
                    self._json_validation_rules = val_rules
                    self._json_validation_loaded = True

                    self.lst_validation_loaded.Items.Clear()
                    for p_name, p_rule in val_rules.items():
                        vals = p_rule.get('allowed_values', [])
                        self.lst_validation_loaded.Items.Add(
                            "{} - {} valori".format(p_name, len(vals)))

                    self.pnl_validation_json.Visibility = Visibility.Visible
                    self.chk_validation.IsChecked = False
                    self.chk_validation.Visibility = Visibility.Collapsed
                    self.pnl_validation_edit.Visibility = Visibility.Collapsed
                else:
                    self.pnl_validation_json.Visibility = Visibility.Collapsed
                    self.chk_validation.Visibility = Visibility.Visible
                
                # Regole discipline salvate
                disc_rules = data.get('discipline_rules', [])
                if disc_rules and isinstance(disc_rules, list):
                    self._json_discipline_rules = [(r['code'], r['desc']) for r in disc_rules
                                                    if isinstance(r, dict) and r.get('code') and r.get('desc')]
                    if self._json_discipline_rules:
                        disc_from_json = True
                
            except Exception:
                pass
        else:
            # JSON non trovato
            self.pnl_json_setup.Visibility = Visibility.Collapsed
            self.pnl_validation_json.Visibility = Visibility.Collapsed
            self.chk_validation.Visibility = Visibility.Visible
            # Mostra la sezione parametri aggiuntivi
            self.pnl_params_section.Visibility = Visibility.Visible
            if not self._param_rows:
                self._add_param_row()
        
        # Gestione sezione Discipline (sempre visibile dopo selezione cartella)
        self.pnl_discipline_section.Visibility = Visibility.Visible
        
        if disc_from_json:
            # Mostra pannello loaded (read-only) con le regole dal JSON
            self.lst_discipline_loaded.Items.Clear()
            for code, desc in self._json_discipline_rules:
                self.lst_discipline_loaded.Items.Add("{} → {}".format(code, desc))
            self.pnl_discipline_loaded.Visibility = Visibility.Visible
            self.pnl_discipline_edit.Visibility = Visibility.Collapsed
            self.chk_modify_discipline.IsChecked = False
        else:
            # Nessuna regola nel JSON: mostra direttamente la tabella editabile
            self.pnl_discipline_loaded.Visibility = Visibility.Collapsed
            self.pnl_discipline_edit.Visibility = Visibility.Visible
            # Aggiungi una riga vuota se non ce ne sono gia'
            if not self._discipline_rows:
                self._add_discipline_row()
    
    # =========================================================================
    # EVENT HANDLERS
    # =========================================================================
    
    def OnSelectFiles(self, sender, args):
        """Apre il dialogo di selezione file multipli."""
        dlg = OpenFileDialog()
        dlg.Title = "Seleziona i file Revit da analizzare"
        dlg.Filter = "File Revit (*.rvt)|*.rvt"
        dlg.Multiselect = True

        if dlg.ShowDialog() == DialogResult.OK:
            for f in dlg.FileNames:
                already_in = False
                for i in range(self.lst_files.Items.Count):
                    if str(self.lst_files.Items[i]) == f:
                        already_in = True
                        break
                if not already_in:
                    self.lst_files.Items.Add(f)
            self._update_file_count()
    
    def OnRemoveFile(self, sender, args):
        if self.lst_files.SelectedIndex >= 0:
            self.lst_files.Items.RemoveAt(self.lst_files.SelectedIndex)
            self._update_file_count()
    
    def OnClearFiles(self, sender, args):
        self.lst_files.Items.Clear()
        self._update_file_count()
    
    def OnBrowseFolder(self, sender, args):
        """Apre il dialogo di selezione cartella e controlla il JSON di setup."""
        selected = self._show_folder_dialog()
        if selected:
            self.txt_output_folder.Text = selected
            self._check_json_setup(selected)

    @staticmethod
    def _show_folder_dialog():
        """Mostra un dialogo moderno per la selezione cartella (compatibile Revit 2022-2025).
        Usa OpenFileDialog con validazione disabilitata per mostrare il dialogo moderno
        anche su .NET Framework 4.8 (Revit 2022-2024), dove FolderBrowserDialog
        mostra solo il dialogo ad albero vecchio stile."""
        try:
            from System.Windows.Forms import OpenFileDialog as OFD, DialogResult as DR
            dlg = OFD()
            dlg.Title = "Seleziona la cartella radice per i CSV"
            dlg.FileName = "Seleziona Cartella"
            dlg.Filter = "Cartella|*.cartella-placeholder"
            dlg.CheckFileExists = False
            dlg.CheckPathExists = True
            dlg.ValidateNames = False
            if dlg.ShowDialog() == DR.OK:
                return os.path.dirname(dlg.FileName)
            return None
        except:
            pass
        # Fallback: FolderBrowserDialog (dialogo ad albero)
        try:
            from System.Windows.Forms import FolderBrowserDialog as FBD, DialogResult as DR
            dlg = FBD()
            dlg.Description = "Seleziona la cartella radice per i CSV"
            dlg.ShowNewFolderButton = True
            if dlg.ShowDialog() == DR.OK:
                return dlg.SelectedPath
            return None
        except:
            return None
    
    def OnDisciplineAdd(self, sender, args):
        """Aggiunge una nuova riga alla tabella discipline."""
        self._add_discipline_row()
    
    def OnDisciplineRemove(self, sender, args):
        """Rimuove l'ultima riga dalla tabella discipline (minimo 1 riga)."""
        if len(self._discipline_rows) > 1:
            self.pnl_discipline_rows.Children.RemoveAt(
                self.pnl_discipline_rows.Children.Count - 1)
            self._discipline_rows.pop()
    
    def OnModifyDisciplineChecked(self, sender, args):
        """Mostra la tabella editabile pre-popolata con le regole JSON."""
        self._modifying_discipline = True
        self._populate_discipline_rows(self._json_discipline_rules)
        self.pnl_discipline_edit.Visibility = Visibility.Visible
    
    def OnModifyDisciplineUnchecked(self, sender, args):
        """Nasconde la tabella editabile e torna alle regole JSON."""
        self._modifying_discipline = False
        self.pnl_discipline_edit.Visibility = Visibility.Collapsed
    
    def OnModifySetupChecked(self, sender, args):
        """Mostra la sezione parametri tabellare pre-popolata con i params JSON."""
        self._modifying_setup = True
        self._populate_param_rows(self._json_params)
        self.pnl_params_section.Visibility = Visibility.Visible
    
    def OnModifySetupUnchecked(self, sender, args):
        """Nasconde la sezione parametri e torna ai params JSON."""
        self._modifying_setup = False
        self.pnl_params_section.Visibility = Visibility.Collapsed
    
    def OnParamAdd(self, sender, args):
        """Aggiunge una nuova riga alla tabella parametri."""
        self._add_param_row()
    
    def OnValidationChecked(self, sender, args):
        self.pnl_validation_edit.Visibility = Visibility.Visible
        if not self._validation_rows:
            self._add_validation_row()
    
    def OnValidationUnchecked(self, sender, args):
        self.pnl_validation_edit.Visibility = Visibility.Collapsed
    
    def OnValidationAdd(self, sender, args):
        """Aggiunge una nuova riga alla tabella validazione."""
        self._add_validation_row()
    
    def OnModifyValidationChecked(self, sender, args):
        """Mostra la tabella editabile pre-popolata con le regole JSON."""
        self._modifying_validation = True
        self._populate_validation_rows(self._json_validation_rules)
        self.chk_validation.IsChecked = True
        self.pnl_validation_edit.Visibility = Visibility.Visible
    
    def OnModifyValidationUnchecked(self, sender, args):
        """Nasconde la tabella editabile e torna alle regole JSON."""
        self._modifying_validation = False
        self.chk_validation.IsChecked = False
        self.pnl_validation_edit.Visibility = Visibility.Collapsed
    
    def OnRun(self, sender, args):
        """Valida e avvia l'elaborazione."""
        if self.lst_files.Items.Count == 0:
            from System.Windows import MessageBox
            MessageBox.Show("Seleziona almeno un file Revit.", "Validazione")
            return
        
        folder = self.txt_output_folder.Text.strip()
        if not folder:
            from System.Windows import MessageBox
            MessageBox.Show("Seleziona una cartella di destinazione.", "Validazione")
            return
        
        if not os.path.isdir(folder):
            from System.Windows import MessageBox
            MessageBox.Show("La cartella selezionata non esiste.\n{}".format(folder), "Validazione")
            return
        
        # Validazione discipline
        # Se le regole sono nel JSON e l'utente non sta modificando, usa quelle
        if self._json_discipline_rules and not getattr(self, '_modifying_discipline', False):
            self.discipline_rules = list(self._json_discipline_rules)
        else:
            # Leggi dalla tabella editabile (obbligatorio almeno una riga)
            disc = self._get_discipline_rules()
            if not disc:
                from System.Windows import MessageBox
                MessageBox.Show("Compila almeno una riga nella tabella Classificazione Disciplina.", "Validazione")
                return
            self.discipline_rules = disc
        
        # Regole di validazione custom
        if self._json_validation_loaded and not self._modifying_validation:
            self.validation_rules = dict(self._json_validation_rules)
        elif self.chk_validation.IsChecked:
            self.validation_rules = self._get_validation_rules()
        else:
            self.validation_rules = {}
        
        # Raccogli i risultati
        self.selected_files = [str(self.lst_files.Items[i]) for i in range(self.lst_files.Items.Count)]
        self.output_folder = folder
        
        # Determina i parametri aggiuntivi
        if self._json_loaded and not self._modifying_setup:
            self.custom_params = list(self._json_params)
        else:
            self.custom_params = self._get_custom_params()
        
        self.result = True
        self.Close()
    
    def OnCancel(self, sender, args):
        self.result = False
        self.Close()


def _resolve_discipline(filename, discipline_rules):
    """Cerca il DisciplineCode nel nome file e restituisce la DisciplineDesc.
    
    Matching case-sensitive.
    
    Args:
        filename: Nome del file (es. 'PRJ-A-Model.rvt')
        discipline_rules: Lista di tuple (code, desc)
    
    Returns:
        str: DisciplineDesc se trovato, 'Undefined' altrimenti
    """
    for code, desc in discipline_rules:
        if code in filename:
            return desc
    return 'Undefined'


def _save_json_setup(folder_path, custom_params, validation_rules=None, discipline_rules=None):
    """Salva il file JSON di setup nella cartella specificata."""
    json_path = os.path.join(folder_path, JSON_SETUP_FILENAME)
    data = {
        'custom_parameters': custom_params,
        'validation_rules': validation_rules or {},
        'discipline_rules': [{'code': c, 'desc': d} for c, d in (discipline_rules or [])],
        'last_modified': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    try:
        with io.open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        OUTPUT.print_md("   \U0001f4be Setup salvato in **{}**".format(JSON_SETUP_FILENAME))
    except Exception as e:
        OUTPUT.print_md("   \u26a0\ufe0f Impossibile salvare il setup JSON: {}".format(str(e)))


# ==============================================================================
# COLOR LEGEND
# ==============================================================================

# Palette fissa di 20 colori moderni e contrastanti per Power BI
_COLOR_PALETTE = [
    '#E63946',  # Vivid Red
    '#377EB8',  # Steel Blue
    '#2A9D8F',  # Teal
    '#984EA3',  # Purple
    '#FF7F00',  # Orange
    '#1B9E77',  # Emerald
    '#E7298A',  # Hot Pink
    '#66A61E',  # Olive Green
    '#4CC9F0',  # Sky Blue
    '#F4A261',  # Sandy Orange
    '#7570B3',  # Lavender
    '#D95F02',  # Burnt Orange
    '#3A86FF',  # Bright Blue
    '#8338EC',  # Violet
    '#06D6A0',  # Mint
    '#FB5607',  # Deep Orange
    '#FF006E',  # Magenta
    '#38B000',  # Vivid Green
    '#0077B6',  # Ocean Blue
    '#A65628',  # Brown
]


def _golden_angle_color(index):
    """Genera un colore HEX usando il golden angle per indici >= 20."""
    import math
    hue = (index * 137.508) % 360
    saturation = 0.70
    lightness = 0.52
    c = (1.0 - abs(2.0 * lightness - 1.0)) * saturation
    x = c * (1.0 - abs((hue / 60.0) % 2.0 - 1.0))
    m = lightness - c / 2.0
    if hue < 60:
        r, g, b = c, x, 0.0
    elif hue < 120:
        r, g, b = x, c, 0.0
    elif hue < 180:
        r, g, b = 0.0, c, x
    elif hue < 240:
        r, g, b = 0.0, x, c
    elif hue < 300:
        r, g, b = x, 0.0, c
    else:
        r, g, b = c, 0.0, x
    return '#{:02X}{:02X}{:02X}'.format(
        int((r + m) * 255),
        int((g + m) * 255),
        int((b + m) * 255)
    )


def _get_color_for_index(index):
    """Restituisce il colore per l'indice dato: palette fissa o golden angle."""
    if index < len(_COLOR_PALETTE):
        return _COLOR_PALETTE[index]
    return _golden_angle_color(index)


def _update_color_legend_csv(output_folder, selected_files):
    """Aggiorna ColorLegend_FileName.csv nella cartella radice.

    - Preserva i colori dei file già presenti
    - Aggiunge nuovi colori per i nuovi file
    - Rimuove le righe per file non più selezionati
    """
    csv_path = os.path.join(output_folder, 'ColorLegend_FileName.csv')

    # Leggi i colori esistenti
    existing_colors = {}  # {filename: hex_color}
    if os.path.isfile(csv_path):
        try:
            with io.open(csv_path, 'r', encoding=CSV_ENCODING) as f:
                reader = csv.DictReader(f, delimiter=CSV_DELIMITER)
                for row in reader:
                    fname = row.get('FileName', '').strip()
                    color = row.get('Color', '').strip()
                    if fname and color:
                        existing_colors[fname] = color
        except Exception as e:
            LOGGER.warning("Errore lettura ColorLegend_FileName.csv: {}".format(str(e)))

    # Costruisci la nuova lista preservando i colori esistenti
    used_colors = set(existing_colors.values())
    next_index = 0
    rows = []

    for fpath in selected_files:
        fname = os.path.basename(fpath).strip()
        if fname in existing_colors:
            rows.append({'FileName': fname, 'Color': existing_colors[fname]})
        else:
            # Trova il prossimo colore non ancora in uso
            while _get_color_for_index(next_index) in used_colors:
                next_index += 1
            new_color = _get_color_for_index(next_index)
            used_colors.add(new_color)
            next_index += 1
            rows.append({'FileName': fname, 'Color': new_color})

    # Scrivi il CSV
    try:
        with io.open(csv_path, 'w', encoding=CSV_ENCODING, newline='') as f:
            writer = csv.DictWriter(f, fieldnames=['FileName', 'Color'],
                                    delimiter=CSV_DELIMITER, lineterminator='\n')
            writer.writeheader()
            writer.writerows(rows)
        OUTPUT.print_md("   \U0001f3a8 ColorLegend_FileName.csv aggiornato ({} file)".format(len(rows)))
    except Exception as e:
        OUTPUT.print_md("   \u26a0\ufe0f Impossibile scrivere ColorLegend_FileName.csv: {}".format(str(e)))


def show_model_report_form():
    """Mostra la form di configurazione e restituisce i parametri.
    
    Returns:
        tuple: (selected_files, output_folder, custom_params, validation_rules, discipline_rules) o None
    """
    form = ModelReportForm()
    form.ShowDialog()
    
    if form.result:
        return (form.selected_files, form.output_folder, form.custom_params,
                form.validation_rules, form.discipline_rules)
    return None


def _load_validation_csv(csv_path):
    """DEPRECATA - Mantenuta per compatibilità, non più usata."""
    return {}


# ==============================================================================
# COSTANTI REGEX PER NAMING CONVENTION (hardcoded)
# ==============================================================================

import re

# FamilyName: e_ + CODICE_MAIUSCOLO (punto come separatore) + _ + NomeDescrittivo
REGEX_FAMILY_NAME = r'^e_[A-Z]+(\.[A-Z]+)*_[A-Za-z0-9 ]+$'

# Type Mark: CODICE (maiuscolo, punti come separatore) + suffisso opzionale + numero
REGEX_TYPE_MARK = r'^[A-Z]+(\.[A-Z]+)*[a-z]*\.?\d{1,3}$'

# Type Name: {Type Mark} - NomeDescrittivo (regex costruita dinamicamente per ogni tipo)
REGEX_TYPE_NAME_TEMPLATE = r'^{type_mark_escaped} - [A-Za-z0-9 ]+$'


def validate_families_data(families_data):
    """Valida il FamilyName delle famiglie con regex hardcoded.
    
    Args:
        families_data: Lista di dizionari delle famiglie (da extract_families_types_instances)
    
    Returns:
        Lista di dizionari per TAB_DataValidation_Families
    """
    validation_rows = []
    
    for fam in families_data:
        family_key = fam.get('FamilyKey', '')
        family_name = fam.get('FamilyName', '')
        
        if not family_name or family_name == '':
            status = 'EMPTY'
        elif re.match(REGEX_FAMILY_NAME, family_name):
            status = 'VALID'
        else:
            status = 'INVALID'
        
        validation_rows.append({
            'FamilyKey': family_key,
            'FieldName': 'FamilyName',
            'FieldValue': family_name,
            'Status': status
        })
    
    return validation_rows


def validate_types_data(types_data, custom_type_params=None, validation_rules=None):
    """Valida i tipi: regex su Type Mark/Type Name + valori ammessi per parametri custom.
    
    La validazione regex e' sempre eseguita (hardcoded).
    La validazione valori ammessi e' eseguita solo per i parametri custom che hanno regole.
    
    Args:
        types_data: Lista di dizionari dei tipi (da extract_families_types_instances)
        custom_type_params: Lista di nomi parametri custom di tipo (opzionale)
        validation_rules: dict {param_name: {'binding': 'type', 'allowed_values': [...]}} (opzionale)
    
    Returns:
        Lista di dizionari per TAB_DataValidation_Types
    """
    if custom_type_params is None:
        custom_type_params = []
    if validation_rules is None:
        validation_rules = {}
    
    validation_rows = []
    
    # Parametri custom di tipo con regole di validazione
    # Il binding (istanza/tipo) viene rilevato automaticamente a runtime da _classify_custom_params
    type_params_with_rules = [p for p in custom_type_params
                              if p in validation_rules
                              and validation_rules[p].get('allowed_values')]
    
    for type_row in types_data:
        type_key = type_row.get('TypeKey', '')
        type_mark = type_row.get('TypeMark', '')
        type_name = type_row.get('TypeName', '')
        
        # --- Validazione regex Type Mark ---
        if not type_mark or type_mark == '':
            tm_status = 'EMPTY'
        elif re.match(REGEX_TYPE_MARK, type_mark):
            tm_status = 'VALID'
        else:
            tm_status = 'INVALID'
        
        validation_rows.append({
            'TypeKey': type_key,
            'FieldName': 'Type Mark',
            'FieldValue': type_mark,
            'Status': tm_status
        })
        
        # --- Validazione regex Type Name (dinamica, basata su Type Mark) ---
        if not type_name or type_name == '':
            tn_status = 'EMPTY'
        elif not type_mark or type_mark == '':
            # Se Type Mark e' vuoto, TypeName e' automaticamente INVALID
            tn_status = 'INVALID'
        else:
            # Costruisci la regex dinamica escapando i caratteri speciali nel Type Mark
            escaped_mark = re.escape(type_mark)
            type_name_regex = REGEX_TYPE_NAME_TEMPLATE.format(type_mark_escaped=escaped_mark)
            if re.match(type_name_regex, type_name):
                tn_status = 'VALID'
            else:
                tn_status = 'INVALID'
        
        validation_rows.append({
            'TypeKey': type_key,
            'FieldName': 'TypeName',
            'FieldValue': type_name,
            'Status': tn_status
        })
        
        # --- Validazione valori ammessi per parametri custom di tipo ---
        for param_name in type_params_with_rules:
            param_value = type_row.get(param_name, '')
            allowed = set(validation_rules[param_name].get('allowed_values', []))
            
            if not param_value or param_value == '':
                status = 'EMPTY'
            elif param_value in allowed:
                status = 'VALID'
            else:
                status = 'INVALID'
            
            validation_rows.append({
                'TypeKey': type_key,
                'FieldName': param_name,
                'FieldValue': param_value,
                'Status': status
            })
    
    return validation_rows


def validate_instances_data(instances_data, custom_instance_params=None, validation_rules=None):
    """Valida i parametri custom di istanza contro i valori ammessi definiti dall'utente.
    
    Args:
        instances_data: Lista di dizionari delle istanze (da extract_families_types_instances)
        custom_instance_params: Lista di nomi parametri custom di istanza (opzionale)
        validation_rules: dict {param_name: {'binding': 'instance', 'allowed_values': [...]}} (opzionale)
    
    Returns:
        Lista di dizionari per TAB_DataValidation_Instances
    """
    if custom_instance_params is None:
        custom_instance_params = []
    if validation_rules is None:
        validation_rules = {}
    
    validation_rows = []
    
    # Parametri custom di istanza con regole di validazione
    # Il binding (istanza/tipo) viene rilevato automaticamente a runtime da _classify_custom_params
    instance_params_with_rules = [p for p in custom_instance_params
                                  if p in validation_rules
                                  and validation_rules[p].get('allowed_values')]
    
    if not instance_params_with_rules:
        return validation_rows
    
    for elem in instances_data:
        element_key = elem.get('ElementKey', '')
        
        for param_name in instance_params_with_rules:
            param_value = elem.get(param_name, '')
            allowed = set(validation_rules[param_name].get('allowed_values', []))
            
            if not param_value or param_value == '':
                status = 'EMPTY'
            elif param_value in allowed:
                status = 'VALID'
            else:
                status = 'INVALID'
            
            validation_rows.append({
                'ElementKey': element_key,
                'FieldName': param_name,
                'FieldValue': param_value,
                'Status': status
            })
    
    return validation_rows


# ==============================================================================
# FUNZIONI HELPER PER FORMATTAZIONE NUMERI
# ==============================================================================

def _format_decimal(value, decimals):
    """Formatta un numero con il numero esatto di decimali specificato.
    
    Args:
        value: Valore numerico da formattare
        decimals: Numero di decimali (1 o 2)
    
    Returns:
        Stringa formattata o stringa vuota se il valore non è valido
    """
    try:
        if value is None or value == "":
            return ""
        if decimals == 1:
            return "{:.1f}".format(float(value))
        elif decimals == 2:
            return "{:.2f}".format(float(value))
        else:
            return str(round(float(value), decimals))
    except:
        return ""


# ==============================================================================
# FUNZIONI DI ESTRAZIONE
# ==============================================================================

def _format_coordinates(x, y, z):
    """Formatta coordinate in metri con formato X=...; Y=...; Z=..."""
    # Converti da piedi a metri (1 piede = 0.3048 metri)
    x_m = round(x * 0.3048, 4)
    y_m = round(y * 0.3048, 4)
    z_m = round(z * 0.3048, 4)
    return "X={}; Y={}; Z={}".format(x_m, y_m, z_m)


def _get_base_points_info(doc):
    """Estrae informazioni su Project Base Point e Survey Point."""
    pbp_coords = ""
    sp_coords = ""
    angle_true_north = ""
    
    try:
        import math
        
        # PROJECT BASE POINT (triangolo) - OST_ProjectBasePoint
        pbp_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).ToElements()
        for bp in pbp_collector:
            try:
                # Prova con la proprietà Position (BasePoint class)
                if hasattr(bp, 'Position'):
                    pos = bp.Position
                    if pos:
                        pbp_coords = _format_coordinates(pos.X, pos.Y, pos.Z)
                
                # Se non funziona, prova con i parametri
                if not pbp_coords:
                    ew_param = bp.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM)
                    ns_param = bp.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM)
                    elev_param = bp.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
                    
                    if ew_param and ns_param and elev_param:
                        pbp_coords = _format_coordinates(
                            ew_param.AsDouble(), 
                            ns_param.AsDouble(), 
                            elev_param.AsDouble()
                        )
                
                # Angolo True North
                angle_param = bp.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM)
                if angle_param and angle_param.HasValue:
                    angle_rad = angle_param.AsDouble()
                    angle_deg = round(math.degrees(angle_rad), 4)
                    angle_true_north = str(angle_deg)
                break
            except Exception as e:
                LOGGER.warning("Errore PBP: {}".format(str(e)))
        
        # SURVEY POINT (cerchio con X) - OST_SharedBasePoint
        sp_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_SharedBasePoint).ToElements()
        for sp in sp_collector:
            try:
                # Prova con la proprietà Position (BasePoint class)
                if hasattr(sp, 'Position'):
                    pos = sp.Position
                    if pos:
                        sp_coords = _format_coordinates(pos.X, pos.Y, pos.Z)
                
                # Se non funziona, prova con i parametri
                if not sp_coords:
                    ew_param = sp.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM)
                    ns_param = sp.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM)
                    elev_param = sp.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
                    
                    if ew_param and ns_param and elev_param:
                        sp_coords = _format_coordinates(
                            ew_param.AsDouble(), 
                            ns_param.AsDouble(), 
                            elev_param.AsDouble()
                        )
                break
            except Exception as e:
                LOGGER.warning("Errore SP: {}".format(str(e)))
                
    except Exception as e:
        LOGGER.warning("Errore estrazione base points: {}".format(str(e)))
    
    # Ritorna nell'ordine corretto: PBP, SP, Angle
    return pbp_coords, sp_coords, angle_true_north


def extract_file_info(processor):
    """Estrae informazioni sul file (TAB_Files)."""
    doc = processor.doc
    
    # Dimensione file (già calcolata nel costruttore)
    file_size_mb = processor.file_size_mb
    
    # Formatta con virgola come separatore decimale (formato italiano)
    # Rimuovi zeri finali non necessari (17.0 -> "17", 17.6 -> "17,6")
    file_size_formatted = "{:.2f}".format(file_size_mb).rstrip('0').rstrip('.').replace('.', ',')
    
    # Starting View - Metodo corretto usando StartingViewSettings
    starting_page_name = ""
    try:
        starting_view_settings = StartingViewSettings.GetStartingViewSettings(doc)
        if starting_view_settings:
            starting_view_id = starting_view_settings.ViewId
            if starting_view_id and starting_view_id != ElementId.InvalidElementId:
                starting_view = doc.GetElement(starting_view_id)
                if starting_view:
                    if isinstance(starting_view, ViewSheet):
                        starting_page_name = "{} - {}".format(
                            starting_view.SheetNumber,
                            starting_view.Name
                        )
                    else:
                        starting_page_name = starting_view.Name
    except Exception as e:
        # Fallback: se StartingViewSettings non è disponibile (versioni vecchie)
        starting_page_name = ""
    
    # Base Points (ritorna: pbp_coords, sp_coords, angle)
    pbp_coords, sp_coords, angle_true_north = _get_base_points_info(doc)
    
    return {
        'FileName': processor.file_name,
        'FileDiscipline': '',  # Compilata dopo, nel main, con _resolve_discipline
        'FilePath': processor.file_path,
        'FileSize_MB': file_size_formatted,
        'OpeningTime': processor.opening_time,
        'StartingPage_Name': starting_page_name,
        'ExtractionDate': EXTRACTION_DATE,
        'IsWorkshared': "YES" if processor.is_workshared else "NO",
        'ProjectBasePoint': pbp_coords,
        'SurveyPoint': sp_coords,
        'AngleTrueNorth': angle_true_north
    }


def compute_file_summary(file_info, links_data, views_data, warnings_data,
                         sheets_data, templates_data, levels_data, grids_data,
                         filters_data, rooms_data, spaces_data, areas_data,
                         tags_data, families_data, types_data, instances_data,
                         purgeable_data, model_in_place_count=0):
    """Calcola le colonne riassuntive per TAB_Files basate sui dati estratti dalle altre tabelle.
    
    Args:
        file_info: Dizionario con le info base del file (da extract_file_info)
        links_data: Lista dati TAB_Links
        views_data: Lista dati TAB_Views
        warnings_data: Lista dati TAB_Warnings
        sheets_data: Lista dati TAB_Sheets
        templates_data: Lista dati TAB_ViewTemplates
        levels_data: Lista dati TAB_Levels
        grids_data: Lista dati TAB_Grids
        filters_data: Lista dati TAB_Filters
        rooms_data: Lista dati TAB_Rooms
        spaces_data: Lista dati TAB_Spaces
        areas_data: Lista dati TAB_Areas
        tags_data: Lista dati TAB_Tags
        families_data: Lista dati TAB_Families
        types_data: Lista dati TAB_Types
        instances_data: Lista dati TAB_Instances
        purgeable_data: Lista dati TAB_PurgeableElements
    
    Returns:
        Dizionario con le colonne riassuntive da aggiungere a file_info
    """
    summary = {}
    
    # --- Conteggi totali ---
    summary['Elements'] = len(instances_data)
    summary['Families'] = len(families_data)
    summary['Types'] = len(types_data)
    summary['PurgeableElements'] = len(purgeable_data)
    summary['Warnings'] = len(warnings_data)
    
    # Viste non template (solo viste normali)
    non_template_views = [v for v in views_data if v.get('IsTemplate') == 'NO']
    summary['Views_HasTemplate(N)'] = len(non_template_views)
    
    # Viste posizionate su un foglio (ReferencingSheet non vuoto)
    views_on_sheet = [v for v in non_template_views if v.get('ReferencingSheet', '')]
    summary['Views_OnSheet'] = len(views_on_sheet)
    
    # Totale tavole
    summary['Sheets'] = len(sheets_data)
    
    # --- Links non pinnati ---
    summary['Links_RVT_Pinned(N)'] = len([l for l in links_data 
                                       if l.get('LinkType') == 'RVT' and l.get('IsPinned') == 'NO'])
    summary['Links_DWG_Pinned(N)'] = len([l for l in links_data 
                                       if l.get('LinkType') == 'DWG' and l.get('IsPinned') == 'NO'])
    
    # --- Livelli ---
    summary['Levels_Pinned(N)'] = len([lv for lv in levels_data if lv.get('IsPinned') == 'NO'])
    
    total_levels = len(levels_data)
    monitored_levels = len([lv for lv in levels_data if lv.get('IsMonitor') == 'YES'])
    summary['Levels_Monitored(N)'] = total_levels - monitored_levels
    
    # --- Griglie ---
    total_grids = len(grids_data)
    if total_grids > 0:
        summary['Grids_Pinned(N)'] = len([g for g in grids_data if g.get('IsPinned') == 'NO'])
    else:
        summary['Grids_Pinned(N)'] = 0
    
    monitored_grids = len([g for g in grids_data if g.get('IsMonitored') == 'YES'])
    summary['Grids_Monitored(N)'] = total_grids - monitored_grids
    
    # --- Viste senza ViewChapter1 o ViewChapter2 (solo non-template) ---
    views_no_chapter = len([v for v in non_template_views 
                            if not v.get('ViewChapter1', '') or not v.get('ViewChapter2', '')])
    summary['Views_VC(N)'] = views_no_chapter
    
    # --- Viste non posizionate su tavola (solo non-template) ---
    views_not_on_sheet = len([v for v in non_template_views 
                              if not v.get('ReferencingSheet', '')])
    summary['Views_OnSheet(N)'] = views_not_on_sheet
    
    # --- Sheets senza ViewChapter1 ---
    sheets_no_chapter1 = len([s for s in sheets_data if not s.get('ViewChapter1', '')])
    summary['Sheets_VC1(N)'] = sheets_no_chapter1
    
    # --- Tag senza Host ---
    summary['Tags_Host(N)'] = len([t for t in tags_data if t.get('HasHost') == 'NO'])
    
    # --- Rooms/Spaces/Areas non confinate, non posizionate o ridondanti ---
    # Condizione OR: IsPlaced="NO" OR IsEnclosed="NO" OR IsRedundant="YES"
    problematic_count = 0
    for room in rooms_data:
        if (room.get('IsPlaced') == 'NO' or 
            room.get('IsEnclosed') == 'NO' or 
            room.get('IsRedundant') == 'YES'):
            problematic_count += 1
    for space in spaces_data:
        if (space.get('IsPlaced') == 'NO' or 
            space.get('IsEnclosed') == 'NO' or 
            space.get('IsRedundant') == 'YES'):
            problematic_count += 1
    # Areas hanno solo IsPlaced (no IsEnclosed/IsRedundant)
    for area in areas_data:
        if area.get('IsPlaced') == 'NO':
            problematic_count += 1
    summary['RoomsAreasSpaces_ConfinedPlaced(N)'] = problematic_count
    
    # --- View Templates non conformi alla naming convention ---
    # Convention: il nome deve iniziare con "e_XXX_" dove XXX e' uno dei codici validi
    valid_vt_prefixes = ['e_3DV_', 'e_CON_', 'e_EXP_', 'e_LNK_', 'e_PRI_', 'e_WIP_', 'e_SCH_']
    non_compliant_vt = 0
    for vt in templates_data:
        name = vt.get('TemplateName', '')
        if not any(name.startswith(prefix) for prefix in valid_vt_prefixes):
            non_compliant_vt += 1
    summary['ViewTemplates_CompliantName(N)'] = non_compliant_vt
    
    # --- Filtri non conformi alla naming convention ---
    # Convention: il nome deve iniziare con "e_XXX_" dove XXX e' uno dei codici validi
    valid_f_prefixes = ['e_ARC_', 'e_GEN_', 'e_ELE_', 'e_MEC_', 'e_MEP_', 'e_STR_']
    non_compliant_f = 0
    for f in filters_data:
        name = f.get('FilterName', '')
        if not any(name.startswith(prefix) for prefix in valid_f_prefixes):
            non_compliant_f += 1
    summary['Filters_CompliantName(N)'] = non_compliant_f
    
    # --- Starting View check ---
    # Controlla se il nome della Starting Page contiene "Starting View"
    starting_page = file_info.get('StartingPage_Name', '')
    summary['StartingView_Correct'] = 'YES' if 'Starting View' in starting_page else 'NO'
    # --- Link RVT con SharedSite "Not Shared" ---
    summary['Links_RVT_LinkedBySharedCoordinates(N)'] = len([l for l in links_data 
                                        if l.get('LinkType') == 'RVT' and l.get('SharedSite') == 'Not Shared'])
    
    # --- Link DWG View-Specific (Workset_(i) che inizia per "View") ---
    summary['Links_DWG_IsViewSpecific'] = len([l for l in links_data 
                                           if l.get('LinkType') == 'DWG' and 
                                           l.get('Workset_(i)', '').startswith('View')])
    # --- Warnings "identical instances in the same place" ---
    # Logica B: per ogni warning group (stesso WarningID) con N elementi coinvolti,
    # i duplicati EFFETTIVI sono N-1 (si esclude l'"originale").
    # Somma (N-1) su tutti i gruppi = totale istanze in eccesso da eliminare.
    duplicate_warning_text = "There are identical instances in the same place"
    dup_warnings = [w for w in warnings_data 
                    if duplicate_warning_text in w.get('WarningDescription', '')]
    dup_by_id = defaultdict(int)
    for w in dup_warnings:
        dup_by_id[w.get('WarningID')] += 1
    summary['Warnings_DuplicateInstance'] = sum(max(0, count - 1) for count in dup_by_id.values())
    
    # --- Famiglie Model In Place ---
    summary['ModelInPlace'] = model_in_place_count

    # --- Elementi nascosti nelle viste su sheet (HiddenElements > 0) ---
    hidden_on_sheets = 0
    for v in views_data:
        he = v.get('HiddenElements', 'ND')
        if he != 'ND':
            try:
                hidden_on_sheets += int(he)
            except:
                pass
    summary['HiddenElements_OnSheets'] = hidden_on_sheets
    summary['HideInView'] = hidden_on_sheets

    return summary


def compute_health_checks(file_info):
    """Calcola i 20 Health Check per un file e ritorna una lista di righe per TAB_HealthChecks.
    
    Args:
        file_info: Dizionario completo del file (da extract_file_info + compute_file_summary)
    
    Returns:
        Lista di dizionari, uno per check, con chiavi:
        CheckID, CheckDescription, FileName, CheckPassed, Value, Score
    """
    checks = []
    file_name = file_info.get('FileName', '')
    
    def _int(key):
        """Legge un intero da file_info in modo sicuro."""
        try:
            return int(file_info.get(key, 0) or 0)
        except:
            return 0
    
    def _add_binary(check_id, description, value, passed_condition, max_score):
        """Aggiunge un check binario: YES se superato, NO altrimenti."""
        checks.append({
            'CheckID': check_id,
            'CheckDescription': description,
            'FileName': file_name,
            'CheckPassed': 'YES' if passed_condition else 'NO',
            'Value': str(value),
            'Score': max_score if passed_condition else 0,
            'MaxScore': max_score
        })
    
    # HC01 — Unpinned Levels (max 5pt)
    v = _int('Levels_Pinned(N)')
    _add_binary('HC01', 'Unpinned Levels', v, v == 0, 5)
    
    # HC02 — Unmonitored Levels (max 4pt)
    v = _int('Levels_Monitored(N)')
    _add_binary('HC02', 'Monitored Levels', v, v == 0, 4)
    
    # HC03 — Unpinned DWG Links (max 4pt)
    v = _int('Links_DWG_Pinned(N)')
    _add_binary('HC03', 'Unpinned DWG Links', v, v == 0, 4)
    
    # HC04 — View Specific DWG Links (max 4pt)
    v = _int('Links_DWG_IsViewSpecific')
    _add_binary('HC04', 'View specific DWG Links', v, v == 0, 4)
    
    # HC05 — Unpinned RVT Links (max 5pt)
    v = _int('Links_RVT_Pinned(N)')
    _add_binary('HC05', 'Unpinned RVT Links', v, v == 0, 5)
    
    # HC06 — Unpinned Grids (max 5pt)
    v = _int('Grids_Pinned(N)')
    _add_binary('HC06', 'Unpinned Grids', v, v == 0, 5)
    
    # HC07 — Unmonitored Grids (max 3pt)
    v = _int('Grids_Monitored(N)')
    _add_binary('HC07', 'Monitor Grid', v, v == 0, 3)
    
    # HC08 — RVT Link without shared site linking method (max 5pt)
    v = _int('Links_RVT_LinkedBySharedCoordinates(N)')
    _add_binary('HC08', 'RVT Link without shared site linking method', v, v == 0, 5)
    
    # HC09 — Views without naming chapter (max 3pt)
    v = _int('Views_VC(N)')
    _add_binary('HC09', 'Views without "u_OTH_ViewChapter1"', v, v == 0, 3)
    
    # HC10 — Views not on sheet (max 5pt, 5 fasce)
    v = _int('Views_OnSheet(N)')
    if v <= 50:
        passed, score = 'Excellent', 5
    elif v <= 80:
        passed, score = 'Good', 4
    elif v <= 100:
        passed, score = 'Sufficient', 3
    elif v <= 150:
        passed, score = 'Poor', 1
    else:
        passed, score = 'Bad', 0
    checks.append({'CheckID': 'HC10', 'CheckDescription': 'Views not on sheet',
                   'FileName': file_name, 'CheckPassed': passed, 'Value': str(v), 'Score': score, 'MaxScore': 5})
    
    # HC11 — Sheets without naming chapter (max 3pt)
    v = _int('Sheets_VC1(N)')
    _add_binary('HC11', 'Sheets without u_OTH_ViewChapter1', v, v == 0, 3)
    
    # HC12 — View Template with unproper name (max 4pt)
    v = _int('ViewTemplates_CompliantName(N)')
    _add_binary('HC12', 'View Template with unproper name', v, v == 0, 4)
    
    # HC13 — View Filter with unproper name (max 4pt)
    v = _int('Filters_CompliantName(N)')
    _add_binary('HC13', 'View Filter with unproper Name', v, v == 0, 4)
    
    # HC14 — Total purgeable elements (max 20pt, 5 fasce)
    v = _int('PurgeableElements')
    if v <= 50:
        passed, score = 'Excellent', 20
    elif v <= 100:
        passed, score = 'Good', 15
    elif v <= 150:
        passed, score = 'Sufficient', 10
    elif v <= 300:
        passed, score = 'Poor', 5
    else:
        passed, score = 'Bad', 0
    checks.append({'CheckID': 'HC14', 'CheckDescription': 'Total purgeable elements',
                   'FileName': file_name, 'CheckPassed': passed, 'Value': str(v), 'Score': score, 'MaxScore': 20})
    
    # HC15 — Tags without Host (max 5pt)
    v = _int('Tags_Host(N)')
    _add_binary('HC15', 'Tags without Host', v, v == 0, 5)
    
    # HC16 — Model in place (max 3pt)
    v = _int('ModelInPlace')
    _add_binary('HC16', 'Model in place', v, v == 0, 3)
    
    # HC17 — Duplicate instance (max 5pt)
    v = _int('Warnings_DuplicateInstance')
    _add_binary('HC17', 'Duplicate instance', v, v == 0, 5)
    
    # HC18 — Rooms/Areas/Spaces unplaced or unbounded (max 4pt)
    v = _int('RoomsAreasSpaces_ConfinedPlaced(N)')
    _add_binary('HC18', 'Rooms/Areas/Spaces unplaced or unbounded', v, v == 0, 4)
    
    # HC19 — File open on starting view (max 4pt) — Value = nome della starting view
    starting_view = file_info.get('StartingPage_Name', '')
    has_sv = file_info.get('StartingView_Correct', 'NO') == 'YES'
    checks.append({
        'CheckID': 'HC19',
        'CheckDescription': 'File open on starting view',
        'FileName': file_name,
        'CheckPassed': 'YES' if has_sv else 'NO',
        'Value': starting_view,
        'Score': 4 if has_sv else 0,
        'MaxScore': 4
    })

    # HC20 — Element Hidden on Sheets (max 5pt)
    v = _int('HiddenElements_OnSheets')
    _add_binary('HC20', 'Element Hidden on Sheets', v, v == 0, 5)

    return checks


def _append_to_snapshot_summary(base_folder, new_rows, current_fieldnames):
    """Gestisce il file TAB_Snapshot_Summary.csv con append incrementale per data.
    
    Logica:
      1. Legge il file esistente (se presente) ed estrae tutti i record
      2. Filtra (rimuove) i record con ExtractionDate della data odierna
      3. Unisce i fieldnames: existing + eventuali nuovi (celle vuote per i vecchi record)
      4. Riscrive il file con: vecchi record filtrati + nuovi record della run corrente
    
    Args:
        base_folder:        Cartella radice scelta dall'utente
        new_rows:           Lista di dict dei record TAB_Files della run corrente
        current_fieldnames: Lista ordinata dei nomi colonna attuali (da TABLE_HEADERS)
    """
    snapshot_path = os.path.join(base_folder, "TAB_Snapshot_Summary.csv")
    today_str = datetime.now().strftime("%Y-%m-%d")
    
    existing_rows = []
    existing_fieldnames = []
    removed_count = 0
    
    # --- Leggi il file esistente ---
    if os.path.exists(snapshot_path):
        try:
            with io.open(snapshot_path, 'r', encoding=CSV_ENCODING) as f:
                reader = csv.DictReader(f, delimiter=CSV_DELIMITER)
                if reader.fieldnames:
                    existing_fieldnames = list(reader.fieldnames)
                for row in reader:
                    extraction_date = row.get('ExtractionDate', '')
                    if extraction_date[:10] == today_str:
                        removed_count += 1  # Record di oggi: da scartare
                    else:
                        existing_rows.append(dict(row))
            if removed_count > 0:
                OUTPUT.print_md("   🗑️ Rimossi {} record di oggi dal Snapshot (verranno riscritti)".format(removed_count))
        except Exception as e:
            OUTPUT.print_md("   ⚠️ Errore lettura Snapshot esistente: {}. Verrà ricreato.".format(str(e)))
            existing_rows = []
            existing_fieldnames = []
    
    # --- Merge fieldnames: mantieni ordine esistente + aggiungi nuove colonne in coda ---
    merged_fieldnames = list(existing_fieldnames)
    for fn in current_fieldnames:
        if fn not in merged_fieldnames:
            merged_fieldnames.append(fn)
    # Se il file non esisteva, usa direttamente i fieldnames correnti
    if not merged_fieldnames:
        merged_fieldnames = list(current_fieldnames)
    
    # --- Scrivi il file finale ---
    try:
        with io.open(snapshot_path, 'w', encoding=CSV_ENCODING, newline='') as f:
            writer = csv.DictWriter(f, fieldnames=merged_fieldnames, delimiter=CSV_DELIMITER,
                                   extrasaction='ignore', lineterminator='\n')
            writer.writeheader()
            # Vecchi record: celle vuote per colonne nuove (DictWriter gestisce extrasaction='ignore'
            # ma per le colonne mancanti nei vecchi record il default di DictWriterè '' se restvalue="")
            for row in existing_rows:
                # Assicura che le nuove colonne abbiano stringa vuota nei vecchi record
                for fn in merged_fieldnames:
                    if fn not in row:
                        row[fn] = ""
                writer.writerow(row)
            # Nuovi record della run corrente
            for row in new_rows:
                writer.writerow(row)
        
        total = len(existing_rows) + len(new_rows)
        OUTPUT.print_md("   ✅ **TAB_Snapshot_Summary** aggiornato: {} record totali ({} precedenti + {} nuovi)".format(
            total, len(existing_rows), len(new_rows)))
    
    except IOError as e:
        if "being used by another process" in str(e) or "cannot access the file" in str(e):
            OUTPUT.print_md("   ⚠️ **TAB_Snapshot_Summary** saltato (file aperto in un altro programma)")
        else:
            OUTPUT.print_md("   ❌ Errore scrittura TAB_Snapshot_Summary: {}".format(str(e)))
    except Exception as e:
        OUTPUT.print_md("   ❌ Errore scrittura TAB_Snapshot_Summary: {}".format(str(e)))


def _prepare_output_folders(base_folder):
    """Gestisce la struttura di storicizzazione delle cartelle di output.
    
    Logica:
      1. Crea (se non esiste) base_folder/CurrentData e base_folder/_Old
      2. Legge TAB_Files.csv in CurrentData e ne estrae ExtractionDate
      3. Se ExtractionDate.date != oggi:
           - Ricava YYYYMMDD dalla ExtractionDate del vecchio file
           - Se _Old/YYYYMMDD NON esiste: copia l'intera CurrentData lì dentro
           - Se _Old/YYYYMMDD esiste già: salta la copia (sovrascrittura diretta)
      4. Restituisce il path di CurrentData (dove scrivere i nuovi CSV)
    
    Args:
        base_folder: Cartella principale scelta dall'utente
    
    Returns:
        Stringa con il path assoluto di CurrentData
    """
    import shutil
    
    current_folder = os.path.join(base_folder, "CurrentData")
    old_folder     = os.path.join(base_folder, "_Old")
    
    # Crea le cartelle se non esistono
    if not os.path.exists(current_folder):
        os.makedirs(current_folder)
        OUTPUT.print_md("📁 Creata cartella **CurrentData**")
    if not os.path.exists(old_folder):
        os.makedirs(old_folder)
        OUTPUT.print_md("📁 Creata cartella **_Old**")
    
    # Controlla se esiste il file TAB_Files.csv in CurrentData
    tab_files_path = os.path.join(current_folder, "TAB_Files.csv")
    
    if not os.path.exists(tab_files_path):
        # Prima esecuzione o run precedente interrotta — procedi direttamente
        OUTPUT.print_md("ℹ️ Nessun dato precedente trovato in **CurrentData**. Prima esecuzione.")
        return current_folder
    
    # Leggi ExtractionDate dalla prima riga dati del vecchio TAB_Files.csv
    old_extraction_date_str = ""
    try:
        with io.open(tab_files_path, 'r', encoding=CSV_ENCODING) as f:
            reader = csv.DictReader(f, delimiter=CSV_DELIMITER)
            for row in reader:
                old_extraction_date_str = row.get('ExtractionDate', '')
                break  # basta la prima riga
    except Exception as e:
        OUTPUT.print_md("⚠️ Impossibile leggere **TAB_Files.csv** esistente: {}. Procedo a sovrascrivere.".format(str(e)))
        return current_folder
    
    if not old_extraction_date_str:
        OUTPUT.print_md("⚠️ **ExtractionDate** non trovata in TAB_Files.csv. Procedo a sovrascrivere.")
        return current_folder
    
    # Estrai solo la parte data (YYYY-MM-DD) dalla stringa "YYYY-MM-DD HH:MM:SS"
    try:
        old_date_part = old_extraction_date_str.strip()[:10]  # "YYYY-MM-DD"
        old_date = datetime.strptime(old_date_part, "%Y-%m-%d").date()
    except Exception as e:
        OUTPUT.print_md("⚠️ Formato **ExtractionDate** non riconosciuto: '{}'. Procedo a sovrascrivere.".format(old_extraction_date_str))
        return current_folder
    
    today = datetime.now().date()
    
    if old_date == today:
        # Stessa giornata: sovrascrittura diretta
        OUTPUT.print_md("ℹ️ Dati di **CurrentData** già aggiornati ad oggi ({}). Sovrascrittura diretta.".format(today))
        return current_folder
    
    # Data diversa: archivia in _Old/YYYYMMDD
    archive_folder_name = old_date.strftime("%Y%m%d")
    archive_path = os.path.join(old_folder, archive_folder_name)
    
    if os.path.exists(archive_path):
        # Cartella già esistente in _Old per quella data: salta l'archiviazione
        OUTPUT.print_md("⚠️ Cartella **_Old/{}** già esistente. Archiviazione saltata. Sovrascrittura diretta in CurrentData.".format(archive_folder_name))
        return current_folder
    
    # Copia l'intera CurrentData in _Old/YYYYMMDD
    try:
        shutil.copytree(current_folder, archive_path)
        OUTPUT.print_md("✅ Dati precedenti archiviati in **_Old/{}**".format(archive_folder_name))
    except Exception as e:
        OUTPUT.print_md("⚠️ Errore durante l'archiviazione in _Old: {}. Procedo a sovrascrivere CurrentData.".format(str(e)))
    
    return current_folder


def _get_file_last_modified_date(file_path):
    """Ottiene la data di ultima modifica del file dal file system.
    Ritorna la data in formato YYYY-MM-DD o 'Non disponibile' se non accessibile."""
    if not file_path:
        return "Non disponibile"
    
    try:
        # Ottieni il timestamp di ultima modifica
        mtime = os.path.getmtime(file_path)
        # Converti in data formattata
        return datetime.fromtimestamp(mtime).strftime("%Y-%m-%d")
    except:
        return "Non disponibile"


def extract_links(processor):
    """Estrae informazioni sui link (TAB_Links)."""
    doc = processor.doc
    links = []
    
    # RVT Links
    rvt_link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    
    for link_instance in rvt_link_instances:
        try:
            link_type = doc.GetElement(link_instance.GetTypeId())
            
            # Salta se il tipo non esiste o non è valido
            if not link_type:
                continue
            
            # Ottieni il path del link
            link_path = ""
            link_file_name = "Unknown RVT Link"
            try:
                external_ref = link_type.GetExternalFileReference()
                if external_ref:
                    link_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(
                        external_ref.GetAbsolutePath()
                    )
                    # Estrai il nome del file dal path
                    if link_path:
                        link_file_name = os.path.basename(link_path)
            except:
                link_path = ""
            
            # Se non abbiamo il path, prova a usare il nome del tipo come fallback
            if link_file_name == "Unknown RVT Link":
                try:
                    if hasattr(link_type, 'Name') and link_type.Name:
                        link_file_name = link_type.Name
                except:
                    pass
            
            # Workset di ISTANZA (solo se workshared)
            instance_workset = "NONE"
            if processor.is_workshared:
                try:
                    ws_param = link_instance.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                    if ws_param and ws_param.HasValue:
                        ws_id = ws_param.AsInteger()
                        workset = doc.GetWorksetTable().GetWorkset(DB.WorksetId(ws_id))
                        if workset:
                            instance_workset = workset.Name
                except:
                    instance_workset = "NONE"
            
            # Workset di TIPO (solo se workshared)
            type_workset = "NONE"
            if processor.is_workshared:
                try:
                    ws_param = link_type.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                    if ws_param and ws_param.HasValue:
                        ws_id = ws_param.AsInteger()
                        workset = doc.GetWorksetTable().GetWorkset(DB.WorksetId(ws_id))
                        if workset:
                            type_workset = workset.Name
                except:
                    type_workset = "NONE"
            
            # FileName del link (dal documento linkato se caricato)
            link_file_name_full = ""
            link_pbp = ""
            link_sp = ""
            linking_method = ""
            
            try:
                linked_doc = link_instance.GetLinkDocument()
                if linked_doc:
                    # Usa Title che include il nome file con estensione
                    link_file_name_full = linked_doc.Title
                    if link_file_name_full and not link_file_name_full.endswith('.rvt'):
                        link_file_name_full += '.rvt'
                    
                    # Estrai Base Points del link (ritorna: pbp, sp, angle)
                    link_pbp_coords, link_sp_coords, _ = _get_base_points_info(linked_doc)
                    link_pbp = link_pbp_coords
                    link_sp = link_sp_coords
            except:
                link_file_name_full = ""
            
            # SharedSite - estratto dal Name dell'istanza del link
            # Il formato del Name può essere:
            #   "NomeFile.rvt : <Not Shared>"
            #   "NomeFile.rvt : location : LocationName"
            #   "NomeFile.rvt : LocationName"
            shared_site = "N/A"
            try:
                instance_name = ""
                try:
                    instance_name = link_instance.Name
                except:
                    try:
                        instance_name = Element.Name.GetValue(link_instance)
                    except:
                        instance_name = ""
                
                if instance_name and " : " in instance_name:
                    # Splitta per " : " e rimuovi la prima parte (nome file)
                    parts = instance_name.split(" : ")
                    # Prendi l'ultima parte (il nome del site)
                    site_part = parts[-1].strip()
                    
                    if "<Not Shared>" in site_part:
                        shared_site = "Not Shared"
                    elif site_part:
                        # Se la penultima parte è "location", il site name è l'ultima parte
                        # Se site_part stesso inizia con "location", rimuovi il prefisso
                        if site_part.lower().startswith("location"):
                            site_part = site_part[len("location"):].strip()
                        shared_site = site_part if site_part else "N/A"
            except:
                shared_site = "N/A"
            
            # Data ultima modifica del file linkato
            last_saved_date = _get_file_last_modified_date(link_path)
            
            links.append({
                'LinkKey': "{} : {}".format(processor.file_name, link_instance.Id.IntegerValue),
                'LinkID': link_instance.Id.IntegerValue,
                'FileName': processor.file_name,
                'LinkName': link_file_name,
                'LinkDiscipline': '',  # Compilata dopo, nel main, con _resolve_discipline
                'LinkPath': link_path,
                'LinkFileName': link_file_name_full,
                'LastSavedDate': last_saved_date,
                'SharedSite': shared_site,
                'ProjectBasePoint': link_pbp,
                'SurveyPoint': link_sp,
                'Workset_(i)': instance_workset,
                'Workset_(t)': type_workset,
                'IsPinned': "YES" if link_instance.Pinned else "NO",
                'LinkType': "RVT"
            })
        except Exception as e:
            LOGGER.warning("Errore estrazione RVT link: {}".format(str(e)))
    
    # CAD Links
    cad_instances = FilteredElementCollector(doc).OfClass(ImportInstance).ToElements()
    
    for cad_instance in cad_instances:
        try:
            if cad_instance.IsLinked:  # Solo link, non import
                cad_type = doc.GetElement(cad_instance.GetTypeId())
                
                # Salta se il tipo non esiste
                if not cad_type:
                    continue
                
                # Path del CAD
                link_path = ""
                cad_file_name = "Unknown CAD"
                try:
                    if hasattr(cad_type, 'GetExternalFileReference'):
                        external_ref = cad_type.GetExternalFileReference()
                        if external_ref:
                            link_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(
                                external_ref.GetAbsolutePath()
                            )
                            # Estrai il nome del file dal path
                            if link_path:
                                cad_file_name = os.path.basename(link_path)
                except:
                    link_path = ""
                
                # Se non abbiamo il path, usa il nome del tipo come fallback
                if cad_file_name == "Unknown CAD":
                    try:
                        if hasattr(cad_type, 'Name') and cad_type.Name:
                            cad_file_name = cad_type.Name
                    except:
                        pass
                
                # Determina il tipo di file
                link_type_str = "CAD"
                try:
                    type_name = cad_file_name.lower()
                    if ".ifc" in type_name:
                        link_type_str = "IFC"
                    elif ".dwg" in type_name:
                        link_type_str = "DWG"
                    elif ".dxf" in type_name:
                        link_type_str = "DXF"
                    elif ".dgn" in type_name:
                        link_type_str = "DGN"
                    elif ".sat" in type_name or ".skp" in type_name:
                        link_type_str = "3D"
                except:
                    link_type_str = "CAD"
                
                # Workset di ISTANZA (solo se workshared)
                instance_workset = "NONE"
                if processor.is_workshared:
                    try:
                        ws_param = cad_instance.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                        if ws_param and ws_param.HasValue:
                            ws_id = ws_param.AsInteger()
                            workset = doc.GetWorksetTable().GetWorkset(DB.WorksetId(ws_id))
                            if workset:
                                instance_workset = workset.Name
                    except:
                        instance_workset = "NONE"
                
                # Workset di TIPO (solo se workshared)
                type_workset = "NONE"
                if processor.is_workshared:
                    try:
                        ws_param = cad_type.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                        if ws_param and ws_param.HasValue:
                            ws_id = ws_param.AsInteger()
                            workset = doc.GetWorksetTable().GetWorkset(DB.WorksetId(ws_id))
                            if workset:
                                type_workset = workset.Name
                    except:
                        type_workset = "NONE"
                
                # Data ultima modifica del file linkato
                last_saved_date = _get_file_last_modified_date(link_path)
                
                links.append({
                    'LinkKey': "{} : {}".format(processor.file_name, cad_instance.Id.IntegerValue),
                    'LinkID': cad_instance.Id.IntegerValue,
                    'FileName': processor.file_name,
                    'LinkName': cad_file_name,
                    'LinkDiscipline': '',  # Compilata dopo, nel main, con _resolve_discipline
                    'LinkPath': link_path,
                    'LinkFileName': cad_file_name,
                    'LastSavedDate': last_saved_date,
                    'SharedSite': "N/A",
                    'ProjectBasePoint': "",
                    'SurveyPoint': "",
                    'Workset_(i)': instance_workset,
                    'Workset_(t)': type_workset,
                    'IsPinned': "YES" if cad_instance.Pinned else "NO",
                    'LinkType': link_type_str
                })
        except Exception as e:
            LOGGER.warning("Errore estrazione CAD link: {}".format(str(e)))
    
    return links


def extract_views(processor):
    """Estrae informazioni sulle viste (TAB_Views) - ESPANSA."""
    doc = processor.doc
    views = []
    
    # Raccogli tutte le viste
    all_views = FilteredElementCollector(doc).OfClass(View).ToElements()

    # Pre-raccogli elementi model (non tipi) per il conteggio HiddenElements
    # Eseguito una sola volta, riutilizzato per tutte le viste su sheet
    model_element_ids = None  # lazy: caricato solo se serve

    for view in all_views:
        try:
            # IsTemplate
            if view.IsTemplate:
                is_template = "YES"
            else:
                is_template = "NO"
            
            # Tipo vista
            view_type = view.ViewType.ToString()
            
            # View Template associato
            template_id = view.ViewTemplateId
            if template_id and template_id != ElementId.InvalidElementId:
                template = doc.GetElement(template_id)
                template_name = template.Name if template else "None"
                template_id_str = template_id.IntegerValue
            else:
                template_name = "None"
                template_id_str = "None"
            
            # 1. HasScopeBox
            has_scope_box = "False"
            try:
                scope_box_param = view.get_Parameter(DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
                if scope_box_param and scope_box_param.HasValue:
                    scope_box_id = scope_box_param.AsElementId()
                    if scope_box_id and scope_box_id != ElementId.InvalidElementId:
                        has_scope_box = "True"
            except:
                has_scope_box = "False"
            
            # 2. IsDependent
            is_dependent = "False"
            try:
                # Le viste dipendenti hanno un parent view
                if hasattr(view, 'GetPrimaryViewId'):
                    parent_id = view.GetPrimaryViewId()
                    if parent_id and parent_id != ElementId.InvalidElementId:
                        is_dependent = "True"
            except:
                is_dependent = "False"
            
            # 3. Phase Filter
            phase_filter_name = ""
            try:
                phase_filter_param = view.get_Parameter(DB.BuiltInParameter.VIEW_PHASE_FILTER)
                if phase_filter_param and phase_filter_param.HasValue:
                    phase_filter_id = phase_filter_param.AsElementId()
                    if phase_filter_id and phase_filter_id != ElementId.InvalidElementId:
                        phase_filter = doc.GetElement(phase_filter_id)
                        if phase_filter:
                            phase_filter_name = phase_filter.Name
            except:
                phase_filter_name = ""
            
            # 4. Phase
            phase_name = ""
            try:
                phase_param = view.get_Parameter(DB.BuiltInParameter.VIEW_PHASE)
                if phase_param and phase_param.HasValue:
                    phase_id = phase_param.AsElementId()
                    if phase_id and phase_id != ElementId.InvalidElementId:
                        phase = doc.GetElement(phase_id)
                        if phase:
                            phase_name = phase.Name
            except:
                phase_name = ""
            
            # 5. Referencing Sheet (numero del foglio su cui è posizionata la vista)
            referencing_sheet = ""
            try:
                sheet_param = view.get_Parameter(DB.BuiltInParameter.VIEWPORT_SHEET_NUMBER)
                if sheet_param and sheet_param.HasValue:
                    referencing_sheet = sheet_param.AsString() or ""
            except:
                referencing_sheet = ""
            
            # 6. Title on Sheet (titolo modificabile nella viewport)
            title_on_sheet = ""
            try:
                title_param = view.get_Parameter(DB.BuiltInParameter.VIEW_DESCRIPTION)
                if title_param and title_param.HasValue:
                    title_on_sheet = title_param.AsString() or ""
            except:
                title_on_sheet = ""
            
            # 7. ViewChapter1 (parametro di progetto)
            view_chapter_1 = ""
            try:
                chapter1_param = view.LookupParameter("u_OTH_ViewChapter1")
                if chapter1_param and chapter1_param.HasValue:
                    view_chapter_1 = chapter1_param.AsString() or ""
            except:
                view_chapter_1 = ""
            
            # 8. ViewChapter2 (parametro di progetto)
            view_chapter_2 = ""
            try:
                chapter2_param = view.LookupParameter("u_OTH_ViewChapter2")
                if chapter2_param and chapter2_param.HasValue:
                    view_chapter_2 = chapter2_param.AsString() or ""
            except:
                view_chapter_2 = ""

            # 9. HiddenElements - conteggio elementi nascosti manualmente (Hide in View > Elements)
            # Calcolato solo per viste non-template posizionate su sheet (con ReferencingSheet)
            # Per le altre viste il valore è "ND"
            hidden_count = "ND"
            if not view.IsTemplate and referencing_sheet:
                try:
                    # Carica gli elementi model una sola volta (lazy)
                    if model_element_ids is None:
                        model_element_ids = FilteredElementCollector(doc) \
                            .WhereElementIsNotElementType() \
                            .ToElementIds()
                    count = 0
                    for eid in model_element_ids:
                        try:
                            elem = doc.GetElement(eid)
                            if elem is not None and elem.IsHidden(view):
                                count += 1
                        except:
                            pass
                    hidden_count = count
                except:
                    hidden_count = "ND"

            views.append({
                'ViewID': view.Id.IntegerValue,
                'FileName': processor.file_name,
                'ViewKey': "{} : {}".format(processor.file_name, view.Id.IntegerValue),
                'ViewName': view.Name,
                'ViewType': view_type,
                'ViewTemplateID': template_id_str,
                'ViewTemplateName': template_name,
                'IsTemplate': is_template,
                'HasScopeBox': has_scope_box,
                'IsDependent': is_dependent,
                'PhaseFilter': phase_filter_name,
                'Phase': phase_name,
                'ReferencingSheet': referencing_sheet,
                'TitleOnSheet': title_on_sheet,
                'ViewChapter1': view_chapter_1,
                'ViewChapter2': view_chapter_2,
                'HiddenElements': hidden_count
            })
        except Exception as e:
            LOGGER.warning("Errore estrazione vista: {}".format(str(e)))
    
    return views


def _load_warnings_severity_csv():
    """Carica il CSV WarningsSeverity.csv dalla cartella dello script.
    
    Returns:
        dict: {description_normalizzata: severity_score}
              es. {'a area was deleted...': '02_Medium', ...}
    """
    severity_lookup = {}
    
    script_dir = os.path.dirname(__file__)
    csv_path = os.path.join(script_dir, 'WarningsSeverity.csv')
    
    if not os.path.isfile(csv_path):
        LOGGER.warning("WarningsSeverity.csv non trovato in: {}".format(script_dir))
        return severity_lookup
    
    try:
        with io.open(csv_path, 'r', encoding=CSV_ENCODING) as f:
            reader = csv.reader(f, delimiter=CSV_DELIMITER)
            rows = list(reader)
        
        if not rows:
            return severity_lookup
        
        # Salta la prima riga (intestazioni: Description;Severity Score)
        for row in rows[1:]:
            if len(row) >= 2:
                desc = row[0].strip()
                severity = row[1].strip()
                if desc and severity:
                    # Normalizza la descrizione (lowercase, strip) per matching
                    severity_lookup[desc.lower()] = severity
    
    except Exception as e:
        LOGGER.warning("Errore caricamento WarningsSeverity.csv: {}".format(str(e)))
    
    return severity_lookup


def _lookup_warning_severity(description, severity_lookup):
    """Cerca la severity di un warning nel dizionario di lookup.
    
    Strategia:
    1. Match esatto (dopo normalizzazione)
    2. Match approssimato: la chiave CSV piu' lunga contenuta nella descrizione
    
    Returns:
        tuple: (severity_score, match_type)
               match_type: 'EXACT', 'APPROX', o 'NOT_FOUND'
    """
    if not severity_lookup:
        return ('00_Unknown', 'NOT_FOUND')
    
    desc_lower = description.strip().lower()
    
    # 1. Match esatto
    if desc_lower in severity_lookup:
        return (severity_lookup[desc_lower], 'EXACT')
    
    # 2. Match approssimato: cerca la chiave CSV piu' lunga contenuta nella descrizione
    best_match = None
    best_len = 0
    
    for csv_desc, severity in severity_lookup.items():
        if csv_desc in desc_lower and len(csv_desc) > best_len:
            best_match = severity
            best_len = len(csv_desc)
    
    if best_match:
        return (best_match, 'APPROX')
    
    # 3. Match inverso: la descrizione e' contenuta in una chiave CSV
    for csv_desc, severity in severity_lookup.items():
        if desc_lower in csv_desc and len(desc_lower) > best_len:
            best_match = severity
            best_len = len(desc_lower)
    
    if best_match:
        return (best_match, 'APPROX')
    
    return ('00_Unknown', 'NOT_FOUND')


def extract_warnings(processor, severity_lookup=None):
    """Estrae warnings (TAB_Warnings).
    
    Args:
        processor: FileProcessor con il documento aperto
        severity_lookup: Dizionario {desc_lower: severity} dal CSV (opzionale)
    """
    doc = processor.doc
    warnings = []
    
    if severity_lookup is None:
        severity_lookup = {}
    
    try:
        failure_messages = doc.GetWarnings()
        
        warning_counter = 1
        for failure in failure_messages:
            # Usa i primi 8 caratteri del filename (senza estensione) per il WarningID
            file_prefix = os.path.splitext(processor.file_name)[0][:8]
            warning_id = "{}_W{:04d}".format(file_prefix, warning_counter)
            
            # Descrizione
            description = failure.GetDescriptionText()
            
            # Severity dal CSV
            severity, match_type = _lookup_warning_severity(description, severity_lookup)
            
            # Failure Definition ID (per categorizzazione)
            failure_def_id = failure.GetFailureDefinitionId()
            failure_guid = str(failure_def_id.Guid) if failure_def_id else ""
            
            # Elementi coinvolti
            element_ids = failure.GetFailingElements()
            additional_ids = failure.GetAdditionalElements()
            
            all_element_ids = list(element_ids) + list(additional_ids)
            
            if all_element_ids:
                for elem_id in all_element_ids:
                    warnings.append({
                        'WarningKey': "{} : {}".format(processor.file_name, warning_id),
                        'WarningID': warning_id,
                        'FileName': processor.file_name,
                        'WarningDescription': description,
                        'WarningSeverity': severity,
                        'WarningDescValidation': match_type,
                        'WarningFailureGUID': failure_guid,
                        'WarningType': _categorize_warning(description),
                        'ElementID': elem_id.IntegerValue
                    })
            else:
                # Warning senza elementi specifici
                warnings.append({
                    'WarningKey': "{} : {}".format(processor.file_name, warning_id),
                    'WarningID': warning_id,
                    'FileName': processor.file_name,
                    'WarningDescription': description,
                    'WarningSeverity': severity,
                    'WarningDescValidation': match_type,
                    'WarningFailureGUID': failure_guid,
                    'WarningType': _categorize_warning(description),
                    'ElementID': "N/A"
                })
            
            warning_counter += 1
            
    except Exception as e:
        LOGGER.warning("Errore estrazione warnings: {}".format(str(e)))
    
    return warnings


def _categorize_warning(description):
    """Categorizza un warning basandosi sulla descrizione."""
    description_lower = description.lower()
    
    if "room" in description_lower:
        return "Room"
    elif "overlap" in description_lower or "duplicate" in description_lower:
        return "Overlap/Duplicate"
    elif "constraint" in description_lower:
        return "Constraint"
    elif "join" in description_lower:
        return "Join"
    elif "wall" in description_lower:
        return "Wall"
    elif "stair" in description_lower or "railing" in description_lower:
        return "Stairs/Railings"
    elif "family" in description_lower:
        return "Family"
    elif "area" in description_lower:
        return "Area"
    elif "level" in description_lower or "grid" in description_lower:
        return "Level/Grid"
    elif "dimension" in description_lower:
        return "Dimension"
    else:
        return "Other"


def get_parameter_value(element, param_name, doc):
    """
    Estrae il valore di un parametro da un elemento.
    Gestisce parametri di istanza e di tipo.
    Restituisce il valore come stringa o stringa vuota se non trovato.
    """
    value = ""
    
    # Prima prova sul parametro di ISTANZA
    try:
        param = element.LookupParameter(param_name)
        if param and param.HasValue:
            value = _extract_param_value(param)
            if value:
                return value
    except:
        pass
    
    # Se non trovato, prova sul TIPO
    try:
        type_id = element.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            elem_type = doc.GetElement(type_id)
            if elem_type:
                type_param = elem_type.LookupParameter(param_name)
                if type_param and type_param.HasValue:
                    value = _extract_param_value(type_param)
    except:
        pass
    
    return value


def _extract_param_value(param):
    """
    Estrae il valore di un parametro in base al suo StorageType.
    Gestisce: String, Integer, Double, Boolean (ElementId per Yes/No).
    """
    try:
        storage_type = param.StorageType
        
        # String
        if storage_type == DB.StorageType.String:
            return param.AsString() or ""
        
        # Integer (include Yes/No che sono 0/1)
        elif storage_type == DB.StorageType.Integer:
            int_value = param.AsInteger()
            # Controlla se è un parametro Yes/No
            try:
                if param.Definition.ParameterType == DB.ParameterType.YesNo:
                    return "Yes" if int_value == 1 else "No"
            except:
                pass
            return str(int_value)
        
        # Double (numeri con virgola)
        elif storage_type == DB.StorageType.Double:
            double_value = param.AsDouble()
            # Formatta con massimo 4 decimali, rimuovendo zeri finali
            return "{:.4f}".format(double_value).rstrip('0').rstrip('.')
        
        # ElementId (per parametri che referenziano altri elementi)
        elif storage_type == DB.StorageType.ElementId:
            elem_id = param.AsElementId()
            if elem_id and elem_id.IntegerValue != -1:
                return str(elem_id.IntegerValue)
            return ""
        else:
            return ""
    except:
        return ""


def _classify_custom_params(doc, custom_params):
    """Classifica i parametri custom come istanza o tipo, verificando che siano parametri di progetto.
    
    Usa doc.ParameterBindings per verificare che i parametri richiesti dall'utente
    siano effettivamente parametri di progetto (non parametri di famiglia).
    
    Args:
        doc: Documento Revit aperto
        custom_params: Lista di nomi di parametri custom richiesti dall'utente
    
    Returns:
        tuple: (instance_params, type_params, invalid_params)
            - instance_params: Lista parametri con InstanceBinding
            - type_params: Lista parametri con TypeBinding
            - invalid_params: Lista parametri non trovati tra i parametri di progetto
    """
    if not custom_params:
        return [], [], []
    
    # Costruisci dizionario dei parametri di progetto dal ParameterBindings
    project_params = {}  # {param_name: 'instance' | 'type'}
    
    try:
        binding_map = doc.ParameterBindings
        iterator = binding_map.ForwardIterator()
        iterator.Reset()
        while iterator.MoveNext():
            try:
                definition = iterator.Key
                binding = iterator.Current
                param_name = definition.Name
                if isinstance(binding, InstanceBinding):
                    project_params[param_name] = 'instance'
                elif isinstance(binding, TypeBinding):
                    project_params[param_name] = 'type'
            except:
                continue
    except Exception as e:
        LOGGER.warning("Errore lettura ParameterBindings: {}".format(str(e)))
    
    instance_params = []
    type_params = []
    invalid_params = []
    
    for p in custom_params:
        if p in project_params:
            if project_params[p] == 'instance':
                instance_params.append(p)
            else:
                type_params.append(p)
        else:
            invalid_params.append(p)
    
    return instance_params, type_params, invalid_params


def extract_worksets(processor):
    """Estrae workset user-defined (TAB_Worksets_UserDefined) - CORRETTO v2."""
    doc = processor.doc
    worksets = []
    
    if not processor.is_workshared:
        return worksets
    
    try:
        # Metodo corretto: usa FilteredWorksetCollector
        from Autodesk.Revit.DB import FilteredWorksetCollector
        
        workset_collector = FilteredWorksetCollector(doc)
        all_worksets = workset_collector.OfKind(WorksetKind.UserWorkset)
        
        workset_table = doc.GetWorksetTable()
        
        for ws in all_worksets:
            # Owner del workset
            owner_name = ""
            try:
                owner_info = workset_table.GetWorksetInfo(ws.Id)
                if owner_info:
                    owner_name = owner_info.Owner
            except:
                owner_name = ""
            
            # IsOpen
            is_open = "NO"
            try:
                if ws.IsOpen:
                    is_open = "YES"
            except:
                is_open = "NO"
            
            worksets.append({
                'WorksetKey': "{} : {}".format(processor.file_name, ws.Id.IntegerValue),
                'WorksetID': ws.Id.IntegerValue,
                'FileName': processor.file_name,
                'WorksetName': ws.Name,
                'IsVisibleInAllViews': "YES" if ws.IsVisibleByDefault else "NO",
                'Owner': owner_name,
                'IsOpen': is_open
            })
            
    except Exception as e:
        LOGGER.warning("Errore estrazione worksets: {}".format(str(e)))
    
    return worksets


def extract_sheets(processor):
    """Estrae tavole (TAB_Sheets) - ESPANSA."""
    doc = processor.doc
    sheets = []
    
    all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    
    for sheet in all_sheets:
        try:
            # Revisione corrente
            revision_ids = sheet.GetAllRevisionIds()
            current_revision_num = ""
            current_revision_name = ""
            
            if revision_ids:
                # Prendi l'ultima revisione
                last_rev_id = revision_ids[revision_ids.Count - 1]
                revision = doc.GetElement(last_rev_id)
                if revision:
                    current_revision_num = revision.SequenceNumber
                    current_revision_name = revision.Description if hasattr(revision, 'Description') else ""
            
            # ViewChapter1 (parametro di progetto)
            view_chapter_1 = ""
            try:
                chapter1_param = sheet.LookupParameter("u_OTH_ViewChapter1")
                if chapter1_param and chapter1_param.HasValue:
                    view_chapter_1 = chapter1_param.AsString() or ""
            except:
                view_chapter_1 = ""
            
            # ViewChapter2 (parametro di progetto)
            view_chapter_2 = ""
            try:
                chapter2_param = sheet.LookupParameter("u_OTH_ViewChapter2")
                if chapter2_param and chapter2_param.HasValue:
                    view_chapter_2 = chapter2_param.AsString() or ""
            except:
                view_chapter_2 = ""
            
            sheets.append({
                'SheetKey': "{} : {}".format(processor.file_name, sheet.Id.IntegerValue),
                'SheetID': sheet.Id.IntegerValue,
                'FileName': processor.file_name,
                'SheetNumber': sheet.SheetNumber,
                'SheetName': sheet.Name,
                'CurrentRevisionNumber': current_revision_num,
                'CurrentRevisionName': current_revision_name,
                'ViewChapter1': view_chapter_1,
                'ViewChapter2': view_chapter_2
            })
        except Exception as e:
            LOGGER.warning("Errore estrazione sheet: {}".format(str(e)))
    
    return sheets


def extract_view_templates(processor):
    """Estrae view templates (TAB_ViewTemplates)."""
    doc = processor.doc
    templates = []
    
    # Raccogli tutti i view template
    all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
    
    # Set di template usati
    used_template_ids = set()
    
    for view in all_views:
        if not view.IsTemplate:
            template_id = view.ViewTemplateId
            if template_id and template_id != ElementId.InvalidElementId:
                used_template_ids.add(template_id.IntegerValue)
    
    # Estrai i template
    for view in all_views:
        try:
            if view.IsTemplate:
                is_used = "YES" if view.Id.IntegerValue in used_template_ids else "NO"
                
                template_data = {}
                template_data['ViewTemplateKey'] = "{} : {}".format(processor.file_name, view.Id.IntegerValue)
                template_data['ViewTemplateID'] = view.Id.IntegerValue
                template_data['FileName'] = processor.file_name
                template_data['TemplateName'] = view.Name
                template_data['ViewType'] = view.ViewType.ToString()
                template_data['IsUsed'] = is_used
                templates.append(template_data)
        except Exception as e:
            LOGGER.warning("Errore estrazione view template: {}".format(str(e)))
    
    return templates


def extract_scope_boxes(processor):
    """Estrae scope boxes (TAB_ScopeBoxes)."""
    doc = processor.doc
    scope_boxes = []
    
    try:
        # Raccogli tutti gli scope box (sono nella categoria OST_VolumeOfInterest)
        sb_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).WhereElementIsNotElementType().ToElements()
        
        for sb in sb_collector:
            try:
                # ScopeBox ID
                sb_id = sb.Id.IntegerValue
                
                # ScopeBox Name
                sb_name = sb.Name if sb.Name else ""
                
                # Workset
                workset_name = "NONE"
                if processor.is_workshared:
                    try:
                        ws_param = sb.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                        if ws_param and ws_param.HasValue:
                            ws_id = ws_param.AsInteger()
                            workset = doc.GetWorksetTable().GetWorkset(DB.WorksetId(ws_id))
                            if workset:
                                workset_name = workset.Name
                    except:
                        workset_name = "NONE"
                
                # IsPinned
                is_pinned = "YES" if sb.Pinned else "NO"
                
                scope_boxes.append({
                    'ScopeBoxKey': "{} : {}".format(processor.file_name, sb_id),
                    'ScopeBoxID': sb_id,
                    'FileName': processor.file_name,
                    'ScopeBoxName': sb_name,
                    'Workset': workset_name,
                    'IsPinned': is_pinned
                })
                
            except Exception as e:
                LOGGER.warning("Errore estrazione scope box: {}".format(str(e)))
        
        OUTPUT.print_md("      ✓ {} scope box estratti".format(len(scope_boxes)))
        
    except Exception as e:
        OUTPUT.print_md("      ⚠️ Errore estrazione scope boxes: {}".format(str(e)))
        LOGGER.error("Errore extract_scope_boxes: {}".format(str(e)))
    
    return scope_boxes


def extract_grids(processor):
    """Estrae griglie (TAB_Grids)."""
    doc = processor.doc
    grids = []
    
    try:
        # Raccogli tutte le griglie
        grid_collector = FilteredElementCollector(doc).OfClass(Grid).ToElements()
        
        for grid in grid_collector:
            try:
                # GridID
                grid_id = grid.Id.IntegerValue
                
                # GridName
                grid_name = grid.Name if grid.Name else ""
                
                # GridType (nome del tipo)
                grid_type_name = ""
                try:
                    type_id = grid.GetTypeId()
                    if type_id and type_id != ElementId.InvalidElementId:
                        grid_type = doc.GetElement(type_id)
                        if grid_type:
                            try:
                                grid_type_name = grid_type.Name
                            except:
                                pass
                            if not grid_type_name:
                                try:
                                    name_param = grid_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                                    if name_param and name_param.HasValue:
                                        grid_type_name = name_param.AsString() or ""
                                except:
                                    pass
                except:
                    pass
                
                # IsPinned
                is_pinned = "YES" if grid.Pinned else "NO"
                
                # IsMonitored, MonitorFileName, MonitorGrid
                is_monitored = "NO"
                monitor_file_name = ""
                monitor_grid_name = ""
                
                try:
                    if grid.IsMonitoringLinkElement():
                        is_monitored = "YES"
                        
                        # Ottieni gli ID dei link monitorati
                        monitored_link_ids = grid.GetMonitoredLinkElementIds()
                        
                        if monitored_link_ids and monitored_link_ids.Count > 0:
                            for monitored_link_id in monitored_link_ids:
                                try:
                                    link_instance = doc.GetElement(monitored_link_id)
                                    if link_instance and isinstance(link_instance, RevitLinkInstance):
                                        link_doc = link_instance.GetLinkDocument()
                                        if link_doc:
                                            monitor_file_name = link_doc.Title
                                            if monitor_file_name and not monitor_file_name.endswith('.rvt'):
                                                monitor_file_name += '.rvt'
                                            
                                            # Cerca la griglia monitorata nel link (stessa posizione/nome)
                                            try:
                                                link_grids = FilteredElementCollector(link_doc).OfClass(Grid).ToElements()
                                                for link_grid in link_grids:
                                                    # Cerca griglia con stesso nome
                                                    if link_grid.Name == grid_name:
                                                        monitor_grid_name = link_grid.Name
                                                        break
                                            except:
                                                pass
                                        break
                                except:
                                    pass
                except:
                    pass
                
                # ScopeBox associato
                scope_box_name = ""
                try:
                    scope_box_param = grid.get_Parameter(DB.BuiltInParameter.DATUM_VOLUME_OF_INTEREST)
                    if scope_box_param and scope_box_param.HasValue:
                        scope_box_id = scope_box_param.AsElementId()
                        if scope_box_id and scope_box_id != ElementId.InvalidElementId:
                            scope_box = doc.GetElement(scope_box_id)
                            if scope_box:
                                scope_box_name = scope_box.Name
                except:
                    scope_box_name = ""
                
                # Workset (campo vuoto se non workshared)
                workset_name = ""
                if processor.is_workshared:
                    try:
                        ws_param = grid.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                        if ws_param and ws_param.HasValue:
                            ws_id = ws_param.AsInteger()
                            workset = doc.GetWorksetTable().GetWorkset(DB.WorksetId(ws_id))
                            if workset:
                                workset_name = workset.Name
                    except:
                        workset_name = ""
                
                grids.append({
                    'GridKey': "{} : {}".format(processor.file_name, grid_id),
                    'GridID': grid_id,
                    'FileName': processor.file_name,
                    'GridName': grid_name,
                    'GridType': grid_type_name,
                    'IsPinned': is_pinned,
                    'IsMonitored': is_monitored,
                    'MonitorFileName': monitor_file_name,
                    'MonitorGrid': monitor_grid_name,
                    'ScopeBox': scope_box_name,
                    'Workset': workset_name
                })
                
            except Exception as e:
                LOGGER.warning("Errore estrazione griglia: {}".format(str(e)))
        
        OUTPUT.print_md("      ✓ {} griglie estratte".format(len(grids)))
        
    except Exception as e:
        OUTPUT.print_md("      ⚠️ Errore estrazione griglie: {}".format(str(e)))
        LOGGER.error("Errore extract_grids: {}".format(str(e)))
    
    return grids


def extract_materials(processor):
    """Estrae materiali (TAB_Materials) - ESPANSA."""
    doc = processor.doc
    materials = []
    
    # Raccogli tutti i materiali
    all_materials = FilteredElementCollector(doc).OfClass(Material).ToElements()
    
    # Trova quali materiali sono usati
    used_material_ids = set()
    try:
        all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        for elem in all_elements:
            try:
                mat_ids = elem.GetMaterialIds(False)
                for mid in mat_ids:
                    used_material_ids.add(mid.IntegerValue)
            except:
                pass
    except:
        pass
    
    # Estrai i materiali
    for mat in all_materials:
        try:
            is_used = "YES" if mat.Id.IntegerValue in used_material_ids else "NO"
            
            # Material Class
            mat_class = ""
            try:
                if hasattr(mat, 'MaterialClass'):
                    mat_class = mat.MaterialClass
            except:
                mat_class = ""
            
            # Material Category
            mat_category = ""
            try:
                if hasattr(mat, 'MaterialCategory'):
                    mat_category = mat.MaterialCategory
            except:
                mat_category = ""
            
            # 1. Material Description
            mat_description = ""
            try:
                desc_param = mat.get_Parameter(DB.BuiltInParameter.ALL_MODEL_DESCRIPTION)
                if desc_param and desc_param.HasValue:
                    mat_description = desc_param.AsString() or ""
            except:
                mat_description = ""
            
            # 2. Material Comments
            mat_comments = ""
            try:
                comments_param = mat.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                if comments_param and comments_param.HasValue:
                    mat_comments = comments_param.AsString() or ""
            except:
                mat_comments = ""
            
            # 3. Material Keywords
            mat_keywords = ""
            try:
                keywords_param = mat.get_Parameter(DB.BuiltInParameter.MATERIAL_PARAM_KEYWORDS)
                if keywords_param and keywords_param.HasValue:
                    mat_keywords = keywords_param.AsString() or ""
            except:
                mat_keywords = ""
            
            # 4. Material Manufacturer
            mat_manufacturer = ""
            try:
                manuf_param = mat.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MANUFACTURER)
                if manuf_param and manuf_param.HasValue:
                    mat_manufacturer = manuf_param.AsString() or ""
            except:
                mat_manufacturer = ""
            
            # 5. Material Model
            mat_model = ""
            try:
                model_param = mat.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MODEL)
                if model_param and model_param.HasValue:
                    mat_model = model_param.AsString() or ""
            except:
                mat_model = ""
            
            # 6. Material Asset Name
            mat_asset_name = ""
            try:
                # Prova a ottenere l'asset di rendering
                if hasattr(mat, 'AppearanceAssetId'):
                    asset_id = mat.AppearanceAssetId
                    if asset_id and asset_id != ElementId.InvalidElementId:
                        asset_elem = doc.GetElement(asset_id)
                        if asset_elem:
                            mat_asset_name = asset_elem.Name
            except:
                mat_asset_name = ""
            
            materials.append({
                'MaterialKey': "{} : {}".format(processor.file_name, mat.Id.IntegerValue),
                'MaterialID': mat.Id.IntegerValue,
                'FileName': processor.file_name,
                'MaterialName': mat.Name,
                'MaterialClass': mat_class,
                'MaterialCategory': mat_category,
                'MaterialDescription': mat_description,
                'MaterialComments': mat_comments,
                'MaterialKeywords': mat_keywords,
                'MaterialManufacturer': mat_manufacturer,
                'MaterialModel': mat_model,
                'MaterialAssetName': mat_asset_name,
                'IsUsed': is_used
            })
        except Exception as e:
            LOGGER.warning("Errore estrazione materiale: {}".format(str(e)))
    
    return materials


def extract_levels(processor):
    """Estrae livelli (TAB_Levels)."""
    doc = processor.doc
    levels = []
    
    try:
        # Raccogli tutti i livelli
        all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
        
        # Pre-carica i link per velocizzare la ricerca del nome file
        link_instances = {}
        try:
            rvt_links = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
            for link_inst in rvt_links:
                try:
                    link_doc = link_inst.GetLinkDocument()
                    if link_doc:
                        link_instances[link_doc.Title] = link_inst
                except:
                    pass
        except:
            pass
        
        for level in all_levels:
            try:
                # LevelID
                level_id = level.Id.IntegerValue
                
                # Level Name
                level_name = level.Name if level.Name else ""
                
                # Level Type (nome del tipo) - CORRETTO
                level_type_name = ""
                try:
                    type_id = level.GetTypeId()
                    if type_id and type_id != ElementId.InvalidElementId:
                        level_type = doc.GetElement(type_id)
                        if level_type:
                            # Metodo 1: Proprietà Name diretta
                            try:
                                level_type_name = level_type.Name
                            except:
                                pass
                            
                            # Metodo 2: Parametro SYMBOL_NAME_PARAM
                            if not level_type_name:
                                try:
                                    name_param = level_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                                    if name_param and name_param.HasValue:
                                        level_type_name = name_param.AsString() or ""
                                except:
                                    pass
                            
                            # Metodo 3: Element.Name.GetValue (IronPython)
                            if not level_type_name:
                                try:
                                    level_type_name = Element.Name.GetValue(level_type) or ""
                                except:
                                    pass
                except:
                    pass
                
                # Level Offset (elevazione formattata con Project Units)
                level_offset = ""
                try:
                    # Ottieni l'elevazione in unità interne (piedi)
                    elevation_internal = level.Elevation
                    
                    # Usa UnitFormatUtils per formattare con le Project Units
                    try:
                        # Revit 2022+ usa ForgeTypeId
                        from Autodesk.Revit.DB import UnitFormatUtils, SpecTypeId
                        level_offset = UnitFormatUtils.Format(
                            doc.GetUnits(), 
                            SpecTypeId.Length, 
                            elevation_internal, 
                            False  # False = non include simbolo unità
                        )
                    except:
                        # Fallback per versioni precedenti
                        try:
                            from Autodesk.Revit.DB import UnitFormatUtils, UnitType
                            level_offset = UnitFormatUtils.Format(
                                doc.GetUnits(), 
                                UnitType.UT_Length, 
                                elevation_internal, 
                                False, 
                                False
                            )
                        except:
                            # Fallback finale: converti manualmente in metri con 4 decimali
                            level_offset = "{:.4f}".format(elevation_internal * 0.3048)
                except:
                    level_offset = ""
                
                # IsPinned
                is_pinned = "NO"
                try:
                    if level.Pinned:
                        is_pinned = "YES"
                except:
                    pass
                
                # IsMonitor, MonitorFileName, MonitorLevel
                is_monitor = "NO"
                monitor_file_name = ""
                monitor_level_name = ""
                
                try:
                    # Verifica se il livello sta monitorando qualcosa
                    if level.IsMonitoringLinkElement():
                        is_monitor = "YES"
                        
                        # Ottieni gli ID degli elementi monitorati nei link
                        monitored_link_ids = level.GetMonitoredLinkElementIds()
                        
                        if monitored_link_ids and monitored_link_ids.Count > 0:
                            # Prendi il primo link monitorato
                            for monitored_link_id in monitored_link_ids:
                                try:
                                    # Ottieni l'istanza del link
                                    link_instance = doc.GetElement(monitored_link_id)
                                    if link_instance and isinstance(link_instance, RevitLinkInstance):
                                        # Nome del file linkato con estensione
                                        link_doc = link_instance.GetLinkDocument()
                                        if link_doc:
                                            monitor_file_name = link_doc.Title
                                            # Assicurati che abbia l'estensione .rvt
                                            if monitor_file_name and not monitor_file_name.endswith('.rvt'):
                                                monitor_file_name += '.rvt'
                                            
                                            # Trova il livello monitorato nel link
                                            monitored_elem_id = level.GetMonitoredLocalElementIds()
                                            # Usa GetMonitoredLinkElementIds per ottenere l'elemento nel link
                                            try:
                                                # Il livello monitorato ha lo stesso nome solitamente
                                                # Cerca nel documento linkato
                                                link_levels = FilteredElementCollector(link_doc).OfClass(Level).ToElements()
                                                for link_level in link_levels:
                                                    # Cerca un livello con elevazione simile
                                                    if abs(link_level.Elevation - level.Elevation) < 0.001:
                                                        monitor_level_name = link_level.Name
                                                        break
                                            except:
                                                pass
                                        break
                                except:
                                    pass
                except:
                    pass
                
                # ScopeBox associato al livello
                scope_box_name = ""
                try:
                    scope_box_param = level.get_Parameter(DB.BuiltInParameter.DATUM_VOLUME_OF_INTEREST)
                    if scope_box_param and scope_box_param.HasValue:
                        scope_box_id = scope_box_param.AsElementId()
                        if scope_box_id and scope_box_id != ElementId.InvalidElementId:
                            scope_box = doc.GetElement(scope_box_id)
                            if scope_box:
                                scope_box_name = scope_box.Name
                except:
                    scope_box_name = ""
                
                # Workset del livello (solo se workshared)
                workset_name = ""
                if processor.is_workshared:
                    try:
                        ws_param = level.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                        if ws_param and ws_param.HasValue:
                            ws_id = ws_param.AsInteger()
                            workset = doc.GetWorksetTable().GetWorkset(DB.WorksetId(ws_id))
                            if workset:
                                workset_name = workset.Name
                    except:
                        workset_name = ""
                
                levels.append({
                    'LevelKey': "{} : {}".format(processor.file_name, level_id),
                    'LevelID': level_id,
                    'FileName': processor.file_name,
                    'LevelName': level_name,
                    'LevelType': level_type_name,
                    'LevelOffset': level_offset,
                    'IsMonitor': is_monitor,
                    'MonitorFileName': monitor_file_name,
                    'MonitorLevel': monitor_level_name,
                    'IsPinned': is_pinned,
                    'ScopeBox': scope_box_name,
                    'Workset': workset_name
                })
                
            except Exception as e:
                LOGGER.warning("Errore estrazione livello: {}".format(str(e)))
        
        OUTPUT.print_md("      ✓ {} livelli estratti".format(len(levels)))
        
    except Exception as e:
        OUTPUT.print_md("      ⚠️ Errore estrazione livelli: {}".format(str(e)))
        LOGGER.error("Errore extract_levels: {}".format(str(e)))
    
    return levels


def extract_filters(processor):
    """Estrae filtri (TAB_Filters)."""
    doc = processor.doc
    filters = []
    
    try:
        # Raccogli tutti i ParameterFilterElement
        all_filters = list(FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements())
        
        if not all_filters:
            OUTPUT.print_md("      ℹ️ Nessun filtro trovato nel modello")
            return filters
        
        # Verifica quali filtri sono usati nelle viste o nei view template
        used_filter_ids = set()
        
        try:
            all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
            
            for view in all_views:
                try:
                    filter_ids = view.GetFilters()
                    for fid in filter_ids:
                        used_filter_ids.add(fid.IntegerValue)
                except:
                    pass
        except Exception as e:
            LOGGER.warning("Errore raccolta filtri usati: {}".format(str(e)))
        
        # Estrai i filtri
        for flt in all_filters:
            try:
                is_used = "YES" if flt.Id.IntegerValue in used_filter_ids else "NO"
                
                filters.append({
                    'FilterKey': "{} : {}".format(processor.file_name, flt.Id.IntegerValue),
                    'FilterID': flt.Id.IntegerValue,
                    'FileName': processor.file_name,
                    'FilterName': flt.Name if flt.Name else "Unnamed",
                    'IsUsed': is_used
                })
            except Exception as e:
                LOGGER.warning("Errore estrazione filtro: {}".format(str(e)))
        
        OUTPUT.print_md("      ✓ {} filtri estratti".format(len(filters)))
        
    except Exception as e:
        OUTPUT.print_md("      ⚠️ Errore estrazione filtri: {}".format(str(e)))
        LOGGER.error("Errore extract_filters: {}".format(str(e)))
    
    return filters


def extract_rooms(processor):
    """Estrae stanze (TAB_Rooms)."""
    doc = processor.doc
    rooms = []
    
    # Fattori di conversione da piedi a metri
    FOOT_TO_METER = 0.3048
    SQFOOT_TO_SQMETER = 0.092903
    CUFOOT_TO_CUMETER = 0.0283168
    
    try:
        # Raccogli tutte le stanze
        room_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).ToElements()
        
        if not room_collector:
            OUTPUT.print_md("      ℹ️ Nessuna stanza trovata nel modello")
            return rooms
        
        for room in room_collector:
            try:
                # RoomID
                room_id = room.Id.IntegerValue
                
                # RoomName
                room_name = ""
                try:
                    name_param = room.get_Parameter(BuiltInParameter.ROOM_NAME)
                    if name_param and name_param.HasValue:
                        room_name = name_param.AsString() or ""
                except:
                    pass
                
                # RoomNumber
                room_number = ""
                try:
                    number_param = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                    if number_param and number_param.HasValue:
                        room_number = number_param.AsString() or ""
                except:
                    pass
                
                # Level
                level_name = ""
                try:
                    level_param = room.get_Parameter(BuiltInParameter.ROOM_LEVEL_ID)
                    if level_param and level_param.HasValue:
                        level_id = level_param.AsElementId()
                        level_elem = doc.GetElement(level_id)
                        if level_elem:
                            level_name = level_elem.Name
                except:
                    pass
                
                # Area (m² - 1 decimale) - Lascia vuoto se "Not Enclosed"
                area_value = ""
                try:
                    area_param = room.get_Parameter(BuiltInParameter.ROOM_AREA)
                    if area_param and area_param.HasValue:
                        area_sqft = area_param.AsDouble()
                        # Se area <= 0, significa "Not Enclosed" - lascia vuoto
                        if area_sqft > 0:
                            area_sqm = area_sqft * SQFOOT_TO_SQMETER
                            area_value = _format_decimal(area_sqm, 1)
                except:
                    pass
                
                # Perimeter (m - 2 decimali)
                perimeter_value = ""
                try:
                    perim_param = room.get_Parameter(BuiltInParameter.ROOM_PERIMETER)
                    if perim_param and perim_param.HasValue:
                        perim_ft = perim_param.AsDouble()
                        if perim_ft > 0:
                            perim_m = perim_ft * FOOT_TO_METER
                            perimeter_value = _format_decimal(perim_m, 2)
                except:
                    pass
                
                # Height (m - 2 decimali)
                height_value = ""
                try:
                    height_param = room.get_Parameter(BuiltInParameter.ROOM_HEIGHT)
                    if height_param and height_param.HasValue:
                        height_ft = height_param.AsDouble()
                        if height_ft > 0:
                            height_m = height_ft * FOOT_TO_METER
                            height_value = _format_decimal(height_m, 2)
                except:
                    pass
                
                # Volume (m³ - 1 decimale) - Lascia vuoto se "Not Computed"
                volume_value = ""
                try:
                    vol_param = room.get_Parameter(BuiltInParameter.ROOM_VOLUME)
                    if vol_param and vol_param.HasValue:
                        vol_cuft = vol_param.AsDouble()
                        # Se volume <= 0, significa "Not Computed" - lascia vuoto
                        if vol_cuft > 0:
                            vol_cum = vol_cuft * CUFOOT_TO_CUMETER
                            volume_value = _format_decimal(vol_cum, 1)
                except:
                    pass
                # Determino lo stato usando proprietà API dirette
                # - Not Placed: Location is None
                # - Redundant: Ha Location, ha boundaries, ma Area = 0
                # - Not Enclosed: Ha Location, nessun boundary, Area = 0
                
                area_display_str = ""  # Per debug
                is_placed = "NO"
                is_enclosed = "NO"
                is_redundant = "NO"
                
                try:
                    # 1. Verifica se è posizionata
                    has_location = room.Location is not None
                    
                    if has_location:
                        is_placed = "YES"
                        
                        # 2. Leggo l'area numerica
                        room_area = room.Area  # In sq feet
                        
                        if room_area > 0:
                            # Ha area > 0 quindi è enclosed e non redundant
                            is_enclosed = "YES"
                            is_redundant = "NO"
                            area_display_str = "Area > 0"
                        else:
                            # Area = 0, verifichiamo se ha boundaries
                            try:
                                # SpatialElementBoundaryOptions
                                options = DB.SpatialElementBoundaryOptions()
                                boundaries = room.GetBoundarySegments(options)
                                
                                if boundaries and boundaries.Count > 0:
                                    # Ha boundaries ma Area = 0 -> Redundant
                                    is_enclosed = "YES"
                                    is_redundant = "YES"
                                    area_display_str = "Redundant (has boundaries, area=0)"
                                else:
                                    # Nessun boundary -> Not Enclosed
                                    is_enclosed = "NO"
                                    is_redundant = "NO"
                                    area_display_str = "Not Enclosed (no boundaries)"
                            except:
                                # Fallback: se non possiamo verificare i boundaries
                                is_enclosed = "NO"
                                area_display_str = "Area = 0 (boundaries check failed)"
                    else:
                        # Non posizionata
                        is_placed = "NO"
                        is_enclosed = "NO"
                        is_redundant = "NO"
                        area_display_str = "Not Placed"
                except Exception as ex:
                    area_display_str = "Error: {}".format(str(ex))
                
                # Phase
                phase_name = ""
                try:
                    phase_param = room.get_Parameter(BuiltInParameter.ROOM_PHASE)
                    if phase_param and phase_param.HasValue:
                        phase_id = phase_param.AsElementId()
                        phase_elem = doc.GetElement(phase_id)
                        if phase_elem:
                            phase_name = phase_elem.Name
                except:
                    pass
                
                # Workset
                workset_name = ""
                if processor.is_workshared:
                    try:
                        ws_param = room.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                        if ws_param and ws_param.HasValue:
                            workset_id = ws_param.AsInteger()
                            try:
                                workset = doc.GetWorksetTable().GetWorkset(DB.WorksetId(workset_id))
                                if workset:
                                    workset_name = workset.Name
                            except:
                                workset_name = "Workset_{}".format(workset_id)
                    except:
                        pass
                
                rooms.append({
                    'RoomKey': "{} : {}".format(processor.file_name, room_id),
                    'RoomID': room_id,
                    'FileName': processor.file_name,
                    'RoomName': room_name,
                    'RoomNumber': room_number,
                    'Level': level_name,
                    'Area_sqm': area_value,
                    'Perimeter_m': perimeter_value,
                    'Height_m': height_value,
                    'Volume_mc': volume_value,
                    'AreaString': area_display_str,
                    'IsPlaced': is_placed,
                    'IsEnclosed': is_enclosed,
                    'IsRedundant': is_redundant,
                    'Phase': phase_name,
                    'Workset': workset_name
                })
                
            except Exception as e:
                LOGGER.warning("Errore estrazione stanza: {}".format(str(e)))
        
        OUTPUT.print_md("      ✓ {} stanze estratte".format(len(rooms)))
        
    except Exception as e:
        OUTPUT.print_md("      ⚠️ Errore estrazione stanze: {}".format(str(e)))
        LOGGER.error("Errore extract_rooms: {}".format(str(e)))
    
    return rooms


def extract_spaces(processor):
    """Estrae vani (TAB_Spaces)."""
    doc = processor.doc
    spaces = []
    
    # Fattori di conversione da piedi a metri
    FOOT_TO_METER = 0.3048
    SQFOOT_TO_SQMETER = 0.092903
    CUFOOT_TO_CUMETER = 0.0283168
    
    try:
        # Raccogli tutti i vani (MEP Spaces)
        space_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).ToElements()
        
        if not space_collector:
            OUTPUT.print_md("      ℹ️ Nessun vano trovato nel modello")
            return spaces
        
        for space in space_collector:
            try:
                # SpaceID
                space_id = space.Id.IntegerValue
                
                # SpaceName
                space_name = ""
                try:
                    name_param = space.get_Parameter(BuiltInParameter.ROOM_NAME)
                    if name_param and name_param.HasValue:
                        space_name = name_param.AsString() or ""
                except:
                    pass
                
                # SpaceNumber
                space_number = ""
                try:
                    number_param = space.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                    if number_param and number_param.HasValue:
                        space_number = number_param.AsString() or ""
                except:
                    pass
                
                # Level
                level_name = ""
                try:
                    level_param = space.get_Parameter(BuiltInParameter.ROOM_LEVEL_ID)
                    if level_param and level_param.HasValue:
                        level_id = level_param.AsElementId()
                        level_elem = doc.GetElement(level_id)
                        if level_elem:
                            level_name = level_elem.Name
                except:
                    pass
                
                # Area (m² - 1 decimale) - Lascia vuoto se "Not Enclosed"
                area_value = ""
                try:
                    area_param = space.get_Parameter(BuiltInParameter.ROOM_AREA)
                    if area_param and area_param.HasValue:
                        area_sqft = area_param.AsDouble()
                        # Se area <= 0, significa "Not Enclosed" - lascia vuoto
                        if area_sqft > 0:
                            area_sqm = area_sqft * SQFOOT_TO_SQMETER
                            area_value = _format_decimal(area_sqm, 1)
                except:
                    pass
                
                # Height (m - 2 decimali)
                height_value = ""
                try:
                    height_param = space.get_Parameter(BuiltInParameter.ROOM_HEIGHT)
                    if height_param and height_param.HasValue:
                        height_ft = height_param.AsDouble()
                        if height_ft > 0:
                            height_m = height_ft * FOOT_TO_METER
                            height_value = _format_decimal(height_m, 2)
                except:
                    pass
                
                # Volume (m³ - 1 decimale) - Lascia vuoto se "Not Computed"
                volume_value = ""
                try:
                    vol_param = space.get_Parameter(BuiltInParameter.ROOM_VOLUME)
                    if vol_param and vol_param.HasValue:
                        vol_cuft = vol_param.AsDouble()
                        # Se volume <= 0, significa "Not Computed" - lascia vuoto
                        if vol_cuft > 0:
                            vol_cum = vol_cuft * CUFOOT_TO_CUMETER
                            volume_value = _format_decimal(vol_cum, 1)
                except:
                    pass
                
                # Zone
                zone_name = ""
                try:
                    zone_param = space.get_Parameter(BuiltInParameter.SPACE_ZONE_NAME)
                    if zone_param and zone_param.HasValue:
                        zone_name = zone_param.AsString() or ""
                except:
                    pass
                # Determino lo stato usando proprietà API dirette
                # - Not Placed: Location is None
                # - Redundant: Ha Location, ha boundaries, ma Area = 0
                # - Not Enclosed: Ha Location, nessun boundary, Area = 0
                
                area_display_str = ""  # Per debug
                is_placed = "NO"
                is_enclosed = "NO"
                is_redundant = "NO"
                
                try:
                    # 1. Verifica se è posizionato
                    has_location = space.Location is not None
                    
                    if has_location:
                        is_placed = "YES"
                        
                        # 2. Leggo l'area numerica
                        space_area = space.Area  # In sq feet
                        
                        if space_area > 0:
                            # Ha area > 0 quindi è enclosed e non redundant
                            is_enclosed = "YES"
                            is_redundant = "NO"
                            area_display_str = "Area > 0"
                        else:
                            # Area = 0, verifichiamo se ha boundaries
                            try:
                                # SpatialElementBoundaryOptions
                                options = DB.SpatialElementBoundaryOptions()
                                boundaries = space.GetBoundarySegments(options)
                                
                                if boundaries and boundaries.Count > 0:
                                    # Ha boundaries ma Area = 0 -> Redundant
                                    is_enclosed = "YES"
                                    is_redundant = "YES"
                                    area_display_str = "Redundant (has boundaries, area=0)"
                                else:
                                    # Nessun boundary -> Not Enclosed
                                    is_enclosed = "NO"
                                    is_redundant = "NO"
                                    area_display_str = "Not Enclosed (no boundaries)"
                            except:
                                # Fallback: se non possiamo verificare i boundaries
                                is_enclosed = "NO"
                                area_display_str = "Area = 0 (boundaries check failed)"
                    else:
                        # Non posizionato
                        is_placed = "NO"
                        is_enclosed = "NO"
                        is_redundant = "NO"
                        area_display_str = "Not Placed"
                except Exception as ex:
                    area_display_str = "Error: {}".format(str(ex))
                
                # Phase
                phase_name = ""
                try:
                    phase_param = space.get_Parameter(BuiltInParameter.ROOM_PHASE)
                    if phase_param and phase_param.HasValue:
                        phase_id = phase_param.AsElementId()
                        phase_elem = doc.GetElement(phase_id)
                        if phase_elem:
                            phase_name = phase_elem.Name
                except:
                    pass
                
                # Workset
                workset_name = ""
                if processor.is_workshared:
                    try:
                        ws_param = space.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                        if ws_param and ws_param.HasValue:
                            workset_id = ws_param.AsInteger()
                            try:
                                workset = doc.GetWorksetTable().GetWorkset(DB.WorksetId(workset_id))
                                if workset:
                                    workset_name = workset.Name
                            except:
                                workset_name = "Workset_{}".format(workset_id)
                    except:
                        pass
                
                spaces.append({
                    'SpaceKey': "{} : {}".format(processor.file_name, space_id),
                    'SpaceID': space_id,
                    'FileName': processor.file_name,
                    'SpaceName': space_name,
                    'SpaceNumber': space_number,
                    'Level': level_name,
                    'Area_sqm': area_value,
                    'Height_m': height_value,
                    'Volume_mc': volume_value,
                    'Zone': zone_name,
                    'AreaString': area_display_str,
                    'IsPlaced': is_placed,
                    'IsEnclosed': is_enclosed,
                    'IsRedundant': is_redundant,
                    'Phase': phase_name,
                    'Workset': workset_name
                })
                
            except Exception as e:
                LOGGER.warning("Errore estrazione vano: {}".format(str(e)))
        
        OUTPUT.print_md("      ✓ {} vani estratti".format(len(spaces)))
        
    except Exception as e:
        OUTPUT.print_md("      ⚠️ Errore estrazione vani: {}".format(str(e)))
        LOGGER.error("Errore extract_spaces: {}".format(str(e)))
    
    return spaces


def extract_areas(processor):
    """Estrae aree (TAB_Areas)."""
    doc = processor.doc
    areas = []
    
    # Fattori di conversione da piedi a metri
    FOOT_TO_METER = 0.3048
    SQFOOT_TO_SQMETER = 0.092903
    
    try:
        # Raccogli tutte le aree
        area_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).ToElements()
        
        if not area_collector:
            OUTPUT.print_md("      ℹ️ Nessuna area trovata nel modello")
            return areas
        
        for area in area_collector:
            try:
                # AreaID
                area_id = area.Id.IntegerValue
                
                # AreaName
                area_name = ""
                try:
                    name_param = area.get_Parameter(BuiltInParameter.ROOM_NAME)
                    if name_param and name_param.HasValue:
                        area_name = name_param.AsString() or ""
                except:
                    pass
                
                # AreaNumber
                area_number = ""
                try:
                    number_param = area.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                    if number_param and number_param.HasValue:
                        area_number = number_param.AsString() or ""
                except:
                    pass
                
                # AreaType (parametro Area Type)
                area_type = ""
                try:
                    # Prova prima con il parametro AREA_TYPE
                    type_param = area.get_Parameter(BuiltInParameter.AREA_TYPE)
                    if type_param and type_param.HasValue:
                        area_type = type_param.AsValueString() or ""
                    # Fallback: prova con AREA_SCHEME_NAME
                    if not area_type:
                        scheme_param = area.get_Parameter(BuiltInParameter.AREA_SCHEME_NAME)
                        if scheme_param and scheme_param.HasValue:
                            area_type = scheme_param.AsString() or ""
                except:
                    pass
                
                # Area (m² - 1 decimale) - Lascia vuoto se "Not Enclosed"
                area_value = ""
                try:
                    area_param = area.get_Parameter(BuiltInParameter.ROOM_AREA)
                    if area_param and area_param.HasValue:
                        area_sqft = area_param.AsDouble()
                        # Se area <= 0, significa "Not Enclosed" - lascia vuoto
                        if area_sqft > 0:
                            area_sqm = area_sqft * SQFOOT_TO_SQMETER
                            area_value = _format_decimal(area_sqm, 1)
                except:
                    pass
                
                # Perimeter (m - 2 decimali)
                perimeter_value = ""
                try:
                    perim_param = area.get_Parameter(BuiltInParameter.ROOM_PERIMETER)
                    if perim_param and perim_param.HasValue:
                        perim_ft = perim_param.AsDouble()
                        if perim_ft > 0:
                            perim_m = perim_ft * FOOT_TO_METER
                            perimeter_value = _format_decimal(perim_m, 2)
                except:
                    pass
                
                # Level (livello su cui è posizionata l'area)
                level_name = ""
                try:
                    level_param = area.get_Parameter(BuiltInParameter.ROOM_LEVEL_ID)
                    if level_param and level_param.HasValue:
                        level_id = level_param.AsElementId()
                        level_elem = doc.GetElement(level_id)
                        if level_elem:
                            level_name = level_elem.Name
                except:
                    pass
                
                # IsPlaced
                is_placed = "NO"
                try:
                    if hasattr(area, 'Location') and area.Location is not None:
                        is_placed = "YES"
                except:
                    pass
                
                # Workset
                workset_name = ""
                if processor.is_workshared:
                    try:
                        ws_param = area.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                        if ws_param and ws_param.HasValue:
                            workset_id = ws_param.AsInteger()
                            try:
                                workset = doc.GetWorksetTable().GetWorkset(DB.WorksetId(workset_id))
                                if workset:
                                    workset_name = workset.Name
                            except:
                                workset_name = "Workset_{}".format(workset_id)
                    except:
                        pass
                
                areas.append({
                    'AreaKey': "{} : {}".format(processor.file_name, area_id),
                    'AreaID': area_id,
                    'FileName': processor.file_name,
                    'AreaName': area_name,
                    'AreaNumber': area_number,
                    'AreaType': area_type,
                    'Area_sqm': area_value,
                    'Perimeter_m': perimeter_value,
                    'Level': level_name,
                    'IsPlaced': is_placed,
                    'Workset': workset_name
                })
                
            except Exception as e:
                LOGGER.warning("Errore estrazione area: {}".format(str(e)))
        
        OUTPUT.print_md("      ✓ {} aree estratte".format(len(areas)))
        
    except Exception as e:
        OUTPUT.print_md("      ⚠️ Errore estrazione aree: {}".format(str(e)))
        LOGGER.error("Errore extract_areas: {}".format(str(e)))
    
    return areas


def extract_tags(processor):
    """
    Estrae tutti i tag dal modello (TAB_Tags).
    Raccoglie tutti i tag da tutte le categorie di tag disponibili.
    
    Args:
        processor: FileProcessor con il documento aperto
    
    Returns:
        Lista di dizionari con i dati dei tag
    """
    doc = processor.doc
    tags = []
    
    try:
        # Lista delle categorie di tag in Revit
        # Usiamo stringhe per compatibilità con diverse versioni di Revit
        tag_category_names = [
            # Architectural Tags
            "OST_RoomTags",
            "OST_AreaTags",
            "OST_DoorTags",
            "OST_WindowTags",
            "OST_WallTags",
            "OST_CurtainWallPanelTags",
            "OST_FloorTags",
            "OST_CeilingTags",
            "OST_RoofTags",
            "OST_StairsTags",
            "OST_RailingTags",
            "OST_ColumnTags",
            "OST_FurnitureTags",
            "OST_CaseworkTags",
            "OST_GenericModelTags",
            "OST_PlantingTags",
            "OST_SiteTags",
            "OST_ParkingTags",
            "OST_SpecialityEquipmentTags",
            "OST_KeynoteTags",
            "OST_MaterialTags",
            "OST_MultiCategoryTags",
            # Structural Tags
            "OST_StructuralFramingTags",
            "OST_StructuralColumnTags",
            "OST_StructuralFoundationTags",
            "OST_StructConnectionTags",
            "OST_RebarTags",
            "OST_FabricAreaTags",
            "OST_TrussTags",
            # MEP Tags
            "OST_SpaceTags",
            "OST_DuctTags",
            "OST_PipeTags",
            "OST_FlexDuctTags",
            "OST_FlexPipeTags",
            "OST_DuctFittingTags",
            "OST_PipeFittingTags",
            "OST_DuctAccessoryTags",
            "OST_PipeAccessoryTags",
            "OST_DuctInsulationsTags",
            "OST_PipeInsulationsTags",
            "OST_DuctTerminalTags",
            "OST_MechanicalEquipmentTags",
            "OST_PlumbingFixtureTags",
            "OST_SprinklerTags",
            "OST_LightingFixtureTags",
            "OST_ElectricalEquipmentTags",
            "OST_ElectricalFixtureTags",
            "OST_CableTrayTags",
            "OST_ConduitTags",
            "OST_DataDeviceTags",
            "OST_CommunicationDeviceTags",
            "OST_FireAlarmDeviceTags",
            "OST_NurseCallDeviceTags",
            "OST_SecurityDeviceTags",
            "OST_TelephoneDeviceTags",
        ]
        
        # Converti i nomi in categorie valide
        valid_categories = []
        for cat_name in tag_category_names:
            try:
                bic = getattr(DB.BuiltInCategory, cat_name, None)
                if bic is not None:
                    valid_categories.append(bic)
            except:
                pass
        
        # Raccogli tutti i tag da tutte le categorie
        for bic in valid_categories:
            try:
                collector = DB.FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
                
                for tag in collector:
                    try:
                        # Escludi tag nelle viste di legenda (sono solo simboli, non veri tag)
                        try:
                            owner_view_id = tag.OwnerViewId
                            if owner_view_id is not None and owner_view_id != DB.ElementId.InvalidElementId:
                                owner_view = doc.GetElement(owner_view_id)
                                if owner_view is not None and owner_view.ViewType == DB.ViewType.Legend:
                                    continue
                        except:
                            pass

                        # TagID
                        tag_id = tag.Id.IntegerValue

                        # ViewID
                        view_id = ""
                        try:
                            if hasattr(tag, 'OwnerViewId'):
                                view_id = tag.OwnerViewId.IntegerValue
                        except:
                            pass
                        
                        # FamilyName e TypeName
                        family_name = ""
                        type_name = ""
                        try:
                            # Metodo 1: Usa Symbol (per FamilyInstance)
                            if hasattr(tag, 'Symbol') and tag.Symbol:
                                symbol = tag.Symbol
                                type_name = symbol.Name if hasattr(symbol, 'Name') else ""
                                if hasattr(symbol, 'Family') and symbol.Family:
                                    family_name = symbol.Family.Name
                            
                            # Metodo 2: Usa GetTypeId se Symbol non funziona
                            if not type_name:
                                tag_type_id = tag.GetTypeId()
                                if tag_type_id and tag_type_id != DB.ElementId.InvalidElementId:
                                    tag_type = doc.GetElement(tag_type_id)
                                    if tag_type:
                                        type_name = tag_type.Name if hasattr(tag_type, 'Name') else ""
                                        if hasattr(tag_type, 'Family') and tag_type.Family:
                                            family_name = tag_type.Family.Name
                                        elif hasattr(tag_type, 'FamilyName'):
                                            family_name = tag_type.FamilyName
                            
                            # Metodo 3: Usa parametri built-in come fallback
                            if not family_name:
                                fam_param = tag.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
                                if fam_param and fam_param.HasValue:
                                    family_name = fam_param.AsValueString() or ""
                            
                            if not type_name:
                                type_param = tag.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM)
                                if type_param and type_param.HasValue:
                                    type_name = type_param.AsValueString() or ""
                        except:
                            pass
                        
                        # TagCategory (nome della categoria del tag)
                        tag_category = ""
                        try:
                            if tag.Category:
                                tag_category = tag.Category.Name
                        except:
                            pass
                        
                        # HasHost - verifica se il tag ha un elemento host
                        has_host = "NO"

                        # --- SpatialElementTag (RoomTag, SpaceTag, AreaTag) ---
                        # Usano TaggedLocalRoomId / TaggedRoomId, NON GetTaggedElementIds()

                        # Metodo 1a: RoomTag.TaggedLocalRoomId (locale)
                        if has_host == "NO":
                            try:
                                rid = tag.TaggedLocalRoomId
                                if rid is not None and rid != DB.ElementId.InvalidElementId:
                                    has_host = "YES"
                            except:
                                pass

                        # Metodo 1b: RoomTag.TaggedRoomId (anche link)
                        if has_host == "NO":
                            try:
                                leid = tag.TaggedRoomId
                                if leid is not None:
                                    if leid.HostElementId != DB.ElementId.InvalidElementId or leid.LinkedElementId != DB.ElementId.InvalidElementId:
                                        has_host = "YES"
                            except:
                                pass

                        # Metodo 1c: SpaceTag - TaggedLocalSpaceId
                        if has_host == "NO":
                            try:
                                sid = tag.TaggedLocalSpaceId
                                if sid is not None and sid != DB.ElementId.InvalidElementId:
                                    has_host = "YES"
                            except:
                                pass

                        # Metodo 1d: AreaTag - TaggedLocalAreaId
                        if has_host == "NO":
                            try:
                                aid = tag.TaggedLocalAreaId
                                if aid is not None and aid != DB.ElementId.InvalidElementId:
                                    has_host = "YES"
                            except:
                                pass

                        # --- IndependentTag (tutti gli altri tag) ---

                        # Metodo 2: GetTaggedElementIds() - copre sia elementi locali che in link
                        if has_host == "NO":
                            try:
                                tagged_refs = tag.GetTaggedElementIds()
                                if tagged_refs is not None and len(list(tagged_refs)) > 0:
                                    has_host = "YES"
                            except:
                                pass

                        # Metodo 3: GetTaggedLocalElementIds() - solo elementi locali
                        if has_host == "NO":
                            try:
                                tagged_ids = tag.GetTaggedLocalElementIds()
                                if tagged_ids is not None and len(list(tagged_ids)) > 0:
                                    has_host = "YES"
                            except:
                                pass

                        # Metodo 4: Host property (fallback generico, es. legende)
                        if has_host == "NO":
                            try:
                                h = tag.Host
                                if h is not None:
                                    has_host = "YES"
                            except:
                                pass
                        
                        tags.append({
                            'TagKey': "{} : {}".format(processor.file_name, tag_id),
                            'TagID': tag_id,
                            'FileName': processor.file_name,
                            'ViewID': view_id,
                            'ViewKey': "{} : {}".format(processor.file_name, view_id),
                            'FamilyName': family_name,
                            'TypeName': type_name,
                            'TagCategory': tag_category,
                            'HasHost': has_host
                        })
                        
                    except Exception as e:
                        LOGGER.warning("Errore estrazione tag singolo: {}".format(str(e)))
                        
            except Exception as e:
                # Categoria non valida o non presente nel modello, ignora silenziosamente
                pass
        
        OUTPUT.print_md("      ✓ {} tag estratti".format(len(tags)))
        
    except Exception as e:
        OUTPUT.print_md("      ⚠️ Errore estrazione tag: {}".format(str(e)))
        LOGGER.error("Errore extract_tags: {}".format(str(e)))
    
    return tags



def extract_families_types_instances(processor, custom_instance_params=None, custom_type_params=None):
    """
    Estrae famiglie, tipi e istanze dal modello.
    Popola TAB_Families, TAB_Types, TAB_Instances in un unico loop sugli elementi.
    
    Args:
        processor: FileProcessor con il documento aperto
        custom_instance_params: Lista di nomi di parametri custom di istanza da estrarre
        custom_type_params: Lista di nomi di parametri custom di tipo da estrarre
    
    Returns:
        tuple: (families_data, types_data, instances_data)
    """
    doc = processor.doc
    
    if custom_instance_params is None:
        custom_instance_params = []
    if custom_type_params is None:
        custom_type_params = []
    
    # Track parametri non trovati (per avviso finale)
    instance_params_not_found = set(custom_instance_params)
    type_params_not_found = set(custom_type_params)
    
    # Dizionari per deduplicazione famiglie e tipi (per file)
    families_dict = {}  # key: family_id -> family_data dict
    types_dict = {}     # key: type_id_int -> type_data dict
    instances = []
    
    OUTPUT.print_md("      ⏳ Raccolta famiglie, tipi e istanze...")
    
    try:
        # Lista COMPLETA di nomi categorie built-in da estrarre
        target_category_names = [
    # ===== ARCHITETTURA =====
    "OST_Walls",
    "OST_Floors",
    "OST_Roofs",
    "OST_Ceilings",
    "OST_Doors",
    "OST_Windows",
    "OST_Furniture",
    "OST_FurnitureSystems",
    "OST_Casework",
    "OST_GenericModel",
    "OST_Columns",
    "OST_Stairs",
    "OST_Ramps",
    "OST_Railings",
    "OST_CurtainWallPanels",
    "OST_CurtainWallMullions",
    "OST_Curtain_Systems",
    "OST_SpecialityEquipment",
    "OST_Mass",
    "OST_Signage",  # Revit 2022+
    
    # ===== SITE / ESTERNO =====
    "OST_Site",
    "OST_Topography",
    "OST_Toposolid",  # Revit 2024+
    "OST_BuildingPad",
    "OST_Parking",
    "OST_Planting",
    "OST_Entourage",
    "OST_Hardscape",  # Revit 2022+
    "OST_Roads",
    
    # ===== STRUTTURALE =====
    "OST_StructuralColumns",
    "OST_StructuralFraming",
    "OST_StructuralFramingSystem",
    "OST_StructuralFoundation",
    "OST_StructuralTruss",
    "OST_StructuralStiffener",
    "OST_Rebar",
    "OST_FabricReinforcement",
    "OST_AreaRein",
    "OST_Coupler",
    "OST_TemporaryStructure",  # Revit 2022+
    "OST_VerticalCirculation",  # Revit 2022+
    
    # ===== MEP - MECHANICAL (HVAC) =====
    "OST_DuctCurves",
    "OST_FlexDuctCurves",
    "OST_DuctFitting",
    "OST_DuctAccessory",
    "OST_DuctInsulations",
    "OST_DuctLinings",
    "OST_DuctTerminal",
    "OST_MechanicalEquipment",
    
    # ===== MEP - PLUMBING =====
    "OST_PipeCurves",
    "OST_FlexPipeCurves",
    "OST_PipeFitting",
    "OST_PipeAccessory",
    "OST_PipeInsulations",
    "OST_PlumbingFixtures",
    "OST_PlumbingEquipment",  # Revit 2022+
    "OST_Sprinklers",
    
    # ===== MEP - ELECTRICAL =====
    "OST_ElectricalEquipment",
    "OST_ElectricalFixtures",
    "OST_LightingFixtures",
    "OST_LightingDevices",
    "OST_CableTray",
    "OST_CableTrayFitting",
    "OST_Conduit",
    "OST_ConduitFitting",
    "OST_Wire",
    
    # ===== MEP - ELECTRICAL DEVICES =====
    "OST_DataDevices",
    "OST_FireAlarmDevices",
    "OST_CommunicationDevices",
    "OST_SecurityDevices",
    "OST_NurseCallDevices",
    "OST_TelephoneDevices",
    "OST_AudioVisualDevices",  # Revit 2022+
    
    # ===== ATTREZZATURE SPECIALI =====
    "OST_FoodServiceEquipment",  # Revit 2022+
    "OST_MedicalEquipment",  # Revit 2022+
    "OST_FireProtection",  # Revit 2022+ (bonus: potrebbe interessarti)
]
        
        # Converti i nomi in BuiltInCategory solo se esistono
        target_categories = []
        for cat_name in target_category_names:
            try:
                cat = getattr(BuiltInCategory, cat_name, None)
                if cat is not None:
                    target_categories.append(cat)
            except:
                pass  # Categoria non esiste in questa versione di Revit
        
        total_count = 0
        
        # Estrai elementi per ogni categoria
        for built_in_cat in target_categories:
            try:
                collector = FilteredElementCollector(doc)\
                    .OfCategory(built_in_cat)\
                    .WhereElementIsNotElementType()
                
                cat_elements = list(collector)
                
                if not cat_elements:
                    continue
                
                # Ottieni il nome della categoria dal primo elemento
                category_name = "Unknown"
                if cat_elements and cat_elements[0].Category:
                    category_name = cat_elements[0].Category.Name
                
                for elem in cat_elements:
                    try:
                        # Salta elementi senza categoria valida
                        if not elem.Category:
                            continue
                        
                        # === TIPO ===
                        type_id = elem.GetTypeId()
                        if not type_id or type_id == ElementId.InvalidElementId:
                            continue
                        
                        elem_type = doc.GetElement(type_id)
                        if not elem_type:
                            continue
                        
                        type_id_int = type_id.IntegerValue
                        
                        # === FAMIGLIA ===
                        family_name = ""
                        family_id = 0
                        is_in_place = "NO"
                        is_system = "YES"
                        
                        # Prova prima: FamilyInstance → loadable family
                        if isinstance(elem, FamilyInstance):
                            try:
                                fam = elem.Symbol.Family
                                if fam:
                                    family_name = fam.Name or ""
                                    family_id = fam.Id.IntegerValue
                                    is_in_place = "YES" if fam.IsInPlace else "NO"
                                    is_system = "NO"
                            except:
                                pass
                        
                        # Fallback: sistema o mancante
                        if not family_name:
                            try:
                                if hasattr(elem_type, 'FamilyName') and elem_type.FamilyName:
                                    family_name = elem_type.FamilyName
                            except:
                                pass
                            if not family_name:
                                try:
                                    fam_param = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                                    if fam_param and fam_param.HasValue:
                                        family_name = fam_param.AsString() or ""
                                except:
                                    pass
                            if not family_name and elem.Category:
                                family_name = elem.Category.Name or ""
                            
                            # Per famiglie di sistema, genera un ID sintetico deterministico (negativo)
                            fam_key_str = "{}|{}".format(category_name, family_name)
                            family_id = -abs(hash(fam_key_str)) % 10000000
                            is_system = "YES"
                            is_in_place = "NO"
                        
                        family_key = "{} : {}".format(processor.file_name, family_id)
                        
                        # Aggiungi famiglia se non già tracciata
                        if family_id not in families_dict:
                            families_dict[family_id] = {
                                'FamilyKey': family_key,
                                'FamilyID': family_id,
                                'FileName': processor.file_name,
                                'FamilyName': family_name,
                                'Category': category_name,
                                'IsInPlace': is_in_place,
                                'IsSystemFamily': is_system,
                                'IsUsed': 'YES',
                            }
                        
                        # === TIPO (deduplicato) ===
                        if type_id_int not in types_dict:
                            # TypeName - estrazione robusta
                            type_name = ""
                            try:
                                if elem_type.Name:
                                    type_name = elem_type.Name
                            except:
                                pass
                            if not type_name:
                                try:
                                    symbol_param = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                                    if symbol_param and symbol_param.HasValue:
                                        type_name = symbol_param.AsString() or ""
                                except:
                                    pass
                            if not type_name:
                                try:
                                    type_param = elem_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                                    if type_param and type_param.HasValue:
                                        type_name = type_param.AsString() or ""
                                except:
                                    pass
                            if not type_name:
                                try:
                                    type_name = Element.Name.GetValue(elem_type) or ""
                                except:
                                    pass
                            
                            # TypeMark
                            type_mark = ""
                            try:
                                mark_param = elem_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK)
                                if mark_param and mark_param.HasValue:
                                    type_mark = mark_param.AsString() or ""
                            except:
                                pass
                            
                            # Description
                            description = ""
                            try:
                                desc_param = elem_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_DESCRIPTION)
                                if desc_param and desc_param.HasValue:
                                    description = desc_param.AsString() or ""
                            except:
                                pass
                            
                            # Family&Type
                            family_and_type = ""
                            if family_name and type_name:
                                family_and_type = "{}: {}".format(family_name, type_name)
                            elif family_name:
                                family_and_type = family_name
                            elif type_name:
                                family_and_type = type_name
                            
                            # Parametri IFC e Classification di tipo
                            def _tp(name):
                                """Legge un parametro dal tipo, stringa vuota se assente."""
                                try:
                                    p = elem_type.LookupParameter(name)
                                    return _extract_param_value(p) if (p and p.HasValue) else ""
                                except:
                                    return ""
                            
                            ifc_type_export_as  = _tp('Export Type to IFC As')
                            ifc_type_predefined = _tp('Type IFC Predefined Type')
                            uni_pr_num   = _tp('Classification.Uniclass.Pr.Number')
                            uni_pr_desc  = _tp('Classification.Uniclass.Pr.Description')
                            uni_ss_num   = _tp('Classification.Uniclass.Ss.Number')
                            uni_ss_desc  = _tp('Classification.Uniclass.Ss.Description')
                            class_code_type  = _tp('ClassificationCode[Type]')
                            class_code2_type = _tp('ClassificationCode(2)[Type]')
                            class_code3_type = _tp('ClassificationCode(3)[Type]')
                            
                            type_key = "{} : {}".format(processor.file_name, type_id_int)
                            
                            type_data = {
                                'TypeKey': type_key,
                                'TypeID': type_id_int,
                                'FileName': processor.file_name,
                                'FamilyKey': family_key,
                                'TypeName': type_name,
                                'Family&Type': family_and_type,
                                'Description': description,
                                'TypeMark': type_mark,
                                'Export Type to IFC As': ifc_type_export_as,
                                'Type IFC Predefined Type': ifc_type_predefined,
                                'Classification.Uniclass.Pr.Number': uni_pr_num,
                                'Classification.Uniclass.Pr.Description': uni_pr_desc,
                                'Classification.Uniclass.Ss.Number': uni_ss_num,
                                'Classification.Uniclass.Ss.Description': uni_ss_desc,
                                'ClassificationCode[Type]': class_code_type,
                                'ClassificationCode(2)[Type]': class_code2_type,
                                'ClassificationCode(3)[Type]': class_code3_type,
                                'IsUsed': 'YES',
                            }
                            
                            # Parametri custom di tipo
                            for param_name in custom_type_params:
                                try:
                                    tp = elem_type.LookupParameter(param_name)
                                    val = _extract_param_value(tp) if (tp and tp.HasValue) else ""
                                    type_data[param_name] = val
                                    if val:
                                        type_params_not_found.discard(param_name)
                                except:
                                    type_data[param_name] = ""
                            
                            types_dict[type_id_int] = type_data
                        
                        # === ISTANZA ===
                        type_key_ref = "{} : {}".format(processor.file_name, type_id_int)
                        
                        # Workset
                        workset_name = "N/A"
                        workset_key = ""
                        if processor.is_workshared:
                            try:
                                ws_param = elem.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                                if ws_param and ws_param.HasValue:
                                    workset_id = ws_param.AsInteger()
                                    try:
                                        workset = doc.GetWorksetTable().GetWorkset(DB.WorksetId(workset_id))
                                        workset_name = workset.Name
                                        workset_key = "{} : {}".format(processor.file_name, workset_id)
                                    except:
                                        workset_name = "Workset_{}".format(workset_id)
                                        workset_key = "{} : {}".format(processor.file_name, workset_id)
                            except:
                                pass

                        # PhaseCreation
                        phase_creation = ""
                        try:
                            phase_param = elem.get_Parameter(DB.BuiltInParameter.PHASE_CREATED)
                            if phase_param and phase_param.HasValue:
                                phase_id = phase_param.AsElementId()
                                if phase_id and phase_id != ElementId.InvalidElementId:
                                    phase = doc.GetElement(phase_id)
                                    if phase:
                                        phase_creation = phase.Name
                        except:
                            pass
                        
                        # PhaseDemolished
                        phase_demolished = ""
                        try:
                            phase_dem_param = elem.get_Parameter(DB.BuiltInParameter.PHASE_DEMOLISHED)
                            if phase_dem_param and phase_dem_param.HasValue:
                                phase_dem_id = phase_dem_param.AsElementId()
                                if phase_dem_id and phase_dem_id != ElementId.InvalidElementId:
                                    phase_dem = doc.GetElement(phase_dem_id)
                                    if phase_dem:
                                        phase_demolished = phase_dem.Name
                        except:
                            pass
                        
                        # IFC istanza
                        ifc_export_as = ""
                        ifc_predefined = ""
                        try:
                            p = elem.LookupParameter('Export to IFC As')
                            if p and p.HasValue:
                                ifc_export_as = _extract_param_value(p)
                        except:
                            pass
                        try:
                            p = elem.LookupParameter('IFC Predefined Type')
                            if p and p.HasValue:
                                ifc_predefined = _extract_param_value(p)
                        except:
                            pass
                        
                        # ClassificationCode istanza
                        class_code  = ""
                        class_code2 = ""
                        class_code3 = ""
                        try:
                            p = elem.LookupParameter('ClassificationCode')
                            if p and p.HasValue:
                                class_code = _extract_param_value(p)
                        except:
                            pass
                        try:
                            p = elem.LookupParameter('ClassificationCode(2)')
                            if p and p.HasValue:
                                class_code2 = _extract_param_value(p)
                        except:
                            pass
                        try:
                            p = elem.LookupParameter('ClassificationCode(3)')
                            if p and p.HasValue:
                                class_code3 = _extract_param_value(p)
                        except:
                            pass
                        
                        instance_data = {
                            'ElementKey': "{} : {}".format(processor.file_name, elem.Id.IntegerValue),
                            'ElementID': elem.Id.IntegerValue,
                            'FileName': processor.file_name,
                            'TypeKey': type_key_ref,
                            'WorksetKey': workset_key,
                            'WorksetName': workset_name,
                            'PhaseCreation': phase_creation,
                            'PhaseDemolished': phase_demolished,
                            'Export to IFC As': ifc_export_as,
                            'IFC Predefined Type': ifc_predefined,
                            'ClassificationCode': class_code,
                            'ClassificationCode(2)': class_code2,
                            'ClassificationCode(3)': class_code3,
                        }
                        
                        # Parametri custom di istanza
                        for param_name in custom_instance_params:
                            try:
                                ip = elem.LookupParameter(param_name)
                                val = _extract_param_value(ip) if (ip and ip.HasValue) else ""
                                instance_data[param_name] = val
                                if val:
                                    instance_params_not_found.discard(param_name)
                            except:
                                instance_data[param_name] = ""
                        
                        instances.append(instance_data)
                        total_count += 1
                        
                    except Exception as e:
                        continue  # Salta elementi problematici
                
            except Exception as e:
                LOGGER.warning("Errore estrazione categoria {}: {}".format(built_in_cat, str(e)))
        
        # ===== SECONDO PASSAGGIO: raccogliere famiglie e tipi SENZA istanze =====
        # Il loop precedente traccia solo famiglie/tipi con almeno un'istanza posizionata.
        # Questo passaggio aggiunge tutti i tipi caricati a modello (anche senza istanze).
        OUTPUT.print_md("      ⏳ Raccolta tipi e famiglie senza istanze...")
        
        for built_in_cat in target_categories:
            try:
                type_collector = FilteredElementCollector(doc)\
                    .OfCategory(built_in_cat)\
                    .WhereElementIsElementType()
                
                for elem_type in type_collector:
                    try:
                        type_id_int = elem_type.Id.IntegerValue
                        
                        # Se il tipo è già stato tracciato dal loop istanze, salta
                        if type_id_int in types_dict:
                            continue
                        
                        # Categoria
                        category_name = "Unknown"
                        if elem_type.Category:
                            category_name = elem_type.Category.Name
                        
                        # === FAMIGLIA ===
                        family_name = ""
                        family_id = 0
                        is_in_place = "NO"
                        is_system = "YES"
                        
                        # Per FamilySymbol (tipi caricabili)
                        try:
                            if hasattr(elem_type, 'Family') and elem_type.Family:
                                fam = elem_type.Family
                                family_name = fam.Name or ""
                                family_id = fam.Id.IntegerValue
                                is_in_place = "YES" if fam.IsInPlace else "NO"
                                is_system = "NO"
                        except:
                            pass
                        
                        # Fallback: sistema
                        if not family_name:
                            try:
                                if hasattr(elem_type, 'FamilyName') and elem_type.FamilyName:
                                    family_name = elem_type.FamilyName
                            except:
                                pass
                            if not family_name:
                                try:
                                    fam_param = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                                    if fam_param and fam_param.HasValue:
                                        family_name = fam_param.AsString() or ""
                                except:
                                    pass
                            if not family_name and elem_type.Category:
                                family_name = elem_type.Category.Name or ""
                            
                            fam_key_str = "{}|{}".format(category_name, family_name)
                            family_id = -abs(hash(fam_key_str)) % 10000000
                            is_system = "YES"
                            is_in_place = "NO"
                        
                        family_key = "{} : {}".format(processor.file_name, family_id)
                        
                        # Aggiungi famiglia se non già tracciata
                        if family_id not in families_dict:
                            families_dict[family_id] = {
                                'FamilyKey': family_key,
                                'FamilyID': family_id,
                                'FileName': processor.file_name,
                                'FamilyName': family_name,
                                'Category': category_name,
                                'IsInPlace': is_in_place,
                                'IsSystemFamily': is_system,
                                'IsUsed': 'NO',
                            }
                        
                        # === TIPO ===
                        type_name = ""
                        try:
                            if elem_type.Name:
                                type_name = elem_type.Name
                        except:
                            pass
                        if not type_name:
                            try:
                                symbol_param = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                                if symbol_param and symbol_param.HasValue:
                                    type_name = symbol_param.AsString() or ""
                            except:
                                pass
                        if not type_name:
                            try:
                                type_param = elem_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                                if type_param and type_param.HasValue:
                                    type_name = type_param.AsString() or ""
                            except:
                                pass
                        if not type_name:
                            try:
                                type_name = Element.Name.GetValue(elem_type) or ""
                            except:
                                pass
                        
                        # TypeMark
                        type_mark = ""
                        try:
                            mark_param = elem_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK)
                            if mark_param and mark_param.HasValue:
                                type_mark = mark_param.AsString() or ""
                        except:
                            pass
                        
                        # Description
                        description = ""
                        try:
                            desc_param = elem_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_DESCRIPTION)
                            if desc_param and desc_param.HasValue:
                                description = desc_param.AsString() or ""
                        except:
                            pass
                        
                        # Family&Type
                        family_and_type = ""
                        if family_name and type_name:
                            family_and_type = "{}: {}".format(family_name, type_name)
                        elif family_name:
                            family_and_type = family_name
                        elif type_name:
                            family_and_type = type_name
                        
                        # Parametri IFC e Classification di tipo
                        def _tp2(name):
                            """Legge un parametro dal tipo, stringa vuota se assente."""
                            try:
                                p = elem_type.LookupParameter(name)
                                return _extract_param_value(p) if (p and p.HasValue) else ""
                            except:
                                return ""
                        
                        ifc_type_export_as  = _tp2('Export Type to IFC As')
                        ifc_type_predefined = _tp2('Type IFC Predefined Type')
                        uni_pr_num   = _tp2('Classification.Uniclass.Pr.Number')
                        uni_pr_desc  = _tp2('Classification.Uniclass.Pr.Description')
                        uni_ss_num   = _tp2('Classification.Uniclass.Ss.Number')
                        uni_ss_desc  = _tp2('Classification.Uniclass.Ss.Description')
                        class_code_type  = _tp2('ClassificationCode[Type]')
                        class_code2_type = _tp2('ClassificationCode(2)[Type]')
                        class_code3_type = _tp2('ClassificationCode(3)[Type]')
                        
                        type_key = "{} : {}".format(processor.file_name, type_id_int)
                        
                        type_data = {
                            'TypeKey': type_key,
                            'TypeID': type_id_int,
                            'FileName': processor.file_name,
                            'FamilyKey': family_key,
                            'TypeName': type_name,
                            'Family&Type': family_and_type,
                            'Description': description,
                            'TypeMark': type_mark,
                            'Export Type to IFC As': ifc_type_export_as,
                            'Type IFC Predefined Type': ifc_type_predefined,
                            'Classification.Uniclass.Pr.Number': uni_pr_num,
                            'Classification.Uniclass.Pr.Description': uni_pr_desc,
                            'Classification.Uniclass.Ss.Number': uni_ss_num,
                            'Classification.Uniclass.Ss.Description': uni_ss_desc,
                            'ClassificationCode[Type]': class_code_type,
                            'ClassificationCode(2)[Type]': class_code2_type,
                            'ClassificationCode(3)[Type]': class_code3_type,
                            'IsUsed': 'NO',
                        }
                        
                        # Parametri custom di tipo
                        for param_name in custom_type_params:
                            try:
                                tp = elem_type.LookupParameter(param_name)
                                val = _extract_param_value(tp) if (tp and tp.HasValue) else ""
                                type_data[param_name] = val
                                if val:
                                    type_params_not_found.discard(param_name)
                            except:
                                type_data[param_name] = ""
                        
                        types_dict[type_id_int] = type_data
                        
                    except:
                        continue
                        
            except Exception as e:
                LOGGER.warning("Errore raccolta tipi categoria {}: {}".format(built_in_cat, str(e)))
        
        OUTPUT.print_md("      ✓ {} istanze estratte, {} famiglie (totali), {} tipi (totali) da {} categorie".format(
            total_count, len(families_dict), len(types_dict), len(target_categories)))
        
        # Avvisa se alcuni parametri personalizzati non sono stati trovati
        for param_name in instance_params_not_found:
            OUTPUT.print_md("      ⚠️ Parametro istanza '**{}**' non trovato in nessun elemento".format(param_name))
        for param_name in type_params_not_found:
            OUTPUT.print_md("      ⚠️ Parametro tipo '**{}**' non trovato in nessun tipo".format(param_name))
        
    except Exception as e:
        OUTPUT.print_md("      ⚠️ Errore estrazione famiglie/tipi/istanze: {}".format(str(e)))
        LOGGER.error("Errore extract_families_types_instances: {}".format(str(e)))
    
    return list(families_dict.values()), list(types_dict.values()), instances




def extract_parameters(processor):
    """Estrae i parametri definiti nei file .rfa per ogni famiglia caricabile (TAB_Parameters).

    Per ogni famiglia caricabile nel modello, elenca tutti i parametri di famiglia
    (non built-in), indicando se sono Shared o di Famiglia, se sono di Type o Instance,
    e il GUID per i parametri shared.

    Args:
        processor: FileProcessor con il documento aperto

    Returns:
        Lista di dizionari con i dati dei parametri per famiglia
    """
    doc = processor.doc
    params_data = []

    try:
        # Raccogli una istanza per famiglia (per ispezionare i parametri di istanza)
        instance_by_family = {}  # family_id -> FamilyInstance
        try:
            all_instances = DB.FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType().ToElements()
            for inst in all_instances:
                try:
                    fam = inst.Symbol.Family
                    fam_id = fam.Id.IntegerValue
                    if fam_id not in instance_by_family:
                        instance_by_family[fam_id] = inst
                except:
                    continue
        except:
            pass

        # Itera tutte le famiglie caricabili
        all_families = DB.FilteredElementCollector(doc).OfClass(Family).ToElements()

        for fam in all_families:
            try:
                family_id = fam.Id.IntegerValue
                family_key = "{} : {}".format(processor.file_name, family_id)

                type_ids = list(fam.GetFamilySymbolIds())
                if not type_ids:
                    continue

                first_type = doc.GetElement(type_ids[0])
                if first_type is None:
                    continue

                # Recupera istanza di riferimento per questa famiglia
                inst = instance_by_family.get(family_id)
                instance_id = inst.Id.IntegerValue if inst is not None else ""

                # --- Parametri di TIPO (dal FamilySymbol) ---
                type_param_names = set()
                for param in first_type.GetOrderedParameters():
                    try:
                        # Escludi parametri built-in (hanno Id negativo)
                        if param.Id.IntegerValue < 0:
                            continue
                        param_name = param.Definition.Name
                        is_shared = "YES" if param.IsShared else "NO"
                        param_guid = ""
                        if is_shared == "YES":
                            try:
                                param_guid = str(param.GUID)
                            except:
                                pass
                        type_param_names.add(param_name)
                        params_data.append({
                            'FamilyKey': family_key,
                            'FamilyID': family_id,
                            'FileName': processor.file_name,
                            'InstanceID': instance_id,
                            'ParameterName': param_name,
                            'IsShared': is_shared,
                            'TypeOrInstance': 'Type',
                            'Param_GUID': param_guid,
                        })
                    except:
                        continue

                # --- Parametri di ISTANZA (da un FamilyInstance) ---
                if inst is not None:
                    for param in inst.GetOrderedParameters():
                        try:
                            if param.Id.IntegerValue < 0:
                                continue
                            param_name = param.Definition.Name
                            if param_name in type_param_names:
                                continue  # gia' contato come Type
                            is_shared = "YES" if param.IsShared else "NO"
                            param_guid = ""
                            if is_shared == "YES":
                                try:
                                    param_guid = str(param.GUID)
                                except:
                                    pass
                            params_data.append({
                                'FamilyKey': family_key,
                                'FamilyID': family_id,
                                'FileName': processor.file_name,
                                'InstanceID': instance_id,
                                'ParameterName': param_name,
                                'IsShared': is_shared,
                                'TypeOrInstance': 'Instance',
                                'Param_GUID': param_guid,
                            })
                        except:
                            continue

            except:
                continue

        OUTPUT.print_md("      ✓ {} record parametri famiglia estratti".format(len(params_data)))

    except Exception as e:
        OUTPUT.print_md("      ⚠️ Errore estrazione parametri: {}".format(str(e)))
        LOGGER.error("Errore extract_parameters: {}".format(str(e)))

    return params_data


def _collect_graphics_styles(doc, geo_element, subcats):
    """Raccoglie ricorsivamente i nomi delle subcategorie (object style) dalla geometria."""
    for geo_obj in geo_element:
        try:
            if isinstance(geo_obj, DB.GeometryInstance):
                instance_geo = geo_obj.GetInstanceGeometry()
                if instance_geo is not None:
                    _collect_graphics_styles(doc, instance_geo, subcats)
            else:
                gs_id = geo_obj.GraphicsStyleId
                if gs_id is not None and gs_id != DB.ElementId.InvalidElementId:
                    gs = doc.GetElement(gs_id)
                    if gs is not None:
                        try:
                            cat = gs.GraphicsStyleCategory
                            if cat is not None:
                                subcats.add(cat.Name)
                        except:
                            pass
        except:
            continue


def extract_object_styles(processor):
    """Estrae gli object style (subcategorie) per ogni famiglia caricabile (TAB_ObjectStyle).

    Per ogni famiglia, ispeziona la geometria di un'istanza rappresentativa
    e raccoglie tutti gli object style (GraphicsStyleCategory) distinti utilizzati.

    Args:
        processor: FileProcessor con il documento aperto

    Returns:
        Lista di dizionari con i dati degli object style per famiglia
    """
    doc = processor.doc
    styles_data = []

    try:
        # Raccogli una istanza per famiglia
        instance_by_family = {}  # family_id -> FamilyInstance
        try:
            all_instances = DB.FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType().ToElements()
            for inst in all_instances:
                try:
                    fam = inst.Symbol.Family
                    fam_id = fam.Id.IntegerValue
                    if fam_id not in instance_by_family:
                        instance_by_family[fam_id] = inst
                except:
                    continue
        except:
            pass

        # Opzioni geometria
        geo_options = DB.Options()
        geo_options.DetailLevel = DB.ViewDetailLevel.Fine

        all_families = DB.FilteredElementCollector(doc).OfClass(Family).ToElements()

        for fam in all_families:
            try:
                family_id = fam.Id.IntegerValue
                family_key = "{} : {}".format(processor.file_name, family_id)

                inst = instance_by_family.get(family_id)
                if inst is None:
                    continue

                # Raccogli subcategorie dalla geometria dell'istanza
                subcats = set()
                try:
                    geo = inst.get_Geometry(geo_options)
                    if geo is not None:
                        _collect_graphics_styles(doc, geo, subcats)
                except:
                    pass

                for subcat_name in sorted(subcats):
                    styles_data.append({
                        'FileName': processor.file_name,
                        'FamilyKey': family_key,
                        'ObjectStyle': subcat_name,
                    })

            except:
                continue

        OUTPUT.print_md("      ✓ {} record object style estratti".format(len(styles_data)))

    except Exception as e:
        OUTPUT.print_md("      ⚠️ Errore estrazione object style: {}".format(str(e)))
        LOGGER.error("Errore extract_object_styles: {}".format(str(e)))

    return styles_data


def extract_purgeable_elements(processor):
    """Estrae elementi purgabili con dettaglio per elemento (TAB_PurgeableElements)."""
    doc = processor.doc
    purgeable = []
    
    try:
        # ===== FAMIGLIE NON USATE =====
        unused_families = _get_unused_families(doc)
        for elem_id, elem_name, revit_category in unused_families:
            purgeable.append({
                'PurgeableElementKey': "{} : {}".format(processor.file_name, elem_id),
                'PurgeableElementID': elem_id,
                'FileName': processor.file_name,
                'Category': 'Families',
                'PurgeableElementName': elem_name,
                'RevitCategory': revit_category
            })
        
        # ===== TIPI NON USATI =====
        unused_types = _get_unused_types(doc)
        for elem_id, elem_name, revit_category in unused_types:
            purgeable.append({
                'FileName': processor.file_name,
                'PurgeableElementKey': "{} : {}".format(processor.file_name, elem_id),
                'PurgeableElementID': elem_id,
                'Category': 'Types',
                'PurgeableElementName': elem_name,
                'RevitCategory': revit_category
            })
        
        # ===== MATERIALI NON USATI =====
        unused_materials = _get_unused_materials(doc)
        for elem_id, elem_name in unused_materials:
            purgeable.append({
                'FileName': processor.file_name,
                'PurgeableElementKey': "{} : {}".format(processor.file_name, elem_id),
                'PurgeableElementID': elem_id,
                'Category': 'Materials',
                'PurgeableElementName': elem_name,
                'RevitCategory': ''
            })
        
        # ===== VIEW TEMPLATES NON USATI =====
        unused_templates = _get_unused_view_templates(doc)
        for elem_id, elem_name in unused_templates:
            purgeable.append({
                'FileName': processor.file_name,
                'PurgeableElementKey': "{} : {}".format(processor.file_name, elem_id),
                'PurgeableElementID': elem_id,
                'Category': 'ViewTemplates',
                'PurgeableElementName': elem_name,
                'RevitCategory': ''
            })
        
        # ===== FILTRI NON USATI =====
        unused_filters = _get_unused_filters(doc)
        for elem_id, elem_name in unused_filters:
            purgeable.append({
                'FileName': processor.file_name,
                'PurgeableElementKey': "{} : {}".format(processor.file_name, elem_id),
                'PurgeableElementID': elem_id,
                'Category': 'Filters',
                'PurgeableElementName': elem_name,
                'RevitCategory': ''
            })
        
        # ===== MODEL GROUPS NON USATI =====
        unused_groups = _get_unused_model_groups(doc)
        for elem_id, elem_name in unused_groups:
            purgeable.append({
                'FileName': processor.file_name,
                'PurgeableElementKey': "{} : {}".format(processor.file_name, elem_id),
                'PurgeableElementID': elem_id,
                'Category': 'ModelGroups',
                'PurgeableElementName': elem_name,
                'RevitCategory': ''
            })
            
    except Exception as e:
        LOGGER.warning("Errore calcolo elementi purgabili: {}".format(str(e)))
    
    return purgeable


def _get_unused_families(doc):
    """Restituisce lista di tuple (family_id, family_name, revit_category) per famiglie caricabili non usate.
    Esclude le famiglie di sistema e le famiglie in-place."""
    unused = []
    try:
        # Raccoglie TUTTI i TypeId usati da TUTTI gli elementi nel modello
        used_type_ids = set()
        all_instances = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        for inst in all_instances:
            try:
                type_id = inst.GetTypeId()
                if type_id and type_id != ElementId.InvalidElementId:
                    used_type_ids.add(type_id.IntegerValue)
            except:
                pass
        
        # Verifica le famiglie caricabili (Family class)
        families = FilteredElementCollector(doc).OfClass(Family).ToElements()
        
        for family in families:
            try:
                # Escludi famiglie in-place (non sono vere famiglie caricabili purgabili)
                try:
                    if family.IsInPlace:
                        continue
                except:
                    pass
                
                # Verifica che sia una famiglia caricabile (IsEditable = può essere modificata)
                # Le famiglie di sistema non hanno la proprietà IsEditable o è False
                try:
                    if not family.IsEditable:
                        continue
                except:
                    pass
                
                family_symbol_ids = family.GetFamilySymbolIds()
                
                # Se la famiglia non ha tipi, non è purgabile in questo modo
                if not family_symbol_ids or len(list(family_symbol_ids)) == 0:
                    continue
                
                has_instances = False
                
                # Verifica se almeno un tipo della famiglia è usato
                for symbol_id in family_symbol_ids:
                    if symbol_id.IntegerValue in used_type_ids:
                        has_instances = True
                        break
                
                if not has_instances:
                    family_name = family.Name if family.Name else ""
                    
                    # Ottieni la categoria Revit della famiglia
                    revit_category = ""
                    try:
                        fam_cat = family.FamilyCategory
                        if fam_cat:
                            revit_category = fam_cat.Name if fam_cat.Name else ""
                    except:
                        pass
                    
                    unused.append((family.Id.IntegerValue, family_name, revit_category))
            except:
                pass
    except:
        pass
    
    return unused


def _get_unused_types(doc):
    """Restituisce lista di tuple (type_id, family_and_type_name) per tipi non usati.
    Il nome è nel formato 'FamilyName : TypeName'.
    Include solo: FamilySymbol + tipi di sistema specifici (Floor, Wall, Ceiling, Duct, Pipe, etc.)"""
    unused = []
    
    # Import dei tipi di sistema specifici
    try:
        from Autodesk.Revit.DB import (
            FamilySymbol,
            WallType,
            FloorType,
            CeilingType
        )
    except:
        pass
    
    # Import tipi MEP (potrebbero non esistere in tutte le versioni)
    try:
        from Autodesk.Revit.DB.Mechanical import DuctType, FlexDuctType
    except:
        DuctType = None
        FlexDuctType = None
    
    try:
        from Autodesk.Revit.DB.Plumbing import PipeType, FlexPipeType
    except:
        PipeType = None
        FlexPipeType = None
    
    try:
        from Autodesk.Revit.DB.Electrical import CableTrayType, ConduitType
    except:
        CableTrayType = None
        ConduitType = None
    
    # Lista di tipi ammessi
    allowed_types = [FamilySymbol, WallType, FloorType, CeilingType]
    if DuctType: allowed_types.append(DuctType)
    if FlexDuctType: allowed_types.append(FlexDuctType)
    if PipeType: allowed_types.append(PipeType)
    if FlexPipeType: allowed_types.append(FlexPipeType)
    if CableTrayType: allowed_types.append(CableTrayType)
    if ConduitType: allowed_types.append(ConduitType)
    
    try:
        types = FilteredElementCollector(doc).WhereElementIsElementType().ToElements()
        
        # Set di tipi usati
        used_type_ids = set()
        instances = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        
        for inst in instances:
            try:
                type_id = inst.GetTypeId()
                if type_id and type_id != ElementId.InvalidElementId:
                    used_type_ids.add(type_id.IntegerValue)
            except:
                pass
        
        for t in types:
            try:
                # Verifica se il tipo è tra quelli ammessi
                is_allowed = False
                for allowed_type in allowed_types:
                    if allowed_type and isinstance(t, allowed_type):
                        is_allowed = True
                        break
                
                if not is_allowed:
                    continue
                
                if t.Id.IntegerValue not in used_type_ids:
                    # Costruisci "Family : Type" name
                    type_name = ""
                    
                    # Metodo PRIMARIO: Usa il parametro SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM
                    # Questo parametro contiene direttamente "FamilyName : TypeName"
                    try:
                        fam_type_param = t.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM)
                        if fam_type_param and fam_type_param.HasValue:
                            type_name = fam_type_param.AsString() or ""
                    except:
                        pass
                    
                    # Fallback: costruisci manualmente
                    if not type_name:
                        family_name = ""
                        elem_type_name = ""
                        
                        # Ottieni il nome del tipo
                        try:
                            elem_type_name = t.Name if t.Name else ""
                        except:
                            elem_type_name = ""
                        
                        # Metodo 1: Se è un FamilySymbol, usa Family.Name
                        try:
                            if isinstance(t, FamilySymbol):
                                family = t.Family
                                if family:
                                    family_name = family.Name if family.Name else ""
                        except:
                            pass
                        
                        # Metodo 2: Parametro SYMBOL_FAMILY_NAME_PARAM
                        if not family_name:
                            try:
                                fam_param = t.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                                if fam_param and fam_param.HasValue:
                                    family_name = fam_param.AsString() or ""
                            except:
                                pass
                        
                        # Metodo 3: Parametro ALL_MODEL_FAMILY_NAME
                        if not family_name:
                            try:
                                fam_param = t.get_Parameter(DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
                                if fam_param and fam_param.HasValue:
                                    family_name = fam_param.AsString() or ""
                            except:
                                pass
                        
                        # Metodo 4: FamilyName property
                        if not family_name:
                            try:
                                if hasattr(t, 'FamilyName'):
                                    family_name = t.FamilyName if t.FamilyName else ""
                            except:
                                pass
                        
                        # Metodo 5: Nome della categoria come fallback per System Families
                        if not family_name:
                            try:
                                cat = t.Category
                                if cat:
                                    family_name = cat.Name if cat.Name else ""
                            except:
                                pass
                        
                        # Costruisci il nome finale
                        if family_name and elem_type_name:
                            type_name = "{} : {}".format(family_name, elem_type_name)
                        elif elem_type_name:
                            type_name = elem_type_name
                        elif family_name:
                            type_name = family_name
                    
                    # Ottieni la categoria Revit del tipo
                    revit_category = ""
                    try:
                        cat = t.Category
                        if cat:
                            revit_category = cat.Name if cat.Name else ""
                    except:
                        pass
                    
                    unused.append((t.Id.IntegerValue, type_name, revit_category))
            except:
                pass
                
    except:
        pass
    
    return unused


def _get_unused_materials(doc):
    """Restituisce lista di tuple (material_id, material_name) per materiali non usati."""
    unused = []
    try:
        materials = FilteredElementCollector(doc).OfClass(Material).ToElements()
        
        # Set di materiali usati
        used_material_ids = set()
        
        # Verifica su elementi
        all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        
        for elem in all_elements:
            try:
                mat_ids = elem.GetMaterialIds(False)
                for mid in mat_ids:
                    used_material_ids.add(mid.IntegerValue)
            except:
                pass
        
        for mat in materials:
            try:
                if mat.Id.IntegerValue not in used_material_ids:
                    mat_name = mat.Name if mat.Name else ""
                    unused.append((mat.Id.IntegerValue, mat_name))
            except:
                pass
                
    except:
        pass
    
    return unused


def _get_unused_view_templates(doc):
    """Restituisce lista di tuple (template_id, template_name) per view template non usati."""
    unused = []
    try:
        all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
        
        used_template_ids = set()
        templates = {}  # id -> view object
        
        for view in all_views:
            try:
                if view.IsTemplate:
                    templates[view.Id.IntegerValue] = view
                else:
                    template_id = view.ViewTemplateId
                    if template_id and template_id != ElementId.InvalidElementId:
                        used_template_ids.add(template_id.IntegerValue)
            except:
                pass
        
        # Trova template non usati
        for template_id, template_view in templates.items():
            if template_id not in used_template_ids:
                template_name = template_view.Name if template_view.Name else ""
                unused.append((template_id, template_name))
        
    except:
        pass
    
    return unused


def _get_unused_filters(doc):
    """Restituisce lista di tuple (filter_id, filter_name) per filtri non usati."""
    unused = []
    try:
        all_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
        
        used_filter_ids = set()
        all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
        
        for view in all_views:
            try:
                filter_ids = view.GetFilters()
                for fid in filter_ids:
                    used_filter_ids.add(fid.IntegerValue)
            except:
                pass
        
        for flt in all_filters:
            try:
                if flt.Id.IntegerValue not in used_filter_ids:
                    filter_name = flt.Name if flt.Name else ""
                    unused.append((flt.Id.IntegerValue, filter_name))
            except:
                pass
                
    except:
        pass
    
    return unused


def _get_unused_model_groups(doc):
    """Restituisce lista di tuple (group_type_id, group_name) per Model Groups non usati."""
    unused = []
    try:
        # Raccogli tutti i GroupType
        all_group_types = FilteredElementCollector(doc).OfClass(GroupType).ToElements()
        
        # Raccogli tutti gli ID dei GroupType che hanno istanze
        used_group_type_ids = set()
        all_groups = FilteredElementCollector(doc).OfClass(Group).ToElements()
        
        for grp in all_groups:
            try:
                type_id = grp.GetTypeId()
                if type_id and type_id != ElementId.InvalidElementId:
                    used_group_type_ids.add(type_id.IntegerValue)
            except:
                pass
        
        # Filtra solo i Model Groups (non Detail Groups)
        for gt in all_group_types:
            try:
                # Verifica se è un Model Group (non Detail Group)
                # I Model Groups hanno categoria OST_IOSModelGroups
                # I Detail Groups hanno categoria OST_IOSDetailGroups
                is_model_group = False
                try:
                    cat = gt.Category
                    if cat:
                        # OST_IOSModelGroups = BuiltInCategory per Model Groups
                        if cat.Id.IntegerValue == int(BuiltInCategory.OST_IOSModelGroups):
                            is_model_group = True
                except:
                    # Fallback: se non riesci a determinare, includi comunque
                    is_model_group = True
                
                if is_model_group and gt.Id.IntegerValue not in used_group_type_ids:
                    group_name = gt.Name if gt.Name else ""
                    unused.append((gt.Id.IntegerValue, group_name))
            except:
                pass
                
    except:
        pass
    
    return unused


# ==============================================================================
# SUMMARY DASHBOARD (modulo esterno)
# ==============================================================================

from summary_dashboard import show_summary_dashboard


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    """Funzione principale - Estrazione completa di tutte le tabelle."""
    
    OUTPUT.print_md("# 🏗️ BIM Data Extractor")
    OUTPUT.print_md("**Modalità**: Estrazione completa di tutte le tabelle")
    OUTPUT.print_md("---")
    
    # 1. Mostra la form XAML unificata
    form_result = show_model_report_form()
    
    if not form_result:
        OUTPUT.print_md("❌ Operazione annullata dall'utente.")
        return
    
    rvt_files, output_folder, custom_params, validation_rules, discipline_rules = form_result
    
    OUTPUT.print_md("## 📁 Selezione File")
    OUTPUT.print_md("✅ Selezionati **{}** file".format(len(rvt_files)))
    
    OUTPUT.print_md("## 📂 Cartella Destinazione")
    OUTPUT.print_md("✅ Destinazione: **{}**".format(output_folder))
    
    OUTPUT.print_md("## 🏷️ Classificazione Disciplina")
    for code, desc in discipline_rules:
        OUTPUT.print_md("   **{}** → {}".format(code, desc))
    
    OUTPUT.print_md("## 🔧 Parametri Aggiuntivi")
    if custom_params:
        OUTPUT.print_md("✅ Parametri aggiuntivi richiesti: **{}**".format(", ".join(custom_params)))
    else:
        OUTPUT.print_md("ℹ️ Nessun parametro aggiuntivo specificato.")
    
    # Mostra regole di validazione utente (se presenti)
    if validation_rules:
        OUTPUT.print_md("## ✅ Validazione Dati (valori ammessi)")
        for param_name, rule in validation_rules.items():
            values = rule.get('allowed_values', [])
            OUTPUT.print_md("   **{}** → {} valori ammessi".format(param_name, len(values)))

    # Salva il JSON di setup (tutto)
    _save_json_setup(output_folder, custom_params, validation_rules, discipline_rules)

    # Aggiorna la color legend nella cartella radice
    OUTPUT.print_md("## 🎨 Color Legend")
    _update_color_legend_csv(output_folder, rvt_files)
    
    # Carica il dizionario severity dei warnings
    OUTPUT.print_md("## ⚠️ Classificazione Warnings")
    severity_lookup = _load_warnings_severity_csv()
    if severity_lookup:
        OUTPUT.print_md("✅ WarningsSeverity.csv caricato: **{}** regole di classificazione".format(len(severity_lookup)))
    else:
        OUTPUT.print_md("⚠️ WarningsSeverity.csv non trovato o vuoto. La colonna WarningSeverity conterrà '00_Unknown'.")
    
    # 2. Prepara struttura cartelle (CurrentData / _Old) e gestisce storicizzazione
    OUTPUT.print_md("## 🗂️ Gestione Archivio")
    current_folder = _prepare_output_folders(output_folder)
    
    # 3. Inizializza writer CSV (scrive in CurrentData)
    csv_writer = CSVWriter(current_folder)

    # 3.1 Pre-check: verifica che nessun CSV di output sia aperto
    locked_files = []
    for table_name in CSVWriter.TABLE_HEADERS.keys():
        csv_path = os.path.join(current_folder, "{}.csv".format(table_name))
        if os.path.isfile(csv_path):
            try:
                with io.open(csv_path, 'a', encoding=CSV_ENCODING):
                    pass
            except IOError:
                locked_files.append("{}.csv".format(table_name))
    # Controlla anche TAB_Snapshot_Summary nella cartella radice
    snapshot_path = os.path.join(output_folder, "TAB_Snapshot_Summary.csv")
    if os.path.isfile(snapshot_path):
        try:
            with io.open(snapshot_path, 'a', encoding=CSV_ENCODING):
                pass
        except IOError:
            locked_files.append("TAB_Snapshot_Summary.csv")

    if locked_files:
        OUTPUT.print_md("")
        OUTPUT.print_md("# <span style='color:red; font-size:24px;'>⚠️ ATTENZIONE - FILE BLOCCATI ⚠️</span>")
        OUTPUT.print_md("")
        OUTPUT.print_md("## <span style='color:red'>I seguenti file CSV sono aperti in un altro programma:</span>")
        OUTPUT.print_md("")
        for lf in locked_files:
            OUTPUT.print_md("- <span style='color:red; font-weight:bold;'>{}</span>".format(lf))
        OUTPUT.print_md("")
        OUTPUT.print_md("## <span style='color:red'>📋 COSA FARE:</span>")
        OUTPUT.print_md("1. <span style='color:red; font-weight:bold;'>Chiudi i file elencati sopra</span> (probabilmente aperti in Excel)")
        OUTPUT.print_md("2. <span style='color:red; font-weight:bold;'>Rilancia lo script</span>")
        OUTPUT.print_md("")
        OUTPUT.print_md("<span style='color:red; font-size:18px; font-weight:bold;'>❌ Estrazione annullata.</span>")
        return

    # 4. Processa i file
    OUTPUT.print_md("## ⚙️ Elaborazione")
    OUTPUT.print_md("---")
    
    app = HOST_APP.app
    
    # Classificazione parametri custom (verrà fatta al primo file)
    custom_instance_params = []
    custom_type_params = []
    params_classified = False
    
    # Progress bar
    total_files = len(rvt_files)
    processed = 0
    errors = 0
    
    with forms.ProgressBar(title="Elaborazione file Revit...", 
                           cancellable=True) as pb:
        
        for i, file_path in enumerate(rvt_files):
            if pb.cancelled:
                OUTPUT.print_md("⚠️ Operazione annullata dall'utente.")
                break
            
            pb.update_progress(i, total_files)
            
            file_name = os.path.basename(file_path)
            OUTPUT.print_md("### 📄 Elaborazione: **{}** ({}/{})".format(
                file_name, i+1, total_files))
            
            # Processa il file
            processor = FileProcessor(file_path, app)
            
            if processor.open_document():
                try:
                    # ===== CLASSIFICAZIONE PARAMETRI CUSTOM (solo al primo file) =====
                    if not params_classified and custom_params:
                        OUTPUT.print_md("   ⏳ Classificazione parametri custom...")
                        inst_p, type_p, invalid_p = _classify_custom_params(
                            processor.doc, custom_params)
                        custom_instance_params = inst_p
                        custom_type_params = type_p
                        
                        # Segnala parametri non validi (non sono parametri di progetto)
                        for p in invalid_p:
                            OUTPUT.print_md("   ❌ **ERRORE**: Il parametro '**{}**' non è un parametro di progetto. Verrà ignorato.".format(p))
                        
                        if custom_instance_params:
                            OUTPUT.print_md("   📋 Parametri di istanza: **{}**".format(
                                ", ".join(custom_instance_params)))
                        if custom_type_params:
                            OUTPUT.print_md("   📋 Parametri di tipo: **{}**".format(
                                ", ".join(custom_type_params)))
                        
                        params_classified = True
                    
                    # ===== ESTRAZIONE DATI (prima raccogliamo tutto, poi calcoliamo il riepilogo) =====
                    OUTPUT.print_md("   ⏳ Estrazione informazioni file...")
                    file_info = extract_file_info(processor)
                    # Compila FileDiscipline
                    file_info['FileDiscipline'] = _resolve_discipline(
                        processor.file_name, discipline_rules)
                    
                    OUTPUT.print_md("   ⏳ Estrazione links...")
                    links_data = extract_links(processor)
                    # Compila LinkDiscipline per ogni link
                    for link_row in links_data:
                        link_name = link_row.get('LinkName', '')
                        link_row['LinkDiscipline'] = _resolve_discipline(
                            link_name, discipline_rules)
                    
                    OUTPUT.print_md("   ⏳ Estrazione viste...")
                    views_data = extract_views(processor)
                    
                    OUTPUT.print_md("   ⏳ Estrazione warnings...")
                    warnings_data = extract_warnings(processor, severity_lookup)
                    
                    OUTPUT.print_md("   ⏳ Estrazione worksets...")
                    worksets_data = extract_worksets(processor)
                    
                    OUTPUT.print_md("   ⏳ Estrazione tavole...")
                    sheets_data = extract_sheets(processor)
                    
                    OUTPUT.print_md("   ⏳ Estrazione view templates...")
                    templates_data = extract_view_templates(processor)
                    
                    OUTPUT.print_md("   ⏳ Estrazione materiali...")
                    materials_data = extract_materials(processor)
                    
                    OUTPUT.print_md("   ⏳ Estrazione livelli...")
                    levels_data = extract_levels(processor)
                    
                    OUTPUT.print_md("   ⏳ Estrazione scope boxes...")
                    scope_boxes_data = extract_scope_boxes(processor)
                    
                    OUTPUT.print_md("   ⏳ Estrazione griglie...")
                    grids_data = extract_grids(processor)
                    
                    OUTPUT.print_md("   ⏳ Estrazione filtri...")
                    filters_data = extract_filters(processor)
                    
                    OUTPUT.print_md("   ⏳ Estrazione stanze...")
                    rooms_data = extract_rooms(processor)
                    
                    OUTPUT.print_md("   ⏳ Estrazione vani...")
                    spaces_data = extract_spaces(processor)
                    
                    OUTPUT.print_md("   ⏳ Estrazione aree...")
                    areas_data = extract_areas(processor)
                    
                    OUTPUT.print_md("   ⏳ Estrazione tag...")
                    tags_data = extract_tags(processor)
                    
                    OUTPUT.print_md("   ⏳ Estrazione famiglie, tipi e istanze...")
                    families_data, types_data, instances_data = extract_families_types_instances(
                        processor, custom_instance_params, custom_type_params)
                    
                    OUTPUT.print_md("   ⏳ Estrazione parametri per famiglia...")
                    parameters_data = extract_parameters(processor)

                    OUTPUT.print_md("   ⏳ Estrazione object style per famiglia...")
                    object_styles_data = extract_object_styles(processor)

                    OUTPUT.print_md("   ⏳ Analisi elementi purgabili...")
                    purgeable_data = extract_purgeable_elements(processor)
                    
                    # ===== CONTEGGIO ISTANZE MODEL IN PLACE =====
                    # Conta le singole istanze appartenenti a famiglie in-place (non le famiglie uniche)
                    model_in_place_count = 0
                    try:
                        all_instances = FilteredElementCollector(processor.doc).OfClass(FamilyInstance).ToElements()
                        for inst in all_instances:
                            try:
                                fam = inst.Symbol.Family
                                if fam and fam.IsInPlace:
                                    model_in_place_count += 1
                            except:
                                pass
                    except Exception as e:
                        LOGGER.warning("Errore conteggio Model In Place: {}".format(str(e)))
                    
                    # ===== CALCOLO RIEPILOGO FILE =====
                    OUTPUT.print_md("   ⏳ Calcolo riepilogo file...")
                    file_summary = compute_file_summary(
                        file_info, links_data, views_data, warnings_data,
                        sheets_data, templates_data, levels_data, grids_data,
                        filters_data, rooms_data, spaces_data, areas_data,
                        tags_data, families_data, types_data, instances_data,
                        purgeable_data, model_in_place_count
                    )
                    file_info.update(file_summary)
                    
                    # ===== CALCOLO HEALTH CHECKS =====
                    health_checks_data = compute_health_checks(file_info)
                    file_info['KPI_HealthScore'] = sum(row.get('Score', 0) for row in health_checks_data)
                    
                    # ===== SCRITTURA DATI NEL CSV =====
                    csv_writer.add_row('TAB_Files', file_info)
                    csv_writer.add_rows('TAB_Links', links_data)
                    csv_writer.add_rows('TAB_Views', views_data)
                    csv_writer.add_rows('TAB_Warnings', warnings_data)
                    csv_writer.add_rows('TAB_Worksets_UserDefined', worksets_data)
                    csv_writer.add_rows('TAB_Sheets', sheets_data)
                    csv_writer.add_rows('TAB_ViewTemplates', templates_data)
                    csv_writer.add_rows('TAB_Materials', materials_data)
                    csv_writer.add_rows('TAB_Levels', levels_data)
                    csv_writer.add_rows('TAB_ScopeBoxes', scope_boxes_data)
                    csv_writer.add_rows('TAB_Grids', grids_data)
                    csv_writer.add_rows('TAB_Filters', filters_data)
                    csv_writer.add_rows('TAB_Rooms', rooms_data)
                    csv_writer.add_rows('TAB_Spaces', spaces_data)
                    csv_writer.add_rows('TAB_Areas', areas_data)
                    csv_writer.add_rows('TAB_Tags', tags_data)
                    csv_writer.add_rows('TAB_Families', families_data)
                    csv_writer.add_rows('TAB_Types', types_data)
                    csv_writer.add_rows('TAB_Instances', instances_data)
                    csv_writer.add_rows('TAB_PurgeableElements', purgeable_data)
                    csv_writer.add_rows('TAB_Parameters', parameters_data)
                    csv_writer.add_rows('TAB_ObjectStyle', object_styles_data)
                    csv_writer.add_rows('TAB_HealthChecks', health_checks_data)
                    
                    # ===== VALIDAZIONE DATI =====
                    OUTPUT.print_md("   ⏳ Validazione naming convention...")
                    
                    # Validazione regex famiglie (sempre eseguita)
                    val_fam = validate_families_data(families_data)
                    csv_writer.add_rows('TAB_DataValidation_Families', val_fam)
                    OUTPUT.print_md("      ✓ {} record validazione famiglie".format(len(val_fam)))
                    
                    # Validazione regex tipi + valori ammessi custom tipo (sempre eseguita)
                    val_type = validate_types_data(types_data, custom_type_params, validation_rules)
                    csv_writer.add_rows('TAB_DataValidation_Types', val_type)
                    OUTPUT.print_md("      ✓ {} record validazione tipi".format(len(val_type)))
                    
                    # Validazione valori ammessi custom istanza (solo se ci sono regole)
                    val_inst = validate_instances_data(instances_data, custom_instance_params, validation_rules)
                    csv_writer.add_rows('TAB_DataValidation_Instances', val_inst)
                    if val_inst:
                        OUTPUT.print_md("      ✓ {} record validazione istanze".format(len(val_inst)))
                    
                    OUTPUT.print_md("   ✅ **{}** elaborato con successo".format(file_name))
                    processed += 1
                    
                except Exception as e:
                    import traceback
                    OUTPUT.print_md("   ❌ **ERRORE** elaborazione **{}**:".format(file_name))
                    OUTPUT.print_md("   ```")
                    OUTPUT.print_md("   {}".format(str(e)))
                    OUTPUT.print_md("   {}".format(traceback.format_exc()))
                    OUTPUT.print_md("   ```")
                    errors += 1
                
                finally:
                    processor.close_document()
            else:
                errors += 1
    
    # 5. Scrivi i CSV
    OUTPUT.print_md("## 💾 Salvataggio CSV")
    OUTPUT.print_md("---")
    
    csv_writer.write_all(
        custom_instance_params=custom_instance_params if custom_instance_params else None,
        custom_type_params=custom_type_params if custom_type_params else None
    )
    
    # 5.1 Aggiorna TAB_Snapshot_Summary nella cartella radice
    OUTPUT.print_md("## 📊 Aggiornamento Snapshot Storico")
    _append_to_snapshot_summary(
        output_folder,
        csv_writer.data.get('TAB_Files', []),
        CSVWriter.TABLE_HEADERS['TAB_Files']
    )
    
    # 5.1 Controlla se ci sono file bloccati
    if csv_writer.blocked_files:
        OUTPUT.print_md("")
        OUTPUT.print_md("---")
        OUTPUT.print_md("# <span style='color:red; font-size:24px;'>⚠️ ATTENZIONE - FILE BLOCCATI ⚠️</span>")
        OUTPUT.print_md("")
        OUTPUT.print_md("## <span style='color:red'>I seguenti file CSV NON sono stati scritti perché aperti in un altro programma:</span>")
        OUTPUT.print_md("")
        for blocked_file in csv_writer.blocked_files:
            OUTPUT.print_md("- <span style='color:red; font-weight:bold;'>{}</span>".format(blocked_file))
        OUTPUT.print_md("")
        OUTPUT.print_md("## <span style='color:red'>📋 COSA FARE:</span>")
        OUTPUT.print_md("1. <span style='color:red; font-weight:bold;'>Chiudi i file elencati sopra</span> (probabilmente aperti in Excel)")
        OUTPUT.print_md("2. <span style='color:red; font-weight:bold;'>Rilancia questo script</span>")
        OUTPUT.print_md("")
        OUTPUT.print_md("---")
    
    # 6. Riepilogo finale
    OUTPUT.print_md("## 📊 Riepilogo")
    OUTPUT.print_md("---")
    OUTPUT.print_md("- **File processati**: {}".format(processed))
    OUTPUT.print_md("- **Errori/Saltati**: {}".format(errors))
    OUTPUT.print_md("- **Cartella output**: {}".format(output_folder))
    OUTPUT.print_md("- **Data estrazione**: {}".format(EXTRACTION_DATE))
    OUTPUT.print_md("")
    
    if csv_writer.blocked_files:
        OUTPUT.print_md("<span style='color:red; font-size:18px; font-weight:bold;'>⚠️ OPERAZIONE COMPLETATA CON AVVISI - Alcuni file non sono stati scritti!</span>")
    else:
        OUTPUT.print_md("✅ **Operazione completata!**")

    # 7. Mostra dashboard di riepilogo
    if processed > 0:
        OUTPUT.print_md("## 📊 Dashboard")
        show_summary_dashboard(csv_writer.data)


# ==============================================================================
# ENTRY POINT
# ==============================================================================

if __name__ == "__main__":
    main()
