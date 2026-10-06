# -*- coding: utf-8 -*-
"""
mepqto_ui.py - finestra XAML del computo MEP.

Cinque schede: elenco prezzi (modificabile: capitolo, sottocapitolo, n. articolo
EPU, prezzario, descrizione, unita' da tendina, prezzo unitario), computo
compilato in automatico (si modifica solo l'override della maggiorazione delle
voci), riepilogo Type Mark con i codici di tipo e d'istanza 1..10 e le loro
descrizioni, regole di misura (modificabili), anomalie. La finestra resta aperta mentre
l'utente completa l'elenco prezzi: Save scrive il file di progetto senza
chiuderla, Export Excel produce il computo in qualsiasi momento. Il modello
Revit viene solo letto.

Le griglie sono legate a System.Data.DataTable: il binding bidirezionale e'
nativo .NET, quindi nessun oggetto IronPython da notificare. Il prezzo
modificabile e' una colonna stringa convertita a mano in ColumnChanging: il
binding WPF usa la cultura en-US e leggerebbe "12,5" come 125.
"""

import os
from collections import OrderedDict
from datetime import datetime

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Data')

import System
from System import Action, DBNull
from System.Data import DataTable, DataRowState
from System.Diagnostics import Process
from System.IO import FileStream, FileMode
from System.Windows import (Window, MessageBox, MessageBoxButton, MessageBoxImage,
                            MessageBoxResult, Visibility)
from System.Windows.Controls import (CheckBox, ContextMenu, DataGridEditingUnit,
                                     DataGridLength, DataGridTextColumn, MenuItem)
from System.Windows.Data import Binding
from System.Windows.Input import Cursors
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Colors
from System.Windows.Threading import DispatcherPriority
from Microsoft.Win32 import OpenFileDialog, SaveFileDialog

from pyrevit import DB, script

import mepqto_model as qm
import mepqto_rules as qr
import mepqto_store as qs
import mepqto_xlsx as qx
from mepqto_wbs_ui import show_wbs_dialog
from mepqto_params_ui import show_parameters_dialog
from mepqto_pricelist_ui import show_pricelist_editor
from mepqto_allowance_ui import show_allowance_dialog

XAML_FILE_NAME = 'MEPQTO_form.xaml'
CONFIG_SECTION = 'ESA_MEPQTO'
TITLE = "MEP Quantity Takeoff"
EURO = u"\u20ac"
CONFIG_SEPARATOR = u"::"

BLACK_BRUSH = SolidColorBrush(Colors.Black)
BLACK_BRUSH.Freeze()
GRAY_BRUSH = SolidColorBrush(Colors.Gray)
GRAY_BRUSH.Freeze()
RED_BRUSH = SolidColorBrush(Colors.Firebrick)
RED_BRUSH.Freeze()

# Colonna della griglia elenco prezzi -> campo dell'archivio
PRICE_FIELDS = {"Chapter": "chapter", "Subchapter": "subchapter",
                "EpuItem": "epu_item", "PriceBook": "price_book",
                "Description": "description", "Unit": "unit", "UnitPrice": "price"}

CLR_STRING = clr.GetClrType(System.String)
CLR_DOUBLE = clr.GetClrType(System.Double)
CLR_INT = clr.GetClrType(System.Int32)
CLR_BOOL = clr.GetClrType(System.Boolean)


# =============================================================================
# UTILITY
# =============================================================================

def cell_text(value):
    if value is None or isinstance(value, DBNull):
        return u""
    return u"{}".format(value).strip()


def escape_like(text):
    """Testo letterale dentro un LIKE di DataView.RowFilter."""
    out = []
    for ch in text:
        if ch in u"[]*%":
            out.append(u"[{}]".format(ch))
        elif ch == u"'":
            out.append(u"''")
        else:
            out.append(ch)
    return u"".join(out)


def model_name(doc):
    """Nome del modello: quello del centrale per i modelli condivisi."""
    key = qs.document_key(doc)
    if key and "://" not in key:
        return os.path.splitext(os.path.basename(key))[0]
    return doc.Title


def price_sort_key(item):
    """Elenco prezzi ordinato per capitolo, sottocapitolo e codice; i vuoti in fondo."""
    chapter = (item.chapter or u"").lower()
    subchapter = (item.subchapter or u"").lower()
    return (chapter == u"", chapter, subchapter == u"", subchapter, item.code)


def plain_number(value):
    """Numero senza zeri inutili: 300.0 -> "300", 5.1 -> "5.1"."""
    if value is None:
        return u""
    text = u"{:.6f}".format(value).rstrip(u"0").rstrip(u".")
    return text or u"0"


def money(value):
    return u"{} {:,.2f}".format(EURO, value)


def new_table(name, columns):
    table = DataTable(name)
    for column_name, clr_type in columns:
        table.Columns.Add(column_name, clr_type)
    return table


# =============================================================================
# SESSIONE
# =============================================================================

class Session(object):
    """Stato del computo condiviso fra la finestra e l'export Excel."""

    def __init__(self, doc):
        self.model_name = model_name(doc)
        self.phase_label = u""
        self.categories_label = u""
        self.generated_at = u""
        self.takeoff = None
        self.lines = []
        # Elenco prezzi: tutti i codici del modello. Computo: solo categorie spuntate.
        self.price_codes = []
        self.items = {}
        self.quantities = OrderedDict()
        # unita' con cui ogni codice del computo e' stato misurato
        self.bill_units = {}
        self.bill_issues = []
        self.bill = None
        # parametri dei livelli WBS attivi (senza i livelli lasciati vuoti)
        self.wbs_labels = []
        self.rules = qr.Rules.defaults()
        self.param_map = qm.ParameterMap.defaults()
        self.type_rows = []
        self.issues = []
        self.project_file = None
        self.price_list_path = None
        self.price_list_count = 0
        self.price_list_error = None
        self.exported_paths = []
        self.saved = False
        self.unsaved_on_close = False

    @property
    def total(self):
        total = 0.0
        for code, quantity in self.quantities.items():
            item = self.items.get(code)
            if item is not None and item.price is not None:
                total += quantity * item.price
        return total

    @property
    def unpriced_codes(self):
        return [code for code in self.quantities
                if self.items.get(code) is None or self.items[code].price is None]


# =============================================================================
# FORM
# =============================================================================

class TakeoffForm(Window):

    def __init__(self, doc):
        self.doc = doc
        self.session = Session(doc)
        self._loading = True
        self._updating = False
        self._phases = OrderedDict()
        self._category_boxes = []
        self._collect = qm.CollectResult()
        self._price_list = qs.PriceList()
        self._store = qs.ProjectStore(None)
        self._price_rows = {}
        self._bill_rows = {}
        # codice -> [(riga del riepilogo, colonna DescN)] da aggiornare quando cambia
        # la descrizione
        self._mark_cells = {}
        self._rules = qr.Rules.defaults()
        self._param_map = qm.ParameterMap.defaults()
        self._cfg = self._get_config()

        self._load_xaml()
        self._build_tables()
        self._init_options()
        self._init_store()
        self._loading = False

        self._collect_model()
        self._refresh()

    # ------------------------------------------------------------------ setup

    def _load_xaml(self):
        xaml_path = script.get_bundle_file(XAML_FILE_NAME)
        if not xaml_path or not os.path.exists(xaml_path):
            xaml_path = os.path.join(os.path.dirname(__file__), XAML_FILE_NAME)

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
        self.MinHeight = root.MinHeight
        self.MinWidth = root.MinWidth
        self.WindowStartupLocation = root.WindowStartupLocation
        self.ResizeMode = root.ResizeMode

        for name in ("txt_price_list", "btn_price_list", "btn_price_edit", "btn_price_reload",
                     "btn_price_clear",
                     "txt_project_file", "btn_project_file",
                     "cbo_phase", "cbo_phase_status", "chk_primary_only", "txt_settings_hint",
                     "txt_wbs", "btn_wbs", "txt_params", "btn_params",
                     "lst_categories", "btn_select_all", "btn_select_none",
                     "tabs", "tab_prices", "tab_bill", "tab_marks", "tab_issues",
                     "txt_search_prices", "btn_clear_prices", "chk_missing_only", "dg_prices",
                     "txt_search_bill", "btn_clear_bill", "btn_allowance", "dg_bill",
                     "txt_search_marks", "btn_clear_marks", "dg_marks", "dg_issues",
                     "tab_rules", "dg_allowance", "dg_duct_weight", "dg_density",
                     "btn_density_add", "btn_density_remove", "btn_rules_defaults",
                     "txt_status", "txt_total", "btn_close", "btn_export", "btn_save"):
            setattr(self, name, root.FindName(name))

        self.btn_price_list.Click += self.OnBrowsePriceList
        self.btn_price_reload.Click += self.OnReloadPriceList
        self.btn_price_edit.Click += self.OnEditPriceList
        self.btn_price_clear.Click += self.OnClearPriceList
        self.btn_project_file.Click += self.OnChangeProjectFile
        self.cbo_phase.SelectionChanged += self.OnCollectOptionsChanged
        self.cbo_phase_status.SelectionChanged += self.OnCollectOptionsChanged
        self.chk_primary_only.Checked += self.OnCollectOptionsChanged
        self.chk_primary_only.Unchecked += self.OnCollectOptionsChanged
        self.btn_select_all.Click += self.OnSelectAll
        self.btn_select_none.Click += self.OnSelectNone
        self.txt_search_prices.TextChanged += self.OnPriceFilterChanged
        self.chk_missing_only.Checked += self.OnPriceFilterChanged
        self.chk_missing_only.Unchecked += self.OnPriceFilterChanged
        self.btn_clear_prices.Click += self.OnClearPriceSearch
        self.txt_search_bill.TextChanged += self.OnBillFilterChanged
        self.btn_clear_bill.Click += self.OnClearBillSearch
        self.btn_allowance.Click += self.OnAllowanceOverride
        # Menu del tasto destro sulle voci del computo: la cella cliccata viene
        # selezionata dalla griglia prima che il menu si apra.
        menu = ContextMenu()
        for header, handler in ((u"Allowance override...", self.OnAllowanceOverride),
                                (u"Remove allowance override", self.OnRemoveAllowanceOverride)):
            item = MenuItem()
            item.Header = header
            item.Click += handler
            menu.Items.Add(item)
        self.dg_bill.ContextMenu = menu
        self.txt_search_marks.TextChanged += self.OnMarkFilterChanged
        self.btn_clear_marks.Click += self.OnClearMarkSearch
        self.btn_close.Click += self.OnCloseClick
        self.btn_export.Click += self.OnExport
        self.btn_save.Click += self.OnSave
        self.btn_density_add.Click += self.OnAddDensity
        self.btn_density_remove.Click += self.OnRemoveDensity
        self.btn_rules_defaults.Click += self.OnRestoreRules
        self.btn_wbs.Click += self.OnWbs
        self.btn_params.Click += self.OnParameters
        self.Closing += self.OnWindowClosing

    def _build_tables(self):
        self.prices_table = new_table("prices", (
            ("Chapter", CLR_STRING), ("Subchapter", CLR_STRING),
            ("EpuItem", CLR_STRING), ("PriceBook", CLR_STRING),
            ("Code", CLR_STRING), ("Description", CLR_STRING), ("Unit", CLR_STRING),
            ("UnitPrice", CLR_STRING), ("NoDescription", CLR_BOOL)))
        # Kind: "group" per la riga di totale di una combinazione WBS, "item" per le voci.
        # W1..W15: valori dei livelli WBS attivi.
        self.bill_table = new_table("bill", [
            ("Kind", CLR_STRING)] + [("W{}".format(i + 1), CLR_STRING)
                                     for i in range(qm.WBS_LEVELS)] + [
            ("TypeMark", CLR_STRING), ("EpuItem", CLR_STRING), ("Code", CLR_STRING),
            ("Description", CLR_STRING), ("Unit", CLR_STRING), ("Quantity", CLR_DOUBLE),
            ("UnitPrice", CLR_DOUBLE), ("Amount", CLR_DOUBLE),
            # voce con override della maggiorazione: la quantita' si colora
            ("AllowanceOverride", CLR_BOOL), ("AllowanceTip", CLR_STRING)])
        # Colonne WBS in testa alla griglia, una per livello, nascoste finche' non servono:
        # l'intestazione e' il nome del parametro, quindi si creano qui e non nell'XAML.
        self._wbs_columns = []
        for index in range(qm.WBS_LEVELS):
            column = DataGridTextColumn()
            column.Binding = Binding("W{}".format(index + 1))
            column.Width = DataGridLength(90)
            column.Visibility = Visibility.Collapsed
            self.dg_bill.Columns.Insert(index, column)
            self._wbs_columns.append(column)
        # Riepilogo Type Mark a matrice: una riga "group" per tipo (categoria, Type Mark,
        # famiglia e tipo, annidata) seguita da una riga "code" per codice. Search, nascosta,
        # porta i dati del gruppo anche sulle righe dei codici, cosi' la ricerca di un Type
        # Mark mostra tutto il gruppo.
        self.marks_table = new_table("marks", (
            ("Kind", CLR_STRING), ("Category", CLR_STRING), ("TypeMark", CLR_STRING),
            ("Types", CLR_STRING), ("Nested", CLR_STRING), ("Slot", CLR_STRING),
            ("Code", CLR_STRING), ("Description", CLR_STRING), ("Search", CLR_STRING)))
        self.issues_table = new_table("issues", (
            ("Kind", CLR_STRING), ("Subject", CLR_STRING), ("Detail", CLR_STRING),
            ("Instances", CLR_INT)))

        self.prices_table.ColumnChanging += self.OnPriceChanging
        self.prices_table.ColumnChanged += self.OnPriceChanged

        self.dg_prices.ItemsSource = self.prices_table.DefaultView
        self.dg_bill.ItemsSource = self.bill_table.DefaultView
        self.dg_marks.ItemsSource = self.marks_table.DefaultView
        self.dg_issues.ItemsSource = self.issues_table.DefaultView

        self.allowance_table = new_table("allowance", (
            ("Key", CLR_STRING), ("Category", CLR_STRING), ("Allowance", CLR_STRING)))
        self.weight_table = new_table("duct_weight", (
            ("Shape", CLR_STRING), ("ShapeLabel", CLR_STRING), ("Limit", CLR_STRING),
            ("Weight", CLR_STRING)))
        self.density_table = new_table("pipe_density", (
            ("TypeMark", CLR_STRING), ("Note", CLR_STRING), ("Density", CLR_STRING)))
        for table in (self.allowance_table, self.weight_table, self.density_table):
            table.ColumnChanging += self.OnRuleChanging
            table.ColumnChanged += self.OnRuleChanged
        self.density_table.RowDeleted += self.OnRuleRowDeleted
        self.dg_allowance.ItemsSource = self.allowance_table.DefaultView
        self.dg_duct_weight.ItemsSource = self.weight_table.DefaultView
        self.dg_density.ItemsSource = self.density_table.DefaultView

        # La colonna a tendina dell'unita' non sta nell'albero visuale: la si cerca
        # fra le colonne della griglia.
        self.col_unit = None
        for column in self.dg_prices.Columns:
            if cell_text(column.Header) == u"Unit":
                self.col_unit = column

    def _init_options(self):
        cfg = self._cfg
        view_phase = self._active_view_phase_name()
        for phase in self.doc.Phases:
            self._phases[phase.Name] = phase
            self.cbo_phase.Items.Add(phase.Name)
        wanted = view_phase or self._cfg_get('last_phase', None)
        if wanted in self._phases:
            self.cbo_phase.SelectedItem = wanted
        elif self.cbo_phase.Items.Count:
            self.cbo_phase.SelectedIndex = self.cbo_phase.Items.Count - 1

        for status in qm.PHASE_STATUS_OPTIONS:
            self.cbo_phase_status.Items.Add(status)
        status = self._cfg_get('last_phase_status', qm.PHASE_STATUS_NEW)
        self.cbo_phase_status.SelectedItem = status if status in qm.PHASE_STATUS_OPTIONS \
            else qm.PHASE_STATUS_NEW

        self.chk_primary_only.IsChecked = bool(self._cfg_get('last_primary_only', True))

        saved = self._cfg_get('last_categories', None) if cfg is not None else None
        wanted_keys = set(saved) if saved else None
        for key, label, _, _ in qm.available_rules():
            box = CheckBox()
            box.Content = label
            box.Tag = key
            box.IsChecked = wanted_keys is None or key in wanted_keys
            box.Checked += self.OnCategoryChanged
            box.Unchecked += self.OnCategoryChanged
            self._category_boxes.append(box)
            self.lst_categories.Items.Add(box)

    def _active_view_phase_name(self):
        try:
            param = self.doc.ActiveView.get_Parameter(DB.BuiltInParameter.VIEW_PHASE)
            if param is not None:
                phase = self.doc.GetElement(param.AsElementId())
                if phase is not None:
                    return phase.Name
        except Exception:
            pass
        return None

    def _init_store(self):
        path = self._remembered_project_file() or qs.default_project_file(self.doc)
        self._open_project_file(path)
        price_path = self._store.price_list_path or self._cfg_get('last_price_list', None)
        if price_path and not self._store.price_list_path:
            # Il percorso preso dalla config entra nel file al prossimo salvataggio,
            # senza segnare modifiche non salvate.
            self._store.price_list_path = price_path
        self._load_price_list(price_path, show_errors=False)

    # ------------------------------------------------------------ config

    def _get_config(self):
        try:
            return script.get_config(CONFIG_SECTION)
        except Exception:
            return None

    def _cfg_get(self, name, default):
        if self._cfg is None:
            return default
        try:
            return self._cfg.get_option(name, default)
        except Exception:
            return default

    def _project_file_map(self):
        entries = self._cfg_get('project_files', None) or []
        mapping = OrderedDict()
        for entry in entries:
            if CONFIG_SEPARATOR in entry:
                key, path = entry.split(CONFIG_SEPARATOR, 1)
                mapping[key] = path
        return mapping

    def _remembered_project_file(self):
        return self._project_file_map().get(qs.document_key(self.doc))

    def _remember_project_file(self, path):
        if self._cfg is None:
            return
        try:
            mapping = self._project_file_map()
            key = qs.document_key(self.doc)
            if path == qs.default_project_file(self.doc):
                mapping.pop(key, None)
            else:
                mapping[key] = path
            # Le ultime 50 associazioni bastano: la lista cresce con i progetti.
            entries = [u"{}{}{}".format(k, CONFIG_SEPARATOR, v) for k, v in mapping.items()]
            self._cfg.project_files = entries[-50:]
            script.save_config()
        except Exception:
            pass

    def _save_settings(self):
        if self._cfg is None:
            return
        try:
            self._cfg.last_phase = self.cbo_phase.SelectedItem
            self._cfg.last_phase_status = self.cbo_phase_status.SelectedItem
            self._cfg.last_primary_only = bool(self.chk_primary_only.IsChecked)
            self._cfg.last_categories = self._selected_keys()
            self._cfg.last_price_list = self.session.price_list_path or u""
            script.save_config()
        except Exception:
            pass

    # ------------------------------------------------------------ archivio

    def _open_project_file(self, path):
        store = qs.ProjectStore(path)
        try:
            store.load()
        except qs.ProjectStoreError as error:
            MessageBox.Show(u"{}\n\nChoose another project file with Change... "
                            u"before saving.".format(error), TITLE,
                            MessageBoxButton.OK, MessageBoxImage.Warning)
            # Mai sovrascrivere un file illeggibile: si riparte senza percorso.
            store = qs.ProjectStore(None)
        self._store = store
        self._rules = qr.Rules.from_dict(store.rules)
        self.session.rules = self._rules
        self._fill_rules_tables()
        self._update_wbs_text()
        self._param_map = qm.ParameterMap.from_dict(store.parameters)
        self.session.param_map = self._param_map
        self._update_params_text()
        self.session.project_file = store.path
        if store.path:
            state = u"" if os.path.isfile(store.path) else u"  (new file, created on Save)"
            self.txt_project_file.Text = store.path + state
        else:
            self.txt_project_file.Text = u"(not set: choose it with Change... before saving)"

    def _load_price_list(self, path, show_errors):
        self.session.price_list_path = path or None
        self.session.price_list_error = None
        self.txt_price_list.Text = path or u""
        if not path:
            self._price_list = qs.PriceList()
            self.session.price_list_count = 0
            return
        self.Cursor = Cursors.Wait
        try:
            self._price_list = qs.load_price_list(path)
        except qs.PriceListError as error:
            self._price_list = qs.PriceList(path)
            self.session.price_list_error = u"{}".format(error)
            if show_errors:
                MessageBox.Show(self.session.price_list_error, TITLE,
                                MessageBoxButton.OK, MessageBoxImage.Warning)
        finally:
            self.Cursor = None
        self.session.price_list_count = len(self._price_list.items)

    def _ask_project_file(self):
        dialog = SaveFileDialog()
        dialog.Title = "Project file of the takeoff"
        dialog.Filter = "MEP QTO project file (*.json)|*.json"
        dialog.OverwritePrompt = False
        current = self._store.path or qs.default_project_file(self.doc)
        if current:
            dialog.InitialDirectory = os.path.dirname(current)
            dialog.FileName = os.path.basename(current)
        else:
            dialog.FileName = self.session.model_name + qs.PROJECT_FILE_SUFFIX
        if dialog.ShowDialog(self) != True:
            return None
        return dialog.FileName

    def _save(self):
        self._commit_edits()
        if not self._store.path:
            path = self._ask_project_file()
            if not path:
                return False
            self._store.path = path
            self.session.project_file = path
            self.txt_project_file.Text = path
            self._remember_project_file(path)
        try:
            merged = self._store.save(self.doc.Application.Username)
        except qs.ProjectStoreError as error:
            MessageBox.Show(u"{}".format(error), TITLE, MessageBoxButton.OK, MessageBoxImage.Error)
            return False

        self.session.saved = True
        self.txt_project_file.Text = self._store.path
        self._save_settings()
        if merged:
            # Qualcun altro aveva salvato nel frattempo: le sue modifiche sono ora
            # nell'archivio e vanno mostrate.
            self._refresh()
            self._set_status(u"Saved {}. Someone else had saved the file meanwhile: "
                             u"their changes were kept and are now shown.".format(
                                 datetime.now().strftime("%H:%M")), False)
        else:
            self._set_status(u"Saved to the project file at {}.".format(
                datetime.now().strftime("%H:%M")), False)
        return True

    # ------------------------------------------------------------ calcolo

    def _current_phase(self):
        name = self.cbo_phase.SelectedItem
        return self._phases.get(name) if name else None

    def _selected_keys(self):
        return [box.Tag for box in self._category_boxes if box.IsChecked]

    def _collect_model(self):
        options = qm.CollectOptions(self._current_phase(),
                                    self.cbo_phase_status.SelectedItem or qm.PHASE_STATUS_NEW,
                                    bool(self.chk_primary_only.IsChecked),
                                    self._wbs_labels(), self._param_map)
        self.Cursor = Cursors.Wait
        try:
            self._collect = qm.collect_records(self.doc, options)
        finally:
            self.Cursor = None

        counts = {}
        for record in self._collect.records:
            counts[record.category_key] = counts.get(record.category_key, 0) + 1
        for box in self._category_boxes:
            box.Content = u"{} ({})".format(qm.category_label(box.Tag), counts.get(box.Tag, 0))

    def _refresh(self):
        """Ricalcola computo e griglie dai record gia' raccolti."""
        self._commit_edits()
        session = self.session
        selected = self._selected_keys()
        session.takeoff = qm.aggregate(self._collect, selected)
        session.price_codes = sorted(set(qm.model_codes(self._collect)) |
                                     set(session.takeoff.codes()))
        session.items = qs.merge_items(session.price_codes,
                                       self._price_list.items, self._store.items)
        session.price_codes.sort(key=lambda code: price_sort_key(session.items[code]))
        session.type_rows = qm.type_rows(self._collect, selected)

        phase = self.cbo_phase.SelectedItem or u"(no phase)"
        session.phase_label = u"{} ({})".format(phase, self.cbo_phase_status.SelectedItem)
        total_categories = len(self._category_boxes)
        session.categories_label = u"all {}".format(total_categories) \
            if len(selected) == total_categories \
            else u"{} of {}".format(len(selected), total_categories)

        self._fill_prices_table()
        self._fill_marks_table()
        self._recompute_bill()
        self._update_hint()

    def _recompute_bill(self):
        """Quantita' del computo dalle unita' delle voci e dalle regole correnti."""
        session = self.session
        bill = qm.compute_bill(session.takeoff, session.items, self._rules,
                               self._store.allowance_overrides)
        session.lines = bill.lines
        session.quantities = bill.quantities
        session.bill_units = bill.units
        session.bill_issues = bill.issues
        session.bill = bill
        session.wbs_labels = self._wbs_labels()
        session.rules = self._rules
        self._fill_bill_table()
        self._refresh_issues()
        self._update_totals()

    def _refresh_issues(self):
        session = self.session
        session.issues = list(session.takeoff.issues) + list(session.bill_issues) + \
            qm.description_issues(session.takeoff, session.items)
        if session.price_list_error:
            session.issues.insert(0, qm.Issue(u"Price list not loaded",
                                              session.price_list_path or u"",
                                              session.price_list_error))
        self._updating = True
        try:
            self.issues_table.Rows.Clear()
            for issue in session.issues:
                row = self.issues_table.NewRow()
                row["Kind"] = issue.kind
                row["Subject"] = issue.subject
                row["Detail"] = issue.detail
                row["Instances"] = issue.count
                self.issues_table.Rows.Add(row)
        finally:
            self._updating = False
        self.tab_issues.Header = u"\u26a0\ufe0f Issues ({})".format(len(session.issues))

    # ------------------------------------------------------------ griglie

    def _unit_choices(self):
        """Valori della tendina: vuoto (= valore del listino), le unita' predefinite e
        quelle non standard arrivate dal listino, perche' nessuna sparisca."""
        extra = set()
        for item in self.session.items.values():
            if item.unit and item.unit not in qs.UNITS:
                extra.add(item.unit)
        return [u""] + list(qs.UNITS) + sorted(extra)

    def _write_price_values(self, row, item):
        row["Chapter"] = item.chapter or u""
        row["Subchapter"] = item.subchapter or u""
        row["EpuItem"] = item.epu_item or u""
        row["PriceBook"] = item.price_book or u""
        row["Description"] = item.description or u""
        row["Unit"] = item.unit or u""
        row["UnitPrice"] = qs.format_decimal(item.price)
        row["NoDescription"] = not item.description

    def _write_bill_values(self, row, item, quantity, unit):
        row["EpuItem"] = item.epu_item or u""
        row["Description"] = item.description or u""
        row["Unit"] = unit or u""
        row["Quantity"] = quantity
        if item.price is None:
            row["UnitPrice"] = DBNull.Value
            row["Amount"] = DBNull.Value
        else:
            row["UnitPrice"] = item.price
            row["Amount"] = quantity * item.price

    def _fill_prices_table(self):
        session = self.session
        if self.col_unit is not None:
            self.col_unit.ItemsSource = self._unit_choices()
        self._price_rows = {}
        self._updating = True
        try:
            self.prices_table.Rows.Clear()
            for code in session.price_codes:
                row = self.prices_table.NewRow()
                row["Code"] = code
                self._write_price_values(row, session.items[code])
                self.prices_table.Rows.Add(row)
                self._price_rows[code] = row
        finally:
            self._updating = False

    def _fill_bill_table(self):
        """Computo per codice o, con i livelli WBS, raggruppato con i subtotali."""
        session = self.session
        labels = list(session.wbs_labels)
        outline = qm.bill_outline(session.bill, labels)
        for index, column in enumerate(self._wbs_columns):
            if index < len(labels):
                column.Header = labels[index]
                column.Visibility = Visibility.Visible
            else:
                column.Visibility = Visibility.Collapsed
        amounts = []
        for entry in outline:
            item = session.items[entry.code] if entry.kind == "item" else None
            amounts.append(entry.quantity * item.price
                           if item is not None and item.price is not None else 0.0)

        self._bill_rows = {}
        self._updating = True
        try:
            self.bill_table.Rows.Clear()
            for index, entry in enumerate(outline):
                row = self.bill_table.NewRow()
                for level, value in enumerate(qm.wbs_cells(entry.key, len(labels))):
                    row["W{}".format(level + 1)] = value
                if entry.kind == "group":
                    end = qm.outline_group_end(outline, index)
                    row["Kind"] = u"group"
                    row["Description"] = entry.text
                    row["Amount"] = sum(amounts[index + 1:end + 1])
                else:
                    row["Kind"] = u"item"
                    row["TypeMark"] = entry.type_mark
                    row["Code"] = entry.code
                    self._write_bill_values(row, session.items[entry.code], entry.quantity,
                                            session.bill_units.get(entry.code, u""))
                    line_key = (entry.type_mark, entry.code)
                    overridden = line_key in session.bill.overridden
                    row["AllowanceOverride"] = overridden
                    row["AllowanceTip"] = self._allowance_tip(line_key) if overridden else u""
                    self._bill_rows.setdefault(entry.code, row)
                self.bill_table.Rows.Add(row)
        finally:
            self._updating = False
        # Con la WBS l'ordine delle righe e' la struttura: niente riordino per colonna.
        self.dg_bill.CanUserSortColumns = not session.wbs_labels

    def _fill_marks_table(self):
        """Una riga di gruppo per tipo, poi una riga per ogni codice valorizzato."""
        session = self.session
        self._mark_cells = {}
        self._updating = True
        try:
            self.marks_table.Rows.Clear()
            for type_row in session.type_rows:
                entries = type_row.code_entries()
                # Separatore che l'utente non puo' digitare nella ricerca.
                group_text = u"\n".join((type_row.category, type_row.type_mark,
                                         type_row.type_label))
                row = self.marks_table.NewRow()
                row["Kind"] = u"group"
                row["Category"] = type_row.category
                row["TypeMark"] = type_row.type_mark
                row["Types"] = type_row.type_label
                row["Nested"] = u"Yes" if type_row.nested else u"No"
                row["Search"] = u"\n".join([group_text] + [code for _, code in entries])
                self.marks_table.Rows.Add(row)
                for label, code in entries:
                    item = session.items.get(code)
                    row = self.marks_table.NewRow()
                    row["Kind"] = u"code"
                    row["Slot"] = label
                    row["Code"] = code
                    row["Description"] = item.description if item else u""
                    row["Search"] = u"\n".join((group_text, code))
                    self.marks_table.Rows.Add(row)
                    self._mark_cells.setdefault(code, []).append((row, "Description"))
        finally:
            self._updating = False

    def _commit_edits(self):
        try:
            self.dg_prices.CommitEdit(DataGridEditingUnit.Cell, True)
            self.dg_prices.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            pass

    def _update_totals(self):
        session = self.session
        unpriced = len(session.unpriced_codes)
        text = u"Total: {}".format(money(session.total))
        if unpriced:
            text += u"   ({} of {} codes without price)".format(unpriced, len(session.quantities))
        self.txt_total.Text = text

    def _update_hint(self):
        session = self.session
        takeoff = session.takeoff
        parts = [u"{} elements counted, {} Type Marks, {} price codes in the bill.".format(
            takeoff.instance_count, len(takeoff.groups), len(session.quantities))]
        if self._collect.skipped_options:
            parts.append(u"{} instances in secondary design options skipped.".format(
                self._collect.skipped_options))
        excluded = len(getattr(self._collect, "excluded", []))
        if excluded:
            parts.append(u"{} elements excluded by the Yes/No include parameter.".format(excluded))
        if session.price_list_error:
            parts.append(u"Price list NOT loaded: see the Issues tab.")
        elif session.price_list_path:
            parts.append(u"Price list: {} items{}.".format(
                session.price_list_count,
                u" ({} duplicate codes ignored)".format(self._price_list.duplicates)
                if self._price_list.duplicates else u""))
        else:
            parts.append(u"No price list: descriptions come from the project file only.")
        self.txt_settings_hint.Text = u" ".join(parts)
        self.txt_settings_hint.Foreground = RED_BRUSH if session.price_list_error else GRAY_BRUSH

    def _set_status(self, text, is_warning):
        self.txt_status.Text = text
        self.txt_status.Foreground = RED_BRUSH if is_warning else GRAY_BRUSH

    def _mark_dirty(self):
        self._set_status(u"Unsaved changes: press Save to write them to the project file.", True)

    def _defer(self, callback):
        """Esegue callback dopo che il DataGrid ha chiuso il commit della cella."""
        self.Dispatcher.BeginInvoke(DispatcherPriority.Background, Action(callback))

    def _warn_later(self, message):
        self._defer(lambda: MessageBox.Show(message, TITLE, MessageBoxButton.OK,
                                            MessageBoxImage.Warning))

    # ------------------------------------------------------------ modifiche elenco prezzi

    def OnPriceChanging(self, sender, args):
        if self._updating:
            return
        name = args.Column.ColumnName
        if name not in PRICE_FIELDS:
            return
        text = cell_text(args.ProposedValue)
        if name != "UnitPrice":
            args.ProposedValue = text
            return
        try:
            value = qs.parse_decimal(text)
        except ValueError:
            value = -1
        if value is not None and value < 0:
            args.ProposedValue = args.Row["UnitPrice"]
            self._warn_later(u"'{}' is not a valid unit price.".format(text))
            return
        args.ProposedValue = qs.format_decimal(value)

    def OnPriceChanged(self, sender, args):
        if self._updating:
            return
        field = PRICE_FIELDS.get(args.Column.ColumnName)
        if field is None:
            return
        code = cell_text(args.Row["Code"])
        text = cell_text(args.Row[args.Column.ColumnName])
        value = qs.parse_decimal(text) if field == "price" else text
        current = self.session.items.get(code)
        if current is not None and getattr(current, field) == value:
            return

        self._store.set_item_field(code, field, value)
        self.session.items[code] = qs.merge_item(code, self._price_list.items, self._store.items)
        self._mark_dirty()
        self._defer(lambda: self._after_price_edit(code))

    def _after_price_edit(self, code):
        session = self.session
        item = session.items.get(code)
        if item is None:
            return
        self._updating = True
        try:
            # Un campo svuotato torna al valore del listino: la riga lo mostra.
            row = self._price_rows.get(code)
            if row is not None:
                self._write_price_values(row, item)
            for mark_row, column in self._mark_cells.get(code, []):
                mark_row[column] = item.description or u""
        finally:
            self._updating = False
        # L'unita' decide come si misurano le categorie lineari: si ricalcola il computo.
        self._recompute_bill()

    # ------------------------------------------------------------ parametri

    def _update_params_text(self):
        param_map = self._param_map
        text = u"{} (types)  |  {} (instances)  |  Yes/No: {}".format(
            param_map.piece_label, param_map.linear_label, param_map.include or u"(none)")
        if param_map.is_default:
            text = u"Default: " + text
        self.txt_params.Text = text
        self.txt_params.ToolTip = u"\n".join(
            [u"Piece categories, type parameters:"] +
            [u"  Code {}: {}".format(i + 1, n) for i, n in enumerate(param_map.piece_codes) if n] +
            [u"All categories, instance parameters:"] +
            [u"  Code {}: {}".format(i + 1, n) for i, n in enumerate(param_map.linear_codes) if n] +
            [u"Include (Yes/No): {} on the instances of all categories".format(
                param_map.include or u"(none)")])

    def OnParameters(self, sender, args):
        self.Cursor = Cursors.Wait
        try:
            names = qm.sample_parameter_names(self.doc, self._collect)
        finally:
            self.Cursor = None
        param_map = show_parameters_dialog(self, self._param_map, names)
        if param_map is None or param_map == self._param_map:
            return
        self._apply_parameters(param_map)

    def _apply_parameters(self, param_map):
        """Nuova mappa: si salva nel progetto e i codici si rileggono dal modello."""
        self._param_map = param_map
        self.session.param_map = param_map
        self._store.set_parameters(param_map.to_dict())
        self._mark_dirty()
        self._update_params_text()
        self._collect_model()
        self._refresh()

    # ------------------------------------------------------------ WBS

    def _wbs_levels(self):
        """I WBS_LEVELS nomi salvati nel file di progetto ("" per i livelli spenti)."""
        levels = list(self._store.wbs or [])[:qm.WBS_LEVELS]
        return levels + [u""] * (qm.WBS_LEVELS - len(levels))

    def _wbs_labels(self):
        return [name for name in self._wbs_levels() if name]

    def _update_wbs_text(self):
        labels = self._wbs_labels()
        if labels:
            self.txt_wbs.Text = u"  >  ".join(labels)
            self.txt_wbs.Foreground = BLACK_BRUSH
            self.txt_wbs.ToolTip = u"\n".join(
                u"Level {}: {}".format(index + 1, name)
                for index, name in enumerate(self._wbs_levels()) if name)
        else:
            self.txt_wbs.Text = u"(none: the bill of quantities is not grouped)"
            self.txt_wbs.Foreground = GRAY_BRUSH
            self.txt_wbs.ToolTip = None

    def OnWbs(self, sender, args):
        self.Cursor = Cursors.Wait
        try:
            names = qm.sample_parameter_names(self.doc, self._collect)
        finally:
            self.Cursor = None
        levels = show_wbs_dialog(self, self._wbs_levels(), names, qm.WBS_LEVELS)
        if levels is None or levels == self._wbs_levels():
            return
        self._apply_wbs(levels)

    def _apply_wbs(self, levels):
        """Nuovi livelli WBS: si salvano nel progetto e si rileggono i valori dal modello."""
        self._store.set_wbs(levels)
        self._mark_dirty()
        self._update_wbs_text()
        self._collect_model()
        self._refresh()
        self.tabs.SelectedItem = self.tab_bill

    # ------------------------------------------------------------ regole

    def _fill_rules_tables(self):
        rules = self._rules
        self._updating = True
        try:
            self.allowance_table.Rows.Clear()
            for key, label, kind, _ in qm.available_rules():
                if kind == qr.KIND_COUNT:
                    continue
                row = self.allowance_table.NewRow()
                row["Key"] = key
                row["Category"] = label
                row["Allowance"] = plain_number(rules.allowance_for(key) * 100.0)
                self.allowance_table.Rows.Add(row)

            self.weight_table.Rows.Clear()
            for shape in (qr.SHAPE_RECT, qr.SHAPE_ROUND):
                for limit, weight in qr.sorted_bands(rules.duct_weight.get(shape, [])):
                    row = self.weight_table.NewRow()
                    row["Shape"] = shape
                    row["ShapeLabel"] = qr.SHAPE_LABELS[shape]
                    row["Limit"] = plain_number(limit)
                    row["Weight"] = plain_number(weight)
                    self.weight_table.Rows.Add(row)

            self.density_table.Rows.Clear()
            for mark, note, density in rules.pipe_density:
                row = self.density_table.NewRow()
                row["TypeMark"] = mark
                row["Note"] = note
                row["Density"] = plain_number(density)
                self.density_table.Rows.Add(row)
        finally:
            self._updating = False

    def _rules_from_tables(self):
        allowance = dict(self._rules.allowance)
        for row in self.allowance_table.Rows:
            value = qs.parse_decimal(cell_text(row["Allowance"]))
            allowance[cell_text(row["Key"])] = (value or 0.0) / 100.0

        duct_weight = {qr.SHAPE_RECT: [], qr.SHAPE_ROUND: []}
        for row in self.weight_table.Rows:
            weight = qs.parse_decimal(cell_text(row["Weight"]))
            if weight is None:
                continue
            duct_weight[cell_text(row["Shape"])].append(
                (qs.parse_decimal(cell_text(row["Limit"])), weight))

        pipe_density = []
        for row in self.density_table.Rows:
            if row.RowState == DataRowState.Deleted:
                continue
            mark = cell_text(row["TypeMark"])
            density = qs.parse_decimal(cell_text(row["Density"]))
            if mark and density is not None:
                pipe_density.append((mark, cell_text(row["Note"]), density))

        return qr.Rules(allowance, dict((k, qr.sorted_bands(v)) for k, v in duct_weight.items()),
                        pipe_density)

    def _apply_rules(self):
        self._rules = self._rules_from_tables()
        self._store.set_rules(self._rules.to_dict())
        self._mark_dirty()
        self._defer(self._recompute_bill)

    def OnRuleChanging(self, sender, args):
        if self._updating:
            return
        name = args.Column.ColumnName
        text = cell_text(args.ProposedValue)
        if name in ("TypeMark", "Note"):
            if name == "TypeMark" and text:
                for row in self.density_table.Rows:
                    if row.RowState != DataRowState.Deleted and row is not args.Row \
                            and cell_text(row["TypeMark"]) == text:
                        args.ProposedValue = args.Row["TypeMark"]
                        self._warn_later(u"Type Mark '{}' is already in the table.".format(text))
                        return
            args.ProposedValue = text
            return
        if name not in ("Allowance", "Limit", "Weight", "Density"):
            return
        try:
            value = qs.parse_decimal(text)
        except ValueError:
            value = -1
        empty_ok = name in ("Limit", "Density")
        if (value is None and not empty_ok) or (value is not None and value < 0) or \
                (value == 0 and name != "Allowance"):
            args.ProposedValue = args.Row[name]
            self._warn_later(u"'{}' is not a valid value.".format(text))
            return
        args.ProposedValue = plain_number(value)

    def OnRuleChanged(self, sender, args):
        if self._updating:
            return
        self._apply_rules()

    def OnRuleRowDeleted(self, sender, args):
        if self._updating:
            return
        self._apply_rules()

    def OnAddDensity(self, sender, args):
        self._updating = True
        try:
            row = self.density_table.NewRow()
            row["TypeMark"] = u""
            row["Note"] = u""
            row["Density"] = u""
            self.density_table.Rows.Add(row)
        finally:
            self._updating = False
        self.dg_density.ScrollIntoView(self.dg_density.Items[self.dg_density.Items.Count - 1])

    def OnRemoveDensity(self, sender, args):
        try:
            self.dg_density.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            pass
        selected = [item.Row for item in self.dg_density.SelectedItems
                    if hasattr(item, "Row")]
        if not selected:
            cell = self.dg_density.CurrentCell
            if cell.Item is not None and hasattr(cell.Item, "Row"):
                selected = [cell.Item.Row]
        for row in selected:
            row.Delete()
        self.density_table.AcceptChanges()

    def OnRestoreRules(self, sender, args):
        self._rules = qr.Rules.defaults()
        self._fill_rules_tables()
        self._store.set_rules(self._rules.to_dict())
        self._mark_dirty()
        self._recompute_bill()

    # ------------------------------------------------------------ filtri

    def OnPriceFilterChanged(self, sender, args):
        self._commit_edits()
        parts = []
        text = escape_like((self.txt_search_prices.Text or u"").strip())
        if text:
            parts.append(u"(Code LIKE '%{0}%' OR Description LIKE '%{0}%' "
                         u"OR Chapter LIKE '%{0}%' OR Subchapter LIKE '%{0}%' "
                         u"OR EpuItem LIKE '%{0}%' OR PriceBook LIKE '%{0}%')".format(text))
        if self.chk_missing_only.IsChecked:
            parts.append(u"NoDescription = true")
        self.prices_table.DefaultView.RowFilter = u" AND ".join(parts)

    def OnClearPriceSearch(self, sender, args):
        self.txt_search_prices.Text = u""

    def OnBillFilterChanged(self, sender, args):
        text = escape_like((self.txt_search_bill.Text or u"").strip())
        if not text:
            self.bill_table.DefaultView.RowFilter = u""
            return
        # La ricerca vale anche sui valori WBS; le righe di totale restano visibili, senza
        # le voci filtrate perdono la loro combinazione.
        columns = ["TypeMark", "EpuItem", "Code", "Description"] + ["W{}".format(i + 1)
                                             for i in range(len(self.session.wbs_labels))]
        self.bill_table.DefaultView.RowFilter = u"Kind = 'group' OR " + u" OR ".join(
            u"{} LIKE '%{}%'".format(column, text) for column in columns)

    def OnClearBillSearch(self, sender, args):
        self.txt_search_bill.Text = u""

    # ------------------------------------------------------------ override maggiorazione

    def _selected_bill_lines(self):
        """(Type Mark, codice) delle voci selezionate nel computo, senza le righe di
        totale WBS e senza ripetizioni (con la WBS una voce compare in piu' righe)."""
        views = [cell.Item for cell in self.dg_bill.SelectedCells]
        if not views and self.dg_bill.CurrentCell.Item is not None:
            views = [self.dg_bill.CurrentCell.Item]
        lines = OrderedDict()
        for view in views:
            if not hasattr(view, "Row"):
                continue
            row = view.Row
            if cell_text(row["Kind"]) != u"item":
                continue
            lines[(cell_text(row["TypeMark"]), cell_text(row["Code"]))] = True
        return list(lines.keys())

    def _category_allowances(self, line_key):
        """Maggiorazioni di categoria della voce, es. "Ducts 30%, Duct Insulation 30%"."""
        return u", ".join(
            u"{} {}%".format(qm.category_label(key),
                             plain_number(self._rules.allowance_for(key) * 100.0))
            for key in self.session.bill.line_categories.get(line_key, []))

    def _allowance_tip(self, line_key):
        fraction = self.session.bill.overridden.get(line_key)
        return u"Allowance override: {}% (category: {})".format(
            plain_number(fraction * 100.0), self._category_allowances(line_key))

    def OnAllowanceOverride(self, sender, args):
        lines = self._selected_bill_lines() if self.session.bill is not None else []
        if not lines:
            MessageBox.Show(u"Select one or more items of the bill of quantities first.",
                            TITLE, MessageBoxButton.OK, MessageBoxImage.Information)
            return
        line_categories = self.session.bill.line_categories
        measured = [key for key in lines if key in line_categories]
        skipped = len(lines) - len(measured)
        if not measured:
            MessageBox.Show(u"The selected items are counted by piece (1 per instance): "
                            u"they have no allowance to override.", TITLE,
                            MessageBoxButton.OK, MessageBoxImage.Information)
            return

        overrides = self._store.allowance_overrides
        text_lines = []
        for key in measured:
            text = u"{}  |  {}  |  {}".format(key[0], key[1], self._category_allowances(key))
            if key in overrides:
                text += u"  |  override {}%".format(plain_number(overrides[key] * 100.0))
            text_lines.append(text)
        if skipped:
            text_lines.append(u"")
            text_lines.append(u"{} selected items are counted by piece and are left out.".format(
                skipped))
        current = set(overrides.get(key) for key in measured)
        initial = plain_number(current.pop() * 100.0) \
            if len(current) == 1 and None not in current else u""

        result = show_allowance_dialog(self, u"\n".join(text_lines), len(measured), initial,
                                       any(key in overrides for key in measured))
        if result is None:
            return
        action, fraction = result
        self._apply_allowance_override(measured, fraction if action == "set" else None)

    def OnRemoveAllowanceOverride(self, sender, args):
        overrides = self._store.allowance_overrides
        lines = [key for key in self._selected_bill_lines() if key in overrides]
        if not lines:
            self._set_status(u"The selected items have no allowance override.", False)
            return
        self._apply_allowance_override(lines, None)

    def _apply_allowance_override(self, lines, fraction):
        """Scrive l'override (None = toglie) nel file di progetto e ricalcola il computo."""
        for type_mark, code in lines:
            self._store.set_allowance_override(type_mark, code, fraction)
        self._recompute_bill()
        if fraction is None:
            done = u"Allowance override removed from {} items".format(len(lines))
        else:
            done = u"Allowance {}% set on {} items".format(plain_number(fraction * 100.0),
                                                            len(lines))
        self._set_status(done + u". Unsaved changes: press Save to write them to the "
                                u"project file.", True)

    def OnMarkFilterChanged(self, sender, args):
        text = escape_like((self.txt_search_marks.Text or u"").strip())
        if not text:
            self.marks_table.DefaultView.RowFilter = u""
            return
        # Search contiene categoria, Type Mark, famiglia e tipo e i codici: un tipo trovato
        # mostra tutti i suoi codici, un codice trovato mostra anche la riga del suo tipo.
        self.marks_table.DefaultView.RowFilter = u"Search LIKE '%{}%'".format(text)

    def OnClearMarkSearch(self, sender, args):
        self.txt_search_marks.Text = u""

    # ------------------------------------------------------------ impostazioni

    def OnCollectOptionsChanged(self, sender, args):
        if self._loading:
            return
        self._collect_model()
        self._refresh()

    def OnCategoryChanged(self, sender, args):
        if self._loading:
            return
        self._refresh()

    def _set_all_categories(self, checked):
        self._loading = True
        try:
            for box in self._category_boxes:
                box.IsChecked = checked
        finally:
            self._loading = False
        self._refresh()

    def OnSelectAll(self, sender, args):
        self._set_all_categories(True)

    def OnSelectNone(self, sender, args):
        self._set_all_categories(False)

    def OnBrowsePriceList(self, sender, args):
        dialog = OpenFileDialog()
        dialog.Title = "Shared price list"
        dialog.Filter = ("Price list (*.json;*.xlsx;*.xlsm;*.csv)|*.json;*.xlsx;*.xlsm;*.csv"
                         "|All files (*.*)|*.*")
        if self.session.price_list_path:
            folder = os.path.dirname(self.session.price_list_path)
            if os.path.isdir(folder):
                dialog.InitialDirectory = folder
        if dialog.ShowDialog(self) != True:
            return
        self._use_price_list(dialog.FileName)

    def _use_price_list(self, path):
        self._load_price_list(path, show_errors=True)
        self._store.set_price_list_path(path)
        self._mark_dirty()
        self._save_settings()
        self._refresh()

    def OnEditPriceList(self, sender, args):
        """Editor del listino JSON. Un listino Excel / CSV in uso viene importato in un
        listino nuovo, da salvare come .json."""
        current = self.session.price_list_path
        is_json = bool(current) and current.lower().endswith(u".json")
        saved = show_pricelist_editor(
            self, path=current if is_json else None,
            import_path=current if current and not is_json and os.path.isfile(current) else None,
            user_name=self.doc.Application.Username)
        if not saved:
            return
        if current and os.path.normcase(saved) == os.path.normcase(current):
            self._load_price_list(current, show_errors=True)
            self._refresh()
            self._set_status(u"Price list reloaded after editing.", False)
            return
        answer = MessageBox.Show(u"Use the saved price list for this takeoff?\n\n{}".format(saved),
                                 TITLE, MessageBoxButton.YesNo, MessageBoxImage.Question)
        if answer == MessageBoxResult.Yes:
            self._use_price_list(saved)

    def OnReloadPriceList(self, sender, args):
        self._load_price_list(self.session.price_list_path, show_errors=True)
        self._refresh()

    def OnClearPriceList(self, sender, args):
        if not self.session.price_list_path:
            return
        self._load_price_list(None, show_errors=False)
        self._store.set_price_list_path(None)
        self._mark_dirty()
        self._refresh()

    def OnChangeProjectFile(self, sender, args):
        self._commit_edits()
        if self._store.is_dirty:
            answer = MessageBox.Show(
                u"Save the changes to the current project file before switching?",
                TITLE, MessageBoxButton.YesNoCancel, MessageBoxImage.Question)
            if answer == MessageBoxResult.Cancel:
                return
            if answer == MessageBoxResult.Yes and not self._save():
                return
        path = self._ask_project_file()
        if not path:
            return
        previous_price_list = self.session.price_list_path
        self._open_project_file(path)
        if self._store.price_list_path and self._store.price_list_path != previous_price_list:
            self._load_price_list(self._store.price_list_path, show_errors=True)
        elif not self._store.price_list_path:
            self._store.price_list_path = previous_price_list
        self._remember_project_file(path)
        self._set_status(u"Project file: {}".format(path), False)
        # Il nuovo file puo' avere altri livelli WBS: i valori si rileggono dal modello.
        self._collect_model()
        self._refresh()

    # ------------------------------------------------------------ bottoni

    def OnSave(self, sender, args):
        self._save()

    def _default_export_path(self):
        name = u"{}_MEPQTO_{}.xlsx".format(datetime.now().strftime("%y%m%d_%H%M%S"),
                                          self.session.model_name)
        folder = None
        for candidate in (self._store.path, qs.document_key(self.doc)):
            if candidate and "://" not in candidate and os.path.isdir(os.path.dirname(candidate)):
                folder = os.path.dirname(candidate)
                break
        return folder, name

    def OnExport(self, sender, args):
        self._commit_edits()
        folder, name = self._default_export_path()
        dialog = SaveFileDialog()
        dialog.Title = "Export the takeoff to Excel"
        dialog.Filter = "Excel workbook (*.xlsx)|*.xlsx"
        dialog.FileName = name
        if folder:
            dialog.InitialDirectory = folder
        if dialog.ShowDialog(self) != True:
            return
        path = dialog.FileName

        self.session.generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
        error_message = None
        self.Cursor = Cursors.Wait
        try:
            qx.export_takeoff(path, self.session, qm.get_element_id_value,
                              qm.category_label, qm.LINEAR_KEYS)
        except qx.FileLockedError:
            error_message = (u"The file is open in another program (Excel?):\n{}\n\n"
                             u"Close it and export again.".format(path))
        except Exception as error:
            error_message = u"Export failed:\n{}".format(error)
        finally:
            self.Cursor = None
        if error_message:
            MessageBox.Show(error_message, TITLE, MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        self.session.exported_paths.append(path)
        self._set_status(u"Exported {} to {}".format(datetime.now().strftime("%H:%M"), path), False)
        answer = MessageBox.Show(u"Takeoff exported:\n{}\n\nOpen it now?".format(path), TITLE,
                                 MessageBoxButton.YesNo, MessageBoxImage.Information)
        if answer == MessageBoxResult.Yes:
            try:
                Process.Start(path)
            except Exception:
                pass

    def OnCloseClick(self, sender, args):
        self.Close()

    def OnWindowClosing(self, sender, args):
        self._commit_edits()
        if self._store.is_dirty:
            answer = MessageBox.Show(
                u"There are unsaved changes (price list fields, rules, allowance overrides "
                u"or settings).\n\n"
                u"Save them to the project file?", TITLE,
                MessageBoxButton.YesNoCancel, MessageBoxImage.Warning)
            if answer == MessageBoxResult.Cancel:
                args.Cancel = True
                return
            if answer == MessageBoxResult.Yes and not self._save():
                args.Cancel = True
                return
            if answer == MessageBoxResult.No:
                self.session.unsaved_on_close = True
        self._save_settings()


def show_takeoff_window(doc):
    """Apre la finestra (modale) e restituisce la Session finale."""
    form = TakeoffForm(doc)
    form.ShowDialog()
    return form.session
