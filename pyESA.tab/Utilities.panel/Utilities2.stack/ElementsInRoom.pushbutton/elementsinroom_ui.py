# -*- coding: utf-8 -*-
"""
elementsinroom_ui.py - finestra XAML di configurazione per ElementsInRoom.

Raccoglie in un solo passaggio fasi, ambito di room ed elementi, categorie,
parametri e tolleranze direzionali, e restituisce un DTO RunConfig allo script
chiamante.

Le room possono venire dal documento corrente o da un RevitLinkInstance: in quel
caso fasi e room sono lette dal documento del link (room_doc), mentre categorie,
parametri di destinazione e fase degli elementi restano del documento corrente.
"""

import os
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window, MessageBox
from System.Windows.Controls import CheckBox
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Colors
from System.IO import FileStream, FileMode

from pyrevit import DB, script

XAML_FILE_NAME = 'ElementsInRoom_form.xaml'
CONFIG_SECTION = 'ESA_ElementsInRoom'

BIC = DB.BuiltInCategory
BIP = DB.BuiltInParameter
ST = DB.StorageType

# Numero di elementi campionati per categoria quando si cercano i parametri
# scrivibili: i collector di Revit sono lazy, quindi ci si ferma subito.
SAMPLE_PER_CATEGORY = 10
SAMPLE_ROOMS = 10

GRAY_BRUSH = SolidColorBrush(Colors.Gray)
GRAY_BRUSH.Freeze()
RED_BRUSH = SolidColorBrush(Colors.Firebrick)
RED_BRUSH.Freeze()


# =============================================================================
# UTILITY
# =============================================================================

def element_id_value(eid):
    """ElementId.IntegerValue e' stato rimosso in Revit 2026 (sostituito da .Value)."""
    if eid is None:
        return -1
    if hasattr(eid, "Value"):
        return eid.Value
    return eid.IntegerValue


def cm_to_internal(cm_value, doc):
    """Converte centimetri in unita' interne Revit (feet).

    Gestisce la differenza API fra Revit < 2022 e >= 2022, come mm_to_internal
    in legend_ui.py.
    """
    if int(doc.Application.VersionNumber) < 2022:
        return DB.UnitUtils.ConvertToInternalUnits(cm_value, DB.DisplayUnitType.DUT_CENTIMETERS)
    return DB.UnitUtils.ConvertToInternalUnits(cm_value, DB.UnitTypeId.Centimeters)


def parse_number(text):
    """float da una stringa utente, accettando la virgola. None se non e' un numero."""
    if text is None:
        return None
    cleaned = text.strip().replace(",", ".")
    if not cleaned:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return None


def collect_rooms(room_doc, phase, active_view_id):
    """Room posizionate della fase indicata nel documento room_doc.

    active_view_id, se dato, limita la raccolta alla vista: possibile solo quando
    room_doc e' il documento corrente (un collector view-scoped non accetta la
    vista di un altro documento).
    """
    if phase is None:
        return []
    phase_id = element_id_value(phase.Id)
    try:
        if active_view_id is not None:
            collector = DB.FilteredElementCollector(room_doc, active_view_id)
        else:
            collector = DB.FilteredElementCollector(room_doc)
        candidates = collector.OfCategory(BIC.OST_Rooms)\
            .WhereElementIsNotElementType()\
            .ToElements()
    except Exception:
        return []

    rooms = []
    for room in candidates:
        try:
            if room.Area <= 0 or room.Location is None:
                continue
            room_phase = room.get_Parameter(BIP.ROOM_PHASE)
            if room_phase is not None and element_id_value(room_phase.AsElementId()) != phase_id:
                continue
        except Exception:
            continue
        rooms.append(room)
    return rooms


def view_scope_supported(doc):
    """La vista attiva puo' fare da filtro solo se e' una vista grafica del modello."""
    try:
        view = doc.ActiveView
        if view is None or isinstance(view, DB.ViewSheet):
            return False
        if view.IsTemplate:
            return False
        return view.ViewType in (DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan,
                                 DB.ViewType.AreaPlan, DB.ViewType.EngineeringPlan,
                                 DB.ViewType.Section, DB.ViewType.Elevation,
                                 DB.ViewType.Detail, DB.ViewType.ThreeD)
    except Exception:
        return False


def model_categories(doc):
    """Categorie di modello a cui si possono legare parametri, Rooms escluse."""
    found = []
    for category in doc.Settings.Categories:
        try:
            if category.CategoryType != DB.CategoryType.Model:
                continue
            if element_id_value(category.Id) == int(BIC.OST_Rooms):
                continue
            if not category.AllowsBoundParameters:
                continue
        except Exception:
            continue
        found.append(category)
    found.sort(key=lambda c: c.Name)
    return found


def room_param_names(rooms):
    """Nomi dei parametri leggibili su un campione di room."""
    names = set()
    for room in rooms[:SAMPLE_ROOMS]:
        for param in room.Parameters:
            try:
                names.add(param.Definition.Name)
            except Exception:
                pass
    return sorted(names)


def writable_text_param_names(elements):
    """Parametri istanza di testo scrivibili trovati sul campione."""
    names = set()
    for element in elements:
        for param in element.Parameters:
            try:
                if param.IsReadOnly or param.StorageType != ST.String:
                    continue
                names.add(param.Definition.Name)
            except Exception:
                pass
    return names


# =============================================================================
# DTO
# =============================================================================

class CategoryItem(object):
    """Riga della lista categorie."""

    def __init__(self, category):
        self.Name = category.Name
        self.Category = category
        self.CategoryId = category.Id
        self.IsSelected = False


class RunConfig(object):
    """Configurazione completa di un run, cosi' come esce dalla finestra."""

    def __init__(self):
        # Fase delle room (documento delle room) e fase degli elementi (documento corrente).
        self.phase = None
        self.element_phase = None
        self.only_active_view = False
        self.only_view_elements = False
        self.room_doc = None
        self.link_instance = None
        self.rooms = []
        self.categories = []
        self.source_name = None
        self.target_name = None
        self.separator = ";"
        self.overwrite = True
        self.retry_at_level = True
        self.local_axes = False
        # Tolleranze: in cm per il report, in piedi per il calcolo.
        self.tol_cm = {"x": 0.0, "y": 0.0, "zup": 0.0, "zdown": 0.0}
        self.tol_ft = {"x": 0.0, "y": 0.0, "zup": 0.0, "zdown": 0.0}

    @property
    def has_tolerance(self):
        return any(value > 0 for value in self.tol_ft.values())


# =============================================================================
# FORM
# =============================================================================

class ElementsInRoomForm(Window):
    """Finestra XAML unica di configurazione."""

    def __init__(self, doc, room_doc, link_instance):
        self.doc = doc
        self.room_doc = room_doc
        self.link_instance = link_instance
        self.link_mode = link_instance is not None
        self.result = False
        self.config = RunConfig()

        self._phases = {}
        self._element_phases = {}
        self._all_items = []
        self._filtered_items = []
        self._param_cache = {}
        self._rooms = []
        self._loading = True

        self._load_xaml()
        self._init_phases()
        self._init_scope()
        self._init_categories()
        self._loading = False

        self._refresh_rooms()
        self._refresh_target_params()
        self._update_category_count()
        self._restore_settings()

    # ------------------------------------------------------------------ setup

    def _load_xaml(self):
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

        self._find_controls(root)
        self._wire_events()

    def _find_controls(self, root):
        self.txt_room_source = root.FindName('txt_room_source')
        self.cbo_phase = root.FindName('cbo_phase')
        self.cbo_element_phase = root.FindName('cbo_element_phase')
        self.rdo_scope_all = root.FindName('rdo_scope_all')
        self.rdo_scope_view = root.FindName('rdo_scope_view')
        self.chk_elements_view = root.FindName('chk_elements_view')
        self.txt_room_count = root.FindName('txt_room_count')

        self.txt_search = root.FindName('txt_search')
        self.btn_clear_search = root.FindName('btn_clear_search')
        self.lst_categories = root.FindName('lst_categories')
        self.thumb_resize = root.FindName('thumb_resize')
        self.btn_select_all = root.FindName('btn_select_all')
        self.btn_select_none = root.FindName('btn_select_none')
        self.txt_category_count = root.FindName('txt_category_count')

        self.cbo_source = root.FindName('cbo_source')
        self.cbo_target = root.FindName('cbo_target')
        self.txt_target_hint = root.FindName('txt_target_hint')
        self.txt_separator = root.FindName('txt_separator')
        self.chk_overwrite = root.FindName('chk_overwrite')

        self.txt_tol_x = root.FindName('txt_tol_x')
        self.txt_tol_y = root.FindName('txt_tol_y')
        self.txt_tol_zup = root.FindName('txt_tol_zup')
        self.txt_tol_zdown = root.FindName('txt_tol_zdown')
        self.rdo_axes_global = root.FindName('rdo_axes_global')
        self.rdo_axes_local = root.FindName('rdo_axes_local')
        self.chk_retry_level = root.FindName('chk_retry_level')

        self.btn_ok = root.FindName('btn_ok')
        self.btn_cancel = root.FindName('btn_cancel')

    def _wire_events(self):
        self.cbo_phase.SelectionChanged += self.OnPhaseChanged
        self.rdo_scope_all.Checked += self.OnScopeChanged
        self.rdo_scope_view.Checked += self.OnScopeChanged
        self.txt_search.TextChanged += self.OnSearchTextChanged
        self.btn_clear_search.Click += self.OnClearSearch
        self.btn_select_all.Click += self.OnSelectAll
        self.btn_select_none.Click += self.OnSelectNone
        self.thumb_resize.DragDelta += self.OnResizeList
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel

    def _active_view_phase_name(self):
        try:
            view_phase_param = self.doc.ActiveView.get_Parameter(BIP.VIEW_PHASE)
            if view_phase_param is not None:
                view_phase = self.doc.GetElement(view_phase_param.AsElementId())
                if view_phase is not None:
                    return view_phase.Name
        except Exception:
            pass
        return None

    @staticmethod
    def _fill_phase_combo(combo, document, default_name):
        """Riempie il combo con le fasi del documento; {nome: Phase}.

        Default: la fase con il nome dato (quella della vista attiva, per nome
        perche' in modalita' link le fasi sono di un altro documento), altrimenti
        l'ultima della sequenza.
        """
        lookup = {}
        combo.Items.Clear()
        for phase in document.Phases:
            lookup[phase.Name] = phase
            combo.Items.Add(phase.Name)
        if default_name in lookup:
            combo.SelectedItem = default_name
        elif combo.Items.Count:
            combo.SelectedIndex = combo.Items.Count - 1
        return lookup

    def _init_phases(self):
        view_phase_name = self._active_view_phase_name()
        self._phases = self._fill_phase_combo(self.cbo_phase, self.room_doc, view_phase_name)
        self._element_phases = self._fill_phase_combo(
            self.cbo_element_phase, self.doc, view_phase_name)

    def _init_scope(self):
        if self.link_mode:
            self.txt_room_source.Text = u"Rooms from link: {}".format(self.link_instance.Name)
        else:
            self.txt_room_source.Text = u"Rooms from the current document."

        self.rdo_scope_all.IsChecked = True
        self.chk_elements_view.IsChecked = False
        if not view_scope_supported(self.doc):
            not_graphical = "The active view is not a graphical model view."
            self.rdo_scope_view.IsEnabled = False
            self.rdo_scope_view.ToolTip = not_graphical
            self.chk_elements_view.IsEnabled = False
            self.chk_elements_view.ToolTip = not_graphical
        if self.link_mode:
            self.rdo_scope_view.IsEnabled = False
            self.rdo_scope_view.ToolTip = (
                "Not available when rooms come from a linked model.")
        self.chk_overwrite.IsChecked = True
        self.chk_retry_level.IsChecked = True
        self.rdo_axes_global.IsChecked = True

    def _init_categories(self):
        self._all_items = [CategoryItem(category)
                           for category in model_categories(self.doc)]
        self._filtered_items = self._all_items[:]
        self._populate_listbox()

    # ------------------------------------------------------------- categorie

    def _populate_listbox(self):
        self.lst_categories.Items.Clear()
        for item in self._filtered_items:
            checkbox = CheckBox()
            checkbox.Content = item.Name
            checkbox.IsChecked = item.IsSelected
            checkbox.Tag = item
            checkbox.Checked += self.OnCheckboxChanged
            checkbox.Unchecked += self.OnCheckboxChanged
            self.lst_categories.Items.Add(checkbox)

    def _filter_categories(self, search_text):
        if not search_text:
            self._filtered_items = self._all_items[:]
        else:
            needle = search_text.lower()
            self._filtered_items = [item for item in self._all_items
                                    if needle in item.Name.lower()]
        self._populate_listbox()

    def _selected_items(self):
        return [item for item in self._all_items if item.IsSelected]

    def _update_category_count(self):
        selected = self._selected_items()
        self.txt_category_count.Foreground = GRAY_BRUSH
        self.txt_category_count.Text = "{} of {} categories selected.".format(
            len(selected), len(self._all_items))

    # ------------------------------------------------------------------ room

    def _current_phase(self):
        name = self.cbo_phase.SelectedItem
        return self._phases.get(name) if name else None

    def _current_element_phase(self):
        name = self.cbo_element_phase.SelectedItem
        return self._element_phases.get(name) if name else None

    def _refresh_rooms(self):
        phase = self._current_phase()
        only_view = bool(self.rdo_scope_view.IsChecked) and not self.link_mode
        view_id = self.doc.ActiveView.Id if only_view else None
        self._rooms = collect_rooms(self.room_doc, phase, view_id)

        if self._rooms:
            self.txt_room_count.Foreground = GRAY_BRUSH
            self.txt_room_count.Text = "{} rooms in the selected phase.".format(len(self._rooms))
        else:
            self.txt_room_count.Foreground = RED_BRUSH
            if only_view:
                self.txt_room_count.Text = (
                    "No rooms: the Rooms category may be turned off in the active "
                    "view's Visibility/Graphics, or no room belongs to this phase.")
            else:
                self.txt_room_count.Text = "No placed rooms in this phase."

        self._refresh_source_params()

    def _refresh_source_params(self):
        previous = self.cbo_source.SelectedItem
        names = room_param_names(self._rooms)
        self.cbo_source.Items.Clear()
        for name in names:
            self.cbo_source.Items.Add(name)
        if previous in names:
            self.cbo_source.SelectedItem = previous
        elif "Name" in names:
            self.cbo_source.SelectedItem = "Name"
        elif names:
            self.cbo_source.SelectedIndex = 0

    # -------------------------------------------------------- parametri target

    def _category_sample(self, category_id):
        """Parametri di testo scrivibili sui primi SAMPLE_PER_CATEGORY elementi
        della categoria, memorizzati. None se la categoria non ha istanze.

        I FilteredElementCollector sono lazy: iterando e fermandosi al decimo
        elemento non si scandisce il modello intero.
        """
        key = element_id_value(category_id)
        if key in self._param_cache:
            return self._param_cache[key]

        sample = []
        try:
            collector = DB.FilteredElementCollector(self.doc)\
                .OfCategoryId(category_id)\
                .WhereElementIsNotElementType()
            for element in collector:
                sample.append(element)
                if len(sample) >= SAMPLE_PER_CATEGORY:
                    break
        except Exception:
            sample = []

        names = writable_text_param_names(sample) if sample else None
        self._param_cache[key] = names
        return names

    def _refresh_target_params(self):
        previous = self.cbo_target.SelectedItem
        selected = self._selected_items()

        # Intersezione: il parametro deve esistere su tutte le categorie scelte.
        # Una categoria senza istanze nel modello non ha nulla da scrivere e non
        # svuota la lista.
        names = None
        for item in selected:
            found = self._category_sample(item.CategoryId)
            if found is None:
                continue
            names = set(found) if names is None else names & found
        names = sorted(names or [])

        self.cbo_target.Items.Clear()
        for name in names:
            self.cbo_target.Items.Add(name)

        if previous in names:
            self.cbo_target.SelectedItem = previous
        elif names:
            self.cbo_target.SelectedIndex = 0

        if not selected:
            self.txt_target_hint.Text = "Select at least one category."
        elif not names:
            self.txt_target_hint.Text = (
                "No writable text instance parameter common to all the selected "
                "categories: a project or shared parameter of type Text bound to "
                "each of them is required.")
        else:
            self.txt_target_hint.Text = ""

    # ------------------------------------------------------------------ eventi

    def OnPhaseChanged(self, sender, args):
        if self._loading:
            return
        self._refresh_rooms()

    def OnScopeChanged(self, sender, args):
        if self._loading:
            return
        self._refresh_rooms()

    def OnSearchTextChanged(self, sender, args):
        self._filter_categories(self.txt_search.Text)

    def OnClearSearch(self, sender, args):
        self.txt_search.Text = ""

    def OnCheckboxChanged(self, sender, args):
        if self._loading or sender is None or sender.Tag is None:
            return
        sender.Tag.IsSelected = bool(sender.IsChecked)
        self._refresh_target_params()
        self._update_category_count()

    def _set_all(self, value):
        # Le spunte si aggiornano sotto _loading: altrimenti ogni CheckBox
        # farebbe scattare un ricalcolo dei parametri di destinazione.
        self._loading = True
        try:
            for item in self._filtered_items:
                item.IsSelected = value
            for checkbox in self.lst_categories.Items:
                if hasattr(checkbox, 'IsChecked'):
                    checkbox.IsChecked = value
        finally:
            self._loading = False
        self._refresh_target_params()
        self._update_category_count()

    def OnSelectAll(self, sender, args):
        self._set_all(True)

    def OnSelectNone(self, sender, args):
        self._set_all(False)

    def OnResizeList(self, sender, args):
        new_height = self.lst_categories.Height + args.VerticalChange
        if 80 <= new_height <= 500:
            self.lst_categories.Height = new_height

    # -------------------------------------------------------------- validazione

    def _tolerance_fields(self):
        return (("x", self.txt_tol_x, u"± X"),
                ("y", self.txt_tol_y, u"± Y"),
                ("zup", self.txt_tol_zup, u"Z up"),
                ("zdown", self.txt_tol_zdown, u"Z down"))

    def _validate_input(self):
        if not self._current_phase():
            return False, u"Select the room phase."
        if not self._current_element_phase():
            return False, u"Select the element phase."
        if not self._rooms:
            return False, u"No room available with the selected phase and scope."
        if not self._selected_items():
            return False, u"Select at least one category to process."
        if not self.cbo_source.SelectedItem:
            return False, u"Select the room parameter to read."
        if not self.cbo_target.SelectedItem:
            return False, u"Select the element parameter to write to."

        for _, textbox, label in self._tolerance_fields():
            value = parse_number(textbox.Text)
            if value is None:
                return False, u"Tolerance '{}' is not a number.".format(label)
            if value < 0:
                return False, u"Tolerance '{}' cannot be negative.".format(label)

        return True, u""

    def _build_config(self):
        config = self.config
        config.phase = self._current_phase()
        config.element_phase = self._current_element_phase()
        config.only_active_view = bool(self.rdo_scope_view.IsChecked) and not self.link_mode
        config.only_view_elements = bool(self.chk_elements_view.IsChecked)
        config.room_doc = self.room_doc
        config.link_instance = self.link_instance
        config.rooms = self._rooms
        config.categories = [item.Category for item in self._selected_items()]
        config.source_name = self.cbo_source.SelectedItem
        config.target_name = self.cbo_target.SelectedItem
        config.separator = self.txt_separator.Text or ";"
        config.overwrite = bool(self.chk_overwrite.IsChecked)
        config.retry_at_level = bool(self.chk_retry_level.IsChecked)
        config.local_axes = bool(self.rdo_axes_local.IsChecked)

        for key, textbox, _ in self._tolerance_fields():
            centimeters = parse_number(textbox.Text) or 0.0
            config.tol_cm[key] = centimeters
            config.tol_ft[key] = cm_to_internal(centimeters, self.doc) if centimeters > 0 else 0.0

    # ------------------------------------------------------------ persistenza

    def _save_settings(self):
        try:
            cfg = script.get_config(CONFIG_SECTION)
            cfg.last_phase = self.cbo_phase.SelectedItem
            cfg.last_element_phase = self.cbo_element_phase.SelectedItem
            cfg.last_scope_view = bool(self.rdo_scope_view.IsChecked)
            cfg.last_elements_view = bool(self.chk_elements_view.IsChecked)
            cfg.last_categories = [item.Name for item in self._selected_items()]
            cfg.last_source = self.cbo_source.SelectedItem
            cfg.last_target = self.cbo_target.SelectedItem
            cfg.last_separator = self.txt_separator.Text
            cfg.last_overwrite = bool(self.chk_overwrite.IsChecked)
            cfg.last_retry = bool(self.chk_retry_level.IsChecked)
            cfg.last_local_axes = bool(self.rdo_axes_local.IsChecked)
            cfg.last_tol_x = self.txt_tol_x.Text
            cfg.last_tol_y = self.txt_tol_y.Text
            cfg.last_tol_zup = self.txt_tol_zup.Text
            cfg.last_tol_zdown = self.txt_tol_zdown.Text
            script.save_config()
        except Exception:
            pass

    def _restore_settings(self):
        try:
            cfg = script.get_config(CONFIG_SECTION)
        except Exception:
            return

        self._loading = True
        try:
            phase_name = cfg.get_option('last_phase', None)
            if phase_name and phase_name in self._phases:
                self.cbo_phase.SelectedItem = phase_name
            element_phase_name = cfg.get_option('last_element_phase', None)
            if element_phase_name and element_phase_name in self._element_phases:
                self.cbo_element_phase.SelectedItem = element_phase_name

            if cfg.get_option('last_scope_view', False) and self.rdo_scope_view.IsEnabled:
                self.rdo_scope_view.IsChecked = True
            if cfg.get_option('last_elements_view', False) and self.chk_elements_view.IsEnabled:
                self.chk_elements_view.IsChecked = True

            self.txt_separator.Text = cfg.get_option('last_separator', ';') or ';'
            self.chk_overwrite.IsChecked = bool(cfg.get_option('last_overwrite', True))
            self.chk_retry_level.IsChecked = bool(cfg.get_option('last_retry', True))
            if cfg.get_option('last_local_axes', False):
                self.rdo_axes_local.IsChecked = True
            self.txt_tol_x.Text = str(cfg.get_option('last_tol_x', '0'))
            self.txt_tol_y.Text = str(cfg.get_option('last_tol_y', '0'))
            self.txt_tol_zup.Text = str(cfg.get_option('last_tol_zup', '0'))
            self.txt_tol_zdown.Text = str(cfg.get_option('last_tol_zdown', '0'))

            saved_categories = cfg.get_option('last_categories', None)
            if saved_categories:
                wanted = set(saved_categories)
                for item in self._all_items:
                    item.IsSelected = item.Name in wanted
                self._populate_listbox()
        except Exception:
            pass
        finally:
            self._loading = False

        try:
            self._refresh_rooms()
            self._refresh_target_params()
            self._update_category_count()
            source_name = cfg.get_option('last_source', None)
            if source_name and source_name in self.cbo_source.Items:
                self.cbo_source.SelectedItem = source_name
            target_name = cfg.get_option('last_target', None)
            if target_name and target_name in self.cbo_target.Items:
                self.cbo_target.SelectedItem = target_name
        except Exception:
            pass

    # ------------------------------------------------------------------- esito

    def OnOK(self, sender, args):
        is_valid, error_message = self._validate_input()
        if not is_valid:
            MessageBox.Show(error_message, "Elements in Room")
            return

        self._build_config()
        self._save_settings()

        self.result = True
        self.Close()

    def OnCancel(self, sender, args):
        self.result = False
        self.Close()


def show_config_form(doc, room_doc=None, link_instance=None):
    """Mostra la finestra e restituisce un RunConfig, oppure None se annullata.

    room_doc: documento da cui leggere le room (il documento del link in modalita'
    link); None equivale a doc. link_instance: RevitLinkInstance scelta, o None.
    """
    form = ElementsInRoomForm(doc, room_doc or doc, link_instance)
    form.ShowDialog()

    if form.result:
        return form.config
    return None
