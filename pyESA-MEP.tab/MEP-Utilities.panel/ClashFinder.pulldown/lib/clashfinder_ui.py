# -*- coding: utf-8 -*-
"""
clashfinder_ui.py - finestra principale di Clash Setup (non modale), pagina unica.

A sinistra la libreria dei test: un clic carica il test nell'editor. A destra
l'editor del test caricato, con due colonne Set A e Set B, ciascuna con:
  - sorgenti: modello aperto (selezione corrente, vista attiva, modello, nulla) e link,
    salvati nel test (i link per nome del file);
  - categorie con spunta e conteggio degli elementi che entrano nell'insieme;
  - filtri di vista del modello (ParameterFilterElement), in AND con le categorie;
  - regole personalizzate sui parametri.
In fondo Refresh, Save as (nome in una finestra a parte), Close e Run, che esegue il
test a schermo, anche con modifiche non salvate.

La finestra resta aperta mentre si lavora in Revit: tutto cio' che legge il modello
o tocca la selezione passa da RevitBridge (ExternalEvent). Gli handler che girano
nel bridge hanno il suffisso _api e ricevono UIApplication.

QuickClashForm (clashfinder_quick_ui) eredita da questa classe e ne sostituisce
_init_controls, _load_editor, _read_editor_raw e _show_summary.
"""

import os
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System import Object
from System.IO import FileStream, FileMode
from System.Windows import (Window, MessageBox, MessageBoxButton, MessageBoxImage,
                            MessageBoxResult, Thickness, GridLength, GridUnitType,
                            FontWeights, TextWrapping, HorizontalAlignment,
                            VerticalAlignment)
from System.Windows.Controls import (CheckBox, ComboBox, TextBox, TextBlock, Button, Grid,
                                     ColumnDefinition, ListBoxItem, StackPanel)
from System.Windows.Input import Keyboard, ModifierKeys
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Colors

from pyrevit import DB, script, forms

import clashfinder_rules as cr
import clashfinder_engine as ce
import clashfinder_report_ui as report_ui

XAML_FILE_NAME = 'ClashFinder_form.xaml'
SAVE_XAML_FILE = 'ClashFinder_save.xaml'
CONFIG_SECTION = 'ESA_ClashFinder'
TITLE = "Clash Setup"
NEW_TEST_NAME = u"New test"

SIDES = ("a", "b")
SET_KEYS = {"a": "set_a", "b": "set_b"}
MATCH_LABELS = [(cr.MATCH_ALL, "All rules"), (cr.MATCH_ANY, "Any rule")]
SOURCE_KEYS = [key for key, _ in ce.SOURCE_LABELS]

GRAY_BRUSH = SolidColorBrush(Colors.Gray)
GRAY_BRUSH.Freeze()
RED_BRUSH = SolidColorBrush(Colors.Firebrick)
RED_BRUSH.Freeze()


def _ask(message, buttons=MessageBoxButton.YesNoCancel, image=MessageBoxImage.Question):
    return MessageBox.Show(message, TITLE, buttons, image)


def _info(message):
    MessageBox.Show(message, TITLE, MessageBoxButton.OK, MessageBoxImage.Information)


def _warn(message):
    MessageBox.Show(message, TITLE, MessageBoxButton.OK, MessageBoxImage.Warning)


def _plural(count, word):
    return u"{} {}{}".format(count, word, u"" if count == 1 else u"s")


def _load_window(window, file_name):
    """Carica l'XAML (accanto ai moduli, in lib/ dello stack) dentro window."""
    xaml_path = os.path.join(os.path.dirname(__file__), file_name)
    if not os.path.exists(xaml_path):
        xaml_path = script.get_bundle_file(file_name)
    stream = FileStream(xaml_path, FileMode.Open)
    try:
        root = XamlReader.Load(stream)
    finally:
        stream.Close()
    window.Content = root.Content
    window.Title = root.Title
    window.Height = root.Height
    window.Width = root.Width
    window.MinHeight = root.MinHeight
    window.MinWidth = root.MinWidth
    window.WindowStartupLocation = root.WindowStartupLocation
    window.ResizeMode = root.ResizeMode
    window.ShowInTaskbar = root.ShowInTaskbar
    return root


class RuleRow(object):
    """Riga dell'editor delle regole: parametro, operatore, valore, rimuovi."""

    def __init__(self, grid, parameter, operator, value):
        self.grid = grid
        self.parameter = parameter
        self.operator = operator
        self.value = value

    def read(self):
        label = self.operator.SelectedItem
        return {"parameter": (self.parameter.Text or u"").strip(),
                "operator": cr.OPERATOR_KEYS.get(label, "equals"),
                "value": self.value.Text or u""}


class SaveDialog(Window):
    """Nome con cui salvare il test (modale sulla finestra principale)."""

    def __init__(self, owner, name, scope_text, title=None, ok_text=None):
        Window.__init__(self)
        root = _load_window(self, SAVE_XAML_FILE)
        self.Owner = owner
        self.result = None
        self.txt_name = root.FindName("txt_name")
        self.txt_name.Text = name
        root.FindName("txt_scope").Text = scope_text
        button = root.FindName("btn_ok")
        if title:
            self.Title = title
            root.FindName("txt_title").Text = title
        if ok_text:
            button.Content = ok_text
        button.Click += self.OnOK
        root.FindName("btn_cancel").Click += lambda s, e: self.Close()
        self.Loaded += lambda s, e: (self.txt_name.Focus(), self.txt_name.SelectAll())

    def OnOK(self, sender, args):
        name = (self.txt_name.Text or u"").strip()
        if not name:
            MessageBox.Show(u"Give the test a name.", TITLE)
            return
        self.result = name
        self.Close()


class ClashFinderForm(Window):

    XAML_FILE = XAML_FILE_NAME

    def __init__(self, doc, bridge):
        self.doc = doc
        self.bridge = bridge
        self.report = None

        self.library_path = cr.library_path()
        self.library = cr.load_library()
        # Nuovi test predefiniti in una libreria gia' esistente (una volta sola).
        added = cr.pending_defaults(self.library)
        if added:
            self.library.extend(added)
            try:
                cr.write_tests_file(self.library_path, self.library)
            except Exception:
                pass
        self.display_names = cr.category_display_names(doc)
        self.all_categories = cr.model_categories(doc)
        self.category_values = dict((key, cr.category_id_value(key))
                                    for key, _ in self.all_categories)
        self.links = ce.list_links(doc)
        self.used_categories = self._categories_in_use()
        self.filter_names = sorted(cr.view_filters(doc).keys(), key=lambda n: n.lower())

        # Selezione corrente di Revit, riletta da _read_selection prima di ogni uso.
        self.loaded_ids = []

        # Editor: test caricato (nome in libreria, None se non salvato) e stato a schermo.
        self.edit_name = None
        self.dirty = False
        self._editor_loading = False
        self._list_loading = False
        self.edit_test = cr.new_test(NEW_TEST_NAME)
        self.edit_cats = {"a": set(), "b": set()}
        self.edit_filters = {"a": set(), "b": set()}
        self.edit_links = {"a": set(), "b": set()}
        self.rule_rows = {"a": [], "b": []}
        self.counts = {"a": {}, "b": {}}
        self.param_names = {"a": list(cr.PSEUDO_FIELDS), "b": list(cr.PSEUDO_FIELDS)}
        self._param_cache = {}
        self._param_pending = {"a": False, "b": False}
        self._summary_pending = False
        # Ultima casella cliccata per elenco, per lo SHIFT + clic su un intervallo.
        self._last_clicked = {}

        Window.__init__(self)
        self._root = _load_window(self, self.XAML_FILE)
        self._find_controls()
        self._init_controls()

        last = self._restore_settings()
        self._fill_tests()
        test = self._library_test(last) or (self.library[0] if self.library else None)
        if test is not None:
            self._select_in_list(test["name"])
            self._load_editor(cr.copy_test(test), test["name"])
        else:
            self._load_editor(cr.new_test(NEW_TEST_NAME), None)

        # Il costruttore gira nel contesto API dello script: prima lettura diretta.
        self._refresh_summary_api(None)

    # ------------------------------------------------------------------ setup

    def _find_controls(self):
        names = ("lst_tests", "btn_up", "btn_down", "btn_new", "btn_delete", "btn_import",
                 "btn_export",
                 "txt_library_path", "pnl_editor", "txt_editor_state",
                 "txt_summary_state", "txt_links_hint",
                 "btn_refresh", "btn_save", "btn_save_as", "btn_close", "btn_run")
        for name in names:
            setattr(self, name, self._root.FindName(name))
        for side in SIDES:
            for name in ("cbo_source_{}", "lst_links_{}",
                         "txt_search_{}", "btn_clear_search_{}", "lst_cats_{}", "thumb_{}",
                         "chk_used_{}", "chk_ticked_{}", "btn_none_{}",
                         "txt_filter_search_{}", "btn_clear_filter_search_{}",
                         "lst_filters_{}", "cbo_match_{}", "pnl_rules_{}", "btn_add_rule_{}"):
                setattr(self, name.format(side), self._root.FindName(name.format(side)))

    def _init_controls(self):
        for side in SIDES:
            combo = getattr(self, "cbo_source_" + side)
            for _, label in ce.SOURCE_LABELS:
                combo.Items.Add(label)
            combo.SelectedIndex = 0
            combo.SelectionChanged += lambda s, e: self._changed()
            combo = getattr(self, "cbo_match_" + side)
            for _, label in MATCH_LABELS:
                combo.Items.Add(label)
            combo.SelectedIndex = 0
            getattr(self, "chk_used_" + side).IsChecked = True
            self._wire_side(side)

        self.txt_library_path.Text = u"Library: {}".format(self.library_path)
        self.txt_library_path.ToolTip = self.library_path

        self.lst_tests.SelectionChanged += self.OnTestSelected
        self.btn_up.Click += lambda s, e: self._move_test(-1)
        self.btn_down.Click += lambda s, e: self._move_test(1)
        self.btn_new.Click += self.OnNewTest
        self.btn_delete.Click += self.OnDeleteTest
        self.btn_import.Click += self.OnImport
        self.btn_export.Click += self.OnExport

        self.btn_refresh.Click += lambda s, e: self._request_summary()
        self.btn_save.Click += lambda s, e: self._save()
        self.btn_save_as.Click += lambda s, e: self._save_as()
        self.btn_close.Click += lambda s, e: self.Close()
        self.btn_run.Click += self.OnRun
        self.Closing += self.OnWindowClosing
        self.Closed += self.OnWindowClosed

    def _wire_side(self, side):
        search = getattr(self, "txt_search_" + side)
        search.TextChanged += lambda s, e: self._fill_categories(side)
        getattr(self, "btn_clear_search_" + side).Click += \
            lambda s, e: setattr(search, "Text", u"")
        getattr(self, "chk_used_" + side).Click += lambda s, e: self._fill_categories(side)
        getattr(self, "chk_ticked_" + side).Click += lambda s, e: self._fill_categories(side)
        getattr(self, "btn_none_" + side).Click += lambda s, e: self._clear_categories(side)
        filter_search = getattr(self, "txt_filter_search_" + side)
        filter_search.TextChanged += lambda s, e: self._fill_filters(side)
        getattr(self, "btn_clear_filter_search_" + side).Click += \
            lambda s, e: setattr(filter_search, "Text", u"")
        getattr(self, "thumb_" + side).DragDelta += \
            lambda s, e: self._resize_list(getattr(self, "lst_cats_" + side), e)
        getattr(self, "cbo_match_" + side).SelectionChanged += lambda s, e: self._changed()
        getattr(self, "btn_add_rule_" + side).Click += \
            lambda s, e: (self._add_rule_row(side, None), self._mark_dirty())

    def _categories_in_use(self):
        """Categorie con almeno un elemento nel modello aperto o in un link caricato."""
        used = set()
        docs = [self.doc] + [link.doc for link in self.links]
        for key, _ in self.all_categories:
            bic = cr.builtin_category(key)
            if bic is None:
                continue
            for document in docs:
                collector = None
                try:
                    collector = DB.FilteredElementCollector(document).OfCategory(bic)\
                        .WhereElementIsNotElementType()
                    if collector.FirstElement() is not None:
                        used.add(key)
                        break
                except Exception:
                    pass
                finally:
                    if collector is not None:
                        collector.Dispose()
        return used

    # ------------------------------------------------------------ contesto API

    def _uidoc_api(self, uiapp):
        """UIDocument del modello della finestra, o None (con avviso) se non e' attivo."""
        try:
            uidoc = uiapp.ActiveUIDocument
            if (uidoc is not None and self.doc.IsValidObject
                    and uidoc.Document.Equals(self.doc)):
                return uidoc
        except Exception:
            pass
        _warn(u"This window works on the model it was opened on. "
              u"Activate that model and try again, or close this window and open "
              u"it again from the other model.")
        return None

    def _read_selection(self, uidoc):
        """Selezione corrente di Revit: sorgente "Current selection" degli insiemi."""
        self.loaded_ids = list(uidoc.Selection.GetElementIds())

    def _refresh_links(self):
        """Rilegge i link (in contesto API): un link ricaricato ha un nuovo Document."""
        links = ce.list_links(self.doc)
        changed = [l.key for l in links] != [l.key for l in self.links]
        self.links = links
        if changed:
            self._fill_link_lists()

    # ================================================================ LIBRERIA

    def _library_test(self, name):
        if not name:
            return None
        for test in self.library:
            if test["name"].lower() == name.lower():
                return test
        return None

    def _fill_tests(self):
        """Elenco dei test: nome in grassetto, sotto le sorgenti di A e B."""
        self._list_loading = True
        try:
            self.lst_tests.Items.Clear()
            for test in self.library:
                panel = StackPanel()
                title = TextBlock()
                title.Text = test["name"]
                title.FontWeight = FontWeights.Bold
                title.TextWrapping = TextWrapping.Wrap
                scope = TextBlock()
                scope.Text = cr.describe_scope(test)
                scope.Foreground = GRAY_BRUSH
                scope.FontSize = 11
                scope.TextWrapping = TextWrapping.Wrap
                panel.Children.Add(title)
                panel.Children.Add(scope)
                item = ListBoxItem()
                item.Content = panel
                item.Tag = test["name"]
                item.ToolTip = u"A: {}\nB: {}".format(
                    cr.describe_set(test["set_a"], self.display_names),
                    cr.describe_set(test["set_b"], self.display_names))
                item.Padding = Thickness(2, 3, 2, 3)
                self.lst_tests.Items.Add(item)
        finally:
            self._list_loading = False

    def _select_in_list(self, name):
        self._list_loading = True
        try:
            self.lst_tests.SelectedItem = None
            for item in self.lst_tests.Items:
                if name and item.Tag.lower() == name.lower():
                    self.lst_tests.SelectedItem = item
                    break
        finally:
            self._list_loading = False

    def _save_library(self):
        try:
            cr.write_tests_file(self.library_path, self.library)
            return True
        except Exception as error:
            _warn(u"The test library could not be saved:\n{}\n\n{}".format(
                self.library_path, error))
            return False

    def OnTestSelected(self, sender, args):
        if self._list_loading:
            return
        item = self.lst_tests.SelectedItem
        if item is None:
            return
        name = item.Tag
        if self.edit_name is not None and name.lower() == self.edit_name.lower():
            return
        if not self._resolve_dirty():
            self._select_in_list(self.edit_name)
            return
        test = self._library_test(name)
        if test is not None:
            self._select_in_list(test["name"])
            self._load_editor(cr.copy_test(test), test["name"])

    def _resolve_dirty(self):
        """True se si puo' proseguire: nessuna modifica, salvata o scartata."""
        if not self.dirty:
            return True
        answer = _ask(u"The test on screen has unsaved changes.\n\n"
                      u"Yes: save them{}.\nNo: discard them.".format(
                          u" to '{}'".format(self.edit_name) if self.edit_name
                          else u" with a name"))
        if answer == MessageBoxResult.Cancel:
            return False
        if answer == MessageBoxResult.Yes:
            return self._save()
        return True

    def OnNewTest(self, sender, args):
        """Chiede il nome e aggiunge subito alla libreria un test vuoto: nessuna
        categoria, link o filtro, entrambi gli insiemi dalla selezione corrente."""
        if not self._resolve_dirty():
            return
        names = [t["name"] for t in self.library]
        test = cr.new_test(cr.unique_name(NEW_TEST_NAME, names))
        for key in SET_KEYS.values():
            test[key]["source"] = cr.SOURCE_SELECTION
        dialog = SaveDialog(self, test["name"], cr.describe_scope(test),
                            title=u"➕ New Test", ok_text=u"Create")
        dialog.ShowDialog()
        name = dialog.result
        if not name:
            return
        if self._library_test(name) is not None:
            _warn(u"A test called '{}' already exists: choose another name.".format(name))
            return
        test["name"] = name
        self.library.append(test)
        if not self._save_library():
            self.library.pop()
            return
        self._fill_tests()
        self._select_in_list(name)
        self._load_editor(cr.copy_test(test), name)

    def _move_test(self, step):
        """Sposta di una posizione il test selezionato; l'ordine si salva subito (vale
        anche per Quick Clash). L'editor non cambia."""
        item = self.lst_tests.SelectedItem
        if item is None:
            _info(u"Select the test to move in the list.")
            return
        names = [t["name"] for t in self.library]
        index = names.index(item.Tag) if item.Tag in names else -1
        target = index + step
        if index < 0 or not 0 <= target < len(self.library):
            return
        self.library[index], self.library[target] = self.library[target], self.library[index]
        self._save_library()
        self._fill_tests()
        self._select_in_list(item.Tag)

    def OnDeleteTest(self, sender, args):
        item = self.lst_tests.SelectedItem
        if item is None:
            _info(u"Select the test to delete in the list.")
            return
        name = item.Tag
        if _ask(u"Delete the test '{}' from the library?".format(name),
                MessageBoxButton.YesNo) != MessageBoxResult.Yes:
            return
        self.library = [t for t in self.library if t["name"] != name]
        self._save_library()
        self._fill_tests()
        if self.edit_name == name:
            # Il test resta a schermo, ma non e' piu' in libreria.
            self.edit_name = None
            self.dirty = True
            self._update_editor_state()

    def OnImport(self, sender, args):
        path = forms.pick_file(file_ext='json', title="Import clash tests")
        if not path:
            return
        try:
            tests = cr.read_tests_file(path)
        except Exception as error:
            _warn(u"The file could not be read as clash tests:\n{}\n\n{}".format(path, error))
            return
        if not tests:
            _info(u"The file contains no clash test.")
            return
        existing = dict((t["name"].lower(), i) for i, t in enumerate(self.library))
        clashing = [t for t in tests if t["name"].lower() in existing]
        replace = False
        if clashing:
            answer = _ask(u"{} of the imported tests have the same name as a saved test.\n\n"
                          u"Yes: replace the saved tests.\n"
                          u"No: keep both (the imported ones are renamed).".format(len(clashing)))
            if answer == MessageBoxResult.Cancel:
                return
            replace = answer == MessageBoxResult.Yes
        for test in tests:
            index = existing.get(test["name"].lower())
            if index is not None and replace:
                self.library[index] = test
            else:
                test["name"] = cr.unique_name(test["name"], [t["name"] for t in self.library])
                self.library.append(test)
        self._save_library()
        self._fill_tests()
        self._select_in_list(self.edit_name)
        _info(u"{} imported.".format(_plural(len(tests), u"test")))

    def OnExport(self, sender, args):
        if not self.library:
            _info(u"There is no test to export.")
            return
        path = forms.save_file(file_ext='json', default_name='ClashTests',
                               title="Export clash tests")
        if not path:
            return
        try:
            cr.write_tests_file(path, self.library)
        except Exception as error:
            _warn(u"The file could not be written:\n{}\n\n{}".format(path, error))
            return
        _info(u"{} exported to:\n{}".format(_plural(len(self.library), u"test"), path))

    def _save(self):
        """Sovrascrive il test caricato; per un test mai salvato chiede il nome."""
        if self.edit_name is None or self._library_test(self.edit_name) is None:
            return self._save_as()
        raw = self._read_editor_raw()
        valid, message = self._validate_rules(raw)
        if not valid:
            _warn(message)
            return False
        raw["name"] = self.edit_name
        test = cr.normalize_test(raw)
        for index, other in enumerate(self.library):
            if other["name"].lower() == self.edit_name.lower():
                self.library[index] = test
                break
        if not self._save_library():
            return False
        self.edit_test = cr.copy_test(test)
        self.dirty = False
        self._fill_tests()
        self._select_in_list(self.edit_name)
        self._update_editor_state()
        return True

    def _save_as(self):
        """Chiede il nome e salva il test a schermo. True se salvato."""
        raw = self._read_editor_raw()
        valid, message = self._validate_rules(raw)
        if not valid:
            _warn(message)
            return False
        dialog = SaveDialog(self, self.edit_name or raw["name"], cr.describe_scope(
            cr.normalize_test(dict(raw, name=u"-"))))
        dialog.ShowDialog()
        name = dialog.result
        if not name:
            return False
        existing = self._library_test(name)
        same = self.edit_name is not None and name.lower() == self.edit_name.lower()
        if existing is not None and not same:
            if _ask(u"A test called '{}' already exists. Overwrite it?".format(
                    existing["name"]), MessageBoxButton.YesNo) != MessageBoxResult.Yes:
                return False
        raw["name"] = name
        test = cr.normalize_test(raw)
        replaced = False
        for index, other in enumerate(self.library):
            if other["name"].lower() == name.lower() or (
                    same and other["name"].lower() == self.edit_name.lower()):
                self.library[index] = test
                replaced = True
                break
        if not replaced:
            self.library.append(test)
        if not self._save_library():
            return False
        self.edit_name = name
        self.edit_test = cr.copy_test(test)
        self.dirty = False
        self._fill_tests()
        self._select_in_list(name)
        self._update_editor_state()
        return True

    # ================================================================== EDITOR

    def _mark_dirty(self):
        if self._editor_loading:
            return
        self.dirty = True
        self._update_editor_state()

    def _changed(self):
        """Modifica che cambia il contenuto degli insiemi: segna e ricalcola."""
        if self._editor_loading:
            return
        self._mark_dirty()
        self._request_summary()

    def _update_editor_state(self):
        """Nome del test caricato nel titolo della finestra, modifiche nella riga di stato."""
        self.Title = u"{} - {}{}".format(TITLE, self.edit_name or u"new test (not saved)",
                                         u" *" if self.dirty else u"")
        self.txt_editor_state.Text = u"Unsaved changes" if self.dirty else u""

    def _load_editor(self, test, name):
        self._editor_loading = True
        try:
            self.edit_test = test
            self.edit_name = name
            for side in SIDES:
                set_data = test[SET_KEYS[side]]
                getattr(self, "cbo_source_" + side).SelectedIndex = \
                    SOURCE_KEYS.index(set_data["source"])
                self.edit_links[side] = set(set_data["links"])
                self.edit_cats[side] = set(set_data["categories"])
                self.edit_filters[side] = set(set_data["view_filters"])
                self.counts[side] = {}
                getattr(self, "cbo_match_" + side).SelectedIndex = \
                    1 if set_data["match"] == cr.MATCH_ANY else 0
                getattr(self, "pnl_rules_" + side).Children.Clear()
                self.rule_rows[side] = []
                for rule in set_data["rules"]:
                    self._add_rule_row(side, rule)
                self._fill_categories(side)
                self._fill_filters(side)
            self._fill_link_lists()
        finally:
            self._editor_loading = False
        self.dirty = False
        self._update_editor_state()
        for side in SIDES:
            self._request_param_names(side)
        self._request_summary()

    def _read_editor_raw(self):
        """Test a schermo, cosi' come e' (regole incomplete comprese)."""
        test = {"name": self.edit_name or NEW_TEST_NAME,
                "ignore_connected": True, "ignore_hosted": True}
        for side in SIDES:
            rules = [row.read() for row in self.rule_rows[side]]
            rules = [r for r in rules if r["parameter"] or r["value"]]
            match_index = getattr(self, "cbo_match_" + side).SelectedIndex
            # Ordine delle categorie: quello dell'elenco, poi le chiavi sconosciute.
            ordered = [key for key, _ in self.all_categories if key in self.edit_cats[side]]
            ordered += sorted(k for k in self.edit_cats[side] if k not in ordered)
            test[SET_KEYS[side]] = {
                "source": self._source(side),
                "links": sorted(self.edit_links[side], key=lambda n: n.lower()),
                "categories": ordered,
                "view_filters": sorted(self.edit_filters[side], key=lambda n: n.lower()),
                "rules": rules,
                "match": MATCH_LABELS[max(match_index, 0)][0]}
        return test

    def _current_test(self):
        """Test a schermo normalizzato (le regole incomplete sono tralasciate)."""
        return cr.normalize_test(self._read_editor_raw())

    def _source(self, side):
        index = getattr(self, "cbo_source_" + side).SelectedIndex
        return SOURCE_KEYS[index] if index >= 0 else cr.SOURCE_SELECTION

    @staticmethod
    def _validate_rules(test):
        for side in SIDES:
            label = side.upper()
            for number, rule in enumerate(test[SET_KEYS[side]]["rules"], 1):
                if not rule["parameter"]:
                    return False, u"Set {}, rule {}: choose or type a parameter.".format(
                        label, number)
                if rule["operator"] in cr.VALUELESS_OPERATORS:
                    continue
                if not rule["value"].strip():
                    return False, u"Set {}, rule {}: type a value.".format(label, number)
                if rule["operator"] in ("greater", "less") and cr.number_of(rule["value"]) is None:
                    return False, u"Set {}, rule {}: '{}' needs a number.".format(
                        label, number, cr.OPERATOR_LABELS[rule["operator"]])
        return True, u""

    # Link ----------------------------------------------------------------------

    def _link_titles(self):
        titles = []
        for link in self.links:
            title = ce.link_title(link)
            if title not in titles:
                titles.append(title)
        return sorted(titles, key=lambda n: n.lower())

    def _fill_link_lists(self):
        """Un elemento per file linkato (tutte le sue istanze); i link del test non
        caricati restano in elenco, spuntati, con la nota."""
        loaded = self._link_titles()
        loading = self._editor_loading
        self._editor_loading = True
        try:
            for side in SIDES:
                listbox = getattr(self, "lst_links_" + side)
                listbox.Items.Clear()
                titles = loaded + sorted((t for t in self.edit_links[side] if t not in loaded),
                                         key=lambda n: n.lower())
                for title in titles:
                    checkbox = CheckBox()
                    checkbox.Content = title if title in loaded else \
                        u"{} (not loaded)".format(title)
                    checkbox.Tag = title
                    checkbox.IsChecked = title in self.edit_links[side]
                    checkbox.Click += self._link_click_handler(side)
                    listbox.Items.Add(checkbox)
        finally:
            self._editor_loading = loading
        self.txt_links_hint.Text = (u"" if loaded else
                                    u"No loaded Revit link in this model: only its own "
                                    u"elements can be tested.")

    def _link_click_handler(self, side):
        def handler(sender, args):
            if sender.IsChecked:
                self.edit_links[side].add(sender.Tag)
            else:
                self.edit_links[side].discard(sender.Tag)
            self._changed()
        return handler

    # Categorie -----------------------------------------------------------------

    def _count_text(self, side, key):
        """Elementi del set per la categoria: solo per le categorie spuntate."""
        if key not in self.edit_cats[side]:
            return u""
        return u"{}".format(self.counts[side].get(self.category_values.get(key)) or 0)

    def _category_row(self, side, key, name):
        """Riga dell'elenco: casella con il nome e, a destra, il conteggio."""
        grid = Grid()
        for width in (GridLength(1, GridUnitType.Star), GridLength.Auto):
            column = ColumnDefinition()
            column.Width = width
            grid.ColumnDefinitions.Add(column)
        checkbox = CheckBox()
        checkbox.Content = name
        checkbox.IsChecked = key in self.edit_cats[side]
        count = TextBlock()
        count.Text = self._count_text(side, key)
        count.Foreground = GRAY_BRUSH
        count.HorizontalAlignment = HorizontalAlignment.Right
        count.VerticalAlignment = VerticalAlignment.Center
        count.Margin = Thickness(10, 0, 4, 0)
        checkbox.Tag = (key, count)
        checkbox.Click += lambda s, e: self._category_clicked(side, s)
        Grid.SetColumn(count, 1)
        grid.Children.Add(checkbox)
        grid.Children.Add(count)
        grid.Tag = (key, count)
        return grid

    def _fill_categories(self, side):
        listbox = getattr(self, "lst_cats_" + side)
        needle = (getattr(self, "txt_search_" + side).Text or u"").strip().lower()
        only_used = bool(getattr(self, "chk_used_" + side).IsChecked)
        only_ticked = bool(getattr(self, "chk_ticked_" + side).IsChecked)
        ticked = self.edit_cats[side]

        entries = []
        known = set()
        for key, name in self.all_categories:
            known.add(key)
            if only_used and key not in self.used_categories and key not in ticked:
                continue
            entries.append((key, name))
        # Categorie salvate nel test ma assenti da questo modello.
        for key in sorted(ticked - known):
            entries.insert(0, (key, u"{} (not in this model)".format(key)))

        listbox.Items.Clear()
        for key, name in entries:
            if needle and needle not in name.lower():
                continue
            if only_ticked and key not in ticked:
                continue
            listbox.Items.Add(self._category_row(side, key, name))

    def _refresh_category_labels(self, side):
        """Aggiorna i conteggi accanto alle categorie senza ricostruire l'elenco."""
        for row in getattr(self, "lst_cats_" + side).Items:
            key, count = row.Tag
            count.Text = self._count_text(side, key)

    def _category_clicked(self, side, checkbox):
        for box in self._shift_range("lst_cats_" + side, checkbox):
            key, count = box.Tag
            if box.IsChecked:
                self.edit_cats[side].add(key)
            else:
                self.edit_cats[side].discard(key)
            count.Text = self._count_text(side, key)
        self._changed()
        self._request_param_names(side)

    def _shift_range(self, list_name, checkbox):
        """Caselle da aggiornare dopo il clic su checkbox: con SHIFT tutte quelle
        dell'elenco (cosi' come e' filtrato) fra l'ultima cliccata e questa, portate
        allo stato di questa; altrimenti solo questa."""
        last = self._last_clicked.get(list_name)
        self._last_clicked[list_name] = checkbox
        if last is None or not (Keyboard.Modifiers & ModifierKeys.Shift):
            return [checkbox]
        boxes = []
        for item in getattr(self, list_name).Items:
            box = item.Children[0] if isinstance(item, Grid) else item
            if isinstance(box, CheckBox):
                boxes.append(box)
        positions = [i for i, box in enumerate(boxes)
                     if Object.ReferenceEquals(box, last) or Object.ReferenceEquals(box, checkbox)]
        if len(positions) != 2:
            return [checkbox]
        selected = boxes[positions[0]:positions[1] + 1]
        for box in selected:
            box.IsChecked = checkbox.IsChecked
        return selected

    def _clear_categories(self, side):
        if not self.edit_cats[side]:
            return
        self.edit_cats[side] = set()
        self._fill_categories(side)
        self._changed()
        self._request_param_names(side)

    @staticmethod
    def _resize_list(listbox, args):
        height = listbox.Height + args.VerticalChange
        if 80 <= height <= 600:
            listbox.Height = height

    # Filtri di vista -----------------------------------------------------------

    def _fill_filters(self, side):
        listbox = getattr(self, "lst_filters_" + side)
        listbox.Items.Clear()
        needle = (getattr(self, "txt_filter_search_" + side).Text or u"").strip().lower()
        names = list(self.filter_names)
        names += sorted((n for n in self.edit_filters[side] if n not in self.filter_names),
                        key=lambda n: n.lower())
        for name in names:
            if needle and needle not in name.lower():
                continue
            checkbox = CheckBox()
            checkbox.Content = name if name in self.filter_names else \
                u"{} (not in this model)".format(name)
            checkbox.Tag = name
            checkbox.IsChecked = name in self.edit_filters[side]
            checkbox.Click += self._filter_click_handler(side)
            listbox.Items.Add(checkbox)
        if not names:
            hint = TextBlock()
            hint.Text = u"This model has no view filters."
            hint.Foreground = GRAY_BRUSH
            hint.FontSize = 11
            listbox.Items.Add(hint)

    def _filter_click_handler(self, side):
        def handler(sender, args):
            for box in self._shift_range("lst_filters_" + side, sender):
                if box.IsChecked:
                    self.edit_filters[side].add(box.Tag)
                else:
                    self.edit_filters[side].discard(box.Tag)
            self._changed()
        return handler

    # Regole --------------------------------------------------------------------

    def _add_rule_row(self, side, rule):
        grid = Grid()
        grid.Margin = Thickness(0, 2, 0, 2)
        for width in (GridLength(1, GridUnitType.Star), GridLength(125),
                      GridLength(1, GridUnitType.Star), GridLength.Auto):
            column = ColumnDefinition()
            column.Width = width
            grid.ColumnDefinitions.Add(column)

        parameter = ComboBox()
        parameter.IsEditable = True
        parameter.Margin = Thickness(0, 0, 5, 0)
        parameter.ToolTip = u"Parameter name: pick it from the list or type it"
        for name in self.param_names[side]:
            parameter.Items.Add(name)
        parameter.Text = rule["parameter"] if rule else u""

        operator = ComboBox()
        operator.Margin = Thickness(0, 0, 5, 0)
        for _, label in cr.OPERATORS:
            operator.Items.Add(label)
        operator.SelectedItem = cr.OPERATOR_LABELS[rule["operator"] if rule else "equals"]

        value = TextBox()
        value.Padding = Thickness(5, 3, 5, 3)
        value.Text = rule["value"] if rule else u""

        remove = Button()
        remove.Content = u"✕"
        remove.Width = 25
        remove.Margin = Thickness(5, 0, 0, 0)
        remove.ToolTip = u"Remove rule"

        for column, control in enumerate((parameter, operator, value, remove)):
            Grid.SetColumn(control, column)
            grid.Children.Add(control)

        row = RuleRow(grid, parameter, operator, value)
        self.rule_rows[side].append(row)
        getattr(self, "pnl_rules_" + side).Children.Add(grid)

        def on_operator(sender, args):
            value.IsEnabled = cr.OPERATOR_KEYS.get(operator.SelectedItem) not in \
                cr.VALUELESS_OPERATORS
            self._changed()

        def on_remove(sender, args):
            self.rule_rows[side].remove(row)
            getattr(self, "pnl_rules_" + side).Children.Remove(grid)
            self._changed()

        on_operator(None, None)
        # Testo: segna la modifica a ogni tasto, ricalcola i conteggi all'uscita dal campo.
        parameter.SelectionChanged += lambda s, e: self._changed()
        parameter.KeyUp += lambda s, e: self._mark_dirty()
        parameter.LostFocus += lambda s, e: self._request_summary()
        operator.SelectionChanged += on_operator
        value.TextChanged += lambda s, e: self._mark_dirty()
        value.LostFocus += lambda s, e: self._request_summary()
        remove.Click += on_remove
        return row

    def _request_param_names(self, side):
        if self._param_pending[side]:
            return
        self._param_pending[side] = True
        self.bridge.run(lambda uiapp: self._update_param_names(side))

    def _update_param_names(self, side):
        """Elenco dei parametri per le regole del lato (in contesto API)."""
        self._param_pending[side] = False
        if not self.doc.IsValidObject:
            return
        keys = tuple(sorted(self.edit_cats[side]))
        names = self._param_cache.get(keys)
        if names is None:
            # Campione nel modello aperto e in ogni link caricato: un test puo' essere
            # usato su elementi di entrambi, e i link hanno spesso parametri propri.
            found = set(cr.parameter_names_for(self.doc, keys))
            for link in self.links:
                try:
                    if link.doc.IsValidObject:
                        found.update(cr.parameter_names_for(link.doc, keys))
                except Exception:
                    continue
            names = sorted(found, key=lambda n: n.lower())
            self._param_cache[keys] = names
        self.param_names[side] = list(cr.PSEUDO_FIELDS) + names
        for row in self.rule_rows[side]:
            current = row.parameter.Text
            loading = self._editor_loading
            self._editor_loading = True
            try:
                row.parameter.Items.Clear()
                for name in self.param_names[side]:
                    row.parameter.Items.Add(name)
                row.parameter.Text = current
            finally:
                self._editor_loading = loading

    # ================================================================ RIEPILOGO

    def _request_summary(self):
        if self._summary_pending:
            return
        self._summary_pending = True
        self.txt_summary_state.Foreground = GRAY_BRUSH
        self.txt_summary_state.Text = u"Updating..."
        self.bridge.run(self._refresh_summary_api)

    def _refresh_summary_api(self, uiapp):
        """Rilegge selezione e link e ricalcola il riepilogo del test a schermo.

        uiapp None: chiamata dal costruttore, gia' in contesto API.
        """
        self._summary_pending = False
        uidoc = revit_uidoc(self.doc) if uiapp is None else self._uidoc_api(uiapp)
        if uidoc is None:
            return
        self._read_selection(uidoc)
        self._refresh_links()
        test = self._current_test()
        summary = ce.summarize(self.doc, [test], self.loaded_ids, self.doc.ActiveView,
                               self.links)
        self._show_summary(test, summary)

    def _show_summary(self, test, summary):
        """Conteggi accanto alle categorie, totali per insieme, riga di stato."""
        empty = []
        for side in SIDES:
            self.counts[side] = summary.by_category[side]
            self._refresh_category_labels(side)
            if summary.host[side] + summary.linked[side] == 0:
                empty.append(side.upper())
        lines = [u"{} selected in Revit. {}".format(
            _plural(len(self.loaded_ids), u"element"), summary.phase_text)]
        if empty:
            lines.append(u"Set {} {} empty.".format(u" and ".join(empty),
                                                    u"is" if len(empty) == 1 else u"are"))
        lines += [note.split(u": ", 1)[-1] for note in summary.notes]
        self.txt_summary_state.Foreground = RED_BRUSH if (
            empty or summary.phase_problem) else GRAY_BRUSH
        self.txt_summary_state.Text = u"\n".join(lines)

    # ================================================================ ESECUZIONE

    def OnRun(self, sender, args):
        valid, message = self._validate_rules(self._read_editor_raw())
        if not valid:
            _warn(message)
            return
        # La selezione si controlla in _before_run: puo' essere cambiata in Revit
        # dopo l'ultimo aggiornamento del riepilogo.
        problem = self._scope_problem(self._current_test(), check_selection=False)
        if problem:
            _info(problem)
            return
        self.bridge.run(self._run_api)

    def _scope_problem(self, test, check_selection):
        """Messaggio se un insieme non puo' contenere nulla, altrimenti None."""
        loaded = set(self._link_titles())
        for side in SIDES:
            set_data = test[SET_KEYS[side]]
            label = side.upper()
            links = [t for t in set_data["links"] if t in loaded]
            if set_data["source"] == cr.SOURCE_NONE and not links:
                return (u"Set {} takes nothing from the open model and no loaded link is "
                        u"ticked for it.".format(label))
            if check_selection and set_data["source"] == cr.SOURCE_SELECTION \
                    and not self.loaded_ids and not links:
                return (u"Set {} is taken from the current selection, but nothing is "
                        u"selected in Revit.\n\nSelect elements, take the set from the "
                        u"active view or the whole model, or tick a link.".format(label))
        return None

    def _before_run(self, uidoc):
        """Rilegge la selezione; True se gli insiemi possono contenere qualcosa."""
        self._read_selection(uidoc)
        # Fase e filtro di fase della vista attiva valgono per ogni sorgente: senza, stop.
        problem = ce.phase_problem(self.doc.ActiveView) or \
            self._scope_problem(self._current_test(), check_selection=True)
        if problem:
            _info(problem)
            return False
        return True

    def _run_api(self, uiapp):
        uidoc = self._uidoc_api(uiapp)
        if uidoc is None:
            return
        self._refresh_links()
        if not self._before_run(uidoc):
            return
        test = self._current_test()

        with forms.ProgressBar(title="Clash Finder: {value} of {max_value} elements",
                               cancellable=True) as progress_bar:
            def progress(done, total):
                if done == total or done % 20 == 0:
                    progress_bar.update_progress(done, total)
                return progress_bar.cancelled

            result = ce.run_tests(self.doc, [test], self.loaded_ids, self.doc.ActiveView,
                                  self.links, progress)

        # Report nella cartella del modello, per il pulsante Last Report; con la
        # selezione usata, per il Refresh del report.
        result.selection_ids = [int(cr.get_element_id_value(i)) for i in self.loaded_ids]
        report_ui.save_run_reports(self.doc, result)

        if self.report is not None:
            try:
                self.report.Close()
            except Exception:
                pass
        self.report = report_ui.show_report(self.doc, result, self.display_names,
                                           uiapp.MainWindowHandle)

    # ================================================================ chiusura

    def OnWindowClosing(self, sender, args):
        if not self._resolve_dirty():
            args.Cancel = True

    def OnWindowClosed(self, sender, args):
        self._save_settings()
        self.bridge.dispose()

    def _save_settings(self):
        try:
            cfg = script.get_config(CONFIG_SECTION)
            cfg.last_test = self.edit_name or u""
            script.save_config()
        except Exception:
            pass

    def _restore_settings(self):
        """Nome dell'ultimo test caricato, o None."""
        try:
            cfg = script.get_config(CONFIG_SECTION)
            return cfg.get_option('last_test', None) or None
        except Exception:
            return None


def revit_uidoc(doc):
    """UIDocument del documento, se e' quello attivo."""
    try:
        from pyrevit import revit
        uidoc = revit.uidoc
        if uidoc is not None and uidoc.Document.Equals(doc):
            return uidoc
    except Exception:
        pass
    return None


def show_main_window(doc, bridge, owner_handle=None):
    form = ClashFinderForm(doc, bridge)
    if owner_handle is not None:
        try:
            from System.Windows.Interop import WindowInteropHelper
            WindowInteropHelper(form).Owner = owner_handle
        except Exception:
            pass
    form.Show()
    return form
