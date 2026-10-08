# -*- coding: utf-8 -*-
"""
mepqto_pricelist_ui.py - editor del listino comune in formato JSON.

Il listino si apre, si modifica in una griglia (capitolo, sottocapitolo, n. articolo
EPU, prezzario di riferimento, codice prezzario, descrizione, unita' da tendina,
prezzo) e si salva in un file .json (struttura in mepqto_store.PriceListDocument).
Si puo' partire da un Excel / CSV con Import.

Le modifiche si ricostruiscono al salvataggio confrontando la griglia con l'istantanea
dell'ultimo caricamento: cosi' al documento arrivano solo i codici aggiunti, cambiati o
cancellati, che sono quelli da riapplicare se nel frattempo qualcun altro ha salvato.
Come nella finestra principale, il prezzo e' una colonna stringa convertita a mano: il
binding WPF usa la cultura en-US e leggerebbe "12,5" come 125.
"""

import os

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Data')

import System
from System import Action, DBNull
from System.Data import DataTable
from System.IO import FileStream, FileMode
from System.Windows import (Window, MessageBox, MessageBoxButton, MessageBoxImage,
                            MessageBoxResult, WindowStartupLocation)
from System.Windows.Controls import DataGridEditingUnit
from System.Windows.Input import Cursors
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Colors
from System.Windows.Threading import DispatcherPriority
from Microsoft.Win32 import OpenFileDialog, SaveFileDialog

from pyrevit import script

import mepqto_store as qs
from mepqto_columns_ui import show_columns_dialog

XAML_FILE_NAME = 'MEPQTO_pricelist.xaml'
TITLE = "Price list editor"
COLUMNS = ("Chapter", "Subchapter", "EpuItem", "PriceBook", "Code", "ShortDescription",
           "Description", "Unit", "UnitPrice")
FIELD_OF = {"Chapter": "chapter", "Subchapter": "subchapter", "EpuItem": "epu_item",
            "PriceBook": "price_book", "ShortDescription": "short_description",
            "Description": "description", "Unit": "unit", "UnitPrice": "price"}
JSON_FILTER = "MEP QTO price list (*.json)|*.json"
IMPORT_FILTER = ("Price list (*.xlsx;*.xlsm;*.csv;*.json)|*.xlsx;*.xlsm;*.csv;*.json"
                 "|All files (*.*)|*.*")

GRAY_BRUSH = SolidColorBrush(Colors.Gray)
GRAY_BRUSH.Freeze()
RED_BRUSH = SolidColorBrush(Colors.Firebrick)
RED_BRUSH.Freeze()


def cell_text(value):
    if value is None or isinstance(value, DBNull):
        return u""
    return u"{}".format(value).strip()


def escape_like(text):
    out = []
    for ch in text:
        if ch in u"[]*%":
            out.append(u"[{}]".format(ch))
        elif ch == u"'":
            out.append(u"''")
        else:
            out.append(ch)
    return u"".join(out)


def item_sort_key(entry):
    code, item = entry
    chapter = (item.get("chapter") or u"").lower()
    subchapter = (item.get("subchapter") or u"").lower()
    return (chapter == u"", chapter, subchapter == u"", subchapter, code)


class PriceListEditor(Window):

    def __init__(self, path=None, import_path=None, user_name=u"", import_layout=None):
        self.saved_path = None
        self._user = user_name
        self._updating = False
        self._dirty = False
        self._document = qs.PriceListDocument(None)
        # codice -> voce all'ultimo caricamento o salvataggio
        self._snapshot = {}

        self._load_xaml()
        self._build_table()
        if path:
            if not self._open(path):
                self._new()
        else:
            self._new()
        if import_path:
            self._import(import_path, ask=False, layout=import_layout)

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
        self.ShowInTaskbar = root.ShowInTaskbar

        for name in ("txt_path", "btn_new", "btn_open", "btn_save_as", "txt_name",
                     "txt_currency", "txt_file_hint", "txt_search", "btn_clear_search",
                     "btn_import", "dg_items", "txt_count", "btn_add", "btn_duplicate",
                     "btn_remove", "txt_status", "btn_close", "btn_save"):
            setattr(self, name, root.FindName(name))

        self.btn_new.Click += self.OnNew
        self.btn_open.Click += self.OnOpen
        self.btn_save_as.Click += self.OnSaveAs
        self.btn_save.Click += self.OnSave
        self.btn_close.Click += self.OnCloseClick
        self.btn_import.Click += self.OnImport
        self.btn_add.Click += self.OnAdd
        self.btn_duplicate.Click += self.OnDuplicate
        self.btn_remove.Click += self.OnRemove
        self.btn_clear_search.Click += self.OnClearSearch
        self.txt_search.TextChanged += self.OnSearchChanged
        self.txt_name.TextChanged += self.OnMetaChanged
        self.txt_currency.TextChanged += self.OnMetaChanged
        self.Closing += self.OnWindowClosing

    def _build_table(self):
        self.table = DataTable("price_list")
        for name in COLUMNS:
            self.table.Columns.Add(name, clr.GetClrType(System.String))
        self.table.ColumnChanging += self.OnCellChanging
        self.table.ColumnChanged += self.OnCellChanged
        self.table.RowDeleted += self.OnRowDeleted
        self.dg_items.ItemsSource = self.table.DefaultView
        self.col_unit = None
        for column in self.dg_items.Columns:
            if cell_text(column.Header) == u"Unit":
                self.col_unit = column

    # ------------------------------------------------------------ documento

    def _fill(self, items):
        self._updating = True
        try:
            self.table.Rows.Clear()
            for code, item in sorted(items.items(), key=item_sort_key):
                self._add_row(code, item)
        finally:
            self._updating = False
        self._update_units()
        self._update_count()

    def _add_row(self, code, item):
        row = self.table.NewRow()
        row["Code"] = code
        row["Chapter"] = item.get("chapter") or u""
        row["Subchapter"] = item.get("subchapter") or u""
        row["EpuItem"] = item.get("epu_item") or u""
        row["PriceBook"] = item.get("price_book") or u""
        row["ShortDescription"] = item.get("short_description") or u""
        row["Description"] = item.get("description") or u""
        row["Unit"] = item.get("unit") or u""
        row["UnitPrice"] = qs.format_decimal(item.get("price"))
        self.table.Rows.Add(row)
        return row

    def _update_units(self):
        if self.col_unit is None:
            return
        extra = set()
        for row in self.table.Rows:
            if row.RowState == System.Data.DataRowState.Deleted:
                continue
            unit = cell_text(row["Unit"])
            if unit and unit not in qs.UNITS:
                extra.add(unit)
        self.col_unit.ItemsSource = [u""] + list(qs.UNITS) + sorted(extra)

    def _set_meta(self, name, currency):
        self._updating = True
        try:
            self.txt_name.Text = name or u""
            self.txt_currency.Text = currency or u"EUR"
        finally:
            self._updating = False

    def _show_path(self):
        document = self._document
        self.txt_path.Text = document.path or u"(new price list, not saved yet)"
        if document.path and document.updated_at:
            self.txt_file_hint.Text = u"Last saved by {} on {}.".format(
                document.updated_by or u"?", document.updated_at)
        else:
            self.txt_file_hint.Text = u"Save the price list as .json, then link it in the " \
                                      u"takeoff window with Browse...."

    def _new(self):
        self._document = qs.PriceListDocument(None)
        self._snapshot = {}
        self._fill({})
        self._set_meta(u"", u"EUR")
        self._dirty = False
        self._show_path()
        self._set_status(u"New price list.", False)

    def _open(self, path):
        document = qs.PriceListDocument(path)
        self.Cursor = Cursors.Wait
        try:
            document.load()
        except qs.PriceListError as error:
            self.Cursor = None
            MessageBox.Show(u"{}".format(error), TITLE, MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            return False
        finally:
            self.Cursor = None
        self._document = document
        self._snapshot = dict((code, dict(item)) for code, item in document.items.items())
        self._fill(document.items)
        self._set_meta(document.name, document.currency)
        self._dirty = False
        self._show_path()
        self._set_status(u"Opened {} items.".format(len(document.items)), False)
        return True

    def _read_table(self):
        """(voci, errore) dalla griglia; errore e' None se la griglia e' valida."""
        items = {}
        empty = 0
        duplicates = []
        for row in self.table.Rows:
            if row.RowState == System.Data.DataRowState.Deleted:
                continue
            code = cell_text(row["Code"])
            if not code:
                if any(cell_text(row[name]) for name in COLUMNS):
                    empty += 1
                continue
            if code in items:
                duplicates.append(code)
                continue
            items[code] = qs.clean_item(dict(
                (field, cell_text(row[column])) for column, field in FIELD_OF.items()))
        if empty:
            return items, u"{} rows have no price book code: fill it in or remove the " \
                          u"rows.".format(empty)
        if duplicates:
            return items, u"Price book codes used more than once: {}.".format(
                u", ".join(sorted(set(duplicates))))
        return items, None

    def _save(self, path=None):
        self._commit_edits()
        items, error = self._read_table()
        if error:
            MessageBox.Show(error, TITLE, MessageBoxButton.OK, MessageBoxImage.Warning)
            return False
        if path is None and not self._document.path:
            path = self._ask_save_path()
            if not path:
                return False

        changed = [code for code, item in items.items() if self._snapshot.get(code) != item]
        deleted = [code for code in self._snapshot if code not in items]
        try:
            merged = self._document.save(items, changed, deleted,
                                         (self.txt_name.Text or u"").strip(),
                                         (self.txt_currency.Text or u"").strip(),
                                         self._user, path)
        except qs.PriceListError as error:
            MessageBox.Show(u"{}".format(error), TITLE, MessageBoxButton.OK,
                            MessageBoxImage.Error)
            return False

        document = self._document
        self._snapshot = dict((code, dict(item)) for code, item in document.items.items())
        if merged:
            self._fill(document.items)
            self._set_status(u"Saved. Someone else had saved the price list meanwhile: "
                             u"their changes were kept and are now shown.", False)
        else:
            self._set_status(u"Saved {} items.".format(len(document.items)), False)
        self._dirty = False
        self.saved_path = document.path
        self._show_path()
        return True

    def _ask_save_path(self):
        dialog = SaveFileDialog()
        dialog.Title = "Save the price list"
        dialog.Filter = JSON_FILTER
        current = self._document.path
        if current:
            dialog.InitialDirectory = os.path.dirname(current)
            dialog.FileName = os.path.basename(current)
        else:
            dialog.FileName = (self.txt_name.Text or u"PriceList").strip() + u".json"
        if dialog.ShowDialog(self) != True:
            return None
        return dialog.FileName

    def _confirm_discard(self):
        """True se si puo' proseguire (modifiche salvate o scartate)."""
        self._commit_edits()
        if not self._dirty:
            return True
        answer = MessageBox.Show(u"The price list has unsaved changes.\n\nSave them?", TITLE,
                                 MessageBoxButton.YesNoCancel, MessageBoxImage.Warning)
        if answer == MessageBoxResult.Cancel:
            return False
        if answer == MessageBoxResult.Yes:
            return self._save()
        return True

    # ------------------------------------------------------------ import

    def _import(self, path, ask=True, layout=None):
        # Excel / CSV: foglio e colonne si scelgono prima di leggere (se non arrivano gia'
        # dalla finestra del computo). Una finestra non ancora mostrata non fa da Owner.
        if qs.is_table_file(path) and layout is None:
            layout = show_columns_dialog(self if self.IsVisible else None, path)
            if layout is None:
                return
        try:
            source = qs.load_price_list(path, layout)
        except qs.PriceListError as error:
            MessageBox.Show(u"{}".format(error), TITLE, MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            return
        rows = {}
        for row in self.table.Rows:
            if row.RowState != System.Data.DataRowState.Deleted:
                rows.setdefault(cell_text(row["Code"]), row)
        existing = [code for code in source.items if code in rows]
        overwrite = False
        if existing and ask:
            answer = MessageBox.Show(
                u"{} of the {} imported codes are already in the price list.\n\n"
                u"Yes: overwrite them with the imported values.\n"
                u"No: keep the current values and add only the new codes.".format(
                    len(existing), len(source.items)),
                TITLE, MessageBoxButton.YesNoCancel, MessageBoxImage.Question)
            if answer == MessageBoxResult.Cancel:
                return
            overwrite = answer == MessageBoxResult.Yes

        added = updated = 0
        for code, values in sorted(source.items.items(), key=item_sort_key):
            item = qs.clean_item(values)
            if code in rows:
                if not overwrite:
                    continue
                row = rows[code]
                row["Chapter"] = item["chapter"]
                row["Subchapter"] = item["subchapter"]
                row["EpuItem"] = item["epu_item"]
                row["PriceBook"] = item["price_book"]
                row["ShortDescription"] = item["short_description"]
                row["Description"] = item["description"]
                row["Unit"] = item["unit"]
                row["UnitPrice"] = qs.format_decimal(item["price"])
                updated += 1
            else:
                rows[code] = self._add_row(code, item)
                added += 1
        if not (self.txt_name.Text or u"").strip():
            self.txt_name.Text = source.sheet_name or os.path.splitext(os.path.basename(path))[0]
        self._dirty = self._dirty or bool(added or updated)
        self._update_units()
        self._update_count()
        message = u"Imported from {}: {} new codes, {} updated.".format(
            os.path.basename(path), added, updated)
        if source.duplicates:
            message += u" {} duplicate codes in the file were ignored.".format(source.duplicates)
        if not self._document.path:
            message += u" Save it as .json to keep it."
        self._set_status(message, False)

    # ------------------------------------------------------------ griglia

    def _commit_edits(self):
        try:
            self.dg_items.CommitEdit(DataGridEditingUnit.Cell, True)
            self.dg_items.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            pass

    def _update_count(self):
        total = len([r for r in self.table.Rows
                     if r.RowState != System.Data.DataRowState.Deleted])
        shown = self.table.DefaultView.Count
        self.txt_count.Text = u"{} items".format(total) if shown == total \
            else u"{} of {} items shown".format(shown, total)

    def _set_status(self, text, is_warning):
        self.txt_status.Text = text
        self.txt_status.Foreground = RED_BRUSH if is_warning else GRAY_BRUSH

    def _mark_dirty(self):
        self._dirty = True
        self._set_status(u"Unsaved changes.", True)

    def _warn_later(self, message):
        self.Dispatcher.BeginInvoke(DispatcherPriority.Background, Action(
            lambda: MessageBox.Show(message, TITLE, MessageBoxButton.OK,
                                    MessageBoxImage.Warning)))

    def _selected_rows(self):
        rows = [item.Row for item in self.dg_items.SelectedItems if hasattr(item, "Row")]
        if not rows:
            cell = self.dg_items.CurrentCell
            if cell.Item is not None and hasattr(cell.Item, "Row"):
                rows = [cell.Item.Row]
        return rows

    def _unique_code(self, base):
        codes = set(cell_text(r["Code"]) for r in self.table.Rows
                    if r.RowState != System.Data.DataRowState.Deleted)
        candidate = u"{}-COPY".format(base)
        index = 2
        while candidate in codes:
            candidate = u"{}-COPY{}".format(base, index)
            index += 1
        return candidate

    # ------------------------------------------------------------ eventi

    def OnCellChanging(self, sender, args):
        if self._updating:
            return
        name = args.Column.ColumnName
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

    def OnCellChanged(self, sender, args):
        if self._updating:
            return
        self._mark_dirty()
        if args.Column.ColumnName == "Unit":
            self._update_units()

    def OnRowDeleted(self, sender, args):
        if self._updating:
            return
        self._mark_dirty()

    def OnMetaChanged(self, sender, args):
        if self._updating:
            return
        self._mark_dirty()

    def OnSearchChanged(self, sender, args):
        self._commit_edits()
        text = escape_like((self.txt_search.Text or u"").strip())
        self.table.DefaultView.RowFilter = (
            u"(Code LIKE '%{0}%' OR Description LIKE '%{0}%' OR Chapter LIKE '%{0}%' "
            u"OR ShortDescription LIKE '%{0}%' "
            u"OR Subchapter LIKE '%{0}%' OR EpuItem LIKE '%{0}%' "
            u"OR PriceBook LIKE '%{0}%')".format(text)) if text else u""
        self._update_count()

    def OnClearSearch(self, sender, args):
        self.txt_search.Text = u""

    def OnAdd(self, sender, args):
        self._commit_edits()
        self.txt_search.Text = u""
        self._updating = True
        try:
            row = self._add_row(u"", {})
        finally:
            self._updating = False
        self._mark_dirty()
        self._update_count()
        view = self.table.DefaultView
        for index in range(view.Count):
            if view[index].Row is row:
                self.dg_items.SelectedItem = view[index]
                self.dg_items.ScrollIntoView(view[index])
                break

    def OnDuplicate(self, sender, args):
        self._commit_edits()
        rows = self._selected_rows()
        if not rows:
            return
        self._updating = True
        try:
            for source in rows:
                item = dict((field, cell_text(source[column]))
                            for column, field in FIELD_OF.items())
                item["price"] = qs.parse_decimal(item["price"]) if item["price"] else None
                self._add_row(self._unique_code(cell_text(source["Code"])), item)
        finally:
            self._updating = False
        self._mark_dirty()
        self._update_count()
        self._set_status(u"Duplicated {} rows at the end of the list: change their codes.".format(
            len(rows)), True)

    def OnRemove(self, sender, args):
        self._commit_edits()
        rows = self._selected_rows()
        if not rows:
            return
        for row in rows:
            row.Delete()
        self.table.AcceptChanges()
        self._update_count()

    def OnNew(self, sender, args):
        if self._confirm_discard():
            self._new()

    def OnOpen(self, sender, args):
        if not self._confirm_discard():
            return
        dialog = OpenFileDialog()
        dialog.Title = "Open a price list"
        dialog.Filter = JSON_FILTER
        if self._document.path:
            dialog.InitialDirectory = os.path.dirname(self._document.path)
        if dialog.ShowDialog(self) == True:
            self._open(dialog.FileName)

    def OnImport(self, sender, args):
        self._commit_edits()
        dialog = OpenFileDialog()
        dialog.Title = "Import a price list"
        dialog.Filter = IMPORT_FILTER
        if dialog.ShowDialog(self) == True:
            self._import(dialog.FileName)

    def OnSave(self, sender, args):
        self._save()

    def OnSaveAs(self, sender, args):
        path = self._ask_save_path()
        if path:
            self._save(path)

    def OnCloseClick(self, sender, args):
        self.Close()

    def OnWindowClosing(self, sender, args):
        if not self._confirm_discard():
            args.Cancel = True


def show_pricelist_editor(owner, path=None, import_path=None, user_name=u"",
                          import_layout=None):
    """Apre l'editor; restituisce il percorso dell'ultimo listino salvato, o None.
    import_layout: foglio e colonne del listino Excel / CSV da importare."""
    form = PriceListEditor(path, import_path, user_name, import_layout)
    if owner is not None:
        form.Owner = owner
        form.WindowStartupLocation = WindowStartupLocation.CenterOwner
    form.ShowDialog()
    return form.saved_path
