# -*- coding: utf-8 -*-
"""
mepqto_columns_ui.py - foglio e colonne di un listino Excel / CSV / PriMus.

Un .xpwe si presenta come un foglio unico (mepqto_store.XPWE_HEADERS): qui si sceglie
fra l'altro quale dei cinque prezzi di PriMus usare.

Si apre quando si sceglie un listino .xlsx / .xlsm / .csv / .xpwe (Browse... ed editor del
listino) e dal pulsante Columns... della finestra del computo. Per ogni campo si sceglie
una colonna del foglio, con l'anteprima delle prime righe. La proposta iniziale e' la
mappatura gia' salvata per lo stesso file, altrimenti quella ricavata dalle intestazioni
note, altrimenti A..E (F..I).

La finestra restituisce un mepqto_store.PriceListLayout, oppure None se annullata.
"""

import os

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Data')

import System
from System.Data import DataTable
from System.IO import FileStream, FileMode
from System.Windows import Window, MessageBox, WindowStartupLocation
from System.Windows.Controls import ComboBox, Grid, RowDefinition, TextBlock
from System.Windows.Input import Cursors
from System.Windows.Markup import XamlReader

from pyrevit import script

import mepqto_store as qs

XAML_FILE_NAME = 'MEPQTO_columns.xaml'
TITLE = "Price list columns"
NONE_COLUMN = u"(none)"
PREVIEW_ROWS = 40
HEADER_TEXT_LENGTH = 28


class ColumnsForm(Window):

    def __init__(self, path, layout=None):
        self.result = None
        self._path = path
        self._loading = True
        self._rows = []
        self._column_count = 0
        self._combos = {}
        self._load_xaml()
        self.txt_file.Text = path

        self._sheets = qs.list_sheets(path)
        initial_sheet = layout.sheet if layout is not None else u""
        if self._sheets:
            self.cbo_sheet.ItemsSource = list(self._sheets)
            self.cbo_sheet.SelectedItem = initial_sheet if initial_sheet in self._sheets \
                else self._sheets[0]
        else:
            self.cbo_sheet.ItemsSource = [u"(PriMus file: the price list)"
                                          if qs.is_xpwe_file(path) else u"(CSV file: one sheet)"]
            self.cbo_sheet.SelectedIndex = 0
            self.cbo_sheet.IsEnabled = False
        self._build_fields()
        self._read_rows()
        if layout is not None:
            self._apply(layout.columns)
        else:
            self._apply(self._proposal())
        self._loading = False

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
        self._styles = root.Resources

        for name in ("txt_file", "cbo_sheet", "grd_fields", "btn_detect", "btn_positions",
                     "txt_preview", "dg_preview", "btn_ok", "btn_cancel"):
            setattr(self, name, root.FindName(name))

        self.cbo_sheet.SelectionChanged += self.OnSheetChanged
        self.btn_detect.Click += self.OnDetect
        self.btn_positions.Click += self.OnPositions
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel

    def _build_fields(self):
        for index, (field, label, required) in enumerate(qs.LAYOUT_FIELDS):
            self.grd_fields.RowDefinitions.Add(RowDefinition())
            text = TextBlock()
            text.Text = label + (u" *" if required else u":")
            text.Style = self._styles["FieldLabel"]
            Grid.SetRow(text, index)
            Grid.SetColumn(text, 0)
            combo = ComboBox()
            combo.Style = self._styles["ComboStyle"]
            combo.ToolTip = u"Column of the {}".format(label.lower())
            Grid.SetRow(combo, index)
            Grid.SetColumn(combo, 1)
            self.grd_fields.Children.Add(text)
            self.grd_fields.Children.Add(combo)
            self._combos[field] = combo

    # ------------------------------------------------------------------ foglio

    def _sheet(self):
        return (self.cbo_sheet.SelectedItem or u"") if self._sheets else u""

    def _read_rows(self):
        """Righe del foglio scelto, anteprima e scelte delle colonne."""
        self.Cursor = Cursors.Wait
        try:
            _, self._rows = qs.read_sheet_rows(self._path, self._sheet())
        except Exception as error:
            self._rows = []
            MessageBox.Show(u"The sheet could not be read:\n{}".format(error), TITLE)
        finally:
            self.Cursor = None
        sample = self._rows[:max(PREVIEW_ROWS, qs.HEADER_SCAN_ROWS)]
        self._column_count = max([len(cells) for cells in sample] + [5])
        header = qs.header_row_index(self._rows)

        # Scelte: "C - Descrizione" (intestazione dalla riga d'intestazione, se c'e').
        choices = []
        for index in range(self._column_count):
            letter = qs.column_letter(index)
            text = u""
            if header is not None and index < len(self._rows[header]):
                text = (self._rows[header][index] or u"").strip().replace(u"\n", u" ")
            if len(text) > HEADER_TEXT_LENGTH:
                text = text[:HEADER_TEXT_LENGTH - 1] + u"…"
            choices.append(u"{} - {}".format(letter, text) if text else letter)
        for field, label, required in qs.LAYOUT_FIELDS:
            combo = self._combos[field]
            selected = self._selected(field)
            combo.ItemsSource = ([] if required else [NONE_COLUMN]) + choices
            self._select(field, selected)

        table = DataTable("preview")
        string_type = clr.GetClrType(System.String)
        for index in range(self._column_count):
            table.Columns.Add(qs.column_letter(index), string_type)
        for cells in self._rows[:PREVIEW_ROWS]:
            row = table.NewRow()
            for index in range(min(len(cells), self._column_count)):
                row[index] = (cells[index] or u"").replace(u"\n", u" ")
            table.Rows.Add(row)
        self.dg_preview.ItemsSource = table.DefaultView
        self.txt_preview.Text = u"\U0001f441️ PREVIEW ({} of {} rows)".format(
            min(len(self._rows), PREVIEW_ROWS), len(self._rows))

    # ------------------------------------------------------------------ mappatura

    def _required(self, field):
        return [required for name, _, required in qs.LAYOUT_FIELDS if name == field][0]

    def _selected(self, field):
        """Indice di colonna scelto per il campo; None per (none) o nessuna scelta."""
        combo = self._combos[field]
        index = combo.SelectedIndex
        if index < 0:
            return None
        if not self._required(field):
            index -= 1
        return index if index >= 0 else None

    def _select(self, field, column):
        combo = self._combos[field]
        offset = 0 if self._required(field) else 1
        if column is None or column >= self._column_count:
            combo.SelectedIndex = -1 if self._required(field) else 0
        else:
            combo.SelectedIndex = column + offset

    def _apply(self, columns):
        for field, _, _ in qs.LAYOUT_FIELDS:
            self._select(field, columns.get(field))

    def _proposal(self):
        guessed = qs.guess_layout(self._rows, self._sheet())
        return guessed.columns if guessed is not None else qs.PriceListLayout.default().columns

    def _validate(self):
        columns = {}
        for field, label, required in qs.LAYOUT_FIELDS:
            index = self._selected(field)
            if index is None:
                if required:
                    return None, u"Choose the column of the {}.".format(label.lower())
                continue
            columns[field] = index
        used = {}
        for field, index in columns.items():
            used.setdefault(index, []).append(field)
        labels = dict((field, label) for field, label, _ in qs.LAYOUT_FIELDS)
        twice = [(index, fields) for index, fields in used.items() if len(fields) > 1]
        if twice:
            index, fields = sorted(twice)[0]
            return None, u"Column {} is used for more than one field: {}.".format(
                qs.column_letter(index), u", ".join(labels[field] for field in fields))
        return columns, None

    # ------------------------------------------------------------------ eventi

    def OnSheetChanged(self, sender, args):
        if self._loading:
            return
        self._read_rows()
        self._apply(self._proposal())

    def OnDetect(self, sender, args):
        guessed = qs.guess_layout(self._rows, self._sheet())
        if guessed is None:
            MessageBox.Show(u"No known header names (Codice, Descrizione, UM, Prezzo...) were "
                            u"found in the first rows of the sheet.", TITLE)
            return
        self._apply(guessed.columns)

    def OnPositions(self, sender, args):
        self._apply(qs.PriceListLayout.default().columns)

    def OnOK(self, sender, args):
        columns, message = self._validate()
        if columns is None:
            MessageBox.Show(message, TITLE)
            return
        self.result = qs.PriceListLayout(self._sheet(), columns)
        self.Close()

    def OnCancel(self, sender, args):
        self.result = None
        self.Close()


def show_columns_dialog(owner, path, layout=None):
    """PriceListLayout scelto per il listino path (layout: mappatura attuale, se e' lo
    stesso file), oppure None se annullata."""
    form = ColumnsForm(path, layout)
    if owner is not None:
        form.Owner = owner
        form.WindowStartupLocation = WindowStartupLocation.CenterOwner
    form.ShowDialog()
    return form.result
