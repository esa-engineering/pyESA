# -*- coding: utf-8 -*-
"""
MEPQTO_columns_ui.py - foglio e colonne di un listino Excel / CSV / PriMus o della
struttura WBS.

Un .xpwe si presenta come un foglio unico (MEPQTO_store.XPWE_HEADERS): qui si sceglie
fra l'altro quale dei cinque prezzi di PriMus usare.

Si apre quando si sceglie un listino .xlsx / .xlsm / .csv / .xpwe (Browse... ed editor del
listino) e dal pulsante Columns... della finestra del computo. Per ogni campo si sceglie
una colonna del foglio, con l'anteprima delle prime righe. La proposta iniziale e' la
mappatura gia' salvata per lo stesso file, altrimenti quella ricavata dalle intestazioni
note, altrimenti A..E (F..I).

La stessa finestra mappa la struttura WBS (show_wbs_columns_dialog): le tabelle dei livelli
stanno affiancate sullo stesso foglio, quindi per ogni livello WBS attivo si sceglie la
coppia di colonne codice / descrizione, con una riga di partenza comune. Le differenze fra
i due usi stanno nelle classi PriceListKind e WbsKind.

La finestra restituisce un MEPQTO_store.PriceListLayout (listino) o WbsLayout (struttura
WBS), oppure None se annullata.
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
from System.Windows import (Window, MessageBox, WindowStartupLocation, Visibility,
                            GridLength, GridUnitType, FontWeights, TextTrimming)
from System.Windows.Controls import (ComboBox, Grid, RowDefinition, ColumnDefinition,
                                     TextBlock)
from System.Windows.Input import Cursors
from System.Windows.Markup import XamlReader

from pyrevit import script

import MEPQTO_store as qs

XAML_FILE_NAME = 'MEPQTO_columns.xaml'
TITLE = "Price list columns"
NONE_COLUMN = u"(none)"
PREVIEW_ROWS = 40
HEADER_TEXT_LENGTH = 28
# colonna dell'anteprima con il numero di riga del foglio (per scegliere la riga di partenza)
ROW_COLUMN = u"Row"


class PriceListKind(object):
    """Listino: campi fissi (LAYOUT_FIELDS), un campo per riga, nessuna riga di partenza."""

    title = u"Price list columns"
    first_row = False
    pairs = False
    sheet_tip = u"Sheet of the workbook that holds the unit price list"
    sheet_hint = None
    columns_hint = None
    positions = (u"A to E", None)
    detect_tip = None
    detect_message = (u"No known header names (Codice, Descrizione, UM, Prezzo...) were "
                      u"found in the first rows of the sheet.")

    def __init__(self):
        self.fields = qs.LAYOUT_FIELDS

    def default(self):
        return qs.PriceListLayout.default()

    def guess(self, rows, sheet):
        return qs.guess_layout(rows, sheet)

    def make(self, sheet, columns, first_row, previous):
        return qs.PriceListLayout(sheet, columns)

    def check(self, columns):
        return None

    def header_row(self, rows, sheet):
        """Riga da cui prendere i nomi delle colonne nelle tendine."""
        return qs.header_row_index(rows)


class WbsKind(object):
    """Struttura WBS: per ogni livello attivo due campi, "code:<nome>" e
    "description:<nome>", mostrati su una riga; riga di partenza comune."""

    title = u"WBS structure columns"
    first_row = True
    pairs = True
    sheet_tip = u"Sheet of the workbook that holds the tables of the WBS levels"
    sheet_hint = (u"Values are read from the first row on (the Row column of the preview "
                  u"numbers the rows of the sheet). In the columns of every level, cells "
                  u"without a code or a description and a header (Codice...) are skipped; "
                  u"the tables can have different lengths. A code repeated in the same "
                  u"level: the first row wins.")
    columns_hint = (u"The tables of the WBS levels sit side by side. For every WBS level "
                    u"pick the column of its code and of its description, or leave both as "
                    u"(none). A value of a WBS parameter is looked up in the codes of its "
                    u"own level, so the same code can mean different things in two levels.")
    positions = (u"Side by side", u"Level 1 in A and B, level 2 in C and D, and so on")
    detect_tip = (u"Find the Codice / Descrizione (or Descrittivo) header pairs: a pair goes "
                  u"to the level whose name is written above it, the others to the "
                  u"remaining levels from left to right; values start after the header row")
    detect_message = (u"No header row with Codice and Descrizione (or Descrittivo) side by "
                      u"side was found in the first rows of the sheet.")

    def __init__(self, level_names):
        self.level_names = list(level_names)
        self.fields = []
        for name in self.level_names:
            self.fields.append((qs.WBS_CODE_FIELD + name, u"{} code".format(name), False))
            self.fields.append((qs.WBS_DESCRIPTION_FIELD + name,
                                u"{} description".format(name), False))

    def default(self):
        return qs.WbsLayout.pairs(self.level_names)

    def guess(self, rows, sheet):
        return qs.guess_wbs_layout(rows, sheet, self.level_names)

    def header_row(self, rows, sheet):
        """La riga Codice / Descrizione trovata, non quella dei titoli dei livelli (che ha
        anch'essa piu' testi)."""
        guessed = self.guess(rows, sheet)
        return guessed.first_row - 2 if guessed is not None else qs.header_row_index(rows)

    def make(self, sheet, columns, first_row, previous):
        # i livelli non mostrati (spenti adesso) restano come erano
        keep = dict((name, pair) for name, pair in previous.levels.items()
                    if name not in self.level_names) if previous is not None else None
        return qs.WbsLayout.from_columns(sheet, columns, first_row, keep)

    def check(self, columns):
        """Codice e descrizione insieme per ogni livello; almeno un livello."""
        mapped = 0
        for name in self.level_names:
            code = qs.WBS_CODE_FIELD + name in columns
            description = qs.WBS_DESCRIPTION_FIELD + name in columns
            if code != description:
                return u"{}: choose both the code and the description column, or " \
                       u"neither.".format(name)
            mapped += 1 if code else 0
        if not mapped:
            return u"Choose the code and description columns of at least one WBS level."
        return None


class ColumnsForm(Window):

    def __init__(self, path, layout=None, kind=None):
        self.result = None
        self._path = path
        self._layout = layout
        self._kind = kind or PriceListKind()
        self._fields = self._kind.fields
        self._loading = True
        self._rows = []
        self._column_count = 0
        self._combos = {}
        self._load_xaml()
        self._apply_kind()
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
        self._apply_layout(layout if layout is not None else self._proposal())
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
                     "txt_preview", "dg_preview", "btn_ok", "btn_cancel", "txt_title",
                     "txt_sheet_hint", "txt_columns_hint", "grd_first_row", "txt_first_row"):
            setattr(self, name, root.FindName(name))

        self.cbo_sheet.SelectionChanged += self.OnSheetChanged
        self.btn_detect.Click += self.OnDetect
        self.btn_positions.Click += self.OnPositions
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel

    def _apply_kind(self):
        """Titoli, note e pulsanti del tipo di file (listino o struttura WBS)."""
        kind = self._kind
        self.Title = kind.title
        self.txt_title.Text = u"\U0001f4d1 " + kind.title
        self.cbo_sheet.ToolTip = kind.sheet_tip
        if kind.sheet_hint:
            self.txt_sheet_hint.Text = kind.sheet_hint
        if kind.columns_hint:
            self.txt_columns_hint.Text = kind.columns_hint
        label, tip = kind.positions
        self.btn_positions.Content = label
        if tip:
            self.btn_positions.ToolTip = tip
        if kind.detect_tip:
            self.btn_detect.ToolTip = kind.detect_tip
        self.grd_first_row.Visibility = Visibility.Visible if kind.first_row \
            else Visibility.Collapsed
        self.txt_first_row.Text = u"1"

    def _combo(self, field, label, row, column):
        combo = ComboBox()
        combo.Style = self._styles["ComboStyle"]
        combo.ToolTip = u"Column of the {}".format(label)
        Grid.SetRow(combo, row)
        Grid.SetColumn(combo, column)
        self.grd_fields.Children.Add(combo)
        self._combos[field] = combo

    def _label(self, text, row, bold=False):
        block = TextBlock()
        block.Text = text
        block.Style = self._styles["FieldLabel"]
        block.TextTrimming = TextTrimming.CharacterEllipsis
        block.ToolTip = text
        if bold:
            block.FontWeight = FontWeights.Bold
        Grid.SetRow(block, row)
        Grid.SetColumn(block, 0)
        self.grd_fields.Children.Add(block)
        return block

    def _build_fields(self):
        if not self._kind.pairs:
            for index, (field, label, required) in enumerate(self._fields):
                self.grd_fields.RowDefinitions.Add(RowDefinition())
                self._label(label + (u" *" if required else u":"), index)
                self._combo(field, label.lower(), index, 1)
            return
        # Struttura WBS: livello | codice | descrizione, una riga per livello.
        self.grd_fields.ColumnDefinitions.Add(ColumnDefinition())
        self.grd_fields.ColumnDefinitions[2].Width = GridLength(1, GridUnitType.Star)
        self.grd_fields.RowDefinitions.Add(RowDefinition())
        self._label(u"WBS level", 0, bold=True)
        for column, text in ((1, u"Code"), (2, u"Description")):
            header = TextBlock()
            header.Text = text
            header.FontWeight = FontWeights.Bold
            header.Margin = System.Windows.Thickness(4, 0, 0, 2)
            Grid.SetRow(header, 0)
            Grid.SetColumn(header, column)
            self.grd_fields.Children.Add(header)
        for index, name in enumerate(self._kind.level_names):
            row = index + 1
            self.grd_fields.RowDefinitions.Add(RowDefinition())
            self._label(name, row)
            self._combo(qs.WBS_CODE_FIELD + name, u"code of {}".format(name), row, 1)
            self._combo(qs.WBS_DESCRIPTION_FIELD + name, u"description of {}".format(name),
                        row, 2)

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
        header = self._kind.header_row(self._rows, self._sheet())

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
        for field, label, required in self._fields:
            combo = self._combos[field]
            selected = self._selected(field)
            combo.ItemsSource = ([] if required else [NONE_COLUMN]) + choices
            self._select(field, selected)

        table = DataTable("preview")
        string_type = clr.GetClrType(System.String)
        table.Columns.Add(ROW_COLUMN, string_type)
        for index in range(self._column_count):
            table.Columns.Add(qs.column_letter(index), string_type)
        for number, cells in enumerate(self._rows[:PREVIEW_ROWS]):
            row = table.NewRow()
            row[ROW_COLUMN] = u"{}".format(number + 1)
            for index in range(min(len(cells), self._column_count)):
                row[index + 1] = (cells[index] or u"").replace(u"\n", u" ")
            table.Rows.Add(row)
        self.dg_preview.ItemsSource = table.DefaultView
        self.txt_preview.Text = u"\U0001f441️ PREVIEW ({} of {} rows)".format(
            min(len(self._rows), PREVIEW_ROWS), len(self._rows))

    # ------------------------------------------------------------------ mappatura

    def _required(self, field):
        return [required for name, _, required in self._fields if name == field][0]

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
        for field, _, _ in self._fields:
            self._select(field, columns.get(field))

    def _apply_layout(self, layout):
        self._apply(layout.columns)
        self.txt_first_row.Text = u"{}".format(getattr(layout, "first_row", 1))

    def _proposal(self):
        guessed = self._kind.guess(self._rows, self._sheet())
        return guessed if guessed is not None else self._kind.default()

    def _first_row(self):
        """(riga di partenza, messaggio d'errore); 1 se il tipo non la prevede."""
        if not self._kind.first_row:
            return 1, None
        text = (self.txt_first_row.Text or u"").strip()
        last = max(len(self._rows), 1)
        try:
            value = int(text)
        except ValueError:
            value = None
        if value is None or value < 1 or value > last:
            return None, u"First row: enter a whole number from 1 to {} (the rows of the " \
                         u"sheet).".format(last)
        return value, None

    def _validate(self):
        columns = {}
        for field, label, required in self._fields:
            index = self._selected(field)
            if index is None:
                if required:
                    return None, u"Choose the column of the {}.".format(label.lower())
                continue
            columns[field] = index
        used = {}
        for field, index in columns.items():
            used.setdefault(index, []).append(field)
        labels = dict((field, label) for field, label, _ in self._fields)
        twice = [(index, fields) for index, fields in used.items() if len(fields) > 1]
        if twice:
            index, fields = sorted(twice)[0]
            return None, u"Column {} is used for more than one field: {}.".format(
                qs.column_letter(index), u", ".join(labels[field] for field in fields))
        message = self._kind.check(columns)
        if message:
            return None, message
        return columns, None

    # ------------------------------------------------------------------ eventi

    def OnSheetChanged(self, sender, args):
        if self._loading:
            return
        self._read_rows()
        self._apply_layout(self._proposal())

    def OnDetect(self, sender, args):
        guessed = self._kind.guess(self._rows, self._sheet())
        if guessed is None:
            MessageBox.Show(self._kind.detect_message, self._kind.title)
            return
        self._apply_layout(guessed)

    def OnPositions(self, sender, args):
        self._apply(self._kind.default().columns)

    def OnOK(self, sender, args):
        columns, message = self._validate()
        first_row, row_message = self._first_row()
        if columns is None or first_row is None:
            MessageBox.Show(message or row_message, self._kind.title)
            return
        self.result = self._kind.make(self._sheet(), columns, first_row, self._layout)
        self.Close()

    def OnCancel(self, sender, args):
        self.result = None
        self.Close()


def show_columns_dialog(owner, path, layout=None, kind=None):
    """PriceListLayout scelto per il listino path (layout: mappatura attuale, se e' lo
    stesso file), oppure None se annullata. kind: WbsKind per la struttura WBS."""
    form = ColumnsForm(path, layout, kind)
    if owner is not None:
        form.Owner = owner
        form.WindowStartupLocation = WindowStartupLocation.CenterOwner
    form.ShowDialog()
    return form.result


def show_wbs_columns_dialog(owner, path, layout, level_names):
    """WbsLayout della struttura WBS: foglio, riga di partenza e colonne di codice e
    descrizione di ogni livello WBS attivo (level_names); None se annullata."""
    return show_columns_dialog(owner, path, layout, WbsKind(level_names))
