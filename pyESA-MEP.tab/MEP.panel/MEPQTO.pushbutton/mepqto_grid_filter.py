# -*- coding: utf-8 -*-
"""
mepqto_grid_filter.py - filtri per colonna, in stile Excel, sulle griglie legate a una
DataTable.

Ogni colonna della griglia riceve nell'intestazione un pulsante a imbuto che apre un
elenco dei valori della colonna, con caselle di spunta, una ricerca e "(Select All)".
L'elenco mostra i valori delle righe che passano gli altri filtri (come in Excel); le
celle vuote compaiono come "(Blanks)". Spuntare tutto toglie il filtro.

GridFilters non tocca la DataView: produce un'espressione per DataView.RowFilter
(expression()) e chiama on_change; e' la finestra a combinarla con la ricerca e con
le proprie regole (le righe di totale del computo, per esempio).
"""

from collections import OrderedDict

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Data')

import System
from System import Action, DBNull
from System.Data import DataView, DataViewRowState
from System.Windows import (HorizontalAlignment, VerticalAlignment, Thickness, FontWeights,
                            Setter, Style, TextWrapping)
from System.Windows.Controls import (Border, Button, CheckBox, ColumnDefinition, Control,
                                     DataGridBoundColumn, DataGridComboBoxColumn, Grid,
                                     ListBox, Orientation, StackPanel, TextBlock, TextBox)
from System.Windows.Controls.Primitives import DataGridColumnHeader, Popup, PlacementMode
from System.Windows.Input import Cursors
from System.Windows.Media import BrushConverter, FontFamily
from System.Windows.Threading import DispatcherPriority

BLANKS = u"(Blanks)"
SELECT_ALL = u"(Select All)"
# Imbuto di Segoe MDL2 Assets (Windows 10 e successivi, quindi tutte le versioni di
# Revit supportate).
FUNNEL = u""
ICON_FONT = FontFamily(u"Segoe MDL2 Assets")

_BRUSHES = BrushConverter()
ACTIVE_BRUSH = _BRUSHES.ConvertFromString("#2D5A8A")
IDLE_BRUSH = _BRUSHES.ConvertFromString("#A0A0A0")
BORDER_BRUSH = _BRUSHES.ConvertFromString("#CCCCCC")
WHITE_BRUSH = _BRUSHES.ConvertFromString("#FFFFFF")
for _brush in (ACTIVE_BRUSH, IDLE_BRUSH, BORDER_BRUSH, WHITE_BRUSH):
    _brush.Freeze()

CLR_DOUBLE = clr.GetClrType(System.Double)
CLR_INT = clr.GetClrType(System.Int32)


def _is_blank(value):
    return value is None or isinstance(value, DBNull) or u"{}".format(value).strip() == u""


def _literal(text):
    return u"'{}'".format(u"{}".format(text).replace(u"'", u"''"))


def binding_field(column):
    """Nome della colonna della DataTable legata a una colonna della griglia."""
    binding = None
    if isinstance(column, DataGridComboBoxColumn):
        binding = column.SelectedItemBinding
    elif isinstance(column, DataGridBoundColumn):
        binding = column.Binding
    try:
        return binding.Path.Path if binding is not None else None
    except Exception:
        return None


class _ColumnFilter(object):
    """Stato del filtro di una colonna: values None = nessun filtro, altrimenti
    l'insieme dei valori ammessi (BLANKS per le celle vuote)."""

    def __init__(self, column, field, numeric):
        self.column = column
        self.field = field
        self.numeric = numeric
        self.values = None
        self.title_block = None
        self.icon = None
        self.button = None

    @property
    def active(self):
        return self.values is not None

    def expression(self):
        if self.values is None:
            return u""
        field = u"[{}]".format(self.field)
        parts = []
        plain = [value for value in self.values if value != BLANKS]
        if plain:
            if self.numeric:
                items = u", ".join(repr(float(value)) for value in plain)
            else:
                items = u", ".join(_literal(value) for value in plain)
            parts.append(u"{} IN ({})".format(field, items))
        if BLANKS in self.values:
            parts.append(u"{} IS NULL".format(field) if self.numeric
                         else u"ISNULL({}, '') = ''".format(field))
        if not parts:
            return u"FALSE"
        return u"({})".format(u" OR ".join(parts))


class GridFilters(object):
    """Filtri per colonna di una DataGrid legata a table.

    on_change(): chiamata quando un filtro cambia; la finestra ricalcola il RowFilter.
    context(): espressione dei filtri esterni (ricerca, caselle), per mostrare
        nell'elenco solo i valori delle righe visibili.
    rows_filter: espressione che individua le righe filtrabili (le voci del computo,
        non le righe di totale): i valori si leggono solo da quelle.
    formatters: {campo: funzione} per il testo dei valori numerici nell'elenco.
    """

    def __init__(self, grid, table, on_change, context=None, rows_filter=u"",
                 formatters=None, skip_fields=()):
        self.grid = grid
        self.table = table
        self.on_change = on_change
        self.context = context or (lambda: u"")
        self.rows_filter = rows_filter
        self.formatters = formatters or {}
        self.filters = OrderedDict()
        self._popup = None

        # L'intestazione occupa tutta la larghezza: il titolo a sinistra, l'imbuto a destra.
        header_style = Style(clr.GetClrType(DataGridColumnHeader))
        if grid.ColumnHeaderStyle is not None:
            header_style.BasedOn = grid.ColumnHeaderStyle
        header_style.Setters.Add(Setter(Control.HorizontalContentAlignmentProperty,
                                        HorizontalAlignment.Stretch))
        grid.ColumnHeaderStyle = header_style

        for column in grid.Columns:
            field = binding_field(column)
            if not field or field in skip_fields or field not in self._column_names():
                continue
            data_type = table.Columns[field].DataType
            numeric = data_type in (CLR_DOUBLE, CLR_INT)
            column_filter = _ColumnFilter(column, field, numeric)
            self._install_header(column_filter, u"{}".format(column.Header or u""))
            self.filters[field] = column_filter

    def _column_names(self):
        return set(column.ColumnName for column in self.table.Columns)

    # ------------------------------------------------------------ intestazioni

    def _install_header(self, column_filter, title):
        grid = Grid()
        grid.ColumnDefinitions.Add(ColumnDefinition())
        auto = ColumnDefinition()
        auto.Width = System.Windows.GridLength.Auto
        grid.ColumnDefinitions.Add(auto)

        block = TextBlock()
        block.Text = title
        block.TextWrapping = TextWrapping.Wrap
        block.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(block, 0)

        icon = TextBlock()
        icon.Text = FUNNEL
        icon.FontFamily = ICON_FONT
        icon.FontSize = 10
        icon.Foreground = IDLE_BRUSH
        button = Button()
        button.Content = icon
        button.Padding = Thickness(3, 1, 3, 1)
        button.Margin = Thickness(4, 0, 0, 0)
        button.Background = None
        button.BorderThickness = Thickness(0)
        button.Cursor = Cursors.Hand
        button.VerticalAlignment = VerticalAlignment.Center
        button.ToolTip = u"Filter this column"
        button.Tag = column_filter.field
        button.Click += self.OnFilterButton
        Grid.SetColumn(button, 1)

        grid.Children.Add(block)
        grid.Children.Add(button)
        column_filter.title_block = block
        column_filter.icon = icon
        column_filter.button = button
        column_filter.column.Header = grid

    def set_title(self, column, title):
        """Cambia il titolo di una colonna gia' filtrabile (le colonne WBS)."""
        for column_filter in self.filters.values():
            if column_filter.column is column:
                column_filter.title_block.Text = title or u""
                return
        column.Header = title

    def _update_header(self, column_filter):
        column_filter.icon.Foreground = ACTIVE_BRUSH if column_filter.active else IDLE_BRUSH
        column_filter.title_block.FontWeight = FontWeights.Bold if column_filter.active \
            else FontWeights.Normal
        column_filter.button.ToolTip = u"Filtered: click to change" if column_filter.active \
            else u"Filter this column"

    # ------------------------------------------------------------ stato

    @property
    def active(self):
        return any(f.active for f in self.filters.values())

    def expression(self, exclude=None):
        """AND dei filtri attivi (escluso il campo exclude); stringa vuota se nessuno."""
        parts = [f.expression() for field, f in self.filters.items()
                 if f.active and field != exclude]
        return u" AND ".join(parts)

    def clear(self, field=None):
        for key, column_filter in self.filters.items():
            if field is None or key == field:
                column_filter.values = None
                self._update_header(column_filter)
        self.on_change()

    def clear_missing_columns(self, visible_fields):
        """Toglie i filtri delle colonne non piu' visibili (livelli WBS spenti)."""
        changed = False
        for key, column_filter in self.filters.items():
            if column_filter.active and key not in visible_fields:
                column_filter.values = None
                self._update_header(column_filter)
                changed = True
        return changed

    # ------------------------------------------------------------ valori

    def _display(self, column_filter, value):
        formatter = self.formatters.get(column_filter.field)
        if formatter is not None:
            try:
                return formatter(value)
            except Exception:
                pass
        return u"{}".format(value)

    def distinct_values(self, field):
        """[(valore, testo)] delle righe che passano ricerca e altri filtri, ordinati;
        (Blanks) in testa."""
        column_filter = self.filters[field]
        parts = [part for part in (self.rows_filter, self.context(),
                                   self.expression(exclude=field)) if part]
        view = DataView(self.table, u" AND ".join(u"({})".format(p) for p in parts), u"",
                        DataViewRowState.CurrentRows)
        seen = OrderedDict()
        has_blank = False
        for row_view in view:
            value = row_view[field]
            if _is_blank(value):
                has_blank = True
                continue
            key = float(value) if column_filter.numeric else u"{}".format(value)
            if key not in seen:
                seen[key] = self._display(column_filter, value)

        def sort_key(key):
            if column_filter.numeric:
                return (0, key, u"")
            try:
                return (0, float(key.replace(u",", u".")), key.lower())
            except ValueError:
                return (1, 0.0, key.lower())

        values = [(key, seen[key]) for key in sorted(seen, key=sort_key)]
        if has_blank:
            values.insert(0, (BLANKS, BLANKS))
        return values

    # ------------------------------------------------------------ elenco a tendina

    def OnFilterButton(self, sender, args):
        self.open(sender.Tag, sender)

    def open(self, field, target):
        column_filter = self.filters[field]
        values = self.distinct_values(field)
        title = column_filter.title_block.Text or field

        boxes = []
        for value, text in values:
            box = CheckBox()
            label = TextBlock()
            label.Text = text
            box.Content = label
            box.Tag = value
            box.IsChecked = not column_filter.active or value in column_filter.values
            boxes.append(box)

        panel = StackPanel()
        panel.Width = 260
        heading = TextBlock()
        heading.Text = u"Filter: {}".format(title)
        heading.FontWeight = FontWeights.Bold
        heading.Margin = Thickness(0, 0, 0, 5)
        panel.Children.Add(heading)

        search_row = Grid()
        search_row.ColumnDefinitions.Add(ColumnDefinition())
        search_row.ColumnDefinitions[0].Width = System.Windows.GridLength.Auto
        search_row.ColumnDefinitions.Add(ColumnDefinition())
        lens = TextBlock()
        lens.Text = u"\U0001f50d"
        lens.VerticalAlignment = VerticalAlignment.Center
        lens.Margin = Thickness(0, 0, 5, 0)
        search = TextBox()
        search.Padding = Thickness(5, 3, 5, 3)
        search.ToolTip = u"Show only the values that contain this text"
        Grid.SetColumn(search, 1)
        search_row.Children.Add(lens)
        search_row.Children.Add(search)
        panel.Children.Add(search_row)

        select_all = CheckBox()
        select_all.Content = SELECT_ALL
        select_all.Margin = Thickness(0, 6, 0, 2)
        panel.Children.Add(select_all)

        list_box = ListBox()
        list_box.Height = 230
        panel.Children.Add(list_box)

        empty = TextBlock()
        empty.Foreground = IDLE_BRUSH
        empty.FontSize = 11
        empty.TextWrapping = TextWrapping.Wrap
        empty.Margin = Thickness(0, 4, 0, 0)
        panel.Children.Add(empty)

        buttons = StackPanel()
        buttons.Orientation = Orientation.Horizontal
        buttons.HorizontalAlignment = HorizontalAlignment.Right
        buttons.Margin = Thickness(0, 8, 0, 0)
        clear_button = Button()
        clear_button.Content = u"Clear Filter"
        clear_button.Padding = Thickness(8, 2, 8, 2)
        clear_button.Margin = Thickness(0, 0, 5, 0)
        clear_button.IsEnabled = column_filter.active
        cancel_button = Button()
        cancel_button.Content = u"Cancel"
        cancel_button.Padding = Thickness(8, 2, 8, 2)
        cancel_button.Margin = Thickness(0, 0, 5, 0)
        ok_button = Button()
        ok_button.Content = u"OK"
        ok_button.Width = 60
        ok_button.Background = ACTIVE_BRUSH
        ok_button.Foreground = WHITE_BRUSH
        for button in (clear_button, cancel_button, ok_button):
            buttons.Children.Add(button)
        panel.Children.Add(buttons)

        border = Border()
        border.BorderBrush = BORDER_BRUSH
        border.BorderThickness = Thickness(1)
        border.Background = WHITE_BRUSH
        border.Padding = Thickness(8)
        border.Child = panel

        popup = Popup()
        popup.Child = border
        popup.PlacementTarget = target
        popup.Placement = PlacementMode.Bottom
        popup.StaysOpen = False

        state = {"syncing": False}

        def visible_boxes():
            return [box for box in list_box.Items]

        def sync_select_all():
            shown = visible_boxes()
            checked = [box for box in shown if box.IsChecked]
            state["syncing"] = True
            try:
                if shown and len(checked) == len(shown):
                    select_all.IsChecked = True
                elif not checked:
                    select_all.IsChecked = False
                else:
                    select_all.IsChecked = None
            finally:
                state["syncing"] = False
            ok_button.IsEnabled = bool(checked)
            empty.Text = u"" if shown else u"No values."

        def fill(text):
            needle = (text or u"").strip().lower()
            list_box.Items.Clear()
            for box in boxes:
                if not needle or needle in box.Content.Text.lower():
                    list_box.Items.Add(box)
            # Come in Excel: con una ricerca si parte da tutti i valori trovati spuntati.
            if needle:
                for box in visible_boxes():
                    box.IsChecked = True
            sync_select_all()

        def on_search(sender, args):
            fill(search.Text)

        def on_select_all(sender, args):
            if state["syncing"]:
                return
            value = bool(select_all.IsChecked)
            for box in visible_boxes():
                box.IsChecked = value
            sync_select_all()

        def on_box(sender, args):
            if not state["syncing"]:
                sync_select_all()

        def on_ok(sender, args):
            shown = visible_boxes()
            chosen = set(box.Tag for box in shown if box.IsChecked)
            if not chosen:
                return
            # Senza ricerca e con tutto spuntato il filtro non serve.
            if not (search.Text or u"").strip() and len(chosen) == len(boxes):
                column_filter.values = None
            else:
                column_filter.values = chosen
            popup.IsOpen = False
            self._update_header(column_filter)
            self.on_change()

        def on_clear(sender, args):
            popup.IsOpen = False
            column_filter.values = None
            self._update_header(column_filter)
            self.on_change()

        def on_cancel(sender, args):
            popup.IsOpen = False

        for box in boxes:
            box.Checked += on_box
            box.Unchecked += on_box
        search.TextChanged += on_search
        select_all.Click += on_select_all
        ok_button.Click += on_ok
        clear_button.Click += on_clear
        cancel_button.Click += on_cancel

        fill(u"")
        self._popup = popup
        popup.IsOpen = True
        search.Dispatcher.BeginInvoke(DispatcherPriority.Input, Action(lambda: search.Focus()))
        return popup
