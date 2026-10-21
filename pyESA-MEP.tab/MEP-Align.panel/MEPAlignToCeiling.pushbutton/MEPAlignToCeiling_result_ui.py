# -*- coding: utf-8 -*-
"""Finestra di esito di MEPAlignToCeiling (MEPAlignToCeiling_result.xaml).

In testa i totali (allineati, gia' a posto, non allineati); sotto la tabella
dei soli elementi NON allineati, con le note. La tabella e' interrogabile
(ricerca su tutte le colonne, filtri per valore nell'intestazione,
ordinamento per colonna).

Navigazione, come in RNWFlowCalculation: il pulsante di una riga (o "Show All
Listed") chiude la finestra con kind = KIND_SHOW e le chiavi delle righe; lo
script seleziona gli elementi, li inquadra e lascia navigare la vista attiva
finche' l'utente non preme Fine o Esc, poi riapre la finestra passando lo
stesso ResultState (ricerca, filtri, riga da riportare in vista). Una finestra
non modale richiederebbe ExternalEvent.

Il modulo non conosce la Revit API. Nessun riferimento a documento o output a
livello di modulo (Rocket mode).
"""
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Data")

import System
from System import DBNull
from System.Data import DataTable
from System.IO import File
from System.Windows import RoutedEventHandler, Visibility
from System.Windows.Controls import Button
from System.Windows.Controls.Primitives import ButtonBase
from System.Windows.Markup import XamlReader

import MEPAlignToCeiling_grid_filter as grid_filter


KIND_CLOSE = "close"
KIND_SHOW = "show"
KIND_REPORT = "report"

# (colonna della DataTable, chiave della riga, tipo .NET). Key resta nascosta:
# lega la riga all'elemento.
COLUMNS = [
    ("Key", "key", System.Int32),
    ("Status", "status", System.String),
    ("Id", "id", System.Double),
    ("Category", "category", System.String),
    ("Type", "type", System.String),
    ("Level", "level", System.String),
    ("Target", "target", System.String),
    ("Move", "move_mm", System.Double),
    ("Buzzer", "buzzer_mm", System.Double),
    ("Notes", "notes", System.String),
]
SEARCH_TEXT_COLUMNS = ("Status", "Category", "Type", "Level", "Target", "Notes")
SEARCH_NUMBER_COLUMNS = ("Id",)


def escape_like(text):
    """Testo letterale dentro un LIKE di DataView.RowFilter."""
    escaped = []
    for char in text:
        if char in u"[]*%":
            escaped.append(u"[{}]".format(char))
        elif char == u"'":
            escaped.append(u"''")
        else:
            escaped.append(char)
    return u"".join(escaped)


def build_table(rows):
    table = DataTable("Result")
    for name, _key, net_type in COLUMNS:
        table.Columns.Add(name, clr.GetClrType(net_type))
    for row in rows:
        data_row = table.NewRow()
        for name, key, _net_type in COLUMNS:
            value = row.get(key)
            data_row[name] = DBNull.Value if value is None else value
        table.Rows.Add(data_row)
    return table


class ResultState(object):
    """Stato della finestra che sopravvive alla navigazione nella vista."""

    def __init__(self):
        self.search = u""
        self.filter_values = {}     # campo -> insieme dei valori ammessi
        self.last_key = None        # riga da riportare in vista alla riapertura


class ResultWindow(object):
    """Totali e tabella dei non allineati. show() ritorna se stesso: kind,
    show_keys e state dicono allo script cosa fare."""

    def __init__(self, xaml_path, info, state=None):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.info = info
        self.state = state if state is not None else ResultState()
        self.kind = KIND_CLOSE
        self.show_keys = []

        find = self.window.FindName
        self.window.Title = info["title"]
        find("SummaryText").Text = info["summary"]
        self.grid = find("ResultGrid")
        self.search_box = find("SearchBox")
        self.row_count = find("RowCount")
        self.empty_text = find("EmptyText")
        # Titolo da impostare PRIMA di GridFilters, che sostituisce le
        # intestazioni con titolo + imbuto.
        find("BuzzerColumn").Header = info["buzzer_header"]

        self.table = build_table(info["rows"])
        self.grid.ItemsSource = self.table.DefaultView
        self.filters = grid_filter.GridFilters(
            self.grid, self.table, self._apply_filter,
            context=self._search_expression,
            formatters={"Id": lambda value: u"{:.0f}".format(value),
                        "Move": lambda value: u"{:+.0f}".format(value),
                        "Buzzer": lambda value: u"{:.0f}".format(value)})
        self._restore_state()

        self.search_box.TextChanged += self._on_search
        find("ClearSearchButton").Click += self._on_clear_search
        find("ClearFiltersButton").Click += self._on_clear_filters
        find("ShowAllButton").Click += self._on_show_all
        find("AdvancedButton").Click += self._on_advanced
        find("CloseButton").Click += self._on_close
        # Un solo gestore per i pulsanti di tutte le righe, come in RNWFlowCalculation.
        self.grid.AddHandler(ButtonBase.ClickEvent, RoutedEventHandler(self._on_row_button))
        self.window.Closed += self._on_window_closed

        self._apply_filter()
        self._scroll_to_last()

    # --- stato ---------------------------------------------------------

    def _restore_state(self):
        self.search_box.Text = self.state.search or u""
        for field, values in self.state.filter_values.items():
            column_filter = self.filters.filters.get(field)
            if column_filter is None:
                continue
            column_filter.values = set(values)
            self.filters._update_header(column_filter)

    def _save_state(self):
        self.state.search = self.search_box.Text or u""
        self.state.filter_values = dict(
            (field, set(f.values)) for field, f in self.filters.filters.items() if f.active)

    def _scroll_to_last(self):
        key = self.state.last_key
        if key is None:
            return
        for item in self.table.DefaultView:
            if item["Key"] == key:
                try:
                    self.grid.SelectedItem = item
                    self.grid.ScrollIntoView(item)
                except Exception:
                    pass
                break

    def _on_window_closed(self, sender, args):
        self._save_state()

    # --- filtri --------------------------------------------------------

    def _search_expression(self):
        text = (self.search_box.Text or u"").strip()
        if not text:
            return u""
        pattern = u"'*{}*'".format(escape_like(text))
        parts = [u"ISNULL([{}], '') LIKE {}".format(name, pattern) for name in SEARCH_TEXT_COLUMNS]
        parts += [u"CONVERT([{}], 'System.String') LIKE {}".format(name, pattern)
                  for name in SEARCH_NUMBER_COLUMNS]
        return u"({})".format(u" OR ".join(parts))

    def _apply_filter(self):
        column_filters = self.filters.expression() if hasattr(self, "filters") else u""
        parts = [part for part in (self._search_expression(), column_filters) if part]
        self.table.DefaultView.RowFilter = u" AND ".join(u"({})".format(p) for p in parts)
        shown = self.table.DefaultView.Count
        self.row_count.Text = u"{} of {} rows".format(shown, self.table.Rows.Count)
        self.empty_text.Visibility = Visibility.Visible if self.table.Rows.Count == 0 \
            else Visibility.Collapsed

    def _on_search(self, sender, args):
        self._apply_filter()

    def _on_clear_search(self, sender, args):
        self.search_box.Text = u""

    def _on_clear_filters(self, sender, args):
        self.search_box.Text = u""
        self.filters.clear()

    # --- navigazione nel modello ---------------------------------------

    @staticmethod
    def _button_key(args):
        """Chiave della riga del pulsante premuto; None per gli altri pulsanti
        della griglia (gli imbuti dei filtri nelle intestazioni)."""
        source = args.OriginalSource
        if not isinstance(source, Button):
            return None
        try:
            return int(source.DataContext["Key"])
        except Exception:
            return None

    def _request_show(self, keys, last_key):
        """Chiude la finestra chiedendo allo script di mostrare gli elementi."""
        if not keys:
            return
        self.kind = KIND_SHOW
        self.show_keys = list(keys)
        self.state.last_key = last_key
        self.window.Close()

    def _on_row_button(self, sender, args):
        key = self._button_key(args)
        if key is None:
            return
        args.Handled = True
        self._request_show([key], key)

    def _on_show_all(self, sender, args):
        keys = [int(item["Key"]) for item in self.table.DefaultView]
        self._request_show(keys, None)

    # --- chiusura ------------------------------------------------------

    def _on_advanced(self, sender, args):
        self.kind = KIND_REPORT
        self.window.Close()

    def _on_close(self, sender, args):
        self.kind = KIND_CLOSE
        self.window.Close()

    def show(self):
        self.window.ShowDialog()
        return self


def show_result(xaml_path, info, state=None):
    """Mostra la finestra; ritorna la finestra chiusa (kind, show_keys, state)."""
    return ResultWindow(xaml_path, info, state).show()
