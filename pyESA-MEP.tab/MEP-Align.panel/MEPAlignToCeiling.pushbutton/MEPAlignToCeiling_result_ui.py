# -*- coding: utf-8 -*-
"""Finestra di esito di MEPAlignToCeiling (MEPAlignToCeiling_result.xaml).

In testa i totali (allineati, gia' a posto, non allineati); sotto la tabella
dei soli elementi NON allineati, con le note. La tabella e' interrogabile
(ricerca su tutte le colonne, filtri per valore nell'intestazione,
ordinamento per colonna).

La finestra e' NON MODALE: resta aperta a comando finito, quindi Revit e'
libero e l'utente puo' spostare gli elementi e cambiarne i parametri mentre la
tabella e' a video. Il pulsante di una riga (o "Show All Listed") chiama
info["show"](chiavi), "Advanced Report" chiama info["report"](): lo script le
esegue tramite un ExternalEvent, l'unico modo di usare la Revit API da una
finestra non modale. set_status() mostra l'esito dell'ultima richiesta.

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
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Markup import XamlReader

import MEPAlignToCeiling_grid_filter as grid_filter

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


class ResultWindow(object):
    """Totali e tabella dei non allineati, in una finestra non modale."""

    def __init__(self, xaml_path, info):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.info = info

        find = self.window.FindName
        self.window.Title = info["title"]
        find("SummaryText").Text = info["summary"]
        self.grid = find("ResultGrid")
        self.search_box = find("SearchBox")
        self.row_count = find("RowCount")
        self.empty_text = find("EmptyText")
        self.status_text = find("StatusText")
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

    def set_status(self, text):
        """Esito dell'ultima richiesta, sotto la tabella."""
        self.status_text.Text = text or u""

    def _on_window_closed(self, sender, args):
        closed = self.info.get("closed")
        if closed is not None:
            closed()

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

    def _request_show(self, keys):
        """Chiede allo script di selezionare e inquadrare gli elementi.

        La finestra resta aperta: l'utente lavora in Revit e torna qui quando
        vuole.
        """
        if not keys:
            self.set_status(u"No row to show.")
            return
        self.set_status(u"Selecting {} element(s) in the model...".format(len(keys)))
        self.info["show"](list(keys))

    def _on_row_button(self, sender, args):
        key = self._button_key(args)
        if key is None:
            return
        args.Handled = True
        self.grid.SelectedItem = args.OriginalSource.DataContext
        self._request_show([key])

    def _on_show_all(self, sender, args):
        self._request_show([int(item["Key"]) for item in self.table.DefaultView])

    # --- report e chiusura ---------------------------------------------

    def _on_advanced(self, sender, args):
        self.set_status(u"Printing the advanced report in the pyRevit output panel...")
        self.info["report"]()

    def _on_close(self, sender, args):
        self.window.Close()

    def show_modeless(self, owner_handle=None):
        """Mostra la finestra senza bloccare Revit.

        Owner = finestra di Revit: resta davanti a Revit e si riduce con lui.
        EnableModelessKeyboardInterop fa arrivare i tasti alla ricerca invece
        che alle scorciatoie di Revit.
        """
        if owner_handle is not None:
            try:
                WindowInteropHelper(self.window).Owner = owner_handle
            except Exception:
                pass
        try:
            clr.AddReference("WindowsFormsIntegration")
            from System.Windows.Forms.Integration import ElementHost
            ElementHost.EnableModelessKeyboardInterop(self.window)
        except Exception:
            pass
        self.window.Show()
        return self
