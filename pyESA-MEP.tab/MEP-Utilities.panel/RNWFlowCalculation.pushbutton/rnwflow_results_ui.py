# -*- coding: utf-8 -*-
"""
rnwflow_results_ui.py - finestra dei risultati per tratto di RNWFlowCalculation.

Una riga per tratto (SectionResult di rnwflow_check). L'utente sceglie da una
tendina la nuova misura fra quelle del tipo di tubazione (PipeSize del tratto,
lette dallo script dai segmenti del tipo) e spunta i tratti da modificare. A
ogni scelta la riga ricalcola portata ammissibile, velocita', riempimento ed
esito con la nuova misura (SectionResult.evaluate_at).

La finestra e' modale e non tocca Revit: si chiude restituendo allo script
un'azione (kind):
  * "apply": le modifiche scelte (tratto, PipeSize);
  * "show": i tubi da selezionare (pulsante della riga o "Select in Model"); lo
    script li seleziona, lascia navigare la vista e riapre la finestra con lo
    stesso ResultsState, cioe' con spunte, misure, ricerca e filtro di prima;
  * None: chiusa senza modifiche.

Griglia legata a una System.Data.DataTable come in MEPQTO: il binding a due
vie lo gestisce .NET, senza oggetti INotifyPropertyChanged in IronPython. La
tendina di ogni riga ha le sue voci nella colonna Options (oggetto List[str]).

Compatibile con Rocket mode: nessun documento o output a livello di modulo.
"""

import os
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Data')

import System
from System import Action
from System.Collections.Generic import List
from System.Data import DataTable, DataRowVersion
from System.Windows import (Window, MessageBox, MessageBoxButton, MessageBoxResult,
                            RoutedEventHandler)
from System.Windows.Controls import DataGridEditingUnit, Button
from System.Windows.Controls.Primitives import ButtonBase
from System.Windows.Markup import XamlReader
from System.Windows.Threading import DispatcherPriority
from System.IO import FileStream, FileMode

from pyrevit import script

import rnwflow_check as rc

XAML_FILE_NAME = 'RNWFlowCalculation_results.xaml'
CONFIG_SECTION = 'ESA_RNWFlowCalculation'
WINDOW_TITLE = 'Rainwater Sections'

MAX_LINES_SHOWN = 10

SEVERITY = {
    rc.STATUS_UNDERSIZED: "error",
    rc.STATUS_BEYOND: "error",
    rc.STATUS_NOT_IN_TABLE: "warning",
    rc.STATUS_LOW_SLOPE: "warning",
    rc.STATUS_HIGH_VELOCITY: "warning",
    rc.STATUS_LOW_VELOCITY: "warning",
    rc.STATUS_OVERSIZED: "info",
    rc.STATUS_OK: "",
    rc.STATUS_NO_FLOW: "none",
}

# Esiti della misura nuova che non richiedono conferma esplicita
NEW_CHECK_PASS = (rc.STATUS_OK, rc.STATUS_NO_FLOW)


def fmt_de(value):
    if value is None:
        return ""
    return rc.fmt_mm(value)


def escape_like(text):
    """Escape di un testo per RowFilter LIKE (stesse regole di MEPQTO)."""
    out = []
    for ch in text:
        if ch in "[]*%":
            out.append("[" + ch + "]")
        elif ch == "'":
            out.append("''")
        else:
            out.append(ch)
    return "".join(out)


def _value(number):
    return number if number is not None else System.DBNull.Value


def _pct(fraction):
    return fraction * 100.0 if fraction is not None else None


def _text(value):
    if value is None or value is System.DBNull.Value:
        return ""
    return str(value)


KIND_APPLY = "apply"
KIND_SHOW = "show"


class ResultsState(object):
    """Stato della tabella che sopravvive alla chiusura della finestra."""

    def __init__(self):
        self.table = None
        self.by_key = {}
        self.search = ""
        self.only_action = None
        self.last_key = None


class ResultsForm(Window):
    """Finestra dei tratti."""

    def __init__(self, sections, settings_text, state=None):
        self.sections = list(sections)
        self.changes = None
        self.kind = None
        self.show_ids = []
        self.state = state if state is not None else ResultsState()
        self._by_key = self.state.by_key
        self._loading = True
        self._column_handler = self.OnColumnChanged
        self._load_xaml()
        self.txt_settings.Text = settings_text
        if self.state.table is None:
            self._build_table()
            self._restore_settings()
        else:
            self.table = self.state.table
            self.table.ColumnChanged += self._column_handler
            self.dg_sections.ItemsSource = self.table.DefaultView
            self.chk_only_action.IsChecked = bool(self.state.only_action)
            self.txt_search.Text = self.state.search or ""
        self._loading = False
        self._apply_filter()

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

        self.txt_settings = root.FindName('txt_settings')
        self.txt_search = root.FindName('txt_search')
        self.btn_clear_search = root.FindName('btn_clear_search')
        self.chk_only_action = root.FindName('chk_only_action')
        self.txt_count = root.FindName('txt_count')
        self.dg_sections = root.FindName('dg_sections')
        self.btn_tick_resize = root.FindName('btn_tick_resize')
        self.btn_untick_all = root.FindName('btn_untick_all')
        self.btn_show = root.FindName('btn_show')
        self.btn_ok = root.FindName('btn_ok')
        self.btn_cancel = root.FindName('btn_cancel')

        self.txt_search.TextChanged += self.OnFilterChanged
        self.btn_clear_search.Click += self.OnClearSearch
        self.chk_only_action.Checked += self.OnFilterChanged
        self.chk_only_action.Unchecked += self.OnFilterChanged
        self.btn_tick_resize.Click += self.OnTickResize
        self.btn_untick_all.Click += self.OnUntickAll
        self.btn_show.Click += self.OnShow
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel
        # Pulsante "seleziona" di ogni riga: un solo gestore sulla griglia.
        self.dg_sections.AddHandler(ButtonBase.ClickEvent, RoutedEventHandler(self.OnGridButton))
        self.Loaded += self.OnWindowLoaded
        self.Closed += self.OnWindowClosed

    def OnWindowLoaded(self, sender, args):
        """Riporta in vista la riga da cui si era partiti per navigare."""
        key = self.state.last_key
        if key is None:
            return
        for item in self.table.DefaultView:
            if item["Key"] == key:
                try:
                    self.dg_sections.ScrollIntoView(item)
                except Exception:
                    pass
                break

    def OnWindowClosed(self, sender, args):
        try:
            self.table.ColumnChanged -= self._column_handler
        except Exception:
            pass
        self.state.table = self.table
        self.state.search = self.txt_search.Text or ""
        self.state.only_action = bool(self.chk_only_action.IsChecked)

    def _build_table(self):
        table = DataTable("Sections")
        columns = [
            ("Key", System.String), ("Apply", System.Boolean), ("System", System.String),
            ("Section", System.String), ("SectionSort", System.Double),
            ("Pipes", System.Int32), ("Run", System.String), ("Slope", System.Double),
            ("Flow", System.Double), ("ModelDn", System.String), ("ModelDe", System.String),
            ("Capacity", System.Double), ("Use", System.Double), ("VModel", System.Double),
            ("FillModel", System.Double), ("Suggested", System.String),
            ("Options", System.Object), ("HasSizes", System.Boolean),
            ("NewSize", System.String), ("NewDeSort", System.Double),
            ("QNew", System.Double), ("VNew", System.Double), ("FillNew", System.Double),
            ("CheckNew", System.String),
            ("Status", System.String), ("Notes", System.String), ("Severity", System.String),
            ("Action", System.Boolean),
        ]
        for name, kind in columns:
            table.Columns.Add(name, kind)

        for index, section in enumerate(self.sections):
            key = str(index)
            self._by_key[key] = section
            options = List[System.String]()
            for size in section.sizes or []:
                options.Add(size.label)
            if section.suggested_size is not None:
                default = section.suggested_size
            else:
                default = section.model_size

            row = table.NewRow()
            row["Key"] = key
            row["Apply"] = bool(section.default_apply)
            row["System"] = section.system
            row["Section"] = section.section if section.section else "-"
            row["SectionSort"] = section.section_sort()
            row["Pipes"] = len(section.pipe_ids)
            row["Run"] = section.run
            row["Slope"] = _value(section.min_slope)
            row["Flow"] = section.flow
            row["ModelDn"] = section.model_dn_label
            row["ModelDe"] = section.model_de_label
            row["Capacity"] = _value(section.capacity)
            row["Use"] = _value(section.use_pct)
            row["VModel"] = _value(section.v_min)
            row["FillModel"] = _value(_pct(section.fill_max))
            row["Suggested"] = fmt_de(section.suggested_de)
            row["Options"] = options
            row["HasSizes"] = options.Count > 0
            row["NewSize"] = default.label if default is not None else ""
            row["Status"] = section.status
            row["Notes"] = "; ".join(section.notes)
            row["Severity"] = SEVERITY.get(section.status, "")
            row["Action"] = bool(section.needs_action)
            self._fill_new(row, section)
            table.Rows.Add(row)

        table.AcceptChanges()
        table.ColumnChanged += self._column_handler
        self.table = table
        self.dg_sections.ItemsSource = table.DefaultView

    # ------------------------------------------------------- misura nuova

    def _size_of(self, row, section):
        label = _text(row["NewSize"])
        for size in section.sizes or []:
            if size.label == label:
                return size
        return None

    def _fill_new(self, row, section):
        """Colonne della misura nuova: Qmax, v, riempimento, esito."""
        size = self._size_of(row, section)
        if size is None:
            for name in ("NewDeSort", "QNew", "VNew", "FillNew"):
                row[name] = System.DBNull.Value
            row["CheckNew"] = ""
            return
        result = section.evaluate_at(size.od_mm)
        row["NewDeSort"] = size.od_mm
        row["QNew"] = _value(result.capacity)
        row["VNew"] = _value(result.velocity)
        row["FillNew"] = _value(_pct(result.fill_actual))
        row["CheckNew"] = result.status

    # --------------------------------------------------------------- settings

    def _restore_settings(self):
        try:
            cfg = script.get_config(CONFIG_SECTION)
            self.chk_only_action.IsChecked = bool(cfg.get_option('results_only_action', True))
        except Exception:
            self.chk_only_action.IsChecked = True

    def _save_settings(self):
        try:
            cfg = script.get_config(CONFIG_SECTION)
            cfg.set_option('results_only_action', bool(self.chk_only_action.IsChecked))
            script.save_config()
        except Exception:
            pass

    # ------------------------------------------------------------------ filtro

    def _apply_filter(self):
        parts = []
        if self.chk_only_action.IsChecked:
            parts.append("Action = true")
        needle = (self.txt_search.Text or "").strip()
        if needle:
            like = escape_like(needle)
            parts.append("(" + " OR ".join(
                "{0} LIKE '%{1}%'".format(col, like)
                for col in ("System", "Section", "Status", "Notes", "Run")) + ")")
        self.table.DefaultView.RowFilter = " AND ".join(parts)
        self._update_count()

    def _update_count(self):
        shown = self.table.DefaultView.Count
        ticked = 0
        pipes = 0
        for row in self.table.Rows:
            if row["Apply"]:
                ticked += 1
                pipes += row["Pipes"]
        self.txt_count.Text = "{0} of {1} sections shown, {2} ticked ({3} pipes)".format(
            shown, self.table.Rows.Count, ticked, pipes)

    def OnFilterChanged(self, sender, args):
        if not self._loading:
            self._apply_filter()

    def OnClearSearch(self, sender, args):
        self.txt_search.Text = ""

    # ------------------------------------------------------------------ righe

    def OnColumnChanged(self, sender, args):
        if self._loading:
            return
        name = args.Column.ColumnName
        if name == "NewSize":
            row = args.Row
            section = self._by_key[row["Key"]]
            self._loading = True
            try:
                self._fill_new(row, section)
                # Scegliere una misura diversa da quella del modello spunta il
                # tratto: e' l'intenzione evidente.
                size = self._size_of(row, section)
                if size is not None and not rc.od_in(size.od_mm, section.model_des) \
                        and not row["Apply"]:
                    row["Apply"] = True
            finally:
                self._loading = False
            self._defer(lambda: self._end_edit(row))
            self._update_count()
        elif name == "Apply":
            self._update_count()

    def _end_edit(self, row):
        """Chiude la modifica della riga, cosi' la griglia mostra subito le colonne
        ricalcolate anche se il binding ha aperto una modifica sul DataRowView."""
        try:
            if row.HasVersion(DataRowVersion.Proposed):
                row.EndEdit()
        except Exception:
            pass

    def _defer(self, func):
        try:
            self.Dispatcher.BeginInvoke(DispatcherPriority.Background, Action(func))
        except Exception:
            func()

    def _set_rows(self, chooser):
        self._commit_edits()
        self._loading = True
        try:
            for row in self.table.Rows:
                chooser(self._by_key[row["Key"]], row)
        finally:
            self._loading = False
        self._update_count()

    def OnTickResize(self, sender, args):
        def tick(section, row):
            if section.default_apply:
                row["Apply"] = True
                row["NewSize"] = section.suggested_size.label
                self._fill_new(row, section)
        self._set_rows(tick)

    def OnUntickAll(self, sender, args):
        def untick(section, row):
            row["Apply"] = False
        self._set_rows(untick)

    def _highlighted_sections(self):
        found = []
        seen = set()
        for cell in self.dg_sections.SelectedCells:
            item = cell.Item
            try:
                key = item["Key"]
            except Exception:
                continue
            if key not in seen:
                seen.add(key)
                found.append(self._by_key[key])
        return found

    def _request_show(self, keys):
        """Chiude la finestra chiedendo allo script di selezionare i tubi delle righe."""
        self._commit_edits()
        ids = []
        for key in keys:
            ids.extend(self._by_key[key].pipe_ids)
        self.kind = KIND_SHOW
        self.show_ids = ids
        self.state.last_key = keys[0]
        self._save_settings()
        self.Close()

    def OnShow(self, sender, args):
        sections = self._highlighted_sections()
        if not sections:
            MessageBox.Show("Highlight one or more rows first, or use the button of a row.",
                            WINDOW_TITLE)
            return
        keys = [key for key, section in self._by_key.items() if section in sections]
        keys.sort(key=lambda k: int(k))
        self._request_show(keys)

    def OnGridButton(self, sender, args):
        source = args.OriginalSource
        if not isinstance(source, Button):
            return
        try:
            key = source.DataContext["Key"]
        except Exception:
            return
        args.Handled = True
        self._request_show([key])

    def _commit_edits(self):
        try:
            self.dg_sections.CommitEdit(DataGridEditingUnit.Cell, True)
            self.dg_sections.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            pass

    # -------------------------------------------------------------- validazione

    def _label(self, section, row):
        return "{0} - section {1} {2}".format(section.system, row["Section"], section.run)

    def _part_warnings(self, changes):
        """Section con le due parti (orizzontale e verticale) trattate diversamente:
        la curva fra le due ha una sola misura e potrebbe bloccare la modifica."""
        chosen = {}
        for section, size in changes:
            chosen.setdefault(section.group_key, []).append((section, size))
        parts = {}
        for section in self.sections:
            parts.setdefault(section.group_key, []).append(section)
        warnings = []
        for key, items in chosen.items():
            siblings = parts.get(key, [])
            if len(siblings) < 2:
                continue
            label = "{0} - section {1}".format(key[0], items[0][0].section or "-")
            if len(items) < len(siblings):
                warnings.append("{0}: only the {1} part is ticked".format(
                    label, items[0][0].run.lower()))
            elif len(set(round(size.od_mm, 1) for _, size in items)) > 1:
                warnings.append("{0}: the two parts get different sizes".format(label))
        return warnings

    def _collect_changes(self):
        """Ritorna (modifiche, errori, esiti non superati).

        Modifica = (SectionResult, PipeSize).
        """
        changes = []
        errors = []
        failing = []
        for row in self.table.Rows:
            if not row["Apply"]:
                continue
            section = self._by_key[row["Key"]]
            size = self._size_of(row, section)
            if size is None:
                if not section.sizes:
                    errors.append("{0}: the sizes of the pipe type are unknown.".format(
                        self._label(section, row)))
                else:
                    errors.append("{0}: choose the new size.".format(self._label(section, row)))
                continue
            check = _text(row["CheckNew"])
            if check and check not in NEW_CHECK_PASS:
                failing.append("{0}: {1} with {2}".format(
                    self._label(section, row), check, size.label))
            changes.append((section, size))
        return changes, errors, failing

    @staticmethod
    def _join(lines):
        text = "\n".join(lines[:MAX_LINES_SHOWN])
        if len(lines) > MAX_LINES_SHOWN:
            text += "\n... and {0} more.".format(len(lines) - MAX_LINES_SHOWN)
        return text

    # ---------------------------------------------------------------- buttons

    def OnOK(self, sender, args):
        self._commit_edits()
        changes, errors, failing = self._collect_changes()
        if errors:
            MessageBox.Show(self._join(errors), WINDOW_TITLE)
            return
        if not changes:
            MessageBox.Show("No section is ticked. Tick Apply on the sections to resize, "
                            "or close the window.", WINDOW_TITLE)
            return
        pipes = sum(len(section.pipe_ids) for section, _ in changes)
        question = "Resize {0} pipe(s) in {1} section(s)?".format(pipes, len(changes))
        if failing:
            question += ("\n\nWith the chosen size these sections still do not pass "
                         "the check:\n" + self._join(failing))
        warnings = self._part_warnings(changes)
        if warnings:
            question += ("\n\nA fitting between the horizontal and the vertical part has "
                         "one size, so these sections may be rolled back:\n" +
                         self._join(warnings))
        answer = MessageBox.Show(question, WINDOW_TITLE, MessageBoxButton.YesNo)
        if answer != MessageBoxResult.Yes:
            return
        self.changes = changes
        self.kind = KIND_APPLY
        self._save_settings()
        self.Close()

    def OnCancel(self, sender, args):
        self.changes = None
        self.kind = None
        self._save_settings()
        self.Close()


def show_results(sections, settings_text, state=None):
    """Mostra la finestra (modale) e la ritorna chiusa: kind, changes, show_ids, state."""
    form = ResultsForm(sections, settings_text, state)
    form.ShowDialog()
    return form
