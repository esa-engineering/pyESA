# -*- coding: utf-8 -*-
"""
rnwflow_ui.py - finestra XAML di configurazione per RNWFlowCalculation.

Raccoglie il tipo di sistema di tubazioni degli scarichi meteorici, la serie
di tubo, il grado di riempimento dei collettori orizzontali e l'ambito, e
restituisce un RunConfig allo script. Le ultime scelte sono salvate nella
config di pyRevit (sezione ESA_RNWFlowCalculation).

La finestra non legge Revit: lo script le passa i tipi di sistema come dati
semplici (SystemTypeItem). Compatibile con Rocket mode: nessun documento o
output a livello di modulo.
"""

import os
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window, MessageBox
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Colors
from System.IO import FileStream, FileMode

from pyrevit import script

import rnwflow_tables as rt

XAML_FILE_NAME = 'RNWFlowCalculation_form.xaml'
CONFIG_SECTION = 'ESA_RNWFlowCalculation'

SCOPE_ALL = "ALL"
SCOPE_VIEW = "VIEW"
SCOPE_SELECTION = "SELECTION"

DEFAULT_FILL = 0.5
DEFAULT_SERIES = rt.SERIES_PE
DEFAULT_VMIN = "0.6"
DEFAULT_VMAX = ""

# Parole che, alla prima apertura, indicano il tipo di sistema da proporre.
RAINWATER_KEYWORDS = ("meteor", "pluv", "bianc", "storm", "rain", "rnw")

GRAY_BRUSH = SolidColorBrush(Colors.Gray)
GRAY_BRUSH.Freeze()
RED_BRUSH = SolidColorBrush(Colors.Firebrick)
RED_BRUSH.Freeze()


class SystemTypeItem(object):
    """Voce della tendina dei tipi di sistema."""

    def __init__(self, name, pipe_count, flow_count):
        self.Name = name
        self.PipeCount = pipe_count
        self.FlowCount = flow_count

    @property
    def label(self):
        return "{0}   ({1} pipes)".format(self.Name, self.PipeCount)


class RunConfig(object):
    """Scelte dell'utente."""

    def __init__(self):
        self.system_type = None
        self.fill = DEFAULT_FILL
        self.series = DEFAULT_SERIES
        self.scope = SCOPE_ALL
        self.vmin = None
        self.vmax = None


def fill_label(fill):
    return "f = {0} ({1:.0f}%) - {2}".format(fill, fill * 100, rt.TABLE_NAMES[fill])


def parse_speed(text):
    """(valore, valido) da un campo velocita': vuoto = (None, True)."""
    cleaned = (text or "").strip().replace(",", ".")
    if not cleaned:
        return None, True
    try:
        value = float(cleaned)
    except ValueError:
        return None, False
    return value, value > 0


def looks_like_rainwater(name):
    lowered = (name or "").lower()
    return any(word in lowered for word in RAINWATER_KEYWORDS)


def default_system_index(items):
    """Tipo proposto senza impostazioni salvate: nome da acque meteoriche,
    altrimenti il tipo con piu' tubi che portano portata."""
    for index, item in enumerate(items):
        if looks_like_rainwater(item.Name):
            return index
    best = 0
    for index, item in enumerate(items):
        if item.FlowCount > items[best].FlowCount:
            best = index
    return best


class RNWFlowForm(Window):
    """Finestra unica di configurazione."""

    def __init__(self, system_items, view_available, selected_pipe_count):
        self.result = False
        self.config = RunConfig()
        self._items = list(system_items)
        self._view_available = view_available
        self._selected_pipe_count = selected_pipe_count
        self._load_xaml()
        self._init_controls()
        self._restore_settings()
        self._update_system_info()

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

        self.cbo_system = root.FindName('cbo_system')
        self.txt_system_info = root.FindName('txt_system_info')
        self.cbo_series = root.FindName('cbo_series')
        self.cbo_fill = root.FindName('cbo_fill')
        self.txt_vmin = root.FindName('txt_vmin')
        self.txt_vmax = root.FindName('txt_vmax')
        self.rdo_scope_all = root.FindName('rdo_scope_all')
        self.rdo_scope_view = root.FindName('rdo_scope_view')
        self.rdo_scope_selection = root.FindName('rdo_scope_selection')
        self.btn_ok = root.FindName('btn_ok')
        self.btn_cancel = root.FindName('btn_cancel')

        self.cbo_system.SelectionChanged += self.OnSystemChanged
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel

    def _init_controls(self):
        for item in self._items:
            self.cbo_system.Items.Add(item.label)
        if self._items:
            self.cbo_system.SelectedIndex = default_system_index(self._items)

        for series in rt.SERIES_ORDER:
            self.cbo_series.Items.Add(rt.SERIES_LABELS[series])
        for fill in rt.FILL_OPTIONS:
            self.cbo_fill.Items.Add(fill_label(fill))
        self.cbo_series.SelectedIndex = rt.SERIES_ORDER.index(DEFAULT_SERIES)
        self.cbo_fill.SelectedIndex = rt.FILL_OPTIONS.index(DEFAULT_FILL)
        self.txt_vmin.Text = DEFAULT_VMIN
        self.txt_vmax.Text = DEFAULT_VMAX

        self.rdo_scope_all.IsChecked = True
        if not self._view_available:
            self.rdo_scope_view.Content = "Only pipes visible in the active view (not a model view)"
            self.rdo_scope_view.IsEnabled = False
        if self._selected_pipe_count > 0:
            self.rdo_scope_selection.Content = \
                "Only the {0} selected pipe(s)".format(self._selected_pipe_count)
            self.rdo_scope_selection.IsChecked = True
        else:
            self.rdo_scope_selection.Content = "Only the selected pipes (no pipe is selected)"
            self.rdo_scope_selection.IsEnabled = False

    # --------------------------------------------------------- tipo di sistema

    def _selected_item(self):
        index = self.cbo_system.SelectedIndex
        if 0 <= index < len(self._items):
            return self._items[index]
        return None

    def _update_system_info(self):
        item = self._selected_item()
        if item is None:
            self.txt_system_info.Foreground = RED_BRUSH
            self.txt_system_info.Text = "No piping system type in the model."
            return
        if item.FlowCount == 0:
            self.txt_system_info.Foreground = RED_BRUSH
            self.txt_system_info.Text = (
                "None of its {0} pipes has Fixture Units: there is no flow to check. "
                "Pick the rainwater drainage system type.").format(item.PipeCount)
            return
        self.txt_system_info.Foreground = GRAY_BRUSH
        self.txt_system_info.Text = "{0} pipes, {1} of them with Fixture Units.".format(
            item.PipeCount, item.FlowCount)

    def OnSystemChanged(self, sender, args):
        self._update_system_info()

    # --------------------------------------------------------------- settings

    def _restore_settings(self):
        try:
            cfg = script.get_config(CONFIG_SECTION)
            names = [item.Name for item in self._items]
            saved = cfg.get_option('system_type', None)
            if saved in names:
                self.cbo_system.SelectedIndex = names.index(saved)
            series = cfg.get_option('series', DEFAULT_SERIES)
            if series in rt.SERIES_ORDER:
                self.cbo_series.SelectedIndex = rt.SERIES_ORDER.index(series)
            fill = float(cfg.get_option('fill', DEFAULT_FILL))
            if fill in rt.FILL_OPTIONS:
                self.cbo_fill.SelectedIndex = rt.FILL_OPTIONS.index(fill)
            self.txt_vmin.Text = str(cfg.get_option('vmin', DEFAULT_VMIN))
            self.txt_vmax.Text = str(cfg.get_option('vmax', DEFAULT_VMAX))
            # La selezione corrente vince sull'ambito salvato.
            if self._selected_pipe_count == 0 and \
                    cfg.get_option('scope', SCOPE_ALL) == SCOPE_VIEW and self._view_available:
                self.rdo_scope_view.IsChecked = True
        except Exception:
            pass

    def _save_settings(self):
        try:
            cfg = script.get_config(CONFIG_SECTION)
            cfg.set_option('system_type', self.config.system_type)
            cfg.set_option('series', self.config.series)
            cfg.set_option('fill', self.config.fill)
            cfg.set_option('scope', self.config.scope)
            cfg.set_option('vmin', self.txt_vmin.Text.strip())
            cfg.set_option('vmax', self.txt_vmax.Text.strip())
            script.save_config()
        except Exception:
            pass

    # ---------------------------------------------------------------- buttons

    def _validate(self):
        """(ok, messaggio): tutti i campi controllati insieme."""
        errors = []
        if self._selected_item() is None:
            errors.append("Choose the piping system type of the rainwater drainage.")
        vmin, vmin_ok = parse_speed(self.txt_vmin.Text)
        vmax, vmax_ok = parse_speed(self.txt_vmax.Text)
        if not vmin_ok:
            errors.append("Minimum velocity must be a positive number, or empty.")
        if not vmax_ok:
            errors.append("Maximum velocity must be a positive number, or empty.")
        if vmin_ok and vmax_ok and vmin is not None and vmax is not None and vmax <= vmin:
            errors.append("Maximum velocity must be higher than the minimum.")
        return (not errors), "\n".join(errors)

    def OnOK(self, sender, args):
        ok, message = self._validate()
        if not ok:
            MessageBox.Show(message, "Rainwater Flow Calculation")
            return
        item = self._selected_item()
        self.config.vmin = parse_speed(self.txt_vmin.Text)[0]
        self.config.vmax = parse_speed(self.txt_vmax.Text)[0]
        self.config.system_type = item.Name
        self.config.series = rt.SERIES_ORDER[max(self.cbo_series.SelectedIndex, 0)]
        self.config.fill = rt.FILL_OPTIONS[max(self.cbo_fill.SelectedIndex, 0)]
        if self.rdo_scope_selection.IsChecked:
            self.config.scope = SCOPE_SELECTION
        elif self.rdo_scope_view.IsChecked:
            self.config.scope = SCOPE_VIEW
        else:
            self.config.scope = SCOPE_ALL
        self._save_settings()
        self.result = True
        self.Close()

    def OnCancel(self, sender, args):
        self.result = False
        self.Close()


def show_dialog(system_items, view_available, selected_pipe_count):
    """Mostra la finestra; ritorna RunConfig oppure None se annullata."""
    form = RNWFlowForm(system_items, view_available, selected_pipe_count)
    form.ShowDialog()
    if not form.result:
        return None
    return form.config
