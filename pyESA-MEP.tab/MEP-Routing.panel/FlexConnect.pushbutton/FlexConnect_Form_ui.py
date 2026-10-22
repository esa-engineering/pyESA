# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader

MIN_GAP_MM = 1.0
DEFAULT_TAP_LABEL = "(routing preferences of the duct type)"
DEFAULT_PATTERN_LABEL = "(type default)"


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


def _fill_combo(combo, names, saved):
    """Riempie il combo; senza tipi nel modello resta vuoto e disabilitato."""
    for name in names:
        combo.Items.Add(name)
    if names:
        combo.SelectedItem = saved if saved in names else names[0]
    else:
        combo.IsEnabled = False


def _fill_pattern_combo(combo, patterns, saved_value):
    """Combo del Flex Pattern: prima voce = default del tipo, poi le voci scoperte [(nome, valore)]."""
    combo.Items.Add(DEFAULT_PATTERN_LABEL)
    selected = DEFAULT_PATTERN_LABEL
    for name, value in patterns:
        combo.Items.Add(name)
        if str(value) == str(saved_value):
            selected = name
    combo.SelectedItem = selected
    combo.IsEnabled = bool(patterns)


def _pattern_value(combo, patterns):
    """Valore intero scelto, come testo ("" = default del tipo)."""
    for name, value in patterns:
        if name == combo.SelectedItem:
            return str(value)
    return ""


class FlexConnectWindow(object):
    """Finestra delle impostazioni di FlexConnect.
    pipe_names / duct_names / tap_names: nomi dei tipi di flexpipe, flexduct e raccordi di derivazione del modello (anche vuoti);
    pipe_patterns / duct_patterns: voci [(nome, valore)] di Flex Pattern (anche vuote);
    settings: dict con flex_type, flex_duct_type, duct_tap, flex_pipe_pattern, flex_duct_pattern, gap_mm e ask_connector.
    show() ritorna il dict aggiornato, o None se annullata."""

    def __init__(self, xaml_path, pipe_names, duct_names, tap_names, pipe_patterns, duct_patterns, settings):
        self.pipe_patterns = pipe_patterns
        self.duct_patterns = duct_patterns
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.result = None
        self.type_combo = self.window.FindName("TypeCombo")
        self.duct_combo = self.window.FindName("DuctTypeCombo")
        _fill_combo(self.type_combo, pipe_names, settings.get("flex_type"))
        _fill_combo(self.duct_combo, duct_names, settings.get("flex_duct_type"))
        self.tap_combo = self.window.FindName("TapCombo")
        _fill_combo(self.tap_combo, [DEFAULT_TAP_LABEL] + list(tap_names), settings.get("duct_tap") or DEFAULT_TAP_LABEL)
        self.pipe_pattern_combo = self.window.FindName("PipePatternCombo")
        self.duct_pattern_combo = self.window.FindName("DuctPatternCombo")
        _fill_pattern_combo(self.pipe_pattern_combo, pipe_patterns, settings.get("flex_pipe_pattern"))
        _fill_pattern_combo(self.duct_pattern_combo, duct_patterns, settings.get("flex_duct_pattern"))
        self.gap_box = self.window.FindName("GapBox")
        self.gap_box.Text = "{:g}".format(settings["gap_mm"])
        self.ask_box = self.window.FindName("AskConnectorBox")
        self.ask_box.IsChecked = bool(settings.get("ask_connector"))
        self.window.FindName("OkButton").Click += self._on_ok
        self.window.FindName("CancelButton").Click += self._on_cancel

    def _on_ok(self, sender, args):
        ok, gap = parse_number(self.gap_box.Text)
        if not ok or gap < MIN_GAP_MM:
            MessageBox.Show("The reducer offset must be a number of at least {:g} mm.".format(MIN_GAP_MM),
                            "FlexConnect")
            return
        self.result = {"flex_type": self.type_combo.SelectedItem or "",
                       "flex_duct_type": self.duct_combo.SelectedItem or "",
                       "duct_tap": "" if self.tap_combo.SelectedItem == DEFAULT_TAP_LABEL else self.tap_combo.SelectedItem,
                       "flex_pipe_pattern": _pattern_value(self.pipe_pattern_combo, self.pipe_patterns),
                       "flex_duct_pattern": _pattern_value(self.duct_pattern_combo, self.duct_patterns),
                       "gap_mm": gap,
                       "ask_connector": bool(self.ask_box.IsChecked)}
        self.window.DialogResult = True

    def _on_cancel(self, sender, args):
        self.window.DialogResult = False

    def show(self):
        self.window.ShowDialog()
        return self.result
