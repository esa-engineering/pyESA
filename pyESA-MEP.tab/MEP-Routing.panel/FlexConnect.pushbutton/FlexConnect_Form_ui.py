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


class FlexConnectWindow(object):
    """Finestra delle impostazioni di FlexConnect.
    pipe_names / duct_names / tap_names: nomi dei tipi di flexpipe, flexduct e raccordi di derivazione del modello (anche vuoti);
    settings: dict con flex_type, flex_duct_type, duct_tap, gap_mm e ask_connector.
    show() ritorna il dict aggiornato, o None se annullata."""

    def __init__(self, xaml_path, pipe_names, duct_names, tap_names, settings):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.result = None
        self.type_combo = self.window.FindName("TypeCombo")
        self.duct_combo = self.window.FindName("DuctTypeCombo")
        _fill_combo(self.type_combo, pipe_names, settings.get("flex_type"))
        _fill_combo(self.duct_combo, duct_names, settings.get("flex_duct_type"))
        self.tap_combo = self.window.FindName("TapCombo")
        _fill_combo(self.tap_combo, [DEFAULT_TAP_LABEL] + list(tap_names), settings.get("duct_tap") or DEFAULT_TAP_LABEL)
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
                       "gap_mm": gap,
                       "ask_connector": bool(self.ask_box.IsChecked)}
        self.window.DialogResult = True

    def _on_cancel(self, sender, args):
        self.window.DialogResult = False

    def show(self):
        self.window.ShowDialog()
        return self.result
