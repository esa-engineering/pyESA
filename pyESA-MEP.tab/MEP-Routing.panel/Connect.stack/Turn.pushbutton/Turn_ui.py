# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


class TurnWindow(object):
    """Finestra delle impostazioni di Turn (SHIFT + CLICK): lunghezza del nuovo tratto e pendenza.
    settings: dict con length_mm e keep_slope; show() ritorna lo stesso dict aggiornato, o None se annullata."""

    def __init__(self, xaml_path, settings, min_length_mm):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.settings = dict(settings)
        self.min_length = min_length_mm
        self.length_box = self.window.FindName("LengthBox")
        self.length_box.Text = "{:g}".format(settings["length_mm"])
        self.keep_slope = self.window.FindName("KeepSlope")
        self.keep_slope.IsChecked = bool(settings["keep_slope"])
        self.window.FindName("OkButton").Click += self._on_ok
        self.window.FindName("CancelButton").Click += self._on_cancel
        self.result = None

    def _on_ok(self, sender, args):
        ok, length = parse_number(self.length_box.Text)
        if not ok or length < self.min_length:
            MessageBox.Show("The length must be a number of at least {:g} mm.".format(self.min_length), "Turn")
            return
        self.settings.update({"length_mm": length, "keep_slope": bool(self.keep_slope.IsChecked)})
        self.result = self.settings
        self.window.DialogResult = True

    def _on_cancel(self, sender, args):
        self.window.DialogResult = False

    def show(self):
        self.window.ShowDialog()
        return self.result
