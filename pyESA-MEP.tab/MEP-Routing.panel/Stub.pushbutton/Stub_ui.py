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


def fill_combo(combo, names, selected_name):
    """Riempie il menu; ritorna False (e lo disabilita) se non ci sono voci."""
    for name in names:
        combo.Items.Add(name)
    if not names:
        combo.IsEnabled = False
        return False
    combo.SelectedIndex = names.index(selected_name) if selected_name in names else 0
    return True


class StubWindow(object):
    """Finestra modale: lunghezza, lunghezze diverse, scelta dei connettori, tipi."""

    def __init__(self, xaml_path, pipe_types, duct_types, settings):
        # pipe_types / duct_types: liste di nomi; settings: dict con le ultime impostazioni
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.length_box = self.window.FindName("LengthBox")
        self.different = self.window.FindName("DifferentCheck")
        self.mode_all = self.window.FindName("ModeAll")
        self.mode_choose = self.window.FindName("ModeChoose")
        self.pipe_combo = self.window.FindName("PipeTypeCombo")
        self.duct_combo = self.window.FindName("DuctTypeCombo")
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

        self.pipe_names, self.duct_names = pipe_types, duct_types
        self.has_pipe_types = fill_combo(self.pipe_combo, pipe_types, settings["pipe_type"])
        self.has_duct_types = fill_combo(self.duct_combo, duct_types, settings["duct_type"])

        self.length_box.Text = "{:g}".format(settings["length_mm"])
        self.different.IsChecked = bool(settings["different"])
        self.mode_choose.IsChecked = bool(settings["choose"])
        self.mode_all.IsChecked = not settings["choose"]

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        ok, length = parse_number(self.length_box.Text)
        if not ok or length <= 0:
            return False, "The length must be a number greater than zero (mm).", None
        if not (self.has_pipe_types or self.has_duct_types):
            return False, "The model has no pipe types and no duct types.", None
        return True, "", {
            "length_mm": length,
            "different": bool(self.different.IsChecked),
            "choose": bool(self.mode_choose.IsChecked),
            "pipe_type": self.pipe_names[self.pipe_combo.SelectedIndex] if self.has_pipe_types else "",
            "duct_type": self.duct_names[self.duct_combo.SelectedIndex] if self.has_duct_types else "",
        }

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Stub")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
