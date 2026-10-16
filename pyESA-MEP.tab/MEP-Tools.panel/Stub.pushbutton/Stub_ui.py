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


# chiave del tipo di canale -> nome del menu nella finestra
DUCT_COMBOS = (("duct_round", "DuctRoundCombo"), ("duct_rect", "DuctRectCombo"), ("duct_oval", "DuctOvalCombo"))


class StubWindow(object):
    """Finestra modale: lunghezza, lunghezze diverse, scelta dei connettori, tipi."""

    def __init__(self, xaml_path, pipe_types, duct_types, settings):
        # pipe_types: lista di nomi; duct_types: {chiave forma: lista di nomi}; settings: ultime impostazioni
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.length_box = self.window.FindName("LengthBox")
        self.different = self.window.FindName("DifferentCheck")
        self.mode_all = self.window.FindName("ModeAll")
        self.mode_choose = self.window.FindName("ModeChoose")
        self.pipe_combo = self.window.FindName("PipeTypeCombo")
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

        self.pipe_names = pipe_types
        self.has_pipe_types = fill_combo(self.pipe_combo, pipe_types, settings["pipe_type"])
        # per forma: (menu, nomi, ci sono tipi)
        self.duct = {}
        for key, combo_name in DUCT_COMBOS:
            combo = self.window.FindName(combo_name)
            names = duct_types.get(key, [])
            self.duct[key] = (combo, names, fill_combo(combo, names, settings[key + "_type"]))
        self.has_duct_types = any(has for _, _, has in self.duct.values())

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
        values = {
            "length_mm": length,
            "different": bool(self.different.IsChecked),
            "choose": bool(self.mode_choose.IsChecked),
            "pipe_type": self.pipe_names[self.pipe_combo.SelectedIndex] if self.has_pipe_types else "",
        }
        for key, (combo, names, has) in self.duct.items():
            values[key + "_type"] = names[combo.SelectedIndex] if has else ""
        return True, "", values

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Spring")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
