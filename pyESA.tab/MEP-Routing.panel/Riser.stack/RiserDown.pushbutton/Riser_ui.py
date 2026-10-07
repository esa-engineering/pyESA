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


TEXTS = {
    "up": {"title": "Riser Up", "header": u"⬆ RISER UP",
           "length": "Length above the selected end", "level": "Up to a level (plus offset)",
           "hint": "goes straight up"},
    "down": {"title": "Riser Down", "header": u"⬇ RISER DOWN",
             "length": "Length below the selected end", "level": "Down to a level (plus offset)",
             "hint": "goes straight down"},
}


class RiserWindow(object):
    """Finestra modale: altezza del tratto verticale (lunghezza oppure livello + offset)."""

    def __init__(self, xaml_path, levels, defaults, direction):
        # levels: lista di tuple (nome, quota_mm); defaults: dict con le ultime impostazioni
        texts = TEXTS[direction]
        self.title = texts["title"]
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.window.Title = texts["title"]
        self.window.FindName("TitleText").Text = texts["header"]
        self.window.FindName("HintBlock").Text = (
            "After OK, click each pipe, duct, cable tray or conduit near the free end you want to extend. "
            "The riser starts from the nearest free end and {}, with an elbow where needed. "
            "Press ESC to finish and get the report. These settings are kept for the next normal "
            "clicks.".format(texts["hint"]))
        self.rb_length = self.window.FindName("ModeLength")
        self.rb_level = self.window.FindName("ModeLevel")
        self.rb_length.Content = texts["length"]
        self.rb_level.Content = texts["level"]
        self.length_box = self.window.FindName("LengthBox")
        self.offset_box = self.window.FindName("OffsetBox")
        self.level_combo = self.window.FindName("LevelCombo")
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

        for name, elev_mm in levels:
            self.level_combo.Items.Add("{} ({:+.0f} mm)".format(name, elev_mm))
        if levels:
            index = 0
            for i, (name, _) in enumerate(levels):
                if name == defaults.get("level_name"):
                    index = i
            self.level_combo.SelectedIndex = index
        else:
            self.rb_level.IsEnabled = False

        self.length_box.Text = "{:g}".format(defaults.get("length_mm", 1000.0))
        self.offset_box.Text = "{:g}".format(defaults.get("offset_mm", 0.0))

        self.rb_length.Checked += self._on_mode
        self.rb_level.Checked += self._on_mode
        use_level = defaults.get("mode") == "level" and bool(levels)
        self.rb_level.IsChecked = use_level
        self.rb_length.IsChecked = not use_level
        self._on_mode(None, None)

    def _on_mode(self, sender, args):
        by_level = bool(self.rb_level.IsChecked)
        self.length_box.IsEnabled = not by_level
        self.level_combo.IsEnabled = by_level
        self.offset_box.IsEnabled = by_level

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        if self.rb_level.IsChecked:
            ok, offset = parse_number(self.offset_box.Text)
            if not ok:
                return False, "The offset must be a number (mm).", None
            if self.level_combo.SelectedIndex < 0:
                return False, "Select a level.", None
            return True, "", {"mode": "level", "offset_mm": offset,
                              "level_index": self.level_combo.SelectedIndex}
        ok, length = parse_number(self.length_box.Text)
        if not ok or length <= 0:
            return False, "The length must be a number greater than zero (mm).", None
        return True, "", {"mode": "length", "length_mm": length}

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, self.title)
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
