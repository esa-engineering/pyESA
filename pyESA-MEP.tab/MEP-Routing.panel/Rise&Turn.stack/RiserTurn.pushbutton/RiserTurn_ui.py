# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader

DIRECTIONS = ("up", "down", "left", "right", "custom")
POSITIONS = ("top", "bottom", "elevation")


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


class RiserTurnWindow(object):
    """Finestra modale: posizione della curva (estremo alto/basso o quota con estremo da spostare), lato del tratto
    orizzontale e sua lunghezza."""

    def __init__(self, xaml_path, run_count, level_names, settings, view_note=None):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        find = self.window.FindName
        self.radios = {"up": find("DirUp"), "down": find("DirDown"), "left": find("DirLeft"),
                       "right": find("DirRight"), "custom": find("DirCustom")}
        self.positions = {"top": find("PosTop"), "bottom": find("PosBottom"), "elevation": find("PosElevation")}
        self.move = {"top": find("MoveTop"), "bottom": find("MoveBottom")}
        self.elevation_panel = find("ElevationPanel")
        self.level_combo = find("LevelCombo")
        self.offset_box = find("OffsetBox")
        self.length_box = find("LengthBox")
        find("SummaryText").Text = (
            "{} vertical runs selected. Each one gets an elbow and a horizontal run.{}".format(
                run_count, "\n\nNote: " + view_note if view_note else ""))
        find("OkButton").Click += self._on_ok
        self.result = None

        direction = settings["direction"] if settings["direction"] in DIRECTIONS else "right"
        self.radios[direction].IsChecked = True
        position = settings["position"] if settings["position"] in POSITIONS else "top"
        self.positions[position].IsChecked = True
        move = settings["move_end"] if settings["move_end"] in self.move else "top"
        self.move[move].IsChecked = True
        for name in level_names:
            self.level_combo.Items.Add(name)
        if settings["level"] in level_names:
            self.level_combo.SelectedIndex = level_names.index(settings["level"])
        elif level_names:
            self.level_combo.SelectedIndex = 0
        self.offset_box.Text = "{:g}".format(settings["offset_mm"])
        self.length_box.Text = "{:g}".format(settings["length_mm"])
        for radio in self.positions.values():
            radio.Checked += self._on_position
        self._on_position(None, None)

    def _on_position(self, sender, args):
        """I campi della quota servono solo con "At an elevation"."""
        self.elevation_panel.IsEnabled = bool(self.positions["elevation"].IsChecked)

    def _checked(self, radios, default):
        for name, radio in radios.items():
            if radio.IsChecked:
                return name
        return default

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        position = self._checked(self.positions, "top")
        offset = 0.0
        level = str(self.level_combo.SelectedItem) if self.level_combo.SelectedIndex >= 0 else ""
        if position == "elevation":
            if not level:
                return False, "Choose a level.", None
            ok, offset = parse_number(self.offset_box.Text)
            if not ok:
                return False, "The offset must be a number (mm).", None
        else:
            ok, value = parse_number(self.offset_box.Text)
            offset = value if ok else 0.0  # non usato: si conserva l'ultimo valore valido
        ok, length = parse_number(self.length_box.Text)
        if not ok or length < 50:
            return False, "The length of the horizontal run must be a number of at least 50 mm.", None
        return True, "", {"direction": self._checked(self.radios, "right"), "position": position, "level": level,
                          "offset_mm": offset, "move_end": self._checked(self.move, "top"), "length_mm": length}

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Riser Turn")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
