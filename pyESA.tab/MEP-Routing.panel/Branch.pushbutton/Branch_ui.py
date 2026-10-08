# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader

DIRECTIONS = ("up", "down", "left", "right")
JUNCTIONS = (
    ("standard", "Tap on ducts, tee on pipes (recommended)"),
    ("routing", "As set in the Routing Preferences of each type"),
    ("tee", "Always tee"),
    ("tap", "Always tap"),
)
MAX_ANGLE = 80.0


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


class BranchWindow(object):
    """Finestra modale: lato dello stacco, quota (livello + offset), lunghezza, angolo, prolungamento oltre la T."""

    def __init__(self, xaml_path, run_count, level_names, settings, view_note=None, warnings=None):
        # warnings(nome livello, offset mm) -> testo con gli elementi che verrebbero scollegati
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.radios = {"up": self.window.FindName("DirUp"), "down": self.window.FindName("DirDown"),
                       "left": self.window.FindName("DirLeft"), "right": self.window.FindName("DirRight")}
        self.level_combo = self.window.FindName("LevelCombo")
        self.offset_box = self.window.FindName("OffsetBox")
        self.length_box = self.window.FindName("LengthBox")
        self.angle_box = self.window.FindName("AngleBox")
        self.stub_box = self.window.FindName("StubBox")
        self.junction_combo = self.window.FindName("JunctionCombo")
        keys = [key for key, _ in JUNCTIONS]
        for _, label in JUNCTIONS:
            self.junction_combo.Items.Add(label)
        self.junction_combo.SelectedIndex = keys.index(settings["junction"]) if settings["junction"] in keys else 0
        self.window.FindName("SummaryText").Text = (
            "{} vertical runs selected. Each one gets a branch on the chosen side, at the chosen elevation.{}".format(
                run_count, "\n\nNote: " + view_note if view_note else ""))
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

        direction = settings["direction"] if settings["direction"] in DIRECTIONS else "right"
        self.radios[direction].IsChecked = True
        for name in level_names:
            self.level_combo.Items.Add(name)
        if settings["level"] in level_names:
            self.level_combo.SelectedIndex = level_names.index(settings["level"])
        elif level_names:
            self.level_combo.SelectedIndex = 0
        self.offset_box.Text = "{:g}".format(settings["offset_mm"])
        self.length_box.Text = "{:g}".format(settings["length_mm"])
        self.angle_box.Text = "{:g}".format(settings["angle_deg"])
        self.stub_box.Text = "{:g}".format(settings["stub_mm"])

        self.warnings = warnings
        self.connection_text = self.window.FindName("ConnectionText")
        self.level_combo.SelectionChanged += self._on_elevation
        self.offset_box.TextChanged += self._on_elevation
        self._on_elevation(None, None)

    def _on_elevation(self, sender, args):
        """Aggiorna l'elenco degli elementi che verrebbero scollegati con la quota inserita."""
        if self.warnings is None or self.level_combo.SelectedIndex < 0:
            return
        ok, offset = parse_number(self.offset_box.Text)
        if not ok:
            self.connection_text.Text = "Enter the offset as a number to see which connections are affected."
            return
        try:
            self.connection_text.Text = self.warnings(str(self.level_combo.SelectedItem), offset)
        except Exception as ex:
            self.connection_text.Text = "Connections could not be checked: {}".format(ex)

    def _direction(self):
        for name in DIRECTIONS:
            if self.radios[name].IsChecked:
                return name
        return "right"

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        if self.level_combo.SelectedIndex < 0:
            return False, "Choose a level.", None
        ok, offset = parse_number(self.offset_box.Text)
        if not ok:
            return False, "The offset must be a number (mm).", None
        ok, length = parse_number(self.length_box.Text)
        if not ok or length < 50:
            return False, "The length of the branch must be a number of at least 50 mm.", None
        ok, angle = parse_number(self.angle_box.Text)
        if not ok or abs(angle) > MAX_ANGLE:
            return False, "The angle must be a number between -{0:g} and {0:g} degrees.".format(MAX_ANGLE), None
        ok, stub = parse_number(self.stub_box.Text)
        if not ok or stub < 50:
            return False, "The riser past the tee must be a number of at least 50 mm.", None
        return True, "", {"direction": self._direction(), "level": str(self.level_combo.SelectedItem),
                          "offset_mm": offset, "length_mm": length, "angle_deg": angle, "stub_mm": stub,
                          "junction": JUNCTIONS[max(self.junction_combo.SelectedIndex, 0)][0]}

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Branch")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
