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


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


class BranchWindow(object):
    """Finestra modale di una categoria (kind: "pipe", "duct", "tray"): lato dello stacco, quota (livello + offset),
    lunghezza e, per tubi e canali, giunzione e prolungamento oltre la T. Tubi: anche direzione libera con un clic.
    Passerelle: nessuna giunzione (sempre T / curva), elenco dei lati disponibili."""

    def __init__(self, xaml_path, kind, run_count, level_names, settings, view_note=None, warnings=None,
                 sides_info=None):
        # warnings(nome livello, offset mm) -> testo con gli elementi che verrebbero scollegati
        self.kind = kind
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.radios = {"up": self.window.FindName("DirUp"), "down": self.window.FindName("DirDown"),
                       "left": self.window.FindName("DirLeft"), "right": self.window.FindName("DirRight")}
        if kind == "pipe":
            self.radios["custom"] = self.window.FindName("DirCustom")
        self.level_combo = self.window.FindName("LevelCombo")
        self.offset_box = self.window.FindName("OffsetBox")
        self.length_box = self.window.FindName("LengthBox")
        self.stub_box = self.window.FindName("StubBox")
        self.junction_combo = self.window.FindName("JunctionCombo")
        self.defaults = settings
        if self.junction_combo is not None:
            keys = [key for key, _ in JUNCTIONS]
            for _, label in JUNCTIONS:
                self.junction_combo.Items.Add(label)
            self.junction_combo.SelectedIndex = keys.index(settings["junction"]) if settings["junction"] in keys else 0
        if sides_info is not None:
            self.window.FindName("SidesText").Text = sides_info
        names = {"pipe": "pipes", "duct": "ducts", "tray": "cable trays"}
        self.window.FindName("SummaryText").Text = (
            "{} vertical {} selected. Each one gets a branch on the chosen side, at the chosen elevation.{}".format(
                run_count, names[kind], "\n\nNote: " + view_note if view_note else ""))
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

        direction = settings["direction"] if settings["direction"] in self.radios else "right"
        self.radios[direction].IsChecked = True
        for name in level_names:
            self.level_combo.Items.Add(name)
        if settings["level"] in level_names:
            self.level_combo.SelectedIndex = level_names.index(settings["level"])
        elif level_names:
            self.level_combo.SelectedIndex = 0
        self.offset_box.Text = "{:g}".format(settings["offset_mm"])
        self.length_box.Text = "{:g}".format(settings["length_mm"])
        if self.stub_box is not None:
            self.stub_box.Text = "{:g}".format(settings["stub_mm"])
        self.reduce_box = self.window.FindName("ReduceBox")  # solo nella finestra dei tubi
        if self.reduce_box is not None:
            self.reduce_box.IsChecked = bool(settings.get("reduce", False))

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
        for name, radio in self.radios.items():
            if radio.IsChecked:
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
        stub = self.defaults["stub_mm"]
        if self.stub_box is not None:
            ok, stub = parse_number(self.stub_box.Text)
            if not ok or stub < 50:
                return False, "The riser past the tee must be a number of at least 50 mm.", None
        junction = self.defaults["junction"]
        if self.junction_combo is not None:
            junction = JUNCTIONS[max(self.junction_combo.SelectedIndex, 0)][0]
        return True, "", {"direction": self._direction(), "level": str(self.level_combo.SelectedItem),
                          "offset_mm": offset, "length_mm": length, "stub_mm": stub, "junction": junction,
                          "reduce": bool(self.reduce_box.IsChecked) if self.reduce_box is not None
                          else bool(self.defaults.get("reduce", False))}

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


