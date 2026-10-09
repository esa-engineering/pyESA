# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader

PRESET_ANGLES = (10, 15, 30, 45, 60, 90)
CUSTOM_LABEL = "Custom"


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


class TurnWindow(object):
    """Finestra modale (SHIFT+CLICK): lunghezza, angolo di curva e pendenza."""

    def __init__(self, xaml_path, settings):
        # settings: dict con length_mm, angle_deg, keep_slope
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.length_box = self.window.FindName("LengthBox")
        self.angle_combo = self.window.FindName("AngleCombo")
        self.angle_box = self.window.FindName("AngleBox")
        self.keep_slope = self.window.FindName("KeepSlope")
        self.horizontal = self.window.FindName("Horizontal")
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

        self.length_box.Text = "{:g}".format(settings["length_mm"])
        for angle in PRESET_ANGLES:
            self.angle_combo.Items.Add("{} degrees".format(angle))
        self.angle_combo.Items.Add(CUSTOM_LABEL)
        angle = settings["angle_deg"]
        if angle in PRESET_ANGLES:
            self.angle_combo.SelectedIndex = PRESET_ANGLES.index(angle)
        else:
            self.angle_combo.SelectedIndex = len(PRESET_ANGLES)
        self.angle_box.Text = "{:g}".format(angle)
        self.angle_combo.SelectionChanged += self._on_angle_mode
        self._on_angle_mode(None, None)

        self.keep_slope.IsChecked = bool(settings["keep_slope"])
        self.horizontal.IsChecked = not settings["keep_slope"]

    def _is_custom(self):
        return self.angle_combo.SelectedIndex == len(PRESET_ANGLES)

    def _on_angle_mode(self, sender, args):
        self.angle_box.IsEnabled = self._is_custom()

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        ok, length = parse_number(self.length_box.Text)
        if not ok or length <= 0:
            return False, "The length must be a number greater than zero (mm).", None
        if self._is_custom():
            ok, angle = parse_number(self.angle_box.Text)
            if not ok or angle <= 0 or angle > 90:
                return False, "The custom angle must be a number greater than 0 and up to 90 degrees.", None
        else:
            angle = float(PRESET_ANGLES[self.angle_combo.SelectedIndex])
        return True, "", {"length_mm": length, "angle_deg": angle,
                          "keep_slope": bool(self.keep_slope.IsChecked)}

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Turn Settings")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
