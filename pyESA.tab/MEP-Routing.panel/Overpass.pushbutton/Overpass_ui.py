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


class OverpassWindow(object):
    """Finestra modale: offset (mm, con segno) e angolo delle salite."""

    def __init__(self, xaml_path, default_offset_mm, default_angle_deg):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.offset_box = self.window.FindName("OffsetBox")
        self.angle_combo = self.window.FindName("AngleCombo")
        self.angle_box = self.window.FindName("AngleBox")
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

        self.offset_box.Text = "{:g}".format(default_offset_mm)
        for angle in PRESET_ANGLES:
            self.angle_combo.Items.Add("{} degrees".format(angle))
        self.angle_combo.Items.Add(CUSTOM_LABEL)

        if default_angle_deg in PRESET_ANGLES:
            self.angle_combo.SelectedIndex = PRESET_ANGLES.index(default_angle_deg)
        else:
            self.angle_combo.SelectedIndex = len(PRESET_ANGLES)
        self.angle_box.Text = "{:g}".format(default_angle_deg)
        self.angle_combo.SelectionChanged += self._on_angle_mode
        self._on_angle_mode(None, None)

    def _is_custom(self):
        return self.angle_combo.SelectedIndex == len(PRESET_ANGLES)

    def _on_angle_mode(self, sender, args):
        self.angle_box.IsEnabled = self._is_custom()

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        ok, offset = parse_number(self.offset_box.Text)
        if not ok:
            return False, "The offset must be a number (mm).", None
        if abs(offset) < 1e-6:
            return False, "The offset must not be zero.", None
        if self._is_custom():
            ok, angle = parse_number(self.angle_box.Text)
            if not ok or angle <= 0 or angle > 90:
                return False, "The custom angle must be a number greater than 0 and up to 90 degrees.", None
        else:
            angle = float(PRESET_ANGLES[self.angle_combo.SelectedIndex])
        return True, "", {"offset_mm": offset, "angle_deg": angle}

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Overpass")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
