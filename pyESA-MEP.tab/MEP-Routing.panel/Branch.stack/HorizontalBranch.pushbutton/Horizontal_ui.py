# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader

PRESET_ANGLES = (90.0, 45.0, 30.0)


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


class HorizontalWindow(object):
    """Unica finestra di Horizontal Branch (SHIFT + CLICK): offset, angolo (90, 45, 30 o personalizzato) e lunghezza
    del tratto orizzontale dopo il cambio di quota. show() ritorna {length_mm, drop_mm, angle_deg} o None."""

    def __init__(self, xaml_path, settings):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        find = self.window.FindName
        self.length_box = find("LengthBox")
        self.offset_box = find("OffsetBox")
        self.custom_box = find("CustomBox")
        self.custom_radio = find("AngCustom")
        self.preset_radios = {90.0: find("Ang90"), 45.0: find("Ang45"), 30.0: find("Ang30")}
        self.length_box.Text = "{:g}".format(settings["length_mm"])
        self.offset_box.Text = "{:g}".format(settings["drop_mm"])
        angle = float(settings["angle_deg"])
        self.custom_box.Text = "{:g}".format(angle)
        if angle in self.preset_radios:
            self.preset_radios[angle].IsChecked = True
        else:
            self.custom_radio.IsChecked = True
        find("OkButton").Click += self._on_ok
        self.result = None

    def _angle(self):
        for angle, radio in self.preset_radios.items():
            if radio.IsChecked:
                return True, angle
        return parse_number(self.custom_box.Text)

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        ok, length = parse_number(self.length_box.Text)
        if not ok or length < 50:
            return False, "The horizontal length must be a number of at least 50 mm.", None
        ok, offset = parse_number(self.offset_box.Text)
        if not ok or (offset != 0 and abs(offset) < 50):
            return False, "The offset must be 0, or a number of at least 50 mm in absolute value (+ up, - down).", None
        ok, angle = self._angle()
        if not ok or angle < 5 or angle > 90:
            return False, "The angle must be a number from 5 to 90 degrees.", None
        return True, "", {"length_mm": length, "drop_mm": offset, "angle_deg": angle}

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Horizontal Branch")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
