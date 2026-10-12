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


class AngleWindow(object):
    """Unica finestra del comando, a ogni avvio: angolo (pulsanti con gli angoli piu' usati o campo libero),
    lunghezza, pendenza e aggiramento delle interferenze. Sostituisce forms.CommandSwitchWindow, che si chiude appena
    perde il focus (subito dopo un clic nel modello)."""

    def __init__(self, xaml_path, settings, angles, min_deg, max_deg):
        # settings: dict con angle_deg, length_mm, keep_slope, avoid_clash; show() ritorna lo stesso dict aggiornato
        from System.Windows.Controls import Button
        from System.Windows import Thickness
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.settings = dict(settings)
        self.min_deg, self.max_deg = min_deg, max_deg
        self.window.FindName("SavedText").Text = "Click an angle, or type a custom one. Saved angle: {:g} degrees.".format(
            settings["angle_deg"])
        self.window.FindName("RangeText").Text = "degrees ({:g} to {:g})".format(min_deg, max_deg)
        self.custom_box = self.window.FindName("CustomBox")
        self.custom_box.Text = "{:g}".format(settings["angle_deg"])
        self.length_box = self.window.FindName("LengthBox")
        self.length_box.Text = "{:g}".format(settings["length_mm"])
        self.keep_slope = self.window.FindName("KeepSlope")
        self.horizontal = self.window.FindName("Horizontal")
        self.keep_slope.IsChecked = bool(settings["keep_slope"])
        self.horizontal.IsChecked = not settings["keep_slope"]
        self.clash_check = self.window.FindName("AvoidClashCheck")
        self.clash_check.IsChecked = bool(settings.get("avoid_clash", False))
        panel = self.window.FindName("AnglePanel")
        for angle in angles:
            button = Button()
            button.Content = u"{:g}°".format(angle)
            button.Width, button.Height = 60, 30
            button.Margin = Thickness(0, 0, 6, 6)
            button.Click += self._picker(angle)
            panel.Children.Add(button)
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

    def _finish(self, angle):
        """Valida le altre opzioni tutte insieme; se va bene chiude la finestra con il risultato."""
        ok, length = parse_number(self.length_box.Text)
        if not ok or length <= 0:
            MessageBox.Show("The length must be a number greater than zero (mm).", "Turn")
            return
        self.settings.update({"angle_deg": angle, "length_mm": length,
                              "keep_slope": bool(self.keep_slope.IsChecked),
                              "avoid_clash": bool(self.clash_check.IsChecked)})
        self.result = self.settings
        self.window.DialogResult = True

    def _picker(self, angle):
        def handler(sender, args):
            self._finish(angle)
        return handler

    def _on_ok(self, sender, args):
        ok, value = parse_number(self.custom_box.Text)
        if not ok or value < self.min_deg or value > self.max_deg:
            MessageBox.Show("The angle must be a number from {:g} to {:g} degrees.".format(
                self.min_deg, self.max_deg), "Turn")
            return
        self._finish(value)

    def show(self):
        self.window.ShowDialog()
        return self.result
