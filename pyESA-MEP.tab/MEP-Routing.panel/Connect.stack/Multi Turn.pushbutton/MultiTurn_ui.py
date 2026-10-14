# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox, Visibility
from System.Windows.Markup import XamlReader

DIRECTIONS = ("up", "down", "left", "right")
DIRECTION_BUTTONS = {"up": "DirUp", "down": "DirDown", "left": "DirLeft", "right": "DirRight"}


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


def describe(settings):
    """Testo breve della scelta salvata (finestra, prompt e report)."""
    if settings.get("mode") == "align":
        return "parallel to a clicked element"
    if settings.get("mode") == "direction":
        return "towards the {} of the view".format(
            {"up": "top", "down": "bottom"}.get(settings.get("direction"), settings.get("direction")))
    return "{:g} degrees".format(settings["angle_deg"])


class TurnWindow(object):
    """Finestra delle impostazioni di Turn e Multi Turn: angolo (pulsanti con gli angoli piu' usati o campo libero),
    direzione nella vista (frecce), lunghezza, pendenza (solo Turn) e allineamento a un elemento cliccato.
    Sostituisce forms.CommandSwitchWindow, che si chiude appena perde il focus (subito dopo un clic nel modello).

    settings: dict con mode ("angle" / "direction" / "align"), angle_deg, direction, length_mm, keep_slope;
    show() ritorna lo stesso dict aggiornato, o None se annullata.
    texts: dict con caption (obbligatorio) e, se diversi dal XAML, title, length_unit, length_header, align_hint,
    info."""

    def __init__(self, xaml_path, settings, texts, angles, min_deg, max_deg, show_slope=True):
        from System.Windows.Controls import Button
        from System.Windows import Thickness
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.settings = dict(settings)
        self.caption = texts["caption"]
        self.min_deg, self.max_deg = min_deg, max_deg
        self.window.Title = texts["caption"]
        # testi propri del comando; quelli non passati restano come nel XAML
        for name, key in (("TitleText", "title"), ("LengthUnit", "length_unit"), ("LengthHeader", "length_header"),
                          ("AlignHint", "align_hint"), ("InfoText", "info")):
            if key in texts:
                self.window.FindName(name).Text = texts[key]
        self.window.FindName("SavedText").Text = "Click an angle or an arrow, or type a custom angle. Last used: {}.".format(
            describe(settings))
        self.window.FindName("RangeText").Text = "degrees ({:g} to {:g})".format(min_deg, max_deg)
        self.custom_box = self.window.FindName("CustomBox")
        self.custom_box.Text = "{:g}".format(settings["angle_deg"])
        self.length_box = self.window.FindName("LengthBox")
        self.length_box.Text = "{:g}".format(settings["length_mm"])
        self.keep_slope = self.window.FindName("KeepSlope")
        self.horizontal = self.window.FindName("Horizontal")
        self.keep_slope.IsChecked = bool(settings.get("keep_slope", True))
        self.horizontal.IsChecked = not settings.get("keep_slope", True)
        if not show_slope:
            self.window.FindName("SlopePanel").Visibility = Visibility.Collapsed
        self.ok_button = self.window.FindName("OkButton")
        self.ok_button.Click += self._on_ok
        self.angle_area = self.window.FindName("AngleArea")
        panel = self.window.FindName("AnglePanel")
        for angle in angles:
            button = Button()
            button.Content = u"{:g}°".format(angle)
            button.Width, button.Height = 52, 30
            button.Margin = Thickness(0, 0, 6, 6)
            button.Click += self._angle_picker(angle)
            panel.Children.Add(button)
        for direction in DIRECTIONS:
            self.window.FindName(DIRECTION_BUTTONS[direction]).Click += self._direction_picker(direction)
        self.align_check = self.window.FindName("AlignCheck")
        self.align_check.IsChecked = settings.get("mode") == "align"
        self.align_check.Checked += self._on_align
        self.align_check.Unchecked += self._on_align
        self._on_align(None, None)
        self.result = None

    def _on_align(self, sender, args):
        """Con l'allineamento l'angolo e la direzione vengono dall'elemento cliccato: angoli e frecce disattivati."""
        align = bool(self.align_check.IsChecked)
        self.angle_area.IsEnabled = not align
        self.ok_button.Content = "Start" if align else "Use custom"

    def _finish(self, mode, angle=None, direction=None):
        """Valida le altre opzioni tutte insieme; se va bene chiude la finestra con il risultato."""
        ok, length = parse_number(self.length_box.Text)
        if not ok or length < 50:
            MessageBox.Show("The length must be a number of at least 50 mm.", self.caption)
            return
        self.settings.update({"mode": mode, "length_mm": length, "keep_slope": bool(self.keep_slope.IsChecked)})
        if angle is not None:
            self.settings["angle_deg"] = angle
        if direction is not None:
            self.settings["direction"] = direction
        self.result = self.settings
        self.window.DialogResult = True

    def _angle_picker(self, angle):
        def handler(sender, args):
            self._finish("angle", angle=angle)
        return handler

    def _direction_picker(self, direction):
        def handler(sender, args):
            self._finish("direction", direction=direction)
        return handler

    def _on_ok(self, sender, args):
        if self.align_check.IsChecked:
            self._finish("align")
            return
        ok, value = parse_number(self.custom_box.Text)
        if not ok or value < self.min_deg or value > self.max_deg:
            MessageBox.Show("The angle must be a number from {:g} to {:g} degrees.".format(
                self.min_deg, self.max_deg), self.caption)
            return
        self._finish("angle", angle=value)

    def show(self):
        self.window.ShowDialog()
        return self.result
