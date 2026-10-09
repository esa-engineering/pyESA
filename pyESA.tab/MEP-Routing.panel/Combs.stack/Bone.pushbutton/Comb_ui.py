# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader

MODES = ("below", "above", "reference")
SIDES = ("Up", "Down")


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


class CombWindow(object):
    """Finestra modale: quote della colonna a parete e distanze."""

    def __init__(self, xaml_path, pipe_labels, settings):
        # pipe_labels: etichette dei tubi nel verso della corsa (#1 = corsa piu' lunga)
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.radios = {"below": self.window.FindName("ModeBelow"),
                       "above": self.window.FindName("ModeAbove"),
                       "reference": self.window.FindName("ModeReference")}
        self.reference_combo = self.window.FindName("ReferenceCombo")
        self.side_combo = self.window.FindName("SideCombo")
        self.factor_box = self.window.FindName("FactorBox")
        self.standoff_box = self.window.FindName("StandoffBox")
        self.length_box = self.window.FindName("LengthBox")
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

        for label in pipe_labels:
            self.reference_combo.Items.Add(label)
        self.reference_combo.SelectedIndex = len(pipe_labels) - 1  # l'ultimo e' l'unico valido
        for side in SIDES:
            self.side_combo.Items.Add(side)
        self.side_combo.SelectedIndex = SIDES.index(settings["side"]) if settings["side"] in SIDES else 0
        self.factor_box.Text = "{:g}".format(settings["factor"])
        self.standoff_box.Text = "{:g}".format(settings["standoff_mm"])
        self.length_box.Text = "{:g}".format(settings["length_mm"])

        for radio in self.radios.values():
            radio.Checked += self._on_mode
        mode = settings["mode"] if settings["mode"] in MODES else "below"
        self.radios[mode].IsChecked = True
        self._on_mode(None, None)

    def _mode(self):
        for name in MODES:
            if self.radios[name].IsChecked:
                return name
        return "below"

    def _on_mode(self, sender, args):
        reference = self._mode() == "reference"
        self.reference_combo.IsEnabled = reference
        self.side_combo.IsEnabled = reference

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        ok, factor = parse_number(self.factor_box.Text)
        if not ok or factor <= 0:
            return False, "The spacing factor must be a number greater than zero.", None
        ok, standoff = parse_number(self.standoff_box.Text)
        if not ok or standoff < 0:
            return False, "The distance from the wall must be a number (mm), zero or more.", None
        ok, length = parse_number(self.length_box.Text)
        if not ok or length < 50:
            return False, "The length along the wall must be a number of at least 50 mm.", None
        return True, "", {"mode": self._mode(), "reference_index": self.reference_combo.SelectedIndex,
                          "side": SIDES[self.side_combo.SelectedIndex],
                          "factor": factor, "standoff_mm": standoff, "length_mm": length}

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Comb")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result


class BendWindow(object):
    """Finestra modale: angolo e lunghezza del tratto dopo la curva."""

    def __init__(self, xaml_path, side_label, settings):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.angle_box = self.window.FindName("AngleBox")
        self.leg_box = self.window.FindName("LegBox")
        if side_label is None:
            side_text = "Next you select the elements, click where the bend starts, then click on the side to turn to."
        else:
            side_text = "{} elements will bend to the {}, towards the side of your click.".format(
                side_label[0], side_label[1])
        self.window.FindName("SideText").Text = side_text
        self.window.FindName("OkButton").Click += self._on_ok
        self.angle_box.Text = "{:g}".format(settings["angle_deg"])
        self.leg_box.Text = "{:g}".format(settings["leg_mm"])
        self.ask_angle = self.window.FindName("AskAngle")
        self.ask_angle.IsChecked = bool(settings.get("ask_angle", False))
        self.result = None

    def _on_ok(self, sender, args):
        ok, angle = parse_number(self.angle_box.Text)
        if not ok or angle < 5 or angle > 90:
            MessageBox.Show("The angle must be a number between 5 and 90 degrees.", "Comb")
            return
        ok, leg = parse_number(self.leg_box.Text)
        if not ok or leg < 50:
            MessageBox.Show("The length after the bend must be a number of at least 50 mm.", "Comb")
            return
        self.result = {"angle_deg": angle, "leg_mm": leg, "ask_angle": bool(self.ask_angle.IsChecked)}
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result


class AngleWindow(object):
    """Finestra modale per scegliere l'angolo al volo: pulsanti con gli angoli piu' usati e un campo libero.
    Sostituisce forms.CommandSwitchWindow, che si chiude appena perde il focus (subito dopo un clic nel modello)."""

    def __init__(self, xaml_path, title, saved_deg, angles, min_deg, max_deg):
        from System.Windows.Controls import Button
        from System.Windows import Thickness
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.window.Title = title
        self.min_deg, self.max_deg = min_deg, max_deg
        self.window.FindName("TitleText").Text = title.upper()
        self.window.FindName("SavedText").Text = "Click an angle, or type a custom one. Saved angle: {:g} degrees.".format(
            saved_deg)
        self.window.FindName("RangeText").Text = "degrees ({:g} to {:g})".format(min_deg, max_deg)
        self.custom_box = self.window.FindName("CustomBox")
        self.custom_box.Text = "{:g}".format(saved_deg)
        panel = self.window.FindName("AnglePanel")
        for angle in angles:
            button = Button()
            button.Content = u"{:g}\u00b0".format(angle)
            button.Width, button.Height = 60, 30
            button.Margin = Thickness(0, 0, 6, 6)
            button.Click += self._picker(angle)
            panel.Children.Add(button)
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

    def _picker(self, angle):
        def handler(sender, args):
            self.result = angle
            self.window.DialogResult = True
        return handler

    def _on_ok(self, sender, args):
        ok, value = parse_number(self.custom_box.Text)
        if not ok or value < self.min_deg or value > self.max_deg:
            MessageBox.Show("The angle must be a number from {:g} to {:g} degrees.".format(
                self.min_deg, self.max_deg), self.window.Title)
            return
        self.result = value
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
