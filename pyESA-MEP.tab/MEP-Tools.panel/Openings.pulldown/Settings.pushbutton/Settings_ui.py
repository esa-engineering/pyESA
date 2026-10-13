# -*- coding: utf-8 -*-
import math

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader

FAMILY_CONTROLS = (("fam_wall_rect", "FamWallRect"), ("fam_wall_round", "FamWallRound"),
                   ("fam_floor_rect", "FamFloorRect"), ("fam_floor_round", "FamFloorRound"))
NUMBER_CONTROLS = (("clearance_mm", "ClearanceBox", "Clearance per side"), ("rounding_mm", "RoundingBox", "Round up to"),
                   ("merge_mm", "MergeBox", "Merge within"), ("extra_depth_mm", "ExtraDepthBox",
                                                              "Extra depth per side"),
                   ("layer_gap_mm", "LayerGapBox", "Join layers across"),
                   ("min_angle_deg", "MinAngleBox", "Minimum angle"))
MAX_ANGLE = 89.0


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


class SettingsWindow(object):
    """Impostazioni di Openings: famiglie (scelte tra i Generic Model caricati, o scritte), misure, opzioni.
    show() ritorna il dict aggiornato, o None se annullata."""

    def __init__(self, xaml_path, settings, defaults, family_names, params_text):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.settings = dict(settings)
        self.defaults = defaults
        find = self.window.FindName
        self.combos = {}
        for key, control in FAMILY_CONTROLS:
            combo = find(control)
            for name in family_names:
                combo.Items.Add(name)
            combo.Text = settings[key]
            self.combos[key] = combo
        self.boxes = {}
        for key, control, _ in NUMBER_CONTROLS:
            box = find(control)
            box.Text = "{:g}".format(settings[key])
            self.boxes[key] = box
        self.slope_text = find("MinAngleSlope")
        self.boxes["min_angle_deg"].TextChanged += self._on_angle
        self._on_angle(None, None)
        self.separate = find("SeparateCheck")
        self.separate.IsChecked = bool(settings["separate"])
        self.misaligned = find("MisalignedCheck")
        self.misaligned.IsChecked = bool(settings["misaligned"])
        find("FamilyHint").Text = ("Generic Model families loaded in this model are listed; a name can also be "
                                   "typed. Not hosted, origin at the centre. The Families button of this menu "
                                   "creates the default ones.")
        find("ParamsInfo").Text = params_text
        find("DefaultsButton").Click += self._on_defaults
        find("OkButton").Click += self._on_ok
        self.result = None

    def _on_angle(self, sender, args):
        """Accanto all'angolo, la pendenza equivalente in percentuale."""
        ok, angle = parse_number(self.boxes["min_angle_deg"].Text)
        if ok and 0 <= angle <= MAX_ANGLE:
            self.slope_text.Text = "degrees (= {:.0f}% slope)".format(math.tan(math.radians(angle)) * 100.0)
        else:
            self.slope_text.Text = "degrees (0 to {:g})".format(MAX_ANGLE)

    def _on_defaults(self, sender, args):
        for key, combo in self.combos.items():
            combo.Text = self.defaults[key]

    def _validate(self):
        """(ok, messaggio, valori): tutti i campi insieme."""
        values = {}
        for key, combo in self.combos.items():
            name = (combo.Text or "").strip()
            if not name:
                return False, "Choose a family for every case.", None
            values[key] = name
        for key, _, label in NUMBER_CONTROLS:
            ok, number = parse_number(self.boxes[key].Text)
            if not ok or number < 0:
                return False, "{}: enter a number of 0 or more.".format(label), None
            if key == "min_angle_deg" and number > MAX_ANGLE:
                return False, "Minimum angle: enter a number from 0 to {:g} degrees.".format(MAX_ANGLE), None
            values[key] = number
        values["separate"] = bool(self.separate.IsChecked)
        values["misaligned"] = bool(self.misaligned.IsChecked)
        return True, "", values

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Openings Settings")
            return
        self.settings.update(values)
        self.result = self.settings
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
