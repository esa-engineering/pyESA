# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox, Visibility, Thickness
from System.Windows.Controls import CheckBox
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


ACTIONS = ("over", "under", "none")
SCOPES = ("distribution", "all")
SCOPE_TEXT = {"distribution": "distribution elements (pipes, ducts, cable trays, conduits, fittings, accessories, "
                              "insulation)",
              "all": "all model elements"}


class ClashWindow(object):
    """Finestra modale, solo all'avvio: azione (sopra / sotto / solo report, distanza, unione degli scavalchi vicini,
    angolo delle salite), avviso con il numero di interferenze trovate e tipo di controllo. Dopo OK lo script risolve
    tutte le interferenze con queste impostazioni. Cambiando il tipo di controllo la finestra si chiude con
    {"rescan": True, "scope": ...}: lo script rifa' la ricerca e la riapre."""

    def __init__(self, xaml_path, clash_count, settings):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.clearance_box = self.window.FindName("ClearanceBox")
        self.merge_box = self.window.FindName("MergeBox")
        self.angle_combo = self.window.FindName("AngleCombo")
        self.angle_box = self.window.FindName("AngleBox")
        self.radios = {"over": self.window.FindName("ActionOver"),
                       "under": self.window.FindName("ActionUnder"),
                       "none": self.window.FindName("ActionNone")}
        self.scope_radios = {"distribution": self.window.FindName("ScopeDistribution"),
                             "all": self.window.FindName("ScopeAll")}
        self.scope = settings.get("scope", "distribution") if settings.get("scope") in SCOPES else "distribution"
        # avviso con il numero di interferenze: si risolvono tutte, niente elenco da spuntare
        header = self.window.FindName("ClashHeader")
        if clash_count:
            header.Text = "⚠ {} CLASHES FOUND".format(clash_count) if clash_count > 1 else "⚠ 1 CLASH FOUND"
            text = ("Between the selected runs and {}. On OK they are all solved with the settings above (Report "
                    "only: nothing is changed, they are listed in the report). Clashes found after each extension "
                    "are solved in the same way. Stop ends the tool.".format(SCOPE_TEXT[self.scope]))
        else:
            header.Text = "✔ NO CLASH FOUND"
            text = ("No clash with {}. On OK you click the point to extend the runs to; clashes found after each "
                    "extension are solved automatically with the settings above.".format(SCOPE_TEXT[self.scope]))
        self.window.FindName("ClashText").Text = text
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

        action = settings["action"] if settings["action"] in ACTIONS else "over"
        self.radios[action].IsChecked = True
        self.clearance_box.Text = "{:g}".format(settings["clearance_mm"])
        self.merge_box.Text = "{:g}".format(settings.get("merge_mm", 2000.0))
        # stato iniziale prima di collegare gli eventi: il cambio successivo fa ripetere la ricerca
        self.scope_radios[self.scope].IsChecked = True
        for radio in self.scope_radios.values():
            radio.Checked += self._on_scope
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

    def _is_custom(self):
        return self.angle_combo.SelectedIndex == len(PRESET_ANGLES)

    def _on_angle_mode(self, sender, args):
        self.angle_box.IsEnabled = self._is_custom()

    def _action(self):
        for name in ACTIONS:
            if self.radios[name].IsChecked:
                return name
        return "over"

    def _on_scope(self, sender, args):
        scope = next((name for name in SCOPES if self.scope_radios[name].IsChecked), self.scope)
        if scope == self.scope:
            return
        self.result = {"rescan": True, "scope": scope}
        self.window.DialogResult = True

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        ok, clearance = parse_number(self.clearance_box.Text)
        if not ok or clearance < 0:
            return False, "The clearance must be a number (mm), zero or more.", None
        ok, merge = parse_number(self.merge_box.Text)
        if not ok or merge < 0:
            return False, "The merge distance must be a number (mm), zero or more.", None
        if self._is_custom():
            ok, angle = parse_number(self.angle_box.Text)
            if not ok or angle <= 0 or angle > 90:
                return False, "The custom angle must be a number greater than 0 and up to 90 degrees.", None
        else:
            angle = float(PRESET_ANGLES[self.angle_combo.SelectedIndex])
        return True, "", {"action": self._action(), "clearance_mm": clearance, "angle_deg": angle,
                          "merge_mm": merge, "scope": self.scope}

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
