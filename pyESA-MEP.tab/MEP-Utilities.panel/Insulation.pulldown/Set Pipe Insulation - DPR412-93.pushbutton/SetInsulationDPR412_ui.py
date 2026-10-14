# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader

# Spessori proposti per lo spessore fisso (mm)
THICKNESS_OPTIONS = [6, 9, 13, 19, 25, 32, 40, 50, 55, 60]
DEFAULT_THICKNESS = 13


def fill_combo(combo, names, selected_name):
    """Riempie il menu e seleziona l'ultima scelta, se ancora presente."""
    for name in names:
        combo.Items.Add(name)
    if names:
        combo.SelectedIndex = names.index(selected_name) if selected_name in names else 0


class SetInsulationDPR412Window(object):
    """Finestra modale: tipi di isolamento a vista / non a vista, spessore fisso opzionale."""

    def __init__(self, xaml_path, insulation_type_names, pipe_count, fitting_count, settings):
        # settings: dict con le ultime impostazioni (visible_type, hidden_type, manual, thickness)
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.visible_combo = self.window.FindName("VisibleCombo")
        self.hidden_combo = self.window.FindName("HiddenCombo")
        self.manual_check = self.window.FindName("ManualCheck")
        self.thickness_combo = self.window.FindName("ThicknessCombo")
        self.window.FindName("OkButton").Click += self._on_ok
        self.manual_check.Checked += self._on_manual_changed
        self.manual_check.Unchecked += self._on_manual_changed
        self.result = None

        self.type_names = insulation_type_names
        fill_combo(self.visible_combo, insulation_type_names, settings["visible_type"])
        fill_combo(self.hidden_combo, insulation_type_names, settings["hidden_type"])
        fill_combo(self.thickness_combo, THICKNESS_OPTIONS,
                   settings["thickness"] if settings["thickness"] in THICKNESS_OPTIONS else DEFAULT_THICKNESS)
        self.manual_check.IsChecked = bool(settings["manual"])
        self._on_manual_changed(None, None)

        self.window.FindName("SelectionHint").Text = (
            "Works on the selection: {} pipes and {} fittings. Elements without insulation get a new one; "
            "existing insulation is updated (type and thickness).".format(pipe_count, fitting_count))

    def _on_manual_changed(self, sender, args):
        self.thickness_combo.IsEnabled = bool(self.manual_check.IsChecked)

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        if self.visible_combo.SelectedIndex < 0 or self.hidden_combo.SelectedIndex < 0:
            return False, "Select both insulation types.", None
        return True, "", {
            "visible_type": self.type_names[self.visible_combo.SelectedIndex],
            "hidden_type": self.type_names[self.hidden_combo.SelectedIndex],
            "manual": bool(self.manual_check.IsChecked),
            "thickness": THICKNESS_OPTIONS[self.thickness_combo.SelectedIndex],
        }

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Set Pipe Insulation DPR 412/93")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
