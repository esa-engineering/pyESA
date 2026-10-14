# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows.Markup import XamlReader


class SettingsWindow(object):
    """Finestra modale delle impostazioni condivise da Match Top / Middle / Bottom: tre opzioni."""

    def __init__(self, xaml_path, settings):
        # settings: dict con override_level e remove_slope
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.level_check = self.window.FindName("OverrideLevelCheck")
        self.slope_check = self.window.FindName("RemoveSlopeCheck")
        self.single_check = self.window.FindName("SingleElementCheck")
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

        self.level_check.IsChecked = bool(settings["override_level"])
        self.slope_check.IsChecked = bool(settings["remove_slope"])
        self.single_check.IsChecked = bool(settings["single_element"])

    def _on_ok(self, sender, args):
        self.result = {
            "override_level": bool(self.level_check.IsChecked),
            "remove_slope": bool(self.slope_check.IsChecked),
            "single_element": bool(self.single_check.IsChecked),
        }
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
