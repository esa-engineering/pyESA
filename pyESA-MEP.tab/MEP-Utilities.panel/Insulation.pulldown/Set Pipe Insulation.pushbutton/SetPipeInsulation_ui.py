# -*- coding: utf-8 -*-
import os

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader
from Microsoft.Win32 import OpenFileDialog

# Spessori proposti per lo spessore fisso (mm)
THICKNESS_OPTIONS = [6, 9, 13, 19, 25, 32, 40, 50, 55, 60]
DEFAULT_THICKNESS = 13


def fill_combo(combo, names, selected_name):
    """Riempie il menu e seleziona l'ultima scelta, se ancora presente."""
    for name in names:
        combo.Items.Add(name)
    if names:
        combo.SelectedIndex = names.index(selected_name) if selected_name in names else 0


class SetPipeInsulationWindow(object):
    """Finestra modale: sistema, ambito, tipi di isolamento, tabella spessori o spessore fisso."""

    def __init__(self, xaml_path, system_names, insulation_type_names, selection_count, settings):
        # settings: dict con le ultime impostazioni (vedi DEFAULTS nello script)
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.system_combo = self.window.FindName("SystemCombo")
        self.mode_all = self.window.FindName("ModeAll")
        self.mode_selected = self.window.FindName("ModeSelected")
        self.visible_combo = self.window.FindName("VisibleCombo")
        self.hidden_combo = self.window.FindName("HiddenCombo")
        self.table_box = self.window.FindName("TableBox")
        self.manual_check = self.window.FindName("ManualCheck")
        self.thickness_combo = self.window.FindName("ThicknessCombo")
        self.window.FindName("BrowseButton").Click += self._on_browse
        self.window.FindName("OkButton").Click += self._on_ok
        self.manual_check.Checked += self._on_manual_changed
        self.manual_check.Unchecked += self._on_manual_changed
        self.result = None

        self.system_names = system_names
        self.type_names = insulation_type_names
        fill_combo(self.system_combo, system_names, settings["system"])
        fill_combo(self.visible_combo, insulation_type_names, settings["visible_type"])
        fill_combo(self.hidden_combo, insulation_type_names, settings["hidden_type"])
        fill_combo(self.thickness_combo, THICKNESS_OPTIONS,
                   settings["thickness"] if settings["thickness"] in THICKNESS_OPTIONS else DEFAULT_THICKNESS)

        self.table_box.Text = settings["table_path"]
        self.manual_check.IsChecked = bool(settings["manual"])
        self._on_manual_changed(None, None)

        if selection_count:
            self.mode_selected.Content = "Only the selected pipes and fittings of this system ({} selected)".format(
                selection_count)
            self.mode_selected.IsChecked = bool(settings["selected_only"])
            self.mode_all.IsChecked = not settings["selected_only"]
        else:
            self.mode_selected.Content = "Only the selected pipes and fittings of this system (nothing selected)"
            self.mode_selected.IsEnabled = False
            self.mode_all.IsChecked = True

    def _on_manual_changed(self, sender, args):
        self.thickness_combo.IsEnabled = bool(self.manual_check.IsChecked)

    def _on_browse(self, sender, args):
        dialog = OpenFileDialog()
        dialog.Title = "Select the insulation thickness table"
        dialog.Filter = "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        current = (self.table_box.Text or "").strip()
        if current and os.path.isdir(os.path.dirname(current)):
            dialog.InitialDirectory = os.path.dirname(current)
        else:
            dialog.InitialDirectory = os.path.expanduser("~\\Documents")
        if dialog.ShowDialog(self.window):
            self.table_box.Text = dialog.FileName

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        if self.system_combo.SelectedIndex < 0:
            return False, "Select a piping system.", None
        if self.visible_combo.SelectedIndex < 0 or self.hidden_combo.SelectedIndex < 0:
            return False, "Select both insulation types.", None
        manual = bool(self.manual_check.IsChecked)
        table_path = (self.table_box.Text or "").strip()
        # la tabella serve solo se lo spessore non e' fisso
        if not manual and not os.path.isfile(table_path):
            return False, "Select an existing thickness table file, or use a fixed thickness.", None
        return True, "", {
            "system": self.system_names[self.system_combo.SelectedIndex],
            "selected_only": bool(self.mode_selected.IsChecked),
            "visible_type": self.type_names[self.visible_combo.SelectedIndex],
            "hidden_type": self.type_names[self.hidden_combo.SelectedIndex],
            "table_path": table_path,
            "manual": manual,
            "thickness": THICKNESS_OPTIONS[self.thickness_combo.SelectedIndex],
        }

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Set Pipe Insulation")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
