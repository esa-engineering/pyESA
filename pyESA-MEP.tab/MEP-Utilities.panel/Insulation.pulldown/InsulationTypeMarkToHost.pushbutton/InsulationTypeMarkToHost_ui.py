# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader


class TypeMarkToHostWindow(object):
    """Finestra modale: categorie, ambito (selezione / tutto), parametro di destinazione."""

    def __init__(self, xaml_path, selected_pipes, selected_ducts, param_names, settings):
        # selected_pipes / selected_ducts: conteggi della selezione corrente
        # param_names: parametri di testo di istanza scrivibili; settings: ultime impostazioni
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.cat_pipe = self.window.FindName("CatPipe")
        self.cat_duct = self.window.FindName("CatDuct")
        self.mode_selected = self.window.FindName("ModeSelected")
        self.mode_all = self.window.FindName("ModeAll")
        self.target_combo = self.window.FindName("TargetCombo")
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None
        self.param_names = param_names

        self.cat_pipe.IsChecked = bool(settings["pipe"])
        self.cat_duct.IsChecked = bool(settings["duct"])

        # come prima: con una selezione si parte da "solo selezionati", altrimenti da "tutti"
        if selected_pipes or selected_ducts:
            self.mode_selected.Content = "Only the selected pipes and ducts ({} pipes, {} ducts)".format(
                selected_pipes, selected_ducts)
            self.mode_selected.IsChecked = True
        else:
            self.mode_selected.Content = "Only the selected pipes and ducts (nothing selected)"
            self.mode_selected.IsEnabled = False
            self.mode_all.IsChecked = True

        for name in param_names:
            self.target_combo.Items.Add(name)
        if param_names:
            last = settings["target_param"]
            self.target_combo.SelectedIndex = param_names.index(last) if last in param_names else 0
        else:
            self.target_combo.IsEnabled = False
            self.window.FindName("TargetHint").Text = (
                "No writable text instance parameter found on pipes and ducts. "
                "Create a project or shared TEXT parameter first.")

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        pipe, duct = bool(self.cat_pipe.IsChecked), bool(self.cat_duct.IsChecked)
        if not (pipe or duct):
            return False, "Select at least one category.", None
        if self.target_combo.SelectedIndex < 0:
            return False, "There is no target parameter to write to.", None
        return True, "", {
            "pipe": pipe,
            "duct": duct,
            "use_selection": bool(self.mode_selected.IsChecked),
            "target_param": self.param_names[self.target_combo.SelectedIndex],
        }

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Insulation Type Mark To Host")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
