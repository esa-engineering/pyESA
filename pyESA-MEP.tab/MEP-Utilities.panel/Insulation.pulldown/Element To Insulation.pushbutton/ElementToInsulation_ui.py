# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader


def split_names(text):
    """Nomi dei parametri separati da ';', senza spazi ai lati e senza voci vuote."""
    return [name.strip() for name in (text or "").split(";") if name.strip()]


class ElementToInsulationWindow(object):
    """Finestra modale: parametri da copiare, ambito (selezione / tutto), categorie."""

    def __init__(self, xaml_path, selection_count, settings):
        # settings: dict con le ultime impostazioni (parameters, pipe, duct)
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.params_box = self.window.FindName("ParamsBox")
        self.mode_selected = self.window.FindName("ModeSelected")
        self.mode_all = self.window.FindName("ModeAll")
        self.cat_pipe = self.window.FindName("CatPipe")
        self.cat_duct = self.window.FindName("CatDuct")
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

        self.params_box.Text = settings["parameters"]
        self.cat_pipe.IsChecked = bool(settings["pipe"])
        self.cat_duct.IsChecked = bool(settings["duct"])

        # come prima: con una selezione si parte da "solo selezionati", altrimenti da "tutti"
        if selection_count:
            self.mode_selected.Content = "Only the insulation of the selected elements ({})".format(selection_count)
            self.mode_selected.IsChecked = True
        else:
            self.mode_selected.Content = "Only the insulation of the selected elements (nothing selected)"
            self.mode_selected.IsEnabled = False
            self.mode_all.IsChecked = True

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        names = split_names(self.params_box.Text)
        if not names:
            return False, "Enter at least one parameter name.", None
        pipe, duct = bool(self.cat_pipe.IsChecked), bool(self.cat_duct.IsChecked)
        if not (pipe or duct):
            return False, "Select at least one category.", None
        return True, "", {
            "parameters": ";".join(names),
            "names": names,
            "use_selection": bool(self.mode_selected.IsChecked),
            "pipe": pipe,
            "duct": duct,
        }

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Element To Insulation")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
