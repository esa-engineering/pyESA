# -*- coding: utf-8 -*-
"""Finestre di Openings Place: cosa selezionare (quando non c'e' selezione) e anteprima prima di creare i fori.
Importato solo da place_main(); nessun documento o output a livello di modulo (Rocket mode)."""
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import Visibility
from System.Windows.Markup import XamlReader

<<<<<<< HEAD
PICK_HOSTS, PICK_MEPS, PICK_AREA = "hosts", "meps", "area"
=======
PICK_HOSTS, PICK_MEPS = "hosts", "meps"
>>>>>>> 21b9cf6a0a03e5f0df727b4d43f9b7745f20f764
# valori dell'anteprima: chiave del dict passato a ConfirmWindow -> TextBlock della finestra
PREVIEW_FIELDS = (("mode", "ModeValue"), ("crossings", "CrossingsValue"), ("openings", "OpeningsValue"),
                  ("existing", "ExistingValue"), ("inside", "InsideValue"), ("extended", "ExtendedValue"))


def _load(xaml_path):
    return XamlReader.Parse(File.ReadAllText(xaml_path))


class PickWindow(object):
<<<<<<< HEAD
    """Cosa selezionare: muri, solai e strutture, elementi MEP, oppure un'area. show() ritorna PICK_HOSTS,
    PICK_MEPS, PICK_AREA o None se annullata."""
=======
    """Cosa selezionare: muri, solai e strutture, oppure elementi MEP. show() ritorna PICK_HOSTS, PICK_MEPS o None
    se annullata."""
>>>>>>> 21b9cf6a0a03e5f0df727b4d43f9b7745f20f764

    def __init__(self, xaml_path, last_pick, settings_text):
        self.window = _load(xaml_path)
        find = self.window.FindName
<<<<<<< HEAD
        self.radios = {PICK_HOSTS: find("HostsRadio"), PICK_MEPS: find("MepsRadio"), PICK_AREA: find("AreaRadio")}
        self.radios.get(last_pick, self.radios[PICK_HOSTS]).IsChecked = True
=======
        self.hosts = find("HostsRadio")
        self.meps = find("MepsRadio")
        (self.meps if last_pick == PICK_MEPS else self.hosts).IsChecked = True
>>>>>>> 21b9cf6a0a03e5f0df727b4d43f9b7745f20f764
        find("SettingsInfo").Text = settings_text
        find("OkButton").Click += self._on_ok
        self.result = None

    def _on_ok(self, sender, args):
<<<<<<< HEAD
        self.result = [key for key, radio in self.radios.items() if radio.IsChecked][0]
=======
        self.result = PICK_MEPS if self.meps.IsChecked else PICK_HOSTS
>>>>>>> 21b9cf6a0a03e5f0df727b4d43f9b7745f20f764
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result


class ConfirmWindow(object):
    """Anteprima dei fori da creare: valori di PREVIEW_FIELDS e problemi di lettura. show() ritorna True per
    creare."""

    def __init__(self, xaml_path, values, problems):
        self.window = _load(xaml_path)
        find = self.window.FindName
        for key, control in PREVIEW_FIELDS:
            find(control).Text = u"{}".format(values.get(key, ""))
        if problems:
            find("ProblemsText").Text = u"\n".join(problems)
        else:
            find("ProblemsBorder").Visibility = Visibility.Collapsed
        find("OkButton").Click += self._on_ok
        self.result = False

    def _on_ok(self, sender, args):
        self.result = True
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
