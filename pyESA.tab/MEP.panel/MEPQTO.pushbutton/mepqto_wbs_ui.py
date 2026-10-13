# -*- coding: utf-8 -*-
"""
mepqto_wbs_ui.py - finestra di scelta dei parametri WBS (fino a 15 livelli).

Ogni livello e' un ComboBox modificabile: si sceglie un parametro fra quelli trovati
su un campione di elementi del modello, oppure se ne scrive il nome. I livelli vuoti
si saltano. La finestra restituisce sempre una lista di WBS_LEVELS nomi (stringa
vuota per i livelli spenti), oppure None se annullata.
"""

import os

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.IO import FileStream, FileMode
from System.Windows import Window, MessageBox, WindowStartupLocation
from System.Windows.Controls import ComboBox, RowDefinition, TextBlock, Grid
from System.Windows.Markup import XamlReader

from pyrevit import script

XAML_FILE_NAME = 'MEPQTO_wbs.xaml'
TITLE = "WBS levels"


class WbsForm(Window):

    def __init__(self, levels, parameter_names, level_count):
        self.result = None
        self._level_count = level_count
        self._combos = []
        self._load_xaml()
        self._build_levels(levels, parameter_names)

    def _load_xaml(self):
        xaml_path = script.get_bundle_file(XAML_FILE_NAME)
        if not xaml_path or not os.path.exists(xaml_path):
            xaml_path = os.path.join(os.path.dirname(__file__), XAML_FILE_NAME)

        Window.__init__(self)

        stream = FileStream(xaml_path, FileMode.Open)
        try:
            root = XamlReader.Load(stream)
        finally:
            stream.Close()

        self.Content = root.Content
        self.Title = root.Title
        self.Height = root.Height
        self.Width = root.Width
        self.MinHeight = root.MinHeight
        self.MinWidth = root.MinWidth
        self.WindowStartupLocation = root.WindowStartupLocation
        self.ResizeMode = root.ResizeMode
        self.ShowInTaskbar = root.ShowInTaskbar
        self._styles = root.Resources

        self.grd_levels = root.FindName('grd_levels')
        self.btn_clear = root.FindName('btn_clear')
        self.btn_ok = root.FindName('btn_ok')
        self.btn_cancel = root.FindName('btn_cancel')
        self.btn_clear.Click += self.OnClear
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel

    def _build_levels(self, levels, parameter_names):
        choices = [u""] + list(parameter_names)
        for index in range(self._level_count):
            self.grd_levels.RowDefinitions.Add(RowDefinition())
            label = TextBlock()
            label.Text = u"Level {}:".format(index + 1)
            label.Style = self._styles["FieldLabel"]
            Grid.SetRow(label, index)
            Grid.SetColumn(label, 0)

            combo = ComboBox()
            combo.IsEditable = True
            combo.Style = self._styles["ComboStyle"]
            combo.ItemsSource = choices
            combo.Text = levels[index] if index < len(levels) else u""
            combo.ToolTip = u"Parameter of WBS level {} (empty = level not used)".format(index + 1)
            Grid.SetRow(combo, index)
            Grid.SetColumn(combo, 1)

            self.grd_levels.Children.Add(label)
            self.grd_levels.Children.Add(combo)
            self._combos.append(combo)

    def values(self):
        return [(combo.Text or u"").strip() for combo in self._combos]

    def _validate(self, values):
        used = [value for value in values if value]
        duplicates = sorted(set(value for value in used if used.count(value) > 1))
        if duplicates:
            return False, u"The same parameter is used on more than one level: {}.".format(
                u", ".join(duplicates))
        return True, u""

    def OnClear(self, sender, args):
        for combo in self._combos:
            combo.Text = u""

    def OnOK(self, sender, args):
        values = self.values()
        is_valid, message = self._validate(values)
        if not is_valid:
            MessageBox.Show(message, TITLE)
            return
        self.result = values
        self.Close()

    def OnCancel(self, sender, args):
        self.result = None
        self.Close()


def show_wbs_dialog(owner, levels, parameter_names, level_count):
    """Lista di level_count nomi (vuoti per i livelli spenti) oppure None se annullata."""
    form = WbsForm(levels, parameter_names, level_count)
    if owner is not None:
        form.Owner = owner
        form.WindowStartupLocation = WindowStartupLocation.CenterOwner
    form.ShowDialog()
    return form.result
