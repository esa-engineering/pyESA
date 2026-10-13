# -*- coding: utf-8 -*-
"""
mepqto_params_ui.py - finestra di mappatura dei parametri del modello.

Tre gruppi, tutti con ComboBox modificabili (si sceglie fra i parametri trovati su un
campione di elementi oppure se ne scrive il nome):

* 10 parametri di tipo con i codici delle categorie a pezzo;
* 10 parametri di istanza con i codici di tutte le categorie (per le a pezzo si
  sommano a quelli di tipo);
* il Si/No d'istanza di inclusione nel computo, per tutte le categorie.

Restituisce una mepqto_model.ParameterMap, oppure None se annullata.
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

import mepqto_model as qm

XAML_FILE_NAME = 'MEPQTO_params.xaml'
TITLE = "Model parameters"


class ParametersForm(Window):

    def __init__(self, param_map, parameter_names):
        self.result = None
        self._choices = [u""] + list(parameter_names)
        self._load_xaml()
        self._piece = self._build_rows(self.grd_piece, u"Type parameter with price code {}")
        self._linear = self._build_rows(self.grd_linear, u"Instance parameter with price code {}")
        self.cbo_include.ItemsSource = self._choices
        self._show(param_map)

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

        for name in ("grd_piece", "grd_linear", "cbo_include",
                     "btn_defaults", "btn_ok", "btn_cancel"):
            setattr(self, name, root.FindName(name))
        self.btn_defaults.Click += self.OnDefaults
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel

    def _build_rows(self, grid, tooltip):
        combos = []
        for index in range(qm.CODE_SLOTS):
            grid.RowDefinitions.Add(RowDefinition())
            label = TextBlock()
            label.Text = u"Code {}:".format(index + 1)
            label.Style = self._styles["FieldLabel"]
            Grid.SetRow(label, index)
            Grid.SetColumn(label, 0)

            combo = ComboBox()
            combo.IsEditable = True
            combo.Style = self._styles["ComboStyle"]
            combo.ItemsSource = self._choices
            combo.ToolTip = tooltip.format(index + 1)
            Grid.SetRow(combo, index)
            Grid.SetColumn(combo, 1)

            grid.Children.Add(label)
            grid.Children.Add(combo)
            combos.append(combo)
        return combos

    def _show(self, param_map):
        for combo, name in zip(self._piece, param_map.piece_codes):
            combo.Text = name
        for combo, name in zip(self._linear, param_map.linear_codes):
            combo.Text = name
        self.cbo_include.Text = param_map.include

    def values(self):
        def texts(combos):
            return [(combo.Text or u"").strip() for combo in combos]
        return qm.ParameterMap(texts(self._piece), texts(self._linear),
                               (self.cbo_include.Text or u"").strip())

    @staticmethod
    def _duplicates(names):
        used = [name for name in names if name]
        return sorted(set(name for name in used if used.count(name) > 1))

    def _validate(self, param_map):
        for title, names in ((u"type parameters", param_map.piece_codes),
                             (u"instance parameters", param_map.linear_codes)):
            duplicates = self._duplicates(names)
            if duplicates:
                return False, u"The same parameter is used twice in the price codes of the " \
                              u"{}: {}.".format(title, u", ".join(duplicates))
        # Sulle categorie a pezzo si leggono entrambi i gruppi: un nome in comune
        # leggerebbe due volte lo stesso parametro.
        shared = sorted(set(n for n in param_map.piece_codes if n) &
                        set(n for n in param_map.linear_codes if n))
        if shared:
            return False, u"These parameters are mapped both as type and as instance price " \
                          u"codes: {}.".format(u", ".join(shared))
        codes = set(name for name in param_map.piece_codes + param_map.linear_codes if name)
        if param_map.include and param_map.include in codes:
            return False, u"'{}' is mapped both as a price code and as the Yes/No " \
                          u"include parameter.".format(param_map.include)
        if not any(param_map.piece_codes) and not any(param_map.linear_codes):
            return False, u"Map at least one price code parameter."
        return True, u""

    def OnDefaults(self, sender, args):
        self._show(qm.ParameterMap.defaults())

    def OnOK(self, sender, args):
        param_map = self.values()
        is_valid, message = self._validate(param_map)
        if not is_valid:
            MessageBox.Show(message, TITLE)
            return
        self.result = param_map
        self.Close()

    def OnCancel(self, sender, args):
        self.result = None
        self.Close()


def show_parameters_dialog(owner, param_map, parameter_names):
    """ParameterMap scelta dall'utente oppure None se annullata."""
    form = ParametersForm(param_map, parameter_names)
    if owner is not None:
        form.Owner = owner
        form.WindowStartupLocation = WindowStartupLocation.CenterOwner
    form.ShowDialog()
    return form.result
