# -*- coding: utf-8 -*-
"""
mepqto_allowance_ui.py - finestra dell'override della maggiorazione sulle voci del
computo.

Mostra le voci scelte (Type Mark, codice, categorie lineari con la loro maggiorazione e
l'eventuale override) e chiede la percentuale. Restituisce ("set", frazione),
("remove", None) oppure None se annullata.
"""

import os

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.IO import FileStream, FileMode
from System.Windows import Window, MessageBox, WindowStartupLocation
from System.Windows.Markup import XamlReader

from pyrevit import script

import mepqto_store as qs

XAML_FILE_NAME = 'MEPQTO_allowance.xaml'
TITLE = "Allowance override"
# Oltre questa percentuale l'override e' quasi certamente un errore di battitura.
MAX_PERCENT = 1000.0


class AllowanceForm(Window):

    def __init__(self, lines_text, line_count, initial_text, can_remove):
        self.result = None
        self._load_xaml()
        self.txt_items_count.Text = u"({})".format(line_count)
        self.txt_items.Text = lines_text
        self.txt_allowance.Text = initial_text
        self.btn_remove.IsEnabled = can_remove
        self.txt_allowance.Focus()
        self.txt_allowance.SelectAll()

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

        for name in ("txt_items_count", "txt_items", "txt_allowance", "btn_remove",
                     "btn_ok", "btn_cancel"):
            setattr(self, name, root.FindName(name))
        self.btn_remove.Click += self.OnRemove
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel

    def OnOK(self, sender, args):
        text = (self.txt_allowance.Text or u"").replace(u"%", u"").strip()
        try:
            percent = qs.parse_decimal(text)
        except ValueError:
            percent = -1
        if percent is None:
            MessageBox.Show(u"Type a percentage, or use Remove Override to go back to the "
                            u"allowance of the category.", TITLE)
            return
        if percent < 0 or percent > MAX_PERCENT:
            MessageBox.Show(u"'{}' is not a valid allowance: use a percentage between 0 and "
                            u"{:.0f}.".format(text, MAX_PERCENT), TITLE)
            return
        self.result = ("set", percent / 100.0)
        self.Close()

    def OnRemove(self, sender, args):
        self.result = ("remove", None)
        self.Close()

    def OnCancel(self, sender, args):
        self.result = None
        self.Close()


def show_allowance_dialog(owner, lines_text, line_count, initial_text, can_remove):
    """("set", frazione), ("remove", None) oppure None se annullata."""
    form = AllowanceForm(lines_text, line_count, initial_text, can_remove)
    if owner is not None:
        form.Owner = owner
        form.WindowStartupLocation = WindowStartupLocation.CenterOwner
    form.ShowDialog()
    return form.result
