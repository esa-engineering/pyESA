# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows.Markup import XamlReader


class ConnectorWindow(object):
    """Finestra di scelta del connettore dell'attrezzatura.
    labels: un testo per riga, gia' in ordine di distanza (la prima e' preselezionata).
    show() ritorna l'indice scelto, o None se annullata."""

    def __init__(self, xaml_path, labels):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.result = None
        self.list_box = self.window.FindName("ConnectorList")
        for label in labels:
            self.list_box.Items.Add(label)
        self.list_box.SelectedIndex = 0
        self.list_box.MouseDoubleClick += self._on_double_click
        self.window.FindName("OkButton").Click += self._on_ok
        self.window.FindName("CancelButton").Click += self._on_cancel

    def _on_ok(self, sender, args):
        if self.list_box.SelectedIndex < 0:
            return
        self.result = self.list_box.SelectedIndex
        self.window.DialogResult = True

    def _on_double_click(self, sender, args):
        self._on_ok(sender, args)

    def _on_cancel(self, sender, args):
        self.window.DialogResult = False

    def show(self):
        self.window.ShowDialog()
        return self.result
