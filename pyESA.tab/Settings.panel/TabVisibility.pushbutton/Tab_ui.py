# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows.Controls import CheckBox
from System.Windows.Markup import XamlReader


class TabWindow(object):
    """Finestra modale: una spunta per scheda del ribbon (spuntata = visibile).
    show() ritorna ({id: visibile}, ricorda_al_riavvio) o None."""

    def __init__(self, xaml_path, items, remember):
        # items: [(id della scheda, titolo, visibile adesso)]; remember: stato iniziale della spunta
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        find = self.window.FindName
        self.list = find("TabList")
        self.remember = find("RememberCheck")
        self.remember.IsChecked = bool(remember)
        self.boxes = []
        for key, title, visible in sorted(items, key=lambda i: i[1].lower()):
            box = CheckBox()
            box.Content = title
            box.Tag = key
            box.IsChecked = bool(visible)
            self.list.Items.Add(box)
            self.boxes.append(box)
        find("SelectAllButton").Click += lambda s, a: self._set_all(True)
        find("DeselectAllButton").Click += lambda s, a: self._set_all(False)
        find("OkButton").Click += self._on_ok
        self.result = None

    def _set_all(self, value):
        for box in self.boxes:
            box.IsChecked = value

    def _on_ok(self, sender, args):
        visible = dict((box.Tag, bool(box.IsChecked)) for box in self.boxes)
        self.result = (visible, bool(self.remember.IsChecked))
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
