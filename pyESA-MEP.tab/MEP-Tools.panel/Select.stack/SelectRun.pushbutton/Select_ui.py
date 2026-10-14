# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import Visibility
from System.Windows.Controls import CheckBox
from System.Windows.Markup import XamlReader


class ExcludeWindow(object):
    """Finestra modale (SHIFT + CLICK): categorie trovate, con il numero di elementi; quelle spuntate vengono
    selezionate (spunta = seleziona). Si salvano pero' le escluse, cosi' una categoria mai vista prima resta
    selezionata. show() ritorna l'insieme degli id di categoria esclusi (non spuntati), o None se annullata."""

    def __init__(self, xaml_path, title, summary, categories, excluded):
        # categories: [(id categoria, nome, numero di elementi)]; excluded: id esclusi la volta precedente
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        find = self.window.FindName
        self.window.Title = "{} - Categories".format(title)
        find("TitleText").Text = title.upper()
        find("SummaryText").Text = summary
        self.list = find("CategoryList")
        self.search = find("SearchBox")
        self.boxes = []
        for category_id, name, count in sorted(categories, key=lambda c: c[1].lower()):
            box = CheckBox()
            box.Content = "{}  ({})".format(name, count)
            box.Tag = (category_id, name)
            box.IsChecked = category_id not in excluded  # spunta = seleziona
            self.list.Items.Add(box)
            self.boxes.append(box)
        self.search.TextChanged += self._on_search
        find("ClearSearchButton").Click += self._on_clear
        find("ResizeThumb").DragDelta += self._on_resize
        find("SelectAllButton").Click += lambda s, a: self._check_visible(True)
        find("DeselectAllButton").Click += lambda s, a: self._check_visible(False)
        find("OkButton").Click += self._on_ok
        self.result = None

    def _on_search(self, sender, args):
        text = (self.search.Text or "").strip().lower()
        for box in self.boxes:
            box.Visibility = Visibility.Visible if text in box.Tag[1].lower() else Visibility.Collapsed

    def _on_clear(self, sender, args):
        self.search.Text = ""

    def _on_resize(self, sender, args):
        height = self.list.Height + args.VerticalChange
        if 80 <= height <= 600:
            self.list.Height = height

    def _check_visible(self, value):
        for box in self.boxes:
            if box.Visibility == Visibility.Visible:
                box.IsChecked = value

    def _on_ok(self, sender, args):
        self.result = set(box.Tag[0] for box in self.boxes if not box.IsChecked)
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
