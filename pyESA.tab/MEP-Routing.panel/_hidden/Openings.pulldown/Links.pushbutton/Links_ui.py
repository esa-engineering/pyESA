# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import GridLength, GridUnitType, Thickness, Visibility, VerticalAlignment
from System.Windows.Controls import CheckBox, ComboBox, ColumnDefinition, Grid, TextBlock
from System.Windows.Markup import XamlReader

REI_CHOICES = ("auto", "rei", "not_rei")


class LinksWindow(object):
    """Elenco dei tipi da verificare: spunta (verifica si'/no), REI (auto / forzato), strutturale (solo lettura)."""

    def __init__(self, xaml_path, entries):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.list = self.window.FindName("TypeList")
        self.search = self.window.FindName("SearchBox")
        self.rows = []
        for entry in entries:
            self._add_row(entry)
        self.search.TextChanged += self._on_search
        self.window.FindName("SelectAllButton").Click += lambda s, a: self._set_all(True)
        self.window.FindName("DeselectAllButton").Click += lambda s, a: self._set_all(False)
        self.window.FindName("OkButton").Click += self._on_ok
        self.window.FindName("SummaryText").Text = "{} types in {} models".format(
            len(entries), len(set(e["model"] for e in entries)))
        self.result = None

    def _add_row(self, entry):
        row = Grid()
        row.Margin = Thickness(0, 1, 0, 1)
        for width in (GridLength(1, GridUnitType.Star), GridLength(190), GridLength(80)):
            column = ColumnDefinition()
            column.Width = width
            row.ColumnDefinitions.Add(column)
        check = CheckBox()
        check.Content = u"[{}] {} | {}  ({}){}".format(entry["model"], entry["category"], entry["type"], entry["count"],
                                                     "  NEW" if entry.get("new") else "")
        check.IsChecked = bool(entry["include"])
        check.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(check, 0)
        combo = ComboBox()
        combo.Items.Add("Auto ({} of {} REI by rules)".format(entry["rei_count"], entry["count"]))
        combo.Items.Add("REI")
        combo.Items.Add("Not REI")
        combo.SelectedIndex = REI_CHOICES.index(entry["rei"]) if entry["rei"] in REI_CHOICES else 0
        combo.Margin = Thickness(6, 0, 6, 0)
        Grid.SetColumn(combo, 1)
        label = TextBlock()
        label.Text = "Structural" if entry["structural"] else ""
        label.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(label, 2)
        row.Children.Add(check)
        row.Children.Add(combo)
        row.Children.Add(label)
        self.list.Items.Add(row)
        self.rows.append((entry, row, check, combo, u"{} {} {}".format(entry["model"], entry["category"],
                                                                      entry["type"]).lower()))

    def _on_search(self, sender, args):
        query = (self.search.Text or "").strip().lower()
        for _, row, _, _, text in self.rows:
            row.Visibility = Visibility.Visible if not query or query in text else Visibility.Collapsed

    def _set_all(self, value):
        for _, row, check, _, _ in self.rows:
            if row.Visibility == Visibility.Visible:
                check.IsChecked = value

    def _on_ok(self, sender, args):
        self.result = {}
        for entry, _, check, combo, _ in self.rows:
            self.result[entry["key"]] = {"include": bool(check.IsChecked),
                                         "rei": REI_CHOICES[max(0, combo.SelectedIndex)]}
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
