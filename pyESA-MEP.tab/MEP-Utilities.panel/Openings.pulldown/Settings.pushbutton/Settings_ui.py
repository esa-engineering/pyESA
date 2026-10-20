# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import FontWeights, GridLength, GridUnitType, MessageBox, TextWrapping, Thickness, \
    VerticalAlignment
from System.Windows.Controls import ColumnDefinition, Grid, TextBlock
from System.Windows.Markup import XamlReader
from System.Windows.Media import Brushes

NUMBER_CONTROLS = (("clearance_mm", "ClearanceBox", "Clearance per side"), ("rounding_mm", "RoundingBox", "Round up to"),
                   ("merge_mm", "MergeBox", "Merge within"), ("extra_depth_mm", "ExtraDepthBox",
                                                              "Extra depth per side"),
                   ("layer_gap_mm", "LayerGapBox", "Join layers across"))
LOADED, NOT_LOADED = u"✔ loaded", u"✖ not loaded"


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


class SettingsWindow(object):
    """Impostazioni di Openings: famiglie (fisse, solo se caricate), misure, MEC/ELE separati.
    families: [(caso, nome della famiglia, caricata: True / False / None senza modello aperto)].
    show() ritorna il dict aggiornato, o None se annullata."""

    def __init__(self, xaml_path, settings, families):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.settings = dict(settings)
        find = self.window.FindName
        rows = find("FamilyRows")
        for case, name, loaded in families:
            rows.Children.Add(self._family_row(case, name, loaded))
        missing = [f for f in families if f[2] is False]
        if any(f[2] is None for f in families):
            hint = "No model open: open a model to check the families."
        elif missing:
            hint = "Load the missing families from the company library: Place&Check needs all four."
        else:
            hint = "All four families are loaded."
        find("FamilyHint").Text = hint
        self.boxes = {}
        for key, control, _ in NUMBER_CONTROLS:
            box = find(control)
            box.Text = "{:g}".format(settings[key])
            self.boxes[key] = box
        self.separate = find("SeparateCheck")
        self.separate.IsChecked = bool(settings["separate"])
        find("OkButton").Click += self._on_ok
        self.result = None

    def _family_row(self, case, name, loaded):
        """Riga: caso (muro, rettangolare...), nome della famiglia, caricata / non caricata."""
        grid = Grid()
        grid.Margin = Thickness(0, 3, 0, 3)
        for width in (GridLength(130), GridLength(1, GridUnitType.Star), GridLength(90)):
            column = ColumnDefinition()
            column.Width = width
            grid.ColumnDefinitions.Add(column)
        cells = [(case, None), (name, Brushes.Gray), (
            "-" if loaded is None else (LOADED if loaded else NOT_LOADED),
            Brushes.Gray if loaded is None else (Brushes.Green if loaded else Brushes.Firebrick))]
        for index, (text, brush) in enumerate(cells):
            block = TextBlock()
            block.Text = text
            block.VerticalAlignment = VerticalAlignment.Center
            block.TextWrapping = TextWrapping.Wrap
            if brush is not None:
                block.Foreground = brush
            if index == 2:
                block.FontWeight = FontWeights.Bold
            Grid.SetColumn(block, index)
            grid.Children.Add(block)
        return grid

    def _validate(self):
        """(ok, messaggio, valori): tutti i campi insieme."""
        values = {}
        for key, _, label in NUMBER_CONTROLS:
            ok, number = parse_number(self.boxes[key].Text)
            if not ok or number < 0:
                return False, "{}: enter a number of 0 or more.".format(label), None
            values[key] = number
        values["separate"] = bool(self.separate.IsChecked)
        return True, "", values

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Openings Settings")
            return
        self.settings.update(values)
        self.result = self.settings
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
