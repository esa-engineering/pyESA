# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


class SpacingWindow(object):
    """Finestra modale: distanza, come misurarla, quale elemento resta fermo."""

    def __init__(self, xaml_path, settings):
        # settings: dict con gap_mm, measure ("faces"|"centerlines"), anchor ("first"|"last"|"centre")
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.gap_box = self.window.FindName("GapBox")
        self.measure_faces = self.window.FindName("MeasureFaces")
        self.measure_centers = self.window.FindName("MeasureCenters")
        self.anchor_first = self.window.FindName("AnchorFirst")
        self.anchor_last = self.window.FindName("AnchorLast")
        self.anchor_centre = self.window.FindName("AnchorCentre")
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

        self.gap_box.Text = "{:g}".format(settings["gap_mm"])
        self.measure_centers.IsChecked = settings["measure"] == "centerlines"
        self.measure_faces.IsChecked = settings["measure"] != "centerlines"
        self.anchor_last.IsChecked = settings["anchor"] == "last"
        self.anchor_centre.IsChecked = settings["anchor"] == "centre"
        self.anchor_first.IsChecked = settings["anchor"] not in ("last", "centre")

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        ok, gap = parse_number(self.gap_box.Text)
        if not ok or gap < 0:
            return False, "The distance must be a number, zero or more (mm).", None
        return True, "", {
            "gap_mm": gap,
            "measure": "centerlines" if self.measure_centers.IsChecked else "faces",
            "anchor": "last" if self.anchor_last.IsChecked else ("centre" if self.anchor_centre.IsChecked else "first"),
        }

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Spacing")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
