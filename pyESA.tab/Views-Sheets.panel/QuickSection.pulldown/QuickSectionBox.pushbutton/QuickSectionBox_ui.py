# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader

NO_TEMPLATE = "<None>"
FORBIDDEN_NAME_CHARS = "\\:{}[]|;<>?`~"  # caratteri non ammessi da Revit nei nomi delle viste


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


class SettingsWindow(object):
    """Finestra modale (SHIFT + CLICK): nome della vista 3D, margine, margini sotto/sopra il view range, template."""

    def __init__(self, xaml_path, settings, template_names, default_name):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        find = self.window.FindName
        self.margin_box = find("MarginBox")
        self.below_box = find("BelowBox")
        self.above_box = find("AboveBox")
        self.template_combo = find("TemplateCombo")
        self.name_box = find("NameBox")
        self.name_box.Text = settings.get("view_name") or ""
        find("NameHint").Text = "Empty: automatic name, one per user ({}).".format(default_name)
        self.margin_box.Text = "{:g}".format(settings["margin_mm"])
        self.below_box.Text = "{:g}".format(settings["below_mm"])
        self.above_box.Text = "{:g}".format(settings["above_mm"])
        templates = [NO_TEMPLATE] + list(template_names)
        for name in templates:
            self.template_combo.Items.Add(name)
        self.template_combo.SelectedIndex = (templates.index(settings["template_name"])
                                             if settings["template_name"] in templates else 0)
        find("OkButton").Click += self._on_ok
        self.result = None

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        name = (self.name_box.Text or "").strip()
        bad = [c for c in FORBIDDEN_NAME_CHARS if c in name]
        if bad:
            return False, "The 3D view name cannot contain these characters: {}".format(" ".join(bad)), None
        ok, margin = parse_number(self.margin_box.Text)
        if not ok or margin < 0:
            return False, "The margin must be a number (mm), zero or more.", None
        ok, below = parse_number(self.below_box.Text)
        if not ok:
            return False, "The margin below the view range must be a number (mm).", None
        ok, above = parse_number(self.above_box.Text)
        if not ok:
            return False, "The margin above the view range must be a number (mm).", None
        return True, "", {"view_name": name, "margin_mm": margin, "below_mm": below, "above_mm": above,
                          "template_name": str(self.template_combo.SelectedItem)}

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Quick SectionBox")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
