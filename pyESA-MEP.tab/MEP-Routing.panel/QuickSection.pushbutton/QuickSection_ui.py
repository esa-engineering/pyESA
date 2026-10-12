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
    """Finestra modale (SHIFT + CLICK): nome della sezione, margine laterale, margini sotto/sopra il view range, tipo
    e template."""

    def __init__(self, xaml_path, settings, type_names, template_names, default_name):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        find = self.window.FindName
        self.margin_box = find("MarginBox")
        self.below_box = find("BelowBox")
        self.above_box = find("AboveBox")
        self.type_combo = find("TypeCombo")
        self.template_combo = find("TemplateCombo")
        self.name_box = find("NameBox")
        self.name_box.Text = settings.get("section_name") or ""
        find("NameHint").Text = "Empty: automatic name, one per user ({}).".format(default_name)
        self.margin_box.Text = "{:g}".format(settings["margin_mm"])
        self.below_box.Text = "{:g}".format(settings["below_mm"])
        self.above_box.Text = "{:g}".format(settings["above_mm"])
        for name in type_names:
            self.type_combo.Items.Add(name)
        if type_names:
            self.type_combo.SelectedIndex = (type_names.index(settings["type_name"])
                                             if settings["type_name"] in type_names else 0)
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
            return False, "The section name cannot contain these characters: {}".format(" ".join(bad)), None
        ok, margin = parse_number(self.margin_box.Text)
        if not ok or margin < 0:
            return False, "The side margin must be a number (mm), zero or more.", None
        ok, below = parse_number(self.below_box.Text)
        if not ok:
            return False, "The margin below the view range must be a number (mm).", None
        ok, above = parse_number(self.above_box.Text)
        if not ok:
            return False, "The margin above the view range must be a number (mm).", None
        if self.type_combo.SelectedIndex < 0:
            return False, "Choose a section type.", None
        return True, "", {"section_name": name, "margin_mm": margin, "below_mm": below, "above_mm": above,
                          "type_name": str(self.type_combo.SelectedItem),
                          "template_name": str(self.template_combo.SelectedItem)}

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Quick Section")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
