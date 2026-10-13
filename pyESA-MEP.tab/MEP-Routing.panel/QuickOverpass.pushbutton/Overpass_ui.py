# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox, Visibility, Thickness
from System.Windows.Controls import CheckBox
from System.Windows.Markup import XamlReader

PRESET_ANGLES = (10, 15, 30, 45, 60, 90)
CUSTOM_LABEL = "Custom"


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


class OverpassWindow(object):
    """Finestra modale: offset (mm, con segno) e angolo delle salite."""

    def __init__(self, xaml_path, default_offset_mm, default_angle_deg):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.offset_box = self.window.FindName("OffsetBox")
        self.angle_combo = self.window.FindName("AngleCombo")
        self.angle_box = self.window.FindName("AngleBox")
        self.window.FindName("OkButton").Click += self._on_ok
        self.result = None

        self.offset_box.Text = "{:g}".format(default_offset_mm)
        for angle in PRESET_ANGLES:
            self.angle_combo.Items.Add("{} degrees".format(angle))
        self.angle_combo.Items.Add(CUSTOM_LABEL)

        if default_angle_deg in PRESET_ANGLES:
            self.angle_combo.SelectedIndex = PRESET_ANGLES.index(default_angle_deg)
        else:
            self.angle_combo.SelectedIndex = len(PRESET_ANGLES)
        self.angle_box.Text = "{:g}".format(default_angle_deg)
        self.angle_combo.SelectionChanged += self._on_angle_mode
        self._on_angle_mode(None, None)

    def _is_custom(self):
        return self.angle_combo.SelectedIndex == len(PRESET_ANGLES)

    def _on_angle_mode(self, sender, args):
        self.angle_box.IsEnabled = self._is_custom()

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        ok, offset = parse_number(self.offset_box.Text)
        if not ok:
            return False, "The offset must be a number (mm).", None
        if abs(offset) < 1e-6:
            return False, "The offset must not be zero.", None
        if self._is_custom():
            ok, angle = parse_number(self.angle_box.Text)
            if not ok or angle <= 0 or angle > 90:
                return False, "The custom angle must be a number greater than 0 and up to 90 degrees.", None
        else:
            angle = float(PRESET_ANGLES[self.angle_combo.SelectedIndex])
        return True, "", {"offset_mm": offset, "angle_deg": angle}

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Overpass")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result


ACTIONS = ("over", "under", "none")


class ClashWindow(object):
    """Finestra modale: interferenze trovate, azione (sopra/sotto/niente), distanza e angolo delle salite."""

    def __init__(self, xaml_path, clash_items, categories, settings):
        # clash_items: [(etichetta, chiave categoria ostacolo)]; categories: [(chiave, nome, controllata)]
        clash_labels = [label for label, _ in clash_items]
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.clash_list = self.window.FindName("ClashList")
        self.search_box = self.window.FindName("SearchBox")
        self.clearance_box = self.window.FindName("ClearanceBox")
        self.angle_combo = self.window.FindName("AngleCombo")
        self.angle_box = self.window.FindName("AngleBox")
        self.radios = {"over": self.window.FindName("ActionOver"),
                       "under": self.window.FindName("ActionUnder"),
                       "none": self.window.FindName("ActionNone")}
        self.window.FindName("SummaryText").Text = (
            "{} clashes found between the selected runs and other MEP elements. Tick the ones to handle, "
            "choose the action, then OK. Stop ends the tool and shows the report.".format(len(clash_labels)))
        self.window.FindName("OkButton").Click += self._on_ok
        self.window.FindName("SelectAllButton").Click += lambda s, a: self._set_all(True)
        self.window.FindName("DeselectAllButton").Click += lambda s, a: self._set_all(False)
        self.window.FindName("ResizeThumb").DragDelta += self._on_resize
        self.search_box.TextChanged += self._on_search
        self.result = None

        # una CheckBox per interferenza, l'indice nella lista originale va in Tag
        self.boxes = []
        self.box_category = {}
        for index, (label, key) in enumerate(clash_items):
            box = CheckBox()
            box.Content = label
            box.Tag = index
            box.IsChecked = True
            self.boxes.append(box)
            self.box_category[index] = key
            self.clash_list.Items.Add(box)

        # una CheckBox per categoria di ostacolo: tolta la spunta, le sue interferenze si accettano
        self.category_boxes = []
        panel = self.window.FindName("CategoryPanel")
        for key, name, checked in categories:
            box = CheckBox()
            count = len([k for k in self.box_category.values() if k == key])
            box.Content = "{} ({} clashes)".format(name, count) if count else "{} (accepted, no clashes listed)".format(name)
            box.Tag = key
            box.IsChecked = checked
            box.Margin = Thickness(0, 2, 0, 2)
            box.Checked += self._on_category
            box.Unchecked += self._on_category
            self.category_boxes.append(box)
            panel.Children.Add(box)
        self._refresh_visibility()

        action = settings["action"] if settings["action"] in ACTIONS else "over"
        self.radios[action].IsChecked = True
        self.clearance_box.Text = "{:g}".format(settings["clearance_mm"])
        for angle in PRESET_ANGLES:
            self.angle_combo.Items.Add("{} degrees".format(angle))
        self.angle_combo.Items.Add(CUSTOM_LABEL)
        angle = settings["angle_deg"]
        if angle in PRESET_ANGLES:
            self.angle_combo.SelectedIndex = PRESET_ANGLES.index(angle)
        else:
            self.angle_combo.SelectedIndex = len(PRESET_ANGLES)
        self.angle_box.Text = "{:g}".format(angle)
        self.angle_combo.SelectionChanged += self._on_angle_mode
        self._on_angle_mode(None, None)

    def _set_all(self, value):
        for box in self.boxes:
            if box.Visibility == Visibility.Visible:
                box.IsChecked = value

    def _accepted_categories(self):
        return set(box.Tag for box in self.category_boxes if not box.IsChecked)

    def _refresh_visibility(self):
        """Interferenze visibili: categoria controllata e testo di ricerca."""
        text = (self.search_box.Text or "").lower()
        accepted = self._accepted_categories()
        for box in self.boxes:
            shown = self.box_category[box.Tag] not in accepted and text in str(box.Content).lower()
            box.Visibility = Visibility.Visible if shown else Visibility.Collapsed

    def _on_category(self, sender, args):
        self._refresh_visibility()

    def _on_search(self, sender, args):
        self._refresh_visibility()

    def _on_resize(self, sender, args):
        self.clash_list.Height = max(80, self.clash_list.Height + args.VerticalChange)

    def _is_custom(self):
        return self.angle_combo.SelectedIndex == len(PRESET_ANGLES)

    def _on_angle_mode(self, sender, args):
        self.angle_box.IsEnabled = self._is_custom()

    def _action(self):
        for name in ACTIONS:
            if self.radios[name].IsChecked:
                return name
        return "over"

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        ok, clearance = parse_number(self.clearance_box.Text)
        if not ok or clearance < 0:
            return False, "The clearance must be a number (mm), zero or more.", None
        if self._is_custom():
            ok, angle = parse_number(self.angle_box.Text)
            if not ok or angle <= 0 or angle > 90:
                return False, "The custom angle must be a number greater than 0 and up to 90 degrees.", None
        else:
            angle = float(PRESET_ANGLES[self.angle_combo.SelectedIndex])
        accepted = self._accepted_categories()
        checked = [box.Tag for box in self.boxes if box.IsChecked and self.box_category[box.Tag] not in accepted]
        return True, "", {"action": self._action(), "checked": checked, "accepted_categories": accepted,
                          "clearance_mm": clearance, "angle_deg": angle}

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Overpass")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
