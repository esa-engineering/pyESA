# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import FontWeights, GridLength, GridUnitType, HorizontalAlignment, MessageBox, TextWrapping, \
    Thickness, Visibility, VerticalAlignment
from System.Windows.Controls import Button, ColumnDefinition, ComboBox, Grid, RowDefinition, TextBlock, TextBox
from System.Windows.Markup import XamlReader

NONE = "(none)"
WHERE = (("both", "Instance or type"), ("instance", "Instance"), ("type", "Type"))
OPS = (("nonempty", "Is not empty"), ("equals", "Equals"), ("contains", "Contains"))
SIZING_FIELDS = (("clearance", "clearance_mm"), ("rounding", "rounding_mm"), ("merge", "merge_mm"),
                 ("min", "min_mm"), ("extra_depth", "extra_depth_mm"))
SIZING_LABELS = {"clearance": "clearance", "rounding": "rounding step", "merge": "merge within",
                 "min": "minimum size", "extra_depth": "extra depth"}
CLASS_LABELS = {"normal": "Normal", "rei": "REI", "struct": "Structural"}
SIMPLE_KEYS = (("h_rect", "SimpleHRect"), ("h_round", "SimpleHRound"), ("v_rect", "SimpleVRect"),
               ("v_round", "SimpleVRound"))
MEP_KEYS = (("pipe", "MepPipe"), ("duct", "MepDuct"), ("tray", "MepTray"), ("conduit", "MepConduit"))
DIM_KEYS = (("width", "DimWidth"), ("height", "DimHeight"), ("depth", "DimDepth"), ("diameter", "DimDiameter"))


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


def fill_combo(combo, keys, selected):
    combo.Items.Clear()
    combo.Items.Add(NONE)
    for key in keys:
        combo.Items.Add(key)
    if selected and selected not in keys:
        combo.Items.Add(selected)  # tipo non piu' caricato: resta visibile, Place lo segnalera'
    combo.SelectedItem = selected if selected else NONE


def combo_value(combo):
    value = combo.SelectedItem
    return u"" if value is None or value == NONE else u"{}".format(value)


def add_column(grid, width):
    column = ColumnDefinition()
    column.Width = width
    grid.ColumnDefinitions.Add(column)


class SetupWindow(object):
    """Configurazione delle forometrie del progetto. callbacks: test_rei(rules) -> testo,
    check_families(keys, dims) -> testo, export(config) -> testo o None, import_setup() -> config o None."""

    def __init__(self, xaml_path, config, family_keys, mep_kinds, cases, param_items, callbacks):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        find = self.window.FindName
        self.keys = family_keys
        self.mep_kinds, self.cases = mep_kinds, cases  # [(chiave, etichetta)]
        self.callbacks = callbacks
        self.mode_simple, self.mode_matrix = find("ModeSimple"), find("ModeMatrix")
        self.simple_panel, self.matrix_panel = find("SimplePanel"), find("MatrixPanel")
        self.simple_combos = dict((key, find(name)) for key, name in SIMPLE_KEYS)
        self.matrix_combos = {}
        self._build_matrix(find("MatrixGrid"))
        self.mep_checks = dict((key, find(name)) for key, name in MEP_KEYS)
        self.dim_boxes = dict((key, find(name)) for key, name in DIM_KEYS)
        self.rules_panel = find("ReiRulesPanel")
        self.rule_rows = []
        self.sizing = {}
        for cls in CLASS_LABELS:
            for short, _ in SIZING_FIELDS:
                self.sizing[(cls, short)] = find("Sz_{}_{}".format(cls, short))
        self.param_boxes = {}
        self._build_params(find("ParamsGrid"), param_items)
        self.obsolete_mark, self.obsolete_delete = find("ObsoleteMark"), find("ObsoleteDelete")
        self.workset_box, self.schedule_box = find("WorksetBox"), find("ScheduleBox")
        self.filters_check, self.apply_now = find("FiltersCheck"), find("ApplyNowCheck")

        self.mode_simple.Checked += lambda s, a: self._update_mode()
        self.mode_matrix.Checked += lambda s, a: self._update_mode()
        find("AddRuleButton").Click += lambda s, a: self._add_rule({})
        find("TestReiButton").Click += self._on_test_rei
        find("CheckFamiliesButton").Click += self._on_check_families
        find("ExportButton").Click += self._on_export
        find("ImportButton").Click += self._on_import
        find("OkButton").Click += self._on_ok
        self.result = None
        self.load(config)

    # --- costruzione ------------------------------------------------------------------------------------------

    def _build_matrix(self, grid):
        add_column(grid, GridLength(140))
        for _ in self.cases:
            add_column(grid, GridLength(190))
        for _ in range(len(self.mep_kinds) + 1):
            grid.RowDefinitions.Add(RowDefinition())
        for col, (case, label) in enumerate(self.cases, 1):
            header = TextBlock()
            header.Text = label
            header.FontWeight = FontWeights.Bold
            header.TextWrapping = TextWrapping.Wrap
            header.Margin = Thickness(4, 0, 4, 4)
            Grid.SetRow(header, 0)
            Grid.SetColumn(header, col)
            grid.Children.Add(header)
        for row, (kind, label) in enumerate(self.mep_kinds, 1):
            text = TextBlock()
            text.Text = label
            text.VerticalAlignment = VerticalAlignment.Center
            Grid.SetRow(text, row)
            Grid.SetColumn(text, 0)
            grid.Children.Add(text)
            for col, (case, _) in enumerate(self.cases, 1):
                combo = ComboBox()
                combo.Margin = Thickness(4, 2, 4, 2)
                Grid.SetRow(combo, row)
                Grid.SetColumn(combo, col)
                grid.Children.Add(combo)
                self.matrix_combos["{}|{}".format(kind, case)] = combo

    def _build_params(self, grid, param_items):
        for row, (key, label) in enumerate(param_items):
            grid.RowDefinitions.Add(RowDefinition())
            text = TextBlock()
            text.Text = label + ":"
            text.VerticalAlignment = VerticalAlignment.Center
            Grid.SetRow(text, row)
            box = TextBox()
            box.Width = 220
            box.Margin = Thickness(0, 3, 0, 3)
            box.HorizontalAlignment = HorizontalAlignment.Left
            Grid.SetRow(box, row)
            Grid.SetColumn(box, 1)
            grid.Children.Add(text)
            grid.Children.Add(box)
            self.param_boxes[key] = box

    def _add_rule(self, rule):
        row = Grid()
        row.Margin = Thickness(0, 2, 0, 2)
        for width in (GridLength(200), GridLength(140), GridLength(120), GridLength(1, GridUnitType.Star),
                      GridLength(34)):
            add_column(row, width)
        name = TextBox()
        name.Text = rule.get("param", "")
        name.ToolTip = "Parameter name"
        where = ComboBox()
        for _, label in WHERE:
            where.Items.Add(label)
        where.SelectedIndex = [k for k, _ in WHERE].index(rule.get("where", "both")) \
            if rule.get("where", "both") in [k for k, _ in WHERE] else 0
        op = ComboBox()
        for _, label in OPS:
            op.Items.Add(label)
        op.SelectedIndex = [k for k, _ in OPS].index(rule.get("op", "nonempty")) \
            if rule.get("op", "nonempty") in [k for k, _ in OPS] else 0
        value = TextBox()
        value.Text = rule.get("value", "")
        value.ToolTip = "Value (for Equals / Contains)"
        remove = Button()
        remove.Content = u"✕"
        for col, control in enumerate((name, where, op, value, remove)):
            control.Margin = Thickness(0 if col == 0 else 6, 0, 0, 0)
            Grid.SetColumn(control, col)
            row.Children.Add(control)
        entry = (row, name, where, op, value)
        remove.Click += lambda s, a, e=entry: self._remove_rule(e)
        self.rules_panel.Children.Add(row)
        self.rule_rows.append(entry)

    def _remove_rule(self, entry):
        self.rules_panel.Children.Remove(entry[0])
        self.rule_rows.remove(entry)

    def _update_mode(self):
        matrix = bool(self.mode_matrix.IsChecked)
        self.matrix_panel.Visibility = Visibility.Visible if matrix else Visibility.Collapsed
        self.simple_panel.Visibility = Visibility.Collapsed if matrix else Visibility.Visible

    # --- dati -------------------------------------------------------------------------------------------------

    def load(self, config):
        self.mode_matrix.IsChecked = config.get("mode") == "matrix"
        self.mode_simple.IsChecked = config.get("mode") != "matrix"
        self._update_mode()
        for key, combo in self.simple_combos.items():
            fill_combo(combo, self.keys, config.get("simple", {}).get(key, ""))
        for key, combo in self.matrix_combos.items():
            fill_combo(combo, self.keys, config.get("matrix", {}).get(key, ""))
        for key, check in self.mep_checks.items():
            check.IsChecked = bool(config.get("mep", {}).get(key, True))
        for key, box in self.dim_boxes.items():
            box.Text = config.get("dims", {}).get(key, "")
        for entry in list(self.rule_rows):
            self._remove_rule(entry)
        for rule in config.get("rei_rules", []):
            self._add_rule(rule)
        if not self.rule_rows:
            self._add_rule({})
        for (cls, short), box in self.sizing.items():
            field = dict(SIZING_FIELDS)[short]
            box.Text = "{:g}".format(float(config.get("rules", {}).get(cls, {}).get(field, 0)))
        for key, box in self.param_boxes.items():
            box.Text = config.get("params", {}).get(key, "")
        self.obsolete_delete.IsChecked = config.get("obsolete") == "delete"
        self.obsolete_mark.IsChecked = config.get("obsolete") != "delete"
        self.workset_box.Text = config.get("workset", "")
        self.schedule_box.Text = config.get("schedule", "")
        self.filters_check.IsChecked = bool(config.get("filters", True))

    def rules(self):
        result = []
        for _, name, where, op, value in self.rule_rows:
            if (name.Text or "").strip():
                result.append({"param": name.Text.strip(), "where": WHERE[max(0, where.SelectedIndex)][0],
                               "op": OPS[max(0, op.SelectedIndex)][0], "value": (value.Text or "").strip()})
        return result

    def collect(self):
        """(ok, messaggio, valori). Tutti i controlli insieme, come nelle altre finestre."""
        errors = []
        rules = self.rules()
        for rule in rules:
            if rule["op"] != "nonempty" and not rule["value"]:
                errors.append("REI rule on '{}': a value is needed for Equals / Contains.".format(rule["param"]))
        sizing = {}
        for cls, cls_label in CLASS_LABELS.items():
            sizing[cls] = {}
            for short, field in SIZING_FIELDS:
                ok, value = parse_number(self.sizing[(cls, short)].Text)
                if not ok or value < 0:
                    errors.append("{} {}: a number, zero or more (mm).".format(cls_label, SIZING_LABELS[short]))
                    value = 0.0
                sizing[cls][field] = value
        if errors:
            return False, "\n".join(errors), None
        return True, "", {
            "mode": "matrix" if self.mode_matrix.IsChecked else "simple",
            "simple": dict((key, combo_value(combo)) for key, combo in self.simple_combos.items()),
            "matrix": dict((key, combo_value(combo)) for key, combo in self.matrix_combos.items()
                           if combo_value(combo)),
            "mep": dict((key, bool(check.IsChecked)) for key, check in self.mep_checks.items()),
            "dims": dict((key, (box.Text or "").strip()) for key, box in self.dim_boxes.items()),
            "rei_rules": rules,
            "rules": sizing,
            "params": dict((key, (box.Text or "").strip()) for key, box in self.param_boxes.items()),
            "obsolete": "delete" if self.obsolete_delete.IsChecked else "mark",
            "workset": (self.workset_box.Text or "").strip(),
            "schedule": (self.schedule_box.Text or "").strip(),
            "filters": bool(self.filters_check.IsChecked),
        }

    # --- pulsanti ---------------------------------------------------------------------------------------------

    def _on_test_rei(self, sender, args):
        rules = self.rules()
        if not rules:
            MessageBox.Show("Write at least one rule first.", "REI rules")
            return
        MessageBox.Show(self.callbacks["test_rei"](rules), "REI rules: test")

    def _on_check_families(self, sender, args):
        if self.mode_matrix.IsChecked:
            keys = [combo_value(c) for c in self.matrix_combos.values()]
        else:
            keys = [combo_value(c) for c in self.simple_combos.values()]
        keys = sorted(set(k for k in keys if k))
        if not keys:
            MessageBox.Show("Choose at least one family first.", "Check families")
            return
        dims = dict((key, (box.Text or "").strip()) for key, box in self.dim_boxes.items())
        MessageBox.Show(self.callbacks["check_families"](keys, dims), "Check families")

    def _on_export(self, sender, args):
        ok, message, values = self.collect()
        if not ok:
            MessageBox.Show(message, "Openings Setup")
            return
        text = self.callbacks["export"](values)
        if text:
            MessageBox.Show(text, "Export")

    def _on_import(self, sender, args):
        config = self.callbacks["import_setup"]()
        if config is not None:
            self.load(config)

    def _on_ok(self, sender, args):
        ok, message, values = self.collect()
        if not ok:
            MessageBox.Show(message, "Openings Setup")
            return
        self.result = (values, bool(self.apply_now.IsChecked))
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
