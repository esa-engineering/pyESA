# -*- coding: utf-8 -*-
__title__ = "Group Params\nTo Elements"
__doc__ = """Version = 1.0
Date    = 28.08.2026
________________________________________________________________
Trasferisce i valori dei parametri dei model group (istanza o tipo)
agli elementi contenuti nei gruppi stessi.

Prima della scrittura viene eseguito un controllo a vuoto: per ogni
categoria presente dentro i gruppi si verifica che il parametro sia
associato, scrivibile e con tipo di dato compatibile. Le anomalie
sono elencate nell'anteprima e nel report finale.
________________________________________________________________
Author(s):
Andrea Patti
"""
__author__ = "Andrea Patti"

import os
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
    Group,
    StorageType,
    Transaction
)
from Autodesk.Revit.UI import TaskDialog

from System.Collections.Generic import Dictionary
from System.Windows import Window
from System.Windows.Markup import XamlReader
from System.IO import StreamReader

from pyrevit import script

# pyRevit entry points
uidoc = __revit__.ActiveUIDocument            # noqa: F821
doc = uidoc.Document
output = script.get_output()

DIALOG_TITLE = "Model Group Params To Elements"
XAML_FILE_NAME = "ModelGroupToElementsForm.xaml"

# Stati usati sia nell'anteprima sia nel report
# StorageType.None non e' scrivibile direttamente in Python ("None" e' riservato)
STORAGE_NONE = getattr(StorageType, "None")

STATE_OK = "OK"
STATE_WARN = "WARN"
STATE_ERROR = "ERROR"

MAX_SAMPLES = 5


# ---------------------------------------------------------------------------
# Helper Revit
# ---------------------------------------------------------------------------

def get_element_id_value(eid):
    """ElementId.IntegerValue e' stato rimosso in Revit 2026."""
    if eid is None:
        return -1
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


MODEL_GROUP_CAT_ID = get_element_id_value(
    ElementId(BuiltInCategory.OST_IOSModelGroups)
)


def is_model_group(el):
    if not isinstance(el, Group):
        return False
    try:
        cat = el.Category
    except Exception:
        return False
    if cat is None:
        return False
    return get_element_id_value(cat.Id) == MODEL_GROUP_CAT_ID


def collect_model_groups():
    """Tutti i model group del modello (anche quelli annidati)."""
    collector = FilteredElementCollector(doc) \
        .OfClass(Group) \
        .WhereElementIsNotElementType() \
        .ToElements()
    return [g for g in collector if is_model_group(g)]


def collect_selected_model_groups():
    groups = []
    try:
        ids = uidoc.Selection.GetElementIds()
    except Exception:
        return groups
    for eid in ids:
        el = doc.GetElement(eid)
        if is_model_group(el):
            groups.append(el)
    return groups


def group_type_name(group):
    try:
        gtype = doc.GetElement(group.GetTypeId())
        if gtype is not None and gtype.Name:
            return gtype.Name
    except Exception:
        pass
    return "<unnamed group type>"


def category_name(el):
    try:
        cat = el.Category
        if cat is not None and cat.Name:
            return cat.Name
    except Exception:
        pass
    return "<no category>"


def get_members(group, recurse):
    """Elementi contenuti nel gruppo, opzionalmente anche nei gruppi annidati."""
    members = []
    seen_groups = set()
    stack = [group]
    while stack:
        current = stack.pop()
        gid = get_element_id_value(current.Id)
        if gid in seen_groups:
            continue
        seen_groups.add(gid)
        try:
            member_ids = current.GetMemberIds()
        except Exception:
            continue
        for mid in member_ids:
            el = doc.GetElement(mid)
            if el is None:
                continue
            if isinstance(el, Group):
                if recurse:
                    stack.append(el)
                continue
            members.append(el)
    return members


def drop_nested_groups(groups):
    """Scarta i gruppi gia' contenuti in un altro gruppo dell'elenco."""
    scope_ids = set(get_element_id_value(g.Id) for g in groups)
    keep = []
    for g in groups:
        nested_in_scope = False
        parent_id = g.GroupId
        guard = 0
        while parent_id is not None \
                and get_element_id_value(parent_id) > 0 \
                and guard < 25:
            if get_element_id_value(parent_id) in scope_ids:
                nested_in_scope = True
                break
            parent = doc.GetElement(parent_id)
            if parent is None:
                break
            parent_id = parent.GroupId
            guard += 1
        if not nested_in_scope:
            keep.append(g)
    return keep


# ---------------------------------------------------------------------------
# Helper parametri
# ---------------------------------------------------------------------------

class ParamRef(object):
    """Un parametro sorgente: istanza del gruppo ('I') o tipo del gruppo ('T')."""

    def __init__(self, scope, name, storage):
        self.scope = scope
        self.name = name
        self.storage = storage

    @property
    def key(self):
        return (self.scope, self.name)

    @property
    def label(self):
        return "{0} [{1}]".format(
            self.name, "Instance" if self.scope == "I" else "Type"
        )


def get_spec_id(param):
    """Identificativo del tipo di dato, per rilevare unita' incompatibili."""
    try:
        definition = param.Definition
        if hasattr(definition, "GetDataType"):
            return definition.GetDataType().TypeId       # Revit 2022+
    except Exception:
        pass
    try:
        return str(param.Definition.ParameterType)       # fallback
    except Exception:
        return None


def read_param_value(param):
    st = param.StorageType
    if st == StorageType.String:
        return param.AsString()
    if st == StorageType.Integer:
        return param.AsInteger()
    if st == StorageType.Double:
        return param.AsDouble()
    if st == StorageType.ElementId:
        return param.AsElementId()
    return None


def is_empty_value(storage, value):
    if storage == StorageType.String:
        return value is None or not value.strip()
    if storage == StorageType.ElementId:
        return value is None or get_element_id_value(value) < 0
    return value is None


def write_param_value(param, storage, value):
    if storage == StorageType.String:
        param.Set(value if value is not None else "")
    elif storage == StorageType.Integer:
        param.Set(int(value))
    elif storage == StorageType.Double:
        param.Set(float(value))
    elif storage == StorageType.ElementId:
        param.Set(value if value is not None else ElementId.InvalidElementId)
    else:
        return False
    return True


def value_signature(storage, value):
    """Chiave di confronto fra istanze dello stesso tipo di gruppo."""
    if value is None:
        return u"<none>"
    if storage == StorageType.ElementId:
        return u"id:{0}".format(get_element_id_value(value))
    if storage == StorageType.Double:
        return u"d:{0:.9f}".format(value)
    return u"v:{0}".format(value)


def value_to_display(storage, value):
    if value is None:
        return u""
    if storage == StorageType.ElementId:
        target = doc.GetElement(value)
        if target is not None:
            try:
                return u"{0} ({1})".format(
                    target.Name, get_element_id_value(value)
                )
            except Exception:
                pass
        return u"{0}".format(get_element_id_value(value))
    if storage == StorageType.Double:
        return u"{0:.6g}".format(value)
    return u"{0}".format(value)


def get_source_param(group, pref):
    """Il parametro sorgente sull'istanza o sul tipo di gruppo."""
    try:
        if pref.scope == "I":
            return group.LookupParameter(pref.name)
        gtype = doc.GetElement(group.GetTypeId())
        if gtype is None:
            return None
        return gtype.LookupParameter(pref.name)
    except Exception:
        return None


def has_type_parameter(el, name):
    """True se il parametro esiste solo sul tipo dell'elemento."""
    try:
        type_id = el.GetTypeId()
        if type_id is None or get_element_id_value(type_id) < 0:
            return False
        el_type = doc.GetElement(type_id)
        if el_type is None:
            return False
        return el_type.LookupParameter(name) is not None
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Motore di analisi / trasferimento
# ---------------------------------------------------------------------------

class Plan(object):
    """Esito del controllo a vuoto, riusato poi dalla transazione."""

    def __init__(self):
        self.rows = {}          # (param, category, status) -> [state, count, samples]
        self.tasks = []         # scritture da eseguire
        self.conflicts = []     # tipi di gruppo con valori discordanti
        self.group_count = 0
        self.type_count = 0
        self.element_count = 0
        self.written = 0
        self.failures = []      # (element_id, param, message)

    def add_row(self, param_label, category, status, state, count, sample_id=None):
        key = (param_label, category, status)
        row = self.rows.get(key)
        if row is None:
            row = [state, 0, []]
            self.rows[key] = row
        row[1] += count
        if sample_id is not None and len(row[2]) < MAX_SAMPLES:
            row[2].append(sample_id)

    def sorted_rows(self):
        order = {STATE_ERROR: 0, STATE_WARN: 1, STATE_OK: 2}
        items = []
        for (param_label, category, status), row in self.rows.items():
            items.append({
                "Parameter": param_label,
                "Category": category,
                "Status": status,
                "State": row[0],
                "Count": row[1],
                "Samples": row[2],
            })
        items.sort(key=lambda r: (order.get(r["State"], 9),
                                  r["Parameter"], r["Category"]))
        return items

    def counters(self):
        ok = warn = err = 0
        for row in self.rows.values():
            if row[0] == STATE_OK:
                ok += row[1]
            elif row[0] == STATE_WARN:
                warn += row[1]
            else:
                err += row[1]
        return ok, warn, err


def build_plan(groups, prefs, opts):
    """Controllo a vuoto: nessuna modifica al modello."""
    plan = Plan()

    if opts["include_nested"]:
        groups = drop_nested_groups(groups)

    plan.group_count = len(groups)

    # istanze raggruppate per tipo di gruppo
    by_type = {}
    for g in groups:
        tid = get_element_id_value(g.GetTypeId())
        by_type.setdefault(tid, []).append(g)
    plan.type_count = len(by_type)

    # valori sorgente per istanza e rilevamento conflitti
    source_values = {}      # (group_id, pref.key) -> (storage, value, spec)
    skip_pairs = set()      # (type_id, pref.key) da saltare per conflitto

    for tid, instances in by_type.items():
        type_label = group_type_name(instances[0])
        for pref in prefs:
            signatures = {}
            for g in instances:
                src = get_source_param(g, pref)
                if src is None:
                    continue
                storage = src.StorageType
                value = read_param_value(src)
                source_values[(get_element_id_value(g.Id), pref.key)] = \
                    (storage, value, get_spec_id(src))
                signatures.setdefault(
                    value_signature(storage, value),
                    value_to_display(storage, value)
                )
            if len(instances) > 1 and len(signatures) > 1:
                plan.conflicts.append({
                    "type": type_label,
                    "instances": len(instances),
                    "param": pref.label,
                    "values": sorted(signatures.values()),
                })
                if opts["skip_conflicts"]:
                    skip_pairs.add((tid, pref.key))

    processed_elements = set()

    for g in groups:
        gid = get_element_id_value(g.Id)
        tid = get_element_id_value(g.GetTypeId())
        type_label = group_type_name(g)
        members = get_members(g, opts["include_nested"])
        for el in members:
            processed_elements.add(get_element_id_value(el.Id))

        for pref in prefs:
            if (tid, pref.key) in skip_pairs:
                plan.add_row(
                    pref.label, type_label,
                    "Skipped: instances of this group type hold different values",
                    STATE_WARN, 1, g.Id
                )
                continue

            source = source_values.get((gid, pref.key))
            if source is None:
                plan.add_row(
                    pref.label, type_label,
                    "Parameter not found on this group (group instances counted)",
                    STATE_WARN, 1, g.Id
                )
                continue

            storage, value, spec = source

            if opts["skip_empty"] and is_empty_value(storage, value):
                plan.add_row(
                    pref.label, type_label,
                    "Empty source value, skipped (group instances counted)",
                    STATE_WARN, 1, g.Id
                )
                continue

            for el in members:
                cat = category_name(el)
                dest = None
                try:
                    dest = el.LookupParameter(pref.name)
                except Exception:
                    dest = None

                if dest is None:
                    if has_type_parameter(el, pref.name):
                        status = ("Parameter exists only as a TYPE parameter "
                                  "of the element, not written")
                    else:
                        status = "Parameter not bound to this category"
                    plan.add_row(pref.label, cat, status,
                                 STATE_ERROR, 1, el.Id)
                    continue

                if dest.IsReadOnly:
                    plan.add_row(pref.label, cat,
                                 "Parameter is read-only on the element",
                                 STATE_ERROR, 1, el.Id)
                    continue

                if dest.StorageType != storage:
                    plan.add_row(
                        pref.label, cat,
                        "Storage type mismatch (group {0} / element {1})".format(
                            storage, dest.StorageType
                        ),
                        STATE_ERROR, 1, el.Id
                    )
                    continue

                dest_spec = get_spec_id(dest)
                if spec and dest_spec and spec != dest_spec:
                    if not opts["allow_unit_mismatch"]:
                        plan.add_row(
                            pref.label, cat,
                            "Different data type / unit, skipped "
                            "(enable the option to force it)",
                            STATE_ERROR, 1, el.Id
                        )
                        continue
                    plan.add_row(pref.label, cat,
                                 "Different data type / unit, transferred anyway",
                                 STATE_WARN, 1, el.Id)

                plan.add_row(pref.label, cat, "Ready to transfer",
                             STATE_OK, 1, el.Id)
                plan.tasks.append({
                    "eid": el.Id,
                    "name": pref.name,
                    "label": pref.label,
                    "category": cat,
                    "storage": storage,
                    "value": value,
                })

    plan.element_count = len(processed_elements)
    return plan


def execute_plan(plan):
    """Scrive i valori pianificati dentro una singola transazione."""
    if not plan.tasks:
        return False

    transaction = Transaction(doc, "Group params to elements")
    transaction.Start()
    try:
        for task in plan.tasks:
            el = doc.GetElement(task["eid"])
            if el is None:
                plan.failures.append(
                    (task["eid"], task["label"], "element not found")
                )
                continue
            try:
                dest = el.LookupParameter(task["name"])
                if dest is None or dest.IsReadOnly:
                    plan.failures.append(
                        (task["eid"], task["label"],
                         "parameter no longer writable")
                    )
                    continue
                if write_param_value(dest, task["storage"], task["value"]):
                    plan.written += 1
                else:
                    plan.failures.append(
                        (task["eid"], task["label"], "unsupported storage type")
                    )
            except Exception as ex:
                plan.failures.append((task["eid"], task["label"], str(ex)))
        transaction.Commit()
        return True
    except Exception as ex:
        if transaction.HasStarted():
            transaction.RollBack()
        plan.failures.append((None, "-", "transaction failed: {0}".format(ex)))
        return False


# ---------------------------------------------------------------------------
# Form
# ---------------------------------------------------------------------------

class ModelGroupToElementsForm(Window):

    def __init__(self):
        self.groups = []
        self.param_refs = []
        self.group_types = []          # (type_id_value, name, [groups])
        self._param_checked = set()
        self._gt_checked = set()
        self._param_boxes = []
        self._gt_boxes = []
        self.plan = None
        self.report = None
        self._loading = False
        self._load_xaml()
        self.reload_scope()

    # -- xaml ---------------------------------------------------------------

    def _load_xaml(self):
        xaml_path = None
        try:
            xaml_path = script.get_bundle_file(XAML_FILE_NAME)
        except Exception:
            xaml_path = None
        if not xaml_path or not os.path.isfile(xaml_path):
            xaml_path = os.path.join(os.path.dirname(__file__), XAML_FILE_NAME)

        Window.__init__(self)
        reader = StreamReader(xaml_path)
        try:
            root = XamlReader.Load(reader.BaseStream)
        finally:
            reader.Close()

        self.Content = root.Content
        self.Title = root.Title
        self.Height = root.Height
        self.Width = root.Width
        self.MinHeight = root.MinHeight
        self.MinWidth = root.MinWidth
        self.WindowStartupLocation = root.WindowStartupLocation
        self.ResizeMode = root.ResizeMode
        self.ShowInTaskbar = root.ShowInTaskbar

        self._find_controls(root)
        self._wire_events()

    def _find_controls(self, root):
        self.lbl_model_name = root.FindName('lbl_model_name')
        self.lbl_scope_info = root.FindName('lbl_scope_info')

        self.rad_scope_all = root.FindName('rad_scope_all')
        self.rad_scope_selection = root.FindName('rad_scope_selection')
        self.btn_refresh = root.FindName('btn_refresh')

        self.chk_include_type_params = root.FindName('chk_include_type_params')
        self.chk_include_nested = root.FindName('chk_include_nested')
        self.chk_skip_empty = root.FindName('chk_skip_empty')
        self.chk_skip_conflicts = root.FindName('chk_skip_conflicts')
        self.chk_allow_unit_mismatch = root.FindName('chk_allow_unit_mismatch')

        self.txt_param_filter = root.FindName('txt_param_filter')
        self.pnl_params = root.FindName('pnl_params')
        self.btn_params_all = root.FindName('btn_params_all')
        self.btn_params_none = root.FindName('btn_params_none')

        self.txt_group_filter = root.FindName('txt_group_filter')
        self.pnl_grouptypes = root.FindName('pnl_grouptypes')
        self.btn_groups_all = root.FindName('btn_groups_all')
        self.btn_groups_none = root.FindName('btn_groups_none')

        self.lst_preview = root.FindName('lst_preview')
        self.lbl_summary = root.FindName('lbl_summary')

        self.btn_analyze = root.FindName('btn_analyze')
        self.btn_transfer = root.FindName('btn_transfer')
        self.btn_close = root.FindName('btn_close')

    def _wire_events(self):
        self.btn_refresh.Click += self.OnRefresh
        self.rad_scope_all.Checked += self.OnScopeChanged
        self.rad_scope_selection.Checked += self.OnScopeChanged
        self.chk_include_type_params.Checked += self.OnScopeChanged
        self.chk_include_type_params.Unchecked += self.OnScopeChanged
        self.txt_param_filter.TextChanged += self.OnParamFilter
        self.txt_group_filter.TextChanged += self.OnGroupFilter
        self.btn_params_all.Click += self.OnParamsAll
        self.btn_params_none.Click += self.OnParamsNone
        self.btn_groups_all.Click += self.OnGroupsAll
        self.btn_groups_none.Click += self.OnGroupsNone
        self.btn_analyze.Click += self.OnAnalyze
        self.btn_transfer.Click += self.OnTransfer
        self.btn_close.Click += self.OnClose

    # -- caricamento dati ---------------------------------------------------

    def _warn(self, message):
        TaskDialog.Show(DIALOG_TITLE, message)

    def reload_scope(self):
        self._loading = True
        try:
            if self.rad_scope_selection.IsChecked == True:
                self.groups = collect_selected_model_groups()
            else:
                self.groups = collect_model_groups()

            self._build_group_types()
            self._build_param_refs()
            self._refresh_header()
            self._render_group_types()
            self._render_params()
            self._invalidate_plan()
        finally:
            self._loading = False

    def _build_group_types(self):
        by_type = {}
        for g in self.groups:
            tid = get_element_id_value(g.GetTypeId())
            entry = by_type.get(tid)
            if entry is None:
                entry = [tid, group_type_name(g), []]
                by_type[tid] = entry
            entry[2].append(g)

        self.group_types = sorted(by_type.values(), key=lambda e: e[1].lower())

        available = set(e[0] for e in self.group_types)
        if not self._gt_checked:
            self._gt_checked = set(available)
        else:
            self._gt_checked = set(
                tid for tid in self._gt_checked if tid in available
            )
            if not self._gt_checked:
                self._gt_checked = set(available)

    def _build_param_refs(self):
        include_type = self.chk_include_type_params.IsChecked == True
        found = {}

        def register(scope, param):
            try:
                name = param.Definition.Name
            except Exception:
                return
            storage = param.StorageType
            if storage == STORAGE_NONE:
                return
            key = (scope, name)
            if key not in found:
                found[key] = ParamRef(scope, name, storage)

        seen_types = set()
        for g in self.groups:
            try:
                for param in g.Parameters:
                    register("I", param)
            except Exception:
                pass
            if include_type:
                tid = get_element_id_value(g.GetTypeId())
                if tid in seen_types:
                    continue
                seen_types.add(tid)
                gtype = doc.GetElement(g.GetTypeId())
                if gtype is None:
                    continue
                try:
                    for param in gtype.Parameters:
                        register("T", param)
                except Exception:
                    pass

        self.param_refs = sorted(
            found.values(), key=lambda p: (p.name.lower(), p.scope)
        )
        available = set(p.key for p in self.param_refs)
        self._param_checked = set(
            k for k in self._param_checked if k in available
        )

    def _refresh_header(self):
        try:
            title = doc.Title
        except Exception:
            title = "-"
        self.lbl_model_name.Text = "Model: {0}".format(title)
        self.lbl_scope_info.Text = (
            "Model groups in scope: {0}  |  group types: {1}  |  "
            "available group parameters: {2}"
        ).format(len(self.groups), len(self.group_types), len(self.param_refs))

    # -- liste con checkbox -------------------------------------------------

    def _make_checkbox(self, content, tooltip, is_checked):
        from System.Windows.Controls import CheckBox
        from System.Windows import Thickness
        box = CheckBox()
        box.Content = content
        box.FontSize = 11
        box.Margin = Thickness(0, 2, 0, 2)
        box.IsChecked = is_checked
        if tooltip:
            box.ToolTip = tooltip
        return box

    def _render_params(self):
        self.pnl_params.Children.Clear()
        self._param_boxes = []

        needle = (self.txt_param_filter.Text or "").strip().lower()
        shown = 0
        for pref in self.param_refs:
            if needle and needle not in pref.name.lower():
                continue
            shown += 1
            box = self._make_checkbox(
                pref.label,
                "Storage type: {0}".format(pref.storage),
                pref.key in self._param_checked
            )

            def on_check(sender, args, key=pref.key):
                if sender.IsChecked == True:
                    self._param_checked.add(key)
                else:
                    self._param_checked.discard(key)
                self._invalidate_plan()

            box.Checked += on_check
            box.Unchecked += on_check
            self.pnl_params.Children.Add(box)
            self._param_boxes.append(box)

        if shown == 0:
            from System.Windows.Controls import TextBlock
            empty = TextBlock()
            empty.Text = "No parameter matches the filter."
            empty.FontSize = 11
            empty.Foreground = self.lbl_scope_info.Foreground
            self.pnl_params.Children.Add(empty)

    def _render_group_types(self):
        self.pnl_grouptypes.Children.Clear()
        self._gt_boxes = []

        needle = (self.txt_group_filter.Text or "").strip().lower()
        shown = 0
        for tid, name, instances in self.group_types:
            if needle and needle not in name.lower():
                continue
            shown += 1
            box = self._make_checkbox(
                "{0}  ({1} inst.)".format(name, len(instances)),
                None,
                tid in self._gt_checked
            )

            def on_check(sender, args, key=tid):
                if sender.IsChecked == True:
                    self._gt_checked.add(key)
                else:
                    self._gt_checked.discard(key)
                self._invalidate_plan()

            box.Checked += on_check
            box.Unchecked += on_check
            self.pnl_grouptypes.Children.Add(box)
            self._gt_boxes.append(box)

        if shown == 0:
            from System.Windows.Controls import TextBlock
            empty = TextBlock()
            empty.Text = "No model group in scope."
            empty.FontSize = 11
            empty.Foreground = self.lbl_scope_info.Foreground
            self.pnl_grouptypes.Children.Add(empty)

    # -- anteprima ----------------------------------------------------------

    def _invalidate_plan(self):
        if self._loading:
            return
        self.plan = None
        self.btn_transfer.IsEnabled = False

    def _current_options(self):
        return {
            "include_nested": self.chk_include_nested.IsChecked == True,
            "skip_empty": self.chk_skip_empty.IsChecked == True,
            "skip_conflicts": self.chk_skip_conflicts.IsChecked == True,
            "allow_unit_mismatch": self.chk_allow_unit_mismatch.IsChecked == True,
        }

    def _checked_params(self):
        return [p for p in self.param_refs if p.key in self._param_checked]

    def _scoped_groups(self):
        return [g for g in self.groups
                if get_element_id_value(g.GetTypeId()) in self._gt_checked]

    def _run_analysis(self):
        prefs = self._checked_params()
        if not prefs:
            self._warn("Select at least one group parameter to transfer.")
            return None
        groups = self._scoped_groups()
        if not groups:
            self._warn("No model group in scope. "
                       "Check the scope and the group type list.")
            return None

        from System.Windows.Input import Cursors
        self.Cursor = Cursors.Wait
        try:
            plan = build_plan(groups, prefs, self._current_options())
        finally:
            self.Cursor = Cursors.Arrow

        self.plan = plan
        self._fill_preview(plan)
        return plan

    def _fill_preview(self, plan):
        self.lst_preview.Items.Clear()
        for item in plan.sorted_rows():
            row = Dictionary[str, object]()
            row["Parameter"] = item["Parameter"]
            row["Category"] = item["Category"]
            row["Count"] = str(item["Count"])
            row["Status"] = item["Status"]
            row["State"] = item["State"]
            self.lst_preview.Items.Add(row)

        ok, warn, err = plan.counters()
        self.btn_transfer.IsEnabled = len(plan.tasks) > 0
        self.lbl_summary.Text = (
            "{0} groups ({1} types), {2} elements analysed - "
            "{3} values ready | {4} warnings | {5} errors"
        ).format(plan.group_count, plan.type_count, plan.element_count,
                 ok, warn, err)

    def _store_report(self, mode, plan):
        self.report = {
            "mode": mode,
            "plan": plan,
            "params": [p.label for p in self._checked_params()],
            "options": self._current_options(),
            "scope": ("Current selection"
                      if self.rad_scope_selection.IsChecked == True
                      else "Whole model"),
        }

    # -- eventi -------------------------------------------------------------

    def OnRefresh(self, sender, args):
        self.reload_scope()

    def OnScopeChanged(self, sender, args):
        if self._loading:
            return
        self.reload_scope()

    def OnParamFilter(self, sender, args):
        if self._loading:
            return
        self._render_params()

    def OnGroupFilter(self, sender, args):
        if self._loading:
            return
        self._render_group_types()

    def OnParamsAll(self, sender, args):
        needle = (self.txt_param_filter.Text or "").strip().lower()
        for pref in self.param_refs:
            if needle and needle not in pref.name.lower():
                continue
            self._param_checked.add(pref.key)
        self._render_params()
        self._invalidate_plan()

    def OnParamsNone(self, sender, args):
        self._param_checked = set()
        self._render_params()
        self._invalidate_plan()

    def OnGroupsAll(self, sender, args):
        needle = (self.txt_group_filter.Text or "").strip().lower()
        for tid, name, _instances in self.group_types:
            if needle and needle not in name.lower():
                continue
            self._gt_checked.add(tid)
        self._render_group_types()
        self._invalidate_plan()

    def OnGroupsNone(self, sender, args):
        self._gt_checked = set()
        self._render_group_types()
        self._invalidate_plan()

    def OnAnalyze(self, sender, args):
        plan = self._run_analysis()
        if plan is None:
            return
        self._store_report("analyze", plan)

    def OnTransfer(self, sender, args):
        plan = self.plan
        if plan is None:
            plan = self._run_analysis()
            if plan is None:
                return
        if not plan.tasks:
            self._warn("There is no value ready to transfer. "
                       "Check the status column.")
            return

        execute_plan(plan)
        self._store_report("transfer", plan)
        self.Close()

    def OnClose(self, sender, args):
        self.Close()


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

def print_report(report):
    plan = report["plan"]
    ok, warn, err = plan.counters()

    output.close_others()
    if report["mode"] == "transfer":
        output.print_md("# Model Group Params To Elements - transfer")
    else:
        output.print_md("# Model Group Params To Elements - analysis only")
        output.print_md("_No change was made to the model._")

    output.print_md("## Summary")
    output.print_md("- Scope: **{0}**".format(report["scope"]))
    output.print_md("- Group instances processed: **{0}** "
                    "(group types: {1})".format(plan.group_count,
                                                plan.type_count))
    output.print_md("- Elements inside the groups: **{0}**"
                    .format(plan.element_count))
    output.print_md("- Parameters: {0}".format(
        ", ".join(report["params"]) if report["params"] else "-"))
    output.print_md("- Ready values: **{0}** | warnings: **{1}** | "
                    "errors: **{2}**".format(ok, warn, err))
    if report["mode"] == "transfer":
        output.print_md("- Values written: **{0}**".format(plan.written))
        output.print_md("- Write failures: **{0}**".format(len(plan.failures)))

    opts = report["options"]
    output.print_md(
        "- Options: nested groups {0}, empty values {1}, "
        "conflicting group types {2}, unit mismatch {3}".format(
            "included" if opts["include_nested"] else "excluded",
            "skipped" if opts["skip_empty"] else "written",
            "skipped" if opts["skip_conflicts"] else "written",
            "allowed" if opts["allow_unit_mismatch"] else "blocked",
        )
    )

    rows = plan.sorted_rows()
    if rows:
        table = []
        for item in rows:
            links = " ".join(
                output.linkify(eid) for eid in item["Samples"]
            ) if item["Samples"] else ""
            status = item["Status"]
            if report["mode"] == "transfer" and status == "Ready to transfer":
                status = "Transferred"
            table.append([
                item["State"],
                item["Parameter"],
                item["Category"],
                item["Count"],
                status,
                links,
            ])
        output.print_md("## Parameter / category check")
        output.print_table(
            table_data=table,
            title="",
            columns=["State", "Parameter", "Category / group type",
                     "Count", "Status", "Samples"]
        )
    else:
        output.print_md("## Parameter / category check")
        output.print_md("No row produced: nothing matched the current scope.")

    if plan.conflicts:
        output.print_md("## Group types with different values across instances")
        output.print_md(
            "Elements inside a group belong to the group definition: writing "
            "a value on a member updates **every instance** of that group type. "
            "The parameters below hold different values on different instances."
        )
        table = []
        for conflict in plan.conflicts:
            table.append([
                conflict["type"],
                conflict["instances"],
                conflict["param"],
                " | ".join(conflict["values"][:6]),
            ])
        output.print_table(
            table_data=table,
            title="",
            columns=["Group type", "Instances", "Parameter", "Values found"]
        )

    if plan.failures:
        output.print_md("## Write failures")
        table = []
        for eid, label, message in plan.failures[:200]:
            table.append([
                output.linkify(eid) if eid is not None else "-",
                label,
                message,
            ])
        output.print_table(
            table_data=table,
            title="",
            columns=["Element", "Parameter", "Message"]
        )
        if len(plan.failures) > 200:
            output.print_md("_...and {0} more._".format(len(plan.failures) - 200))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

try:
    if doc.IsFamilyDocument:
        TaskDialog.Show(
            DIALOG_TITLE,
            "Model groups do not exist in a family document.\n"
            "Open a project model and run the tool again."
        )
    elif not collect_model_groups():
        TaskDialog.Show(
            DIALOG_TITLE,
            "No model group found in the current model."
        )
    else:
        form = ModelGroupToElementsForm()
        form.ShowDialog()
        if form.report is not None:
            print_report(form.report)

except Exception as ex:
    TaskDialog.Show("{0} - Error".format(DIALOG_TITLE), str(ex))
