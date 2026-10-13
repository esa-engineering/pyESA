# -*- coding: utf-8 -*-
__title__   = "Check Type Mark\nDuplicates"
__doc__     = """Version = 3.0
Date    = 08.10.2026
Checks duplicated values of a type parameter on the categories you choose
and shows only the duplicated ones in an editable table (values can be
changed and applied directly).

In the window you can:
- choose the parameter to check (any text type parameter of the chosen categories)
- select the rows and press Select in model to select their instances in Revit
- press Refresh to read the model again (unapplied edits are discarded)

CLICK: starts with the Type Mark parameter.
SHIFT + CLICK: starts with the type parameter e_IDA_TypeMark_1 (if present).
Author(s): ESA Engineering
"""
__author__  = "ESA Engineering"

import io
import os.path as op
import clr
for _asm in ("WindowsBase", "PresentationCore", "PresentationFramework", "System.Data"):
    try:
        clr.AddReference(_asm)
    except Exception:
        pass

from System.Collections.Generic import List
from System.Data import DataTable
from System.Windows.Markup import XamlReader

from pyrevit import revit, script, DB, forms

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

XAML_FILE_NAME = "CheckTypeMarkDuplicates_form.xaml"

TYPE_MARK_LABEL = "Type Mark"             # Type Mark built-in (il nome reale dipende dalla lingua di Revit)
CUSTOM_PARAM_NAME = "e_IDA_TypeMark_1"    # parametro di tipo custom, preselezionato con SHIFT + CLICK
START_PARAM = CUSTOM_PARAM_NAME if __shiftclick__ else TYPE_MARK_LABEL  # noqa: F821


# -------------------------------
# HELPER
# -------------------------------
def get_element_id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


def get_name(elem):
    # Element.Name via descriptor: evita problemi di binding su alcuni ElementType
    try:
        return DB.Element.Name.GetValue(elem)
    except Exception:
        return elem.Name


def is_type_mark(param):
    return getattr(param.Definition, "BuiltInParameter", None) == DB.BuiltInParameter.ALL_MODEL_TYPE_MARK


def get_param(elem_type, name):
    """Parametro di tipo per nome; TYPE_MARK_LABEL indica il Type Mark built-in."""
    if name == TYPE_MARK_LABEL:
        p = elem_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK)
        if p is None:
            p = elem_type.LookupParameter("Type Mark")
        return p
    return elem_type.LookupParameter(name)


def get_value(elem_type, name):
    p = get_param(elem_type, name)
    if p is None or p.StorageType != DB.StorageType.String:
        return None
    return (p.AsString() or "").strip()


def collect_types_by_category():
    """Dict {id categoria: (categoria, [ElementType])} per i tipi con Type Mark."""
    result = {}
    for et in DB.FilteredElementCollector(doc).WhereElementIsElementType().ToElements():
        cat = et.Category
        if cat is None:
            continue
        if get_param(et, TYPE_MARK_LABEL) is None:
            continue
        result.setdefault(get_element_id_value(cat.Id), (cat, []))[1].append(et)
    return result


def collect_param_names(types_by_cat, cat_keys):
    """Nomi dei parametri di tipo di testo presenti nelle categorie scelte (Type Mark per primo)."""
    names = set()
    has_type_mark = False
    for key in cat_keys:
        for et in types_by_cat[key][1]:
            for p in et.Parameters:
                if p.StorageType != DB.StorageType.String:
                    continue
                if is_type_mark(p):
                    has_type_mark = True
                else:
                    names.add(p.Definition.Name)
    ordered = sorted(names, key=lambda n: n.lower())
    return ([TYPE_MARK_LABEL] if has_type_mark else []) + ordered


def ask_categories(types_by_cat):
    """Selezione categorie; ritorna lista di chiavi (id categoria) o None."""
    labels = {}
    for key, (cat, types) in types_by_cat.items():
        labels["{}  ({} types)".format(cat.Name, len(types))] = key
    picked = forms.SelectFromList.show(
        sorted(labels.keys()),
        title="Select categories to check",
        multiselect=True,
        button_name="Check",
    )
    if not picked:
        return None
    return [labels[p] for p in picked]


def find_duplicates(types_by_cat, cat_keys, param_name):
    """Lista di tuple (categoria, tipo, valore) dei soli duplicati per categoria."""
    rows = []
    for key in cat_keys:
        cat, types = types_by_cat[key]
        groups = {}
        for et in types:
            value = get_value(et, param_name)
            if value:
                groups.setdefault(value, []).append(et)
        for value, group in groups.items():
            if len(group) > 1:
                for et in group:
                    rows.append((cat, et, value))
    rows.sort(key=lambda r: (r[0].Name, r[2], get_name(r[1])))
    return rows


def collect_instances(cat_keys, types_by_cat):
    """Dict {id tipo: [ElementId delle istanze]}."""
    result = {}
    for key in cat_keys:
        cat = types_by_cat[key][0]
        for el in (DB.FilteredElementCollector(doc)
                   .OfCategoryId(cat.Id)
                   .WhereElementIsNotElementType()
                   .ToElements()):
            result.setdefault(get_element_id_value(el.GetTypeId()), []).append(el.Id)
    return result


# -------------------------------
# FINESTRA
# -------------------------------
class DuplicatesWindow(object):
    def __init__(self, types_by_cat, cat_keys):
        self.types_by_cat = types_by_cat
        self.cat_keys = cat_keys
        self.param_name = START_PARAM
        self.type_by_id = {}   # id (str) -> ElementType
        self.original = {}     # id (str) -> valore originale
        self.instances = {}    # id tipo (int) -> [ElementId]
        self.loading = False   # True mentre si riempie il menu: ignora SelectionChanged

        xaml_path = script.get_bundle_file(XAML_FILE_NAME) or \
            op.join(op.dirname(__file__), XAML_FILE_NAME)
        with io.open(xaml_path, "r", encoding="utf-8") as f:
            self.win = XamlReader.Parse(f.read())

        self.dg = self.win.FindName("dg_types")
        self.cmb_param = self.win.FindName("cmb_param")
        self.lbl_status = self.win.FindName("lbl_status")
        self.btn_apply = self.win.FindName("btn_apply")
        self.btn_close = self.win.FindName("btn_close")
        self.btn_refresh = self.win.FindName("btn_refresh")
        self.btn_select = self.win.FindName("btn_select")
        self.btn_apply.Click += self.on_apply
        self.btn_close.Click += lambda s, e: self.win.Close()
        self.btn_refresh.Click += self.on_refresh
        self.btn_select.Click += self.on_select
        self.cmb_param.SelectionChanged += self.on_param_changed

        self.fill_params()
        self.refresh()

    def fill_params(self):
        """Riempie il menu dei parametri mantenendo la scelta corrente, se ancora disponibile."""
        names = collect_param_names(self.types_by_cat, self.cat_keys)
        self.loading = True
        self.cmb_param.Items.Clear()
        for n in names:
            self.cmb_param.Items.Add(n)
        if self.param_name not in names:
            self.param_name = TYPE_MARK_LABEL if TYPE_MARK_LABEL in names else (names[0] if names else None)
        if self.param_name is not None:
            self.cmb_param.SelectedItem = self.param_name
        self.loading = False

    def refresh(self):
        param = self.param_name
        rows = find_duplicates(self.types_by_cat, self.cat_keys, param) if param else []
        self.instances = collect_instances(self.cat_keys, self.types_by_cat)

        table = DataTable("Duplicates")
        for col in ("Value", "Category", "Family", "TypeName", "Instances", "TypeId", "Sep"):
            table.Columns.Add(col, clr.GetClrType(str))

        self.type_by_id = {}
        self.original = {}
        prev_group = None
        for cat, et, value in rows:
            # le righe sono ordinate per (categoria, valore): Sep = "1" sulla prima riga di ogni nuovo gruppo
            group = (get_element_id_value(cat.Id), value)
            sep = "1" if prev_group is not None and group != prev_group else "0"
            prev_group = group
            tid = get_element_id_value(et.Id)
            sid = str(tid)
            self.type_by_id[sid] = et
            self.original[sid] = value
            try:
                family = et.FamilyName
            except Exception:
                family = ""
            table.Rows.Add(value, cat.Name, family, get_name(et), str(len(self.instances.get(tid, []))), sid, sep)

        self.table = table
        self.dg.ItemsSource = table.DefaultView
        label = param or "-"
        self.dg.Columns[0].Header = label
        self.win.Title = "Check Duplicates: {}".format(label)
        n_values = len(set((r[0].Name, r[2]) for r in rows))
        self.lbl_status.Text = "{} duplicated {} value(s) on {} type(s).".format(n_values, label, len(rows))
        self.btn_apply.IsEnabled = len(rows) > 0
        self.btn_select.IsEnabled = len(rows) > 0

    def on_param_changed(self, sender, args):
        if self.loading or self.cmb_param.SelectedItem is None:
            return
        self.param_name = str(self.cmb_param.SelectedItem)
        self.refresh()

    def on_refresh(self, sender, args):
        """Rilegge il modello: tipi, parametri disponibili e duplicati. Le modifiche non applicate si perdono."""
        self.types_by_cat = collect_types_by_category()
        self.cat_keys = [k for k in self.cat_keys if k in self.types_by_cat]
        self.fill_params()
        self.refresh()
        self.lbl_status.Text = "Refreshed. " + self.lbl_status.Text

    def on_select(self, sender, args):
        """Seleziona in Revit le istanze dei tipi scelti nella tabella e chiude la finestra."""
        self.dg.CommitEdit()
        picked = list(self.dg.SelectedItems)
        if not picked:
            self.lbl_status.Text = "Select one or more rows first."
            return
        ids = List[DB.ElementId]()
        for item in picked:
            for eid in self.instances.get(int(item["TypeId"]), []):
                ids.Add(eid)
        if ids.Count == 0:
            self.lbl_status.Text = "The selected types have no instances in the model."
            return
        uidoc.Selection.SetElementIds(ids)
        self.win.Close()

    def on_apply(self, sender, args):
        self.dg.CommitEdit()
        param_name = self.param_name
        changes = []
        for row in self.table.Rows:
            sid = row["TypeId"]
            new = (row["Value"] or "").strip()
            if new != self.original.get(sid):
                changes.append((self.type_by_id[sid], new))
        if not changes:
            self.lbl_status.Text = "No changes to apply."
            return

        failed = []
        with revit.Transaction("Edit " + param_name):
            for et, new in changes:
                p = get_param(et, param_name)
                try:
                    if p is None or p.IsReadOnly:
                        raise Exception("read-only")
                    p.Set(new)
                except Exception as ex:
                    failed.append("{} ({})".format(get_name(et), ex))

        self.refresh()
        msg = "{} type(s) updated.".format(len(changes) - len(failed))
        if failed:
            msg += " Failed: " + "; ".join(failed) + "."
        self.lbl_status.Text = msg + " " + self.lbl_status.Text

    def show(self):
        self.win.ShowDialog()


# -------------------------------
# MAIN
# -------------------------------
def main():
    types_by_cat = collect_types_by_category()
    if not types_by_cat:
        forms.alert("No element types with a '{}' parameter found.".format(TYPE_MARK_LABEL),
                    exitscript=True)

    cat_keys = ask_categories(types_by_cat)
    if cat_keys is None:
        script.exit()

    DuplicatesWindow(types_by_cat, cat_keys).show()


main()
