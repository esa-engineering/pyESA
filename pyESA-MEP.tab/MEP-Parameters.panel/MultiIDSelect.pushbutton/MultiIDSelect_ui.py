# -*- coding: utf-8 -*-
import re
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Data")

import System
from System.IO import File
from System.Data import DataTable
from System.Windows import MessageBox, Clipboard, Visibility
from System.Windows.Controls import CheckBox, DataGridTextColumn
from System.Windows.Data import Binding
from System.Windows.Markup import XamlReader

from pyrevit import script

TITLE = "Multi ID Read/Select"

# colonne fisse della tabella: (intestazione, chiave nella riga)
FIXED_COLUMNS = [("Model", "model"), ("Id", "id"), ("Category", "category"),
                 ("Family", "family"), ("Type", "type")]


def parse_ids(text):
    """Estrae gli Id numerici dal testo, senza doppioni e nell'ordine di scrittura."""
    seen = set()
    result = []
    for token in re.findall(r"\d+", text or ""):
        value = int(token)
        if value not in seen:
            seen.add(value)
            result.append(value)
    return result


def group_by_model(rows):
    groups = []
    index = {}
    for r in rows:
        if r["model"] not in index:
            index[r["model"]] = len(groups)
            groups.append((r["model"], []))
        groups[index[r["model"]]][1].append(r)
    return groups


def format_clipboard(rows):
    groups = group_by_model(rows)
    if len(groups) == 1:
        return ", ".join(str(r["id"]) for r in groups[0][1])
    return "\n".join("{}: {}".format(m, ", ".join(str(r["id"]) for r in items)) for m, items in groups)


class MultiIDSelectWindow(object):
    """Finestra non modale. Le azioni su Revit passano dalle callback dello script
    (on_load(), on_select(ids)), che le eseguono via ExternalEvent e poi richiamano
    set_rows() / set_status()."""

    def __init__(self, xaml_path, config_name, on_load, on_select):
        self.config_name = config_name
        self.on_load = on_load
        self.on_select = on_select
        self.rows = []
        self.param_boxes = []
        self.chosen = set()

        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        find = self.window.FindName
        self.ids_box = find("IdsBox")
        self.ids_count = find("IdsCountLabel")
        self.status_text = find("StatusText")
        self.grid = find("SelectionGrid")
        self.selection_count = find("SelectionCountLabel")
        self.param_list = find("ParamList")
        self.param_search = find("ParamSearchBox")
        self.copy_btn = find("CopyButton")
        self.copy_table_btn = find("CopyTableButton")

        self._load_config()
        self._update_count()

        self.ids_box.TextChanged += lambda s, a: self._update_count()
        find("PasteButton").Click += self._on_paste
        find("SelectButton").Click += self._on_select
        find("LoadButton").Click += lambda s, a: self.on_load()
        find("ClearParamsButton").Click += self._on_clear_params
        find("CloseButton").Click += lambda s, a: self.window.Close()
        self.copy_btn.Click += self._on_copy
        self.copy_table_btn.Click += self._on_copy_table
        self.param_search.TextChanged += self._on_param_search
        self.window.Closing += lambda s, a: self._save_config()

    # --- config ---------------------------------------------------------
    def _load_config(self):
        try:
            text = script.get_config(self.config_name).get_option("params", "")
            self.chosen = set(p for p in text.split("|") if p)
        except Exception:
            pass

    def _save_config(self):
        try:
            cfg = script.get_config(self.config_name)
            cfg.params = "|".join(sorted(self.chosen))
            script.save_config()
        except Exception:
            pass

    # --- chiamate dallo script -------------------------------------------
    def set_rows(self, rows):
        """rows: dict con le chiavi di FIXED_COLUMNS piu' 'params' {nome: valore}."""
        self.rows = rows
        self._fill_param_list()
        self._rebuild_grid()
        self.selection_count.Text = "{} element(s) in {} model(s)".format(
            len(rows), len(group_by_model(rows)))
        self.copy_btn.IsEnabled = bool(rows)
        self.copy_table_btn.IsEnabled = bool(rows)

    def set_status(self, text):
        self.status_text.Text = text or ""

    def show(self):
        self.window.Show()

    # --- tabella ---------------------------------------------------------
    def _fill_param_list(self):
        names = set()
        for r in self.rows:
            names.update(r["params"].keys())
        names.update(self.chosen)       # restano visibili anche se assenti nella selezione
        self.param_list.Items.Clear()
        self.param_boxes = []
        for name in sorted(names, key=lambda n: n.lower()):
            cb = CheckBox()
            cb.Content = name
            cb.IsChecked = name in self.chosen
            cb.Checked += self._on_param_toggled
            cb.Unchecked += self._on_param_toggled
            self.param_list.Items.Add(cb)
            self.param_boxes.append(cb)
        self._on_param_search(None, None)

    def _param_columns(self):
        return [str(cb.Content) for cb in self.param_boxes if cb.IsChecked]

    def _rebuild_grid(self):
        params = self._param_columns()
        headers = [h for h, k in FIXED_COLUMNS] + params
        table = DataTable()
        for i, h in enumerate(headers):
            # nomi neutri C0..Cn: i nomi dei parametri possono contenere caratteri
            # che rompono il path del Binding
            col_type = clr.GetClrType(System.Int64) if h == "Id" else clr.GetClrType(System.String)
            table.Columns.Add("C{}".format(i), col_type)
        for r in self.rows:
            values = [r[k] for h, k in FIXED_COLUMNS] + [r["params"].get(p, "") for p in params]
            row = table.NewRow()
            for i, v in enumerate(values):
                row[i] = System.Int64(v) if headers[i] == "Id" else (v or "")
            table.Rows.Add(row)
        self.grid.Columns.Clear()
        for i, h in enumerate(headers):
            col = DataGridTextColumn()
            col.Header = h
            col.Binding = Binding("C{}".format(i))
            self.grid.Columns.Add(col)
        self.grid.ItemsSource = table.DefaultView

    def _on_param_toggled(self, sender, args):
        name = str(sender.Content)
        if sender.IsChecked:
            self.chosen.add(name)
        else:
            self.chosen.discard(name)
        self._rebuild_grid()

    def _on_param_search(self, sender, args):
        text = (self.param_search.Text or "").strip().lower()
        for cb in self.param_boxes:
            visible = (not text) or (text in str(cb.Content).lower())
            cb.Visibility = Visibility.Visible if visible else Visibility.Collapsed

    def _on_clear_params(self, sender, args):
        self.chosen = set()
        for cb in self.param_boxes:
            cb.Unchecked -= self._on_param_toggled
            cb.IsChecked = False
            cb.Unchecked += self._on_param_toggled
        self._rebuild_grid()

    # --- pulsanti --------------------------------------------------------
    def _update_count(self):
        self.ids_count.Text = "{} Id(s) recognised".format(len(parse_ids(self.ids_box.Text)))

    def _on_paste(self, sender, args):
        try:
            if Clipboard.ContainsText():
                self.ids_box.Text = Clipboard.GetText()
        except Exception:
            pass

    def _on_select(self, sender, args):
        ids = parse_ids(self.ids_box.Text)
        if not ids:
            MessageBox.Show("Enter at least one element Id.", TITLE)
            return
        self.on_select(ids)

    def _on_copy(self, sender, args):
        try:
            Clipboard.SetText(format_clipboard(self.rows))
        except Exception:
            MessageBox.Show("Clipboard not available, try again.", TITLE)

    def _on_copy_table(self, sender, args):
        params = self._param_columns()
        lines = ["\t".join([h for h, k in FIXED_COLUMNS] + params)]
        for r in self.rows:
            values = [r[k] for h, k in FIXED_COLUMNS] + [r["params"].get(p, "") for p in params]
            lines.append("\t".join(u"{}".format(v if v is not None else "") for v in values))
        try:
            Clipboard.SetText("\r\n".join(lines))
        except Exception:
            MessageBox.Show("Clipboard not available, try again.", TITLE)
