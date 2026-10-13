# -*- coding: utf-8 -*-
"""
mepqto_layout_ui.py - disposizione della scheda Bill of quantities: livelli sulle righe e
livelli WBS sulle colonne.

Tre elenchi: campi disponibili, righe, colonne. Un campo si aggiunge in coda alle righe o
(solo i livelli WBS) alle colonne: l'ordine di inserimento e' l'ordine di raggruppamento,
e Up / Down lo cambiano. La finestra restituisce (chiavi delle righe, chiavi delle
colonne), oppure None se annullata.
"""

import os

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.IO import FileStream, FileMode
from System.Windows import Window, MessageBox, WindowStartupLocation
from System.Windows.Controls import ListBoxItem, TextBlock
from System.Windows.Markup import XamlReader

from pyrevit import script

import mepqto_model as qm

XAML_FILE_NAME = 'MEPQTO_layout.xaml'
TITLE = "Bill of quantities layout"


class LayoutForm(Window):

    def __init__(self, options, rows, columns):
        self.result = None
        # options: [(chiave, etichetta)] nell'ordine di group_options
        self._options = list(options)
        self._labels = dict(self._options)
        self._rows = [key for key in rows if key in self._labels]
        self._columns = [key for key in columns
                         if key in self._labels and qm.wbs_level_index(key) is not None
                         and key not in self._rows]
        self._load_xaml()
        self._refresh()

    def _load_xaml(self):
        xaml_path = script.get_bundle_file(XAML_FILE_NAME)
        if not xaml_path or not os.path.exists(xaml_path):
            xaml_path = os.path.join(os.path.dirname(__file__), XAML_FILE_NAME)

        Window.__init__(self)

        stream = FileStream(xaml_path, FileMode.Open)
        try:
            root = XamlReader.Load(stream)
        finally:
            stream.Close()

        self.Content = root.Content
        self.Title = root.Title
        self.Height = root.Height
        self.Width = root.Width
        self.MinHeight = root.MinHeight
        self.MinWidth = root.MinWidth
        self.WindowStartupLocation = root.WindowStartupLocation
        self.ResizeMode = root.ResizeMode
        self.ShowInTaskbar = root.ShowInTaskbar

        for name in ("lst_available", "btn_add_rows", "btn_add_columns",
                     "lst_rows", "btn_rows_up", "btn_rows_down", "btn_rows_remove",
                     "lst_columns", "btn_columns_up", "btn_columns_down", "btn_columns_remove",
                     "btn_default", "btn_ok", "btn_cancel"):
            setattr(self, name, root.FindName(name))

        self.btn_add_rows.Click += self.OnAddRows
        self.btn_add_columns.Click += self.OnAddColumns
        self.lst_available.MouseDoubleClick += self.OnAddRows
        self.btn_rows_up.Click += lambda s, a: self._move(self._rows, self.lst_rows, -1)
        self.btn_rows_down.Click += lambda s, a: self._move(self._rows, self.lst_rows, 1)
        self.btn_rows_remove.Click += lambda s, a: self._remove(self._rows, self.lst_rows)
        self.btn_columns_up.Click += lambda s, a: self._move(self._columns, self.lst_columns, -1)
        self.btn_columns_down.Click += lambda s, a: self._move(self._columns, self.lst_columns, 1)
        self.btn_columns_remove.Click += lambda s, a: self._remove(self._columns,
                                                                   self.lst_columns)
        self.btn_default.Click += self.OnDefault
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel

    # ------------------------------------------------------------------ elenchi

    def _item(self, key):
        # Testo in un TextBlock: un '_' nel nome di un parametro WBS non sparisce.
        block = TextBlock()
        block.Text = self._labels[key]
        item = ListBoxItem()
        item.Content = block
        item.Tag = key
        return item

    def _fill(self, list_box, keys, selected=None):
        list_box.Items.Clear()
        for key in keys:
            item = self._item(key)
            list_box.Items.Add(item)
            if key == selected:
                item.IsSelected = True

    def _available(self):
        used = set(self._rows) | set(self._columns)
        return [key for key, _ in self._options if key not in used]

    def _refresh(self, rows_selected=None, columns_selected=None):
        self._fill(self.lst_available, self._available())
        self._fill(self.lst_rows, self._rows, rows_selected)
        self._fill(self.lst_columns, self._columns, columns_selected)

    def _selected_available(self):
        # nell'ordine dell'elenco, cosi' una selezione multipla entra in ordine
        return [item.Tag for item in self.lst_available.Items if item.IsSelected]

    def _move(self, keys, list_box, step):
        item = list_box.SelectedItem
        if item is None:
            return
        index = keys.index(item.Tag)
        target = index + step
        if not 0 <= target < len(keys):
            return
        keys[index], keys[target] = keys[target], keys[index]
        if keys is self._rows:
            self._refresh(rows_selected=item.Tag)
        else:
            self._refresh(columns_selected=item.Tag)

    def _remove(self, keys, list_box):
        item = list_box.SelectedItem
        if item is None:
            return
        keys.remove(item.Tag)
        self._refresh()

    # ------------------------------------------------------------------ eventi

    def OnAddRows(self, sender, args):
        chosen = self._selected_available()
        if not chosen:
            return
        self._rows.extend(chosen)
        self._refresh()

    def OnAddColumns(self, sender, args):
        chosen = self._selected_available()
        if not chosen:
            return
        wbs = [key for key in chosen if qm.wbs_level_index(key) is not None]
        if len(wbs) != len(chosen):
            MessageBox.Show(u"Only WBS levels can go on the columns.", TITLE)
        if wbs:
            self._columns.extend(wbs)
            self._refresh()

    def OnDefault(self, sender, args):
        self._rows = [key for key in qm.DEFAULT_GROUP_BY if key in self._labels]
        self._columns = []
        self._refresh()

    def OnOK(self, sender, args):
        self.result = (list(self._rows), list(self._columns))
        self.Close()

    def OnCancel(self, sender, args):
        self.result = None
        self.Close()


def show_layout_dialog(owner, options, rows, columns):
    """(chiavi delle righe, chiavi delle colonne) oppure None se annullata."""
    form = LayoutForm(options, rows, columns)
    if owner is not None:
        form.Owner = owner
        form.WindowStartupLocation = WindowStartupLocation.CenterOwner
    form.ShowDialog()
    return form.result
