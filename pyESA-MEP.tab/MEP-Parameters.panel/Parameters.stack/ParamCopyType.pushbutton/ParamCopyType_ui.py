# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox, Visibility
from System.Windows.Controls import CheckBox
from System.Windows.Markup import XamlReader

from pyrevit import script


class ParamCopyWindow(object):
    def __init__(self, xaml_path, config_name, source_label, type_label, params):
        self.config_name = config_name
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.source_label = self.window.FindName("SourceLabel")
        self.type_label = self.window.FindName("TypeLabel")
        self.param_list = self.window.FindName("ParamList")
        self.search_box = self.window.FindName("SearchBox")
        self.count_label = self.window.FindName("CountLabel")
        self.overwrite_check = self.window.FindName("OverwriteCheck")
        self.select_all_btn = self.window.FindName("SelectAllButton")
        self.deselect_all_btn = self.window.FindName("DeselectAllButton")
        self.ok_button = self.window.FindName("OkButton")

        self.source_label.Text = source_label
        self.type_label.Text = type_label
        self.result = None

        # una CheckBox per parametro, con il dict del parametro in .Tag
        self.boxes = []
        for p in params:
            cb = CheckBox()
            cb.Content = p["label"]
            cb.Tag = p
            cb.Checked += self._on_check_changed
            cb.Unchecked += self._on_check_changed
            self.param_list.Items.Add(cb)
            self.boxes.append(cb)

        self._load_config()
        self._update_count()

        self.search_box.TextChanged += self._on_search
        self.select_all_btn.Click += self._on_select_all
        self.deselect_all_btn.Click += self._on_deselect_all
        self.ok_button.Click += self._on_ok

    def _load_config(self):
        try:
            cfg = script.get_config(self.config_name)
            self.overwrite_check.IsChecked = bool(cfg.get_option("overwrite", "1") == "1")
        except Exception:
            pass

    def _save_config(self):
        try:
            cfg = script.get_config(self.config_name)
            cfg.overwrite = "1" if self.overwrite_check.IsChecked else "0"
            script.save_config()
        except Exception:
            pass

    def _update_count(self):
        chosen = len([cb for cb in self.boxes if cb.IsChecked])
        self.count_label.Text = "{} of {} parameters selected".format(chosen, len(self.boxes))

    def _on_check_changed(self, sender, args):
        self._update_count()

    def _on_search(self, sender, args):
        # filtra la lista sul testo digitato (visibilita' degli item)
        text = (self.search_box.Text or "").strip().lower()
        for cb in self.boxes:
            visible = (not text) or (text in str(cb.Content).lower())
            cb.Visibility = Visibility.Visible if visible else Visibility.Collapsed

    def _set_visible_checked(self, state):
        for cb in self.boxes:
            if cb.Visibility == Visibility.Visible:
                cb.IsChecked = state

    def _on_select_all(self, sender, args):
        self._set_visible_checked(True)

    def _on_deselect_all(self, sender, args):
        self._set_visible_checked(False)

    def _on_ok(self, sender, args):
        selected = [cb.Tag for cb in self.boxes if cb.IsChecked]
        if not selected:
            MessageBox.Show("Select at least one type parameter to copy.", "Param Copy Type")
            return
        self._save_config()
        self.result = {
            "params": selected,
            "overwrite": bool(self.overwrite_check.IsChecked),
        }
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
