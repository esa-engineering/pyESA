# -*- coding: utf-8 -*-
__title__ = "Classification\nTool"
__author__ = "Claude + Antonio Miano"

import os
import codecs
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window
from System.Windows.Controls import TreeViewItem, ComboBoxItem
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode

from pyrevit import forms, revit, DB

doc = revit.doc
selection = revit.get_selection()
SCRIPT_DIR = os.path.dirname(__file__)

# Syntax options
SYNTAX_COMPOSED, SYNTAX_CODE, SYNTAX_DESCRIPTION = 0, 1, 2


class TreeNode:
    __slots__ = ['code', 'title', 'full_code_parts', 'children', 'children_order']
    
    def __init__(self, code, full_code_parts):
        self.code = code
        self.title = ""
        self.full_code_parts = full_code_parts
        self.children = {}
        self.children_order = []


class ClassificationWriterWindow(Window):
    
    def __init__(self, selected_elements):
        self.selected_elements = selected_elements
        self.csv_data = {}
        self.current_tree_root = None
        self.all_parameters = []
        self.selected_node = None
        self.selected_parameter = None
        self.is_ifc = False
        self.result = None
        
        self._load_xaml()
        self._load_parameters()
        self._load_csv_files()
    
    def _load_xaml(self):
        Window.__init__(self)
        stream = FileStream(os.path.join(SCRIPT_DIR, "ClassificationTool_script.xaml"), FileMode.Open)
        root = XamlReader.Load(stream)
        stream.Close()
        
        for prop in ['Content', 'Title', 'Height', 'Width', 'MinHeight', 'MinWidth', 
                     'WindowStartupLocation', 'Background', 'ResizeMode', 'ShowInTaskbar']:
            setattr(self, prop, getattr(root, prop))
        
        # Find controls
        controls = ['cmbClassification', 'txtSearch', 'btnClearSearch', 'treeClassification',
                    'thumbResize', 'btnCompactAll', 'btnExpandAll', 'cmbParameters', 
                    'cmbSyntax', 'btnOk', 'btnCancel']
        for name in controls:
            setattr(self, name, root.FindName(name))
        
        # Wire events
        self.cmbClassification.SelectionChanged += self.OnClassificationChanged
        self.txtSearch.GotFocus += lambda s,e: setattr(self.txtSearch, 'Text', '') if self.txtSearch.Text == "Search..." else None
        self.txtSearch.LostFocus += lambda s,e: setattr(self.txtSearch, 'Text', 'Search...') if not self.txtSearch.Text else None
        self.txtSearch.TextChanged += self.OnSearchChanged
        self.btnClearSearch.Click += lambda s,e: setattr(self.txtSearch, 'Text', '')
        self.treeClassification.SelectedItemChanged += self.OnTreeSelectionChanged
        self.thumbResize.DragDelta += lambda s,e: setattr(self.treeClassification, 'Height', 
            max(100, min(600, self.treeClassification.Height + e.VerticalChange)))
        self.btnCompactAll.Click += lambda s,e: self._set_tree_expanded(False)
        self.btnExpandAll.Click += lambda s,e: self._set_tree_expanded(True)
        self.cmbParameters.SelectionChanged += self.OnParameterChanged
        self.btnOk.Click += self.OnOK
        self.btnCancel.Click += lambda s,e: self.Close()
    
    def _load_csv_files(self):
        csv_files = sorted([f[:-4] for f in os.listdir(SCRIPT_DIR) if f.lower().endswith('.csv')])
        for name in csv_files:
            self.cmbClassification.Items.Add(name)
        if csv_files:
            self.cmbClassification.SelectedIndex = 0
    
    def _load_parameters(self):
        if not self.selected_elements:
            return
        
        param_sets = []
        for elem in self.selected_elements:
            params = set()
            for p in elem.Parameters:
                if p.StorageType == DB.StorageType.String and not p.IsReadOnly:
                    params.add(("", p.Definition.Name, False))
            elem_type = doc.GetElement(elem.GetTypeId())
            if elem_type:
                for p in elem_type.Parameters:
                    if p.StorageType == DB.StorageType.String and not p.IsReadOnly:
                        params.add(("[Type] ", p.Definition.Name, True))
            param_sets.append(params)
        
        if param_sets:
            common = param_sets[0]
            for ps in param_sets[1:]:
                common &= ps
            self.all_parameters = sorted(common, key=lambda x: (x[0] != "", x[1].lower()))
            for prefix, name, is_type in self.all_parameters:
                item = ComboBoxItem()
                item.Content = prefix + name
                item.Tag = (name, is_type)
                self.cmbParameters.Items.Add(item)
    
    def _parse_csv(self, csv_filename):
        csv_path = os.path.join(SCRIPT_DIR, csv_filename + ".csv")
        root = TreeNode("", [])
        is_ifc = "IFC" in csv_filename.upper()
        
        content = None
        for enc in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']:
            try:
                with codecs.open(csv_path, 'r', enc) as f:
                    content = f.read()
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        
        if content is None:
            with open(csv_path, 'rb') as f:
                content = f.read().decode('utf-8', errors='replace')
        
        for line in content.splitlines():
            parts = [p.strip() for p in line.strip().split(';') if p.strip()]
            if not parts:
                continue
            
            if is_ifc:
                # IFC: tutte le colonne sono gerarchia
                hierarchy = parts
                description = parts[-1]
            else:
                # Uniclass: tutte tranne l'ultima sono gerarchia
                hierarchy = parts if len(parts) == 1 else parts[:-1]
                description = parts[-1]
            
            current = root
            for i, code in enumerate(hierarchy):
                if code not in current.children:
                    current.children[code] = TreeNode(code, hierarchy[:i+1])
                    current.children_order.append(code)
                current = current.children[code]
            current.title = description
        
        return root
    
    def _build_tree_view(self, root_node, filter_text=""):
        self.treeClassification.Items.Clear()
        ft = filter_text.lower().strip() if filter_text else ""
        
        def should_show(node):
            if not ft:
                return True
            text = node.code.lower() if self.is_ifc else (node.code + " " + node.title).lower()
            return ft in text or any(should_show(c) for c in node.children.values())
        
        def add_nodes(parent_item, parent_node):
            has_match = False
            for code in parent_node.children_order:
                child = parent_node.children[code]
                if not should_show(child):
                    continue
                
                item = TreeViewItem()
                item.Header = child.code if self.is_ifc else ("{} - {}".format(child.code, child.title) if child.title else child.code)
                item.Tag = child
                
                child_match = add_nodes(item, child)
                this_match = ft and ft in (child.code.lower() if self.is_ifc else (child.code + " " + child.title).lower())
                
                if ft and (this_match or child_match):
                    item.IsExpanded = True
                    has_match = True
                
                (self.treeClassification if parent_item is None else parent_item).Items.Add(item)
            return has_match
        
        add_nodes(None, root_node)
    
    def _set_tree_expanded(self, expanded):
        def recurse(item):
            item.IsExpanded = expanded
            for i in range(item.Items.Count):
                if isinstance(item.Items[i], TreeViewItem):
                    recurse(item.Items[i])
        for i in range(self.treeClassification.Items.Count):
            if isinstance(self.treeClassification.Items[i], TreeViewItem):
                recurse(self.treeClassification.Items[i])
    
    def _update_ok(self):
        self.btnOk.IsEnabled = self.selected_node is not None and self.selected_parameter is not None
    
    def OnClassificationChanged(self, sender, args):
        if not self.cmbClassification.SelectedItem:
            return
        csv_name = str(self.cmbClassification.SelectedItem)
        self.is_ifc = "IFC" in csv_name.upper()
        
        if csv_name not in self.csv_data:
            self.csv_data[csv_name] = self._parse_csv(csv_name)
        
        self.current_tree_root = self.csv_data[csv_name]
        self._build_tree_view(self.current_tree_root)
        self.selected_node = None
        
        if self.is_ifc:
            self.cmbSyntax.SelectedIndex = SYNTAX_DESCRIPTION
            self.cmbSyntax.IsEnabled = False
        else:
            self.cmbSyntax.IsEnabled = True
        
        self._update_ok()
    
    def OnSearchChanged(self, sender, args):
        search = self.txtSearch.Text
        if search == "Search..." or not self.current_tree_root:
            search = ""
        self._build_tree_view(self.current_tree_root, search)
        self.selected_node = None
        self._update_ok()
    
    def OnTreeSelectionChanged(self, sender, args):
        self.selected_node = self.treeClassification.SelectedItem.Tag if self.treeClassification.SelectedItem else None
        self._update_ok()
    
    def OnParameterChanged(self, sender, args):
        item = self.cmbParameters.SelectedItem
        self.selected_parameter = item.Tag if item and hasattr(item, 'Tag') else None
        self._update_ok()
    
    def OnOK(self, sender, args):
        if not self.selected_node or not self.selected_parameter:
            return
        self.result = {
            'node': self.selected_node,
            'parameter': self.selected_parameter,
            'csv_name': str(self.cmbClassification.SelectedItem),
            'syntax': self.cmbSyntax.SelectedIndex
        }
        self.DialogResult = True
        self.Close()


def build_value(node, csv_name, syntax):
    code = "_".join(node.full_code_parts)
    desc = node.title or ""
    if syntax == SYNTAX_CODE:
        return code
    if syntax == SYNTAX_DESCRIPTION:
        return desc
    # SYNTAX_COMPOSED
    prefix = "[{}]".format(csv_name.split('_')[0])
    return "{}{}:{}".format(prefix, code, desc) if desc else "{}{}".format(prefix, code)


def main():
    elements = list(selection.elements)
    if not elements:
        forms.alert("Please select elements first.", title="No Selection", exitscript=True)
    
    window = ClassificationWriterWindow(elements)
    if not window.ShowDialog() or not window.result:
        return
    
    r = window.result
    node, (param_name, is_type), csv_name, syntax = r['node'], r['parameter'], r['csv_name'], r['syntax']
    value = build_value(node, csv_name, syntax)
    
    if not value:
        forms.alert("Selected item has no description.", title="Warning")
        return
    
    with revit.Transaction("Write Classification"):
        count = 0
        for elem in elements:
            target = doc.GetElement(elem.GetTypeId()) if is_type else elem
            if target:
                param = target.LookupParameter(param_name)
                if param and not param.IsReadOnly:
                    param.Set(value)
                    count += 1
        forms.alert("Written to {} element(s).\n\nValue: {}".format(count, value), title="Success")


if __name__ == "__main__":
    main()
