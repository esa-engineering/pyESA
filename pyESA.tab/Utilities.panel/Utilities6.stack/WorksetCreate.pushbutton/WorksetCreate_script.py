# -*- coding: utf-8 -*-
"""Create user worksets in the current Revit model.

Names can be loaded from an Excel file (.xlsx, first column),
from a text file (comma / semicolon / newline separated) or typed
manually in the input box.
"""

__title__ = "Workset\nCreate"
__author__ = "pyESA"

import os
import re
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xml')

from Autodesk.Revit.DB import (
    Transaction,
    Workset,
    WorksetTable,
    FilteredWorksetCollector,
    WorksetKind
)
from Autodesk.Revit.UI import TaskDialog

from System.Collections.Generic import Dictionary
from System.Windows import Window
from System.Windows.Markup import XamlReader
from System.IO import StreamReader, File
from System.Text import Encoding, UTF8Encoding
from System.Xml import XmlDocument
from Microsoft.Win32 import OpenFileDialog

# pyRevit entry points
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

DIALOG_TITLE = "Workset Create"

# Characters Revit does not allow inside a workset name
INVALID_CHARS = '\\:{}[]|;<>?`~'
MAX_NAME_LEN = 100


# ---------------------------------------------------------------------------
# xlsx reader (pure zip + xml, no Excel required)
# ---------------------------------------------------------------------------

SS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def _open_archive(path):
    clr.AddReference('System.IO.Compression')
    clr.AddReference('System.IO.Compression.FileSystem')
    from System.IO.Compression import ZipFile
    return ZipFile.OpenRead(path)


def _read_entry(archive, name):
    """Return the text content of a zip entry, or None."""
    entry = archive.GetEntry(name)
    if entry is None:
        # zip entries are case sensitive: fall back to a tolerant lookup
        lname = name.lower()
        for e in archive.Entries:
            if e.FullName.lower() == lname:
                entry = e
                break
    if entry is None:
        return None
    stream = entry.Open()
    try:
        reader = StreamReader(stream, Encoding.UTF8)
        try:
            return reader.ReadToEnd()
        finally:
            reader.Close()
    finally:
        stream.Close()


def _children(node, local_name):
    """Direct child elements with the given local name."""
    out = []
    for child in node.ChildNodes:
        if child.NodeType.ToString() == 'Element' \
                and child.LocalName == local_name:
            out.append(child)
    return out


def _collect_text(node, acc):
    """Recursively collect <t> text, ignoring phonetic runs <rPh>."""
    for child in node.ChildNodes:
        if child.NodeType.ToString() != 'Element':
            continue
        if child.LocalName == 'rPh':
            continue
        if child.LocalName == 't':
            acc.append(child.InnerText)
        else:
            _collect_text(child, acc)


def _load_shared_strings(archive):
    xml = _read_entry(archive, 'xl/sharedStrings.xml')
    if not xml:
        return []
    docx = XmlDocument()
    docx.LoadXml(xml)
    strings = []
    for si in _children(docx.DocumentElement, 'si'):
        acc = []
        _collect_text(si, acc)
        strings.append(''.join(acc))
    return strings


def get_xlsx_sheets(path):
    """Return an ordered list of (sheet_name, sheet_xml_part_path)."""
    archive = _open_archive(path)
    try:
        # relationship id -> part path
        rels = {}
        rel_xml = _read_entry(archive, 'xl/_rels/workbook.xml.rels')
        if rel_xml:
            rdoc = XmlDocument()
            rdoc.LoadXml(rel_xml)
            for rel in _children(rdoc.DocumentElement, 'Relationship'):
                rid = rel.GetAttribute('Id')
                target = rel.GetAttribute('Target')
                if not target:
                    continue
                if target.startswith('/'):
                    target = target[1:]
                elif not target.startswith('xl/'):
                    target = 'xl/' + target.replace('../', '')
                rels[rid] = target

        wb_xml = _read_entry(archive, 'xl/workbook.xml')
        if not wb_xml:
            return []
        wdoc = XmlDocument()
        wdoc.LoadXml(wb_xml)

        sheets = []
        for sheets_node in _children(wdoc.DocumentElement, 'sheets'):
            for sh in _children(sheets_node, 'sheet'):
                name = sh.GetAttribute('name')
                rid = sh.GetAttribute('r:id') or sh.GetAttribute('id')
                part = rels.get(rid)
                if part is None:
                    part = 'xl/worksheets/sheet{0}.xml'.format(len(sheets) + 1)
                sheets.append((name, part))
        return sheets
    finally:
        archive.Dispose()


def _cell_column(ref):
    """'B12' -> 'B'."""
    letters = ''
    for ch in ref:
        if ch.isalpha():
            letters += ch.upper()
        else:
            break
    return letters


def _cell_value(cell, shared):
    ctype = cell.GetAttribute('t')
    if ctype == 's':
        vs = _children(cell, 'v')
        if not vs:
            return ''
        try:
            idx = int(vs[0].InnerText)
        except ValueError:
            return ''
        if 0 <= idx < len(shared):
            return shared[idx]
        return ''
    if ctype == 'inlineStr':
        acc = []
        _collect_text(cell, acc)
        return ''.join(acc)
    vs = _children(cell, 'v')
    if not vs:
        acc = []
        _collect_text(cell, acc)
        return ''.join(acc)
    raw = vs[0].InnerText
    # numbers stored as 100.0 should read as "100"
    try:
        fval = float(raw)
        if fval == int(fval):
            return str(int(fval))
    except (ValueError, OverflowError):
        pass
    return raw


def read_xlsx_first_column(path, sheet_part, skip_header):
    """Return the list of raw strings found in column A of a sheet."""
    archive = _open_archive(path)
    try:
        shared = _load_shared_strings(archive)
        xml = _read_entry(archive, sheet_part)
        if not xml:
            return []
        sdoc = XmlDocument()
        sdoc.LoadXml(xml)

        values = []
        for sheet_data in _children(sdoc.DocumentElement, 'sheetData'):
            for row in _children(sheet_data, 'row'):
                for cell in _children(row, 'c'):
                    ref = cell.GetAttribute('r')
                    if ref and _cell_column(ref) != 'A':
                        continue
                    values.append(_cell_value(cell, shared))
                    break  # first column only
        if skip_header and values:
            values = values[1:]
        return values
    finally:
        archive.Dispose()


# ---------------------------------------------------------------------------
# Text parsing
# ---------------------------------------------------------------------------

def split_names(text):
    """Split a free text blob on commas, semicolons and line breaks."""
    if not text:
        return []
    parts = re.split(r'[,;\r\n\t]+', text)
    return [p.strip() for p in parts]


def read_text_file(path):
    """Read a text file as UTF-8, falling back to the system codepage."""
    data = File.ReadAllBytes(path)
    try:
        strict_utf8 = UTF8Encoding(False, True)
        content = strict_utf8.GetString(data)
        # strip a leading BOM if present
        if content and content[0] == u'\ufeff':
            content = content[1:]
    except Exception:
        content = Encoding.Default.GetString(data)
    return split_names(content)


# ---------------------------------------------------------------------------
# Revit helpers
# ---------------------------------------------------------------------------

def get_user_worksets():
    """Return {lowercase name: Workset} for user worksets."""
    result = {}
    collector = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    for ws in collector:
        result[ws.Name.strip().lower()] = ws
    return result


def validate_name(name):
    """Return None if valid, otherwise an error message."""
    if not name:
        return "empty name"
    if len(name) > MAX_NAME_LEN:
        return "name longer than {0} characters".format(MAX_NAME_LEN)
    bad = [c for c in name if c in INVALID_CHARS]
    if bad:
        return "invalid character(s): {0}".format(' '.join(sorted(set(bad))))
    return None


# ---------------------------------------------------------------------------
# Form
# ---------------------------------------------------------------------------

class WorksetCreateForm(Window):
    def __init__(self):
        self._rows = []
        self._sheets = []
        self._load_xaml()

    # -- xaml ---------------------------------------------------------------

    def _load_xaml(self):
        script_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(script_dir, 'WorksetCreateForm.xaml')

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
        self._refresh_header()

    def _find_controls(self, root):
        self.lbl_model_name = root.FindName('lbl_model_name')
        self.lbl_existing = root.FindName('lbl_existing')
        self.lbl_summary = root.FindName('lbl_summary')
        self.lst_preview = root.FindName('lst_preview')

        self.txt_excel_path = root.FindName('txt_excel_path')
        self.cmb_sheet = root.FindName('cmb_sheet')
        self.chk_skip_header = root.FindName('chk_skip_header')
        self.btn_browse_excel = root.FindName('btn_browse_excel')
        self.btn_load_excel = root.FindName('btn_load_excel')

        self.txt_txt_path = root.FindName('txt_txt_path')
        self.btn_browse_txt = root.FindName('btn_browse_txt')
        self.btn_load_txt = root.FindName('btn_load_txt')

        self.txt_manual = root.FindName('txt_manual')
        self.btn_load_manual = root.FindName('btn_load_manual')

        self.btn_clear = root.FindName('btn_clear')
        self.btn_create = root.FindName('btn_create')
        self.btn_close = root.FindName('btn_close')

    def _wire_events(self):
        self.btn_browse_excel.Click += self.OnBrowseExcel
        self.btn_load_excel.Click += self.OnLoadExcel
        self.btn_browse_txt.Click += self.OnBrowseTxt
        self.btn_load_txt.Click += self.OnLoadTxt
        self.btn_load_manual.Click += self.OnLoadManual
        self.btn_clear.Click += self.OnClear
        self.btn_create.Click += self.OnCreate
        self.btn_close.Click += self.OnClose

    # -- ui helpers ---------------------------------------------------------

    def _refresh_header(self):
        try:
            title = doc.Title
        except Exception:
            title = "-"
        self.lbl_model_name.Text = "Model: {0}".format(title)
        self.lbl_existing.Text = "Existing user worksets: {0}".format(
            len(get_user_worksets())
        )

    def _warn(self, message):
        TaskDialog.Show(DIALOG_TITLE, message)

    def _build_rows(self, raw_names):
        """Turn raw strings into preview rows with a status."""
        existing = get_user_worksets()
        seen = set()
        rows = []
        index = 0

        for raw in raw_names:
            name = (raw or '').strip()
            if not name:
                continue
            index += 1
            key = name.lower()

            err = validate_name(name)
            if err:
                state, status = 'INVALID', "Invalid - {0}".format(err)
            elif key in existing:
                state, status = 'EXISTING', "Already exists in the model"
            elif key in seen:
                state, status = 'DUPLICATE', "Duplicate in the list"
            else:
                state, status = 'NEW', "Will be created"
                seen.add(key)

            row = Dictionary[str, object]()
            row["Index"] = str(index)
            row["Name"] = name
            row["Status"] = status
            row["State"] = state
            rows.append(row)

        return rows

    def _set_rows(self, raw_names, source_label):
        self._rows = self._build_rows(raw_names)

        self.lst_preview.Items.Clear()
        for row in self._rows:
            self.lst_preview.Items.Add(row)

        counts = {'NEW': 0, 'EXISTING': 0, 'DUPLICATE': 0, 'INVALID': 0}
        for row in self._rows:
            counts[row["State"]] += 1

        self.btn_create.IsEnabled = counts['NEW'] > 0

        if not self._rows:
            self.lbl_summary.Text = \
                "{0}: no valid name found.".format(source_label)
        else:
            self.lbl_summary.Text = (
                "{0} - {1} to create | {2} already existing | "
                "{3} duplicated | {4} invalid"
            ).format(
                source_label,
                counts['NEW'], counts['EXISTING'],
                counts['DUPLICATE'], counts['INVALID']
            )

    # -- events -------------------------------------------------------------

    def OnBrowseExcel(self, sender, args):
        dlg = OpenFileDialog()
        dlg.Title = "Select the Excel file"
        dlg.Filter = "Excel workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm"
        result = dlg.ShowDialog()
        if not result:
            return

        path = dlg.FileName
        self.cmb_sheet.Items.Clear()
        self._sheets = []

        try:
            self._sheets = get_xlsx_sheets(path)
        except Exception as ex:
            self._warn(
                "Unable to read the workbook.\n\n"
                "Only the .xlsx / .xlsm format is supported "
                "(the old .xls is not).\n\n{0}".format(ex)
            )
            return

        if not self._sheets:
            self._warn("No sheet found in the selected workbook.")
            return

        self.txt_excel_path.Text = path
        for name, _part in self._sheets:
            self.cmb_sheet.Items.Add(name)
        self.cmb_sheet.SelectedIndex = 0

    def OnLoadExcel(self, sender, args):
        path = self.txt_excel_path.Text
        if not path or not os.path.isfile(path):
            self._warn("Select an Excel file first.")
            return
        idx = self.cmb_sheet.SelectedIndex
        if idx < 0 or idx >= len(self._sheets):
            self._warn("Select a sheet first.")
            return

        sheet_name, sheet_part = self._sheets[idx]
        try:
            values = read_xlsx_first_column(
                path, sheet_part, self.chk_skip_header.IsChecked == True
            )
        except Exception as ex:
            self._warn("Error while reading the sheet.\n\n{0}".format(ex))
            return

        self._set_rows(
            values,
            "Excel [{0}]".format(sheet_name)
        )

    def OnBrowseTxt(self, sender, args):
        dlg = OpenFileDialog()
        dlg.Title = "Select the text file"
        dlg.Filter = "Text file (*.txt;*.csv)|*.txt;*.csv|All files (*.*)|*.*"
        result = dlg.ShowDialog()
        if not result:
            return
        self.txt_txt_path.Text = dlg.FileName

    def OnLoadTxt(self, sender, args):
        path = self.txt_txt_path.Text
        if not path or not os.path.isfile(path):
            self._warn("Select a text file first.")
            return
        try:
            values = read_text_file(path)
        except Exception as ex:
            self._warn("Error while reading the file.\n\n{0}".format(ex))
            return
        self._set_rows(values, "Text file")

    def OnLoadManual(self, sender, args):
        values = split_names(self.txt_manual.Text)
        self._set_rows(values, "Manual input")

    def OnClear(self, sender, args):
        self._rows = []
        self.lst_preview.Items.Clear()
        self.btn_create.IsEnabled = False
        self.lbl_summary.Text = "No list loaded."

    def OnCreate(self, sender, args):
        to_create = [r["Name"] for r in self._rows if r["State"] == 'NEW']
        if not to_create:
            self._warn("There is no new workset to create.")
            return

        created = []
        failed = []

        t = Transaction(doc, "Create Worksets")
        t.Start()
        try:
            for name in to_create:
                try:
                    if not WorksetTable.IsWorksetNameUnique(doc, name):
                        failed.append((name, "name already in use"))
                        continue
                    Workset.Create(doc, name)
                    created.append(name)
                except Exception as ex:
                    failed.append((name, str(ex)))
            t.Commit()
        except Exception as ex:
            if t.HasStarted():
                t.RollBack()
            self._warn("Transaction failed, nothing was created.\n\n{0}".format(ex))
            return

        self._refresh_header()

        # rebuild the preview so the created ones show as existing
        current = [r["Name"] for r in self._rows]
        self._set_rows(current, "After creation")

        lines = ["Worksets created: {0}".format(len(created))]
        if created:
            lines.append("")
            lines.extend("  + " + n for n in created)
        if failed:
            lines.append("")
            lines.append("Not created: {0}".format(len(failed)))
            lines.extend("  - {0} ({1})".format(n, e) for n, e in failed)
        TaskDialog.Show(DIALOG_TITLE, "\n".join(lines))

    def OnClose(self, sender, args):
        self.Close()


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

try:
    if doc.IsFamilyDocument:
        TaskDialog.Show(
            DIALOG_TITLE,
            "Worksets cannot be created in a family document.\n"
            "Open a project model and run the tool again."
        )
    elif not doc.IsWorkshared:
        TaskDialog.Show(
            DIALOG_TITLE,
            "The current model is not workshared.\n\n"
            "Enable Collaborate > Worksets first, then run this tool "
            "again to create the user worksets."
        )
    else:
        form = WorksetCreateForm()
        form.ShowDialog()

except Exception as ex:
    TaskDialog.Show(
        "{0} - Error".format(DIALOG_TITLE),
        str(ex)
    )
