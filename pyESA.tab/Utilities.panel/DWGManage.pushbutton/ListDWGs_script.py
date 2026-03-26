# -*- coding: UTF-8 -*-
"""Lists all imported DWG instances in a WPF DataGrid with selection and deletion."""
import os
import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ImportInstance,
    BuiltInParameter,
    WorksharingUtils,
    Transaction,
    ElementId,
    UnitUtils,
    UnitTypeId,
)
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection

from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxResult
from System.Windows.Markup import XamlReader
from System.IO import StreamReader

# pyRevit entry points
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


class DwgRow(object):
    """Data model for a single DWG import instance row."""
    def __init__(self, dwg_name, created_by, element_id, workset_name,
                 hosted_level, offset_z, view_specific, view_name):
        self.DwgName = dwg_name
        self.CreatedBy = created_by
        self.ElementId = element_id
        self.WorksetName = workset_name
        self.HostedLevel = hosted_level
        self.OffsetZ = offset_z
        self.ViewSpecific = view_specific
        self.ViewName = view_name


def collect_dwg_rows():
    """Collect all imported DWG instances and return a list of DwgRow objects."""
    dwgs = (
        FilteredElementCollector(doc)
        .OfClass(ImportInstance)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    rows = []
    for dwg in dwgs:
        if dwg.IsLinked:
            continue

        dwg_name = dwg.Parameter[BuiltInParameter.IMPORT_SYMBOL_NAME].AsString()

        try:
            created_by = WorksharingUtils.GetWorksharingTooltipInfo(
                doc, dwg.Id
            ).Creator
        except Exception:
            created_by = "N/A"

        element_id = dwg.Id.IntegerValue

        # Workset
        workset_name = "N/A"
        try:
            wt = doc.GetWorksetTable()
            if wt is not None:
                ws = wt.GetWorkset(dwg.WorksetId)
                if ws is not None:
                    workset_name = ws.Name
        except Exception:
            pass

        # Hosted level
        level = doc.GetElement(dwg.LevelId)
        hosted_level = level.Name if level else "N/A"

        # Offset Z
        offset_z = "N/A"
        try:
            transform = dwg.GetTransform()
            if transform:
                z_feet = transform.Origin.Z
                z_meters = UnitUtils.ConvertFromInternalUnits(
                    z_feet, UnitTypeId.Meters
                )
                offset_z = "{:.3f} m ({:.3f} ft)".format(z_meters, z_feet)
        except Exception:
            pass

        # View Specific / View Name
        if workset_name.startswith("View"):
            view_specific = "YES"
            view_name = workset_name[5:] if len(workset_name) > 5 else workset_name
        else:
            view_specific = "NO"
            view_name = ""

        rows.append(DwgRow(
            dwg_name, created_by, element_id, workset_name,
            hosted_level, offset_z, view_specific, view_name
        ))

    return rows


class DwgManagerForm(Window):
    """WPF window for managing imported DWGs."""

    def __init__(self):
        self._collection = ObservableCollection[object]()
        self._load_xaml()

    def _load_xaml(self):
        script_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(script_dir, "Form.xaml")

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

        # Find controls
        self.dg_dwgs = root.FindName("dg_dwgs")
        self.btn_delete = root.FindName("btn_delete")
        self.btn_close = root.FindName("btn_close")
        self.lbl_status = root.FindName("lbl_status")

        # Populate data
        self._load_data()

        # Bind
        self.dg_dwgs.ItemsSource = self._collection
        self.dg_dwgs.SelectionChanged += self._on_selection_changed
        self.btn_delete.Click += self._on_delete_click
        self.btn_close.Click += self._on_close_click

    def _load_data(self):
        """Load DWG data into the ObservableCollection."""
        rows = collect_dwg_rows()
        self._collection.Clear()
        for row in rows:
            self._collection.Add(row)

        count = len(rows)
        if count == 0:
            self.lbl_status.Text = "Nessun DWG importato trovato nel modello"
        else:
            self.lbl_status.Text = "Trovati {} DWG importati".format(count)

    def _on_selection_changed(self, sender, args):
        """Sync DataGrid selection to Revit selection."""
        selected_items = list(self.dg_dwgs.SelectedItems)
        selected_count = len(selected_items)

        self.btn_delete.IsEnabled = selected_count > 0

        if selected_count > 0:
            ids = List[ElementId]()
            for row in selected_items:
                ids.Add(ElementId(int(row.ElementId)))
            try:
                uidoc.Selection.SetElementIds(ids)
            except Exception:
                pass

    def _on_delete_click(self, sender, args):
        """Delete selected DWG elements after confirmation."""
        selected_items = list(self.dg_dwgs.SelectedItems)
        if not selected_items:
            return

        count = len(selected_items)
        result = MessageBox.Show(
            "Eliminare {} DWG selezionat{}?\n\n"
            "Questa operazione non puo' essere annullata.".format(
                count, "o" if count == 1 else "i"
            ),
            "Conferma eliminazione",
            MessageBoxButton.YesNo
        )

        if result != MessageBoxResult.Yes:
            return

        ids_to_delete = [
            ElementId(int(row.ElementId)) for row in selected_items
        ]

        t = Transaction(doc, "Elimina DWG importati")
        t.Start()
        try:
            for eid in ids_to_delete:
                doc.Delete(eid)
            t.Commit()
        except Exception as ex:
            if t.HasStarted():
                t.RollBack()
            MessageBox.Show(
                "Errore durante l'eliminazione:\n{}".format(ex),
                "Errore"
            )
            return

        # Remove from collection
        for row in selected_items:
            self._collection.Remove(row)

        remaining = self._collection.Count
        self.lbl_status.Text = "Eliminati {} DWG. Rimangono {} DWG importati".format(
            count, remaining
        )

    def _on_close_click(self, sender, args):
        self.Close()


# --- Entry point ---
try:
    if doc is None:
        TaskDialog.Show("DWG Manage", "Nessun documento aperto.")
    else:
        form = DwgManagerForm()
        form.ShowDialog()
except Exception as ex:
    TaskDialog.Show("DWG Manage - Errore", str(ex))
