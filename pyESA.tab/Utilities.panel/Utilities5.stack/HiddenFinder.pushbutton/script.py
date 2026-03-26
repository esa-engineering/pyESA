import os
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Transaction,
    ElementId,
    Element,
    TemporaryViewMode
)
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List, Dictionary

from System.Windows import Window
from System.Windows.Markup import XamlReader
from System.IO import StreamReader

# pyRevit entry points
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
active_view = doc.ActiveView


def get_id_value(eid):
    """Return the integer value of an ElementId (works in Revit 2022-2026+)."""
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def get_element_info(elem):
    """Return (category, family, type) strings for an element."""
    cat_name = "Unknown"
    family_name = "Unknown"
    type_name = "Unknown"

    try:
        if elem.Category:
            cat_name = elem.Category.Name
    except Exception:
        pass

    try:
        type_id = elem.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            elem_type = doc.GetElement(type_id)
            if elem_type is not None:
                try:
                    fn = elem_type.FamilyName
                    if fn:
                        family_name = fn
                except Exception:
                    pass
                try:
                    tn = Element.Name.GetValue(elem_type)
                    if tn:
                        type_name = tn
                except Exception:
                    try:
                        tn = elem_type.Name
                        if tn:
                            type_name = tn
                    except Exception:
                        pass
    except Exception:
        pass

    return (cat_name, family_name, type_name)


def collect_hidden_elements():
    """Collect per-element hidden items in the active view."""
    result = []
    view = active_view

    collector = FilteredElementCollector(doc) \
        .WhereElementIsNotElementType()

    for elem in collector:
        cat = elem.Category
        if cat is None:
            continue
        try:
            if not elem.IsHidden(view):
                continue
            if view.GetCategoryHidden(cat.Id):
                continue

            cat_name, family_name, type_name = get_element_info(elem)
            id_str = str(get_id_value(elem.Id))

            d = Dictionary[str, object]()
            d["Category"] = cat_name
            d["Family"] = family_name
            d["Type"] = type_name
            d["Id"] = id_str
            d["ElementId"] = elem.Id
            result.append(d)
        except Exception:
            continue

    result.sort(key=lambda x: (x["Category"], int(x["Id"])))
    return result


def enable_reveal_hidden():
    """Enable Reveal Hidden Elements mode on the active view."""
    if active_view.IsInTemporaryViewMode(
        TemporaryViewMode.RevealHiddenElements
    ):
        return
    t = Transaction(doc, "Enable Reveal Hidden")
    t.Start()
    try:
        active_view.EnableRevealHiddenMode()
        t.Commit()
    except Exception:
        if t.HasStarted():
            t.RollBack()


def disable_reveal_hidden():
    """Disable Reveal Hidden Elements mode if active."""
    if not active_view.IsInTemporaryViewMode(
        TemporaryViewMode.RevealHiddenElements
    ):
        return
    t = Transaction(doc, "Disable Reveal Hidden")
    t.Start()
    try:
        active_view.DisableTemporaryViewMode(
            TemporaryViewMode.RevealHiddenElements
        )
        t.Commit()
    except Exception:
        if t.HasStarted():
            t.RollBack()


class HiddenElementsForm(Window):
    def __init__(self):
        self._hidden_items = []
        self._updating = False
        self._load_xaml()

    def _load_xaml(self):
        script_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(script_dir, 'HiddenElementsForm.xaml')

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
        self._populate_list()

    def _find_controls(self, root):
        self.lbl_view_name = root.FindName('lbl_view_name')
        self.lbl_count = root.FindName('lbl_count')
        self.lst_elements = root.FindName('lst_elements')
        self.btn_unhide = root.FindName('btn_unhide')
        self.btn_close = root.FindName('btn_close')

    def _wire_events(self):
        self.lst_elements.SelectionChanged += self.OnSelectionChanged
        self.btn_unhide.Click += self.OnUnhideSelected
        self.btn_close.Click += self.OnClose
        self.Closing += self.OnWindowClosing

    def _populate_list(self):
        """Collect hidden elements and fill ListView."""
        self._updating = True
        try:
            self._hidden_items = collect_hidden_elements()

            self.lst_elements.Items.Clear()
            for item in self._hidden_items:
                self.lst_elements.Items.Add(item)

            count = len(self._hidden_items)
            self.lbl_count.Text = "Elementi nascosti: {0}".format(count)
            self.lbl_view_name.Text = "Vista corrente: {0}".format(
                active_view.Name
            )
        finally:
            self._updating = False

    def OnSelectionChanged(self, sender, args):
        """Select corresponding elements in the Revit model."""
        if self._updating:
            return
        selected = self.lst_elements.SelectedItems
        if selected.Count == 0:
            return

        ids = List[ElementId]()
        for item in selected:
            ids.Add(item["ElementId"])

        try:
            uidoc.Selection.SetElementIds(ids)
        except Exception:
            pass

    def OnUnhideSelected(self, sender, args):
        """Unhide all selected elements in the current view."""
        selected = self.lst_elements.SelectedItems
        if selected.Count == 0:
            return

        ids = List[ElementId]()
        for item in selected:
            ids.Add(item["ElementId"])

        t = Transaction(doc, "Unhide Elements")
        t.Start()
        try:
            active_view.UnhideElements(ids)
            t.Commit()
        except Exception as ex:
            if t.HasStarted():
                t.RollBack()
            TaskDialog.Show(
                "Hidden Elements Inspector",
                "Errore unhide:\n{0}".format(str(ex))
            )
            return

        # Refresh list after unhiding
        self._populate_list()

    def OnWindowClosing(self, sender, args):
        """Disable Reveal Hidden when window closes."""
        disable_reveal_hidden()

    def OnClose(self, sender, args):
        self.Close()


def is_graphical_view(view):
    from Autodesk.Revit.DB import ViewType
    non_graphical = [
        ViewType.Schedule,
        ViewType.ColumnSchedule,
        ViewType.PanelSchedule,
        ViewType.DrawingSheet,
        ViewType.Report,
        ViewType.ProjectBrowser,
        ViewType.SystemBrowser,
        ViewType.Undefined,
        ViewType.Internal,
    ]
    return view.ViewType not in non_graphical


# --- Main entry point ---
try:
    if active_view is None or not is_graphical_view(active_view):
        TaskDialog.Show(
            "Hidden Elements Inspector",
            "La vista attiva non e' una vista grafica.\n"
            "Attiva una vista di pianta, sezione, 3D o simile."
        )
    else:
        # Collect first to decide if reveal is needed
        test_items = collect_hidden_elements()

        # Enable Reveal Hidden if there are hidden elements
        if len(test_items) > 0:
            enable_reveal_hidden()

        # Show modal form (all API calls work directly)
        form = HiddenElementsForm()
        form.ShowDialog()

except Exception as ex:
    TaskDialog.Show(
        "Hidden Elements Inspector - Errore",
        str(ex)
    )
