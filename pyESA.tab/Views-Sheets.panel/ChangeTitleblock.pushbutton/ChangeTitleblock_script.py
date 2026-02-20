# -*- coding: utf-8 -*-
"""Gestione Cartigli - Cambia il tipo di cartiglio nelle tavole selezionate."""

__title__   = "Change\nTitleblocks"
__author__  = "Andrea Patti"

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import (
    Window, Visibility, Thickness, GridLength,
    FontWeights, VerticalAlignment, HorizontalAlignment,
    WindowStartupLocation, ResizeMode, GridUnitType,
    SizeToContent, CornerRadius
)
from System.Windows.Controls import (
    Grid, RowDefinition, ColumnDefinition,
    StackPanel, ScrollViewer, Border, TextBlock, TextBox,
    Button, CheckBox, ComboBox, ComboBoxItem,
    Orientation, ScrollBarVisibility, Dock, DockPanel
)
from System.Windows.Media import (
    Brushes, SolidColorBrush, Color
)

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    ViewSheet,
    Transaction,
)
from pyrevit import revit, forms

doc   = revit.doc
uidoc = revit.uidoc

# ==============================================================================
#  Colori
# ==============================================================================

CLR_PRIMARY    = SolidColorBrush(Color.FromArgb(255, 63, 81, 181))
CLR_PRIMARY_DK = SolidColorBrush(Color.FromArgb(255, 48, 63, 159))
CLR_SECONDARY  = SolidColorBrush(Color.FromArgb(255, 117, 117, 117))
CLR_SEC_DK     = SolidColorBrush(Color.FromArgb(255, 97, 97, 97))
CLR_BG         = SolidColorBrush(Color.FromArgb(255, 245, 245, 245))
CLR_BORDER     = SolidColorBrush(Color.FromArgb(255, 204, 204, 204))

# ==============================================================================
#  Funzioni di utilita Revit
# ==============================================================================

def get_all_sheets():
    sheets = (
        FilteredElementCollector(doc)
        .OfClass(ViewSheet)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    return sorted(sheets, key=lambda s: s.SheetNumber)


def get_titleblock_on_sheet(sheet):
    collector = (
        FilteredElementCollector(doc, sheet.Id)
        .OfCategory(BuiltInCategory.OST_TitleBlocks)
        .WhereElementIsNotElementType()
    )
    tbs = list(collector)
    return tbs[0] if tbs else None


def get_all_titleblock_types():
    symbols = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_TitleBlocks)
        .WhereElementIsElementType()
        .ToElements()
    )
    result = {}
    for sym in symbols:
        fam_name = sym.FamilyName or ""
        type_param = sym.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        type_str = type_param.AsString() if type_param else "Senza nome"
        full_name = "{} : {}".format(fam_name, type_str)
        result[full_name] = sym.Id
    return result


def titleblock_display_name(tb_instance):
    if tb_instance is None:
        return "<nessun cartiglio>"
    sym = doc.GetElement(tb_instance.GetTypeId())
    if sym is None:
        return "<tipo sconosciuto>"
    fam_name = sym.FamilyName or ""
    type_param = sym.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
    type_str = type_param.AsString() if type_param else "Senza nome"
    return "{} : {}".format(fam_name, type_str)


# ==============================================================================
#  Helper per costruire i controlli
# ==============================================================================

def make_button(text, bg, bg_hover, width=None):
    btn = Button()
    btn.Content    = text
    btn.FontSize   = 13
    btn.Padding    = Thickness(16, 8, 16, 8)
    btn.Background = bg
    btn.Foreground = Brushes.White
    btn.BorderThickness = Thickness(0)
    if width:
        btn.Width = width

    def on_enter(s, e):
        s.Background = bg_hover
    def on_leave(s, e):
        s.Background = bg

    btn.MouseEnter += on_enter
    btn.MouseLeave += on_leave
    return btn


def make_small_button(text, bg, bg_hover):
    btn = make_button(text, bg, bg_hover)
    btn.FontSize = 11.5
    btn.Padding  = Thickness(10, 5, 10, 5)
    return btn


# ==============================================================================
#  Classe della finestra
# ==============================================================================

class TitleBlockManagerWindow(Window):

    def __init__(self):
        self.Title  = "Gestione Cartigli"
        self.Width  = 950
        self.Height = 620
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResizeWithGrip
        self.Background = CLR_BG

        # -- Dati --------------------------------------------------------------
        self.all_sheets    = get_all_sheets()
        self.tb_types      = get_all_titleblock_types()
        self.tb_type_names = sorted(self.tb_types.keys())
        self.sheet_map     = {}   # int_id -> (checkbox, sheet, tb_instance)
        self.combo_map     = {}   # int_id -> ComboBox

        # -- Costruisci UI -----------------------------------------------------
        self._build_ui()

    def _build_ui(self):
        # Griglia principale
        main = Grid()
        main.Margin = Thickness(16)
        for h in ["Auto", "Auto", "*", "Auto", "Auto"]:
            rd = RowDefinition()
            if h == "*":
                rd.Height = GridLength(1, GridUnitType.Star)
            else:
                rd.Height = GridLength.Auto
            main.RowDefinitions.Add(rd)

        # ---- Riga 0: Titolo fase 1 ----
        title1 = TextBlock()
        title1.Text       = "Fase 1 - Seleziona le tavole"
        title1.FontSize   = 16
        title1.FontWeight = FontWeights.SemiBold
        title1.Foreground = CLR_PRIMARY
        title1.Margin     = Thickness(0, 0, 0, 8)
        Grid.SetRow(title1, 0)
        main.Children.Add(title1)

        # ---- Riga 1: Filtro + pulsanti selezione ----
        filter_grid = Grid()
        filter_grid.Margin = Thickness(0, 0, 0, 6)
        for w in ["Auto", "*", "Auto", "Auto"]:
            cd = ColumnDefinition()
            if w == "*":
                cd.Width = GridLength(1, GridUnitType.Star)
            else:
                cd.Width = GridLength.Auto
            filter_grid.ColumnDefinitions.Add(cd)

        lbl_filter = TextBlock()
        lbl_filter.Text = "Filtra:"
        lbl_filter.VerticalAlignment = VerticalAlignment.Center
        lbl_filter.Margin   = Thickness(0, 0, 6, 0)
        lbl_filter.FontSize = 12.5
        Grid.SetColumn(lbl_filter, 0)
        filter_grid.Children.Add(lbl_filter)

        self.txt_filter = TextBox()
        self.txt_filter.Padding  = Thickness(4)
        self.txt_filter.FontSize = 12.5
        self.txt_filter.TextChanged += self.filter_changed
        Grid.SetColumn(self.txt_filter, 1)
        filter_grid.Children.Add(self.txt_filter)

        btn_sel = make_small_button("Seleziona tutto", CLR_SECONDARY, CLR_SEC_DK)
        btn_sel.Margin = Thickness(8, 0, 0, 0)
        btn_sel.Click += self.select_all_click
        Grid.SetColumn(btn_sel, 2)
        filter_grid.Children.Add(btn_sel)

        btn_desel = make_small_button("Deseleziona tutto", CLR_SECONDARY, CLR_SEC_DK)
        btn_desel.Margin = Thickness(4, 0, 0, 0)
        btn_desel.Click += self.deselect_all_click
        Grid.SetColumn(btn_desel, 3)
        filter_grid.Children.Add(btn_desel)

        Grid.SetRow(filter_grid, 1)
        main.Children.Add(filter_grid)

        # ---- Riga 2: Lista tavole ----
        border_sheets = Border()
        border_sheets.BorderBrush     = CLR_BORDER
        border_sheets.BorderThickness = Thickness(1)
        border_sheets.CornerRadius    = CornerRadius(4)
        border_sheets.Background      = Brushes.White
        border_sheets.Margin          = Thickness(0, 0, 0, 8)

        sv_sheets = ScrollViewer()
        sv_sheets.VerticalScrollBarVisibility = ScrollBarVisibility.Auto

        self.stack_sheets = StackPanel()
        self.stack_sheets.Margin = Thickness(8, 6, 8, 6)
        sv_sheets.Content = self.stack_sheets
        border_sheets.Child = sv_sheets

        Grid.SetRow(border_sheets, 2)
        main.Children.Add(border_sheets)

        # Popola le tavole
        self._populate_sheets()

        # ---- Riga 3: Pannello fase 2 (nascosto) ----
        self.phase2_border = Border()
        self.phase2_border.Visibility      = Visibility.Collapsed
        self.phase2_border.BorderBrush     = CLR_PRIMARY
        self.phase2_border.BorderThickness = Thickness(1)
        self.phase2_border.CornerRadius    = CornerRadius(4)
        self.phase2_border.Background      = Brushes.White
        self.phase2_border.Margin          = Thickness(0, 0, 0, 8)
        self.phase2_border.MaxHeight       = 280

        phase2_grid = Grid()
        for h in ["Auto", "Auto", "*"]:
            rd = RowDefinition()
            if h == "*":
                rd.Height = GridLength(1, GridUnitType.Star)
            else:
                rd.Height = GridLength.Auto
            phase2_grid.RowDefinitions.Add(rd)

        title2 = TextBlock()
        title2.Text       = "Fase 2 - Assegna il nuovo cartiglio"
        title2.FontSize   = 16
        title2.FontWeight = FontWeights.SemiBold
        title2.Foreground = CLR_PRIMARY
        title2.Margin     = Thickness(12, 10, 12, 4)
        Grid.SetRow(title2, 0)
        phase2_grid.Children.Add(title2)

        # Riga "Applica a tutte"
        apply_all_grid = Grid()
        apply_all_grid.Margin = Thickness(12, 2, 12, 6)
        for w in ["Auto", "*", "Auto"]:
            cd = ColumnDefinition()
            if w == "*":
                cd.Width = GridLength(1, GridUnitType.Star)
            else:
                cd.Width = GridLength.Auto
            apply_all_grid.ColumnDefinitions.Add(cd)

        lbl_aa = TextBlock()
        lbl_aa.Text = "Applica a tutte:"
        lbl_aa.VerticalAlignment = VerticalAlignment.Center
        lbl_aa.FontSize = 12
        lbl_aa.Margin   = Thickness(0, 0, 8, 0)
        Grid.SetColumn(lbl_aa, 0)
        apply_all_grid.Children.Add(lbl_aa)

        self.cmb_apply_all = ComboBox()
        self.cmb_apply_all.FontSize = 12
        self.cmb_apply_all.Padding  = Thickness(4, 3, 4, 3)
        Grid.SetColumn(self.cmb_apply_all, 1)
        apply_all_grid.Children.Add(self.cmb_apply_all)

        btn_aa = make_small_button("Applica", CLR_SECONDARY, CLR_SEC_DK)
        btn_aa.Margin = Thickness(8, 0, 0, 0)
        btn_aa.Click += self.apply_all_click
        Grid.SetColumn(btn_aa, 2)
        apply_all_grid.Children.Add(btn_aa)

        Grid.SetRow(apply_all_grid, 1)
        phase2_grid.Children.Add(apply_all_grid)

        # Griglia assegnamenti
        sv_assign = ScrollViewer()
        sv_assign.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        sv_assign.Margin = Thickness(0, 0, 0, 8)

        self.grid_assignments = Grid()
        self.grid_assignments.Margin = Thickness(12, 0, 12, 4)
        for w in [200, 250, 30]:
            cd = ColumnDefinition()
            cd.Width = GridLength(w)
            self.grid_assignments.ColumnDefinitions.Add(cd)
        cd_star = ColumnDefinition()
        cd_star.Width = GridLength(1, GridUnitType.Star)
        self.grid_assignments.ColumnDefinitions.Add(cd_star)

        sv_assign.Content = self.grid_assignments
        Grid.SetRow(sv_assign, 2)
        phase2_grid.Children.Add(sv_assign)

        self.phase2_border.Child = phase2_grid
        Grid.SetRow(self.phase2_border, 3)
        main.Children.Add(self.phase2_border)

        # Popola combo "applica a tutte"
        self._populate_apply_all_combo()

        # ---- Riga 4: Pulsanti azione ----
        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right

        self.btn_next = make_button("Avanti >>", CLR_PRIMARY, CLR_PRIMARY_DK, 120)
        self.btn_next.Click += self.next_click
        btn_panel.Children.Add(self.btn_next)

        self.btn_apply = make_button("Applica modifiche", CLR_PRIMARY, CLR_PRIMARY_DK, 180)
        self.btn_apply.Margin     = Thickness(8, 0, 0, 0)
        self.btn_apply.Visibility = Visibility.Collapsed
        self.btn_apply.Click += self.apply_click
        btn_panel.Children.Add(self.btn_apply)

        btn_cancel = make_button("Annulla", CLR_SECONDARY, CLR_SEC_DK, 100)
        btn_cancel.Margin = Thickness(8, 0, 0, 0)
        btn_cancel.Click += self.cancel_click
        btn_panel.Children.Add(btn_cancel)

        Grid.SetRow(btn_panel, 4)
        main.Children.Add(btn_panel)

        self.Content = main

    # -- Popolazione lista tavole ----------------------------------------------
    def _populate_sheets(self):
        self.stack_sheets.Children.Clear()
        self.sheet_map.clear()
        for sheet in self.all_sheets:
            tb = get_titleblock_on_sheet(sheet)
            tb_name = titleblock_display_name(tb)
            label = "{} - {}   [{}]".format(
                sheet.SheetNumber, sheet.Name, tb_name
            )
            cb = CheckBox()
            cb.Content  = label
            cb.Tag      = sheet.Id.IntegerValue
            cb.FontSize = 12.5
            cb.Margin   = Thickness(0, 2, 0, 2)
            self.stack_sheets.Children.Add(cb)
            self.sheet_map[sheet.Id.IntegerValue] = (cb, sheet, tb)

    def _populate_apply_all_combo(self):
        self.cmb_apply_all.Items.Clear()
        for name in self.tb_type_names:
            item = ComboBoxItem()
            item.Content = name
            self.cmb_apply_all.Items.Add(item)

    # -- Eventi Fase 1 ---------------------------------------------------------
    def filter_changed(self, sender, args):
        text = self.txt_filter.Text.lower()
        for child in self.stack_sheets.Children:
            if isinstance(child, CheckBox):
                visible = text in child.Content.lower()
                child.Visibility = (
                    Visibility.Visible if visible
                    else Visibility.Collapsed
                )

    def select_all_click(self, sender, args):
        for child in self.stack_sheets.Children:
            if isinstance(child, CheckBox) \
               and child.Visibility == Visibility.Visible:
                child.IsChecked = True

    def deselect_all_click(self, sender, args):
        for child in self.stack_sheets.Children:
            if isinstance(child, CheckBox):
                child.IsChecked = False

    # -- Passaggio alla Fase 2 -------------------------------------------------
    def next_click(self, sender, args):
        selected = []
        for child in self.stack_sheets.Children:
            if isinstance(child, CheckBox) and child.IsChecked:
                key = child.Tag
                if key in self.sheet_map:
                    selected.append(self.sheet_map[key])
        if not selected:
            forms.alert("Seleziona almeno una tavola.", title="Attenzione")
            return
        self._build_assignment_grid(selected)
        self.phase2_border.Visibility = Visibility.Visible
        self.btn_next.Visibility      = Visibility.Collapsed
        self.btn_apply.Visibility     = Visibility.Visible

    def _build_assignment_grid(self, selected):
        grid = self.grid_assignments
        grid.RowDefinitions.Clear()
        to_remove = [c for c in grid.Children]
        for c in to_remove:
            grid.Children.Remove(c)

        # Intestazioni
        rd = RowDefinition()
        rd.Height = GridLength.Auto
        grid.RowDefinitions.Add(rd)
        headers = ["Tavola", "Cartiglio attuale", "", "Nuovo cartiglio"]
        for col, text in enumerate(headers):
            tb = TextBlock()
            tb.Text       = text
            tb.FontWeight = FontWeights.SemiBold
            tb.FontSize   = 12
            tb.Margin     = Thickness(0, 0, 0, 6)
            Grid.SetRow(tb, 0)
            Grid.SetColumn(tb, col)
            grid.Children.Add(tb)

        self.combo_map.clear()

        for idx, (cb, sheet, tb_inst) in enumerate(selected):
            row = idx + 1
            rd = RowDefinition()
            rd.Height = GridLength.Auto
            grid.RowDefinitions.Add(rd)

            # Col 0 - Tavola
            lbl = TextBlock()
            lbl.Text              = "{} - {}".format(sheet.SheetNumber, sheet.Name)
            lbl.FontSize          = 12
            lbl.Margin            = Thickness(0, 3, 8, 3)
            lbl.VerticalAlignment = VerticalAlignment.Center
            Grid.SetRow(lbl, row)
            Grid.SetColumn(lbl, 0)
            grid.Children.Add(lbl)

            # Col 1 - Cartiglio attuale
            current_name = titleblock_display_name(tb_inst)
            lbl2 = TextBlock()
            lbl2.Text              = current_name
            lbl2.FontSize          = 12
            lbl2.Margin            = Thickness(0, 3, 8, 3)
            lbl2.Foreground        = Brushes.Gray
            lbl2.VerticalAlignment = VerticalAlignment.Center
            Grid.SetRow(lbl2, row)
            Grid.SetColumn(lbl2, 1)
            grid.Children.Add(lbl2)

            # Col 2 - Freccia
            arrow = TextBlock()
            arrow.Text                  = ">>"
            arrow.FontSize              = 13
            arrow.Margin                = Thickness(0, 3, 0, 3)
            arrow.VerticalAlignment     = VerticalAlignment.Center
            arrow.HorizontalAlignment   = HorizontalAlignment.Center
            Grid.SetRow(arrow, row)
            Grid.SetColumn(arrow, 2)
            grid.Children.Add(arrow)

            # Col 3 - ComboBox
            cmb = ComboBox()
            cmb.FontSize = 12
            cmb.Margin   = Thickness(0, 2, 0, 2)
            cmb.Padding  = Thickness(4, 3, 4, 3)

            no_change = ComboBoxItem()
            no_change.Content = "-- Nessuna modifica --"
            cmb.Items.Add(no_change)

            for name in self.tb_type_names:
                item = ComboBoxItem()
                item.Content = name
                if name == current_name:
                    item.FontWeight = FontWeights.Bold
                cmb.Items.Add(item)

            cmb.SelectedIndex = 0
            cmb.Tag = sheet.Id.IntegerValue
            Grid.SetRow(cmb, row)
            Grid.SetColumn(cmb, 3)
            grid.Children.Add(cmb)
            self.combo_map[sheet.Id.IntegerValue] = cmb

    # -- Applica a tutte -------------------------------------------------------
    def apply_all_click(self, sender, args):
        sel = self.cmb_apply_all.SelectedItem
        if sel is None:
            return
        target_name = sel.Content
        for sid, cmb in self.combo_map.items():
            for i in range(cmb.Items.Count):
                if cmb.Items[i].Content == target_name:
                    cmb.SelectedIndex = i
                    break

    # -- Applicazione finale ---------------------------------------------------
    def apply_click(self, sender, args):
        changes = []
        for sid, cmb in self.combo_map.items():
            if cmb.SelectedIndex <= 0:
                continue
            new_name = cmb.SelectedItem.Content
            new_type_id = self.tb_types.get(new_name)
            if new_type_id is None:
                continue
            cb, sheet, tb_inst = self.sheet_map[sid]
            if tb_inst is None:
                continue
            if tb_inst.GetTypeId() == new_type_id:
                continue
            changes.append((sheet, tb_inst, new_type_id))

        if not changes:
            forms.alert(
                "Nessuna modifica da applicare.\n"
                "Verifica di aver scelto un nuovo cartiglio "
                "diverso da quello attuale.",
                title="Nessuna modifica"
            )
            return

        n = len(changes)
        msg = "Stai per modificare il cartiglio di {} tavol{}.\nContinuare?".format(
            n, "a" if n == 1 else "e"
        )
        if not forms.alert(msg, yes=True, no=True, title="Conferma"):
            return

        t = Transaction(doc, "Cambia Cartigli")
        t.Start()
        try:
            for sheet, tb_inst, new_type_id in changes:
                tb_inst.ChangeTypeId(new_type_id)
            t.Commit()
            forms.alert(
                "Operazione completata!\n"
                "{} cartigl{} modificat{}.".format(
                    n,
                    "io" if n == 1 else "i",
                    "o" if n == 1 else "i",
                ),
                title="Successo"
            )
            self.Close()
        except Exception as ex:
            t.RollBack()
            forms.alert(
                "Errore durante la modifica:\n{}".format(str(ex)),
                title="Errore"
            )

    def cancel_click(self, sender, args):
        self.Close()


# ==============================================================================
#  Avvio
# ==============================================================================

sheets = get_all_sheets()
if not sheets:
    forms.alert("Non sono presenti tavole nel progetto.", title="Attenzione")
else:
    tb_types = get_all_titleblock_types()
    if not tb_types:
        forms.alert(
            "Non sono presenti tipi di cartiglio nel progetto.",
            title="Attenzione"
        )
    else:
        window = TitleBlockManagerWindow()
        window.ShowDialog()