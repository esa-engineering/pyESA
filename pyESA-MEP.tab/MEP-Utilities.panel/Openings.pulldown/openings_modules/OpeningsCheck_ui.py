# -*- coding: utf-8 -*-
"""Finestre dei report di Openings (pulsanti New Report e Last Report, e Place con elementi gia' associati):
ambito (New Report), riepilogo con le due schede (Text Report, Interactive Report), Text Report con l'export Excel,
Interactive Report. Nessun accesso a Revit: openings_report passa testi e Id interi, e un oggetto actions (call:
select, prepare_recreate, recreate, move, delete, refresh, report_data) che fa il lavoro sul modello.
Una sola copia, in openings_modules."""
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")

from Microsoft.Win32 import OpenFileDialog, SaveFileDialog
from System.Diagnostics import Process
from System.IO import Directory, File, Path
from System.Windows import CornerRadius, FontWeights, GridLength, MessageBox, MessageBoxButton, MessageBoxResult, \
    Thickness, TextWrapping, VerticalAlignment, Visibility
from System.Windows.Controls import Border, CheckBox, ColumnDefinition, Grid, Orientation, Separator, StackPanel, \
    TextBlock, TreeViewItem
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Markup import XamlReader
from System.Windows.Media import BrushConverter
from System.Windows.Forms import FolderBrowserDialog, DialogResult as FormsResult
from System.Windows.Input import Keyboard, ModifierKeys

import OpeningsReport_xlsx

SCOPE_ALL, SCOPE_MEP, SCOPE_ARC, SCOPE_AREA = "all", "mep", "arc", "area"
SHOW_TEXT, SHOW_WINDOW = "text", "window"
THIRD_EXPORT, THIRD_IMPORT = "export", "import"  # terzo pulsante del riepilogo (ThirdAction)
TITLE = "Openings Report"


def _load(xaml_path):
    return XamlReader.Parse(File.ReadAllText(xaml_path))


def _brush(color):
    return BrushConverter().ConvertFromString(color)


def _value_row(label, value):
    """Riga etichetta / valore (colonna fissa + colonna *), come le righe dei campi dello stile di casa."""
    grid = Grid()
    grid.Margin = Thickness(0, 4, 0, 4)
    first = ColumnDefinition()
    first.Width = GridLength(200)
    grid.ColumnDefinitions.Add(first)
    grid.ColumnDefinitions.Add(ColumnDefinition())
    name = TextBlock()
    name.Text = label
    name.VerticalAlignment = VerticalAlignment.Center
    text = TextBlock()
    text.Text = u"{}".format(value)
    text.FontWeight = FontWeights.Bold
    text.TextWrapping = TextWrapping.Wrap
    Grid.SetColumn(text, 1)
    grid.Children.Add(name)
    grid.Children.Add(text)
    return grid


def _fill_totals(panel, rows):
    for row in rows:  # None = linea di separazione
        if row is None:
            line = Separator()
            line.Margin = Thickness(0, 4, 0, 4)
            panel.Children.Add(line)
        else:
            panel.Children.Add(_value_row(row[0], row[1]))


def _set_owner(window, owner):
    if owner:
        # sopra la finestra di Revit (si riduce a icona con lei), ma non la blocca
        WindowInteropHelper(window).Owner = owner


class ScopeWindow(object):
    """Ambito del New Report (tutto il modello, selezione MEP, selezione ARC, area). show() ritorna l'ambito o None
    se annullata."""

    def __init__(self, xaml_path, last_scope, selected_meps, selected_hosts):
        self.window = _load(xaml_path)
        find = self.window.FindName
        self.radios = {SCOPE_ALL: find("AllRadio"), SCOPE_MEP: find("MepRadio"), SCOPE_ARC: find("ArcRadio"),
                       SCOPE_AREA: find("AreaRadio")}
        if selected_meps:
            last_scope = SCOPE_MEP
        elif selected_hosts:
            last_scope = SCOPE_ARC
        self.radios.get(last_scope, self.radios[SCOPE_ALL]).IsChecked = True
        find("MepHint").Text = ("Uses the {} MEP elements and openings already selected.".format(selected_meps)
                                if selected_meps else "You select pipes, ducts, cable trays and conduits after "
                                                      "the button.") + " Only the openings associated with them."
        find("ArcHint").Text = ("Uses the {} walls, floors and structural elements already selected.".format(
            selected_hosts) if selected_hosts else "You select walls, floors and structural elements (also in "
                                                   "links) after the button.") + " Only the openings on them."
        find("OkButton").Click += self._on_ok
        self.result = None

    def _on_ok(self, sender, args):
        self.result = [key for key, radio in self.radios.items() if radio.IsChecked][0]
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result


class ThirdAction(object):
    """Terzo pulsante del riepilogo. kind: THIRD_EXPORT (si sceglie una cartella) o THIRD_IMPORT (si sceglie un file
    .json); hint: testo sotto il titolo; folder: cartella proposta; run(percorso) -> {"message", e per l'Import
    "title", "rows", "warning" del report caricato}; un errore va in un messaggio."""

    def __init__(self, kind, hint, folder, run):
        self.kind, self.hint, self.folder, self.run = kind, hint, folder, run


class SummaryWindow(object):
    """Totali (del New Report appena fatto, dell'ultimo report o del report importato), due schede (Text Report e
    Interactive Report) e il terzo pulsante (ThirdAction: Export / Import). enabled: le due schede sono attive
    (spente se non c'e' ancora un report da mostrare; l'Import le accende). imported: report importato (testo
    dell'Interactive Report in sola lettura). show() ritorna SHOW_TEXT, SHOW_WINDOW o None (chiusa)."""

    def __init__(self, xaml_path, title, rows, warning, third=None, enabled=True, imported=False):
        self.window = _load(xaml_path)
        find = self.window.FindName
        self.find = find
        if imported:  # l'avviso di un report importato non riguarda parametri non scritti
            find("WarningTitle").Text = "CHECK"
        self._set_report(title, rows, warning)
        find("TextButton").Click += self._on_text
        find("OkButton").Click += self._on_window
        self._enable(enabled)
        if imported:
            find("WindowHint").Text = ("Show and select the elements of this model and the openings that cite them "
                                       "(through the linked Openings model). Read-only.")
        self.third = third
        if third is None:
            find("ThirdButton").Visibility = Visibility.Collapsed
        else:
            export = third.kind == THIRD_EXPORT
            find("ExportPanel").Visibility = Visibility.Visible if export else Visibility.Collapsed
            find("ImportPanel").Visibility = Visibility.Collapsed if export else Visibility.Visible
            find("ExportHint" if export else "ImportHint").Text = third.hint
            find("ThirdButton").Click += self._on_third
        self.result = None

    def _set_report(self, title, rows, warning):
        find = self.find
        find("SummaryTitle").Text = title.upper()
        find("TotalsPanel").Children.Clear()
        _fill_totals(find("TotalsPanel"), rows)
        if warning:  # riquadro di avviso solo se c'e' qualcosa da dire
            find("WarningText").Text = warning
            find("WarningBorder").Visibility = Visibility.Visible
        else:
            find("WarningBorder").Visibility = Visibility.Collapsed

    def _enable(self, enabled):
        for name in ("TextButton", "OkButton"):
            self.find(name).IsEnabled = enabled

    def _on_third(self, sender, args):
        third = self.third
        if third.kind == THIRD_EXPORT:
            dialog = FolderBrowserDialog()
            dialog.Description = "Folder for the report read by the MEP / ARC models"
            if third.folder and Directory.Exists(third.folder):
                dialog.SelectedPath = third.folder
            if dialog.ShowDialog() != FormsResult.OK:
                return
            path = dialog.SelectedPath
        else:
            dialog = OpenFileDialog()
            dialog.Title = "Openings report exported from the Openings model"
            dialog.Filter = "Openings report (*.json)|*.json"
            if third.folder and Directory.Exists(third.folder):
                dialog.InitialDirectory = third.folder
            if not dialog.ShowDialog():
                return
            path = dialog.FileName
        try:
            outcome = third.run(path)
        except Exception as ex:
            MessageBox.Show(u"{} failed:\n{}".format("Export" if third.kind == THIRD_EXPORT else "Import", ex), TITLE)
            return
        third.folder = path if third.kind == THIRD_EXPORT else Path.GetDirectoryName(path)
        status = self.find("ThirdStatus")
        status.Text = outcome.get("message", u"")
        status.Visibility = Visibility.Visible
        if "rows" in outcome:  # Import: il riepilogo mostra il report caricato
            self._set_report(outcome["title"], outcome["rows"], outcome.get("warning", ""))
            self._enable(True)

    def _on_text(self, sender, args):
        self.result = SHOW_TEXT
        self.window.DialogResult = True

    def _on_window(self, sender, args):
        self.result = SHOW_WINDOW
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result


class TextReportWindow(object):
    """Text Report non modale: riepilogo, una sezione per stato (ogni foro con i suoi elementi e il loro stato; i
    Valid in forma compatta), dettagli; Export to Excel (OpeningsReport_xlsx). data: vedi report_data() nello
    script. Non tocca il modello: niente ExternalEvent."""

    MAX_ITEMS = 300  # per sezione; oltre, il resto e' nell'export Excel

    def __init__(self, xaml_path, data, owner=None):
        self.window = _load(xaml_path)
        self.data = data
        find = self.window.FindName
        _set_owner(self.window, owner)
        find("HeaderTitle").Text = u"OPENINGS {}".format(data["title"].upper())
        find("SummaryDate").Text = u"  -  {}".format(data["date"]) if data.get("date") else u""
        _fill_totals(find("TotalsPanel"), data["summary"])
        if data["warning"]:
            find("WarningText").Text = data["warning"]
        else:
            find("WarningBorder").Visibility = Visibility.Collapsed
        panel = find("SectionsPanel")
        for section in data["sections"]:
            panel.Children.Add(self._section(section))
        find("DetailsText").Text = u"\n".join(data["details"] + data["log"])
        self.status = find("StatusText")
        find("ExportButton").Click += self._on_export
        find("CloseButton").Click += self._on_close

    # -- contenuto ---------------------------------------------------------------------------------------------
    def _text(self, text, style=None, bold=False, margin=None):
        block = TextBlock()
        if style:
            block.Style = self.window.FindResource(style)
        block.Text = text
        block.TextWrapping = TextWrapping.Wrap
        if bold:
            block.FontWeight = FontWeights.Bold
        if margin is not None:
            block.Margin = margin
        return block

    def _columns(self, values, widths, indent=0):
        """Riga a colonne: larghezze fisse per tutte tranne l'ultima (*)."""
        grid = Grid()
        grid.Margin = Thickness(indent, 1, 0, 1)
        for width in widths:
            column = ColumnDefinition()
            column.Width = GridLength(width)
            grid.ColumnDefinitions.Add(column)
        grid.ColumnDefinitions.Add(ColumnDefinition())
        for index, value in enumerate(values):
            block = TextBlock()
            block.Text = value
            block.FontSize = 11
            block.TextWrapping = TextWrapping.Wrap
            Grid.SetColumn(block, index)
            grid.Children.Add(block)
        return grid

    def _section(self, section):
        border = Border()
        border.BorderBrush = _brush("#CCCCCC")
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(5)
        border.Padding = Thickness(10)
        border.Margin = Thickness(0, 0, 0, 10)
        stack = StackPanel()
        border.Child = stack
        items = section["items"]
        stack.Children.Add(self._text(u"{} ({})".format(section["title"].upper(), len(items)), "SectionHeader"))
        if not items:
            stack.Children.Add(self._text(section.get("empty", u"None."), "HintText"))
            return border
        for item in items[:self.MAX_ITEMS]:
            stack.Children.Add(self._compact(item) if item.get("compact") else self._item(item))
        if len(items) > self.MAX_ITEMS:
            stack.Children.Add(self._text(u"... and {} more: see the Excel export.".format(
                len(items) - self.MAX_ITEMS), "HintText", margin=Thickness(0, 6, 0, 0)))
        return border

    def _item(self, item):
        box = StackPanel()
        box.Margin = Thickness(0, 6, 0, 6)
        box.Children.Add(self._text(item["title"], bold=True))
        for note in item["notes"]:
            box.Children.Add(self._text(note, "HintText", margin=Thickness(12, 0, 0, 0)))
        for row in item["rows"]:
            box.Children.Add(self._columns(row, (90, 110), indent=12))
        return box

    def _compact(self, item):
        elements = u", ".join(row[0] for row in item["rows"]) or u"-"
        return self._columns((item["title"], elements), (430,))

    # -- azioni ------------------------------------------------------------------------------------------------
    def _on_export(self, sender, args):
        dialog = SaveFileDialog()
        dialog.Title = "Export the report to Excel"
        dialog.Filter = "Excel workbook (*.xlsx)|*.xlsx"
        dialog.FileName = u"{} - Openings {}.xlsx".format(self.data.get("model", u"Model"), self.data["title"])
        if not dialog.ShowDialog(self.window):
            return
        path = dialog.FileName
        try:
            OpeningsReport_xlsx.write_report(path, self.data)
        except OpeningsReport_xlsx.FileLockedError:
            MessageBox.Show(u"The file is open in Excel: close it and export again.", TITLE)
            return
        except Exception as ex:
            MessageBox.Show(u"The export failed:\n{}".format(ex), TITLE)
            return
        self.status.Text = u"Exported to {}".format(path)
        if MessageBox.Show(u"Report exported:\n{}\n\nOpen it now?".format(path), TITLE,
                           MessageBoxButton.YesNo) == MessageBoxResult.Yes:
            try:
                Process.Start(path)
            except Exception:
                pass

    def _on_close(self, sender, args):
        self.window.Close()

    def show(self):
        """Non modale: ritorna subito, Revit resta usabile."""
        self.window.Show()


class OpeningEntry(object):
    """Riga di un foro nel report: key = Id del foro (testo), elements = [(testo, [Id da selezionare])]."""

    def __init__(self, key, label, detail, select_ids, elements, can_recreate=False, can_move=False, checkable=True,
                 can_delete=False):
        self.key, self.label, self.detail = key, label, detail
        self.select_ids, self.elements = list(select_ids), list(elements)
        self.can_recreate, self.can_move, self.checkable = can_recreate, can_move, checkable
        self.can_delete = can_delete
        self.search = u" ".join([label] + [text for text, _ in elements]).lower()


class ReportWindow(object):
    """Report interattivo NON modale: albero stato -> foro -> elementi, groups = [(titolo, [OpeningEntry])]. Le spunte
    stanno sui fori; Recreate e Move lavorano sui fori spuntati, Show / Select sul nodo scelto.
    Revit resta usabile mentre la finestra e' aperta: ogni azione sul modello passa da actions.call(nome, argomenti,
    fatto), che la esegue nel contesto API (ExternalEvent) e poi chiama fatto(risultato, errore) sullo stesso thread
    della finestra. Mentre una richiesta e' in coda i pulsanti delle azioni sono spenti."""

    EXPANDED = ("Missing openings", "Not Associated", "To Be Refreshed")  # gruppi aperti all'inizio
    BUTTONS = ("ShowButton", "SelectCloseButton", "RecreateButton", "MoveButton", "DeleteButton", "TextButton",
               "RefreshButton")

    def __init__(self, xaml_path, groups, actions, owner=None, text_xaml=None, readonly=False):
        self.window = _load(xaml_path)
        find = self.window.FindName
        if readonly:
            # report importato in un modello MEP / ARC: niente spunte, azioni sui fori e Refresh (solo Show / Select)
            for name in ("ActionsBorder", "TickButtons", "RefreshButton"):
                find(name).Visibility = Visibility.Collapsed
            find("StatusText").Text = ("Report imported from the Openings model: read-only. Show / Select picks the "
                                       "elements of this model and, if the Openings model is linked here, the "
                                       "openings.")
        self.actions = actions
        self.owner, self.text_xaml = owner, text_xaml  # per aprire il Text Report da qui
        _set_owner(self.window, owner)
        self.tree = find("OpeningsTree")
        self.search = find("SearchBox")
        self.detail = find("DetailText")
        self.status = find("StatusText")
        self.buttons = [find(name) for name in self.BUTTONS]
        self.checked = set()
        self.boxes = {}  # chiave del foro -> casella nell'albero attuale
        self.anchor = None  # chiave dell'ultima casella cliccata: Shift + clic spunta da li' a quella cliccata
        self.expanded = set(self.EXPANDED)
        self.search.TextChanged += self._on_search
        self.tree.SelectedItemChanged += self._on_selected
        find("SelectAllButton").Click += self._on_select_all
        find("DeselectAllButton").Click += self._on_deselect_all
        find("ShowButton").Click += self._on_show
        find("SelectCloseButton").Click += self._on_select_close
        find("RecreateButton").Click += self._on_recreate
        find("MoveButton").Click += self._on_move
        find("DeleteButton").Click += self._on_delete
        find("TextButton").Click += self._on_text
        find("RefreshButton").Click += self._on_refresh
        find("CloseButton").Click += self._on_close  # IsCancel chiude da solo solo le finestre modali
        self._set_groups(groups)

    # -- albero ------------------------------------------------------------------------------------------------
    def _set_groups(self, groups):
        self.groups = groups
        keys = set(e.key for _, entries in groups for e in entries)
        self.checked &= keys
        self._fill()

    def _visible(self):
        text = (self.search.Text or u"").strip().lower()
        return [(title, [e for e in entries if not text or text in e.search]) for title, entries in self.groups]

    def _fill(self):
        for item in self.tree.Items:  # si tiene lo stato aperto/chiuso dei gruppi
            title = item.Tag[1]
            if item.IsExpanded:
                self.expanded.add(title)
            else:
                self.expanded.discard(title)
        self.tree.Items.Clear()
        self.boxes = {}
        for title, entries in self._visible():
            group = TreeViewItem()
            header = TextBlock()
            header.Text = u"{} ({})".format(title, len(entries))
            header.FontWeight = FontWeights.Bold
            group.Header = header
            group.Tag = ("group", title, entries)
            group.IsExpanded = title in self.expanded
            for entry in entries:
                group.Items.Add(self._entry_item(entry))
            self.tree.Items.Add(group)
        self.detail.Text = "Pick an opening or an element in the list."

    def _entry_item(self, entry):
        item = TreeViewItem()
        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        if entry.checkable:
            box = CheckBox()
            box.Tag = entry
            box.IsChecked = entry.key in self.checked
            box.Margin = Thickness(0, 0, 6, 0)
            box.VerticalAlignment = VerticalAlignment.Center
            box.Checked += self._on_check
            box.Unchecked += self._on_uncheck
            box.Click += self._on_box_click
            self.boxes[entry.key] = box
            row.Children.Add(box)
        label = TextBlock()
        label.Text = entry.label
        label.VerticalAlignment = VerticalAlignment.Center
        row.Children.Add(label)
        item.Header = row
        item.Tag = ("entry", entry.label, entry)
        for text, ids in entry.elements:
            child = TreeViewItem()
            child.Header = text
            child.Tag = ("element", text, ids)
            item.Items.Add(child)
        return item

    def _on_check(self, sender, args):
        self.checked.add(sender.Tag.key)

    def _on_uncheck(self, sender, args):
        self.checked.discard(sender.Tag.key)

    def _shown_keys(self):
        """Chiavi dei fori con la casella, nell'ordine in cui si vedono: solo i gruppi aperti (e la ricerca)."""
        keys = []
        for group in self.tree.Items:
            if group.IsExpanded:
                keys.extend(item.Tag[2].key for item in group.Items if item.Tag[2].key in self.boxes)
        return keys

    def _on_box_click(self, sender, args):
        """Shift + clic: tutti i fori visibili tra l'ultima casella cliccata e questa prendono lo stato di questa (come
        in Esplora risorse). Click arriva dopo che la casella ha cambiato stato."""
        key = sender.Tag.key
        if (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift and self.anchor is not None:
            shown = self._shown_keys()
            if self.anchor in shown and key in shown:
                first, last = sorted((shown.index(self.anchor), shown.index(key)))
                state = bool(sender.IsChecked)
                for other in shown[first:last + 1]:
                    self.boxes[other].IsChecked = state  # Checked / Unchecked aggiornano self.checked
                self.status.Text = u"{} {} openings.".format("Ticked" if state else "Unticked", last - first + 1)
        self.anchor = key

    def _on_search(self, sender, args):
        self._fill()

    def _on_selected(self, sender, args):
        node = self.tree.SelectedItem
        if node is None:
            return
        kind = node.Tag[0]
        if kind == "group":
            self.detail.Text = u"{}: {} items.".format(node.Tag[1], len(node.Tag[2]))
        elif kind == "entry":
            self.detail.Text = node.Tag[2].detail
        else:
            self.detail.Text = node.Tag[1]

    def _on_select_all(self, sender, args):
        for _, entries in self._visible():
            self.checked.update(e.key for e in entries if e.checkable)
        self._fill()

    def _on_deselect_all(self, sender, args):
        self.checked.clear()
        self._fill()

    # -- richieste a Revit -------------------------------------------------------------------------------------
    def _busy(self, text):
        for button in self.buttons:
            button.IsEnabled = False
        self.status.Text = text

    def _idle(self):
        for button in self.buttons:
            button.IsEnabled = True

    def _call(self, name, args, done, text="Working in Revit..."):
        """Accoda una richiesta; done(risultato) solo se riesce, un errore va in un messaggio."""
        self._busy(text)

        def finished(result, error):
            self._idle()
            if error:
                self.status.Text = u"Not done: {}".format(error)
                MessageBox.Show(u"The action was not done:\n{}".format(error), TITLE)
                return
            done(result)

        self.actions.call(name, args, finished)

    def _show_groups(self, outcome):
        """outcome: (messaggio, gruppi ricalcolati)."""
        message, groups = outcome
        self.status.Text = message
        self.checked.clear()
        self._set_groups(groups)

    # -- azioni ------------------------------------------------------------------------------------------------
    def _selected_ids(self):
        node = self.tree.SelectedItem
        if node is None:
            return []
        kind = node.Tag[0]
        if kind == "group":
            return [i for e in node.Tag[2] for i in e.select_ids]
        if kind == "entry":
            return node.Tag[2].select_ids
        return node.Tag[2]

    def _on_show(self, sender, args):
        ids = self._selected_ids()
        if not ids:
            MessageBox.Show("Pick an opening, an element or a status group in the list first.", TITLE)
            return
        self._call("select", (ids, True), lambda result: setattr(self.status, "Text", result))

    def _on_select_close(self, sender, args):
        ids = self._selected_ids()
        if not ids:
            MessageBox.Show("Pick an opening, an element or a status group in the list first.", TITLE)
            return
        self._call("select", (ids, False), lambda result: self.window.Close())

    def _ticked(self, flag):
        return [e for _, entries in self.groups for e in entries if e.key in self.checked and getattr(e, flag)]

    def _on_recreate(self, sender, args):
        entries = self._ticked("can_recreate")
        if not entries:
            MessageBox.Show("Tick the openings to recreate: Not Associated / To Be "
                            "Refreshed openings with elements still crossing them.", TITLE)
            return
        keys = [e.key for e in entries]
        # prima si ricontrollano i fori spuntati (il modello puo' essere cambiato), poi anteprima e conferma
        self._call("prepare_recreate", (keys,), lambda result: self._confirm_recreate(keys, result),
                   "Checking the ticked openings again first...")

    def _confirm_recreate(self, keys, result):
        preview, groups = result
        self.checked |= set(keys)  # le spunte restano sui fori ancora presenti (_set_groups toglie gli altri)
        self._set_groups(groups)
        message = (u"Recreate {} openings?\n\n{}\n\nEvery new opening gets a new Id and status Valid; the Ids on the "
                   u"MEP elements are updated.".format(len(keys), preview or u"(nothing to recreate)"))
        if MessageBox.Show(message, TITLE, MessageBoxButton.YesNo) != MessageBoxResult.Yes:
            self.status.Text = "Cancelled; the ticked openings were checked again."
            return
        self._call("recreate", (keys,), self._show_groups, "Creating / recreating...")

    def _on_move(self, sender, args):
        entries = self._ticked("can_move")
        if not entries:
            MessageBox.Show("Tick the openings to move: openings with elements partly or fully outside them.", TITLE)
            return
        message = (u"Move {} openings onto their elements? They are checked again first; sizes and associations "
                   u"do not change.".format(len(entries)))
        if MessageBox.Show(message, TITLE, MessageBoxButton.YesNo) != MessageBoxResult.Yes:
            return
        self._call("move", ([e.key for e in entries],), self._show_groups, "Moving...")

    def _on_delete(self, sender, args):
        entries = self._ticked("can_delete")
        if not entries:
            MessageBox.Show("Tick the Not Associated openings to delete.", TITLE)
            return
        keys = [e.key for e in entries]
        others = len([key for key in self.checked if key not in keys])
        message = (u"Delete {} Not Associated openings?\n\n{}\n\nTheir elements stay in the report. Undo in Revit "
                   u"brings them back.".format(len(keys), u", ".join(keys)))
        if others:
            message += u"\n\n{} other ticked items are not Not Associated openings and are left as they are.".format(
                others)
        if MessageBox.Show(message, TITLE, MessageBoxButton.YesNo) != MessageBoxResult.Yes:
            return
        self._call("delete", (keys,), self._show_groups, "Deleting...")

    def _on_refresh(self, sender, args):
        self._call("refresh", (), self._show_groups, "New Report on the scope...")

    def _on_text(self, sender, args):
        self._call("report_data", (), self._open_text, "Preparing the text report...")

    def _open_text(self, data):
        TextReportWindow(self.text_xaml, data, self.owner).show()
        self.status.Text = "Text Report opened (same list, read-only, with Export to Excel)."

    def _on_close(self, sender, args):
        self.window.Close()

    def show(self):
        """Non modale: ritorna subito, Revit resta usabile."""
        self.window.Show()
