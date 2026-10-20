# -*- coding: utf-8 -*-
"""
clashfinder_report_ui.py - finestra del report di Clash Finder (non modale).

Elenco espandibile (TreeView costruito in codice), diviso per coppia di categorie
("Duct Fittings x Walls") con una linea spessa sopra ogni gruppo:

    gruppo di categorie (n)
      Clash #12
        A  livello | categoria | famiglia : tipo | id
        B  livello | categoria | famiglia : tipo | id

I numeri seguono l'ordine dell'elenco (gruppi, poi ordine di ricerca) e sono gli stessi
dell'export CSV. Ogni riga porta in Tag le coppie (ElementInfo, lato) da mostrare.

Un clic su una riga, via RevitBridge:
  - colora gli elementi host della riga nella vista attiva, Set A verde e Set B rosso
    (override di elemento; quelli di prima tornano all'originale). Gli elementi linkati
    non si possono colorare: Revit non ha override sul singolo elemento di un link;
  - in una vista 3D adatta la section box agli elementi, margine 10 cm;
  - con "Isolate selected" attivo li isola temporaneamente (per un elemento linkato si
    isola l'istanza del link);
  - inquadra gli elementi.
La selezione di Revit non viene cambiata: l'evidenziazione blu coprirebbe i colori.
Colori e isolamento vengono tolti chiudendo il report.

Il report lavora sui dati raccolti al momento dell'esecuzione (ElementInfo) e rilegge
il modello solo per mostrare gli elementi. Le modifiche alla vista (colori, section
box, isolamento) stanno in una transazione per clic.
"""

import io
import os
from collections import OrderedDict

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Collections.Generic import List
from System.IO import FileStream, FileMode
from System.Windows import (Window, MessageBox, MessageBoxButton, MessageBoxImage,
                            FontWeights, FontStyles, TextWrapping, Thickness, Clipboard,
                            VerticalAlignment)
from System.Windows.Controls import TreeViewItem, TextBlock, Border, StackPanel, Orientation
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Color, Colors

from pyrevit import DB, script, forms

import clashfinder_rules as cr
import clashfinder_engine as ce
from clashfinder_bridge import RevitBridge

XAML_FILE_NAME = 'ClashFinder_report.xaml'
TITLE = "Clash Report"

SECTION_BOX_MARGIN_FT = 10.0 / 30.48     # 10 cm
ZOOM_PAD_FT = 1.0                        # ~30 cm attorno allo zoom su box

# Colori dei due insiemi: in Revit (override) e nell'elenco (quadratino accanto all'elemento).
SET_COLORS = {"a": (0, 170, 0), "b": (220, 0, 0)}
SIDE_LABELS = {"a": u"A", "b": u"B"}


def _brush(rgb):
    brush = SolidColorBrush(Color.FromRgb(rgb[0], rgb[1], rgb[2]))
    brush.Freeze()
    return brush


GRAY_BRUSH = SolidColorBrush(Colors.Gray)
GRAY_BRUSH.Freeze()
GROUP_BRUSH = _brush((0x2D, 0x5A, 0x8A))
SIDE_BRUSHES = dict((side, _brush(rgb)) for side, rgb in SET_COLORS.items())


def _warn(message):
    MessageBox.Show(message, TITLE, MessageBoxButton.OK, MessageBoxImage.Warning)


def _plural(count, word, plural=None):
    if count == 1:
        return u"1 {}".format(word)
    return u"{} {}".format(count, plural or word + u"s")


def _count_text(total, linked):
    """"120" oppure "120, 300 linked"."""
    return u"{}, {} linked".format(total, linked) if linked else cr.text(total)


def _box_text(box):
    """Box in metri, coordinate interne: "[x y z .. x y z]"."""
    if box is None:
        return u"[no box]"
    low, high = box
    return u"[{:.2f} {:.2f} {:.2f} .. {:.2f} {:.2f} {:.2f} m]".format(
        *[v * 0.3048 for v in tuple(low) + tuple(high)])


def element_row_text(info):
    """livello | categoria | famiglia : tipo | id, con il link davanti se linkato."""
    parts = [info.level or u"(no level)", info.category, info.family_type,
             u"id {}".format(info.id_value)]
    text = u"  |  ".join(parts)
    if info.is_link:
        text = u"[{}]  {}".format(info.model, text)
    return text


def search_text(info):
    return u" ".join((info.model, info.category, info.family_type, info.level,
                      cr.text(info.id_value))).lower()


def union_box(boxes):
    """Box che contiene tutti i (min, max) dati, o None."""
    boxes = [b for b in boxes if b is not None]
    if not boxes:
        return None
    low = [min(b[0][i] for b in boxes) for i in range(3)]
    high = [max(b[1][i] for b in boxes) for i in range(3)]
    return low, high


def group_key(clash):
    return (clash.a.category or u"", clash.b.category or u"")


class ClashReportForm(Window):

    def __init__(self, doc, result, display_names):
        self.doc = doc
        self.result = result
        self.display_names = display_names
        self.bridge = RevitBridge(TITLE)
        # Override originali degli elementi colorati: {id vista: {id elemento:
        # (ElementId, OverrideGraphicSettings)}}. Servono per il ripristino.
        self._colored = {}
        self._isolated_views = set()
        self._solid_fill = False
        self._group_borders = []

        # Ordine di visualizzazione: per gruppo di categorie, poi ordine di ricerca.
        # I numeri seguono quest'ordine (gli stessi nell'export).
        self.numbers = {}
        self.ordered = {}
        for test_result in result.tests:
            ordered = sorted(test_result.clashes, key=lambda c: (
                group_key(c)[0].lower(), group_key(c)[1].lower()))
            self.ordered[id(test_result)] = ordered
            for number, clash in enumerate(ordered, 1):
                self.numbers[id(clash)] = number

        self._load_xaml()
        self._fill_header()
        self._fill_tree()

    def _load_xaml(self):
        # Il report si apre da un ExternalEvent, a comando finito: get_bundle_file
        # (che usa il percorso del comando in corso) solo come ripiego.
        xaml_path = os.path.join(os.path.dirname(__file__), XAML_FILE_NAME)
        if not os.path.exists(xaml_path):
            xaml_path = script.get_bundle_file(XAML_FILE_NAME)

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

        for name in ("txt_result", "txt_result_hint", "txt_search", "btn_clear_search",
                     "tree", "btn_copy_id", "btn_expand", "btn_collapse", "btn_isolate",
                     "btn_export", "btn_close"):
            setattr(self, name, root.FindName(name))

        self.txt_search.TextChanged += lambda s, e: self._fill_tree()
        self.btn_clear_search.Click += lambda s, e: setattr(self.txt_search, "Text", u"")
        self.btn_copy_id.Click += self.OnCopyId
        self.btn_expand.Click += lambda s, e: self._expand_all(True)
        self.btn_collapse.Click += lambda s, e: self._expand_all(False)
        self.tree.SelectedItemChanged += self.OnTreeSelection
        self.tree.SizeChanged += lambda s, e: self._fit_group_borders()
        self.btn_isolate.Click += self.OnIsolateToggled
        self.btn_export.Click += self.OnExport
        self.btn_close.Click += lambda s, e: self.Close()
        self.Closed += self.OnWindowClosed

    # ----------------------------------------------------------------- contenuto

    def _fill_header(self):
        result = self.result
        sources = dict(ce.SOURCE_LABELS)
        names = u", ".join(r.name for r in result.tests)
        text = u"{}: {} ({:.1f} s).".format(
            names, _plural(result.clash_count, u"clash", u"clashes"), result.seconds)
        if result.cancelled:
            text += u" Stopped by the user: the result is partial."
        self.txt_result.Text = text

        unsupported = sum(r.unsupported for r in result.tests)
        connected = sum(r.ignored_connected for r in result.tests)
        hosted = sum(r.ignored_hosted for r in result.tests)
        scope = result.scope

        def scope_text(source, link_keys):
            parts = []
            if source != ce.SOURCE_NONE:
                parts.append(sources[source].lower())
            labels = [result.link_labels.get(key, key) for key in link_keys]
            if labels:
                parts.append(u"links {}".format(u", ".join(labels)))
            return u" + ".join(parts) or u"nothing"

        hint = u"Set A from {}, Set B from {}. Model: {}.".format(
            scope_text(scope.source_a, scope.links_a),
            scope_text(scope.source_b, scope.links_b), self.doc.Title)
        details = []
        if unsupported:
            details.append(u"{} skipped (no solid that Revit can test)".format(
                _plural(unsupported, u"element")))
        if connected:
            details.append(u"{} ignored as connected".format(_plural(connected, u"pair")))
        if hosted:
            details.append(u"{} ignored as element and host".format(_plural(hosted, u"pair")))
        if details:
            hint += u" " + u"; ".join(details) + u"."
        self.txt_result_hint.Text = hint

    def _filtered(self, clashes):
        needle = (self.txt_search.Text or u"").strip().lower()
        if not needle:
            return clashes
        return [c for c in clashes
                if needle in search_text(c.a) or needle in search_text(c.b)]

    def _fill_tree(self):
        self.tree.Items.Clear()
        self._group_borders = []
        several = len(self.result.tests) > 1
        for test_result in self.result.tests:
            clashes = self._filtered(self.ordered[id(test_result)])
            if several:
                header = u"{}: {}".format(test_result.name,
                                          _plural(len(clashes), u"clash", u"clashes"))
                parent = self._text_item(header, self._pairs(clashes), bold=True)
                parent.IsExpanded = True
                self.tree.Items.Add(parent)
                add = parent.Items.Add
            else:
                add = self.tree.Items.Add
            for note in self._test_notes(test_result):
                add(self._text_item(u"ℹ️ " + note, [], gray=True))
            if not test_result.clashes:
                add(self._text_item(u"No clashes.", [], gray=True))
            elif not clashes:
                add(self._text_item(u"No clash matches the search.", [], gray=True))
            groups = OrderedDict()
            for clash in clashes:
                groups.setdefault(group_key(clash), []).append(clash)
            for (cat_a, cat_b), members in groups.items():
                add(self._group_item(cat_a, cat_b, members))
        self._fit_group_borders()

    def _group_item(self, cat_a, cat_b, members):
        """Gruppo di categorie: linea spessa sopra, nome e numero di interferenze."""
        text = TextBlock()
        text.Text = u"{}  ×  {}   ({})".format(cat_a or u"?", cat_b or u"?", len(members))
        text.FontWeight = FontWeights.Bold
        text.Foreground = GROUP_BRUSH
        text.Margin = Thickness(0, 4, 0, 2)
        border = Border()
        border.BorderBrush = GROUP_BRUSH
        border.BorderThickness = Thickness(0, 2, 0, 0)
        border.Margin = Thickness(0, 6, 0, 0)
        border.Child = text
        self._group_borders.append(border)

        item = TreeViewItem()
        item.Header = border
        item.Tag = self._pairs(members)
        for clash in members:
            item.Items.Add(self._clash_item(clash))
        return item

    def _clash_item(self, clash):
        """Riga "Clash #n"; espandendola, i due elementi."""
        item = self._text_item(u"Clash #{}".format(self.numbers[id(clash)]),
                               [(clash.a, "a"), (clash.b, "b")], bold=True)
        item.Items.Add(self._element_item(clash.a, "a"))
        item.Items.Add(self._element_item(clash.b, "b"))
        return item

    @staticmethod
    def _element_item(info, side):
        """Elemento di una interferenza, con il quadratino del colore del suo insieme."""
        panel = StackPanel()
        panel.Orientation = Orientation.Horizontal
        mark = Border()
        mark.Width = 10
        mark.Height = 10
        mark.Background = SIDE_BRUSHES[side]
        mark.Margin = Thickness(0, 0, 6, 0)
        mark.VerticalAlignment = VerticalAlignment.Center
        label = TextBlock()
        label.Text = u"{}   {}".format(SIDE_LABELS[side], element_row_text(info))
        panel.Children.Add(mark)
        panel.Children.Add(label)
        item = TreeViewItem()
        item.Header = panel
        item.Tag = [(info, side)]
        return item

    @staticmethod
    def _pairs(clashes):
        """(info, lato) distinti delle interferenze, Set A prima di Set B."""
        found = OrderedDict()
        for clash in clashes:
            found.setdefault(clash.a.key, (clash.a, "a"))
            found.setdefault(clash.b.key, (clash.b, "b"))
        return list(found.values())

    @staticmethod
    def _text_item(text, pairs, bold=False, gray=False):
        item = TreeViewItem()
        item.Tag = pairs
        if gray:
            # Note: a capo dentro la finestra, testo intero anche nel tooltip.
            block = TextBlock()
            block.Text = text
            block.TextWrapping = TextWrapping.Wrap
            block.MaxWidth = 620
            block.Foreground = GRAY_BRUSH
            block.FontStyle = FontStyles.Italic
            item.Header = block
            item.ToolTip = text
        else:
            item.Header = text
            if bold:
                item.FontWeight = FontWeights.Bold
        return item

    def _fit_group_borders(self):
        """Linea dei gruppi larga quanto l'elenco."""
        width = max(200.0, self.tree.ActualWidth - 60.0)
        for border in self._group_borders:
            border.Width = width

    @staticmethod
    def _test_notes(test_result):
        """Note del test e diagnostica dei confronti con i link."""
        notes = [test_result.note] if test_result.note else []
        if test_result.linked_a or test_result.linked_b:
            notes.append(u"Linked elements: {} in Set A, {} in Set B; {} with overlapping "
                         u"bounding boxes, {} tested on their solids.".format(
                             test_result.linked_a, test_result.linked_b,
                             _plural(test_result.link_box_hits, u"pair"),
                             test_result.link_pairs_tested))
        if test_result.link_without_box:
            notes.append(u"{} of Set B without a bounding box.".format(
                _plural(test_result.link_without_box, u"linked element")))
        if test_result.nearest is not None and not test_result.clashes:
            gap, item_a, item_b = test_result.nearest
            notes.append(u"Closest pair with a linked element: {} cm apart (bounding boxes). "
                         u"A {} id {} {}, B {} id {} {}.".format(
                             u"{:.1f}".format(gap * 30.48),
                             item_a.source.label, item_a.value, _box_text(item_a.box()),
                             item_b.source.label, item_b.value, _box_text(item_b.box())))
        if test_result.geometry_errors:
            notes.append(u"{} in the geometric checks (first: {}).".format(
                _plural(test_result.geometry_errors, u"error"), test_result.first_error))
        return notes

    def _expand_all(self, value):
        def walk(items):
            for item in items:
                if isinstance(item, TreeViewItem):
                    item.IsExpanded = value
                    walk(item.Items)
        walk(self.tree.Items)

    def OnCopyId(self, sender, args):
        item = self.tree.SelectedItem
        pairs = list(item.Tag) if item is not None and item.Tag else []
        if not pairs:
            self.txt_result_hint.Text = u"Select a clash or an element first."
            return
        text = u", ".join(cr.text(info.id_value) for info, _ in pairs)
        try:
            Clipboard.SetText(text)
        except Exception as error:
            _warn(u"The ids could not be copied: {}".format(error))
            return
        self.txt_result_hint.Text = u"Copied: {}".format(text)

    # ---------------------------------------------------------------- selezione

    def OnTreeSelection(self, sender, args):
        item = self.tree.SelectedItem
        if item is None or not item.Tag:
            return
        pairs = list(item.Tag)
        isolate = bool(self.btn_isolate.IsChecked)
        self.bridge.run(lambda uiapp: self._show_api(uiapp, pairs, isolate))

    def OnIsolateToggled(self, sender, args):
        if self.btn_isolate.IsChecked:
            item = self.tree.SelectedItem
            if item is not None and item.Tag:
                pairs = list(item.Tag)
                self.bridge.run(lambda uiapp: self._show_api(uiapp, pairs, True))
        else:
            self.bridge.run(self._end_isolation_api)

    def _uidoc_api(self, uiapp):
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None or not self.doc.IsValidObject or not uidoc.Document.Equals(self.doc):
            _warn(u"This report belongs to the model '{}'. Activate it to navigate "
                  u"the clashes.".format(self.doc.Title if self.doc.IsValidObject else u""))
            return None
        return uidoc

    def _show_api(self, uiapp, pairs, isolate):
        uidoc = self._uidoc_api(uiapp)
        if uidoc is None:
            return
        host = []        # (elemento, lato)
        linked = []      # (istanza di link, elemento)
        boxes = []
        for info, side in pairs:
            if info.is_link:
                pair = self._linked_element(info)
                if pair is None:
                    continue
                linked.append(pair)
            else:
                element = self.doc.GetElement(info.id)
                if element is None:
                    continue
                host.append((element, side))
            boxes.append(info.box)
        if not host and not linked:
            self.txt_result_hint.Text = (u"The elements of this row no longer exist "
                                         u"(or their link is not loaded).")
            return
        box = union_box(boxes)
        view = self.doc.ActiveView
        self._update_view(view, host, linked, box, isolate)
        if linked and box is not None:
            self._zoom_to_box(uidoc, box)
        elif host:
            ids = List[DB.ElementId]()
            for element, _ in host:
                ids.Add(element.Id)
            uidoc.ShowElements(ids)

    def _linked_element(self, info):
        """(istanza di link, elemento nel link), o None se non piu' disponibile."""
        try:
            instance = self.doc.GetElement(info.link_id)
            if not isinstance(instance, DB.RevitLinkInstance):
                return None
            link_doc = instance.GetLinkDocument()
            element = link_doc.GetElement(info.id) if link_doc is not None else None
            return (instance, element) if element is not None else None
        except Exception:
            return None

    def _zoom_to_box(self, uidoc, box):
        low, high = box
        pad = ZOOM_PAD_FT
        try:
            view_id = self.doc.ActiveView.Id
            for ui_view in uidoc.GetOpenUIViews():
                if ui_view.ViewId == view_id:
                    ui_view.ZoomAndCenterRectangle(
                        DB.XYZ(low[0] - pad, low[1] - pad, low[2] - pad),
                        DB.XYZ(high[0] + pad, high[1] + pad, high[2] + pad))
                    return
        except Exception:
            pass

    # ------------------------------------------------------------- vista attiva

    def _update_view(self, view, host, linked, box, isolate):
        """Colori, section box (viste 3D) e isolamento in una sola transazione."""
        if view is None or view.IsTemplate or isinstance(view, DB.ViewSheet):
            self.txt_result_hint.Text = u"The active view is not a model view."
            return
        notes = []
        transaction = DB.Transaction(self.doc, "Clash Finder view")
        try:
            transaction.Start()
            notes.append(self._apply_colors(view, host))
            if isinstance(view, DB.View3D) and box is not None:
                notes.append(self._fit_section_box(view, box))
            if isolate:
                notes.append(self._isolate(view, host, linked))
            transaction.Commit()
        except Exception as error:
            if transaction.HasStarted() and not transaction.HasEnded():
                transaction.RollBack()
            notes.append(u"View not updated: {}".format(error))
        notes = [n for n in notes if n]
        if linked and not host:
            notes.append(u"Linked elements cannot be coloured: they are only zoomed to.")
        self.txt_result_hint.Text = u" ".join(notes)

    def _solid_fill_id(self):
        """Id del retino pieno di disegno, cercato una volta sola (None se assente)."""
        if self._solid_fill is False:
            self._solid_fill = None
            collector = DB.FilteredElementCollector(self.doc).OfClass(DB.FillPatternElement)
            try:
                for element in collector:
                    try:
                        pattern = element.GetFillPattern()
                        if pattern.IsSolidFill and \
                                pattern.Target == DB.FillPatternTarget.Drafting:
                            self._solid_fill = element.Id
                            break
                    except Exception:
                        continue
            finally:
                collector.Dispose()
        return self._solid_fill

    def _color_override(self, side):
        rgb = SET_COLORS[side]
        color = DB.Color(rgb[0], rgb[1], rgb[2])
        override = DB.OverrideGraphicSettings()
        override.SetProjectionLineColor(color)
        override.SetCutLineColor(color)
        solid = self._solid_fill_id()
        if solid is not None:
            override.SetSurfaceForegroundPatternId(solid)
            override.SetSurfaceForegroundPatternColor(color)
            override.SetCutForegroundPatternId(solid)
            override.SetCutForegroundPatternColor(color)
        return override

    def _apply_colors(self, view, host):
        """Verde Set A, rosso Set B; gli elementi colorati prima tornano all'originale."""
        view_key = cr.get_element_id_value(view.Id)
        saved = self._colored.setdefault(view_key, {})
        wanted = OrderedDict((cr.get_element_id_value(e.Id), (e, side)) for e, side in host)
        try:
            for value in [v for v in saved if v not in wanted]:
                element_id, original = saved.pop(value)
                view.SetElementOverrides(element_id, original)
            overrides = {"a": self._color_override("a"), "b": self._color_override("b")}
            for value, (element, side) in wanted.items():
                if value not in saved:
                    saved[value] = (element.Id, view.GetElementOverrides(element.Id))
                view.SetElementOverrides(element.Id, overrides[side])
        except Exception as error:
            if cr.get_element_id_value(view.ViewTemplateId) > 0:
                return (u"Colours not set: the view template of the active view controls "
                        u"Visibility/Graphics.")
            return u"Colours not set: {}".format(error)
        return None

    @staticmethod
    def _fit_section_box(view, box):
        low, high = box
        margin = SECTION_BOX_MARGIN_FT
        section = DB.BoundingBoxXYZ()
        section.Min = DB.XYZ(low[0] - margin, low[1] - margin, low[2] - margin)
        section.Max = DB.XYZ(high[0] + margin, high[1] + margin, high[2] + margin)
        try:
            view.SetSectionBox(section)
            if not view.IsSectionBoxActive:
                view.IsSectionBoxActive = True
        except Exception as error:
            return u"Section box not set: {}".format(error)
        return None

    def _isolate(self, view, host, linked):
        """Isolamento temporaneo: elementi host e istanze dei link degli elementi linkati."""
        ids = List[DB.ElementId]()
        seen = set()
        for element in [e for e, _ in host] + [i for i, _ in linked]:
            value = cr.get_element_id_value(element.Id)
            if value not in seen:
                seen.add(value)
                ids.Add(element.Id)
        try:
            if view.IsInTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate):
                view.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate)
            view.IsolateElementsTemporary(ids)
            self._isolated_views.add(cr.get_element_id_value(view.Id))
        except Exception as error:
            return u"Elements not isolated: {}".format(error)
        if linked:
            return u"Linked elements: the whole link is isolated with them."
        return None

    def _end_isolation_api(self, uiapp=None):
        """Toglie l'isolamento temporaneo dalle viste in cui e' stato messo."""
        if not self._isolated_views or not self.doc.IsValidObject:
            self._isolated_views = set()
            return
        transaction = DB.Transaction(self.doc, "Clash Finder end isolation")
        try:
            transaction.Start()
            for view in self._views(self._isolated_views):
                try:
                    if view.IsInTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate):
                        view.DisableTemporaryViewMode(
                            DB.TemporaryViewMode.TemporaryHideIsolate)
                except Exception:
                    continue
            transaction.Commit()
            self._isolated_views = set()
        except Exception as error:
            if transaction.HasStarted() and not transaction.HasEnded():
                transaction.RollBack()
            self.txt_result_hint.Text = u"Isolation not ended: {}".format(error)

    def _restore_colors_api(self, uiapp=None):
        """Rimette gli override originali degli elementi colorati, in tutte le viste."""
        if not self._colored or not self.doc.IsValidObject:
            self._colored = {}
            return
        transaction = DB.Transaction(self.doc, "Clash Finder restore colours")
        try:
            transaction.Start()
            views = dict((cr.get_element_id_value(v.Id), v) for v in self._views(self._colored))
            for view_key, saved in self._colored.items():
                view = views.get(view_key)
                if view is None:
                    continue
                for element_id, original in saved.values():
                    try:
                        view.SetElementOverrides(element_id, original)
                    except Exception:
                        continue
            transaction.Commit()
            self._colored = {}
        except Exception as error:
            if transaction.HasStarted() and not transaction.HasEnded():
                transaction.RollBack()
            self.txt_result_hint.Text = u"Colours not restored: {}".format(error)

    def _views(self, view_keys):
        wanted = set(view_keys)
        found = []
        collector = DB.FilteredElementCollector(self.doc).OfClass(DB.View)
        try:
            for view in collector:
                if cr.get_element_id_value(view.Id) in wanted:
                    found.append(view)
        finally:
            collector.Dispose()
        return found

    # ------------------------------------------------------------------- export

    def OnExport(self, sender, args):
        path = forms.save_file(file_ext='csv', default_name='ClashReport',
                               title="Export clash report")
        if not path:
            return
        headers = [u"Test", u"Clash",
                   u"A Model", u"A Id", u"A Category", u"A Family : Type", u"A Level",
                   u"B Model", u"B Id", u"B Category", u"B Family : Type", u"B Level"]
        lines = [u";".join(headers)]
        for test_result in self.result.tests:
            for clash in self.ordered[id(test_result)]:
                cells = [test_result.name, cr.text(self.numbers[id(clash)])]
                for info in (clash.a, clash.b):
                    cells += [info.model, cr.text(info.id_value), info.category,
                              info.family_type, info.level]
                lines.append(u";".join(_csv_cell(c) for c in cells))
        try:
            # utf-8-sig: Excel riconosce la codifica dal BOM.
            with io.open(path, "w", encoding="utf-8-sig") as handle:
                handle.write(u"\r\n".join(lines) + u"\r\n")
        except Exception as error:
            _warn(u"The file could not be written:\n{}\n\n{}".format(path, error))
            return
        MessageBox.Show(u"{} exported to:\n{}".format(
            _plural(self.result.clash_count, u"clash", u"clashes"), path),
            TITLE, MessageBoxButton.OK, MessageBoxImage.Information)

    # ----------------------------------------------------------------- chiusura

    def OnWindowClosed(self, sender, args):
        if not self._colored and not self._isolated_views:
            self.bridge.dispose()
            return

        # Colori e isolamento vanno tolti prima di chiudere. L'ExternalEvent non viene
        # eliminato dal suo stesso handler: resta inattivo e lo libera Revit.
        def restore(uiapp):
            self._restore_colors_api(uiapp)
            self._end_isolation_api(uiapp)
        self.bridge.run(restore)


def _csv_cell(value):
    value = cr.text(value)
    if u";" in value or u'"' in value or u"\n" in value:
        return u'"{}"'.format(value.replace(u'"', u'""'))
    return value


def show_report(doc, result, display_names, owner_handle=None):
    """Apre il report non modale e lo restituisce. Da chiamare in contesto API."""
    form = ClashReportForm(doc, result, display_names)
    if owner_handle is not None:
        try:
            from System.Windows.Interop import WindowInteropHelper
            WindowInteropHelper(form).Owner = owner_handle
        except Exception:
            pass
    form.Show()
    return form
