# -*- coding: utf-8 -*-
"""Finestra di dialogo di MEPAlignToCeiling (MEPAlignToCeiling_form.xaml).

Nessun riferimento a documento, vista o output a livello di modulo: in
Rocket mode questo modulo resta in sys.modules fra un comando e l'altro.
Il modulo non conosce la Revit API: riceve etichette e conteggi gia' pronti
e restituisce un dizionario con le scelte.
"""
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox, Thickness, Visibility
from System.Windows.Controls import CheckBox
from System.Windows.Markup import XamlReader


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


class AlignToCeilingWindow(object):
    """Categorie, tolleranza (mm) e opzioni."""

    def __init__(self, xaml_path, info):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.result = None
        self._info = info
        self._checks = []

        find = self.window.FindName
        self.category_list = find("CategoryList")
        self.category_count = find("CategoryCount")
        self.search_box = find("SearchBox")
        self.tolerance_box = find("ToleranceBox")
        self.tolerance_hint = find("ToleranceHint")
        self.include_links = find("IncludeLinks")
        self.skip_connected = find("SkipConnected")

        find("SelectionText").Text = info["selection_text"]
        find("SelectionDetail").Text = info["detail_text"]
        find("BuzzerText").Text = info["buzzer_text"]
        # La nota sui rilevatori serve solo se ce n'e' qualcuno nella selezione.
        if not info["buzzers"]:
            find("BuzzerHint").Visibility = Visibility.Collapsed

        for name, label, count, flagged, checked in info["categories"]:
            check = CheckBox()
            check.Content = u"{}  ({} selected, {} for ceiling)".format(label, count, flagged)
            check.Tag = name
            check.IsChecked = checked
            check.Margin = Thickness(2, 2, 2, 2)
            check.Checked += self._on_category_toggled
            check.Unchecked += self._on_category_toggled
            self._checks.append(check)
        self._populate_list()

        self.tolerance_box.Text = u"{:g}".format(info["tolerance_mm"])
        self.include_links.IsChecked = info["include_links"]
        self.skip_connected.IsChecked = info["skip_connected"]

        self.search_box.TextChanged += self._on_search
        self.tolerance_box.TextChanged += self._on_tolerance_changed
        find("ClearSearchButton").Click += self._on_clear_search
        find("SelectAllButton").Click += self._on_select_all
        find("DeselectAllButton").Click += self._on_deselect_all
        find("OkButton").Click += self._on_ok

        self._refresh_count()
        self._refresh_tolerance_hint()

    # --- elenco delle categorie ----------------------------------------

    def _passes_search(self, item):
        text = (self.search_box.Text or u"").strip().lower()
        if not text:
            return True
        return text in (u"{}".format(item.Content)).lower()

    def _visible_checks(self):
        return [c for c in self._checks if self._passes_search(c)]

    def _populate_list(self):
        """Ripopola la lista con le caselle che passano la ricerca.

        Le CheckBox sono sempre le stesse istanze, quindi la spunta di una
        categoria nascosta dalla ricerca si conserva. Items.Filter non
        funziona con elementi UI aggiunti direttamente a Items.
        """
        self.category_list.Items.Clear()
        for check in self._visible_checks():
            self.category_list.Items.Add(check)

    def _refresh_count(self):
        selected = [c for c in self._checks if c.IsChecked]
        elements = 0
        flagged = 0
        for name, _label, count, ceiling, _checked in self._info["categories"]:
            if any(c.Tag == name for c in selected):
                elements += count
                flagged += ceiling
        self.category_count.Text = u"{} of {} categories, {} elements ({} for ceiling)".format(
            len(selected), len(self._checks), elements, flagged)

    def _on_category_toggled(self, sender, args):
        self._refresh_count()

    def _on_search(self, sender, args):
        self._populate_list()

    def _on_clear_search(self, sender, args):
        self.search_box.Text = u""

    def _on_select_all(self, sender, args):
        for check in self._visible_checks():
            check.IsChecked = True
        self._refresh_count()

    def _on_deselect_all(self, sender, args):
        for check in self._visible_checks():
            check.IsChecked = False
        self._refresh_count()

    # --- tolleranza ----------------------------------------------------

    def _refresh_tolerance_hint(self):
        # Solo una riga di stato: il metodo della retta di analisi e' nel README.
        ok, value = parse_number(self.tolerance_box.Text)
        if ok and value > 0:
            self.tolerance_hint.Text = u"Now: search line {:g} mm long, centred on the insertion point.".format(
                value * self._info["line_factor"])
        else:
            self.tolerance_hint.Text = u"Enter a tolerance greater than zero."

    def _on_tolerance_changed(self, sender, args):
        self._refresh_tolerance_hint()

    # --- conferma ------------------------------------------------------

    def _validate(self):
        """(ok, messaggio, valori)."""
        ok, tolerance = parse_number(self.tolerance_box.Text)
        if not ok:
            return False, u"The tolerance must be a number (mm).", None
        if tolerance <= 0:
            return False, u"The tolerance must be greater than zero.", None
        if tolerance > self._info["max_tolerance_mm"]:
            return False, u"The tolerance looks out of scale (over {:g} mm).".format(
                self._info["max_tolerance_mm"]), None
        categories = [c.Tag for c in self._checks if c.IsChecked]
        if not categories:
            return False, u"Select at least one category to align.", None
        return True, u"", {
            "categories": categories,
            "tolerance_mm": tolerance,
            "include_links": bool(self.include_links.IsChecked),
            "skip_connected": bool(self.skip_connected.IsChecked),
        }

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, u"MEP Align to Ceiling")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result


def show_dialog(xaml_path, info):
    """Mostra la finestra; ritorna il dizionario delle scelte o None."""
    return AlignToCeilingWindow(xaml_path, info).show()
