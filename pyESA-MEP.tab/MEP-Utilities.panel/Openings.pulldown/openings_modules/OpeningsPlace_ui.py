# -*- coding: utf-8 -*-
"""Finestre di Openings Place&Check: cosa selezionare e quali link leggere (quando non c'e' selezione), anteprima
prima di creare i fori, esito (anche "nessun foro da aggiungere"). Nessun accesso a Revit. Importato solo da
place_main(); nessun documento o output a livello di modulo (Rocket mode)."""
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import FontWeights, GridLength, TextWrapping, Thickness, VerticalAlignment, Visibility
from System.Windows.Controls import CheckBox, ColumnDefinition, Grid, TextBlock
from System.Windows.Markup import XamlReader

PICK_HOSTS, PICK_MEPS, PICK_AREA = "hosts", "meps", "area"
# valori dell'anteprima: chiave del dict passato a ConfirmWindow -> TextBlock della finestra
PREVIEW_FIELDS = (("mode", "ModeValue"), ("crossings", "CrossingsValue"), ("openings", "OpeningsValue"),
                  ("existing", "ExistingValue"), ("inside", "InsideValue"), ("extended", "ExtendedValue"))


def _load(xaml_path):
    return XamlReader.Parse(File.ReadAllText(xaml_path))


def _checkbox(tag, text, checked):
    """CheckBox con testo a capo e il valore in Tag (liste multi-selezione dello stile di casa)."""
    box = CheckBox()
    label = TextBlock()
    label.Text = text
    label.TextWrapping = TextWrapping.Wrap
    box.Content = label
    box.Tag = tag
    box.IsChecked = checked
    return box


class PickWindow(object):
    """Cosa selezionare (muri, solai e strutture, elementi MEP, oppure un'area) e quali link leggere. links:
    [(chiave, percorso)] dei link caricati; skipped: chiavi da non spuntare; phase: fase letta (solo testo). show()
    ritorna PICK_HOSTS, PICK_MEPS, PICK_AREA o None se annullata; dopo, self.skipped = link non spuntati."""

    def __init__(self, xaml_path, last_pick, settings_text, links=(), skipped=(), phase=""):
        self.window = _load(xaml_path)
        find = self.window.FindName
        self.radios = {PICK_HOSTS: find("HostsRadio"), PICK_MEPS: find("MepsRadio"), PICK_AREA: find("AreaRadio")}
        self.radios.get(last_pick, self.radios[PICK_HOSTS]).IsChecked = True
        find("SettingsInfo").Text = settings_text
        find("PhaseText").Text = u"Phase read: {} (phase of the active view; each link in the phase with the same " \
                                 u"name).".format(phase or "none")
        self.links = self._fill(find, "Links", [(key, u"{}  ({})".format(key, path or "path unknown"),
                                                 key not in skipped) for key, path in links])
        find("SelectAllButton").Click += lambda sender, args: self._tick_all(self.links, True)
        find("DeselectAllButton").Click += lambda sender, args: self._tick_all(self.links, False)
        find("OkButton").Click += self._on_ok
        self.result = None
        self.skipped = set(skipped)

    @staticmethod
    def _fill(find, name, rows):
        """Riempie la lista <name>List con [(tag, testo, spuntato)]; senza righe mostra No<name>Text."""
        boxes = []
        for tag, text, checked in rows:
            box = _checkbox(tag, text, checked)
            find(name + "List").Items.Add(box)
            boxes.append(box)
        if not boxes:
            for control in (name + "List", name + "Buttons"):
                find(control).Visibility = Visibility.Collapsed
            find("No{}Text".format(name)).Visibility = Visibility.Visible
        return boxes

    @staticmethod
    def _tick_all(boxes, value):
        for box in boxes:
            box.IsChecked = value

    def _on_ok(self, sender, args):
        self.result = [key for key, radio in self.radios.items() if radio.IsChecked][0]
        self.skipped = set(box.Tag for box in self.links if not box.IsChecked)
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result


def _value_row(label, value):
    """Riga etichetta / valore (colonna fissa + colonna *), come le righe dei campi dello stile di casa."""
    grid = Grid()
    grid.Margin = Thickness(0, 3, 0, 3)
    first = ColumnDefinition()
    first.Width = GridLength(210)
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


class ResultWindow(object):
    """Esito di Place&Check, al posto del pannello di output di pyRevit: titolo, messaggio, riepilogo [(etichetta,
    valore)], avvisi, fori creati [testo] (ids: i loro Id), dettagli [testo]. show() ritorna True se si e' chiesto di
    selezionare i fori creati (la selezione la fa chi chiama, dopo la chiusura)."""

    def __init__(self, xaml_path, header, message, summary, openings=(), ids=(), warning="", details=()):
        self.window = _load(xaml_path)
        find = self.window.FindName
        find("HeaderText").Text = header
        find("MessageText").Text = message
        for label, value in summary:
            find("SummaryPanel").Children.Add(_value_row(label, value))
        if warning:
            find("WarningText").Text = warning
        else:
            find("WarningBorder").Visibility = Visibility.Collapsed
        for line in openings:
            row = TextBlock()
            row.Text = line
            row.TextWrapping = TextWrapping.Wrap
            find("OpeningsList").Items.Add(row)
        if not openings:
            find("OpeningsBorder").Visibility = Visibility.Collapsed
        find("DetailsText").Text = u"\n".join(details)
        select = find("SelectButton")
        if ids:
            select.Click += self._on_select
        else:
            select.Visibility = Visibility.Collapsed
        self.result = False

    def _on_select(self, sender, args):
        self.result = True
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result


class ConfirmWindow(object):
    """Anteprima dei fori da creare: valori di PREVIEW_FIELDS e problemi di lettura. show() ritorna True per
    creare."""

    def __init__(self, xaml_path, values, problems):
        self.window = _load(xaml_path)
        find = self.window.FindName
        for key, control in PREVIEW_FIELDS:
            find(control).Text = u"{}".format(values.get(key, ""))
        if problems:
            find("ProblemsText").Text = u"\n".join(problems)
        else:
            find("ProblemsBorder").Visibility = Visibility.Collapsed
        find("OkButton").Click += self._on_ok
        self.result = False

    def _on_ok(self, sender, args):
        self.result = True
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
