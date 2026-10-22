# -*- coding: utf-8 -*-
"""
RNWFlowCalculation_Export_ui.py - finestra di impostazione dell'esportazione,
aperta dal pulsante Export della finestra dei risultati.

L'utente sceglie il formato (Excel o PDF, e per il PDF il formato della
pagina) e le colonne da esportare per ciascuna delle tre schede: la tabella dei
sistemi di General, Sections e Fixtures. Le colonne sono quelle dei cataloghi
di RNWFlowCalculation_Report; le caselle si creano in codice, con la chiave
della colonna in Tag. Le schede si esportano sempre tutte e tre, quindi ogni
tabella deve avere almeno una colonna.

Le scelte si salvano nella config (sezione ESA_RNWFlowCalculation) come chiavi
separate da virgole; chiavi sconosciute (colonne tolte in una versione
successiva) si ignorano.

La finestra non tocca Revit e non scrive file: restituisce un ExportChoice o
None. Compatibile con Rocket mode: nessuno stato di modulo.
"""

import os
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window, MessageBox, Thickness
from System.Windows.Controls import CheckBox, StackPanel
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode

from pyrevit import script

import RNWFlowCalculation_Report as rp
import RNWFlowCalculation_Pdf as rpdf

XAML_FILE_NAME = 'RNWFlowCalculation_Export.xaml'
CONFIG_SECTION = 'ESA_RNWFlowCalculation'
WINDOW_TITLE = 'Export Calculation'

FORMAT_EXCEL = "xlsx"
FORMAT_PDF = "pdf"

# (chiave di RNWFlowCalculation_Pdf.PAPER_SIZES, voce della tendina)
PAPERS = [("A4", "A4 landscape"), ("A3", "A3 landscape")]

# Tabelle: (nome, catalogo, opzione della config)
GENERAL = "general"
SECTIONS = "sections"
FIXTURES = "fixtures"
TABLES = [
    (GENERAL, rp.GENERAL_COLUMNS, "export_general_columns"),
    (SECTIONS, rp.SECTION_COLUMNS, "export_section_columns"),
    (FIXTURES, rp.FIXTURE_COLUMNS, "export_fixture_columns"),
]
TABLE_LABELS = {GENERAL: "systems table of the General tab", SECTIONS: "Sections tab",
                FIXTURES: "Fixtures tab"}


class ExportChoice(object):
    """Scelte della finestra: formato, pagina del PDF, chiavi delle colonne."""

    def __init__(self, file_format, paper, columns):
        self.file_format = file_format
        self.paper = paper
        self.columns = dict(columns)        # nome della tabella -> [chiavi]

    @property
    def extension(self):
        return "." + self.file_format


def _saved_keys(cfg, option, catalog):
    """Chiavi salvate e ancora nel catalogo; le predefinite se non ce ne sono."""
    valid = set(column.key for column in catalog)
    try:
        text = cfg.get_option(option, "") if cfg is not None else ""
    except Exception:
        text = ""
    keys = [key for key in (text or "").split(",") if key in valid]
    return keys or rp.default_keys(catalog)


class ExportForm(Window):
    def __init__(self):
        self.choice = None
        self._boxes = {}
        self._load_xaml()
        self._restore()

    def _load_xaml(self):
        xaml_path = script.get_bundle_file(XAML_FILE_NAME)
        if not xaml_path or not os.path.exists(xaml_path):
            xaml_path = os.path.join(os.path.dirname(__file__), XAML_FILE_NAME)

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
        self.ShowInTaskbar = root.ShowInTaskbar

        for name in ('rb_excel', 'rb_pdf', 'lbl_paper', 'cbo_paper', 'pnl_general',
                     'pnl_sections', 'pnl_fixtures', 'btn_sections_default',
                     'btn_sections_all', 'btn_sections_none', 'btn_fixtures_default',
                     'btn_fixtures_all', 'btn_fixtures_none', 'btn_ok', 'btn_cancel'):
            setattr(self, name, root.FindName(name))

        for _key, label in PAPERS:
            self.cbo_paper.Items.Add(label)
        panels = {GENERAL: self.pnl_general, SECTIONS: self.pnl_sections,
                  FIXTURES: self.pnl_fixtures}
        for table, catalog, _option in TABLES:
            # Sections ha due colonne (UniformGrid): due StackPanel, cosi' le
            # caselle si leggono dall'alto in basso nell'ordine della griglia.
            if table == SECTIONS:
                halves = [StackPanel(), StackPanel()]
                for half in halves:
                    panels[table].Children.Add(half)
                split = (len(catalog) + 1) // 2
            boxes = []
            for index, column in enumerate(catalog):
                box = CheckBox()
                box.Content = column.header
                box.Tag = column.key
                box.Margin = Thickness(0, 3, 10, 3)
                if table == SECTIONS:
                    halves[0 if index < split else 1].Children.Add(box)
                else:
                    panels[table].Children.Add(box)
                boxes.append(box)
            self._boxes[table] = boxes

        self.rb_excel.Checked += self.OnFormatChanged
        self.rb_pdf.Checked += self.OnFormatChanged
        self.btn_sections_default.Click += lambda sender, args: self._set_default(SECTIONS)
        self.btn_sections_all.Click += lambda sender, args: self._set_all(SECTIONS, True)
        self.btn_sections_none.Click += lambda sender, args: self._set_all(SECTIONS, False)
        self.btn_fixtures_default.Click += lambda sender, args: self._set_default(FIXTURES)
        self.btn_fixtures_all.Click += lambda sender, args: self._set_all(FIXTURES, True)
        self.btn_fixtures_none.Click += lambda sender, args: self._set_all(FIXTURES, False)
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel

    # --------------------------------------------------------------- settings

    def _restore(self):
        try:
            cfg = script.get_config(CONFIG_SECTION)
        except Exception:
            cfg = None
        try:
            file_format = cfg.get_option('export_format', FORMAT_EXCEL) if cfg else FORMAT_EXCEL
            paper = cfg.get_option('export_paper', rpdf.PAPER_DEFAULT) if cfg \
                else rpdf.PAPER_DEFAULT
        except Exception:
            file_format, paper = FORMAT_EXCEL, rpdf.PAPER_DEFAULT
        keys = [key for key, _label in PAPERS]
        self.cbo_paper.SelectedIndex = keys.index(paper) if paper in keys else 0
        if file_format == FORMAT_PDF:
            self.rb_pdf.IsChecked = True
        else:
            self.rb_excel.IsChecked = True
        self._update_paper()
        for table, catalog, option in TABLES:
            self._check(table, _saved_keys(cfg, option, catalog))

    def _save(self, choice):
        try:
            cfg = script.get_config(CONFIG_SECTION)
            cfg.set_option('export_format', choice.file_format)
            cfg.set_option('export_paper', choice.paper)
            for table, _catalog, option in TABLES:
                cfg.set_option(option, ",".join(choice.columns[table]))
            script.save_config()
        except Exception:
            pass

    # ------------------------------------------------------------------ caselle

    def _check(self, table, keys):
        chosen = set(keys)
        for box in self._boxes[table]:
            box.IsChecked = box.Tag in chosen

    def _set_all(self, table, value):
        for box in self._boxes[table]:
            box.IsChecked = value

    def _set_default(self, table):
        catalog = dict((name, items) for name, items, _option in TABLES)[table]
        self._check(table, rp.default_keys(catalog))

    def _checked(self, table):
        return [box.Tag for box in self._boxes[table] if box.IsChecked]

    def _update_paper(self):
        is_pdf = bool(self.rb_pdf.IsChecked)
        self.cbo_paper.IsEnabled = is_pdf
        self.lbl_paper.Opacity = 1.0 if is_pdf else 0.5

    def OnFormatChanged(self, sender, args):
        self._update_paper()

    # ------------------------------------------------------------------ pulsanti

    def OnOK(self, sender, args):
        columns = {}
        missing = []
        for table, _catalog, _option in TABLES:
            columns[table] = self._checked(table)
            if not columns[table]:
                missing.append(TABLE_LABELS[table])
        if missing:
            MessageBox.Show("Choose at least one column for the {0}.".format(
                " and the ".join(missing)), WINDOW_TITLE)
            return
        file_format = FORMAT_PDF if self.rb_pdf.IsChecked else FORMAT_EXCEL
        index = self.cbo_paper.SelectedIndex
        paper = PAPERS[index][0] if 0 <= index < len(PAPERS) else rpdf.PAPER_DEFAULT
        self.choice = ExportChoice(file_format, paper, columns)
        self._save(self.choice)
        self.Close()

    def OnCancel(self, sender, args):
        self.choice = None
        self.Close()


def ask_export_options(owner=None):
    """Mostra la finestra (modale); ExportChoice, o None se annullata."""
    form = ExportForm()
    if owner is not None:
        try:
            form.Owner = owner
        except Exception:
            pass
    form.ShowDialog()
    return form.choice
