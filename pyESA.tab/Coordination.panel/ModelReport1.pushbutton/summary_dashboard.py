# -*- coding: utf-8 -*-
"""
Summary Dashboard - Dashboard di riepilogo post-estrazione
Mostra KPI aggregati, tabella file e health checks per file.
"""

import os

# WPF imports
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode

# pyRevit imports
from pyrevit import script

OUTPUT = script.get_output()
LOGGER = script.get_logger()


class SummaryDashboard(Window):
    """Dashboard di riepilogo post-estrazione con KPI, tabella file e health checks."""

    # Colonne della tabella file: (chiave, colore_barra)
    TABLE_COLUMNS = [
        ('FileName',          None),
        ('Score',             '#4CAF50'),
        ('FileSize_MB',       '#1976D2'),
        ('Warnings',          '#D32F2F'),
        ('UnMaterials',       '#E65100'),
        ('PurgeableElements', '#D32F2F'),
        ('UnVT',              '#E65100'),
        ('Links',             '#1976D2'),
        ('Sheets',            '#1976D2'),
        ('Views',             '#1976D2'),
        ('Elements',          '#9E9E9E'),
    ]

    # Colori pallino per CheckPassed
    STATUS_COLORS = {
        'YES':        '#4CAF50',
        'NO':         '#D32F2F',
        'Excellent':  '#2E7D32',
        'Good':       '#4CAF50',
        'Sufficient': '#FF9800',
        'Poor':       '#E65100',
        'Bad':        '#D32F2F',
    }

    def __init__(self):
        Window.__init__(self)
        self._csv_data = None
        self._rows = []
        self._load_xaml()

    def _load_xaml(self):
        """Carica il file XAML della dashboard."""
        script_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(script_dir, 'SummaryDashboard.xaml')

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
        self.Background = root.Background

        # Riferimenti KPI
        self.kpi_files = root.FindName('kpi_files')
        self.kpi_size = root.FindName('kpi_size')
        self.kpi_warnings = root.FindName('kpi_warnings')
        self.kpi_purgeable = root.FindName('kpi_purgeable')
        self.kpi_sheets = root.FindName('kpi_sheets')
        self.kpi_views = root.FindName('kpi_views')
        self.kpi_links = root.FindName('kpi_links')
        self.table_body = root.FindName('table_body')

        # Riferimenti Health Checks
        self.cmb_file_select = root.FindName('cmb_file_select')
        self.txt_health_score = root.FindName('txt_health_score')
        self.health_checks_body = root.FindName('health_checks_body')

        # Event handler ComboBox
        self.cmb_file_select.SelectionChanged += self._on_file_selected

    def populate(self, csv_data):
        """Popola la dashboard con i dati estratti."""
        from System.Windows.Controls import Grid as WpfGrid, ColumnDefinition, TextBlock as WpfTextBlock
        from System.Windows.Controls import Border as WpfBorder
        from System.Windows import GridLength, GridUnitType, HorizontalAlignment, VerticalAlignment, Thickness, CornerRadius
        from System.Windows.Media import BrushConverter

        self._csv_data = csv_data
        self._converter = BrushConverter()

        files_data = csv_data.get('TAB_Files', [])
        warnings_data = csv_data.get('TAB_Warnings', [])
        materials_data = csv_data.get('TAB_Materials', [])
        templates_data = csv_data.get('TAB_ViewTemplates', [])
        links_data = csv_data.get('TAB_Links', [])
        views_data = csv_data.get('TAB_Views', [])

        # Calcola conteggi per file
        self._rows = []
        for f in files_data:
            fname = f.get('FileName', '')
            row = {
                'FileName': fname,
                'Score': self._to_int(f.get('KPI_HealthScore', 0)),
                'FileSize_MB': self._parse_size(f.get('FileSize_MB', 0)),
                'Warnings': len([w for w in warnings_data if w.get('FileName', '') == fname]),
                'UnMaterials': len([m for m in materials_data if m.get('FileName', '') == fname and m.get('IsUsed', '') == 'NO']),
                'PurgeableElements': self._to_int(f.get('PurgeableElements', 0)),
                'UnVT': len([t for t in templates_data if t.get('FileName', '') == fname and t.get('IsUsed', '') == 'NO']),
                'Links': len([l for l in links_data if l.get('FileName', '') == fname]),
                'Sheets': self._to_int(f.get('Sheets', 0)),
                'Views': len([v for v in views_data if v.get('FileName', '') == fname]),
                'Elements': self._to_int(f.get('Elements', 0)),
            }
            self._rows.append(row)

        self._rows.sort(key=lambda r: r['FileName'])

        # KPI totali
        self.kpi_files.Text = str(len(self._rows))
        self.kpi_size.Text = str(int(sum(r['FileSize_MB'] for r in self._rows)))
        self.kpi_warnings.Text = self._format_number(sum(r['Warnings'] for r in self._rows))
        self.kpi_purgeable.Text = self._format_number(sum(r['PurgeableElements'] for r in self._rows))
        self.kpi_sheets.Text = str(sum(r['Sheets'] for r in self._rows))
        self.kpi_views.Text = self._format_number(sum(r['Views'] for r in self._rows))
        self.kpi_links.Text = str(sum(r['Links'] for r in self._rows))

        # Valori massimi per barre
        max_values = {}
        for key, _ in self.TABLE_COLUMNS:
            if key == 'FileName':
                continue
            vals = [r.get(key, 0) for r in self._rows]
            max_values[key] = max(vals) if vals else 1

        # Popola tabella file
        conv = self._converter
        col_widths = [3, 1.1, 1.1, 1.1, 1.3, 1.1, 1.1, 0.9, 0.9, 0.9, 1.1]

        for row_idx, row in enumerate(self._rows):
            row_bg = conv.ConvertFromString("#FFFFFF" if row_idx % 2 == 0 else "#FAFAFA")

            row_border = WpfBorder()
            row_border.Background = row_bg

            row_grid = WpfGrid()
            for w in col_widths:
                cd = ColumnDefinition()
                cd.Width = GridLength(w, GridUnitType.Star)
                row_grid.ColumnDefinitions.Add(cd)

            for col_idx, (key, bar_color) in enumerate(self.TABLE_COLUMNS):
                value = row.get(key, '')

                if key == 'FileName':
                    tb = WpfTextBlock()
                    tb.Text = str(value)
                    tb.Foreground = conv.ConvertFromString("#333333")
                    tb.FontSize = 12
                    tb.VerticalAlignment = VerticalAlignment.Center
                    tb.Margin = Thickness(10, 7, 10, 7)
                    tb.TextTrimming = tb.TextTrimming.CharacterEllipsis
                    WpfGrid.SetColumn(tb, col_idx)
                    row_grid.Children.Add(tb)
                else:
                    cell_grid = WpfGrid()
                    cell_grid.Margin = Thickness(4, 3, 4, 3)

                    max_val = max_values.get(key, 1)
                    ratio = float(value) / float(max_val) if max_val > 0 else 0

                    # Barra proporzionale
                    bar_grid = WpfGrid()
                    cd1 = ColumnDefinition()
                    cd2 = ColumnDefinition()
                    if ratio > 0:
                        cd1.Width = GridLength(ratio * 100, GridUnitType.Star)
                        cd2.Width = GridLength((1 - ratio) * 100 + 0.001, GridUnitType.Star)
                    else:
                        cd1.Width = GridLength(0, GridUnitType.Pixel)
                        cd2.Width = GridLength(1, GridUnitType.Star)
                    bar_grid.ColumnDefinitions.Add(cd1)
                    bar_grid.ColumnDefinitions.Add(cd2)

                    bar_fill = WpfBorder()
                    bar_fill.Background = conv.ConvertFromString(bar_color)
                    bar_fill.Opacity = 0.2
                    bar_fill.CornerRadius = CornerRadius(3)
                    bar_fill.Margin = Thickness(0, 1, 2, 1)
                    WpfGrid.SetColumn(bar_fill, 0)
                    bar_grid.Children.Add(bar_fill)

                    cell_grid.Children.Add(bar_grid)

                    # Testo valore
                    tb = WpfTextBlock()
                    if key == 'FileSize_MB':
                        tb.Text = str(int(value)) if value else "0"
                    else:
                        tb.Text = self._format_number(int(value)) if value else "0"
                    tb.Foreground = conv.ConvertFromString("#333333")
                    tb.FontSize = 12
                    tb.HorizontalAlignment = HorizontalAlignment.Right
                    tb.VerticalAlignment = VerticalAlignment.Center
                    tb.Margin = Thickness(6, 4, 6, 4)
                    cell_grid.Children.Add(tb)

                    WpfGrid.SetColumn(cell_grid, col_idx)
                    row_grid.Children.Add(cell_grid)

            row_border.Child = row_grid
            self.table_body.Children.Add(row_border)

        # Popola ComboBox file
        file_names = [r['FileName'] for r in self._rows]
        for fname in file_names:
            self.cmb_file_select.Items.Add(fname)
        if file_names:
            self.cmb_file_select.SelectedIndex = 0

    def _on_file_selected(self, sender, args):
        """Aggiorna la sezione Health Checks quando si seleziona un file."""
        selected = self.cmb_file_select.SelectedItem
        if not selected or not self._csv_data:
            return
        fname = str(selected)
        health_data = self._csv_data.get('TAB_HealthChecks', [])
        file_checks = [h for h in health_data if h.get('FileName', '') == fname]

        # Calcola score totale
        total_score = sum(self._to_int(h.get('Score', 0)) for h in file_checks)
        self.txt_health_score.Text = "{} / 100".format(total_score)

        # Colore score
        if total_score >= 80:
            self.txt_health_score.Foreground = self._converter.ConvertFromString('#2E7D32')
        elif total_score >= 50:
            self.txt_health_score.Foreground = self._converter.ConvertFromString('#E65100')
        else:
            self.txt_health_score.Foreground = self._converter.ConvertFromString('#D32F2F')

        # Ricostruisci tabella health checks
        self._populate_health_checks(file_checks)

    def _populate_health_checks(self, checks):
        """Popola la tabella dei health checks."""
        from System.Windows.Controls import Grid as WpfGrid, ColumnDefinition, TextBlock as WpfTextBlock
        from System.Windows.Controls import Border as WpfBorder
        from System.Windows import GridLength, GridUnitType, HorizontalAlignment, VerticalAlignment, Thickness, CornerRadius
        from System.Windows.Shapes import Ellipse

        conv = self._converter
        self.health_checks_body.Children.Clear()

        for row_idx, check in enumerate(checks):
            row_bg = conv.ConvertFromString("#FFFFFF" if row_idx % 2 == 0 else "#FAFAFA")

            row_border = WpfBorder()
            row_border.Background = row_bg

            row_grid = WpfGrid()
            for w in [3, 1.5, 1.5, 1]:
                cd = ColumnDefinition()
                cd.Width = GridLength(w, GridUnitType.Star)
                row_grid.ColumnDefinitions.Add(cd)

            # Col 0: Check Description
            tb_desc = WpfTextBlock()
            tb_desc.Text = str(check.get('CheckDescription', ''))
            tb_desc.Foreground = conv.ConvertFromString("#333333")
            tb_desc.FontSize = 12
            tb_desc.VerticalAlignment = VerticalAlignment.Center
            tb_desc.Margin = Thickness(10, 7, 10, 7)
            WpfGrid.SetColumn(tb_desc, 0)
            row_grid.Children.Add(tb_desc)

            # Col 1: Check Passed (pallino + testo)
            passed = str(check.get('CheckPassed', ''))
            status_color = self.STATUS_COLORS.get(passed, '#9E9E9E')

            from System.Windows.Controls import StackPanel as WpfStackPanel
            sp = WpfStackPanel()
            sp.Orientation = sp.Orientation.Horizontal
            sp.HorizontalAlignment = HorizontalAlignment.Center
            sp.VerticalAlignment = VerticalAlignment.Center

            dot = Ellipse()
            dot.Width = 10
            dot.Height = 10
            dot.Fill = conv.ConvertFromString(status_color)
            dot.Margin = Thickness(0, 0, 6, 0)
            sp.Children.Add(dot)

            tb_passed = WpfTextBlock()
            tb_passed.Text = passed
            tb_passed.Foreground = conv.ConvertFromString(status_color)
            tb_passed.FontSize = 12
            from System.Windows import FontWeights
            tb_passed.FontWeight = FontWeights.SemiBold
            tb_passed.VerticalAlignment = VerticalAlignment.Center
            sp.Children.Add(tb_passed)

            WpfGrid.SetColumn(sp, 1)
            row_grid.Children.Add(sp)

            # Col 2: Value
            tb_val = WpfTextBlock()
            tb_val.Text = str(check.get('Value', ''))
            tb_val.Foreground = conv.ConvertFromString("#333333")
            tb_val.FontSize = 12
            tb_val.HorizontalAlignment = HorizontalAlignment.Right
            tb_val.VerticalAlignment = VerticalAlignment.Center
            tb_val.Margin = Thickness(10, 7, 10, 7)
            WpfGrid.SetColumn(tb_val, 2)
            row_grid.Children.Add(tb_val)

            # Col 3: Score
            score = self._to_int(check.get('Score', 0))
            max_score = self._to_int(check.get('MaxScore', 0))
            tb_score = WpfTextBlock()
            tb_score.Text = "{} / {}".format(score, max_score) if max_score > 0 else str(score)
            tb_score.Foreground = conv.ConvertFromString("#333333")
            tb_score.FontSize = 12
            tb_score.FontWeight = FontWeights.SemiBold
            tb_score.HorizontalAlignment = HorizontalAlignment.Right
            tb_score.VerticalAlignment = VerticalAlignment.Center
            tb_score.Margin = Thickness(10, 7, 10, 7)
            WpfGrid.SetColumn(tb_score, 3)
            row_grid.Children.Add(tb_score)

            row_border.Child = row_grid
            self.health_checks_body.Children.Add(row_border)

    @staticmethod
    def _to_int(val):
        """Converte un valore a intero in modo sicuro."""
        try:
            return int(val)
        except:
            return 0

    @staticmethod
    def _parse_size(val):
        """Converte FileSize_MB gestendo sia punto che virgola come separatore decimale."""
        try:
            return float(str(val).replace(',', '.'))
        except:
            return 0.0

    @staticmethod
    def _to_float(val):
        """Converte un valore a float in modo sicuro."""
        try:
            return float(val)
        except:
            return 0.0

    @staticmethod
    def _format_number(num):
        """Formatta numeri grandi (es. 10290 -> 10.29K)."""
        try:
            num = int(num)
        except:
            return str(num)
        if num >= 1000:
            return "{:.2f}K".format(num / 1000.0)
        return str(num)


def show_summary_dashboard(csv_data):
    """Crea e mostra la dashboard di riepilogo (non modale)."""
    try:
        dashboard = SummaryDashboard()
        dashboard.populate(csv_data)
        dashboard.Show()
    except Exception as e:
        OUTPUT.print_md("⚠️ Errore apertura dashboard: {}".format(str(e)))
        LOGGER.error("Errore SummaryDashboard: {}".format(str(e)))
