# -*- coding: utf-8 -*-
import os

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Data")

import System
from System.Data import DataTable
from System.IO import File
from System.Windows import MessageBox
from System.Windows.Controls import DataGridLength, DataGridLengthUnitType, DataGridTextColumn
from System.Windows.Data import Binding
from System.Windows.Markup import XamlReader
from Microsoft.Win32 import OpenFileDialog

import SetPipeInsulation_Table as insulation_table

# Spessori proposti per lo spessore fisso (mm)
THICKNESS_OPTIONS = [6, 9, 13, 19, 25, 32, 40, 50, 55, 60]
DEFAULT_THICKNESS = 13


def fill_combo(combo, names, selected_name):
    """Riempie il menu e seleziona l'ultima scelta, se ancora presente."""
    for name in names:
        combo.Items.Add(name)
    if names:
        combo.SelectedIndex = names.index(selected_name) if selected_name in names else 0


class SetPipeInsulationWindow(object):
    """Finestra modale: sistema, ambito, tipi di isolamento, tabella spessori (DPR 412/93 interna o file)
    con anteprima, oppure spessore fisso."""

    def __init__(self, xaml_path, system_names, insulation_type_names, selection_count, settings):
        # settings: dict con le ultime impostazioni (vedi DEFAULTS nello script)
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        self.system_combo = self.window.FindName("SystemCombo")
        self.mode_all = self.window.FindName("ModeAll")
        self.mode_selected = self.window.FindName("ModeSelected")
        self.visible_combo = self.window.FindName("VisibleCombo")
        self.hidden_combo = self.window.FindName("HiddenCombo")
        self.builtin_check = self.window.FindName("BuiltinCheck")
        self.table_box = self.window.FindName("TableBox")
        self.browse_button = self.window.FindName("BrowseButton")
        self.preview_grid = self.window.FindName("PreviewGrid")
        self.preview_status = self.window.FindName("PreviewStatus")
        self.manual_check = self.window.FindName("ManualCheck")
        self.thickness_combo = self.window.FindName("ThicknessCombo")
        self.browse_button.Click += self._on_browse
        self.window.FindName("OkButton").Click += self._on_ok
        self.manual_check.Checked += self._on_manual_changed
        self.manual_check.Unchecked += self._on_manual_changed
        self.builtin_check.Checked += self._on_builtin_changed
        self.builtin_check.Unchecked += self._on_builtin_changed
        self.table_box.LostFocus += self._on_table_box_lost_focus
        self.result = None

        # Tabella mostrata in anteprima: righe lette dal file (o DPR 412/93) e percorso letto
        self.table_rows = []
        self.table_error = ""
        self.loaded_path = None

        self.system_names = system_names
        self.type_names = insulation_type_names
        fill_combo(self.system_combo, system_names, settings["system"])
        fill_combo(self.visible_combo, insulation_type_names, settings["visible_type"])
        fill_combo(self.hidden_combo, insulation_type_names, settings["hidden_type"])
        fill_combo(self.thickness_combo, THICKNESS_OPTIONS,
                   settings["thickness"] if settings["thickness"] in THICKNESS_OPTIONS else DEFAULT_THICKNESS)

        self.table_box.Text = settings["table_path"]
        self.manual_check.IsChecked = bool(settings["manual"])
        self._on_manual_changed(None, None)
        self.builtin_check.IsChecked = bool(settings["builtin_table"])
        self._on_builtin_changed(None, None)

        if selection_count:
            self.mode_selected.Content = "Only the selected pipes and fittings of this system ({} selected)".format(
                selection_count)
            self.mode_selected.IsChecked = bool(settings["selected_only"])
            self.mode_all.IsChecked = not settings["selected_only"]
        else:
            self.mode_selected.Content = "Only the selected pipes and fittings of this system (nothing selected)"
            self.mode_selected.IsEnabled = False
            self.mode_all.IsChecked = True

    def _on_manual_changed(self, sender, args):
        self.thickness_combo.IsEnabled = bool(self.manual_check.IsChecked)

    def _on_builtin_changed(self, sender, args):
        # con la tabella interna il percorso del file non serve: campo e pulsante disattivati (grigi)
        builtin = bool(self.builtin_check.IsChecked)
        self.table_box.IsEnabled = not builtin
        self.browse_button.IsEnabled = not builtin
        if builtin:
            rows, columns = insulation_table.builtin_table()
            self._show_preview(rows, columns)
            self.preview_status.Text = "DPR 412/93 table built into the tool: {} rows.".format(len(rows))
        else:
            self.loaded_path = None
            self._load_table_file()

    def _on_table_box_lost_focus(self, sender, args):
        if not self.builtin_check.IsChecked:
            self._load_table_file()

    def _on_browse(self, sender, args):
        dialog = OpenFileDialog()
        dialog.Title = "Select the insulation thickness table"
        dialog.Filter = "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        current = (self.table_box.Text or "").strip()
        if current and os.path.isdir(os.path.dirname(current)):
            dialog.InitialDirectory = os.path.dirname(current)
        else:
            dialog.InitialDirectory = os.path.expanduser("~\\Documents")
        if dialog.ShowDialog(self.window):
            self.table_box.Text = dialog.FileName
            self._load_table_file()

    def _load_table_file(self):
        """Legge il file indicato nel campo e aggiorna l'anteprima; non rilegge lo stesso percorso."""
        path = (self.table_box.Text or "").strip()
        if path == self.loaded_path:
            return
        self.loaded_path = path
        self.table_rows = []
        self.table_error = ""
        if not path:
            self.table_error = "Select a thickness table file, or use a fixed thickness."
            self._show_preview([], [])
            self.preview_status.Text = "No table file selected."
            return
        if not os.path.isfile(path):
            self.table_error = "Select an existing thickness table file, or use a fixed thickness."
            self._show_preview([], [])
            self.preview_status.Text = "File not found."
            return
        try:
            rows, columns, skipped_cells = insulation_table.read_table_file(path)
        except insulation_table.TableError as e:
            self.table_error = str(e)
            self._show_preview([], [])
            self.preview_status.Text = str(e)
            return
        except Exception as e:
            self.table_error = "Error reading the table file:\n{}".format(str(e))
            self._show_preview([], [])
            self.preview_status.Text = "The file cannot be read."
            return

        self.table_rows = rows
        self._show_preview(rows, columns)
        if not rows:
            self.table_error = "The thickness table has no valid row."
            self.preview_status.Text = "No valid row in the file."
            return
        status = "{} rows, thermal zones: {}.".format(len(rows), ", ".join(columns[1:]))
        if skipped_cells == 1:
            status += " 1 cell is not a whole number and is ignored."
        elif skipped_cells:
            status += " {} cells are not whole numbers and are ignored.".format(skipped_cells)
        self.preview_status.Text = status

    def _show_preview(self, rows, columns):
        """Mostra la tabella nella griglia: una colonna per intestazione, una riga per diametro."""
        table = DataTable("preview")
        string_type = clr.GetClrType(System.String)
        # nomi di colonna neutri (C0, C1...): "Non riscaldato" con lo spazio non va bene in un Binding
        for index in range(len(columns)):
            table.Columns.Add("C{}".format(index), string_type)
        for row in rows:
            data_row = table.NewRow()
            for index, column in enumerate(columns):
                data_row[index] = str(row[column]) if column in row else ""
            table.Rows.Add(data_row)

        self.preview_grid.Columns.Clear()
        for index, column in enumerate(columns):
            grid_column = DataGridTextColumn()
            grid_column.Header = "{} (mm)".format(column) if column == insulation_table.DIAMETER_COLUMN else column
            grid_column.Binding = Binding("C{}".format(index))
            # l'ultima colonna riempie lo spazio; le altre larghe quanto l'intestazione
            last = index == len(columns) - 1
            grid_column.Width = DataGridLength(1, DataGridLengthUnitType.Star if last else DataGridLengthUnitType.Auto)
            grid_column.MinWidth = 60
            self.preview_grid.Columns.Add(grid_column)
        self.preview_grid.ItemsSource = table.DefaultView

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        if self.system_combo.SelectedIndex < 0:
            return False, "Select a piping system.", None
        if self.visible_combo.SelectedIndex < 0 or self.hidden_combo.SelectedIndex < 0:
            return False, "Select both insulation types.", None
        manual = bool(self.manual_check.IsChecked)
        builtin = bool(self.builtin_check.IsChecked)
        table_path = (self.table_box.Text or "").strip()
        table_rows = []
        # la tabella serve solo se lo spessore non e' fisso
        if not manual:
            if builtin:
                table_rows = insulation_table.builtin_table()[0]
            else:
                self._load_table_file()  # percorso scritto a mano e confermato con Invio
                if self.table_error:
                    return False, self.table_error, None
                table_rows = self.table_rows
        return True, "", {
            "system": self.system_names[self.system_combo.SelectedIndex],
            "selected_only": bool(self.mode_selected.IsChecked),
            "visible_type": self.type_names[self.visible_combo.SelectedIndex],
            "hidden_type": self.type_names[self.hidden_combo.SelectedIndex],
            "builtin_table": builtin,
            "table_path": table_path,
            "table": table_rows,
            "manual": manual,
            "thickness": THICKNESS_OPTIONS[self.thickness_combo.SelectedIndex],
        }

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Set Pipe Insulation")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
