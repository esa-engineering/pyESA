# -*- coding: utf-8 -*-
"""
RelinkModels_Form_ui.py - finestra principale di RelinkModels.

Un'unica finestra, due azioni:
- Check: legge i link dei modelli chiusi e dice per ognuno se il file esiste al
  percorso salvato e in che stato era il link all'ultimo salvataggio. Se e' data
  una cartella, cerca anche il file li' (per nome) e propone il nuovo percorso.
- Apply: riscrive tutti i link spuntati, anche quelli che oggi funzionano.

La tabella e' un outline in un solo DataTable, come la distinta di MEPQTO: per
ogni modello una riga di intestazione (Kind = H) con i totali, sotto i suoi link
(Kind = L). Il doppio clic su un'intestazione chiude/apre il gruppo impostando la
colonna nascosta Hidden; la spunta di un'intestazione vale per tutti i suoi link.
Niente ordinamento per colonna: romperebbe l'outline.
"""

import os
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Data')
clr.AddReference('System.Windows.Forms')

from System import Boolean, Int32, String
from System.Data import DataTable
from System.IO import FileStream, FileMode
from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult
from System.Windows.Input import Mouse, Cursors
from System.Windows.Markup import XamlReader
from System.Windows.Forms import OpenFileDialog, DialogResult

from pyrevit import script

import RelinkModels_Match as rm
import RelinkModels_Core as rc

XAML_FILE_NAME = 'RelinkModels_Form.xaml'
CONFIG_SECTION = 'ESA_RelinkModels'
MAX_SAVED_MODELS = 50

COL_APPLY = "Apply"
COL_MODEL = "Model"
COL_LINK = "Link"
COL_KIND = "Type"
COL_EXISTS = "File exists"
COL_STATE = "Saved state"
COL_MATCH = "In new folder"
COL_RESULT = "Result"
COL_OLD = "Current path"
COL_NEW = "New path"
# colonne nascoste
COL_IDX = "Idx"            # indice in self._records; -1 sulle intestazioni
COL_ROW = "Kind"           # ROW_HEADER / ROW_LINK (nome usato dal DataTrigger della XAML)
COL_GROUP = "Group"        # percorso del modello: lega link e intestazione
COL_HIDDEN = "Hidden"      # True per i link di un gruppo chiuso

HIDDEN_COLUMNS = (COL_IDX, COL_ROW, COL_GROUP, COL_HIDDEN)
READ_ONLY_COLUMNS = (COL_MODEL, COL_LINK, COL_KIND, COL_EXISTS, COL_STATE, COL_MATCH,
                     COL_RESULT, COL_OLD)

ROW_HEADER = "H"
ROW_LINK = "L"
GLYPH_OPEN = u"▼"     # triangolo in giu'
GLYPH_CLOSED = u"▶"   # triangolo a destra

YES = "Yes"
NO = "No"
MATCH_BY_HAND = "Picked by hand"
LOADED = "Loaded"

RESULT_WRITTEN = "Written"
RESULT_FAILED = "Failed"
RESULT_UNREAD = "Cannot read"
RESULT_NO_PATH = "No new path"
RESULT_PENDING = "Rewritten, not yet opened in Revit"
RESULT_CENTRAL = "Written, saved as new central"
RESULT_CENTRAL_FAILED = "Written, but saving as new central failed"


def _pick_folder(title):
    """Dialogo cartella moderno (OpenFileDialog senza validazione), come ModelReport1."""
    try:
        dlg = OpenFileDialog()
        dlg.Title = title
        dlg.FileName = "Select folder"
        dlg.Filter = "Folder|*.folder-placeholder"
        dlg.CheckFileExists = False
        dlg.CheckPathExists = True
        dlg.ValidateNames = False
        if dlg.ShowDialog() == DialogResult.OK:
            return os.path.dirname(dlg.FileName)
    except Exception:
        pass
    return None


class RelinkForm(Window):
    """Finestra di Relink Models."""

    def __init__(self, app):
        self._app = app         # Application di Revit, per aprire i modelli (opzione nuova central)
        self._records = []      # LinkRecord, nella posizione indicata da Idx
        self._tds = {}          # percorso modello -> TransmissionData letto dal Check
        self._table = None
        self._unreadable = 0
        self._has_folder = False
        self._collapsed = set() # gruppi chiusi (percorso modello), restano chiusi a un nuovo Check
        self._syncing = False   # evita che le scritture da codice rilancino ColumnChanged
        self._busy = False
        self._load_xaml()
        self._restore_settings()

    # ------------------------------------------------------------------ setup

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

        for name in ('ModelsList', 'AddFilesButton', 'RemoveButton', 'ClearButton',
                     'RootBox', 'BrowseRootButton', 'SubfoldersCheck', 'RevitCheck',
                     'CadCheck', 'BackupCheck', 'CentralCheck', 'ShowNotFoundCheck',
                     'LinksGrid', 'SummaryLabel',
                     'PickFileButton', 'CheckAllButton', 'CheckNoneButton',
                     'CancelButton', 'CheckButton', 'ApplyButton'):
            setattr(self, name, root.FindName(name))

        self.AddFilesButton.Click += self.OnAddFiles
        self.RemoveButton.Click += self.OnRemove
        self.ClearButton.Click += self.OnClear
        self.BrowseRootButton.Click += self.OnBrowseRoot
        self.PickFileButton.Click += self.OnPickFile
        self.CheckAllButton.Click += self.OnCheckAll
        self.CheckNoneButton.Click += self.OnCheckNone
        self.CancelButton.Click += self.OnClose
        self.CheckButton.Click += self.OnCheck
        self.ApplyButton.Click += self.OnApply
        self.LinksGrid.AutoGeneratingColumn += self.OnAutoColumn
        self.LinksGrid.MouseDoubleClick += self.OnGridDoubleClick
        self.ShowNotFoundCheck.Checked += self.OnShowNotFound
        self.ShowNotFoundCheck.Unchecked += self.OnShowNotFound

    # --------------------------------------------------------------- settings

    def _restore_settings(self):
        models = []
        try:
            cfg = script.get_config(CONFIG_SECTION)
            raw = getattr(cfg, 'models', '') or ''
            models = [m for m in raw.split('|') if m]
            self.RootBox.Text = getattr(cfg, 'root', '') or ''
            self.SubfoldersCheck.IsChecked = bool(getattr(cfg, 'subfolders', True))
            self.RevitCheck.IsChecked = bool(getattr(cfg, 'revit_links', True))
            self.CadCheck.IsChecked = bool(getattr(cfg, 'cad_links', False))
            self.BackupCheck.IsChecked = bool(getattr(cfg, 'backup', True))
            self.ShowNotFoundCheck.IsChecked = bool(getattr(cfg, 'show_not_found', False))
        except Exception:
            self.SubfoldersCheck.IsChecked = True
            self.RevitCheck.IsChecked = True
            self.CadCheck.IsChecked = False
            self.BackupCheck.IsChecked = True
        for m in models[:MAX_SAVED_MODELS]:
            if os.path.exists(m):
                self.ModelsList.Items.Add(m)

    def _save_settings(self):
        try:
            cfg = script.get_config(CONFIG_SECTION)
            cfg.models = '|'.join([str(m) for m in self.ModelsList.Items][:MAX_SAVED_MODELS])
            cfg.root = self.RootBox.Text or ''
            cfg.subfolders = bool(self.SubfoldersCheck.IsChecked)
            cfg.revit_links = bool(self.RevitCheck.IsChecked)
            cfg.cad_links = bool(self.CadCheck.IsChecked)
            cfg.backup = bool(self.BackupCheck.IsChecked)
            cfg.show_not_found = bool(self.ShowNotFoundCheck.IsChecked)
            script.save_config()
        except Exception:
            pass

    # ----------------------------------------------------------- model list

    def OnAddFiles(self, sender, args):
        dlg = OpenFileDialog()
        dlg.Title = "Select the Revit models to check"
        dlg.Filter = "Revit models (*.rvt)|*.rvt"
        dlg.Multiselect = True
        if dlg.ShowDialog() == DialogResult.OK:
            known = [str(m).lower() for m in self.ModelsList.Items]
            for f in dlg.FileNames:
                if f.lower() not in known:
                    self.ModelsList.Items.Add(f)
            self._invalidate_check()

    def OnRemove(self, sender, args):
        for item in list(self.ModelsList.SelectedItems):
            self.ModelsList.Items.Remove(item)
        self._invalidate_check()

    def OnClear(self, sender, args):
        self.ModelsList.Items.Clear()
        self._invalidate_check()

    def OnBrowseRoot(self, sender, args):
        folder = _pick_folder("Select the folder with the moved files")
        if folder:
            self.RootBox.Text = folder
            self._invalidate_check()

    def _invalidate_check(self):
        """Dopo una modifica agli input la tabella non e' piu' valida."""
        self._records = []
        self._tds = {}
        self._table = None
        self.LinksGrid.ItemsSource = None
        self.ApplyButton.IsEnabled = False
        self.SummaryLabel.Text = ""

    # ----------------------------------------------------------------- check

    def _kinds(self):
        kinds = []
        if self.RevitCheck.IsChecked:
            kinds.append(rc.KIND_REVIT)
        if self.CadCheck.IsChecked:
            kinds.append(rc.KIND_CAD)
        return kinds

    def _extensions(self, kinds):
        exts = []
        if rc.KIND_REVIT in kinds:
            exts.append(".rvt")
        if rc.KIND_CAD in kinds:
            exts.extend([".dwg", ".dxf", ".dgn"])
        return exts

    def _check_inputs(self):
        """(bool, messaggio): validazione unica prima del Check. La cartella e' facoltativa."""
        if self.ModelsList.Items.Count == 0:
            return False, "Add at least one model."
        root = (self.RootBox.Text or "").strip()
        if root and not os.path.isdir(root):
            return False, "The new folder does not exist. Choose another one or leave it empty."
        if not self._kinds():
            return False, "Tick at least one link type."
        return True, ""

    def _new_table(self):
        table = DataTable()
        table.Columns.Add(COL_APPLY, Boolean)
        for title in (COL_MODEL, COL_LINK, COL_KIND, COL_EXISTS, COL_STATE, COL_MATCH,
                      COL_RESULT, COL_OLD, COL_NEW):
            table.Columns.Add(title, String)
        table.Columns.Add(COL_IDX, Int32)
        table.Columns.Add(COL_ROW, String)
        table.Columns.Add(COL_GROUP, String)
        table.Columns.Add(COL_HIDDEN, Boolean)
        return table

    def _add_header(self, table, model, result=""):
        row = table.NewRow()
        row[COL_APPLY] = False
        row[COL_MODEL] = os.path.basename(model)
        for title in (COL_LINK, COL_KIND, COL_EXISTS, COL_STATE, COL_MATCH, COL_OLD, COL_NEW):
            row[title] = ""
        row[COL_RESULT] = result
        row[COL_IDX] = -1
        row[COL_ROW] = ROW_HEADER
        row[COL_GROUP] = model
        row[COL_HIDDEN] = False
        table.Rows.Add(row)

    def OnCheck(self, sender, args):
        ok, message = self._check_inputs()
        if not ok:
            MessageBox.Show(message, "Relink Models", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        self._save_settings()
        kinds = self._kinds()
        root = (self.RootBox.Text or "").strip()
        self._has_folder = bool(root)
        index = rm.scan_folder(root, bool(self.SubfoldersCheck.IsChecked),
                               self._extensions(kinds)) if root else {}

        table = self._new_table()
        self._records = []
        self._tds = {}
        exists_cache = {}      # un percorso di rete irraggiungibile si interroga una volta sola
        unreadable = 0
        for model in [str(m) for m in self.ModelsList.Items]:
            try:
                records, td = rc.read_links(model, kinds)
            except Exception as ex:
                unreadable += 1
                self._add_header(table, model,
                                 "{}: {}".format(RESULT_UNREAD, str(ex).split("\n")[0]))
                continue
            self._tds[model] = td
            self._add_header(table, model)
            for rec in sorted(records, key=lambda r: r.name.lower()):
                key = rec.old_path.lower()
                if key not in exists_cache:
                    exists_cache[key] = os.path.isfile(rec.old_path)
                self._records.append(rec)
                row = table.NewRow()
                row[COL_MODEL] = ""
                row[COL_LINK] = rec.name
                row[COL_KIND] = rec.kind
                row[COL_EXISTS] = YES if exists_cache[key] else NO
                row[COL_STATE] = rec.load_state
                # il Check legge il percorso dell'ultimo salvataggio: un link gia' riscritto
                # lo mostra finche' Revit non apre e salva il modello
                row[COL_RESULT] = "{}: {}".format(RESULT_PENDING, rec.pending_path) \
                    if rec.pending_path else ""
                row[COL_OLD] = rec.old_path
                if root:
                    new_path, match, count = rm.pick_candidate(rec.old_path, index)
                    row[COL_NEW] = new_path
                    row[COL_MATCH] = match if match != rm.STATUS_AMBIGUOUS else \
                        "{} ({} files)".format(match, count)
                    # sostituzione in ogni caso: anche i link che oggi funzionano
                    row[COL_APPLY] = bool(new_path)
                else:
                    row[COL_NEW] = ""
                    row[COL_MATCH] = ""
                    row[COL_APPLY] = False
                row[COL_IDX] = len(self._records) - 1
                row[COL_ROW] = ROW_LINK
                row[COL_GROUP] = model
                row[COL_HIDDEN] = model in self._collapsed
                table.Rows.Add(row)

        table.ColumnChanged += self.OnTableColumnChanged
        self._table = table
        self._unreadable = unreadable
        self._refresh_headers()
        self._apply_filter()
        self.LinksGrid.ItemsSource = table.DefaultView
        self._update_summary()

    # --------------------------------------------------------------- outline

    def _link_rows(self, group=None):
        for row in self._table.Rows:
            if row[COL_ROW] == ROW_LINK and (group is None or row[COL_GROUP] == group):
                yield row

    def _refresh_headers(self):
        """Ricalcola i totali delle intestazioni dai loro link."""
        if self._table is None:
            return
        stats = {}
        for row in self._link_rows():
            s = stats.setdefault(row[COL_GROUP], {'n': 0, 'missing': 0, 'unloaded': 0,
                                                  'found': 0, 'ticked': 0, 'written': 0,
                                                  'failed': 0, 'outside': 0, 'pending': 0})
            s['n'] += 1
            if row[COL_EXISTS] == NO:
                s['missing'] += 1
            if row[COL_STATE] != LOADED:
                s['unloaded'] += 1
            if row[COL_NEW]:
                s['found'] += 1
            if row[COL_MATCH] == rm.STATUS_NOT_FOUND:
                s['outside'] += 1
            if row[COL_APPLY]:
                s['ticked'] += 1
            result = row[COL_RESULT] or ""
            if result in (RESULT_WRITTEN, RESULT_CENTRAL):
                s['written'] += 1
            elif result.startswith(RESULT_FAILED) or result.startswith(RESULT_CENTRAL_FAILED):
                s['failed'] += 1
            elif result.startswith(RESULT_PENDING):
                s['pending'] += 1

        self._syncing = True
        try:
            for row in self._table.Rows:
                if row[COL_ROW] != ROW_HEADER:
                    continue
                group = row[COL_GROUP]
                glyph = GLYPH_CLOSED if group in self._collapsed else GLYPH_OPEN
                row[COL_MODEL] = u"{} {}".format(glyph, os.path.basename(group))
                s = stats.get(group)
                if s is None:
                    continue            # modello illeggibile o senza link: resta il Result
                row[COL_LINK] = "{} links".format(s['n'])
                row[COL_EXISTS] = "{} missing".format(s['missing'])
                row[COL_STATE] = "{} not loaded".format(s['unloaded'])
                if self._has_folder:
                    text = "{} found".format(s['found'])
                    if s['outside']:
                        text += ", {} not in folder".format(s['outside'])
                    row[COL_MATCH] = text
                parts = []
                if s['written']:
                    parts.append("{} written".format(s['written']))
                if s['failed']:
                    parts.append("{} failed".format(s['failed']))
                if s['pending']:
                    parts.append("{} rewritten, not yet opened".format(s['pending']))
                row[COL_RESULT] = ", ".join(parts)
                row[COL_APPLY] = s['ticked'] > 0
        finally:
            self._syncing = False

    def OnTableColumnChanged(self, sender, args):
        """La spunta di un'intestazione vale per tutti i link del modello che hanno un nuovo percorso."""
        if self._syncing or args.Column.ColumnName != COL_APPLY:
            return
        row = args.Row
        self._syncing = True
        try:
            if row[COL_ROW] == ROW_HEADER:
                value = bool(row[COL_APPLY])
                for link in self._link_rows(row[COL_GROUP]):
                    if link[COL_NEW]:
                        link[COL_APPLY] = value
        finally:
            self._syncing = False
        self._refresh_headers()
        self._update_summary()

    def OnGridDoubleClick(self, sender, args):
        item = self.LinksGrid.SelectedItem
        if item is None or self._table is None:
            return
        try:
            if item[COL_ROW] != ROW_HEADER:
                return
            group = item[COL_GROUP]
        except Exception:
            return
        if group in self._collapsed:
            self._collapsed.remove(group)
        else:
            self._collapsed.add(group)
        hidden = group in self._collapsed
        self._syncing = True
        try:
            for link in self._link_rows(group):
                link[COL_HIDDEN] = hidden
        finally:
            self._syncing = False
        self._refresh_headers()

    def _apply_filter(self):
        """Gruppi chiusi e, con una cartella, link a file non presenti li' (quelli non spostati)."""
        if self._table is None:
            return
        expr = "[{}] = false".format(COL_HIDDEN)
        if self._has_folder and not self.ShowNotFoundCheck.IsChecked:
            expr += " AND ([{}] = '{}' OR [{}] <> '{}')".format(
                COL_ROW, ROW_HEADER, COL_MATCH, rm.STATUS_NOT_FOUND)
        self._table.DefaultView.RowFilter = expr

    def OnShowNotFound(self, sender, args):
        self._apply_filter()
        self._update_summary()

    def OnAutoColumn(self, sender, args):
        header = str(args.Column.Header)
        if header in HIDDEN_COLUMNS:
            args.Cancel = True
        elif header in READ_ONLY_COLUMNS:
            # sola lettura solo per l'utente: DataColumn.ReadOnly bloccherebbe anche il codice
            args.Column.IsReadOnly = True

    def _update_summary(self):
        if self._table is None:
            return
        total = ticked = missing = not_loaded = outside = 0
        for row in self._link_rows():
            total += 1
            if row[COL_APPLY]:
                ticked += 1
            if row[COL_EXISTS] == NO:
                missing += 1
            if row[COL_STATE] != LOADED:
                not_loaded += 1
            if row[COL_MATCH] == rm.STATUS_NOT_FOUND:
                outside += 1
        models = self.ModelsList.Items.Count
        text = "{} models, {} links: {} files missing, {} not loaded".format(
            models, total, missing, not_loaded)
        if self._has_folder:
            text += ". {} ticked".format(ticked)
            if outside and not self.ShowNotFoundCheck.IsChecked:
                text += ", {} not in the new folder (hidden)".format(outside)
        if self._unreadable:
            text += ". {} models could not be read".format(self._unreadable)
        self.SummaryLabel.Text = text
        self.ApplyButton.IsEnabled = total > 0

    # ------------------------------------------------------------ grid tools

    def OnCheckAll(self, sender, args):
        self._set_all(True)

    def OnCheckNone(self, sender, args):
        self._set_all(False)

    def _set_all(self, value):
        if self._table is None:
            return
        self.LinksGrid.CommitEdit()
        self._syncing = True
        try:
            for row in self._link_rows():
                if row[COL_NEW]:
                    row[COL_APPLY] = value
        finally:
            self._syncing = False
        self._refresh_headers()
        self._update_summary()

    def OnPickFile(self, sender, args):
        """Sceglie a mano il file per le righe selezionate (casi ambigui o non trovati)."""
        if self._table is None or self.LinksGrid.SelectedItems.Count == 0:
            return
        dlg = OpenFileDialog()
        dlg.Title = "Select the new file"
        dlg.Filter = "Linked files|*.rvt;*.dwg;*.dxf;*.dgn|All files|*.*"
        if (self.RootBox.Text or "").strip():
            dlg.InitialDirectory = self.RootBox.Text.strip()
        if dlg.ShowDialog() != DialogResult.OK:
            return
        self._syncing = True
        try:
            for item in self.LinksGrid.SelectedItems:
                if item[COL_ROW] == ROW_LINK:
                    item[COL_NEW] = dlg.FileName
                    item[COL_APPLY] = True
                    item[COL_MATCH] = MATCH_BY_HAND
        finally:
            self._syncing = False
        self._refresh_headers()
        self._update_summary()

    # ----------------------------------------------------------------- apply

    def OnApply(self, sender, args):
        if self._table is None or self._busy:
            return
        self.LinksGrid.CommitEdit()

        # righe da scrivere, raggruppate per modello
        jobs = {}
        self._syncing = True
        try:
            for row in self._link_rows():
                if not row[COL_APPLY]:
                    continue
                new_path = (row[COL_NEW] or "").strip()
                if not new_path or not os.path.isfile(new_path):
                    row[COL_RESULT] = RESULT_NO_PATH
                    row[COL_APPLY] = False
                    continue
                jobs.setdefault(row[COL_GROUP], []).append((row, new_path))
        finally:
            self._syncing = False
        if not jobs:
            self._refresh_headers()
            MessageBox.Show("Nothing to write: tick at least one row with an existing new file.",
                            "Relink Models", MessageBoxButton.OK, MessageBoxImage.Information)
            self._update_summary()
            return

        count = sum(len(v) for v in jobs.values())
        central = bool(self.CentralCheck.IsChecked)
        message = "Rewrite {} links in {} closed models?".format(count, len(jobs))
        message += ("\nA backup copy is made first." if self.BackupCheck.IsChecked
                    else "\nNo backup copy will be made.")
        if central:
            message += ("\n\nEach model is then opened and saved as a new central. This takes "
                        "minutes per model, and local files of the old central must be recreated.")
        answer = MessageBox.Show(message, "Relink Models", MessageBoxButton.OKCancel,
                                 MessageBoxImage.Question)
        if answer != MessageBoxResult.OK:
            return

        self._busy = True
        self._syncing = True
        Mouse.OverrideCursor = Cursors.Wait
        failures = 0
        try:
            for model, items in jobs.items():
                changes = [(self._records[row[COL_IDX]], path) for row, path in items]
                written = []
                try:
                    results = rc.write_links(model, self._tds[model], changes,
                                             bool(self.BackupCheck.IsChecked))
                    for row, path in items:
                        ok, detail = results.get(id(self._records[row[COL_IDX]]), (False, "not re-read"))
                        if ok:
                            row[COL_RESULT] = RESULT_PENDING if central else RESULT_WRITTEN
                            row[COL_APPLY] = False
                            written.append((row, path))
                        else:
                            row[COL_RESULT] = "{}: {}".format(RESULT_FAILED, detail)
                            failures += 1
                except Exception as ex:
                    for row, path in items:
                        row[COL_RESULT] = "{}: {}".format(RESULT_FAILED, str(ex).split("\n")[0])
                        failures += 1
                if central and written:
                    failures += self._save_central(model, written)
        finally:
            Mouse.OverrideCursor = None
            self._syncing = False
            self._busy = False
        self._refresh_headers()
        self._update_summary()
        if failures:
            MessageBox.Show("{} links were not written. The reason is in the Result column.".format(failures),
                            "Relink Models", MessageBoxButton.OK, MessageBoxImage.Warning)

    def _save_central(self, model, written):
        """Secondo passo dell'opzione nuova central. Ritorna il numero di righe fallite."""
        changes = [(self._records[row[COL_IDX]], path) for row, path in written]
        try:
            results = rc.save_as_new_central(self._app, model, changes)
        except Exception as ex:
            for row, path in written:
                row[COL_RESULT] = "{}: {}".format(RESULT_CENTRAL_FAILED, str(ex).split("\n")[0])
            return len(written)
        failed = 0
        for row, path in written:
            ok, detail = results.get(id(self._records[row[COL_IDX]]), (False, "not re-read"))
            if ok:
                row[COL_RESULT] = RESULT_CENTRAL
                row[COL_OLD] = path
                row[COL_EXISTS] = YES
            else:
                row[COL_RESULT] = "{}: {}".format(RESULT_CENTRAL_FAILED, detail)
                failed += 1
        return failed

    def OnClose(self, sender, args):
        self._save_settings()
        self.Close()


def show_relink_form(app):
    """Mostra la finestra (modale). `app` e' l'Application di Revit."""
    form = RelinkForm(app)
    form.ShowDialog()
