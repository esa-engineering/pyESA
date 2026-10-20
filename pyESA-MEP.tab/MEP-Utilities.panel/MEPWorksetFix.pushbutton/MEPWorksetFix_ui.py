# -*- coding: utf-8 -*-
"""
MEPWorksetFix_ui.py - finestra XAML di configurazione per MEPWorksetFix.

Per ognuna delle quattro sezioni (meccanici, idraulici, antincendio, elettrici)
l'utente sceglie il workset utente di destinazione, oppure "(do not change)";
sceglie poi l'ambito degli elementi (modello, vista attiva, selezione).
Restituisce un RunConfig allo script chiamante.

Le regole di assegnazione stanno nello script: la finestra riceve solo i testi
descrittivi da mostrare sotto ogni sezione, cosi' non possono divergere.

Compatibile con Rocket mode: nessun documento a livello di modulo, nessuna cache.
"""

import os
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window, MessageBox, GridLength, GridUnitType, Thickness
from System.Windows.Controls import Button, ColumnDefinition, ComboBox, Grid, TextBlock
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode

from pyrevit import DB, script

XAML_FILE_NAME = 'MEPWorksetFix_form.xaml'
CONFIG_SECTION = 'ESA_MEPWorksetFix'

# Voce della tendina che lascia invariati gli elementi della sezione.
NOT_SET_LABEL = '(do not change)'
# Voce vuota delle tendine dei workset da non toccare.
NO_WORKSET_LABEL = '(none)'
# Colonne delle righe create da codice: le stesse della tabella dei workset di
# destinazione nel XAML, cosi' le tendine stanno in colonna.
LABEL_COLUMN_WIDTH = 190
COMBO_COLUMN_WIDTH = 280
REMOVE_GLYPH = u'\u2715'

SECTION_KEYS = ('mechanical', 'plumbing', 'fire', 'electrical')

SCOPE_MODEL = 'model'
SCOPE_VIEW = 'view'
SCOPE_SELECTION = 'selection'

# Parole chiave per proporre un workset alla prima apertura, quando non c'e'
# ancora una scelta salvata. Solo una proposta: l'utente la vede prima di OK.
GUESS_KEYWORDS = {
    'mechanical': ('MECH', 'MECC', 'HVAC'),
    'plumbing': ('PLUMB', 'IDRA', 'IDRI', 'SANIT'),
    'fire': ('FIRE', 'ANTINC', 'SPRINK'),
    'electrical': ('ELEC', 'ELET'),
}


# =============================================================================
# UTILITY
# =============================================================================

def user_worksets(doc):
    """Lista ordinata (nome, id intero) dei workset utente del documento."""
    result = []
    try:
        collector = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
        for workset in collector:
            result.append((workset.Name, workset.Id.IntegerValue))
    except Exception:
        pass
    result.sort(key=lambda pair: pair[0].lower())
    return result


def view_scope_supported(doc):
    """La vista attiva puo' fare da filtro solo se e' una vista grafica del modello."""
    try:
        view = doc.ActiveView
        if view is None or isinstance(view, DB.ViewSheet):
            return False
        if view.IsTemplate:
            return False
        return view.ViewType in (DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan,
                                 DB.ViewType.EngineeringPlan, DB.ViewType.Section,
                                 DB.ViewType.Elevation, DB.ViewType.Detail,
                                 DB.ViewType.ThreeD)
    except Exception:
        return False


class RunConfig(object):
    """Esito della finestra.

    targets: {chiave sezione: id intero del workset, oppure None = non cambiare}
    target_names: {chiave sezione: nome del workset, oppure None}
    excluded: set degli id interi dei workset i cui elementi non si spostano
    excluded_names: nomi degli stessi workset, nell'ordine della finestra
    scope: SCOPE_MODEL / SCOPE_VIEW / SCOPE_SELECTION
    """

    def __init__(self):
        self.targets = {}
        self.target_names = {}
        self.excluded = set()
        self.excluded_names = []
        self.scope = SCOPE_MODEL


# =============================================================================
# FINESTRA
# =============================================================================

class MEPWorksetFixForm(Window):
    """Finestra XAML unica di configurazione."""

    def __init__(self, doc, hints, selection_count):
        self.doc = doc
        self.hints = hints
        self.selection_count = selection_count
        self.result = False
        self.config = RunConfig()

        self._worksets = user_worksets(doc)
        self._workset_ids = dict(self._worksets)
        # Righe dei workset da non toccare: (Grid della riga, ComboBox).
        self._excluded_rows = []
        self._styles = {}

        self._load_xaml()
        self._init_combos()
        self._init_scope()
        self._restore_settings()
        if not self._excluded_rows:
            self._add_excluded_row()

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

        # Gli stili servono alle righe create da codice: il Content passa a
        # self, le risorse restano sulla Window caricata.
        for style_key in ('FieldLabel', 'ComboStyle'):
            try:
                self._styles[style_key] = root.Resources[style_key]
            except Exception:
                pass

        self._find_controls(root)
        self._wire_events()

    def _find_controls(self, root):
        self.combos = {}
        self.hint_blocks = {}
        for key in SECTION_KEYS:
            self.combos[key] = root.FindName('cbo_' + key)
            self.hint_blocks[key] = root.FindName('txt_hint_' + key)

        self.rdo_scope_model = root.FindName('rdo_scope_model')
        self.rdo_scope_view = root.FindName('rdo_scope_view')
        self.rdo_scope_selection = root.FindName('rdo_scope_selection')

        self.pnl_excluded = root.FindName('pnl_excluded')
        self.btn_add_excluded = root.FindName('btn_add_excluded')

        self.btn_ok = root.FindName('btn_ok')
        self.btn_cancel = root.FindName('btn_cancel')

    def _wire_events(self):
        self.btn_add_excluded.Click += self.OnAddExcluded
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel

    def _init_combos(self):
        names = [name for name, _ in self._worksets]
        for key in SECTION_KEYS:
            combo = self.combos[key]
            combo.Items.Add(NOT_SET_LABEL)
            for name in names:
                combo.Items.Add(name)
            combo.SelectedIndex = 0
            self.hint_blocks[key].Text = self.hints.get(key, '')

    def _init_scope(self):
        if not view_scope_supported(self.doc):
            self.rdo_scope_view.IsEnabled = False
            self.rdo_scope_view.Content = "Only elements visible in the active view (not a model view)"
        if self.selection_count:
            self.rdo_scope_selection.Content = "Only the current selection ({} elements)".format(
                self.selection_count)
        else:
            self.rdo_scope_selection.IsEnabled = False
            self.rdo_scope_selection.Content = "Only the current selection (nothing selected)"
        self.rdo_scope_model.IsChecked = True

    # ------------------------------------------------- workset da non toccare

    def _add_excluded_row(self, selected_name=None):
        """Aggiunge una riga etichetta + tendina + pulsante di rimozione."""
        row = Grid()
        row.Margin = Thickness(0, 5, 0, 0)
        for width in (GridLength(LABEL_COLUMN_WIDTH), GridLength(COMBO_COLUMN_WIDTH),
                      GridLength.Auto, GridLength(1, GridUnitType.Star)):
            column = ColumnDefinition()
            column.Width = width
            row.ColumnDefinitions.Add(column)

        label = TextBlock()
        label.Text = "Workset:"
        if 'FieldLabel' in self._styles:
            label.Style = self._styles['FieldLabel']
        Grid.SetColumn(label, 0)
        row.Children.Add(label)

        combo = ComboBox()
        if 'ComboStyle' in self._styles:
            combo.Style = self._styles['ComboStyle']
        combo.ToolTip = "Elements in this workset are never moved"
        combo.Items.Add(NO_WORKSET_LABEL)
        for name, _ in self._worksets:
            combo.Items.Add(name)
        combo.SelectedIndex = 0
        if selected_name and selected_name in combo.Items:
            combo.SelectedItem = selected_name
        Grid.SetColumn(combo, 1)
        row.Children.Add(combo)

        remove = Button()
        remove.Content = REMOVE_GLYPH
        remove.Width = 25
        remove.Margin = Thickness(5, 0, 0, 0)
        remove.ToolTip = "Remove this workset from the list"
        remove.Tag = row
        remove.Click += self.OnRemoveExcluded
        Grid.SetColumn(remove, 2)
        row.Children.Add(remove)

        self.pnl_excluded.Children.Add(row)
        self._excluded_rows.append((row, combo))
        return combo

    def _excluded_names(self):
        """Nomi scelti nelle tendine, senza vuoti ne' doppioni, in ordine."""
        names = []
        for _, combo in self._excluded_rows:
            name = combo.SelectedItem
            if name and name != NO_WORKSET_LABEL and name not in names:
                names.append(name)
        return names

    def OnAddExcluded(self, sender, args):
        combo = self._add_excluded_row()
        combo.Focus()

    def OnRemoveExcluded(self, sender, args):
        row = sender.Tag
        # L'ultima riga non si toglie: si svuota, cosi' la tendina resta.
        if len(self._excluded_rows) <= 1:
            for _, combo in self._excluded_rows:
                combo.SelectedIndex = 0
            return
        self.pnl_excluded.Children.Remove(row)
        self._excluded_rows = [pair for pair in self._excluded_rows if pair[0] is not row]

    def _guess_workset(self, key):
        for name, _ in self._worksets:
            upper = name.upper()
            for keyword in GUESS_KEYWORDS.get(key, ()):
                if keyword in upper:
                    return name
        return None

    # ------------------------------------------------------------ persistenza

    def _save_settings(self):
        try:
            cfg = script.get_config(CONFIG_SECTION)
            for key in SECTION_KEYS:
                setattr(cfg, 'target_' + key, self.combos[key].SelectedItem or NOT_SET_LABEL)
            cfg.excluded_worksets = self._excluded_names()
            # La selezione dipende dal momento: si ricorda solo modello / vista.
            cfg.last_scope_view = bool(self.rdo_scope_view.IsChecked)
            script.save_config()
        except Exception:
            pass

    def _restore_settings(self):
        cfg = None
        try:
            cfg = script.get_config(CONFIG_SECTION)
        except Exception:
            pass

        for key in SECTION_KEYS:
            saved = None
            if cfg is not None:
                try:
                    saved = cfg.get_option('target_' + key, None)
                except Exception:
                    saved = None
            if saved is None:
                saved = self._guess_workset(key)
            if saved and saved in self.combos[key].Items:
                self.combos[key].SelectedItem = saved

        # Workset da non toccare: una riga per ogni nome salvato ancora presente
        # nel modello (i nomi valgono anche fra progetti diversi).
        saved_excluded = []
        if cfg is not None:
            try:
                saved_excluded = cfg.get_option('excluded_worksets', []) or []
            except Exception:
                saved_excluded = []
        for name in saved_excluded:
            if name in self._workset_ids:
                self._add_excluded_row(name)

        try:
            if cfg is not None and cfg.get_option('last_scope_view', False) \
                    and self.rdo_scope_view.IsEnabled:
                self.rdo_scope_view.IsChecked = True
        except Exception:
            pass
        # Una selezione fatta prima di lanciare il comando indica l'intenzione.
        if self.selection_count:
            self.rdo_scope_selection.IsChecked = True

    # ------------------------------------------------------------------- esito

    def _validate_input(self):
        chosen = [key for key in SECTION_KEYS
                  if self.combos[key].SelectedItem not in (None, NOT_SET_LABEL)]
        if not self._worksets:
            return False, "The model has no user worksets."
        if not chosen:
            return False, "Choose a workset for at least one section."
        excluded = set(self._excluded_names())
        clashes = []
        for key in chosen:
            name = self.combos[key].SelectedItem
            if name in excluded and name not in clashes:
                clashes.append(name)
        if clashes:
            return False, ("These worksets are both a target and a workset to leave untouched:\n\n"
                           "{}\n\nRemove them from one of the two lists.").format(
                "\n".join(clashes))
        return True, ""

    def _build_config(self):
        config = RunConfig()
        for key in SECTION_KEYS:
            name = self.combos[key].SelectedItem
            if name in (None, NOT_SET_LABEL):
                config.targets[key] = None
                config.target_names[key] = None
            else:
                config.targets[key] = self._workset_ids[name]
                config.target_names[key] = name
        config.excluded_names = self._excluded_names()
        config.excluded = set(self._workset_ids[name] for name in config.excluded_names)
        if self.rdo_scope_selection.IsChecked:
            config.scope = SCOPE_SELECTION
        elif self.rdo_scope_view.IsChecked:
            config.scope = SCOPE_VIEW
        else:
            config.scope = SCOPE_MODEL
        self.config = config

    def OnOK(self, sender, args):
        is_valid, error_message = self._validate_input()
        if not is_valid:
            MessageBox.Show(error_message, "MEP Workset Fix")
            return

        self._build_config()
        self._save_settings()

        self.result = True
        self.Close()

    def OnCancel(self, sender, args):
        self.result = False
        self.Close()


def show_config_form(doc, hints, selection_count=0):
    """Mostra la finestra e restituisce un RunConfig, oppure None se annullata.

    hints: {chiave sezione: testo che descrive cosa finisce in quella sezione}
    selection_count: elementi selezionati al lancio (0 disabilita l'ambito selezione).
    """
    form = MEPWorksetFixForm(doc, hints, selection_count)
    form.ShowDialog()

    if form.result:
        return form.config
    return None
