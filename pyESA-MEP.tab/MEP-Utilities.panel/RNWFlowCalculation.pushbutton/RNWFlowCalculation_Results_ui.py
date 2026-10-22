# -*- coding: utf-8 -*-
"""
RNWFlowCalculation_Results_ui.py - finestra dei risultati di RNWFlowCalculation.

Tre schede (la finestra si apre su Sections).

General: riepilogo dai terminali (general_summary, Python puro): superficie
coperta = somma delle superfici di raccolta, portata raccolta = somma delle
portate, indice pluviometrico ricavato da ogni terminale come Q / (A x C1 x C2)
(lo calcola la famiglia, lo script non lo conosce), e l'elenco dei sistemi con
superficie e portata. Valori letti dal modello all'apertura.

Sections: una riga per parte di tratto (SectionResult di RNWFlowCalculation_Check) e una
per terminale (FixtureItem: bocchettone con la sua portata), nell'ordine della
rete da monte a valle (colonna "#", da RNWFlowCalculation_Check.order_network). L'utente
sceglie da una tendina la nuova misura fra quelle del tipo di tubazione
(PipeSize del tratto, lette dallo script dai segmenti del tipo); a ogni scelta la
riga ricalcola portata ammissibile, velocita', riempimento ed esito con la nuova
misura (SectionResult.evaluate_at). Non ci sono spunte: si modificano tutte le
righe la cui misura nuova e' diversa da quella del modello (colonna Changed,
righe colorate). Per questo la tendina parte dalla misura suggerita solo nei
tratti da ridimensionare (SectionResult.default_apply: sottodimensionati, De fuori
tabella, velocita' alta), altrimenti dalla misura del modello: nessun tratto
cambia senza una scelta dell'utente o una verifica non superata.
La colonna Sim. slope simula sulle righe orizzontali una pendenza delle tabelle (0,5 - 10 cm/m):
la riga si verifica a quella pendenza, diventa gialla (Simulated) e non cambia il
modello, perche' cambiare la pendenza dei tubi romperebbe i collegamenti.

Fixtures: i terminali collegati ai tratti calcolati, con livello e offset
(FixtureRecord.level, dallo script), famiglia e tipo e i parametri
FIXTURE_PARAMS, tutti modificabili. I dati arrivano dallo script gia'
ridotti a testo (FixtureData: valori di istanza per terminale, valori di tipo per
tipo); un parametro di istanza cambia solo la sua riga, uno di tipo tutte le
righe dello stesso tipo, come in Revit (pending_type). Cambiare tipo ricarica i
valori di tipo della riga. La finestra restituisce le modifiche (FixtureEdits) e
lo script le scrive.

La finestra e' modale e non tocca Revit: si chiude restituendo allo script
un'azione (kind):
  * "apply": misure nuove (changes: tratto, PipeSize) e modifiche ai terminali
    (fixture_edits);
  * "recalc": le sole modifiche ai terminali; lo script le scrive, ricalcola la
    rete con le portate aggiornate da Revit e riapre la finestra con
    ResultsState.carry(), che conserva le misure scelte a mano;
  * "show": gli elementi da selezionare (pulsante di una riga); lo script li
    seleziona, lascia navigare la vista e riapre la finestra con lo stesso
    ResultsState, cioe' con misure, modifiche, ricerca, filtro e scheda di prima;
  * None: chiusa senza modifiche.

Export apre la finestra di impostazione (RNWFlowCalculation_Export_ui: formato
Excel o PDF, colonne di ogni scheda), poi build_report() riduce le tre schede a
un report neutro (RNWFlowCalculation_Report: tutte le righe nell'ordine della
rete, filtro e ricerca non contano) che RNWFlowCalculation_Xlsx o
RNWFlowCalculation_Pdf scrivono; titolo e impostazioni da ExportInfo (script).

Griglie legate a System.Data.DataTable come in MEPQTO: il binding a due vie lo
gestisce .NET, senza oggetti INotifyPropertyChanged in IronPython. Le tendine
hanno le voci nella colonna Options / TypeOptions (oggetto List[str]). Gli id
(terminali, tipi) sono tenuti come testo: in Revit 2026 sono Int64 e come chiavi
di dizionario e' piu' sicuro confrontare stringhe.

Compatibile con Rocket mode: nessun documento o output a livello di modulo.
"""

import os
import datetime
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Data')

import System
from System import Action
from System.Diagnostics import Process, ProcessStartInfo
from Microsoft.Win32 import SaveFileDialog
from System.Collections.Generic import List
from System.Data import DataTable, DataRowVersion
from System.Windows import (Window, MessageBox, MessageBoxButton, MessageBoxResult,
                            RoutedEventHandler)
from System.Windows.Controls import DataGridEditingUnit, Button
from System.Windows.Controls.Primitives import ButtonBase
from System.Windows.Markup import XamlReader
from System.Windows.Threading import DispatcherPriority
from System.IO import FileStream, FileMode

from pyrevit import script

import RNWFlowCalculation_Check as rc
import RNWFlowCalculation_Report as rp
import RNWFlowCalculation_Xlsx as rx
import RNWFlowCalculation_Pdf as rpdf
import RNWFlowCalculation_Export_ui as rexp

XAML_FILE_NAME = 'RNWFlowCalculation_Results.xaml'
CONFIG_SECTION = 'ESA_RNWFlowCalculation'
WINDOW_TITLE = 'Rainwater Sections'

MAX_LINES_SHOWN = 10

SEVERITY = {
    rc.STATUS_UNDERSIZED: "error",
    rc.STATUS_BEYOND: "error",
    rc.STATUS_NOT_IN_TABLE: "warning",
    rc.STATUS_LOW_SLOPE: "warning",
    rc.STATUS_HIGH_VELOCITY: "warning",
    rc.STATUS_LOW_VELOCITY: "warning",
    rc.STATUS_OVERSIZED: "info",
    rc.STATUS_OK: "",
    rc.STATUS_NO_FLOW: "none",
}

# Esiti della misura nuova che non richiedono conferma esplicita
NEW_CHECK_PASS = (rc.STATUS_OK, rc.STATUS_NO_FLOW)

KIND_APPLY = "apply"
KIND_SHOW = "show"
KIND_RECALC = "recalc"

# Ordine delle schede nella XAML; la finestra si apre su Sections.
TAB_GENERAL = 0
TAB_SECTIONS = 1
TAB_FIXTURES = 2

# Indice pluviometrico ricavato dai terminali: valori entro questa tolleranza
# relativa contano come un valore solo (arrotondamenti dei parametri).
INTENSITY_TOLERANCE = 0.005
# l/(s m2) -> mm/h
INTENSITY_TO_MMH = 3600.0

# Tendina della pendenza simulata: prima voce = pendenza del modello (nessuna
# simulazione), poi le pendenze di RNWFlowCalculation_Check.SIMULATED_SLOPES [cm/m = %].
SLOPE_MODEL = "Model"


def slope_label(slope):
    return "{0:.1f}".format(slope)

# Parametri dei terminali nella scheda Fixtures: (nome del parametro in Revit,
# prefisso delle colonne della DataTable). Il codice elemento compare anche nella
# scheda Sections e nell'export.
PARAM_AREA = "u_PLU_Area_1"
PARAM_C1 = "C1"
PARAM_C2 = "C2"
PARAM_CODE = "e_MEC_ElementCode_1"
FIXTURE_PARAMS = [(PARAM_AREA, "Area"), (PARAM_C1, "C1"), (PARAM_C2, "C2"), (PARAM_CODE, "Code")]

WHERE_INSTANCE = "Instance"
WHERE_TYPE = "Type"
WHERE_MISSING = "Missing"

VALUE_NUMBER = "number"
VALUE_INTEGER = "integer"
VALUE_TEXT = "text"


def fmt_de(value):
    if value is None:
        return ""
    return rc.fmt_mm(value)


def escape_like(text):
    """Escape di un testo per RowFilter LIKE (stesse regole di MEPQTO)."""
    out = []
    for ch in text:
        if ch in "[]*%":
            out.append("[" + ch + "]")
        elif ch == "'":
            out.append("''")
        else:
            out.append(ch)
    return "".join(out)


def _value(number):
    return number if number is not None else System.DBNull.Value


def _pct(fraction):
    return fraction * 100.0 if fraction is not None else None


def _text(value):
    if value is None or value is System.DBNull.Value:
        return ""
    return str(value)


def _number(value):
    if value is None or value is System.DBNull.Value:
        return None
    return float(value)


def parse_decimal(text):
    """Numero da un testo con virgola o punto decimale; None se non valido.

    Con entrambi i separatori quello decimale e' l'ultimo (1.234,5 e 1,234.5).
    """
    if text is None:
        return None
    cleaned = str(text).strip().replace(" ", "")
    if not cleaned:
        return None
    if "," in cleaned and "." in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    else:
        cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except (ValueError, TypeError):
        return None


def number_text(value):
    """Numero come testo, punto decimale, al massimo 6 decimali, senza zeri finali.
    Lo script lo usa anche per i valori letti dal modello."""
    text = "{0:.6f}".format(round(value, 6)).rstrip("0").rstrip(".")
    return "0" if text in ("", "-0") else text


def values_differ(original, current, kind):
    """True se il valore e' cambiato: i numeri si confrontano come numeri, cosi'
    "0,90" non risulta diverso da "0.9"."""
    if kind in (VALUE_NUMBER, VALUE_INTEGER):
        a = parse_decimal(original)
        b = parse_decimal(current)
        if a is None or b is None:
            return (a is None) != (b is None)
        return abs(a - b) > 1e-9 * max(1.0, abs(a), abs(b))
    return (original or "").strip() != (current or "").strip()


# =============================================================================
# DATI SCAMBIATI CON LO SCRIPT
# =============================================================================

class ResultsState(object):
    """Stato delle tabelle che sopravvive alla chiusura della finestra."""

    def __init__(self):
        self.table = None
        self.fixture_table = None
        self.by_key = {}
        self.defaults = {}          # chiave di riga -> misura proposta all'apertura
        self.pending_type = {}      # (id tipo, parametro) -> testo
        self.preset = {}            # SectionResult.key -> misura scelta a mano
        self.preset_slope = {}      # SectionResult.key -> pendenza simulata
        self.search = ""
        self.only_action = None
        self.last_key = None
        self.tab = TAB_SECTIONS

    def carry(self):
        """Stato per la finestra riaperta dopo un ricalcolo: tabelle nuove (portate e
        terminali sono cambiati), ma restano le misure scelte a mano (per
        SectionResult.key, che non cambia), la ricerca, il filtro e la scheda."""
        new = ResultsState()
        if self.table is not None:
            for row in self.table.Rows:
                entry = self.by_key.get(row["Key"])
                if entry is None or getattr(entry, "is_fixture", False):
                    continue
                label = _text(row["NewSize"])
                if label != self.defaults.get(row["Key"], ""):
                    new.preset[entry.key] = label
                slope = _text(row["SimSlope"])
                if slope and slope != SLOPE_MODEL:
                    new.preset_slope[entry.key] = slope
        new.search = self.search
        new.only_action = self.only_action
        new.tab = self.tab
        return new


class ExportInfo(object):
    """Dati dello script per l'esportazione: titolo, impostazioni, cartella e
    nome proposto (l'estensione si cambia secondo il formato scelto)."""

    def __init__(self, title, settings_rows, folder, file_name):
        self.title = title
        self.settings_rows = list(settings_rows)
        self.folder = folder
        self.file_name = file_name


class ParamValue(object):
    """Valore di un parametro di un terminale, gia' convertito in testo dallo
    script (numeri nelle unita' di visualizzazione, area in m2)."""

    def __init__(self, text, where, kind=VALUE_TEXT, editable=True):
        self.text = text or ""
        self.where = where
        self.kind = kind
        self.editable = editable


MISSING = ParamValue("", WHERE_MISSING, VALUE_TEXT, False)


class FixtureRecord(object):
    """Un terminale: i suoi valori di istanza ({parametro: ParamValue}) e il tipo.
    level = "Livello, +offset [mm]", level_sort = quota assoluta [mm] (ordinamento)."""

    def __init__(self, element_id, system, flow, type_id, instance_values, notes=None,
                 level="", level_sort=None):
        self.id = str(element_id)
        self.system = system
        self.flow = flow
        self.type_id = str(type_id)
        self.instance_values = dict(instance_values)
        self.notes = list(notes or [])
        self.level = level or ""
        self.level_sort = level_sort


class FixtureData(object):
    """Terminali (in ordine di rete), tipi disponibili [(id, "Famiglia : Tipo")] e
    valori di tipo {id tipo: {parametro: ParamValue}}."""

    def __init__(self, records, types, type_values):
        self.records = list(records)
        self.types = [(str(type_id), label) for type_id, label in types]
        self.type_values = dict((str(type_id), values) for type_id, values in type_values.items())


class FixtureEdits(object):
    """Modifiche da scrivere: tipo per terminale, parametri di tipo, di istanza."""

    def __init__(self):
        self.types = {}             # id terminale -> id tipo
        self.type_params = {}       # (id tipo, parametro) -> testo
        self.instance = {}          # (id terminale, parametro) -> testo

    @property
    def count(self):
        return len(self.types) + len(self.type_params) + len(self.instance)


# =============================================================================
# SCHEDA GENERAL (Python puro)
# =============================================================================

class SystemSummary(object):
    """Un sistema calcolato: terminali, superficie [m2] e portata [l/s] raccolte."""

    def __init__(self, name):
        self.name = name
        self.fixtures = 0
        self.area = None
        self.flow = None


class GeneralSummary(object):
    """Dati della scheda General. Valori letti dal modello all'apertura della
    finestra: le modifiche ai terminali non ancora scritte non contano, come la
    portata Q della scheda Fixtures."""

    def __init__(self):
        self.systems = []           # SystemSummary, per nome
        self.fixtures = 0
        self.area = None
        self.flow = None
        self.no_area = 0            # terminali senza superficie
        self.no_flow = 0            # terminali senza portata leggibile
        self.intensities = []       # l/(s m2), una per terminale con Q, A, C1, C2 > 0


def _model_number(record, type_values, name):
    """Valore numerico del parametro letto dal modello (istanza, poi tipo); None se
    manca, e' vuoto o e' un testo."""
    value = record.instance_values.get(name)
    if value is None:
        value = type_values.get(record.type_id, {}).get(name)
    if value is None or value.kind == VALUE_TEXT:
        return None
    return parse_decimal(value.text)


def general_summary(sections, fixture_data):
    """Superficie e portata per sistema e totali, dai terminali (FixtureData).

    L'indice pluviometrico non e' un dato dello script (lo calcola la famiglia):
    si ricava da ogni terminale come Q / (A x C1 x C2), cioe' supponendo che la
    famiglia calcoli Q = i x A x C1 x C2.
    """
    summary = GeneralSummary()
    by_name = {}

    def system(name):
        item = by_name.get(name)
        if item is None:
            item = by_name[name] = SystemSummary(name)
        return item

    # Anche i sistemi senza terminali compaiono nell'elenco
    for entry in sections:
        system(entry.system)
    for record in fixture_data.records:
        item = system(record.system)
        item.fixtures += 1
        summary.fixtures += 1
        area = _model_number(record, fixture_data.type_values, PARAM_AREA)
        if area is None:
            summary.no_area += 1
        else:
            item.area = (item.area or 0.0) + area
            summary.area = (summary.area or 0.0) + area
        flow = record.flow
        if flow is None:
            summary.no_flow += 1
        else:
            item.flow = (item.flow or 0.0) + flow
            summary.flow = (summary.flow or 0.0) + flow
        c1 = _model_number(record, fixture_data.type_values, PARAM_C1)
        c2 = _model_number(record, fixture_data.type_values, PARAM_C2)
        if flow and area and c1 and c2 and flow > 0 and area > 0 and c1 > 0 and c2 > 0:
            summary.intensities.append(flow / (area * c1 * c2))
    summary.systems = sorted(by_name.values(), key=lambda s: (s.name or "").lower())
    return summary


def intensity_texts(values):
    """(testo in l/(s m2), testo in mm/h, True se i terminali danno valori diversi)."""
    if not values:
        return "-", "-", False
    low = min(values)
    high = max(values)
    if high - low <= INTENSITY_TOLERANCE * high:
        mean = sum(values) / len(values)
        return ("{0:.4f}".format(mean), "{0:.1f}".format(mean * INTENSITY_TO_MMH), False)
    return ("{0:.4f} - {1:.4f}".format(low, high),
            "{0:.1f} - {1:.1f}".format(low * INTENSITY_TO_MMH, high * INTENSITY_TO_MMH), True)


def number_or_dash(value, digits):
    if value is None:
        return "-"
    return "{0:.{1}f}".format(float(value), digits)


def export_severity(status):
    """Gravita' di un esito per il report (colore della cella in Excel e PDF)."""
    if status == rc.STATUS_OK:
        return rp.SEV_PASS
    severity = SEVERITY.get(status, "")
    return severity if severity in (rp.SEV_ERROR, rp.SEV_WARNING, rp.SEV_INFO) else None


def number_or_text(value):
    """Numero se il testo lo e' (De e DN del modello, De suggerito), altrimenti
    il testo ("110/125", "-"); None se vuoto."""
    text = _text(value).strip()
    if not text:
        return None
    number = parse_decimal(text)
    return number if number is not None else text


def general_notes(summary):
    """Note della scheda General (finestra ed export)."""
    notes = ["{0} system(s), {1} fixture(s).".format(len(summary.systems), summary.fixtures)]
    if summary.no_area:
        notes.append("{0} fixture(s) without a collection area ({1}) are not in the "
                     "covered area.".format(summary.no_area, PARAM_AREA))
    if summary.no_flow:
        notes.append("{0} fixture(s) without a readable flow are not in the collected "
                     "flow.".format(summary.no_flow))
    used = len(summary.intensities)
    if used:
        notes.append("Rainfall intensity derived from {0} fixture(s).".format(used))
        if used < summary.fixtures:
            notes.append("{0} fixture(s) with no flow, collection area, C1 or C2 are "
                         "not used for it.".format(summary.fixtures - used))
        _per_m2, _per_hour, varies = intensity_texts(summary.intensities)
        if varies:
            notes.append("The fixtures give different values: check the rainfall "
                         "intensity in the fixture families.")
    elif summary.fixtures:
        notes.append("The rainfall intensity cannot be derived: no fixture has flow, "
                     "collection area, C1 and C2 all above zero.")
    return notes


def open_file(path):
    """Apre il file con il programma predefinito. Serve UseShellExecute: in
    .NET 8 (Revit 2025 e successivi) Process.Start(path) non apre i documenti."""
    info = ProcessStartInfo(path)
    info.UseShellExecute = True
    Process.Start(info)


def is_change(section, size):
    """True se la misura cambia il tratto: diversa da quella del modello, o tratto
    con tubi di misure diverse (la misura scelta li uniforma)."""
    if size is None:
        return False
    return len(section.model_des) != 1 or not rc.od_in(size.od_mm, section.model_des)


# =============================================================================
# FINESTRA
# =============================================================================

class ResultsForm(Window):
    """Finestra dei tratti e dei terminali."""

    def __init__(self, sections, settings_text, state=None, export_info=None,
                 fixture_data=None):
        self._export_info = export_info
        self.sections = list(sections)
        self.changes = None
        self.fixture_edits = None
        self.kind = None
        self.show_ids = []
        self.state = state if state is not None else ResultsState()
        self._by_key = self.state.by_key
        self._loading = True
        self._column_handler = self.OnColumnChanged
        self._fixture_changing = self.OnFixtureColumnChanging
        self._fixture_changed = self.OnFixtureColumnChanged

        data = fixture_data if fixture_data is not None else FixtureData([], [], {})
        self._records = dict((record.id, record) for record in data.records)
        self._fixture_records = data.records
        self._types = data.types
        self._type_values = data.type_values
        self._type_label = dict(data.types)
        self._label_type = dict((label, type_id) for type_id, label in data.types)
        self._prefix_param = dict((prefix, name) for name, prefix in FIXTURE_PARAMS)

        self._load_xaml()
        self.txt_settings.Text = settings_text
        self._build_general(data)
        if self.state.table is None:
            self._build_table()
            self._build_fixture_table()
            if self.state.only_action is None:
                self._restore_settings()
            else:
                self.chk_only_action.IsChecked = bool(self.state.only_action)
            self.txt_search.Text = self.state.search or ""
        else:
            self.table = self.state.table
            self.ftable = self.state.fixture_table
            self.dg_sections.ItemsSource = self.table.DefaultView
            self.dg_fixtures.ItemsSource = self.ftable.DefaultView
            self.chk_only_action.IsChecked = bool(self.state.only_action)
            self.txt_search.Text = self.state.search or ""
        self.table.ColumnChanged += self._column_handler
        self.ftable.ColumnChanging += self._fixture_changing
        self.ftable.ColumnChanged += self._fixture_changed
        self._fixture_section_rows = {}
        self._fixture_ids = {}
        for row in self.table.Rows:
            entry = self._by_key[row["Key"]]
            if getattr(entry, "is_fixture", False):
                self._fixture_section_rows[str(entry.id)] = row
                self._fixture_ids[str(entry.id)] = entry.id
        self.btn_recalc.IsEnabled = self.ftable.Rows.Count > 0
        self._loading = False
        self._apply_filter()
        self._update_fixture_count()

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

        for name in ('txt_settings', 'tabs', 'txt_search', 'btn_clear_search',
                     'chk_only_action', 'txt_count', 'dg_sections', 'txt_fixture_count',
                     'dg_fixtures', 'txt_total_area', 'txt_intensity', 'txt_intensity_mmh',
                     'txt_total_flow', 'txt_general_notes', 'dg_systems',
                     'btn_export', 'btn_recalc', 'btn_ok', 'btn_cancel'):
            setattr(self, name, root.FindName(name))

        self.txt_search.TextChanged += self.OnFilterChanged
        self.btn_clear_search.Click += self.OnClearSearch
        self.chk_only_action.Checked += self.OnFilterChanged
        self.chk_only_action.Unchecked += self.OnFilterChanged
        self.btn_export.Click += self.OnExport
        self.btn_export.IsEnabled = self._export_info is not None
        self.btn_recalc.Click += self.OnRecalc
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel
        # Pulsante "seleziona" di ogni riga: un solo gestore per griglia.
        self.dg_sections.AddHandler(ButtonBase.ClickEvent,
                                    RoutedEventHandler(self.OnSectionButton))
        self.dg_fixtures.AddHandler(ButtonBase.ClickEvent,
                                    RoutedEventHandler(self.OnFixtureButton))
        self.dg_fixtures.BeginningEdit += self.OnFixtureBeginningEdit
        self.Loaded += self.OnWindowLoaded
        self.Closing += self.OnWindowClosing
        self.Closed += self.OnWindowClosed

    def OnWindowLoaded(self, sender, args):
        """Riporta la scheda e la riga da cui si era partiti per navigare."""
        try:
            self.tabs.SelectedIndex = self.state.tab if self.state.tab is not None \
                else TAB_SECTIONS
        except Exception:
            pass
        key = self.state.last_key
        if key is None:
            return
        table, grid = (self.ftable, self.dg_fixtures) if key.startswith("F") \
            else (self.table, self.dg_sections)
        for item in table.DefaultView:
            if item["Key"] == key:
                try:
                    grid.ScrollIntoView(item)
                except Exception:
                    pass
                break

    def OnWindowClosing(self, sender, args):
        """Chiudere senza applicare perde le modifiche ai terminali: si chiede."""
        if self.kind is not None:
            return
        edits = self._fixture_edits()
        if not edits.count:
            return
        answer = MessageBox.Show(
            "{0} fixture change(s) are not written to the model.\n\nClose and discard "
            "them?".format(edits.count), WINDOW_TITLE, MessageBoxButton.YesNo)
        if answer != MessageBoxResult.Yes:
            args.Cancel = True

    def OnWindowClosed(self, sender, args):
        try:
            self.table.ColumnChanged -= self._column_handler
            self.ftable.ColumnChanging -= self._fixture_changing
            self.ftable.ColumnChanged -= self._fixture_changed
        except Exception:
            pass
        self.state.table = self.table
        self.state.fixture_table = self.ftable
        self.state.search = self.txt_search.Text or ""
        self.state.only_action = bool(self.chk_only_action.IsChecked)
        try:
            self.state.tab = self.tabs.SelectedIndex
        except Exception:
            pass

    # ------------------------------------------------------- scheda General

    def _build_general(self, fixture_data):
        """Riepilogo: superficie, indice pluviometrico, portata e sistemi. Si
        ricostruisce a ogni apertura: dipende solo dai dati letti dallo script."""
        summary = general_summary(self.sections, fixture_data)
        self._summary = summary
        self.txt_total_area.Text = number_or_dash(summary.area, 2)
        self.txt_total_flow.Text = number_or_dash(summary.flow, 2)
        per_m2, per_hour, _varies = intensity_texts(summary.intensities)
        self.txt_intensity.Text = per_m2
        self.txt_intensity_mmh.Text = per_hour
        self.txt_general_notes.Text = " ".join(general_notes(summary))

        table = DataTable("Systems")
        for name, kind in (("System", System.String), ("Fixtures", System.Int32),
                           ("Area", System.Double), ("Flow", System.Double)):
            table.Columns.Add(name, kind)
        for item in summary.systems:
            row = table.NewRow()
            row["System"] = item.name
            row["Fixtures"] = item.fixtures
            row["Area"] = _value(item.area)
            row["Flow"] = _value(item.flow)
            table.Rows.Add(row)
        table.AcceptChanges()
        self.dg_systems.ItemsSource = table.DefaultView

    # ------------------------------------------------------- tabella Sections

    def _build_table(self):
        table = DataTable("Sections")
        columns = [
            ("Key", System.String), ("System", System.String), ("Code", System.String),
            ("PipeType", System.String),
            ("Section", System.String), ("SectionSort", System.Double),
            ("Pipes", System.Int32), ("Run", System.String), ("Slope", System.Double),
            ("Flow", System.Double), ("ModelDn", System.String), ("ModelDe", System.String),
            ("Capacity", System.Double), ("Use", System.Double), ("VModel", System.Double),
            ("FillModel", System.Double), ("Suggested", System.String),
            ("Options", System.Object), ("HasSizes", System.Boolean),
            ("NewSize", System.String), ("NewDeSort", System.Double),
            ("QNew", System.Double), ("VNew", System.Double), ("FillNew", System.Double),
            ("CheckNew", System.String), ("SeverityNew", System.String),
            ("Changed", System.Boolean),
            ("SimSlope", System.String), ("SlopeOptions", System.Object),
            ("CanSimulate", System.Boolean), ("Simulated", System.Boolean),
            ("Status", System.String), ("Notes", System.String), ("Severity", System.String),
            ("Action", System.Boolean), ("Order", System.Int32), ("Kind", System.String),
            ("FedBy", System.String),
        ]
        for name, kind in columns:
            table.Columns.Add(name, kind)

        slope_options = List[System.String]()
        slope_options.Add(SLOPE_MODEL)
        for slope in rc.SIMULATED_SLOPES:
            slope_options.Add(slope_label(slope))
        no_options = List[System.String]()

        for index, section in enumerate(self.sections):
            key = str(index)
            self._by_key[key] = section
            if getattr(section, "is_fixture", False):
                table.Rows.Add(self._fixture_row(table, key, index, section))
                continue
            options = List[System.String]()
            for size in section.sizes or []:
                options.Add(size.label)
            # Senza spunte si applica ogni misura diversa da quella del modello:
            # la proposta parte dal suggerimento solo dove il tratto va cambiato.
            if section.default_apply:
                default = section.suggested_size
            else:
                default = section.model_size
            default_label = default.label if default is not None else ""
            self.state.defaults[key] = default_label
            preset = self.state.preset.get(section.key)
            chosen = preset if preset is not None and \
                (preset == "" or preset in list(options)) else default_label

            row = table.NewRow()
            row["Key"] = key
            row["System"] = section.system
            row["Code"] = getattr(section, "element_code", "") or ""
            row["PipeType"] = getattr(section, "pipe_type", "") or ""
            row["Section"] = section.section if section.section else "-"
            row["SectionSort"] = section.section_sort()
            row["Pipes"] = len(section.pipe_ids)
            row["Run"] = section.run
            row["Slope"] = _value(section.min_slope)
            row["Flow"] = section.flow
            row["ModelDn"] = section.model_dn_label
            row["ModelDe"] = section.model_de_label
            row["Capacity"] = _value(section.capacity)
            row["Use"] = _value(section.use_pct)
            row["VModel"] = _value(section.v_min)
            row["FillModel"] = _value(_pct(section.fill_max))
            row["Suggested"] = fmt_de(section.suggested_de)
            row["Options"] = options
            row["HasSizes"] = options.Count > 0
            row["NewSize"] = chosen
            # La pendenza si simula solo sui tratti orizzontali
            can_simulate = section.run == rc.ORIENT_HORIZONTAL
            row["CanSimulate"] = can_simulate
            row["SlopeOptions"] = slope_options if can_simulate else no_options
            preset_slope = self.state.preset_slope.get(section.key)
            row["SimSlope"] = preset_slope if can_simulate and preset_slope in list(slope_options) \
                else (SLOPE_MODEL if can_simulate else "")
            row["Status"] = section.status
            row["Notes"] = "; ".join(section.notes)
            row["Severity"] = SEVERITY.get(section.status, "")
            row["Action"] = bool(section.needs_action)
            row["Order"] = index + 1
            row["Kind"] = "section"
            self._fill_new(row, section)
            table.Rows.Add(row)

        self._fill_fed_by(table)
        table.AcceptChanges()
        self.table = table
        self.dg_sections.ItemsSource = table.DefaultView

    def _fill_fed_by(self, table):
        """Colonna "Fed by": le righe che confluiscono direttamente in ogni riga,
        da RNWFlowCalculation_Check.order_network, in forma breve: "Fixture <id>" per un
        terminale, "<Section>V" / "<Section>O" per la parte verticale / orizzontale
        di una Section, unite da "+" (per esempio "1V+2O"); "-" se a monte non c'e'
        nulla."""
        def label(entry):
            if getattr(entry, "is_fixture", False):
                return "Fixture {0}".format(entry.id)
            letter = "V" if entry.run == rc.ORIENT_STACK else "O"
            if entry.section:
                return "{0}{1}".format(entry.section, letter)
            return "pipe {0} {1}".format(entry.pipe_ids[0], letter)

        for row in table.Rows:
            entry = self._by_key[row["Key"]]
            sources = getattr(entry, "upstream", None) or []
            row["FedBy"] = "+".join(label(source) for source in sources) or "-"

    def _fixture_row(self, table, key, index, item):
        """Riga di un terminale: portata e attacco, nessuna misura da scegliere."""
        row = table.NewRow()
        row["Key"] = key
        row["System"] = item.system
        row["Code"] = getattr(item, "element_code", "") or ""
        row["Section"] = "-"
        row["SectionSort"] = item.section_sort()
        row["Run"] = item.run
        row["Flow"] = _value(item.flow)
        row["ModelDn"] = fmt_de(item.dn_mm)
        row["ModelDe"] = ""
        row["Suggested"] = ""
        row["Options"] = List[System.String]()
        row["HasSizes"] = False
        row["NewSize"] = ""
        row["CheckNew"] = ""
        row["SeverityNew"] = ""
        row["Changed"] = False
        row["SimSlope"] = ""
        row["SlopeOptions"] = List[System.String]()
        row["CanSimulate"] = False
        row["Simulated"] = False
        row["Status"] = item.status
        row["Notes"] = "; ".join([item.label] + item.notes)
        row["Severity"] = "fixture"
        row["Action"] = False
        row["Order"] = index + 1
        row["Kind"] = "fixture"
        self.state.defaults[key] = ""
        return row

    # ------------------------------------------------------- misura nuova

    def _size_of(self, row, section):
        label = _text(row["NewSize"])
        for size in section.sizes or []:
            if size.label == label:
                return size
        return None

    @staticmethod
    def _sim_slope(row):
        """Pendenza simulata della riga [cm/m]; None = pendenza del modello."""
        text = _text(row["SimSlope"])
        if not text or text == SLOPE_MODEL:
            return None
        return parse_decimal(text)

    def _fill_new(self, row, section):
        """Colonne della misura nuova: Qmax, v, riempimento, esito, cambio.

        Con una pendenza simulata la riga si verifica a quella pendenza, il De
        suggerito segue la simulazione e la riga non cambia il modello
        (Simulated): cambiare la pendenza nel modello sposterebbe i tubi e
        romperebbe quasi certamente i collegamenti, quindi la si lascia all'utente.
        """
        size = self._size_of(row, section)
        slope = self._sim_slope(row)
        simulated = slope is not None
        row["Simulated"] = simulated
        row["Changed"] = is_change(section, size) and not simulated
        suggested = section.suggested_at(slope) if simulated else section.suggested_de
        row["Suggested"] = fmt_de(suggested)
        if size is None:
            for name in ("NewDeSort", "QNew", "VNew", "FillNew"):
                row[name] = System.DBNull.Value
            row["CheckNew"] = ""
            row["SeverityNew"] = ""
            return
        result = section.evaluate_at(size.od_mm, slope)
        row["NewDeSort"] = size.od_mm
        row["QNew"] = _value(result.capacity)
        row["VNew"] = _value(result.velocity)
        row["FillNew"] = _value(_pct(result.fill_actual))
        row["CheckNew"] = result.status
        row["SeverityNew"] = SEVERITY.get(result.status, "")

    # ------------------------------------------------------ tabella Fixtures

    def _build_fixture_table(self):
        table = DataTable("Fixtures")
        columns = [
            ("Key", System.String), ("Order", System.Int32), ("System", System.String),
            ("Level", System.String), ("LevelSort", System.Double),
            ("FixtureId", System.String), ("TypeId", System.String),
            ("TypeLabel", System.String), ("TypeOptions", System.Object),
            ("Flow", System.Double), ("Notes", System.String), ("Dirty", System.Boolean),
        ]
        for _name, prefix in FIXTURE_PARAMS:
            columns.extend([(prefix, System.String), (prefix + "From", System.String),
                            (prefix + "Where", System.String), (prefix + "Edit", System.Boolean),
                            (prefix + "Num", System.Boolean)])
        for name, kind in columns:
            table.Columns.Add(name, kind)

        options = List[System.String]()
        for _type_id, label in self._types:
            options.Add(label)
        order_of = {}
        for key, entry in self._by_key.items():
            if getattr(entry, "is_fixture", False):
                order_of[str(entry.id)] = int(key) + 1

        for record in self._fixture_records:
            row = table.NewRow()
            row["Key"] = "F" + record.id
            row["Order"] = order_of.get(record.id, 0)
            row["System"] = record.system
            row["Level"] = record.level or "-"
            row["LevelSort"] = _value(record.level_sort)
            row["FixtureId"] = record.id
            row["TypeId"] = record.type_id
            row["TypeLabel"] = self._type_label.get(record.type_id, "")
            row["TypeOptions"] = options
            row["Flow"] = _value(record.flow)
            row["Notes"] = "; ".join(record.notes)
            row["Dirty"] = False
            self._fill_fixture_params(row, record, record.type_id)
            table.Rows.Add(row)
        table.AcceptChanges()
        self.ftable = table
        self.dg_fixtures.ItemsSource = table.DefaultView

    def _resolve(self, record, type_id, name):
        """(ParamValue, testo mostrato) del parametro per il terminale con quel
        tipo: prima l'istanza, poi il tipo (con le modifiche di tipo in sospeso)."""
        value = record.instance_values.get(name)
        if value is not None:
            return value, value.text
        typed = self._type_values.get(type_id, {}).get(name)
        if typed is not None:
            return typed, self.state.pending_type.get((type_id, name), typed.text)
        return MISSING, ""

    @staticmethod
    def _from_label(value):
        if value.where == WHERE_MISSING:
            return "missing"
        if not value.editable:
            return "read-only " + value.where.lower()
        return value.where.lower()

    def _fill_fixture_params(self, row, record, type_id, keep_instance=False):
        """Valori della riga per il tipo dato. Con keep_instance (cambio di tipo)
        i valori di istanza gia' modificati restano."""
        for name, prefix in FIXTURE_PARAMS:
            value, text = self._resolve(record, type_id, name)
            if keep_instance and value.where == WHERE_INSTANCE:
                text = _text(row[prefix])
            row[prefix] = text
            row[prefix + "Where"] = value.where
            row[prefix + "Edit"] = bool(value.editable)
            row[prefix + "Num"] = value.kind != VALUE_TEXT
            row[prefix + "From"] = self._from_label(value)

    def _original(self, row, name, prefix):
        """ParamValue letto dal modello per la cella, None se non c'e'."""
        record = self._records[_text(row["FixtureId"])]
        where = _text(row[prefix + "Where"])
        if where == WHERE_INSTANCE:
            return record.instance_values.get(name)
        if where == WHERE_TYPE:
            return self._type_values.get(_text(row["TypeId"]), {}).get(name)
        return None

    def _row_dirty(self, row):
        record = self._records[_text(row["FixtureId"])]
        if _text(row["TypeId"]) != record.type_id:
            return True
        for name, prefix in FIXTURE_PARAMS:
            original = self._original(row, name, prefix)
            if original is not None and values_differ(original.text, _text(row[prefix]),
                                                      original.kind):
                return True
        return False

    def _sync_code(self, row):
        """Il codice elemento di un terminale compare anche nella scheda Sections."""
        target = self._fixture_section_rows.get(_text(row["FixtureId"]))
        if target is None:
            return
        code = _text(row["Code"])
        if _text(target["Code"]) != code:
            target["Code"] = code

    def OnFixtureBeginningEdit(self, sender, args):
        """Le celle di un parametro mancante o in sola lettura non si modificano."""
        prefix = args.Column.SortMemberPath
        if prefix not in self._prefix_param:
            return
        try:
            editable = bool(args.Row.Item[prefix + "Edit"])
        except Exception:
            return
        if not editable:
            args.Cancel = True

    def OnFixtureColumnChanging(self, sender, args):
        """Valida i valori numerici alla conferma della cella (virgola o punto)."""
        if self._loading:
            return
        prefix = args.Column.ColumnName
        if prefix not in self._prefix_param:
            return
        row = args.Row
        if not row[prefix + "Num"]:
            return
        text = _text(args.ProposedValue).strip()
        number = parse_decimal(text)
        problem = None
        if number is None:
            problem = "'{0}' is not a number.".format(text)
        elif number < 0:
            problem = "{0} cannot be negative.".format(self._prefix_param[prefix])
        else:
            original = self._original(row, self._prefix_param[prefix], prefix)
            if original is not None and original.kind == VALUE_INTEGER \
                    and abs(number - round(number)) > 1e-9:
                problem = "{0} takes whole numbers only.".format(self._prefix_param[prefix])
        if problem is None:
            # Testo normalizzato (punto decimale), come quelli letti dal modello
            args.ProposedValue = number_text(number)
            return
        args.ProposedValue = row[prefix]
        message = "{0}\nThe value is not changed.".format(problem)
        self._defer(lambda: MessageBox.Show(message, WINDOW_TITLE))

    def OnFixtureColumnChanged(self, sender, args):
        if self._loading:
            return
        column = args.Column.ColumnName
        row = args.Row
        record = self._records.get(_text(row["FixtureId"]))
        if record is None:
            return
        touched = []
        self._loading = True
        try:
            if column == "TypeLabel":
                type_id = self._label_type.get(_text(row["TypeLabel"]))
                if type_id is None:
                    return
                row["TypeId"] = type_id
                self._fill_fixture_params(row, record, type_id, keep_instance=True)
                touched.append(row)
            elif column in self._prefix_param:
                touched.append(row)
                if _text(row[column + "Where"]) == WHERE_TYPE:
                    # Parametro di tipo: cambia per tutte le righe di quel tipo.
                    name = self._prefix_param[column]
                    type_id = _text(row["TypeId"])
                    text = _text(row[column])
                    self.state.pending_type[(type_id, name)] = text
                    for other in self.ftable.Rows:
                        if _text(other["TypeId"]) == type_id and \
                                _text(other[column + "Where"]) == WHERE_TYPE:
                            if _text(other[column]) != text:
                                other[column] = text
                            touched.append(other)
            else:
                return
            for item in touched:
                item["Dirty"] = self._row_dirty(item)
                self._sync_code(item)
        finally:
            self._loading = False
        self._defer(lambda: self._end_edit(row))
        self._defer(self._update_fixture_count)

    def _fixture_edits(self):
        """FixtureEdits con i soli valori diversi da quelli del modello."""
        self._commit_edits()
        edits = FixtureEdits()
        for row in self.ftable.Rows:
            fixture_id = _text(row["FixtureId"])
            record = self._records.get(fixture_id)
            if record is None:
                continue
            type_id = _text(row["TypeId"])
            if type_id != record.type_id:
                edits.types[fixture_id] = type_id
            for name, prefix in FIXTURE_PARAMS:
                if _text(row[prefix + "Where"]) != WHERE_INSTANCE:
                    continue
                original = record.instance_values.get(name)
                text = _text(row[prefix])
                if original is not None and values_differ(original.text, text, original.kind):
                    edits.instance[(fixture_id, name)] = text.strip()
        for (type_id, name), text in self.state.pending_type.items():
            original = self._type_values.get(type_id, {}).get(name)
            if original is not None and values_differ(original.text, text, original.kind):
                edits.type_params[(type_id, name)] = text.strip()
        return edits

    def _update_fixture_count(self):
        dirty = sum(1 for row in self.ftable.Rows if row["Dirty"])
        self.txt_fixture_count.Text = "{0} fixture(s), {1} with changes to write".format(
            self.ftable.Rows.Count, dirty)

    # --------------------------------------------------------------- settings

    def _restore_settings(self):
        try:
            cfg = script.get_config(CONFIG_SECTION)
            self.chk_only_action.IsChecked = bool(cfg.get_option('results_only_action', True))
        except Exception:
            self.chk_only_action.IsChecked = True

    def _save_settings(self):
        try:
            cfg = script.get_config(CONFIG_SECTION)
            cfg.set_option('results_only_action', bool(self.chk_only_action.IsChecked))
            script.save_config()
        except Exception:
            pass

    # ------------------------------------------------------------------ filtro

    def _apply_filter(self):
        parts = []
        if self.chk_only_action.IsChecked:
            # I terminali restano sempre visibili: danno il punto di partenza dei
            # rami; restano visibili anche le righe che cambiano misura.
            parts.append("(Action = true OR Kind = 'fixture' OR Changed = true "
                         "OR Simulated = true)")
        needle = (self.txt_search.Text or "").strip()
        if needle:
            like = escape_like(needle)
            parts.append("(" + " OR ".join(
                "{0} LIKE '%{1}%'".format(col, like)
                for col in ("System", "Code", "PipeType", "Section", "Status", "Notes", "Run",
                            "FedBy")) + ")")
        self.table.DefaultView.RowFilter = " AND ".join(parts)
        self._update_count()

    def _update_count(self):
        shown = self.table.DefaultView.Count
        changed = 0
        pipes = 0
        simulated = 0
        for row in self.table.Rows:
            if row["Changed"] is True:      # DBNull non conta
                changed += 1
                pipes += row["Pipes"]
            if row["Simulated"] is True:
                simulated += 1
        text = "{0} of {1} rows shown; {2} section part(s) change size ({3} pipes)".format(
            shown, self.table.Rows.Count, changed, pipes)
        if simulated:
            text += "; {0} with a simulated slope (not changed)".format(simulated)
        self.txt_count.Text = text

    def OnFilterChanged(self, sender, args):
        if not self._loading:
            self._apply_filter()

    def OnClearSearch(self, sender, args):
        self.txt_search.Text = ""

    # ------------------------------------------------------------------ righe

    def OnColumnChanged(self, sender, args):
        if self._loading:
            return
        if args.Column.ColumnName not in ("NewSize", "SimSlope"):
            return
        row = args.Row
        section = self._by_key[row["Key"]]
        self._loading = True
        try:
            self._fill_new(row, section)
        finally:
            self._loading = False
        self._defer(lambda: self._end_edit(row))
        self._update_count()

    def _end_edit(self, row):
        """Chiude la modifica della riga, cosi' la griglia mostra subito le colonne
        ricalcolate anche se il binding ha aperto una modifica sul DataRowView."""
        try:
            if row.HasVersion(DataRowVersion.Proposed):
                row.EndEdit()
        except Exception:
            pass

    def _defer(self, func):
        try:
            self.Dispatcher.BeginInvoke(DispatcherPriority.Background, Action(func))
        except Exception:
            func()

    def _request_show(self, ids, last_key):
        """Chiude la finestra chiedendo allo script di selezionare gli elementi."""
        self._commit_edits()
        self.kind = KIND_SHOW
        self.show_ids = list(ids)
        self.state.last_key = last_key
        self._save_settings()
        self.Close()

    @staticmethod
    def _button_key(args):
        source = args.OriginalSource
        if not isinstance(source, Button):
            return None
        try:
            return _text(source.DataContext["Key"])
        except Exception:
            return None

    def OnSectionButton(self, sender, args):
        key = self._button_key(args)
        if not key:
            return
        args.Handled = True
        self._request_show(self._by_key[key].pipe_ids, key)

    def OnFixtureButton(self, sender, args):
        key = self._button_key(args)
        if not key:
            return
        args.Handled = True
        fixture_id = self._fixture_ids.get(key[1:])
        if fixture_id is None:
            return
        self._request_show([fixture_id], key)

    # ---------------------------------------------------------------- Export

    def _general_export(self, choice):
        """Scheda General del report: riepilogo, note e tabella dei sistemi."""
        summary = self._summary
        per_m2, per_hour, _varies = intensity_texts(summary.intensities)
        pairs = rp.Pairs("Project", [
            (u"Total covered area [m\u00b2]", summary.area if summary.area is not None else "-"),
            (u"Rainfall intensity [l/(s\u00b7m\u00b2)]", per_m2),
            ("Rainfall intensity [mm/h]", per_hour),
            ("Total collected flow [l/s]", summary.flow if summary.flow is not None else "-"),
        ])
        notes = general_notes(summary)
        pending = self._fixture_edits().count
        if pending:
            notes.append("{0} fixture change(s) not written to the model yet are included "
                         "in the Fixtures tab.".format(pending))
        rows = [rp.Row({"system": item.name, "fixtures": item.fixtures, "area": item.area,
                        "flow": item.flow}) for item in summary.systems]
        table = rp.build_table("Systems", rp.GENERAL_COLUMNS, choice.columns[rexp.GENERAL],
                               rows)
        return rp.Tab(rp.TAB_GENERAL, [pairs, rp.Text(" ".join(notes)), table])

    def _section_export_rows(self):
        """Righe della scheda Sections, nell'ordine della rete (#), tutte (il
        filtro e la ricerca della griglia non contano). Le colonne New sono la
        misura di calcolo (la tendina New size) con la pendenza simulata."""
        result = []
        for row in sorted(self.table.Rows, key=lambda r: r["Order"]):
            entry = self._by_key[row["Key"]]
            values = {
                "order": row["Order"],
                "system": entry.system,
                "code": _text(row["Code"]),
                "fed_by": _text(row["FedBy"]),
                "flow": entry.flow,
                "status": _text(row["Status"]),
                "notes": _text(row["Notes"]),
            }
            if getattr(entry, "is_fixture", False):
                values.update({"section": "-", "run": rc.RUN_FIXTURE, "model_dn": entry.dn_mm,
                               "check_new": "-"})
                result.append(rp.Row(values, rp.ROW_FIXTURE))
                continue
            size = self._size_of(row, entry)
            capacity = _number(row["QNew"])
            check = _text(row["CheckNew"]) or "-"
            use_new = None
            if capacity and entry.flow is not None:
                use_new = entry.flow / capacity * 100.0
            values.update({
                "section": entry.section or "-",
                "pipes": len(entry.pipe_ids),
                "pipe_type": _text(row["PipeType"]),
                "run": entry.run,
                "slope": entry.min_slope,
                "model_dn": number_or_text(row["ModelDn"]),
                "model_de": number_or_text(row["ModelDe"]),
                "capacity": _number(row["Capacity"]),
                "use": _number(row["Use"]),
                "v_model": _number(row["VModel"]),
                "fill_model": _number(row["FillModel"]),
                "suggested": number_or_text(row["Suggested"]),
                "new_dn": size.dn_mm if size is not None else None,
                "new_de": size.od_mm if size is not None else None,
                "sim_slope": self._sim_slope(row),
                "q_new": capacity,
                "use_new": use_new,
                "v_new": _number(row["VNew"]),
                "fill_new": _number(row["FillNew"]),
                "check_new": check,
            })
            kind = rp.ROW_NO_FLOW if _text(row["Severity"]) == "none" else None
            result.append(rp.Row(values, kind, {"status": export_severity(values["status"]),
                                                "check_new": export_severity(check)}))
        return result

    def _fixture_export_rows(self):
        """Righe della scheda Fixtures con i valori della griglia, modifiche non
        ancora scritte comprese (colonna pending); numeri come numeri."""
        result = []
        for row in sorted(self.ftable.Rows, key=lambda r: r["Order"]):
            values = {
                "order": row["Order"],
                "system": _text(row["System"]),
                "level": _text(row["Level"]),
                "id": _text(row["FixtureId"]),
                "type": _text(row["TypeLabel"]),
                "flow": _number(row["Flow"]),
                "pending": "Yes" if row["Dirty"] is True else "",
                "notes": _text(row["Notes"]),
            }
            for _name, prefix in FIXTURE_PARAMS:
                key = prefix.lower()        # Area, C1, C2, Code -> chiavi del catalogo
                text = _text(row[prefix])
                number = parse_decimal(text) if row[prefix + "Num"] is True else None
                values[key] = number if number is not None else text
                values[key + "_from"] = _text(row[prefix + "From"])
            result.append(rp.Row(values))
        return result

    def build_report(self, choice):
        """Le tre schede con le colonne scelte (ExportChoice), piu' le impostazioni."""
        self._commit_edits()
        info = self._export_info
        tabs = [
            self._general_export(choice),
            rp.Tab(rp.TAB_SECTIONS, [rp.build_table(
                "", rp.SECTION_COLUMNS, choice.columns[rexp.SECTIONS],
                self._section_export_rows())]),
            rp.Tab(rp.TAB_FIXTURES, [rp.build_table(
                "", rp.FIXTURE_COLUMNS, choice.columns[rexp.FIXTURES],
                self._fixture_export_rows())]),
        ]
        footer = "RNW Flow Calculation - exported {0}".format(
            datetime.datetime.now().strftime("%d.%m.%Y %H:%M"))
        return rp.Report(info.title, footer, tabs, info.settings_rows)

    def OnExport(self, sender, args):
        if self._export_info is None:
            return
        self._commit_edits()
        choice = rexp.ask_export_options(self)
        if choice is None:
            return
        is_pdf = choice.file_format == rexp.FORMAT_PDF
        dialog = SaveFileDialog()
        if is_pdf:
            dialog.Title = "Export the calculation to PDF"
            dialog.Filter = "PDF document (*.pdf)|*.pdf"
        else:
            dialog.Title = "Export the calculation to Excel"
            dialog.Filter = "Excel workbook (*.xlsx)|*.xlsx"
        dialog.DefaultExt = choice.extension
        dialog.FileName = os.path.splitext(self._export_info.file_name)[0] + choice.extension
        if self._export_info.folder and os.path.isdir(self._export_info.folder):
            dialog.InitialDirectory = self._export_info.folder
        if dialog.ShowDialog(self) != True:     # noqa: E712 (Nullable[bool] di .NET)
            return
        path = dialog.FileName
        try:
            report = self.build_report(choice)
            if is_pdf:
                rpdf.export_report(path, report, choice.paper)
            else:
                rx.export_report(path, report)
        except (rx.FileLockedError, rpdf.FileLockedError):
            MessageBox.Show("The file is open in another program: close it and export "
                            "again.\n" + path, WINDOW_TITLE)
            return
        except Exception as error:
            MessageBox.Show("The export failed:\n{0}".format(error), WINDOW_TITLE)
            return
        answer = MessageBox.Show("Calculation exported to\n{0}\n\nOpen it now?".format(path),
                                 WINDOW_TITLE, MessageBoxButton.YesNo)
        if answer == MessageBoxResult.Yes:
            try:
                open_file(path)
            except Exception:
                pass

    def _commit_edits(self):
        for grid in (self.dg_sections, self.dg_fixtures):
            try:
                grid.CommitEdit(DataGridEditingUnit.Cell, True)
                grid.CommitEdit(DataGridEditingUnit.Row, True)
            except Exception:
                pass

    # -------------------------------------------------------------- validazione

    def _label(self, section, row):
        return "{0} - section {1} {2}".format(section.system, row["Section"], section.run)

    def _part_warnings(self, changes):
        """Section con le due parti (orizzontale e verticale) trattate diversamente:
        la curva fra le due ha una sola misura e potrebbe bloccare la modifica."""
        chosen = {}
        for section, size in changes:
            chosen.setdefault(section.group_key, []).append((section, size))
        parts = {}
        for section in self.sections:
            if section.group_key is not None:
                parts.setdefault(section.group_key, []).append(section)
        warnings = []
        for key, items in chosen.items():
            siblings = parts.get(key, [])
            if len(siblings) < 2:
                continue
            label = "{0} - section {1}".format(key[0], items[0][0].section or "-")
            if len(items) < len(siblings):
                warnings.append("{0}: only the {1} part changes size".format(
                    label, items[0][0].run.lower()))
            elif len(set(round(size.od_mm, 1) for _, size in items)) > 1:
                warnings.append("{0}: the two parts get different sizes".format(label))
        return warnings

    def _collect_changes(self):
        """Ritorna (modifiche, esiti non superati).

        Modifica = (SectionResult, PipeSize) per ogni riga la cui misura nuova e'
        diversa da quella del modello.
        """
        changes = []
        failing = []
        for row in sorted(self.table.Rows, key=lambda r: r["Order"]):
            section = self._by_key[row["Key"]]
            if getattr(section, "is_fixture", False):
                continue
            size = self._size_of(row, section)
            # Una riga con pendenza simulata non cambia il modello
            if row["Simulated"] is True or not is_change(section, size):
                continue
            check = _text(row["CheckNew"])
            if check and check not in NEW_CHECK_PASS:
                failing.append("{0}: {1} with {2}".format(
                    self._label(section, row), check, size.label))
            changes.append((section, size))
        return changes, failing

    @staticmethod
    def _join(lines):
        text = "\n".join(lines[:MAX_LINES_SHOWN])
        if len(lines) > MAX_LINES_SHOWN:
            text += "\n... and {0} more.".format(len(lines) - MAX_LINES_SHOWN)
        return text

    # ---------------------------------------------------------------- buttons

    def OnOK(self, sender, args):
        self._commit_edits()
        changes, failing = self._collect_changes()
        edits = self._fixture_edits()
        simulated = sum(1 for row in self.table.Rows if row["Simulated"] is True)
        simulated_note = ("{0} row(s) with a simulated slope (yellow) are not changed: "
                          "change their slope in the model, then run the tool "
                          "again.".format(simulated)) if simulated else ""
        if not changes and not edits.count:
            MessageBox.Show("Nothing to apply: every new size is the size of the model and "
                            "no fixture value was edited.\n\nChoose a new size in the "
                            "Sections tab or edit the fixtures in the Fixtures tab, or close "
                            "the window." + ("\n\n" + simulated_note if simulated else ""),
                            WINDOW_TITLE)
            return
        lines = ["Apply these changes to the model?", ""]
        if changes:
            pipes = sum(len(section.pipe_ids) for section, _ in changes)
            lines.append("- resize {0} pipe(s) in {1} section part(s) (coloured "
                         "rows)".format(pipes, len(changes)))
        if edits.count:
            lines.append("- write {0} fixture change(s)".format(edits.count))
        question = "\n".join(lines)
        if simulated:
            question += "\n\n" + simulated_note
        if changes and edits.count:
            question += ("\n\nThe sizes were calculated with the current flows. To check "
                         "them with the new fixture values, answer No and use Apply and "
                         "Recalculate first.")
        if failing:
            question += ("\n\nWith the chosen size these sections still do not pass "
                         "the check:\n" + self._join(failing))
        warnings = self._part_warnings(changes)
        if warnings:
            question += ("\n\nA fitting between the horizontal and the vertical part has "
                         "one size, so these sections may be rolled back:\n" +
                         self._join(warnings))
        answer = MessageBox.Show(question, WINDOW_TITLE, MessageBoxButton.YesNo)
        if answer != MessageBoxResult.Yes:
            return
        self.changes = changes
        self.fixture_edits = edits
        self.kind = KIND_APPLY
        self._save_settings()
        self.Close()

    def OnRecalc(self, sender, args):
        edits = self._fixture_edits()
        if not edits.count:
            MessageBox.Show("No fixture value was edited: edit them in the Fixtures tab "
                            "first.", WINDOW_TITLE)
            return
        answer = MessageBox.Show(
            "Write {0} fixture change(s) to the model and recalculate the network?\n\n"
            "Revit updates the flows of the network and the window opens again with the "
            "new results. The sizes you chose by hand are kept; the pipes are not "
            "resized yet.".format(edits.count), WINDOW_TITLE, MessageBoxButton.YesNo)
        if answer != MessageBoxResult.Yes:
            return
        self.fixture_edits = edits
        self.kind = KIND_RECALC
        self._save_settings()
        self.Close()

    def OnCancel(self, sender, args):
        self.changes = None
        self.fixture_edits = None
        self._save_settings()
        self.Close()


def show_results(sections, settings_text, state=None, export_info=None, fixture_data=None):
    """Mostra la finestra (modale) e la ritorna chiusa: kind, changes,
    fixture_edits, show_ids, state."""
    form = ResultsForm(sections, settings_text, state, export_info, fixture_data)
    form.ShowDialog()
    return form
