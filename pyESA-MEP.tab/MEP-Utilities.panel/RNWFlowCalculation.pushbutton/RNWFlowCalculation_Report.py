# -*- coding: utf-8 -*-
"""
RNWFlowCalculation_Report.py - contenuto dell'esportazione del calcolo, uguale per
Excel (RNWFlowCalculation_Xlsx) e PDF (RNWFlowCalculation_Pdf).

Il report ha le tre schede della finestra dei risultati, General, Sections e
Fixtures, piu' le impostazioni del calcolo. Ogni scheda e' una lista di blocchi:
  * Pairs: coppie (voce, valore), per il riepilogo;
  * Text: una nota;
  * Table: colonne (Column) e righe (Row).
I due writer traducono la gravita' delle celle (SEV_*) e il tipo di riga (ROW_*)
nei loro colori, gli stessi della finestra.

Le colonne esportabili di ogni tabella sono nei cataloghi GENERAL_COLUMNS,
SECTION_COLUMNS e FIXTURE_COLUMNS: la finestra di esportazione
(RNWFlowCalculation_Export_ui) fa scegliere quali includere, per chiave. Le
righe arrivano dalla finestra dei risultati come dizionari {chiave: valore} con
tutte le chiavi del catalogo; build_table() tiene le sole colonne scelte.

Python puro: nessun import di Revit ne' di WPF; nessuno stato di modulo
(Rocket mode).
"""

TEXT = "text"
NUMBER = "number"

# Gravita' di una cella (stessi valori di SEVERITY nella finestra dei risultati)
SEV_ERROR = "error"
SEV_WARNING = "warning"
SEV_INFO = "info"
SEV_PASS = "pass"

# Tipo di riga
ROW_FIXTURE = "fixture"
ROW_NO_FLOW = "none"

TAB_GENERAL = "General"
TAB_SECTIONS = "Sections"
TAB_FIXTURES = "Fixtures"


class Column(object):
    """Colonna esportabile.

    key: chiave del valore nella riga; header: intestazione (inglese, come la
    griglia); kind: TEXT o NUMBER; digits: decimali dei numeri; default: inclusa
    se l'utente non ha ancora scelto; width: larghezza in Excel [caratteri];
    severity: True se la cella prende il colore dell'esito (Row.severities).
    """

    def __init__(self, key, header, kind=TEXT, digits=0, default=True, width=10,
                 severity=False):
        self.key = key
        self.header = header
        self.kind = kind
        self.digits = digits
        self.default = default
        self.width = width
        self.severity = severity


class Row(object):
    """Riga di una tabella: valori per chiave, tipo di riga, gravita' per chiave."""

    def __init__(self, values, kind=None, severities=None):
        self.values = values
        self.kind = kind
        self.severities = severities or {}


class Pairs(object):
    """Coppie (voce, valore); un valore numerico si scrive con PAIR_DIGITS
    decimali (e resta un numero in Excel)."""

    def __init__(self, title, rows):
        self.title = title
        self.rows = list(rows)          # [(voce, valore)]


PAIR_DIGITS = 2


def pair_text(value):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return u"{0:.{1}f}".format(float(value), PAIR_DIGITS)
    return u"{0}".format(value if value is not None else u"")


class Text(object):
    def __init__(self, text):
        self.text = text


class Table(object):
    def __init__(self, title, columns, rows):
        self.title = title
        self.columns = list(columns)
        self.rows = list(rows)


class Tab(object):
    def __init__(self, name, blocks):
        self.name = name
        self.blocks = list(blocks)


class Report(object):
    """title: titolo (modello); footer: riga in fondo alle pagine del PDF;
    tabs: [Tab]; settings: [(voce, valore)] (foglio Settings in Excel, in coda
    alla scheda General nel PDF)."""

    def __init__(self, title, footer, tabs, settings):
        self.title = title
        self.footer = footer
        self.tabs = list(tabs)
        self.settings = list(settings)


# =============================================================================
# CATALOGHI DELLE COLONNE
# =============================================================================

GENERAL_COLUMNS = [
    Column("system", "System", width=28),
    Column("fixtures", "Fixtures", NUMBER, 0, width=10),
    Column("area", u"Covered area [m\u00b2]", NUMBER, 2, width=16),
    Column("flow", "Q [l/s]", NUMBER, 2, width=12),
]

# Le colonne della misura nuova (New ...) sono il risultato del calcolo; quelle
# del modello (Model ..., Status) sono lo stato attuale e partono escluse, come
# nell'export Excel delle versioni precedenti.
SECTION_COLUMNS = [
    Column("order", "#", NUMBER, 0, width=5),
    Column("system", "System", width=14),
    Column("code", "Element code", width=16),
    Column("section", "Section", width=9),
    Column("pipes", "Pipes", NUMBER, 0, width=7),
    Column("pipe_type", "Pipe Type", width=14),
    Column("run", "Run", width=11),
    Column("fed_by", "Fed by", width=16),
    Column("slope", "Min slope [cm/m]", NUMBER, 2, width=10),
    Column("flow", "Q [l/s]", NUMBER, 2, width=9),
    Column("model_dn", "Model DN [mm]", NUMBER, 0, default=False, width=9),
    Column("model_de", "Model De [mm]", NUMBER, 0, default=False, width=9),
    Column("capacity", "Qmax model [l/s]", NUMBER, 1, default=False, width=10),
    Column("use", "Use model [%]", NUMBER, 0, default=False, width=8),
    Column("v_model", "v model [m/s]", NUMBER, 2, default=False, width=8),
    Column("fill_model", "Fill model [%]", NUMBER, 0, default=False, width=8),
    Column("suggested", "Suggested De [mm]", NUMBER, 0, default=False, width=10),
    Column("status", "Status", default=False, width=20, severity=True),
    Column("new_dn", "New DN [mm]", NUMBER, 0, width=9),
    Column("new_de", "New De [mm]", NUMBER, 0, width=9),
    Column("sim_slope", "Sim. slope [%]", NUMBER, 2, width=10),
    Column("q_new", "Qmax new [l/s]", NUMBER, 1, width=10),
    Column("use_new", "Use new [%]", NUMBER, 0, width=8),
    Column("v_new", "v new [m/s]", NUMBER, 2, width=8),
    Column("fill_new", "Fill new [%]", NUMBER, 0, width=8),
    Column("check_new", "Check new", width=22, severity=True),
    Column("notes", "Notes", width=42),
]

FIXTURE_COLUMNS = [
    Column("order", "#", NUMBER, 0, width=5),
    Column("system", "System", width=14),
    Column("level", "Level, offset [mm]", width=20),
    Column("id", "Id", width=10),
    Column("type", "Family and type", width=36),
    Column("flow", "Q [l/s]", NUMBER, 2, width=9),
    Column("area", u"Collection area [m\u00b2]", NUMBER, 2, width=12),
    Column("area_from", "Area from", default=False, width=12),
    Column("c1", "Runoff coeff. C1", NUMBER, 2, width=10),
    Column("c1_from", "C1 from", default=False, width=12),
    Column("c2", "Risk coeff. C2", NUMBER, 2, width=10),
    Column("c2_from", "C2 from", default=False, width=12),
    Column("code", "Element code", width=18),
    Column("code_from", "Code from", default=False, width=12),
    Column("pending", "Changes not written", default=False, width=12),
    Column("notes", "Notes", width=36),
]


def default_keys(catalog):
    return [column.key for column in catalog if column.default]


def columns_for(catalog, keys):
    """Colonne del catalogo con chiave in keys, nell'ordine del catalogo."""
    chosen = set(keys or [])
    return [column for column in catalog if column.key in chosen]


def build_table(title, catalog, keys, rows):
    """Table con le sole colonne scelte; le righe restano complete (i writer
    leggono i valori per chiave)."""
    return Table(title, columns_for(catalog, keys), rows)


def cell_text(column, value):
    """Testo di una cella: numeri con i decimali della colonna, vuoto per None."""
    if value is None:
        return u""
    if isinstance(value, bool):
        return u"Yes" if value else u""
    if isinstance(value, (int, float)):
        return u"{0:.{1}f}".format(float(value), column.digits)
    return u"{0}".format(value)
