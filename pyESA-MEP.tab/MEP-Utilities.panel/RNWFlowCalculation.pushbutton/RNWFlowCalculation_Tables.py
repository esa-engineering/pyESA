# -*- coding: utf-8 -*-
"""
RNWFlowCalculation_Tables.py - tabelle di dimensionamento delle reti di acque meteoriche.

Dati trascritti dal manuale tecnico allegato alla richiesta (cap. 6):

* Tabella 6.2 - portata massima delle colonne pluviali con grado di riempimento
  f = 0,33, per diametro esterno De e serie di tubo (PE, PP, Triplus, Silere).
* Tabelle 6.3 / 6.4 / 6.5 - velocita' v [m/s] e portata Q [l/s] dei collettori
  orizzontali in funzione della pendenza i [cm/m], per f = 0,5 / 0,6 / 0,7.

Modulo puro Python, senza import di Revit: si puo' provare con CPython.
Le etichette visibili all'utente sono in inglese (vedi CLAUDE.md, "English out").

Regole di lettura adottate:

* pendenza fra due righe di tabella -> si usa la riga con pendenza inferiore
  (portata ammissibile minore, quindi a favore di sicurezza);
* pendenza oltre 10 cm/m -> si usa la riga 10 cm/m e si segnala;
* pendenza sotto 0,5 cm/m -> fuori tabella, il tratto non e' verificabile;
* diametro suggerito = il minimo De fra quelli candidati (le misure del tipo di
  tubazione, passate dallo script) con Q tabellare >= Q di progetto;
* colonna della tabella 6.2 = la serie del tubo, riconosciuta dallo script dal
  materiale del segmento (series_from_text); se non e' riconosciuta si usa il
  valore piu' cautelativo fra le colonne disponibili per quel De.
"""

import math
import re

# -----------------------------------------------------------------------------
# Serie di tubo (colonne della tabella 6.2)
# -----------------------------------------------------------------------------

SERIES_PE = "PE"
SERIES_PP = "PP"
SERIES_TRIPLUS = "TRIPLUS"
SERIES_SILERE = "SILERE"

SERIES_ORDER = [SERIES_PE, SERIES_PP, SERIES_TRIPLUS, SERIES_SILERE]

SERIES_LABELS = {
    SERIES_PE: "Polyethylene (PE)",
    SERIES_PP: "Polypropylene (PP)",
    SERIES_TRIPLUS: "Triplus",
    SERIES_SILERE: "Silere",
}

STACK_FILL = 0.33

# Tabella 6.2: De -> (PE, PP, Triplus, Silere); None = diametro non disponibile.
_STACK_ROWS = [
    (32, 0.3, 0.4, 0.4, None),
    (40, 0.6, 0.7, 0.7, None),
    (50, 1.2, 1.4, 1.4, None),
    (56, 1.7, None, None, None),
    (58, None, None, None, 1.7),
    (63, 2.4, None, None, None),
    (75, 4.0, 4.3, 4.1, None),
    (78, None, None, None, 3.5),
    (90, 6.5, 7.1, 6.7, 6.1),
    (110, 11.2, 12.1, 11.7, 10.5),
    (125, 15.7, 17.0, 16.4, None),
    (135, None, None, None, 18.9),
    (160, 30.3, 32.8, 31.7, 30.9),
    (200, 57.4, None, 57.4, None),
    (250, 104.1, None, 104.1, None),
    (315, 192.8, None, None, None),
]

# {serie: {De: Qmax}}
STACK_TABLE = {}
for _row in _STACK_ROWS:
    for _index, _series in enumerate(SERIES_ORDER):
        _value = _row[_index + 1]
        if _value is not None:
            STACK_TABLE.setdefault(_series, {})[_row[0]] = _value

# Diametri disponibili per serie (dalla tabella 6.2): limitano anche i
# diametri suggeriti per i collettori orizzontali.
SERIES_DIAMETERS = dict((s, sorted(STACK_TABLE[s].keys())) for s in SERIES_ORDER)

# -----------------------------------------------------------------------------
# Collettori orizzontali (tabelle 6.3, 6.4, 6.5)
# -----------------------------------------------------------------------------

HORIZONTAL_DE = [110, 125, 135, 160, 200, 250, 315]

FILL_OPTIONS = [0.5, 0.6, 0.7]

# Ogni riga: (pendenza cm/m, [v, Q] per ciascun De di HORIZONTAL_DE)
_HORIZONTAL_ROWS = {
    0.5: [
        (0.5, [0.5, 1.7, 0.5, 2.6, 0.5, 2.6, 0.6, 5.3, 0.7, 9.9, 0.9, 18.2, 1.0, 33.9]),
        (1.0, [0.7, 2.4, 0.7, 3.7, 0.7, 3.7, 0.9, 7.5, 1.1, 14.0, 1.2, 25.7, 1.5, 48.0]),
        (1.5, [0.8, 2.9, 0.9, 4.5, 0.9, 4.5, 1.1, 9.1, 1.3, 17.2, 1.5, 31.4, 1.8, 58.8]),
        (2.0, [0.9, 3.3, 1.0, 5.2, 1.0, 5.2, 1.3, 10.6, 1.5, 19.8, 1.7, 36.3, 2.1, 67.8]),
        (2.5, [1.0, 3.7, 1.2, 5.9, 1.2, 5.9, 1.4, 11.8, 1.7, 22.2, 2.0, 40.6, 2.3, 75.8]),
        (3.0, [1.1, 4.1, 1.3, 6.4, 1.3, 6.4, 1.5, 12.9, 1.8, 24.3, 2.1, 44.5, 2.5, 83.1]),
        (3.5, [1.2, 4.4, 1.4, 6.9, 1.4, 6.9, 1.7, 14.0, 2.0, 26.2, 2.3, 48.0, 2.7, 89.7]),
        (4.0, [1.3, 4.7, 1.5, 7.4, 1.5, 7.4, 1.8, 14.9, 2.1, 28.0, 2.5, 51.3, 2.9, 95.9]),
        (4.5, [1.4, 5.0, 1.6, 7.8, 1.6, 7.8, 1.9, 15.8, 2.2, 29.7, 2.6, 54.5, 3.1, 101.8]),
        (5.0, [1.5, 5.3, 1.6, 8.3, 1.6, 8.3, 2.0, 16.7, 2.4, 31.4, 2.8, 57.4, 3.2, 107.3]),
        (6.0, [1.6, 5.8, 1.8, 9.1, 1.8, 9.1, 2.2, 18.3, 2.6, 34.3, 3.0, 62.9, 3.6, 117.5]),
        (7.0, [1.7, 6.2, 2.0, 9.8, 2.0, 9.8, 2.4, 19.8, 2.8, 37.1, 3.3, 67.9, 3.8, 126.9]),
        (8.0, [1.8, 6.7, 2.1, 10.5, 2.1, 10.5, 2.5, 21.1, 3.0, 39.7, 3.5, 72.6, 4.1, 135.7]),
        (9.0, [2.0, 7.1, 2.2, 11.1, 2.2, 11.1, 2.7, 22.4, 3.2, 42.1, 3.7, 77.0, 4.4, 143.9]),
        (10.0, [2.1, 7.5, 2.3, 11.7, 2.3, 11.7, 2.8, 23.6, 3.3, 44.3, 3.9, 81.2, 4.6, 151.7]),
    ],
    0.6: [
        (0.5, [0.5, 2.3, 0.6, 3.5, 0.6, 3.5, 0.7, 7.1, 0.8, 13.4, 0.9, 24.5, 1.1, 45.7]),
        (1.0, [0.7, 3.2, 0.8, 5.0, 0.8, 5.0, 1.0, 10.1, 1.1, 18.9, 1.3, 34.6, 1.6, 64.6]),
        (1.5, [0.9, 3.9, 1.0, 6.1, 1.0, 6.1, 1.2, 12.4, 1.4, 23.2, 1.6, 42.4, 1.9, 79.1]),
        (2.0, [1.0, 4.5, 1.1, 7.1, 1.1, 7.1, 1.4, 14.3, 1.6, 26.8, 1.9, 49.0, 2.2, 91.3]),
        (2.5, [1.1, 5.1, 1.3, 7.9, 1.3, 7.9, 1.5, 16.0, 1.8, 29.9, 2.1, 54.7, 2.5, 102.1]),
        (3.0, [1.2, 5.5, 1.4, 8.7, 1.4, 8.7, 1.7, 17.5, 2.0, 32.8, 2.3, 60.0, 2.7, 111.9]),
        (3.5, [1.3, 6.0, 1.5, 9.4, 1.5, 9.4, 1.8, 18.9, 2.1, 35.4, 2.5, 64.8, 2.9, 120.8]),
        (4.0, [1.4, 6.4, 1.6, 10.0, 1.6, 10.0, 1.9, 20.2, 2.3, 37.9, 2.7, 69.2, 3.1, 129.2]),
        (4.5, [1.5, 6.8, 1.7, 10.6, 1.7, 10.6, 2.0, 21.4, 2.4, 40.2, 2.8, 73.4, 3.3, 137.0]),
        (5.0, [1.6, 7.2, 1.8, 11.2, 1.8, 11.2, 2.2, 22.6, 2.5, 42.3, 3.0, 77.4, 3.5, 144.4]),
        (6.0, [1.7, 7.8, 2.0, 12.3, 2.0, 12.3, 2.4, 24.7, 2.8, 46.4, 3.3, 84.8, 3.8, 158.2]),
        (7.0, [1.9, 8.5, 2.1, 13.3, 2.1, 13.3, 2.5, 26.7, 3.0, 50.1, 3.5, 91.6, 4.1, 170.9]),
        (8.0, [2.0, 9.1, 2.3, 14.2, 2.3, 14.2, 2.7, 28.6, 3.2, 53.6, 3.8, 97.9, 4.4, 182.7]),
        (9.0, [2.1, 9.6, 2.4, 15.0, 2.4, 15.0, 2.9, 30.3, 3.4, 56.8, 4.0, 103.9, 4.7, 193.8]),
        (10.0, [2.2, 10.1, 2.5, 15.9, 2.5, 15.9, 3.0, 31.9, 3.6, 59.9, 4.2, 109.5, 4.9, 204.3]),
    ],
    0.7: [
        (0.5, [0.5, 2.8, 0.6, 4.4, 0.6, 4.4, 0.7, 8.9, 0.8, 16.7, 1.0, 30.6, 1.2, 57.0]),
        (1.0, [0.7, 4.0, 0.8, 6.3, 0.8, 6.3, 1.0, 12.6, 1.2, 23.7, 1.4, 43.2, 1.6, 80.6]),
        (1.5, [0.9, 4.9, 1.0, 7.7, 1.0, 7.7, 1.2, 15.5, 1.5, 29.0, 1.7, 52.9, 2.0, 98.7]),
        (2.0, [1.0, 5.7, 1.2, 8.9, 1.2, 8.9, 1.4, 17.9, 1.7, 33.5, 2.0, 61.1, 2.3, 113.9]),
        (2.5, [1.2, 6.3, 1.3, 9.9, 1.3, 9.9, 1.6, 20.0, 1.9, 37.4, 2.2, 68.3, 2.6, 127.4]),
        (3.0, [1.3, 7.0, 1.4, 10.9, 1.4, 10.9, 1.7, 21.9, 2.1, 41.0, 2.4, 74.9, 2.8, 139.6]),
        (3.5, [1.4, 7.5, 1.6, 11.7, 1.6, 11.7, 1.9, 23.6, 2.2, 44.3, 2.6, 80.9, 3.1, 150.7]),
        (4.0, [1.5, 8.0, 1.7, 12.6, 1.7, 12.6, 2.0, 25.3, 2.4, 47.3, 2.8, 86.4, 3.3, 161.1]),
        (4.5, [1.6, 8.5, 1.8, 13.3, 1.8, 13.3, 2.1, 26.8, 2.5, 50.2, 3.0, 91.7, 3.5, 170.9]),
        (5.0, [1.7, 9.0, 1.9, 14.0, 1.9, 14.0, 2.3, 28.2, 2.7, 52.9, 3.1, 96.6, 3.6, 180.2]),
        (6.0, [1.8, 9.8, 2.1, 15.4, 2.1, 15.4, 2.5, 30.9, 2.9, 58.0, 3.4, 105.9, 4.0, 197.4]),
        (7.0, [2.0, 10.6, 2.2, 16.6, 2.2, 16.6, 2.7, 33.4, 3.1, 62.6, 3.7, 114.4, 4.3, 213.2]),
        (8.0, [2.1, 11.4, 2.4, 17.8, 2.4, 17.8, 2.9, 35.7, 3.4, 66.9, 3.9, 122.3, 4.6, 227.9]),
        (9.0, [2.2, 12.0, 2.5, 18.8, 2.5, 18.8, 3.0, 37.9, 3.6, 71.0, 4.2, 129.7, 4.9, 241.7]),
        (10.0, [2.3, 12.7, 2.6, 19.8, 2.6, 19.8, 3.2, 39.9, 3.8, 74.8, 4.4, 136.7, 5.2, 254.8]),
    ],
}

HORIZONTAL_SLOPES = [row[0] for row in _HORIZONTAL_ROWS[0.5]]
MIN_SLOPE = HORIZONTAL_SLOPES[0]
MAX_SLOPE = HORIZONTAL_SLOPES[-1]

# {fill: {slope: {De: (v, Q)}}}
HORIZONTAL_TABLE = {}
for _fill, _rows in _HORIZONTAL_ROWS.items():
    _by_slope = {}
    for _slope, _values in _rows:
        _cells = {}
        for _index, _de in enumerate(HORIZONTAL_DE):
            _cells[_de] = (_values[2 * _index], _values[2 * _index + 1])
        _by_slope[_slope] = _cells
    HORIZONTAL_TABLE[_fill] = _by_slope

TABLE_NAMES = {
    0.5: "Table 6.3",
    0.6: "Table 6.4",
    0.7: "Table 6.5",
}
STACK_TABLE_NAME = "Table 6.2"

# Tolleranza per riconoscere un De di tabella dal diametro esterno del modello.
DE_MATCH_TOLERANCE_MM = 1.0


# -----------------------------------------------------------------------------
# Funzioni di lettura
# -----------------------------------------------------------------------------

def match_de(outer_diameter_mm, candidates):
    """De di tabella piu' vicino al diametro esterno, entro la tolleranza; None se nessuno."""
    if outer_diameter_mm is None:
        return None
    best = None
    best_delta = None
    for de in candidates:
        delta = abs(de - outer_diameter_mm)
        if delta <= DE_MATCH_TOLERANCE_MM and (best_delta is None or delta < best_delta):
            best = de
            best_delta = delta
    return best


def slope_row(slope_cm_m):
    """Riga di tabella da usare per una pendenza in cm/m.

    Ritorna (pendenza di tabella, motivo) dove motivo e' None, "BELOW" (pendenza
    sotto il minimo: nessuna riga) o "ABOVE" (oltre il massimo: si usa l'ultima).
    """
    if slope_cm_m is None:
        return None, "BELOW"
    # Arrotondamento: 1 % calcolato dalla geometria puo' valere 0.9999999.
    value = round(slope_cm_m, 3)
    if value < MIN_SLOPE:
        return None, "BELOW"
    if value > MAX_SLOPE:
        return MAX_SLOPE, "ABOVE"
    chosen = MIN_SLOPE
    for slope in HORIZONTAL_SLOPES:
        if slope <= value:
            chosen = slope
    return chosen, None


def horizontal_cell(fill, table_slope, de):
    """(v, Q) della tabella orizzontale, None se il De non e' in tabella."""
    if table_slope is None or de is None:
        return None
    return HORIZONTAL_TABLE[fill][table_slope].get(de)


def horizontal_diameters(series):
    """De dei collettori orizzontali disponibili nella serie scelta."""
    available = set(SERIES_DIAMETERS.get(series, []))
    return [de for de in HORIZONTAL_DE if de in available]


def series_from_text(text):
    """Serie (colonna della tabella 6.2) riconosciuta da un testo (nome del
    materiale, del segmento o del tipo di tubazione); None se non riconosciuta."""
    lowered = " {0} ".format((text or "").lower())
    if "silere" in lowered:
        return SERIES_SILERE
    if "triplus" in lowered:
        return SERIES_TRIPLUS

    def token(word):
        return re.search(r"(^|[^a-z])" + word + r"([^a-z]|$)", lowered) is not None

    if "polipropilene" in lowered or "polypropylene" in lowered or token("pp"):
        return SERIES_PP
    if "polietilene" in lowered or "polyethylene" in lowered or token("pe") \
            or token("hdpe") or token("pead") or token("pe-hd"):
        return SERIES_PE
    return None


def suggest_horizontal(fill, table_slope, flow, candidates=None, vmax=None):
    """Minimo De fra i candidati (default: tutti quelli della tabella) con
    Q >= flow. Ritorna (De, v, Q) o None.

    Con vmax, il minimo De che rispetta anche la velocita' massima alla portata
    di progetto; se nessun De la rispetta si torna al solo criterio di portata.
    """
    if table_slope is None:
        return None
    if candidates is None:
        candidates = HORIZONTAL_DE
    adequate = []
    for de in sorted(de for de in candidates if de in HORIZONTAL_DE):
        v, q = HORIZONTAL_TABLE[fill][table_slope][de]
        if q >= flow:
            adequate.append((de, v, q))
    if not adequate:
        return None
    if vmax is not None:
        for candidate in adequate:
            state = partial_state(flow, candidate[0], table_slope)
            if state is not None and state[1] <= vmax:
                return candidate
    return adequate[0]


def stack_capacity(series, de):
    """Qmax della colonna (tabella 6.2), None se il De non esiste nella serie.

    Serie None (non riconosciuta): il valore minimo fra le colonne che hanno quel
    De, a favore di sicurezza.
    """
    if de is None:
        return None
    if series is None:
        values = [STACK_TABLE[s][de] for s in SERIES_ORDER if de in STACK_TABLE.get(s, {})]
        return min(values) if values else None
    return STACK_TABLE.get(series, {}).get(de)


def suggest_stack(flow, series, candidates=None):
    """Minimo De fra i candidati (default: quelli della serie) con Qmax >= flow.
    Ritorna (De, Qmax) o None."""
    if candidates is None:
        candidates = SERIES_DIAMETERS.get(series, []) if series else all_stack_diameters()
    for de in sorted(candidates):
        q = stack_capacity(series, de)
        if q is not None and q >= flow:
            return de, q
    return None


def all_stack_diameters():
    """Tutti i De della tabella 6.2, qualunque serie."""
    return [row[0] for row in _STACK_ROWS]


# -----------------------------------------------------------------------------
# Riempimento parziale e velocita' alla portata di progetto
# -----------------------------------------------------------------------------
#
# La v delle tabelle 6.3 - 6.5 vale al grado di riempimento della tabella, cioe'
# alla portata Q tabellare, non alla portata reale del tubo. Le tabelle seguono
# la relazione di moto uniforme a riempimento parziale di un tubo circolare
# (Gauckler-Strickler / Manning, coefficiente costante): per esempio
# Q(f=0,7) / Q(f=0,5) = 1,674 in teoria, 1,67 in tabella per De 110 e 160.
# Quindi per ogni De e pendenza si ricavano dalle tre tabelle portata e velocita'
# a tubo pieno (media delle tre stime, per smorzare l'arrotondamento a un
# decimale delle v), e alla portata reale si risolve il grado di riempimento
# y = h/D e la velocita' v = v_pieno x (R(y) / R_pieno)^(2/3).

# Grado di riempimento con la portata massima (Q/Q_pieno = 1,076 a y = 0,938).
_Y_MAX_FLOW = 0.938


def fill_ratios(y):
    """(A/A_pieno, v/v_pieno, Q/Q_pieno) per grado di riempimento y = h/D."""
    if y <= 0:
        return 0.0, 0.0, 0.0
    if y >= 1:
        return 1.0, 1.0, 1.0
    theta = 2.0 * math.acos(1.0 - 2.0 * y)
    area = (theta - math.sin(theta)) / (2.0 * math.pi)
    radius = 1.0 - math.sin(theta) / theta
    velocity = radius ** (2.0 / 3.0)
    return area, velocity, area * velocity


_MAX_FLOW_RATIO = fill_ratios(_Y_MAX_FLOW)[2]


def full_pipe(de, table_slope):
    """(Q_pieno [l/s], v_pieno [m/s]) dalle tre tabelle; None se De fuori tabella."""
    if table_slope is None:
        return None
    flows = []
    velocities = []
    for fill in FILL_OPTIONS:
        cell = HORIZONTAL_TABLE[fill][table_slope].get(de)
        if cell is None:
            return None
        _area, v_ratio, q_ratio = fill_ratios(fill)
        velocities.append(cell[0] / v_ratio)
        flows.append(cell[1] / q_ratio)
    return sum(flows) / len(flows), sum(velocities) / len(velocities)


def partial_state(flow, de, table_slope):
    """(grado di riempimento y, velocita' m/s) alla portata flow.

    None se il De o la pendenza sono fuori tabella. Se la portata supera quella
    massima a pelo libero il riempimento vale 1.0 e la velocita' None (tubo in
    pressione: la tabella non vale piu').
    """
    full = full_pipe(de, table_slope)
    if full is None or flow is None or flow <= 0:
        return None
    q_full, v_full = full
    ratio = flow / q_full
    if ratio >= _MAX_FLOW_RATIO:
        return 1.0, None
    low, high = 0.0, _Y_MAX_FLOW
    for _ in range(60):
        mid = (low + high) / 2.0
        if fill_ratios(mid)[2] < ratio:
            low = mid
        else:
            high = mid
    y = (low + high) / 2.0
    return y, v_full * fill_ratios(y)[1]
