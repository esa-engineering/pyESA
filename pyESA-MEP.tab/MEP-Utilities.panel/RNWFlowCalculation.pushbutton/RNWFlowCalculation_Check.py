# -*- coding: utf-8 -*-
"""
RNWFlowCalculation_Check.py - verifica dei tubi e aggregazione per tratto (Section).

Modulo puro Python (nessun import di Revit): lo script legge il modello e
riduce ogni tubo a un PipeCheck con dati semplici; qui si fanno la verifica
contro le tabelle (RNWFlowCalculation_Tables) e il raggruppamento per tratto.

Tratto = (nome del sistema, parametro Section del tubo, orientamento). Revit
numera i tratti per sistema: un tratto e' una sequenza di tubi a portata e
dimensione costanti. Ogni Section si divide in una parte orizzontale (tubi
orizzontali e inclinati fino a STACK_MIN_ANGLE_DEG, quindi anche i 45 gradi) e
una verticale, verificate e modificate come righe separate. group_key riunisce
le righe della stessa Section: lo script le modifica nella stessa transazione.
I tubi senza Section formano un tratto ciascuno.

Verifica di un tubo a un De (evaluate): portata ammissibile dalla tabella,
grado di riempimento e velocita' alla portata di progetto (solo orizzontali,
RNWFlowCalculation_Tables.partial_state), confronto con le velocita' minima e massima
scelte dall'utente. La stessa funzione serve per il De del modello e per il De
scelto nella tendina della finestra dei risultati.

Regole di aggregazione del tratto:

* portata = massima portata dei suoi tubi;
* stato = il piu' grave fra i tubi; OK / OVERSIZED si decidono sul tratto;
* De suggerito = il massimo De suggerito fra i tubi con portata; nessuno se un
  tubo non e' verificabile (pendenza sotto tabella, portata oltre tabella);
* portata ammissibile = la minima fra i tubi; velocita' minima e massima e
  riempimento massimo fra i tubi orizzontali;
* misure disponibili = quelle comuni ai tubi del tratto (dai segmenti del tipo
  di tubazione, vedi lo script).
"""

import RNWFlowCalculation_Tables as rt

# Inclinazione sull'orizzontale oltre la quale un tubo e' una colonna: i tratti a
# 45 gradi restano orizzontali (verificati con le tabelle 6.3 - 6.5, riga 10 cm/m).
STACK_MIN_ANGLE_DEG = 60.0

# Inclinazione da cui un tubo orizzontale e' un tratto a 45 gradi: si verifica come
# gli altri (riga 10 cm/m), ma senza la nota "Slope over 10 cm/m", che su questi
# tubi comparirebbe sempre. Fra 10 cm/m e questo angolo la nota resta.
RUN_45_MIN_ANGLE_DEG = 30.0

# Pendenze [cm/m] della simulazione nella finestra dei risultati: tutte le righe
# delle tabelle 6.3 - 6.5 (0,5 - 5 ogni 0,5, poi 6 - 10 ogni 1)
SIMULATED_SLOPES = list(rt.HORIZONTAL_SLOPES)

# Testo della colonna Pipe Type per un tipo di tubazione senza Type Mark
NO_TYPE_MARK = "(no type mark)"

ORIENT_HORIZONTAL = "Horizontal"
ORIENT_STACK = "Vertical"
ORIENT_MIXED = "Mixed"

STATUS_OK = "OK"
STATUS_OVERSIZED = "OVERSIZED"
STATUS_UNDERSIZED = "UNDERSIZED"
STATUS_NOT_IN_TABLE = "DE NOT IN TABLE"
STATUS_LOW_SLOPE = "SLOPE UNDER 0.5 cm/m"
STATUS_BEYOND = "FLOW BEYOND TABLE"
STATUS_HIGH_VELOCITY = "HIGH VELOCITY"
STATUS_LOW_VELOCITY = "LOW VELOCITY"
STATUS_NO_FLOW = "NO FLOW"

# Dal piu' grave al meno grave
STATUS_ORDER = [STATUS_UNDERSIZED, STATUS_BEYOND, STATUS_NOT_IN_TABLE,
                STATUS_LOW_SLOPE, STATUS_HIGH_VELOCITY, STATUS_LOW_VELOCITY,
                STATUS_OVERSIZED, STATUS_OK, STATUS_NO_FLOW]

# Stati per cui si propone di cambiare diametro (tratti spuntati in apertura)
STATUS_TO_RESIZE = (STATUS_UNDERSIZED, STATUS_NOT_IN_TABLE, STATUS_HIGH_VELOCITY)

# Stati che non richiedono nessuna azione
STATUS_PASS = (STATUS_OK, STATUS_NO_FLOW)

# Tolleranza [mm] per riconoscere un diametro esterno fra le misure disponibili
OD_TOLERANCE_MM = 1.0

# Differenza di portata [l/s] oltre la quale i tubi di un tratto sono segnalati
FLOW_TOLERANCE = 0.01


class Limits(object):
    """Parametri della verifica scelti dall'utente."""

    def __init__(self, fill, series, vmin=None, vmax=None):
        self.fill = fill
        self.series = series
        self.vmin = vmin
        self.vmax = vmax


class PipeSize(object):
    """Una misura disponibile per il tubo: diametro esterno e nominale [mm].

    nominal e' il valore da scrivere nel parametro Diameter (unita' interne);
    resta opaco per questo modulo.
    """

    __slots__ = ("od_mm", "dn_mm", "nominal")

    def __init__(self, od_mm, dn_mm, nominal=None):
        self.od_mm = od_mm
        self.dn_mm = dn_mm
        self.nominal = nominal

    @property
    def label(self):
        return "De {0}  /  DN {1}".format(fmt_mm(self.od_mm), fmt_mm(self.dn_mm))


def fmt_mm(value):
    if value is None:
        return "-"
    value = float(value)
    if abs(value - round(value)) < 0.05:
        return "{0:.0f}".format(value)
    return "{0:.1f}".format(value)


def od_in(value_mm, ods):
    """True se value_mm e' fra i diametri ods, entro la tolleranza."""
    if value_mm is None or ods is None:
        return False
    return any(abs(od - value_mm) <= OD_TOLERANCE_MM for od in ods)


def find_size(sizes, od_mm):
    """Misura di sizes con diametro esterno od_mm (entro la tolleranza), o None."""
    if sizes is None or od_mm is None:
        return None
    best = None
    for size in sizes:
        delta = abs(size.od_mm - od_mm)
        if delta <= OD_TOLERANCE_MM and (best is None or delta < abs(best.od_mm - od_mm)):
            best = size
    return best


class PipeCheck(object):
    """Un tubo ridotto a dati semplici, piu' il risultato della verifica."""

    def __init__(self, pipe_id, system, system_type, level, section, flow,
                 angle, slope, model_de_mm, model_dn_mm, sizes, z_mid=None, series=None,
                 code=None, type_mark=None):
        self.id = pipe_id
        self.system = system
        self.system_type = system_type
        self.level = level
        self.section = section            # stringa, None se il tubo non ha Section
        self.flow = flow if flow is not None else 0.0
        self.angle = angle
        self.model_de_mm = model_de_mm
        self.model_dn_mm = model_dn_mm
        self.sizes = sizes                # [PipeSize] disponibili, None se ignote
        self.z_mid = z_mid                # quota del punto medio [mm], per il verso del flusso
        self.series = series              # colonna della tabella 6.2, None se non riconosciuta
        self.code = code                  # e_MEC_ElementCode_1, None se assente o vuoto
        self.type_mark = type_mark        # Type Mark del tipo di tubazione, None se vuoto
        if angle is not None and angle > STACK_MIN_ANGLE_DEG:
            self.orientation = ORIENT_STACK
            self.slope = None
            self.table_slope = None
            self.slope_reason = None
        else:
            self.orientation = ORIENT_HORIZONTAL
            self.slope = slope if slope is not None else 0.0
            self.table_slope, self.slope_reason = rt.slope_row(self.slope)
        self.model_de = None
        self.capacity = None
        self.velocity = None
        self.fill_actual = None
        self.suggested_de = None
        self.suggested_capacity = None
        self.status = STATUS_OK
        self.notes = []


class Evaluation(object):
    """Esito di un tubo (o di un tratto) a un De dato."""

    __slots__ = ("capacity", "velocity", "fill_actual", "status")

    def __init__(self, capacity=None, velocity=None, fill_actual=None, status=STATUS_OK):
        self.capacity = capacity
        self.velocity = velocity
        self.fill_actual = fill_actual
        self.status = status


def table_de(row, de_mm):
    """De di tabella corrispondente a de_mm per l'orientamento del tubo."""
    if row.orientation == ORIENT_STACK:
        return rt.match_de(de_mm, rt.all_stack_diameters())
    return rt.match_de(de_mm, rt.HORIZONTAL_DE)


def evaluate(row, de_mm, limits, table_slope=None):
    """Verifica il tubo come se avesse diametro esterno de_mm.

    Stati possibili: NO FLOW, SLOPE UNDER 0.5, DE NOT IN TABLE (De non in
    tabella o non nella serie), UNDERSIZED, HIGH / LOW VELOCITY, OK.
    table_slope: riga di tabella simulata per un tubo orizzontale, al posto di
    quella della pendenza del modello.
    """
    if row.flow <= 0:
        return Evaluation(status=STATUS_NO_FLOW)
    de = table_de(row, de_mm)
    if row.orientation == ORIENT_STACK:
        capacity = rt.stack_capacity(row_series(row, limits), de)
        if capacity is None:
            return Evaluation(status=STATUS_NOT_IN_TABLE)
        if capacity < row.flow:
            return Evaluation(capacity, status=STATUS_UNDERSIZED)
        return Evaluation(capacity, status=STATUS_OK)

    slope = table_slope if table_slope is not None else row.table_slope
    if slope is None:
        return Evaluation(status=STATUS_LOW_SLOPE)
    cell = rt.horizontal_cell(limits.fill, slope, de)
    if cell is None:
        return Evaluation(status=STATUS_NOT_IN_TABLE)
    capacity = cell[1]
    state = rt.partial_state(row.flow, de, slope)
    fill_actual, velocity = state if state is not None else (None, None)
    if capacity < row.flow:
        return Evaluation(capacity, velocity, fill_actual, STATUS_UNDERSIZED)
    if velocity is not None and limits.vmax is not None and velocity > limits.vmax:
        return Evaluation(capacity, velocity, fill_actual, STATUS_HIGH_VELOCITY)
    if velocity is not None and limits.vmin is not None and velocity < limits.vmin:
        return Evaluation(capacity, velocity, fill_actual, STATUS_LOW_VELOCITY)
    return Evaluation(capacity, velocity, fill_actual, STATUS_OK)


def row_series(row, limits):
    """Serie del tubo (dal segmento); in mancanza quella dei limiti (di solito None)."""
    return row.series if row.series is not None else limits.series


def _candidates(row, table):
    """De di tabella fra le misure del tipo di tubazione del tubo; None se ignote."""
    if not row.sizes:
        return None
    found = set()
    for size in row.sizes:
        de = rt.match_de(size.od_mm, table)
        if de is not None:
            found.add(de)
    return sorted(found) or None


def check_pipe(row, limits):
    """Verifica il tubo al De del modello e ne imposta stato e De suggerito.

    I De suggeriti si scelgono fra le misure del tipo di tubazione del tubo; per
    le colonne la portata ammissibile e' quella della serie del tubo (colonna
    della tabella 6.2), o la piu' cautelativa se la serie non e' riconosciuta.
    """
    series = row_series(row, limits)
    row.model_de = table_de(row, row.model_de_mm)
    result = evaluate(row, row.model_de_mm, limits)
    row.capacity = result.capacity
    row.velocity = result.velocity
    row.fill_actual = result.fill_actual

    if result.status == STATUS_NO_FLOW:
        row.status = STATUS_NO_FLOW
        return row
    if result.status == STATUS_LOW_SLOPE:
        row.status = STATUS_LOW_SLOPE
        return row
    run_45 = row.angle is not None and row.angle >= RUN_45_MIN_ANGLE_DEG
    if row.slope_reason == "ABOVE" and not run_45:
        row.notes.append("Slope over {0:.0f} cm/m: {0:.0f} cm/m row used".format(rt.MAX_SLOPE))

    if row.orientation == ORIENT_STACK:
        if series is None:
            row.notes.append("Pipe series not recognized from the pipe segment: the most "
                             "conservative column of Table 6.2 is used")
        elif row.model_de is not None and rt.stack_capacity(series, row.model_de) is None:
            row.notes.append("De {0} not in Table 6.2 for {1}".format(
                row.model_de, rt.SERIES_LABELS[series]))
        suggestion = rt.suggest_stack(row.flow, series,
                                      _candidates(row, rt.all_stack_diameters()))
        if suggestion is not None:
            row.suggested_de, row.suggested_capacity = suggestion
    else:
        suggestion = rt.suggest_horizontal(limits.fill, row.table_slope, row.flow,
                                           _candidates(row, rt.HORIZONTAL_DE), limits.vmax)
        if suggestion is not None:
            row.suggested_de = suggestion[0]
            row.suggested_capacity = suggestion[2]

    status = result.status
    if status in (STATUS_UNDERSIZED, STATUS_NOT_IN_TABLE) and row.suggested_de is None:
        status = STATUS_BEYOND
    if status == STATUS_OK and row.suggested_de is not None and row.model_de is not None \
            and row.model_de > row.suggested_de:
        status = STATUS_OVERSIZED
    row.status = status

    if status == STATUS_LOW_VELOCITY and row.suggested_de is not None:
        state = rt.partial_state(row.flow, row.suggested_de, row.table_slope)
        if state is not None and state[1] is not None and \
                (limits.vmin is None or state[1] < limits.vmin):
            row.notes.append("Velocity stays low at every De: increase the slope")
    if status == STATUS_HIGH_VELOCITY:
        state = rt.partial_state(row.flow, row.suggested_de, row.table_slope) \
            if row.suggested_de is not None else None
        if state is None or state[1] is None or state[1] > limits.vmax:
            row.notes.append("Velocity stays high at every De: reduce the slope")
    return row


def _common_sizes(rows):
    """Misure presenti per tutti i tubi (per diametro esterno); None se una e' ignota."""
    common = None
    for row in rows:
        if row.sizes is None:
            return None
        if common is None:
            common = sorted(row.sizes, key=lambda s: s.od_mm)
        else:
            common = [s for s in common if find_size(row.sizes, s.od_mm) is not None]
    return common or []


def _worst(statuses):
    return min(statuses, key=STATUS_ORDER.index)


class SectionResult(object):
    """Un tratto: i suoi tubi e la proposta di diametro."""

    def __init__(self, key, system, section, rows, limits):
        self.key = key
        self.group_key = key[:2]
        self.system = system
        self.section = section
        self.rows = rows
        self.limits = limits
        self.pipe_ids = [r.id for r in rows]
        quotes = [r.z_mid for r in rows if r.z_mid is not None]
        self.z = sum(quotes) / len(quotes) if quotes else None
        self.system_type = rows[0].system_type
        self.level = rows[0].level
        self.flow = max(r.flow for r in rows)

        orientations = set(r.orientation for r in rows)
        self.run = orientations.pop() if len(orientations) == 1 else ORIENT_MIXED
        slopes = [r.slope for r in rows if r.orientation == ORIENT_HORIZONTAL
                  and r.slope is not None]
        self.min_slope = min(slopes) if slopes else None

        self.model_des = sorted(set(round(r.model_de_mm, 1) for r in rows
                                    if r.model_de_mm is not None))
        self.model_dns = sorted(set(round(r.model_dn_mm, 1) for r in rows
                                    if r.model_dn_mm is not None))
        # Codice elemento dei tubi del tratto (valori diversi uniti da "/")
        self.element_code = "/".join(sorted(set(r.code for r in rows if r.code)))
        # Type Mark dei tipi di tubazione del tratto, piu' tipi uniti da " - "; un
        # tipo senza Type Mark compare come "(no type mark)"
        self.pipe_type = " - ".join(sorted(set(r.type_mark or NO_TYPE_MARK for r in rows)))

        active = [r for r in rows if r.status != STATUS_NO_FLOW]
        self.status = _worst([r.status for r in rows])

        capacities = [r.capacity for r in active]
        self.capacity = min(capacities) if active and all(c is not None for c in capacities) \
            else None
        velocities = [r.velocity for r in active if r.velocity is not None]
        self.v_min = min(velocities) if velocities else None
        self.v_max = max(velocities) if velocities else None
        fills = [r.fill_actual for r in active if r.fill_actual is not None]
        self.fill_max = max(fills) if fills else None

        suggested = [r.suggested_de for r in active]
        self.suggested_de = max(suggested) if active and all(s is not None for s in suggested) \
            else None

        # OK / OVERSIZED si decidono sul tratto, non sul singolo tubo: un tubo a
        # portata minore puo' risultare sovradimensionato mentre il De del
        # tratto (quello del tubo piu' caricato) e' proprio quello del modello.
        if self.status in (STATUS_OK, STATUS_OVERSIZED):
            model_des = [r.model_de for r in active if r.model_de is not None]
            if self.suggested_de is not None and model_des \
                    and min(model_des) > self.suggested_de:
                self.status = STATUS_OVERSIZED
            else:
                self.status = STATUS_OK

        self.sizes = _common_sizes(rows)
        self.model_size = find_size(self.sizes, self.model_des[0]) \
            if len(self.model_des) == 1 else None
        self.suggested_size = find_size(self.sizes, self.suggested_de)

        notes = []
        for r in rows:
            for note in r.notes:
                if note not in notes:
                    notes.append(note)
        if section is None:
            notes.append("No Section value: pipe checked on its own")
        if max(r.flow for r in rows) - min(r.flow for r in rows) > FLOW_TOLERANCE:
            notes.append("Pipes of the section carry different flows: the highest is used")
        if len(self.model_des) > 1:
            notes.append("Pipes of the section have different diameters")
        if self.sizes is None:
            notes.append("Pipe type sizes unknown")
        elif self.suggested_de is not None and self.suggested_size is None:
            notes.append("Suggested De not among the sizes of the pipe type")
        # Un tubo sottodimensionato e' in pressione: la velocita' a pelo libero
        # non esiste e la nota confonderebbe.
        if self.run != ORIENT_STACK and self.v_min is None and \
                self.status in (STATUS_OK, STATUS_OVERSIZED):
            notes.append("Velocity not available")
        self.notes = notes

        self.default_apply = (
            self.status in STATUS_TO_RESIZE
            and self.suggested_size is not None
            and not od_in(self.suggested_size.od_mm, self.model_des))

    @staticmethod
    def _table_slope(slope):
        """Riga di tabella per una pendenza simulata [cm/m]; None = pendenza del modello."""
        return rt.slope_row(slope)[0] if slope is not None else None

    def evaluate_at(self, de_mm, slope=None):
        """Esito del tratto se tutti i suoi tubi avessero diametro esterno de_mm e,
        con slope [cm/m], quella pendenza (simulazione: i tubi orizzontali del
        tratto si verificano tutti a quella riga di tabella)."""
        active = [r for r in self.rows if r.status != STATUS_NO_FLOW]
        if not active:
            return Evaluation(status=STATUS_NO_FLOW)
        table_slope = self._table_slope(slope)
        results = [evaluate(r, de_mm, self.limits, table_slope) for r in active]
        capacities = [e.capacity for e in results]
        velocities = [e.velocity for e in results if e.velocity is not None]
        fills = [e.fill_actual for e in results if e.fill_actual is not None]
        return Evaluation(
            capacity=min(capacities) if all(c is not None for c in capacities) else None,
            velocity=min(velocities) if velocities else None,
            fill_actual=max(fills) if fills else None,
            status=_worst([e.status for e in results]))

    def suggested_at(self, slope):
        """De suggerito del tratto orizzontale alla pendenza simulata [cm/m], fra le
        misure del tipo di tubazione; None se un tubo non trova un De in tabella."""
        table_slope = self._table_slope(slope)
        if table_slope is None or self.run == ORIENT_STACK:
            return None
        found = []
        for r in self.rows:
            if r.status == STATUS_NO_FLOW:
                continue
            suggestion = rt.suggest_horizontal(self.limits.fill, table_slope, r.flow,
                                               _candidates(r, rt.HORIZONTAL_DE),
                                               self.limits.vmax)
            if suggestion is None:
                return None
            found.append(suggestion[0])
        return max(found) if found else None

    @property
    def use_pct(self):
        if self.capacity is None or self.capacity <= 0:
            return None
        return self.flow / self.capacity * 100.0

    @property
    def needs_action(self):
        return self.status not in STATUS_PASS

    @property
    def model_de_label(self):
        return "/".join(fmt_mm(d) for d in self.model_des) or "-"

    @property
    def model_dn_label(self):
        return "/".join(fmt_mm(d) for d in self.model_dns) or "-"

    def section_sort(self):
        """Chiave numerica del Section per l'ordinamento (testo in coda).

        Controllo esplicito di None: in IronPython float(None) solleva una
        NullReferenceException di .NET, non un TypeError.
        """
        if self.section is None:
            return 1e12
        try:
            return float(self.section)
        except Exception:
            return 1e12


RUN_ORDER = {ORIENT_HORIZONTAL: 0, ORIENT_STACK: 1}


def group_sections(rows, limits):
    """Tratti nell'ordine sistema, Section, orizzontale prima del verticale.

    Una riga per (sistema, Section, orientamento); i tubi senza Section in coda
    al sistema, uno per riga.
    """
    groups = {}
    order = []
    for row in rows:
        section_key = row.section if row.section else "#{0}".format(row.id)
        key = (row.system, section_key, row.orientation)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(row)
    sections = [SectionResult(key, key[0], groups[key][0].section or None, groups[key], limits)
                for key in order]
    sections.sort(key=lambda s: (s.system.lower(), s.section is None,
                                 s.section_sort(), s.section or "", s.pipe_ids[0] if s.section is None else 0,
                                 RUN_ORDER.get(s.run, 2)))
    return sections


# =============================================================================
# TERMINALI E ORDINE DELLA RETE
# =============================================================================

STATUS_TERMINAL = "TERMINAL"
RUN_FIXTURE = "Fixture"


class FixtureItem(object):
    """Un terminale (plumbing fixture, per esempio un bocchettone) della rete.

    Ha gli stessi attributi di un SectionResult che servono alla finestra, ma non
    si puo' modificare: nessuna misura, nessuna spunta.
    """

    is_fixture = True

    def __init__(self, element_id, system, label, flow, dn_mm, z=None, notes=None, code=None):
        self.id = element_id
        self.element_code = code or ""
        self.key = ("fixture", element_id)
        self.group_key = None
        self.pipe_ids = [element_id]
        self.system = system
        self.label = label
        self.flow = flow
        self.dn_mm = dn_mm
        self.z = z
        self.notes = list(notes or [])
        self.section = None
        self.run = RUN_FIXTURE
        self.status = STATUS_TERMINAL
        self.sizes = None
        self.default_apply = False
        self.needs_action = False

    def section_sort(self):
        return 1e12


def _flow_of(entry):
    return entry.flow if entry.flow is not None else 0.0


def order_network(sections, fixtures, row_edges, fixture_edges):
    """Ordina tratti e terminali dal punto di monte al punto di valle.

    sections: [SectionResult]; fixtures: [FixtureItem];
    row_edges: coppie (chiave tratto, chiave tratto) collegate (attraverso raccordi);
    fixture_edges: coppie (chiave terminale, chiave tratto).

    Verso del flusso fra due tratti collegati: a valle quello con portata maggiore
    (la portata cresce verso valle); a parita' di portata quello piu' basso; poi la
    chiave. I terminali sono sempre a monte. L'elenco parte da ogni punto di valle
    (tratto senza tratti a valle, il piu' carico per primo) e risale i rami
    scrivendo prima cio' che sta a monte: ogni ramo compare dal terminale in giu',
    prima del tratto in cui confluisce. Ritorna l'elenco ordinato.

    Ogni elemento riceve anche upstream (le righe che confluiscono direttamente
    in esso) e downstream (quelle in cui confluisce), per la colonna "Fed by".
    """
    entries = {}
    for entry in list(sections) + list(fixtures):
        entries[entry.key] = entry
    upstream = dict((key, set()) for key in entries)
    downstream = dict((key, set()) for key in entries)

    def add(up, down):
        if up == down or up not in entries or down not in entries:
            return
        upstream[down].add(up)
        downstream[up].add(down)

    tolerance = 0.01
    for a, b in row_edges:
        if a not in entries or b not in entries or a == b:
            continue
        ea, eb = entries[a], entries[b]
        fa, fb = _flow_of(ea), _flow_of(eb)
        if abs(fa - fb) > tolerance:
            up, down = (a, b) if fa < fb else (b, a)
        elif ea.z is not None and eb.z is not None and abs(ea.z - eb.z) > 1.0:
            up, down = (a, b) if ea.z > eb.z else (b, a)
        else:
            up, down = (a, b) if str(a) < str(b) else (b, a)
        if up in downstream.get(down, set()):
            continue        # gia' collegati nel verso opposto: niente anelli
        add(up, down)
    for fixture_key, row_key in fixture_edges:
        add(fixture_key, row_key)

    def upstream_sort(key):
        entry = entries[key]
        return (0 if getattr(entry, "is_fixture", False) else 1, _flow_of(entry), str(key))

    order = []
    placed = set()

    def visit(root):
        # Visita in post-ordine iterativa (niente ricorsione profonda in IronPython).
        stack = [(root, False)]
        while stack:
            key, expanded = stack.pop()
            if key in placed:
                continue
            if expanded:
                placed.add(key)
                order.append(entries[key])
                continue
            stack.append((key, True))
            for up in sorted(upstream[key], key=upstream_sort, reverse=True):
                if up not in placed:
                    stack.append((up, False))

    sinks = [key for key in entries if not downstream[key]]
    sinks.sort(key=lambda k: ((entries[k].system or "").lower(), -_flow_of(entries[k]), str(k)))
    for key in sinks:
        visit(key)
    for key in sorted(entries, key=lambda k: (_flow_of(entries[k]), str(k))):
        if key not in placed:
            visit(key)
    position = dict((entry.key, index) for index, entry in enumerate(order))
    for key, entry in entries.items():
        entry.upstream = [entries[k] for k in sorted(upstream[key], key=lambda k: position[k])]
        entry.downstream = [entries[k] for k in sorted(downstream[key], key=lambda k: position[k])]
    return order
