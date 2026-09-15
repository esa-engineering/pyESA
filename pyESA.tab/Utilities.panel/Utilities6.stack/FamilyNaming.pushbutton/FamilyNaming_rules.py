# -*- coding: utf-8 -*-
"""Motore di composizione dei nomi: regole pure, nessuna API di Revit.

Prende i codici scelti nel form e restituisce le stringhe finali. Sta separato
dal modulo di mappatura perche' quello e' dichiarativo e si corregge senza
sapere programmare, questo invece e' logica e va letto come tale.

Tutte le funzioni lavorano su stringhe unicode gia' pulite: la validazione
degli input sta in validate(), che il form chiama prima di comporre.
"""

import re

import FamilyNaming_data as DATA
import FamilyNaming_map as MAP

SEP = u"_"

# Caratteri che Revit rifiuta nei nomi di famiglia e di tipo.
INVALID_NAME_CHARS = u"\\:{}[]|;<>?`~"

MAX_DESCRIPTION = 40

# I due blocchi che non si aggiungono al nome ma sostituiscono il blocco
# dimensionale: il pianerottolo che eredita lo spessore dalla rampa, e la
# rampa piena fino a terra.
DIM_SUBSTITUTES = {
    "ln_shape": u"SR",
    "rm_shape": u"SL",
}


# ---------------------------------------------------------------------------
# Accesso alle tabelle
# ---------------------------------------------------------------------------

def table(name):
    """Righe di una tabella del foglio DV, tupla vuota se il nome non esiste."""
    return DATA.TABLES.get(name, ())


def resolve_table(spec, cat_code):
    """Risolve una voce 'g1' o 'g2' della mappa.

    E' una stringa quando la scheda ha una lista sola, un dizionario quando i
    sottogruppi hanno liste distinte: Walls separa Basic Wall da Curtain Wall,
    Roofs separa Basic Roof da Sloped Glazing, Stairs ha tre liste Group1.
    """
    if isinstance(spec, dict):
        return spec.get(cat_code, u"")
    return spec or u""


def row_by_code(rows, code):
    for row in rows:
        if row and row[0] == code:
            return row
    return None


def label_of(rows, code):
    row = row_by_code(rows, code)
    return row[1] if row and len(row) > 1 else code


def description_of(rows, code):
    row = row_by_code(rows, code)
    return row[2] if row and len(row) > 2 else u""


def tm_digit_of(rows, code):
    """Prima cifra della colonna TM Code: '1xx' -> '1'. Vuoto se assente."""
    row = row_by_code(rows, code)
    if row and len(row) > 3 and row[3]:
        return row[3][0]
    return u""


# ---------------------------------------------------------------------------
# Type Mark
# ---------------------------------------------------------------------------

def type_mark_prefix(sheet, cat_code, g1_code, g2_code, manual_code=None):
    """Prefisso del Type Mark e numero di cifre del sequenziale.

    Restituisce (prefisso_completo, cifre_sequenziale). Il prefisso e' gia'
    comprensivo del trattino e dell'eventuale cifra di gruppo, cosi' che il
    Type Mark sia semplicemente prefisso + sequenziale formattato.

        DR.HN.IN  ->  ("DR-1", 2)    il sequenziale sono le ultime due cifre
        FN.CO.SE  ->  ("SE-",  3)    tutte e tre le cifre sono sequenziale
        BW.FR.IN  ->  ("BW-1", 2)
        Group1=OT ->  ("WD-",  3)    prefisso materiale, scelto dall'utente
    """
    # Il Type Mark manuale vince su tutto: Group1 = Other, oppure il campo Use
    # compilato sulle Structural Framing.
    if manual_code:
        return manual_code + u"-", 3

    mode = sheet.get("tm_from")

    if mode == "g2_code":
        return g2_code + u"-", 3

    if mode == "g2_digit":
        rows = table(resolve_table(sheet.get("g2"), cat_code))
        digit = tm_digit_of(rows, g2_code)
        return cat_code + u"-" + digit, 2

    # 'g1_digit': tutte le schede System e Structural
    rows = table(resolve_table(sheet.get("g1"), cat_code))
    digit = tm_digit_of(rows, g1_code)
    return cat_code + u"-" + digit, 2


def format_type_mark(prefix, seq_digits, sequential, use_xx=False):
    """Compone il Type Mark finale.

    Con use_xx il sequenziale non e' ancora deciso e si scrive con delle X,
    tante quante sono le cifre che prenderanno il suo posto.
    """
    if use_xx:
        return prefix + (u"X" * seq_digits)
    if sequential is None:
        return prefix
    fmt = u"{0:0" + str(seq_digits) + u"d}"
    return prefix + fmt.format(int(sequential))


def type_mark_regex(prefix, seq_digits):
    """Riconosce i Type Mark gia' assegnati con lo stesso prefisso.

    Serve a cercare il primo sequenziale libero. La ricerca e' per prefisso e
    non per categoria: sulle sette categorie a Group1 condiviso il codice di
    categoria non compare nel Type Mark, quindi e' il prefisso a definire
    l'ambito di unicita'.
    """
    return re.compile(u"^" + re.escape(prefix) + u"([0-9]{" + str(seq_digits) + u"})$")


def next_sequential(existing_marks, prefix, seq_digits):
    """Massimo sequenziale in uso piu' uno, 1 se non ce n'e' nessuno."""
    rx = type_mark_regex(prefix, seq_digits)
    best = 0
    for mark in existing_marks:
        if not mark:
            continue
        match = rx.match(mark.strip())
        if match:
            value = int(match.group(1))
            if value > best:
                best = value
    return best + 1


def sequential_overflow(sequential, seq_digits):
    """True quando il progressivo non entra piu' nelle cifre disponibili."""
    return sequential is not None and sequential > (10 ** seq_digits) - 1


# ---------------------------------------------------------------------------
# Blocco dimensionale
# ---------------------------------------------------------------------------

def format_dim_block(kind, values, substitute=None):
    """Compone il blocco dimensionale.

    values e' un dizionario di stringhe gia' in millimetri interi, o testo
    libero per le sezioni con designazione normalizzata. substitute, quando
    presente, prende il posto dell'intero blocco: SR sui pianerottoli che
    ereditano lo spessore dalla rampa, SL sulle rampe piene.
    """
    if substitute:
        return substitute

    def val(key):
        return (values.get(key) or u"").strip()

    if kind == "none":
        return u""
    if kind == "free":
        return val("text")
    if kind == "t":
        return (u"T" + val("t")) if val("t") else u""
    if kind == "h":
        return (u"H" + val("h")) if val("h") else u""
    if kind == "wxh":
        return u"{0}x{1}".format(val("w"), val("h"))
    if kind == "wxl":
        return u"{0}x{1}".format(val("w"), val("l"))
    if kind == "wxt":
        return u"{0}x{1}".format(val("w"), val("t"))
    if kind == "nfin_t":
        return u"{0}.{1}".format(val("nfin"), val("t"))
    if kind == "wxdxh":
        base = u"{0}x{1}".format(val("w"), val("d"))
        # l'altezza si scrive solo quando distingue due tipi
        return base + (u"x" + val("h") if val("h") else u"")
    if kind in ("wxd_or_d", "bxh_or_d"):
        if val("dia"):
            return u"D" + val("dia")
        first, second = ("w", "d") if kind == "wxd_or_d" else ("b", "h")
        return u"{0}x{1}".format(val(first), val(second))
    return u""


def dim_fields(kind, round_section=False):
    """Misure richieste da un blocco, nell'ordine in cui vanno chieste.

    Ogni voce e' (chiave, etichetta, obbligatoria).
    """
    if kind == "none":
        return ()
    if kind == "free":
        return (("text", u"Designation or size", True),)
    if kind == "t":
        return (("t", u"Thickness", True),)
    if kind == "h":
        return (("h", u"Height", True),)
    if kind == "wxh":
        return (("w", u"Width", True), ("h", u"Height", True))
    if kind == "wxl":
        return (("w", u"Width", True), ("l", u"Length", True))
    if kind == "wxt":
        return (("w", u"Width", True), ("t", u"Thickness", True))
    if kind == "nfin_t":
        return (("t", u"Total thickness", True),)
    if kind == "wxdxh":
        return (("w", u"Width", True), ("d", u"Depth", True),
                ("h", u"Height", False))
    if kind == "wxd_or_d":
        if round_section:
            return (("dia", u"Diameter", True),)
        return (("w", u"Width", True), ("d", u"Depth", True))
    if kind == "bxh_or_d":
        if round_section:
            return (("dia", u"Diameter", True),)
        return (("b", u"Width", True), ("h", u"Section height", True))
    return ()


# ---------------------------------------------------------------------------
# Composizione dei nomi
# ---------------------------------------------------------------------------

def _triplet(cat_code, g1_code, g2_code, in_place):
    """Il cuore del nome: [(I)]Cat.G1.G2.

    Il marcatore degli elementi in place e' attaccato al codice di categoria
    senza separatore, ed e' l'unico blocco che non e' preceduto da underscore.
    """
    marker = u"(I)" if in_place else u""
    return u"{0}{1}.{2}.{3}".format(marker, cat_code, g1_code, g2_code)


def compose_family_name(sheet, values):
    """Nome della famiglia. Solo per le schede a schema 'loadable'."""
    parts = [MAP.AUTHOR_CODE]
    parts.append(_triplet(values.get("cat", u""), values.get("g1", u""),
                          values.get("g2", u""), values.get("in_place")))

    for block in sheet.get("blocks", ()):
        if block in DIM_SUBSTITUTES:
            continue
        chunk = _family_block(block, values)
        if chunk:
            parts.append(chunk)

    return SEP.join(parts)


def _family_block(block, values):
    """Traduce un blocco opzionale nella stringa che finisce nel nome."""
    if block == "leaves":
        leaves = values.get("leaves")
        # un componente nidificato o un simbolo non hanno ante: il campo si
        # spegne e il blocco sparisce dal nome
        if leaves in (None, u""):
            return u""
        return u"L" + unicode_str(leaves)
    if block == "cw_hosted":
        return u"CW" if values.get("cw_hosted") else u""
    if block == "entrance":
        return u"EN" if values.get("entrance") else u""
    if block == "rei":
        return (values.get("rei") or u"").strip()
    if block == "use":
        return (values.get("use") or u"").strip()
    if block == "role":
        return (values.get("role") or u"").strip()
    if block == "custom":
        return u"CT" if values.get("custom") else u""
    if block == "manufacturer":
        return (values.get("manufacturer") or u"").strip()
    if block == "brand":
        return (values.get("brand") or u"").strip()
    if block == "description":
        return (values.get("description") or u"").strip()
    return u""


def compose_type_name(sheet, values):
    """Nome del tipo.

    Sulle schede 'loadable' e' TypeMark _ blocco dimensionale _ [Descrizione]:
    il nome della famiglia porta gia' la classificazione, il tipo porta le
    misure. Sulle schede 'system' non esiste un nome famiglia da comporre e il
    nome del tipo porta tutto, terna compresa.
    """
    dim = format_dim_block(
        sheet.get("dim"),
        values.get("dim", {}),
        _active_substitute(sheet, values),
    )

    if sheet.get("schema") == "loadable":
        parts = [values.get("type_mark", u"")]
        if dim:
            parts.append(dim)
        type_desc = (values.get("type_description") or u"").strip()
        if type_desc:
            parts.append(type_desc)
        return SEP.join(p for p in parts if p)

    # schema 'system'
    parts = [MAP.AUTHOR_CODE, values.get("type_mark", u"")]
    parts.append(_triplet(values.get("cat", u""), values.get("g1", u""),
                          values.get("g2", u""), values.get("in_place")))
    if dim:
        parts.append(dim)

    for block in sheet.get("blocks", ()):
        if block in DIM_SUBSTITUTES:
            continue
        chunk = _type_block(block, values)
        if chunk:
            parts.append(chunk)

    return SEP.join(p for p in parts if p)


def _type_block(block, values):
    if block == "rei":
        return (values.get("rei") or u"").strip()
    if block == "wi":
        # i due suffissi si scrivono attaccati: un elemento che ha entrambi
        # finisce con _WI
        return (u"W" if values.get("waterproof") else u"") + \
               (u"I" if values.get("insulation") else u"")
    if block == "rail_top":
        return (values.get("rail_top") or u"").strip()
    if block == "underside":
        return (values.get("underside") or u"").strip()
    if block == "description":
        return (values.get("description") or u"").strip()
    return u""


def _active_substitute(sheet, values):
    for block in sheet.get("blocks", ()):
        if block in DIM_SUBSTITUTES and values.get(block):
            return DIM_SUBSTITUTES[block]
    return None


def compose_in_place_name(family_name, type_name):
    """Nome unico per un elemento modellato in place.

    Un in place non ha due nomi indipendenti su cui scrivere, quindi i due
    blocchi si uniscono in una stringa sola. Sulle categorie di sistema il
    nome della famiglia non esiste come regola: resta il solo nome del tipo.
    """
    if not family_name:
        return type_name
    if not type_name:
        return family_name
    return u"{0} - {1}".format(family_name, type_name)


# ---------------------------------------------------------------------------
# Validazione
# ---------------------------------------------------------------------------

def unicode_str(value):
    if value is None:
        return u""
    try:
        return unicode(value)          # noqa: F821  IronPython 2.7
    except NameError:
        return str(value)


def invalid_chars(text):
    """Caratteri vietati da Revit trovati nel testo, ordinati e senza ripetizioni."""
    return sorted(set(ch for ch in (text or u"") if ch in INVALID_NAME_CHARS))


def validate(sheet, values):
    """Controlla i campi prima di comporre.

    Restituisce una lista di messaggi. Lista vuota significa che si puo'
    procedere. I controlli sono tutti bloccanti: la descrizione non viene
    sanitizzata, gli spazi sono ammessi, quindi l'unico limite sui testi
    liberi e' quello che Revit stesso impone.
    """
    problems = []

    if not values.get("g1"):
        problems.append(u"Group 1 is not set.")
    if not values.get("g2"):
        problems.append(u"Group 2 is not set.")

    mark = (values.get("type_mark") or u"").strip()
    if not mark:
        problems.append(u"The Type Mark is empty.")

    # blocco dimensionale
    substitute = _active_substitute(sheet, values)
    if not substitute:
        round_section = _is_round(sheet, values)
        dim_values = values.get("dim", {})
        for key, label, required in dim_fields(sheet.get("dim"), round_section):
            if required and not (dim_values.get(key) or u"").strip():
                if _dim_is_optional(sheet, values):
                    continue
                problems.append(u"{0} is required by the dimensional block.".format(label))

    # descrizioni
    for key, label in (("description", u"Family description"),
                       ("type_description", u"Type description")):
        text = (values.get(key) or u"").strip()
        if len(text) > MAX_DESCRIPTION:
            problems.append(
                u"{0} is {1} characters long, the limit is {2}.".format(
                    label, len(text), MAX_DESCRIPTION))

    # Manufacturer e Custom non convivono: uno standard su misura non ha un
    # produttore a catalogo
    if values.get("manufacturer"):
        if values.get("custom") or values.get("g1") == u"CT":
            problems.append(
                u"Manufacturer cannot be used together with Custom.")

    # caratteri vietati, su tutto quello che finisce in un nome
    for key, label in (("manufacturer", u"Manufacturer"),
                       ("brand", u"Brand"),
                       ("rei", u"REI"),
                       ("description", u"Family description"),
                       ("type_description", u"Type description"),
                       ("type_mark", u"Type Mark")):
        bad = invalid_chars(values.get(key))
        if bad:
            problems.append(
                u"{0} contains characters Revit does not allow: {1}".format(
                    label, u" ".join(bad)))

    return problems


def _is_round(sheet, values):
    """True quando il Group1 dichiara una sezione tonda: si quota il diametro."""
    return values.get("g1") in (sheet.get("round_g1") or ())


def _dim_is_optional(sheet, values):
    """True quando il Group1 scelto non ha blocco dimensionale.

    Unico caso oggi: il pannello di sistema vuoto, che non ha spessore perche'
    ospita una famiglia caricabile.
    """
    return values.get("g1") in (sheet.get("dim_optional_g1") or ())
