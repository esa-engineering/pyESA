# -*- coding: utf-8 -*-
"""
mepqto_rules.py - regole di misura delle categorie lineari e loro valori di default.

Le formule vengono dai fogli QTO aziendali (XXXXEN_QTO_Ducts, _Pipes, _Flex Ducts,
_Cable Trays, _Conduits):

* canali    m  = L
            mq = perimetro x L, perimetro = pi*D (circolari) o 2*(W+H) (rettangolari
                 e ovali, come nel foglio)
            kg = kg/mq della lamiera (tabella per lato maggiore, diversa per circolari
                 e rettangolari) x mq
* tubi      m  = L
            mq = pi*De x L (superficie esterna, non presente nel foglio)
            kg = densita' per Type Mark x pi/4*(De^2 - Di^2) x L
* flessibili m = L, mq = perimetro x L
* passerelle e conduit: sempre m = L
* isolanti  mq = perimetro esterno dell'isolante x L: pi*(D+2s), 2*(W+H+4s) sui
                 canali rettangolari, pi*(De+2s) sui tubi; m = L

Ogni quantita' lineare e' maggiorata della percentuale della sua categoria per
raccordi e sfridi (nei fogli il 30%), oppure di quella dell'override impostato sulla
singola voce del computo. Le tre tabelle sono modificabili nella scheda Rules e si
salvano nel file di progetto: questo modulo non tocca Revit.
"""

import math

M = u"m"
MQ = u"mq"
KG = u"kg"

KIND_COUNT = "count"
KIND_DUCT = "duct"
KIND_FLEX_DUCT = "flexduct"
KIND_PIPE = "pipe"
KIND_FLEX_PIPE = "flexpipe"
KIND_TRAY = "tray"
KIND_CONDUIT = "conduit"
KIND_DUCT_INSULATION = "ductins"
KIND_PIPE_INSULATION = "pipeins"

# Unita' ammesse per ogni tipo di misura (None = qualsiasi: le categorie a pezzo
# contano le istanze qualunque sia l'unita' della voce).
ALLOWED_UNITS = {
    KIND_COUNT: None,
    KIND_DUCT: (M, MQ, KG),
    KIND_FLEX_DUCT: (M, MQ),
    KIND_PIPE: (M, MQ, KG),
    KIND_FLEX_PIPE: (M, MQ),
    KIND_TRAY: (M,),
    KIND_CONDUIT: (M,),
    KIND_DUCT_INSULATION: (MQ, M),
    KIND_PIPE_INSULATION: (MQ, M),
}

# Categorie computate sempre nella stessa unita', qualunque sia quella della voce.
FIXED_UNIT = {KIND_TRAY: M, KIND_CONDUIT: M}

SHAPE_RECT = "rect"
SHAPE_ROUND = "round"
SHAPE_LABELS = {SHAPE_RECT: u"Rectangular / oval", SHAPE_ROUND: u"Round"}

# Maggiorazione per raccordi e sfridi, come frazione (0.30 = 30%).
DEFAULT_ALLOWANCE = {
    "OST_DuctCurves": 0.30,
    "OST_FlexDuctCurves": 0.30,
    "OST_PipeCurves": 0.30,
    "OST_FlexPipeCurves": 0.30,
    "OST_CableTray": 0.0,
    "OST_Conduit": 0.0,
    "OST_DuctInsulations": 0.30,
    "OST_PipeInsulations": 0.30,
}

# kg/mq della lamiera per fascia di lato maggiore [mm]: la fascia vale quando il lato
# maggiore (o il diametro) e' MINORE del limite; l'ultima fascia non ha limite.
# Valori di '02_REGOLE DI CALCOLO' del foglio Ducts.
DEFAULT_DUCT_WEIGHT = {
    SHAPE_RECT: [(300.0, 5.1), (750.0, 6.7), (1200.0, 8.2), (2000.0, 9.8), (None, 12.0)],
    SHAPE_ROUND: [(300.0, 4.8), (750.0, 6.4), (1200.0, 8.0), (2000.0, 9.6), (None, 12.0)],
}

# Densita' [kg/mc] per Type Mark, dal foglio Pipes ('2.2 Codifica Tubazioni').
DEFAULT_PIPE_DENSITY = [
    (u"CS01", u"Carbon steel, L series EN 10255", 7820.0),
    (u"CS02", u"Carbon steel, M series UNI EN 10255", 7950.0),
    (u"CS03", u"Carbon steel, H series UNI EN 10255", 7850.0),
    (u"CS04", u"Carbon steel, over size EN 10216", 7900.0),
    (u"GS01", u"Galvanized steel, L series EN 10255", 8300.0),
    (u"GS02", u"Galvanized steel, M series UNI EN 10255", 8100.0),
    (u"GS03", u"Galvanized steel, H series UNI EN 10255", 8050.0),
    (u"GS04", u"Galvanized steel, over size EN 10216", 7900.0),
]

# Esiti di measure() oltre al valore.
PROBLEM_UNIT = "unit"                # unita' mancante o non ammessa: non computato
PROBLEM_UNIT_FORCED = "unit_forced"  # unita' diversa da quella fissa: computato in m
PROBLEM_GEOMETRY = "geometry"        # dimensioni mancanti per l'unita' richiesta
PROBLEM_DENSITY = "density"          # tubo al kg senza densita' per il Type Mark
PROBLEM_HOST = "host"                # isolante su raccordo/accessorio: non misurabile


# =============================================================================
# TABELLE
# =============================================================================

def _float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def sorted_bands(bands):
    """Fasce ordinate per limite crescente, quella senza limite in fondo."""
    return sorted(bands, key=lambda band: (band[0] is None, band[0] or 0.0))


class Rules(object):
    """Maggiorazioni per categoria, kg/mq dei canali, densita' dei tubi."""

    def __init__(self, allowance, duct_weight, pipe_density):
        self.allowance = allowance          # {chiave categoria: frazione}
        self.duct_weight = duct_weight      # {forma: [(limite mm o None, kg/mq)]}
        self.pipe_density = pipe_density    # [(Type Mark, nota, kg/mc)]

    @classmethod
    def defaults(cls):
        return cls(dict(DEFAULT_ALLOWANCE),
                   dict((shape, list(bands)) for shape, bands in DEFAULT_DUCT_WEIGHT.items()),
                   list(DEFAULT_PIPE_DENSITY))

    @classmethod
    def from_dict(cls, data):
        """Regole dal file di progetto; quello che manca o e' illeggibile prende il default."""
        rules = cls.defaults()
        if not isinstance(data, dict):
            return rules

        for key, value in (data.get("allowance") or {}).items():
            fraction = _float(value)
            if key in rules.allowance and fraction is not None and fraction >= 0:
                rules.allowance[key] = fraction

        for shape, bands in (data.get("duct_weight") or {}).items():
            if shape not in DEFAULT_DUCT_WEIGHT or not isinstance(bands, list):
                continue
            parsed = []
            for band in bands:
                if not isinstance(band, (list, tuple)) or len(band) != 2:
                    continue
                limit = None if band[0] is None else _float(band[0])
                weight = _float(band[1])
                if weight is not None and (band[0] is None or limit is not None):
                    parsed.append((limit, weight))
            if parsed:
                rules.duct_weight[shape] = sorted_bands(parsed)

        densities = data.get("pipe_density")
        if isinstance(densities, list):
            parsed = []
            for entry in densities:
                if not isinstance(entry, (list, tuple)) or len(entry) != 3:
                    continue
                density = _float(entry[2])
                if entry[0] and density is not None:
                    parsed.append((u"{}".format(entry[0]), u"{}".format(entry[1] or u""), density))
            rules.pipe_density = parsed
        return rules

    def to_dict(self):
        return {
            "allowance": dict(self.allowance),
            "duct_weight": dict((shape, [list(band) for band in sorted_bands(bands)])
                                for shape, bands in self.duct_weight.items()),
            "pipe_density": [list(entry) for entry in self.pipe_density],
        }

    def allowance_for(self, category_key):
        return self.allowance.get(category_key, 0.0)

    def duct_kg_per_m2(self, shape, max_side_mm):
        if not max_side_mm or max_side_mm <= 0:
            return None
        for limit, weight in sorted_bands(self.duct_weight.get(shape, [])):
            if limit is None or max_side_mm < limit:
                return weight
        return None

    def density_for(self, type_mark):
        for mark, _, density in self.pipe_density:
            if mark == type_mark:
                return density
        return None


# =============================================================================
# GEOMETRIA E MISURA
# =============================================================================

class Geometry(object):
    """Dimensioni lette dal modello, in mm e m. host: per gli isolanti, la geometria
    dell'elemento isolato (None se e' un raccordo o un accessorio)."""

    __slots__ = ("length_m", "diameter_mm", "width_mm", "height_mm",
                 "outer_mm", "inner_mm", "thickness_mm", "host")

    def __init__(self, length_m=0.0, diameter_mm=0.0, width_mm=0.0, height_mm=0.0,
                 outer_mm=0.0, inner_mm=0.0, thickness_mm=0.0, host=None):
        self.length_m = length_m or 0.0
        self.diameter_mm = diameter_mm or 0.0
        self.width_mm = width_mm or 0.0
        self.height_mm = height_mm or 0.0
        self.outer_mm = outer_mm or 0.0
        self.inner_mm = inner_mm or 0.0
        self.thickness_mm = thickness_mm or 0.0
        self.host = host

    @property
    def is_round(self):
        return self.diameter_mm > 0

    @property
    def shape(self):
        return SHAPE_ROUND if self.is_round else SHAPE_RECT

    @property
    def max_side_mm(self):
        return max(self.diameter_mm, self.width_mm, self.height_mm)

    @property
    def section_perimeter_m(self):
        """Perimetro di un canale: pi*D o 2*(W+H). 0 se mancano le dimensioni."""
        if self.is_round:
            return math.pi * self.diameter_mm / 1000.0
        if self.width_mm > 0 and self.height_mm > 0:
            return 2.0 * (self.width_mm + self.height_mm) / 1000.0
        return 0.0

    @property
    def pipe_outer_mm(self):
        return self.outer_mm or self.diameter_mm


def _insulation_perimeter_m(kind, geometry):
    host = geometry.host
    t = geometry.thickness_mm
    if host is None or t <= 0:
        return 0.0
    if kind == KIND_PIPE_INSULATION:
        outer = host.pipe_outer_mm
        return math.pi * (outer + 2.0 * t) / 1000.0 if outer > 0 else 0.0
    if host.is_round:
        return math.pi * (host.diameter_mm + 2.0 * t) / 1000.0
    if host.width_mm > 0 and host.height_mm > 0:
        return 2.0 * (host.width_mm + host.height_mm + 4.0 * t) / 1000.0
    return 0.0


def effective_unit(kind, unit):
    """(unita' usata, problema) per una voce di quell'unita' su quel tipo di misura."""
    allowed = ALLOWED_UNITS.get(kind)
    if allowed is None:
        return unit, None
    if kind in FIXED_UNIT:
        fixed = FIXED_UNIT[kind]
        return fixed, (PROBLEM_UNIT_FORCED if unit and unit != fixed else None)
    if unit in allowed:
        return unit, None
    return None, PROBLEM_UNIT


def measure(kind, geometry, unit, type_mark, category_key, rules, allowance=None):
    """(quantita' gia' maggiorata o None, unita' usata, problema o None).

    allowance: maggiorazione (frazione) che sostituisce quella della categoria, per le
    voci del computo con un override; None = maggiorazione della categoria.
    """
    if kind == KIND_COUNT:
        return 1.0, unit, None

    used, problem = effective_unit(kind, unit)
    if used is None:
        return None, None, problem
    if geometry is None:
        return None, used, PROBLEM_GEOMETRY

    length = geometry.length_m
    value = None
    if kind in (KIND_DUCT_INSULATION, KIND_PIPE_INSULATION) and geometry.host is None:
        return None, used, PROBLEM_HOST

    if used == M:
        value = length
    elif used == MQ:
        if kind in (KIND_DUCT, KIND_FLEX_DUCT):
            perimeter = geometry.section_perimeter_m
        elif kind in (KIND_PIPE, KIND_FLEX_PIPE):
            perimeter = math.pi * geometry.pipe_outer_mm / 1000.0
        else:
            perimeter = _insulation_perimeter_m(kind, geometry)
        value = perimeter * length if perimeter > 0 else None
    elif used == KG:
        if kind == KIND_DUCT:
            kg_m2 = rules.duct_kg_per_m2(geometry.shape, geometry.max_side_mm)
            perimeter = geometry.section_perimeter_m
            if kg_m2 is not None and perimeter > 0:
                value = kg_m2 * perimeter * length
        elif kind == KIND_PIPE:
            density = rules.density_for(type_mark)
            if density is None:
                return None, used, PROBLEM_DENSITY
            outer = geometry.pipe_outer_mm / 1000.0
            inner = geometry.inner_mm / 1000.0
            if outer > inner > 0:
                value = density * math.pi / 4.0 * (outer ** 2 - inner ** 2) * length

    if value is None or length <= 0:
        return None, used, PROBLEM_GEOMETRY
    if allowance is None:
        allowance = rules.allowance_for(category_key)
    return value * (1.0 + allowance), used, problem
