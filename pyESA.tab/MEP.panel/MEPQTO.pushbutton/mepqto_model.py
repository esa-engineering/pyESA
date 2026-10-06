# -*- coding: utf-8 -*-
"""
mepqto_model.py - raccolta dal modello e aggregazione del computo MEP.

La parte Revit (collect_records) legge una volta sola gli elementi delle categorie
computabili e li riduce a record semplici, con la geometria gia' convertita in mm
e m per le categorie lineari. Tutto il resto (aggregazione per Type Mark,
quantita' per codice, anomalie) lavora sui record in Python puro: la finestra
ricalcola al volo quando cambiano categorie, unita' delle voci o regole, senza
tornare sul modello.

Misure: le categorie a pezzo contano le istanze (1 istanza = 1 pezzo); quelle
lineari si misurano in m, mq o kg secondo l'unita' della voce, con le formule e
le maggiorazioni di mepqto_rules.
"""

from collections import OrderedDict

from System.Collections.Generic import List

from pyrevit import DB

import mepqto_rules as qr

BIP = DB.BuiltInParameter
ST = DB.StorageType

FT_TO_MM = 304.8
FT_TO_M = 0.3048

# Chiave BuiltInCategory, etichetta inglese mostrata in UI ed Excel, tipo di misura.
CATEGORY_RULES = (
    ("OST_DuctTerminal", "Air Terminals", qr.KIND_COUNT),
    ("OST_CommunicationDevices", "Communication Devices", qr.KIND_COUNT),
    ("OST_ConduitFitting", "Conduit Fittings", qr.KIND_COUNT),
    ("OST_DataDevices", "Data Devices", qr.KIND_COUNT),
    ("OST_DuctAccessory", "Duct Accessories", qr.KIND_COUNT),
    ("OST_ElectricalEquipment", "Electrical Equipment", qr.KIND_COUNT),
    ("OST_ElectricalFixtures", "Electrical Fixtures", qr.KIND_COUNT),
    ("OST_FireAlarmDevices", "Fire Alarm Devices", qr.KIND_COUNT),
    ("OST_LightingDevices", "Lighting Devices", qr.KIND_COUNT),
    ("OST_LightingFixtures", "Lighting Fixtures", qr.KIND_COUNT),
    ("OST_MechanicalEquipment", "Mechanical Equipment", qr.KIND_COUNT),
    ("OST_NurseCallDevices", "Nurse Call Devices", qr.KIND_COUNT),
    ("OST_PipeAccessory", "Pipe Accessories", qr.KIND_COUNT),
    ("OST_PlumbingFixtures", "Plumbing Fixtures", qr.KIND_COUNT),
    ("OST_SecurityDevices", "Security Devices", qr.KIND_COUNT),
    ("OST_Sprinklers", "Sprinklers", qr.KIND_COUNT),
    ("OST_SpecialityEquipment", "Specialty Equipment", qr.KIND_COUNT),
    ("OST_DuctCurves", "Ducts", qr.KIND_DUCT),
    ("OST_FlexDuctCurves", "Flex Ducts", qr.KIND_FLEX_DUCT),
    ("OST_PipeCurves", "Pipes", qr.KIND_PIPE),
    ("OST_FlexPipeCurves", "Flex Pipes", qr.KIND_FLEX_PIPE),
    ("OST_CableTray", "Cable Trays", qr.KIND_TRAY),
    ("OST_Conduit", "Conduits", qr.KIND_CONDUIT),
    ("OST_DuctInsulations", "Duct Insulation", qr.KIND_DUCT_INSULATION),
    ("OST_PipeInsulations", "Pipe Insulation", qr.KIND_PIPE_INSULATION),
)

LINEAR_KEYS = tuple(key for key, _, kind in CATEGORY_RULES if kind != qr.KIND_COUNT)

# Categorie che possono ospitare un isolante misurabile (i raccordi no).
DUCT_HOST_KEYS = ("OST_DuctCurves", "OST_FlexDuctCurves")
PIPE_HOST_KEYS = ("OST_PipeCurves", "OST_FlexPipeCurves")

CODE_SLOTS = 10
# Posizioni dei codici di un elemento: 0..9 dai parametri di tipo, 10..19 da quelli
# d'istanza. Il riepilogo Type Mark ha una coppia di colonne per posizione.
TOTAL_SLOTS = 2 * CODE_SLOTS
INSTANCE_SLOT_OFFSET = CODE_SLOTS
# Default dei parametri, modificabili con il pulsante Parameters... (ParameterMap).
# Codici di tipo: solo categorie a pezzo (con ripiego sull'istanza con lo stesso nome).
PRICE_CODE_PARAMS = tuple(u"e_DAT_PriceCode_{}".format(i) for i in range(1, CODE_SLOTS + 1))
# Codici d'istanza: tutte le categorie. Per le lineari sono gli unici, perche' lo stesso
# tipo cambia voce con il diametro o la sezione; per quelle a pezzo si sommano ai codici
# di tipo.
INSTANCE_PRICE_CODE_PARAMS = tuple(u"e_DAT_PriceCode_i_{}".format(i)
                                   for i in range(1, CODE_SLOTS + 1))
# Si/No d'istanza per togliere elementi dal computo, su tutte le categorie. Solo il No
# esplicito esclude.
INCLUDE_PARAM = u"e_DAT_BOQ_i"

# Etichette delle anomalie: finiscono in UI ed Excel, quindi in inglese.
ISSUE_NO_TYPE_MARK = "No Type Mark"
ISSUE_NO_CODE = "No price code"
ISSUE_INCONSISTENT = "Inconsistent price codes"
ISSUE_NO_DESCRIPTION = "Code without description"
ISSUE_NO_PARAMS = "Price code parameters not found"
ISSUE_UNIT = "Unit not valid for the category"
ISSUE_UNIT_FORCED = "Unit replaced"
ISSUE_GEOMETRY = "Dimensions missing"
ISSUE_DENSITY = "No pipe density"
ISSUE_HOST = "Insulation on fittings not measured"
ISSUE_EXCLUDED = "Excluded from the bill"

# WBS: fino a 15 livelli, ognuno legato a un parametro scelto dall'utente.
WBS_LEVELS = 15
WBS_NOT_SET = u"(not set)"

PHASE_STATUS_NEW = "New"
PHASE_STATUS_NEW_EXISTING = "New + Existing"
PHASE_STATUS_OPTIONS = (PHASE_STATUS_NEW, PHASE_STATUS_NEW_EXISTING)


# =============================================================================
# UTILITY
# =============================================================================

def get_element_id_value(eid):
    """ElementId.IntegerValue e' stato rimosso in Revit 2026 (sostituito da .Value)."""
    if eid is None:
        return -1
    if hasattr(eid, "Value"):
        return eid.Value
    return eid.IntegerValue


def element_name(element):
    """Nome di un elemento: el.Name fallisce in silenzio su molte sottoclassi in IronPython."""
    try:
        return DB.Element.Name.GetValue(element) or ""
    except Exception:
        try:
            return element.Name or ""
        except Exception:
            return ""


def available_rules():
    """CATEGORY_RULES filtrate sulle BuiltInCategory esistenti nella versione di Revit."""
    rules = []
    for key, label, kind in CATEGORY_RULES:
        bic = getattr(DB.BuiltInCategory, key, None)
        if bic is not None:
            rules.append((key, label, kind, bic))
    return rules


def category_label(key):
    for rule_key, label, _ in CATEGORY_RULES:
        if rule_key == key:
            return label
    return key


def category_kind(key):
    for rule_key, _, kind in CATEGORY_RULES:
        if rule_key == key:
            return kind
    return qr.KIND_COUNT


def _param_text(param):
    """Valore testuale di un parametro, stringa vuota se assente o vuoto."""
    if param is None or not param.HasValue:
        return u""
    try:
        if param.StorageType == ST.String:
            value = param.AsString()
        else:
            value = param.AsValueString()
    except Exception:
        return u""
    return (value or u"").strip()


def _double_ft(element, *names):
    """Primo parametro numerico non nullo fra i BuiltInParameter indicati, in piedi."""
    for name in names:
        bip = getattr(BIP, name, None)
        if bip is None:
            continue
        try:
            param = element.get_Parameter(bip)
            if param is not None and param.HasValue and param.StorageType == ST.Double:
                value = param.AsDouble()
                if value:
                    return value
        except Exception:
            continue
    return 0.0


# =============================================================================
# RACCOLTA (lato Revit)
# =============================================================================

class InstanceRecord(object):
    """Un elemento computabile, ridotto ai soli dati che servono al computo."""

    __slots__ = ("element_id", "category_key", "kind", "type_mark", "type_label",
                 "slots", "codes", "nested", "geometry", "wbs")

    def __init__(self, element_id, category_key, type_mark, type_label, slots, nested,
                 geometry=None, wbs=()):
        self.element_id = element_id
        self.category_key = category_key
        self.kind = category_kind(category_key)
        self.type_mark = type_mark
        self.type_label = type_label
        # slots: TOTAL_SLOTS valori nella posizione del loro parametro ("" se vuoto):
        # prima i 10 codici di tipo, poi i 10 d'istanza
        self.slots = tuple(slots)
        # codes: tutti i codici, tipo e istanza, in ordine di parametro, senza vuoti ne'
        # duplicati (un codice sia sul tipo sia sull'istanza conta una volta)
        self.codes = codes_from_slots(self.slots)
        self.nested = nested
        # geometry: qr.Geometry per le categorie lineari, None per quelle a pezzo
        self.geometry = geometry
        # wbs: un valore per ogni livello WBS attivo ("" se l'elemento non ne ha)
        self.wbs = tuple(wbs)


def _names_label(names):
    """Etichetta compatta di una lista di parametri: "e_DAT_PriceCode_1..10" per una
    serie numerata completa, altrimenti i nomi separati da virgola."""
    used = [name for name in names if name]
    if not used:
        return u"(none)"
    if len(used) > 2:
        stem = used[0][:-1]
        if used[0].endswith(u"1") and all(name == u"{}{}".format(stem, index + 1)
                                          for index, name in enumerate(used)):
            return u"{}1..{}".format(stem, len(used))
    return u", ".join(used)


class ParameterMap(object):
    """Parametri letti dal modello: codici delle voci (piece_codes sul tipo delle
    categorie a pezzo, linear_codes sull'istanza di tutte le categorie) e Si/No
    d'istanza di inclusione nel computo (include, tutte le categorie).

    linear_codes conserva il nome storico perche' e' la chiave del file di progetto;
    include si salva come "linear_include" per lo stesso motivo. La vecchia chiave
    "piece_include" (Si/No di tipo) si ignora.
    """

    def __init__(self, piece_codes, linear_codes, include):
        self.piece_codes = self._slots(piece_codes)
        self.linear_codes = self._slots(linear_codes)
        self.include = (include or u"").strip()

    @staticmethod
    def _slots(names):
        names = [(u"{}".format(name or u"")).strip() for name in (names or [])][:CODE_SLOTS]
        return tuple(names + [u""] * (CODE_SLOTS - len(names)))

    @classmethod
    def defaults(cls):
        return cls(PRICE_CODE_PARAMS, INSTANCE_PRICE_CODE_PARAMS, INCLUDE_PARAM)

    @classmethod
    def from_dict(cls, data):
        """Mappa dal file di progetto; le chiavi mancanti prendono il default."""
        default = cls.defaults()
        if not isinstance(data, dict):
            return default

        def pick(key, fallback):
            value = data.get(key)
            return fallback if value is None else value

        return cls(pick("piece_codes", default.piece_codes),
                   pick("linear_codes", default.linear_codes),
                   pick("linear_include", default.include))

    def to_dict(self):
        return {"piece_codes": list(self.piece_codes), "linear_codes": list(self.linear_codes),
                "linear_include": self.include}

    def __eq__(self, other):
        return isinstance(other, ParameterMap) and self.to_dict() == other.to_dict()

    def __ne__(self, other):
        return not self.__eq__(other)

    @property
    def is_default(self):
        return self == ParameterMap.defaults()

    @property
    def piece_label(self):
        return _names_label(self.piece_codes)

    @property
    def linear_label(self):
        return _names_label(self.linear_codes)


def flag_is_no(param):
    """True solo per un Si/No valorizzato a No (o un testo "No"/"False"/"0")."""
    if param is None:
        return False
    try:
        if not param.HasValue:
            return False
        if param.StorageType == ST.Integer:
            return param.AsInteger() == 0
        return _param_text(param).lower() in (u"no", u"false", u"0", u"n")
    except Exception:
        return False


class CollectOptions(object):
    def __init__(self, phase=None, phase_status=PHASE_STATUS_NEW, primary_only=True,
                 wbs_names=(), param_map=None):
        self.phase = phase
        self.phase_status = phase_status
        self.primary_only = primary_only
        # parametri dei livelli WBS attivi, gia' senza i livelli lasciati vuoti
        self.wbs_names = tuple(wbs_names)
        self.param_map = param_map or ParameterMap.defaults()


class CollectResult(object):
    def __init__(self):
        self.records = []
        # True se almeno un tipo o un'istanza porta uno dei parametri dei codici.
        self.params_found = False
        self.skipped_options = 0
        self.param_map = ParameterMap.defaults()
        # elementi tolti dal Si/No: [(chiave categoria, etichetta tipo, origine, ElementId)]
        self.excluded = []


class _TypeInfo(object):
    """Dati letti una volta per tipo."""

    __slots__ = ("type_mark", "label", "codes", "instance_names")

    def __init__(self, type_mark, label, codes, instance_names):
        self.type_mark = type_mark
        self.label = label
        # codes: {indice parametro: codice} letti sul tipo.
        self.codes = codes
        # Parametri dei codici che il tipo non ha: si cercano sull'istanza.
        self.instance_names = instance_names


def _read_type(elem_type, result, param_map):
    names = [(index, name) for index, name in enumerate(param_map.piece_codes) if name]
    if elem_type is None:
        return _TypeInfo(u"", u"", {}, tuple(names))

    type_mark = _param_text(elem_type.get_Parameter(BIP.ALL_MODEL_TYPE_MARK))
    try:
        family_name = elem_type.FamilyName or u""
    except Exception:
        family_name = u""
    label = u"{} : {}".format(family_name, element_name(elem_type)) if family_name \
        else element_name(elem_type)

    codes = {}
    instance_names = []
    for index, name in names:
        param = elem_type.LookupParameter(name)
        if param is None:
            instance_names.append((index, name))
            continue
        result.params_found = True
        value = _param_text(param)
        if value:
            codes[index] = value
    return _TypeInfo(type_mark, label, codes, tuple(instance_names))


def codes_from_slots(slots):
    ordered = []
    for code in slots:
        if code and code not in ordered:
            ordered.append(code)
    return tuple(ordered)


def _read_slots(type_info, instance):
    """Categorie a pezzo: i 10 codici di tipo nella posizione del loro parametro; se il
    tipo non ha il parametro, si legge quello d'istanza con lo stesso nome."""
    slots = [u""] * CODE_SLOTS
    for index, code in type_info.codes.items():
        slots[index] = code
    for index, name in type_info.instance_names:
        param = instance.LookupParameter(name)
        if param is None:
            continue
        value = _param_text(param)
        if value:
            slots[index] = value
    return tuple(slots)


def _read_instance_slots(element, result, names=INSTANCE_PRICE_CODE_PARAMS):
    """Tutte le categorie: i codici d'istanza nella posizione del loro parametro."""
    slots = []
    for name in names:
        if not name:
            slots.append(u"")
            continue
        param = element.LookupParameter(name)
        if param is not None:
            result.params_found = True
        slots.append(_param_text(param))
    return tuple(slots)


def _instance_has_code_params(type_info, instance):
    for _, name in type_info.instance_names:
        if instance.LookupParameter(name) is not None:
            return True
    return False


def _phase_filter(options):
    if options.phase is None:
        return None
    statuses = List[DB.ElementOnPhaseStatus]()
    statuses.Add(DB.ElementOnPhaseStatus.New)
    if options.phase_status == PHASE_STATUS_NEW_EXISTING:
        statuses.Add(DB.ElementOnPhaseStatus.Existing)
    return DB.ElementPhaseStatusFilter(options.phase.Id, statuses)


# --- geometria ---------------------------------------------------------------

def _curve_geometry(element, key):
    """Lunghezza e sezione di canali, tubi, flessibili, passerelle e conduit."""
    length = _double_ft(element, "CURVE_ELEM_LENGTH") * FT_TO_M
    if key in DUCT_HOST_KEYS:
        return qr.Geometry(
            length_m=length,
            diameter_mm=_double_ft(element, "RBS_CURVE_DIAMETER_PARAM") * FT_TO_MM,
            width_mm=_double_ft(element, "RBS_CURVE_WIDTH_PARAM") * FT_TO_MM,
            height_mm=_double_ft(element, "RBS_CURVE_HEIGHT_PARAM") * FT_TO_MM)
    if key in PIPE_HOST_KEYS:
        return qr.Geometry(
            length_m=length,
            diameter_mm=_double_ft(element, "RBS_PIPE_DIAMETER_PARAM") * FT_TO_MM,
            outer_mm=_double_ft(element, "RBS_PIPE_OUTER_DIAMETER") * FT_TO_MM,
            inner_mm=_double_ft(element, "RBS_PIPE_INNER_DIAM_PARAM") * FT_TO_MM)
    return qr.Geometry(length_m=length)


def _insulation_geometry(doc, insulation, key_by_cat_id, host_cache):
    """Spessore e lunghezza dell'isolante, con la sezione dell'elemento che riveste.

    host resta None quando l'isolante riveste un raccordo o un accessorio: nei fogli
    QTO quei pezzi non si misurano e li copre la maggiorazione percentuale.
    """
    try:
        thickness = (insulation.Thickness or 0.0) * FT_TO_MM
    except Exception:
        thickness = _double_ft(insulation, "RBS_INSULATION_THICKNESS_FOR_DUCT",
                               "RBS_INSULATION_THICKNESS_FOR_PIPE") * FT_TO_MM
    host_geometry = None
    try:
        host_id = insulation.HostElementId
        host_key_value = get_element_id_value(host_id)
        if host_key_value in host_cache:
            host_geometry = host_cache[host_key_value]
        else:
            host = doc.GetElement(host_id)
            host_key = None
            if host is not None and host.Category is not None:
                host_key = key_by_cat_id.get(int(get_element_id_value(host.Category.Id)))
            if host_key in DUCT_HOST_KEYS or host_key in PIPE_HOST_KEYS:
                host_geometry = _curve_geometry(host, host_key)
            host_cache[host_key_value] = host_geometry
    except Exception:
        host_geometry = None

    length = _double_ft(insulation, "CURVE_ELEM_LENGTH") * FT_TO_M
    if not length and host_geometry is not None:
        length = host_geometry.length_m
    return qr.Geometry(length_m=length, thickness_mm=thickness, host=host_geometry)


class _WbsReader(object):
    """Valori WBS di un elemento: istanza, poi tipo, poi il "genitore" (la famiglia che
    contiene un'annidata condivisa, l'elemento rivestito da un isolante) e il suo tipo.
    I valori di tipo si memorizzano per tipo e parametro."""

    def __init__(self, doc, names):
        self.doc = doc
        self.names = names
        self._type_values = {}

    def _type_value(self, element, name):
        try:
            type_id = element.GetTypeId()
        except Exception:
            return u""
        key = (get_element_id_value(type_id), name)
        if key not in self._type_values:
            elem_type = self.doc.GetElement(type_id)
            self._type_values[key] = _param_text(elem_type.LookupParameter(name)) \
                if elem_type is not None else u""
        return self._type_values[key]

    def _value(self, element, name):
        return _param_text(element.LookupParameter(name)) or self._type_value(element, name)

    def values(self, element, parent=None):
        out = []
        for name in self.names:
            value = self._value(element, name)
            if not value and parent is not None:
                value = self._value(parent, name)
            out.append(value)
        return tuple(out)


def _wbs_parent(doc, element, kind):
    try:
        if kind == qr.KIND_COUNT:
            return element.SuperComponent
        if kind in (qr.KIND_DUCT_INSULATION, qr.KIND_PIPE_INSULATION):
            return doc.GetElement(element.HostElementId)
    except Exception:
        pass
    return None


def sample_parameter_names(doc, collect_result, per_category=10):
    """Nomi dei parametri (istanza e tipo) presenti su un campione di elementi per
    categoria: sono le scelte proposte nella finestra WBS."""
    names = set()
    seen_types = set()
    counts = {}

    def add(parameters):
        for param in parameters:
            try:
                names.add(param.Definition.Name)
            except Exception:
                pass

    for record in collect_result.records:
        count = counts.get(record.category_key, 0)
        if count >= per_category:
            continue
        counts[record.category_key] = count + 1
        try:
            element = doc.GetElement(record.element_id)
            if element is None:
                continue
            add(element.Parameters)
            type_id = element.GetTypeId()
            type_key = get_element_id_value(type_id)
            if type_key not in seen_types:
                seen_types.add(type_key)
                elem_type = doc.GetElement(type_id)
                if elem_type is not None:
                    add(elem_type.Parameters)
        except Exception:
            continue
    return sorted((n for n in names if n), key=lambda n: n.lower())


def collect_records(doc, options):
    """Elementi di tutte le categorie computabili, filtrati per fase e opzione.

    Le categorie scelte nella finestra si applicano dopo, sui record: cambiare
    la spunta di una categoria non richiede una nuova lettura del modello.
    """
    result = CollectResult()
    rules = available_rules()
    if not rules:
        return result

    key_by_cat_id = {}
    bics = List[DB.BuiltInCategory]()
    for key, _, _, bic in rules:
        bics.Add(bic)
        key_by_cat_id[int(bic)] = key

    collector = DB.FilteredElementCollector(doc)\
        .WherePasses(DB.ElementMulticategoryFilter(bics))\
        .WhereElementIsNotElementType()
    phase_filter = _phase_filter(options)
    if phase_filter is not None:
        collector = collector.WherePasses(phase_filter)

    param_map = options.param_map
    result.param_map = param_map
    type_cache = {}
    host_cache = {}
    wbs_reader = _WbsReader(doc, options.wbs_names) if options.wbs_names else None
    try:
        for element in collector:
            try:
                category = element.Category
                if category is None:
                    continue
                key = key_by_cat_id.get(int(get_element_id_value(category.Id)))
                if key is None:
                    continue
                kind = category_kind(key)
                if kind == qr.KIND_COUNT and not isinstance(element, DB.FamilyInstance):
                    continue

                if options.primary_only:
                    option = element.DesignOption
                    if option is not None and not option.IsPrimary:
                        result.skipped_options += 1
                        continue

                type_id = element.GetTypeId()
                type_key = get_element_id_value(type_id)
                type_info = type_cache.get(type_key)
                if type_info is None:
                    type_info = _read_type(doc.GetElement(type_id), result, param_map)
                    type_cache[type_key] = type_info

                # Si/No di inclusione: sull'istanza, per tutte le categorie.
                if param_map.include and \
                        flag_is_no(element.LookupParameter(param_map.include)):
                    result.excluded.append((key, type_info.label, param_map.include,
                                            element.Id))
                    continue

                geometry = None
                nested = False
                # Codici d'istanza su tutte le categorie; quelli di tipo solo sulle a pezzo
                # (per le lineari lo stesso tipo cambia voce con la sezione).
                instance_slots = _read_instance_slots(element, result, param_map.linear_codes)
                if kind == qr.KIND_COUNT:
                    if type_info.instance_names and not result.params_found:
                        if _instance_has_code_params(type_info, element):
                            result.params_found = True
                    nested = element.SuperComponent is not None
                    slots = _read_slots(type_info, element) + instance_slots
                else:
                    slots = (u"",) * CODE_SLOTS + instance_slots
                    if kind in (qr.KIND_DUCT_INSULATION, qr.KIND_PIPE_INSULATION):
                        geometry = _insulation_geometry(doc, element, key_by_cat_id, host_cache)
                    else:
                        geometry = _curve_geometry(element, key)

                wbs = ()
                if wbs_reader is not None:
                    wbs = wbs_reader.values(element, _wbs_parent(doc, element, kind))

                result.records.append(InstanceRecord(
                    element.Id, key, type_info.type_mark, type_info.label,
                    slots, nested, geometry, wbs))
            except Exception:
                continue
    finally:
        collector.Dispose()

    return result


# =============================================================================
# AGGREGAZIONE (Python puro)
# =============================================================================

class TypeMarkGroup(object):
    """Tutti gli elementi computati con lo stesso Type Mark."""

    def __init__(self, type_mark):
        self.type_mark = type_mark
        self.category_keys = []
        self.type_labels = []
        self.element_ids = []
        self.nested_count = 0
        # codice -> record degli elementi che lo portano
        self.code_records = OrderedDict()
        # tupla dei codici di tipo -> numero di elementi a pezzo, per riconoscere i
        # conflitti. I codici d'istanza non contano: possono cambiare da elemento a
        # elemento (per le lineari con il diametro o la sezione).
        self.code_sets = OrderedDict()
        self.no_code_ids = []
        # elementi lineari senza codice d'istanza in un Type Mark che altrove ne ha
        self.linear_no_code_ids = []
        self.has_piece = False
        self.has_linear = False

    @property
    def code_ids(self):
        return OrderedDict((code, [r.element_id for r in records])
                           for code, records in self.code_records.items())

    @property
    def count(self):
        return len(self.element_ids)

    @property
    def categories_text(self):
        return u", ".join(category_label(key) for key in self.category_keys)

    @property
    def types_text(self):
        return u"; ".join(self.type_labels)


class Issue(object):
    def __init__(self, kind, subject, detail, element_ids=None):
        self.kind = kind
        self.subject = subject
        self.detail = detail
        self.element_ids = list(element_ids or [])

    @property
    def count(self):
        return len(self.element_ids)


class Takeoff(object):
    """Esito dell'aggregazione per le categorie scelte."""

    def __init__(self):
        self.groups = OrderedDict()
        self.issues = []
        self.instance_count = 0

    def codes(self):
        """Tutti i codici usati, in ordine di prima comparsa."""
        found = OrderedDict()
        for group in self.groups.values():
            for code in group.code_records:
                found[code] = True
        return list(found.keys())


def aggregate(collect_result, selected_keys):
    takeoff = Takeoff()
    selected = set(selected_keys)
    missing_mark_ids = OrderedDict()

    records = [r for r in collect_result.records if r.category_key in selected]
    records.sort(key=lambda r: (r.type_mark.lower(), r.category_key, r.type_label))

    for record in records:
        takeoff.instance_count += 1
        if not record.type_mark:
            missing_mark_ids.setdefault(record.type_label, []).append(record.element_id)
            continue

        group = takeoff.groups.get(record.type_mark)
        if group is None:
            group = TypeMarkGroup(record.type_mark)
            takeoff.groups[record.type_mark] = group

        group.element_ids.append(record.element_id)
        if record.category_key not in group.category_keys:
            group.category_keys.append(record.category_key)
        if record.type_label not in group.type_labels:
            group.type_labels.append(record.type_label)
        if record.nested:
            group.nested_count += 1

        linear = record.kind != qr.KIND_COUNT
        if linear:
            group.has_linear = True
            if not record.codes:
                group.linear_no_code_ids.append(record.element_id)
        else:
            group.has_piece = True
            type_codes = codes_from_slots(record.slots[:INSTANCE_SLOT_OFFSET])
            group.code_sets[type_codes] = group.code_sets.get(type_codes, 0) + 1
        if not record.codes:
            group.no_code_ids.append(record.element_id)
        for code in record.codes:
            group.code_records.setdefault(code, []).append(record)

    param_map = getattr(collect_result, "param_map", None) or ParameterMap.defaults()
    piece_label = param_map.piece_label
    linear_label = param_map.linear_label

    excluded = OrderedDict()
    for key, type_label, param_name, element_id in getattr(collect_result, "excluded", []):
        if key in selected:
            excluded.setdefault((key, type_label, param_name), []).append(element_id)
    for (key, type_label, param_name), ids in excluded.items():
        takeoff.issues.append(Issue(
            ISSUE_EXCLUDED, type_label or u"(unknown type)",
            u"{}: {} = No on the instance, left out of the bill.".format(
                category_label(key), param_name), ids))

    for type_label, ids in missing_mark_ids.items():
        takeoff.issues.append(Issue(
            ISSUE_NO_TYPE_MARK, type_label or u"(unknown type)",
            u"Elements excluded from the takeoff: the type has no Type Mark.", ids))

    for group in takeoff.groups.values():
        if not group.code_records:
            names = u" / ".join(
                ([piece_label + u" (type)"] if group.has_piece else []) +
                [linear_label + u" (instance)"])
            takeoff.issues.append(Issue(
                ISSUE_NO_CODE, group.type_mark,
                u"No {} value on {}.".format(names, group.types_text),
                group.element_ids))
            continue
        if group.linear_no_code_ids:
            takeoff.issues.append(Issue(
                ISSUE_NO_CODE, group.type_mark,
                u"{} of {} elements have no {} value (instance parameter): they are not "
                u"counted.".format(len(group.linear_no_code_ids), group.count,
                                   linear_label),
                group.linear_no_code_ids))
        if len(group.code_sets) > 1:
            variants = [u"[{}] x{}".format(u", ".join(codes) or u"none", count)
                        for codes, count in group.code_sets.items()]
            takeoff.issues.append(Issue(
                ISSUE_INCONSISTENT, group.type_mark,
                u"Elements with this Type Mark carry different type codes: {}. "
                u"Each code is counted only on the elements that carry it.".format(
                    u"; ".join(variants)),
                group.element_ids))

    if takeoff.instance_count and not collect_result.params_found:
        takeoff.issues.insert(0, Issue(
            ISSUE_NO_PARAMS, u"{} / {}".format(piece_label, linear_label),
            u"None of the counted elements has the price code parameters "
            u"({} on the types of the piece categories, {} on the instances of all "
            u"categories): load the shared parameters into the families or the project, "
            u"or map other parameters with Parameters....".format(piece_label, linear_label),
            []))
    return takeoff


# =============================================================================
# COMPUTO (Python puro)
# =============================================================================

class DetailLine(object):
    """Riga Type Mark x codice del computo."""

    def __init__(self, group, code, count, quantity):
        self.group = group
        self.code = code
        self.count = count
        self.quantity = quantity


class Bill(object):
    def __init__(self):
        self.lines = []
        self.quantities = OrderedDict()
        # unita' con cui il codice e' stato misurato (quella della voce, oppure "m" per
        # passerelle e conduit quando la voce non ne ha una)
        self.units = {}
        self.issues = []
        # combinazione WBS -> {(Type Mark, codice): quantita'}; con la WBS spenta la
        # chiave e' (). Il computo ha una riga per Type Mark e codice.
        self.wbs_quantities = OrderedDict()
        # (Type Mark, codice) -> chiavi delle categorie lineari che la voce misura: solo
        # queste voci hanno una maggiorazione e ammettono un override.
        self.line_categories = OrderedDict()
        # (Type Mark, codice) -> maggiorazione dell'override applicato (frazione)
        self.overridden = {}


def _problem_issue(problem, code, unit, record):
    """(tipo anomalia, soggetto, dettaglio) per un problema di misura."""
    label = category_label(record.category_key)
    allowed = qr.ALLOWED_UNITS.get(record.kind) or ()
    if problem == qr.PROBLEM_UNIT:
        return (ISSUE_UNIT, code,
                u"{} are measured in {}, but the code has unit '{}': not counted. "
                u"Set the unit in the Price list.".format(
                    label, u" / ".join(allowed), unit or u"(empty)"))
    if problem == qr.PROBLEM_UNIT_FORCED:
        return (ISSUE_UNIT_FORCED, code,
                u"{} are always measured in m: the code has unit '{}' and was counted "
                u"in m.".format(label, unit))
    if problem == qr.PROBLEM_DENSITY:
        return (ISSUE_DENSITY, record.type_mark,
                u"Pipe Type Mark '{}' is not in the density table of the Rules tab: "
                u"code {} not counted in kg.".format(record.type_mark, code))
    if problem == qr.PROBLEM_HOST:
        return (ISSUE_HOST, code,
                u"{} on fittings or accessories has no measurable section: not counted "
                u"(the allowance for fittings covers it, as in the QTO sheets).".format(label))
    return (ISSUE_GEOMETRY, code,
            u"{} without the dimensions needed for '{}' (length, size or insulation "
            u"thickness): not counted.".format(label, unit or u""))


def wbs_key(values):
    """Combinazione WBS di un elemento fino al suo ultimo livello valorizzato.

    I livelli vuoti in coda si tolgono (le voci stanno sotto l'ultimo livello che
    l'elemento ha); quelli vuoti in mezzo restano e diventano "(not set)". Un elemento
    senza alcun valore finisce in un unico gruppo "(not set)".
    """
    if not values:
        return ()
    trimmed = list(values)
    while trimmed and not trimmed[-1]:
        trimmed.pop()
    return tuple(trimmed) if trimmed else (u"",)


def _wbs_sort_key(key):
    # vuoti in fondo; a parita' di prefisso il gruppo piu' corto viene prima
    return tuple((value == u"", value.lower()) for value in key)


def compute_bill(takeoff, items, rules, overrides=None):
    """Quantita' per codice secondo l'unita' della voce, con le maggiorazioni.

    overrides: {(Type Mark, codice): frazione} che sostituisce la maggiorazione della
    categoria sugli elementi lineari di quella voce. Le categorie a pezzo restano a 1.
    """
    bill = Bill()
    problems = OrderedDict()
    overrides = overrides or {}

    for group in takeoff.groups.values():
        for code, records in group.code_records.items():
            item = items.get(code)
            unit = item.unit if item is not None else u""
            quantity = 0.0
            used_units = set()
            line_key = (group.type_mark, code)
            override = overrides.get(line_key)
            for record in records:
                # La riga esiste anche quando l'elemento non si puo' misurare (unita'
                # mancante, dimensioni assenti): resta a 0 e si vede nel computo.
                per_line = bill.wbs_quantities.setdefault(wbs_key(record.wbs), OrderedDict())
                per_line.setdefault(line_key, 0.0)
                allowance = None
                if record.kind != qr.KIND_COUNT:
                    categories = bill.line_categories.setdefault(line_key, [])
                    if record.category_key not in categories:
                        categories.append(record.category_key)
                    if override is not None:
                        allowance = override
                        bill.overridden[line_key] = override
                value, used, problem = qr.measure(
                    record.kind, record.geometry, unit, record.type_mark,
                    record.category_key, rules, allowance)
                if problem:
                    kind, subject, detail = _problem_issue(problem, code, unit, record)
                    entry = problems.setdefault((kind, subject, record.category_key),
                                                [detail, []])
                    entry[1].append(record.element_id)
                if value is None:
                    continue
                quantity += value
                if used:
                    used_units.add(used)
                per_line[line_key] += value
            bill.lines.append(DetailLine(group, code, len(records), quantity))
            bill.quantities[code] = bill.quantities.get(code, 0.0) + quantity
            if not unit and len(used_units) == 1:
                bill.units[code] = list(used_units)[0]
            else:
                bill.units.setdefault(code, unit)

    for (kind, subject, _), (detail, ids) in problems.items():
        bill.issues.append(Issue(kind, subject, detail, ids))
    return bill


COMBINATION_TOTAL = u"Total"


class OutlineGroup(object):
    """Riga di una combinazione WBS: i valori dei livelli e il totale delle sue voci."""
    kind = "group"

    def __init__(self, key):
        self.key = key
        self.text = COMBINATION_TOTAL


class OutlineItem(object):
    """Voce del computo dentro una combinazione WBS."""
    kind = "item"

    def __init__(self, key, type_mark, code, quantity):
        self.key = key
        self.type_mark = type_mark
        self.code = code
        self.quantity = quantity


def wbs_cells(key, level_count):
    """Valori delle colonne WBS di una riga: "(not set)" per i livelli vuoti in mezzo
    (e per l'elemento senza alcun valore), vuoto per i livelli oltre l'ultimo."""
    cells = [value or WBS_NOT_SET for value in key]
    return cells + [u""] * (level_count - len(cells))


def _line_sort_key(line_key):
    type_mark, code = line_key
    return (type_mark.lower(), type_mark, code.lower(), code)


def bill_outline(bill, labels):
    """Righe del computo in ordine: per ogni combinazione WBS una riga di totale e poi le
    sue voci, una per Type Mark e codice (ordinate per Type Mark, poi codice). Senza
    livelli WBS, solo le voci."""
    if bill is None:
        return []
    if not labels:
        merged = OrderedDict()
        for per_line in bill.wbs_quantities.values():
            for line_key, quantity in per_line.items():
                merged[line_key] = merged.get(line_key, 0.0) + quantity
        return [OutlineItem((), line_key[0], line_key[1], merged[line_key])
                for line_key in sorted(merged, key=_line_sort_key)]
    entries = []
    for key in sorted(bill.wbs_quantities, key=_wbs_sort_key):
        entries.append(OutlineGroup(key))
        per_line = bill.wbs_quantities[key]
        for line_key in sorted(per_line, key=_line_sort_key):
            entries.append(OutlineItem(key, line_key[0], line_key[1], per_line[line_key]))
    return entries


def outline_group_end(entries, index):
    """Indice dell'ultima voce della combinazione che inizia in entries[index]."""
    end = index
    for position in range(index + 1, len(entries)):
        if entries[position].kind == "group":
            break
        end = position
    return end


def model_codes(collect_result):
    """Codici usati nel modello (tutte le categorie raccolte, solo elementi con Type Mark),
    ordinati: sono le righe dell'elenco prezzi."""
    found = set()
    for record in collect_result.records:
        if record.type_mark:
            found.update(record.codes)
    return sorted(found)


class TypeRow(object):
    """Riga del riepilogo Type Mark: categoria, Type Mark, tipo, annidata, 10 codici di
    tipo e 10 d'istanza."""

    __slots__ = ("category_key", "type_mark", "type_label", "nested", "slots", "element_ids")

    def __init__(self, record):
        self.category_key = record.category_key
        self.type_mark = record.type_mark
        self.type_label = record.type_label
        self.nested = record.nested
        self.slots = record.slots
        self.element_ids = []

    @property
    def category(self):
        return category_label(self.category_key)


def type_rows(collect_result, selected_keys):
    """Una riga per categoria / Type Mark / famiglia e tipo / annidata / codici.

    Gli elementi senza Type Mark compaiono in fondo alla loro categoria: sono esclusi
    dal computo ma il riepilogo serve proprio a vederli.
    """
    selected = set(selected_keys)
    rows = OrderedDict()
    for record in collect_result.records:
        if record.category_key not in selected:
            continue
        key = (record.category_key, record.type_mark, record.type_label,
               record.nested, record.slots)
        row = rows.get(key)
        if row is None:
            row = TypeRow(record)
            rows[key] = row
        row.element_ids.append(record.element_id)
    return sorted(rows.values(), key=lambda r: (
        r.category.lower(), r.type_mark == u"", r.type_mark.lower(),
        r.type_label.lower(), r.nested))


def used_slot_count(rows):
    """(coppie di tipo, coppie d'istanza) da mostrare: per ciascun gruppo, fino
    all'ultimo slot valorizzato."""
    last_type = last_instance = 0
    for row in rows:
        for index, code in enumerate(row.slots):
            if not code:
                continue
            if index < INSTANCE_SLOT_OFFSET:
                last_type = max(last_type, index + 1)
            else:
                last_instance = max(last_instance, index - INSTANCE_SLOT_OFFSET + 1)
    return last_type, last_instance


def description_issues(takeoff, merged_items):
    """Codici usati nel modello che non hanno descrizione ne' nel listino ne' nel progetto."""
    issues = []
    for code in takeoff.codes():
        item = merged_items.get(code)
        if item is not None and item.description:
            continue
        ids = []
        marks = []
        for group in takeoff.groups.values():
            if code in group.code_records:
                ids.extend(r.element_id for r in group.code_records[code])
                marks.append(group.type_mark)
        issues.append(Issue(
            ISSUE_NO_DESCRIPTION, code,
            u"Not in the price list and not described in the project file. "
            u"Used by: {}.".format(u", ".join(marks)), ids))
    return issues
