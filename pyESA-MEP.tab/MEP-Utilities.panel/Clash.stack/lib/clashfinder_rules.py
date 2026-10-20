# -*- coding: utf-8 -*-
"""
clashfinder_rules.py - test di clash: modello dati, file JSON, filtri sugli elementi.

Un test confronta due insiemi (Set A, Set B). Ogni insieme e' definito da:
  - sorgente nel modello aperto ("selection", "view", "model", "none") e link da cui
    prendere altri elementi, per nome del file (Document.Title): valgono in ogni progetto
    che carica gli stessi link;
  - categorie (nomi BuiltInCategory, "OST_DuctCurves": indipendenti dalla lingua di Revit);
    nessuna categoria = qualsiasi categoria di modello;
  - filtri di vista del modello aperto (ParameterFilterElement, per nome): l'elemento deve
    passarne almeno uno (sue categorie e sue regole), in AND con le categorie;
  - regole sui parametri (parametro, operatore, valore), combinate con "all" o "any".

I test sono dizionari semplici, cosi' come stanno nel file JSON:

    {"name": "Ducts vs Walls",
     "set_a": {"source": "selection", "links": [], "categories": ["OST_DuctCurves"],
               "view_filters": [], "match": "all",
               "rules": [{"parameter": "System Classification",
                          "operator": "contains", "value": "Supply"}]},
     "set_b": {"source": "none", "links": ["ARC_Model"], "categories": ["OST_Walls"],
               "view_filters": ["Walls - structural"], "match": "all", "rules": []},
     "ignore_connected": true, "ignore_hosted": true}

I file scritti prima di sorgenti, link e filtri si leggono senza conversione: Set A
dalla selezione, Set B dal modello intero, nessun link, nessun filtro.

La libreria personale sta in un file universale di pyRevit (non viene ripulito
all'avvio); import ed export usano lo stesso formato.
"""

import io
import json
import os
import re

from System.Collections.Generic import List

from pyrevit import DB, script

FILE_FORMAT = "ESA_ClashFinder_Tests"
FILE_VERSION = 1
LIBRARY_FILE_ID = "ESA_ClashFinder_tests"

MATCH_ALL = "all"
MATCH_ANY = "any"

# Sorgenti di un insieme nel modello aperto (le etichette sono in clashfinder_engine).
SOURCE_SELECTION = "selection"
SOURCE_VIEW = "view"
SOURCE_MODEL = "model"
SOURCE_NONE = "none"
SOURCES = (SOURCE_SELECTION, SOURCE_VIEW, SOURCE_MODEL, SOURCE_NONE)
DEFAULT_SOURCES = {"set_a": SOURCE_SELECTION, "set_b": SOURCE_MODEL}

# Operatori: chiave salvata nel JSON -> etichetta mostrata.
OPERATORS = [
    ("equals", "equals"),
    ("not_equals", "does not equal"),
    ("contains", "contains"),
    ("not_contains", "does not contain"),
    ("begins", "begins with"),
    ("ends", "ends with"),
    ("greater", "is greater than"),
    ("less", "is less than"),
    ("has_value", "has a value"),
    ("no_value", "has no value"),
]
OPERATOR_LABELS = dict(OPERATORS)
OPERATOR_KEYS = dict((label, key) for key, label in OPERATORS)
VALUELESS_OPERATORS = ("has_value", "no_value")

# Pseudo-parametri, indipendenti dalla lingua di Revit.
FIELD_FAMILY = "[Family]"
FIELD_TYPE = "[Type]"
FIELD_WORKSET = "[Workset]"
PSEUDO_FIELDS = (FIELD_FAMILY, FIELD_TYPE, FIELD_WORKSET)

NUMBER_PATTERN = re.compile(r"[-+]?\d+(?:[.,]\d+)?")

# Libreria iniziale, creata solo se il file non esiste ancora.
STRUCTURE = ["OST_Walls", "OST_Floors", "OST_StructuralFraming",
             "OST_StructuralColumns", "OST_Roofs", "OST_Ceilings"]
DEFAULT_TESTS = [
    {"name": "Ducts vs Structure",
     "set_a": {"categories": ["OST_DuctCurves", "OST_DuctFitting", "OST_DuctAccessory",
                              "OST_FlexDuctCurves"], "match": MATCH_ALL, "rules": []},
     "set_b": {"categories": STRUCTURE, "match": MATCH_ALL, "rules": []},
     "ignore_connected": True, "ignore_hosted": True},
    {"name": "Pipes vs Structure",
     "set_a": {"categories": ["OST_PipeCurves", "OST_PipeFitting", "OST_PipeAccessory",
                              "OST_FlexPipeCurves"], "match": MATCH_ALL, "rules": []},
     "set_b": {"categories": STRUCTURE, "match": MATCH_ALL, "rules": []},
     "ignore_connected": True, "ignore_hosted": True},
    {"name": "Ducts vs Pipes",
     "set_a": {"categories": ["OST_DuctCurves", "OST_DuctFitting"], "match": MATCH_ALL,
               "rules": []},
     "set_b": {"categories": ["OST_PipeCurves", "OST_PipeFitting"], "match": MATCH_ALL,
               "rules": []},
     "ignore_connected": True, "ignore_hosted": True},
]


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


def text(value):
    if value is None:
        return u""
    return u"{}".format(value)


def category_key(category):
    """Nome BuiltInCategory di una categoria ("OST_Walls"), o None."""
    if category is None:
        return None
    try:
        value = int(get_element_id_value(category.Id))
        if value >= 0:
            return None
        import System
        return System.Enum.ToObject(DB.BuiltInCategory, value).ToString()
    except Exception:
        return None


def builtin_category(key):
    try:
        return getattr(DB.BuiltInCategory, key)
    except Exception:
        return None


def category_id_value(key):
    """Valore intero dell'ElementId della categoria, o None se il nome non esiste."""
    bic = builtin_category(key)
    if bic is None:
        return None
    try:
        return int(bic)
    except Exception:
        return None


def model_categories(doc):
    """[(key, nome visualizzato)] delle categorie di modello, ordinate per nome."""
    found = []
    for category in doc.Settings.Categories:
        try:
            if category.CategoryType != DB.CategoryType.Model:
                continue
            key = category_key(category)
            if key is None:
                continue
            found.append((key, category.Name))
        except Exception:
            continue
    found.sort(key=lambda pair: pair[1].lower())
    return found


def category_display_names(doc):
    """{key: nome visualizzato} per tutte le categorie del documento."""
    names = {}
    for category in doc.Settings.Categories:
        key = category_key(category)
        if key:
            names[key] = category.Name
    return names


def number_of(value_text):
    """Primo numero contenuto nel testo ("300 mm" -> 300.0), o None."""
    if not value_text:
        return None
    match = NUMBER_PATTERN.search(value_text)
    if not match:
        return None
    try:
        return float(match.group(0).replace(",", "."))
    except ValueError:
        return None


# =============================================================================
# TEST (dizionari) - normalizzazione e copia
# =============================================================================

def _names(values):
    """Lista di stringhe non vuote, senza doppioni, nell'ordine dato."""
    found = []
    for value in values or []:
        value = text(value).strip()
        if value and value not in found:
            found.append(value)
    return found


def normalize_set(data, default_source=SOURCE_MODEL):
    data = data if isinstance(data, dict) else {}
    categories = _names(data.get("categories"))
    rules = []
    for rule in data.get("rules") or []:
        if not isinstance(rule, dict):
            continue
        parameter = text(rule.get("parameter")).strip()
        operator = text(rule.get("operator")).strip()
        if not parameter or operator not in OPERATOR_LABELS:
            continue
        rules.append({"parameter": parameter, "operator": operator,
                      "value": text(rule.get("value"))})
    match = data.get("match")
    if match not in (MATCH_ALL, MATCH_ANY):
        match = MATCH_ALL
    source = data.get("source")
    if source not in SOURCES:
        source = default_source
    return {"source": source, "links": _names(data.get("links")),
            "categories": categories, "view_filters": _names(data.get("view_filters")),
            "rules": rules, "match": match}


def normalize_test(data):
    """Test valido o None (serve almeno un nome)."""
    if not isinstance(data, dict):
        return None
    name = text(data.get("name")).strip()
    if not name:
        return None
    return {
        "name": name,
        "set_a": normalize_set(data.get("set_a"), DEFAULT_SOURCES["set_a"]),
        "set_b": normalize_set(data.get("set_b"), DEFAULT_SOURCES["set_b"]),
        "ignore_connected": bool(data.get("ignore_connected", True)),
        "ignore_hosted": bool(data.get("ignore_hosted", True)),
    }


def copy_test(test):
    return normalize_test(json.loads(json.dumps(test)))


def new_test(name):
    return normalize_test({"name": name})


def unique_name(name, existing):
    """name, oppure "name (2)", "name (3)"... se gia' usato."""
    taken = set(n.lower() for n in existing)
    if name.lower() not in taken:
        return name
    index = 2
    while u"{} ({})".format(name, index).lower() in taken:
        index += 1
    return u"{} ({})".format(name, index)


def set_is_empty(set_data):
    return (not set_data["categories"] and not set_data["rules"]
            and not set_data["view_filters"])


def describe_set(set_data, display_names):
    """Testo breve: categorie e numero di regole."""
    names = [display_names.get(key, key) for key in set_data["categories"]]
    if not names:
        label = u"Any model category"
    elif len(names) <= 3:
        label = u", ".join(names)
    else:
        label = u"{}, +{} more".format(u", ".join(names[:3]), len(names) - 3)
    extras = []
    if set_data["view_filters"]:
        count = len(set_data["view_filters"])
        extras.append(u"{} filter{}".format(count, u"" if count == 1 else u"s"))
    if set_data["rules"]:
        count = len(set_data["rules"])
        extras.append(u"{} rule{}".format(count, u"" if count == 1 else u"s"))
    if extras:
        label += u" ({})".format(u", ".join(extras))
    return label


SOURCE_SHORT = {SOURCE_SELECTION: u"selection", SOURCE_VIEW: u"active view",
                SOURCE_MODEL: u"model", SOURCE_NONE: u""}


def describe_sources(set_data):
    """Da dove viene l'insieme, breve: "selection + ARC_Model"."""
    parts = []
    if SOURCE_SHORT.get(set_data["source"]):
        parts.append(SOURCE_SHORT[set_data["source"]])
    parts += set_data["links"]
    return u" + ".join(parts) or u"nothing"


def describe_scope(test):
    """Riga mostrata sotto il nome del test: "A: selection | B: ARC_Model"."""
    return u"A: {}  |  B: {}".format(describe_sources(test["set_a"]),
                                     describe_sources(test["set_b"]))


# =============================================================================
# FILE JSON
# =============================================================================

def library_path():
    return script.get_universal_data_file(LIBRARY_FILE_ID, "json")


def _json_number(value):
    """Numeri .NET (System.Int64, System.Double...) che json non riconosce."""
    convert = int if "Int" in type(value).__name__ else float
    try:
        return convert(value)
    except Exception:
        pass
    raise TypeError("{!r} is not JSON serializable".format(value))


def write_json_file(path, data):
    """JSON indentato via file temporaneo.

    ensure_ascii=True in IronPython 2.7 prova a decodificare come UTF-8 stringhe gia'
    unicode e fallisce sulle lettere accentate: si serializza senza escape e si fa
    l'escape \\uXXXX a mano (come mepqto_store.write_json_file). Solleva IOError / OSError.
    """
    payload = json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True,
                         default=_json_number)
    payload = u"".join(c if ord(c) < 128 else u"\\u{:04x}".format(ord(c)) for c in payload)
    temp_path = path + ".tmp"
    with io.open(temp_path, "w", encoding="utf-8") as handle:
        handle.write(unicode(payload))  # noqa: F821 - IronPython 2.7
    if os.path.exists(path):
        os.remove(path)
    os.rename(temp_path, path)


def read_tests_file(path):
    """Lista di test validi letti dal file. Solleva IOError / ValueError."""
    with io.open(path, "r", encoding="utf-8-sig") as handle:
        data = json.loads(handle.read())
    if isinstance(data, list):
        raw = data
    elif isinstance(data, dict) and isinstance(data.get("tests"), list):
        raw = data["tests"]
    elif isinstance(data, dict) and data.get("name"):
        raw = [data]
    else:
        raise ValueError("The file does not contain clash tests.")
    tests = [t for t in (normalize_test(item) for item in raw) if t is not None]
    return tests


def write_tests_file(path, tests):
    write_json_file(path, {"format": FILE_FORMAT, "version": FILE_VERSION,
                           "tests": tests})


def load_library():
    """Test della libreria personale; alla prima apertura quelli di DEFAULT_TESTS."""
    path = library_path()
    if not os.path.exists(path):
        return [normalize_test(t) for t in DEFAULT_TESTS]
    try:
        return read_tests_file(path)
    except Exception:
        return []


def save_library(tests):
    write_tests_file(library_path(), tests)


# =============================================================================
# VALUTAZIONE SUGLI ELEMENTI
# =============================================================================

def view_filters(doc):
    """{nome: ParameterFilterElement} dei filtri di vista del documento."""
    found = {}
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement)
    try:
        for element in collector:
            try:
                found[DB.Element.Name.GetValue(element)] = element
            except Exception:
                continue
    finally:
        collector.Dispose()
    return found


class ViewFilterRule(object):
    """Un filtro di vista applicato a un elemento: sue categorie e sue regole.

    Le regole sono quelle di GetElementFilter() del modello aperto. Su un elemento linkato
    funzionano i parametri di sistema (stesso id in ogni documento); i parametri di
    progetto e condivisi hanno id diversi nel link e non trovano corrispondenza.
    """

    def __init__(self, element):
        self.name = DB.Element.Name.GetValue(element)
        self.category_values = set(int(get_element_id_value(eid))
                                   for eid in element.GetCategories())
        try:
            self.element_filter = element.GetElementFilter()
        except Exception:
            self.element_filter = None

    def passes(self, element, category_value):
        if category_value not in self.category_values:
            return False
        if self.element_filter is None:
            return True
        try:
            return self.element_filter.PassesFilter(element)
        except Exception:
            return False


class CompiledSet(object):
    """Un insieme pronto da applicare agli elementi di un documento.

    host_doc: documento in cui cercare i filtri di vista (il modello aperto). Senza,
    i filtri sono ignorati (campionamento dei parametri).
    """

    def __init__(self, set_data, host_doc=None):
        self.data = set_data
        self.category_values = set()
        self.unknown_categories = []
        for key in set_data["categories"]:
            value = category_id_value(key)
            if value is None:
                self.unknown_categories.append(key)
            else:
                self.category_values.add(value)
        self.rules = set_data["rules"]
        self.match_all = set_data["match"] != MATCH_ANY
        # Categorie indicate ma nessuna esistente: l'insieme non contiene nulla.
        self.impossible = bool(set_data["categories"]) and not self.category_values

        self.filters = []
        self.unknown_filters = []
        if host_doc is not None and set_data["view_filters"]:
            available = view_filters(host_doc)
            for name in set_data["view_filters"]:
                if name in available:
                    self.filters.append(ViewFilterRule(available[name]))
                else:
                    self.unknown_filters.append(name)
            # Filtri indicati ma nessuno presente in questo progetto: nulla passa.
            if not self.filters:
                self.impossible = True

    def category_filter(self):
        """ElementMulticategoryFilter sulle categorie, o None se sono tutte ammesse."""
        if not self.category_values:
            return None
        ids = List[DB.ElementId]()
        for value in self.category_values:
            ids.Add(DB.ElementId(builtin_category_from_value(value)))
        return DB.ElementMulticategoryFilter(ids)

    def matches(self, element, doc):
        if self.impossible or not is_model_element(element):
            return False
        value = int(get_element_id_value(element.Category.Id))
        if self.category_values and value not in self.category_values:
            return False
        if self.filters and not any(f.passes(element, value) for f in self.filters):
            return False
        if not self.rules:
            return True
        results = (rule_matches(element, doc, rule) for rule in self.rules)
        return all(results) if self.match_all else any(results)


def builtin_category_from_value(value):
    import System
    return System.Enum.ToObject(DB.BuiltInCategory, int(value))


# Marchio dei DirectShape creati dal report come copia degli elementi linkati: la
# clash detection li ignora e il report li cancella (anche quelli rimasti orfani).
PREVIEW_APP_ID = "ESA_ClashFinder_Preview"


def is_preview(element):
    try:
        return isinstance(element, DB.DirectShape) and element.ApplicationId == PREVIEW_APP_ID
    except Exception:
        return False


def is_model_element(element):
    """Istanza di una categoria di modello, non vincolata a una vista."""
    try:
        if element is None or isinstance(element, DB.ElementType) or is_preview(element):
            return False
        category = element.Category
        if category is None or category.CategoryType != DB.CategoryType.Model:
            return False
        if get_element_id_value(element.OwnerViewId) > 0:
            return False
        return True
    except Exception:
        return False


def collect(doc, compiled, view_id=None):
    """Elementi del documento (o della vista) che appartengono all'insieme."""
    if compiled.impossible:
        return []
    collector = None
    try:
        if view_id is not None:
            collector = DB.FilteredElementCollector(doc, view_id)
        else:
            collector = DB.FilteredElementCollector(doc)
        collector = collector.WhereElementIsNotElementType()
        category_filter = compiled.category_filter()
        if category_filter is not None:
            collector = collector.WherePasses(category_filter)
        return [e for e in collector if compiled.matches(e, doc)]
    finally:
        if collector is not None:
            collector.Dispose()


def parameter_text(element, doc, name):
    """(trovato, testo) del parametro: prima l'istanza, poi il tipo.

    Testo come lo mostra Revit (AsString o AsValueString, unita' comprese).
    """
    if name in PSEUDO_FIELDS:
        return pseudo_field_text(element, doc, name)
    param = None
    try:
        param = element.LookupParameter(name)
        if param is None:
            type_element = doc.GetElement(element.GetTypeId())
            if type_element is not None:
                param = type_element.LookupParameter(name)
    except Exception:
        param = None
    if param is None:
        return False, None
    try:
        if not param.HasValue:
            return True, None
        if param.StorageType == DB.StorageType.String:
            return True, param.AsString()
        shown = param.AsValueString()
        if shown is not None:
            return True, shown
        if param.StorageType == DB.StorageType.Double:
            return True, text(param.AsDouble())
        if param.StorageType == DB.StorageType.Integer:
            return True, text(param.AsInteger())
        return True, None
    except Exception:
        return True, None


def pseudo_field_text(element, doc, name):
    try:
        if name == FIELD_WORKSET:
            if not doc.IsWorkshared:
                return True, None
            workset = doc.GetWorksetTable().GetWorkset(element.WorksetId)
            return True, workset.Name if workset is not None else None
        type_element = doc.GetElement(element.GetTypeId())
        if type_element is None:
            return True, None
        if name == FIELD_TYPE:
            return True, type_name(type_element)
        family = getattr(type_element, "FamilyName", None)
        return True, family
    except Exception:
        return True, None


def type_name(type_element):
    """Nome del tipo.

    type_element.Name in IronPython solleva eccezione su alcune sottoclassi
    (FamilySymbol, WallType...): si passa dal getter di Element, poi dal parametro.
    """
    try:
        return DB.Element.Name.GetValue(type_element)
    except Exception:
        pass
    for bip in (DB.BuiltInParameter.ALL_MODEL_TYPE_NAME, DB.BuiltInParameter.SYMBOL_NAME_PARAM):
        try:
            param = type_element.get_Parameter(bip)
            if param is not None and param.AsString():
                return param.AsString()
        except Exception:
            continue
    return None


def rule_matches(element, doc, rule):
    found, value = parameter_text(element, doc, rule["parameter"])
    operator = rule["operator"]
    has_value = found and value not in (None, u"")
    if operator == "has_value":
        return has_value
    if operator == "no_value":
        return not has_value
    if not found:
        # Parametro assente: soddisfa solo le negazioni.
        return operator in ("not_equals", "not_contains")
    current = (value or u"").strip().lower()
    wanted = text(rule["value"]).strip().lower()
    if operator == "equals":
        return current == wanted
    if operator == "not_equals":
        return current != wanted
    if operator == "contains":
        return wanted in current
    if operator == "not_contains":
        return wanted not in current
    if operator == "begins":
        return current.startswith(wanted)
    if operator == "ends":
        return current.endswith(wanted)
    left = number_of(current)
    right = number_of(wanted)
    if left is None or right is None:
        return False
    if operator == "greater":
        return left > right
    if operator == "less":
        return left < right
    return False


def parameter_names_for(doc, category_keys, sample_size=10):
    """Nomi dei parametri (istanza e tipo) su un campione di elementi delle categorie.

    Nessuna categoria: campione delle istanze di modello qualsiasi. Il collector e' lazy,
    si interrompe dopo sample_size elementi per categoria.
    """
    names = set()
    groups = [[key] for key in category_keys] if category_keys else [[]]
    for keys in groups:
        compiled = CompiledSet({"categories": keys, "rules": [], "match": MATCH_ALL})
        if compiled.impossible:
            continue
        collector = None
        count = 0
        try:
            collector = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()
            category_filter = compiled.category_filter()
            if category_filter is not None:
                collector = collector.WherePasses(category_filter)
            for element in collector:
                if not is_model_element(element):
                    continue
                _add_param_names(element, names)
                type_element = doc.GetElement(element.GetTypeId())
                if type_element is not None:
                    _add_param_names(type_element, names)
                count += 1
                if count >= sample_size:
                    break
        except Exception:
            pass
        finally:
            if collector is not None:
                collector.Dispose()
    return sorted(names, key=lambda n: n.lower())


def _add_param_names(element, names):
    for param in element.Parameters:
        try:
            names.add(param.Definition.Name)
        except Exception:
            pass
