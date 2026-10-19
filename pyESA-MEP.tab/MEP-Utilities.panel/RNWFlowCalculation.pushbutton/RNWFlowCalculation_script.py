# -*- coding: utf-8 -*-
__title__ = "RNW Flow\nCalculation"
__doc__ = """Version = 4.0
Date    = 16.10.2026
_____________________________________________________________________
Description:
Rainwater network sizing. The rainwater fixtures write their design flow
[l/s] in the Fixture Units of their connector and Revit sums it along the
network, so the Fixture Units value of every pipe is its design flow.
The flow is compared with the tables of the technical manual: Tables
6.3 / 6.4 / 6.5 for sloped horizontal pipes (fill degree chosen in the
dialog), Table 6.2 for vertical stacks (fill degree 0.33). Sloped pipes
are also checked for velocity and fill degree at the design flow.
The results are listed per section (System + Section parameter) in a
window where the user picks the new size from the sizes of the pipe type;
the rows whose new size differs from the model are changed, after
confirmation. A second tab lists the connected fixtures with their
collection area, C1, C2, element code and type, all editable; the tool
can write them and recalculate the network with the new flows. The fittings and accessories connected to the resized
pipes are resized too, as with the Diameter field of the options bar:
connector size, or the family parameter tied to it. Each section is listed as a
horizontal part (horizontal and 45 degree pipes) and a vertical part.
_____________________________________________________________________
Author(s): ESA Engineering
"""
__author__ = "ESA Engineering"

# -----------------------------------------------------------------------------
# Note di sviluppo (in italiano, come il resto dei commenti del repo)
#
# La portata di ogni tubo e' il valore del parametro Fixture Units
# (RBS_PIPE_FIXTURE_UNITS_PARAM), letto come l/s: le famiglie dei bocchettoni
# scrivono la portata nelle unita' di carico del connettore e Revit la somma
# lungo la rete.
#
# Flusso:
#   1. raccolta dei tubi e dei tipi di sistema di tubazioni;
#   2. finestra delle impostazioni (rnwflow_ui): tipo di sistema e sistemi da
#      calcolare (uno o piu'); la serie di tubo non si sceglie, si legge per ogni
#      tubo dal materiale del suo segmento (pipe_series);
#   3. ogni tubo diventa un PipeCheck (dati semplici: portata, geometria, De,
#      DN, Section, misure del tipo di tubazione) verificato da rnwflow_check;
#      le misure vengono dalle preferenze di instradamento del tipo (tutti i
#      segmenti, ciascuno nel suo campo di diametri), vedi pipe_type_sizes;
#   4. raggruppamento per tratto (sistema + Section), terminali (plumbing
#      fixtures) collegati con la loro portata, ordine della rete da monte a valle
#      (network_topology + rnwflow_check.order_network) e finestra dei risultati
#      (rnwflow_results_ui), dove l'utente sceglie le misure nuove (si
#      modificano le righe con misura diversa da quella del modello) e i
#      parametri dei terminali (scheda Fixtures, read_fixture_data);
#   4b. "Apply and Recalculate": write_fixture_edits scrive i parametri dei
#      terminali, Revit aggiorna le portate, analyze rilegge la rete e la
#      finestra si riapre (ResultsState.carry conserva le misure scelte);
#   5. modifica: parametri dei terminali in sospeso, poi tubi, poi raccordi e
#      accessori collegati (resize_fittings);
#      una Transaction per Section (le parti orizzontale e verticale insieme)
#      dentro un TransactionGroup, con un
#      IFailuresPreprocessor dedicato (non il FailureSwallower di pyRevit, vedi
#      la docstring di ResizeFailurePreprocessor). Un tratto che fallisce torna
#      com'era e non annulla gli altri;
#   6. report solo in uscita e solo se una Section o un parametro dei
#      terminali non si e' potuto modificare.
#
# Navigazione: il pulsante di una riga (scheda Sections o Fixtures) chiude la finestra,
# seleziona i tubi, li inquadra e lascia l'utente in una selezione (PickObjects)
# durante la quale puo' navigare la vista attiva; con Fine o Esc la finestra si
# riapre con lo stesso stato. Una finestra non modale richiederebbe ExternalEvent
# e lascerebbe modificare il modello mentre la tabella mostra dati vecchi.
# -----------------------------------------------------------------------------

import math
import os
import re

from System.Collections.Generic import List

from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, forms, DB

import rnwflow_tables as rt
import rnwflow_check as rc
import rnwflow_ui as rui
import rnwflow_results_ui as rres

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

BIP = DB.BuiltInParameter

FT_TO_MM = 304.8

NO_SYSTEM_TYPE = "(No piping system)"

# Parametro che identifica il tratto. Si cerca per nome (vale anche per un
# parametro condiviso "Section"), poi come parametro di sistema RBS_SECTION
# (che in un Revit italiano si chiama "Sezione").
SECTION_PARAM = "Section"

TRANSACTION_GROUP_NAME = "Resize rainwater sections"
FIXTURE_TRANSACTION_NAME = "Edit rainwater fixtures"

# Codice elemento: colonna della scheda Sections e dell'export, modificabile per
# i terminali nella scheda Fixtures (istanza, poi tipo).
ELEMENT_CODE_PARAM = rres.PARAM_CODE

# Categorie dei raccordi e degli accessori da adeguare ai tubi modificati
FITTING_CATEGORY = int(DB.BuiltInCategory.OST_PipeFitting)
ACCESSORY_CATEGORY = int(DB.BuiltInCategory.OST_PipeAccessory)

# Tolleranza [mm] sul raggio fra connettore del raccordo e connettore del tubo
RADIUS_TOLERANCE_MM = 0.5

RESULT_CHANGED = "RESIZED"
RESULT_UNCHANGED = "ALREADY AT NEW DE"
RESULT_FAILED = "FAILED"


# =============================================================================
# UTILITY
# =============================================================================

def get_element_id_value(eid):
    if eid is None:
        return -1
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


def element_name(el):
    """Nome di un elemento.

    In IronPython el.Name fallisce su varie classi (Name dichiarato piu' volte
    nella gerarchia); su PipingSystemType falliscono sia el.Name sia
    Element.Name.GetValue, quindi si ripiega sui parametri del nome del tipo.
    L'ultimo ripiego contiene l'id, cosi' un fallimento resta riconoscibile.
    """
    if el is None:
        return ""
    try:
        name = el.Name
        if name:
            return name
    except Exception:
        pass
    try:
        name = DB.Element.Name.GetValue(el)
        if name:
            return name
    except Exception:
        pass
    for bip in (BIP.SYMBOL_NAME_PARAM, BIP.ALL_MODEL_TYPE_NAME):
        try:
            param = el.get_Parameter(bip)
            if param is not None and param.HasValue and param.AsString():
                return param.AsString()
        except Exception:
            pass
    return "<{0}>".format(get_element_id_value(el.Id))


def fmt(value, digits=2):
    # float(): in IronPython "{:.2f}" su un int solleva ValueError
    if value is None:
        return "-"
    return "{0:.{1}f}".format(float(value), digits)


def fmt_de(value):
    if value is None:
        return "-"
    value = float(value)
    if abs(value - round(value)) < 0.05:
        return "{0:.0f}".format(value)
    return "{0:.1f}".format(value)


def level_name(pipe):
    try:
        if pipe.ReferenceLevel is not None:
            return element_name(pipe.ReferenceLevel)
    except Exception:
        pass
    return "-"


def system_name(pipe):
    param = pipe.get_Parameter(BIP.RBS_SYSTEM_NAME_PARAM)
    if param is not None and param.HasValue:
        return param.AsString() or "-"
    return "-"


def system_type_name(pipe, cache):
    param = pipe.get_Parameter(BIP.RBS_PIPING_SYSTEM_TYPE_PARAM)
    if param is None or not param.HasValue:
        return NO_SYSTEM_TYPE
    type_id = param.AsElementId()
    key = get_element_id_value(type_id)
    if key not in cache:
        system_type = doc.GetElement(type_id) if key > 0 else None
        cache[key] = element_name(system_type) if system_type is not None else NO_SYSTEM_TYPE
    return cache[key]


def pipe_flow(pipe):
    """Portata del tubo [l/s] = Fixture Units; None se il parametro manca."""
    param = pipe.get_Parameter(BIP.RBS_PIPE_FIXTURE_UNITS_PARAM)
    if param is None or not param.HasValue:
        return None
    try:
        return param.AsDouble()
    except Exception:
        return None


def section_value(pipe):
    """Valore del parametro Section come testo; None se assente o vuoto."""
    param = pipe.LookupParameter(SECTION_PARAM)
    if param is None:
        bip = getattr(BIP, "RBS_SECTION", None)
        if bip is not None:
            param = pipe.get_Parameter(bip)
    if param is None or not param.HasValue:
        return None
    storage = param.StorageType
    try:
        if storage == DB.StorageType.Integer:
            value = param.AsInteger()
            return str(value) if value > 0 else None
        if storage == DB.StorageType.String:
            text = (param.AsString() or "").strip()
            return text or None
        text = (param.AsValueString() or "").strip()
        return text or None
    except Exception:
        return None


def outer_diameter_mm(pipe):
    param = pipe.get_Parameter(BIP.RBS_PIPE_OUTER_DIAMETER)
    if param is None or not param.HasValue:
        return None
    return param.AsDouble() * FT_TO_MM


def pipe_geometry(pipe):
    """(angolo sull'orizzontale in gradi, pendenza in cm/m) dalla location curve."""
    try:
        curve = pipe.Location.Curve
        start = curve.GetEndPoint(0)
        end = curve.GetEndPoint(1)
    except Exception:
        return None, None
    dz = abs(end.Z - start.Z)
    dxy = math.sqrt((end.X - start.X) ** 2 + (end.Y - start.Y) ** 2)
    if dxy < 1e-9:
        return 90.0, None
    angle = math.degrees(math.atan2(dz, dxy))
    return angle, dz / dxy * 100.0


def nominal_diameter_mm(pipe):
    param = pipe.get_Parameter(BIP.RBS_PIPE_DIAMETER_PARAM)
    if param is None or not param.HasValue:
        return None
    return param.AsDouble() * FT_TO_MM


def pipe_segment(pipe):
    try:
        segment = pipe.PipeSegment
        if segment is not None:
            return segment
    except Exception:
        pass
    try:
        param = pipe.get_Parameter(BIP.RBS_PIPE_SEGMENT_PARAM)
        if param is not None and param.HasValue:
            return doc.GetElement(param.AsElementId())
    except Exception:
        pass
    return None


def _listed(size):
    """True se la misura compare negli elenchi delle dimensioni di Revit."""
    try:
        return bool(size.UsedInSizeLists)
    except Exception:
        return True


def _add_sizes(sizes, segment, low=None, high=None):
    """Aggiunge a sizes le misure del segmento nel campo [low, high] di diametro
    nominale (unita' interne), senza ripetere un diametro esterno gia' presente."""
    for size in segment.GetSizes():
        if not _listed(size):
            continue
        nominal = size.NominalDiameter
        if low is not None and nominal < low - 1e-9:
            continue
        if high is not None and nominal > high + 1e-9:
            continue
        od_mm = size.OuterDiameter * FT_TO_MM
        if rc.find_size(sizes, od_mm) is None:
            sizes.append(rc.PipeSize(od_mm, nominal * FT_TO_MM, nominal))


def _routing_sizes(pipe_type):
    """Misure delle regole Segments delle preferenze di instradamento, in ordine
    di priorita': una misura vale per il primo segmento che la ammette."""
    manager = pipe_type.RoutingPreferenceManager
    group = DB.RoutingPreferenceRuleGroupType.Segments
    sizes = []
    for index in range(manager.GetNumberOfRules(group)):
        rule = manager.GetRule(group, index)
        segment = doc.GetElement(rule.MEPPartId)
        if segment is None:
            continue
        low = None
        high = None
        for c_index in range(rule.NumberOfCriteria):
            criterion = rule.GetCriterion(c_index)
            if isinstance(criterion, DB.PrimarySizeCriterion):
                low = criterion.MinimumSize
                high = criterion.MaximumSize
        _add_sizes(sizes, segment, low, high)
    return sizes


def pipe_type_sizes(pipe, cache):
    """[PipeSize] che il tipo di tubazione del tubo puo' usare, per De crescente.

    Dalle preferenze di instradamento del tipo; se non sono leggibili, dal
    segmento attuale del tubo. None se nessuna delle due fonti e' leggibile.
    Cache per tipo di tubazione.
    """
    type_id = pipe.GetTypeId()
    key = get_element_id_value(type_id)
    if key in cache:
        return cache[key]
    sizes = None
    try:
        pipe_type = doc.GetElement(type_id)
        if pipe_type is not None:
            sizes = _routing_sizes(pipe_type) or None
    except Exception:
        sizes = None
    if sizes is None:
        segment = pipe_segment(pipe)
        if segment is not None:
            try:
                sizes = []
                _add_sizes(sizes, segment)
                sizes = sizes or None
            except Exception:
                sizes = None
    if sizes is not None:
        sizes.sort(key=lambda size: size.od_mm)
    cache[key] = sizes
    return sizes


# =============================================================================
# RACCOLTA E VERIFICA
# =============================================================================

def collect_pipes():
    collector = DB.FilteredElementCollector(doc)\
        .OfClass(DB.Plumbing.Pipe)\
        .WhereElementIsNotElementType()
    return list(collector)


def selection_systems(type_of_system):
    """Sistemi degli elementi selezionati all'avvio (tubi, raccordi, accessori,
    terminali): nomi letti da RBS_SYSTEM_NAME_PARAM (piu' sistemi separati da
    virgola) e tenuti solo se sono sistemi dei tubi del modello.

    Ritorna (numero di elementi selezionati, [(nome del sistema, tipo)]).
    """
    ids = list(uidoc.Selection.GetElementIds())
    found = []
    for element_id in ids:
        element = doc.GetElement(element_id)
        if element is None:
            continue
        try:
            param = element.get_Parameter(BIP.RBS_SYSTEM_NAME_PARAM)
        except Exception:
            param = None
        if param is None or not param.HasValue:
            continue
        for name in (param.AsString() or "").split(","):
            name = name.strip()
            if name in type_of_system and (name, type_of_system[name]) not in found:
                found.append((name, type_of_system[name]))
    return len(ids), found


def system_type_items(pipes, type_names, system_names):
    """Tipi di sistema con i loro sistemi (istanze), e per ciascuno il numero di
    tubi e di tubi con Fixture Units."""
    counts = {}
    systems = {}
    for pipe in pipes:
        pipe_id = get_element_id_value(pipe.Id)
        type_name = type_names[pipe_id]
        system = system_names[pipe_id]
        flow = pipe_flow(pipe)
        has_flow = 1 if flow is not None and flow > 0 else 0
        total, with_flow = counts.get(type_name, (0, 0))
        counts[type_name] = (total + 1, with_flow + has_flow)
        per_type = systems.setdefault(type_name, {})
        s_total, s_flow = per_type.get(system, (0, 0))
        per_type[system] = (s_total + 1, s_flow + has_flow)
    items = []
    for type_name in sorted(counts.keys(), key=lambda n: n.lower()):
        system_items = [rui.SystemItem(name, values[0], values[1])
                        for name, values in sorted(systems[type_name].items(),
                                                   key=lambda kv: kv[0].lower())]
        items.append(rui.SystemTypeItem(type_name, counts[type_name][0],
                                        counts[type_name][1], system_items))
    return items


def pipe_series(pipe, cache):
    """(serie, descrizione) del tubo, per la colonna della tabella 6.2.

    Si legge dal materiale del segmento del tubo; se non si riconosce, dal nome
    del segmento, poi dal nome del tipo di tubazione. Serie None = non
    riconosciuta. Cache per segmento e tipo.
    """
    segment = pipe_segment(pipe)
    key = (get_element_id_value(segment.Id) if segment is not None else None,
           get_element_id_value(pipe.GetTypeId()))
    if key in cache:
        return cache[key]
    texts = []
    if segment is not None:
        try:
            material = doc.GetElement(segment.MaterialId)
            if material is not None:
                texts.append(("material", element_name(material)))
        except Exception:
            pass
        texts.append(("segment", element_name(segment)))
    try:
        texts.append(("pipe type", element_name(doc.GetElement(pipe.GetTypeId()))))
    except Exception:
        pass
    result = (None, "not recognized ({0})".format(
        ", ".join("{0} '{1}'".format(kind, text) for kind, text in texts) or "no segment"))
    for kind, text in texts:
        series = rt.series_from_text(text)
        if series is not None:
            result = (series, "{0} from {1} '{2}'".format(rt.SERIES_LABELS[series], kind, text))
            break
    cache[key] = result
    return result


def _param_text(param):
    """Valore di un parametro come testo; None se vuoto."""
    if param is None or not param.HasValue:
        return None
    try:
        if param.StorageType == DB.StorageType.String:
            text = param.AsString()
        else:
            text = param.AsValueString()
    except Exception:
        return None
    text = (text or "").strip()
    return text or None


def element_code(element, type_cache=None):
    """e_MEC_ElementCode_1 dell'elemento: parametro di istanza, poi di tipo."""
    param = element.LookupParameter(ELEMENT_CODE_PARAM)
    if param is not None:
        return _param_text(param)
    try:
        type_id = element.GetTypeId()
        key = get_element_id_value(type_id)
    except Exception:
        return None
    if type_cache is not None and key in type_cache:
        return type_cache[key]
    element_type = doc.GetElement(type_id) if key > 0 else None
    value = None
    if element_type is not None:
        value = _param_text(element_type.LookupParameter(ELEMENT_CODE_PARAM))
    if type_cache is not None:
        type_cache[key] = value
    return value


def pipe_type_mark(pipe, cache):
    """Type Mark del tipo di tubazione (ALL_MODEL_TYPE_MARK); None se vuoto. Cache
    per tipo."""
    type_id = pipe.GetTypeId()
    key = get_element_id_value(type_id)
    if key not in cache:
        pipe_type = doc.GetElement(type_id) if key > 0 else None
        param = pipe_type.get_Parameter(BIP.ALL_MODEL_TYPE_MARK) if pipe_type is not None \
            else None
        cache[key] = _param_text(param) if param is not None else None
    return cache[key]


def build_check(pipe, type_name, limits, size_cache, series_cache, code_cache=None,
                mark_cache=None):
    angle, slope = pipe_geometry(pipe)
    series, _description = pipe_series(pipe, series_cache)
    row = rc.PipeCheck(
        pipe_id=get_element_id_value(pipe.Id),
        system=system_name(pipe),
        system_type=type_name,
        level=level_name(pipe),
        section=section_value(pipe),
        flow=pipe_flow(pipe),
        angle=angle,
        slope=slope,
        model_de_mm=outer_diameter_mm(pipe),
        model_dn_mm=nominal_diameter_mm(pipe),
        sizes=pipe_type_sizes(pipe, size_cache),
        z_mid=pipe_mid_z(pipe),
        series=series,
        code=element_code(pipe, code_cache),
        type_mark=pipe_type_mark(pipe, mark_cache if mark_cache is not None else {}),
    )
    return rc.check_pipe(row, limits)


def pipe_mid_z(pipe):
    """Quota del punto medio del tubo [mm]."""
    try:
        curve = pipe.Location.Curve
        return (curve.GetEndPoint(0).Z + curve.GetEndPoint(1).Z) / 2.0 * FT_TO_MM
    except Exception:
        return None


# -----------------------------------------------------------------------------
# Terminali e topologia della rete
# -----------------------------------------------------------------------------

PLUMBING_FIXTURE_CATEGORY = int(DB.BuiltInCategory.OST_PlumbingFixtures)


def fixture_flow(fixture):
    """Portata del terminale [l/s] = Fixture Units dei suoi connettori; None se illeggibili."""
    total = None
    fu_id = DB.ElementId(BIP.RBS_PIPE_FIXTURE_UNITS_PARAM)
    for connector in piping_connectors(fixture):
        try:
            value = connector.GetMEPConnectorInfo().GetConnectorParameterValue(fu_id)
        except Exception:
            continue
        if value is None:
            continue
        try:
            total = (total or 0.0) + float(value.Value)
        except Exception:
            continue
    return total


def network_topology(pipes):
    """Collegamenti della rete fra i tubi dell'elenco, attraverso i raccordi.

    Ritorna (nodi, coppie (id terminale, id tubo), {id terminale: elemento}).
    Un nodo e' l'insieme dei tubi che si incontrano in un raccordo (o in una
    catena di raccordi): due per una curva, tre per una braga. La visita si ferma
    ai tubi fuori dall'elenco.
    """
    ids = set(get_element_id_value(pipe.Id) for pipe in pipes)
    nodes = set()
    fixture_pairs = set()
    fixtures = {}
    for pipe in pipes:
        pipe_id = get_element_id_value(pipe.Id)
        for connector in piping_connectors(pipe):
            stack = list(linked_connectors(connector))
            visited = set()
            node = set([pipe_id])
            while stack:
                ref = stack.pop()
                owner = ref.Owner
                owner_id = get_element_id_value(owner.Id)
                if is_pipe(owner):
                    if owner_id in ids:
                        node.add(owner_id)
                    continue
                if category_value(owner) == PLUMBING_FIXTURE_CATEGORY:
                    fixtures[owner_id] = owner
                    fixture_pairs.add((owner_id, pipe_id))
                    continue
                if not is_fitting(owner) or owner_id in visited:
                    continue
                visited.add(owner_id)
                for other in piping_connectors(owner):
                    try:
                        if other.Id == ref.Id:
                            continue
                    except Exception:
                        pass
                    stack.extend(linked_connectors(other))
            if len(node) > 1:
                nodes.add(frozenset(node))
    return nodes, fixture_pairs, fixtures


def fixture_item(fixture):
    flow = fixture_flow(fixture)
    connectors = piping_connectors(fixture)
    dn = (radius_mm(connectors[0]) or 0) * 2.0 if connectors else None
    try:
        z = fixture.Location.Point.Z * FT_TO_MM
    except Exception:
        z = None
    notes = [] if flow is not None else ["Fixture Units not readable from the connector"]
    return rc.FixtureItem(get_element_id_value(fixture.Id), system_name(fixture),
                          fitting_label(fixture), flow, dn, z, notes, element_code(fixture))


def ordered_entries(sections, pipes_by_id):
    """Tratti e terminali dal punto di monte al punto di valle.

    Ritorna (elenco ordinato, {id: elemento} dei terminali).
    """
    row_of = {}
    pipes = []
    for section in sections:
        for pipe_id in section.pipe_ids:
            row_of[pipe_id] = section.key
            pipes.append(pipes_by_id[pipe_id])
    nodes, fixture_pairs, fixtures = network_topology(pipes)
    by_key = dict((section.key, section) for section in sections)

    def downstream_rank(key):
        section = by_key[key]
        return (section.flow or 0.0, -(section.z if section.z is not None else 0.0))

    # In ogni nodo il tratto di valle e' quello piu' carico (a parita', il piu'
    # basso): gli altri vi confluiscono. I rami di una braga restano fratelli.
    row_edges = set()
    for node in nodes:
        rows_here = set(row_of[pipe_id] for pipe_id in node)
        if len(rows_here) < 2:
            continue
        down = max(rows_here, key=downstream_rank)
        for key in rows_here:
            if key != down:
                row_edges.add((key, down))
    items = [fixture_item(fixture) for fixture in fixtures.values()]
    fixture_edges = [(("fixture", fixture_id), row_of[pipe_id])
                     for fixture_id, pipe_id in fixture_pairs]
    return rc.order_network(sections, items, row_edges, fixture_edges), fixtures


# -----------------------------------------------------------------------------
# Parametri dei terminali (scheda Fixtures)
# -----------------------------------------------------------------------------

def _display_unit(param):
    """Unita' con cui il valore si mostra e si scrive: m2 per un'area, altrimenti
    l'unita' di visualizzazione del parametro; None se non ha unita'."""
    try:
        spec = param.Definition.GetDataType()
        if spec.TypeId.split("-")[0] == DB.SpecTypeId.Area.TypeId.split("-")[0]:
            return DB.UnitTypeId.SquareMeters
    except Exception:
        pass
    try:
        return param.GetUnitTypeId()
    except Exception:
        return None


number_text = rres.number_text


def param_value(param, where):
    """Il parametro ridotto a rres.ParamValue (testo nelle unita' di _display_unit)."""
    editable = not param.IsReadOnly
    storage = param.StorageType
    try:
        if storage == DB.StorageType.String:
            return rres.ParamValue(param.AsString() or "", where, rres.VALUE_TEXT, editable)
        if storage == DB.StorageType.Double:
            text = ""
            if param.HasValue:
                value = param.AsDouble()
                unit = _display_unit(param)
                if unit is not None:
                    try:
                        value = DB.UnitUtils.ConvertFromInternalUnits(value, unit)
                    except Exception:
                        pass
                text = number_text(value)
            return rres.ParamValue(text, where, rres.VALUE_NUMBER, editable)
        if storage == DB.StorageType.Integer:
            text = str(param.AsInteger()) if param.HasValue else ""
            return rres.ParamValue(text, where, rres.VALUE_INTEGER, editable)
    except Exception:
        pass
    return rres.ParamValue(param.AsValueString() or "", where, rres.VALUE_TEXT, False)


def symbol_label(symbol):
    try:
        return "{0} : {1}".format(symbol.Family.Name, element_name(symbol))
    except Exception:
        return element_name(symbol)


def fixture_types(fixtures):
    """{id testo: FamilySymbol} dei tipi di plumbing fixture del documento, piu'
    quelli dei terminali collegati se mancassero."""
    symbols = {}
    collector = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_PlumbingFixtures).WhereElementIsElementType()
    for symbol in collector:
        if isinstance(symbol, DB.FamilySymbol):
            symbols[str(get_element_id_value(symbol.Id))] = symbol
    for fixture in fixtures.values():
        key = str(get_element_id_value(fixture.GetTypeId()))
        if key not in symbols:
            symbol = doc.GetElement(fixture.GetTypeId())
            if symbol is not None:
                symbols[key] = symbol
    return symbols


def read_fixture_data(entries, fixtures, symbols):
    """Dati della scheda Fixtures: terminali in ordine di rete, valori di istanza
    per terminale e valori di tipo per tipo (rres.FixtureData)."""
    names = [name for name, _prefix in rres.FIXTURE_PARAMS]
    types = []
    type_values = {}
    for key, symbol in symbols.items():
        types.append((key, symbol_label(symbol)))
        values = {}
        for name in names:
            param = symbol.LookupParameter(name)
            if param is not None:
                values[name] = param_value(param, rres.WHERE_TYPE)
        type_values[key] = values
    types.sort(key=lambda item: item[1].lower())
    records = []
    for entry in entries:
        if not getattr(entry, "is_fixture", False):
            continue
        fixture = fixtures.get(entry.id)
        if fixture is None:
            continue
        instance = {}
        for name in names:
            param = fixture.LookupParameter(name)
            if param is not None:
                instance[name] = param_value(param, rres.WHERE_INSTANCE)
        type_key = str(get_element_id_value(fixture.GetTypeId()))
        notes = list(entry.notes)
        missing = [name for name in names
                   if name not in instance and name not in type_values.get(type_key, {})]
        if missing:
            notes.append("Parameter(s) not found: " + ", ".join(missing))
        records.append(rres.FixtureRecord(entry.id, entry.system, entry.flow, type_key,
                                          instance, notes))
    return rres.FixtureData(records, types, type_values)


def _write_param(param, text):
    """Scrive il testo nel parametro; ritorna il motivo dell'errore o None."""
    if param is None:
        return "parameter not found"
    if param.IsReadOnly:
        return "parameter is read-only"
    try:
        storage = param.StorageType
        if storage == DB.StorageType.String:
            done = param.Set(text)
        elif storage in (DB.StorageType.Double, DB.StorageType.Integer):
            number = rres.parse_decimal(text)
            if number is None:
                return "'{0}' is not a number".format(text)
            if storage == DB.StorageType.Integer:
                done = param.Set(int(round(number)))
            else:
                unit = _display_unit(param)
                if unit is not None:
                    try:
                        number = DB.UnitUtils.ConvertToInternalUnits(number, unit)
                    except Exception:
                        pass
                done = param.Set(float(number))
        else:
            return "parameter type not supported"
    except Exception as error:
        return _error_text(error)
    return None if done else "Revit refused the value"


def write_fixture_edits(edits, fixtures, symbols):
    """Scrive tipo e parametri dei terminali in una Transaction. Ritorna gli errori.

    Ordine: cambio di tipo, poi parametri di tipo, poi di istanza (un parametro di
    istanza si cerca sul terminale gia' col tipo nuovo).
    """
    errors = []
    by_id = dict((str(key), fixture) for key, fixture in fixtures.items())
    preprocessor = ResizeFailurePreprocessor()
    transaction = DB.Transaction(doc, FIXTURE_TRANSACTION_NAME)
    handling = transaction.GetFailureHandlingOptions()
    handling = handling.SetForcedModalHandling(False)
    handling = handling.SetClearAfterRollback(True)
    handling = handling.SetFailuresPreprocessor(preprocessor)
    transaction.SetFailureHandlingOptions(handling)
    transaction.Start()
    try:
        for fixture_id, type_id in sorted(edits.types.items()):
            fixture = by_id.get(fixture_id)
            symbol = symbols.get(type_id)
            if fixture is None or symbol is None:
                errors.append("Fixture {0}: type not found".format(fixture_id))
                continue
            try:
                if not symbol.IsActive:
                    symbol.Activate()
                fixture.ChangeTypeId(symbol.Id)
            except Exception as error:
                errors.append("Fixture {0}: type not changed to {1}: {2}".format(
                    fixture_id, symbol_label(symbol), _error_text(error)))
        for (type_id, name), text in sorted(edits.type_params.items()):
            symbol = symbols.get(type_id)
            problem = _write_param(symbol.LookupParameter(name) if symbol else None, text)
            if problem:
                errors.append("Type {0}: {1} not written: {2}".format(
                    symbol_label(symbol) if symbol else type_id, name, problem))
        for (fixture_id, name), text in sorted(edits.instance.items()):
            fixture = by_id.get(fixture_id)
            problem = _write_param(fixture.LookupParameter(name) if fixture else None, text)
            if problem:
                errors.append("Fixture {0}: {1} not written: {2}".format(
                    fixture_id, name, problem))
        doc.Regenerate()
        status = transaction.Commit()
        if status == DB.TransactionStatus.RolledBack:
            errors.insert(0, "Revit rolled the fixture changes back: nothing was written")
            errors.extend(preprocessor.messages)
    except Exception as error:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        errors.append("Fixture changes not written: {0}".format(_error_text(error)))
    return errors


# =============================================================================
# MODIFICA DEL MODELLO
# =============================================================================

class ResizeFailurePreprocessor(DB.IFailuresPreprocessor):
    """Registra gli avvisi di Revit senza applicare nessuna risoluzione.

    Non si usa il FailureSwallower di pyRevit: la sua lista RESOLUTION_TYPES
    include UnlockConstraints e DeleteElements, quindi potrebbe sbloccare
    vincoli o cancellare raccordi dell'utente. Qui si eliminano soltanto i
    warning privi di risoluzioni (aprirebbero un popup a ogni tratto) e si
    chiede il rollback in presenza di errori. Stesso schema di MEPAlignToWall.
    """

    def __init__(self):
        self.messages = []

    def PreprocessFailures(self, failuresAccessor):
        has_error = False
        try:
            failures = failuresAccessor.GetFailureMessages()
        except Exception:
            return DB.FailureProcessingResult.Continue
        for failure in failures:
            description = "-"
            try:
                description = failure.GetDescriptionText()
            except Exception:
                pass
            severity = None
            try:
                severity = failure.GetSeverity()
            except Exception:
                pass
            is_error = severity == DB.FailureSeverity.Error
            if is_error:
                has_error = True
            record = ("Error: " if is_error else "Warning: ") + description
            if record not in self.messages:
                self.messages.append(record)
            if severity == DB.FailureSeverity.Warning:
                try:
                    if not failure.HasResolutions():
                        failuresAccessor.DeleteWarning(failure)
                except Exception:
                    pass
        if has_error:
            return DB.FailureProcessingResult.ProceedWithRollBack
        return DB.FailureProcessingResult.Continue


class ChangeResult(object):
    def __init__(self, section, size):
        self.section = section
        self.size = size
        self.fittings = None
        self.result = RESULT_FAILED
        self.changed = 0
        self.messages = []


def resize_pipe(pipe, size):
    """Imposta sul tubo il diametro nominale della misura scelta.

    Ritorna True se il tubo e' cambiato, False se aveva gia' quel De.
    """
    current = outer_diameter_mm(pipe)
    if current is not None and abs(current - size.od_mm) < 0.05:
        return False
    param = pipe.get_Parameter(BIP.RBS_PIPE_DIAMETER_PARAM)
    if param is None or param.IsReadOnly:
        raise ValueError("the Diameter parameter of pipe {0} is read-only".format(
            get_element_id_value(pipe.Id)))
    param.Set(size.nominal)
    return True


def verify_resized(pipes, size):
    """Dopo la rigenerazione ogni tubo deve avere il De scelto: Revit sceglie il
    segmento dalle preferenze di instradamento e potrebbe assegnare un'altra
    misura. In quel caso si annulla il tratto."""
    for pipe in pipes:
        current = outer_diameter_mm(pipe)
        if current is None or abs(current - size.od_mm) > rc.OD_TOLERANCE_MM:
            raise ValueError(
                "Revit gave pipe {0} De {1} instead of {2}: check the routing preferences "
                "of the pipe type".format(get_element_id_value(pipe.Id), fmt_de(current),
                                          fmt_de(size.od_mm)))


# -----------------------------------------------------------------------------
# Raccordi e accessori collegati ai tubi modificati
# -----------------------------------------------------------------------------
#
# I tubi hanno sempre la misura del calcolo: per far tornare un raccordo non si
# cambia mai la misura di un tubo. Come fa Revit quando si cambia a mano la misura
# di un raccordo: il raccordo prende la misura dei tubi modificati e, dove resta
# collegato a un tubo che non cambia, Revit mette una riduzione.
#
# Per ogni raccordo o accessorio collegato a un tubo modificato:
#   1. se i raggi verso gli elementi gia' alla misura giusta coincidono, nulla;
#   2. raggio del connettore impostato direttamente (Connector.Radius), come il
#      campo "Diameter" della barra delle opzioni;
#   2b. connettore staccato e ricollegato: collegando un tubo, Revit adegua il
#      raccordo alla sua misura;
#   3. parametro associato al raggio / diametro del connettore (per esempio NR
#      nelle famiglie e_PP.GG), risalendo le formule semplici; un parametro di
#      tipo si soddisfa cambiando tipo (_try_associated);
#   in 2, 2b e 3, se il raccordo cambiato non combacia piu' con un tubo che
#   resta com'e', fra i due si inserisce una riduzione con NewTransitionFitting
#   dalle preferenze di instradamento (_bridge), come fa Revit; se Revit l'ha gia'
#   inserita da solo, la si riconosce e la si accetta;
#   4. se il raccordo non accetta la misura nuova (per esempio una riduzione
#      bloccata verso un bocchettone), resta com'e' e si inserisce una riduzione
#      fra il raccordo e il tubo modificato (_try_bridge_wrong);
#   5. un raccordo attaccato direttamente a un raccordo appena cambiato entra a
#      sua volta nella coda (adeguamento a catena).
# Nessun raccordo esistente viene cancellato, nessun tubo spostato. Ogni
# tentativo e' in una SubTransaction; se nulla riesce si solleva un errore con la
# spiegazione e la Section viene annullata.
#
# SHIFT + clic (run_fitting_test): si sceglie un raccordo e un DN; lo script prova
# i tentativi 2 e 3 in una transazione che poi annulla, e il report mostra cosa e'
# collegato a ogni connettore prima e dopo ciascun tentativo.

# Parametri del connettore (nella famiglia) e fattore rispetto al raggio
CONNECTOR_SIZE_PARAMS = (("CONNECTOR_RADIUS", 1.0), ("CONNECTOR_DIAMETER", 2.0))

# Numero massimo di raccordi toccati da una modifica (protezione dalla catena)
MAX_FITTINGS_PER_CHANGE = 500


def piping_connectors(element):
    """Connettori fisici del dominio tubazioni di un tubo o di una famiglia."""
    manager = None
    try:
        if isinstance(element, DB.MEPCurve):
            manager = element.ConnectorManager
        else:
            mep = element.MEPModel
            manager = mep.ConnectorManager if mep is not None else None
    except Exception:
        manager = None
    if manager is None:
        return []
    found = []
    for connector in manager.Connectors:
        try:
            if connector.ConnectorType == DB.ConnectorType.Logical:
                continue
            if connector.Domain != DB.Domain.DomainPiping:
                continue
        except Exception:
            continue
        found.append(connector)
    # Il ConnectorManager non garantisce l'ordine: si ordina per Id del connettore,
    # che e' stabile, cosi' "connettore 1" resta lo stesso fra un elenco e l'altro.
    try:
        found.sort(key=lambda connector: connector.Id)
    except Exception:
        pass
    return found


def connector_by_id(element, connector_id):
    for connector in piping_connectors(element):
        try:
            if connector.Id == connector_id:
                return connector
        except Exception:
            continue
    return None


def linked_connectors(connector):
    """Connettori fisicamente collegati (sistemi e connessioni logiche esclusi)."""
    found = []
    try:
        refs = connector.AllRefs
    except Exception:
        return found
    for ref in refs:
        try:
            if ref.ConnectorType == DB.ConnectorType.Logical:
                continue
            owner = ref.Owner
        except Exception:
            continue
        if owner is None or isinstance(owner, DB.MEPSystem):
            continue
        if owner.Id == connector.Owner.Id:
            continue
        found.append(ref)
    return found


def category_value(element):
    try:
        return get_element_id_value(element.Category.Id)
    except Exception:
        return None


def is_fitting(element):
    return category_value(element) in (FITTING_CATEGORY, ACCESSORY_CATEGORY)


def is_pipe(element):
    return isinstance(element, DB.Plumbing.Pipe)


def part_type_name(fitting):
    try:
        return str(fitting.MEPModel.PartType)
    except Exception:
        return ""


def fitting_label(fitting):
    try:
        symbol = fitting.Symbol
        name = "{0} : {1}".format(symbol.Family.Name, element_name(symbol))
    except Exception:
        name = element_name(fitting)
    return "{0} {1}".format(name, get_element_id_value(fitting.Id))


def radius_mm(connector):
    try:
        return connector.Radius * FT_TO_MM
    except Exception:
        return None


def dn_label(connector):
    radius = radius_mm(connector)
    return "DN {0}".format(fmt_de(radius * 2.0)) if radius is not None else "DN ?"


def same_radius(a, b):
    ra = radius_mm(a)
    rb = radius_mm(b)
    return ra is not None and rb is not None and abs(ra - rb) <= RADIUS_TOLERANCE_MM


class FittingLink(object):
    """Un connettore del raccordo e l'elemento (tubo o raccordo) a cui e' collegato."""

    def __init__(self, connector, neighbor, neighbor_connector):
        self.connector = connector
        self.neighbor = neighbor
        self.neighbor_connector = neighbor_connector
        self.neighbor_id = get_element_id_value(neighbor.Id) if neighbor is not None else None
        self.to_pipe = neighbor is not None and is_pipe(neighbor)

    @property
    def matches(self):
        return self.neighbor_connector is None or same_radius(self.connector, self.neighbor_connector)


def fitting_links(fitting):
    links = []
    for connector in piping_connectors(fitting):
        neighbor = None
        neighbor_connector = None
        for ref in linked_connectors(connector):
            owner = ref.Owner
            if is_pipe(owner) or is_fitting(owner):
                neighbor = owner
                neighbor_connector = ref
                if is_pipe(owner):
                    break
        links.append(FittingLink(connector, neighbor, neighbor_connector))
    return links


def connected_fittings(pipes):
    """Raccordi e accessori collegati ai tubi, nell'ordine in cui si incontrano."""
    found = {}
    order = []
    for pipe in pipes:
        for connector in piping_connectors(pipe):
            for ref in linked_connectors(connector):
                owner = ref.Owner
                if not is_fitting(owner):
                    continue
                key = get_element_id_value(owner.Id)
                if key not in found:
                    found[key] = owner
                    order.append(key)
    return [found[key] for key in order]


def _wrong(fitting, authority):
    """Collegamenti verso elementi gia' alla misura giusta, con raggio diverso."""
    return [link for link in fitting_links(fitting)
            if link.neighbor_id in authority and not link.matches]


def _pipe_snapshot(fitting, authority):
    """Tubi non modificati che ora combaciano col raccordo: devono restare cosi'."""
    return set(link.neighbor_id for link in fitting_links(fitting)
               if link.to_pipe and link.neighbor_id not in authority and link.matches)


def _broken(fitting, snapshot):
    return [link for link in fitting_links(fitting)
            if link.to_pipe and link.neighbor_id in snapshot and not link.matches]


def _rollback(sub):
    try:
        if sub.GetStatus() == DB.TransactionStatus.Started:
            sub.RollBack()
    except Exception:
        pass


class FittingReport(object):
    def __init__(self):
        self.checked = 0
        self.resized = 0
        self.bridged = 0        # riduzioni inserite verso tubi che restano com'erano
        self.notes = []
        self.fittings = []      # raccordi toccati e riduzioni nuove, per la verifica finale


# Riduzioni inserite durante il tentativo in corso: (raccordo nuovo, id tubo).
_BRIDGES = []


def _bip(name):
    return getattr(BIP, name, None)


class FamilyConnector(object):
    """Connettore della famiglia: origine (coordinate di famiglia) e parametri
    associati alla sua misura [(nome, fattore, di istanza, formula)]."""

    def __init__(self, origin):
        self.origin = origin
        self.params = []


class FamilyInfo(object):
    """Quanto serve di una famiglia di raccordo: i connettori e, per ogni
    parametro di famiglia, (di istanza, testo della formula o "")."""

    def __init__(self):
        self.connectors = []
        self.params = {}


def scan_fitting_families(fittings):
    """{id famiglia: FamilyInfo}.

    Apre ogni famiglia (senza salvarla) e legge quale parametro di famiglia e'
    associato al raggio o al diametro di ciascun connettore, dove si trova il
    connettore, e per ogni parametro se e' di istanza e la sua formula. Va
    chiamata fuori da ogni transazione. Una famiglia che non si apre vale None.
    """
    families = {}
    for fitting in fittings:
        try:
            family = fitting.Symbol.Family
        except Exception:
            continue
        families[get_element_id_value(family.Id)] = family
    result = {}
    for key, family in families.items():
        family_doc = None
        info = FamilyInfo()
        try:
            family_doc = doc.EditFamily(family)
            manager = family_doc.FamilyManager
            try:
                family_params = list(manager.Parameters)
            except Exception:
                family_params = []
            for family_param in family_params:
                try:
                    info.params[family_param.Definition.Name] = (
                        bool(family_param.IsInstance), family_param.Formula or "")
                except Exception:
                    continue
            collector = DB.FilteredElementCollector(family_doc).OfClass(DB.ConnectorElement)
            for element in collector:
                try:
                    origin = element.Origin
                except Exception:
                    origin = None
                entry = FamilyConnector(origin)
                for bip_name, factor in CONNECTOR_SIZE_PARAMS:
                    bip = _bip(bip_name)
                    param = element.get_Parameter(bip) if bip is not None else None
                    if param is None:
                        continue
                    associated = manager.GetAssociatedFamilyParameter(param)
                    if associated is None:
                        continue
                    entry.params.append((associated.Definition.Name, factor,
                                         bool(associated.IsInstance), associated.Formula or ""))
                info.connectors.append(entry)
            result[key] = info
        except Exception:
            result[key] = None
        finally:
            if family_doc is not None:
                try:
                    family_doc.Close(False)
                except Exception:
                    pass
    return result


# Numero massimo di elementi visitati per trovare le famiglie dei raccordi
MAX_NETWORK_ELEMENTS = 20000


def network_fittings(pipes):
    """Raccordi e accessori della rete collegata ai tubi (attraverso tubi e
    raccordi), per sapere prima delle transazioni quali famiglie leggere: con la
    regola del lato maggiore la catena di adeguamento puo' arrivare lontano."""
    found = {}
    seen = set()
    queue = list(pipes)
    while queue and len(seen) < MAX_NETWORK_ELEMENTS:
        element = queue.pop()
        key = get_element_id_value(element.Id)
        if key in seen:
            continue
        seen.add(key)
        if is_fitting(element):
            found[key] = element
        for connector in piping_connectors(element):
            for ref in linked_connectors(connector):
                owner = ref.Owner
                if (is_pipe(owner) or is_fitting(owner)) and                         get_element_id_value(owner.Id) not in seen:
                    queue.append(owner)
    return list(found.values())


def _associated_params(fitting, connector, family_map):
    """[(parametro di istanza, fattore, descrizione)] associati alla misura del connettore."""
    found = []
    try:
        info = connector.GetMEPConnectorInfo()
    except Exception:
        info = None
    if info is not None:
        for bip_name, factor in CONNECTOR_SIZE_PARAMS:
            bip = _bip(bip_name)
            if bip is None:
                continue
            try:
                family_id = info.GetAssociateFamilyParameterId(DB.ElementId(bip))
            except Exception:
                continue
            key = get_element_id_value(family_id)
            if key is None or key < 0:
                continue
            for param in fitting.Parameters:
                try:
                    if get_element_id_value(param.Id) == key:
                        found.append((param, factor))
                except Exception:
                    continue
    if not found and family_map:
        for name, factor, _is_instance, _formula in _family_params_of(fitting, connector,
                                                                       family_map):
            param = fitting.LookupParameter(name)
            if param is not None:
                found.append((param, factor))
    unique = []
    names = set()
    for param, factor in found:
        name = _param_name(param)
        if name not in names:
            names.add(name)
            unique.append((param, factor))
    return unique


def _family_info(fitting, family_map):
    try:
        return family_map.get(get_element_id_value(fitting.Symbol.Family.Id)) \
            if family_map else None
    except Exception:
        return None


def _family_entries(fitting, family_map):
    info = _family_info(fitting, family_map)
    return info.connectors if info is not None else None


# Distanza massima [mm] fra connettore dell'istanza e connettore della famiglia
CONNECTOR_MATCH_MM = 5.0


def _family_params_of(fitting, connector, family_map):
    """Parametri di famiglia associati a QUESTO connettore dell'istanza.

    Il connettore dell'istanza si riporta nelle coordinate della famiglia con la
    trasformazione del raccordo e si abbina al connettore di famiglia piu'
    vicino. Se l'abbinamento non riesce si usano i parametri solo quando tutti i
    connettori della famiglia hanno lo stesso (per esempio NR su una curva).
    """
    entries = _family_entries(fitting, family_map) or []
    if not entries:
        return []
    try:
        local = fitting.GetTransform().Inverse.OfPoint(connector.Origin)
    except Exception:
        local = None
    if local is not None:
        best = None
        best_distance = None
        for entry in entries:
            if entry.origin is None:
                continue
            try:
                distance = entry.origin.DistanceTo(local) * FT_TO_MM
            except Exception:
                continue
            if best_distance is None or distance < best_distance:
                best = entry
                best_distance = distance
        if best is not None and best_distance <= CONNECTOR_MATCH_MM:
            return best.params
    names = set(tuple(entry.params) for entry in entries)
    if len(names) == 1:
        return entries[0].params
    return []


def _param_name(param):
    try:
        return param.Definition.Name
    except Exception:
        return "?"


def _family_entries_text(fitting, family_map):
    entries = _family_entries(fitting, family_map)
    if entries is None:
        return "family not read"
    params = []
    for entry in entries:
        for item in entry.params:
            if item not in params:
                params.append(item)
    if not params:
        return "no family parameter is associated with the connector size"
    texts = []
    for name, factor, is_instance, formula in params:
        text = "{0} ({1}{2}{3})".format(
            name, "radius" if factor == 1.0 else "diameter",
            "" if is_instance else ", type parameter",
            ", formula '{0}'".format(formula) if formula else "")
        texts.append(text)
    return ", ".join(texts)


def _check(fitting, authority, snapshot):
    """(esito, motivo): esito True se il raccordo ora combacia con i tubi modificati.

    Se il raccordo cambiato non combacia piu' con un tubo che resta com'e', si
    inserisce fra i due una riduzione dalle preferenze di instradamento, come fa
    Revit (_bridge); le riduzioni inserite finiscono in _BRIDGES.
    """
    broken = _broken(fitting, snapshot)
    if broken:
        ids = ", ".join(sorted(set(str(l.neighbor_id) for l in broken)))
        try:
            _BRIDGES.extend(_bridge(fitting, broken))
        except Exception as error:
            return False, ("it changes the connector joined to pipe {0}, which keeps its size, "
                           "and Revit could not add a transition there ({1})".format(
                               ids, _error_text(error)))
        if _broken(fitting, snapshot):
            return False, "pipe {0} still does not match the fitting after adding a " \
                          "transition".format(ids)
    wrong = _wrong(fitting, authority)
    if wrong:
        return False, "the connectors did not reach the size ({0})".format(
            "; ".join("{0} to {1} {2}".format(dn_label(l.connector),
                                               "pipe" if l.to_pipe else "fitting", l.neighbor_id)
                      for l in wrong))
    return True, ""


def _error_text(error):
    """Tipo e messaggio completo di un errore di Revit, su una riga."""
    try:
        name = error.clsException.GetType().Name
    except Exception:
        name = type(error).__name__
    text = " ".join(str(error).split())
    return "{0}: {1}".format(name, text)[:240]


def _try_reconnect(fitting, authority, snapshot, log):
    """Strategia 2b: stacca e ricollega i connettori fuori misura, cosi' Revit
    adegua il raccordo all'elemento collegato."""
    wrong = _wrong(fitting, authority)
    if not wrong:
        return True
    sub = DB.SubTransaction(doc)
    sub.Start()
    try:
        for link in wrong:
            link.connector.DisconnectFrom(link.neighbor_connector)
        doc.Regenerate()
        for link in wrong:
            link.neighbor_connector.ConnectTo(link.connector)
        doc.Regenerate()
        ok, reason = _check(fitting, authority, snapshot)
        if ok:
            sub.Commit()
            return True
        log.append("reconnecting: " + reason)
    except Exception as error:
        log.append("reconnecting refused by Revit ({0})".format(_error_text(error)))
    _rollback(sub)
    return False


def _try_connector_radius(fitting, authority, snapshot, log):
    """Strategia 2: raggio del connettore, come la barra delle opzioni."""
    sub = DB.SubTransaction(doc)
    sub.Start()
    try:
        for link in _wrong(fitting, authority):
            link.connector.Radius = link.neighbor_connector.Radius
        doc.Regenerate()
        ok, reason = _check(fitting, authority, snapshot)
        if ok:
            sub.Commit()
            return True
        log.append("connector size: " + reason)
    except Exception as error:
        log.append("connector size refused by Revit ({0})".format(_error_text(error)))
    _rollback(sub)
    return False


_NUMBER = r"([0-9]+(?:[.,][0-9]+)?)"


def _parse_linear(formula, params):
    """(nome, scala) se la formula e' 'X', 'X / k', 'X * k' o 'k * X', con X un
    parametro della famiglia: il parametro vale scala x X. None altrimenti."""
    text = formula.strip()
    while text.startswith("(") and text.endswith(")"):
        text = text[1:-1].strip()

    def name_of(raw):
        raw = raw.strip().strip("[]").strip()
        return raw if raw in params else None

    if name_of(text):
        return name_of(text), 1.0
    match = re.match(r"^(.+?)\s*([*/])\s*" + _NUMBER + r"$", text)
    if match and name_of(match.group(1)):
        number = float(match.group(3).replace(",", "."))
        if number == 0:
            return None
        scale = 1.0 / number if match.group(2) == "/" else number
        return name_of(match.group(1)), scale
    match = re.match(r"^" + _NUMBER + r"\s*\*\s*(.+)$", text)
    if match and name_of(match.group(2)):
        return name_of(match.group(2)), float(match.group(1).replace(",", "."))
    return None


def _follow(info, name, value, depth=0):
    """Risale le formule dal parametro associato al connettore fino al parametro
    che le guida. Ritorna (tipo, nome, valore) con tipo "instance", "type" o
    "blocked" (nome = motivo)."""
    is_instance, formula = info.params.get(name, (True, ""))
    if not formula:
        return ("instance" if is_instance else "type", name, value)
    if depth >= 4:
        return ("blocked", "{0}: formula chain too long".format(name), None)
    parsed = _parse_linear(formula, info.params)
    if parsed is None:
        return ("blocked", "{0} has the formula '{1}', which the tool cannot invert".format(
            name, formula), None)
    driver, scale = parsed
    return _follow(info, driver, value / scale, depth + 1)


def _drivers(fitting, link, family_map):
    """Parametri da impostare per portare il connettore del link alla misura
    dell'elemento collegato: [(tipo, nome, valore)] come _follow."""
    target = link.neighbor_connector.Radius
    info = _family_info(fitting, family_map)
    if info is not None:
        entries = _family_params_of(fitting, link.connector, family_map)
        if entries:
            return [_follow(info, name, target * factor) for name, factor, _i, _f in entries]
    found = []
    for param, factor in _associated_params(fitting, link.connector, None):
        name = _param_name(param)
        if param.IsReadOnly:
            found.append(("blocked", "{0} is locked on the instance (tied to the connector, "
                                     "or driven by a formula)".format(name), None))
        else:
            found.append(("instance", name, target * factor))
    return found


def _type_candidates(fitting, requirements):
    """Tipi della stessa famiglia con i parametri di tipo ai valori richiesti."""
    try:
        symbol = fitting.Symbol
        ids = list(symbol.Family.GetFamilySymbolIds())
    except Exception:
        return []
    tolerance = RADIUS_TOLERANCE_MM / FT_TO_MM
    found = []
    for symbol_id in ids:
        if get_element_id_value(symbol_id) == get_element_id_value(symbol.Id):
            continue
        candidate = doc.GetElement(symbol_id)
        if candidate is None:
            continue
        ok = True
        for name, value in requirements:
            param = candidate.LookupParameter(name)
            try:
                if param is None or abs(param.AsDouble() - value) > tolerance:
                    ok = False
                    break
            except Exception:
                ok = False
                break
        if ok:
            found.append(candidate)
    return found


def _try_associated(fitting, authority, snapshot, family_map, log, used):
    """Strategia 3: parametri associati alla misura dei connettori.

    Dal parametro associato al connettore si risale l'eventuale formula fino al
    parametro che guida (per esempio NR2 = ND2 / 2 -> ND2). Un parametro di
    istanza si imposta; un parametro di tipo richiede un altro tipo della stessa
    famiglia con quel valore, che viene assegnato al raccordo.
    """
    plan = {}          # nome parametro di istanza -> valore
    type_reqs = {}     # nome parametro di tipo -> valore
    for link in _wrong(fitting, authority):
        for kind, name, value in _drivers(fitting, link, family_map):
            if kind == "blocked":
                log.append(name)
                return False
            target = plan if kind == "instance" else type_reqs
            if name in target and abs(target[name] - value) * FT_TO_MM > RADIUS_TOLERANCE_MM:
                log.append("{0} would need two different values: the connectors it drives "
                           "need different sizes".format(name))
                return False
            target[name] = value
    if not plan and not type_reqs:
        log.append("associated parameter: none found ({0})".format(
            _family_entries_text(fitting, family_map)))
        return False

    params = []
    for name, value in plan.items():
        param = fitting.LookupParameter(name)
        if param is None or param.IsReadOnly:
            log.append("{0} is locked on the instance (tied to the connector, or driven by a "
                       "formula)".format(name))
            return False
        params.append((param, value))
    text = ", ".join("{0} = {1} mm".format(name, fmt_de(value * FT_TO_MM))
                     for name, value in sorted(list(plan.items()) + list(type_reqs.items())))

    candidates = [None]
    if type_reqs:
        candidates = _type_candidates(fitting, list(type_reqs.items()))
        if not candidates:
            log.append("{0} is a type parameter and no other type of the family has {1}".format(
                ", ".join(sorted(type_reqs)), text))
            return False

    reasons = []
    for candidate in candidates:
        sub = DB.SubTransaction(doc)
        sub.Start()
        try:
            if candidate is not None:
                try:
                    if not candidate.IsActive:
                        candidate.Activate()
                except Exception:
                    pass
                fitting.ChangeTypeId(candidate.Id)
            for param, value in params:
                param.Set(value)
            doc.Regenerate()
            ok, reason = _check(fitting, authority, snapshot)
            if ok:
                sub.Commit()
                if candidate is not None:
                    used.append("type changed to {0}".format(element_name(candidate)))
                if params:
                    used.append(", ".join("{0} = {1} mm".format(
                        _param_name(param), fmt_de(value * FT_TO_MM)) for param, value in params))
                return True
            reasons.append(("type {0}: ".format(element_name(candidate)) if candidate else "")
                           + reason)
        except Exception as error:
            reasons.append("{0}refused by Revit ({1})".format(
                "type {0} ".format(element_name(candidate)) if candidate else "",
                _error_text(error)))
        _rollback(sub)
    log.append("{0}: {1}".format(text, "; ".join(reasons)))
    return False


def connector_inventory(fitting):
    """Ogni connettore del raccordo: misura e a cosa e' collegato, per i messaggi."""
    texts = []
    for connector in piping_connectors(fitting):
        size = dn_label(connector)
        refs = linked_connectors(connector)
        if not refs:
            texts.append("{0} free (not connected)".format(size))
            continue
        owner = refs[0].Owner
        if is_pipe(owner):
            what = "pipe {0} ({1})".format(get_element_id_value(owner.Id), dn_label(refs[0]))
        elif is_fitting(owner):
            what = "{0} ({1})".format(fitting_label(owner), dn_label(refs[0]))
        else:
            try:
                category = owner.Category.Name
            except Exception:
                category = "element"
            what = "{0} {1} (not a pipe or a fitting)".format(category,
                                                              get_element_id_value(owner.Id))
        texts.append("{0} to {1}".format(size, what))
    return texts


def _connector_near(element, point):
    """Connettore di tubazione dell'elemento piu' vicino al punto."""
    best = None
    best_distance = None
    for connector in piping_connectors(element):
        try:
            distance = connector.Origin.DistanceTo(point)
        except Exception:
            continue
        if best_distance is None or distance < best_distance:
            best = connector
            best_distance = distance
    return best


def _bridge(fitting, links):
    """Inserisce una riduzione fra il tubo di ogni link e il connettore del
    raccordo (misure diverse), con NewTransitionFitting e la famiglia delle
    preferenze di instradamento, come fa Revit.

    NewTransitionFitting vuole come primo connettore quello di un tubo ("The owner
    should be (flex) duct or pipe"): il tubo va sempre per primo. Se Revit non la
    accetta con i connettori ancora collegati, li si stacca e si riprova. Ritorna
    [(riduzione, id tubo)]; solleva l'errore se non riesce.
    """
    created = []
    for link in links:
        if not link.to_pipe:
            raise ValueError("{0} meets {1} directly: a transition can only be added toward "
                             "a pipe".format(fitting_label(fitting), fitting_label(link.neighbor)))
        pipe = link.neighbor
        pipe_connector = link.neighbor_connector
        fitting_connector = link.connector
        connector_id = fitting_connector.Id
        point = pipe_connector.Origin
        try:
            transition = doc.Create.NewTransitionFitting(pipe_connector, fitting_connector)
        except Exception:
            try:
                fitting_connector.DisconnectFrom(pipe_connector)
            except Exception:
                pass
            doc.Regenerate()
            # Dopo la rigenerazione si riprendono i connettori: quello del raccordo
            # per Id, quello del tubo per posizione.
            fitting_connector = connector_by_id(fitting, connector_id) or fitting_connector
            pipe_connector = _connector_near(pipe, point) or pipe_connector
            transition = doc.Create.NewTransitionFitting(pipe_connector, fitting_connector)
        doc.Regenerate()
        created.append((transition, link.neighbor_id))
    return created


def _try_bridge_wrong(fitting, authority, log):
    """Strategia 4: il raccordo resta com'e' e si inserisce una riduzione fra il
    raccordo e ogni elemento modificato che non combacia."""
    wrong = _wrong(fitting, authority)
    if not wrong:
        return True
    sub = DB.SubTransaction(doc)
    sub.Start()
    try:
        _BRIDGES.extend(_bridge(fitting, wrong))
        doc.Regenerate()
        still = _wrong(fitting, authority)
        if not still:
            sub.Commit()
            return True
        log.append("transition between the fitting and the resized pipe: the connectors still "
                   "do not match")
    except Exception as error:
        log.append("transition between the fitting and the resized pipe refused by Revit "
                   "({0})".format(_error_text(error)))
    _rollback(sub)
    return False


def _new_neighbors(fitting, before_ids):
    """Raccordi ora collegati al raccordo che prima non lo erano (riduzioni
    inserite da Revit o dallo script)."""
    found = []
    for link in fitting_links(fitting):
        if link.neighbor is not None and not link.to_pipe and link.neighbor_id not in before_ids:
            found.append(link.neighbor)
    return found


def _resize_fitting(fitting, authority, family_map, report):
    """Porta il raccordo alla misura degli elementi in authority, o solleva
    ValueError con la spiegazione. Ritorna le riduzioni inserite (da verificare)."""
    label = fitting_label(fitting)
    del _BRIDGES[:]
    before_ids = set(link.neighbor_id for link in fitting_links(fitting) if link.neighbor is not None)
    snapshot = _pipe_snapshot(fitting, authority)
    wrong = _wrong(fitting, authority)
    targets = sorted(set(round(link.neighbor_connector.Radius * 2.0 * FT_TO_MM, 1)
                         for link in wrong))
    target_text = ", ".join("DN " + fmt_de(t) for t in targets)
    inventory = connector_inventory(fitting)
    log = []
    how = None
    if _try_connector_radius(fitting, authority, snapshot, log):
        how = "connector size = {0}".format(target_text)
    elif _try_reconnect(fitting, authority, snapshot, log):
        how = "resized by reconnecting it ({0})".format(target_text)
    else:
        used = []
        if _try_associated(fitting, authority, snapshot, family_map, log, used):
            how = ", ".join(used)
        elif _try_bridge_wrong(fitting, authority, log):
            how = "kept as it is (it does not take {0}); a transition was added toward the " \
                  "resized pipe".format(target_text)
    if how is None:
        message = "{0} could not be resized to {1}. Tried: {2}. Connectors: {3}.".format(
            label, target_text, "; ".join(log), "; ".join(inventory))
        raise ValueError(message)
    report.resized += 1
    report.notes.append("{0}: {1}".format(label, how))
    added = _new_neighbors(fitting, before_ids)
    for transition, element_id in _BRIDGES:
        report.notes.append("{0} added between {1} and element {2}".format(
            fitting_label(transition), label, element_id))
    added_ids = set(get_element_id_value(f.Id) for f in added)
    script_ids = set(get_element_id_value(t.Id) for t, _e in _BRIDGES)
    for other in added:
        if get_element_id_value(other.Id) not in script_ids:
            report.notes.append("Revit added {0} next to {1}".format(fitting_label(other), label))
    report.bridged += len(added_ids | script_ids)
    return added + [t for t, _e in _BRIDGES if get_element_id_value(t.Id) not in added_ids]


def resize_fittings(pipes, family_map=None):
    """Adegua raccordi e accessori collegati ai tubi modificati, a catena.

    authority = elementi gia' alla misura giusta: i tubi modificati e i raccordi
    gia' sistemati. Ritorna un FittingReport; solleva ValueError (Section da
    annullare) se un raccordo non si riesce ad adeguare. Le misure dei tubi non
    cambiano mai.
    """
    report = FittingReport()
    authority = set(get_element_id_value(pipe.Id) for pipe in pipes)
    queue = connected_fittings(pipes)
    seen = set()
    while queue:
        fitting = queue.pop(0)
        fitting_id = get_element_id_value(fitting.Id)
        if fitting_id in seen:
            continue
        seen.add(fitting_id)
        report.checked += 1
        if report.checked > MAX_FITTINGS_PER_CHANGE:
            raise ValueError("more than {0} fittings to resize: stopped".format(
                MAX_FITTINGS_PER_CHANGE))
        if _wrong(fitting, authority):
            for added in _resize_fitting(fitting, authority, family_map, report):
                added_id = get_element_id_value(added.Id)
                seen.add(added_id)
                authority.add(added_id)
                report.fittings.append(added)
        authority.add(fitting_id)
        report.fittings.append(fitting)
        # Raccordi collegati direttamente che ora non combaciano: in coda.
        for link in fitting_links(fitting):
            if link.neighbor is not None and not link.to_pipe \
                    and link.neighbor_id not in seen and not link.matches:
                queue.append(link.neighbor)
    return report


def verify_fittings(report):
    """Ogni raccordo toccato deve combaciare con tutti gli elementi collegati,
    anche quelli non modificati."""
    for fitting in report.fittings:
        for link in fitting_links(fitting):
            if link.neighbor is None or link.matches:
                continue
            if link.to_pipe:
                raise ValueError(
                    "{0} now joins pipe {1}, which has a different size: resize that part "
                    "too or add a transition".format(fitting_label(fitting), link.neighbor_id))
            raise ValueError("{0} does not match the size of {1}".format(
                fitting_label(fitting), fitting_label(link.neighbor)))


def _change_groups(changes):
    """Modifiche riunite per Section (sistema + Section): le parti orizzontale e
    verticale della stessa Section si modificano nella stessa transazione, cosi'
    la curva fra le due vede subito entrambe le misure nuove."""
    groups = {}
    order = []
    for section, size in changes:
        key = section.group_key
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append((section, size))
    return [groups[key] for key in order]


def apply_changes(changes, pipes_by_id):
    """Una Transaction per Section dentro un TransactionGroup. Ritorna [ChangeResult]."""
    results = []
    # Le famiglie dei raccordi si leggono prima: EditFamily non si puo' chiamare
    # con una transazione aperta.
    changed_pipes = [pipes_by_id[pipe_id] for section, _size in changes
                     for pipe_id in section.pipe_ids]
    family_map = scan_fitting_families(network_fittings(changed_pipes))
    group = DB.TransactionGroup(doc, TRANSACTION_GROUP_NAME)
    group.Start()
    try:
        for items in _change_groups(changes):
            outcomes = [ChangeResult(section, size) for section, size in items]
            results.extend(outcomes)
            first = items[0][0]
            preprocessor = ResizeFailurePreprocessor()
            transaction = DB.Transaction(
                doc, "Resize section {0} {1}".format(first.system, first.section or ""))
            handling = transaction.GetFailureHandlingOptions()
            handling = handling.SetForcedModalHandling(False)
            handling = handling.SetClearAfterRollback(True)
            handling = handling.SetFailuresPreprocessor(preprocessor)
            transaction.SetFailureHandlingOptions(handling)
            transaction.Start()
            try:
                counts = []
                all_pipes = []
                for outcome in outcomes:
                    pipes = [pipes_by_id[pipe_id] for pipe_id in outcome.section.pipe_ids]
                    counts.append(sum(1 for pipe in pipes if resize_pipe(pipe, outcome.size)))
                    all_pipes.extend(pipes)
                fittings = None
                if sum(counts):
                    doc.Regenerate()
                    for outcome in outcomes:
                        verify_resized([pipes_by_id[i] for i in outcome.section.pipe_ids],
                                       outcome.size)
                    fittings = resize_fittings(all_pipes, family_map)
                    doc.Regenerate()
                    for outcome in outcomes:
                        verify_resized([pipes_by_id[i] for i in outcome.section.pipe_ids],
                                       outcome.size)
                    verify_fittings(fittings)
                status = transaction.Commit()
                for index, outcome in enumerate(outcomes):
                    outcome.messages.extend(preprocessor.messages)
                    if status == DB.TransactionStatus.RolledBack:
                        outcome.result = RESULT_FAILED
                        outcome.messages.insert(0, "Revit rolled the section back")
                        continue
                    outcome.changed = counts[index]
                    outcome.result = RESULT_CHANGED if counts[index] else RESULT_UNCHANGED
                    if index == 0:
                        outcome.fittings = fittings
                        if fittings is not None:
                            outcome.messages.extend(fittings.notes)
                    else:
                        outcome.messages.append("Fittings handled with the other part of "
                                                "the section")
            except Exception as error:
                if transaction.GetStatus() == DB.TransactionStatus.Started:
                    transaction.RollBack()
                for outcome in outcomes:
                    outcome.result = RESULT_FAILED
                    outcome.messages.append(str(error))
                    outcome.messages.extend(preprocessor.messages)
        group.Assimilate()
    except Exception:
        if group.GetStatus() == DB.TransactionStatus.Started:
            group.RollBack()
        raise
    return results


# =============================================================================
# REPORT
# =============================================================================

def scope_label(config):
    if config.from_selection:
        return "whole systems of the elements selected at start"
    return "whole systems chosen in the dialog"


def speed_label(config):
    parts = []
    if config.vmin is not None:
        parts.append("min {0} m/s".format(config.vmin))
    if config.vmax is not None:
        parts.append("max {0} m/s".format(config.vmax))
    return ", ".join(parts) if parts else "not checked"


def settings_text(config):
    return ("System type: {0}   |   Systems: {9}   |   Horizontal pipes: f = {2} ({3})   |   "
            "Velocity: {8}   |   Vertical: f = {4} ({5}, steeper than {6:.0f} deg)   |   "
            "Scope: {7}   |   Flow = Fixture Units [l/s]   |   Pipe series: {1}").format(
        config.system_type, config.series_summary, config.fill,
        rt.TABLE_NAMES[config.fill], rt.STACK_FILL, rt.STACK_TABLE_NAME,
        rc.STACK_MIN_ANGLE_DEG, scope_label(config), speed_label(config),
        ", ".join(config.systems))


def settings_rows(config):
    """Impostazioni del calcolo come coppie (voce, valore), per report ed Excel."""
    return [
        ("Flow of each pipe", "Fixture Units value, read as l/s"),
        ("System type", config.system_type),
        ("Systems", ", ".join(config.systems)),
        ("Scope", scope_label(config)),
        ("Pipe series (Table 6.2 column, from the pipe segments)", config.series_summary),
        ("Horizontal pipes", "f = {0}, {1}".format(config.fill, rt.TABLE_NAMES[config.fill])),
        ("Velocity of sloped pipes at the design flow", speed_label(config)),
        ("Vertical pipes", "f = {0}, {1} (steeper than {2:.0f} deg; 45 deg runs are "
                           "horizontal)".format(rt.STACK_FILL, rt.STACK_TABLE_NAME,
                                                rc.STACK_MIN_ANGLE_DEG)),
    ]


def export_info(config):
    """Dati per l'esportazione Excel della finestra dei risultati."""
    import datetime
    model = doc.Title or "Model"
    if model.lower().endswith(".rvt"):
        model = model[:-4]
    try:
        folder = os.path.dirname(doc.PathName) if doc.PathName else ""
    except Exception:
        folder = ""
    now = datetime.datetime.now()
    rows = [("Model", model), ("Date", now.strftime("%d.%m.%Y %H:%M"))] + settings_rows(config)
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in model)
    name = "{0}_RainwaterCalculation_{1}.xlsx".format(safe, now.strftime("%Y%m%d"))
    return rres.ExportInfo("Rainwater flow calculation - " + model, rows, folder, name)


def print_settings(config):
    output.print_table(
        table_data=[list(row) for row in settings_rows(config)],
        title="Settings",
        columns=["Setting", "Value"],
    )


def fitting_summary(report):
    if report is None:
        return "-"
    kept = report.checked - report.resized
    text = "{0} connected: {1} resized, {2} already matching".format(
        report.checked, report.resized, kept)
    if report.bridged:
        text += "; {0} transition(s) added toward pipes that keep their size".format(
            report.bridged)
    return text


def print_changes(results, pipes_by_id):
    output.print_md("## Changes")
    data = []
    for outcome in results:
        section = outcome.section
        ids = [pipes_by_id[i].Id for i in section.pipe_ids]
        data.append([
            section.system,
            "{0} {1}".format(section.section or "-", section.run),
            output.linkify(ids, title="{0} pipe(s)".format(len(ids))),
            "DN {0} / De {1}".format(section.model_dn_label, section.model_de_label),
            outcome.size.label,
            outcome.result,
            outcome.changed,
            fitting_summary(outcome.fittings),
            "; ".join(outcome.messages),
        ])
    output.print_table(
        table_data=data,
        columns=["System", "Section", "Pipes", "Model size [mm]", "New size [mm]",
                 "Result", "Pipes changed", "Fittings", "Messages"],
    )
    failed = [o for o in results if o.result == RESULT_FAILED]
    if failed:
        output.print_md(":warning: {0} section(s) were not changed: Revit rolled them back "
                        "gave a different size, or a fitting could not be resized (see "
                        "Messages). The other sections "
                        "were changed.".format(len(failed)))
    output.print_md("Run the tool again to check the network with the new diameters: Revit "
                    "recalculates fixture units and sections after the change.")


def report_header():
    output.print_md("# Rainwater Flow Calculation")


# =============================================================================
# MAIN
# =============================================================================

class OnlyThesePipes(ISelectionFilter):
    """Durante la navigazione si possono cliccare solo i tubi della riga."""

    def __init__(self, ids):
        self.ids = set(ids)

    def AllowElement(self, element):
        try:
            return get_element_id_value(element.Id) in self.ids
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return False


NAVIGATE_PROMPT = ("Rainwater Sections: the elements of the row are selected. Navigate the view "
                   "(pan, zoom, orbit), then press Finish or Esc to return to the table")


def navigate(ids, elements_by_id):
    """Seleziona e inquadra tubi (o terminali), poi lascia navigare la vista attiva
    finche' l'utente non preme Fine o Esc."""
    pipes = [elements_by_id[i] for i in ids
             if i in elements_by_id and elements_by_id[i].IsValidObject]
    if not pipes:
        return
    element_ids = List[DB.ElementId]([pipe.Id for pipe in pipes])
    uidoc.Selection.SetElementIds(element_ids)
    try:
        uidoc.ShowElements(element_ids)
    except Exception:
        pass
    references = List[DB.Reference]([DB.Reference(pipe) for pipe in pipes])
    try:
        uidoc.Selection.PickObjects(ObjectType.Element, OnlyThesePipes(ids), NAVIGATE_PROMPT,
                                    references)
    except OperationCanceledException:
        pass
    except Exception:
        pass
    # La selezione resta sui tubi della riga anche dopo il ritorno alla tabella.
    try:
        uidoc.Selection.SetElementIds(element_ids)
    except Exception:
        pass


def print_error_report(config, results, pipes_by_id, fixture_errors=None):
    """Report in uscita, solo se una Section o un parametro dei terminali non e'
    stato modificato."""
    output.close_others()
    output.set_title("Rainwater Flow Calculation")
    report_header()
    failed = [o for o in results if o.result == RESULT_FAILED]
    if failed:
        output.print_md(":warning: **{0} of {1} section part(s) could not be resized.** "
                        "The reason is in the Messages column; the other parts were "
                        "changed.".format(len(failed), len(results)))
    if fixture_errors:
        output.print_md(":warning: **{0} fixture change(s) could not be written.** The "
                        "other changes were written.".format(len(fixture_errors)))
        output.print_table(table_data=[[message] for message in fixture_errors],
                           title="Fixtures", columns=["Message"])
    print_settings(config)
    if results:
        print_changes(results, pipes_by_id)


class OnlyFittings(ISelectionFilter):
    def AllowElement(self, element):
        return is_fitting(element)

    def AllowReference(self, reference, position):
        return False


def _connected_text(connector):
    refs = linked_connectors(connector)
    if not refs:
        return "free"
    owner = refs[0].Owner
    if is_pipe(owner):
        return "pipe {0} ({1})".format(get_element_id_value(owner.Id), dn_label(refs[0]))
    if is_fitting(owner):
        return "{0} ({1})".format(fitting_label(owner), dn_label(refs[0]))
    try:
        category = owner.Category.Name
    except Exception:
        category = "element"
    return "{0} {1}".format(category, get_element_id_value(owner.Id))


def _state_text(fitting):
    """Stato del raccordo per la prova: misura di ogni connettore e cosa vi e' collegato."""
    if not fitting.IsValidObject:
        return "the fitting no longer exists"
    lines = []
    for index, connector in enumerate(piping_connectors(fitting)):
        lines.append("{0}. {1} to {2}".format(index + 1, dn_label(connector),
                                             _connected_text(connector)))
    return "<br>".join(lines)


def _test_attempts(fitting, connector, radius, family_map):
    """[(nome, azione)] dei tentativi da provare sul connettore scelto."""
    def by_connector():
        connector.Radius = radius

    def by_parameter():
        info = _family_info(fitting, family_map)
        entries = _family_params_of(fitting, connector, family_map) if info is not None else []
        done = []
        if entries:
            for name, factor, _i, _f in entries:
                kind, driver, value = _follow(info, name, radius * factor)
                if kind == "blocked":
                    raise ValueError(driver)
                if kind == "type":
                    raise ValueError("{0} is a type parameter".format(driver))
                param = fitting.LookupParameter(driver)
                if param is None:
                    raise ValueError("{0} not found on the instance".format(driver))
                param.Set(value)
                done.append("{0} = {1} mm".format(driver, fmt_de(value * FT_TO_MM)))
        else:
            for param, factor in _associated_params(fitting, connector, None):
                param.Set(radius * factor)
                done.append("{0} = {1} mm".format(_param_name(param),
                                                  fmt_de(radius * factor * FT_TO_MM)))
        if not done:
            raise ValueError("no associated parameter found ({0})".format(
                _family_entries_text(fitting, family_map)))
        return ", ".join(done)

    def by_parameter_and_transition():
        detail = by_parameter()
        doc.Regenerate()
        mismatched = [link for link in fitting_links(fitting)
                      if link.to_pipe and not link.matches]
        if not mismatched:
            return detail + "; no pipe left with a different size"
        added = _bridge(fitting, mismatched)
        return detail + "; transition added toward pipe " + ", ".join(
            "{0} ({1})".format(element_id, fitting_label(transition))
            for transition, element_id in added)

    return [("Connector size, as the Diameter field of the options bar", by_connector),
            ("Associated family parameter (for example NR)", by_parameter),
            ("Associated family parameter, then a transition toward the pipes that keep their "
             "size (what the tool does when it resizes)", by_parameter_and_transition)]


def run_fitting_test():
    """SHIFT + clic: prova a cambiare la misura di un connettore di un raccordo,
    in una transazione che poi si annulla, e mostra cosa succede."""
    try:
        reference = uidoc.Selection.PickObject(
            ObjectType.Element, OnlyFittings(), "Pick the pipe fitting to test (Esc to cancel)")
    except OperationCanceledException:
        return
    fitting = doc.GetElement(reference.ElementId)
    connectors = piping_connectors(fitting)
    if not connectors:
        forms.alert("The element has no pipe connectors.", title="Fitting test")
        return
    labels = ["{0}. {1} to {2}".format(i + 1, dn_label(c), _connected_text(c))
              for i, c in enumerate(connectors)]
    chosen = forms.SelectFromList.show(labels, title="Connector to resize", button_name="Select",
                                       multiselect=False)
    if not chosen:
        return
    connector_index = labels.index(chosen)
    connector_id = connectors[connector_index].Id
    text = forms.ask_for_string(
        default=fmt_de((radius_mm(connectors[connector_index]) or 0) * 2),
        prompt="Nominal diameter DN [mm] to give to that connector",
        title="Fitting test")
    if not text:
        return
    try:
        dn_mm = float(text.replace(",", "."))
    except ValueError:
        forms.alert("Not a number: {0}".format(text), title="Fitting test")
        return
    radius = dn_mm / 2.0 / FT_TO_MM

    family_map = scan_fitting_families([fitting])
    before = _state_text(fitting)
    rows = []
    preprocessor = ResizeFailurePreprocessor()
    transaction = DB.Transaction(doc, "RNW fitting test (rolled back)")
    handling = transaction.GetFailureHandlingOptions()
    handling = handling.SetForcedModalHandling(False)
    handling = handling.SetClearAfterRollback(True)
    handling = handling.SetFailuresPreprocessor(preprocessor)
    transaction.SetFailureHandlingOptions(handling)
    transaction.Start()
    try:
        for name, action in _test_attempts(fitting, connector_by_id(fitting, connector_id),
                                           radius, family_map):
            sub = DB.SubTransaction(doc)
            sub.Start()
            try:
                detail = action()
                doc.Regenerate()
                result = "accepted" + (": " + detail if detail else "")
            except Exception as error:
                result = "refused: " + _error_text(error)
            rows.append([name, result, _state_text(fitting)])
            _rollback(sub)
    finally:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()

    output.close_others()
    output.set_title("Fitting test")
    output.print_md("# Fitting test (nothing was changed)")
    output.print_md("**{0}**, part type {1}. Connector {2} to DN {3}.".format(
        fitting_label(fitting), part_type_name(fitting) or "-", connector_index + 1, fmt_de(dn_mm)))
    output.print_md("Size parameters of the family: {0}".format(
        _family_entries_text(fitting, family_map)))
    output.print_md("Before: <br>{0}".format(before))
    output.print_table(table_data=rows, columns=["Attempt", "Result", "Connectors after the attempt"])
    if preprocessor.messages:
        output.print_md("Revit messages: {0}".format("; ".join(preprocessor.messages)))
    output.print_md("If a new reducer appears in the connectors after an attempt, Revit adds "
                    "transitions by itself when the fitting size changes.")


class Analysis(object):
    """Risultato di un calcolo della rete, per la finestra e per le modifiche."""

    def __init__(self, entries, fixtures, symbols, fixture_data, elements_by_id):
        self.entries = entries
        self.fixtures = fixtures            # {id: plumbing fixture}
        self.symbols = symbols              # {id testo: FamilySymbol}
        self.fixture_data = fixture_data
        self.elements_by_id = elements_by_id


def analyze(config, pipes_by_id, type_names, system_names):
    """Verifica i tubi scelti e ordina la rete; None se non c'e' nessun tubo.

    Rilegge portate e parametri dal modello: dopo la scrittura dei terminali
    (Apply and Recalculate) da' le portate aggiornate da Revit.
    """
    limits = rc.Limits(config.fill, None, config.vmin, config.vmax)
    size_cache = {}
    series_cache = {}
    code_cache = {}
    mark_cache = {}
    wanted_systems = set(config.systems)
    rows = []
    for pipe_id, pipe in pipes_by_id.items():
        if type_names[pipe_id] != config.system_type:
            continue
        if system_names[pipe_id] not in wanted_systems:
            continue
        rows.append(build_check(pipe, type_names[pipe_id], limits, size_cache, series_cache,
                                code_cache, mark_cache))
    descriptions = sorted(set(description for _series, description in series_cache.values()))
    config.series_summary = "; ".join(descriptions) or "-"
    if not rows:
        return None

    sections = rc.group_sections(rows, limits)
    entries, fixtures = ordered_entries(sections, pipes_by_id)
    elements_by_id = dict(pipes_by_id)
    elements_by_id.update(fixtures)
    symbols = fixture_types(fixtures)
    fixture_data = read_fixture_data(entries, fixtures, symbols)
    return Analysis(entries, fixtures, symbols, fixture_data, elements_by_id)


def main():
    if int(doc.Application.VersionNumber) < 2022:
        forms.alert("This tool requires Revit 2022 or later.", exitscript=True)

    if __shiftclick__:  # noqa: F821
        run_fitting_test()
        return

    pipes = collect_pipes()
    if not pipes:
        forms.alert("The model has no pipes.", title="Rainwater Flow Calculation")
        return

    type_cache = {}
    pipes_by_id = dict((get_element_id_value(p.Id), p) for p in pipes)
    type_names = dict((pid, system_type_name(p, type_cache)) for pid, p in pipes_by_id.items())
    system_names = dict((pid, system_name(p)) for pid, p in pipes_by_id.items())

    # Con una selezione all'avvio si calcolano i sistemi interi degli elementi
    # selezionati: la finestra li propone gia' spuntati.
    type_of_system = {}
    for pipe_id in pipes_by_id:
        type_of_system.setdefault(system_names[pipe_id], type_names[pipe_id])
    selected_count, preselected = selection_systems(type_of_system)

    config = rui.show_dialog(system_type_items(pipes, type_names, system_names),
                             preselected, selected_count)
    if config is None:
        script.exit()

    analysis = analyze(config, pipes_by_id, type_names, system_names)
    if analysis is None:
        forms.alert("No pipe of the chosen systems was found.",
                    sub_msg="Systems: {0}.".format(", ".join(config.systems)),
                    title="Rainwater Flow Calculation")
        return

    # Finestra, navigazione e ricalcolo si alternano finche' l'utente non applica
    # o chiude.
    state = None
    fixture_errors = []
    while True:
        form = rres.show_results(analysis.entries, settings_text(config), state,
                                 export_info(config), analysis.fixture_data)
        state = form.state
        if form.kind == rres.KIND_SHOW:
            navigate(form.show_ids, analysis.elements_by_id)
            continue
        if form.kind == rres.KIND_RECALC:
            errors = write_fixture_edits(form.fixture_edits, analysis.fixtures,
                                         analysis.symbols)
            if errors:
                fixture_errors.extend(errors)
                forms.alert("{0} fixture change(s) could not be written.".format(len(errors)),
                            sub_msg="\n".join(errors[:10]) +
                            "\n\nThe report at the end lists them all.",
                            title="Rainwater Flow Calculation")
            refreshed = analyze(config, pipes_by_id, type_names, system_names)
            if refreshed is None:
                break
            analysis = refreshed
            state = state.carry()
            continue
        break

    results = []
    if form.kind == rres.KIND_APPLY:
        if form.fixture_edits is not None and form.fixture_edits.count:
            fixture_errors.extend(write_fixture_edits(form.fixture_edits, analysis.fixtures,
                                                      analysis.symbols))
        if form.changes:
            results = apply_changes(form.changes, pipes_by_id)
    if fixture_errors or any(o.result == RESULT_FAILED for o in results):
        print_error_report(config, results, pipes_by_id, fixture_errors)


main()
