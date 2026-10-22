# -*- coding: utf-8 -*-
"""Logica Revit di FlexConnect condivisa tra tubi e condotti: selezione, connettori, creazione del flex.

Il collegamento sul fianco di un condotto (presa) sta in FlexConnect_Duct.py.
Nessun documento o UIDocument a livello di modulo (Rocket mode): arrivano dai parametri.
"""
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System.Collections.Generic import List
from Autodesk.Revit.DB.Plumbing import Pipe, FlexPipe, FlexPipeType, PipeInsulation, PipeSystemType, PipingSystemType
from Autodesk.Revit.DB.Mechanical import Duct, FlexDuct, FlexDuctType, DuctInsulation, DuctSystemType, MechanicalSystemType
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

from pyrevit import DB, revit

PartType = DB.PartType

MM_PER_FOOT = 304.8
MIN_FLEX_LENGTH_FT = 0.05   # sotto i ~15 mm Revit rifiuta il flex
DIAMETER_TOLERANCE_FT = 1e-4

WILDCARD_SYSTEMS = (PipeSystemType.Global, PipeSystemType.UndefinedSystemType,
                    DuctSystemType.Global, DuctSystemType.UndefinedSystemType)

PIPE_EQUIPMENT = (DB.BuiltInCategory.OST_MechanicalEquipment, DB.BuiltInCategory.OST_PlumbingEquipment)
DUCT_EQUIPMENT = (DB.BuiltInCategory.OST_DuctTerminal, DB.BuiltInCategory.OST_MechanicalEquipment)
ALL_EQUIPMENT = (DB.BuiltInCategory.OST_DuctTerminal, DB.BuiltInCategory.OST_MechanicalEquipment,
                 DB.BuiltInCategory.OST_PlumbingEquipment)


class FlexConnectError(Exception):
    """Errore atteso, con il testo (inglese) da mostrare all'utente."""


class FlexConnectCancelled(Exception):
    """L'utente ha chiuso la finestra di scelta del connettore."""


class Kind(object):
    """Differenze tra il caso tubo e il caso condotto."""

    def __init__(self, name, domain, rigid, flex, flex_type, insulation, size_param, system_param, system_class):
        self.name = name
        self.domain = domain
        self.rigid = rigid
        self.flex = flex
        self.flex_type = flex_type
        self.insulation = insulation
        self.size_param = size_param
        self.system_param = system_param
        self.system_class = system_class


PIPE = Kind("pipe", DB.Domain.DomainPiping, Pipe, FlexPipe, FlexPipeType, PipeInsulation,
            DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM, PipingSystemType)
DUCT = Kind("duct", DB.Domain.DomainHvac, Duct, FlexDuct, FlexDuctType, DuctInsulation,
            DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM, DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM, MechanicalSystemType)


def kind_of(trunk):
    return DUCT if isinstance(trunk, Duct) else PIPE


def get_id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


# ---------------------------------------------------------------- selezione

def is_pipe(element):
    return isinstance(element, Pipe)


def is_duct(element):
    return isinstance(element, Duct)


def is_trunk(element):
    return is_pipe(element) or is_duct(element)


def in_categories(element, categories):
    category = element.Category
    if category is None:
        return False
    cat_id = get_id_value(category.Id)
    return any(cat_id == int(c) for c in categories)


class PredicateFilter(ISelectionFilter):
    def __init__(self, predicate):
        self.predicate = predicate

    def AllowElement(self, element):
        return self.predicate(element)

    def AllowReference(self, reference, position):
        # con PointOnElement Revit chiede anche il consenso sul riferimento (faccia, spigolo):
        # il filtro sull'elemento e' gia' stato applicato da AllowElement
        return True


def _pick(uidoc, predicate, prompt):
    """(elemento, punto cliccato)."""
    ref = uidoc.Selection.PickObject(ObjectType.PointOnElement, PredicateFilter(predicate), prompt)
    return uidoc.Document.GetElement(ref), ref.GlobalPoint


def preselected_trunk(uidoc):
    """Il tubo o condotto gia' selezionato, se la selezione e' esattamente uno."""
    ids = list(uidoc.Selection.GetElementIds())
    if len(ids) != 1:
        return None
    element = uidoc.Document.GetElement(ids[0])
    return element if is_trunk(element) else None


def _equipment_predicate_for(trunk):
    categories = DUCT_EQUIPMENT if is_duct(trunk) else PIPE_EQUIPMENT
    return lambda e: in_categories(e, categories)


def _trunk_predicate_for(equipment):
    if in_categories(equipment, (DB.BuiltInCategory.OST_DuctTerminal,)):
        return is_duct
    if in_categories(equipment, (DB.BuiltInCategory.OST_PlumbingEquipment,)):
        return is_pipe
    return is_trunk


def pick_elements(uidoc):
    """(tubo o condotto, punto cliccato su di esso o None se preselezionato, attrezzatura).

    L'ordine e' libero: prima il tubo/condotto o prima l'attrezzatura/terminale."""
    trunk = preselected_trunk(uidoc)
    if trunk is not None:
        equipment, _ = _pick(uidoc, _equipment_predicate_for(trunk), "Select the equipment or air terminal")
        return trunk, None, equipment

    first, first_pt = _pick(uidoc, lambda e: is_trunk(e) or in_categories(e, ALL_EQUIPMENT),
                            "Select the pipe or duct, or the equipment or air terminal")
    if is_trunk(first):
        equipment, _ = _pick(uidoc, _equipment_predicate_for(first), "Select the equipment or air terminal")
        return first, first_pt, equipment
    trunk, trunk_pt = _pick(uidoc, _trunk_predicate_for(first), "Select the pipe or duct")
    return trunk, trunk_pt, first


def pick_next_duct_connection(uidoc, duct):
    """Per i collegamenti successivi sullo stesso condotto: (punto, terminale), o None con Esc."""
    from Autodesk.Revit.Exceptions import OperationCanceledException
    try:
        equipment, _ = _pick(uidoc, _equipment_predicate_for(duct),
                             "Select the next air terminal or equipment (Esc to finish)")
        _, point = _pick(uidoc, lambda e: e.Id == duct.Id, "Click the point on the duct")
    except OperationCanceledException:
        return None
    return point, equipment


def pick_point_on(uidoc, element):
    _, point = _pick(uidoc, lambda e: e.Id == element.Id, "Click the point on the duct")
    return point


# ---------------------------------------------------------------- tipi flex

def collect_flex_types(doc, kind):
    """[(nome, ElementId)] dei tipi di flex del tipo indicato, in ordine alfabetico."""
    types = [(revit.query.get_name(t), t.Id)
             for t in DB.FilteredElementCollector(doc).OfClass(kind.flex_type)]
    types.sort(key=lambda item: item[0].lower())
    return types


# ---------------------------------------------------------------- connettori

def _connector_manager(element):
    manager = getattr(element, "ConnectorManager", None)
    if manager is not None:
        return manager
    model = getattr(element, "MEPModel", None)
    return model.ConnectorManager if model is not None else None


def free_connectors(element, domain):
    """Connettori liberi dell'elemento nel dominio indicato (tubazioni o canali)."""
    manager = _connector_manager(element)
    if manager is None:
        return []
    result = []
    for conn in manager.Connectors:
        if conn.ConnectorType != DB.ConnectorType.End:
            continue
        if conn.Domain != domain or conn.IsConnected:
            continue
        result.append(conn)
    return result


def system_of(conn):
    try:
        if conn.Domain == DB.Domain.DomainPiping:
            return conn.PipeSystemType
        return conn.DuctSystemType
    except Exception:
        return None


def is_wildcard(system):
    return system is None or system in WILDCARD_SYSTEMS


def select_target(entries, targets, chooser):
    """(payload, connettore bersaglio): la coppia piu' vicina dello stesso sistema.

    entries: [(sistema, punto, payload)] lato tubo/condotto; targets: connettori liberi dell'attrezzatura.
    Prima si cercano le coppie con lo stesso sistema; solo se non ce ne sono si accettano i connettori
    Global / Undefined. Con piu' connettori candidati e un chooser, decide il chooser: riceve la lista
    [(distanza, payload, connettore)] ordinata per distanza e ritorna un suo elemento, o None per annullare."""
    exact, wildcard = [], []
    for system, point, payload in entries:
        for target in targets:
            target_system = system_of(target)
            dist = point.DistanceTo(target.Origin)
            if system == target_system and not is_wildcard(system):
                exact.append((dist, payload, target))
            elif is_wildcard(system) or is_wildcard(target_system):
                wildcard.append((dist, payload, target))

    pairs = exact or wildcard
    if not pairs:
        raise FlexConnectError("The equipment has no free connector of the same system type.")

    # un candidato per connettore dell'attrezzatura, con l'origine piu' vicina
    nearest = {}
    for item in pairs:
        key = item[2].Id
        if key not in nearest or item[0] < nearest[key][0]:
            nearest[key] = item
    candidates = sorted(nearest.values(), key=lambda item: item[0])

    best = candidates[0]
    if chooser is not None and len(candidates) > 1:
        best = chooser(candidates)
        if best is None:
            raise FlexConnectCancelled()
    return best[1], best[2]


def find_connector_pair(trunk, equipment, chooser=None):
    """(estremita' libera del tubo/condotto, connettore dell'attrezzatura)."""
    kind = kind_of(trunk)
    ends = free_connectors(trunk, kind.domain)
    if not ends:
        raise FlexConnectError("The {} has no free end.".format(kind.name))
    targets = free_connectors(equipment, kind.domain)
    if not targets:
        raise FlexConnectError("The equipment has no free {} connector.".format(
            "piping" if kind is PIPE else "duct"))
    return select_target([(system_of(c), c.Origin, c) for c in ends], targets, chooser)


def connector_size_text(conn):
    try:
        if conn.Shape == DB.ConnectorProfileType.Round:
            return "{:g} mm".format(round(conn.Radius * 2.0 * MM_PER_FOOT, 1))
        return "{:g} x {:g} mm".format(round(conn.Width * MM_PER_FOOT, 1), round(conn.Height * MM_PER_FOOT, 1))
    except Exception:
        return "?"


def describe_candidate(candidate):
    """Testo di una riga della finestra di scelta: id connettore, sistema, misura, distanza."""
    dist, _, target = candidate
    return "Connector {}  |  {}  |  {}  |  {:g} mm away".format(
        target.Id, system_of(target), connector_size_text(target), round(dist * MM_PER_FOOT))


def nearest_connector(element, point):
    return min(element.ConnectorManager.Connectors, key=lambda c: c.Origin.DistanceTo(point))


# ---------------------------------------------------------------- creazione

def set_size(kind, element, diameter):
    param = element.get_Parameter(kind.size_param)
    if param is None or param.IsReadOnly:
        raise FlexConnectError("The size of the new element cannot be set.")
    param.Set(diameter)


def trunk_reference(kind, trunk):
    """(livello, tipo di sistema) del tubo o condotto di partenza."""
    level_param = trunk.get_Parameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM)
    level_id = level_param.AsElementId() if level_param is not None else trunk.ReferenceLevel.Id
    system_param = trunk.get_Parameter(kind.system_param)
    system_id = system_param.AsElementId() if system_param is not None else DB.ElementId.InvalidElementId
    if system_id == DB.ElementId.InvalidElementId:
        raise FlexConnectError("The {} has no system type.".format(kind.name))
    return level_id, system_id


def create_flex(kind, doc, system_id, flex_type_id, level_id, start_pt, end_pt, start_tangent, end_tangent):
    """Flex da start_pt a end_pt con le tangenti indicate (fallback senza tangenti)."""
    points = List[DB.XYZ]([start_pt, end_pt])
    try:
        return kind.flex.Create(doc, system_id, flex_type_id, level_id, start_tangent, end_tangent, points)
    except Exception:
        return kind.flex.Create(doc, system_id, flex_type_id, level_id, points)


MAX_PATTERN_VALUE = 8       # valori interi provati per scoprire le voci di "Flex Pattern"


def probe_flex_patterns(doc, kind, flex_type_id):
    """[(nome, valore intero)] delle voci del parametro "Flex Pattern" per il tipo di flex.

    L'API non elenca le voci: si crea un flex temporaneo, si provano i valori 0..7 leggendo il testo
    mostrato da Revit e si annulla tutto (transazione con RollBack). Lista vuota se non e' possibile
    (nessun livello o tipo di sistema, parametro non intero)."""
    level_id = DB.FilteredElementCollector(doc).OfClass(DB.Level).FirstElementId()
    system_id = DB.FilteredElementCollector(doc).OfClass(kind.system_class).FirstElementId()
    if level_id == DB.ElementId.InvalidElementId or system_id == DB.ElementId.InvalidElementId:
        return []
    found = []
    transaction = DB.Transaction(doc, "FlexConnect probe")
    transaction.Start()
    try:
        points = List[DB.XYZ]([DB.XYZ(0, 0, 0), DB.XYZ(10, 0, 0)])
        flex = kind.flex.Create(doc, system_id, flex_type_id, level_id, points)
        param = flex.get_Parameter(DB.BuiltInParameter.RBS_FLEX_PATTERN_PARAM)
        if param is not None and param.StorageType == DB.StorageType.Integer:
            seen = set()
            for value in range(MAX_PATTERN_VALUE):
                try:
                    param.Set(value)
                except Exception:
                    continue
                name = param.AsValueString()
                if param.AsInteger() != value or not name or name.isdigit() or name in seen:
                    continue
                seen.add(name)
                found.append((name, value))
    except Exception:
        found = []
    finally:
        transaction.RollBack()
    return found


def apply_pattern(flex, pattern):
    """Imposta "Flex Pattern" sul flex appena creato. pattern: valore intero o None (default del tipo)."""
    if pattern is None:
        return
    param = flex.get_Parameter(DB.BuiltInParameter.RBS_FLEX_PATTERN_PARAM)
    if param is None or param.IsReadOnly:
        raise FlexConnectError("The flex pattern cannot be set.")
    param.Set(int(pattern))


def copy_insulation(kind, doc, source, targets):
    """Applica ai nuovi elementi l'isolamento dell'elemento di partenza (stesso tipo e spessore), se c'e'."""
    ids = DB.InsulationLiningBase.GetInsulationIds(doc, source.Id)
    if ids.Count == 0:
        return
    insulation = doc.GetElement(list(ids)[0])
    for target in targets:
        kind.insulation.Create(doc, target.Id, insulation.GetTypeId(), insulation.Thickness)


def _needs_reducer(trunk_conn, eq_conn):
    if eq_conn.Shape != DB.ConnectorProfileType.Round:
        return True
    return abs(trunk_conn.Radius * 2.0 - eq_conn.Radius * 2.0) > DIAMETER_TOLERANCE_FT


def connect_end(doc, trunk, trunk_conn, eq_conn, flex_type_id, gap_mm, pattern=None):
    """Collega l'estremita' libera di un tubo o condotto all'attrezzatura. Da chiamare in una transazione.

    Stessa misura: il flex va dall'estremita' al connettore dell'attrezzatura.
    Misura (o forma) diversa: sull'attrezzatura si crea una riduzione seguita da un tratto rigido corto,
    lungo quanto il diametro, della misura del tubo/condotto; il flex va dall'estremita' al tratto corto.
    Se l'elemento di partenza e' isolato, flex e tratto corto ricevono lo stesso isolamento."""
    kind = kind_of(trunk)
    level_id, system_id = trunk_reference(kind, trunk)
    diameter = trunk_conn.Radius * 2.0
    out_dir = eq_conn.CoordinateSystem.BasisZ

    target_conn = eq_conn
    target_pt = eq_conn.Origin
    new_elements = []
    if _needs_reducer(trunk_conn, eq_conn):
        gap = max(gap_mm, 1.0) / MM_PER_FOOT
        stub_start = eq_conn.Origin + out_dir.Multiply(gap)
        stub_end = stub_start + out_dir.Multiply(diameter)  # lunghezza = DN
        stub = kind.rigid.Create(doc, system_id, trunk.GetTypeId(), level_id, stub_start, stub_end)
        set_size(kind, stub, diameter)
        doc.Regenerate()
        doc.Create.NewTransitionFitting(nearest_connector(stub, stub_start), eq_conn)
        target_conn = nearest_connector(stub, stub_end)
        target_pt = stub_end
        new_elements.append(stub)

    if trunk_conn.Origin.DistanceTo(target_pt) < MIN_FLEX_LENGTH_FT:
        raise FlexConnectError("The end is too close to the equipment connector for a flex.")

    flex = create_flex(kind, doc, system_id, flex_type_id, level_id, trunk_conn.Origin, target_pt,
                       trunk_conn.CoordinateSystem.BasisZ, target_conn.CoordinateSystem.BasisZ.Negate())
    set_size(kind, flex, diameter)
    apply_pattern(flex, pattern)
    doc.Regenerate()
    nearest_connector(flex, trunk_conn.Origin).ConnectTo(trunk_conn)
    nearest_connector(flex, target_pt).ConnectTo(target_conn)
    new_elements.append(flex)
    copy_insulation(kind, doc, trunk, new_elements)
    return flex
