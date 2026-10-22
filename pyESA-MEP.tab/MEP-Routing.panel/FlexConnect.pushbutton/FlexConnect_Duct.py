# -*- coding: utf-8 -*-
"""FlexConnect sui condotti: estremita' libera (come per i tubi) o presa sul fianco del condotto.

La presa e' un raccordo creato da Document.NewTakeoffFitting sul condotto principale, partendo dal
connettore finale del flex. Il flex ha la misura del connettore del terminale, perche' un condotto
rettangolare non ha un diametro da ereditare.
"""
import FlexConnect_Build as fb

END_FACE_MM = 100.0         # un clic entro questa distanza (lungo l'asse) dalla faccia di un'estremita' libera vale come estremita'
TAP_CLEARANCE_MM = 50.0     # il flex finisce a questa distanza dalla superficie del condotto, la presa fa il resto
MIN_LATERAL_FT = 0.01
TAP_PART_TYPES = (fb.PartType.TapAdjustable, fb.PartType.TapPerpendicular,
                  fb.PartType.SpudAdjustable, fb.PartType.SpudPerpendicular)


def plan_connection(duct, point, equipment, chooser):
    """Decide il tipo di collegamento. Ritorna ("end", estremita', connettore terminale, None)
    oppure ("tap", None, connettore terminale, punto sull'asse del condotto)."""
    domain = fb.DUCT.domain
    targets = fb.free_connectors(equipment, domain)
    if not targets:
        raise fb.FlexConnectError("The equipment has no free duct connector.")

    # estremita': il clic sulla faccia terminale cade sul piano del connettore, quindi la distanza
    # lungo l'asse e' ~0; un clic sul fianco, anche vicino all'estremita', e' piu' lontano
    ends = fb.free_connectors(duct, domain)
    if ends:
        def axial_distance(conn):
            return abs((point - conn.Origin).DotProduct(conn.CoordinateSystem.BasisZ))
        end = min(ends, key=axial_distance)
        if axial_distance(end) <= END_FACE_MM / fb.MM_PER_FOOT:
            end_conn, eq_conn = fb.select_target([(fb.system_of(end), end.Origin, end)], targets, chooser)
            return "end", end_conn, eq_conn, None

    axis_pt = duct.Location.Curve.Project(point).XYZPoint
    duct_system = fb.system_of(next(iter(duct.ConnectorManager.Connectors)))
    _, eq_conn = fb.select_target([(duct_system, axis_pt, None)], targets, chooser)
    return "tap", None, eq_conn, axis_pt


def _half_extent(duct, horizontal):
    """Mezza dimensione della sezione verso il fianco, in piedi.
    Tondo: il raggio. Rettangolare: mezza larghezza se il fianco e' verticale (condotto non ruotato),
    altrimenti mezzo lato maggiore."""
    diameter = duct.get_Parameter(fb.DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
    if diameter is not None and diameter.HasValue:
        return diameter.AsDouble() / 2.0
    width = duct.get_Parameter(fb.DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble()
    height = duct.get_Parameter(fb.DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble()
    return width / 2.0 if horizontal else max(width, height) / 2.0


def _side_direction(axis, to_eq):
    """(direzione dello stacco, True se orizzontale). Con il condotto orizzontale lo stacco e' sul fianco
    verticale, a quota centrale: perpendicolare all'asse, in orizzontale, dal lato del terminale.
    Con il condotto verticale non c'e' un fianco privilegiato: si usa la componente laterale."""
    side = axis.CrossProduct(fb.DB.XYZ.BasisZ)
    if side.GetLength() > 1e-6:
        side = side.Normalize()
        return (side if side.DotProduct(to_eq) >= 0 else side.Negate()), True
    lateral = to_eq - axis.Multiply(to_eq.DotProduct(axis))
    if lateral.GetLength() < MIN_LATERAL_FT:
        return None, False
    return lateral.Normalize(), False


def connect_end(doc, duct, end_conn, eq_conn, flex_type_id, gap_mm):
    """Flex dall'estremita' libera di un condotto al terminale. In una transazione.

    Il flex ha sempre la misura del connettore del terminale. Se l'estremita' del condotto ha forma o
    misura diversa, il flex parte a gap_mm dall'estremita' e una transizione (NewTransitionFitting,
    il primo connettore e' quello del flex) lo collega al condotto; altrimenti si collega direttamente.
    Ritorna True se e' stata inserita la transizione."""
    kind = fb.DUCT
    if eq_conn.Shape != fb.DB.ConnectorProfileType.Round:
        raise fb.FlexConnectError("The connector is not round: a flex duct cannot take its size.")
    level_id, system_id = fb.trunk_reference(kind, duct)
    diameter = eq_conn.Radius * 2.0
    out_dir = end_conn.CoordinateSystem.BasisZ

    same_profile = (end_conn.Shape == fb.DB.ConnectorProfileType.Round and
                    abs(end_conn.Radius * 2.0 - diameter) <= fb.DIAMETER_TOLERANCE_FT)
    start_pt = end_conn.Origin
    if not same_profile:
        start_pt = end_conn.Origin + out_dir.Multiply(max(gap_mm, 1.0) / fb.MM_PER_FOOT)
    if start_pt.DistanceTo(eq_conn.Origin) < fb.MIN_FLEX_LENGTH_FT:
        raise fb.FlexConnectError("The duct end is too close to the connector for a flex.")

    flex = fb.create_flex(kind, doc, system_id, flex_type_id, level_id, start_pt, eq_conn.Origin,
                          out_dir, eq_conn.CoordinateSystem.BasisZ.Negate())
    fb.set_size(kind, flex, diameter)
    doc.Regenerate()
    flex_start = fb.nearest_connector(flex, start_pt)
    if same_profile:
        flex_start.ConnectTo(end_conn)
    else:
        doc.Create.NewTransitionFitting(flex_start, end_conn)
    fb.nearest_connector(flex, eq_conn.Origin).ConnectTo(eq_conn)
    fb.copy_insulation(kind, doc, duct, [flex])
    return not same_profile


def connect_tap(doc, duct, axis_pt, eq_conn, flex_type_id):
    """Flex dal connettore del terminale al fianco del condotto, chiuso da una presa. In una transazione.
    Ritorna False: la presa non ha riduzioni."""
    kind = fb.DUCT
    if eq_conn.Shape != fb.DB.ConnectorProfileType.Round:
        raise fb.FlexConnectError("The connector is not round: a flex duct cannot take its size.")
    level_id, system_id = fb.trunk_reference(kind, duct)
    diameter = eq_conn.Radius * 2.0

    axis = duct.Location.Curve.Direction
    to_eq = eq_conn.Origin - axis_pt
    side, horizontal = _side_direction(axis, to_eq)
    if side is None or abs(to_eq.DotProduct(side)) < MIN_LATERAL_FT:
        raise fb.FlexConnectError("The connector is aligned with the duct axis: no side tap is possible.")
    end_pt = axis_pt + side.Multiply(_half_extent(duct, horizontal) + TAP_CLEARANCE_MM / fb.MM_PER_FOOT)

    if eq_conn.Origin.DistanceTo(end_pt) < fb.MIN_FLEX_LENGTH_FT:
        raise fb.FlexConnectError("The connector is too close to the duct for a flex.")

    flex = fb.create_flex(kind, doc, system_id, flex_type_id, level_id, eq_conn.Origin, end_pt,
                          eq_conn.CoordinateSystem.BasisZ, side.Negate())
    fb.set_size(kind, flex, diameter)
    doc.Regenerate()
    fb.nearest_connector(flex, eq_conn.Origin).ConnectTo(eq_conn)
    doc.Create.NewTakeoffFitting(fb.nearest_connector(flex, end_pt), duct)
    fb.copy_insulation(kind, doc, duct, [flex])
    return False


def build(doc, duct, plan, flex_type_id, gap_mm, tap_symbol_id=None):
    """Esegue il piano. Ritorna True se e' stata inserita una riduzione (transizione).
    tap_symbol_id: famiglia di derivazione da usare per la presa (None = preferenze
    di instradamento del tipo di condotto)."""
    mode, end_conn, eq_conn, axis_pt = plan
    if mode == "end":
        return connect_end(doc, duct, end_conn, eq_conn, flex_type_id, gap_mm)
    if tap_symbol_id is None:
        return connect_tap(doc, duct, axis_pt, eq_conn, flex_type_id)
    return _with_tap_family(doc, duct, tap_symbol_id,
                            lambda: connect_tap(doc, duct, axis_pt, eq_conn, flex_type_id))


def collect_tap_symbols(doc):
    """[(etichetta "Famiglia : Tipo", ElementId)] dei raccordi di derivazione (tap, spud) del progetto."""
    part_ints = [int(p) for p in TAP_PART_TYPES]
    symbols = []
    collector = fb.DB.FilteredElementCollector(doc).OfCategory(fb.DB.BuiltInCategory.OST_DuctFitting)
    for symbol in collector.OfClass(fb.DB.FamilySymbol):
        family = symbol.Family
        param = family.get_Parameter(fb.DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE)
        if param is None or param.AsInteger() not in part_ints:
            continue
        symbols.append(("{} : {}".format(family.Name, fb.revit.query.get_name(symbol)), symbol.Id))
    symbols.sort(key=lambda item: item[0].lower())
    return symbols


def _with_tap_family(doc, duct, symbol_id, action):
    """Esegue action con la famiglia scelta in testa alle regole Junctions del tipo di condotto.

    La regola e' temporanea: viene tolta prima di uscire, quindi il tipo di condotto resta com'era.
    Da chiamare dentro la transazione del collegamento (se questa fallisce, torna tutto indietro)."""
    symbol = doc.GetElement(symbol_id)
    if not symbol.IsActive:
        symbol.Activate()
    manager = doc.GetElement(duct.GetTypeId()).RoutingPreferenceManager
    group = fb.DB.RoutingPreferenceRuleGroupType.Junctions
    manager.AddRule(group, fb.DB.RoutingPreferenceRule(symbol_id, "FlexConnect"), 0)
    try:
        return action()
    finally:
        manager.RemoveRule(group, 0)
