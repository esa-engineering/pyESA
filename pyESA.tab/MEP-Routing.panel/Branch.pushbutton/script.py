# -*- coding: utf-8 -*-
__title__ = "Branch"
__doc__ = """Version = 0.1
Date    = 08.10.2026
________________________________________________________________
Description:
Creates a horizontal or inclined branch from vertical runs of pipe,
duct or cable tray (cable trays: only from the narrow side).

Instruction:
1. Select the vertical runs (or select them first) and press Finish
2. In the window choose the side of the branch, as seen on screen in
   the current plan view, the elevation (level + offset), the length
   and the angle of the branch
3. Elevation along the riser: a tee (the riser is split) or a tap.
   Junction option in the window: tap on ducts and tee on pipes
   (default), as in the Routing Preferences of the type (the other
   junction is tried if the first fails), always tee, always tap.
   Cable trays always get a tee.
   Elevation above or below the end of the riser: you are asked
   whether to turn with an elbow or to extend the riser and place a tee
   (cable trays always turn with a curve). What is already connected
   to that end is disconnected: the window lists those elements, and
   the report lists them again

The branch keeps the type and size of the riser. If a fitting does not
fit, that riser is left unchanged and the report says why.
The new parts keep the parameters and the insulation of the starting element.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import math

try:
    import traceback
except ImportError:
    traceback = None

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB.Plumbing import Pipe, PipeInsulation, PipingSystemType, PlumbingUtils
from Autodesk.Revit.DB.Mechanical import Duct, DuctInsulation, DuctLining, MechanicalSystemType, MechanicalUtils
from Autodesk.Revit.DB.Electrical import CableTray
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, DB, forms

import Branch_ui

XAML_FILE_NAME = "Branch_form.xaml"
CONFIG_NAME = "ESA_Branch"
DEFAULTS = {"direction": "right", "level": "", "offset_mm": 3000.0, "length_mm": 500.0,
            "angle_deg": 0.0, "stub_mm": 300.0, "junction": "standard"}

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


REPORT_CONFIG = "ESA_Report"  # impostazione condivisa da tutti i pulsanti, cambiata col pulsante Report


def want_report(total, problems, noun):
    """True se va stampato il report completo. Modalita' "problems" (predefinita): solo se qualcosa non e' andato a
    buon fine; "always": sempre; "never": mai, resta la notifica. Altrimenti notifica breve (e nessuna finestra)."""
    try:
        mode = script.get_config(REPORT_CONFIG).get_option("mode", "problems")
    except Exception:
        mode = "problems"
    if mode == "always" or (mode == "problems" and problems):
        return True
    if not total:
        return False
    text = "{} {}".format(total - problems, noun) + (", {} not done".format(problems) if problems else "")
    try:
        script.toast(text, title="pyESA")
    except Exception:
        return problems > 0  # notifica non disponibile: con dei problemi meglio il report che niente
    return False

MM_PER_FT = 304.8
VERTICAL_TOL = 0.999
MIN_REMAINING_MM = 20.0  # lunghezza minima di ogni tratto dopo l'inserimento dei raccordi
STAGE = [""]             # passo in corso, per i messaggi di errore
PHYSICAL_TYPES = (DB.ConnectorType.End, DB.ConnectorType.Curve, DB.ConnectorType.Physical)
ELBOW, EXTEND_TEE = "Elbow", "Extend and tee"

CATEGORY_NAMES = ((Pipe, "Pipe"), (Duct, "Duct"), (CableTray, "Cable Tray"))


class BranchError(Exception):
    """Errore atteso: il messaggio va nel report."""


def is_vertical_run(element):
    if not isinstance(element, (Pipe, Duct, CableTray)):
        return False
    curve = element.Location.Curve
    return isinstance(curve, DB.Line) and abs(curve.Direction.Z) > VERTICAL_TOL


class VerticalRunFilter(ISelectionFilter):
    """Tubi, canali e passerelle verticali."""

    def AllowElement(self, element):
        try:
            return is_vertical_run(element)
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return False


def category_name(element):
    for cls, name in CATEGORY_NAMES:
        if isinstance(element, cls):
            return name
    return "?"


def level_elevation(level):
    # ProjectElevation e' coerente con la geometria (coordinate interne), Elevation puo' riferirsi al Survey Point
    value = getattr(level, "ProjectElevation", None)
    return level.Elevation if value is None else value


def get_level_id(element):
    param = element.get_Parameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM)
    if param is not None and param.AsElementId() != DB.ElementId.InvalidElementId:
        return param.AsElementId()
    ref_level = getattr(element, "ReferenceLevel", None)
    return ref_level.Id if ref_level is not None else None


def nearest_connector(element, point):
    best, best_dist = None, None
    for c in connector_manager(element).Connectors:
        if c.ConnectorType not in PHYSICAL_TYPES:
            continue
        dist = c.Origin.DistanceTo(point)
        if best is None or dist < best_dist:
            best, best_dist = c, dist
    return best


def connector_manager(element):
    """ConnectorManager di un tratto MEP o di un raccordo (FamilyInstance)."""
    manager = getattr(element, "ConnectorManager", None)
    if manager is None:
        model = getattr(element, "MEPModel", None)
        manager = model.ConnectorManager if model is not None else None
    return manager


def free_connector_near(element, point):
    """Connettore non collegato, il piu' vicino al punto."""
    manager = connector_manager(element)
    if manager is None:
        return None
    best, best_dist = None, None
    for c in manager.Connectors:
        if c.IsConnected:
            continue
        dist = c.Origin.DistanceTo(point)
        if best is None or dist < best_dist:
            best, best_dist = c, dist
    return best


def physical_refs(conn, element):
    """Connettori fisici di altri elementi collegati a conn (esclude il connettore logico del sistema)."""
    return [r for r in conn.AllRefs if r.Owner.Id != element.Id and r.ConnectorType in PHYSICAL_TYPES]


# parametri istanza copiati dal tratto di partenza: quelli di progetto/condivisi e pochi built-in sicuri.
# Esclusi Mark (i duplicati generano avvisi), quote, offset, lunghezze e sistema: si ricalcolano da soli.
COPIED_BUILTIN = (DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, DB.BuiltInParameter.ELEM_PARTITION_PARAM,
                  DB.BuiltInParameter.PHASE_CREATED, DB.BuiltInParameter.PHASE_DEMOLISHED)
COPYABLE_STORAGE = (DB.StorageType.String, DB.StorageType.Integer, DB.StorageType.Double,
                    DB.StorageType.ElementId)


def copy_parameters(source, target):
    """Copia i parametri istanza dal tratto di partenza al nuovo. Ritorna il numero di valori copiati."""
    copied = 0
    for param in source.Parameters:
        try:
            if param.IsReadOnly or not param.HasValue or param.StorageType not in COPYABLE_STORAGE:
                continue
            bip = getattr(param.Definition, "BuiltInParameter", None)
            if bip is not None and bip != DB.BuiltInParameter.INVALID:
                if bip not in COPIED_BUILTIN:
                    continue
                dest = target.get_Parameter(bip)
            else:
                dest = target.LookupParameter(param.Definition.Name)
            if dest is None or dest.IsReadOnly or dest.StorageType != param.StorageType:
                continue
            if param.StorageType == DB.StorageType.String:
                value = param.AsString()
                if value is None:
                    continue
                dest.Set(value)
            elif param.StorageType == DB.StorageType.Integer:
                dest.Set(param.AsInteger())
            elif param.StorageType == DB.StorageType.Double:
                dest.Set(param.AsDouble())
            else:
                dest.Set(param.AsElementId())
            copied += 1
        except Exception:
            continue
    return copied


def insulation_specs(source):
    """Isolamento (e rivestimento interno dei condotti) del tratto di partenza: [(classe, getter, id tipo, spessore)]."""
    if isinstance(source, Pipe):
        kinds = ((PipeInsulation, DB.InsulationLiningBase.GetInsulationIds),)
    elif isinstance(source, Duct):
        kinds = ((DuctInsulation, DB.InsulationLiningBase.GetInsulationIds),
                 (DuctLining, DB.InsulationLiningBase.GetLiningIds))
    else:
        return []
    specs = []
    for cls, getter in kinds:
        for insulation_id in getter(doc, source.Id):
            insulation = doc.GetElement(insulation_id)
            specs.append((cls, getter, insulation.GetTypeId(), insulation.Thickness))
    return specs


def apply_insulation(specs, hosts):
    """Crea lo stesso isolamento sui nuovi tratti e sui raccordi. Ritorna il numero di elementi isolati."""
    made = 0
    for host in hosts:
        if host is None:
            continue
        for cls, getter, type_id, thickness in specs:
            try:
                if len(list(getter(doc, host.Id))) > 0:
                    continue
                cls.Create(doc, host.Id, type_id, thickness)
                made += 1
            except Exception:
                continue
    return made


def inherit(source, new_elements, fittings):
    """Parametri e isolamento del tratto di partenza sui nuovi elementi. Ritorna una nota per il report."""
    try:
        count = sum(copy_parameters(source, e) for e in new_elements)
        specs = insulation_specs(source)
        text = "{} parameter values copied".format(count)
        if specs:
            text += ", insulation on {} elements".format(
                apply_insulation(specs, list(new_elements) + list(fittings)))
        return text
    except Exception as ex:
        return "parameters and insulation not copied: {}".format(ex)


def check_remaining_space(parts):
    """parts: [(nome, elemento)]. Dopo i raccordi ogni tratto deve avere ancora una lunghezza utile."""
    doc.Regenerate()
    for name, element in parts:
        remaining = element.Location.Curve.Length * MM_PER_FT
        if remaining < MIN_REMAINING_MM:
            raise BranchError("Not enough space for the fitting: the {} would be only {:.0f} mm long "
                              "(minimum {:.0f} mm)".format(name, remaining, MIN_REMAINING_MM))


def flat(vector):
    return DB.XYZ(vector.X, vector.Y, 0).Normalize()


def view_axes():
    """Destra e alto dello schermo nella vista attiva. Una pianta legata a uno scope box ruotato (o con il
    taglio ruotato) si vede ruotata: l'orientamento vero e' quello del riquadro di taglio, se attivo.
    Ritorna (destra, alto, angolo di rotazione rispetto al progetto in gradi, origine dei dati)."""
    view = doc.ActiveView
    if not isinstance(view, DB.ViewPlan):
        return DB.XYZ.BasisX, DB.XYZ.BasisY, 0.0, "project axes (the active view is not a plan)"
    source = "view direction"
    right = flat(view.RightDirection)
    try:
        if view.CropBoxActive:
            crop_right = flat(view.CropBox.Transform.BasisX)
            if crop_right.DotProduct(right) < math.cos(math.radians(0.5)):
                source = "rotated crop region (scope box)"
            right = crop_right
    except Exception:
        pass
    up = DB.XYZ.BasisZ.CrossProduct(right)  # pianta vista dall'alto: alto = Z x destra
    angle = math.degrees(math.atan2(right.Y, right.X))
    return right, up, angle, source


def screen_direction(name, axes):
    """Direzione orizzontale del lato scelto, come si vede sullo schermo (axes da view_axes)."""
    right, up = axes[0], axes[1]
    return {"right": right, "left": right.Negate(), "up": up, "down": up.Negate()}[name]


def rotation_note(axes):
    """Avviso per la finestra e il report quando la vista non e' allineata al progetto, None altrimenti."""
    angle = axes[2]
    if abs(angle) < 0.5 and abs(abs(angle) - 180.0) > 0.5:
        return None
    return ("The active view is rotated by {:.1f} degrees from the project axes ({}). Left, right, up and down "
            "follow the screen, not the project axes.".format(angle, axes[3]))


def param_value(element, bip):
    param = element.get_Parameter(bip)
    return param.AsDouble() if param is not None and param.HasValue and param.AsDouble() > 0 else None


def set_param(element, bip, value):
    param = element.get_Parameter(bip)
    if param is not None and not param.IsReadOnly and value:
        param.Set(value)


def is_round(element):
    if isinstance(element, Pipe):
        return True
    if isinstance(element, Duct):
        return param_value(element, DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM) is None
    return False


def branch_direction(riser, wanted):
    """Direzione dello stacco: libera per le sezioni tonde, agganciata al lato piu' vicino per le rettangolari.
    Ritorna (direzione, asse): asse "w" = lungo la larghezza della sezione, "h" = lungo l'altezza, None = tonda."""
    if is_round(riser):
        return wanted, None
    curve = riser.Location.Curve
    top = max((curve.GetEndPoint(0), curve.GetEndPoint(1)), key=lambda p: p.Z)
    system = nearest_connector(riser, top).CoordinateSystem
    ax, ay = flat(system.BasisX), flat(system.BasisY)  # larghezza lungo BasisX, altezza lungo BasisY
    candidates = ((ax, "w"), (ax.Negate(), "w"), (ay, "h"), (ay.Negate(), "h"))
    direction, axis = max(candidates, key=lambda c: c[0].DotProduct(wanted))
    if isinstance(riser, CableTray) and axis != "w":
        raise BranchError("Cable trays branch only from the narrow side: choose a side along the width of the tray")
    return direction, axis


def system_type_id(riser, bip, cls):
    param = riser.get_Parameter(bip)
    if param is not None and param.AsElementId() != DB.ElementId.InvalidElementId:
        return param.AsElementId()
    first = DB.FilteredElementCollector(doc).OfClass(cls).FirstElementId()
    if first == DB.ElementId.InvalidElementId:
        raise BranchError("No system type found in the model for the new branch")
    return first


def create_branch(riser, start, end, level_id, axis):
    """Stacco da start a end, con tipo e sezione del tratto verticale."""
    type_id = riser.GetTypeId()
    if isinstance(riser, Pipe):
        branch = Pipe.Create(doc, system_type_id(riser, DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM,
                                                 PipingSystemType), type_id, level_id, start, end)
        set_param(branch, DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM,
                  param_value(riser, DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM))
    elif isinstance(riser, Duct):
        branch = Duct.Create(doc, system_type_id(riser, DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM,
                                                 MechanicalSystemType), type_id, level_id, start, end)
        if axis is None:
            set_param(branch, DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM,
                      param_value(riser, DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM))
        else:
            width = param_value(riser, DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
            height = param_value(riser, DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
            # la faccia da cui esce lo stacco: larghezza orizzontale dello stacco = lato della faccia
            along, across = (width, height) if axis == "w" else (height, width)
            set_param(branch, DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM, across)
            set_param(branch, DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM, along)
    else:
        branch = CableTray.Create(doc, type_id, start, end, level_id)
        set_param(branch, DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM,
                  param_value(riser, DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM))
        set_param(branch, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM,
                  param_value(riser, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM))
        doc.Regenerate()
        # passerella di costa: la larghezza deve stare nel piano della T, come quella del tratto verticale
        DB.ElementTransformUtils.RotateElement(doc, branch.Id, DB.Line.CreateBound(start, end), math.pi / 2)
    doc.Regenerate()
    return branch


def break_tray(tray, point):
    """Divide una passerella in point (non esiste BreakCurve per le passerelle). Ritorna l'id del nuovo pezzo."""
    curve = tray.Location.Curve
    a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
    far_conn = nearest_connector(tray, b)
    far_info = [(r.Owner.Id, r.Origin) for r in physical_refs(far_conn, tray)]
    for ref in physical_refs(far_conn, tray):
        far_conn.DisconnectFrom(ref)
    copy_id = list(DB.ElementTransformUtils.CopyElement(doc, tray.Id, DB.XYZ.Zero))[0]
    copy = doc.GetElement(copy_id)
    tray.Location.Curve = DB.Line.CreateBound(a, point)
    copy.Location.Curve = DB.Line.CreateBound(point, b)
    doc.Regenerate()
    conn = nearest_connector(copy, b)
    for owner_id, origin in far_info:
        target = free_connector_near(doc.GetElement(owner_id), origin)
        if target is None:
            raise BranchError("No free connector found on the element connected to the end of the cable tray")
        conn.ConnectTo(target)
    return copy_id


def split_riser(riser, point):
    """Divide il tratto verticale in point. Ritorna (connettore del pezzo originale, connettore del nuovo, nuovo pezzo)."""
    if isinstance(riser, Pipe):
        new_id = PlumbingUtils.BreakCurve(doc, riser.Id, point)
    elif isinstance(riser, Duct):
        new_id = MechanicalUtils.BreakCurve(doc, riser.Id, point)
    else:
        new_id = break_tray(riser, point)
    doc.Regenerate()
    other = doc.GetElement(new_id)
    conn_a, conn_b = nearest_connector(riser, point), nearest_connector(other, point)
    if conn_a.IsConnectedTo(conn_b):
        conn_a.DisconnectFrom(conn_b)
    return conn_a, conn_b, other


def move_end(element, old_end, new_point):
    """Sposta l'estremo old_end del tratto in new_point, mantenendo l'altro estremo e il verso."""
    curve = element.Location.Curve
    a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
    if a.DistanceTo(old_end) < b.DistanceTo(old_end):
        element.Location.Curve = DB.Line.CreateBound(new_point, b)
    else:
        element.Location.Curve = DB.Line.CreateBound(a, new_point)
    doc.Regenerate()


def classify(riser, z):
    """Dove cade la quota: "inside" (lungo il tratto), oppure oltre l'estremo "top" / "bottom"."""
    curve = riser.Location.Curve
    low, high = sorted((curve.GetEndPoint(0), curve.GetEndPoint(1)), key=lambda p: p.Z)
    min_ft = MIN_REMAINING_MM / MM_PER_FT
    if low.Z + min_ft < z < high.Z - min_ft:
        return "inside", low, high
    return ("top" if z >= (low.Z + high.Z) / 2.0 else "bottom"), low, high


def tip_refs(riser, where):
    """Connettori fisici collegati all'estremo oltre cui cade la quota (alto o basso)."""
    curve = riser.Location.Curve
    low, high = sorted((curve.GetEndPoint(0), curve.GetEndPoint(1)), key=lambda p: p.Z)
    tip = high if where == "top" else low
    conn = nearest_connector(riser, tip)
    return conn, physical_refs(conn, riser)


def connection_warnings(risers, z):
    """Testo per la finestra: elementi collegati agli estremi che verrebbero scollegati con questa quota."""
    lines = []
    for riser in risers:
        where = classify(riser, z)[0]
        if where == "inside":
            continue
        for ref in tip_refs(riser, where)[1]:
            owner = ref.Owner
            lines.append("{} {}: {} end connected to {} {} - it will be disconnected".format(
                category_name(riser), riser.Id, where, owner.Category.Name if owner.Category else "element",
                owner.Id))
    if not lines:
        return "No connected element would be disconnected at this elevation."
    return "At this elevation the final configuration has priority over what is already there:\n" + "\n".join(lines)


def build_branch(riser, settings, z, level_id, wanted, beyond_choice):
    """Costruisce lo stacco su un tratto. Ritorna (caso, stacco, nuovi elementi, raccordi, elementi scollegati)."""
    where, low, high = classify(riser, z)
    direction, axis = branch_direction(riser, wanted)
    angle = math.radians(settings["angle_deg"])
    run = direction.Multiply(math.cos(angle)).Add(DB.XYZ.BasisZ.Multiply(math.sin(angle)))
    point = DB.XYZ(low.X, low.Y, z)
    end = point.Add(run.Multiply(settings["length_mm"] / MM_PER_FT))

    disconnected = []
    if where != "inside":
        # le passerelle (come Riser Up/Down) fanno solo curve: mai prolungamento con la T
        if isinstance(riser, CableTray):
            beyond_choice = ELBOW
        if beyond_choice is None:
            raise BranchError("Skipped: the elevation is beyond the end of the riser and no choice was made")
        tip = high if where == "top" else low
        STAGE[0] = "disconnecting the {} end of the riser".format(where)
        # vale la configurazione finale voluta: quanto e' gia' collegato all'estremo viene scollegato
        conn, refs = tip_refs(riser, where)
        for ref in refs:
            disconnected.append("{} {}".format(ref.Owner.Category.Name if ref.Owner.Category else "element",
                                               ref.Owner.Id))
            conn.DisconnectFrom(ref)

    if where != "inside" and beyond_choice == ELBOW:
        STAGE[0] = "extending the riser to the elevation"
        move_end(riser, tip, point)
        STAGE[0] = "creating the branch"
        branch = create_branch(riser, point, end, level_id, axis)
        STAGE[0] = "placing the elbow"
        elbow = doc.Create.NewElbowFitting(nearest_connector(riser, point), nearest_connector(branch, point))
        check_remaining_space([("riser", riser), ("branch", branch)])
        return "Elbow at the {} end".format(where), branch, [branch], [elbow], disconnected

    case = ""
    if where != "inside":
        STAGE[0] = "extending the riser past the elevation"
        sign = 1.0 if where == "top" else -1.0
        move_end(riser, tip, point.Add(DB.XYZ(0, 0, sign * settings["stub_mm"] / MM_PER_FT)))
        case = "Riser extended at the {} end, ".format(where)
    STAGE[0] = "creating the branch"
    branch = create_branch(riser, point, end, level_id, axis)
    STAGE[0] = "placing the junction"
    kind, fitting, others = place_junction(riser, point, branch, settings["junction"])
    parts = [("riser", riser), ("branch", branch)] + [("riser above the junction", o) for o in others]
    check_remaining_space(parts)
    return case + kind, branch, others + [branch], [fitting], disconnected


def junction_preference(riser):
    """Innesto preferito nelle Routing Preferences del tipo: "tee", "tap" o None (passerelle, non leggibile)."""
    try:
        preferred = doc.GetElement(riser.GetTypeId()).RoutingPreferenceManager.PreferredJunctionType
    except Exception:
        return None
    return "tap" if preferred == DB.PreferredJunctionType.Tap else "tee"


def junction_order(riser, mode):
    """Innesti da provare, in ordine. Ritorna (ordine, innesto atteso, origine della scelta)."""
    if isinstance(riser, CableTray):
        return ("tee",), "tee", "cable trays have no taps"
    if mode == "routing":
        preferred = junction_preference(riser)
        # solo qui si prova anche l'altro innesto: la scelta e' lasciata al modello
        return (("tap", "tee") if preferred == "tap" else ("tee", "tap")), preferred, "Routing Preferences"
    if mode == "standard":
        kind = "tap" if isinstance(riser, Duct) else "tee"
        return (kind,), kind, "tap on ducts, tee on pipes"
    return (mode,), mode, "chosen in the window"


def place_junction(riser, point, branch, mode):
    """T (con il tratto diviso) o stacco (tap, senza dividere), secondo la scelta della finestra.
    Ritorna (descrizione, raccordo, nuovi pezzi del tratto)."""
    order, preferred, source = junction_order(riser, mode)
    errors = []
    for kind in order:
        sub = DB.SubTransaction(doc)
        sub.Start()
        try:
            if kind == "tee":
                conn_a, conn_b, other = split_riser(riser, point)
                fitting = doc.Create.NewTeeFitting(conn_a, conn_b, nearest_connector(branch, point))
                others = [other]
            else:
                fitting = doc.Create.NewTakeoffFitting(nearest_connector(branch, point), riser)
                others = []
            sub.Commit()
            label = "Tee" if kind == "tee" else "Tap"
            if preferred is not None and kind != preferred:
                label += " (the preferred {} could not be placed)".format(preferred)
            return label, fitting, others
        except Exception as ex:
            sub.RollBack()
            errors.append("{}: {}".format(kind, ex))
    raise BranchError("No junction could be placed ({}; {}). Check that the {} family is loaded and set for this "
                      "size in the Routing Preferences of the type".format(
                          "; ".join(errors), source, " / ".join(order)))


def process(riser, settings, z, wanted, beyond_choice):
    result = {"element": riser, "category": category_name(riser), "case": "", "branch": None,
              "fittings": 0, "status": ""}
    level_id = get_level_id(riser)
    if level_id is None:
        result["status"] = "Unchanged. No reference level on the riser"
        return result
    transaction = DB.Transaction(doc, "Branch")
    transaction.Start()
    try:
        case, branch, new_elements, fittings, disconnected = build_branch(riser, settings, z, level_id, wanted,
                                                                          beyond_choice)
        STAGE[0] = "copying parameters and insulation"
        note = inherit(riser, new_elements, fittings)
        transaction.Commit()
    except BranchError as ex:
        transaction.RollBack()
        result["status"] = "Unchanged. {}".format(ex)
        return result
    except Exception as ex:
        transaction.RollBack()
        result["status"] = "Unchanged. The branch could not be built (step: {}): {}".format(STAGE[0], ex)
        result["detail"] = traceback.format_exc() if traceback else ""
        return result
    status = "Created ({})".format(note)
    if disconnected:
        status += ". Disconnected from: {}".format(", ".join(disconnected))
    result.update({"case": case, "branch": branch, "fittings": len(fittings), "status": status})
    return result


def get_selection():
    selected = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    risers = [e for e in selected if is_vertical_run(e)]
    if risers:
        return risers
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, VerticalRunFilter(),
                                           "Select the vertical pipes, ducts or cable trays, then press Finish")
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def load_settings(config):
    settings = dict(DEFAULTS)
    try:
        for key in ("direction", "level", "junction"):
            settings[key] = config.get_option(key, DEFAULTS[key])
        for key in ("offset_mm", "length_mm", "angle_deg", "stub_mm"):
            settings[key] = float(config.get_option(key, DEFAULTS[key]))
    except Exception:
        pass
    return settings


def save_settings(config, settings):
    try:
        for key in DEFAULTS:
            setattr(config, key, settings[key])
        script.save_config()
    except Exception:
        pass


def print_report(results, settings, level, note):
    output.print_md("# Branch report")
    output.print_md("Side **{}** (on screen) | level **{}** {:+g} mm | length **{:g} mm** | angle **{:g} degrees** | junction **{}**".format(
        settings["direction"], level.Name, settings["offset_mm"], settings["length_mm"], settings["angle_deg"],
        settings["junction"]))
    if note:
        output.print_md("**View:** {}".format(note))
    done =[r for r in results if r["status"].startswith("Created")]
    output.print_md("Risers: **{}** | branches created: **{}** | left unchanged: **{}**".format(
        len(results), len(done), len(results) - len(done)))
    rows = [[i, output.linkify(r["element"].Id), r["category"], r["case"],
             output.linkify(r["branch"].Id) if r["branch"] is not None else "", r["fittings"] or "", r["status"]]
            for i, r in enumerate(results, 1)]
    output.print_table(table_data=rows, columns=["#", "Riser", "Category", "Case", "Branch", "Fittings", "Status"])
    for i, r in enumerate(results, 1):
        if r.get("detail"):
            output.print_md("**Technical detail for riser #{}**".format(i))
            output.print_code(r["detail"])


def main():
    risers = get_selection()
    if not risers:
        script.exit()

    levels = sorted(DB.FilteredElementCollector(doc).OfClass(DB.Level), key=level_elevation)
    names = [lv.Name for lv in levels]
    config = script.get_config(CONFIG_NAME)
    settings = load_settings(config)
    if settings["level"] not in names:
        own = doc.GetElement(get_level_id(risers[0])) if get_level_id(risers[0]) is not None else None
        settings["level"] = own.Name if own is not None else (names[0] if names else "")
    axes = view_axes()
    note = rotation_note(axes)

    def warnings(level_name, offset_mm):
        z_ft = level_elevation(levels[names.index(level_name)]) + offset_mm / MM_PER_FT
        return connection_warnings(risers, z_ft)

    settings = Branch_ui.BranchWindow(script.get_bundle_file(XAML_FILE_NAME), len(risers), names, settings,
                                      note, warnings).show()
    if settings is None:
        script.exit()
    save_settings(config, settings)

    level = levels[names.index(settings["level"])]
    z = level_elevation(level) + settings["offset_mm"] / MM_PER_FT
    wanted = screen_direction(settings["direction"], axes)

    # le passerelle fanno sempre la curva: la scelta riguarda solo tubi e canali
    beyond = [r for r in risers if classify(r, z)[0] != "inside" and not isinstance(r, CableTray)]
    beyond_choice = None
    if beyond:
        beyond_choice = forms.CommandSwitchWindow.show(
            [ELBOW, EXTEND_TEE],
            message="The elevation is above or below the end of {} of the {} risers (pipes and ducts). "
                    "Turn there with an elbow, or extend the riser and place a tee? "
                    "Cable trays always turn with a curve.".format(len(beyond), len(risers)))
        if beyond_choice is None:
            script.exit()

    results = [process(r, settings, z, wanted, beyond_choice) for r in risers]
    problems = len([r for r in results if not r["status"].startswith("Created")])
    if want_report(len(results), problems, "branches created"):
        print_report(results, settings, level, note)


main()
