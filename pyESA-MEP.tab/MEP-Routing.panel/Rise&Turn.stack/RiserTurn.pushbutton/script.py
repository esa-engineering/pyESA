# -*- coding: utf-8 -*-
__title__ = "Riser\nTurn"
__doc__ = """Version = 0.2
Date    = 12.10.2026
________________________________________________________________
Description:
Turns vertical runs of pipe, duct or cable tray into a horizontal run,
with an elbow.

Instruction:
1. Select the vertical runs (or select them first) and press Finish
2. In the window choose:
   - where the elbow goes: at the top end, at the bottom end, or at an
     elevation (level + offset). With an elevation you choose which end
     moves there: the riser is extended or shortened up to it
   - the side of the horizontal run, as seen on screen in the current
     plan view, or Custom: after OK you click a direction in the view
     (one direction for all the risers, from the first one). Pipes snap
     to the nearest multiple of 45 degrees from the view axes;
     rectangular ducts and cable trays snap to the nearest side
   - the length of the horizontal run, from the riser axis
3. Cable trays: from a narrow side the horizontal run stays on edge,
   from a wide side it comes out flat. Rectangular ducts and cable trays
   keep exactly the section of the riser through the elbow
4. Staggering, computed in plan and independent of the view: each
   horizontal run is a corridor along its real direction (after snapping
   to a face), as wide as the riser across that direction. If it crosses
   another selected riser farther ahead, the run passes the elbow of that
   riser by their distance in plan along the run (at least enough for
   the sections not to touch): above it when the elbow is at the top
   (the riser is below the elbow), below it when the elbow is at the
   bottom (the riser rises above the elbow). The riser farthest in the
   direction of the turn keeps the elevation. The moved end is the one
   where the elbow goes

If the end where the elbow goes is already connected, or the elevation
leaves the riser too short, that riser is left unchanged and the report
says why. The new parts keep the type, size, parameters and insulation
of the riser.
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
from Autodesk.Revit.DB.Plumbing import Pipe, PipeInsulation, PipingSystemType
from Autodesk.Revit.DB.Mechanical import Duct, DuctInsulation, DuctLining, MechanicalSystemType
from Autodesk.Revit.DB.Electrical import CableTray
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, DB, forms

import RiserTurn_ui

XAML_FILE_NAME = "RiserTurn_form.xaml"
CONFIG_NAME = "ESA_RiserTurn"
DEFAULTS = {"direction": "right", "position": "top", "level": "", "offset_mm": 3000.0, "move_end": "top",
            "length_mm": 1000.0}

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
MIN_REMAINING_MM = 20.0  # lunghezza minima di ogni tratto dopo l'inserimento della curva
STAGE = [""]             # passo in corso, per i messaggi di errore
PHYSICAL_TYPES = (DB.ConnectorType.End, DB.ConnectorType.Curve, DB.ConnectorType.Physical)
SNAP_DEG = 45.0          # passo dell'aggancio della direzione cliccata per i tubi

CATEGORY_NAMES = ((Pipe, "Pipe"), (Duct, "Duct"), (CableTray, "Cable Tray"))


class TurnError(Exception):
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


def connector_manager(element):
    """ConnectorManager di un tratto MEP o di un raccordo (FamilyInstance)."""
    manager = getattr(element, "ConnectorManager", None)
    if manager is None:
        model = getattr(element, "MEPModel", None)
        manager = model.ConnectorManager if model is not None else None
    return manager


def nearest_connector(element, point):
    best, best_dist = None, None
    for c in connector_manager(element).Connectors:
        if c.ConnectorType not in PHYSICAL_TYPES:
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
    """parts: [(nome, elemento)]. Dopo la curva ogni tratto deve avere ancora una lunghezza utile."""
    doc.Regenerate()
    for name, element in parts:
        remaining = element.Location.Curve.Length * MM_PER_FT
        if remaining < MIN_REMAINING_MM:
            raise TurnError("Not enough space for the elbow: the {} would be only {:.0f} mm long "
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


def tray_width_direction(tray):
    """Direzione orizzontale della larghezza di una passerella verticale. Il verso BasisX/BasisY del connettore non
    e' affidabile per le passerelle: si confrontano le estensioni della geometria con larghezza e altezza."""
    curve = tray.Location.Curve
    top = max((curve.GetEndPoint(0), curve.GetEndPoint(1)), key=lambda p: p.Z)
    system = nearest_connector(tray, top).CoordinateSystem
    ax, ay = flat(system.BasisX), flat(system.BasisY)
    width = param_value(tray, DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
    height = param_value(tray, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
    if not width or not height or abs(width - height) < 1.0 / MM_PER_FT:
        return ax  # sezione quasi quadrata: nessun modo di distinguere, resta l'ipotesi storica
    try:
        lows, highs = [1e9, 1e9], [-1e9, -1e9]
        pending = list(tray.get_Geometry(DB.Options()))
        while pending:
            item = pending.pop()
            if isinstance(item, DB.GeometryInstance):
                pending.extend(item.GetInstanceGeometry())
            elif isinstance(item, DB.Solid):
                for edge in item.Edges:
                    for end in (0, 1):
                        p = edge.AsCurve().GetEndPoint(end)
                        for i, axis in enumerate((ax, ay)):
                            v = p.X * axis.X + p.Y * axis.Y
                            lows[i], highs[i] = min(lows[i], v), max(highs[i], v)
        ext_x, ext_y = highs[0] - lows[0], highs[1] - lows[1]
        if ext_x < 0 or ext_y < 0:
            return ax
        keep = abs(ext_x - width) + abs(ext_y - height)
        swap = abs(ext_x - height) + abs(ext_y - width)
        return ay if swap < keep else ax
    except Exception:
        return ax


def run_direction(riser, wanted):
    """Direzione del tratto orizzontale: libera per le sezioni tonde, agganciata al lato piu' vicino per le
    rettangolari. Passerelle: lato stretto ("w") = tratto di costa, lato largo ("h") = tratto piatto.
    Ritorna (direzione, asse): "w" = lungo la larghezza, "h" = lungo l'altezza, None = tonda."""
    if is_round(riser):
        return wanted, None
    if isinstance(riser, CableTray):
        width_dir = tray_width_direction(riser)
        across = DB.XYZ.BasisZ.CrossProduct(width_dir)
        candidates = ((width_dir, "w"), (width_dir.Negate(), "w"), (across, "h"), (across.Negate(), "h"))
        return max(candidates, key=lambda c: c[0].DotProduct(wanted))
    curve = riser.Location.Curve
    top = max((curve.GetEndPoint(0), curve.GetEndPoint(1)), key=lambda p: p.Z)
    system = nearest_connector(riser, top).CoordinateSystem
    ax, ay = flat(system.BasisX), flat(system.BasisY)  # larghezza lungo BasisX, altezza lungo BasisY
    candidates = ((ax, "w"), (ax.Negate(), "w"), (ay, "h"), (ay.Negate(), "h"))
    return max(candidates, key=lambda c: c[0].DotProduct(wanted))


def system_type_id(riser, bip, cls):
    param = riser.get_Parameter(bip)
    if param is not None and param.AsElementId() != DB.ElementId.InvalidElementId:
        return param.AsElementId()
    first = DB.FilteredElementCollector(doc).OfClass(cls).FirstElementId()
    if first == DB.ElementId.InvalidElementId:
        raise TurnError("No system type found in the model for the new run")
    return first


def rotated_copy(riser, start, end):
    """Tratto orizzontale da start a end con la sezione del montante ribaltata rigidamente: si copia il montante
    (verticale, sezione gia' orientata), lo si accorcia a un tratto verticale lungo quanto il nuovo tratto e lo si
    ribalta di 90 gradi attorno a un asse orizzontale. La sezione arriva alla curva esattamente come esce dal
    montante: niente scambi larghezza/altezza da indovinare (canali rettangolari) e passerelle di costa possibili
    (Revit non ruota una passerella orizzontale attorno al proprio asse, "cannot be rotated in this manner")."""
    run = end.Subtract(start)
    length = run.GetLength()
    direction = run.Normalize()
    copy_id = list(DB.ElementTransformUtils.CopyElement(doc, riser.Id, DB.XYZ.Zero))[0]
    branch = doc.GetElement(copy_id)
    branch.Location.Curve = DB.Line.CreateBound(start, start.Add(DB.XYZ.BasisZ.Multiply(length)))
    doc.Regenerate()
    # rotazione di +90 gradi attorno a Z x direzione: porta Z su direzione (Rodrigues)
    pivot = DB.Line.CreateBound(start, start.Add(DB.XYZ.BasisZ.CrossProduct(direction)))
    DB.ElementTransformUtils.RotateElement(doc, copy_id, pivot, math.pi / 2)
    doc.Regenerate()
    mark = branch.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
    if mark is not None and not mark.IsReadOnly:
        mark.Set("")  # la copia eredita il Mark del montante: duplicato
    curve = branch.Location.Curve
    if min(curve.GetEndPoint(0).DistanceTo(end), curve.GetEndPoint(1).DistanceTo(end)) > 1.0 / MM_PER_FT:
        raise TurnError("The horizontal run did not end up in the expected position")
    return branch


def create_run(riser, start, end, level_id, axis):
    """Tratto orizzontale da start a end, con tipo e sezione del montante. Sezioni tonde: nuovo tratto con lo stesso
    diametro. Canali rettangolari e passerelle (entrambi i lati): copia ribaltata del montante, cosi' la sezione
    combacia con la curva (dal lato largo una passerella esce piatta, dal lato stretto di costa)."""
    if axis is not None:
        return rotated_copy(riser, start, end)
    type_id = riser.GetTypeId()
    if isinstance(riser, Pipe):
        new = Pipe.Create(doc, system_type_id(riser, DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM,
                                              PipingSystemType), type_id, level_id, start, end)
        set_param(new, DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM,
                  param_value(riser, DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM))
    else:
        new = Duct.Create(doc, system_type_id(riser, DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM,
                                              MechanicalSystemType), type_id, level_id, start, end)
        set_param(new, DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM,
                  param_value(riser, DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM))
    doc.Regenerate()
    return new


def move_end(element, old_end, new_point):
    """Sposta l'estremo old_end del tratto in new_point, mantenendo l'altro estremo e il verso."""
    curve = element.Location.Curve
    a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
    if a.DistanceTo(old_end) < b.DistanceTo(old_end):
        element.Location.Curve = DB.Line.CreateBound(new_point, b)
    else:
        element.Location.Curve = DB.Line.CreateBound(a, new_point)
    doc.Regenerate()


def ends(riser):
    curve = riser.Location.Curve
    return sorted((curve.GetEndPoint(0), curve.GetEndPoint(1)), key=lambda p: p.Z)


def owner_label(ref):
    owner = ref.Owner
    return "{} {}".format(owner.Category.Name if owner.Category else "element", owner.Id)


def build_turn(riser, plan, wanted, length_mm, level_id):
    """Curva e tratto orizzontale su un montante. plan: dict con "which" (estremo, "top"/"bottom") e "z" (quota
    della curva, None = all'estremo com'e'). Ritorna (caso, nuovo tratto, curva)."""
    low, high = ends(riser)
    which, z = plan["which"], plan["z"]
    tip = high if which == "top" else low

    # l'estremo dove va la curva deve essere libero: niente modifiche a cio' che esiste gia'
    STAGE[0] = "checking the {} end".format(which)
    refs = physical_refs(nearest_connector(riser, tip), riser)
    if refs:
        raise TurnError("The {} end is connected to {}: nothing changed. Disconnect it or choose the other "
                        "end".format(which, ", ".join(owner_label(r) for r in refs)))

    case = "Elbow at the {} end".format(which)
    if z is not None:
        min_ft = MIN_REMAINING_MM / MM_PER_FT
        if which == "top" and z < low.Z + min_ft:
            raise TurnError("The elevation is below the bottom end of the riser (or too close to it): moving the "
                            "top end there is not possible")
        if which == "bottom" and z > high.Z - min_ft:
            raise TurnError("The elevation is above the top end of the riser (or too close to it): moving the "
                            "bottom end there is not possible")
        STAGE[0] = "moving the {} end to the elevation".format(which)
        delta_mm = (z - tip.Z) * MM_PER_FT
        new_tip = DB.XYZ(tip.X, tip.Y, z)
        move_end(riser, tip, new_tip)
        tip = new_tip
        grows = (delta_mm > 0) == (which == "top")
        if abs(delta_mm) >= 0.5:
            case = "Elbow at {:.0f} mm, riser {} by {:.0f} mm at the {} end".format(
                z * MM_PER_FT, "extended" if grows else "shortened", abs(delta_mm), which)

    if plan.get("staggered"):
        case += " ({})".format(plan["staggered"])
    direction, axis = run_direction(riser, wanted)
    end = tip.Add(direction.Multiply(length_mm / MM_PER_FT))
    STAGE[0] = "creating the horizontal run"
    new = create_run(riser, tip, end, level_id, axis)
    STAGE[0] = "placing the elbow"
    elbow = doc.Create.NewElbowFitting(nearest_connector(riser, tip), nearest_connector(new, tip))
    check_remaining_space([("riser", riser), ("horizontal run", new)])
    if isinstance(riser, CableTray):
        case += ", on edge" if axis == "w" else ", flat"
    return case, new, elbow


class RevitErrors(DB.IFailuresPreprocessor):
    """Errori di Revit al commit: invece della finestra bloccante, la transazione viene annullata e i messaggi vanno
    nel report. Gli avvisi passano come sempre."""

    def __init__(self):
        self.messages = []

    def PreprocessFailures(self, accessor):
        for failure in accessor.GetFailureMessages():
            if failure.GetSeverity() != DB.FailureSeverity.Warning:
                self.messages.append(failure.GetDescriptionText())
        if self.messages:
            return DB.FailureProcessingResult.ProceedWithRollBack
        return DB.FailureProcessingResult.Continue


def fitting_label(fitting):
    try:
        return "{}: {}".format(fitting.Symbol.FamilyName, fitting.Name)
    except Exception:
        return "elbow"


def plan_points(riser):
    """Punti della geometria del montante (vertici degli spigoli dei solidi), per l'ingombro in pianta.
    Se la geometria non e' leggibile: i quattro punti della sezione attorno all'asse, con il lato maggiore."""
    points = []
    try:
        pending = list(riser.get_Geometry(DB.Options()))
        while pending:
            item = pending.pop()
            if isinstance(item, DB.GeometryInstance):
                pending.extend(item.GetInstanceGeometry())
            elif isinstance(item, DB.Solid):
                for edge in item.Edges:
                    curve = edge.AsCurve()
                    points.append(curve.GetEndPoint(0))
                    points.append(curve.GetEndPoint(1))
    except Exception:
        points = []
    if points:
        return points
    values = [param_value(riser, bip) for bip in (
        DB.BuiltInParameter.RBS_PIPE_OUTER_DIAMETER, DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM,
        DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM, DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM,
        DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)]
    half = max([v for v in values if v] or [0.0]) / 2.0
    c = ends(riser)[0]
    return [DB.XYZ(c.X + dx, c.Y + dy, c.Z) for dx in (-half, half) for dy in (-half, half)]


def extent(points, axis):
    """Intervallo (min, max) dei punti proiettati sull'asse orizzontale."""
    values = [p.X * axis.X + p.Y * axis.Y for p in points]
    return min(values), max(values)


def plan_elevations(risers, settings, z, wanted):
    """Estremo e quota della curva di ogni montante, con lo sfalsamento per evitare le interferenze.
    Geometria in pianta, indipendente dalla vista: ogni tratto orizzontale e' un corridoio che parte dall'asse del
    suo montante, nella direzione reale del tratto (quella agganciata alla faccia) e largo quanto l'ingombro vero
    del montante di traverso a quella direzione. Se il corridoio attraversa un altro montante che sta piu' avanti,
    alla stessa quota il tratto lo colpirebbe: il tratto passa oltre la curva di quel montante, della distanza in
    pianta tra i due lungo il tratto (almeno quanto serve perche' le sezioni non si tocchino). Sopra se la curva e'
    in alto (il montante sta sotto), sotto se la curva e' in basso (il montante sale sopra la curva). Il montante
    piu' avanti nella direzione della curva tiene quindi la quota base. Ritorna {id del montante: plan}."""
    which = settings["move_end"] if settings["position"] == "elevation" else settings["position"]
    length = settings["length_mm"] / MM_PER_FT
    info = {}
    for riser in risers:
        low, high = ends(riser)
        tip = high if which == "top" else low
        direction = flat(run_direction(riser, wanted)[0])  # direzione reale, dopo l'aggancio alla faccia
        lateral = DB.XYZ.BasisZ.CrossProduct(direction)
        points = plan_points(riser)
        own_along = extent(points, direction)
        info[get_id(riser)] = {"riser": riser, "tip": tip, "dir": direction, "lat": lateral, "points": points,
                               "lat_range": extent(points, lateral), "along_range": own_along,
                               # altezza del tratto orizzontale = ingombro del montante lungo il tratto
                               "height": own_along[1] - own_along[0],
                               "base": z if z is not None else tip.Z}

    def hits(i):
        """Montanti attraversati dal corridoio del tratto di i: [(id, distanza lungo il tratto)]."""
        tip_along = i["tip"].X * i["dir"].X + i["tip"].Y * i["dir"].Y
        found = []
        for key, j in info.items():
            if j is i:
                continue
            j_along = extent(j["points"], i["dir"])
            j_lat = extent(j["points"], i["lat"])
            ahead = j_along[1] > i["along_range"][1] and j_along[0] < tip_along + length
            side_by_side = j_lat[0] < i["lat_range"][1] and j_lat[1] > i["lat_range"][0]
            if ahead and side_by_side:
                j_center = (j_along[0] + j_along[1]) / 2.0
                found.append((key, j_center - tip_along))
        return found

    elevations, notes, visiting = {}, {}, set()
    # curva in alto: il montante sta sotto la curva, il tratto passa SOPRA la curva dell'altro montante.
    # Curva in basso: il montante sale sopra la curva, quindi il tratto deve passare SOTTO, altrimenti attraversa
    # il montante dell'altro
    sign = 1.0 if which == "top" else -1.0

    def elevation(key):
        if key in elevations:
            return elevations[key]
        i = info[key]
        value = i["base"]
        if key not in visiting:  # direzioni diverse potrebbero creare un ciclo: si interrompe
            visiting.add(key)
            for other, distance in hits(i):
                step = max(distance, (i["height"] + info[other]["height"]) / 2.0)
                needed = elevation(other) + sign * step
                if sign * (needed - value) > 1e-6:
                    value = needed
                    notes[key] = "{} {:+.0f} mm to pass {} riser {}".format(
                        "raised" if sign > 0 else "lowered", (value - i["base"]) * MM_PER_FT,
                        "over" if sign > 0 else "under", other)
            visiting.discard(key)
        elevations[key] = value
        return value

    plans = {}
    for key, i in info.items():
        value = elevation(key)
        moved = key in notes
        plans[key] = {"which": which, "z": value if (z is not None or moved) else None,
                      "staggered": notes.get(key, "")}
    return plans


def get_id(element):
    eid = element.Id
    return eid.Value if hasattr(eid, "Value") else eid.IntegerValue


def process(riser, plan, wanted, length_mm):
    result = {"element": riser, "category": category_name(riser), "case": "", "new": None, "fitting": "",
              "status": ""}
    level_id = get_level_id(riser)
    if level_id is None:
        result["status"] = "Unchanged. No reference level on the riser"
        return result
    transaction = DB.Transaction(doc, "Riser Turn")
    errors = RevitErrors()
    options = transaction.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(errors)
    options.SetClearAfterRollback(True)
    transaction.SetFailureHandlingOptions(options)
    transaction.Start()
    try:
        case, new, elbow = build_turn(riser, plan, wanted, length_mm, level_id)
        STAGE[0] = "copying parameters and insulation"
        note = inherit(riser, [new], [elbow])
        STAGE[0] = "committing"
        if transaction.Commit() != DB.TransactionStatus.Committed:
            result["status"] = "Unchanged. Revit refused the result: {}".format(
                "; ".join(errors.messages) or "the transaction was rolled back")
            return result
    except TurnError as ex:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        result["status"] = "Unchanged. {}".format(ex)
        return result
    except Exception as ex:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        result["status"] = "Unchanged. The turn could not be built (step: {}): {}".format(STAGE[0], ex)
        result["detail"] = traceback.format_exc() if traceback else ""
        return result
    result.update({"case": case, "new": new, "fitting": fitting_label(elbow),
                   "status": "Created ({})".format(note)})
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


def pick_direction(origin, axes, snap_45):
    """Direzione orizzontale dall'asse del montante al punto cliccato. snap_45: agganciata al multiplo di 45 gradi
    piu' vicino rispetto agli assi della vista (tubi). Ritorna (direzione, angolo in gradi rispetto alla destra)."""
    try:
        picked = uidoc.Selection.PickPoint("Click in the direction of the horizontal run")
    except OperationCanceledException:
        script.exit()
    vector = DB.XYZ(picked.X - origin.X, picked.Y - origin.Y, 0)
    if vector.GetLength() < 1e-6:
        forms.alert("The clicked point is on the axis of the riser: no direction. Run the command again and click "
                    "away from the riser.", exitscript=True)
    right, up = axes[0], axes[1]
    angle = math.degrees(math.atan2(vector.DotProduct(up), vector.DotProduct(right)))
    if snap_45:
        angle = round(angle / SNAP_DEG) * SNAP_DEG
    rad = math.radians(angle)
    return right.Multiply(math.cos(rad)).Add(up.Multiply(math.sin(rad))).Normalize(), angle


def load_settings(config):
    settings = dict(DEFAULTS)
    try:
        for key in ("direction", "position", "level", "move_end"):
            settings[key] = config.get_option(key, DEFAULTS[key])
        for key in ("offset_mm", "length_mm"):
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


def position_text(settings, level):
    if settings["position"] != "elevation":
        return "at the {} end".format(settings["position"])
    return "at level {} {:+g} mm, moving the {} end".format(level.Name, settings["offset_mm"], settings["move_end"])


def print_report(results, settings, level, direction_text, note):
    output.print_md("# Riser Turn report")
    output.print_md("Elbow **{}** | side **{}** | horizontal run **{:g} mm** from the riser axis".format(
        position_text(settings, level), direction_text, settings["length_mm"]))
    if note:
        output.print_md("**View:** {}".format(note))
    done = [r for r in results if r["status"].startswith("Created")]
    output.print_md("Risers: **{}** | turned: **{}** | left unchanged: **{}**".format(
        len(results), len(done), len(results) - len(done)))
    rows = [[i, output.linkify(r["element"].Id), r["category"], r["case"],
             output.linkify(r["new"].Id) if r["new"] is not None else "", r["fitting"], r["status"]]
            for i, r in enumerate(results, 1)]
    output.print_table(table_data=rows,
                       columns=["#", "Riser", "Category", "Case", "Horizontal run", "Elbow", "Status"])
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

    xaml_path = script.get_bundle_file(XAML_FILE_NAME)
    settings = RiserTurn_ui.RiserTurnWindow(xaml_path, len(risers), names, settings, note).show()
    if settings is None:
        script.exit()
    save_settings(config, settings)

    level = levels[names.index(settings["level"])] if settings["level"] in names else None
    z = None
    if settings["position"] == "elevation":
        z = level_elevation(level) + settings["offset_mm"] / MM_PER_FT

    if settings["direction"] == "custom":
        # i tubi si agganciano ai 45 gradi; canali rettangolari e passerelle si agganciano comunque a un lato
        wanted, angle = pick_direction(ends(risers[0])[0], axes, snap_45=True)
        direction_text = "clicked, {:g} degrees from screen right".format(angle)
    else:
        wanted = screen_direction(settings["direction"], axes)
        direction_text = "{} (on screen)".format(settings["direction"])

    plans = plan_elevations(risers, settings, z, wanted)
    results = [process(r, plans[get_id(r)], wanted, settings["length_mm"]) for r in risers]
    problems = len([r for r in results if not r["status"].startswith("Created")])
    if want_report(len(results), problems, "risers turned"):
        print_report(results, settings, level, direction_text, note)


main()
