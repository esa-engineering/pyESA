# -*- coding: utf-8 -*-
__title__ = "Horizontal\nBranch"
__doc__ = """Version = 0.6
Date    = 12.10.2026
________________________________________________________________
Description:
Creates branches at 90 degrees from horizontal runs of pipe, duct or
cable tray (also sloped up to 60 degrees), click by click. Branches
from vertical runs are made by Vertical Branch.

- click the point of a run where the branch starts, then click on one of
  its sides: a branch at 90 degrees, horizontal, is created there with a
  tee (the run is split) or a tap, as saved for the category in
  Vertical Branch (tap on ducts and tee on pipes by default; cable trays
  always get a tee and keep the orientation of their section)
- keep clicking to add more branches, also on the same run; ESC to finish
- pipe branches keep the size of the run
- optional change of level: a short stub with the tee or tap, then an inclined run at the
  chosen angle (90, 45, 30 or custom) up or down to the offset (+ up, - down), then the
  horizontal run of the set length after the change of level. Ducts and cable trays keep
  the tee on the stub before the change of level. Set in SHIFT + CLICK; 0 = straight branch
- optional return to the starting level: after the horizontal run, a second inclined run
  (same angle, opposite direction) brings the branch back to the level it started from, then a last horizontal
  run of the same length
- length of the horizontal run (after the change of level, or the whole branch with 0 offset),
  from the run axis / elbow: the saved one (default 500 mm)
- SHIFT + CLICK opens one window with the offset, the angle and the length
If a fitting does not fit, the run is left unchanged and the report
says why.
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
from System.Collections.Generic import List
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB.Plumbing import Pipe, PipeInsulation, PipeSegment, PipingSystemType, PlumbingUtils
from Autodesk.Revit.DB.Mechanical import Duct, DuctInsulation, DuctLining, MechanicalSystemType, MechanicalUtils
from Autodesk.Revit.DB.Electrical import CableTray
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, DB, forms

import Branch_ui
import Horizontal_ui

HORIZONTAL_XAML = "Horizontal_form.xaml"

XAML_FILES = {"duct": "Branch_form.xaml", "pipe": "BranchPipe_form.xaml", "tray": "BranchTray_form.xaml"}
KIND_LABELS = {"duct": "Ducts", "pipe": "Pipes", "tray": "Cable trays"}
CONFIG_PREFIX = {"duct": "", "pipe": "pipe_", "tray": "tray_"}  # i condotti mantengono le chiavi storiche
CONFIG_NAME = "ESA_Branch"
DEFAULTS = {"direction": "right", "level": "", "offset_mm": 3000.0, "length_mm": 500.0,
            "stub_mm": 300.0, "junction": "standard", "reduce": False}  # reduce: solo tubi, stacco di una misura in meno

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


def is_vertical_pipe(element):
    return isinstance(element, Pipe) and is_vertical_run(element)


class VerticalPipeFilter(VerticalRunFilter):
    """Solo tubi verticali (stacco con braga)."""

    def AllowElement(self, element):
        try:
            return is_vertical_pipe(element)
        except Exception:
            return False


def category_name(element):
    for cls, name in CATEGORY_NAMES:
        if isinstance(element, cls):
            return name
    return "?"


def kind_of(element):
    """Categoria per la finestra: "pipe", "duct" o "tray"."""
    if isinstance(element, Pipe):
        return "pipe"
    return "tray" if isinstance(element, CableTray) else "duct"


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


def screen_name(vector, axes):
    """Nome del lato dello schermo (up, down, left, right) piu' vicino al vettore."""
    return max(("up", "down", "left", "right"), key=lambda n: screen_direction(n, axes).DotProduct(vector))


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


def branch_direction(riser, wanted):
    """Direzione dello stacco: libera per le sezioni tonde, agganciata al lato piu' vicino per le rettangolari.
    Passerelle: lato stretto ("w", lungo la larghezza) = T con ramo di costa; lato largo ("h") = ramo piatto non
    collegato. Ritorna (direzione, asse): "w" = lungo la larghezza, "h" = lungo l'altezza, None = tonda."""
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


def tray_sides_text(trays, axes):
    """Per la finestra: lati dello schermo da cui puo' uscire lo stacco di ogni passerella."""
    lines = []
    for tray in trays:
        width_dir = tray_width_direction(tray)
        across = DB.XYZ.BasisZ.CrossProduct(width_dir)
        lines.append("Cable tray {}: narrow sides (tee, on edge) {} / {}; wide sides (flat, not connected) "
                     "{} / {}".format(tray.Id, screen_name(width_dir, axes), screen_name(width_dir.Negate(), axes),
                                      screen_name(across, axes), screen_name(across.Negate(), axes)))
    return "Sides of each tray, as seen on screen:\n" + "\n".join(lines)


def system_type_id(riser, bip, cls):
    param = riser.get_Parameter(bip)
    if param is not None and param.AsElementId() != DB.ElementId.InvalidElementId:
        return param.AsElementId()
    first = DB.FilteredElementCollector(doc).OfClass(cls).FirstElementId()
    if first == DB.ElementId.InvalidElementId:
        raise BranchError("No system type found in the model for the new branch")
    return first


def tray_on_edge(riser, start, end):
    """Passerella orizzontale di costa da start a end. Revit non ruota una passerella orizzontale attorno al proprio
    asse ("cannot be rotated in this manner"): si copia il montante (verticale, sezione gia' orientata), lo si
    accorcia a un tratto verticale lungo quanto lo stacco e lo si ribalta di 90 gradi attorno a un asse orizzontale.
    La larghezza, orizzontale sul montante, finisce in verticale."""
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
    if curve.GetEndPoint(1).DistanceTo(end) > 1.0 / MM_PER_FT and curve.GetEndPoint(0).DistanceTo(end) > 1.0 / MM_PER_FT:
        raise BranchError("The on-edge branch did not end up in the expected position")
    return branch


def tray_turned_branch(run, start, end):
    """Stacco a 90 gradi da una passerella orizzontale o inclinata, nel piano della T: si copia il tratto, lo si
    mette lungo la propria direzione a partire da start e lo si ruota di 90 gradi attorno all'asse perpendicolare
    a tratto e stacco. La sezione resta orientata come quella del tratto (piatta, di costa o inclinata)."""
    branch_run = end.Subtract(start)
    length = branch_run.GetLength()
    branch_dir = branch_run.Normalize()
    run_dir = run.Location.Curve.Direction
    copy_id = list(DB.ElementTransformUtils.CopyElement(doc, run.Id, DB.XYZ.Zero))[0]
    branch = doc.GetElement(copy_id)
    branch.Location.Curve = DB.Line.CreateBound(start, start.Add(run_dir.Multiply(length)))
    doc.Regenerate()
    # +90 gradi attorno a run x branch porta run_dir su branch_dir (Rodrigues, vettori ortogonali)
    normal = run_dir.CrossProduct(branch_dir)
    DB.ElementTransformUtils.RotateElement(doc, copy_id, DB.Line.CreateBound(start, start.Add(normal)),
                                           math.pi / 2)
    doc.Regenerate()
    mark = branch.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
    if mark is not None and not mark.IsReadOnly:
        mark.Set("")  # la copia eredita il Mark del tratto: duplicato
    curve = branch.Location.Curve
    if min(curve.GetEndPoint(0).DistanceTo(end), curve.GetEndPoint(1).DistanceTo(end)) > 1.0 / MM_PER_FT:
        raise BranchError("The branch did not end up in the expected position")
    return branch


def smaller_pipe_size(pipe):
    """Diametro nominale immediatamente inferiore a quello del tubo, tra le misure del suo segmento (Routing
    Preferences del tipo) abilitate negli elenchi delle misure. Ritorna (diametro, nota per il report);
    diametro None se il tubo e' gia' alla misura minima o il segmento non e' leggibile."""
    current = param_value(pipe, DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
    try:
        segment = pipe.PipeSegment
        sizes = [s.NominalDiameter for s in segment.GetSizes() if s.UsedInSizeLists]
    except Exception as ex:
        return None, "same size: the pipe segment could not be read ({})".format(ex)
    smaller = [d for d in sizes if current is not None and d < current - 1e-6]
    if not smaller:
        return None, "same size: DN {:g} is the smallest size of the segment".format((current or 0) * MM_PER_FT)
    new = max(smaller)
    return new, "branch reduced DN {:g} to DN {:g}".format(current * MM_PER_FT, new * MM_PER_FT)


def create_branch(riser, start, end, level_id, axis, diameter=None):
    """Stacco da start a end, con tipo e sezione del tratto di partenza. diameter: solo tubi, diametro nominale
    dello stacco se diverso da quello del tratto (stacco ridotto)."""
    type_id = riser.GetTypeId()
    if isinstance(riser, Pipe):
        branch = Pipe.Create(doc, system_type_id(riser, DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM,
                                                 PipingSystemType), type_id, level_id, start, end)
        set_param(branch, DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM,
                  diameter or param_value(riser, DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM))
    elif isinstance(riser, Duct):
        branch = Duct.Create(doc, system_type_id(riser, DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM,
                                                 MechanicalSystemType), type_id, level_id, start, end)
        if axis is None or (axis == "same" and is_round(riser)):
            set_param(branch, DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM,
                      param_value(riser, DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM))
        elif axis == "same":
            # stacco da un tratto orizzontale: stessa larghezza e altezza del tratto
            set_param(branch, DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM,
                      param_value(riser, DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM))
            set_param(branch, DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM,
                      param_value(riser, DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM))
        else:
            width = param_value(riser, DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
            height = param_value(riser, DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
            # la faccia da cui esce lo stacco: larghezza orizzontale dello stacco = lato della faccia
            along, across = (width, height) if axis == "w" else (height, width)
            set_param(branch, DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM, across)
            set_param(branch, DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM, along)
    elif axis == "w":
        branch = tray_on_edge(riser, start, end)
    elif axis == "same":
        # stacco da una passerella orizzontale o inclinata: CableTray.Create la farebbe sempre piatta
        branch = tray_turned_branch(riser, start, end)
    else:
        branch = CableTray.Create(doc, type_id, start, end, level_id)
        set_param(branch, DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM,
                  param_value(riser, DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM))
        set_param(branch, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM,
                  param_value(riser, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM))
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


def build_branch(riser, settings, z, level_id, wanted):
    """Costruisce lo stacco su un tratto. Ritorna (caso, stacco, nuovi elementi, raccordi, elementi scollegati)."""
    where, low, high = classify(riser, z)
    direction, axis = branch_direction(riser, wanted)
    point = DB.XYZ(low.X, low.Y, z)
    end = point.Add(direction.Multiply(settings["length_mm"] / MM_PER_FT))

    if isinstance(riser, CableTray) and axis == "h":
        # ramo piatto accostato alla faccia larga, senza raccordo: il montante non viene toccato
        half = (param_value(riser, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM) or 0.0) / 2.0
        if settings["length_mm"] / MM_PER_FT <= half + MIN_REMAINING_MM / MM_PER_FT:
            raise BranchError("The length of the branch must be more than half the depth of the tray "
                              "({:.0f} mm)".format(half * MM_PER_FT))
        STAGE[0] = "creating the branch"
        branch = create_branch(riser, point.Add(direction.Multiply(half)), end, level_id, axis)
        case = "Flat branch, not connected"
        if where != "inside":
            case += " (elevation beyond the {} end of the riser)".format(where)
        return case, branch, [branch], [], []

    disconnected = []
    case = ""
    if where != "inside":
        # quota oltre l'estremo: il montante viene prolungato e si mette la T (le curve sono del comando Turn)
        tip = high if where == "top" else low
        STAGE[0] = "disconnecting the {} end of the riser".format(where)
        # vale la configurazione finale voluta: quanto e' gia' collegato all'estremo viene scollegato
        conn, refs = tip_refs(riser, where)
        for ref in refs:
            disconnected.append("{} {}".format(ref.Owner.Category.Name if ref.Owner.Category else "element",
                                               ref.Owner.Id))
            conn.DisconnectFrom(ref)
        STAGE[0] = "extending the riser past the elevation"
        sign = 1.0 if where == "top" else -1.0
        move_end(riser, tip, point.Add(DB.XYZ(0, 0, sign * settings["stub_mm"] / MM_PER_FT)))
        case = "Riser extended at the {} end, ".format(where)
    diameter, size_note = None, ""
    if isinstance(riser, Pipe) and settings.get("reduce"):
        diameter, size_note = smaller_pipe_size(riser)
        size_note = ", " + size_note
    STAGE[0] = "creating the branch"
    branch = create_branch(riser, point, end, level_id, axis, diameter)
    STAGE[0] = "placing the junction"
    kind, fitting, others = place_junction(riser, point, branch, settings["junction"])
    parts = [("riser", riser), ("branch", branch)] + [("riser above the junction", o) for o in others]
    check_remaining_space(parts)
    return case + kind + size_note, branch, others + [branch], [fitting], disconnected


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


class RevitErrors(DB.IFailuresPreprocessor):
    """Errori di Revit al commit (es. "cannot be rotated in this manner"): invece della finestra bloccante, la
    transazione viene annullata e i messaggi vanno nel report. Gli avvisi passano come sempre."""

    def __init__(self):
        self.messages = []

    def PreprocessFailures(self, accessor):
        for failure in accessor.GetFailureMessages():
            if failure.GetSeverity() != DB.FailureSeverity.Warning:
                self.messages.append(failure.GetDescriptionText())
        if self.messages:
            return DB.FailureProcessingResult.ProceedWithRollBack
        return DB.FailureProcessingResult.Continue


def process(riser, settings, z, wanted):
    result = {"element": riser, "category": category_name(riser), "case": "", "branch": None,
              "fittings": 0, "status": ""}
    level_id = get_level_id(riser)
    if level_id is None:
        result["status"] = "Unchanged. No reference level on the riser"
        return result
    transaction = DB.Transaction(doc, "Branch")
    errors = RevitErrors()
    options = transaction.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(errors)
    options.SetClearAfterRollback(True)
    transaction.SetFailureHandlingOptions(options)
    transaction.Start()
    try:
        case, branch, new_elements, fittings, disconnected = build_branch(riser, settings, z, level_id, wanted)
        STAGE[0] = "copying parameters and insulation"
        note = inherit(riser, new_elements, fittings)
        STAGE[0] = "committing"
        if transaction.Commit() != DB.TransactionStatus.Committed:
            result["status"] = "Unchanged. Revit refused the result: {}".format(
                "; ".join(errors.messages) or "the transaction was rolled back")
            return result
    except BranchError as ex:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        result["status"] = "Unchanged. {}".format(ex)
        return result
    except Exception as ex:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        result["status"] = "Unchanged. The branch could not be built (step: {}): {}".format(STAGE[0], ex)
        result["detail"] = traceback.format_exc() if traceback else ""
        return result
    status = "Created ({})".format(note)
    if disconnected:
        status += ". Disconnected from: {}".format(", ".join(disconnected))
    result.update({"case": case, "branch": branch, "fittings": len(fittings), "status": status})
    return result


def get_selection(pipes_only=False):
    accept = is_vertical_pipe if pipes_only else is_vertical_run
    selected = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    risers = [e for e in selected if accept(e)]
    if risers:
        return risers
    if pipes_only:
        selection_filter, prompt = VerticalPipeFilter(), "Select the vertical pipes, then press Finish"
    else:
        selection_filter = VerticalRunFilter()
        prompt = "Select the vertical pipes, ducts or cable trays, then press Finish"
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, selection_filter, prompt)
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def load_settings(config, kind):
    prefix = CONFIG_PREFIX[kind]
    settings = dict(DEFAULTS)
    try:
        for key in ("direction", "level", "junction"):
            settings[key] = config.get_option(prefix + key, DEFAULTS[key])
        for key in ("offset_mm", "length_mm", "stub_mm"):
            settings[key] = float(config.get_option(prefix + key, DEFAULTS[key]))
        reduce_ = config.get_option(prefix + "reduce", DEFAULTS["reduce"])
        if not isinstance(reduce_, bool):
            reduce_ = str(reduce_).strip().lower() in ("true", "1", "yes")
        settings["reduce"] = reduce_
    except Exception:
        pass
    return settings


def save_settings(config, settings, kind):
    try:
        for key in DEFAULTS:
            setattr(config, CONFIG_PREFIX[kind] + key, settings[key])
        script.save_config()
    except Exception:
        pass


def print_report(results, settings, level, note, kind):
    output.print_md("# Branch report: {}".format(KIND_LABELS[kind].lower()))
    output.print_md("Side **{}** (on screen) | level **{}** {:+g} mm | length **{:g} mm** | junction **{}**".format(
        "custom (clicked)" if settings["direction"] == "custom" else settings["direction"], level.Name,
        settings["offset_mm"], settings["length_mm"], "tee" if kind == "tray" else settings["junction"]))
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


def pick_direction(origin):
    """Direzione orizzontale dall'asse del tubo al punto cliccato nella vista."""
    try:
        picked = uidoc.Selection.PickPoint("Click in the direction of the branch")
    except OperationCanceledException:
        script.exit()
    vector = DB.XYZ(picked.X - origin.X, picked.Y - origin.Y, 0)
    if vector.GetLength() < 1e-6:
        forms.alert("The clicked point is on the axis of the riser: no direction. Run the command again and click "
                    "away from the riser.", exitscript=True)
    return vector.Normalize()


def run_group(kind, risers, levels, names, config, axes, note):
    """Finestra e stacchi per una categoria. Ritorna (risultati, impostazioni, livello) oppure None se annullata."""
    settings = load_settings(config, kind)
    if settings["level"] not in names:
        own = doc.GetElement(get_level_id(risers[0])) if get_level_id(risers[0]) is not None else None
        settings["level"] = own.Name if own is not None else (names[0] if names else "")
    if kind != "pipe" and settings["direction"] == "custom":
        settings["direction"] = "right"  # il clic libero e' solo dei tubi

    def warnings(level_name, offset_mm):
        z_ft = level_elevation(levels[names.index(level_name)]) + offset_mm / MM_PER_FT
        return connection_warnings(risers, z_ft)

    sides_info = tray_sides_text(risers, axes) if kind == "tray" else None
    settings = Branch_ui.BranchWindow(script.get_bundle_file(XAML_FILES[kind]), kind, len(risers), names, settings,
                                      note, warnings, sides_info).show()
    if settings is None:
        return None
    save_settings(config, settings, kind)

    level = levels[names.index(settings["level"])]
    z = level_elevation(level) + settings["offset_mm"] / MM_PER_FT
    if settings["direction"] == "custom":
        wanted = pick_direction(risers[0].Location.Curve.GetEndPoint(0))
    else:
        wanted = screen_direction(settings["direction"], axes)

    results = [process(r, settings, z, wanted) for r in risers]
    return results, settings, level


def main():
    risers = get_selection()
    if not risers:
        script.exit()

    levels = sorted(DB.FilteredElementCollector(doc).OfClass(DB.Level), key=level_elevation)
    names = [lv.Name for lv in levels]
    config = script.get_config(CONFIG_NAME)
    axes = view_axes()
    note = rotation_note(axes)

    # una finestra per categoria, in sequenza; se se ne annulla una le altre gia' fatte restano e vanno nel report
    done = []
    for kind in ("pipe", "duct", "tray"):
        group = [r for r in risers if kind_of(r) == kind]
        if not group:
            continue
        outcome = run_group(kind, group, levels, names, config, axes, note)
        if outcome is None:
            if not done:
                script.exit()
            break
        done.append((kind,) + outcome)

    all_results = [r for _, results, _, _ in done for r in results]
    problems = len([r for r in all_results if not r["status"].startswith("Created")])
    if want_report(len(all_results), problems, "branches created"):
        for kind, results, settings, level in done:
            print_report(results, settings, level, note, kind)


# ----------------------------------------------------------------------------------------------------------------
# Tratti orizzontali (e inclinati fino a 60 gradi): un clic sul punto del tratto, un clic sul lato, stacco a 90 gradi
# ----------------------------------------------------------------------------------------------------------------
VERTICAL_MODE, HORIZONTAL_MODE = "Vertical runs", "Horizontal runs"
MAX_SLOPE_DEG = 60.0
H_LENGTH_KEY, H_LENGTH_DEFAULT = "h_length_mm", 500.0  # tratto orizzontale fino al gomito
H_DROP_KEY, H_DROP_DEFAULT = "h_drop_mm", 0.0  # dislivello del tratto inclinato: + sale, - scende, 0 = nessuno
H_ANGLE_KEY, H_ANGLE_DEFAULT = "h_angle_deg", 90.0  # angolo del tratto inclinato dall'orizzontale
H_RETURN_KEY, H_RETURN_DEFAULT = "h_return", False  # dopo l'orizzontale, tratto inclinato che riporta alla quota di partenza
def is_horizontal_run(element):
    """Tubo, condotto o passerella rettilinea con pendenza fino a 60 gradi dall'orizzontale."""
    if not isinstance(element, (Pipe, Duct, CableTray)):
        return False
    curve = element.Location.Curve
    return isinstance(curve, DB.Line) and abs(curve.Direction.Z) <= math.sin(math.radians(MAX_SLOPE_DEG))


class HorizontalRunFilter(ISelectionFilter):
    """Tubi, canali e passerelle orizzontali o inclinati fino a 60 gradi."""

    def AllowElement(self, element):
        try:
            return is_horizontal_run(element)
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return False


def point_on_axis(run, picked):
    """Punto dell'asse del tratto corrispondente al clic. In pianta la Z del clic non conta: si proietta in XY."""
    curve = run.Location.Curve
    a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
    if isinstance(doc.ActiveView, DB.ViewPlan):
        d = b.Subtract(a)
        t = ((picked.X - a.X) * d.X + (picked.Y - a.Y) * d.Y) / (d.X ** 2 + d.Y ** 2)
        t = min(max(t, 0.0), 1.0)
        return a.Add(d.Multiply(t))
    return curve.Project(picked).XYZPoint


def side_direction(run, point, side_click):
    """Direzione orizzontale a 90 gradi dal tratto, verso il lato cliccato. None se il clic e' sull'asse."""
    run_xy = flat(run.Location.Curve.Direction)
    perpendicular = DB.XYZ.BasisZ.CrossProduct(run_xy)
    side = (side_click.X - point.X) * perpendicular.X + (side_click.Y - point.Y) * perpendicular.Y
    if abs(side) < 1.0 / MM_PER_FT:
        return None
    return perpendicular if side > 0 else perpendicular.Negate()


def section_half_height(element):
    """Meta' altezza della sezione del condotto (raggio se tondo), in piedi."""
    for bip in (DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM, DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM):
        value = param_value(element, bip)
        if value:
            return value / 2.0
    raise BranchError("Cannot read the section of the duct")


def horizontal_axis(cs):
    """(nome, vettore) dell'asse orizzontale del connettore: X o Y, quello che non e' verticale. In un canale di costa
    e' Y (l'altezza), in uno piano e' X (la larghezza)."""
    if abs(cs.BasisX.Z) < 0.01:
        return "X", cs.BasisX
    return "Y", cs.BasisY


def align_vertical_section(run, vertical, point):
    """Ruota il canale verticale attorno al proprio asse: in pianta la sezione deve coincidere con quella del tratto
    di partenza. Se il tratto e' di costa la sua larghezza e' verticale e la dimensione orizzontale e' l'altezza:
    senza questa rotazione Revit orienta il verticale sugli assi del progetto e le due misure risultano scambiate."""
    if is_round(run):
        return
    doc.Regenerate()
    name, run_axis = horizontal_axis(nearest_connector(run, point).CoordinateSystem)
    new_cs = nearest_connector(vertical, point).CoordinateSystem
    new_axis = new_cs.BasisX if name == "X" else new_cs.BasisY
    angle = math.atan2(run_axis.Y, run_axis.X) - math.atan2(new_axis.Y, new_axis.X)
    # sezione simmetrica: basta allineare la direzione, a meno di 180 gradi
    while angle > math.pi / 2:
        angle -= math.pi
    while angle <= -math.pi / 2:
        angle += math.pi
    if abs(angle) > 1e-6:
        axis = DB.Line.CreateBound(point, point.Add(DB.XYZ.BasisZ))
        DB.ElementTransformUtils.RotateElement(doc, vertical.Id, axis, angle)
        doc.Regenerate()


def tray_is_on_edge(tray):
    """True se la passerella e' di costa: la larghezza e' verticale. Si confronta l'estensione verticale della
    geometria (nel piano della sezione) con larghezza e altezza. Sezione quasi quadrata: considerata piatta."""
    width = param_value(tray, DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
    height = param_value(tray, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
    if not width or not height or abs(width - height) < 1.0 / MM_PER_FT:
        return False
    direction = tray.Location.Curve.Direction
    up = DB.XYZ.BasisZ.Subtract(direction.Multiply(DB.XYZ.BasisZ.DotProduct(direction)))
    if up.GetLength() < 1e-6:
        return False
    up = up.Normalize()
    try:
        low, high = 1e9, -1e9
        pending = list(tray.get_Geometry(DB.Options()))
        while pending:
            item = pending.pop()
            if isinstance(item, DB.GeometryInstance):
                pending.extend(item.GetInstanceGeometry())
            elif isinstance(item, DB.Solid):
                for edge in item.Edges:
                    for end in (0, 1):
                        value = edge.AsCurve().GetEndPoint(end).DotProduct(up)
                        low, high = min(low, value), max(high, value)
        extent = high - low
        if extent < 0:
            return False
        return abs(extent - width) < abs(extent - height)
    except Exception:
        return False


def optional_elbow(first, second, point):
    """Gomito tra due tratti di passerella, se Revit riesce a metterlo (sezioni non sempre compatibili): altrimenti
    None e i tratti restano non collegati. Ritorna (gomito o None, nota)."""
    sub = DB.SubTransaction(doc)
    sub.Start()
    try:
        elbow = doc.Create.NewElbowFitting(nearest_connector(first, point), nearest_connector(second, point))
        sub.Commit()
        return elbow, ""
    except Exception:
        sub.RollBack()
        return None, "no elbow at {:.0f} mm: runs not connected".format(point.Z * MM_PER_FT)


def build_tray_branch(run, point, direction, length_mm, junction, drop_mm, angle_deg, level_id, on_edge):
    """Stacco da passerella dove il T esiste solo se lo stacco giace nel piano di tratto e larghezza: passerella
    piatta, stacchi di lato (gestiti altrove); passerella di costa, stacchi verticali. Negli altri casi solo nuovi
    tratti, senza T. Ritorna (descrizione, stacco, pezzi del tratto, nuovi tratti, raccordi, note)."""
    height = param_value(run, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM) or 0.0
    width = param_value(run, DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM) or 0.0
    half_v = (width if on_edge else height) / 2.0  # semi-estensione verticale della sezione
    half_h = (height if on_edge else width) / 2.0  # semi-estensione orizzontale, perpendicolare al tratto
    drop = drop_mm / MM_PER_FT
    inclined = bool(drop_mm)
    vertical = inclined and abs(angle_deg - 90.0) < 1e-6
    length = length_mm / MM_PER_FT
    notes = []

    if vertical and not on_edge:
        # piatta, stacco verticale: perpendicolare alla larghezza, nessun T. Parte dalla faccia superiore o inferiore
        STAGE[0] = "creating the vertical tray"
        start = DB.XYZ(point.X, point.Y, point.Z + (half_v if drop > 0 else -half_v))
        top = start.Add(DB.XYZ(0, 0, drop))
        branch = create_branch(run, start, top, level_id, "same")
        # larghezza lungo il tratto di partenza, come quella dell'orizzontale che segue: i gomiti la richiedono cosi'
        DB.ElementTransformUtils.RotateElement(
            doc, branch.Id, DB.Line.CreateBound(start, start.Add(DB.XYZ.BasisZ)), math.pi / 2)
        doc.Regenerate()
        STAGE[0] = "creating the horizontal run after the change of level"
        final = create_branch(run, top, top.Add(direction.Multiply(length)), level_id, "same")
        elbow, note = optional_elbow(branch, final, top)
        return ("Cable tray: new runs only (no tee perpendicular to the width)", branch, [], [branch, final],
                [elbow] if elbow is not None else [], [note] if note else [])

    if vertical and on_edge:
        # di costa, stacco verticale: giace nel piano di tratto e larghezza, il T esiste
        STAGE[0] = "creating the vertical tray"
        top = point.Add(DB.XYZ(0, 0, drop))
        branch = create_branch(run, point, top, level_id, "same")
        STAGE[0] = "placing the junction"
        kind, fitting, others = place_junction(run, point, branch, junction)
        check_remaining_space([("run", run), ("vertical run", branch)] + [("rest of the run", o) for o in others])
        STAGE[0] = "creating the horizontal run after the change of level"
        # dopo la discesa la passerella e' piatta: la larghezza del verticale e' lungo il tratto di partenza
        final = create_branch(run, top, top.Add(direction.Multiply(length)), level_id, None)
        elbow, note = optional_elbow(branch, final, top)
        return (kind, branch, others, [branch, final], [fitting] + ([elbow] if elbow is not None else []),
                [note] if note else [])

    if not inclined:
        # di costa, stacco di lato senza offset: nessun T, solo il tratto dalla faccia laterale
        start = point.Add(direction.Multiply(half_h))
        branch = create_branch(run, start, start.Add(direction.Multiply(length)), level_id, "same")
        return "Cable tray on edge: new run only (no tee perpendicular to the width)", branch, [], [branch], [], []

    # Il tratto corto nasce lungo (TRAY_LONG_STUB_FT): T e gomito si mangiano parte della sua lunghezza e con 50 mm
    # non ci starebbero (Revit rifiuta al commit: "opposite direction or insufficient space"). Solo alla fine, con
    # tutto collegato, si accorcia spostando indietro inclinato, finale e gomiti.
    if not on_edge:
        # piatta, stacco inclinato: T, 50 mm di orizzontale dopo il raccordo, salita o discesa fino alla quota con
        # l'angolo scelto, tratto finale. Il T sta nel piano di tratto e larghezza, come lo stacco laterale
        STAGE[0] = "creating the stub and the junction"
        start = point
        kind, fitting = None, None
    else:
        # di costa, stacco di lato: nessun T, solo tratti. Parte dalla faccia laterale, sempre di costa
        STAGE[0] = "creating the horizontal stub"
        start = point.Add(direction.Multiply(half_h))
        kind, fitting, others = "Cable tray on edge: new runs only (no tee perpendicular to the width)", None, []
    stub_end = start.Add(direction.Multiply(TRAY_LONG_STUB_FT))
    stub = create_branch(run, start, stub_end, level_id, "same")
    if not on_edge:
        kind, fitting, others = place_junction(run, point, stub, junction)
        check_remaining_space([("run", run), ("stub", stub)] + [("rest of the run", o) for o in others])
    # i tratti seguenti sono copie del corto, per conservare l'orientamento della sezione (piatta o di costa)
    STAGE[0] = "creating the inclined run"
    tilt = math.radians(angle_deg) * (1.0 if drop > 0 else -1.0)
    slope_len = abs(drop) / math.sin(math.radians(angle_deg))
    top = stub_end.Add(direction.Multiply(abs(drop) / math.tan(math.radians(angle_deg)))).Add(DB.XYZ(0, 0, drop))
    slope = tray_piece(stub, stub_end, direction, slope_len, tilt)
    STAGE[0] = "creating the horizontal run after the change of level"
    final = tray_piece(stub, top, direction, length, 0.0)
    STAGE[0] = "placing the elbows"
    elbow1, note1 = optional_elbow(stub, slope, stub_end)
    elbow2, note2 = optional_elbow(slope, final, top)
    STAGE[0] = "shortening the stub to {:g} mm".format(TRAY_STUB_FT * MM_PER_FT)
    shorten_stub(stub, start, direction, [slope, final, elbow1, elbow2], elbow1 is not None)
    return (kind, stub, others, [stub, slope, final],
            [f for f in (fitting, elbow1, elbow2) if f is not None], [n for n in (note1, note2) if n])


TRAY_STUB_FT = 50.0 / MM_PER_FT  # orizzontale visibile dopo il T (o dalla faccia, di costa), prima della salita
TRAY_LONG_STUB_FT = 1000.0 / MM_PER_FT  # lunghezza provvisoria perche' T e gomito abbiano spazio


def shorten_stub(stub, start, direction, movers, connected):
    """Porta il tratto corto a TRAY_STUB_FT di lunghezza visibile (tra T, o faccia, e gomito) spostando indietro,
    tutti insieme, inclinato, finale e gomiti: il tratto corto, collegato al gomito, si accorcia da solo. Senza
    gomito lo si accorcia a mano, tenendo il verso della curva (invertirlo fa rifiutare il risultato a Revit)."""
    doc.Regenerate()
    curve = stub.Location.Curve
    delta = curve.Length - TRAY_STUB_FT
    if delta <= 1.0 / MM_PER_FT:
        return
    ids = List[DB.ElementId]([m.Id for m in movers if m is not None])
    DB.ElementTransformUtils.MoveElements(doc, ids, direction.Multiply(-delta))
    if not connected:
        a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
        if a.DistanceTo(start) < b.DistanceTo(start):
            stub.Location.Curve = DB.Line.CreateBound(a, a.Add(direction.Multiply(TRAY_STUB_FT)))
        else:
            stub.Location.Curve = DB.Line.CreateBound(b.Add(direction.Multiply(TRAY_STUB_FT)), b)
    doc.Regenerate()


def tray_piece(base, start, direction, length, tilt_rad):
    """Copia del tratto base (orizzontale, lungo direction) messa a start, lunga length e inclinata di tilt_rad
    (positivo = verso l'alto) attorno all'asse orizzontale perpendicolare a direction: la rotazione non cambia
    l'orientamento della sezione, piatta o di costa."""
    copy_id = list(DB.ElementTransformUtils.CopyElement(doc, base.Id, DB.XYZ.Zero))[0]
    piece = doc.GetElement(copy_id)
    piece.Location.Curve = DB.Line.CreateBound(start, start.Add(direction.Multiply(length)))
    doc.Regenerate()
    if abs(tilt_rad) > 1e-9:
        side = DB.XYZ.BasisZ.CrossProduct(direction)  # una rotazione positiva attorno a side abbassa il tratto
        DB.ElementTransformUtils.RotateElement(doc, copy_id, DB.Line.CreateBound(start, start.Add(side)), -tilt_rad)
    mark = piece.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
    if mark is not None and not mark.IsReadOnly:
        mark.Set("")  # la copia eredita il Mark: duplicato
    doc.Regenerate()
    return piece



STUB_MIN_MM = 150.0  # tratto corto dopo lo stacco, dove stanno il T o lo stacco: per condotti e passerelle, il doppio


def stub_ft(run):
    """Lunghezza del tratto corto dopo lo stacco, in piedi: almeno STUB_MIN_MM, o il doppio dell'altezza sezione."""
    try:
        section_ft = 2.0 * section_half_height(run)
    except BranchError:
        section_ft = 0.0
    return max(STUB_MIN_MM / MM_PER_FT, 2.0 * section_ft)


def connect_elbow(first, second, point):
    """Gomito tra due tratti che si incontrano in point."""
    try:
        return doc.Create.NewElbowFitting(nearest_connector(first, point), nearest_connector(second, point))
    except Exception as ex:
        raise BranchError("No elbow could be placed after the change of level ({})".format(ex))


def add_inclined(branch, run, elbow_point, direction, drop_mm, angle_deg, level_id):
    """Tratto inclinato dal gomito alla quota finale: sale o scende di drop_mm con l'angolo dato dall'orizzontale,
    nella direzione dello stacco. Il gomito tra lo stacco orizzontale e l'inclinato sta in elbow_point. Ritorna
    (tratto inclinato, gomito, punto finale). A 90 gradi il tratto e' verticale."""
    drop = drop_mm / MM_PER_FT
    run_ft = 0.0 if abs(angle_deg - 90.0) < 1e-6 else abs(drop) / math.tan(math.radians(angle_deg))
    end = elbow_point.Add(direction.Multiply(run_ft)).Add(DB.XYZ(0, 0, drop))
    slope = create_branch(run, elbow_point, end, level_id, "same")
    try:
        elbow = doc.Create.NewElbowFitting(nearest_connector(branch, elbow_point),
                                           nearest_connector(slope, elbow_point))
    except Exception as ex:
        raise BranchError("No elbow of {:g} degrees could be placed with the routing preferences of this type "
                          "({})".format(angle_deg, ex))
    return slope, elbow, end


def tray_return(final, direction, drop_mm, angle_deg):
    """Passerella: dall'estremo del tratto finale, tratto inclinato (stesso angolo, verso opposto) che riporta alla
    quota di partenza, poi un altro orizzontale della stessa lunghezza, con i gomiti se Revit riesce a metterli.
    Ritorna (tratti, gomiti, note)."""
    ends = [final.Location.Curve.GetEndPoint(i) for i in (0, 1)]
    start = max(ends, key=lambda p: p.DotProduct(direction))
    drop = -drop_mm / MM_PER_FT
    tilt = math.radians(angle_deg) * (1.0 if drop > 0 else -1.0)
    piece = tray_piece(final, start, direction, abs(drop) / math.sin(math.radians(angle_deg)), tilt)
    elbow, note = optional_elbow(final, piece, start)
    low = max([piece.Location.Curve.GetEndPoint(i) for i in (0, 1)], key=lambda p: p.DotProduct(direction))
    last = tray_piece(final, low, direction, final.Location.Curve.Length, 0.0)
    elbow2, note2 = optional_elbow(piece, last, low)
    return [piece, last], [e for e in (elbow, elbow2) if e is not None], [n for n in (note, note2) if n]


def process_horizontal(run, point, direction, length_mm, junction, drop_mm=0.0, angle_deg=90.0, return_back=False):
    """Stacco in point: tratto orizzontale di length_mm, poi, se drop_mm e' diverso da zero, un tratto inclinato
    che porta alla quota finale con angle_deg. Con return_back, dopo l'orizzontale un secondo tratto inclinato (stesso
    angolo, verso opposto) riporta alla quota di partenza. Con T (tratto diviso) o tap secondo l'innesto salvato per
    la categoria."""
    result = {"element": run, "category": category_name(run), "case": "", "branch": None, "fittings": 0,
              "status": ""}
    level_id = get_level_id(run)
    if level_id is None:
        result["status"] = "Unchanged. No reference level on the run"
        return result
    transaction = DB.Transaction(doc, "Branch")
    errors = RevitErrors()
    options = transaction.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(errors)
    options.SetClearAfterRollback(True)
    transaction.SetFailureHandlingOptions(options)
    transaction.Start()
    try:
        inclined = bool(drop_mm)
        return_back = return_back and inclined
        tray_notes = []

        def add_return(final, top):
            """Tubi e condotti: dall'estremo dell'orizzontale, tratto inclinato opposto fino alla quota di partenza, poi un
            altro orizzontale della stessa lunghezza. Ritorna (tratti, gomiti)."""
            STAGE[0] = "creating the return run to the starting level"
            end = top.Add(direction.Multiply(length_mm / MM_PER_FT))
            slope, elbow, low = add_inclined(final, run, end, direction, -drop_mm, angle_deg, level_id)
            STAGE[0] = "creating the last horizontal run"
            last = create_branch(run, low, low.Add(direction.Multiply(length_mm / MM_PER_FT)), level_id, "same")
            elbow2 = connect_elbow(slope, last, low)
            doc.Regenerate()
            check_remaining_space([("horizontal run", final), ("return run", slope), ("last run", last)])
            return [slope, last], [elbow, elbow2]

        from_below = inclined and isinstance(run, Duct) and abs(angle_deg - 90.0) < 1e-6
        if from_below:
            # condotto a 90 gradi: stacco dal basso, verticale dalla faccia inferiore del condotto, con una presa
            if drop_mm > 0:
                # salendo dalla faccia inferiore il tratto attraverserebbe il condotto stesso
                raise BranchError("A duct branch at 90 degrees from below must go down: use a negative offset")
            STAGE[0] = "creating the duct from below"
            half = section_half_height(run)
            bottom = DB.XYZ(point.X, point.Y, point.Z - half)
            top = bottom.Add(DB.XYZ(0, 0, drop_mm / MM_PER_FT))
            branch = create_branch(run, bottom, top, level_id, "same")
            align_vertical_section(run, branch, bottom)
            STAGE[0] = "placing the take-off from below"
            fitting = doc.Create.NewTakeoffFitting(nearest_connector(branch, bottom), run)
            kind, others = "Tap from below", []
            check_remaining_space([("run", run), ("branch", branch)])
            # poi, come negli altri casi, l'orizzontale di length_mm dalla quota raggiunta, con un gomito
            STAGE[0] = "creating the horizontal run after the change of level"
            final = create_branch(run, top, top.Add(direction.Multiply(length_mm / MM_PER_FT)), level_id, "same")
            elbow = connect_elbow(branch, final, top)
            doc.Regenerate()
            check_remaining_space([("vertical run", branch), ("horizontal run", final)])
            parts, fittings = [branch, final], [fitting, elbow]
            if return_back:
                back, back_elbows = add_return(final, top)
                parts.extend(back)
                fittings.extend(back_elbows)
        elif isinstance(run, CableTray) and (tray_is_on_edge(run) or inclined):
            # passerelle: il T solo se lo stacco giace nel piano di tratto e larghezza (di costa: stacchi verticali)
            kind, branch, others, parts, fittings, tray_notes = build_tray_branch(
                run, point, direction, length_mm, junction, drop_mm, angle_deg, level_id, tray_is_on_edge(run))
            if return_back:
                STAGE[0] = "creating the return run to the starting level"
                back, back_elbows, back_notes = tray_return(parts[-1], direction, drop_mm, angle_deg)
                parts = parts + back
                fittings = fittings + back_elbows
                tray_notes.extend(back_notes)
            doc.Regenerate()
            check_remaining_space([("new run", p) for p in parts] + [("rest of the run", o) for o in others])
        elif inclined and isinstance(run, Pipe):
            # tubi: il T e' ruotato dell'angolo scelto e lo stacco prosegue inclinato fino alla quota di offset;
            # da li' si continua con l'orizzontale di length_mm
            STAGE[0] = "creating the inclined branch"
            drop = drop_mm / MM_PER_FT
            run_ft = 0.0 if abs(angle_deg - 90.0) < 1e-6 else abs(drop) / math.tan(math.radians(angle_deg))
            top = point.Add(direction.Multiply(run_ft)).Add(DB.XYZ(0, 0, drop))
            slope = create_branch(run, point, top, level_id, "same")
            STAGE[0] = "placing the junction"
            kind, fitting, others = place_junction(run, point, slope, junction)
            check_remaining_space([("run", run), ("inclined branch", slope)] +
                                  [("rest of the run", o) for o in others])
            STAGE[0] = "creating the horizontal run after the change of level"
            final = create_branch(run, top, top.Add(direction.Multiply(length_mm / MM_PER_FT)), level_id, "same")
            elbow = connect_elbow(slope, final, top)
            doc.Regenerate()
            check_remaining_space([("inclined branch", slope), ("horizontal run", final)])
            branch = slope
            parts, fittings = [slope, final], [fitting, elbow]
            if return_back:
                back, back_elbows = add_return(final, top)
                parts.extend(back)
                fittings.extend(back_elbows)
        else:
            if inclined:
                # prima il tratto corto con T (o stacco), poi il cambio di quota, poi l'orizzontale di length_mm
                STAGE[0] = "creating the stub and the junction"
                stub_end = point.Add(direction.Multiply(stub_ft(run)))
                branch = create_branch(run, point, stub_end, level_id, "same")
                kind, fitting, others = place_junction(run, point, branch, junction)
                check_remaining_space([("run", run), ("stub", branch)] + [("rest of the run", o) for o in others])
                STAGE[0] = "creating the inclined run and its elbow"
                slope, elbow, top = add_inclined(branch, run, stub_end, direction, drop_mm, angle_deg, level_id)
                STAGE[0] = "creating the horizontal run after the change of level"
                final = create_branch(run, top, top.Add(direction.Multiply(length_mm / MM_PER_FT)), level_id, "same")
                elbow2 = connect_elbow(slope, final, top)
                doc.Regenerate()
                check_remaining_space([("inclined run", slope), ("horizontal run", final)])
                parts, fittings = [branch, slope, final], [fitting, elbow, elbow2]
                if return_back:
                    back, back_elbows = add_return(final, top)
                    parts.extend(back)
                    fittings.extend(back_elbows)
            else:
                STAGE[0] = "creating the branch"
                end = point.Add(direction.Multiply(length_mm / MM_PER_FT))
                branch = create_branch(run, point, end, level_id, "same")
                STAGE[0] = "placing the junction"
                kind, fitting, others = place_junction(run, point, branch, junction)
                check_remaining_space([("run", run), ("branch", branch)] + [("rest of the run", o) for o in others])
                parts, fittings = [branch], [fitting]
        STAGE[0] = "copying parameters and insulation"
        note = "; ".join([inherit(run, others + parts, fittings)] + tray_notes)
        STAGE[0] = "committing"
        if transaction.Commit() != DB.TransactionStatus.Committed:
            result["status"] = "Unchanged. Revit refused the result: {}".format(
                "; ".join(errors.messages) or "the transaction was rolled back")
            if inclined and isinstance(run, (CableTray, Duct)):
                result["status"] += (" (the inclined run may be too short for its two elbows: try a larger offset or "
                                     "a steeper angle)")
            return result
    except BranchError as ex:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        result["status"] = "Unchanged. {}".format(ex)
        return result
    except Exception as ex:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        result["status"] = "Unchanged. The branch could not be built (step: {}): {}".format(STAGE[0], ex)
        result["detail"] = traceback.format_exc() if traceback else ""
        return result
    result.update({"case": kind, "branch": branch, "fittings": len(fittings),
                   "status": "Created ({})".format(note)})
    return result


def print_horizontal_report(results, length_mm, drop_mm=0.0, angle_deg=90.0, return_back=False):
    output.print_md("# Branch report: horizontal runs")
    if drop_mm:
        output.print_md("Pipes: junction at **{:g} degrees**, inclined branch up to an offset of **{:+g} mm**, then "
                        "a horizontal run of **{:g} mm**. Ducts and cable trays: short stub with the junction, then "
                        "the inclined run and the horizontal run".format(angle_deg, drop_mm, length_mm))
        if return_back:
            output.print_md("After the horizontal run, a return run at the same angle brings the branch back to the "
                            "starting level, followed by a last horizontal run of the same length")
    else:
        output.print_md("Branches at 90 degrees | length **{:g} mm** from the run axis".format(length_mm))
    done = [r for r in results if r["status"].startswith("Created")]
    output.print_md("Branches requested: **{}** | created: **{}** | not created: **{}**".format(
        len(results), len(done), len(results) - len(done)))
    if not results:
        output.print_md("No run was picked.")
        return
    rows = [[i, output.linkify(r["element"].Id), r["category"], r["case"],
             output.linkify(r["branch"].Id) if r["branch"] is not None else "", r["status"]]
            for i, r in enumerate(results, 1)]
    output.print_table(table_data=rows, columns=["#", "Run", "Category", "Junction", "Branch", "Status"])
    for i, r in enumerate(results, 1):
        if r.get("detail"):
            output.print_md("**Technical detail for branch #{}**".format(i))
            output.print_code(r["detail"])


def main_horizontal():
    config = script.get_config(CONFIG_NAME)
    try:
        length_mm = float(config.get_option(H_LENGTH_KEY, H_LENGTH_DEFAULT))
        drop_mm = float(config.get_option(H_DROP_KEY, H_DROP_DEFAULT))
        angle_deg = float(config.get_option(H_ANGLE_KEY, H_ANGLE_DEFAULT))
        return_back = str(config.get_option(H_RETURN_KEY, H_RETURN_DEFAULT)).lower() in ("true", "1")
    except Exception:
        length_mm, drop_mm, angle_deg = H_LENGTH_DEFAULT, H_DROP_DEFAULT, H_ANGLE_DEFAULT
        return_back = H_RETURN_DEFAULT
    if __shiftclick__:  # noqa: F821
        values = Horizontal_ui.HorizontalWindow(script.get_bundle_file(HORIZONTAL_XAML), {
            "length_mm": length_mm, "drop_mm": drop_mm, "angle_deg": angle_deg, "return_back": return_back}).show()
        if values is None:
            script.exit()
        length_mm, drop_mm, angle_deg = values["length_mm"], values["drop_mm"], values["angle_deg"]
        return_back = values["return_back"]
        try:
            setattr(config, H_LENGTH_KEY, length_mm)
            setattr(config, H_DROP_KEY, drop_mm)
            setattr(config, H_ANGLE_KEY, angle_deg)
            setattr(config, H_RETURN_KEY, return_back)
            script.save_config()
        except Exception:
            pass
    junctions = {kind: load_settings(config, kind)["junction"] for kind in ("pipe", "duct", "tray")}

    results = []
    run_filter = HorizontalRunFilter()
    while True:
        try:
            reference = uidoc.Selection.PickObject(
                ObjectType.Element, run_filter,
                "Branch ({:g} mm): click the point of the run where the branch starts (ESC to finish)".format(
                    length_mm))
        except OperationCanceledException:
            break
        run = doc.GetElement(reference.ElementId)
        point = point_on_axis(run, reference.GlobalPoint or run.Location.Curve.Evaluate(0.5, True))
        try:
            side_click = uidoc.Selection.PickPoint("Branch: click on the side of the run where the branch goes "
                                                   "(ESC to skip this one)")
        except OperationCanceledException:
            continue  # lato annullato: questo stacco viene saltato, si continua
        direction = side_direction(run, point, side_click)
        if direction is None:
            results.append({"element": run, "category": category_name(run), "case": "", "branch": None,
                            "fittings": 0, "status": "Unchanged. The side click is on the run axis: no side"})
            continue
        results.append(process_horizontal(run, point, direction, length_mm, junctions[kind_of(run)],
                                          drop_mm, angle_deg, return_back))

    problems = len([r for r in results if not r["status"].startswith("Created")])
    if want_report(len(results), problems, "branches created"):
        print_horizontal_report(results, length_mm, drop_mm, angle_deg, return_back and bool(drop_mm))


main_horizontal()
