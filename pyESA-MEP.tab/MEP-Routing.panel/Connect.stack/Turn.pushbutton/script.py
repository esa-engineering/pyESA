# -*- coding: utf-8 -*-
__title__ = "Turn"
__doc__ = """Version = 0.7
Date    = 16.10.2026
________________________________________________________________
Description:
Adds a run of fixed length that starts at the free end of a pipe, duct,
cable tray or conduit and runs parallel to another element (wall, beam,
pipe, duct, grid, vertical face...), joined with an elbow.

Instruction:
- CLICK: starts straight away with the saved settings
- SHIFT + CLICK: window with the length of the new run (default
  1000 mm) and a tick to keep the slope of the existing run (default
  off: the new run is horizontal). Both are saved
- Click the element to extend, near its free end
- Second click: a point to the left or right of the run gives a classic
  90 degree turn to that side
- To follow an element instead, press ESC at the second click and click
  the element to align to (also in linked RVT and IFC models). An element
  with an axis is followed along its axis, anything else along the clicked
  face. If it is perpendicular to the run, the side of the click sets the
  side of the turn. ESC again skips the element, ESC on the first click
  finishes
- The old run is not changed: the new run starts exactly at its free end,
  is horizontal (or keeps its slope, if ticked), and the elbow joins the two
- Ducts and cable trays are copied and rotated, so width and height stay
  as in the original (also for a tray or duct on edge)

Turns of more than 90 degrees are refused. Vertical ends are skipped.
The new parts keep the parameters and the insulation of the starting
element.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import math

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB.Plumbing import Pipe, PipeInsulation
from Autodesk.Revit.DB.Mechanical import Duct, DuctInsulation, DuctLining
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, forms, DB

import Turn_ui

CONFIG_NAME = "ESA_MEPTurn"
XAML_FILE_NAME = "Turn_form.xaml"
DEFAULT_LENGTH_MM = 1000.0
MIN_LENGTH_MM = 50.0
MIN_ANGLE, MAX_ANGLE = 0.1, 90.0

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
MIN_REMAINING_MM = 20.0  # lunghezza minima dei tratti dopo l'inserimento del gomito

CATEGORY_NAMES = (
    (Pipe, "Pipe"),
    (Duct, "Duct"),
    (CableTray, "Cable Tray"),
    (Conduit, "Conduit"),
)


class MepCurveFilter(ISelectionFilter):
    """Accetta solo tubi, condotti, passerelle e conduit."""

    def AllowElement(self, element):
        return isinstance(element, (Pipe, Duct, CableTray, Conduit))

    def AllowReference(self, reference, point):
        return False


class TurnError(Exception):
    """Errore atteso (es. spazio insufficiente): il messaggio va nel report."""


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


def nearest_free_connector(element, point):
    # in pianta la Z di GlobalPoint non e' significativa: si confronta solo XY
    planar = isinstance(doc.ActiveView, DB.ViewPlan)
    best, best_dist = None, None
    for conn in element.ConnectorManager.UnusedConnectors:
        origin = conn.Origin
        if planar:
            dist = ((origin.X - point.X) ** 2 + (origin.Y - point.Y) ** 2) ** 0.5
        else:
            dist = origin.DistanceTo(point)
        if best is None or dist < best_dist:
            best, best_dist = conn, dist
    return best


def nearest_connector(element, point):
    best, best_dist = None, None
    for c in element.ConnectorManager.Connectors:
        dist = c.Origin.DistanceTo(point)
        if best is None or dist < best_dist:
            best, best_dist = c, dist
    return best


def find_fitting(conn, element):
    """Raccordo (FamilyInstance) collegato al connettore, se esiste."""
    for ref in conn.AllRefs:
        owner = ref.Owner
        if owner.Id != element.Id and isinstance(owner, DB.FamilyInstance):
            return owner
    return None


def fitting_label(fitting):
    try:
        return "{}: {}".format(fitting.Symbol.FamilyName, fitting.Name)
    except Exception:
        return "elbow"


def copy_conduit_diameter(source, target):
    for bip in (DB.BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM,
                DB.BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM):
        src = source.get_Parameter(bip)
        tgt = target.get_Parameter(bip)
        if src and tgt and not tgt.IsReadOnly:
            tgt.Set(src.AsDouble())
            return


def run_direction_at(element, conn):
    """Direzione di percorrenza: dall'altra estremita' verso quella del connettore."""
    curve = element.Location.Curve
    a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
    if a.DistanceTo(conn.Origin) < b.DistanceTo(conn.Origin):
        return a.Subtract(b).Normalize()
    return b.Subtract(a).Normalize()


def side_of(travel, origin, click):
    """Lato del clic rispetto alla direzione di percorrenza, visto dall'alto: "left", "right" o None se il clic e'
    in linea con il tratto (in pianta conta solo XY)."""
    cross = travel.X * (click.Y - origin.Y) - travel.Y * (click.X - origin.X)
    horizontal_len = math.sqrt(travel.X ** 2 + travel.Y ** 2)
    if abs(cross) / horizontal_len < 1.0 / MM_PER_FT:  # meno di 1 mm dalla linea del tratto
        return None
    return "left" if cross > 0 else "right"


# ----------------------------------------------------------------------------------------------------------------
# Elemento di riferimento: il nuovo tratto ha l'asse coincidente con il suo (anche in un link)
# ----------------------------------------------------------------------------------------------------------------
VERTICAL_FACE_TOL = 0.1  # sotto questa componente orizzontale la normale della faccia cliccata non da' una direzione
PERPENDICULAR_SIN = math.sin(math.radians(1.0))  # entro 1 grado dalla perpendicolare il verso lo decide il clic


def flat_forward(travel):
    """Direzione di percorrenza in pianta (orizzontale, unitaria)."""
    horizontal_len = math.sqrt(travel.X ** 2 + travel.Y ** 2)
    return DB.XYZ(travel.X / horizontal_len, travel.Y / horizontal_len, 0)


def signed_angle_to(forward, target):
    """Angolo con segno in gradi da forward a target (orizzontali): positivo = sinistra visto dall'alto."""
    return math.degrees(math.atan2(forward.CrossProduct(target).Z, forward.DotProduct(target)))


def hit_face_normal(reference, element, link_instance):
    """Normale (coordinate host) della faccia piana cliccata, o None se non e' leggibile. Per gli elementi di un link
    la faccia si legge dal documento del link e la normale si porta nel modello host."""
    try:
        if link_instance is not None:
            face = element.GetGeometryObjectFromReference(reference.CreateReferenceInLink())
        else:
            face = element.GetGeometryObjectFromReference(reference)
        if not isinstance(face, DB.PlanarFace):
            return None
        normal = face.FaceNormal
        return link_instance.GetTotalTransform().OfVector(normal) if link_instance is not None else normal
    except Exception:
        return None


def pick_alignment(prompt):
    """Clic su un elemento di riferimento, anche in un link RVT o IFC. Ritorna (elemento, transform del link o None,
    nome del modello, punto cliccato, normale della faccia cliccata o None), o None con ESC."""
    try:
        reference = uidoc.Selection.PickObject(ObjectType.PointOnElement, prompt)
    except OperationCanceledException:
        return None
    host = doc.GetElement(reference.ElementId)
    hit = reference.GlobalPoint
    if isinstance(host, DB.RevitLinkInstance):
        link_doc = host.GetLinkDocument()
        if link_doc is None or reference.LinkedElementId == DB.ElementId.InvalidElementId:
            raise TurnError("The element clicked in the linked model could not be read")
        element = link_doc.GetElement(reference.LinkedElementId)
        return element, host.GetTotalTransform(), link_doc.Title, hit, hit_face_normal(reference, element, host)
    return host, None, "host model", hit, hit_face_normal(reference, host, None)


def element_axis(element, transform, hit):
    """Asse orizzontale dell'elemento vicino al punto cliccato (muri, travi, tubi, canali, griglie, linee): ritorna
    (tangente orizzontale unitaria, punto sull'asse) in coordinate host, o (None, None) se non c'e'."""
    location = getattr(element, "Location", None)
    if isinstance(location, DB.LocationCurve):
        curve = location.Curve
    elif isinstance(element, DB.Grid):
        curve = element.Curve
    elif isinstance(element, DB.CurveElement):
        curve = element.GeometryCurve
    else:
        return None, None
    if transform is not None:
        curve = curve.CreateTransformed(transform)
    projection = curve.Project(hit)
    if projection is None:
        return None, None
    direction = curve.ComputeDerivatives(projection.Parameter, False).BasisX
    flat = DB.XYZ(direction.X, direction.Y, 0)
    if flat.GetLength() < 1e-9:
        return None, None
    return flat.Normalize(), projection.XYZPoint


def reference_line(pick, origin, travel):
    """Direzione di riferimento in pianta: ritorna (direzione unitaria, nota). Dei due versi lungo l'elemento si tiene quello che curva al massimo 90 gradi; se l'elemento e' perpendicolare
    al tratto, il lato lo decide il punto cliccato (a sinistra o a destra del tratto)."""
    element, transform, model, hit, face_normal = pick
    category = element.Category.Name if element.Category is not None else "Element"
    label = "{} {} ({})".format(category, element.Id, model)
    tangent, on_line = element_axis(element, transform, hit)
    if tangent is None:
        # senza asse (pilastro, pannello, DirectShape IFC): lungo la faccia cliccata
        if face_normal is None:
            raise TurnError("{} has no axis and the clicked face could not be read: click a wall, beam, pipe or "
                            "duct, or a vertical face".format(label))
        flat = DB.XYZ(-face_normal.Y, face_normal.X, 0)
        if flat.GetLength() < VERTICAL_FACE_TOL:
            raise TurnError("The clicked face of {} is horizontal (floor, soffit, top of an element): it gives no "
                            "direction. Click a vertical face".format(label))
        tangent = flat.Normalize()
        note = "Parallel to the clicked face of {}".format(label)
    else:
        note = "Parallel to the axis of {}".format(label)
    forward = flat_forward(travel)
    dot = forward.DotProduct(tangent)
    if abs(dot) < PERPENDICULAR_SIN:
        side = side_of(travel, origin, hit)
        if side is None:
            raise TurnError("{} is perpendicular to the run: click it to the left or to the right of the run, on "
                            "the side to turn to".format(label))
        left = DB.XYZ(-forward.Y, forward.X, 0)
        if (tangent.DotProduct(left) > 0) != (side == "left"):
            tangent = tangent.Negate()
    elif dot < 0:
        tangent = tangent.Negate()
    return tangent, note


def right_angle_line(origin, travel):
    """Curva classica a 90 gradi: clic libero, il lato del punto rispetto al tratto (visto dall'alto) decide se si
    gira a sinistra o a destra. Ritorna (direzione unitaria, nota), o (None, "") con ESC."""
    try:
        point = uidoc.Selection.PickPoint("Turn: click a point on the side to turn to (90 degrees). "
                                          "ESC: choose an element to be parallel to instead")
    except OperationCanceledException:
        return None, ""
    except Exception as ex:
        raise TurnError("Cannot pick a point in this view (set a work plane or use a plan view): {}".format(ex))
    side = side_of(travel, origin, point)
    if side is None:
        raise TurnError("Click to the left or to the right of the run, on the side to turn to")
    forward = flat_forward(travel)
    left = DB.XYZ(-forward.Y, forward.X, 0)
    return (left if side == "left" else left.Negate()), "90 degrees to the {} (clicked point)".format(side)


def turn_direction(travel, tangent, keep_slope):
    """Direzione del nuovo tratto: la tangente di riferimento in pianta, orizzontale (predefinito) o con la stessa
    pendenza del tratto di partenza (stessa salita per unita' di percorso orizzontale)."""
    if not keep_slope:
        return tangent
    horizontal_len = math.sqrt(travel.X ** 2 + travel.Y ** 2)
    return DB.XYZ(tangent.X, tangent.Y, travel.Z / horizontal_len).Normalize()


def check_angle(signed_deg):
    if abs(signed_deg) < MIN_ANGLE:
        raise TurnError("The new direction is only {:.1f} degrees from the run: the minimum is {:g} degrees".format(
            abs(signed_deg), MIN_ANGLE))
    if abs(signed_deg) > MAX_ANGLE + 0.01:
        raise TurnError("The new direction is {:.1f} degrees from the run: more than {:g} degrees needs two "
                        "turns".format(abs(signed_deg), MAX_ANGLE))


def section_turned_copy(element, conn, signed_deg, length_mm, keep_slope):
    """Canale o passerella nuovi con la stessa sezione, larghezza e altezza non scambiate anche se posati di costa:
    Duct.Create e CableTray.Create li rifarebbero sempre "in piano" e Revit non ruota un tratto orizzontale attorno al
    proprio asse. Si copia il tratto di partenza, lo si mette in prosecuzione dall'estremo libero e lo si ruota
    attorno alla verticale passante per l'estremo: la rotazione attorno a Z conserva orientamento e pendenza."""
    travel = run_direction_at(element, conn)
    if not keep_slope:
        travel = flat_forward(travel)  # nuovo tratto orizzontale: la rotazione attorno a Z lo lascia tale
    start = conn.Origin
    copy_id = list(DB.ElementTransformUtils.CopyElement(doc, element.Id, DB.XYZ.Zero))[0]
    new_element = doc.GetElement(copy_id)
    new_element.Location.Curve = DB.Line.CreateBound(start, start.Add(travel.Multiply(length_mm / MM_PER_FT)))
    doc.Regenerate()
    axis = DB.Line.CreateBound(start, start.Add(DB.XYZ.BasisZ))
    DB.ElementTransformUtils.RotateElement(doc, copy_id, axis, math.radians(signed_deg))
    mark = new_element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
    if mark is not None and not mark.IsReadOnly:
        mark.Set("")  # la copia eredita il Mark: duplicato
    doc.Regenerate()
    return new_element


def create_segment(element, conn, end_point, level_id, signed_deg, length_mm, keep_slope):
    """Crea il nuovo tratto dal connettore a end_point (senza garantire il raccordo). Canali e passerelle sono copie
    ruotate, per non invertire larghezza e altezza."""
    type_id = element.GetTypeId()
    if isinstance(element, Pipe):
        return Pipe.Create(doc, type_id, level_id, conn, end_point)
    if isinstance(element, (Duct, CableTray)):
        return section_turned_copy(element, conn, signed_deg, length_mm, keep_slope)
    new_element = Conduit.Create(doc, type_id, conn.Origin, end_point, level_id)
    copy_conduit_diameter(element, new_element)
    return new_element


def ensure_elbow(element, conn, new_element):
    """Garantisce il gomito tra il tratto esistente e quello nuovo. Ritorna il gomito."""
    fitting = find_fitting(conn, element)
    if fitting is not None:
        return fitting
    new_conn = nearest_connector(new_element, conn.Origin)
    if new_conn.IsConnectedTo(conn):
        conn.DisconnectFrom(new_conn)
    return doc.Create.NewElbowFitting(conn, new_conn)


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
    """Isolamento (e rivestimento interno dei condotti) del tratto di partenza: [(classe, id tipo, spessore)]."""
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


def check_remaining_space(element, new_element):
    """Dopo il gomito, i due tratti devono avere ancora una lunghezza utile."""
    doc.Regenerate()
    for target, name in ((element, "existing element"), (new_element, "new run")):
        remaining = target.Location.Curve.Length * MM_PER_FT
        if remaining < MIN_REMAINING_MM:
            raise TurnError("Not enough space for the elbow: the {} would be only {:.0f} mm long "
                            "(minimum {:.0f} mm)".format(name, remaining, MIN_REMAINING_MM))



def process_pick(reference, length_mm, keep_slope):
    """Elabora un elemento. Ritorna un dict con l'esito (usato per il report), o None se l'elemento e' saltato."""
    element = doc.GetElement(reference.ElementId)
    result = {"element": element, "category": category_name(element), "status": "", "ok": False,
              "new": None, "fitting": "", "level": "", "z_mm": None, "note": "", "angle": None}

    click = reference.GlobalPoint
    if click is None:
        click = element.Location.Curve.Evaluate(0.5, True)

    conn = nearest_free_connector(element, click)
    if conn is None:
        result["status"] = "No free end (both ends are connected)"
        return result
    travel = run_direction_at(element, conn)
    if abs(travel.Z) > VERTICAL_TOL:
        result["status"] = "Vertical end: right and left are not defined"
        return result
    level_id = get_level_id(element)
    if level_id is None:
        result["status"] = "No reference level on the element"
        return result
    result["level"] = doc.GetElement(level_id).Name

    # direzione di riferimento: a ogni elemento, prima di aprire la transazione
    try:
        # predefinito: curva classica a 90 gradi dal lato del punto cliccato; con ESC si sceglie un elemento di
        # riferimento (anche in un link), con un secondo ESC l'elemento viene saltato
        tangent, result["note"] = right_angle_line(conn.Origin, travel)
        if tangent is None:
            pick = pick_alignment("Turn: click the element the new run must be parallel to, also in a linked "
                                  "model (ESC to skip this element)")
            if pick is None:
                return None
            tangent, result["note"] = reference_line(pick, conn.Origin, travel)
        signed_deg = signed_angle_to(flat_forward(travel), tangent)
        check_angle(signed_deg)
    except TurnError as ex:
        result["status"] = "Not created. {}".format(ex)
        return result
    result["angle"] = abs(signed_deg)
    side = "left" if signed_deg > 0 else "right"
    end_point = conn.Origin.Add(turn_direction(travel, tangent, keep_slope).Multiply(length_mm / MM_PER_FT))

    transaction = DB.Transaction(doc, "Turn {}".format(side))
    transaction.Start()
    try:
        new_element = create_segment(element, conn, end_point, level_id, signed_deg, length_mm, keep_slope)
        fitting = ensure_elbow(element, conn, new_element)
        check_remaining_space(element, new_element)
        inherited = inherit(element, [new_element], [fitting])
        transaction.Commit()
    except TurnError as ex:
        transaction.RollBack()
        result["status"] = "Not created. {}".format(ex)
        return result
    except Exception as ex:
        transaction.RollBack()
        result["status"] = ("Not created. The elbow could not be placed (not enough space, or no valid "
                            "fitting in the routing preferences): {}".format(ex))
        return result

    result.update({"ok": True, "status": "Created ({})".format(inherited), "new": new_element,
                   "fitting": fitting_label(fitting), "z_mm": end_point.Z * MM_PER_FT})
    return result


def fmt(value):
    return "" if value is None else "{:.1f}".format(value)


def print_report(results, settings):
    output.print_md("# Turn report")
    output.print_md("Settings: **{:g} mm, {}**".format(
        settings["length_mm"], "keep slope" if settings["keep_slope"] else "horizontal"))
    done = [r for r in results if r["ok"]]
    output.print_md("Elements picked: **{}** | runs created: **{}** | not processed: **{}**".format(
        len(results), len(done), len(results) - len(done)))
    if not results:
        output.print_md("No element was picked.")
        return

    rows = []
    for i, r in enumerate(results, 1):
        rows.append([
            i, output.linkify(r["element"].Id), r["category"], r["status"], fmt(r["angle"]),
            output.linkify(r["new"].Id) if r["new"] is not None else "", r["fitting"],
            fmt(r["z_mm"]), r["level"], r["note"],
        ])
    output.print_table(
        table_data=rows,
        columns=["#", "Element", "Category", "Status", "Angle [deg]", "New element",
                 "Fitting", "End Z, internal coordinates [mm]", "Reference level", "Direction"])


def load_settings(config):
    """Ultime impostazioni salvate: lunghezza del nuovo tratto e pendenza (0 = orizzontale, predefinito)."""
    settings = {"length_mm": DEFAULT_LENGTH_MM, "keep_slope": False}
    try:
        settings["length_mm"] = float(config.get_option("length_mm", DEFAULT_LENGTH_MM))
        keep = config.get_option("keep_slope", False)
        if not isinstance(keep, bool):
            keep = str(keep).strip().lower() in ("true", "1", "yes")
        settings["keep_slope"] = keep
    except Exception:
        pass
    return settings


def ask_settings(settings):
    """Finestra con lunghezza e pendenza (solo con SHIFT + CLICK). Ritorna le impostazioni o None se annullata."""
    return Turn_ui.TurnWindow(script.get_bundle_file(XAML_FILE_NAME), settings, MIN_LENGTH_MM).show()


def pick_loop(settings):
    results = []
    selection_filter = MepCurveFilter()
    prompt = "Turn ({:g} mm, {}): click near the free end (ESC to finish)".format(
        settings["length_mm"], "keep slope" if settings["keep_slope"] else "horizontal")
    while True:
        try:
            reference = uidoc.Selection.PickObject(ObjectType.Element, selection_filter, prompt)
        except OperationCanceledException:
            break
        result = process_pick(reference, settings["length_mm"], settings["keep_slope"])
        if result is not None:  # None: clic sul riferimento annullato con ESC, l'elemento viene saltato
            results.append(result)
    return results


def main():
    config = script.get_config(CONFIG_NAME)
    settings = load_settings(config)
    if __shiftclick__:  # noqa: F821
        settings = ask_settings(settings)
        if settings is None:
            script.exit()
        try:
            config.length_mm = settings["length_mm"]
            config.keep_slope = settings["keep_slope"]
            script.save_config()
        except Exception:
            pass

    results = pick_loop(settings)

    problems = len([r for r in results if not r["ok"]])
    if want_report(len(results), problems, "runs created"):
        print_report(results, settings)


main()
