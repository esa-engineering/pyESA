# -*- coding: utf-8 -*-
"""Funzioni su un tubo selezionato come collettore (Branch Tools, sezione Pipe).

- place_wye: braga inserita nel tubo, con rotazione e angolo dati (0 = derivazione orizzontale a destra
  guardando verso valle, 90 = in alto). Prima dalle routing preferences (NewTeeFitting con un tubo di
  derivazione provvisorio); se Revit la rifiuta, la famiglia della braga inserita e orientata a mano.
- connect: un altro tubo collegato al collettore con una braga a 45 gradi rispetto al collettore,
  "da sopra" (angolo sopra il collettore, discesa a 45 nel suo piano verticale verso valle) o
  "di lato" (dall'estremita' del tubo, inclinazione che ne risulta). Il collettore non cambia quota.
Verso valle: in discesa se il collettore e' in pendenza, altrimenti il verso del tubo; reverse lo inverte.
Nessun documento a livello di modulo (Rocket mode).
"""
import math

from Autodesk.Revit.DB import (BuiltInCategory, BuiltInParameter, ConnectorType, ElementTransformUtils,
                               FamilySymbol, FilteredElementCollector, Line, PartType,
                               RoutingPreferenceRuleGroupType, SubTransaction, XYZ)
from Autodesk.Revit.DB.Plumbing import Pipe, PlumbingUtils
from Autodesk.Revit.DB.Structure import StructuralType

import BranchTools_Align as fa
import BranchTools_Place as fp
import BranchTools_Rotate as fr
from BranchTools_SideBranch import end_at, set_diameter

LEVEL_TOLERANCE = 1e-4      # pendenza (componente Z della direzione) sotto cui il collettore e' piano
END_MARGIN = 0.1            # piedi: distanza minima della braga dalle estremita' del collettore


class PipeError(Exception):
    pass


def pipe_ids(pipe):
    """(tipo di sistema, tipo di tubo, livello) del tubo, per crearne altri uguali."""
    system_id = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
    level = pipe.ReferenceLevel
    return system_id, pipe.GetTypeId(), level.Id if level is not None else pipe.LevelId


def insert_junction(doc, main_pipe, point, flow, branch_dir, angle_deg, branch_diameter, symbol=None):
    """Braga nel collettore con il centro nel punto e la derivazione libera nella direzione branch_dir.
    symbol: tipo di braga scelto dall'utente, inserito e orientato direttamente. Senza scelta, prima le routing
    preferences del tipo di tubo e, se Revit le rifiuta, la famiglia con Part Type Wye. Ritorna la braga.
    Va chiamata in una transazione."""
    if symbol is None:
        fitting = _junction_from_preferences(doc, main_pipe, point, branch_dir, branch_diameter)
        if fitting is not None:
            return fitting
        symbol = wye_symbol(doc, main_pipe)
        if symbol is None:
            raise PipeError("Revit refused the junction of the routing preferences and no family has Part Type "
                            "Wye. Choose the wye family in the Pipe section.")
    return _junction_from_family(doc, main_pipe, point, flow, branch_dir, angle_deg, branch_diameter, symbol)


def symbol_label(symbol):
    return u"{} : {}".format(symbol.Family.Name, symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())


BRANCH_PART_TYPES = set(int(t) for t in (PartType.Tee, PartType.Cross, PartType.Wye,
                                          PartType.LateralTee, PartType.LateralCross))


def pipe_fitting_symbols(doc):
    """[(etichetta "Famiglia : Tipo", tipo)] dei raccordi di tubazione a T, a croce o a braga del progetto
    (Part Type Tee, Cross, Wye, LateralTee, LateralCross), in ordine alfabetico."""
    symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).OfClass(FamilySymbol)
    return sorted(((symbol_label(s), s) for s in symbols if _part_type(s) in BRANCH_PART_TYPES),
                  key=lambda item: item[0].lower())


def _junction_from_preferences(doc, main_pipe, point, branch_dir, branch_diameter):
    """NewTeeFitting con un tubo di derivazione provvisorio, poi eliminato. In una sotto-transazione: se Revit
    rifiuta ("angle between them is too small or too large", visto con le braghe degli scarichi ESA) il modello
    torna com'era e si ritorna None."""
    sub = SubTransaction(doc)
    sub.Start()
    try:
        system_id, type_id, level_id = pipe_ids(main_pipe)
        length = max(1.0, 4.0 * main_pipe.Diameter)
        branch_pipe = Pipe.Create(doc, system_id, type_id, level_id, point, point + branch_dir.Multiply(length))
        set_diameter(branch_pipe, branch_diameter)
        doc.Regenerate()
        piece = doc.GetElement(PlumbingUtils.BreakCurve(doc, main_pipe.Id, point))
        doc.Regenerate()
        first, second = end_at(main_pipe, point), end_at(piece, point)
        branch = end_at(branch_pipe, point)
        fitting = None
        # l'angolo e' 45 gradi rispetto a un connettore del collettore e 135 rispetto all'altro: si provano
        # i due ordini
        for main_a, main_b in ((first, second), (second, first)):
            try:
                fitting = doc.Create.NewTeeFitting(main_a, main_b, branch)
                break
            except Exception:
                pass
        if fitting is None:
            sub.RollBack()
            return None
        doc.Delete(branch_pipe.Id)     # la derivazione resta libera
        sub.Commit()
        return fitting
    except Exception:
        if sub.HasStarted() and not sub.HasEnded():
            sub.RollBack()
        return None


def _part_type(symbol):
    param = symbol.Family.get_Parameter(BuiltInParameter.FAMILY_CONTENT_PART_TYPE)
    return param.AsInteger() if param is not None else None


def wye_symbol(doc, pipe):
    """Tipo di braga: la giunzione delle routing preferences del tipo di tubo se e' una braga (Part Type Wye),
    altrimenti la prima famiglia di raccordo con Part Type Wye caricata nel progetto, o None."""
    wye = int(PartType.Wye)
    try:
        manager = doc.GetElement(pipe.GetTypeId()).RoutingPreferenceManager
        group = RoutingPreferenceRuleGroupType.Junctions
        for index in range(manager.GetNumberOfRules(group)):
            symbol = doc.GetElement(manager.GetRule(group, index).MEPPartId)
            if isinstance(symbol, FamilySymbol) and _part_type(symbol) == wye:
                return symbol
    except Exception:
        pass
    for symbol in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).OfClass(FamilySymbol):
        if _part_type(symbol) == wye:
            return symbol
    return None


def _set_radius(connector, radius):
    try:
        connector.Radius = radius
    except Exception:
        pass    # misura guidata da una formula della famiglia: resta quella


def _rotate_between(doc, element_id, origin, v_from, v_to, hint):
    """Ruota l'elemento attorno a origin in modo che v_from vada su v_to. hint: vettore non parallelo a v_from,
    per l'asse del mezzo giro quando v_from e v_to sono opposti."""
    cross = v_from.CrossProduct(v_to)
    sine = cross.GetLength()
    cosine = v_from.DotProduct(v_to)
    if sine < 1e-9:
        if cosine > 0:
            return
        axis = (hint - v_from.Multiply(hint.DotProduct(v_from))).Normalize()
        angle = math.pi
    else:
        axis = cross.Multiply(1.0 / sine)
        angle = math.atan2(sine, cosine)
    ElementTransformUtils.RotateElement(doc, element_id, Line.CreateUnbound(origin, axis), angle)


def _closest_on_first(p1, d1, p2, d2):
    """Punto della retta (p1, d1) piu' vicino alla retta (p2, d2); direzioni unitarie, non parallele."""
    w0 = p1 - p2
    b = d1.DotProduct(d2)
    d = d1.DotProduct(w0)
    e = d2.DotProduct(w0)
    return p1 + d1.Multiply((b * e - d) / (1.0 - b * b))


def _junction_from_family(doc, main_pipe, point, flow, branch_dir, angle_deg, branch_diameter, symbol):
    """Famiglia della braga inserita nel punto, con Angle e misure, orientata (collettore sull'asse del tubo,
    derivazione in branch_dir), poi il collettore spezzato e collegato come in Place copies."""
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()
    fitting = doc.Create.NewFamilyInstance(point, symbol, StructuralType.NonStructural)
    param = fr.angle_parameter(fitting)
    if param is not None:
        param.Set(math.radians(angle_deg))
    doc.Regenerate()
    frame = fr.build_frame(fitting)
    if frame.main_ids is None or frame.branch_id is None:
        raise PipeError(u"{} is not a wye: it has no main and branch connectors.".format(symbol_label(symbol)))
    for connector_id in frame.main_ids:
        _set_radius(fr.connector_by_id(fitting, connector_id), main_pipe.Diameter / 2.0)
    _set_radius(fr.connector_by_id(fitting, frame.branch_id), branch_diameter / 2.0)
    doc.Regenerate()

    def read():
        first, second = [fr.connector_by_id(fitting, cid) for cid in frame.main_ids]
        branch = fr.connector_by_id(fitting, frame.branch_id)
        axis = (second.Origin - first.Origin).Normalize()
        out = branch.CoordinateSystem.BasisZ
        # verso dell'asse come il flusso: la derivazione sta verso monte rispetto all'asse come branch_dir
        # rispetto al flusso
        if (axis.DotProduct(out) < 0) != (flow.DotProduct(branch_dir) < 0):
            axis = axis.Negate()
        return first, branch, axis, out

    origin = fitting.Location.Point
    _, _, axis, out = read()
    _rotate_between(doc, fitting.Id, origin, axis, flow, out)
    doc.Regenerate()
    _, _, axis, out = read()
    # attorno al collettore, finche' la derivazione e' nella direzione voluta
    source = (out - flow.Multiply(out.DotProduct(flow))).Normalize()
    target = (branch_dir - flow.Multiply(branch_dir.DotProduct(flow))).Normalize()
    turn = math.atan2(source.CrossProduct(target).DotProduct(flow), source.DotProduct(target))
    if abs(turn) > 1e-9:
        ElementTransformUtils.RotateElement(doc, fitting.Id, Line.CreateUnbound(origin, flow), turn)
        doc.Regenerate()
    first, branch, axis, out = read()
    # centro della braga (incrocio tra asse del collettore e asse della derivazione) sul punto
    center = _closest_on_first(first.Origin, axis, branch.Origin, out)
    ElementTransformUtils.MoveElement(doc, fitting.Id, point - center)
    doc.Regenerate()
    ends = [fr.connector_by_id(fitting, cid).Origin for cid in frame.main_ids]
    if not all(fp._contains(main_pipe, end) for end in ends):
        raise PipeError("The wye does not fit on the main here: too close to the end of the pipe.")
    fp.split_and_connect(doc, main_pipe, fitting, ends)
    return fitting


class MainPipe(object):
    """Il tubo selezionato come collettore. reason: None se le funzioni si possono usare, altrimenti il motivo.
    point: punto di inserimento della braga (dove si e' cliccato il tubo, o la mezzeria)."""

    def __init__(self, doc, uidoc, element, picked_point):
        self.doc = doc
        self.uidoc = uidoc
        self.pipe = element if isinstance(element, Pipe) else None
        self.reverse = False
        self.symbol_id = None       # tipo di braga scelto nella finestra; None = routing preferences
        self.point = None
        self.from_click = picked_point is not None
        self.reason = self._check(picked_point)

    @property
    def ok(self):
        return self.reason is None

    def _check(self, picked_point):
        if self.pipe is None:
            return "select a pipe."
        curve = self.pipe.Location.Curve
        if not isinstance(curve, Line):
            return "the pipe is not straight."
        start, finish = curve.GetEndPoint(0), curve.GetEndPoint(1)
        direction = (finish - start).Normalize()
        if abs(direction.Z) > 0.99:
            return "the pipe is vertical."
        if picked_point is None:
            self.point = (start + finish).Multiply(0.5)
        else:
            self.point = start + direction.Multiply((picked_point - start).DotProduct(direction))
        return None

    def directions(self):
        """(f verso valle, U alto perpendicolare a f, R destra guardando verso valle, testo per la riga di stato)."""
        curve = self.pipe.Location.Curve
        d = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
        if abs(d.Z) > LEVEL_TOLERANCE:
            f = d if d.Z < 0 else d.Negate()
            text = "downhill"
        else:
            f = d
            text = "pipe direction (level main)"
        if self.reverse:
            f = f.Negate()
            text += ", reversed"
        up = (XYZ.BasisZ - f.Multiply(XYZ.BasisZ.DotProduct(f))).Normalize()
        return f, up, f.CrossProduct(up), text

    def _symbol(self):
        return self.doc.GetElement(self.symbol_id) if self.symbol_id is not None else None

    def _station_ok(self, point):
        curve = self.pipe.Location.Curve
        start, finish = curve.GetEndPoint(0), curve.GetEndPoint(1)
        along = (point - start).DotProduct((finish - start).Normalize())
        return END_MARGIN < along < start.DistanceTo(finish) - END_MARGIN

    # -------------------------------------------------------------- braga

    def place_wye(self, roll_deg, angle_deg):
        """Inserisce la braga nel punto. Ritorna (braga o None, messaggio d'errore o None)."""
        if not self.ok:
            return None, self.reason
        if not self._station_ok(self.point):
            return None, "the point is too close to the end of the pipe."
        f, up, right, _ = self.directions()
        roll, angle = math.radians(roll_deg), math.radians(angle_deg)
        # derivazione verso monte: e' la direzione da cui arriva l'acqua nella braga
        branch_dir = f.Negate().Multiply(math.cos(angle)) + \
            (right.Multiply(math.cos(roll)) + up.Multiply(math.sin(roll))).Multiply(math.sin(angle))
        doc = self.doc
        main = self.pipe
        point = self.point
        result = {}

        def action():
            fitting = insert_junction(doc, main, point, f, branch_dir, angle_deg, main.Diameter, self._symbol())
            result["id"] = fitting.Id

        error = fr.run_step(doc, self.uidoc, "Place Wye", action)
        if error:
            return None, error
        return doc.GetElement(result["id"]), None

    # -------------------------------------------------------------- collegamento di un altro tubo

    def connect(self, other, picked_point, from_above):
        """Collega il tubo other al collettore con una braga a 45 gradi. picked_point: dove si e' cliccato other,
        per scegliere l'estremita' libera piu' vicina. Ritorna None se riuscito, altrimenti il messaggio."""
        if not self.ok:
            return self.reason
        if not isinstance(other, Pipe) or other.Id == self.pipe.Id:
            return "pick another pipe."
        ends = [c for c in other.ConnectorManager.Connectors
                if c.ConnectorType == ConnectorType.End and not c.IsConnected]
        if not ends:
            return "the picked pipe has no open end."
        end = min(ends, key=lambda c: c.Origin.DistanceTo(picked_point))
        f, up, right, _ = self.directions()
        origin = self.pipe.Location.Curve.GetEndPoint(0)
        try:
            corner = self._above_corner(end, origin, f) if from_above else end.Origin
        except PipeError as err:
            return str(err)
        # braga a 45 gradi rispetto al collettore: a valle della proiezione dell'angolo, di quanto l'angolo
        # dista dall'asse del collettore (tan 45 = 1)
        offset = corner - origin
        along = offset.DotProduct(f)
        across = (offset - f.Multiply(along)).GetLength()
        if across < 1e-3:
            return "the pipe end is on the axis of the main."
        junction = origin + f.Multiply(along + across)
        if not self._station_ok(junction):
            return "the wye would fall outside the main pipe: move the pipe or extend the main."
        doc = self.doc
        main = self.pipe

        def action():
            if from_above:
                fa.stretch(other, end, corner)
                doc.Regenerate()
            direction = (corner - junction).Normalize()
            fitting = insert_junction(doc, main, junction, f, direction, 45.0, other.Diameter, self._symbol())
            doc.Regenerate()
            # tubo dalla derivazione della braga (libera) all'angolo, poi curva sull'altro tubo
            frame = fr.build_frame(fitting)
            wye_branch = fr.connector_by_id(fitting, frame.branch_id) if frame.branch_id is not None else None
            if wye_branch is None or wye_branch.IsConnected:
                raise PipeError("The wye has no free branch.")
            start = wye_branch.Origin
            if (corner - start).DotProduct(direction) < 0.05:
                raise PipeError("The pipe is too close to the main for a wye: move it further away.")
            system_id, type_id, level_id = pipe_ids(other)
            branch = Pipe.Create(doc, system_id, type_id, level_id, start, corner)
            set_diameter(branch, other.Diameter)
            doc.Regenerate()
            end_at(branch, start).ConnectTo(wye_branch)
            branch_end = end_at(branch, corner)
            other_end = end_at(other, corner)
            if branch_end.CoordinateSystem.BasisZ.DotProduct(other_end.CoordinateSystem.BasisZ) < -0.9999:
                branch_end.ConnectTo(other_end)     # gia' in linea: nessuna curva
            else:
                doc.Create.NewElbowFitting(branch_end, other_end)

        return fr.run_step(doc, self.uidoc, "Connect Pipe", action)

    def _above_corner(self, end, origin, f):
        """Punto dell'asse dell'altro tubo che in pianta sta sull'asse del collettore."""
        e = end.Origin
        n = end.CoordinateSystem.BasisZ
        det = n.Y * f.X - n.X * f.Y
        if abs(det) < 1e-6:
            raise PipeError("the pipe runs parallel to the main in plan: use From the side.")
        dx, dy = origin.X - e.X, origin.Y - e.Y
        t = (dy * f.X - dx * f.Y) / det
        corner = e + n.Multiply(t)
        main_point = origin + f.Multiply((corner - origin).DotProduct(f))
        clearance = self.pipe.Diameter / 2.0 + end.Radius
        if corner.Z - main_point.Z < clearance:
            raise PipeError("the pipe is not above the main.")
        return corner
