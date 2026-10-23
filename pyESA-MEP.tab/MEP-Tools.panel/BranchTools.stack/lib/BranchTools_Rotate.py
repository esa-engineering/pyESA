# -*- coding: utf-8 -*-
"""Rotazione dei raccordi e parametro Angle (Branch Tools).

La rotazione e' assoluta: si misura l'angolo della derivazione attorno all'asse del raccordo
(0 = orizzontale, 90 = verso l'alto, -90 = verso il basso, sul lato verso cui punta all'apertura).
L'asse e il riferimento sono fissati all'apertura della finestra (FittingFrame).
Nessun documento a livello di modulo (Rocket mode): tutto arriva come argomento.
"""
import math

from Autodesk.Revit.DB import (BuiltInCategory, ConnectorType, ElementTransformUtils,
                               FailureProcessingResult, FailureSeverity, IFailuresPreprocessor, Line,
                               StorageType, Transaction, TransactionStatus, XYZ)

TOLERANCE = 1e-6            # piedi, per allineamento e collinearita'
VERTICAL_LIMIT = 0.99       # |asse . Z| oltre il quale l'asse e' considerato verticale
ANGLE_PARAMETER_NAMES = ("Angle", "Angolo")
PHYSICAL_TYPES = (ConnectorType.End, ConnectorType.Curve, ConnectorType.Physical)
FITTING_CATEGORIES = (BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_DuctFitting,
                      BuiltInCategory.OST_CableTrayFitting, BuiltInCategory.OST_ConduitFitting)


def get_element_id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


FITTING_CATEGORY_IDS = set(int(c) for c in FITTING_CATEGORIES)
PIPE_FITTING_ID = int(BuiltInCategory.OST_PipeFitting)


def is_fitting(element):
    category = element.Category if element is not None else None
    return category is not None and get_element_id_value(category.Id) in FITTING_CATEGORY_IDS


def is_pipe_fitting(element):
    return element.Category is not None and get_element_id_value(element.Category.Id) == PIPE_FITTING_ID


def wrap_degrees(angle):
    """Angolo riportato in [-180, 180)."""
    return (angle + 180.0) % 360.0 - 180.0


# ------------------------------------------------------------------ connettori

def end_connectors(fitting):
    """Connettori di estremita' ordinati per Id, cosi' l'ordine e' lo stesso a ogni lettura."""
    try:
        manager = fitting.MEPModel.ConnectorManager
    except Exception:
        return []
    if manager is None:
        return []
    return sorted([c for c in manager.Connectors if c.ConnectorType == ConnectorType.End], key=lambda c: c.Id)


def connector_by_id(fitting, connector_id):
    for connector in end_connectors(fitting):
        if connector.Id == connector_id:
            return connector
    return None


def connectors_of(element):
    """Connettori di un tubo o di un raccordo (lista vuota se l'elemento non ne ha)."""
    try:
        manager = element.ConnectorManager if hasattr(element, "ConnectorManager") else element.MEPModel.ConnectorManager
    except Exception:
        return []
    return list(manager.Connectors) if manager is not None else []


def connected_owner(connector, owner_id):
    """Elemento collegato fisicamente al connettore (i riferimenti logici non contano), o None."""
    if not connector.IsConnected:
        return None
    for ref in connector.AllRefs:
        if ref.Owner.Id != owner_id and ref.ConnectorType in PHYSICAL_TYPES:
            return ref.Owner
    return None


def _on_line(connector, origin, direction):
    """True se il connettore sta sulla retta (origin, direction) ed e' parallelo ad essa."""
    if direction.CrossProduct(connector.CoordinateSystem.BasisZ).GetLength() > TOLERANCE:
        return False
    return (connector.Origin - origin).CrossProduct(direction).GetLength() < TOLERANCE * 100


def _main_pair(connectors):
    """Coppia di connettori allineati e opposti (collettore di T e Y, raccordi diritti), o None."""
    for i, first in enumerate(connectors):
        direction = first.CoordinateSystem.BasisZ
        for second in connectors[i + 1:]:
            if direction.DotProduct(second.CoordinateSystem.BasisZ) < -1 + TOLERANCE and \
                    _on_line(second, first.Origin, direction):
                return first, second
    return None


# ------------------------------------------------------------------ asse e lettura dell'angolo

class FittingFrame(object):
    """Asse di rotazione e riferimento di un raccordo, fissati all'apertura della finestra.
    axis/h/u: terna destra (h orizzontale, u verso l'alto, axis = h x u); reference_id: connettore
    la cui direzione si misura (derivazione o estremita' libera), None = asse X del raccordo;
    branch_id: connettore fuori asse (derivazione di T e Y), None per curve e raccordi diritti."""

    def __init__(self, fitting, reason=None):
        self.fitting = fitting
        self.reason = reason
        self.origin = self.axis = self.h = self.u = None
        self.reference_id = self.branch_id = None
        self.main_ids = None        # Id dei due connettori del collettore, se ci sono

    @property
    def ok(self):
        return self.reason is None

    def reference_direction(self):
        """Direzione misurata: il connettore di riferimento, altrimenti il primo asse del raccordo
        non parallelo all'asse di rotazione (un riferimento parallelo darebbe sempre 0)."""
        if self.reference_id is not None:
            connector = connector_by_id(self.fitting, self.reference_id)
            if connector is not None:
                return connector.CoordinateSystem.BasisZ
        transform = self.fitting.GetTransform()
        for candidate in (transform.BasisX, transform.BasisY, transform.BasisZ):
            if candidate.CrossProduct(self.axis).GetLength() > 0.1:
                return candidate
        return transform.BasisX

    def angle(self):
        """Angolo attuale della derivazione attorno all'asse, in gradi."""
        r = self.reference_direction()
        return math.degrees(math.atan2(r.DotProduct(self.u), r.DotProduct(self.h)))


def _perpendicular_frame(axis):
    """(h, u) perpendicolari all'asse: u e' la parte di Z (o di Y, con asse verticale) normale all'asse."""
    up = XYZ.BasisZ if abs(axis.DotProduct(XYZ.BasisZ)) < VERTICAL_LIMIT else XYZ.BasisY
    u = (up - axis.Multiply(up.DotProduct(axis))).Normalize()
    return u.CrossProduct(axis), u


def build_frame(fitting, zero_side=None):
    """Asse: il collettore se c'e' una coppia di connettori allineati, altrimenti l'estremita' collegata
    (o il primo connettore se libero). Non ruotabile (reason) se pinned o collegato fuori dall'asse:
    asse e riferimenti restano calcolati, servono a Place copies per una braga con lo stacco collegato."""
    connectors = end_connectors(fitting)
    if not connectors:
        return FittingFrame(fitting, "no connectors")
    connected = [c for c in connectors if connected_owner(c, fitting.Id) is not None]
    pair = _main_pair(connectors)
    axis_connector = pair[0] if pair else (connected[0] if connected else connectors[0])
    origin = axis_connector.Origin
    axis = axis_connector.CoordinateSystem.BasisZ.Normalize()
    reason = None
    if fitting.Pinned:
        reason = "pinned"
    elif any(not _on_line(other, origin, axis) for other in connected):
        reason = "connected on more than one axis"

    frame = FittingFrame(fitting, reason)
    frame.origin = origin
    if pair:
        frame.main_ids = (pair[0].Id, pair[1].Id)
    off_axis = [c for c in connectors if not _on_line(c, origin, axis)]
    if off_axis:
        frame.reference_id = off_axis[0].Id
        if pair:
            frame.branch_id = off_axis[0].Id
    frame.axis = axis
    frame.h, frame.u = _perpendicular_frame(axis)
    # lato dello 0: quello verso cui punta la derivazione all'apertura (angolo iniziale tra -90 e 90), oppure
    # zero_side se dato (braga appena inserita da un tubo: 0 = a destra guardando verso valle);
    # se h punta dall'altra parte si gira l'asse, e con lui h
    side = zero_side if zero_side is not None else frame.reference_direction()
    if side.DotProduct(frame.h) < 0:
        frame.axis = axis.Negate()
        frame.h, frame.u = _perpendicular_frame(frame.axis)
    return frame


# ------------------------------------------------------------------ transazioni dei passi

class _QuietFailures(IFailuresPreprocessor):
    """Anteprima dello slider: niente finestre di avviso a ogni passo, un errore annulla il passo."""

    def PreprocessFailures(self, accessor):
        for message in accessor.GetFailureMessages():
            if message.GetSeverity() == FailureSeverity.Warning:
                accessor.DeleteWarning(message)
            else:
                return FailureProcessingResult.ProceedWithRollBack
        return FailureProcessingResult.Continue


def run_step(doc, uidoc, name, action):
    """Esegue action() in una transazione silenziosa (avvisi soppressi, un errore annulla tutto il passo).
    Ritorna None se riuscito, altrimenti il messaggio d'errore. Va usata dentro un TransactionGroup."""
    t = Transaction(doc, name)
    options = t.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(_QuietFailures())
    t.SetFailureHandlingOptions(options)
    t.Start()
    try:
        action()
    except Exception as err:
        t.RollBack()
        return str(err)
    if t.Commit() != TransactionStatus.Committed:
        return "Revit refused the change."
    try:
        uidoc.RefreshActiveView()
    except Exception:
        pass
    return None


class Rotator(object):
    """Porta tutti i raccordi ruotabili allo stesso angolo assoluto (gradi, vedi FittingFrame.angle)."""

    def __init__(self, doc, uidoc, fittings, zero_side=None):
        self.doc = doc
        self.uidoc = uidoc
        self.frames = [build_frame(f, zero_side) for f in fittings]
        self.rotatable = [f for f in self.frames if f.ok]
        self.skipped = [f.reason for f in self.frames if not f.ok]
        self.start = self.value()

    def value(self):
        """Angolo attuale del primo raccordo ruotabile (quello mostrato dallo slider)."""
        return self.rotatable[0].angle() if self.rotatable else 0.0

    def rotate_to(self, angle):
        """Ritorna (angolo effettivo, messaggio d'errore o None)."""
        moves = []
        for frame in self.rotatable:
            delta = wrap_degrees(angle - frame.angle())
            if abs(delta) > 1e-9:
                moves.append((frame, delta))
        if not moves:
            return self.value(), None

        def rotate():
            for frame, delta in moves:
                line = Line.CreateUnbound(frame.origin, frame.axis)
                ElementTransformUtils.RotateElement(self.doc, frame.fitting.Id, line, math.radians(delta))

        error = run_step(self.doc, self.uidoc, "Rotate Fitting", rotate)
        return self.value(), error


# ------------------------------------------------------------------ parametro Angle

def angle_parameter(fitting):
    """Parametro di istanza dell'angolo del raccordo (es. derivazione a Y), o None se manca o non e' modificabile."""
    for name in ANGLE_PARAMETER_NAMES:
        param = fitting.LookupParameter(name)
        if param is not None and not param.IsReadOnly and param.StorageType == StorageType.Double:
            return param
    return None


class AngleSetter(object):
    """Imposta il parametro Angle (gradi, valore assoluto) su tutti i raccordi che lo hanno."""

    def __init__(self, doc, uidoc, fittings):
        self.doc = doc
        self.uidoc = uidoc
        self.fittings = [f for f in fittings if angle_parameter(f) is not None]
        self.start_values = [self.current(f) for f in self.fittings]
        self.start = self.value()

    def current(self, fitting):
        return math.degrees(angle_parameter(fitting).AsDouble())

    def value(self):
        """Valore attuale del primo raccordo (quello mostrato dallo slider)."""
        return self.current(self.fittings[0]) if self.fittings else 0.0

    def set_to(self, angle):
        """Ritorna (valore effettivo dopo la rigenerazione, messaggio d'errore o None)."""
        if not self.fittings:
            return 0.0, None

        def set_angle():
            for fitting in self.fittings:
                if not angle_parameter(fitting).Set(math.radians(angle)):
                    raise ValueError("The family does not accept {:g} degrees.".format(angle))

        error = run_step(self.doc, self.uidoc, "Set Fitting Angle", set_angle)
        return self.value(), error
