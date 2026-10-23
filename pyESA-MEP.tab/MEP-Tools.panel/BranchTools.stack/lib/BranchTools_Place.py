# -*- coding: utf-8 -*-
"""Copie di uno stacco laterale lungo lo stesso collettore (Branch Tools, Place copies).

La braga e lo stacco collegato alla derivazione sono copiati con una traslazione lungo l'asse del
collettore fino al punto cliccato: rotazione e direzioni restano identiche. Con il collettore in pendenza
la copia di norma torna alla quota dell'originale (curva e tratti dopo il tubo inclinato spostati lungo
quel tubo); altrimenti tiene l'offset dal collettore. Il collettore cliccato viene spezzato alle estremita' della nuova braga, il
pezzo in mezzo eliminato e le due parti collegate alla braga.
Nessun documento a livello di modulo (Rocket mode).
"""
from Autodesk.Revit.DB import ElementId, ElementTransformUtils, FamilyInstance, Line, LocationPoint
from Autodesk.Revit.DB.Plumbing import Pipe, PlumbingUtils
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from System.Collections.Generic import List

import BranchTools_Rotate as fr

MAX_CHAIN = 30              # elementi dello stacco: oltre, la derivazione porta a una rete, non a uno stacco
POINT_TOLERANCE = 1e-4      # piedi, per estremita' coincidenti
LINE_TOLERANCE = 0.01       # piedi (circa 3 mm), per riconoscere i tubi sullo stesso asse
END_MARGIN = 0.01           # piedi: spazio minimo tra la nuova braga e la fine del tubo cliccato
MIN_PIPE_LENGTH = 0.05      # piedi (circa 15 mm): tubo inclinato minimo quando la copia tiene la quota
MM_FEET = 1.0 / 304.8       # un millimetro in piedi, per i messaggi


class PlaceError(Exception):
    pass


def branch_chain(fitting, branch_connector):
    """Elementi collegati alla derivazione (tubi, curve, raccordi dello stacco), senza la braga."""
    start = fr.connected_owner(branch_connector, fitting.Id)
    seen = set([fr.get_element_id_value(fitting.Id)])
    chain = []
    todo = [start]
    while todo:
        element = todo.pop()
        if element is None or fr.get_element_id_value(element.Id) in seen:
            continue
        seen.add(fr.get_element_id_value(element.Id))
        chain.append(element)
        if len(chain) > MAX_CHAIN:
            raise PlaceError("The branch is connected to other runs: only a free side branch can be copied.")
        for connector in fr.connectors_of(element):
            todo.append(fr.connected_owner(connector, element.Id))
    return chain


def slide_branch(doc, branch, sloped, rest_ids, along):
    """Sposta lungo il tubo inclinato (direzione della derivazione) gli elementi rest_ids, cioe' curva e tratti
    successivi, di along piedi: il tubo inclinato si allunga (along > 0) o si accorcia, pendenza e direzioni
    restano, la quota cambia di along * d.Z. branch: connettore della braga; sloped: primo tubo dello stacco.
    Va chiamata in una transazione."""
    direction = branch.CoordinateSystem.BasisZ
    near_point = branch.Origin
    far = max(sloped.ConnectorManager.Connectors, key=lambda c: c.Origin.DistanceTo(near_point))
    far_point = far.Origin + direction.Multiply(along)
    if (far_point - near_point).DotProduct(direction) < MIN_PIPE_LENGTH:
        raise PlaceError("The sloped pipe would be too short.")
    if rest_ids:
        ElementTransformUtils.MoveElements(doc, List[ElementId](rest_ids), direction.Multiply(along))
    doc.Regenerate()
    # di norma Revit allunga il tubo collegato alla curva spostata; se non lo fa, lo si allunga qui
    far = max(sloped.ConnectorManager.Connectors, key=lambda c: c.Origin.DistanceTo(near_point))
    if far.Origin.DistanceTo(far_point) > POINT_TOLERANCE:
        sloped.Location.Curve = Line.CreateBound(near_point, far_point)
        doc.Regenerate()
        far = max(sloped.ConnectorManager.Connectors, key=lambda c: c.Origin.DistanceTo(near_point))
        if not far.IsConnected:
            for eid in rest_ids:
                for other in fr.connectors_of(doc.GetElement(eid)):
                    if not other.IsConnected and other.Origin.DistanceTo(far.Origin) < POINT_TOLERANCE:
                        far.ConnectTo(other)


def _on_axis_line(point, origin, axis):
    return (point - origin).CrossProduct(axis).GetLength() < LINE_TOLERANCE


class MainPipeFilter(ISelectionFilter):
    """Solo tubi sullo stesso asse del collettore della braga."""

    def __init__(self, origin, axis):
        self.origin = origin
        self.axis = axis

    def AllowElement(self, element):
        if not isinstance(element, Pipe):
            return False
        try:
            curve = element.Location.Curve
        except Exception:
            return False
        return (_on_axis_line(curve.GetEndPoint(0), self.origin, self.axis) and
                _on_axis_line(curve.GetEndPoint(1), self.origin, self.axis))

    def AllowReference(self, reference, point):
        return True


class BranchCopier(object):
    """Copia della braga del SideBranch e del suo stacco in un punto dello stesso collettore."""

    def __init__(self, doc, uidoc, side_branch):
        self.doc = doc
        self.uidoc = uidoc
        frame = side_branch.frame
        self.fitting = frame.fitting
        self.origin = frame.origin
        self.axis = frame.axis
        self.main_points = [fr.connector_by_id(self.fitting, cid).Origin for cid in frame.main_ids]
        self.branch_id = frame.branch_id
        branch = fr.connector_by_id(self.fitting, frame.branch_id)
        chain = branch_chain(self.fitting, branch)
        self.source_ids = [self.fitting.Id] + [e.Id for e in chain]
        self.location = self.fitting.Location.Point
        self.stage = ""             # fase dell'ultima copia, riportata nei messaggi d'errore

    def _station(self, point):
        """Ascissa del punto lungo l'asse del collettore."""
        return (point - self.origin).DotProduct(self.axis)

    def shift_to(self, point):
        """Traslazione che porta il centro della braga sulla proiezione del punto sull'asse."""
        center = sum(self._station(p) for p in self.main_points) / 2.0
        return self.axis.Multiply(self._station(point) - center)

    def place(self, pipe, shift, keep_level):
        """Copia braga e stacco traslati di shift e li inserisce nel tubo. Va chiamata in una transazione.
        keep_level: lo stacco copiato torna alla quota dell'originale (vedi _keep_level); altrimenti tiene
        l'offset dal collettore, che in pendenza cambia quota da una copia all'altra."""
        doc = self.doc
        self.stage = "check"
        ends = [p + shift for p in self.main_points]
        curve = pipe.Location.Curve
        start, finish = curve.GetEndPoint(0), curve.GetEndPoint(1)
        length = start.DistanceTo(finish)
        direction = (finish - start).Normalize()
        for point in ends:
            along = (point - start).DotProduct(direction)
            if along < END_MARGIN or along > length - END_MARGIN:
                raise PlaceError("The new wye does not fit on this pipe: too close to its end or to a fitting.")

        self.stage = "copy"
        new_ids = ElementTransformUtils.CopyElements(doc, List[ElementId](self.source_ids), shift)
        target = self.location + shift
        new_fitting = None
        for eid in new_ids:
            element = doc.GetElement(eid)
            if isinstance(element, FamilyInstance) and isinstance(element.Location, LocationPoint) and \
                    element.Location.Point.DistanceTo(target) < POINT_TOLERANCE:
                new_fitting = element
                break
        if new_fitting is None:
            raise PlaceError("The copied wye was not found.")
        self._connect_copy(new_fitting, new_ids)
        if keep_level:
            self.stage = "level"
            self._keep_level(new_fitting, new_ids, shift.Z)

        self.stage = "split"
        split_and_connect(doc, pipe, new_fitting, ends)
        return new_fitting

    def _keep_level(self, new_fitting, new_ids, drop):
        """Riporta lo stacco copiato alla quota dell'originale. La braga sta piu' in alto o in basso di drop
        (collettore in pendenza): curva e tratti successivi si spostano lungo il tubo inclinato di quanto
        serve a compensarlo, il tubo inclinato si allunga o si accorcia, pendenza e direzioni restano."""
        if abs(drop) < 1e-6:
            return
        doc = self.doc
        branch = fr.connector_by_id(new_fitting, self.branch_id)
        direction = branch.CoordinateSystem.BasisZ
        if abs(direction.Z) < 1e-3:
            raise PlaceError("The branch is horizontal, so it cannot keep the source elevation. "
                             "Untick Copies keep this elevation.")
        sloped = fr.connected_owner(branch, new_fitting.Id)
        if not isinstance(sloped, Pipe):
            raise PlaceError("The branch does not start with a pipe, so it cannot keep the source elevation. "
                             "Untick Copies keep this elevation.")
        along = -drop / direction.Z
        length = sloped.Location.Curve.Length
        if length + along < MIN_PIPE_LENGTH:
            # a monte sul collettore in pendenza la copia nasce piu' in alto: per tornare alla quota il tubo
            # inclinato si accorcia, e oltre una certa distanza non ne resta abbastanza
            raise PlaceError("Here the sloped pipe must be {:.0f} mm shorter to keep the elevation, but it is only "
                             "{:.0f} mm long. Click closer to the source, lengthen the branch first, or untick "
                             "Copies keep this elevation.".format(-along / MM_FEET, length / MM_FEET))
        rest = [eid for eid in new_ids
                if eid != new_fitting.Id and eid != sloped.Id and fr.connectors_of(doc.GetElement(eid))]
        slide_branch(doc, branch, sloped, rest, along)

    def _connect_copy(self, new_fitting, new_ids):
        """Se la copia non ha tenuto il collegamento tra braga e stacco, lo rifa' sui connettori coincidenti."""
        others = [self.doc.GetElement(eid) for eid in new_ids if eid != new_fitting.Id]
        for connector in fr.end_connectors(new_fitting):
            if connector.IsConnected:
                continue
            for element in others:
                for other in fr.connectors_of(element):
                    if not other.IsConnected and other.Origin.DistanceTo(connector.Origin) < POINT_TOLERANCE:
                        other.ConnectTo(connector)



def _contains(pipe, point):
    if pipe is None or not pipe.IsValidObject:
        return False
    curve = pipe.Location.Curve
    start, finish = curve.GetEndPoint(0), curve.GetEndPoint(1)
    length = start.DistanceTo(finish)
    along = (point - start).DotProduct((finish - start).Normalize())
    return POINT_TOLERANCE < along < length - POINT_TOLERANCE


def _has_end(pipe, point):
    if pipe is None or not pipe.IsValidObject:
        return False
    curve = pipe.Location.Curve
    return min(curve.GetEndPoint(0).DistanceTo(point), curve.GetEndPoint(1).DistanceTo(point)) < POINT_TOLERANCE


def split_and_connect(doc, pipe, fitting, ends):
    """Inserisce nel tubo un raccordo gia' posato sul suo asse: due tagli nei punti ends (le estremita' del
    collettore del raccordo), via il pezzo in mezzo, le due parti collegate ai connettori liberi coincidenti
    del raccordo. Va chiamata in una transazione."""
    # gli oggetti si rileggono per Id: un oggetto eliminato non va piu' toccato (InvalidObjectException)
    pipe_id = pipe.Id
    piece_ids = [pipe_id, PlumbingUtils.BreakCurve(doc, pipe_id, ends[0])]
    holder = [eid for eid in piece_ids if _contains(doc.GetElement(eid), ends[1])]
    if not holder:
        raise PlaceError("The main pipe could not be split.")
    piece_ids.append(PlumbingUtils.BreakCurve(doc, holder[0], ends[1]))
    middle = [eid for eid in piece_ids
              if _has_end(doc.GetElement(eid), ends[0]) and _has_end(doc.GetElement(eid), ends[1])]
    if len(middle) != 1:
        raise PlaceError("The main pipe could not be split.")
    middle_value = fr.get_element_id_value(middle[0])
    doc.Delete(middle[0])
    pieces = [doc.GetElement(eid) for eid in piece_ids if fr.get_element_id_value(eid) != middle_value]
    pieces = [p for p in pieces if p is not None and p.IsValidObject]
    for connector in fr.end_connectors(fitting):
        if connector.IsConnected:
            continue
        for piece in pieces:
            for end in piece.ConnectorManager.Connectors:
                if not end.IsConnected and end.Origin.DistanceTo(connector.Origin) < POINT_TOLERANCE:
                    end.ConnectTo(connector)


def place_copies(doc, uidoc, side_branch, keep_level):
    """Ciclo di clic sul collettore: una copia per clic, Esc termina. keep_level: vedi BranchCopier.place.
    Ritorna (copie riuscite, lista dei motivi delle copie non riuscite)."""
    try:
        copier = BranchCopier(doc, uidoc, side_branch)
    except PlaceError as err:
        return 0, [str(err)]
    selection_filter = MainPipeFilter(copier.origin, copier.axis)
    placed = 0
    failures = []
    while True:
        try:
            reference = uidoc.Selection.PickObject(ObjectType.PointOnElement, selection_filter,
                                                   "Click the main where the next branch goes. Esc to finish.")
        except OperationCanceledException:
            break
        pipe = doc.GetElement(reference.ElementId)
        shift = copier.shift_to(reference.GlobalPoint)
        error = fr.run_step(doc, uidoc, "Place Side Branch", lambda: copier.place(pipe, shift, keep_level))
        if error:
            failures.append("{} (step: {})".format(error, copier.stage))
        else:
            placed += 1
    return placed, failures
