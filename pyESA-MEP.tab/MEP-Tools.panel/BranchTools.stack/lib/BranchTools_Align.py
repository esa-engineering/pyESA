# -*- coding: utf-8 -*-
"""Allineamento e collegamento di due estremita' libere (Branch Tools, Align).

L'elemento selezionato (tubo, canale, passerella, tubo protettivo o raccordo) viene spostato
perpendicolarmente all'asse di un'estremita' libera dell'elemento scelto dall'utente (run_with; senza
scelta, il piu' vicino entro 2 m), parallela e rivolta verso la sua, finche' le due estremita' sono in asse. Poi la distanza lungo l'asse si chiude (tubo allungato, o
raccordo spostato) e le due estremita' si collegano; tra tubi di misura diversa Revit mette una riduzione.
Nessun documento a livello di modulo (Rocket mode).
"""
from Autodesk.Revit.DB import (BoundingBoxIntersectsFilter, BuiltInCategory, ConnectorProfileType,
                               ConnectorType, ElementMulticategoryFilter, ElementTransformUtils,
                               FilteredElementCollector, Line, LogicalAndFilter, MEPCurve, Outline, XYZ)
from System.Collections.Generic import List

import BranchTools_Rotate as fr

SEARCH_FEET = 2000.0 / 304.8    # raggio di ricerca dell'altra estremita': 2 m
PARALLEL_TOLERANCE = 1e-3       # |n1 x n2| sotto cui due estremita' sono parallele
MM = 1.0 / 304.8
POINT_TOLERANCE = 1e-4          # piedi, per estremita' coincidenti
ALIGNED = "the ends are already aligned."

ALIGN_CATEGORIES = (BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_PipeFitting,
                    BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_DuctAccessory,
                    BuiltInCategory.OST_CableTray, BuiltInCategory.OST_CableTrayFitting,
                    BuiltInCategory.OST_Conduit, BuiltInCategory.OST_ConduitFitting)


def is_alignable(element):
    """Tubi, canali, passerelle, tubi protettivi e raccordi: gli elementi che la finestra accetta."""
    return isinstance(element, MEPCurve) or fr.is_fitting(element)


def open_ends(element):
    return [c for c in fr.connectors_of(element) if c.ConnectorType == ConnectorType.End and not c.IsConnected]


class Aligner(object):
    """reason: None se c'e' una coppia di estremita' da allineare, altrimenti il motivo (riga di stato).
    shift: spostamento dell'elemento selezionato che mette in asse le due estremita'."""

    def __init__(self, doc, uidoc, element):
        self.doc = doc
        self.uidoc = uidoc
        self.element = element
        self.target = None          # elemento scelto dall'utente a cui allinearsi (None = il piu' vicino)
        self.shift = None
        self.pair = None            # (Id connettore mio, Id elemento dell'altro, Id connettore dell'altro)
        self.distance = 0.0
        self.reason = None if open_ends(element) else "the selected element has no open end."

    def run_with(self, target):
        """Allinea e collega l'elemento all'elemento target scelto dall'utente.
        Ritorna None se riuscito, altrimenti il messaggio per l'avviso."""
        if fr.get_element_id_value(target.Id) == fr.get_element_id_value(self.element.Id):
            return "Pick another element, not the selected one."
        self.target = target
        self.reason = self.find()
        if not self.can_run():
            return "Not aligned: {}".format(self.reason)
        error = self.align()
        if error:
            return "Not aligned: {}".format(error)
        error = self.connect()
        if error:
            return "Aligned, not connected: {}".format(error)
        return None

    @property
    def ok(self):
        return self.reason is None

    def _candidates(self):
        """Elementi MEP vicini (entro SEARCH_FEET dal riquadro dell'elemento), escluso l'elemento stesso."""
        box = self.element.get_BoundingBox(None)
        if box is None:
            return []
        pad = XYZ(SEARCH_FEET, SEARCH_FEET, SEARCH_FEET)
        categories = List[BuiltInCategory]()
        for category in ALIGN_CATEGORIES:
            categories.Add(category)
        element_filter = LogicalAndFilter(ElementMulticategoryFilter(categories),
                                          BoundingBoxIntersectsFilter(Outline(box.Min - pad, box.Max + pad)))
        own = fr.get_element_id_value(self.element.Id)
        return [e for e in FilteredElementCollector(self.doc).WherePasses(element_filter).WhereElementIsNotElementType()
                if fr.get_element_id_value(e.Id) != own]

    def find(self):
        """Cerca la coppia piu' vicina: un'estremita' libera dell'elemento e una libera di un altro elemento,
        parallele e rivolte una verso l'altra, dello stesso dominio (tubazioni, canali...)."""
        self.shift = None
        mine = open_ends(self.element)
        if not mine:
            return "the selected element has no open end."
        best = None
        candidates = [self.target] if self.target is not None else self._candidates()
        for other_element in candidates:
            for other in open_ends(other_element):
                for end in mine:
                    if end.Domain != other.Domain:
                        continue
                    n1 = end.CoordinateSystem.BasisZ
                    n2 = other.CoordinateSystem.BasisZ
                    if n1.CrossProduct(n2).GetLength() > PARALLEL_TOLERANCE or n1.DotProduct(n2) > 0:
                        continue
                    # rivolte una verso l'altra: l'altra estremita' sta davanti alla mia
                    if (other.Origin - end.Origin).DotProduct(n1) < -1e-6:
                        continue
                    distance = end.Origin.DistanceTo(other.Origin)
                    # con l'elemento scelto dall'utente niente limite di distanza
                    if (self.target is not None or distance <= SEARCH_FEET) and \
                            (best is None or distance < best[0]):
                        best = (distance, end, other)
        if best is None:
            if self.target is not None:
                return "the picked element has no open end parallel to and facing this one."
            return "no open end facing this one within 2 m."
        distance, end, other = best
        # la coppia si ricorda per Id: dopo lo spostamento gli oggetti Connector vanno riletti
        self.pair = (end.Id, other.Owner.Id, other.Id)
        gap = other.Origin - end.Origin
        axis = other.CoordinateSystem.BasisZ
        self.shift = gap - axis.Multiply(gap.DotProduct(axis))
        self.distance = distance
        if self.shift.GetLength() < 1e-6:
            return ALIGNED
        return None

    def can_run(self):
        """True se c'e' una coppia da allineare o gia' allineata da collegare."""
        return self.reason is None or self.reason == ALIGNED

    def describe(self):
        if not self.ok:
            return "Not available: {}".format(self.reason)
        return ""

    def align(self):
        """Sposta l'elemento selezionato. Ritorna None se riuscito, altrimenti il messaggio d'errore."""
        if self.reason == ALIGNED:
            return None
        if not self.ok:
            return self.reason
        shift = self.shift
        error = fr.run_step(self.doc, self.uidoc, "Align Ends",
                            lambda: ElementTransformUtils.MoveElement(self.doc, self.element.Id, shift))
        self.reason = self.find()
        return error

    def _pair_connectors(self):
        end_id, other_owner_id, other_id = self.pair
        end = [c for c in fr.connectors_of(self.element) if c.Id == end_id][0]
        other_owner = self.doc.GetElement(other_owner_id)
        other = [c for c in fr.connectors_of(other_owner) if c.Id == other_id][0]
        return end, other

    def connect(self):
        """Chiude la distanza lungo l'asse e collega le due estremita' allineate. Ritorna None se riuscito,
        altrimenti il messaggio d'errore. Passo separato dall'allineamento: se fallisce, l'allineamento resta."""
        if self.reason != ALIGNED:
            return self.reason or "the ends are not in line."
        doc = self.doc

        def action():
            end, other = self._pair_connectors()
            gap = other.Origin - end.Origin
            curves = isinstance(end.Owner, MEPCurve) and isinstance(other.Owner, MEPCurve)
            if not same_size(end, other):
                if curves and gap.GetLength() > POINT_TOLERANCE:
                    # due tubi di misura diversa: la riduzione la sceglie Revit dalle routing preferences
                    doc.Create.NewTransitionFitting(end, other)
                    return
                raise AlignError("the two ends have different sizes.")
            if gap.GetLength() > POINT_TOLERANCE:
                if isinstance(end.Owner, MEPCurve):
                    stretch(end.Owner, end, other.Origin)
                elif isinstance(other.Owner, MEPCurve):
                    stretch(other.Owner, other, end.Origin)
                else:
                    ElementTransformUtils.MoveElement(doc, self.element.Id, gap)
                doc.Regenerate()
                end, other = self._pair_connectors()
            end.ConnectTo(other)

        error = fr.run_step(doc, self.uidoc, "Connect Ends", action)
        self.reason = self.find()
        return error


class AlignError(Exception):
    pass


def same_size(first, second):
    """Stessa forma e stessa misura (raggio, o larghezza e altezza)."""
    try:
        if first.Shape != second.Shape:
            return False
        if first.Shape == ConnectorProfileType.Round:
            return abs(first.Radius - second.Radius) < 1e-6
        return abs(first.Width - second.Width) < 1e-6 and abs(first.Height - second.Height) < 1e-6
    except Exception:
        return True


def stretch(curve_element, connector, target):
    """Sposta in target l'estremita' del tubo (o canale, passerella...) che sta sul connettore."""
    line = curve_element.Location.Curve
    start, finish = line.GetEndPoint(0), line.GetEndPoint(1)
    if start.DistanceTo(connector.Origin) < finish.DistanceTo(connector.Origin):
        curve_element.Location.Curve = Line.CreateBound(target, finish)
    else:
        curve_element.Location.Curve = Line.CreateBound(start, target)
