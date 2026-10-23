# -*- coding: utf-8 -*-
"""Lunghezza del tubo inclinato di uno stacco esistente (Branch Tools, Branch length).

Curva e tratti successivi scorrono lungo il tubo inclinato (BranchTools_Place.slide_branch): cambia la
lunghezza L del tubo inclinato e con lei la quota H dello stacco, mentre braga e collettore restano fermi.
H e' lo scorrimento: fondo interno del tratto dopo la curva meno fondo interno del collettore.
Nessun documento a livello di modulo (Rocket mode).
"""
from Autodesk.Revit.DB.Plumbing import Pipe

import BranchTools_Place as fp
import BranchTools_Rotate as fr
from BranchTools_SideBranch import MM, inner_radius

MIN_LENGTH_MM = 20.0
MAX_LENGTH_MM = 2000.0


class BranchLength(object):
    """reason: None se la braga selezionata ha uno stacco che comincia con un tubo, altrimenti il motivo.
    Si ricarica con load() dopo Create branch."""

    def __init__(self, doc, uidoc, side_branch):
        self.doc = doc
        self.uidoc = uidoc
        self.side_branch = side_branch
        self.start = 0.0
        self.reason = self.load()

    @property
    def ok(self):
        return self.reason is None

    def load(self):
        side_branch = self.side_branch
        if side_branch.reason is not None:
            return side_branch.reason
        if not side_branch.branch_connected():
            return "create the branch first."
        fitting = side_branch.frame.fitting
        branch = self._branch()
        sloped = fr.connected_owner(branch, fitting.Id)
        if not isinstance(sloped, Pipe):
            return "the branch does not start with a pipe."
        try:
            chain = fp.branch_chain(fitting, branch)
        except fp.PlaceError as err:
            return str(err)
        self.sloped_id = sloped.Id
        self.rest_ids = [e.Id for e in chain if e.Id != sloped.Id and fr.connectors_of(e)]
        self.start = self.value()
        self.reason = None
        return None

    def _branch(self):
        return fr.connector_by_id(self.side_branch.frame.fitting, self.side_branch.frame.branch_id)

    def _sloped(self):
        return self.doc.GetElement(self.sloped_id)

    def value(self):
        """Lunghezza attuale del tubo inclinato, in mm."""
        return self._sloped().Location.Curve.Length / MM

    def max_value(self):
        return max(MAX_LENGTH_MM, 3.0 * self.start)

    def _next_pipe_end(self):
        """(estremita' del tratto dopo la curva vicina alla curva, raggio interno), o il tubo inclinato
        se lo stacco finisce li'."""
        sloped = self._sloped()
        near = self._branch().Origin
        corner = max((c.Origin for c in sloped.ConnectorManager.Connectors), key=lambda p: p.DistanceTo(near))
        pipes = [self.doc.GetElement(eid) for eid in self.rest_ids]
        pipes = [p for p in pipes if isinstance(p, Pipe)]
        if not pipes:
            return corner, inner_radius(sloped)

        def nearest_end(pipe):
            curve = pipe.Location.Curve
            return min((curve.GetEndPoint(0), curve.GetEndPoint(1)), key=lambda p: p.DistanceTo(corner))

        pipe = min(pipes, key=lambda p: nearest_end(p).DistanceTo(corner))
        return nearest_end(pipe), inner_radius(pipe)

    def offset(self):
        """Scorrimento attuale dello stacco (H) rispetto al collettore, in mm."""
        frame = self.side_branch.frame
        branch = self._branch()
        main_point = frame.origin + frame.axis.Multiply((branch.Origin - frame.origin).DotProduct(frame.axis))
        end, radius = self._next_pipe_end()
        main_invert = main_point.Z - inner_radius(self.side_branch.main_pipe)
        return (end.Z - radius - main_invert) / MM

    def length_for_offset(self, offset_mm):
        """Lunghezza del tubo inclinato che porta lo stacco allo scorrimento offset_mm."""
        rise = self._branch().CoordinateSystem.BasisZ.Z
        if abs(rise) < 1e-3:
            raise fp.PlaceError("The branch is horizontal: its length does not change the elevation.")
        return self.value() + (offset_mm - self.offset()) / rise

    def set_to(self, length_mm):
        """Ritorna (lunghezza effettiva in mm, messaggio d'errore o None)."""
        along = (length_mm - self.value()) * MM
        if abs(along) < 1e-6:
            return self.value(), None
        branch = self._branch()
        error = fr.run_step(self.doc, self.uidoc, "Set Branch Length",
                            lambda: fp.slide_branch(self.doc, branch, self._sloped(), self.rest_ids, along))
        return self.value(), error
