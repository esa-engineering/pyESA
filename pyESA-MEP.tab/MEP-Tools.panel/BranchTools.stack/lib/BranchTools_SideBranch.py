# -*- coding: utf-8 -*-
"""Stacco laterale da una braga (Branch Tools).

Di norma la braga resta con la rotazione impostata dall'utente. Con una pendenza scelta, invece, la
braga e' ruotata attorno al collettore finche' la derivazione sale (o scende) con quell'angolo. Il
tratto orizzontale va perpendicolare al collettore (predefinito) o nella direzione in pianta della
derivazione; la curva prende l'angolo tra le due direzioni. Poi tubo inclinato fino al piano della quota voluta, curva e tubo orizzontale. La curva sta
nel punto in cui l'asse della derivazione incontra quel piano, quindi la quota del tratto orizzontale
e' esatta qualunque sia la misura dei raccordi. La quota e' data come scorrimento: differenza tra il
fondo interno del collettore e quello dello stacco.
Nessun documento a livello di modulo (Rocket mode).
"""
import math

from Autodesk.Revit.DB import BuiltInParameter, Transaction, XYZ
from Autodesk.Revit.DB.Plumbing import Pipe

import BranchTools_Rotate as fr

MM = 1.0 / 304.8            # da millimetri a piedi
ELBOW_ANGLES = (15, 22.5, 30, 45)
FROM_ROTATION = 0           # valore salvato per "curva dalla rotazione impostata"


class SideBranchError(Exception):
    pass


def inner_radius(pipe):
    param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM)
    if param is not None and param.AsDouble() > 0:
        return param.AsDouble() / 2.0
    return pipe.Diameter / 2.0


def set_diameter(pipe, diameter):
    pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter)


def end_at(pipe, point):
    """Connettore del tubo piu' vicino al punto."""
    return min(pipe.ConnectorManager.Connectors, key=lambda c: c.Origin.DistanceTo(point))


class SideBranch(object):
    """Stacco laterale dal raccordo dell'unico FittingFrame del Rotator.
    reason: None se la braga e' adatta (Y o T su un tubo collettore), altrimenti il motivo.
    create_reason: None se si puo' creare lo stacco (derivazione libera, braga ruotabile);
    place_reason(): None se si puo' copiare lo stacco (derivazione collegata). Testi per la riga di stato."""

    def __init__(self, doc, uidoc, rotator):
        self.doc = doc
        self.uidoc = uidoc
        self.frame = None
        self.main_pipe = None
        self.reason = self._check(rotator)
        self.create_reason = self.reason
        if self.reason is None:
            if self.branch_connected():
                self.create_reason = "the branch is already connected."
            elif not self.frame.ok:
                self.create_reason = "the fitting cannot be rotated ({}).".format(self.frame.reason)

    @property
    def ok(self):
        """True se lo stacco si puo' creare."""
        return self.create_reason is None

    def _check(self, rotator):
        if len(rotator.frames) != 1:
            return "select one wye or tee only."
        frame = rotator.frames[0]
        if not fr.is_pipe_fitting(frame.fitting):
            return "pipe fittings only."
        if frame.axis is None or frame.branch_id is None or frame.main_ids is None:
            return "the fitting has no branch (wye or tee)."
        for connector_id in frame.main_ids:
            owner = fr.connected_owner(fr.connector_by_id(frame.fitting, connector_id), frame.fitting.Id)
            if isinstance(owner, Pipe):
                self.main_pipe = owner
                break
        if self.main_pipe is None:
            return "the fitting is not connected to a main pipe."
        self.frame = frame
        return None

    def branch_connected(self):
        return fr.connected_owner(self._branch(), self.frame.fitting.Id) is not None

    def place_reason(self):
        if self.reason is not None:
            return self.reason
        if not self.branch_connected():
            return "create the branch first."
        return None

    # -------------------------------------------------------------- dati del collettore

    def _branch(self):
        return fr.connector_by_id(self.frame.fitting, self.frame.branch_id)

    def _pipe_ids(self):
        main = self.main_pipe
        system_id = main.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
        level = main.ReferenceLevel
        return system_id, main.GetTypeId(), level.Id if level is not None else main.LevelId

    def _probe_inner_radius(self, diameter):
        """Raggio interno di un tubo del tipo del collettore con quel diametro: tubo di prova, poi annullato.
        Va chiamata fuori da una transazione."""
        system_id, type_id, level_id = self._pipe_ids()
        t = Transaction(self.doc, "Probe Pipe Size")
        t.Start()
        try:
            pipe = Pipe.Create(self.doc, system_id, type_id, level_id, XYZ(0, 0, 0), XYZ(1, 0, 0))
            set_diameter(pipe, diameter)
            self.doc.Regenerate()
            return inner_radius(pipe)
        finally:
            t.RollBack()

    def target_z(self, offset_mm):
        """(quota dell'asse del tratto orizzontale, quota dell'asse del collettore in corrispondenza della braga).
        Non cambia ruotando la braga: la proiezione sull'asse del collettore resta la stessa."""
        branch = self._branch()
        axis, origin = self.frame.axis, self.frame.origin
        main_point = origin + axis.Multiply((branch.Origin - origin).DotProduct(axis))
        branch_inner = self._probe_inner_radius(2.0 * branch.Radius)
        z = main_point.Z - inner_radius(self.main_pipe) + offset_mm * MM + branch_inner
        return z, main_point.Z

    def target_angle(self, offset_mm, elbow_deg):
        """Angolo di rotazione (convenzione di FittingFrame.angle) che inclina la derivazione di elbow_deg
        verso l'alto se lo stacco sta sopra il collettore, verso il basso se sta sotto."""
        target, main_z = self.target_z(offset_mm)
        if abs(target - main_z) < 1e-6:
            raise SideBranchError("The branch axis would be at the level of the main: no slope needed.")
        elevation = math.radians(elbow_deg) * (1.0 if target > main_z else -1.0)
        frame = self.frame
        r = self._branch().CoordinateSystem.BasisZ
        along = r.DotProduct(frame.axis)
        across = (r - frame.axis.Multiply(along)).GetLength()
        if across < 1e-6 or abs(frame.u.Z) < 1e-6:
            raise SideBranchError("Side branches need a main that is not vertical.")
        # direzione della derivazione ruotata di phi: along * axis + across * (cos phi * h + sin phi * u)
        sine = (math.sin(elevation) - along * frame.axis.Z) / (across * frame.u.Z)
        if abs(sine) > 1:
            raise SideBranchError("A {:g} degree slope is steeper than this branch can reach. "
                                  "Use a smaller slope or a larger branch angle.".format(elbow_deg))
        return math.degrees(math.asin(sine))

    def elevation(self):
        """Inclinazione attuale della derivazione sull'orizzontale, in gradi (+ verso l'alto)."""
        z = self._branch().CoordinateSystem.BasisZ.Z
        return math.degrees(math.asin(max(-1.0, min(1.0, z))))

    def run_direction(self, direction, perpendicular):
        """Direzione del tratto orizzontale: in pianta come la derivazione, oppure perpendicolare al
        collettore sul lato verso cui va la derivazione."""
        flat = XYZ(direction.X, direction.Y, 0)
        if flat.GetLength() < 1e-6:
            raise SideBranchError("The branch is vertical: no horizontal direction.")
        if not perpendicular:
            return flat.Normalize()
        side = XYZ.BasisZ.CrossProduct(self.frame.axis)
        if side.GetLength() < 1e-6:
            raise SideBranchError("Side branches need a main that is not vertical.")
        side = side.Normalize()
        across = flat.DotProduct(side)
        if abs(across) < 1e-3 * flat.GetLength():
            raise SideBranchError("The branch runs along the main: rotate it to one side first.")
        return side if across > 0 else side.Negate()

    def elbow_angle(self, perpendicular):
        """Angolo della curva tra la derivazione nella posizione attuale e il tratto orizzontale, in gradi."""
        direction = self._branch().CoordinateSystem.BasisZ
        run = self.run_direction(direction, perpendicular)
        return math.degrees(math.acos(max(-1.0, min(1.0, direction.DotProduct(run)))))

    # -------------------------------------------------------------- creazione

    def create(self, offset_mm, length_mm, perpendicular):
        """Crea tubo inclinato, curva e tubo orizzontale dalla derivazione nella posizione attuale
        (la rotazione non viene cambiata). perpendicular: tratto orizzontale perpendicolare al collettore
        invece che nella direzione in pianta della derivazione.
        Ritorna None se riuscito, altrimenti il messaggio d'errore."""
        try:
            target, main_z = self.target_z(offset_mm)
        except Exception as err:
            return str(err)
        system_id, type_id, level_id = self._pipe_ids()
        doc = self.doc

        def build():
            branch = self._branch()
            start = branch.Origin
            direction = branch.CoordinateSystem.BasisZ
            if abs(direction.Z) < 1e-6:
                raise SideBranchError("The branch is horizontal: rotate it up or down first.")
            if (target > main_z) != (direction.Z > 0):
                raise SideBranchError("The branch points {} but the offset is {}: rotate the wye {} "
                                      "or change the sign of the offset.".format(
                                          "down" if direction.Z < 0 else "up",
                                          "up" if target > main_z else "down",
                                          "up" if target > main_z else "down"))
            distance = (target - start.Z) / direction.Z
            if distance <= 0:
                raise SideBranchError("The branch end is already past the branch level. "
                                      "Use a larger offset or a smaller rotation.")
            corner = start + direction.Multiply(distance)
            end = corner + self.run_direction(direction, perpendicular).Multiply(length_mm * MM)
            diameter = 2.0 * branch.Radius
            sloped_pipe = Pipe.Create(doc, system_id, type_id, level_id, start, corner)
            flat_pipe = Pipe.Create(doc, system_id, type_id, level_id, corner, end)
            set_diameter(sloped_pipe, diameter)
            set_diameter(flat_pipe, diameter)
            doc.Regenerate()
            end_at(sloped_pipe, start).ConnectTo(branch)
            doc.Create.NewElbowFitting(end_at(sloped_pipe, corner), end_at(flat_pipe, corner))

        return fr.run_step(doc, self.uidoc, "Create Side Branch", build)
