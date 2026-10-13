# -*- coding: utf-8 -*-
__title__ = "Remove\nBranch"
__doc__ = """Version = 1.0
Date    = 13.10.2026
________________________________________________________________
Description:
Removes branches from a main run and restores the main run: pipes,
ducts, cable trays, conduits and flex, also sloped.

Instruction:
- Select one or more elements of the branches to remove (or select them
  before clicking the button), then Finish
- Each branch is followed in both directions and deleted with its
  fittings, accessories and sub-branches, up to:
  - a terminal or a piece of equipment: it stays
  - the tee (or wye) where the branch leaves a run: the tee is removed
    and that run is restored as one piece. If the two sides of the run
    have a different size (the tee was also a reduction), a transition
    is placed where the tee was, so the sizes stay as they were
  - a tap (take-off): the tap is removed, the main run is not touched
- A confirmation shows what will be deleted and restored
- A report lists the result

Several branches on the same run work, also with reductions between
them. Crosses (4-way fittings) are not handled: the branch is left
unchanged and reported.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB.Plumbing import Pipe, FlexPipe
from Autodesk.Revit.DB.Mechanical import Duct, FlexDuct
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Collections.Generic import List

from pyrevit import revit, script, DB, forms

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
SIZE_TOL = 0.5 / MM_PER_FT      # due sezioni sono uguali entro mezzo millimetro
PARALLEL_TOL = 0.999            # coseno minimo tra le direzioni dei due lati di un tratto da unire
CURVE_CLASSES = (Pipe, Duct, CableTray, Conduit, FlexPipe, FlexDuct)
STRAIGHT_CLASSES = (Pipe, Duct, CableTray, Conduit)
PHYSICAL_TYPES = (DB.ConnectorType.End, DB.ConnectorType.Curve, DB.ConnectorType.Physical)
FITTING_CATEGORIES = (DB.BuiltInCategory.OST_PipeFitting, DB.BuiltInCategory.OST_DuctFitting,
                      DB.BuiltInCategory.OST_CableTrayFitting, DB.BuiltInCategory.OST_ConduitFitting)
ACCESSORY_CATEGORIES = (DB.BuiltInCategory.OST_PipeAccessory, DB.BuiltInCategory.OST_DuctAccessory)


class RestoreError(Exception):
    """Errore atteso nel ripristino di un tratto: il messaggio va nel report."""


def get_id(element_id):
    return element_id.Value if hasattr(element_id, "Value") else element_id.IntegerValue  # Revit 2026+ / <= 2025


def category_is(element, categories):
    if element.Category is None:
        return False
    value = get_id(element.Category.Id)
    return any(value == int(c) for c in categories)


def label(element):
    category = element.Category.Name if element.Category is not None else "Element"
    return "{} {}".format(category, get_id(element.Id))


# ----------------------------------------------------------------------------------------------------------------
# Connettori
# ----------------------------------------------------------------------------------------------------------------
def connector_manager(element):
    manager = getattr(element, "ConnectorManager", None)
    if manager is None:
        model = getattr(element, "MEPModel", None)
        manager = model.ConnectorManager if model is not None else None
    return manager


def connectors(element):
    """Connettori fisici dell'elemento (estremi e, sui tratti, quelli delle prese lungo il tratto)."""
    manager = connector_manager(element)
    if manager is None:
        return []
    return [c for c in manager.Connectors if c.ConnectorType in PHYSICAL_TYPES]


def end_connectors(element):
    return [c for c in connectors(element) if c.ConnectorType == DB.ConnectorType.End]


def partners(conn):
    """Connettori fisici di altri elementi collegati a conn (escluso il connettore logico del sistema)."""
    if not conn.IsConnected:
        return []
    return [r for r in conn.AllRefs if r.Owner.Id != conn.Owner.Id and r.ConnectorType in PHYSICAL_TYPES]


def same_connector(a, b):
    return a.Owner.Id == b.Owner.Id and a.Id == b.Id


def free_connector_near(element, point):
    best, best_dist = None, None
    for c in connectors(element):
        if c.IsConnected:
            continue
        dist = c.Origin.DistanceTo(point)
        if best is None or dist < best_dist:
            best, best_dist = c, dist
    return best


def same_size(a, b):
    """Stessa sezione (forma e misure) dei due connettori."""
    try:
        if a.Shape != b.Shape:
            return False
        if a.Shape == DB.ConnectorProfileType.Round:
            return abs(a.Radius - b.Radius) < SIZE_TOL
        return abs(a.Width - b.Width) < SIZE_TOL and abs(a.Height - b.Height) < SIZE_TOL
    except Exception:
        return False


def size_text(conn):
    try:
        if conn.Shape == DB.ConnectorProfileType.Round:
            return "D{:.0f}".format(conn.Radius * 2 * MM_PER_FT)
        return "{:.0f}x{:.0f}".format(conn.Width * MM_PER_FT, conn.Height * MM_PER_FT)
    except Exception:
        return "?"


# ----------------------------------------------------------------------------------------------------------------
# Classificazione degli elementi incontrati lungo il ramo
# ----------------------------------------------------------------------------------------------------------------
def junction_roles(fitting):
    """T o braga: (connettore 1 del tratto, connettore 2 del tratto, connettore dello stacco). I due del tratto sono
    quelli opposti (stesso asse, versi contrari). None se non ci sono tre connettori con una coppia opposta;
    "cross" con quattro o piu' connettori."""
    ends = end_connectors(fitting)
    if len(ends) >= 4:
        return "cross"
    if len(ends) != 3:
        return None
    for i in range(3):
        for j in range(i + 1, 3):
            if ends[i].CoordinateSystem.BasisZ.DotProduct(ends[j].CoordinateSystem.BasisZ) < -PARALLEL_TOL:
                k = 3 - i - j
                return ends[i], ends[j], ends[k]
    return None


def is_tap(fitting):
    """Presa (tap / take-off): uno dei suoi connettori e' collegato lungo un tratto, non a un suo estremo."""
    for c in connectors(fitting):
        for p in partners(c):
            if p.ConnectorType == DB.ConnectorType.Curve:
                return True
    return False


def classify(element):
    """curve, fitting, accessory, tap, junction, cross o terminal (tutto il resto: resta)."""
    if isinstance(element, CURVE_CLASSES):
        return "curve"
    if isinstance(element, DB.FamilyInstance):
        if category_is(element, FITTING_CATEGORIES):
            if is_tap(element):
                return "tap"
            roles = junction_roles(element)
            if roles == "cross":
                return "cross"
            return "junction" if roles else "fitting"
        if category_is(element, ACCESSORY_CATEGORIES):
            return "accessory"
    return "terminal"


class BranchFilter(ISelectionFilter):
    """Tratti, raccordi semplici e accessori: non i T (sono il limite del ramo) ne' i terminali."""

    def AllowElement(self, element):
        try:
            return classify(element) in ("curve", "fitting", "accessory")
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return False


# ----------------------------------------------------------------------------------------------------------------
# Raccolta: cosa cancellare, quali T ripristinare, quali prese togliere
# ----------------------------------------------------------------------------------------------------------------
class Plan(object):
    def __init__(self):
        self.delete = []          # id da cancellare (rami, raccordi, accessori, T interni al ramo, prese)
        self.roots = []           # id dei T da cui partono i rami: tolti e tratto ripristinato
        self.taps = []            # id delle prese da cui partono i rami: tolte, il principale resta
        self.kept = []            # terminali e apparecchi raggiunti: restano
        self.problems = []        # (elemento, messaggio)
        self.seen = set()


def collect(start, plan):
    """Segue il ramo da start in tutte le direzioni. Ritorna False (e nulla aggiunto al piano) se il ramo non va
    toccato: nessun T o presa d'origine trovato, o una croce lungo il ramo."""
    local = Plan()
    local.seen = set(plan.seen)
    queue = [start]
    while queue:
        element = queue.pop()
        key = get_id(element.Id)
        if key in local.seen:
            continue
        local.seen.add(key)
        local.delete.append(element.Id)
        for conn in connectors(element):
            for ref in partners(conn):
                neighbour = ref.Owner
                n_key = get_id(neighbour.Id)
                if n_key in local.seen:
                    continue
                kind = classify(neighbour)
                if kind in ("curve", "fitting", "accessory"):
                    queue.append(neighbour)
                elif kind == "tap":
                    if conn.ConnectorType == DB.ConnectorType.Curve:
                        queue.append(neighbour)  # presa lungo un tratto del ramo: e' un sotto-ramo, va via
                    else:
                        local.seen.add(n_key)  # presa d'origine: si toglie, il principale non si tocca
                        local.taps.append(neighbour.Id)
                elif kind == "junction":
                    branch_conn = junction_roles(neighbour)[2]
                    if same_connector(branch_conn, ref):
                        local.seen.add(n_key)  # T d'origine: il ramo arriva dal suo stacco
                        local.roots.append(neighbour.Id)
                    else:
                        queue.append(neighbour)  # T interno al ramo (sotto-ramo): va via con tutto il resto
                elif kind == "cross":
                    plan.problems.append((start, "a cross ({}) is on the branch: not handled, branch left "
                                                 "unchanged".format(label(neighbour))))
                    return False
                else:
                    local.seen.add(n_key)
                    local.kept.append(neighbour.Id)
    if not local.roots and not local.taps:
        plan.problems.append((start, "no tee or tap found where the branch leaves a run: it is not a branch (or it "
                                     "is the main run itself). Nothing changed"))
        return False
    plan.seen = local.seen
    for name in ("delete", "roots", "taps", "kept"):
        getattr(plan, name).extend(getattr(local, name))
    return True


# ----------------------------------------------------------------------------------------------------------------
# Ripristino del tratto dove c'era il T
# ----------------------------------------------------------------------------------------------------------------
def endpoints(element):
    curve = element.Location.Curve
    return curve.GetEndPoint(0), curve.GetEndPoint(1)


def far_end(element, point):
    a, b = endpoints(element)
    return a if a.DistanceTo(point) > b.DistanceTo(point) else b


def move_end(element, old_point, new_point):
    """Sposta l'estremo del tratto piu' vicino a old_point in new_point, tenendo il verso della curva."""
    a, b = endpoints(element)
    if a.DistanceTo(old_point) < b.DistanceTo(old_point):
        element.Location.Curve = DB.Line.CreateBound(new_point, b)
    else:
        element.Location.Curve = DB.Line.CreateBound(a, new_point)
    doc.Regenerate()


def has_taps(element):
    """Il tratto porta delle prese lungo di se' (che sparirebbero cancellandolo)."""
    return any(c.ConnectorType == DB.ConnectorType.Curve and c.IsConnected for c in connectors(element))


def collinear(a, b):
    da = endpoints(a)[1].Subtract(endpoints(a)[0]).Normalize()
    db_ = endpoints(b)[1].Subtract(endpoints(b)[0]).Normalize()
    return abs(da.DotProduct(db_)) > PARALLEL_TOL


def tee_center(tee, conn_a, conn_b):
    """Centro del T sull'asse del tratto: il punto d'inserimento proiettato sulla retta dei due connettori."""
    a, b = conn_a.Origin, conn_b.Origin
    location = getattr(tee.Location, "Point", None)
    if location is None:
        return a.Add(b).Divide(2.0)
    axis = b.Subtract(a)
    t = location.Subtract(a).DotProduct(axis) / axis.DotProduct(axis)
    return a.Add(axis.Multiply(t))


def merge(keep, drop, center):
    """Unisce due tratti uguali e allineati: keep si allunga fino all'estremo lontano di drop, che si cancella;
    quanto era collegato all'estremo lontano di drop passa a keep."""
    target = far_end(drop, center)
    links = []
    for conn in end_connectors(drop):
        if conn.Origin.DistanceTo(target) < 1e-6:
            links = [(p.Owner.Id, p.Origin) for p in partners(conn)]
    doc.Delete(drop.Id)
    doc.Regenerate()
    move_end(keep, center, target)
    conn = free_connector_near(keep, target)
    for owner_id, origin in links:
        other = free_connector_near(doc.GetElement(owner_id), origin)
        if other is None:
            raise RestoreError("the element after the joined run could not be reconnected")
        conn.ConnectTo(other)


def join(a, b, conn_a, conn_b):
    """Collega due estremi coincidenti: riduzione se le sezioni sono diverse, altrimenti giunto o collegamento
    diretto. Ritorna la descrizione."""
    if not same_size(conn_a, conn_b):
        try:
            doc.Create.NewTransitionFitting(conn_a, conn_b)
            return "transition {} / {} placed where the tee was".format(size_text(conn_a), size_text(conn_b))
        except Exception as ex:
            raise RestoreError("the two sides have different sizes ({} / {}) and no transition could be placed: "
                               "{}".format(size_text(conn_a), size_text(conn_b), ex))
    try:
        doc.Create.NewUnionFitting(conn_a, conn_b)
        return "union placed where the tee was (the two sides are of different types)"
    except Exception:
        conn_a.ConnectTo(conn_b)
        return "connected directly where the tee was"


def restore(tee_id):
    """Toglie il T e ripristina il tratto. Ritorna la descrizione; RestoreError se non si puo'."""
    tee = doc.GetElement(tee_id)
    if tee is None:
        return "tee already removed"
    run_a, run_b, _ = junction_roles(tee)
    center = tee_center(tee, run_a, run_b)
    sides = []
    for conn in (run_a, run_b):
        refs = partners(conn)
        sides.append((refs[0].Owner, refs[0].Origin) if refs else (None, conn.Origin))
    doc.Delete(tee.Id)
    doc.Regenerate()
    (a, point_a), (b, point_b) = sides
    if a is None and b is None:
        return "tee removed (nothing was connected to the run)"
    if a is None or b is None:
        return "tee removed; one side of the run was open, nothing to join"
    a_curve, b_curve = isinstance(a, STRAIGHT_CLASSES), isinstance(b, STRAIGHT_CLASSES)
    if not a_curve and not b_curve:
        raise RestoreError("on both sides of the tee there is a fitting ({} and {}): join them by hand".format(
            label(a), label(b)))
    if a_curve and b_curve:
        if not collinear(a, b):
            raise RestoreError("the two sides of the run are not aligned")
        conn_a, conn_b = free_connector_near(a, point_a), free_connector_near(b, point_b)
        if (same_size(conn_a, conn_b) and type(a) == type(b) and a.GetTypeId() == b.GetTypeId()
                and not (has_taps(a) and has_taps(b))):
            keep, drop = (b, a) if has_taps(b) or (not has_taps(a) and
                                                   b.Location.Curve.Length > a.Location.Curve.Length) else (a, b)
            merge(keep, drop, center)
            return "run restored as one piece ({}, {})".format(label(keep), size_text(conn_a))
        move_end(a, point_a, center)
        move_end(b, point_b, center)
        return join(a, b, free_connector_near(a, center), free_connector_near(b, center))
    # da un lato un tratto, dall'altro un raccordo (es. una curva subito dopo il T): il tratto arriva al raccordo
    curve, curve_point, fitting, fitting_point = (a, point_a, b, point_b) if a_curve else (b, point_b, a, point_a)
    fitting_conn = free_connector_near(fitting, fitting_point)
    move_end(curve, curve_point, fitting_conn.Origin)
    curve_conn = free_connector_near(curve, fitting_conn.Origin)
    if same_size(curve_conn, fitting_conn):
        # es. una riduzione subito dopo il T: il tratto si ricollega direttamente al raccordo, senza giunti
        curve_conn.ConnectTo(fitting_conn)
        return "run extended and connected directly to {}".format(label(fitting))
    return join(curve, fitting, curve_conn, fitting_conn)


# ----------------------------------------------------------------------------------------------------------------
# Comando
# ----------------------------------------------------------------------------------------------------------------
def get_selection():
    """Selezione corrente (solo gli elementi validi), o selezione con clic multipli."""
    selected = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    selected = [e for e in selected if e is not None and BranchFilter().AllowElement(e)]
    if selected:
        return selected
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, BranchFilter(),
                                           "Remove Branch: select elements of the branches to remove, then Finish")
    except OperationCanceledException:
        script.exit()
    return [doc.GetElement(r.ElementId) for r in refs]


def count_problems(rows, plan):
    return len([r for r in rows if r[2] not in ("Restored", "Removed")]) + len(plan.problems)


def print_report(rows, plan, deleted):
    output.print_md("# Remove Branch report")
    output.print_md("Elements deleted: **{}** | runs restored: **{}** | taps removed: **{}** | problems: **{}**"
                    .format(deleted, len([r for r in rows if r[2] == "Restored"]), len(plan.taps),
                            count_problems(rows, plan)))
    if plan.kept:
        output.print_md("Terminals and equipment kept: {}".format(
            ", ".join(output.linkify(i) for i in plan.kept)))
    table = [[i, link, kind, status, text] for i, (link, kind, status, text) in enumerate(rows, 1)]
    for element, message in plan.problems:
        table.append([len(table) + 1, output.linkify(element.Id), "Branch", "Unchanged", message])
    if table:
        output.print_table(table_data=table, columns=["#", "Element", "Kind", "Result", "Detail"])


def main():
    selected = get_selection()
    plan = Plan()
    for element in selected:
        if get_id(element.Id) in plan.seen:
            continue  # gia' raccolto con un altro elemento dello stesso ramo
        collect(element, plan)

    if not plan.delete:
        print_report([], plan, 0)
        return

    message = ("Delete {} elements of the branches, remove {} tees restoring their runs and remove {} taps?\n"
               "Terminals and equipment at the end of the branches stay ({}).").format(
        len(plan.delete), len(plan.roots), len(plan.taps), len(plan.kept))
    if plan.problems:
        message += "\n{} selected branches will be left unchanged (see the report).".format(len(plan.problems))
    uidoc.Selection.SetElementIds(List[DB.ElementId](plan.delete))
    if not forms.alert(message, title="Remove Branch", yes=True, no=True):
        uidoc.Selection.SetElementIds(List[DB.ElementId]())
        script.exit()

    rows = []
    transaction = DB.Transaction(doc, "Remove Branch")
    transaction.Start()
    try:
        ids = [i for i in plan.delete + plan.taps if doc.GetElement(i) is not None]
        deleted = len(ids)
        doc.Delete(List[DB.ElementId](ids))
        doc.Regenerate()
        # un T alla volta, ognuno in una sotto-transazione: se uno non si ripristina resta com'era, gli altri no.
        # Ogni T legge i collegamenti attuali: dopo un'unione il tratto allungato e' gia' collegato al T successivo
        for tee_id in plan.roots:
            link = "Tee {}".format(get_id(tee_id))
            sub = DB.SubTransaction(doc)
            sub.Start()
            try:
                text = restore(tee_id)
                sub.Commit()
                rows.append((link, "Tee", "Restored", text))
            except Exception as ex:
                sub.RollBack()
                rows.append((output.linkify(tee_id), "Tee", "Tee left in place", str(ex)))
        for tap_id in plan.taps:
            rows.append(("Tap {}".format(get_id(tap_id)), "Tap", "Removed", "tap removed, main run unchanged"))
        transaction.Commit()
    except Exception:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    uidoc.Selection.SetElementIds(List[DB.ElementId]())

    if want_report(len(rows) + len(plan.problems), count_problems(rows, plan), "branches removed"):
        print_report(rows, plan, deleted)


main()
