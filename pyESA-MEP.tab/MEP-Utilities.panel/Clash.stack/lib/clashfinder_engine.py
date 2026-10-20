# -*- coding: utf-8 -*-
"""
clashfinder_engine.py - risoluzione degli insiemi e ricerca delle interferenze.

Stesso criterio dell'Interference Check di Revit: intersezione dei solidi.

Ogni insieme mette insieme elementi del modello aperto (dalla sorgente scelta:
selezione caricata, vista attiva, intero modello o nulla) e, a scelta, elementi dei
modelli linkati. Ogni elemento e' un Item che porta la sua ModelSource (documento e
trasformazione del link). Le coppie si provano cosi':

  host  x host   ElementIntersectsElementFilter (nativo, il caso veloce)
  link  x host   ElementIntersectsSolidFilter con i solidi del link portati nelle
                 coordinate del modello aperto, sugli elementi host candidati
  host  x link   come sopra, a parti invertite, sul solo elemento host
  link  x link   BooleanOperationsUtils.Intersect fra i solidi trasformati

Prima del test sui solidi c'e' sempre il filtro sul bounding box: nativo
(BoundingBoxIntersectsFilter) verso gli elementi host, una griglia in Python
(BoxGrid) verso quelli linkati.

Gli id di documenti diversi si sovrappongono: chiavi e cache usano (indice della
sorgente, id). Nessuna transazione: il modello non viene modificato.
"""

import io
import json
import os
import re
import time

from System.Collections.Generic import List

from pyrevit import DB

import clashfinder_rules as cr

SOURCE_SELECTION = cr.SOURCE_SELECTION
SOURCE_VIEW = cr.SOURCE_VIEW
SOURCE_MODEL = cr.SOURCE_MODEL
SOURCE_NONE = cr.SOURCE_NONE
SOURCE_LABELS = [(SOURCE_SELECTION, "Current selection"),
                 (SOURCE_VIEW, "Active view"),
                 (SOURCE_MODEL, "Whole model"),
                 (SOURCE_NONE, "Nothing (links only)")]

# Volume minimo dell'intersezione fra solidi linkati (piedi cubi, ~0.3 cm3).
MIN_VOLUME_FT3 = 1e-5
# Lato delle celle della griglia dei bounding box linkati (piedi, ~3 m).
GRID_CELL_FT = 10.0
GRID_MAX_CELLS = 512
# Fino a questo numero di elementi linkati in Set B il filtro sui box e' un confronto
# diretto con tutti (niente griglia): piu' lento ma senza approssimazioni.
BRUTE_FORCE_LINKED = 5000


class ClashCancelled(Exception):
    pass


# =============================================================================
# SORGENTI (modello aperto e link)
# =============================================================================

class ModelSource(object):
    """Documento da cui vengono gli elementi. index 0 = modello aperto."""

    def __init__(self, index, doc, instance=None, label=None):
        self.index = index
        self.doc = doc
        self.instance = instance
        self.label = label or doc.Title
        self.transform = None
        self.key = "host"
        if instance is not None:
            self.key = instance.UniqueId
            try:
                transform = instance.GetTotalTransform()
                if not transform.IsIdentity:
                    self.transform = transform
            except Exception:
                pass

    @property
    def is_link(self):
        return self.instance is not None

    def to_host(self, point):
        return self.transform.OfPoint(point) if self.transform is not None else point


def list_links(doc):
    """ModelSource dei link caricati (istanze dirette del modello aperto), per nome."""
    found = []
    collector = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)
    for instance in collector:
        try:
            link_doc = instance.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue
        found.append((instance, link_doc))
    collector.Dispose()

    # Piu' istanze dello stesso file: si distinguono con il numero d'ordine.
    titles = {}
    for _, link_doc in found:
        titles[link_doc.Title] = titles.get(link_doc.Title, 0) + 1
    seen = {}
    sources = []
    for instance, link_doc in sorted(found, key=lambda pair: pair[1].Title.lower()):
        label = link_doc.Title
        if titles[label] > 1:
            seen[label] = seen.get(label, 0) + 1
            label = u"{} #{}".format(label, seen[label])
        sources.append(ModelSource(len(sources) + 1, link_doc, instance, label))
    return sources


class RunScope(object):
    """Da dove vengono Set A e Set B: sorgente nel modello aperto e link scelti.

    links_a / links_b sono le key (UniqueId delle istanze) dei link caricati;
    missing_a / missing_b i nomi dei link chiesti dal test ma non caricati.
    """

    def __init__(self, source_a, source_b, links_a=(), links_b=(),
                 missing_a=(), missing_b=()):
        self.source_a = source_a
        self.source_b = source_b
        self.links_a = list(links_a)
        self.links_b = list(links_b)
        self.missing_a = list(missing_a)
        self.missing_b = list(missing_b)


def link_title(source):
    try:
        return source.doc.Title
    except Exception:
        return source.label


def scope_of(test, links):
    """RunScope del test: i link sono salvati per nome del file e valgono per tutte le
    istanze caricate di quel file."""
    parts = []
    for set_key in ("set_a", "set_b"):
        wanted = test[set_key]["links"]
        keys = [l.key for l in links if link_title(l) in wanted]
        titles = set(link_title(l) for l in links)
        parts.append((test[set_key]["source"], keys, [w for w in wanted if w not in titles]))
    (source_a, keys_a, missing_a), (source_b, keys_b, missing_b) = parts
    return RunScope(source_a, source_b, keys_a, keys_b, missing_a, missing_b)


# =============================================================================
# ELEMENTI
# =============================================================================

def _box_of(element, source):
    """(min, max) del bounding box nelle coordinate del modello aperto, o None."""
    try:
        box = element.get_BoundingBox(None)
    except Exception:
        box = None
    if box is None:
        return None
    if source.transform is None:
        return ((box.Min.X, box.Min.Y, box.Min.Z), (box.Max.X, box.Max.Y, box.Max.Z))
    xs, ys, zs = [], [], []
    for x in (box.Min.X, box.Max.X):
        for y in (box.Min.Y, box.Max.Y):
            for z in (box.Min.Z, box.Max.Z):
                point = source.transform.OfPoint(DB.XYZ(x, y, z))
                xs.append(point.X)
                ys.append(point.Y)
                zs.append(point.Z)
    return ((min(xs), min(ys), min(zs)), (max(xs), max(ys), max(zs)))


def _collect_solids(geometry, solids):
    if geometry is None:
        return
    for item in geometry:
        if isinstance(item, DB.Solid):
            try:
                if item.Volume > 1e-9:
                    solids.append(item)
            except Exception:
                pass
        elif isinstance(item, DB.GeometryInstance):
            _collect_solids(item.GetInstanceGeometry(), solids)


def element_solids(element, transform):
    """Solidi dell'elemento (dettaglio Fine), nelle coordinate del modello aperto."""
    options = DB.Options()
    options.DetailLevel = DB.ViewDetailLevel.Fine
    options.ComputeReferences = False
    options.IncludeNonVisibleObjects = False
    solids = []
    try:
        _collect_solids(element.get_Geometry(options), solids)
    except Exception:
        return []
    if transform is None:
        return solids
    moved = []
    for solid in solids:
        try:
            moved.append(DB.SolidUtils.CreateTransformed(solid, transform))
        except Exception:
            pass
    return moved


class Item(object):
    """Un elemento di un insieme, con la sua sorgente. Box e solidi calcolati al bisogno."""

    __slots__ = ("source", "element", "value", "key", "_box", "_solids")

    def __init__(self, source, element):
        self.source = source
        self.element = element
        self.value = cr.get_element_id_value(element.Id)
        self.key = (source.index, self.value)
        self._box = False
        self._solids = None

    @property
    def is_link(self):
        return self.source.is_link

    def box(self):
        if self._box is False:
            self._box = _box_of(self.element, self.source)
        return self._box

    def solids(self):
        if self._solids is None:
            self._solids = element_solids(self.element, self.source.transform)
        return self._solids


def boxes_overlap(first, second):
    (a_min, a_max), (b_min, b_max) = first, second
    return (a_min[0] <= b_max[0] and b_min[0] <= a_max[0]
            and a_min[1] <= b_max[1] and b_min[1] <= a_max[1]
            and a_min[2] <= b_max[2] and b_min[2] <= a_max[2])


def box_gap(first, second):
    """Distanza (piedi) fra due box; 0 se si toccano o si sovrappongono."""
    (a_min, a_max), (b_min, b_max) = first, second
    total = 0.0
    for i in range(3):
        gap = max(0.0, b_min[i] - a_max[i], a_min[i] - b_max[i])
        total += gap * gap
    return total ** 0.5


def outline_of(box):
    return DB.Outline(DB.XYZ(*box[0]), DB.XYZ(*box[1]))


class BoxGrid(object):
    """Indice spaziale semplice sui bounding box (celle cubiche di GRID_CELL_FT).

    Un box che coprirebbe piu' di GRID_MAX_CELLS celle (un solaio intero) va in una
    lista a parte, sempre restituita come candidata.
    """

    def __init__(self):
        self.cells = {}
        self.large = []

    @staticmethod
    def _range(box):
        low, high = box
        return [(int(low[i] // GRID_CELL_FT), int(high[i] // GRID_CELL_FT)) for i in range(3)]

    def add(self, index, box):
        (x0, x1), (y0, y1), (z0, z1) = self._range(box)
        if (x1 - x0 + 1) * (y1 - y0 + 1) * (z1 - z0 + 1) > GRID_MAX_CELLS:
            self.large.append(index)
            return
        for x in range(x0, x1 + 1):
            for y in range(y0, y1 + 1):
                for z in range(z0, z1 + 1):
                    self.cells.setdefault((x, y, z), []).append(index)

    def query(self, box):
        found = set(self.large)
        (x0, x1), (y0, y1), (z0, z1) = self._range(box)
        if (x1 - x0 + 1) * (y1 - y0 + 1) * (z1 - z0 + 1) > GRID_MAX_CELLS:
            for members in self.cells.values():
                found.update(members)
            return found
        for x in range(x0, x1 + 1):
            for y in range(y0, y1 + 1):
                for z in range(z0, z1 + 1):
                    members = self.cells.get((x, y, z))
                    if members:
                        found.update(members)
        return found


# =============================================================================
# INSIEMI
# =============================================================================

def view_usable(view):
    try:
        return (view is not None and not view.IsTemplate
                and not isinstance(view, DB.ViewSheet)
                and view.ViewType not in (DB.ViewType.Schedule, DB.ViewType.Legend,
                                          DB.ViewType.DraftingView,
                                          DB.ViewType.ProjectBrowser,
                                          DB.ViewType.SystemBrowser,
                                          DB.ViewType.Report))
    except Exception:
        return False


# Stati di fase senza filtro: None = elemento senza fasi (sempre preso); Past e Future
# (demolito prima, costruito dopo) non esistono nella fase della vista.
_PHASE_NONE = getattr(DB.ElementOnPhaseStatus, "None")
_PHASE_OUTSIDE = (DB.ElementOnPhaseStatus.Past, DB.ElementOnPhaseStatus.Future)


def phase_problem(view):
    """Messaggio se la vista non ha fase / filtro di fase, altrimenti None."""
    return PhaseScope(None, view).problem


class PhaseScope(object):
    """Fase e filtro di fase della vista attiva, applicati a ogni sorgente.

    Un elemento entra se il filtro della vista lo mostra (PhaseStatusPresentation diversa
    da DontShow); filtro "<none>": tutti tranne Past e Future. Nei link si usa la fase con
    lo stesso nome; un link senza quella fase resta fuori (missing_links).
    """

    def __init__(self, doc, view):
        self.phase = None
        self.phase_filter = None
        self.problem = None
        self.missing_links = []
        self._link_phases = {}
        phase_param = filter_param = None
        try:
            if view is not None:
                phase_param = view.get_Parameter(DB.BuiltInParameter.VIEW_PHASE)
                filter_param = view.get_Parameter(DB.BuiltInParameter.VIEW_PHASE_FILTER)
        except Exception:
            phase_param = filter_param = None
        view_doc = doc if doc is not None else (view.Document if view is not None else None)
        if phase_param is not None and view_doc is not None:
            self.phase = view_doc.GetElement(phase_param.AsElementId())
        if filter_param is not None and view_doc is not None:
            # "<none>" e' un valore valido: nessun filtro.
            self.phase_filter = view_doc.GetElement(filter_param.AsElementId())
        if not isinstance(self.phase, DB.Phase) or filter_param is None:
            self.phase = None
            self.problem = (u"The active view has no Phase / Phase Filter. Clash detection "
                            u"uses the phase of the active view: activate a model view "
                            u"(plan, section, elevation or 3D) and try again.")

    def describe(self):
        if self.problem:
            return u"No phase: the active view has no Phase / Phase Filter."
        name = DB.Element.Name.GetValue(self.phase)
        filter_name = (DB.Element.Name.GetValue(self.phase_filter)
                       if isinstance(self.phase_filter, DB.PhaseFilter) else u"<none>")
        return u"Phase {}, phase filter {}.".format(name, filter_name)

    def host_phase_id(self):
        return self.phase.Id if self.phase is not None else None

    def link_phase_id(self, link):
        """Fase del link con il nome di quella della vista, o None."""
        if self.phase is None:
            return None
        if link.key not in self._link_phases:
            wanted = DB.Element.Name.GetValue(self.phase).lower()
            found = None
            try:
                for phase in link.doc.Phases:
                    if DB.Element.Name.GetValue(phase).lower() == wanted:
                        found = phase.Id
                        break
            except Exception:
                found = None
            self._link_phases[link.key] = found
            if found is None:
                self.missing_links.append(link.label)
        return self._link_phases[link.key]

    def passes(self, element, phase_id):
        try:
            status = element.GetPhaseStatus(phase_id)
        except Exception:
            return True
        if status == _PHASE_NONE:
            return True
        if status in _PHASE_OUTSIDE:
            return False
        if not isinstance(self.phase_filter, DB.PhaseFilter):
            return True
        try:
            return (self.phase_filter.GetPhaseStatusPresentation(status)
                    != DB.PhaseStatusPresentation.DontShow)
        except Exception:
            return True


class SetResolver(object):
    """Risolve gli insiemi dei test sulle sorgenti scelte.

    links: tutte le ModelSource dei link disponibili; le scelte di un insieme sono le
    loro key. Gli elementi dei link vengono dall'intero modello linkato; quando la
    sorgente dell'insieme e' la vista attiva, da quanto del link e' visibile nella
    vista (FilteredElementCollector(doc, view, link), Revit 2024+; prima l'intero link).
    """

    def __init__(self, doc, loaded_ids, view, links=()):
        self.doc = doc
        self.view = view
        self.host = ModelSource(0, doc)
        self.link_list = list(links)
        self.links = dict((source.key, source) for source in links)
        self._loaded = None
        self._loaded_ids = loaded_ids or []
        self.view_link_fallback = False
        # Fase della vista attiva per tutte le sorgenti. Senza fase (problem) gli insiemi
        # non vengono filtrati: il run e' bloccato prima, il riepilogo lo segnala.
        self.phases = PhaseScope(doc, view)

    def loaded_elements(self):
        if self._loaded is None:
            elements = []
            for eid in self._loaded_ids:
                element = self.doc.GetElement(eid)
                if element is not None:
                    elements.append(element)
            self._loaded = elements
        return self._loaded

    def view_usable(self):
        return view_usable(self.view)

    def resolve_host(self, compiled, source):
        elements = self._resolve_host(compiled, source)
        phase_id = self.phases.host_phase_id()
        if phase_id is None:
            return elements
        return [e for e in elements if self.phases.passes(e, phase_id)]

    def _resolve_host(self, compiled, source):
        if source == SOURCE_NONE:
            return []
        if source == SOURCE_SELECTION:
            return [e for e in self.loaded_elements() if compiled.matches(e, self.doc)]
        if source == SOURCE_VIEW:
            if not self.view_usable():
                return []
            return cr.collect(self.doc, compiled, self.view.Id)
        return cr.collect(self.doc, compiled)

    def resolve_link(self, compiled, link, in_view):
        if compiled.impossible:
            return []
        phase_id = None
        if self.phases.phase is not None:
            phase_id = self.phases.link_phase_id(link)
            if phase_id is None:
                return []
        collector = None
        if in_view and self.view_usable():
            try:
                collector = DB.FilteredElementCollector(self.doc, self.view.Id, link.instance.Id)
            except Exception:
                collector = None
                self.view_link_fallback = True
        if collector is None:
            collector = DB.FilteredElementCollector(link.doc)
        try:
            collector = collector.WhereElementIsNotElementType()
            category_filter = compiled.category_filter()
            if category_filter is not None:
                collector = collector.WherePasses(category_filter)
            return [e for e in collector if compiled.matches(e, link.doc)
                    and (phase_id is None or self.phases.passes(e, phase_id))]
        finally:
            collector.Dispose()

    def resolve(self, compiled, source, link_keys):
        """(items host, items linkati) dell'insieme."""
        host_items = [Item(self.host, e) for e in self.resolve_host(compiled, source)]
        link_items = []
        for key in link_keys:
            link = self.links.get(key)
            if link is None:
                continue
            for element in self.resolve_link(compiled, link, source == SOURCE_VIEW):
                link_items.append(Item(link, element))
        return host_items, link_items


# =============================================================================
# DESCRIZIONE ELEMENTI
# =============================================================================

def element_label(element, doc):
    """(categoria, "Famiglia : Tipo") dell'elemento."""
    category = u""
    family_type = u""
    try:
        category = element.Category.Name if element.Category is not None else u""
    except Exception:
        pass
    try:
        type_element = doc.GetElement(element.GetTypeId())
        if type_element is not None:
            family = getattr(type_element, "FamilyName", None) or u""
            type_name = DB.Element.Name.GetValue(type_element)
            family_type = u"{} : {}".format(family, type_name) if family else type_name
        else:
            family_type = DB.Element.Name.GetValue(element)
    except Exception:
        pass
    return category, family_type


def element_level(element, doc):
    try:
        level_id = element.LevelId
        if level_id is None or level_id == DB.ElementId.InvalidElementId:
            for bip in (DB.BuiltInParameter.RBS_START_LEVEL_PARAM,
                        DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                        DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                        DB.BuiltInParameter.SCHEDULE_LEVEL_PARAM):
                param = element.get_Parameter(bip)
                if param is not None and param.AsElementId() != DB.ElementId.InvalidElementId:
                    level_id = param.AsElementId()
                    break
        level = doc.GetElement(level_id) if level_id is not None else None
        return DB.Element.Name.GetValue(level) if level is not None else u""
    except Exception:
        return u""


class ElementInfo(object):
    """Dati di un elemento per report ed export, letti una volta sola.

    id e' l'ElementId nel documento dell'elemento (quello del link, se linkato);
    link_id l'ElementId dell'istanza di link nel modello aperto, o None; box il
    bounding box nelle coordinate del modello aperto.
    """

    __slots__ = ("id", "id_value", "key", "category", "category_value", "family_type",
                 "level", "model", "link_id", "link_uid", "box")

    def __init__(self, item):
        element = item.element
        doc = item.source.doc
        self.id = element.Id
        self.id_value = item.value
        self.key = item.key
        self.category, self.family_type = element_label(element, doc)
        # Id della categoria: per le categorie di sistema e' lo stesso in ogni documento,
        # quindi vale anche per gli override della vista del modello aperto.
        try:
            self.category_value = int(cr.get_element_id_value(element.Category.Id))
        except Exception:
            self.category_value = None
        self.level = element_level(element, doc)
        self.model = item.source.label
        self.link_id = item.source.instance.Id if item.is_link else None
        self.link_uid = item.source.instance.UniqueId if item.is_link else None
        self.box = item.box()

    @property
    def is_link(self):
        return self.link_id is not None

    def label(self):
        text = u"{} | {} | {}".format(self.category, self.family_type, self.id_value)
        if self.is_link:
            text = u"[{}] {}".format(self.model, text)
        return text

    def to_dict(self):
        # int(): da Revit 2024 gli id sono System.Int64, che json non sa scrivere.
        category_value = self.category_value
        return {"id": int(self.id_value), "source": int(self.key[0]),
                "category": self.category,
                "category_value": int(category_value) if category_value is not None else None,
                "family_type": self.family_type, "level": self.level, "model": self.model,
                "link_id": int(cr.get_element_id_value(self.link_id)) if self.link_id else None,
                "link_uid": self.link_uid,
                "box": ([[float(v) for v in self.box[0]], [float(v) for v in self.box[1]]]
                        if self.box else None)}

    @classmethod
    def from_dict(cls, data):
        info = cls.__new__(cls)
        info.id_value = int(data["id"])
        info.id = make_element_id(info.id_value)
        info.key = (int(data.get("source") or 0), info.id_value)
        info.category = data.get("category") or u""
        info.category_value = data.get("category_value")
        info.family_type = data.get("family_type") or u""
        info.level = data.get("level") or u""
        info.model = data.get("model") or u""
        link_id = data.get("link_id")
        info.link_id = make_element_id(int(link_id)) if link_id is not None else None
        info.link_uid = data.get("link_uid")
        box = data.get("box")
        info.box = (tuple(box[0]), tuple(box[1])) if box else None
        return info


def make_element_id(value):
    """ElementId da un intero: Int64 da Revit 2024, int prima."""
    try:
        import System
        return DB.ElementId(System.Int64(value))
    except Exception:
        return DB.ElementId(int(value))


class Clash(object):
    __slots__ = ("test_name", "a", "b")

    def __init__(self, test_name, a_info, b_info):
        self.test_name = test_name
        self.a = a_info
        self.b = b_info


class TestResult(object):
    def __init__(self, test):
        self.test = test
        self.name = test["name"]
        self.count_a = 0
        self.count_b = 0
        self.linked_a = 0
        self.linked_b = 0
        self.unsupported = 0
        self.ignored_connected = 0
        self.ignored_hosted = 0
        self.clashes = []
        self.note = u""
        # Diagnostica dei confronti con i link: coppie provate sui solidi, errori
        # delle prove geometriche (prima eccezione conservata per il report).
        self.link_pairs_tested = 0
        self.link_box_hits = 0
        self.link_without_box = 0
        self.geometry_errors = 0
        self.first_error = u""
        # Coppia con un elemento linkato piu' vicina fra i box: (distanza in piedi,
        # item A, item B). Dice se il filtro sui box scarta coppie che si toccano.
        self.nearest = None

    def track_nearest(self, gap, item_a, item_b):
        if self.nearest is None or gap < self.nearest[0]:
            self.nearest = (gap, item_a, item_b)

    def record_error(self, error):
        self.geometry_errors += 1
        if not self.first_error:
            self.first_error = u"{}: {}".format(type(error).__name__, error)


TEST_RESULT_FIELDS = ("count_a", "count_b", "linked_a", "linked_b", "unsupported",
                      "ignored_connected", "ignored_hosted", "note", "link_pairs_tested",
                      "link_box_hits", "link_without_box", "geometry_errors", "first_error")
SCOPE_FIELDS = ("source_a", "source_b", "links_a", "links_b", "missing_a", "missing_b")


class RunResult(object):
    """Esito di un run. model / model_key / created / saved_path servono al report
    salvato su file (Last Report); created e' l'ora del run."""

    def __init__(self, scope, links):
        self.tests = []
        self.seconds = 0.0
        self.cancelled = False
        self.scope = scope
        self.link_labels = dict((source.key, source.label) for source in links)
        self.model = u""
        self.model_key = u""
        self.created = time.strftime("%Y-%m-%d %H:%M")
        self.saved_path = None
        # Percorsi in cui il report e' stato salvato e nota per l'intestazione.
        self.report_paths = []
        self.save_note = u""
        # Id della selezione di Revit usata dal run (sorgente "Current selection"):
        # il Refresh del report rifa' il test sugli stessi elementi.
        self.selection_ids = []
        # Fase e filtro di fase della vista usata ("Phase X, phase filter Y.").
        self.phase_text = u""

    @property
    def clash_count(self):
        return sum(len(r.clashes) for r in self.tests)


# =============================================================================
# RELAZIONI DA IGNORARE
# =============================================================================

def connector_manager(element):
    try:
        if isinstance(element, DB.MEPCurve):
            return element.ConnectorManager
        mep_model = getattr(element, "MEPModel", None)
        if mep_model is not None:
            return mep_model.ConnectorManager
    except Exception:
        pass
    return None


def connected_ids(element):
    """Valori degli id degli elementi collegati ai connettori MEP (stesso documento)."""
    found = set()
    manager = connector_manager(element)
    if manager is None:
        return found
    try:
        for connector in manager.Connectors:
            try:
                if not connector.IsConnected:
                    continue
                for ref in connector.AllRefs:
                    owner = ref.Owner
                    if owner is not None:
                        found.add(cr.get_element_id_value(owner.Id))
            except Exception:
                continue
    except Exception:
        pass
    found.discard(cr.get_element_id_value(element.Id))
    return found


def host_id_value(element):
    try:
        host = getattr(element, "Host", None)
        if host is not None:
            return cr.get_element_id_value(host.Id)
    except Exception:
        pass
    return None


def intersect_supported(element):
    """Elemento con geometria che ElementIntersectsElementFilter sa confrontare."""
    try:
        return DB.ElementIntersectsFilter.IsElementSupported(element)
    except AttributeError:
        # Metodo non esposto: si prova comunque, il filtro fallisce da solo.
        return True
    except Exception:
        return False


# =============================================================================
# PROVE GEOMETRICHE
# =============================================================================

def _ids_list(ids):
    collection = List[DB.ElementId]()
    for eid in ids:
        collection.Add(eid)
    return collection


def host_hits_by_element(doc, element, box, host_ids):
    """Id degli elementi host (fra host_ids) che intersecano un elemento host."""
    collector = None
    try:
        collector = DB.FilteredElementCollector(doc, host_ids)\
            .WherePasses(DB.BoundingBoxIntersectsFilter(outline_of(box)))\
            .WherePasses(DB.ElementIntersectsElementFilter(element))
        return list(collector.ToElementIds())
    except Exception:
        return []
    finally:
        if collector is not None:
            collector.Dispose()


def host_hits_by_solids(doc, solids, box, host_ids, diag):
    """Id degli elementi host (fra host_ids) che intersecano uno dei solidi.

    Se ElementIntersectsSolidFilter rifiuta un solido, i candidati del bounding box
    sono provati con l'intersezione booleana sui loro solidi.
    """
    found = {}
    candidates = None
    for solid in solids:
        collector = None
        try:
            collector = DB.FilteredElementCollector(doc, host_ids)\
                .WherePasses(DB.BoundingBoxIntersectsFilter(outline_of(box)))\
                .WherePasses(DB.ElementIntersectsSolidFilter(solid))
            for eid in collector.ToElementIds():
                found[cr.get_element_id_value(eid)] = eid
        except Exception as error:
            diag.record_error(error)
            if candidates is None:
                candidates = _bbox_candidates(doc, box, host_ids)
            for element in candidates:
                value = cr.get_element_id_value(element.Id)
                if value not in found and \
                        solids_intersect([solid], element_solids(element, None), diag):
                    found[value] = element.Id
        finally:
            if collector is not None:
                collector.Dispose()
    return list(found.values())


def _bbox_candidates(doc, box, host_ids):
    collector = None
    try:
        collector = DB.FilteredElementCollector(doc, host_ids)\
            .WherePasses(DB.BoundingBoxIntersectsFilter(outline_of(box)))
        return list(collector.ToElements())
    except Exception:
        return []
    finally:
        if collector is not None:
            collector.Dispose()


def host_element_hits_solids(doc, element, solids, diag):
    """True se l'elemento host interseca uno dei solidi (booleana se il filtro fallisce)."""
    single = _ids_list([element.Id])
    own_solids = None
    for solid in solids:
        collector = None
        try:
            collector = DB.FilteredElementCollector(doc, single)\
                .WherePasses(DB.ElementIntersectsSolidFilter(solid))
            # Confronto per valore: l'operatore != fra ElementId non e' affidabile
            # in IronPython.
            if collector.GetElementCount() > 0:
                return True
        except Exception as error:
            diag.record_error(error)
            if own_solids is None:
                own_solids = element_solids(element, None)
            if solids_intersect([solid], own_solids, diag):
                return True
        finally:
            if collector is not None:
                collector.Dispose()
    return False


def solids_intersect(first, second, diag=None):
    for solid_a in first:
        for solid_b in second:
            try:
                common = DB.BooleanOperationsUtils.ExecuteBooleanOperation(
                    solid_a, solid_b, DB.BooleanOperationsType.Intersect)
                if common is not None and common.Volume > MIN_VOLUME_FT3:
                    return True
            except Exception as error:
                if diag is not None:
                    diag.record_error(error)
                continue
    return False


# =============================================================================
# ESECUZIONE
# =============================================================================

def _resolve_test(resolver, test, scope=None):
    scope = scope or scope_of(test, resolver.link_list)
    set_a = cr.CompiledSet(test["set_a"], resolver.doc)
    set_b = cr.CompiledSet(test["set_b"], resolver.doc)
    a_host, a_link = resolver.resolve(set_a, scope.source_a, scope.links_a)
    b_host, b_link = resolver.resolve(set_b, scope.source_b, scope.links_b)
    note = set_note(set_a, set_b, len(a_host) + len(a_link), len(b_host) + len(b_link),
                    resolver, scope)
    return a_host, a_link, b_host, b_link, note


class Summary(object):
    """Riepilogo dei test spuntati, per insieme, senza cercare le interferenze.

    rows[side]         [(modello, categoria, n)] dell'unione degli insiemi dei test
    by_category[side]  {id della categoria: n}, modello aperto e link sommati
    host[side]         elementi del modello aperto; linked[side] elementi dei link
    notes              note per test ("Ducts vs Walls: Set A is empty")
    """

    def __init__(self):
        self.rows = {"a": [], "b": []}
        self.by_category = {"a": {}, "b": {}}
        self.host = {"a": 0, "b": 0}
        self.linked = {"a": 0, "b": 0}
        self.notes = []
        self.phase_text = u""
        self.phase_problem = None


def _category_name(element):
    try:
        return element.Category.Name if element.Category is not None else u"(no category)"
    except Exception:
        return u"(no category)"


def _category_value(element):
    try:
        return int(cr.get_element_id_value(element.Category.Id))
    except Exception:
        return None


def summarize(doc, tests, loaded_ids, view, links):
    """Riepilogo dei test (sorgenti e link presi da ciascun test)."""
    resolver = SetResolver(doc, loaded_ids, view, links)
    members = {"a": {}, "b": {}}
    summary = Summary()
    for test in tests:
        a_host, a_link, b_host, b_link, note = _resolve_test(resolver, test)
        if note:
            summary.notes.append(u"{}: {}".format(test["name"], note))
        for side, items in (("a", a_host + a_link), ("b", b_host + b_link)):
            for item in items:
                members[side][item.key] = item

    for side in ("a", "b"):
        counts = {}
        for item in members[side].values():
            if item.is_link:
                summary.linked[side] += 1
            else:
                summary.host[side] += 1
            key = (item.source.index, item.source.label, _category_name(item.element))
            counts[key] = counts.get(key, 0) + 1
            value = _category_value(item.element)
            summary.by_category[side][value] = summary.by_category[side].get(value, 0) + 1
        # Prima il modello aperto, poi i link; dentro, categorie in ordine alfabetico.
        for key in sorted(counts, key=lambda k: (k[0], k[2].lower())):
            summary.rows[side].append((key[1], key[2], counts[key]))
    summary.phase_text = resolver.phases.describe()
    summary.phase_problem = resolver.phases.problem
    return summary


def set_note(set_a, set_b, count_a, count_b, resolver, scope):
    notes = []
    if SOURCE_VIEW in (scope.source_a, scope.source_b) and not resolver.view_usable():
        notes.append(u"The active view is not a model view")
    if resolver.phases.missing_links:
        notes.append(u"Links without a phase named like the active view's, left out: "
                     u"{}".format(u", ".join(resolver.phases.missing_links)))
    if resolver.view_link_fallback:
        notes.append(u"Linked elements taken from the whole link (needs Revit 2024+ "
                     u"to limit them to the view)")
    for label, compiled, missing in (("A", set_a, scope.missing_a),
                                     ("B", set_b, scope.missing_b)):
        if cr.set_is_empty(compiled.data):
            notes.append(u"Set {} has no category, filter or rule: any model element".format(
                label))
        if compiled.unknown_categories:
            notes.append(u"Set {}: unknown categories {}".format(
                label, u", ".join(compiled.unknown_categories)))
        if compiled.unknown_filters:
            notes.append(u"Set {}: view filters not in this model: {}".format(
                label, u", ".join(compiled.unknown_filters)))
        if missing:
            notes.append(u"Set {}: links not loaded: {}".format(label, u", ".join(missing)))
    if not count_a:
        notes.append(u"Set A is empty")
    if not count_b:
        notes.append(u"Set B is empty")
    return u"; ".join(notes)


def run_tests(doc, tests, loaded_ids, view, links, progress=None):
    """Esegue i test e restituisce un RunResult.

    progress(done, total) viene chiamato a ogni elemento di Set A; se restituisce True
    l'esecuzione si ferma (il risultato parziale resta).
    """
    started = time.time()
    resolver = SetResolver(doc, loaded_ids, view, links)
    result = RunResult(scope_of(tests[0], links) if tests else RunScope(
        SOURCE_NONE, SOURCE_NONE), links)

    prepared = []
    total = 0
    for test in tests:
        a_host, a_link, b_host, b_link, note = _resolve_test(resolver, test)
        test_result = TestResult(test)
        test_result.count_a = len(a_host) + len(a_link)
        test_result.count_b = len(b_host) + len(b_link)
        test_result.linked_a = len(a_link)
        test_result.linked_b = len(b_link)
        test_result.note = note
        result.tests.append(test_result)
        prepared.append((test_result, a_host + a_link, b_host, b_link))
        total += test_result.count_a

    infos = {}
    done = [0]
    try:
        for test_result, items_a, b_host, b_link in prepared:
            _run_one(doc, test_result, items_a, b_host, b_link, infos, done, total, progress)
    except ClashCancelled:
        result.cancelled = True
    result.seconds = time.time() - started
    result.phase_text = resolver.phases.describe()
    return result


def _info(infos, item):
    info = infos.get(item.key)
    if info is None:
        info = ElementInfo(item)
        infos[item.key] = info
    return info


def _run_one(doc, test_result, items_a, b_host, b_link, infos, done, total, progress):
    test = test_result.test
    if not items_a or not (b_host or b_link):
        done[0] += len(items_a)
        return

    unsupported = set()

    # Set B nel modello aperto: id per i collector nativi.
    host_by_value = {}
    host_ids = List[DB.ElementId]()
    for item in b_host:
        if not intersect_supported(item.element):
            unsupported.add(item.key)
            continue
        if item.value not in host_by_value:
            host_by_value[item.value] = item
            host_ids.Add(item.element.Id)

    # Set B nei link: confronto diretto dei box se sono pochi, altrimenti griglia.
    brute = len(b_link) <= BRUTE_FORCE_LINKED
    grid = BoxGrid()
    for index, item in enumerate(b_link):
        box = item.box()
        if box is None:
            test_result.link_without_box += 1
        elif not brute:
            grid.add(index, box)

    def link_candidates(box, item_a):
        indexes = range(len(b_link)) if brute else grid.query(box)
        found = []
        for index in indexes:
            item_b = b_link[index]
            box_b = item_b.box()
            if box_b is None or item_b.key == item_a.key:
                continue
            gap = box_gap(box, box_b)
            test_result.track_nearest(gap, item_a, item_b)
            if gap <= 1e-6:
                found.append(item_b)
        test_result.link_box_hits += len(found)
        return found

    seen_pairs = set()
    for item_a in items_a:
        done[0] += 1
        if progress is not None and progress(done[0], total):
            raise ClashCancelled()
        box = item_a.box()
        if box is None:
            continue

        hits = []
        if item_a.is_link:
            solids_a = item_a.solids()
            if not solids_a:
                unsupported.add(item_a.key)
                continue
            if host_ids.Count:
                for eid in host_hits_by_solids(doc, solids_a, box, host_ids, test_result):
                    hits.append(host_by_value.get(cr.get_element_id_value(eid)))
            for item_b in link_candidates(box, item_a):
                solids_b = item_b.solids()
                if not solids_b:
                    unsupported.add(item_b.key)
                    continue
                test_result.link_pairs_tested += 1
                if solids_intersect(solids_a, solids_b, test_result):
                    hits.append(item_b)
        else:
            if not intersect_supported(item_a.element):
                unsupported.add(item_a.key)
                continue
            if host_ids.Count:
                for eid in host_hits_by_element(doc, item_a.element, box, host_ids):
                    hits.append(host_by_value.get(cr.get_element_id_value(eid)))
            for item_b in link_candidates(box, item_a):
                solids_b = item_b.solids()
                if not solids_b:
                    unsupported.add(item_b.key)
                    continue
                test_result.link_pairs_tested += 1
                if host_element_hits_solids(doc, item_a.element, solids_b, test_result):
                    hits.append(item_b)

        if not hits:
            continue

        connected = None
        a_host = None
        for item_b in hits:
            if item_b is None or item_b.key == item_a.key:
                continue
            # Un elemento presente in entrambi gli insiemi darebbe la coppia due volte.
            pair = (min(item_a.key, item_b.key), max(item_a.key, item_b.key))
            if pair in seen_pairs:
                continue
            seen_pairs.add(pair)
            same_source = item_a.source.index == item_b.source.index
            if test["ignore_connected"] and same_source:
                if connected is None:
                    connected = connected_ids(item_a.element)
                if item_b.value in connected:
                    test_result.ignored_connected += 1
                    continue
            if test["ignore_hosted"] and same_source:
                if a_host is None:
                    a_host = host_id_value(item_a.element) or -1
                if a_host == item_b.value or host_id_value(item_b.element) == item_a.value:
                    test_result.ignored_hosted += 1
                    continue
            test_result.clashes.append(
                Clash(test_result.name, _info(infos, item_a), _info(infos, item_b)))

    test_result.unsupported = len(unsupported)


# =============================================================================
# REPORT SU FILE (Last Report, Export / Import JSON)
# =============================================================================

REPORT_FORMAT = "ESA_ClashFinder_Report"
REPORT_VERSION = 1
CONFIG_SECTION = "ESA_ClashFinder"
# Cartella dei report per modello, come "<model_key>::<cartella>" (al massimo 50).
REPORT_FOLDERS_KEY = "report_folders"
MAX_REPORT_FOLDERS = 50
OLD_FOLDER = "_old"


def model_key(doc):
    """Identita' del modello: percorso del centrale, poi del file, poi il titolo."""
    try:
        if doc.IsWorkshared:
            central = doc.GetWorksharingCentralModelPath()
            if central is not None:
                return DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(central)
    except Exception:
        pass
    return doc.PathName or doc.Title


def _model_entries(option):
    """Voci "<model_key>::<valore>" di un'opzione per modello della config."""
    from pyrevit import script
    try:
        raw = script.get_config(CONFIG_SECTION).get_option(option, []) or []
    except Exception:
        return []
    return [cr.text(entry) for entry in raw if u"::" in cr.text(entry)]


def _model_value(doc, option):
    key = model_key(doc)
    for entry in _model_entries(option):
        entry_key, _, value = entry.partition(u"::")
        if entry_key == key:
            return value
    return None


def _set_model_value(doc, option, value):
    """Il modello va in cima; al massimo MAX_REPORT_FOLDERS voci."""
    from pyrevit import script
    key = model_key(doc)
    entries = [e for e in _model_entries(option) if e.partition(u"::")[0] != key]
    entries.insert(0, u"{}::{}".format(key, value))
    try:
        cfg = script.get_config(CONFIG_SECTION)
        cfg.set_option(option, entries[:MAX_REPORT_FOLDERS])
        script.save_config()
    except Exception:
        pass


def report_folder(doc):
    """Cartella dei report scelta per il modello, o None."""
    return _model_value(doc, REPORT_FOLDERS_KEY) or None


def set_report_folder(doc, folder):
    _set_model_value(doc, REPORT_FOLDERS_KEY, folder)


# Workset delle copie dei linkati nei modelli condivisi, per modello: quello attivo
# (default) o un workset esistente, per nome.
PREVIEW_WORKSETS_KEY = "preview_worksets"
WORKSET_ACTIVE = u"<active>"


def preview_workset_choice(doc):
    # Vuoto: la voce del vecchio workset dedicato, ora tolto; vale come workset attivo.
    return _model_value(doc, PREVIEW_WORKSETS_KEY) or WORKSET_ACTIVE


def set_preview_workset_choice(doc, value):
    _set_model_value(doc, PREVIEW_WORKSETS_KEY, value)


def user_worksets(doc):
    """Nomi dei workset utente del modello (vuoto se non condiviso)."""
    if not doc.IsWorkshared:
        return []
    names = []
    try:
        for workset in DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset):
            names.append(workset.Name)
    except Exception:
        return []
    return sorted(names, key=lambda n: n.lower())


def report_file_name(test_name):
    """<nome del test>.json, senza i caratteri vietati da Windows."""
    name = re.sub(u'[<>:"/\\\\|?*\\x00-\\x1f]+', u"_", cr.text(test_name)).strip(u" .")
    return (name or u"Clash test")[:80] + u".json"


def _same_file(first, second):
    return os.path.normcase(os.path.abspath(first)) == os.path.normcase(os.path.abspath(second))


def _archive(path):
    """Sposta il report gia' presente in _old, con data e ora del file nel nome."""
    folder = os.path.join(os.path.dirname(path), OLD_FOLDER)
    if not os.path.isdir(folder):
        os.makedirs(folder)
    stem = os.path.splitext(os.path.basename(path))[0]
    stamp = time.strftime("%Y%m%d_%H%M%S", time.localtime(os.path.getmtime(path)))
    target = os.path.join(folder, u"{}_{}.json".format(stem, stamp))
    number = 2
    while os.path.exists(target):
        target = os.path.join(folder, u"{}_{}_{}.json".format(stem, stamp, number))
        number += 1
    os.rename(path, target)


def save_reports(folder, result, doc):
    """Un file per test nella cartella, <test>.json; quello che c'era va in _old.
    Un file gia' scritto per questo stesso report viene solo riscritto. Restituisce
    i percorsi scritti. Solleva IOError / OSError."""
    if not os.path.isdir(folder):
        os.makedirs(folder)
    own = [p for p in [result.saved_path] + list(result.report_paths) if p]
    paths = []
    for test_result in result.tests:
        single = RunResult(result.scope, [])
        single.link_labels = result.link_labels
        single.seconds = result.seconds
        single.cancelled = result.cancelled
        single.created = result.created
        single.selection_ids = result.selection_ids
        single.phase_text = result.phase_text
        single.tests = [test_result]
        path = os.path.join(folder, report_file_name(test_result.name))
        if os.path.exists(path) and not any(_same_file(path, p) for p in own):
            _archive(path)
        save_report(path, single, doc)
        paths.append(path)
    return paths


def pair_key(clash):
    """Identita' di un'interferenza fra due run: (istanza di link, id) dei due elementi."""
    return (clash.a.link_uid or u"", int(clash.a.id_value),
            clash.b.link_uid or u"", int(clash.b.id_value))


def latest_report(folder):
    """Il report piu' recente della cartella (non in _old), o None."""
    best = None
    best_time = None
    try:
        names = os.listdir(folder)
    except Exception:
        return None
    for name in names:
        path = os.path.join(folder, name)
        if not name.lower().endswith(".json") or not os.path.isfile(path):
            continue
        modified = os.path.getmtime(path)
        if best_time is None or modified > best_time:
            best, best_time = path, modified
    return best


def save_report(path, result, doc):
    """Scrive il report su file. Gli elementi sono scritti una volta sola e le
    interferenze li richiamano per indice. Solleva IOError / OSError."""
    elements = []
    index_of = {}

    def index(info):
        if info.key not in index_of:
            index_of[info.key] = len(elements)
            elements.append(info.to_dict())
        return index_of[info.key]

    tests = []
    for test_result in result.tests:
        data = {"name": test_result.name, "test": test_result.test,
                "clashes": [[index(c.a), index(c.b)] for c in test_result.clashes]}
        for field in TEST_RESULT_FIELDS:
            data[field] = getattr(test_result, field)
        tests.append(data)
    payload = {
        "format": REPORT_FORMAT, "version": REPORT_VERSION,
        "model": doc.Title, "model_key": model_key(doc), "created": result.created,
        "seconds": result.seconds, "cancelled": result.cancelled,
        "scope": dict((f, getattr(result.scope, f)) for f in SCOPE_FIELDS),
        "link_labels": result.link_labels,
        "selection_ids": [int(value) for value in result.selection_ids],
        "phase": result.phase_text,
        "elements": elements, "tests": tests}
    cr.write_json_file(path, payload)


def load_report(path):
    """RunResult letto da file. Solleva IOError / ValueError."""
    with io.open(path, "r", encoding="utf-8-sig") as handle:
        data = json.loads(handle.read())
    if not isinstance(data, dict) or data.get("format") != REPORT_FORMAT:
        raise ValueError("The file is not a Clash Finder report.")
    infos = [ElementInfo.from_dict(item) for item in data.get("elements") or []]
    scope_data = data.get("scope") or {}
    scope = RunScope(scope_data.get("source_a", SOURCE_NONE),
                     scope_data.get("source_b", SOURCE_NONE),
                     scope_data.get("links_a") or [], scope_data.get("links_b") or [],
                     scope_data.get("missing_a") or [], scope_data.get("missing_b") or [])
    result = RunResult(scope, [])
    result.link_labels = data.get("link_labels") or {}
    result.seconds = float(data.get("seconds") or 0.0)
    result.cancelled = bool(data.get("cancelled"))
    result.model = data.get("model") or u""
    result.model_key = data.get("model_key") or u""
    result.created = data.get("created") or u""
    result.saved_path = path
    result.selection_ids = [int(value) for value in data.get("selection_ids") or []]
    result.phase_text = data.get("phase") or u""
    for test_data in data.get("tests") or []:
        test = cr.normalize_test(test_data.get("test")) or cr.new_test(
            test_data.get("name") or u"Test")
        test_result = TestResult(test)
        for field in TEST_RESULT_FIELDS:
            if field in test_data:
                setattr(test_result, field, test_data[field])
        for a_index, b_index in test_data.get("clashes") or []:
            test_result.clashes.append(Clash(test_result.name, infos[a_index], infos[b_index]))
        result.tests.append(test_result)
    return result
