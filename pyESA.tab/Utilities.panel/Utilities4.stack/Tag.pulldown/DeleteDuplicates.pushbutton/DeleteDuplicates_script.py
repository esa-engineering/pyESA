# -*- coding: utf-8 -*-
# Intestazione dello script con metadati
__title__ = "Delete\nDuplicate Tags"
__doc__ = """Version = 1.0
Date    = 28.08.2026
________________________________________________________________
Analizza le viste selezionate, segnala i tag duplicati (piu tag
che puntano allo stesso elemento nella stessa vista) ed elimina
quelli confermati dall'utente.
________________________________________________________________
CLICK        : apre la finestra con la sola vista attiva preselezionata
SHIFT + CLICK: apre la finestra con tutte le viste preselezionate
________________________________________________________________
Author(s): Andrea Patti
"""
__author__ = "Andrea Patti"

# -------------------------------
# SEZIONE IMPORT MODULI
# -------------------------------
import os.path as op

import clr
# Gli assembly WPF servono per costruire in codice le checkbox di viste e
# risultati. Referenziati esplicitamente per non dipendere dall'ordine di
# import ne da cosa il motore IronPython di pyRevit ha gia caricato.
for _asm in ("WindowsBase", "PresentationCore", "PresentationFramework"):
    try:
        clr.AddReference(_asm)
    except Exception:
        pass

from System.Collections.Generic import List
from System.Windows import Thickness, Visibility
from System.Windows import Controls

from pyrevit import revit, script, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    IndependentTag,
)

# -------------------------------
# SEZIONE INIZIALIZZAZIONE
# -------------------------------
doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

if doc is None:
    forms.alert("Nessun documento aperto.", title="Documento mancante",
                exitscript=True)

if doc.IsFamilyDocument:
    forms.alert("Lo script lavora solo su un progetto, non su una famiglia.",
                title="Documento non valido", exitscript=True)

# Grafica della finestra, nella stessa cartella dello script.
XAML_FILE_NAME = "DeleteDuplicatesUI.xaml"

# Righe di dettaglio stampate nella tabella del report: oltre questa soglia
# la tabella viene troncata e il troncamento e' dichiarato esplicitamente.
REPORT_ROW_LIMIT = 400

MM_TO_FT = 1.0 / 304.8

# Categorie dei tag che non derivano da IndependentTag (SpatialElementTag).
SPATIAL_TAG_CATEGORIES = (
    ("OST_RoomTags", "tag di locale"),
    ("OST_MEPSpaceTags", "tag di spazio MEP"),
    ("OST_AreaTags", "tag di area"),
)

# Attributi da cui ricavare l'elemento taggato, in ordine di preferenza.
# Sono diversi tra IndependentTag e le sottoclassi di SpatialElementTag, e
# cambiano tra le versioni di Revit: si prova quello che esiste.
LINKID_ATTRS = ("TaggedElementId", "TaggedRoomId", "TaggedSpaceId",
                "TaggedAreaId")
ELEMID_ATTRS = ("TaggedLocalElementId", "TaggedLocalRoomId")
ELEMENT_ATTRS = ("Room", "Space", "Area")

KEEP_FIRST = "Il piu vecchio"
KEEP_LAST = "Il piu recente"

# Avvisi raccolti durante l'analisi e stampati in coda al report
WARNINGS = []


def warn(message):
    """Registra un avviso: niente degradazione silenziosa."""
    if message not in WARNINGS:
        WARNINGS.append(message)


# -------------------------------------------------------------------------
# SEZIONE HELPER - COMPATIBILITA API
# -------------------------------------------------------------------------
def get_element_id_value(eid):
    """Valore intero di un ElementId.

    Revit 2026 ha rimosso IntegerValue in favore di Value (Int64).
    Restituisce None se l'id non e' leggibile.
    """
    if eid is None:
        return None
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    try:
        return eid.IntegerValue   # Revit <= 2025
    except Exception:
        return None


def make_element_id(value):
    """ElementId a partire dal valore intero, None se non costruibile."""
    if value is None:
        return None
    try:
        return ElementId(value)
    except Exception:
        return None


def get_builtin_category(name):
    """BuiltInCategory per nome, None se la versione non la conosce."""
    return getattr(BuiltInCategory, name, None)


# -------------------------------------------------------------------------
# SEZIONE HELPER - LETTURA DEI TAG
# -------------------------------------------------------------------------
def link_pair(link_element_id):
    """Normalizza un LinkElementId in (id_istanza_link, id_elemento, linkato).

    Per un elemento del modello corrente l'id istanza vale 0.
    Restituisce None se il riferimento non punta a nulla di valido.
    """
    if link_element_id is None:
        return None
    linked = get_element_id_value(
        getattr(link_element_id, "LinkedElementId", None))
    host = get_element_id_value(
        getattr(link_element_id, "HostElementId", None))
    inst = get_element_id_value(
        getattr(link_element_id, "LinkInstanceId", None))
    if linked is not None and linked > 0:
        return (inst if (inst is not None and inst > 0) else 0, linked, True)
    if host is not None and host > 0:
        return (0, host, False)
    return None


def get_tag_targets(tag):
    """Elenco dei riferimenti di un tag, come tuple (istanza, elemento, linkato).

    Un tag multi-riferimento (Revit 2022+) restituisce piu tuple: il confronto
    a valle usa l'insieme completo, cosi due multi-tag sono duplicati solo se
    puntano esattamente agli stessi elementi.
    """
    pairs = []

    getter = getattr(tag, "GetTaggedElementIds", None)
    if getter is not None:
        try:
            for link_id in getter():
                pair = link_pair(link_id)
                if pair is not None and pair not in pairs:
                    pairs.append(pair)
        except Exception:
            pass
    if pairs:
        return pairs

    # Ripiego sulle proprieta singole: SpatialElementTag e le versioni
    # precedenti al multi-riferimento.
    for attr in LINKID_ATTRS:
        try:
            pair = link_pair(getattr(tag, attr, None))
        except Exception:
            pair = None
        if pair is not None:
            return [pair]

    for attr in ELEMID_ATTRS:
        try:
            value = get_element_id_value(getattr(tag, attr, None))
        except Exception:
            value = None
        if value is not None and value > 0:
            return [(0, value, False)]

    for attr in ELEMENT_ATTRS:
        try:
            element = getattr(tag, attr, None)
            value = get_element_id_value(element.Id) if element else None
        except Exception:
            value = None
        if value is not None and value > 0:
            return [(0, value, False)]

    return []


def tag_position(tag):
    """Posizione della testa del tag, None se non ricavabile."""
    try:
        head = getattr(tag, "TagHeadPosition", None)
        if head is not None:
            return head
    except Exception:
        pass
    try:
        location = tag.Location
        if location is not None and hasattr(location, "Point"):
            return location.Point
    except Exception:
        pass
    return None


def points_near(point_a, point_b, tolerance_ft):
    if point_a is None or point_b is None:
        return False
    try:
        return point_a.DistanceTo(point_b) <= tolerance_ft
    except Exception:
        return False


# -------------------------------------------------------------------------
# SEZIONE HELPER - DESCRIZIONI PER IL REPORT
# -------------------------------------------------------------------------
_link_doc_cache = {}
_target_label_cache = {}
_type_label_cache = {}


def get_link_doc(link_instance_value):
    """Documento di un RevitLinkInstance, None se scaricato o non trovato."""
    if link_instance_value in _link_doc_cache:
        return _link_doc_cache[link_instance_value]

    link_doc = None
    eid = make_element_id(link_instance_value)
    if eid is not None:
        try:
            instance = doc.GetElement(eid)
            if instance is not None:
                link_doc = instance.GetLinkDocument()
        except Exception:
            link_doc = None
    _link_doc_cache[link_instance_value] = link_doc
    return link_doc


def element_label(element):
    """Etichetta leggibile di un elemento: categoria, nome, marca/numero."""
    parts = []
    try:
        if element.Category is not None:
            parts.append(element.Category.Name)
    except Exception:
        pass
    try:
        name = element.Name
        if name:
            parts.append(name)
    except Exception:
        pass
    for bip_name in ("ROOM_NUMBER", "ALL_MODEL_MARK"):
        bip = getattr(BuiltInParameter, bip_name, None)
        if bip is None:
            continue
        try:
            param = element.get_Parameter(bip)
            if param is not None:
                value = param.AsString()
                if value:
                    parts.append("[{}]".format(value))
                    break
        except Exception:
            pass
    return " - ".join(parts) if parts else None


def describe_target(pair):
    """Descrizione dell'elemento taggato, con il nome del link se serve."""
    if pair in _target_label_cache:
        return _target_label_cache[pair]

    instance_value, element_value, is_linked = pair
    label = None
    source_doc = get_link_doc(instance_value) if is_linked else doc

    if source_doc is not None:
        eid = make_element_id(element_value)
        if eid is not None:
            try:
                element = source_doc.GetElement(eid)
                if element is not None:
                    label = element_label(element)
            except Exception:
                label = None

    if not label:
        label = "Elemento {}".format(element_value)
    else:
        label = "{} (Id {})".format(label, element_value)

    if is_linked:
        link_name = "link"
        eid = make_element_id(instance_value)
        if eid is not None:
            try:
                instance = doc.GetElement(eid)
                if instance is not None:
                    link_name = instance.Name
            except Exception:
                pass
        label = "{} > {}".format(link_name, label)

    _target_label_cache[pair] = label
    return label


def describe_type(tag):
    """Nome famiglia e tipo del tag."""
    type_value = get_element_id_value(tag.GetTypeId())
    if type_value in _type_label_cache:
        return _type_label_cache[type_value]

    label = "Tipo {}".format(type_value)
    try:
        tag_type = doc.GetElement(tag.GetTypeId())
        if tag_type is not None:
            family = getattr(tag_type, "FamilyName", None)
            name = None
            try:
                name = tag_type.Name
            except Exception:
                name = None
            if family and name:
                label = "{} : {}".format(family, name)
            elif name:
                label = name
    except Exception:
        pass

    _type_label_cache[type_value] = label
    return label


def view_label(view):
    try:
        return "{} [{}]".format(view.Name, view.ViewType)
    except Exception:
        return "Vista {}".format(get_element_id_value(view.Id))


# -------------------------------------------------------------------------
# SEZIONE RACCOLTA TAG
# -------------------------------------------------------------------------
def collect_tags_by_view():
    """Tutti i tag del modello raggruppati per vista di appartenenza.

    Una sola passata sul documento: i tag sono elementi view-specific, quindi
    OwnerViewId basta a smistarli senza interrogare vista per vista.
    Restituisce (dizionario id_vista -> lista tag, dizionario id_vista -> vista).
    """
    collectors = []
    try:
        collectors.append(FilteredElementCollector(doc)
                          .OfClass(IndependentTag)
                          .WhereElementIsNotElementType())
    except Exception as err:
        warn("Raccolta dei tag standard non riuscita: {}".format(err))

    for bic_name, description in SPATIAL_TAG_CATEGORIES:
        bic = get_builtin_category(bic_name)
        if bic is None:
            warn("Categoria {} non disponibile in questa versione di Revit: "
                 "i {} non sono stati analizzati.".format(bic_name, description))
            continue
        try:
            collectors.append(FilteredElementCollector(doc)
                              .OfCategory(bic)
                              .WhereElementIsNotElementType())
        except Exception as err:
            warn("Raccolta dei {} non riuscita: {}".format(description, err))

    seen = set()
    tags_by_view = {}
    views_by_key = {}

    for collector in collectors:
        for tag in collector:
            tag_value = get_element_id_value(tag.Id)
            if tag_value is None or tag_value in seen:
                continue
            seen.add(tag_value)

            owner_id = getattr(tag, "OwnerViewId", None)
            owner_value = get_element_id_value(owner_id)
            if owner_value is None or owner_value <= 0:
                continue

            if owner_value not in views_by_key:
                view = None
                try:
                    view = doc.GetElement(owner_id)
                except Exception:
                    view = None
                if view is None or getattr(view, "IsTemplate", False):
                    views_by_key[owner_value] = None
                else:
                    views_by_key[owner_value] = view
            if views_by_key[owner_value] is None:
                continue

            tags_by_view.setdefault(owner_value, []).append(tag)

    for key in list(views_by_key.keys()):
        if views_by_key[key] is None:
            views_by_key.pop(key)

    return tags_by_view, views_by_key


# -------------------------------------------------------------------------
# SEZIONE ANALISI
# -------------------------------------------------------------------------
class DuplicateGroup(object):
    """Un insieme di tag della stessa vista che puntano allo stesso elemento."""

    def __init__(self, view, tags, keep, targets, type_label):
        self.view = view
        self.view_name = view_label(view)
        self.tags = tags
        self.keep = keep
        self.targets = targets
        self.type_label = type_label
        keep_value = get_element_id_value(keep.Id)
        self.to_delete = [t for t in tags
                          if get_element_id_value(t.Id) != keep_value]
        # Gli ElementId vengono conservati qui perche restano leggibili anche
        # dopo l'eliminazione, mentre gli Element diventano invalidi: il report
        # viene stampato a valle della transazione.
        self.keep_id = keep.Id
        self.delete_ids = [t.Id for t in self.to_delete]

    def target_label(self):
        return " + ".join([describe_target(p) for p in self.targets])

    def description(self):
        return "{}  |  {}  |  {}  |  {} tag, {} da eliminare".format(
            self.view_name,
            self.target_label(),
            self.type_label,
            len(self.tags),
            len(self.to_delete))


def cluster_by_position(tags, tolerance_ft):
    """Raggruppa i tag per vicinanza della testa.

    Un tag di cui non si riesce a leggere la posizione resta isolato: meglio
    non eliminarlo che eliminarlo per un dato mancante.
    """
    positions = [tag_position(t) for t in tags]
    clusters = []
    used = [False] * len(tags)

    for i in range(len(tags)):
        if used[i]:
            continue
        used[i] = True
        current = [i]
        cursor = 0
        while cursor < len(current):
            index = current[cursor]
            cursor += 1
            for j in range(len(tags)):
                if used[j]:
                    continue
                if points_near(positions[index], positions[j], tolerance_ft):
                    used[j] = True
                    current.append(j)
        clusters.append([tags[k] for k in current])

    return clusters


def analyse(views, options):
    """Cerca i duplicati nelle viste indicate. Non modifica il modello.

    Restituisce (lista di DuplicateGroup, dizionario di statistiche).
    """
    groups = []
    stats = {
        "views": len(views),
        "tags": 0,
        "skipped_pinned": 0,
        "skipped_linked": 0,
        "skipped_orphan": 0,
        "isolated_no_position": 0,
    }

    tolerance_ft = options["tolerance_mm"] * MM_TO_FT

    for view in views:
        view_key = get_element_id_value(view.Id)
        tags = TAGS_BY_VIEW.get(view_key, [])
        buckets = {}

        for tag in tags:
            stats["tags"] += 1

            if not options["include_pinned"]:
                try:
                    if tag.Pinned:
                        stats["skipped_pinned"] += 1
                        continue
                except Exception:
                    pass

            targets = get_tag_targets(tag)
            if not targets:
                stats["skipped_orphan"] += 1
                continue

            if not options["include_linked"]:
                if any(pair[2] for pair in targets):
                    stats["skipped_linked"] += 1
                    continue

            target_key = tuple(sorted([(p[0], p[1]) for p in targets]))
            if options["same_type"]:
                bucket_key = (target_key, get_element_id_value(tag.GetTypeId()))
            else:
                bucket_key = (target_key, None)
            buckets.setdefault(bucket_key, []).append((tag, targets))

        for bucket_key in buckets:
            entries = buckets[bucket_key]
            if len(entries) < 2:
                continue

            bucket_tags = [e[0] for e in entries]
            targets = entries[0][1]

            if options["use_position"]:
                clusters = cluster_by_position(bucket_tags, tolerance_ft)
            else:
                clusters = [bucket_tags]

            for cluster in clusters:
                if len(cluster) < 2:
                    if options["use_position"] and len(bucket_tags) > 1:
                        if tag_position(cluster[0]) is None:
                            stats["isolated_no_position"] += 1
                    continue
                ordered = sorted(
                    cluster, key=lambda t: get_element_id_value(t.Id))
                keep = ordered[0] if options["keep_first"] else ordered[-1]
                if options["same_type"]:
                    type_label = describe_type(ordered[0])
                else:
                    labels = []
                    for t in ordered:
                        label = describe_type(t)
                        if label not in labels:
                            labels.append(label)
                    type_label = " / ".join(labels)
                groups.append(
                    DuplicateGroup(view, ordered, keep, targets, type_label))

    groups.sort(key=lambda g: (g.view_name, g.target_label()))
    return groups, stats


# -------------------------------------------------------------------------
# SEZIONE RACCOLTA DATI INIZIALE
# -------------------------------------------------------------------------
TAGS_BY_VIEW, VIEWS_BY_KEY = collect_tags_by_view()

if not TAGS_BY_VIEW:
    output.close_others()
    output.print_md("# Delete Duplicate Tags")
    output.print_md("Nel modello **{}** non e' stato trovato nessun tag "
                    "(tag standard, di locale, di spazio MEP o di area) "
                    "collocato in una vista.".format(doc.Title))
    if WARNINGS:
        output.print_md("## Avvisi")
        for message in WARNINGS:
            output.print_md("- {}".format(message))
    script.exit()

# Ordine di presentazione: nome vista, cosi la lista e' navigabile.
CANDIDATE_VIEWS = sorted(VIEWS_BY_KEY.values(), key=lambda v: view_label(v))

active_view_key = None
try:
    if doc.ActiveView is not None:
        active_view_key = get_element_id_value(doc.ActiveView.Id)
except Exception:
    active_view_key = None

# pyRevit inietta __shiftclick__ a runtime
preselect_all = bool(globals().get("__shiftclick__", False))  # noqa: F821


# -------------------------------------------------------------------------
# SEZIONE FINESTRA
# -------------------------------------------------------------------------
class DuplicateTagsWindow(forms.WPFWindow):
    """Finestra unica: selezione viste, criteri, risultati, eliminazione.

    L'analisi gira dentro la finestra ed e' sola lettura. L'eliminazione non
    parte da qui: la finestra si chiude registrando i gruppi confermati e la
    transazione viene aperta dallo script, come nel resto dell'estensione.
    """

    # Attributo di classe: gli handler XAML possono scattare durante il
    # popolamento iniziale, prima che __init__ abbia finito.
    _ready = False

    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)

        self.result = None          # gruppi confermati per l'eliminazione
        self.analysis = None        # (gruppi, statistiche, opzioni) dell'ultima analisi
        self.view_checks = []
        self.result_checks = []

        self.header_hint.Text = (
            "Sono elencate solo le viste che contengono almeno un tag. "
            "L'analisi cerca i tag che puntano allo stesso elemento nella "
            "stessa vista e non modifica il modello.")

        self.keep_combo.Items.Add(KEEP_FIRST)
        self.keep_combo.Items.Add(KEEP_LAST)
        self.keep_combo.SelectedIndex = 0

        for view in CANDIDATE_VIEWS:
            key = get_element_id_value(view.Id)
            check = Controls.CheckBox()
            check.Content = "{}  -  {} tag".format(
                view_label(view), len(TAGS_BY_VIEW.get(key, [])))
            check.Tag = view
            check.Margin = Thickness(2, 2, 2, 2)
            check.IsChecked = preselect_all or (key == active_view_key)
            check.Checked += self.view_toggled
            check.Unchecked += self.view_toggled
            self.views_panel.Children.Add(check)
            self.view_checks.append(check)

        self._ready = True
        self.refresh_view_count()

    # ------------- helper -------------
    def refresh_view_count(self):
        count = len([c for c in self.view_checks if c.IsChecked])
        self.views_count.Text = "{} di {} viste selezionate".format(
            count, len(self.view_checks))

    def refresh_result_count(self):
        selected = [c for c in self.result_checks if c.IsChecked]
        tags = sum([len(c.Tag.to_delete) for c in selected])
        self.results_count.Text = "{} di {} gruppi selezionati, {} tag da " \
                                  "eliminare".format(len(selected),
                                                     len(self.result_checks),
                                                     tags)
        self.btn_delete.IsEnabled = bool(selected)

    def clear_results(self):
        """Invalida i risultati: i criteri o le viste sono cambiati."""
        self.analysis = None
        self.result_checks = []
        self.results_panel.Children.Clear()
        self.results_count.Text = ""
        self.btn_delete.IsEnabled = False
        self.btn_res_all.IsEnabled = False
        self.btn_res_none.IsEnabled = False
        self.results_summary.Text = ("Nessuna analisi eseguita. Seleziona le "
                                     "viste e premi Analizza.")

    def parse_tolerance(self):
        """Tolleranza in mm. Accetta virgola o punto.
        Restituisce (valore, messaggio_errore)."""
        raw = (self.tol_input.Text or "").strip().replace(",", ".")
        if not raw:
            return None, "Inserisci la tolleranza di sovrapposizione in mm."
        try:
            value = float(raw)
        except ValueError:
            return None, "La tolleranza non e' un numero valido."
        if value < 0:
            return None, "La tolleranza non puo essere negativa."
        if value > 5000.0:
            return None, "La tolleranza sembra fuori scala (> 5000 mm)."
        return value, None

    def read_options(self):
        """Opzioni correnti della finestra. Restituisce (opzioni, errore)."""
        tolerance, error = self.parse_tolerance()
        if error and self.opt_position.IsChecked:
            return None, error
        return {
            "same_type": bool(self.crit_same_type.IsChecked),
            "include_linked": bool(self.opt_linked.IsChecked),
            "include_pinned": bool(self.opt_pinned.IsChecked),
            "use_position": bool(self.opt_position.IsChecked),
            "tolerance_mm": tolerance if tolerance is not None else 0.0,
            "keep_first": (str(self.keep_combo.SelectedItem) == KEEP_FIRST),
        }, None

    # ------------- handler XAML -------------
    def view_toggled(self, sender, args):
        if not self._ready:
            return
        self.refresh_view_count()
        self.clear_results()

    def result_toggled(self, sender, args):
        if self._ready:
            self.refresh_result_count()

    def filter_changed(self, sender, args):
        if not self._ready:
            return
        needle = (self.views_filter.Text or "").strip().lower()
        for check in self.view_checks:
            visible = (not needle) or (needle in str(check.Content).lower())
            check.Visibility = \
                Visibility.Visible if visible else Visibility.Collapsed

    def criteria_changed(self, sender, args):
        if self._ready:
            self.clear_results()

    def select_all_click(self, sender, args):
        for check in self.view_checks:
            check.IsChecked = True

    def select_none_click(self, sender, args):
        for check in self.view_checks:
            check.IsChecked = False

    def select_filtered_click(self, sender, args):
        for check in self.view_checks:
            check.IsChecked = (check.Visibility == Visibility.Visible)

    def results_all_click(self, sender, args):
        for check in self.result_checks:
            check.IsChecked = True

    def results_none_click(self, sender, args):
        for check in self.result_checks:
            check.IsChecked = False

    def analizza_click(self, sender, args):
        views = [c.Tag for c in self.view_checks if c.IsChecked]
        if not views:
            forms.alert("Seleziona almeno una vista.", title="Nessuna vista")
            return

        options, error = self.read_options()
        if error:
            forms.alert(error, title="Tolleranza")
            return

        groups, stats = analyse(views, options)

        self.result_checks = []
        self.results_panel.Children.Clear()

        for group in groups:
            check = Controls.CheckBox()
            check.Content = group.description()
            check.Tag = group
            check.IsChecked = True
            check.Margin = Thickness(2, 2, 2, 2)
            check.Checked += self.result_toggled
            check.Unchecked += self.result_toggled
            self.results_panel.Children.Add(check)
            self.result_checks.append(check)

        self.analysis = (groups, stats, options)

        if groups:
            total = sum([len(g.to_delete) for g in groups])
            self.results_summary.Text = (
                "Trovati {} gruppi di tag duplicati in {} viste: {} tag "
                "verrebbero eliminati. Deseleziona i gruppi che vuoi "
                "conservare.".format(
                    len(groups),
                    len(set([get_element_id_value(g.view.Id) for g in groups])),
                    total))
        else:
            self.results_summary.Text = (
                "Nessun tag duplicato trovato nelle {} viste analizzate "
                "({} tag esaminati) con i criteri impostati.".format(
                    stats["views"], stats["tags"]))

        self.btn_res_all.IsEnabled = bool(groups)
        self.btn_res_none.IsEnabled = bool(groups)
        self.refresh_result_count()

    def elimina_click(self, sender, args):
        selected = [c.Tag for c in self.result_checks if c.IsChecked]
        if not selected:
            forms.alert("Seleziona almeno un gruppo da ripulire.",
                        title="Nessun gruppo selezionato")
            return

        total = sum([len(g.to_delete) for g in selected])
        confirmed = forms.alert(
            "Verranno eliminati {} tag in {} gruppi.\n\n"
            "In ogni gruppo resta un solo tag. L'operazione e' annullabile "
            "con Undo di Revit dopo la chiusura della finestra.\n\n"
            "Procedo?".format(total, len(selected)),
            title="Conferma eliminazione", yes=True, no=True)
        if not confirmed:
            return

        self.result = selected
        self.Close()

    def chiudi_click(self, sender, args):
        self.result = None
        self.Close()


# -------------------------------------------------------------------------
# SEZIONE AVVIO FINESTRA
# -------------------------------------------------------------------------
xaml_path = None
try:
    xaml_path = script.get_bundle_file(XAML_FILE_NAME)
except Exception:
    pass
if not xaml_path:
    xaml_path = op.join(op.dirname(__file__), XAML_FILE_NAME)

if not op.isfile(xaml_path):
    forms.alert("File grafica non trovato:\n\n{}\n\nDeve stare nella stessa "
                "cartella dello script.".format(xaml_path),
                title="Grafica mancante", exitscript=True)

window = DuplicateTagsWindow(xaml_path)
window.show_dialog()

if window.analysis is None:
    # Finestra chiusa senza analisi: nessun report da produrre.
    script.exit()

groups, stats, options = window.analysis
selected_groups = window.result or []


# -------------------------------------------------------------------------
# SEZIONE ELIMINAZIONE
# -------------------------------------------------------------------------
def delete_groups(target_groups):
    """Elimina i tag in eccesso. Restituisce (eliminati, lista di fallimenti).

    I valori degli id vengono letti prima della cancellazione: dopo, gli
    oggetti Element non sono piu interrogabili.
    """
    requested = []      # (tag, valore_id)
    for group in target_groups:
        for tag in group.to_delete:
            requested.append((tag, get_element_id_value(tag.Id)))

    deleted = 0
    failures = []

    with revit.Transaction("Delete duplicate tags"):
        ids = List[ElementId]()
        for tag, _ in requested:
            try:
                if tag.Pinned:
                    tag.Pinned = False
            except Exception:
                pass
            ids.Add(tag.Id)

        try:
            removed = doc.Delete(ids)
            # Delete restituisce anche gli elementi cancellati per dipendenza:
            # si contano solo i tag effettivamente richiesti.
            removed_values = set()
            if removed is not None:
                for eid in removed:
                    removed_values.add(get_element_id_value(eid))
            deleted = len([v for _, v in requested if v in removed_values])
            for tag, value in requested:
                if value not in removed_values:
                    failures.append((value, "non eliminato da Revit"))
        except Exception as batch_error:
            # Ripiego elemento per elemento, per isolare i casi non eliminabili
            # (elemento posseduto da un altro utente, tag gia rimosso, ecc.).
            warn("Eliminazione in blocco non riuscita ({}), i tag sono stati "
                 "eliminati uno alla volta.".format(batch_error))
            deleted = 0
            failures = []
            for tag, value in requested:
                eid = make_element_id(value)
                try:
                    if eid is None or doc.GetElement(eid) is None:
                        failures.append((value, "tag gia rimosso"))
                        continue
                    single = List[ElementId]()
                    single.Add(eid)
                    doc.Delete(single)
                    deleted += 1
                except Exception as err:
                    failures.append((value, str(err)))

    return deleted, failures


deleted_count = 0
delete_failures = []
if selected_groups:
    deleted_count, delete_failures = delete_groups(selected_groups)


# -------------------------------------------------------------------------
# SEZIONE REPORT
# -------------------------------------------------------------------------
output.close_others()
output.print_md("# Delete Duplicate Tags")
output.print_md("Modello: **{}**".format(doc.Title))

criteria_lines = [
    "- Criterio: **{}**".format(
        "stesso elemento taggato e stesso tipo di tag" if options["same_type"]
        else "stesso elemento taggato, qualsiasi tipo di tag"),
    "- Tag di elementi linkati: **{}**".format(
        "inclusi" if options["include_linked"] else "esclusi"),
    "- Tag bloccati (pin): **{}**".format(
        "inclusi" if options["include_pinned"] else "esclusi"),
    "- Tag da mantenere in ogni gruppo: **{}**".format(
        "il piu vecchio (Id piu basso)" if options["keep_first"]
        else "il piu recente (Id piu alto)"),
]
if options["use_position"]:
    criteria_lines.append(
        "- Vincolo di sovrapposizione: **attivo**, tolleranza {:.1f} mm "
        "sulla testa del tag".format(options["tolerance_mm"]))
else:
    criteria_lines.append("- Vincolo di sovrapposizione: **non attivo**")

output.print_md("## Criteri usati")
for line in criteria_lines:
    output.print_md(line)

output.print_md("## Analisi")
output.print_md("- Viste analizzate: **{}**".format(stats["views"]))
output.print_md("- Tag esaminati: **{}**".format(stats["tags"]))
output.print_md("- Gruppi di duplicati trovati: **{}**".format(len(groups)))
output.print_md("- Tag in eccesso individuati: **{}**".format(
    sum([len(g.to_delete) for g in groups])))

if stats["skipped_pinned"]:
    output.print_md("- Tag bloccati (pin) ignorati: **{}**".format(
        stats["skipped_pinned"]))
if stats["skipped_linked"]:
    output.print_md("- Tag di elementi linkati ignorati: **{}**".format(
        stats["skipped_linked"]))
if stats["skipped_orphan"]:
    output.print_md("- Tag senza elemento associato, ignorati: **{}**".format(
        stats["skipped_orphan"]))
if stats["isolated_no_position"]:
    output.print_md("- Tag esclusi dal confronto perche la posizione della "
                    "testa non e' leggibile: **{}**".format(
                        stats["isolated_no_position"]))

if not groups:
    output.print_md("## Esito")
    output.print_md("**Nessun tag duplicato** con i criteri impostati. "
                    "Se ti aspettavi dei duplicati, prova il criterio "
                    "'qualsiasi tipo di tag' o disattiva il vincolo di "
                    "sovrapposizione.")
else:
    selected_ids = set()
    for group in selected_groups:
        selected_ids.add(id(group))

    rows = []
    for group in groups:
        if id(group) in selected_ids:
            state = "eliminato" if deleted_count else "selezionato"
        else:
            state = "conservato"
        deleted_links = ", ".join(
            [output.linkify(eid) for eid in group.delete_ids])
        rows.append([
            group.view_name,
            group.target_label(),
            group.type_label,
            str(len(group.tags)),
            output.linkify(group.keep_id),
            deleted_links,
            state,
        ])

    truncated = 0
    if len(rows) > REPORT_ROW_LIMIT:
        truncated = len(rows) - REPORT_ROW_LIMIT
        rows = rows[:REPORT_ROW_LIMIT]

    output.print_md("## Dettaglio dei duplicati")
    output.print_table(
        table_data=rows,
        columns=["Vista", "Elemento taggato", "Tipo di tag", "N. tag",
                 "Tag mantenuto", "Tag in eccesso", "Stato"])
    if truncated:
        output.print_md("_Tabella troncata: altri **{}** gruppi non sono "
                        "elencati. L'eliminazione, se confermata, li ha "
                        "comunque trattati tutti._".format(truncated))
    if deleted_count:
        output.print_md("_I link della colonna 'Tag in eccesso' non sono piu "
                        "risolvibili per i tag gia eliminati: restano come "
                        "traccia degli Id rimossi._")

    output.print_md("## Esito")
    if not selected_groups:
        output.print_md("Nessuna eliminazione richiesta: il modello **non e' "
                        "stato modificato**.")
    else:
        output.print_md("Tag eliminati: **{}** su {} richiesti, in {} "
                        "gruppi.".format(
                            deleted_count,
                            sum([len(g.to_delete) for g in selected_groups]),
                            len(selected_groups)))
        if delete_failures:
            output.print_md("### Tag non eliminati")
            output.print_table(
                table_data=[[str(v), m] for v, m in delete_failures],
                columns=["Id tag", "Motivo"])

if WARNINGS:
    output.print_md("## Avvisi")
    for message in WARNINGS:
        output.print_md("- {}".format(message))
