# -*- coding: utf-8 -*-
"""Crea abachi (schedule) multipli, uno per ogni categoria selezionata.

CLICK
    Selezione manuale delle categorie e dei parametri (nativi, di progetto,
    condivisi). Il report finale elenca i parametri scartati perche' non
    associati a una determinata categoria.

SHIFT + CLICK
    Usa un abaco esistente come template: campi, ordinamento/raggruppamento
    e formattazione vengono replicati sulle categorie selezionate. Anche in
    questo caso gli scarti finiscono nel report.

Il nome dell'abaco e' composto da <prefisso digitato> + <nome categoria>.

Motore: IronPython 2.7 (pyRevit)
Autore: ESA Engineering
"""

__title__ = "Crea\nAbachi"
__author__ = "ESA Engineering"
__min_revit_ver__ = 2019
__highlight__ = "new"

import os
import traceback
from collections import OrderedDict

import clr

clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Action
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Input import Cursors
from System.Windows.Threading import DispatcherPriority

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Category,
    FilteredElementCollector,
    ParameterElement,
    ScheduleFilter,
    ScheduleSortGroupField,
    SharedParameterElement,
    Transaction,
    View,
    ViewSchedule,
)

from pyrevit import forms
from pyrevit import revit
from pyrevit import script


# ---------------------------------------------------------------------------
# CONTESTO
# ---------------------------------------------------------------------------

doc = revit.doc
logger = script.get_logger()
output = script.get_output()

try:
    SHIFT = __shiftclick__          # noqa: F821  (iniettata da pyRevit)
except NameError:
    SHIFT = False

try:
    from pyrevit import EXEC_PARAMS
    BUNDLE_DIR = EXEC_PARAMS.command_path
except Exception:
    BUNDLE_DIR = os.path.dirname(__file__)

XAML_MAIN = os.path.join(BUNDLE_DIR, "MainWindow.xaml")
XAML_TEMPLATE = os.path.join(BUNDLE_DIR, "TemplateWindow.xaml")
XAML_CONFLICT = os.path.join(BUNDLE_DIR, "ConflictWindow.xaml")

KEY_SEP = "\x1f"
INVALID_NAME_CHARS = "\\:{}[]|;<>?`~"

# Categorie non presenti in doc.Settings.Categories ma schedulabili
EXTRA_BIC_NAMES = ("OST_Views", "OST_Sheets")

# Proprieta' di ScheduleDefinition replicate dal template.
# Tutte in lettura/scrittura; l'assegnazione e' comunque protetta da try.
DEF_PROPS = (
    "IsItemized",
    "ShowTitle",
    "ShowHeaders",
    "ShowGrandTotal",
    "ShowGrandTotalTitle",
    "ShowGrandTotalCount",
    "GrandTotalTitle",
    "IncludeLinkedFiles",
    "ShowGridLines",
)

# Proprieta' di ScheduleField replicate dal template.
# HasTotals non e' inclusa: deprecata dal 2017 e rimossa dalle versioni
# recenti; il suo ruolo e' coperto da DisplayType.
FIELD_PROPS = (
    "ColumnHeading",
    "IsHidden",
    "HorizontalAlignment",
    "VerticalAlignment",
    "HeadingOrientation",
    "DisplayType",
    "SheetColumnWidth",
    "GridColumnWidth",
)

CALC_FIELD_TYPES = ("Formula", "Percentage", "CombinedParameter")


# ---------------------------------------------------------------------------
# HELPER GENERICI
# ---------------------------------------------------------------------------

def eid_value(eid):
    """Valore intero di un ElementId (compatibile Revit <=2023 e >=2024)."""
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def sanitize_name(name):
    """Rimuove i caratteri non ammessi da Revit nei nomi delle viste."""
    if not name:
        return ""
    cleaned = "".join([ch for ch in name if ch not in INVALID_NAME_CHARS])
    return cleaned.strip()


def make_key(param_id_value, field_type, name):
    return "{0}{3}{1}{3}{2}".format(param_id_value, field_type, name, KEY_SEP)


def param_origin(document, param_id):
    """Nativo / Condiviso / Progetto."""
    try:
        value = eid_value(param_id)
    except Exception:
        return "n/d"
    if value < 0:
        return "Nativo"
    try:
        element = document.GetElement(param_id)
    except Exception:
        element = None
    if element is None:
        return "n/d"
    if isinstance(element, SharedParameterElement):
        return "Condiviso"
    if isinstance(element, ParameterElement):
        return "Progetto"
    return "n/d"


# ---------------------------------------------------------------------------
# CATEGORIE E CAMPI SCHEDULABILI
# ---------------------------------------------------------------------------

class FieldInfo(object):
    """Descrittore serializzabile di un SchedulableField."""

    def __init__(self, key, name, param_id_value, field_type, origin):
        self.key = key
        self.name = name
        self.param_id_value = param_id_value
        self.field_type = field_type
        self.origin = origin


def is_valid_schedule_category(document, cat_id):
    """True / False / None (API non disponibile in questa versione)."""
    for args in ((cat_id,), (document, cat_id)):
        try:
            return bool(ViewSchedule.IsValidCategoryForSchedule(*args))
        except Exception:
            continue
    return None


def get_schedulable_categories(document):
    """Elenco ordinato di Category per cui e' possibile creare un abaco."""
    result = []
    seen = set()

    for cat in document.Settings.Categories:
        try:
            if cat is None:
                continue
            cat_id = cat.Id
            value = eid_value(cat_id)
            if value >= 0:                       # solo categorie built-in
                continue
            if value in seen:
                continue
            if str(cat.CategoryType) not in ("Model", "Annotation"):
                continue
            valid = is_valid_schedule_category(document, cat_id)
            if valid is False:
                continue
            if valid is None and not cat.AllowsBoundParameters:
                continue
            seen.add(value)
            result.append(cat)
        except Exception:
            continue

    for bic_name in EXTRA_BIC_NAMES:
        try:
            bic = getattr(BuiltInCategory, bic_name)
            cat = Category.GetCategory(document, bic)
            if cat is None:
                continue
            value = eid_value(cat.Id)
            if value in seen:
                continue
            if is_valid_schedule_category(document, cat.Id) is False:
                continue
            seen.add(value)
            result.append(cat)
        except Exception:
            continue

    result.sort(key=lambda c: c.Name.lower())
    return result


def probe_categories(document, cat_ids, cache):
    """Popola `cache` con {valore_cat_id: {key: FieldInfo}}.

    Crea abachi temporanei in una transazione annullata: e' l'unico modo
    affidabile per ottenere l'elenco completo dei campi schedulabili di
    una categoria (nativi + di progetto + condivisi).
    Il valore None indica una categoria non schedulabile.
    """
    todo = [cid for cid in cat_ids if eid_value(cid) not in cache]
    if not todo:
        return

    trans = Transaction(document, "pyRevit - lettura campi schedulabili")
    trans.Start()
    try:
        for cat_id in todo:
            fields = OrderedDict()
            try:
                temp = ViewSchedule.CreateSchedule(document, cat_id)
                for sfield in temp.Definition.GetSchedulableFields():
                    try:
                        name = sfield.GetName(document)
                    except Exception:
                        name = None
                    if not name:
                        continue
                    pid = eid_value(sfield.ParameterId)
                    ftype = str(sfield.FieldType)
                    key = make_key(pid, ftype, name)
                    if key in fields:
                        continue
                    fields[key] = FieldInfo(
                        key, name, pid, ftype,
                        param_origin(document, sfield.ParameterId))
            except Exception as err:
                logger.debug("Categoria non schedulabile %s: %s", cat_id, err)
                fields = None
            cache[eid_value(cat_id)] = fields
    finally:
        try:
            trans.RollBack()
        except Exception:
            pass


def match_schedulable(document, schedulable_fields, param_id_value,
                      field_type, name):
    """Trova il SchedulableField corrispondente, con fallback progressivi."""
    for sfield in schedulable_fields:
        try:
            if eid_value(sfield.ParameterId) == param_id_value \
                    and str(sfield.FieldType) == field_type:
                return sfield
        except Exception:
            continue

    if param_id_value != -1:
        for sfield in schedulable_fields:
            try:
                if eid_value(sfield.ParameterId) == param_id_value:
                    return sfield
            except Exception:
                continue

    for sfield in schedulable_fields:
        try:
            if sfield.GetName(document) == name \
                    and str(sfield.FieldType) == field_type:
                return sfield
        except Exception:
            continue

    for sfield in schedulable_fields:
        try:
            if sfield.GetName(document) == name:
                return sfield
        except Exception:
            continue

    return None


# ---------------------------------------------------------------------------
# NOMI E CONFLITTI
# ---------------------------------------------------------------------------

def get_existing_views(document):
    """{nome_minuscolo: View} per tutte le viste non template."""
    result = {}
    for view in FilteredElementCollector(document).OfClass(View).ToElements():
        try:
            if view.IsTemplate:
                continue
            result[view.Name.strip().lower()] = view
        except Exception:
            continue
    return result


def make_unique_name(base, taken):
    if base.strip().lower() not in taken:
        return base
    index = 1
    while True:
        candidate = "{0} ({1})".format(base, index)
        if candidate.strip().lower() not in taken:
            return candidate
        index += 1
        if index > 999:
            return candidate


# ---------------------------------------------------------------------------
# REPORT
# ---------------------------------------------------------------------------

class Report(object):

    def __init__(self, mode, prefix):
        self.mode = mode
        self.prefix = prefix
        self.rows = []
        self.issues = []
        self.notes = []

    def add_row(self, category, name, element_id, added, dropped, status):
        self.rows.append({
            "category": category,
            "name": name,
            "id": element_id,
            "added": added,
            "dropped": dropped,
            "status": status,
        })

    def add_issue(self, category, kind, detail):
        self.issues.append({
            "category": category,
            "kind": kind,
            "detail": detail,
        })

    def add_note(self, text):
        self.notes.append(text)

    # -- rendering ----------------------------------------------------------

    def render(self):
        output.set_title("Report creazione abachi")
        output.print_md("# Report creazione abachi")
        output.print_md(
            "**Modalita**: {0}&nbsp;&nbsp;&nbsp;"
            "**Prefisso**: `{1}`".format(
                self.mode, self.prefix if self.prefix else "(nessuno)"))

        created = [r for r in self.rows if r["status"].startswith("OK")]
        output.print_md(
            "**Abachi creati**: {0} / {1} categorie richieste&nbsp;&nbsp;&nbsp;"
            "**Segnalazioni**: {2}".format(
                len(created), len(self.rows), len(self.issues)))

        if self.rows:
            output.print_md("## Esito per categoria")
            data = []
            for row in self.rows:
                if row["id"] is not None:
                    label = output.linkify(row["id"], row["name"])
                else:
                    label = row["name"]
                data.append([
                    row["category"],
                    label,
                    str(row["added"]),
                    str(row["dropped"]),
                    row["status"],
                ])
            output.print_table(
                table_data=data,
                columns=["Categoria", "Abaco", "Campi aggiunti",
                         "Campi scartati", "Esito"])

        if self.issues:
            output.print_md("## Segnalazioni")
            data = []
            for issue in self.issues:
                data.append([
                    issue["category"],
                    issue["kind"],
                    issue["detail"].replace("|", "/"),
                ])
            output.print_table(
                table_data=data,
                columns=["Categoria", "Tipo", "Dettaglio"])
        else:
            output.print_md("## Segnalazioni\n\nNessuna. "
                            "Tutti i campi richiesti sono stati aggiunti.")

        if self.notes:
            output.print_md("## Note")
            for note in self.notes:
                output.print_md("- " + note)


# ---------------------------------------------------------------------------
# UTILITY UI
# ---------------------------------------------------------------------------

def ui_pump(window):
    """Forza un giro di rendering della finestra (operazioni sincrone lunghe)."""
    noop = Action(lambda: None)
    try:
        window.Dispatcher.Invoke(noop, DispatcherPriority.Background)
        return
    except Exception:
        pass
    try:
        window.Dispatcher.Invoke(DispatcherPriority.Background, noop)
    except Exception:
        pass


def ui_busy(window, state):
    try:
        window.Cursor = Cursors.Wait if state else None
    except Exception:
        pass
    ui_pump(window)


# ---------------------------------------------------------------------------
# MODELLI PER LE LISTE WPF
# ---------------------------------------------------------------------------

class CatItem(object):

    def __init__(self, category):
        self.category = category
        self.cat_id = category.Id
        self.name = category.Name
        self.checked = False
        self.tooltip = "{0} (id {1})".format(category.Name,
                                             eid_value(category.Id))


class ParamItem(object):

    def __init__(self, field_info):
        self.info = field_info
        self.key = field_info.key
        self.name = field_info.name
        self.origin = field_info.origin
        self.checked = False
        self.order = None
        self.count = 0
        self.total = 0
        self.coverage = ""
        self.tooltip = ""


class SchedItem(object):

    def __init__(self, schedule, subtitle):
        self.schedule = schedule
        self.name = schedule.Name
        self.subtitle = subtitle


class FieldItem(object):

    def __init__(self, index, name, origin):
        self.idx = str(index)
        self.name = name
        self.origin = origin


class ConflictItem(object):

    ACTIONS = ("rename", "overwrite", "skip")

    def __init__(self, name, subtitle, category_name, cat_id, existing_view):
        self.name = name
        self.subtitle = subtitle
        self.category_name = category_name
        self.cat_id = cat_id
        self.existing_view = existing_view
        self.action_index = 0

    @property
    def action(self):
        return ConflictItem.ACTIONS[self.action_index]


# ---------------------------------------------------------------------------
# FINESTRA 1 - SELEZIONE MANUALE
# ---------------------------------------------------------------------------

class ManualWindow(forms.WPFWindow):

    def __init__(self, document):
        forms.WPFWindow.__init__(self, XAML_MAIN)
        self.doc = document
        self.result = None
        self._loaded = False
        self._order_seq = 1
        self._cache = {}
        self._param_state = {}
        self._all_params = []

        self._all_cats = [CatItem(c) for c in get_schedulable_categories(document)]
        self._rebuild_cat_list()
        self._refresh_params()
        self._update_preview()
        self._loaded = True

    # -- categorie ----------------------------------------------------------

    def _visible_cats(self):
        query = (self.cat_search_tb.Text or "").strip().lower()
        if not query:
            return list(self._all_cats)
        return [c for c in self._all_cats if query in c.name.lower()]

    def _update_cat_header(self):
        selected = len([c for c in self._all_cats if c.checked])
        self.cat_header_tb.Text = \
            "1. Categorie - {0} selezionate su {1}".format(
                selected, len(self._all_cats))

    def _rebuild_cat_list(self):
        self.cat_lb.ItemsSource = ObservableCollection[object](self._visible_cats())
        self._update_cat_header()

    def on_cat_search(self, sender, args):
        if not self._loaded:
            return
        self._rebuild_cat_list()

    def on_cat_click(self, sender, args):
        item = sender.DataContext
        if item is None:
            return
        item.checked = bool(sender.IsChecked)
        self._update_cat_header()
        self._refresh_params()
        self._update_preview()

    def on_cat_select_all(self, sender, args):
        for item in self._visible_cats():
            item.checked = True
        self._rebuild_cat_list()
        self._refresh_params()
        self._update_preview()

    def on_cat_clear(self, sender, args):
        for item in self._all_cats:
            item.checked = False
        self._rebuild_cat_list()
        self._refresh_params()
        self._update_preview()

    # -- parametri ----------------------------------------------------------

    def _refresh_params(self):
        selected_cats = [c for c in self._all_cats if c.checked]
        if not selected_cats:
            self._all_params = []
            self.param_lb.ItemsSource = ObservableCollection[object]()
            self.param_header_tb.Text = "2. Parametri"
            self.status_tb.Text = "Seleziona almeno una categoria."
            self._update_order_text()
            return

        self.status_tb.Text = "Lettura dei parametri schedulabili in corso..."
        ui_busy(self, True)
        try:
            probe_categories(self.doc, [c.cat_id for c in selected_cats],
                             self._cache)
        finally:
            ui_busy(self, False)

        union = OrderedDict()
        counts = {}
        unsupported = []
        for cat_item in selected_cats:
            fields = self._cache.get(eid_value(cat_item.cat_id))
            if fields is None:
                unsupported.append(cat_item.name)
                continue
            for key, field_info in fields.items():
                if key not in union:
                    union[key] = field_info
                counts[key] = counts.get(key, 0) + 1

        total = len(selected_cats) - len(unsupported)
        items = []
        for key, field_info in union.items():
            item = ParamItem(field_info)
            state = self._param_state.get(key)
            if state:
                item.checked, item.order = state
            item.count = counts.get(key, 0)
            item.total = total
            item.coverage = "{0}/{1}".format(item.count, total)
            item.tooltip = (
                "{0}\nOrigine: {1}\nTipo campo: {2}\n"
                "Disponibile in {3} categorie su {4} selezionate".format(
                    item.name, item.origin, field_info.field_type,
                    item.count, total))
            items.append(item)

        items.sort(key=lambda i: i.name.lower())
        self._all_params = items

        if unsupported:
            self.status_tb.Text = \
                "Categorie senza abaco disponibile: {0}".format(
                    ", ".join(unsupported))
        else:
            self.status_tb.Text = ""

        self._apply_param_filter()
        self._update_order_text()

    def _apply_param_filter(self):
        query = (self.param_search_tb.Text or "").strip().lower()
        only_common = bool(self.only_common_cb.IsChecked)
        visible = []
        for item in self._all_params:
            if only_common and item.total and item.count < item.total:
                continue
            if query and query not in item.name.lower():
                continue
            visible.append(item)
        self.param_lb.ItemsSource = ObservableCollection[object](visible)
        self.param_header_tb.Text = \
            "2. Parametri - {0} visibili su {1} (nativi / progetto / condivisi)".format(
                len(visible), len(self._all_params))

    def on_param_search(self, sender, args):
        if not self._loaded:
            return
        self._apply_param_filter()

    def on_only_common(self, sender, args):
        if not self._loaded:
            return
        self._apply_param_filter()

    def _set_checked(self, item, value):
        item.checked = value
        if value:
            if item.order is None:
                item.order = self._order_seq
                self._order_seq += 1
        else:
            item.order = None
        self._param_state[item.key] = (item.checked, item.order)

    def on_param_click(self, sender, args):
        item = sender.DataContext
        if item is None:
            return
        self._set_checked(item, bool(sender.IsChecked))
        self._update_order_text()

    def on_param_select_all(self, sender, args):
        for item in self.param_lb.ItemsSource:
            self._set_checked(item, True)
        self._apply_param_filter()
        self._update_order_text()

    def on_param_clear(self, sender, args):
        for item in self._all_params:
            self._set_checked(item, False)
        self._apply_param_filter()
        self._update_order_text()

    def _selected_params(self):
        chosen = [p for p in self._all_params if p.checked]
        chosen.sort(key=lambda p: (p.order if p.order is not None else 99999,
                                   p.name.lower()))
        return chosen

    def _update_order_text(self):
        chosen = self._selected_params()
        if not chosen:
            self.order_tb.Text = "Nessun parametro selezionato."
            return
        names = [p.name for p in chosen]
        text = "  >  ".join(names[:25])
        if len(names) > 25:
            text += "  >  ... (+{0})".format(len(names) - 25)
        self.order_tb.Text = "{0} campi:  {1}".format(len(names), text)

    # -- prefisso -----------------------------------------------------------

    def _update_preview(self):
        prefix = sanitize_name(self.prefix_tb.Text or "")
        selected = [c for c in self._all_cats if c.checked]
        sample = selected[0].name if selected else "NomeCategoria"
        self.preview_tb.Text = "Nome risultante: {0}{1}".format(prefix, sample)

    def on_prefix_changed(self, sender, args):
        if not self._loaded:
            return
        self._update_preview()

    # -- azioni -------------------------------------------------------------

    def on_cancel(self, sender, args):
        self.result = None
        self.Close()

    def on_create(self, sender, args):
        cats = [c for c in self._all_cats if c.checked]
        params = self._selected_params()
        if not cats:
            forms.alert("Seleziona almeno una categoria.", title="Crea Abachi")
            return
        if not params:
            if not forms.alert(
                    "Nessun parametro selezionato: gli abachi verranno "
                    "creati senza campi.\nVuoi procedere comunque?",
                    title="Crea Abachi", yes=True, no=True):
                return
        self.result = {
            "prefix": sanitize_name(self.prefix_tb.Text or ""),
            "categories": [(c.cat_id, c.name) for c in cats],
            "params": [p.info for p in params],
        }
        self.Close()


# ---------------------------------------------------------------------------
# FINESTRA 2 - DA TEMPLATE
# ---------------------------------------------------------------------------

def collect_template_schedules(document):
    """Abachi utilizzabili come template."""
    result = []
    for sched in FilteredElementCollector(document).OfClass(ViewSchedule).ToElements():
        try:
            if sched.IsTemplate:
                continue
            if sched.IsTitleblockRevisionSchedule:
                continue
            definition = sched.Definition
            if definition is None:
                continue
            if definition.IsKeySchedule:
                continue
            cat_name = "categoria non definita"
            try:
                cat = Category.GetCategory(document, definition.CategoryId)
                if cat is not None:
                    cat_name = cat.Name
            except Exception:
                pass
            kind = "Computo materiali" if definition.IsMaterialTakeoff else "Abaco"
            n_fields = len(list(definition.GetFieldOrder()))
            subtitle = "{0} - {1} - {2} campi".format(kind, cat_name, n_fields)
            result.append(SchedItem(sched, subtitle))
        except Exception:
            continue
    result.sort(key=lambda s: s.name.lower())
    return result


def read_template_fields(document, schedule):
    """Estrae la descrizione ordinata dei campi del template.

    Restituisce (specs, fid_to_key).
    """
    definition = schedule.Definition
    specs = []
    fid_to_key = {}

    for field_id in definition.GetFieldOrder():
        try:
            field = definition.GetField(field_id)
        except Exception:
            continue

        try:
            name = field.GetName()
        except Exception:
            name = ""
        try:
            param_id = field.ParameterId
            param_id_value = eid_value(param_id)
        except Exception:
            param_id = None
            param_id_value = -1
        field_type = str(field.FieldType)

        is_calc = field_type in CALC_FIELD_TYPES
        try:
            is_calc = is_calc or bool(field.IsCalculatedField)
        except Exception:
            pass

        key = make_key(param_id_value, field_type, name)
        # chiavi duplicate: rendile univoche mantenendo l'ordine
        base_key = key
        suffix = 1
        while key in fid_to_key.values():
            key = "{0}#{1}".format(base_key, suffix)
            suffix += 1

        props = {}
        for prop in FIELD_PROPS:
            try:
                props[prop] = getattr(field, prop)
            except Exception:
                continue
        try:
            fmt = field.GetFormatOptions()
        except Exception:
            fmt = None
        try:
            style = field.GetStyle()
        except Exception:
            style = None

        specs.append({
            "key": key,
            "name": name,
            "param_id_value": param_id_value,
            "field_type": field_type,
            "origin": param_origin(document, param_id) if param_id else "n/d",
            "is_calc": is_calc,
            "props": props,
            "format": fmt,
            "style": style,
        })
        fid_to_key[field_id.IntegerValue] = key

    return specs, fid_to_key


class TemplateWindow(forms.WPFWindow):

    def __init__(self, document):
        forms.WPFWindow.__init__(self, XAML_TEMPLATE)
        self.doc = document
        self.result = None
        self._loaded = False

        self._all_scheds = collect_template_schedules(document)
        self._all_cats = [CatItem(c) for c in get_schedulable_categories(document)]

        self._refresh_scheds()
        self._rebuild_cat_list()
        self._update_preview()
        self.field_header_tb.Text = "Campi del template"
        self._loaded = True

        if not self._all_scheds:
            self.status_tb.Text = \
                "Nessun abaco utilizzabile come template nel progetto."

    # -- abachi -------------------------------------------------------------

    def _visible_scheds(self):
        query = (self.sched_search_tb.Text or "").strip().lower()
        if not query:
            return list(self._all_scheds)
        return [s for s in self._all_scheds
                if query in s.name.lower() or query in s.subtitle.lower()]

    def _refresh_scheds(self):
        self.sched_lb.ItemsSource = ObservableCollection[object](self._visible_scheds())
        self.sched_header_tb.Text = \
            "1. Abaco template - {0} disponibili".format(len(self._all_scheds))

    def on_sched_search(self, sender, args):
        if not self._loaded:
            return
        self._refresh_scheds()

    def on_sched_selected(self, sender, args):
        item = self.sched_lb.SelectedItem
        if item is None:
            self.field_lb.ItemsSource = ObservableCollection[object]()
            self.field_header_tb.Text = "Campi del template"
            return
        specs, _ = read_template_fields(self.doc, item.schedule)
        rows = []
        for index, spec in enumerate(specs, start=1):
            origin = spec["origin"]
            if spec["is_calc"]:
                origin = "Calcolato"
            rows.append(FieldItem(index, spec["name"], origin))
        self.field_lb.ItemsSource = ObservableCollection[object](rows)
        self.field_header_tb.Text = \
            "Campi del template - {0}".format(len(rows))
        calc = len([s for s in specs if s["is_calc"]])
        if calc:
            self.status_tb.Text = (
                "Attenzione: il template contiene {0} campi calcolati/combinati, "
                "che non possono essere replicati automaticamente.".format(calc))
        else:
            self.status_tb.Text = ""

    # -- categorie ----------------------------------------------------------

    def _visible_cats(self):
        query = (self.cat_search_tb.Text or "").strip().lower()
        if not query:
            return list(self._all_cats)
        return [c for c in self._all_cats if query in c.name.lower()]

    def _update_cat_header(self):
        selected = len([c for c in self._all_cats if c.checked])
        self.cat_header_tb.Text = \
            "2. Categorie - {0} su {1}".format(selected, len(self._all_cats))

    def _rebuild_cat_list(self):
        self.cat_lb.ItemsSource = ObservableCollection[object](self._visible_cats())
        self._update_cat_header()

    def on_cat_search(self, sender, args):
        if not self._loaded:
            return
        self._rebuild_cat_list()

    def on_cat_click(self, sender, args):
        item = sender.DataContext
        if item is None:
            return
        item.checked = bool(sender.IsChecked)
        self._update_cat_header()
        self._update_preview()

    def on_cat_select_all(self, sender, args):
        for item in self._visible_cats():
            item.checked = True
        self._rebuild_cat_list()
        self._update_preview()

    def on_cat_clear(self, sender, args):
        for item in self._all_cats:
            item.checked = False
        self._rebuild_cat_list()
        self._update_preview()

    # -- prefisso -----------------------------------------------------------

    def _update_preview(self):
        prefix = sanitize_name(self.prefix_tb.Text or "")
        selected = [c for c in self._all_cats if c.checked]
        sample = selected[0].name if selected else "NomeCategoria"
        self.preview_tb.Text = "Nome risultante: {0}{1}".format(prefix, sample)

    def on_prefix_changed(self, sender, args):
        if not self._loaded:
            return
        self._update_preview()

    # -- azioni -------------------------------------------------------------

    def on_cancel(self, sender, args):
        self.result = None
        self.Close()

    def on_create(self, sender, args):
        item = self.sched_lb.SelectedItem
        cats = [c for c in self._all_cats if c.checked]
        if item is None:
            forms.alert("Seleziona un abaco da usare come template.",
                        title="Crea Abachi")
            return
        if not cats:
            forms.alert("Seleziona almeno una categoria.", title="Crea Abachi")
            return
        self.result = {
            "prefix": sanitize_name(self.prefix_tb.Text or ""),
            "template": item.schedule,
            "categories": [(c.cat_id, c.name) for c in cats],
            "copy_sort": bool(self.opt_sort_cb.IsChecked),
            "copy_format": bool(self.opt_format_cb.IsChecked),
            "copy_filters": bool(self.opt_filters_cb.IsChecked),
        }
        self.Close()


# ---------------------------------------------------------------------------
# FINESTRA 3 - CONFLITTI DI NOME
# ---------------------------------------------------------------------------

class ConflictWindow(forms.WPFWindow):

    def __init__(self, items):
        forms.WPFWindow.__init__(self, XAML_CONFLICT)
        self.items = items
        self.confirmed = False
        self._loaded = False
        self._refresh()
        self.warn_tb.Text = (
            "{0} nomi in conflitto. 'Sovrascrivi' elimina definitivamente "
            "la vista esistente.".format(len(items)))
        self._loaded = True

    def _refresh(self):
        self.conflict_lb.ItemsSource = ObservableCollection[object](self.items)

    def on_action_changed(self, sender, args):
        item = sender.DataContext
        if item is None:
            return
        index = sender.SelectedIndex
        if index is None or index < 0:
            return
        item.action_index = index

    def _set_all(self, index):
        for item in self.items:
            item.action_index = index
        self._refresh()

    def on_all_rename(self, sender, args):
        self._set_all(0)

    def on_all_overwrite(self, sender, args):
        self._set_all(1)

    def on_all_skip(self, sender, args):
        self._set_all(2)

    def on_cancel(self, sender, args):
        self.confirmed = False
        self.Close()

    def on_ok(self, sender, args):
        self.confirmed = True
        self.Close()


# ---------------------------------------------------------------------------
# PIANIFICAZIONE DEI NOMI
# ---------------------------------------------------------------------------

def build_plan(document, prefix, categories, report, protected_view=None):
    """Determina nome finale e azione per ogni categoria.

    `protected_view` e' una vista che non puo' essere sovrascritta (in
    modalita' template e' l'abaco usato come modello: eliminarlo a meta'
    ciclo farebbe fallire tutte le categorie successive).

    Restituisce (plan, annullato_dall_utente).
    plan = lista di dict {cat_id, cat_name, name, delete_view}
    """
    existing = get_existing_views(document)
    taken = set(existing.keys())
    protected_id = None
    if protected_view is not None:
        try:
            protected_id = eid_value(protected_view.Id)
        except Exception:
            protected_id = None

    entries = []
    conflicts = []
    for cat_id, cat_name in categories:
        base = sanitize_name("{0}{1}".format(prefix, cat_name))
        if not base:
            base = sanitize_name(cat_name) or "Abaco"
        entry = {
            "cat_id": cat_id,
            "cat_name": cat_name,
            "base": base,
            "name": base,
            "delete_view": None,
            "skip": False,
        }
        clash = existing.get(base.strip().lower())
        if clash is not None:
            is_sched = isinstance(clash, ViewSchedule)
            subtitle = "Categoria: {0} - vista esistente: {1}".format(
                cat_name, "abaco" if is_sched else type(clash).__name__)
            conflict = ConflictItem(base, subtitle, cat_name, cat_id, clash)
            conflicts.append(conflict)
            entry["conflict"] = conflict
        entries.append(entry)

    if conflicts:
        window = ConflictWindow(conflicts)
        window.show_dialog()
        if not window.confirmed:
            return None, True

    for entry in entries:
        conflict = entry.get("conflict")
        if conflict is None:
            entry["name"] = make_unique_name(entry["base"], taken)
            taken.add(entry["name"].strip().lower())
            continue

        action = conflict.action
        if action == "skip":
            entry["skip"] = True
            report.add_row(entry["cat_name"], entry["base"], None, 0, 0,
                           "SALTATO - nome gia' esistente")
            report.add_issue(entry["cat_name"], "Nome duplicato",
                             "Abaco non creato su richiesta dell'utente: "
                             "'{0}' esiste gia'.".format(entry["base"]))
            continue

        if action == "overwrite":
            existing_view = conflict.existing_view
            reason = None
            try:
                same_as_template = (protected_id is not None
                                    and eid_value(existing_view.Id) == protected_id)
            except Exception:
                same_as_template = False

            if not isinstance(existing_view, ViewSchedule):
                reason = "'{0}' non e' un abaco".format(entry["base"])
            elif same_as_template:
                reason = ("'{0}' e' l'abaco usato come template"
                          .format(entry["base"]))

            if reason is None:
                entry["delete_view"] = existing_view
                entry["name"] = entry["base"]
                taken.add(entry["name"].strip().lower())
                report.add_issue(entry["cat_name"], "Sovrascrittura",
                                 "La vista '{0}' esistente e' stata eliminata "
                                 "e ricreata.".format(entry["base"]))
            else:
                entry["name"] = make_unique_name(entry["base"], taken)
                taken.add(entry["name"].strip().lower())
                report.add_issue(
                    entry["cat_name"], "Sovrascrittura non possibile",
                    "{0}: e' stato usato il nome '{1}'.".format(
                        reason, entry["name"]))
            continue

        entry["name"] = make_unique_name(entry["base"], taken)
        taken.add(entry["name"].strip().lower())
        report.add_issue(entry["cat_name"], "Nome duplicato",
                         "'{0}' esiste gia': rinominato in '{1}'.".format(
                             entry["base"], entry["name"]))

    return [e for e in entries if not e["skip"]], False


# ---------------------------------------------------------------------------
# CREAZIONE - MODALITA MANUALE
# ---------------------------------------------------------------------------

def run_manual(document, config):
    report = Report("Selezione manuale (click)", config["prefix"])
    params = config["params"]

    plan, cancelled = build_plan(document, config["prefix"],
                                 config["categories"], report)
    if cancelled:
        return None
    if not plan:
        return report

    trans = Transaction(document, "pyRevit - Crea abachi")
    trans.Start()
    try:
        for entry in plan:
            create_one_manual(document, entry, params, report)
        trans.Commit()
    except Exception:
        try:
            trans.RollBack()
        except Exception:
            pass
        raise
    return report


def create_one_manual(document, entry, params, report):
    cat_name = entry["cat_name"]
    try:
        if entry["delete_view"] is not None:
            document.Delete(entry["delete_view"].Id)

        schedule = ViewSchedule.CreateSchedule(document, entry["cat_id"])
    except Exception as err:
        report.add_row(cat_name, entry["name"], None, 0, len(params),
                       "ERRORE - abaco non creato")
        report.add_issue(cat_name, "Creazione fallita", str(err))
        return

    try:
        schedule.Name = entry["name"]
    except Exception as err:
        report.add_issue(cat_name, "Rinomina",
                         "Nome '{0}' rifiutato da Revit ({1}). "
                         "Mantenuto '{2}'.".format(entry["name"], err,
                                                   schedule.Name))
        entry["name"] = schedule.Name

    definition = schedule.Definition
    schedulable = list(definition.GetSchedulableFields())

    added = 0
    dropped = []
    for field_info in params:
        sfield = match_schedulable(document, schedulable,
                                   field_info.param_id_value,
                                   field_info.field_type,
                                   field_info.name)
        if sfield is None:
            dropped.append(field_info.name)
            continue
        try:
            definition.AddField(sfield)
            added += 1
        except Exception as err:
            dropped.append("{0} ({1})".format(field_info.name, err))

    if dropped:
        report.add_issue(
            cat_name, "Parametri non disponibili",
            "{0} campi scartati: {1}".format(len(dropped), "; ".join(dropped)))

    status = "OK" if not dropped else "OK con scarti"
    report.add_row(cat_name, entry["name"], schedule.Id, added,
                   len(dropped), status)


# ---------------------------------------------------------------------------
# CREAZIONE - MODALITA TEMPLATE
# ---------------------------------------------------------------------------

def run_template(document, config):
    template = config["template"]
    report = Report("Da template (shift+click) - '{0}'".format(template.Name),
                    config["prefix"])

    specs, fid_to_key = read_template_fields(document, template)
    is_takeoff = False
    try:
        is_takeoff = bool(template.Definition.IsMaterialTakeoff)
    except Exception:
        pass

    calc_fields = [s["name"] for s in specs if s["is_calc"]]
    if calc_fields:
        report.add_note(
            "Campi calcolati/combinati presenti nel template e non "
            "replicabili tramite API: {0}.".format(", ".join(calc_fields)))
    if not config["copy_filters"]:
        report.add_note("I filtri del template non sono stati copiati "
                        "(opzione disattivata).")

    plan, cancelled = build_plan(document, config["prefix"],
                                 config["categories"], report,
                                 protected_view=template)
    if cancelled:
        return None
    if not plan:
        return report

    trans = Transaction(document, "pyRevit - Crea abachi da template")
    trans.Start()
    try:
        for entry in plan:
            create_one_from_template(document, entry, template, specs,
                                     fid_to_key, is_takeoff, config, report)
        trans.Commit()
    except Exception:
        try:
            trans.RollBack()
        except Exception:
            pass
        raise
    return report


def create_one_from_template(document, entry, template, specs, fid_to_key,
                             is_takeoff, config, report):
    cat_name = entry["cat_name"]
    try:
        if entry["delete_view"] is not None:
            document.Delete(entry["delete_view"].Id)

        if is_takeoff:
            schedule = ViewSchedule.CreateMaterialTakeoff(document, entry["cat_id"])
        else:
            schedule = ViewSchedule.CreateSchedule(document, entry["cat_id"])
    except Exception as err:
        report.add_row(cat_name, entry["name"], None, 0, len(specs),
                       "ERRORE - abaco non creato")
        report.add_issue(cat_name, "Creazione fallita", str(err))
        return

    try:
        schedule.Name = entry["name"]
    except Exception as err:
        report.add_issue(cat_name, "Rinomina",
                         "Nome '{0}' rifiutato da Revit ({1}). "
                         "Mantenuto '{2}'.".format(entry["name"], err,
                                                   schedule.Name))
        entry["name"] = schedule.Name

    definition = schedule.Definition
    schedulable = list(definition.GetSchedulableFields())

    # le proprieta' generali vanno impostate prima dei campi: alcune
    # opzioni di colonna (es. DisplayType) dipendono da IsItemized
    if config["copy_format"]:
        copy_definition_properties(template.Definition, definition)

    added = 0
    dropped = []
    key_to_fid = {}

    for spec in specs:
        if spec["is_calc"]:
            dropped.append("{0} (campo calcolato)".format(spec["name"]))
            continue

        sfield = match_schedulable(document, schedulable,
                                   spec["param_id_value"],
                                   spec["field_type"], spec["name"])
        if sfield is None:
            dropped.append(spec["name"])
            continue

        try:
            new_field = definition.AddField(sfield)
        except Exception as err:
            dropped.append("{0} ({1})".format(spec["name"], err))
            continue

        added += 1
        key_to_fid[spec["key"]] = new_field.FieldId

        if config["copy_format"]:
            copy_field_properties(spec, new_field)

    # template di vista
    if config["copy_format"]:
        try:
            if template.ViewTemplateId is not None \
                    and eid_value(template.ViewTemplateId) > 0:
                schedule.ViewTemplateId = template.ViewTemplateId
        except Exception as err:
            report.add_issue(cat_name, "Template di vista",
                             "Non applicato: {0}".format(err))

    # ordinamento e raggruppamento
    if config["copy_sort"]:
        lost = copy_sort_group(template.Definition, definition,
                               fid_to_key, key_to_fid)
        if lost:
            report.add_issue(
                cat_name, "Ordinamento/raggruppamento",
                "Regole scartate perche' il campo non esiste nella "
                "categoria: {0}".format("; ".join(lost)))

    # filtri
    if config["copy_filters"]:
        lost = copy_filters(template.Definition, definition,
                            fid_to_key, key_to_fid)
        if lost:
            report.add_issue(
                cat_name, "Filtri",
                "Filtri scartati: {0}".format("; ".join(lost)))

    if dropped:
        report.add_issue(
            cat_name, "Campi non disponibili",
            "{0} campi scartati: {1}".format(len(dropped), "; ".join(dropped)))

    status = "OK" if not dropped else "OK con scarti"
    report.add_row(cat_name, entry["name"], schedule.Id, added,
                   len(dropped), status)


def copy_field_properties(spec, new_field):
    for prop, value in spec["props"].items():
        try:
            setattr(new_field, prop, value)
        except Exception:
            continue
    if spec["format"] is not None:
        try:
            new_field.SetFormatOptions(spec["format"])
        except Exception:
            pass
    if spec["style"] is not None:
        try:
            new_field.SetStyle(spec["style"])
        except Exception:
            pass


def copy_definition_properties(src_def, dst_def):
    for prop in DEF_PROPS:
        try:
            setattr(dst_def, prop, getattr(src_def, prop))
        except Exception:
            continue


def copy_sort_group(src_def, dst_def, fid_to_key, key_to_fid):
    lost = []
    new_rules = List[ScheduleSortGroupField]()
    try:
        source_rules = src_def.GetSortGroupFields()
    except Exception:
        return lost

    for rule in source_rules:
        try:
            key = fid_to_key.get(rule.FieldId.IntegerValue)
            new_fid = key_to_fid.get(key) if key else None
            if new_fid is None:
                try:
                    label = src_def.GetField(rule.FieldId).GetName()
                except Exception:
                    label = "campo sconosciuto"
                lost.append(label)
                continue
            new_rule = ScheduleSortGroupField(new_fid, rule.SortOrder)
            for prop in ("ShowHeader", "ShowFooter", "ShowFooterTitle",
                         "ShowFooterCount", "ShowBlankLine"):
                try:
                    setattr(new_rule, prop, getattr(rule, prop))
                except Exception:
                    continue
            new_rules.Add(new_rule)
        except Exception as err:
            lost.append(str(err))

    if new_rules.Count:
        try:
            dst_def.SetSortGroupFields(new_rules)
        except Exception as err:
            lost.append("applicazione regole fallita: {0}".format(err))
    return lost


def build_filter_candidates(src_filter, new_field_id):
    """ScheduleFilter.FieldId e' in sola lettura: il filtro va ricostruito.

    Il tipo del valore non e' esposto direttamente, quindi si provano i
    getter disponibili in ordine di specificita' e si restituiscono i
    ScheduleFilter candidati da tentare con AddFilter.
    """
    filter_type = src_filter.FilterType
    candidates = []
    for getter in ("GetElementIdValue", "GetStringValue",
                   "GetDoubleValue", "GetIntegerValue"):
        try:
            value = getattr(src_filter, getter)()
        except Exception:
            continue
        if value is None:
            continue
        try:
            candidates.append(ScheduleFilter(new_field_id, filter_type, value))
        except Exception:
            continue
    # filtri senza valore (es. "ha un valore" / "non ha valore")
    try:
        candidates.append(ScheduleFilter(new_field_id, filter_type))
    except Exception:
        pass
    return candidates


def copy_filters(src_def, dst_def, fid_to_key, key_to_fid):
    lost = []
    try:
        source_filters = src_def.GetFilters()
    except Exception:
        return lost

    for sched_filter in source_filters:
        label = "filtro"
        try:
            label = src_def.GetField(sched_filter.FieldId).GetName()
        except Exception:
            pass
        try:
            key = fid_to_key.get(sched_filter.FieldId.IntegerValue)
            new_fid = key_to_fid.get(key) if key else None
            if new_fid is None:
                lost.append("{0} (campo assente)".format(label))
                continue

            applied = False
            for candidate in build_filter_candidates(sched_filter, new_fid):
                try:
                    dst_def.AddFilter(candidate)
                    applied = True
                    break
                except Exception:
                    continue
            if not applied:
                lost.append("{0} (valore non replicabile)".format(label))
        except Exception as err:
            lost.append("{0} ({1})".format(label, err))
    return lost


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    if doc is None:
        forms.alert("Nessun documento aperto.", title="Crea Abachi",
                    exitscript=True)
    if doc.IsFamilyDocument:
        forms.alert("Comando disponibile solo nei documenti di progetto.",
                    title="Crea Abachi", exitscript=True)

    for path in (XAML_MAIN, XAML_TEMPLATE, XAML_CONFLICT):
        if not os.path.isfile(path):
            forms.alert("File di interfaccia mancante:\n{0}".format(path),
                        title="Crea Abachi", exitscript=True)

    if SHIFT:
        window = TemplateWindow(doc)
        window.show_dialog()
        if not window.result:
            return
        report = run_template(doc, window.result)
    else:
        window = ManualWindow(doc)
        window.show_dialog()
        if not window.result:
            return
        report = run_manual(doc, window.result)

    if report is None:
        return
    report.render()


try:
    main()
except Exception as exc:                                  # noqa: BLE001
    logger.error(traceback.format_exc())
    forms.alert("Errore imprevisto durante l'esecuzione:\n\n{0}\n\n"
                "Il dettaglio completo e' nel log di pyRevit.".format(exc),
                title="Crea Abachi")
