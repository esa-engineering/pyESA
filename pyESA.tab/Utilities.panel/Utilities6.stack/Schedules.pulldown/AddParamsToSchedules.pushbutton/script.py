# -*- coding: utf-8 -*-
"""Aggiunge uno o piu' parametri a una o piu' schedule.

Sorgenti dei parametri selezionabili:
  - parametri di progetto e condivisi (doc.ParameterBindings);
  - parametri nativi di Revit, ricavati dall'unione dei campi schedulabili
    (ScheduleDefinition.GetSchedulableFields) di tutte le schedule del progetto.

Se un parametro non e' disponibile (non schedulabile) per la categoria della
singola schedule, viene saltato senza interrompere l'elaborazione e riportato
nel report finale insieme al motivo.

Compatibile con IronPython 2.7 (engine di default pyRevit) e con CPython3.
"""

__title__ = "Aggiungi Parametri\nalle Schedule"
__author__ = "ESA Engineering"
__min_revit_ver__ = 2019
__doc__ = ("Seleziona una o piu' schedule e uno o piu' parametri (di progetto, "
           "condivisi o nativi di Revit) da aggiungere come campi.\n"
           "I parametri non applicabili alla categoria della schedule vengono "
           "saltati e riportati nel report finale.")

import traceback

from System.Collections.Generic import List

from pyrevit import revit, DB, forms, script
from pyrevit.forms import WPFWindow


doc = revit.doc
logger = script.get_logger()

# valori di esito usati anche dagli stili XAML del report
OUT_ADDED = u"Aggiunto"
OUT_EXISTS = u"Già presente"
OUT_NOT_AVAILABLE = u"Non disponibile"
OUT_ERROR = u"Errore"

# tipi di parametro
KIND_PROJECT = u"Progetto"
KIND_SHARED = u"Condiviso"
KIND_NATIVE = u"Nativo"

# traduzione di ScheduleFieldType
FIELD_TYPE_IT = {
    "Instance": u"Istanza",
    "ElementType": u"Tipo",
    "Count": u"Conteggio",
    "ViewBased": u"Da vista",
    "ProjectInfo": u"Info progetto",
    "MaterialQuantity": u"Quantità materiale",
    "Formula": u"Formula",
    "Percentage": u"Percentuale",
    "CombinedParameter": u"Parametro combinato",
}


# ---------------------------------------------------------------------------
# helper compatibilita' API
# ---------------------------------------------------------------------------
def eid_value(element_id):
    """ElementId -> int, compatibile Revit <=2023 (IntegerValue) e >=2024 (Value)."""
    if element_id is None:
        return None
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def safe_str(value, default=u""):
    if value is None:
        return default
    try:
        return unicode(value)  # noqa  IronPython 2.7
    except NameError:
        return str(value)
    except Exception:
        return default


def api_get(obj, name, default=None):
    """getattr difensivo: le proprieta' non presenti nella versione di API in uso
    non devono mai far scartare l'elemento."""
    try:
        return getattr(obj, name)
    except Exception:
        return default


def category_name(cat_id):
    """Nome leggibile di una categoria a partire dal suo ElementId."""
    if cat_id is None:
        return u"-"
    val = eid_value(cat_id)
    if val is None or val == -1:
        return u"Nessuna categoria"
    if val == -2000001:            # OST_MultiCategory
        return u"Multi-categoria"
    try:
        cat = DB.Category.GetCategory(doc, cat_id)
        if cat is not None:
            return cat.Name
    except Exception:
        pass
    return u"Categoria {}".format(val)


def param_type_name(definition):
    """Tipo dato del parametro, con fallback tra API vecchie e nuove."""
    try:
        return DB.LabelUtils.GetLabelForSpec(definition.GetDataType())
    except Exception:
        pass
    try:
        return safe_str(api_get(definition, "ParameterType"), u"-")
    except Exception:
        return u"-"


def field_type_name(schedulable_field):
    raw = safe_str(api_get(schedulable_field, "FieldType"), u"")
    return FIELD_TYPE_IT.get(raw, raw or u"-")


# ---------------------------------------------------------------------------
# modelli dati per le liste
# ---------------------------------------------------------------------------
class ScheduleItem(object):
    """Riga della lista schedule."""

    def __init__(self, view_schedule):
        self.view = view_schedule
        self.checked = False
        self._cb = None                       # CheckBox della riga, se a video
        self.idx = None                       # indice usato dall'indice dei campi
        self.name = api_get(view_schedule, "Name", u"<senza nome>")
        sdef = view_schedule.Definition
        self.category = category_name(api_get(sdef, "CategoryId"))
        flags = []
        if api_get(sdef, "IsKeySchedule", False):
            flags.append(u"Key Schedule")
        if api_get(sdef, "IsMaterialTakeoff", False):
            flags.append(u"Material Takeoff")
        try:
            n_fields = len(list(sdef.GetFieldOrder()))
        except Exception:
            n_fields = 0
        flags.append(u"{} campi".format(n_fields))
        self.info = u"{}  |  {}".format(self.category, u"  |  ".join(flags))

    def __str__(self):
        return self.name


class ParamItem(object):
    """Riga della lista parametri, sia di progetto/condivisi sia nativi."""

    def __init__(self, name, param_id, kind, detail,
                 category_ids=None, sched_set=None, sched_total=0):
        self.name = name
        self.param_id = param_id
        self.kind = kind                      # Progetto / Condiviso / Nativo
        self.checked = False
        self._cb = None
        self.category_ids = category_ids or set()
        self.sched_set = sched_set or set()   # indici delle schedule dove e' schedulabile
        avail = u"disponibile in {}/{} schedule".format(
            len(self.sched_set), sched_total) if sched_total else u""
        parts = [kind] + [p for p in detail if p] + ([avail] if avail else [])
        self.info = u"  |  ".join(parts)

    @property
    def is_native(self):
        return self.kind == KIND_NATIVE

    def applies_to_category(self, cat_id_value):
        """Solo per i parametri di progetto/condivisi: il binding copre la categoria?"""
        if not self.category_ids:
            return True
        return cat_id_value in self.category_ids

    def available_in(self, schedule_indices):
        """True se schedulabile in tutte le schedule indicate."""
        if not schedule_indices:
            return True
        return schedule_indices.issubset(self.sched_set)

    def __str__(self):
        return self.name


class ResultItem(object):
    """Riga del report finale."""

    def __init__(self, schedule, parameter, outcome, detail):
        self.schedule = schedule
        self.parameter = parameter
        self.outcome = outcome
        self.detail = detail


# ---------------------------------------------------------------------------
# raccolta dati dal modello
# ---------------------------------------------------------------------------
def collect_schedules():
    """Schedule utente del progetto, escluse view template, revision schedule dei
    cartigli e keynote schedule interne.

    Ritorna (items, totale_esaminate, scartate): nessuna schedule viene mai
    scartata in silenzio.
    """
    items = []
    skipped = []
    total = 0
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)
    for vs in collector:
        total += 1
        name = api_get(vs, "Name", u"<senza nome>")
        try:
            if api_get(vs, "IsTemplate", False):
                continue
            # NB: queste due proprieta' stanno su ViewSchedule, non su ScheduleDefinition
            if api_get(vs, "IsTitleblockRevisionSchedule", False):
                continue
            if api_get(vs, "IsInternalKeynoteSchedule", False):
                continue
            if vs.Definition is None:
                skipped.append(u"{}: Definition non disponibile".format(name))
                continue
            items.append(ScheduleItem(vs))
        except Exception as ex:
            logger.debug(traceback.format_exc())
            skipped.append(u"{}: {}".format(name, safe_str(ex)))
    items.sort(key=lambda i: i.name.lower())
    for i, item in enumerate(items):
        item.idx = i
    return items, total, skipped


def field_key(param_id_value, name):
    """Chiave di deduplica: l'id del parametro se valido, altrimenti il nome."""
    if param_id_value is not None and param_id_value != -1:
        return param_id_value
    return u"nome:" + name


def build_schedulable_index(schedule_items):
    """Unione dei campi schedulabili di tutte le schedule.

    Ritorna (index, interrotto) dove index e' {chiave: {name, pid, ftype, scheds}}
    e scheds e' l'insieme degli indici delle schedule in cui il campo e'
    schedulabile. E' la sorgente dei parametri nativi ed e' anche cio' che
    permette di dire in quante schedule un parametro e' effettivamente usabile.
    """
    index = {}
    total = len(schedule_items)
    progress = None
    try:
        progress = forms.ProgressBar(
            title=u"Analisi campi schedulabili... {value}/{max_value}",
            cancellable=True, step=max(1, total // 60))
    except Exception:
        logger.debug("ProgressBar non disponibile")

    cancelled = False
    try:
        for i, s_item in enumerate(schedule_items):
            if progress is not None and api_get(progress, "cancelled", False):
                cancelled = True
                break
            try:
                fields = list(s_item.view.Definition.GetSchedulableFields())
            except Exception as ex:
                logger.debug("Campi schedulabili non leggibili per %s: %s",
                             s_item.name, ex)
                fields = []
            for sfield in fields:
                try:
                    pid = eid_value(api_get(sfield, "ParameterId"))
                    try:
                        fname = sfield.GetName(doc)
                    except Exception:
                        fname = None
                    if not fname:
                        continue
                    key = field_key(pid, fname)
                    entry = index.get(key)
                    if entry is None:
                        entry = {"name": fname, "pid": pid,
                                 "ftype": field_type_name(sfield),
                                 "scheds": set()}
                        index[key] = entry
                    entry["scheds"].add(s_item.idx)
                except Exception:
                    logger.debug(traceback.format_exc())
            if progress is not None:
                try:
                    progress.update_progress(i + 1, total)
                except Exception:
                    pass
    finally:
        if progress is not None:
            try:
                progress.Close()
            except Exception:
                pass
    return index, cancelled


def collect_parameters(index, sched_total):
    """Parametri di progetto/condivisi (dai binding) piu' parametri nativi
    (dall'indice dei campi schedulabili).

    Ritorna (items, totale_binding_esaminati, scartati).
    """
    items = []
    skipped = []
    total = 0
    seen_ids = set()
    seen_names = set()

    # --- parametri di progetto e condivisi --------------------------------
    iterator = doc.ParameterBindings.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        total += 1
        name = u"<parametro sconosciuto>"
        try:
            definition = iterator.Key
            binding = iterator.Current
            if definition is None or binding is None:
                skipped.append(u"binding vuoto")
                continue
            name = api_get(definition, "Name", name)
            pid = api_get(definition, "Id")
            pid_val = eid_value(pid)

            is_shared = False
            if pid is not None:
                is_shared = isinstance(doc.GetElement(pid), DB.SharedParameterElement)

            cat_ids = set()
            try:
                for cat in binding.Categories:
                    cat_ids.add(eid_value(cat.Id))
            except Exception:
                pass

            entry = index.get(field_key(pid_val, name))
            items.append(ParamItem(
                name=name,
                param_id=pid,
                kind=KIND_SHARED if is_shared else KIND_PROJECT,
                detail=[u"Istanza" if isinstance(binding, DB.InstanceBinding) else u"Tipo",
                        param_type_name(definition),
                        u"{} categorie".format(len(cat_ids))],
                category_ids=cat_ids,
                sched_set=entry["scheds"] if entry else set(),
                sched_total=sched_total))
            if pid_val is not None:
                seen_ids.add(pid_val)
            seen_names.add(name)
        except Exception as ex:
            logger.debug(traceback.format_exc())
            skipped.append(u"{}: {}".format(name, safe_str(ex)))

    # --- parametri nativi -------------------------------------------------
    for entry in index.values():
        pid = entry["pid"]
        if pid is not None and pid in seen_ids:
            continue
        if entry["name"] in seen_names:
            continue
        items.append(ParamItem(
            name=entry["name"],
            param_id=DB.ElementId(pid) if pid is not None and pid != -1 else None,
            kind=KIND_NATIVE,
            detail=[entry["ftype"]],
            sched_set=entry["scheds"],
            sched_total=sched_total))

    items.sort(key=lambda i: i.name.lower())
    return items, total, skipped


def diagnostic_message(header, total, skipped):
    """Messaggio d'errore che spiega perche' la lista e' vuota."""
    msg = u"{}\n\nElementi esaminati nel documento: {}".format(header, total)
    if skipped:
        msg += u"\nScartati: {}\n\nPrimi motivi:\n- {}".format(
            len(skipped), u"\n- ".join(skipped[:6]))
    return msg


# ---------------------------------------------------------------------------
# logica di aggiunta campi
# ---------------------------------------------------------------------------
def build_schedulable_maps(sdef):
    """Mappe (id parametro -> SchedulableField) e (nome -> SchedulableField)."""
    by_id = {}
    by_name = {}
    try:
        fields = list(sdef.GetSchedulableFields())
    except Exception:
        fields = []
    for sfield in fields:
        try:
            key = eid_value(sfield.ParameterId)
            if key is not None and key not in by_id:
                by_id[key] = sfield
        except Exception:
            pass
        try:
            fname = sfield.GetName(doc)
            if fname and fname not in by_name:
                by_name[fname] = sfield
        except Exception:
            pass
    return by_id, by_name


def build_existing_maps(sdef):
    """Id parametro e nomi dei campi gia' presenti nella schedule."""
    ids = set()
    names = set()
    try:
        for fid in sdef.GetFieldOrder():
            field = sdef.GetField(fid)
            val = eid_value(field.ParameterId)
            if val is not None:
                ids.add(val)
            try:
                names.add(field.GetName())
            except Exception:
                pass
    except Exception:
        pass
    return ids, names


def add_parameters(schedule_items, param_items, hide_fields=False):
    """Aggiunge i parametri alle schedule. Ritorna la lista di ResultItem."""
    results = []

    tgroup = DB.TransactionGroup(doc, "Aggiungi parametri alle schedule")
    tgroup.Start()
    try:
        for s_item in schedule_items:
            sdef = s_item.view.Definition
            sched_cat = eid_value(api_get(sdef, "CategoryId"))
            by_id, by_name = build_schedulable_maps(sdef)
            existing_ids, existing_names = build_existing_maps(sdef)

            trans = DB.Transaction(doc, "Campi schedule: {}".format(s_item.name))
            trans.Start()
            added_any = False
            try:
                for p_item in param_items:
                    pid_value = eid_value(p_item.param_id)

                    # 1) campo gia' presente nella schedule
                    if pid_value in existing_ids or p_item.name in existing_names:
                        results.append(ResultItem(
                            s_item.name, p_item.name, OUT_EXISTS,
                            u"Il campo è già presente nella schedule."))
                        continue

                    # 2) parametro non schedulabile per questa schedule
                    sfield = by_id.get(pid_value)
                    if sfield is None:
                        sfield = by_name.get(p_item.name)
                    if sfield is None:
                        if not p_item.applies_to_category(sched_cat):
                            reason = (u"Parametro non associato alla categoria "
                                      u"'{}' della schedule.".format(s_item.category))
                        else:
                            reason = (u"Parametro non presente tra i campi "
                                      u"schedulabili di questa schedule "
                                      u"(categoria '{}').".format(s_item.category))
                        results.append(ResultItem(
                            s_item.name, p_item.name, OUT_NOT_AVAILABLE, reason))
                        continue

                    # 3) aggiunta effettiva
                    try:
                        field = sdef.AddField(sfield)
                        if hide_fields and field is not None:
                            field.IsHidden = True
                        if pid_value is not None:
                            existing_ids.add(pid_value)
                        existing_names.add(p_item.name)
                        added_any = True
                        results.append(ResultItem(
                            s_item.name, p_item.name, OUT_ADDED,
                            u"Campo aggiunto{}.".format(
                                u" (nascosto)" if hide_fields else u"")))
                    except Exception as ex:
                        logger.debug(traceback.format_exc())
                        results.append(ResultItem(
                            s_item.name, p_item.name, OUT_ERROR,
                            u"Errore API Revit: {}".format(safe_str(ex))))

                if added_any:
                    trans.Commit()
                else:
                    trans.RollBack()
            except Exception as ex:
                # errore non gestito sulla singola schedule: annulla solo questa
                if trans.HasStarted() and not trans.HasEnded():
                    trans.RollBack()
                logger.debug(traceback.format_exc())
                results.append(ResultItem(
                    s_item.name, u"-", OUT_ERROR,
                    u"Schedule non modificata: {}".format(safe_str(ex))))

        tgroup.Assimilate()
    except Exception as ex:
        if tgroup.HasStarted() and not tgroup.HasEnded():
            tgroup.RollBack()
        logger.error("Operazione annullata: %s", safe_str(ex))
        results.append(ResultItem(u"-", u"-", OUT_ERROR,
                                  u"Operazione annullata: {}".format(safe_str(ex))))
    return results


# ---------------------------------------------------------------------------
# finestra di report
# ---------------------------------------------------------------------------
class ReportWindow(WPFWindow):
    def __init__(self, results):
        WPFWindow.__init__(self, "ReportWindow.xaml")
        self.results = results
        self._refresh()

    def _counts(self):
        counts = {OUT_ADDED: 0, OUT_EXISTS: 0,
                  OUT_NOT_AVAILABLE: 0, OUT_ERROR: 0}
        for res in self.results:
            if res.outcome in counts:
                counts[res.outcome] += 1
        return counts

    def _visible_results(self):
        if self.only_issues.IsChecked:
            return [r for r in self.results if r.outcome != OUT_ADDED]
        return list(self.results)

    def _refresh(self):
        data = List[object]()
        for res in self._visible_results():
            data.Add(res)
        self.results_grid.ItemsSource = data

        counts = self._counts()
        self.summary_text.Text = (
            u"Aggiunti: {}     Già presenti: {}     "
            u"Non disponibili: {}     Errori: {}".format(
                counts[OUT_ADDED], counts[OUT_EXISTS],
                counts[OUT_NOT_AVAILABLE], counts[OUT_ERROR]))

    def _as_text(self, separator=u"\t"):
        lines = [separator.join([u"Schedule", u"Parametro", u"Esito", u"Dettaglio"])]
        for res in self.results:
            lines.append(separator.join([
                safe_str(res.schedule), safe_str(res.parameter),
                safe_str(res.outcome), safe_str(res.detail)]))
        return u"\r\n".join(lines)

    # -- eventi -------------------------------------------------------------
    def filter_changed(self, sender, args):
        self._refresh()

    def copy_click(self, sender, args):
        try:
            from System.Windows import Clipboard
            Clipboard.SetText(self._as_text(u"\t"))
            forms.toast("Report copiato negli appunti.")
        except Exception as ex:
            forms.alert(u"Copia non riuscita: {}".format(safe_str(ex)))

    def export_click(self, sender, args):
        try:
            path = forms.save_file(file_ext="csv",
                                   default_name="report_parametri_schedule")
            if not path:
                return
            handle = open(path, "wb")
            try:
                handle.write(self._as_text(u";").encode("utf-8-sig"))
            finally:
                handle.close()
            forms.alert(u"Report salvato in:\n{}".format(path))
        except Exception as ex:
            forms.alert(u"Esportazione non riuscita: {}".format(safe_str(ex)))

    def close_click(self, sender, args):
        self.Close()


# ---------------------------------------------------------------------------
# finestra principale
# ---------------------------------------------------------------------------
class AddParamsWindow(WPFWindow):
    def __init__(self, schedule_items, param_items, scan_incomplete=False):
        WPFWindow.__init__(self, "AddParamsWindow.xaml")
        self.all_schedules = schedule_items
        self.all_params = param_items
        self.scan_incomplete = scan_incomplete
        self.confirmed = False
        self._bind(self.schedules_list, self.all_schedules)
        self._bind(self.params_list, self._visible_params())
        self._update_counters()

    # -- helper -------------------------------------------------------------
    @staticmethod
    def _bind(listbox, items):
        data = List[object]()
        for item in items:
            data.Add(item)
        listbox.ItemsSource = data

    @staticmethod
    def _text_filter(items, text):
        text = (text or u"").strip().lower()
        if not text:
            return items
        tokens = text.split()
        return [i for i in items
                if all(tok in (i.name + u" " + i.info).lower() for tok in tokens)]

    @staticmethod
    def _set_checked(items, value):
        """Aggiorna stato logico e, se la riga e' a video, anche la CheckBox."""
        for item in items:
            item.checked = value
            checkbox = getattr(item, "_cb", None)
            if checkbox is not None:
                try:
                    checkbox.IsChecked = value
                except Exception:
                    pass

    def _ready(self):
        """False mentre lo XAML e' ancora in caricamento."""
        return getattr(self, "all_params", None) is not None

    def _visible_schedules(self):
        return self._text_filter(self.all_schedules, self.schedules_filter.Text)

    def _visible_params(self):
        items = self.all_params
        # filtro sorgente
        if self.src_project.IsChecked:
            items = [i for i in items if not i.is_native]
        elif self.src_native.IsChecked:
            items = [i for i in items if i.is_native]
        # filtro disponibilita' rispetto alle schedule selezionate
        if self.only_common.IsChecked:
            sel = self.selected_schedule_indices()
            if sel:
                items = [i for i in items if i.available_in(sel)]
        return self._text_filter(items, self.params_filter.Text)

    def selected_schedules(self):
        return [i for i in self.all_schedules if i.checked]

    def selected_schedule_indices(self):
        return set(i.idx for i in self.selected_schedules() if i.idx is not None)

    def selected_params(self):
        return [i for i in self.all_params if i.checked]

    def _update_counters(self):
        n_sched = len(self.selected_schedules())
        self.schedules_count.Text = u"{} selezionate / {} totali".format(
            n_sched, len(self.all_schedules))

        note = u""
        if self.only_common.IsChecked and not n_sched:
            note = u"  |  seleziona prima le schedule per applicare il filtro"
        if self.scan_incomplete:
            note += u"  |  scansione interrotta: elenco nativi parziale"
        self.params_count.Text = u"{} selezionati  |  {} mostrati su {}{}".format(
            len(self.selected_params()), len(self._visible_params()),
            len(self.all_params), note)

    # -- eventi -------------------------------------------------------------
    def item_loaded(self, sender, args):
        """Aggancia la CheckBox all'elemento e ne ripristina lo stato.

        Sostituisce il databinding su IsChecked: cosi' il comportamento non
        dipende da INotifyPropertyChanged ne' dall'engine Python in uso, e
        resta corretto con la virtualizzazione attiva.
        """
        item = getattr(sender, "DataContext", None)
        if item is None:
            return
        try:
            item._cb = sender
            sender.IsChecked = bool(item.checked)
        except Exception:
            pass

    def _on_item_toggle(self, sender, value):
        item = getattr(sender, "DataContext", None)
        if item is None:
            return
        item.checked = value
        # cambiare la selezione delle schedule cambia il filtro di disponibilita'
        if isinstance(item, ScheduleItem) and self.only_common.IsChecked:
            self._bind(self.params_list, self._visible_params())
        self._update_counters()

    def item_checked(self, sender, args):
        self._on_item_toggle(sender, True)

    def item_unchecked(self, sender, args):
        self._on_item_toggle(sender, False)

    def schedules_filter_changed(self, sender, args):
        if not self._ready():
            return
        self._bind(self.schedules_list, self._visible_schedules())

    def params_filter_changed(self, sender, args):
        if not self._ready():
            return
        self._bind(self.params_list, self._visible_params())
        self._update_counters()

    def params_source_changed(self, sender, args):
        if not self._ready():
            return
        self._bind(self.params_list, self._visible_params())
        self._update_counters()

    def schedules_check_all(self, sender, args):
        self._set_checked(self._visible_schedules(), True)
        if self.only_common.IsChecked:
            self._bind(self.params_list, self._visible_params())
        self._update_counters()

    def schedules_uncheck_all(self, sender, args):
        self._set_checked(self.all_schedules, False)
        if self.only_common.IsChecked:
            self._bind(self.params_list, self._visible_params())
        self._update_counters()

    def params_check_all(self, sender, args):
        self._set_checked(self._visible_params(), True)
        self._update_counters()

    def params_uncheck_all(self, sender, args):
        self._set_checked(self.all_params, False)
        self._update_counters()

    def run_click(self, sender, args):
        if not self.selected_schedules():
            forms.alert("Seleziona almeno una schedule.", title="Selezione mancante")
            return
        if not self.selected_params():
            forms.alert("Seleziona almeno un parametro.", title="Selezione mancante")
            return
        self.confirmed = True
        self.Close()

    def cancel_click(self, sender, args):
        self.confirmed = False
        self.Close()


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def main():
    if doc is None:
        forms.alert("Nessun documento attivo.", exitscript=True)
    if doc.IsFamilyDocument:
        forms.alert("Il comando funziona solo nei file di progetto (.rvt).",
                    exitscript=True)

    schedules, sched_total, sched_skipped = collect_schedules()
    if not schedules:
        forms.alert(diagnostic_message(
            u"Nessuna schedule utilizzabile trovata nel progetto.",
            sched_total, sched_skipped), exitscript=True)
    if sched_skipped:
        logger.debug("Schedule scartate: %s", sched_skipped)

    index, cancelled = build_schedulable_index(schedules)

    params, param_total, param_skipped = collect_parameters(index, len(schedules))
    if not params:
        forms.alert(diagnostic_message(
            u"Nessun parametro utilizzabile trovato nel modello.",
            param_total, param_skipped), exitscript=True)
    if param_skipped:
        logger.debug("Parametri scartati: %s", param_skipped)

    window = AddParamsWindow(schedules, params, scan_incomplete=cancelled)
    window.show_dialog()
    if not window.confirmed:
        script.exit()

    results = add_parameters(window.selected_schedules(),
                             window.selected_params(),
                             hide_fields=bool(window.hidden_fields.IsChecked))

    ReportWindow(results).show_dialog()


if __name__ == "__main__":
    main()
