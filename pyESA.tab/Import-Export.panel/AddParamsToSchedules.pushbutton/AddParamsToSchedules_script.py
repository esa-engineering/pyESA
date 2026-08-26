# -*- coding: utf-8 -*-
"""Aggiunge uno o piu' parametri di progetto/condivisi a una o piu' schedule.

Se un parametro non e' disponibile (non schedulabile) per la categoria della
singola schedule, viene saltato senza interrompere l'elaborazione e riportato
nel report finale insieme al motivo.

Compatibile con IronPython 2.7 (engine di default pyRevit) e con CPython3.
"""

__title__ = "Aggiungi Parametri\nalle Schedule"
__author__ = "ESA Engineering"
__min_revit_ver__ = 2019
__doc__ = ("Seleziona una o piu' schedule e uno o piu' parametri di progetto/"
           "condivisi da aggiungere come campi.\n"
           "I parametri non applicabili alla categoria della schedule vengono "
           "saltati e riportati nel report finale.")

import os
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


def safe_str(value, default=""):
    if value is None:
        return default
    try:
        return unicode(value)  # noqa  IronPython 2.7
    except NameError:
        return str(value)
    except Exception:
        return default


def category_name(cat_id):
    """Nome leggibile di una categoria a partire dal suo ElementId."""
    if cat_id is None:
        return "-"
    try:
        cat = DB.Category.GetCategory(doc, cat_id)
        if cat is not None:
            return cat.Name
    except Exception:
        pass
    val = eid_value(cat_id)
    if val is None or val == -1:
        return "-"
    return "Categoria {}".format(val)


def param_type_name(definition):
    """Tipo del parametro, con fallback tra API vecchie e nuove."""
    try:
        spec_id = definition.GetDataType()
        return DB.LabelUtils.GetLabelForSpec(spec_id)
    except Exception:
        pass
    try:
        return safe_str(definition.ParameterType)
    except Exception:
        return "-"


# ---------------------------------------------------------------------------
# modelli dati per il databinding WPF
# ---------------------------------------------------------------------------
class ScheduleItem(object):
    """Riga della lista schedule."""

    def __init__(self, view_schedule):
        self.view = view_schedule
        self.checked = False
        self.name = view_schedule.Name
        sdef = view_schedule.Definition
        self.category = category_name(sdef.CategoryId)
        flags = []
        if sdef.IsKeySchedule:
            flags.append("Key Schedule")
        try:
            if sdef.IsMaterialTakeoff:
                flags.append("Material Takeoff")
        except Exception:
            pass
        try:
            n_fields = len(list(sdef.GetFieldOrder()))
        except Exception:
            n_fields = 0
        flags.append("{} campi".format(n_fields))
        self.info = "{}  |  {}".format(self.category, "  |  ".join(flags))

    def __str__(self):
        return self.name


class ParamItem(object):
    """Riga della lista parametri (parametri di progetto e condivisi)."""

    def __init__(self, definition, binding, is_shared):
        self.definition = definition
        self.checked = False
        self.name = definition.Name
        self.param_id = getattr(definition, "Id", None)
        self.is_shared = is_shared
        self.is_instance = isinstance(binding, DB.InstanceBinding)
        self.category_ids = set()
        self.category_names = []
        try:
            for cat in binding.Categories:
                self.category_ids.add(eid_value(cat.Id))
                self.category_names.append(cat.Name)
        except Exception:
            pass
        self.category_names.sort()
        self.kind = "Condiviso" if is_shared else "Progetto"
        self.binding_kind = "Istanza" if self.is_instance else "Tipo"
        self.info = "{}  |  {}  |  {}  |  {} categorie".format(
            self.kind,
            self.binding_kind,
            param_type_name(definition),
            len(self.category_ids),
        )

    def applies_to_category(self, cat_id_value):
        if not self.category_ids:
            return True
        return cat_id_value in self.category_ids

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
    """Tutte le schedule utente del progetto, escluse template e revision schedule."""
    items = []
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)
    for vs in collector:
        try:
            if vs.IsTemplate:
                continue
            sdef = vs.Definition
            if sdef is None:
                continue
            if sdef.IsTitleblockRevisionSchedule:
                continue
        except Exception:
            continue
        try:
            items.append(ScheduleItem(vs))
        except Exception as ex:
            logger.debug("Schedule ignorata: %s", ex)
    items.sort(key=lambda i: i.name.lower())
    return items


def collect_parameters():
    """Parametri di progetto e condivisi presenti nelle ParameterBindings del documento."""
    items = []
    bindings = doc.ParameterBindings
    iterator = bindings.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        try:
            definition = iterator.Key
            binding = iterator.Current
            if definition is None or binding is None:
                continue
            is_shared = False
            pid = getattr(definition, "Id", None)
            if pid is not None:
                param_elem = doc.GetElement(pid)
                is_shared = isinstance(param_elem, DB.SharedParameterElement)
            items.append(ParamItem(definition, binding, is_shared))
        except Exception as ex:
            logger.debug("Parametro ignorato: %s", ex)
    items.sort(key=lambda i: i.name.lower())
    return items


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
            view = s_item.view
            sdef = view.Definition
            sched_cat = eid_value(sdef.CategoryId)
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
                            reason = ("Parametro non associato alla categoria "
                                      "'{}' della schedule.".format(s_item.category))
                        else:
                            reason = ("Parametro non presente tra i campi "
                                      "schedulabili di questa schedule "
                                      "(categoria '{}').".format(s_item.category))
                        results.append(ResultItem(
                            s_item.name, p_item.name, OUT_NOT_AVAILABLE, reason))
                        continue

                    # 3) aggiunta effettiva
                    try:
                        field = sdef.AddField(sfield)
                        if hide_fields and field is not None:
                            field.IsHidden = True
                        existing_ids.add(pid_value)
                        existing_names.add(p_item.name)
                        added_any = True
                        results.append(ResultItem(
                            s_item.name, p_item.name, OUT_ADDED,
                            "Campo aggiunto{}.".format(
                                " (nascosto)" if hide_fields else "")))
                    except Exception as ex:
                        logger.debug(traceback.format_exc())
                        results.append(ResultItem(
                            s_item.name, p_item.name, OUT_ERROR,
                            "Errore API Revit: {}".format(safe_str(ex))))

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
                    s_item.name, "-", OUT_ERROR,
                    "Schedule non modificata: {}".format(safe_str(ex))))

        tgroup.Assimilate()
    except Exception as ex:
        if tgroup.HasStarted() and not tgroup.HasEnded():
            tgroup.RollBack()
        logger.error("Operazione annullata: %s", safe_str(ex))
        results.append(ResultItem("-", "-", OUT_ERROR,
                                  "Operazione annullata: {}".format(safe_str(ex))))
    return results


# ---------------------------------------------------------------------------
# finestra di report
# ---------------------------------------------------------------------------
class ReportWindow(WPFWindow):
    def __init__(self, results):
        WPFWindow.__init__(self, "ReportWindow.xaml")
        self.results = results
        self._refresh()

    # -- helper -------------------------------------------------------------
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

    def _as_text(self, separator="\t"):
        lines = [separator.join(["Schedule", "Parametro", "Esito", "Dettaglio"])]
        for res in self.results:
            lines.append(separator.join([
                safe_str(res.schedule), safe_str(res.parameter),
                safe_str(res.outcome), safe_str(res.detail)]))
        return "\r\n".join(lines)

    # -- eventi -------------------------------------------------------------
    def filter_changed(self, sender, args):
        self._refresh()

    def copy_click(self, sender, args):
        try:
            from System.Windows import Clipboard
            Clipboard.SetText(self._as_text("\t"))
            forms.toast("Report copiato negli appunti.")
        except Exception as ex:
            forms.alert("Copia non riuscita: {}".format(safe_str(ex)))

    def export_click(self, sender, args):
        try:
            path = forms.save_file(file_ext="csv", default_name="report_parametri_schedule")
            if not path:
                return
            content = self._as_text(";")
            handle = open(path, "wb")
            try:
                handle.write(content.encode("utf-8-sig"))
            finally:
                handle.close()
            forms.alert("Report salvato in:\n{}".format(path))
        except Exception as ex:
            forms.alert("Esportazione non riuscita: {}".format(safe_str(ex)))

    def close_click(self, sender, args):
        self.Close()


# ---------------------------------------------------------------------------
# finestra principale
# ---------------------------------------------------------------------------
class AddParamsWindow(WPFWindow):
    def __init__(self, schedule_items, param_items):
        WPFWindow.__init__(self, "AddParamsWindow.xaml")
        self.all_schedules = schedule_items
        self.all_params = param_items
        self.confirmed = False
        self._bind(self.schedules_list, self.all_schedules)
        self._bind(self.params_list, self.all_params)
        self._update_counters()

    # -- helper -------------------------------------------------------------
    @staticmethod
    def _bind(listbox, items):
        data = List[object]()
        for item in items:
            data.Add(item)
        listbox.ItemsSource = data

    @staticmethod
    def _filter(items, text):
        text = (text or "").strip().lower()
        if not text:
            return items
        tokens = text.split()
        result = []
        for item in items:
            haystack = (item.name + " " + item.info).lower()
            if all(tok in haystack for tok in tokens):
                result.append(item)
        return result

    def _visible_schedules(self):
        return self._filter(self.all_schedules, self.schedules_filter.Text)

    def _visible_params(self):
        return self._filter(self.all_params, self.params_filter.Text)

    def selected_schedules(self):
        return [i for i in self.all_schedules if i.checked]

    def selected_params(self):
        return [i for i in self.all_params if i.checked]

    def _update_counters(self):
        self.schedules_count.Text = u"{} selezionate / {} totali".format(
            len(self.selected_schedules()), len(self.all_schedules))
        self.params_count.Text = u"{} selezionati / {} totali".format(
            len(self.selected_params()), len(self.all_params))

    # -- eventi -------------------------------------------------------------
    def item_checked(self, sender, args):
        try:
            sender.DataContext.checked = True
        except Exception:
            pass
        self._update_counters()

    def item_unchecked(self, sender, args):
        try:
            sender.DataContext.checked = False
        except Exception:
            pass
        self._update_counters()

    def schedules_filter_changed(self, sender, args):
        self._bind(self.schedules_list, self._visible_schedules())

    def params_filter_changed(self, sender, args):
        self._bind(self.params_list, self._visible_params())

    def schedules_check_all(self, sender, args):
        for item in self._visible_schedules():
            item.checked = True
        self._bind(self.schedules_list, self._visible_schedules())
        self._update_counters()

    def schedules_uncheck_all(self, sender, args):
        for item in self.all_schedules:
            item.checked = False
        self._bind(self.schedules_list, self._visible_schedules())
        self._update_counters()

    def params_check_all(self, sender, args):
        for item in self._visible_params():
            item.checked = True
        self._bind(self.params_list, self._visible_params())
        self._update_counters()

    def params_uncheck_all(self, sender, args):
        for item in self.all_params:
            item.checked = False
        self._bind(self.params_list, self._visible_params())
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

    schedules = collect_schedules()
    if not schedules:
        forms.alert("Nessuna schedule trovata nel progetto.", exitscript=True)

    params = collect_parameters()
    if not params:
        forms.alert("Nessun parametro di progetto o condiviso trovato nel modello.",
                    exitscript=True)

    window = AddParamsWindow(schedules, params)
    window.show_dialog()
    if not window.confirmed:
        script.exit()

    sel_schedules = window.selected_schedules()
    sel_params = window.selected_params()
    hide_fields = bool(window.hidden_fields.IsChecked)

    results = add_parameters(sel_schedules, sel_params, hide_fields=hide_fields)

    ReportWindow(results).show_dialog()


if __name__ == "__main__":
    main()
