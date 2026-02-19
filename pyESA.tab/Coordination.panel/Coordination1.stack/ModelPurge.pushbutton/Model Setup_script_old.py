# -*- coding: utf-8 -*-
# Intestazione dello script con metadati
__title__   = "Model\nSetup"
__doc__     = """Version = 1.1
Date    = 19.05.2025
________________________________________________________________
Elimina View Templates, Schedule, Legende e Filtri
Lo script permette di eliminare elementi selezionati dall'utente, 
saltando quelli in uso (applicati a viste o inseriti in tavole).
________________________________________________________________
Author(s):
Andrea Patti
"""

import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult

from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


def get_sheets_with_views():
    """Restituisce un set con le viste contenute nelle tavole."""
    views_on_sheets = set()
    
    # Raccoglie tutte le tavole
    all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    
    for sheet in all_sheets:
        # Raccoglie i viewports in ogni tavola (viste grafiche e legende)
        viewports = FilteredElementCollector(doc, sheet.Id).OfClass(Viewport).ToElements()
        for vp in viewports:
            views_on_sheets.add(vp.ViewId)
        
        # Raccoglie schedules sulla tavola
        schedules = FilteredElementCollector(doc, sheet.Id).OfClass(ScheduleSheetInstance).ToElements()
        for sched in schedules:
            views_on_sheets.add(sched.ScheduleId)
    
    return views_on_sheets


def get_views_with_templates():
    """Restituisce un dizionario dei view template e le viste che li utilizzano."""
    template_dict = {}
    
    # Raccoglie tutte le viste
    all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
    
    for view in all_views:
        # Salta viste che non possono avere template
        if not hasattr(view, "ViewTemplateId") or view.IsTemplate:
            continue
        
        template_id = view.ViewTemplateId
        
        # Se la vista ha un template assegnato
        if not template_id.IntegerValue == -1:
            if template_id not in template_dict:
                template_dict[template_id] = []
            
            template_dict[template_id].append(view.Id)
    
    return template_dict


def delete_view_templates():
    """Elimina i view template selezionati, saltando quelli in uso."""
    # Raccoglie tutti i view template
    all_templates = [v for v in FilteredElementCollector(doc).OfClass(View).ToElements() 
                    if v.IsTemplate]
    
    if not all_templates:
        forms.alert("Non ci sono view template nel progetto.", title="Nessun Template")
        return
    
    # Ottiene i template in uso
    templates_in_use = get_views_with_templates()
    
    # Crea lista di template con indicazione se sono in uso
    templates_data = []
    for template in all_templates:
        in_use = template.Id in templates_in_use
        templates_data.append({
            'name': template.Name,
            'id': template.Id,
            'in_use': in_use
        })
    
    # Ordinamento alfabetico
    templates_data.sort(key=lambda x: x['name'])
    
    # Crea lista di selezione con indicazione se in uso (testo rosso)
    template_options = {}
    for t in templates_data:
        if t['in_use']:
            display_name = '<span style="color:red">{} [IN USO]</span>'.format(t['name'])
        else:
            display_name = t['name']
        
        template_options[display_name] = t
    
    # Finestra di dialogo per la selezione
    selected_names = forms.SelectFromList.show(
        sorted(template_options.keys()),
        title="Seleziona View Template da eliminare",
        multiselect=True,
        button_name="Elimina View Template"
    )
    
    if not selected_names:
        return
    
    # Contatori per report
    deleted_count = 0
    skipped_count = 0
    skipped_names = []
    
    # Transazione per eliminazione
    with revit.Transaction("Elimina View Template"):
        for name in selected_names:
            template = template_options[name]
            
            # Salta template in uso
            if template['in_use']:
                skipped_count += 1
                skipped_names.append(template['name'])
                continue
            
            # Elimina template
            try:
                doc.Delete(template['id'])
                deleted_count += 1
            except Exception as e:
                print("Errore nell'eliminazione di {}: {}".format(template['name'], str(e)))
    
    # Messaggio di riepilogo
    result_message = "View Template eliminati: {}\n".format(deleted_count)
    if skipped_count > 0:
        result_message += "\nView Template saltati (in uso): {}\n".format(skipped_count)
        result_message += "- " + "\n- ".join(skipped_names)
    
    forms.alert(result_message, title="Eliminazione Completata")


def delete_schedules():
    """Elimina le schedule selezionate, saltando quelle inserite in tavole."""
    # Apri la finestra di output di PyRevit per informazioni diagnostiche
    output.close_others()
    output.print_md("# Diagnostica Schedule")
    
    # Raccoglie tutte le schedule
    all_schedules = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
    output.print_md("## Numero totale di Schedule trovate: {}".format(len(all_schedules)))
    
    # Testa la possibilità di eliminazione
    schedules_data = []
    
    # Ottiene le viste contenute nelle tavole
    views_on_sheets = get_sheets_with_views()
    
    # Raccogli tutte le schedule con un nome
    for s in all_schedules:
        try:
            name = s.Name
            # Se la schedule ha un nome, considerala potenzialmente eliminabile
            on_sheet = s.Id in views_on_sheets
            schedules_data.append({
                'name': name,
                'id': s.Id,
                'in_sheet': on_sheet
            })
        except Exception as e:
            output.print_md("⚠️ Schedule senza nome o accesso negato: {}".format(str(e)))
    
    output.print_md("## Schedule trovate: {}".format(len(schedules_data)))
    
    if not schedules_data:
        forms.alert("Non è stato possibile accedere a nessuna schedule nel progetto.", title="Errore di Accesso")
        return
    
    # Ordinamento alfabetico
    schedules_data.sort(key=lambda x: x['name'])
    
    # Crea lista di selezione con indicazione se in tavola (testo rosso)
    schedule_options = {}
    for s in schedules_data:
        if s['in_sheet']:
            display_name = '<span style="color:red">{} [IN TAVOLA]</span>'.format(s['name'])
        else:
            display_name = s['name']
        
        schedule_options[display_name] = s
    
    # Finestra di dialogo per la selezione
    selected_names = forms.SelectFromList.show(
        sorted(schedule_options.keys()),
        title="Seleziona Schedule da eliminare",
        multiselect=True,
        button_name="Elimina Schedule"
    )
    
    if not selected_names:
        return
    
    # Contatori per report
    deleted_count = 0
    skipped_count = 0
    skipped_names = []
    error_count = 0
    error_names = []
    
    # Transazione per eliminazione
    with revit.Transaction("Elimina Schedule"):
        for name in selected_names:
            schedule = schedule_options[name]
            
            # Salta schedule in tavola
            if schedule['in_sheet']:
                skipped_count += 1
                skipped_names.append(schedule['name'])
                continue
            
            # Prova ad eliminare la schedule
            try:
                # Tentativo diretto di eliminazione
                doc.Delete(schedule['id'])
                deleted_count += 1
                output.print_md("✅ Eliminata: {}".format(schedule['name']))
            except Exception as e:
                error_count += 1
                error_names.append(schedule['name'])
                output.print_md("❌ Errore nell'eliminazione di {}: {}".format(schedule['name'], str(e)))
    
    # Messaggio di riepilogo
    result_message = "Schedule eliminate: {}\n".format(deleted_count)
    if skipped_count > 0:
        result_message += "\nSchedule saltate (in tavola): {}\n".format(skipped_count)
        result_message += "- " + "\n- ".join(skipped_names)
    if error_count > 0:
        result_message += "\nSchedule non eliminate per errori: {}\n".format(error_count)
        result_message += "- " + "\n- ".join(error_names)
        result_message += "\nControlla la finestra di diagnostica per maggiori dettagli."
    
    forms.alert(result_message, title="Eliminazione Completata")


def delete_legends():
    """Elimina le legende selezionate, saltando quelle inserite in tavole."""
    # Raccoglie tutte le legende
    all_legends = FilteredElementCollector(doc).OfClass(View).ToElements()
    all_legends = [view for view in all_legends if view.ViewType == ViewType.Legend]
    
    if not all_legends:
        forms.alert("Non ci sono legende nel progetto.", title="Nessuna Legenda")
        return
    
    # Ottiene le viste contenute nelle tavole
    views_on_sheets = get_sheets_with_views()
    
    # Crea lista di legende con indicazione se sono in tavola
    legends_data = []
    for legend in all_legends:
        in_sheet = legend.Id in views_on_sheets
        legends_data.append({
            'name': legend.Name,
            'id': legend.Id,
            'in_sheet': in_sheet
        })
    
    # Ordinamento alfabetico
    legends_data.sort(key=lambda x: x['name'])
    
    # Crea lista di selezione con indicazione se in tavola (testo rosso)
    legend_options = {}
    for l in legends_data:
        if l['in_sheet']:
            display_name = '<span style="color:red">{} [IN TAVOLA]</span>'.format(l['name'])
        else:
            display_name = l['name']
        
        legend_options[display_name] = l
    
    # Finestra di dialogo per la selezione
    selected_names = forms.SelectFromList.show(
        sorted(legend_options.keys()),
        title="Seleziona Legende da eliminare",
        multiselect=True,
        button_name="Elimina Legende"
    )
    
    if not selected_names:
        return
    
    # Contatori per report
    deleted_count = 0
    skipped_count = 0
    skipped_names = []
    error_count = 0
    error_names = []
    
    # Transazione per eliminazione
    with revit.Transaction("Elimina Legende"):
        for name in selected_names:
            legend = legend_options[name]
            
            # Salta legende in tavola
            if legend['in_sheet']:
                skipped_count += 1
                skipped_names.append(legend['name'])
                continue
            
            # Elimina legenda
            try:
                doc.Delete(legend['id'])
                deleted_count += 1
            except Exception as e:
                error_count += 1
                error_names.append(legend['name'])
                output.print_md("❌ Errore nell'eliminazione di {}: {}".format(legend['name'], str(e)))
    
    # Messaggio di riepilogo
    result_message = "Legende eliminate: {}\n".format(deleted_count)
    if skipped_count > 0:
        result_message += "\nLegende saltate (in tavola): {}\n".format(skipped_count)
        result_message += "- " + "\n- ".join(skipped_names)
    if error_count > 0:
        result_message += "\nLegende non eliminate per errori: {}\n".format(error_count)
        result_message += "- " + "\n- ".join(error_names)
        result_message += "\nControlla la finestra di diagnostica per maggiori dettagli."
    
    forms.alert(result_message, title="Eliminazione Completata")


# ----- NUOVA FUNZIONALITÀ: ELIMINAZIONE FILTRI -----

def get_view_name_safely(view):
    """Ottiene il nome di una vista in modo sicuro."""
    try:
        return view.Name
    except Exception:
        return "Vista senza nome"


def can_view_have_filters(view):
    """Controlla se una vista può avere filtri applicati."""
    try:
        # Le viste che possono avere filtri in genere hanno il metodo GetFilters()
        filter_ids = view.GetFilters()
        return True
    except Exception:
        return False


def check_filter_usage(filter_id):
    """
    Verifica se il filtro è applicato in qualsiasi view template o vista.
    Restituisce una lista di tuple (elemento, è_template) per tutti gli usi del filtro.
    """
    usage_list = []
    
    # Ottieni tutte le viste in modo sicuro
    try:
        all_views = list(FilteredElementCollector(doc).OfClass(View).ToElements())
    except Exception as e:
        output.print_md("❌ Errore durante il recupero delle viste: {}".format(str(e)))
        return usage_list
    
    # Controlla prima i view template
    try:
        view_templates = [view for view in all_views if view.IsTemplate]
        
        for view_template in view_templates:
            if can_view_have_filters(view_template):
                try:
                    filter_ids = view_template.GetFilters()
                    if filter_id in filter_ids:
                        usage_list.append((view_template, True))  # True indica che è un template
                except Exception as e:
                    output.print_md("⚠️ Errore durante il controllo del template {}: {}".format(
                        get_view_name_safely(view_template), str(e)))
                    continue
    except Exception as e:
        output.print_md("❌ Errore durante il controllo dei view template: {}".format(str(e)))
    
    # Controlla le viste normali
    try:
        normal_views = [view for view in all_views if not view.IsTemplate]
        
        for view in normal_views:
            if can_view_have_filters(view):
                try:
                    filter_ids = view.GetFilters()
                    if filter_id in filter_ids:
                        usage_list.append((view, False))  # False indica che non è un template
                except Exception as e:
                    output.print_md("⚠️ Errore durante il controllo della vista {}: {}".format(
                        get_view_name_safely(view), str(e)))
                    continue
    except Exception as e:
        output.print_md("❌ Errore durante il controllo delle viste normali: {}".format(str(e)))
    
    return usage_list  # Restituisce la lista di tutte le viste/template che usano il filtro


def remove_filter_from_views(filter_id, views_and_templates):
    """
    Rimuove il filtro specificato da tutte le viste e template nella lista fornita.
    Restituisce il numero di viste/template modificati con successo.
    """
    success_count = 0
    
    for view, is_template in views_and_templates:
        try:
            # Ottiene i filtri e le override correnti
            filter_ids = view.GetFilters()
            
            if filter_id in filter_ids:
                # Rimuovi il filtro
                view.RemoveFilter(filter_id)
                success_count += 1
                output.print_md("✅ Filtro rimosso da {} '{}'".format(
                    "View Template" if is_template else "Vista", 
                    get_view_name_safely(view)))
        except Exception as e:
            output.print_md("❌ Errore durante la rimozione del filtro da {}: {}".format(
                get_view_name_safely(view), str(e)))
    
    return success_count


def delete_filters():
    """Elimina i filtri selezionati, con opzione di rimuoverli da viste/template se in uso."""
    # Apri la finestra di output di PyRevit per informazioni diagnostiche
    output.close_others()
    output.print_md("# Eliminazione Filtri")
    
    # Ottiene tutti i filtri
    try:
        all_filters = list(FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements())
    except Exception as e:
        output.print_md("❌ Errore durante il recupero dei filtri: {}".format(str(e)))
        forms.alert("Errore durante il recupero dei filtri.", title="Errore")
        return
    
    if not all_filters:
        forms.alert("Non ci sono filtri nel progetto.", title="Nessun Filtro")
        return
    
    # Prepara i dati dei filtri e verifica quali sono in uso
    filters_data = []
    for filter_element in all_filters:
        filter_name = filter_element.Name
        usage_list = check_filter_usage(filter_element.Id)
        in_use = len(usage_list) > 0
        
        filters_data.append({
            'name': filter_name,
            'element': filter_element,
            'id': filter_element.Id,
            'in_use': in_use,
            'usage': usage_list
        })
    
    # Ordina alfabeticamente
    filters_data.sort(key=lambda x: x['name'])
    
    # Crea lista di selezione con indicazione se in uso (testo rosso)
    filter_options = {}
    for f in filters_data:
        if f['in_use']:
            display_name = '<span style="color:red">{} [IN USO]</span>'.format(f['name'])
        else:
            display_name = f['name']
        
        filter_options[display_name] = f
    
    # Mostra la finestra di dialogo per la selezione dei filtri
    selected_names = forms.SelectFromList.show(
        sorted(filter_options.keys()),
        title="Seleziona i filtri da eliminare",
        multiselect=True,
        button_name="Elimina filtri selezionati"
    )
    
    if not selected_names:
        return
    
    # Raccogli filtri problematici (in uso)
    problematic_filters = {}
    for name in selected_names:
        filter_data = filter_options[name]
        if filter_data['in_use']:
            problematic_filters[filter_data['name']] = filter_data
    
    # Se ci sono filtri problematici, mostra un messaggio di avviso
    modified_views_count = 0
    if problematic_filters:
        error_message = "I seguenti filtri sono applicati a view template o viste:\n\n"
        for filter_name, data in problematic_filters.items():
            error_message += "Filtro: '{}'\n".format(filter_name)
            for view, is_template in data['usage']:
                view_name = get_view_name_safely(view)
                view_type = "View Template" if is_template else "Vista"
                error_message += "  - Applicato in {} '{}'\n".format(view_type, view_name)
            error_message += "\n"
        
        error_message += "Vuoi procedere con l'eliminazione dei filtri, rimuovendoli automaticamente da tutti i view template e viste in cui sono applicati?"
        
        # Chiede all'utente se vuole procedere comunque
        result = forms.alert(
            error_message,
            title="Filtri applicati",
            yes=True,
            no=True,
            ok=False
        )
        
        if not result:
            forms.alert("Operazione annullata dall'utente.", title="Operazione annullata")
            return
    
    # Se l'utente ha deciso di procedere o non ci sono filtri problematici
    with revit.Transaction("Elimina filtri"):
        delete_count = 0
        error_count = 0
        error_names = []
        
        # Prima rimuovi i filtri dalle viste dove sono applicati
        for filter_name, data in problematic_filters.items():
            filter_element = data['element']
            usage_list = data['usage']
            
            modified_views_count += remove_filter_from_views(filter_element.Id, usage_list)
        
        # Poi elimina i filtri selezionati
        for name in selected_names:
            filter_data = filter_options[name]
            filter_element = filter_data['element']
            filter_name = filter_data['name']
            
            try:
                doc.Delete(filter_element.Id)
                delete_count += 1
                output.print_md("✅ Filtro eliminato: {}".format(filter_name))
            except Exception as e:
                error_count += 1
                error_names.append(filter_name)
                output.print_md("❌ Errore durante l'eliminazione del filtro {}: {}".format(
                    filter_name, str(e)))
    
    # Messaggio di riepilogo
    result_message = "Filtri eliminati: {}\n".format(delete_count)
    
    if modified_views_count > 0:
        result_message += "\nViste/View Template modificati: {}\n".format(modified_views_count)
    
    if error_count > 0:
        result_message += "\nFiltri non eliminati per errori: {}\n".format(error_count)
        result_message += "- " + "\n- ".join(error_names)
        result_message += "\nControlla la finestra di diagnostica per maggiori dettagli."
    
    forms.alert(result_message, title="Eliminazione Completata")


# Menu principale
options = {
    "Elimina View Template": delete_view_templates,
    "Elimina Schedule": delete_schedules,
    "Elimina Legende": delete_legends,
    "Elimina Filtri": delete_filters  # Aggiunta la nuova opzione
}

selected_option = forms.CommandSwitchWindow.show(
    options.keys(),
    message="Seleziona l'operazione da eseguire:"
)

if selected_option:
    options[selected_option]()