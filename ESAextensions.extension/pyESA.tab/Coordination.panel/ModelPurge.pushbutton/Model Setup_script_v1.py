# -*- coding: utf-8 -*-
"""
Elimina View Templates, Schedule e Legende
Lo script permette di eliminare elementi selezionati dall'utente, 
saltando quelli in uso (applicati a viste o inseriti in tavole).
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
    
    # Crea lista di selezione con indicazione se in uso
    template_options = {t['name'] + (" [IN USO]" if t['in_use'] else ""): t for t in templates_data}
    
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
    deletable_schedules = []
    non_deletable_schedules = []
    
    # Ottiene le viste contenute nelle tavole
    views_on_sheets = get_sheets_with_views()
    
    # Primo tentativo: raccogli tutte le schedule con un nome
    for s in all_schedules:
        try:
            name = s.Name
            # Se la schedule ha un nome, considerala potenzialmente eliminabile
            on_sheet = s.Id in views_on_sheets
            deletable_schedules.append({
                'obj': s,
                'name': name,
                'id': s.Id,
                'in_sheet': on_sheet,
                'view_type': s.ViewType
            })
        except Exception as e:
            output.print_md("⚠️ Schedule senza nome o accesso negato: {}".format(str(e)))
            non_deletable_schedules.append(s)
    
    output.print_md("## Schedule con nome: {}".format(len(deletable_schedules)))
    output.print_md("## Schedule senza accesso al nome: {}".format(len(non_deletable_schedules)))
    
    if not deletable_schedules:
        forms.alert("Non è stato possibile accedere a nessuna schedule nel progetto.", title="Errore di Accesso")
        return
    
    # Ordinamento alfabetico
    deletable_schedules.sort(key=lambda x: x['name'])
    
    # Crea lista di selezione con indicazione se in tavola
    schedule_options = {s['name'] + (" [IN TAVOLA]" if s['in_sheet'] else ""): s for s in deletable_schedules}
    
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
                
                # Stampa dettagli aggiuntivi per il debug
                output.print_md("Dettagli schedule:")
                output.print_md("- ID: {}".format(schedule['id'].IntegerValue))
                output.print_md("- Tipo di vista: {}".format(schedule['view_type']))
                try:
                    editable = schedule['obj'].IsEditable()
                    output.print_md("- Modificabile: {}".format(editable))
                except Exception as e2:
                    output.print_md("- Errore nel controllo di modificabilità: {}".format(str(e2)))
    
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


def delete_schedules_forced():
    """Tenta di eliminare le schedule, ignorando i controlli di sistema."""
    # Apri la finestra di output di PyRevit per informazioni diagnostiche
    output.close_others()
    output.print_md("# Eliminazione Schedule (Modalità Forzata)")
    
    # Raccoglie tutte le schedule
    all_schedules = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
    output.print_md("## Numero totale di Schedule trovate: {}".format(len(all_schedules)))
    
    # Archivia le schedule con i loro nomi
    schedules_data = []
    
    # Ottiene le viste contenute nelle tavole
    views_on_sheets = get_sheets_with_views()
    
    for s in all_schedules:
        try:
            name = s.Name
            in_sheet = s.Id in views_on_sheets
            schedules_data.append({
                'name': name,
                'id': s.Id,
                'in_sheet': in_sheet
            })
        except:
            # Salta le schedule senza nome accessibile
            pass
    
    # Ordinamento alfabetico
    schedules_data.sort(key=lambda x: x['name'])
    
    # Mostra tutte le schedule trovate
    output.print_md("## Schedule trovate: {}".format(len(schedules_data)))
    for s in schedules_data:
        output.print_md("- {} {}".format(s['name'], "(in tavola)" if s['in_sheet'] else ""))
    
    if not schedules_data:
        forms.alert("Non ci sono schedule con nomi accessibili nel progetto.", title="Nessuna Schedule")
        return
    
    # Crea lista di selezione con indicazione se in tavola
    schedule_options = {s['name'] + (" [IN TAVOLA]" if s['in_sheet'] else ""): s for s in schedules_data}
    
    # Finestra di dialogo per la selezione
    selected_names = forms.SelectFromList.show(
        sorted(schedule_options.keys()),
        title="Seleziona Schedule da eliminare (FORZATO)",
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
    
    # Chiedi conferma per l'eliminazione forzata
    confirm = forms.alert(
        "Stai per eliminare {} schedule selezionate. Le schedule in tavola saranno saltate. Continuare?".format(len(selected_names)),
        title="Conferma Eliminazione",
        yes=True, no=True
    )
    
    if not confirm:
        return
    
    # Transazione per eliminazione
    with revit.Transaction("Elimina Schedule (Forzato)"):
        for name in selected_names:
            schedule = schedule_options[name]
            
            # Salta schedule in tavola
            if schedule['in_sheet']:
                skipped_count += 1
                skipped_names.append(schedule['name'])
                continue
            
            # Tenta l'eliminazione
            try:
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
    
    # Crea lista di selezione con indicazione se in tavola
    legend_options = {l['name'] + (" [IN TAVOLA]" if l['in_sheet'] else ""): l for l in legends_data}
    
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
                print("Errore nell'eliminazione di {}: {}".format(legend['name'], str(e)))
    
    # Messaggio di riepilogo
    result_message = "Legende eliminate: {}\n".format(deleted_count)
    if skipped_count > 0:
        result_message += "\nLegende saltate (in tavola): {}\n".format(skipped_count)
        result_message += "- " + "\n- ".join(skipped_names)
    
    forms.alert(result_message, title="Eliminazione Completata")


# Menu principale
options = {
    "Elimina View Template": delete_view_templates,
    "Elimina Schedule (Normale)": delete_schedules,
    "Elimina Schedule (Forzato)": delete_schedules_forced,
    "Elimina Legende": delete_legends
}

selected_option = forms.CommandSwitchWindow.show(
    options.keys(),
    message="Seleziona l'operazione da eseguire:"
)

if selected_option:
    options[selected_option]()