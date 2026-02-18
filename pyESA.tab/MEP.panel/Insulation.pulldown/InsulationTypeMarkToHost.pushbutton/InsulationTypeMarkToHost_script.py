# -*- coding: utf-8 -*-
__title__ = "Insulation TypeMark\nTo Host"
__doc__ = """Version = 2.0
Date = 14.01.2026
________________________________________________________________
Copia il Type Mark dell'isolamento in un parametro di istanza
dell'elemento (tubazione o canalizzazione).

- L'utente sceglie se operare su tubazioni, canali o entrambi
- L'utente sceglie se operare su elementi selezionati o tutti
- Esclude automaticamente elementi senza isolamento
- Copia il Type Mark dall'isolamento all'elemento
________________________________________________________________
Author: Andrea Patti
"""

from pyrevit import revit, DB, forms, script
from rpw.ui.forms import FlexForm, Label, ComboBox, CheckBox, Separator, Button

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

def get_all_pipes():
    """Ottiene tutte le tubazioni del modello"""
    return list(DB.FilteredElementCollector(doc)
                .OfCategory(DB.BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .ToElements())

def get_all_ducts():
    """Ottiene tutte le canalizzazioni del modello"""
    return list(DB.FilteredElementCollector(doc)
                .OfCategory(DB.BuiltInCategory.OST_DuctCurves)
                .WhereElementIsNotElementType()
                .ToElements())

def get_selected_pipes():
    """Ottiene le tubazioni selezionate"""
    selection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
    pipes = []
    for elem in selection:
        if elem and elem.Category and elem.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_PipeCurves):
            pipes.append(elem)
    return pipes

def get_selected_ducts():
    """Ottiene le canalizzazioni selezionate"""
    selection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
    ducts = []
    for elem in selection:
        if elem and elem.Category and elem.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_DuctCurves):
            ducts.append(elem)
    return ducts

def get_instance_text_parameters(elements):
    """Ottiene tutti i parametri di istanza di tipo testo scrivibili comuni"""
    if not elements:
        return []
    
    param_names = set()
    for element in elements[:min(10, len(elements))]:
        for param in element.Parameters:
            if (not param.IsReadOnly and 
                param.StorageType == DB.StorageType.String):
                param_names.add(param.Definition.Name)
    return sorted(list(param_names))

def main():
    # Ottieni gli elementi selezionati (se presenti)
    selected_pipes = get_selected_pipes()
    selected_ducts = get_selected_ducts()
    has_selection = bool(selected_pipes or selected_ducts)
    
    # Opzioni per la modalità di selezione
    if has_selection:
        selection_options = ["Solo elementi selezionati", "Tutti gli elementi del progetto"]
    else:
        selection_options = ["Tutti gli elementi del progetto"]
    
    # Form per scegliere la modalità e le categorie
    components = [
        Label('Seleziona le categorie su cui operare:'),
        CheckBox('cat_pipe', 'Tubazioni', True),
        CheckBox('cat_duct', 'Canalizzazioni', True),
        Separator(),
        Label('Modalità di selezione:'),
        ComboBox('selection_mode', selection_options, selection_options[0]),
        Separator(),
        Button('OK')
    ]
    
    form = FlexForm('Copia Type Mark Isolamento', components)
    form.show()
    
    if not form.values:
        script.exit()
    
    # Categorie selezionate
    process_pipes = form.values.get('cat_pipe', False)
    process_ducts = form.values.get('cat_duct', False)
    
    if not process_pipes and not process_ducts:
        forms.alert('Nessuna categoria selezionata!', exitscript=True)
    
    # Determina la modalità di selezione
    use_selection = form.values['selection_mode'] == "Solo elementi selezionati"
    
    # Ottieni gli elementi da processare
    pipes_to_process = []
    ducts_to_process = []
    
    if process_pipes:
        if use_selection:
            pipes_to_process = selected_pipes
            if not pipes_to_process:
                forms.alert('Nessuna tubazione selezionata!', exitscript=True)
        else:
            pipes_to_process = get_all_pipes()
    
    if process_ducts:
        if use_selection:
            ducts_to_process = selected_ducts
            if not ducts_to_process:
                forms.alert('Nessuna canalizzazione selezionata!', exitscript=True)
        else:
            ducts_to_process = get_all_ducts()
    
    if not pipes_to_process and not ducts_to_process:
        forms.alert('Nessun elemento trovato!', exitscript=True)
    
    output.print_md('## Analisi elementi in corso...')
    output.print_md('---')
    if pipes_to_process:
        output.print_md('Tubazioni totali da analizzare: **{}**'.format(len(pipes_to_process)))
    if ducts_to_process:
        output.print_md('Canalizzazioni totali da analizzare: **{}**'.format(len(ducts_to_process)))
    
    # Ottieni tutti gli isolamenti
    pipe_insulations = []
    duct_insulations = []
    
    if process_pipes:
        pipe_insulations = list(DB.FilteredElementCollector(doc)
                               .OfCategory(DB.BuiltInCategory.OST_PipeInsulations)
                               .WhereElementIsNotElementType()
                               .ToElements())
        output.print_md('Isolamenti tubazioni nel modello: **{}**'.format(len(pipe_insulations)))
    
    if process_ducts:
        duct_insulations = list(DB.FilteredElementCollector(doc)
                               .OfCategory(DB.BuiltInCategory.OST_DuctInsulations)
                               .WhereElementIsNotElementType()
                               .ToElements())
        output.print_md('Isolamenti canalizzazioni nel modello: **{}**'.format(len(duct_insulations)))
    
    # Crea dizionari: element_id -> insulation
    pipe_insulation_dict = {}
    for insulation in pipe_insulations:
        host_id = insulation.HostElementId
        if host_id and host_id != DB.ElementId.InvalidElementId:
            pipe_insulation_dict[host_id] = insulation
    
    duct_insulation_dict = {}
    for insulation in duct_insulations:
        host_id = insulation.HostElementId
        if host_id and host_id != DB.ElementId.InvalidElementId:
            duct_insulation_dict[host_id] = insulation
    
    # Filtra gli elementi che hanno isolamento
    pipes_with_insulation = []
    pipes_without_insulation = []
    ducts_with_insulation = []
    ducts_without_insulation = []
    
    for pipe in pipes_to_process:
        if pipe.Id in pipe_insulation_dict:
            pipes_with_insulation.append(pipe)
        else:
            pipes_without_insulation.append(pipe)
    
    for duct in ducts_to_process:
        if duct.Id in duct_insulation_dict:
            ducts_with_insulation.append(duct)
        else:
            ducts_without_insulation.append(duct)
    
    output.print_md('---')
    if pipes_to_process:
        output.print_md('**Tubazioni:**')
        output.print_md('- Con isolamento: **{}**'.format(len(pipes_with_insulation)))
        output.print_md('- Senza isolamento: **{}**'.format(len(pipes_without_insulation)))
    
    if ducts_to_process:
        output.print_md('**Canalizzazioni:**')
        output.print_md('- Con isolamento: **{}**'.format(len(ducts_with_insulation)))
        output.print_md('- Senza isolamento: **{}**'.format(len(ducts_without_insulation)))
    
    output.print_md('---')
    
    # Verifica che ci siano elementi con isolamento
    elements_with_insulation = pipes_with_insulation + ducts_with_insulation
    
    if not elements_with_insulation:
        message = 'Nessun elemento con isolamento trovato!\n\n'
        if pipes_to_process:
            message += 'Tubazioni senza isolamento: {}\n'.format(len(pipes_without_insulation))
        if ducts_to_process:
            message += 'Canalizzazioni senza isolamento: {}'.format(len(ducts_without_insulation))
        forms.alert(message, title='Nessun isolamento trovato', exitscript=True)
    
    # Ottieni i parametri disponibili
    available_params = get_instance_text_parameters(elements_with_insulation)
    
    if not available_params:
        forms.alert('Nessun parametro di istanza di testo disponibile negli elementi!\n\n'
                   'Crea un parametro di progetto o condiviso di tipo TESTO.',
                   title='Parametri non trovati', exitscript=True)
    
    # Form per selezionare il parametro di destinazione
    param_components = [
        Label('Seleziona il parametro di destinazione:'),
        ComboBox('target_param', available_params, available_params[0]),
        Separator(),
        Button('OK')
    ]
    
    param_form = FlexForm('Parametro di destinazione', param_components)
    param_form.show()
    
    if not param_form.values:
        script.exit()
    
    selected_param = param_form.values['target_param']
    
    # Inizia la transazione
    t = DB.Transaction(doc, 'Copia Type Mark Isolamento')
    t.Start()
    
    success_count = 0
    error_count = 0
    no_typemark_count = 0
    
    output.print_md('## Risultati elaborazione')
    output.print_md('Parametro di destinazione: **{}**'.format(selected_param))
    output.print_md('---')
    
    try:
        # Processa le tubazioni
        if pipes_with_insulation:
            output.print_md('### Tubazioni')
            for pipe in pipes_with_insulation:
                try:
                    # Ottieni l'isolamento della tubazione
                    insulation = pipe_insulation_dict[pipe.Id]
                    
                    # Ottieni il tipo di isolamento
                    insulation_type = doc.GetElement(insulation.GetTypeId())
                    
                    if not insulation_type:
                        error_count += 1
                        continue
                    
                    # Ottieni il Type Mark
                    type_mark_param = insulation_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK)
                    
                    if type_mark_param and type_mark_param.HasValue:
                        type_mark = type_mark_param.AsString()
                        
                        if type_mark and type_mark.strip():
                            # Trova il parametro nella tubazione
                            pipe_param = pipe.LookupParameter(selected_param)
                            
                            if pipe_param and not pipe_param.IsReadOnly:
                                pipe_param.Set(type_mark)
                                success_count += 1
                                
                                type_name = insulation_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                                output.print_md('✓ Pipe ID {}: Type Mark = **{}** (Isolamento: {})'.format(
                                    output.linkify(pipe.Id), type_mark, type_name))
                            else:
                                error_count += 1
                                output.print_md('✗ Pipe ID {}: Parametro non trovato o in sola lettura'.format(
                                    output.linkify(pipe.Id)))
                        else:
                            no_typemark_count += 1
                            type_name = insulation_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                            output.print_md('⚠ Pipe ID {}: Type Mark vuoto per isolamento **{}**'.format(
                                output.linkify(pipe.Id), type_name))
                    else:
                        no_typemark_count += 1
                        type_name = insulation_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                        output.print_md('⚠ Pipe ID {}: Type Mark non disponibile per isolamento **{}**'.format(
                            output.linkify(pipe.Id), type_name))
                        
                except Exception as e:
                    error_count += 1
                    output.print_md('✗ Pipe ID {}: Errore - {}'.format(
                        output.linkify(pipe.Id), str(e)))
        
        # Processa le canalizzazioni
        if ducts_with_insulation:
            output.print_md('---')
            output.print_md('### Canalizzazioni')
            for duct in ducts_with_insulation:
                try:
                    # Ottieni l'isolamento della canalizzazione
                    insulation = duct_insulation_dict[duct.Id]
                    
                    # Ottieni il tipo di isolamento
                    insulation_type = doc.GetElement(insulation.GetTypeId())
                    
                    if not insulation_type:
                        error_count += 1
                        continue
                    
                    # Ottieni il Type Mark
                    type_mark_param = insulation_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK)
                    
                    if type_mark_param and type_mark_param.HasValue:
                        type_mark = type_mark_param.AsString()
                        
                        if type_mark and type_mark.strip():
                            # Trova il parametro nella canalizzazione
                            duct_param = duct.LookupParameter(selected_param)
                            
                            if duct_param and not duct_param.IsReadOnly:
                                duct_param.Set(type_mark)
                                success_count += 1
                                
                                type_name = insulation_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                                output.print_md('✓ Duct ID {}: Type Mark = **{}** (Isolamento: {})'.format(
                                    output.linkify(duct.Id), type_mark, type_name))
                            else:
                                error_count += 1
                                output.print_md('✗ Duct ID {}: Parametro non trovato o in sola lettura'.format(
                                    output.linkify(duct.Id)))
                        else:
                            no_typemark_count += 1
                            type_name = insulation_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                            output.print_md('⚠ Duct ID {}: Type Mark vuoto per isolamento **{}**'.format(
                                output.linkify(duct.Id), type_name))
                    else:
                        no_typemark_count += 1
                        type_name = insulation_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                        output.print_md('⚠ Duct ID {}: Type Mark non disponibile per isolamento **{}**'.format(
                            output.linkify(duct.Id), type_name))
                        
                except Exception as e:
                    error_count += 1
                    output.print_md('✗ Duct ID {}: Errore - {}'.format(
                        output.linkify(duct.Id), str(e)))
        
        t.Commit()
        
    except Exception as e:
        t.RollBack()
        forms.alert('Errore durante l\'elaborazione: {}'.format(str(e)), exitscript=True)
    
    # Stampa riepilogo
    output.print_md('---')
    output.print_md('## Riepilogo')
    output.print_md('- Modalità: **{}**'.format(form.values['selection_mode']))
    output.print_md('- Categorie: **{}**'.format(
        ', '.join([c for c in ['Tubazioni' if process_pipes else None, 
                                'Canalizzazioni' if process_ducts else None] if c])))
    if use_selection:
        if selected_pipes:
            output.print_md('- Tubazioni selezionate: **{}**'.format(len(selected_pipes)))
        if selected_ducts:
            output.print_md('- Canalizzazioni selezionate: **{}**'.format(len(selected_ducts)))
    output.print_md('- Elementi con isolamento trovati: **{}**'.format(len(elements_with_insulation)))
    output.print_md('- ✓ **Successo**: {}'.format(success_count))
    output.print_md('- ⚠ **Type Mark vuoto**: {}'.format(no_typemark_count))
    output.print_md('- ✗ **Errori**: {}'.format(error_count))
    output.print_md('---')

if __name__ == '__main__':
    main()