# -*- coding: utf-8 -*-
__title__   = "Element To\nInsulation"
__doc__     = """Version = 1.0
Date    = 06.05.2025
________________________________________________________________ 
Recupera i tubi, raccordi tubazioni (rigide e flessibili), canalizzazioni (rigide e flessibili) e raccordi canalizzazioni 
che hanno l'isolante e copia e incolla il valore dei parametri definiti dall'utente dai relativi elementi host agli isolamenti.
 
ATT1: Inserire i vari parametri separati da un ; senza spazi intermedi a dividere
ATT2: È possibile scegliere su quali categorie di elementi applicare le modifiche
ATT3: È possibile scegliere se operare solo sugli elementi selezionati o su tutti gli elementi del progetto
ATT4: Supporta parametri di tipo testo, numeri e booleani
________________________________________________________________
Author(s): 
Andrea Patti, Tommaso Lorenzi 
"""
 
#REFERENCES
import pyrevit
from pyrevit import revit, script, DB
from pyrevit import forms
from rpw.ui.forms import TaskDialog, FlexForm, Label, TextBox, CheckBox, Separator, Button, ComboBox
from Autodesk.Revit.DB import Transaction, ElementId, StorageType, FilteredElementCollector, BuiltInCategory
import System
from System.Collections.Generic import List
 
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

# Ottieni gli elementi selezionati
selection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
    
# Form per l'inserimento dei parametri e selezione delle categorie
categories = [
    {"name": "Tubazioni", "checked": True, "category": BuiltInCategory.OST_PipeInsulations},
    {"name": "Raccordi Tubazioni", "checked": True, "category": BuiltInCategory.OST_PipeFittingInsulation},
    {"name": "Canali", "checked": True, "category": BuiltInCategory.OST_DuctInsulations},
    {"name": "Raccordi Canali", "checked": True, "category": BuiltInCategory.OST_DuctFittingInsulation}
]

# Opzioni per la modalità di selezione
selection_options = ["Solo elementi selezionati", "Tutti gli elementi del progetto"]

components = [
    Label('Parametri da copiare (separati da ;)'),
    TextBox('source_p'),
    Separator(),
    Label('Modalità di selezione:'),
    ComboBox('selection_mode', selection_options, selection_options[0] if selection else selection_options[1]),
    Separator(),
    Label('Seleziona le categorie su cui operare:'),
    CheckBox('cat_pipe', 'Isolamenti Tubazioni', categories[0]["checked"]),
#    CheckBox('cat_pipe_fitting', 'Raccordi Tubazioni', categories[1]["checked"]),
    CheckBox('cat_duct', 'Isolamenti Canali', categories[2]["checked"]),
#    CheckBox('cat_duct_fitting', 'Raccordi Canali', categories[3]["checked"]),
    Separator(),
    Button('OK')
]

form = FlexForm('Parameter copy from element to insulation', components)
form.show()

if form.values:
    # Inizia la transazione per modificare il documento
    t = Transaction(doc, "Set Insulation Parameter")
    t.Start()

    parameters = form.values['source_p']
    parameters_split = parameters.split(';')
    
    # Categorie selezionate
    selected_categories = []
    
    if form.values['cat_pipe']:
        selected_categories.append(BuiltInCategory.OST_PipeInsulations)
#    if form.values['cat_pipe_fitting']:
#        selected_categories.append(BuiltInCategory.OST_PipeFittingInsulation)
    if form.values['cat_duct']:
        selected_categories.append(BuiltInCategory.OST_DuctInsulations)
#    if form.values['cat_duct_fitting']:
#        selected_categories.append(BuiltInCategory.OST_DuctFittingInsulation)
    
    if not selected_categories:
        forms.alert('Nessuna categoria selezionata.', exitscript=True)
        t.RollBack()
        script.exit()
    
    # Determina la modalità di selezione
    use_selection = form.values['selection_mode'] == "Solo elementi selezionati"
    
    # Dizionario per tenere traccia delle categorie
    category_counts = {
        BuiltInCategory.OST_PipeInsulations: 0,
    #    BuiltInCategory.OST_PipeFittingInsulation: 0,
        BuiltInCategory.OST_DuctInsulations: 0,
    #    BuiltInCategory.OST_DuctFittingInsulation: 0
    }
    
    # Lista per tutti gli isolamenti da processare
    matching_insulations = []
    
    if use_selection:
        # Modalità elementi selezionati
        if not selection:
            forms.alert('Nessun elemento selezionato. Seleziona almeno un elemento o cambia modalità.', exitscript=True)
            t.RollBack()
            script.exit()
        
        # Ottieni gli ID degli elementi host selezionati
        selected_host_ids = [elem.Id for elem in selection]
        
        # Per ogni categoria selezionata, trova gli isolamenti corrispondenti
        for category in selected_categories:
            insulations = FilteredElementCollector(doc)\
                        .OfCategory(category)\
                        .WhereElementIsNotElementType()\
                        .ToElements()
            
            # Filtra gli isolamenti che hanno un host tra gli elementi selezionati
            for insulation in insulations:
                host_id = insulation.HostElementId
                if host_id in selected_host_ids:
                    matching_insulations.append(insulation)
                    category_counts[category] += 1
    else:
        # Modalità tutti gli elementi
        # Per ogni categoria selezionata, trova tutti gli isolamenti
        for category in selected_categories:
            insulations = FilteredElementCollector(doc)\
                        .OfCategory(category)\
                        .WhereElementIsNotElementType()\
                        .ToElements()
            
            for insulation in insulations:
                if insulation.HostElementId != ElementId.InvalidElementId:
                    matching_insulations.append(insulation)
                    category_counts[category] += 1
    
    # Stampa il conteggio per categoria
    output.print_md("## Isolamenti trovati per elementi selezionati:")
    if BuiltInCategory.OST_PipeInsulations in selected_categories:
        output.print_md("- Isolamenti Tubazioni: {}".format(category_counts[BuiltInCategory.OST_PipeInsulations]))
#    if BuiltInCategory.OST_PipeFittingInsulation in selected_categories:
#        output.print_md("- Raccordi Tubazioni: {}".format(category_counts[BuiltInCategory.OST_PipeFittingInsulation]))
    if BuiltInCategory.OST_DuctInsulations in selected_categories:
        output.print_md("- Isolamenti Canali: {}".format(category_counts[BuiltInCategory.OST_DuctInsulations]))
#    if BuiltInCategory.OST_DuctFittingInsulation in selected_categories:
#        output.print_md("- Raccordi Canali: {}".format(category_counts[BuiltInCategory.OST_DuctFittingInsulation]))
    
    # Contatori per il report finale
    updated_count = 0
    errors_count = 0

    # Ciclo for per ogni isolante che corrisponde a un host selezionato
    for insulation in matching_insulations:
        host_id = insulation.HostElementId
        host = doc.GetElement(host_id)
        
        # Verifica che l'host sia valido
        if not host:
            continue

        # Flag per verificare se almeno un parametro è stato copiato
        param_copied = False

        # Ciclo for per estrarre i valori dagli elementi host e copiarli sugli isolanti
        for param_name in parameters_split:
            # Recupera il parametro dall'elemento Host
            host_param = host.LookupParameter(param_name)
            # Verifica che il parametro esista nell'host
            if not host_param:
                continue
            # Recupera il parametro nell'insulation
            insulation_param = insulation.LookupParameter(param_name)
            # Verifica che il parametro esista nell'insulation
            if not insulation_param:
                continue
            # Verifica che il parametro sia scrivibile
            if insulation_param.IsReadOnly:
                continue
            try:
                # Copia il valore in base al tipo di parametro
                storage_type = host_param.StorageType
                if storage_type == StorageType.String:
                    host_value = host_param.AsString()
                    if host_value:  # Verifica che il valore non sia vuoto
                        insulation_param.Set(host_value)
                        param_copied = True
                elif storage_type == StorageType.Integer:
                    host_value = host_param.AsInteger()
                    insulation_param.Set(host_value)
                    param_copied = True
                elif storage_type == StorageType.Double:
                    host_value = host_param.AsDouble()
                    insulation_param.Set(host_value)
                    param_copied = True
                elif storage_type == StorageType.ElementId:
                    host_value = host_param.AsElementId()
                    insulation_param.Set(host_value)
                    param_copied = True
            except Exception as e:
                errors_count += 1
                output.print_md("Errore durante la copia del parametro '{}' per l'elemento ID: {}: {}".format(
                    param_name, insulation.Id.IntegerValue, str(e)))
        if param_copied:
            updated_count += 1

    # Chiude la transazione
    t.Commit()

    output.print_md("---------------------------------------------------------------------------")
    output.print_md("**Riepilogo operazioni:**")
    output.print_md("- Modalità: {}".format(form.values['selection_mode']))
    if use_selection:
        output.print_md("- Elementi selezionati: {}".format(len(selection)))
    output.print_md("- Elementi isolanti trovati: {}".format(len(matching_insulations)))
    output.print_md("- Elementi aggiornati con successo: {}".format(updated_count))
    if errors_count > 0:
        output.print_md("- Errori riscontrati: {}".format(errors_count))
    output.print_md("---------------------------------------------------------------------------")

else:
    script.exit()