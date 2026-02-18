# -*- coding: utf-8 -*-
# Intestazione dello script con metadati
__title__   = "Set Insulation\nWorkset"
__doc__     = """Version = 1.0
Date    = 09.05.2025
________________________________________________________________
Corregge i workset degli isolamenti.
Lo script verifica che il workset di ogni isolamento coincida con
quello dell'elemento a cui è associato (canale, raccordo, tubazione).
Se non coincide, imposta il workset dell'isolamento uguale a quello dell'elemento correlato.
________________________________________________________________
Author(s):
Andrea Patti
"""

import clr
from System.Collections.Generic import List, Dictionary
from System import DateTime

# Import Revit API
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

# Import pyRevit API
from pyrevit import revit, DB, forms, script

# Logger per report
output = script.get_output()
logger = script.get_logger()

# Ottieni documento attivo
doc = revit.doc

# Funzione per ottenere l'elemento di riferimento da un isolamento
def get_reference_element(insulation):
    try:
        # Ottieni l'id dell'elemento di riferimento
        ref_id = insulation.HostElementId
        if ref_id and not ref_id.IntegerValue == -1:
            return doc.GetElement(ref_id)
        return None
    except Exception as ex:
        logger.error("Errore nell'ottenere l'elemento di riferimento: {}".format(ex))
        return None

# Funzione per ottenere il nome del workset
def get_workset_name(workset_id):
    try:
        workset_table = doc.GetWorksetTable()
        workset = workset_table.GetWorkset(workset_id)
        return workset.Name
    except Exception as ex:
        logger.error("Errore nell'ottenere il nome del workset: {}".format(ex))
        return "Errore"

# Funzione per ottenere il nome della categoria
def get_category_name(element):
    try:
        if hasattr(element, "Category") and element.Category:
            return element.Category.Name
        return "Categoria sconosciuta"
    except:
        return "Categoria sconosciuta"

# Funzione principale
def correct_insulation_worksets():
    start_time = DateTime.Now
    
    # Colleziona tutti gli isolamenti (Pipe Insulations e Duct Insulations)
    pipe_insulations = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeInsulations).WhereElementIsNotElementType().ToElements()
    duct_insulations = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctInsulations).WhereElementIsNotElementType().ToElements()
    
    all_insulations = list(pipe_insulations) + list(duct_insulations)
    total_insulations = len(all_insulations)
    
    # Raggruppa elementi per workset di destinazione
    # Struttura: { workset_id_dest: { workset_id_origine: [elementi] } }
    workset_groups = {}
    
    # Elementi da processare e contatori
    elements_to_process = []
    skipped = 0
    
    # Analizza ogni isolamento e raggruppa per workset
    for insulation in all_insulations:
        ref_element = get_reference_element(insulation)
        if not ref_element:
            skipped += 1
            continue
        
        insulation_workset_id = insulation.WorksetId.IntegerValue
        ref_element_workset_id = ref_element.WorksetId.IntegerValue
        
        if insulation_workset_id != ref_element_workset_id:
            # Aggiungi a elements_to_process per report
            elements_to_process.append({
                'insulation': insulation,
                'ref_element': ref_element,
                'original_workset_id': ref_element.WorksetId,
                'insulation_workset_id': insulation.WorksetId
            })
            
            # Aggiungi al gruppo di workset usando dizionari Python standard
            if ref_element_workset_id not in workset_groups:
                workset_groups[ref_element_workset_id] = {}
                
            if insulation_workset_id not in workset_groups[ref_element_workset_id]:
                workset_groups[ref_element_workset_id][insulation_workset_id] = []
                
            workset_groups[ref_element_workset_id][insulation_workset_id].append({
                'insulation': insulation,
                'ref_element': ref_element
            })
    
    if not elements_to_process:
        processing_time = (DateTime.Now - start_time).TotalSeconds
        forms.alert(
            "Nessun isolamento da correggere trovato.\nTutti gli isolamenti ({}) hanno già il workset corretto.\n\nTempo di elaborazione: {:.2f} secondi".format(
                total_insulations, processing_time
            ), 
            title="Correzione Workset Isolamenti"
        )
        return
    
    # Chiedi conferma prima di procedere
    result = forms.alert(
        "Trovati {} isolamenti da correggere su {} totali.\nProcedere con la correzione?".format(
            len(elements_to_process), total_insulations
        ),
        title="Conferma correzione",
        yes=True, no=True
    )
    
    if not result:
        forms.alert("Operazione annullata dall'utente.", title="Correzione Workset Isolamenti")
        return
    
    # Contatori per report
    corrected = 0
    errors = 0
    
    # Tabella per report dettagliato
    output.print_md("## Report correzione workset isolamenti")
    table_data = []
    
    # Processa gli elementi per workset
    with revit.Transaction("Correzione Workset Isolamenti"):
        # Per ogni workset di destinazione
        for dest_workset_id in workset_groups:
            dest_workset_name = get_workset_name(WorksetId(dest_workset_id))
            
            # Per ogni workset di origine
            for source_workset_id in workset_groups[dest_workset_id]:
                source_workset_name = get_workset_name(WorksetId(source_workset_id))
                elements = workset_groups[dest_workset_id][source_workset_id]
                
                # Raggruppa per tipo di elemento
                element_types = {}
                for item in elements:
                    ref_element = item['ref_element']
                    element_type = get_category_name(ref_element)
                    
                    if element_type not in element_types:
                        element_types[element_type] = []
                    
                    element_types[element_type].append(ref_element.Id)
                
                # Elabora ogni tipo di elemento
                output.print_md("### Elaborazione gruppo: {} → {} ({} elementi)".format(
                    source_workset_name, dest_workset_name, len(elements)
                ))
                
                for element_type, element_ids in element_types.items():
                    batch_size = 50  # Elabora in batch di 50 elementi
                    total_batches = (len(element_ids) + batch_size - 1) // batch_size
                    
                    for batch_idx in range(total_batches):
                        start_idx = batch_idx * batch_size
                        end_idx = min((batch_idx + 1) * batch_size, len(element_ids))
                        current_batch = element_ids[start_idx:end_idx]
                        
                        # Step 1: Sposta gli elementi nel workset di origine (dell'isolamento)
                        for element_id in current_batch:
                            element = doc.GetElement(element_id)
                            element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).Set(source_workset_id)
                        
                        # Step 2: Riporta gli elementi nel workset originale
                        for element_id in current_batch:
                            element = doc.GetElement(element_id)
                            element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).Set(dest_workset_id)
                        
                        output.print_md("- Batch {}/{} ({} elementi di tipo '{}') elaborato".format(
                            batch_idx + 1, total_batches, len(current_batch), element_type
                        ))
                
                # Verifica e aggiorna il report
                for item in elements:
                    insulation = item['insulation']
                    ref_element = item['ref_element']
                    
                    # Ricarica l'isolamento per ottenere il nuovo workset
                    updated_insulation = doc.GetElement(insulation.Id)
                    updated_workset_id = updated_insulation.WorksetId.IntegerValue
                    
                    if updated_workset_id == dest_workset_id:
                        corrected += 1
                        
                        # Aggiungi alla tabella per il report
                        table_data.append([
                            "ID: " + str(insulation.Id.IntegerValue),
                            get_category_name(ref_element),
                            source_workset_name,
                            dest_workset_name
                        ])
                    else:
                        errors += 1
                        logger.warning("Non è stato possibile correggere l'isolamento con ID: {}".format(insulation.Id))
    
    # Calcola il tempo di elaborazione
    processing_time = (DateTime.Now - start_time).TotalSeconds
    
    # Report finale
    output.print_md("### Riepilogo:")
    output.print_md("- Isolamenti totali: {}".format(total_insulations))
    output.print_md("- Isolamenti da correggere: {}".format(len(elements_to_process)))
    output.print_md("- Isolamenti corretti con successo: {}".format(corrected))
    output.print_md("- Errori riscontrati: {}".format(errors))
    output.print_md("- Tempo di elaborazione: {:.2f} secondi".format(processing_time))
    
    if corrected > 0:
        output.print_md("\n### Dettaglio isolamenti corretti:")
        output.print_table(
            table_data,
            columns=["ID Isolamento", "Tipo Elemento", "Workset Originale", "Nuovo Workset"]
        )
    
    # Messaggio finale
#    if corrected > 0:
#        forms.alert(
#            "Correzione completata.\n\n{} isolamenti su {} sono stati corretti.\nTempo di elaborazione: {:.2f} secondi".format(
#                corrected, len(elements_to_process), processing_time
#            ),
#            title="Correzione Workset Isolamenti"
#        )
#    else:
#        forms.alert(
#            "Nessun isolamento è stato corretto.\nSi sono verificati {} errori.\nTempo di elaborazione: {:.2f} secondi".format(
#                errors, processing_time
#            ),
#            title="Correzione Workset Isolamenti"
#        )

# Esegui il programma principale
if __name__ == "__main__":
    correct_insulation_worksets()