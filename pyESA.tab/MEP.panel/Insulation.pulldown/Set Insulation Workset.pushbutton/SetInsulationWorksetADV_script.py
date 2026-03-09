# -*- coding: utf-8 -*-
__title__   = "Set Insulation\nWorkset"
__doc__     = """Version = 1.1
Date    = 03.03.2026
________________________________________________________________
Corregge i workset degli isolamenti.
Lo script verifica che il workset di ogni isolamento coincida con
quello dell'elemento a cui e associato (canale, raccordo, tubazione).
Se non coincide, imposta il workset dell'isolamento uguale a quello
dell'elemento correlato.

Compatibile con Revit 2020-2025 e Revit 2026+.
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

# ---------------------------------------------------------------------------
# Compatibilita ElementId: Revit <= 2025 usa .IntegerValue, Revit >= 2026 usa .Value
# ---------------------------------------------------------------------------
def get_element_id_value(eid):
    """Restituisce il valore numerico di un ElementId, compatibile con tutte le versioni."""
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


def create_element_id(value):
    """Crea un ElementId da un valore numerico, compatibile con tutte le versioni."""
    try:
        return ElementId(int(value))
    except Exception:
        from System import Int64
        return ElementId(Int64(value))


# Funzione per ottenere l'elemento di riferimento da un isolamento
def get_reference_element(insulation):
    try:
        ref_id = insulation.HostElementId
        if ref_id and get_element_id_value(ref_id) != -1:
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
    pipe_insulations = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_PipeInsulations) \
        .WhereElementIsNotElementType().ToElements()
    duct_insulations = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_DuctInsulations) \
        .WhereElementIsNotElementType().ToElements()

    all_insulations = list(pipe_insulations) + list(duct_insulations)
    total_insulations = len(all_insulations)

    # Raggruppa elementi per workset di destinazione
    workset_groups = {}
    elements_to_process = []
    skipped = 0

    # Analizza ogni isolamento e raggruppa per workset
    for insulation in all_insulations:
        ref_element = get_reference_element(insulation)
        if not ref_element:
            skipped += 1
            continue

        insulation_workset_id = get_element_id_value(insulation.WorksetId)
        ref_element_workset_id = get_element_id_value(ref_element.WorksetId)

        if insulation_workset_id != ref_element_workset_id:
            elements_to_process.append({
                'insulation': insulation,
                'ref_element': ref_element,
                'original_workset_id': ref_element.WorksetId,
                'insulation_workset_id': insulation.WorksetId
            })

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
            "Nessun isolamento da correggere trovato.\n"
            "Tutti gli isolamenti ({}) hanno gia il workset corretto.\n\n"
            "Tempo di elaborazione: {:.2f} secondi".format(
                total_insulations, processing_time),
            title="Correzione Workset Isolamenti"
        )
        return

    # Chiedi conferma prima di procedere
    result = forms.alert(
        "Trovati {} isolamenti da correggere su {} totali.\n"
        "Procedere con la correzione?".format(
            len(elements_to_process), total_insulations),
        title="Conferma correzione",
        yes=True, no=True
    )

    if not result:
        forms.alert("Operazione annullata dall'utente.",
                     title="Correzione Workset Isolamenti")
        return

    corrected = 0
    errors = 0

    output.print_md("## Report correzione workset isolamenti")
    table_data = []

    with revit.Transaction("Correzione Workset Isolamenti"):
        for dest_workset_id in workset_groups:
            dest_workset_name = get_workset_name(WorksetId(dest_workset_id))

            for source_workset_id in workset_groups[dest_workset_id]:
                source_workset_name = get_workset_name(WorksetId(source_workset_id))
                elements = workset_groups[dest_workset_id][source_workset_id]

                element_types = {}
                for item in elements:
                    ref_element = item['ref_element']
                    element_type = get_category_name(ref_element)
                    if element_type not in element_types:
                        element_types[element_type] = []
                    element_types[element_type].append(ref_element.Id)

                output.print_md(
                    "### Elaborazione gruppo: {} -> {} ({} elementi)".format(
                        source_workset_name, dest_workset_name, len(elements)))

                for element_type, element_ids in element_types.items():
                    batch_size = 50
                    total_batches = (len(element_ids) + batch_size - 1) // batch_size

                    for batch_idx in range(total_batches):
                        start_idx = batch_idx * batch_size
                        end_idx = min((batch_idx + 1) * batch_size, len(element_ids))
                        current_batch = element_ids[start_idx:end_idx]

                        # Step 1: Sposta gli elementi nel workset dell'isolamento
                        for element_id in current_batch:
                            element = doc.GetElement(element_id)
                            element.get_Parameter(
                                BuiltInParameter.ELEM_PARTITION_PARAM
                            ).Set(source_workset_id)

                        # Step 2: Riporta gli elementi nel workset originale
                        for element_id in current_batch:
                            element = doc.GetElement(element_id)
                            element.get_Parameter(
                                BuiltInParameter.ELEM_PARTITION_PARAM
                            ).Set(dest_workset_id)

                        output.print_md(
                            "- Batch {}/{} ({} elementi di tipo '{}') elaborato".format(
                                batch_idx + 1, total_batches,
                                len(current_batch), element_type))

                # Verifica e aggiorna il report
                for item in elements:
                    insulation = item['insulation']
                    ref_element = item['ref_element']

                    updated_insulation = doc.GetElement(insulation.Id)
                    updated_workset_id = get_element_id_value(updated_insulation.WorksetId)

                    if updated_workset_id == dest_workset_id:
                        corrected += 1
                        table_data.append([
                            "ID: " + str(get_element_id_value(insulation.Id)),
                            get_category_name(ref_element),
                            source_workset_name,
                            dest_workset_name
                        ])
                    else:
                        errors += 1
                        logger.warning(
                            "Non e stato possibile correggere "
                            "l'isolamento con ID: {}".format(
                                get_element_id_value(insulation.Id)))

    processing_time = (DateTime.Now - start_time).TotalSeconds

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
            columns=["ID Isolamento", "Tipo Elemento",
                      "Workset Originale", "Nuovo Workset"]
        )


# Esegui il programma principale
if __name__ == "__main__":
    correct_insulation_worksets()