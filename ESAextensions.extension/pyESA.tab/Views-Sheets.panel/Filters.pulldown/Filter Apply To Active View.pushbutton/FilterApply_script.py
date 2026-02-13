# -*- coding: utf-8 -*-
# Intestazione dello script con metadati
__title__   = "Filter Apply\nActive View"
__doc__     = """Version = 1.2
Date    = 19.05.2025
________________________________________________________________
Applica filtri alla vista attiva e assegna colori casuali con pattern solid fill.
1- Selezionare uno o più filtri da applicare alla vista attiva
2- Se alla vista attiva è applicato un view template, il programma permette di sciegliere se applicare i filtri al view template. 
In caso negativo, l'operazione viene annullata.
3- Selezionare se applicare la sostituzione grafica al pattern di proiezione, al pattern di taglio o a entrambi.
4- Controlla se il tipo di vista supporta le sostituzioni grafiche.
________________________________________________________________
Author(s):
Andrea Patti
"""

import random
import clr

# Importa le librerie necessarie di Revit
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *

# Importa le librerie di PyRevit
from pyrevit import forms
from pyrevit import revit

# Ottieni il documento corrente e la vista attiva
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

def get_solid_fill_pattern_id(doc):
    """Ottiene l'ID del pattern solid fill cercando tra i pattern disponibili."""
    # Ottieni tutti i fill pattern
    fill_patterns = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
    
    # Prima cerca il pattern con nome "Solid fill"
    for pattern in fill_patterns:
        if pattern.Name == "Solid fill" or pattern.Name == "Solid Fill":
            return pattern.Id
        
    # Se non trova "Solid fill", cerca qualsiasi pattern che potrebbe essere solido
    for pattern in fill_patterns:
        # In alcune versioni di Revit, possiamo identificare un pattern solido dal suo nome
        if "solid" in pattern.Name.lower():
            return pattern.Id
    
    # Se ancora non trova nulla, prende il primo pattern (soluzione di ripiego)
    if fill_patterns.Count > 0:
        return fill_patterns.FirstElement().Id
    
    # Se non ci sono pattern disponibili, restituisce None
    return None

def apply_overrides(view_to_modify, filter_id, solid_fill_id, random_color, apply_to):
    """
    Applica le sostituzioni grafiche al filtro in base alle scelte dell'utente.
    
    Parameters:
    - view_to_modify: Vista o template a cui applicare il filtro
    - filter_id: ID del filtro da applicare
    - solid_fill_id: ID del pattern solid fill
    - random_color: Colore casuale da applicare
    - apply_to: Stringa che indica dove applicare le sostituzioni 
                ('projection', 'cut', o 'both')
    """
    # Ottieni le sostituzioni grafiche esistenti per il filtro (se presenti)
    existing_overrides = view_to_modify.GetFilterOverrides(filter_id)
    
    # Crea un nuovo oggetto OverrideGraphicSettings basato su quello esistente
    override_settings = OverrideGraphicSettings(existing_overrides)
    
    # Applica le sostituzioni in base alla scelta dell'utente
    if apply_to == 'projection' or apply_to == 'both':
        override_settings.SetSurfaceForegroundPatternVisible(True)
        override_settings.SetSurfaceForegroundPatternId(solid_fill_id)
        override_settings.SetSurfaceForegroundPatternColor(random_color)
    
    if apply_to == 'cut' or apply_to == 'both':
        override_settings.SetCutForegroundPatternVisible(True)
        override_settings.SetCutForegroundPatternId(solid_fill_id)
        override_settings.SetCutForegroundPatternColor(random_color)
    
    # Applica le impostazioni al filtro nella vista/template
    view_to_modify.SetFilterOverrides(filter_id, override_settings)

def main():
    # Verifica se è applicato un view template
    view_template_id = active_view.ViewTemplateId
    if not view_template_id.Equals(ElementId.InvalidElementId):
        template_name = doc.GetElement(view_template_id).Name
        
        # Chiedi all'utente cosa vuole fare
        question = "Alla vista corrente è applicato il view template '{}'.\nCosa vuoi fare?".format(template_name)
        options = ["Applica filtri al view template", "Annulla operazione"]
        selected_option = forms.alert(question, options=options)
        
        # Se l'utente ha scelto di annullare, termina lo script
        if selected_option == "Annulla operazione" or selected_option is None:
            forms.alert("Operazione annullata dall'utente.", exitscript=True)
            return
        
        # Altrimenti, continua con l'applicazione dei filtri al view template
        # Aggiorna la vista di destinazione al view template
        view_to_modify = doc.GetElement(view_template_id)
    else:
        view_to_modify = active_view
    
    # Verifica se la vista supporta le sostituzioni grafiche
    try:
        # Tenta di ottenere i filtri già applicati alla vista
        # Questo genererà un'eccezione se la vista non supporta le sostituzioni grafiche
        filters_ids = view_to_modify.GetFilters()
    except Exception:
        forms.alert('Questo tipo di vista non supporta le sostituzioni grafiche. Operazione annullata.', exitscript=True)
        return
    
    # Ottieni l'ID del solid fill pattern
    solid_fill_id = get_solid_fill_pattern_id(doc)
    if not solid_fill_id:
        forms.alert('Non è stato possibile trovare un pattern solido. Operazione annullata.', exitscript=True)
    
    # Ottieni tutti i filtri disponibili nel documento
    all_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
    
    # Crea una lista di nomi dei filtri per la selezione
    filter_names = {f.Name: f for f in all_filters}
    
    # Mostra una finestra di dialogo per selezionare i filtri
    selected_filter_names = forms.SelectFromList.show(
        sorted(filter_names.keys()),
        title='Seleziona i filtri da applicare',
        multiselect=True,
        button_name='Applica Filtri'
    )
    
    # Se nessun filtro è stato selezionato, esci
    if not selected_filter_names:
        forms.alert('Nessun filtro selezionato. Operazione annullata.', exitscript=True)
    
    # Chiedi all'utente dove applicare la sostituzione grafica
    pattern_options = {
        'Proiezione': 'projection', 
        'Taglio': 'cut', 
        'Entrambi': 'both'
    }
    
    selected_pattern_option = forms.CommandSwitchWindow.show(
        pattern_options.keys(),
        message='Seleziona dove applicare la sostituzione grafica:'
    )
    
    # Se l'utente annulla, esci
    if not selected_pattern_option:
        forms.alert('Nessuna opzione selezionata. Operazione annullata.', exitscript=True)
        return
    
    # Converti l'opzione selezionata nel valore corrispondente
    apply_to = pattern_options[selected_pattern_option]
    
    # Ottieni gli oggetti filtro selezionati
    selected_filters = [filter_names[name] for name in selected_filter_names]
    
    # Inizia una transazione
    with revit.Transaction('Applica filtri con colori casuali'):
        # Per ogni filtro selezionato
        for filter_elem in selected_filters:
            # Aggiungi il filtro alla vista/template selezionato se non esiste già
            if not view_to_modify.IsFilterApplied(filter_elem.Id):
                view_to_modify.AddFilter(filter_elem.Id)
            
            # Crea un colore casuale
            random_color = Color(
                random.randint(0, 255),
                random.randint(0, 255),
                random.randint(0, 255)
            )
            
            # Applica le sostituzioni grafiche in base alla scelta dell'utente
            apply_overrides(
                view_to_modify, 
                filter_elem.Id, 
                solid_fill_id, 
                random_color, 
                apply_to
            )
    
    # Prepara il messaggio di conferma
    target_type = "view template" if not view_template_id.Equals(ElementId.InvalidElementId) else "vista attiva"
    pattern_type_msg = {
        'projection': 'pattern di proiezione',
        'cut': 'pattern di taglio',
        'both': 'pattern di proiezione e di taglio'
    }
    
    # Mostra un messaggio di conferma
    forms.alert(
        'Operazione completata con successo!\n' + 
        'Sono stati applicati {} filtri con colori casuali al {}.\n'.format(len(selected_filters), target_type) +
        'Sostituzioni grafiche applicate al: {}.'.format(pattern_type_msg[apply_to])
    )

if __name__ == '__main__':
    main()