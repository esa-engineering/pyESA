# -*- coding: utf-8 -*-
# Intestazione dello script con metadati
__title__   = "Filter\nCreate"
__doc__     = """Version = 1.0
Date    = 16.05.2025
________________________________________________________________
Creazione automatica filtri per viste basati su parametro di progetto.
1- Seleziona il parametro di progetto
2- Inserisci il prefisso per i filtri
3- Seleziona il file txt con i valori che può assumere il parametro
________________________________________________________________
Author(s):
Andrea Patti
"""

import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# Importazione per la classe List
clr.AddReference("System.Collections")
from System.Collections.Generic import List

from pyrevit import forms
from pyrevit import revit
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# Funzione per ottenere tutti i parametri di progetto disponibili
def ottieni_parametri_progetto():
    parametri = {}
    parametri_progetto = FilteredElementCollector(doc).OfClass(ParameterElement)
    
    for param in parametri_progetto:
        definition = param.GetDefinition()
        nome = definition.Name
        
        # Verifica se il parametro è associato ad almeno una categoria
        param_bind = doc.ParameterBindings.get_Item(definition)
        if param_bind and param_bind.Categories.Size > 0:
            parametri[nome] = param
    
    return parametri

# Funzione per ottenere categorie associate al parametro
def ottieni_categorie_parametro(param_elemento):
    try:
        # Ottieni le categorie a cui è associato il parametro
        param_bind = doc.ParameterBindings.get_Item(param_elemento.GetDefinition())
        
        if param_bind is None:
            forms.alert("Il parametro non è associato a nessuna categoria.", title="Errore")
            return None
            
        return [cat.Name for cat in param_bind.Categories]
    
    except Exception as e:
        forms.alert("Errore: " + str(e), title="Errore")
        return None

# Funzione per trovare l'ID del parametro in una categoria
def trova_parametro_id(nome_parametro, categoria_id):
    try:
        # Ottieni un elemento della categoria per trovare i parametri disponibili
        collector = FilteredElementCollector(doc).OfCategoryId(categoria_id).WhereElementIsNotElementType()
        if collector.GetElementCount() > 0:
            elemento = collector.FirstElement()
            for param in elemento.Parameters:
                if param.Definition.Name == nome_parametro:
                    return param.Definition.Id
    except:
        pass
    return None

# Funzione principale
def main():
    # Step 1: Ottenere tutti i parametri di progetto
    parametri = ottieni_parametri_progetto()
    
    if not parametri:
        forms.alert("Nessun parametro di progetto trovato.", title="Errore")
        return
    
    # Step 2: Mostrare tendina di selezione per il parametro
    opzioni_parametri = sorted(parametri.keys())
    nome_parametro = forms.SelectFromList.show(
        opzioni_parametri,
        title="Seleziona Parametro di Progetto",
        button_name="Seleziona"
    )
    
    if not nome_parametro:
        return
        
    param_elemento = parametri[nome_parametro]
    
    # Step 3: Richiedere il prefisso per i filtri
    prefisso_filtro = forms.ask_for_string(
        default="",
        prompt="Inserisci il prefisso per i nomi dei filtri:",
        title="Prefisso Filtri"
    )
    
    if not prefisso_filtro:
        return
    
    # Step 4: Richiesta file txt con i valori
    file_percorso = forms.pick_file(file_ext='txt', title='Seleziona file con lista valori')
    
    if not file_percorso:
        return
    
    # Leggi il file e ottieni i valori
    try:
        with open(file_percorso, 'r') as file:
            contenuto = file.read()
            valori = [valore.strip() for valore in contenuto.split(',')]
    except Exception as e:
        forms.alert("Errore nella lettura del file: " + str(e), title="Errore")
        return
    
    # Step 5: Verifica a quali categorie è associato il parametro
    categorie_nomi = ottieni_categorie_parametro(param_elemento)
    
    if not categorie_nomi:
        return
    
    # Mostra all'utente le categorie associate
    forms.alert("Il parametro è associato alle seguenti categorie:\n" + "\n".join(categorie_nomi), 
                title="Categorie associate")
    
    # Step 6: Crea i filtri
    with revit.Transaction("Creazione Filtri per " + nome_parametro):
        filtri_creati = 0
        filtri_esistenti_count = 0
        
        for valore in valori:
            if not valore:  # Salta valori vuoti
                continue
                
            nome_filtro = prefisso_filtro + "_" + valore
            
            # Verifica se il filtro esiste già
            esistente = False
            filtri_esistenti = FilteredElementCollector(doc).OfClass(ParameterFilterElement)
            for filtro in filtri_esistenti:
                if filtro.Name == nome_filtro:
                    esistente = True
                    break
            
            if esistente:
                print("Il filtro '{}' esiste già.".format(nome_filtro))
                filtri_esistenti_count += 1
                continue
            
            # Crea lista di categorie da usare per il filtro
            categorie_filtro = List[ElementId]()
            param_ids = []
            
            for nome_categoria in categorie_nomi:
                try:
                    categoria = None
                    for cat in doc.Settings.Categories:
                        if cat.Name == nome_categoria:
                            categoria = cat
                            break
                    
                    if categoria and categoria.AllowsBoundParameters:
                        cat_id = ElementId(categoria.Id.IntegerValue)
                        categorie_filtro.Add(cat_id)
                        
                        # Trova l'ID del parametro per questa categoria
                        param_id = trova_parametro_id(nome_parametro, cat_id)
                        if param_id and param_id not in param_ids:
                            param_ids.append(param_id)
                            
                except Exception as e:
                    print("Errore per categoria {}: {}".format(nome_categoria, str(e)))
            
            # Se non ci sono categorie valide, salta
            if categorie_filtro.Count == 0:
                print("Nessuna categoria valida trovata per il filtro '{}'.".format(nome_filtro))
                continue
                
            # Se non è stato possibile trovare il parametro, salta
            if not param_ids:
                print("Parametro '{}' non trovato in nessuna categoria.".format(nome_parametro))
                continue
            
            # Crea il filtro
            try:
                # Usa il primo ID del parametro trovato
                param_id = param_ids[0]
                
                # Crea la regola del filtro
                regola = ParameterFilterRuleFactory.CreateEqualsRule(
                    param_id, 
                    valore,
                    True  # case sensitive
                )
                
                # Crea il filtro logico utilizzando la regola
                element_filter = ElementParameterFilter(regola)
                
                # Crea il filtro
                nuovo_filtro = ParameterFilterElement.Create(
                    doc, 
                    nome_filtro, 
                    categorie_filtro, 
                    element_filter
                )
                
                print("Filtro creato: {}".format(nome_filtro))
                filtri_creati += 1
                
            except Exception as e:
                print("Errore nella creazione del filtro '{}': {}".format(nome_filtro, str(e)))
    
    # Messaggio di completamento
    forms.alert("Processo completato.\nFiltri creati: {}\nFiltri già esistenti: {}".format(
        filtri_creati, filtri_esistenti_count), 
        title="Completato")

# Esegui la funzione principale
if __name__ == "__main__":
    main()