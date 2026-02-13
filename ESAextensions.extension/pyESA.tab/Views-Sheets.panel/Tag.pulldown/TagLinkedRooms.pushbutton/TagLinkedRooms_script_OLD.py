# -*- coding: utf-8 -*-
# Intestazione dello script con metadati
__title__   = "Linked Room Tag\nin Multiple Views"
__doc__     = """Version = 1.0
Date    = 08.05.2025
________________________________________________________________
Tagga tutte le rooms visibili nelle viste selezionate con il tag scelto
da un modello linkato selezionato
________________________________________________________________
Author(s): Tommaso Lorenzi, Andrea Patti   
"""

# -------------------------------
# SEZIONE IMPORT MODULI
# -------------------------------
# Importazione dei moduli necessari da pyRevit e Revit API

import pyrevit
from pyrevit import revit, script, DB
from pyrevit import forms
from rpw.ui.forms import TaskDialog, FlexForm, Label, TextBox, CheckBox, Separator, Button
from Autodesk.Revit.DB import Transaction, RevitLinkInstance, FilteredElementCollector, BuiltInCategory, Reference, LinkElementId, UV

# -------------------------------
# SEZIONE INIZIALIZZAZIONE
# -------------------------------
# Ottenimento del documento attivo di Revit
doc = revit.doc

# -------------------------------
# SEZIONE SELEZIONE MODELLO LINKATO
# -------------------------------
# Recupera tutti i modelli linkati nel progetto
link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()

# Verifica se ci sono modelli linkati
if not link_instances:
    forms.alert("Non ci sono modelli linkati nel progetto corrente.", exitscript=True)

# Prepara la lista per la selezione
link_names = [link.Name.split(" : ")[0] for link in link_instances]
link_dict = dict(zip(link_names, link_instances))

# Mostra all'utente una finestra di dialogo per selezionare il modello linkato
selected_link_name = forms.SelectFromList.show(
    sorted(link_names),
    multiselect=False,
    button_name='Seleziona Modello Linkato'
)

# Verifica se è stato selezionato un modello
if not selected_link_name:
    forms.alert("Nessun modello linkato selezionato. Operazione annullata.", exitscript=True)

# Recupera l'istanza del modello linkato selezionato
selected_link = link_dict[selected_link_name]

# Ottieni il documento del modello linkato
link_doc = selected_link.GetLinkDocument()

# -------------------------------
# SEZIONE SELEZIONE VISTE
# -------------------------------
# Recupera tutte le viste del progetto
all_views = revit.query.get_all_views(doc=doc)

# Filtra solo le floor plan e le ordina per nome
floor_plans = [view for view in all_views if view.ViewType == DB.ViewType.FloorPlan]
floor_plans.sort(key=lambda x:x.Name)

# Mostra all'utente una finestra di dialogo per selezionare le viste
views_selected = forms.SelectFromList.show(
    floor_plans,
    multiselect=True,  # Permette selezione multipla
    name_attr='Name',  # Mostra il nome della vista nell'interfaccia
    button_name='Seleziona Viste'  # Testo del pulsante di conferma
)

# Verifica se sono state selezionate viste
if not views_selected:
    forms.alert("Nessuna vista selezionata. Operazione annullata.", exitscript=True)

# -------------------------------
# SEZIONE SELEZIONE ROOM TAG
# -------------------------------
# Recupera tutti i simboli di room tag nel progetto
room_tags = list(DB.FilteredElementCollector(doc)
               .OfClass(DB.FamilySymbol)
               .OfCategory(DB.BuiltInCategory.OST_RoomTags))

# Ordina i tag per nome combinato famiglia+tipo
room_tags.sort(key=lambda x: x.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString())

# Prepara una lista per i nomi dei tag formattati
room_tags_list = []

# Popola la lista con i nomi formattati "Famiglia - Tipo"
for tag in room_tags:
    family_name = tag.FamilyName  # Nome della famiglia
    type_name = tag.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()  # Nome del tipo
    room_tags_list.append(family_name + " - " + type_name)  # Formatto e aggiungo alla lista

# Mostra all'utente una finestra per selezionare il tipo di tag
selected_tag = forms.SelectFromList.show(
    room_tags_list,
    multiselect=False,  # Selezione singola
    button_name='Seleziona Room Tag'  # Testo del pulsante
)

# Verifica se è stato selezionato un tag
if not selected_tag:
    forms.alert("Nessun Room Tag selezionato. Operazione annullata.", exitscript=True)

# -------------------------------
# SEZIONE CONTEGGIO ROOMS NEL MODELLO LINKATO
# -------------------------------
# Raccogli tutte le rooms nel modello linkato
rooms_collector = FilteredElementCollector(link_doc) \
                    .OfCategory(BuiltInCategory.OST_Rooms) \
                    .WhereElementIsNotElementType() \
                    .ToElements()

# Stampa il numero totale di rooms nel modello linkato
print("N. di rooms presenti nel modello linkato '" + selected_link_name + "': " + str(len(rooms_collector)))
print("---------------------------------------------------------------------------")

# Inizializza il contatore dei tag
n_tag = int()

# -------------------------------
# SEZIONE OPERAZIONI PRINCIPALI
# -------------------------------

# Stampa il numero di viste selezionate su cui andrà a lavorare lo script
print("N. floor plan selezionate: " + str(len(views_selected)))

# Trova il RoomTag corrispondente alla selezione dell'utente
tag_type = None
for fs in FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(BuiltInCategory.OST_RoomTags):
    fam_param = fs.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
    type_param = fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    fam_type_param = fam_param + " - " + type_param
    
    if fam_type_param == selected_tag:
        tag_type = fs
        break

# Gestione errore se il tag non viene trovato
if tag_type is None:
    forms.alert("Errore: Non è stato trovato il Room Tag selezionato.", exitscript=True)
else:
    # Assicurati che il tipo di tag sia attivo
    if not tag_type.IsActive:
        t_activate = Transaction(doc, "Activate Tag Type")
        t_activate.Start()
        tag_type.Activate()
        t_activate.Commit()
    
    # Inizia la transazione per modificare il documento
    t = Transaction(doc, "Tag Rooms from Linked Model")
    t.Start()

    # Itera su tutte le viste selezionate
    for view in views_selected:
        # Ottieni il livello associato alla vista corrente
        view_level_id = view.GenLevel.Id if hasattr(view, 'GenLevel') and view.GenLevel is not None else None
        
        # Contatore per le stanze taggate nella vista corrente
        view_tag_count = 0
        
        # Stampa informazioni iniziali
        print("Elaborazione vista: " + view.Name + " - Livello: " + (view.GenLevel.Name if hasattr(view, 'GenLevel') and view.GenLevel is not None else "Nessun livello"))
        
        # Itera su tutte le stanze nel modello linkato
        for room in rooms_collector:
            try:
                # Verifica se la stanza appartiene al livello della vista corrente
                room_level_id = room.LevelId
                
                # Confronto dei livelli (nota: questo è un controllo semplificato, 
                # potrebbe richiedere una logica più complessa per mappare i livelli tra modelli)
                if view_level_id and room_level_id:
                    # Confrontiamo i nomi dei livelli invece degli ID poiché gli ID potrebbero essere diversi tra modelli
                    view_level_name = doc.GetElement(view_level_id).Name
                    room_level_name = link_doc.GetElement(room_level_id).Name
                    
                    # Se i nomi dei livelli non corrispondono, salta questa stanza
                    if view_level_name != room_level_name:
                        continue
                
                location = room.Location
                if isinstance(location, DB.LocationPoint):
                    point = location.Point
                    
                    # Trasforma il punto dal sistema di coordinate del modello linkato a quello del modello corrente
                    transformed_point = selected_link.GetTotalTransform().OfPoint(point)
                    
                    # Crea un riferimento all'elemento linkato
                    # Utilizziamo un metodo alternativo che funziona in più versioni dell'API
                    linked_elem_id = room.Id
                    linked_elem_ref = DB.Reference(room)
                    
                    # Verifica se la stanza è visibile nella vista corrente
                    # (Questa è una verifica semplificata, potrebbe essere necessario un controllo più accurato)
                    bbox = room.get_BoundingBox(None)
                    if bbox:
                        # Crea il tag per la stanza nel modello linkato
                        try:
                            # Creiamo un nuovo tag per la stanza collegata
                            # Convertiamo il punto 3D in un punto UV per la vista
                            uv_point = DB.UV(transformed_point.X, transformed_point.Y)
                            
                            # Controlliamo se la stanza è effettivamente visibile nella vista corrente
                            # Questo richiede verificare se il bounding box della stanza si interseca con i limiti della vista
                            bbox = room.get_BoundingBox(None)
                            
                            if bbox:  # Verifica che la stanza abbia un bounding box
                                # Creare un tag room per un elemento collegato
                                linked_tag = doc.Create.NewRoomTag(
                                    DB.LinkElementId(selected_link.Id, room.Id),
                                    uv_point,
                                    view.Id
                                )
                            
                            # Se il tag è stato creato con successo, impostiamo il tipo e incrementiamo il contatore
                            if linked_tag:
                                # Impostiamo il tipo di tag se necessario (potrebbe essere richiesto un cambio nel tipo)
                                linked_tag.ChangeTypeId(tag_type.Id)
                                view_tag_count += 1
                                n_tag += 1
                        except Exception as inner_e:
                            print("Errore nella creazione del tag: " + str(inner_e))
            except Exception as e:
                print("Errore nel taggare la stanza {}: {}".format(room.Id, str(e)))
        
        # Stampa informazioni di debug con nome vista e numero di rooms taggate nella vista
        print("Elaborazione vista: " + view.Name + " - N. Rooms taggate nella vista: " + str(view_tag_count))
    
    # Chiude la transazione
    t.Commit()
    
    # Stampa il riepilogo finale
    print("---------------------------------------------------------------------------")
    print(str(n_tag) + " Tag aggiunti con successo.")