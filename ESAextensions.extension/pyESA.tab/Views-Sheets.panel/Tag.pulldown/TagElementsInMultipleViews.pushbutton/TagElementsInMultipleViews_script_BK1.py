# -*- coding: utf-8 -*-
# Intestazione dello script con metadati
__title__   = "Tag Category Elements in\nMultiple Views"
__doc__     = """Version = 1.0
Date    = 09.05.2025
________________________________________________________________
Tagga tutti gli elementi della categoria selezionata visibili nelle viste selezionate con il tag scelto
con possibilità di aggiungere una leader line di tipo Attached End.

N.B. La lunghezza della leader line che viene utilizzata dallo script è quella impostata in Revit nel
comando "Tag All" , scheda "Annotate", gruppo "Tag", opzione "Tag All" e non è modificabile dallo script.
________________________________________________________________
Author(s):
Andrea Patti, Tommaso Lorenzi
"""

# -------------------------------
# SEZIONE IMPORT MODULI
# -------------------------------
# Importazione dei moduli necessari da pyRevit e Revit API

import pyrevit
from pyrevit import revit, script, DB
from pyrevit import forms
from rpw.ui.forms import TaskDialog, FlexForm, Label, TextBox, CheckBox, Separator, Button
from Autodesk.Revit.DB import Transaction, BuiltInCategory, ElementId, XYZ
from System.Collections.Generic import List

# -------------------------------
# SEZIONE INIZIALIZZAZIONE
# -------------------------------
# Ottenimento del documento attivo di Revit
doc = revit.doc

# -------------------------------
# SEZIONE SELEZIONE CATEGORIA
# -------------------------------
# Definizione delle categorie supportate con relativi BuiltInCategory e BuiltInParameter
supported_categories = [
    {"name": "Rooms", "category": BuiltInCategory.OST_Rooms, "tag_category": BuiltInCategory.OST_RoomTags},
    {"name": "Doors", "category": BuiltInCategory.OST_Doors, "tag_category": BuiltInCategory.OST_DoorTags},
    {"name": "Windows", "category": BuiltInCategory.OST_Windows, "tag_category": BuiltInCategory.OST_WindowTags},
    {"name": "Walls", "category": BuiltInCategory.OST_Walls, "tag_category": BuiltInCategory.OST_WallTags},
    {"name": "Furniture", "category": BuiltInCategory.OST_Furniture, "tag_category": BuiltInCategory.OST_FurnitureTags},
    {"name": "Mechanical Equipment", "category": BuiltInCategory.OST_MechanicalEquipment, "tag_category": BuiltInCategory.OST_MechanicalEquipmentTags},
    {"name": "Plumbing Fixtures", "category": BuiltInCategory.OST_PlumbingFixtures, "tag_category": BuiltInCategory.OST_PlumbingFixtureTags},
    {"name": "Electrical Fixtures", "category": BuiltInCategory.OST_ElectricalFixtures, "tag_category": BuiltInCategory.OST_ElectricalFixtureTags},
    {"name": "Electrical Equipment", "category": BuiltInCategory.OST_ElectricalEquipment, "tag_category": BuiltInCategory.OST_ElectricalEquipmentTags},
    {"name": "Pipes", "category": BuiltInCategory.OST_PipeCurves, "tag_category": BuiltInCategory.OST_PipeTags},
    {"name": "Ducts", "category": BuiltInCategory.OST_DuctCurves, "tag_category": BuiltInCategory.OST_DuctTags}
]

# Estrai solo i nomi delle categorie per la selezione
category_names = [cat["name"] for cat in supported_categories]

# Mostra all'utente una finestra di dialogo per selezionare la categoria
selected_category_name = forms.SelectFromList.show(
    category_names,
    multiselect=False,
    button_name='Select Category'
)

# Se l'utente non seleziona una categoria, termina lo script
if not selected_category_name:
    script.exit()

# Trova la categoria selezionata nel dizionario
selected_category = next((cat for cat in supported_categories if cat["name"] == selected_category_name), None)

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
    button_name='Select'  # Testo del pulsante di conferma
)

# Se l'utente non seleziona le viste, termina lo script
if not views_selected:
    script.exit()

# -------------------------------
# SEZIONE SELEZIONE TAG
# -------------------------------
# Recupera tutti i simboli di tag per la categoria selezionata
tags = list(DB.FilteredElementCollector(doc)
           .OfClass(DB.FamilySymbol)
           .OfCategory(selected_category["tag_category"]))

# Ordina i tag per nome combinato famiglia+tipo
tags.sort(key=lambda x: x.FamilyName + " - " + x.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString())

# Prepara una lista per i nomi dei tag formattati
tags_list = []

# Popola la lista con i nomi formattati "Famiglia - Tipo"
for tag in tags:
    family_name = tag.FamilyName  # Nome della famiglia
    type_name = tag.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()  # Nome del tipo
    tags_list.append(family_name + " - " + type_name)  # Formatto e aggiungo alla lista

    # Se non ci sono tag disponibili per la categoria, avvisa l'utente e termina
if not tags_list:
    TaskDialog.Show("Errore", "Non sono stati trovati tag per la categoria {}.".format(selected_category_name))
    script.exit()

# Mostra all'utente una finestra per selezionare il tipo di tag
selected_tag = forms.SelectFromList.show(
    tags_list,
    multiselect=False,  # Selezione singola
    button_name='Select {} Tag'.format(selected_category_name)  # Testo del pulsante
)

# Se l'utente non seleziona un tag, termina lo script
if not selected_tag:
    script.exit()

# -------------------------------
# SEZIONE OPZIONI LEADER LINE
# -------------------------------
# Crea un form per raccogliere le preferenze di leader line
components = [
    Label("Impostazioni Leader Line:"),
    CheckBox("use_leader", "Usa Leader Line (tipo Attached End)", default=True),
    Separator(),
    Button("OK")
]

# Mostra il form e ottieni i valori
form = FlexForm("Opzioni Leader Line", components)
form.show()

# Estrai i valori dal form
use_leader = form.values["use_leader"]

# -------------------------------
# SEZIONE CONTEGGIO ELEMENTI
# -------------------------------
# Raccogli tutti gli elementi della categoria selezionata nel progetto
elements_collector = DB.FilteredElementCollector(doc) \
                    .OfCategory(selected_category["category"]) \
                    .WhereElementIsNotElementType() \
                    .ToElements()

# Stampa il numero totale di elementi
print("N. di {} presenti a progetto: {}".format(selected_category_name, len(elements_collector)))
print("---------------------------------------------------------------------------")

# Inizializza il contatore dei tag
n_tag = 0

# -------------------------------
# SEZIONE OPERAZIONI PRINCIPALI
# -------------------------------

# Stampa il numero di viste selezionate su cui andrà a lavorare lo script
print("N. floor plan selezionate: {}".format(len(views_selected)))

# Trova il tag corrispondente alla selezione dell'utente
tag_type = None
for fs in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(selected_category["tag_category"]):
    fam_param = fs.FamilyName
    type_param = fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    fam_type_param = fam_param + " - " + type_param
    
    if fam_type_param == selected_tag:
        tag_type = fs
        break

# Gestione errore se il tag non viene trovato
if tag_type is None:
    print("Errore: Tag selezionato non trovato.")
    script.exit()

# Assicurarsi che il tipo di tag sia attivo
if not tag_type.IsActive:
    t_activate = Transaction(doc, "Activate Tag Type")
    t_activate.Start()
    tag_type.Activate()
    t_activate.Commit()

# Inizia la transazione per modificare il documento
t = Transaction(doc, "Tag {}".format(selected_category_name))
t.Start()

try:
    # Itera su tutte le viste selezionate
    for view in views_selected:
        # Trova tutti gli elementi visibili nella vista corrente
        visible_elements = DB.FilteredElementCollector(doc, view.Id) \
                            .OfCategory(selected_category["category"]) \
                            .WhereElementIsNotElementType() \
                            .ToElements()
        
        # Stampa informazioni di debug con nome vista e numero di elementi taggati nella vista
        print("Elaborazione vista: {} - N. {} nella vista: {}".format(view.Name, selected_category_name, len(visible_elements)))
        
        # Contatore locale per questa vista
        view_tags = 0
        
        # Itera su tutti gli elementi visibili
        for element in visible_elements:
            try:
                # Gestione diversa a seconda della categoria
                if selected_category["name"] == "Rooms":
                    # Per le stanze, usa il LocationPoint
                    location = element.Location
                    if isinstance(location, DB.LocationPoint):
                        point = location.Point
                        
                        # Crea il tag nella posizione della stanza
                        if use_leader:
                            # Quando usiamo la leader line di tipo "Attached End"
                            # Calcola un punto a una certa distanza per il posizionamento del tag
                            # Creiamo un vettore nella direzione Y
                            direction = XYZ(0, 1, 0)
                            tag = DB.IndependentTag.Create(
                                doc, 
                                tag_type.Id, 
                                view.Id, 
                                DB.Reference(element), 
                                True,
                                DB.TagOrientation.Horizontal, 
                                point
                            )        
                        else:
                            # Crea il tag senza leader
                            tag = DB.IndependentTag.Create(
                                doc, 
                                tag_type.Id, 
                                view.Id, 
                                DB.Reference(element), 
                                False,
                                DB.TagOrientation.Horizontal, 
                                point
                            )
                        
                        view_tags += 1
                else:
                    # Per altri elementi, usa il metodo generico
                    # Ottieni il punto di inserimento dell'elemento
                    if hasattr(element, "Location"):
                        location = element.Location
                        if isinstance(location, DB.LocationPoint):
                            point = location.Point
                        elif isinstance(location, DB.LocationCurve):
                            # Per elementi con curve (come muri), usa il punto medio
                            curve = location.Curve
                            point = curve.Evaluate(0.5, True)
                        else:
                            # Se non ha Location riconoscibile, usa il bounding box
                            bbox = element.get_BoundingBox(view)
                            if bbox:
                                point = (bbox.Min + bbox.Max) / 2
                            else:
                                # Salta l'elemento se non possiamo ottenere una posizione
                                continue
                        
                        # Crea il tag nella posizione dell'elemento
                        if use_leader:
                            # Quando usiamo la leader line di tipo "Attached End"
                            # Calcola un punto a una certa distanza per il posizionamento del tag
                            # Creiamo un vettore nella direzione Y
                            direction = XYZ(0, 1, 0)
                            tag = DB.IndependentTag.Create(
                                doc, 
                                tag_type.Id, 
                                view.Id, 
                                DB.Reference(element), 
                                True,
                                DB.TagOrientation.Horizontal, 
                                point
                            )
                        else:
                            # Crea il tag senza leader
                            tag = DB.IndependentTag.Create(
                                doc, 
                                tag_type.Id, 
                                view.Id, 
                                DB.Reference(element), 
                                False,
                                DB.TagOrientation.Horizontal, 
                                point
                            )
                        
                        view_tags += 1
            except Exception as e:
                print("Errore nel taggare {}: {}".format(element.Id, str(e)))
        
        # Incrementa il contatore totale dei tag
        n_tag += view_tags
        print("Tag aggiunti in questa vista: {}".format(view_tags))
    
    # Chiude la transazione
    t.Commit()
    
    # Stampa il riepilogo finale
    print("---------------------------------------------------------------------------")
    print("{} Tag aggiunti con successo per la categoria {}.".format(n_tag, selected_category_name))
    if use_leader:
        print("Tag creati con leader line di tipo Attached End.")
        print("Nota: La lunghezza delle leader line potrebbe essere limitata dalle impostazioni di Revit.")
    else:
        print("Tag creati senza leader line.")

except Exception as e:
    # In caso di errore, annulla la transazione
    t.RollBack()
    print("Errore durante l'esecuzione: {}".format(str(e)))
    import traceback
    print(traceback.format_exc())