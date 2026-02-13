# -*- coding: utf-8 -*-
# Intestazione dello script con metadati
__title__   = "Linked Room Tag\nin Multiple Views"
__doc__     = """Version = 2.0
Date    = 29.05.2025
________________________________________________________________
Tagga tutte le rooms visibili nelle viste selezionate con il tag scelto
da un modello linkato selezionato.
NUOVO: Controllo preciso basato su punto di inserimento 3D e view range
________________________________________________________________
Author(s): Tommaso Lorenzi, Andrea Patti   
Improvements: 3D point detection with view range volume comparison
"""

# -------------------------------
# SEZIONE IMPORT MODULI
# -------------------------------
import pyrevit
from pyrevit import revit, script, DB
from pyrevit import forms
from rpw.ui.forms import TaskDialog
from Autodesk.Revit.DB import Transaction, RevitLinkInstance, FilteredElementCollector, BuiltInCategory, Reference, LinkElementId, UV
import sys

# -------------------------------
# SEZIONE INIZIALIZZAZIONE
# -------------------------------
doc = revit.doc
uidoc = revit.uidoc

# -------------------------------
# FUNZIONI HELPER MODIFICATE
# -------------------------------
def check_existing_room_tag(room, view, link_instance, doc):
    """
    Verifica se esiste già un tag per la stanza specificata nella vista
    """
    try:
        # Ottieni tutti i room tags nella vista
        room_tags_in_view = FilteredElementCollector(doc, view.Id) \
                           .OfCategory(BuiltInCategory.OST_RoomTags) \
                           .WhereElementIsNotElementType() \
                           .ToElements()
        
        # Verifica ogni tag
        for tag in room_tags_in_view:
            try:
                # Ottieni l'elemento taggato
                tagged_room_ref = tag.TaggedRoomId
                
                # Verifica se è un elemento linkato
                if tagged_room_ref and hasattr(tagged_room_ref, 'LinkedElementId'):
                    # Confronta gli ID
                    if (tagged_room_ref.LinkInstanceId == link_instance.Id and 
                        tagged_room_ref.LinkedElementId == room.Id):
                        return True
                
                # Per compatibilità con versioni precedenti, prova anche questo metodo
                tagged_room = tag.Room
                if tagged_room and tagged_room.Id == room.Id:
                    return True
                    
            except:
                continue
        
        return False
    except Exception as e:
        # In caso di errore, assumiamo che non ci siano tag duplicati
        return False

def get_room_insertion_point_3d(room, link_doc):
    """
    Ottiene il punto di inserimento della room con coordinate XYZ complete
    e calcola il punto a metà altezza
    """
    try:
        # Ottieni il punto di inserimento della room
        location = room.Location
        if not isinstance(location, DB.LocationPoint):
            # Se non ha LocationPoint, usa il centro del bounding box
            bbox = room.get_BoundingBox(None)
            if bbox:
                insertion_point = (bbox.Min + bbox.Max) / 2
            else:
                return None
        else:
            insertion_point = location.Point
        
        # Ottieni l'altezza della room
        room_height = get_room_height(room)
        if room_height is None:
            room_height = 3.0  # Altezza di default in metri
        
        # Calcola il punto a metà altezza
        # Il punto Z sarà: elevazione del livello + metà altezza room
        room_level = link_doc.GetElement(room.LevelId)
        if room_level:
            base_elevation = room_level.Elevation
            mid_height_z = base_elevation + (room_height / 2.0)
        else:
            mid_height_z = insertion_point.Z + (room_height / 2.0)
        
        # Crea il punto 3D a metà altezza
        mid_height_point = DB.XYZ(insertion_point.X, insertion_point.Y, mid_height_z)
        
        return mid_height_point
        
    except Exception as e:
        print("  ⚠ Errore nel calcolo punto 3D per room {}: {}".format(room.Id, str(e)))
        return None

def get_room_height(room):
    """
    Ottiene l'altezza della room dai parametri
    """
    try:
        # Prova diversi parametri per l'altezza
        height_param = room.get_Parameter(DB.BuiltInParameter.ROOM_HEIGHT)
        if height_param and height_param.HasValue:
            return height_param.AsDouble()
        
        # Parametro alternativo
        height_param = room.LookupParameter("Height")
        if height_param and height_param.HasValue:
            return height_param.AsDouble()
        
        # Se non trova l'altezza, usa il bounding box
        bbox = room.get_BoundingBox(None)
        if bbox:
            return bbox.Max.Z - bbox.Min.Z
            
        return None
        
    except:
        return None

def get_view_range_volume(view):
    """
    Ottiene il volume definito dal view range della vista
    Restituisce un dizionario con le elevazioni di top, cut plane e bottom
    """
    try:
        view_range = view.GetViewRange()
        if not view_range:
            return None
        
        # Ottieni il livello associato alla vista
        view_level = view.GenLevel
        if not view_level:
            return None
            
        base_elevation = view_level.Elevation
        
        # Ottieni le elevazioni del view range
        # Top
        top_level_id = view_range.GetLevelId(DB.PlanViewPlane.TopClipPlane)
        top_offset = view_range.GetOffset(DB.PlanViewPlane.TopClipPlane)
        if top_level_id != DB.ElementId.InvalidElementId:
            top_level = doc.GetElement(top_level_id)
            top_elevation = top_level.Elevation + top_offset
        else:
            top_elevation = base_elevation + top_offset
        
        # Bottom
        bottom_level_id = view_range.GetLevelId(DB.PlanViewPlane.ViewDepthPlane)
        bottom_offset = view_range.GetOffset(DB.PlanViewPlane.ViewDepthPlane)
        if bottom_level_id != DB.ElementId.InvalidElementId:
            bottom_level = doc.GetElement(bottom_level_id)
            bottom_elevation = bottom_level.Elevation + bottom_offset
        else:
            bottom_elevation = base_elevation + bottom_offset
        
        # Cut Plane
        cut_level_id = view_range.GetLevelId(DB.PlanViewPlane.CutPlane)
        cut_offset = view_range.GetOffset(DB.PlanViewPlane.CutPlane)
        if cut_level_id != DB.ElementId.InvalidElementId:
            cut_level = doc.GetElement(cut_level_id)
            cut_elevation = cut_level.Elevation + cut_offset
        else:
            cut_elevation = base_elevation + cut_offset
        
        return {
            'top': top_elevation,
            'bottom': bottom_elevation,
            'cut': cut_elevation,
            'base': base_elevation
        }
        
    except Exception as e:
        print("  ⚠ Errore nell'ottenere view range: {}".format(str(e)))
        return None

def is_point_in_view_range_volume(point_3d, view, link_transform):
    """
    Verifica se un punto 3D è all'interno del volume definito dal view range
    e dal crop region della vista
    """
    try:
        # Trasforma il punto dalle coordinate del link a quelle del modello host
        transformed_point = link_transform.OfPoint(point_3d)
        
        # 1. Controllo del crop region (piano XY)
        if view.CropBoxActive:
            crop_box = view.CropBox
            if (transformed_point.X < crop_box.Min.X or transformed_point.X > crop_box.Max.X or
                transformed_point.Y < crop_box.Min.Y or transformed_point.Y > crop_box.Max.Y):
                return False, "Fuori crop region"
        
        # 2. Controllo del view range (asse Z)
        view_range_info = get_view_range_volume(view)
        if not view_range_info:
            return False, "View range non disponibile"
        
        point_z = transformed_point.Z
        
        # Verifica se il punto è nel range verticale
        if point_z < view_range_info['bottom'] or point_z > view_range_info['top']:
            return False, "Fuori view range (Z: {:.2f}, Range: {:.2f} - {:.2f})".format(
                point_z, view_range_info['bottom'], view_range_info['top'])
        
        return True, "Visibile"
        
    except Exception as e:
        return False, "Errore nel controllo: {}".format(str(e))

def get_room_center_point(room):
    """
    Ottiene il punto centrale della stanza (mantenuto per compatibilità)
    """
    location = room.Location
    if isinstance(location, DB.LocationPoint):
        return location.Point
    
    # Se non ha LocationPoint, usa il centro del bounding box
    bbox = room.get_BoundingBox(None)
    if bbox:
        return (bbox.Min + bbox.Max) / 2
    
    return None

def match_levels_by_elevation(view_level, room_level, tolerance=0.5):
    """
    Confronta i livelli basandosi sull'elevazione invece che sul nome
    """
    try:
        view_elev = view_level.Elevation
        room_elev = room_level.Elevation
        return abs(view_elev - room_elev) < tolerance
    except:
        # Fallback al confronto per nome
        return view_level.Name == room_level.Name

# -------------------------------
# SEZIONE SELEZIONE MODELLO LINKATO
# -------------------------------
link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()

if not link_instances:
    forms.alert("Non ci sono modelli linkati nel progetto corrente.", exitscript=True)

# Filtra solo i link caricati
loaded_links = [link for link in link_instances if link.GetLinkDocument() is not None]

if not loaded_links:
    forms.alert("Non ci sono modelli linkati caricati nel progetto.", exitscript=True)

link_names = [link.Name.split(" : ")[0] for link in loaded_links]
link_dict = dict(zip(link_names, loaded_links))

selected_link_name = forms.SelectFromList.show(
    sorted(link_names),
    multiselect=False,
    button_name='Seleziona Modello Linkato',
    title='Selezione Modello Linkato'
)

if not selected_link_name:
    script.exit()

selected_link = link_dict[selected_link_name]
link_doc = selected_link.GetLinkDocument()

# -------------------------------
# SEZIONE SELEZIONE VISTE
# -------------------------------
all_views = FilteredElementCollector(doc).OfClass(DB.View).ToElements()
floor_plans = [v for v in all_views if v.ViewType == DB.ViewType.FloorPlan and not v.IsTemplate]
floor_plans.sort(key=lambda x: x.Name)

if not floor_plans:
    forms.alert("Non ci sono Floor Plans disponibili nel progetto.", exitscript=True)

views_selected = forms.SelectFromList.show(
    floor_plans,
    multiselect=True,
    name_attr='Name',
    button_name='Seleziona Viste',
    title='Selezione Floor Plans'
)

if not views_selected:
    script.exit()

# -------------------------------
# SEZIONE SELEZIONE ROOM TAG
# -------------------------------
room_tags = list(DB.FilteredElementCollector(doc)
               .OfClass(DB.FamilySymbol)
               .OfCategory(DB.BuiltInCategory.OST_RoomTags))

if not room_tags:
    forms.alert("Non ci sono Room Tags caricati nel progetto.", exitscript=True)

room_tags.sort(key=lambda x: x.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString())

room_tags_list = []
tag_dict = {}

for tag in room_tags:
    family_name = tag.FamilyName
    type_name = tag.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    full_name = "{} - {}".format(family_name, type_name)
    room_tags_list.append(full_name)
    tag_dict[full_name] = tag

selected_tag_name = forms.SelectFromList.show(
    room_tags_list,
    multiselect=False,
    button_name='Seleziona Room Tag',
    title='Selezione Room Tag'
)

if not selected_tag_name:
    script.exit()

tag_type = tag_dict[selected_tag_name]

# -------------------------------
# SEZIONE RACCOLTA ROOMS
# -------------------------------
rooms_collector = FilteredElementCollector(link_doc) \
                    .OfCategory(BuiltInCategory.OST_Rooms) \
                    .WhereElementIsNotElementType() \
                    .ToElements()

# Filtra solo le stanze posizionate
placed_rooms = [r for r in rooms_collector if r.Area > 0]

print("Modello linkato: {}".format(selected_link_name))
print("Rooms totali: {} (di cui {} posizionate)".format(len(rooms_collector), len(placed_rooms)))
print("-" * 80)

# -------------------------------
# SEZIONE OPERAZIONI PRINCIPALI
# -------------------------------
output = script.get_output()

# Attiva il tipo di tag se necessario
if not tag_type.IsActive:
    t_activate = Transaction(doc, "Activate Tag Type")
    t_activate.Start()
    tag_type.Activate()
    doc.Regenerate()
    t_activate.Commit()

# Transazione principale
t = Transaction(doc, "Tag Rooms from Linked Model")
t.Start()

try:
    total_tags = 0
    total_duplicates = 0
    total_out_of_range = 0
    link_transform = selected_link.GetTotalTransform()
    
    # Progress bar
    with forms.ProgressBar(title='Tagging Rooms...', cancellable=True) as pb:
        step = 0
        total_steps = len(views_selected) * len(placed_rooms)
        
        for view in views_selected:
            view_tags = 0
            skipped_rooms = 0
            duplicate_tags = 0
            out_of_range = 0
            
            print("\nElaborazione vista: {}".format(view.Name))
            
            # Ottieni informazioni del view range
            view_range_info = get_view_range_volume(view)
            if view_range_info:
                print("  View Range - Top: {:.2f}, Bottom: {:.2f}, Cut: {:.2f}".format(
                    view_range_info['top'], view_range_info['bottom'], view_range_info['cut']))
            
            # Ottieni il livello della vista
            view_level = view.GenLevel
            if not view_level:
                print("  ⚠ Vista senza livello associato, saltata")
                continue
            
            for room in placed_rooms:
                # Update progress
                if pb.cancelled:
                    t.RollBack()
                    script.exit()
                
                pb.update_progress(step, total_steps)
                step += 1
                
                try:
                    # Verifica il livello della stanza
                    room_level = link_doc.GetElement(room.LevelId)
                    if not room_level:
                        continue
                    
                    # NUOVO: Ottieni il punto 3D di inserimento a metà altezza
                    room_point_3d = get_room_insertion_point_3d(room, link_doc)
                    if not room_point_3d:
                        skipped_rooms += 1
                        continue
                    
                    # NUOVO: Verifica se il punto è nel volume del view range
                    is_visible, visibility_reason = is_point_in_view_range_volume(
                        room_point_3d, view, link_transform)
                    
                    if not is_visible:
                        out_of_range += 1
                        if "Fuori view range" in visibility_reason:
                            # Print dettagliato solo per debug
                            pass
                        continue
                    
                    # Verifica se esiste già un tag per questa stanza
                    if check_existing_room_tag(room, view, selected_link, doc):
                        duplicate_tags += 1
                        total_duplicates += 1
                        continue
                    
                    # Trasforma il punto per il posizionamento del tag
                    transformed_point = link_transform.OfPoint(room_point_3d)
                    
                    # Crea il punto UV per il tag (proiezione sul piano della vista)
                    uv_point = UV(transformed_point.X, transformed_point.Y)
                    
                    # Crea il tag
                    try:
                        # Metodo per Revit 2020+
                        linked_elem_id = LinkElementId(selected_link.Id, room.Id)
                        new_tag = doc.Create.NewRoomTag(linked_elem_id, uv_point, view.Id)
                        
                        if new_tag:
                            # Cambia il tipo di tag se necessario
                            if new_tag.GetTypeId() != tag_type.Id:
                                new_tag.ChangeTypeId(tag_type.Id)
                            
                            view_tags += 1
                            total_tags += 1
                    
                    except Exception as tag_error:
                        # Prova metodo alternativo per versioni precedenti
                        try:
                            new_tag = doc.Create.NewRoomTag(room, uv_point, view)
                            if new_tag:
                                new_tag.ChangeTypeId(tag_type.Id)
                                view_tags += 1
                                total_tags += 1
                        except:
                            print("  ⚠ Impossibile creare tag per room ID: {}".format(room.Id))
                
                except Exception as e:
                    print("  ⚠ Errore room {}: {}".format(room.Id, str(e)))
            
            total_out_of_range += out_of_range
            print("  ✓ Tags creati: {} | Fuori view range: {} | Tag duplicati saltati: {}".format(
                view_tags, out_of_range, duplicate_tags))
    
    t.Commit()
    
    # Riepilogo finale
    print("\n" + "=" * 80)
    print("RIEPILOGO OPERAZIONE")
    print("=" * 80)
    print("✓ Tags totali creati: {}".format(total_tags))
    print("✓ Tags duplicati saltati: {}".format(total_duplicates))
    print("✓ Rooms fuori view range: {}".format(total_out_of_range))
    print("✓ Viste elaborate: {}".format(len(views_selected)))
    print("✓ Room Tag utilizzato: {}".format(selected_tag_name))
    print("METODO: Controllo 3D con punto inserimento a metà altezza")
    
    # Mostra riepilogo in dialog
    forms.alert("Operazione completata!\n\n"
                "Tags creati: {}\n"
                "Tags duplicati saltati: {}\n"
                "Rooms fuori view range: {}\n"
                "Viste elaborate: {}\n\n"
                "Metodo: Controllo 3D preciso".format(
                    total_tags, total_duplicates, total_out_of_range, len(views_selected)),
                title="Tag Rooms - Completato")

except Exception as e:
    t.RollBack()
    print("\n⚠ ERRORE CRITICO: {}".format(str(e)))
    forms.alert("Errore durante l'operazione:\n\n{}".format(str(e)),
                title="Errore",
                exitscript=True)