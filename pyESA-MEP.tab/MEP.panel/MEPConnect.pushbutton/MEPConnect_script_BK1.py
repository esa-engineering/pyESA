# -*- coding: utf-8 -*-
__title__ = 'MEP\nConnect'
__doc__ = """Version = 1.0
Date = 23.05.2025
________________________________________________________________
Universal connector for Pipes, Ducts, Cable Trays and Conduits
Slope: Default 90 degrees - OPTIMIZED VERSION

Instruction:
- Select any element (Pipe, Duct, Cable Tray or Conduit)
- The script will automatically detect the category and proceed
- Select elements in order to connect them
- Press ESC to end

Differently from the Trim/Extend, it can also connect
elements at different elevation and aligned each other.
When the selected elements are not aligned, the tool will
create a new element to connect the two closest extremities.
_______________________________________________________________
Author: 
ESA Engineering Srl
"""

import math

import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct
from Autodesk.Revit.DB.Electrical import CableTray, Conduit

from pyrevit import revit, DB
from pyrevit import PyRevitException, PyRevitIOError
from pyrevit import forms

doc = revit.doc

### UTILITY FUNCTIONS

# Cache per evitare ripetute conversioni
ELEMENT_TYPE_CACHE = {}
CATEGORY_MAPPINGS = {
    "PIPE": ("Pipe", DB.BuiltInCategory.OST_PipeCurves),
    "DUCT": ("Duct", DB.BuiltInCategory.OST_DuctCurves),
    "CABLETRAY": ("Cable Tray", DB.BuiltInCategory.OST_CableTray),
    "CONDUIT": ("Conduit", DB.BuiltInCategory.OST_Conduit)
}

def get_element_category(element):
    """Determina la categoria dell'elemento selezionato con cache"""
    element_type = type(element)
    
    if element_type not in ELEMENT_TYPE_CACHE:
        if isinstance(element, Pipe):
            ELEMENT_TYPE_CACHE[element_type] = "PIPE"
        elif isinstance(element, Duct):
            ELEMENT_TYPE_CACHE[element_type] = "DUCT"
        elif isinstance(element, CableTray):
            ELEMENT_TYPE_CACHE[element_type] = "CABLETRAY"
        elif isinstance(element, Conduit):
            ELEMENT_TYPE_CACHE[element_type] = "CONDUIT"
        else:
            ELEMENT_TYPE_CACHE[element_type] = None
    
    return ELEMENT_TYPE_CACHE[element_type]

def get_category_info(category):
    """Restituisce nome e BuiltInCategory in una sola chiamata"""
    return CATEGORY_MAPPINGS.get(category, ("Unknown", None))

### COMMON FUNCTIONS

def placeElbow(p1, p2, category=None):
    """Versione ottimizzata con early exit e meno allocazioni"""
    try:
        # Usa generator expression invece di list comprehension per risparmiare memoria
        connectors_p1 = p1.ConnectorManager.Connectors
        connectors_p2 = p2.ConnectorManager.Connectors
        
        # Trova la distanza minima senza creare liste intermedie
        min_dist = float('inf')
        best_c1, best_c2 = None, None
        
        for c1 in connectors_p1:
            for c2 in connectors_p2:
                dist = c1.Origin.DistanceTo(c2.Origin)
                if dist < min_dist:
                    min_dist = dist
                    best_c1, best_c2 = c1, c2
        
        # Per le cable trays e conduit, usa un approccio diverso
        if category in ["CABLETRAY", "CONDUIT"]:
            try:
                return doc.Create.NewTransitionFitting(best_c1, best_c2)
            except:
                try:
                    return doc.Create.NewElbowFitting(best_c1, best_c2)
                except:
                    best_c1.ConnectTo(best_c2)
                    return None
        else:
            return doc.Create.NewElbowFitting(best_c1, best_c2)
    except:
        # Fallback ottimizzato
        try:
            best_c1.ConnectTo(best_c2)
            return None
        except:
            return None

def get_connectingLine(ln1, ln2):
    """Versione ottimizzata con meno allocazioni"""
    pts1_0, pts1_1 = ln1.GetEndPoint(0), ln1.GetEndPoint(1)
    pts2_0, pts2_1 = ln2.GetEndPoint(0), ln2.GetEndPoint(1)
    
    # Calcola tutte le distanze senza creare liste intermedie
    distances = [
        (pts1_0.DistanceTo(pts2_0), pts1_0, pts2_0),
        (pts1_0.DistanceTo(pts2_1), pts1_0, pts2_1),
        (pts1_1.DistanceTo(pts2_0), pts1_1, pts2_0),
        (pts1_1.DistanceTo(pts2_1), pts1_1, pts2_1)
    ]
    
    # Trova il minimo
    min_dist, pt1, pt2_orig = min(distances)
    pt2 = DB.XYZ(pt2_orig.X, pt2_orig.Y, pt1.Z)
    return DB.Line.CreateBound(pt1, pt2)

def testFloat(s):
    """Versione ottimizzata con early return"""
    try:
        return float(s)
    except:
        return None

### OPTIMIZED CONNECTOR FUNCTIONS

def get_closest_unused_connectors(element1, element2):
    """Funzione ottimizzata per trovare i connettori non utilizzati più vicini"""
    unused1 = element1.ConnectorManager.UnusedConnectors
    unused2 = element2.ConnectorManager.UnusedConnectors
    
    min_dist = float('inf')
    best_conn1, best_conn2 = None, None
    
    for u1 in unused1:
        for u2 in unused2:
            dist = u1.Origin.DistanceTo(u2.Origin)
            if dist < min_dist:
                min_dist = dist
                best_conn1, best_conn2 = u1, u2
    
    return best_conn1, best_conn2

def get_closest_unused_connector(element, point):
    """Trova il connettore non utilizzato più vicino a un punto"""
    unused = element.ConnectorManager.UnusedConnectors
    min_dist = float('inf')
    best_conn = None
    
    for u in unused:
        dist = u.Origin.DistanceTo(point)
        if dist < min_dist:
            min_dist = dist
            best_conn = u
    
    return best_conn

### PIPE SPECIFIC FUNCTIONS

def connectPipesWithPipe(doc, p1, p2):
    """Versione ottimizzata"""
    conn01, conn02 = get_closest_unused_connectors(p1, p2)
    levId = p1.Parameter[DB.BuiltInParameter.RBS_START_LEVEL_PARAM].AsElementId()
    return Pipe.Create(doc, p1.GetTypeId(), levId, conn01, conn02)

def createPipe(doc, p1, pt1, pt2):
    """Versione ottimizzata"""
    conn01 = get_closest_unused_connector(p1, pt1)
    levId = p1.Parameter[DB.BuiltInParameter.RBS_START_LEVEL_PARAM].AsElementId()
    return Pipe.Create(doc, p1.GetTypeId(), levId, conn01, pt2)

### DUCT SPECIFIC FUNCTIONS

def connectDuctsWithDuct(doc, p1, p2):
    """Versione ottimizzata"""
    conn01, conn02 = get_closest_unused_connectors(p1, p2)
    levId = p1.Parameter[DB.BuiltInParameter.RBS_START_LEVEL_PARAM].AsElementId()
    return Duct.Create(doc, p1.GetTypeId(), levId, conn01, conn02)

def createDuct(doc, p1, pt1, pt2):
    """Versione ottimizzata"""
    conn01 = get_closest_unused_connector(p1, pt1)
    levId = p1.Parameter[DB.BuiltInParameter.RBS_START_LEVEL_PARAM].AsElementId()
    return Duct.Create(doc, p1.GetTypeId(), levId, conn01, pt2)

### CABLE TRAY SPECIFIC FUNCTIONS

def copyTrayDimensions(sourceTray, targetTray, slopeDef=None):
    """Versione ottimizzata con early exit"""
    try:
        width_param = sourceTray.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
        height_param = sourceTray.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
        
        if not (width_param and height_param):
            return
            
        source_width = width_param.AsDouble()
        source_height = height_param.AsDouble()
        
        target_width = targetTray.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
        target_height = targetTray.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
        
        # Se slope è 90 gradi (verticale), scambia width e height
        if slopeDef is not None and slopeDef == 0:
            if target_width and not target_width.IsReadOnly:
                target_width.Set(source_height)
            if target_height and not target_height.IsReadOnly:
                target_height.Set(source_width)
        else:
            if target_width and not target_width.IsReadOnly:
                target_width.Set(source_width)
            if target_height and not target_height.IsReadOnly:
                target_height.Set(source_height)
                    
    except Exception as e:
        print("Errore nella copia delle dimensioni: {}".format(str(e)))

def connectTraysWithTray(doc, p1, p2, slopeDef=None):
    """Versione ottimizzata"""
    conn01, conn02 = get_closest_unused_connectors(p1, p2)
    levId = p1.Parameter[DB.BuiltInParameter.RBS_START_LEVEL_PARAM].AsElementId()
    newTray = CableTray.Create(doc, p1.GetTypeId(), conn01.Origin, conn02.Origin, levId)
    copyTrayDimensions(p1, newTray, slopeDef)
    return newTray

def createTray(doc, p1, pt1, pt2, slopeDef=None):
    """Versione ottimizzata"""
    conn01 = get_closest_unused_connector(p1, pt1)
    levId = p1.Parameter[DB.BuiltInParameter.RBS_START_LEVEL_PARAM].AsElementId()
    newTray = CableTray.Create(doc, p1.GetTypeId(), conn01.Origin, pt2, levId)
    copyTrayDimensions(p1, newTray, slopeDef)
    return newTray

### CONDUIT SPECIFIC FUNCTIONS

def copyConduitDimensions(sourceConduit, targetConduit):
    """Copia le dimensioni del conduit (diametro)"""
    try:
        # Per i conduit usiamo il parametro del diametro
        diameter_param = sourceConduit.get_Parameter(DB.BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)
        
        if not diameter_param:
            # Prova con il parametro alternativo per il diametro esterno
            diameter_param = sourceConduit.get_Parameter(DB.BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM)
        
        if diameter_param:
            source_diameter = diameter_param.AsDouble()
            
            # Imposta il diametro sul conduit di destinazione
            target_diameter = targetConduit.get_Parameter(DB.BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)
            if not target_diameter:
                target_diameter = targetConduit.get_Parameter(DB.BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM)
            
            if target_diameter and not target_diameter.IsReadOnly:
                target_diameter.Set(source_diameter)
                    
    except Exception as e:
        print("Errore nella copia del diametro del conduit: {}".format(str(e)))

def connectConduitsWithConduit(doc, p1, p2, slopeDef=None):
    """Connette due conduit con un nuovo conduit"""
    conn01, conn02 = get_closest_unused_connectors(p1, p2)
    levId = p1.Parameter[DB.BuiltInParameter.RBS_START_LEVEL_PARAM].AsElementId()
    newConduit = Conduit.Create(doc, p1.GetTypeId(), conn01.Origin, conn02.Origin, levId)
    copyConduitDimensions(p1, newConduit)
    return newConduit

def createConduit(doc, p1, pt1, pt2, slopeDef=None):
    """Crea un nuovo conduit"""
    conn01 = get_closest_unused_connector(p1, pt1)
    levId = p1.Parameter[DB.BuiltInParameter.RBS_START_LEVEL_PARAM].AsElementId()
    newConduit = Conduit.Create(doc, p1.GetTypeId(), conn01.Origin, pt2, levId)
    copyConduitDimensions(p1, newConduit)
    return newConduit

### MAIN CONNECTION LOGIC

# Dizionario delle funzioni per evitare if multipli
CONNECT_FUNCTIONS = {
    "PIPE": (connectPipesWithPipe, createPipe),
    "DUCT": (connectDuctsWithDuct, createDuct),
    "CABLETRAY": (connectTraysWithTray, createTray),
    "CONDUIT": (connectConduitsWithConduit, createConduit)
}

def process_connection(element1, element2, category, slopeDef, nr):
    """Versione ottimizzata della logica di connessione"""
    
    ln1 = element1.Location.Curve
    ln2 = element2.Location.Curve
    baseNewLine = get_connectingLine(ln1, ln2)

    # DEFINE VALUE FOR INCLINATION
    deltaH = math.fabs(ln2.GetEndPoint(0).Z - baseNewLine.GetEndPoint(1).Z)
    Co = math.tan(math.radians(slopeDef)) * deltaH
    oppDir = baseNewLine.Direction.Multiply(Co)

    category_name, _ = get_category_info(category)
    connect_func, create_func = CONNECT_FUNCTIONS[category]

    with revit.Transaction('MEP Connect {} {}'.format(category_name, nr)):
        # CHECK ALIGNMENT
        if (ln1.Direction.IsAlmostEqualTo(baseNewLine.Direction) and 
            ln1.Direction.IsAlmostEqualTo(ln2.Direction) and 
            deltaH == 0):
            
            baseNewLine_new = DB.Line.CreateBound(ln1.GetEndPoint(0), ln2.GetEndPoint(1))
            element1.Location.Curve = baseNewLine_new
            doc.Delete(element2.Id)

        elif ln1.Direction.IsAlmostEqualTo(baseNewLine.Direction):
            if baseNewLine.GetEndPoint(0).IsAlmostEqualTo(ln1.GetEndPoint(0)):
                baseNewLine_new = DB.Line.CreateBound(ln1.GetEndPoint(1),
                            baseNewLine.GetEndPoint(1).Subtract(oppDir))
            else:
                baseNewLine_new = DB.Line.CreateBound(ln1.GetEndPoint(0),
                            baseNewLine.GetEndPoint(1).Subtract(oppDir))
            
            element1.Location.Curve = baseNewLine_new
            
            # Create connecting element
            if category in ["CABLETRAY", "CONDUIT"]:
                element3 = connect_func(doc, element1, element2, slopeDef)
            else:
                element3 = connect_func(doc, element1, element2)
            
            # CONNECT
            placeElbow(element1, element3, category)
            placeElbow(element3, element2, category)

        else:
            # CREATE NEW ELEMENT
            if category in ["CABLETRAY", "CONDUIT"]:
                element3 = create_func(doc, element1,
                        baseNewLine.GetEndPoint(0),
                        baseNewLine.GetEndPoint(1).Subtract(oppDir), slopeDef)
            else:
                element3 = create_func(doc, element1,
                        baseNewLine.GetEndPoint(0),
                        baseNewLine.GetEndPoint(1).Subtract(oppDir))
            
            # CONNECT THEM ALL
            placeElbow(element1, element3, category)
            
            try:
                if category in ["CABLETRAY", "CONDUIT"]:
                    element4 = connect_func(doc, element3, element2, slopeDef)
                else:
                    element4 = connect_func(doc, element3, element2)
                
                placeElbow(element3, element4, category)
                placeElbow(element4, element2, category)
            except:
                # in case element3 is at same elevation of element2
                placeElbow(element3, element2, category)
        
        baseNewLine.Dispose()

### MAIN EXECUTION

def get_slope_selection():
    """Selezione pendenza ottimizzata"""
    slope_options = [30, 45, 60, 90]
    slopeDef = forms.CommandSwitchWindow.show(slope_options, 
                                            message='Select connector slope [degrees]:')
    return (90 - testFloat(slopeDef)) if slopeDef else 0

def process_elements_of_category(category, category_name, slopeDef, nr, first_element=None):
    """Versione ottimizzata del processamento elementi"""
    _, target_category = get_category_info(category)
    
    with forms.ProgressBar(title='Select {}s to connect - press Esc to stop'.format(category_name), 
                          cancellable=True) as pb:
        
        selected_elements = [first_element] if first_element else []
        prog = 50 if first_element else 0
        
        pb.update_progress(prog, 100)
        
        for element in revit.get_picked_elements_by_category(target_category, 
                                                           "Select {} element".format(category_name)):
            prog += 50
            pb.update_progress(prog, 100)
            selected_elements.append(element)
            
            if element is None:
                return nr, True  # Utente ha premuto ESC, termina tutto
            elif prog >= 100:
                prog = 0
                element1, element2 = selected_elements[0], selected_elements[1]
                selected_elements = []
                
                # Verifica categoria con early exit
                if not (get_element_category(element1) == category and 
                       get_element_category(element2) == category):
                    forms.alert("Tutti gli elementi devono essere della stessa categoria: {}".format(category_name))
                    continue
                
                nr += 1
                process_connection(element1, element2, category, slopeDef, nr)
                
                return nr, False  # Continua ma rianalizza categoria
    
    return nr, True  # Fine naturale del loop

try:
    nr = 0
    
    # SELEZIONE DELLA PENDENZA UNA SOLA VOLTA ALL'INIZIO
#    print("Seleziona la pendenza per tutte le connessioni...")
    global_slope = get_slope_selection()
#    print("Pendenza selezionata: {} gradi".format(90 - global_slope))
    
    while True:
        # Selezione elemento per determinare/rianalizzare la categoria
#        print("Seleziona un elemento per determinare la categoria...")
        first_element = revit.pick_element("Seleziona un elemento (Pipe, Duct, Cable Tray o Conduit)")
        
        if not first_element:
#            print("Nessun elemento selezionato. Script terminato.")
            break
        
        # Determina la categoria
        category = get_element_category(first_element)
        
        if not category:
            forms.alert("Elemento non supportato. Seleziona un Pipe, Duct, Cable Tray o Conduit.")
            continue  # Riprova con un altro elemento
        
        category_name, _ = get_category_info(category)
#        print("Categoria rilevata: {}".format(category_name))
        
        # Processa gli elementi di questa categoria usando la pendenza globale
        nr, should_exit = process_elements_of_category(category, category_name, global_slope, nr, first_element)
        
        if should_exit:
            break
        
        # Se arriviamo qui, significa che è stata completata una connessione
        # e dobbiamo rianalizzare la categoria per la prossima operazione
#        print("Operazione completata. Pronto per una nuova selezione...")

except Exception as e:
    forms.alert("Errore durante l'esecuzione: {}".format(str(e)), exitscript=True)