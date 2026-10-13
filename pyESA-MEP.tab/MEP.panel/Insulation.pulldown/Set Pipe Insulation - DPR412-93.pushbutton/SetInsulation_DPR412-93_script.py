# -*- coding: utf-8 -*-
__title__ = 'Set Pipe Insulation\nDPR 412/93'
__author__ = 'Andrea Patti'
__doc__ = """Version = 1.0
________________________________________________________________
Date    = 16.05.2025
Aggiunge o modifica l'isolamento delle tubazioni e dei raccordi selezionati in base alla visibilità e alla zona termica.
________________________________________________________________
Author(s):
Andrea Patti
"""

import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import *
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, DB, UI
from pyrevit import forms
import System
from System.Windows.Forms import Form, Label, ComboBox, Button, CheckBox
from System.Drawing import Size, Point

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# Tabella degli spessori di isolamento in base al diametro e alla zona termica
tabella_isolanti = [
    {"Diameter": 10, "Interno": 6, "Non riscaldato": 13, "Interrato": 13, "Esterno": 25},
    {"Diameter": 15, "Interno": 9, "Non riscaldato": 13, "Interrato": 13, "Esterno": 25},
    {"Diameter": 20, "Interno": 9, "Non riscaldato": 19, "Interrato": 19, "Esterno": 32},
    {"Diameter": 25, "Interno": 9, "Non riscaldato": 19, "Interrato": 19, "Esterno": 32},
    {"Diameter": 32, "Interno": 9, "Non riscaldato": 19, "Interrato": 19, "Esterno": 32},
    {"Diameter": 40, "Interno": 13, "Non riscaldato": 25, "Interrato": 25, "Esterno": 40},
    {"Diameter": 50, "Interno": 13, "Non riscaldato": 25, "Interrato": 25, "Esterno": 40},
    {"Diameter": 65, "Interno": 19, "Non riscaldato": 25, "Interrato": 25, "Esterno": 50},
    {"Diameter": 80, "Interno": 19, "Non riscaldato": 32, "Interrato": 32, "Esterno": 55},
    {"Diameter": 100, "Interno": 19, "Non riscaldato": 32, "Interrato": 32, "Esterno": 60},
    {"Diameter": 125, "Interno": 19, "Non riscaldato": 32, "Interrato": 32, "Esterno": 60},
    {"Diameter": 150, "Interno": 19, "Non riscaldato": 32, "Interrato": 32, "Esterno": 60},
    {"Diameter": 200, "Interno": 19, "Non riscaldato": 32, "Interrato": 32, "Esterno": 60},
    {"Diameter": 250, "Interno": 19, "Non riscaldato": 32, "Interrato": 32, "Esterno": 60}
]

# Zone termiche valide dalla tabella
zone_termiche_valide = ["Interno", "Non riscaldato", "Interrato", "Esterno"]

# Funzione per ottenere tutti i tipi di isolamento disponibili
def get_insulation_types():
    # Per tubazioni
    pipe_insulation_types = FilteredElementCollector(doc)\
                           .OfClass(PipeInsulationType)\
                           .ToElements()
    
    # Utilizziamo il metodo corretto per ottenere il nome del tipo
    result = {}
    for insulation_type in pipe_insulation_types:
        try:
            # Ottieni l'elemento di tipo
            type_name = Element.Name.GetValue(insulation_type)
            if not type_name:
                # Prova con il parametro di tipo
                type_param = insulation_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                if type_param:
                    type_name = type_param.AsString()
                else:
                    type_name = "Tipo Isolamento " + insulation_type.Id.ToString()
            
            result[type_name] = insulation_type
        except Exception as e:
            print("Errore nel recuperare il nome del tipo di isolamento: {}".format(str(e)))
    
    return result

# Funzione per verificare se una tubazione ha già un isolamento
def get_pipe_insulation(pipe):
    try:
        # Cerca eventuali elementi di isolamento associati alla tubazione
        collector = FilteredElementCollector(doc).OfClass(PipeInsulation)
        for insulation in collector:
            if insulation.HostElementId == pipe.Id:
                return insulation
    except Exception as e:
        print("Errore nel verificare l'isolamento: {}".format(str(e)))
    
    return None

# Funzione per verificare se un raccordo ha già un isolamento
def get_fitting_insulation(fitting):
    try:
        # Cerca eventuali elementi di isolamento associati al raccordo
        collector = FilteredElementCollector(doc).OfClass(PipeInsulation)
        for insulation in collector:
            if insulation.HostElementId == fitting.Id:
                return insulation
    except Exception as e:
        print("Errore nel verificare l'isolamento del raccordo: {}".format(str(e)))
    
    return None

# Funzione per modificare l'isolamento esistente
def modify_existing_insulation(insulation, thickness_feet, insulation_type_id):
    success = False
    error_message = ""
    
    try:
        # Cerca il parametro "Thickness" o equivalente in italiano
        thickness_param = None
        for param in insulation.Parameters:
            param_name = param.Definition.Name.lower()
            if "thickness" in param_name or "spessore" in param_name:
                thickness_param = param
                print("Trovato parametro di spessore: {}".format(param.Definition.Name))
                break
        
        if thickness_param and thickness_param.StorageType == StorageType.Double:
            if thickness_param.IsReadOnly:
                print("Il parametro di spessore è in sola lettura")
            else:
                print("Impostazione del parametro di spessore a: {} piedi".format(thickness_feet))
                thickness_param.Set(thickness_feet)
                success = True
        else:
            print("Parametro di spessore non trovato o non è di tipo Double")
            
        # Prova a cambiare il tipo di isolamento
        try:
            current_type_id = insulation.GetTypeId()
            if current_type_id != insulation_type_id:
                insulation.ChangeTypeId(insulation_type_id)
                print("Tipo di isolamento modificato")
                success = True
        except Exception as type_e:
            print("Errore durante la modifica del tipo: {}".format(str(type_e)))
            error_message += " Errore tipo: " + str(type_e)
            
    except Exception as e:
        error_message = str(e)
        print("Errore durante la modifica dell'isolamento: {}".format(error_message))
    
    return success, error_message

# Ottiene elementi di tubazione e raccordi dalla selezione
def get_pipes_and_fittings_from_selection():
    selection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
    pipes = []
    fittings = []
    
    pipe_fitting_category_id = ElementId(BuiltInCategory.OST_PipeFitting)
    
    for elem in selection:
        if isinstance(elem, Pipe):
            pipes.append(elem)
        elif isinstance(elem, FamilyInstance):
            # Verifica se il FamilyInstance è un raccordo di tubazione
            if elem.Category and elem.Category.Id == pipe_fitting_category_id:
                fittings.append(elem)
    
    return pipes, fittings

# Funzione per verificare se un elemento è "a vista" in base al parametro e_EAN_IsVisible_1
def is_element_visible(element):
    try:
        # Cerca il parametro e_EAN_IsVisible_1
        for param in element.Parameters:
            if param.Definition.Name == "e_EAN_IsVisible_1":
                if param.StorageType == StorageType.Integer:
                    # Per i parametri booleani, 1 = True, 0 = False
                    return param.AsInteger() == 1
                return False  # Se il tipo non è corretto, considera come non visibile
        
        # Se il parametro non esiste, considera come non visibile
        return False
    except Exception as e:
        print("Errore nel verificare la visibilità dell'elemento: {}".format(str(e)))
        return False

# Funzione per ottenere la zona termica di un elemento
def get_thermal_position(element):
    try:
        # Cerca il parametro e_EAN_ThermalPosition_1
        for param in element.Parameters:
            if param.Definition.Name == "e_EAN_ThermalPosition_1" and param.HasValue:
                if param.StorageType == StorageType.String:
                    value = param.AsString()
                    # Verifica che il valore sia tra quelli validi
                    if value in zone_termiche_valide:
                        return value, True
                    else:
                        print("Zona termica '{}' non valida. Zone valide: {}".format(
                            value, ", ".join(zone_termiche_valide)))
                        return value, False
        
        # Se il parametro non esiste o non ha un valore valido
        print("Parametro 'e_EAN_ThermalPosition_1' non trovato o non valido")
        return "", False
    except Exception as e:
        print("Errore nel verificare la zona termica dell'elemento: {}".format(str(e)))
        return "", False

# Funzione per ottenere il diametro di una tubazione o di un raccordo
def get_pipe_diameter(element):
    try:
        diameter_mm = 0
        
        if isinstance(element, Pipe):
            # Per le tubazioni, ottieni il diametro dalla proprietà Diameter
            diameter_feet = element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
            diameter_mm = int(round(diameter_feet * 304.8))  # Converti da piedi a mm e arrotonda
        elif isinstance(element, FamilyInstance):
            # Per i raccordi, cerca il parametro del diametro
            # Prova diversi parametri comuni per il diametro
            diameter_param = None
            for param_name in ["Diametro", "Diameter", "Nominal Diameter", "Diametro nominale"]:
                for param in element.Parameters:
                    if param.Definition.Name == param_name and param.HasValue:
                        diameter_param = param
                        break
                if diameter_param:
                    break
            
            if diameter_param and diameter_param.StorageType == StorageType.Double:
                diameter_feet = diameter_param.AsDouble()
                diameter_mm = int(round(diameter_feet * 304.8))  # Converti da piedi a mm e arrotonda
            else:
                # Prova a ottenere il diametro dal connettore
                if hasattr(element, "MEPModel") and element.MEPModel:
                    connectors = element.MEPModel.ConnectorManager.Connectors
                    if connectors and connectors.Size > 0:
                        for connector in connectors:
                            if connector.Shape == ConnectorProfileType.Round:
                                diameter_feet = connector.Radius * 2
                                diameter_mm = int(round(diameter_feet * 304.8))
                                break
        
        # Se non riesci a trovare il diametro, restituisci 0 e segnala
        if diameter_mm <= 0:
            print("Impossibile determinare il diametro dell'elemento")
            return 0, False
        
        return diameter_mm, True
    except Exception as e:
        print("Errore nel determinare il diametro: {}".format(str(e)))
        return 0, False

# Funzione per ottenere lo spessore dell'isolamento dalla tabella
def get_insulation_thickness(diameter_mm, thermal_position):
    try:
        # Verifica che la posizione termica sia valida
        if thermal_position not in zone_termiche_valide:
            print("Posizione termica '{}' non presente nella tabella".format(thermal_position))
            return 0, False
        
        # Trova la riga con il diametro più vicino
        closest_row = None
        min_diff = float('inf')
        
        for row in tabella_isolanti:
            diff = abs(row["Diameter"] - diameter_mm)
            if diff < min_diff:
                min_diff = diff
                closest_row = row
        
        if closest_row:
            thickness_mm = closest_row[thermal_position]
            return thickness_mm, True
        else:
            print("Nessuna corrispondenza trovata nella tabella per il diametro {} mm".format(diameter_mm))
            return 0, False
    except Exception as e:
        print("Errore nella ricerca dello spessore dalla tabella: {}".format(str(e)))
        return 0, False

# Classe per la finestra di dialogo personalizzata
class IsolationSelectionForm(Form):
    def __init__(self, insulation_types_list):
        self.Text = "Selezione tipi di isolamento"
        self.Size = Size(450, 230)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        
        # Etichette
        lbl_visible = Label()
        lbl_visible.Text = "Tipo isolamento per tubazioni A VISTA:"
        lbl_visible.Location = Point(20, 20)
        lbl_visible.Size = Size(250, 20)
        self.Controls.Add(lbl_visible)
        
        lbl_hidden = Label()
        lbl_hidden.Text = "Tipo isolamento per tubazioni NON A VISTA:"
        lbl_hidden.Location = Point(20, 80)
        lbl_hidden.Size = Size(250, 20)
        self.Controls.Add(lbl_hidden)
        
        # ComboBox per isolamento tubazioni a vista
        self.cb_visible = ComboBox()
        self.cb_visible.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self.cb_visible.Location = Point(20, 40)
        self.cb_visible.Size = Size(400, 30)
        for type_name in insulation_types_list:
            self.cb_visible.Items.Add(type_name)
        if self.cb_visible.Items.Count > 0:
            self.cb_visible.SelectedIndex = 0
        self.Controls.Add(self.cb_visible)
        
        # ComboBox per isolamento tubazioni non a vista
        self.cb_hidden = ComboBox()
        self.cb_hidden.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self.cb_hidden.Location = Point(20, 100)
        self.cb_hidden.Size = Size(400, 30)
        for type_name in insulation_types_list:
            self.cb_hidden.Items.Add(type_name)
        if self.cb_hidden.Items.Count > 0:
            self.cb_hidden.SelectedIndex = 0
        self.Controls.Add(self.cb_hidden)
        
        # Checkbox per spessore manuale
        self.chk_manual_thickness = CheckBox()
        self.chk_manual_thickness.Text = "Specifica manualmente lo spessore dell'isolamento"
        self.chk_manual_thickness.Location = Point(20, 140)
        self.chk_manual_thickness.Size = Size(300, 20)
        self.chk_manual_thickness.Checked = False
        self.chk_manual_thickness.CheckedChanged += self.on_manual_thickness_changed
        self.Controls.Add(self.chk_manual_thickness)
        
        # ComboBox per spessore isolamento (inizialmente nascosto)
        self.cb_thickness = ComboBox()
        self.cb_thickness.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self.cb_thickness.Location = Point(320, 140)
        self.cb_thickness.Size = Size(100, 30)
        self.cb_thickness.Visible = False
        thickness_options = [6, 9, 13, 19, 25, 32, 40, 50, 55, 60]
        for thickness in thickness_options:
            self.cb_thickness.Items.Add(thickness)
        self.cb_thickness.SelectedIndex = 2  # Imposta 13mm come predefinito
        self.Controls.Add(self.cb_thickness)
        
        # Pulsante OK
        self.btn_ok = Button()
        self.btn_ok.Text = "OK"
        self.btn_ok.DialogResult = System.Windows.Forms.DialogResult.OK
        self.btn_ok.Location = Point(185, 170)
        self.btn_ok.Size = Size(80, 30)
        self.Controls.Add(self.btn_ok)
        
        # Pulsante Annulla
        self.btn_cancel = Button()
        self.btn_cancel.Text = "Annulla"
        self.btn_cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        self.btn_cancel.Location = Point(275, 170)
        self.btn_cancel.Size = Size(80, 30)
        self.Controls.Add(self.btn_cancel)
        
        self.AcceptButton = self.btn_ok
        self.CancelButton = self.btn_cancel
    
    def on_manual_thickness_changed(self, sender, args):
        self.cb_thickness.Visible = self.chk_manual_thickness.Checked
    
    def get_selected_values(self):
        return {
            "visible_type": self.cb_visible.SelectedItem,
            "hidden_type": self.cb_hidden.SelectedItem,
            "manual_thickness": self.chk_manual_thickness.Checked,
            "thickness": self.cb_thickness.SelectedItem if self.chk_manual_thickness.Checked else None
        }

# Ottieni tubazioni e raccordi selezionati
pipes, fittings = get_pipes_and_fittings_from_selection()

if not pipes and not fittings:
    forms.alert('Seleziona almeno una tubazione o un raccordo prima di eseguire lo script.', exitscript=True)

# Ottieni i tipi di isolamento disponibili
insulation_types_dict = get_insulation_types()

if not insulation_types_dict:
    forms.alert('Non ci sono tipi di isolamento disponibili nel progetto.', exitscript=True)

# Mostra la finestra di dialogo personalizzata
insulation_types_list = sorted(insulation_types_dict.keys())
dialog = IsolationSelectionForm(insulation_types_list)
result = dialog.ShowDialog()

if result != System.Windows.Forms.DialogResult.OK:
    dialog.Dispose()
    forms.alert('Operazione annullata.', exitscript=True)

# Ottieni i valori selezionati
selected_values = dialog.get_selected_values()
dialog.Dispose()

# Recupera i tipi di isolamento selezionati
visible_type_name = selected_values["visible_type"]
hidden_type_name = selected_values["hidden_type"]
manual_thickness = selected_values["manual_thickness"]
manual_thickness_mm = selected_values["thickness"] if manual_thickness else None

visible_insulation_type = insulation_types_dict[visible_type_name]
hidden_insulation_type = insulation_types_dict[hidden_type_name]

# Preparazione per il conteggio delle operazioni
pipes_added = 0
pipes_modified = 0
pipes_skipped = 0
pipes_error = 0

fittings_added = 0
fittings_modified = 0
fittings_skipped = 0
fittings_error = 0

# Esegui le operazioni in una transazione
with revit.Transaction("Gestione isolamento tubazioni e raccordi"):
    # Elabora le tubazioni
    for pipe_index in range(len(pipes)):
        pipe = pipes[pipe_index]
        try:
            pipe_id_str = pipe.Id.ToString()
            
            # Determina se la tubazione è "a vista"
            is_visible = is_element_visible(pipe)
            
            # Scegli il tipo di isolamento appropriato
            insulation_type_id = visible_insulation_type.Id if is_visible else hidden_insulation_type.Id
            
            # Determina lo spessore dell'isolamento
            if manual_thickness:
                # Usa lo spessore specificato manualmente
                thickness_mm = manual_thickness_mm
                valid_thickness = True
            else:
                # Ottieni il diametro della tubazione
                diameter_mm, valid_diameter = get_pipe_diameter(pipe)
                
                if not valid_diameter:
                    print("Tubazione {}: Diametro non valido, tubazione saltata".format(pipe_id_str))
                    pipes_skipped += 1
                    continue
                
                # Ottieni la zona termica della tubazione
                thermal_position, valid_position = get_thermal_position(pipe)
                
                if not valid_position:
                    print("Tubazione {}: Zona termica '{}' non valida, tubazione saltata".format(
                        pipe_id_str, thermal_position))
                    pipes_skipped += 1
                    continue
                
                # Cerca lo spessore appropriato nella tabella
                thickness_mm, valid_thickness = get_insulation_thickness(diameter_mm, thermal_position)
                
                if not valid_thickness:
                    print("Tubazione {}: Spessore non trovato nella tabella, tubazione saltata".format(pipe_id_str))
                    pipes_skipped += 1
                    continue
                
                print("Tubazione {}: Diametro = {} mm, Zona termica = {}, Spessore isolamento = {} mm".format(
                    pipe_id_str, diameter_mm, thermal_position, thickness_mm))
            
            # Converti mm in piedi (unità interne di Revit)
            thickness_feet = thickness_mm / 304.8
            
            # Verifica se la tubazione ha già un isolamento
            existing_insulation = get_pipe_insulation(pipe)
            
            if not existing_insulation:
                # Se non ha isolamento, aggiungilo
                try:
                    PipeInsulation.Create(doc, pipe.Id, insulation_type_id, thickness_feet)
                    pipes_added += 1
                    print("Isolamento aggiunto alla tubazione {} ({}) con spessore {} mm.".format(
                        pipe_id_str, "a vista" if is_visible else "non a vista", thickness_mm))
                except Exception as add_e:
                    print("Errore nell'aggiungere isolamento alla tubazione {}: {}".format(pipe_id_str, str(add_e)))
                    pipes_error += 1
            else:
                # Se ha già un isolamento, prova a modificarlo
                success, error = modify_existing_insulation(existing_insulation, thickness_feet, insulation_type_id)
                
                if success:
                    pipes_modified += 1
                    print("Isolamento modificato per la tubazione {} ({}) con spessore {} mm.".format(
                        pipe_id_str, "a vista" if is_visible else "non a vista", thickness_mm))
                else:
                    # Se la modifica fallisce, prova l'approccio elimina-e-ricrea
                    try:
                        print("Tentativo di eliminare e ricreare l'isolamento per la tubazione {}".format(pipe_id_str))
                        doc.Delete(existing_insulation.Id)
                        PipeInsulation.Create(doc, pipe.Id, insulation_type_id, thickness_feet)
                        pipes_modified += 1
                        print("Isolamento ricreato per la tubazione {} ({}) con spessore {} mm.".format(
                            pipe_id_str, "a vista" if is_visible else "non a vista", thickness_mm))
                    except Exception as recreate_e:
                        print("Errore nella ricreazione dell'isolamento: {}".format(str(recreate_e)))
                        pipes_error += 1
        except Exception as e:
            print("Errore generico con la tubazione {}: {}".format(pipe.Id, str(e)))
            pipes_error += 1
    
    # Elabora i raccordi
    for fitting_index in range(len(fittings)):
        fitting = fittings[fitting_index]
        try:
            fitting_id_str = fitting.Id.ToString()
            
            # Determina se il raccordo è "a vista"
            is_visible = is_element_visible(fitting)
            
            # Scegli il tipo di isolamento appropriato
            insulation_type_id = visible_insulation_type.Id if is_visible else hidden_insulation_type.Id
            
            # Determina lo spessore dell'isolamento
            if manual_thickness:
                # Usa lo spessore specificato manualmente
                thickness_mm = manual_thickness_mm
                valid_thickness = True
            else:
                # Ottieni il diametro del raccordo
                diameter_mm, valid_diameter = get_pipe_diameter(fitting)
                
                if not valid_diameter:
                    print("Raccordo {}: Diametro non valido, raccordo saltato".format(fitting_id_str))
                    fittings_skipped += 1
                    continue
                
                # Ottieni la zona termica del raccordo
                thermal_position, valid_position = get_thermal_position(fitting)
                
                if not valid_position:
                    print("Raccordo {}: Zona termica '{}' non valida, raccordo saltato".format(
                        fitting_id_str, thermal_position))
                    fittings_skipped += 1
                    continue
                
                # Cerca lo spessore appropriato nella tabella
                thickness_mm, valid_thickness = get_insulation_thickness(diameter_mm, thermal_position)
                
                if not valid_thickness:
                    print("Raccordo {}: Spessore non trovato nella tabella, raccordo saltato".format(fitting_id_str))
                    fittings_skipped += 1
                    continue
                
                print("Raccordo {}: Diametro = {} mm, Zona termica = {}, Spessore isolamento = {} mm".format(
                    fitting_id_str, diameter_mm, thermal_position, thickness_mm))
            
            # Converti mm in piedi (unità interne di Revit)
            thickness_feet = thickness_mm / 304.8
            
            # Verifica se il raccordo ha già un isolamento
            existing_insulation = get_fitting_insulation(fitting)
            
            if not existing_insulation:
                # Se non ha isolamento, aggiungilo
                try:
                    PipeInsulation.Create(doc, fitting.Id, insulation_type_id, thickness_feet)
                    fittings_added += 1
                    print("Isolamento aggiunto al raccordo {} ({}) con spessore {} mm.".format(
                        fitting_id_str, "a vista" if is_visible else "non a vista", thickness_mm))
                except Exception as add_e:
                    print("Errore nell'aggiungere isolamento al raccordo {}: {}".format(fitting_id_str, str(add_e)))
                    fittings_error += 1
            else:
                # Se ha già un isolamento, prova a modificarlo
                success, error = modify_existing_insulation(existing_insulation, thickness_feet, insulation_type_id)
                
                if success:
                    fittings_modified += 1
                    print("Isolamento modificato per il raccordo {} ({}) con spessore {} mm.".format(
                        fitting_id_str, "a vista" if is_visible else "non a vista", thickness_mm))
                else:
                    # Se la modifica fallisce, prova l'approccio elimina-e-ricrea
                    try:
                        print("Tentativo di eliminare e ricreare l'isolamento per il raccordo {}".format(fitting_id_str))
                        doc.Delete(existing_insulation.Id)
                        PipeInsulation.Create(doc, fitting.Id, insulation_type_id, thickness_feet)
                        fittings_modified += 1
                        print("Isolamento ricreato per il raccordo {} ({}) con spessore {} mm.".format(
                            fitting_id_str, "a vista" if is_visible else "non a vista", thickness_mm))
                    except Exception as recreate_e:
                        print("Errore nella ricreazione dell'isolamento: {}".format(str(recreate_e)))
                        fittings_error += 1
        except Exception as e:
            print("Errore generico con il raccordo {}: {}".format(fitting.Id, str(e)))
            fittings_error += 1

# Mostra un messaggio di riepilogo
total_processed = pipes_added + pipes_modified + fittings_added + fittings_modified
total_skipped = pipes_skipped + fittings_skipped
summary_message = "Operazione completata:\n\n"
summary_message += "TUBAZIONI:\n"
summary_message += "- {} tubazioni isolate\n".format(pipes_added)
summary_message += "- {} isolamenti modificati\n".format(pipes_modified)
if pipes_skipped > 0:
    summary_message += "- {} tubazioni saltate\n".format(pipes_skipped)
if pipes_error > 0:
    summary_message += "- {} operazioni non riuscite\n".format(pipes_error)

summary_message += "\nRACCORDI:\n"
summary_message += "- {} raccordi isolati\n".format(fittings_added)
summary_message += "- {} isolamenti modificati\n".format(fittings_modified)
if fittings_skipped > 0:
    summary_message += "- {} raccordi saltati\n".format(fittings_skipped)
if fittings_error > 0:
    summary_message += "- {} operazioni non riuscite\n".format(fittings_error)

summary_message += "\nTotale: {} elementi elaborati, {} elementi saltati".format(
    total_processed, total_skipped)

forms.alert(summary_message)