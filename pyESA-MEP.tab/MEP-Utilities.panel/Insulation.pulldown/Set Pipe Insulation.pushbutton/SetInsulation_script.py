# -*- coding: utf-8 -*-
__title__ = 'Set Pipe\nInsulation\nBy System'
__author__ = 'Andrea Patti'
__doc__ = """Version = 2.1
Date = 14.10.2026
________________________________________________________________
Adds or updates the insulation of the pipes and pipe fittings of a
piping system. Thickness comes from an external table, by nominal
diameter and thermal zone, or is a fixed value.
1- Choose the piping system
2- Choose the scope: all elements of the system, or the selected ones
3- Choose the insulation type for exposed and concealed pipes
4- Choose the thickness table file (template: ALT + CLICK on the
   button opens the folder with insulation_table_template.txt)
5- Optionally use a fixed thickness instead of the table
________________________________________________________________
Author(s):
Andrea Patti
"""

import os
import csv
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import *
from pyrevit import revit, DB
from pyrevit import forms, script
import System

import SetPipeInsulation_ui

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

XAML_FILE_NAME = "SetPipeInsulation_form.xaml"
CONFIG_NAME = "ESA_SetPipeInsulation"
DEFAULTS = {"system": "", "selected_only": False, "visible_type": "", "hidden_type": "",
            "table_path": "", "manual": False, "thickness": 13}

# Zone termiche valide (intestazioni della tabella e valori di e_EAN_ThermalPosition_1)
zone_termiche_valide = ["Interno", "Non riscaldato", "Interrato", "Esterno"]


def get_element_id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


def to_bool(value):
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ("true", "1", "yes")


def load_settings(config):
    settings = dict(DEFAULTS)
    try:
        for key in ("system", "visible_type", "hidden_type", "table_path"):
            settings[key] = config.get_option(key, DEFAULTS[key]) or ""
        settings["selected_only"] = to_bool(config.get_option("selected_only", DEFAULTS["selected_only"]))
        settings["manual"] = to_bool(config.get_option("manual", DEFAULTS["manual"]))
        settings["thickness"] = int(config.get_option("thickness", DEFAULTS["thickness"]))
    except Exception:
        pass
    return settings


def save_settings(config, values):
    try:
        for key in DEFAULTS:
            setattr(config, key, values[key])
        script.save_config()
    except Exception:
        pass


# Funzione per ottenere i tipi di sistema di tubazioni dagli elementi nel modello
def get_piping_system_types():
    """Ottiene i tipi di sistema di tubazioni dagli elementi nel modello"""
    # Dizionario per memorizzare i valori dei tipi di sistema
    system_values = {}

    # Ottieni tutte le tubazioni
    try:
        pipes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
        for pipe in pipes:
            try:
                # Ottieni il parametro System Type
                param = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
                if param and param.HasValue:
                    # Ottieni il valore del parametro come stringa
                    param_string = param.AsValueString()
                    if param_string and param_string not in system_values:
                        system_values[param_string] = param_string
            except Exception as e:
                print("Error reading the system type of a pipe: " + str(e))
    except Exception as e:
        print("Error collecting the pipes: " + str(e))

    # Se non abbiamo trovato nulla, usa i valori di PipeSystemType come fallback
    if not system_values:
        print("No system type found on the pipes: using the PipeSystemType values instead")
        try:
            for value in System.Enum.GetValues(PipeSystemType):
                name = value.ToString()
                if name != "UndefinedSystemType" and name != "Fitting" and name != "Global":
                    system_values[name] = name
        except Exception as e:
            print("Error reading the PipeSystemType values: " + str(e))

    return system_values


# Funzione per ottenere elementi in base al nome del sistema
def get_elements_by_system_name(system_name):
    """Ottiene elementi in base al nome del sistema"""
    pipes = []
    fittings = []

    # Categorie da controllare
    categories = {
        BuiltInCategory.OST_PipeCurves: pipes,    # Tubazioni
        BuiltInCategory.OST_PipeFitting: fittings # Raccordi
    }

    for category, collection in categories.items():
        try:
            elements = FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().ToElements()
            # Per ogni elemento, controlla se corrisponde al sistema selezionato
            for element in elements:
                try:
                    param = element.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
                    if param and param.HasValue and param.AsValueString() == system_name:
                        collection.append(element)
                except Exception as e:
                    print("Error reading element {}: {}".format(get_element_id_value(element.Id), str(e)))
                    continue
        except Exception as e:
            print("Error collecting category {}: {}".format(category, str(e)))

    print("Elements found in system '{}': {} pipes, {} fittings".format(system_name, len(pipes), len(fittings)))
    return pipes, fittings


# Funzione per caricare la tabella degli isolanti da un file di testo
def load_insulation_table(file_path):
    tabella_isolanti = []
    try:
        with open(file_path, 'r') as file:
            # Leggi il file usando csv reader
            reader = csv.reader(file, delimiter='\t')  # Assumiamo che il file sia delimitato da tab
            headers = next(reader)  # Leggi l'intestazione

            # Verifica che l'intestazione contenga almeno "Diameter" e alcune zone termiche
            if "Diameter" not in headers:
                forms.alert('The table has no "Diameter" column. Check the file format '
                            '(tab-separated, see insulation_table_template.txt).', exitscript=True)

            # Controlla le zone termiche disponibili nell'intestazione
            available_zones = [h for h in headers if h in zone_termiche_valide]
            if not available_zones:
                forms.alert('The table has no thermal zone column ({}). Check the file format.'.format(
                    ", ".join(zone_termiche_valide)), exitscript=True)

            for row in reader:
                if len(row) < len(headers):  # Verifica che la riga sia completa
                    continue

                row_dict = {}
                for i, header in enumerate(headers):
                    try:
                        if header == "Diameter" or header in zone_termiche_valide:
                            row_dict[header] = int(row[i])
                    except ValueError:
                        # Se la conversione fallisce, salta il valore
                        print("Table: value '{}' in column {} is not a whole number, skipped".format(row[i], header))

                if "Diameter" in row_dict:
                    tabella_isolanti.append(row_dict)

        # Ordina la tabella per diametro
        tabella_isolanti.sort(key=lambda x: x["Diameter"])

        return tabella_isolanti
    except Exception as e:
        forms.alert('Error reading the table file:\n{}'.format(str(e)), exitscript=True)
        return []


# Funzione per ottenere tutti i tipi di isolamento disponibili
def get_insulation_types():
    pipe_insulation_types = FilteredElementCollector(doc)\
                           .OfClass(PipeInsulationType)\
                           .ToElements()

    result = {}
    for insulation_type in pipe_insulation_types:
        try:
            type_name = Element.Name.GetValue(insulation_type)
            if not type_name:
                # Prova con il parametro di tipo
                type_param = insulation_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                if type_param:
                    type_name = type_param.AsString()
                else:
                    type_name = "Insulation type " + insulation_type.Id.ToString()

            result[type_name] = insulation_type
        except Exception as e:
            print("Error reading an insulation type name: {}".format(str(e)))

    return result


# Funzione per verificare se una tubazione o un raccordo ha gia' un isolamento
def get_existing_insulation(element):
    try:
        collector = FilteredElementCollector(doc).OfClass(PipeInsulation)
        for insulation in collector:
            if insulation.HostElementId == element.Id:
                return insulation
    except Exception as e:
        print("Error looking for the insulation of element {}: {}".format(get_element_id_value(element.Id), str(e)))

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
                break

        if thickness_param and thickness_param.StorageType == StorageType.Double:
            if thickness_param.IsReadOnly:
                print("The thickness parameter of insulation {} is read-only".format(get_element_id_value(insulation.Id)))
            else:
                thickness_param.Set(thickness_feet)
                success = True
        else:
            print("Insulation {}: thickness parameter not found or not a length".format(get_element_id_value(insulation.Id)))

        # Prova a cambiare il tipo di isolamento
        try:
            if insulation.GetTypeId() != insulation_type_id:
                insulation.ChangeTypeId(insulation_type_id)
                success = True
        except Exception as type_e:
            print("Error changing the insulation type: {}".format(str(type_e)))
            error_message += " Type error: " + str(type_e)

    except Exception as e:
        error_message = str(e)
        print("Error updating the insulation: {}".format(error_message))

    return success, error_message


# Funzione per verificare se un elemento e' "a vista" in base al parametro e_EAN_IsVisible_1
def is_element_visible(element):
    try:
        for param in element.Parameters:
            if param.Definition.Name == "e_EAN_IsVisible_1":
                if param.StorageType == StorageType.Integer:
                    # Per i parametri booleani, 1 = True, 0 = False
                    return param.AsInteger() == 1
                return False  # Se il tipo non e' corretto, considera come non visibile

        # Se il parametro non esiste, considera come non visibile
        return False
    except Exception as e:
        print("Error reading e_EAN_IsVisible_1: {}".format(str(e)))
        return False


# Funzione per ottenere la zona termica di un elemento
def get_thermal_position(element):
    try:
        for param in element.Parameters:
            if param.Definition.Name == "e_EAN_ThermalPosition_1" and param.HasValue:
                if param.StorageType == StorageType.String:
                    value = param.AsString()
                    # Verifica che il valore sia tra quelli validi
                    return value, value in zone_termiche_valide

        # Se il parametro non esiste o non ha un valore valido
        return "", False
    except Exception as e:
        print("Error reading e_EAN_ThermalPosition_1: {}".format(str(e)))
        return "", False


# Funzione per ottenere il diametro di una tubazione o di un raccordo
def get_pipe_diameter(element):
    try:
        diameter_mm = 0

        if isinstance(element, Pipe):
            # Per le tubazioni, ottieni il diametro dalla proprieta' Diameter
            diameter_feet = element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
            diameter_mm = int(round(diameter_feet * 304.8))  # Converti da piedi a mm e arrotonda
        elif isinstance(element, FamilyInstance):
            # Per i raccordi, prova diversi parametri comuni per il diametro
            diameter_param = None
            for param_name in ["Diametro", "Diameter", "Nominal Diameter", "Diametro nominale"]:
                for param in element.Parameters:
                    if param.Definition.Name == param_name and param.HasValue:
                        diameter_param = param
                        break
                if diameter_param:
                    break

            if diameter_param and diameter_param.StorageType == StorageType.Double:
                diameter_mm = int(round(diameter_param.AsDouble() * 304.8))
            else:
                # Prova a ottenere il diametro dal connettore
                if hasattr(element, "MEPModel") and element.MEPModel:
                    connectors = element.MEPModel.ConnectorManager.Connectors
                    if connectors and connectors.Size > 0:
                        for connector in connectors:
                            if connector.Shape == ConnectorProfileType.Round:
                                diameter_mm = int(round(connector.Radius * 2 * 304.8))
                                break

        if diameter_mm <= 0:
            return 0, False

        return diameter_mm, True
    except Exception as e:
        print("Error reading the diameter: {}".format(str(e)))
        return 0, False


# Funzione per ottenere lo spessore dell'isolamento dalla tabella
def get_insulation_thickness(diameter_mm, thermal_position, tabella_isolanti):
    try:
        if thermal_position not in zone_termiche_valide:
            return 0, False

        # Trova la riga con il diametro piu' vicino
        closest_row = None
        min_diff = float('inf')
        for row in tabella_isolanti:
            diff = abs(row["Diameter"] - diameter_mm)
            if diff < min_diff:
                min_diff = diff
                closest_row = row

        if closest_row and thermal_position in closest_row:
            return closest_row[thermal_position], True
        return 0, False
    except Exception as e:
        print("Error looking up the thickness in the table: {}".format(str(e)))
        return 0, False


def process_elements(elements, label, visible_type, hidden_type, manual_thickness_mm, tabella_isolanti):
    """Aggiunge o modifica l'isolamento; ritorna i contatori (added, modified, skipped, error)."""
    counts = {"added": 0, "modified": 0, "skipped": 0, "error": 0}
    for element in elements:
        try:
            element_id = get_element_id_value(element.Id)

            # Determina se l'elemento e' "a vista" e scegli il tipo di isolamento
            is_visible = is_element_visible(element)
            insulation_type_id = visible_type.Id if is_visible else hidden_type.Id
            exposure = "exposed" if is_visible else "concealed"

            # Determina lo spessore dell'isolamento
            if manual_thickness_mm is not None:
                thickness_mm = manual_thickness_mm
            else:
                diameter_mm, valid_diameter = get_pipe_diameter(element)
                if not valid_diameter:
                    print("{} {}: diameter not found, skipped".format(label, element_id))
                    counts["skipped"] += 1
                    continue

                thermal_position, valid_position = get_thermal_position(element)
                if not valid_position:
                    print("{} {}: thermal zone '{}' not valid (allowed: {}), skipped".format(
                        label, element_id, thermal_position, ", ".join(zone_termiche_valide)))
                    counts["skipped"] += 1
                    continue

                thickness_mm, valid_thickness = get_insulation_thickness(diameter_mm, thermal_position, tabella_isolanti)
                if not valid_thickness:
                    print("{} {}: no thickness in the table for {} mm / {}, skipped".format(
                        label, element_id, diameter_mm, thermal_position))
                    counts["skipped"] += 1
                    continue

            # Converti mm in piedi (unita' interne di Revit)
            thickness_feet = thickness_mm / 304.8

            existing_insulation = get_existing_insulation(element)
            if not existing_insulation:
                # Se non ha isolamento, aggiungilo
                try:
                    PipeInsulation.Create(doc, element.Id, insulation_type_id, thickness_feet)
                    counts["added"] += 1
                    print("{} {}: insulation added ({}, {} mm)".format(label, element_id, exposure, thickness_mm))
                except Exception as add_e:
                    print("{} {}: error adding the insulation: {}".format(label, element_id, str(add_e)))
                    counts["error"] += 1
            else:
                # Se ha gia' un isolamento, prova a modificarlo
                success, error = modify_existing_insulation(existing_insulation, thickness_feet, insulation_type_id)
                if success:
                    counts["modified"] += 1
                    print("{} {}: insulation updated ({}, {} mm)".format(label, element_id, exposure, thickness_mm))
                else:
                    # Se la modifica fallisce, prova l'approccio elimina-e-ricrea
                    try:
                        doc.Delete(existing_insulation.Id)
                        PipeInsulation.Create(doc, element.Id, insulation_type_id, thickness_feet)
                        counts["modified"] += 1
                        print("{} {}: insulation recreated ({}, {} mm)".format(label, element_id, exposure, thickness_mm))
                    except Exception as recreate_e:
                        print("{} {}: error recreating the insulation: {}".format(label, element_id, str(recreate_e)))
                        counts["error"] += 1
        except Exception as e:
            print("{} {}: error: {}".format(label, element.Id, str(e)))
            counts["error"] += 1
    return counts


def main():
    # Ottieni i tipi di sistema di tubazioni
    system_types = get_piping_system_types()
    if not system_types:
        forms.alert("No piping system type found in the model.", exitscript=True)
    system_names = sorted(system_types.keys())

    # Ottieni i tipi di isolamento disponibili
    insulation_types_dict = get_insulation_types()
    if not insulation_types_dict:
        forms.alert('The project has no pipe insulation type.', exitscript=True)
    insulation_types_list = sorted(insulation_types_dict.keys())

    selection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]

    config = script.get_config(CONFIG_NAME)
    values = SetPipeInsulation_ui.SetPipeInsulationWindow(
        script.get_bundle_file(XAML_FILE_NAME), system_names, insulation_types_list,
        len(selection), load_settings(config)).show()
    if values is None:
        script.exit()
    save_settings(config, values)

    selected_system = values["system"]
    visible_insulation_type = insulation_types_dict[values["visible_type"]]
    hidden_insulation_type = insulation_types_dict[values["hidden_type"]]
    manual_thickness_mm = values["thickness"] if values["manual"] else None
    table_file_path = values["table_path"]

    # Carica la tabella degli isolanti dal file (non serve con lo spessore fisso)
    tabella_isolanti = []
    if manual_thickness_mm is None:
        tabella_isolanti = load_insulation_table(table_file_path)
        if not tabella_isolanti:
            forms.alert('The thickness table has no valid row.', exitscript=True)

    # Ottieni gli elementi in base al sistema selezionato o alla selezione corrente
    if not values["selected_only"]:
        pipes, fittings = get_elements_by_system_name(selected_system)
        if not pipes and not fittings:
            forms.alert('No element found in system "{}".'.format(selected_system), exitscript=True)
    else:
        pipes = []
        fittings = []
        pipe_fitting_category_id = ElementId(BuiltInCategory.OST_PipeFitting)
        for elem in selection:
            param = elem.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
            in_system = param and param.HasValue and param.AsValueString() == selected_system
            if not in_system:
                continue
            if isinstance(elem, Pipe):
                pipes.append(elem)
            elif isinstance(elem, FamilyInstance) and elem.Category and elem.Category.Id == pipe_fitting_category_id:
                fittings.append(elem)

        if not pipes and not fittings:
            forms.alert('The selection contains no pipe or fitting of system "{}".'.format(selected_system), exitscript=True)

    # Esegui le operazioni in una transazione
    with revit.Transaction("Set pipe and fitting insulation"):
        pipe_counts = process_elements(pipes, "Pipe", visible_insulation_type, hidden_insulation_type,
                                       manual_thickness_mm, tabella_isolanti)
        fitting_counts = process_elements(fittings, "Fitting", visible_insulation_type, hidden_insulation_type,
                                          manual_thickness_mm, tabella_isolanti)

    # Riepilogo nella finestra di output
    output.print_md("---")
    output.print_md("## Summary")
    output.print_md("- System: **{}**".format(selected_system))
    output.print_md("- Scope: **{}**".format("selected elements" if values["selected_only"] else "all elements of the system"))
    if manual_thickness_mm is not None:
        output.print_md("- Thickness: **{} mm** (fixed)".format(manual_thickness_mm))
    else:
        output.print_md("- Thickness table: **{}**".format(os.path.basename(table_file_path)))
    output.print_table(
        table_data=[["Pipes", pipe_counts["added"], pipe_counts["modified"], pipe_counts["skipped"], pipe_counts["error"]],
                    ["Fittings", fitting_counts["added"], fitting_counts["modified"], fitting_counts["skipped"], fitting_counts["error"]]],
        columns=["", "Insulated", "Updated", "Skipped", "Failed"])
    if pipe_counts["skipped"] or fitting_counts["skipped"]:
        output.print_md("Skipped elements are listed above, with the reason.")


# Gestione degli errori a livello di script
try:
    main()
except Exception as e:
    import traceback
    error_msg = "An error occurred:\n" + str(e) + "\n\nTraceback:\n" + traceback.format_exc()
    print(error_msg)
    forms.alert(error_msg, exitscript=True)
