# -*- coding: utf-8 -*-
__title__ = 'Set Pipe Insulation\nDPR 412/93'
__author__ = 'Andrea Patti'
__doc__ = """Version = 1.1
Date    = 14.10.2026
________________________________________________________________
Adds or updates the insulation of the selected pipes and pipe
fittings, by exposure (e_EAN_IsVisible_1) and thermal zone
(e_EAN_ThermalPosition_1), with the DPR 412/93 thickness table
built into the tool, or with a fixed thickness.
________________________________________________________________
Author(s):
Andrea Patti
"""

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import *
from pyrevit import revit, DB
from pyrevit import forms, script

import SetInsulationDPR412_ui

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

XAML_FILE_NAME = "SetInsulationDPR412_form.xaml"
CONFIG_NAME = "ESA_SetInsulationDPR412"
DEFAULTS = {"visible_type": "", "hidden_type": "", "manual": False, "thickness": 13}

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
        settings["visible_type"] = config.get_option("visible_type", "") or ""
        settings["hidden_type"] = config.get_option("hidden_type", "") or ""
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
            # Verifica se il FamilyInstance e' un raccordo di tubazione
            if elem.Category and elem.Category.Id == pipe_fitting_category_id:
                fittings.append(elem)

    return pipes, fittings


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
def get_insulation_thickness(diameter_mm, thermal_position):
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

        if closest_row:
            return closest_row[thermal_position], True
        return 0, False
    except Exception as e:
        print("Error looking up the thickness in the table: {}".format(str(e)))
        return 0, False


def process_elements(elements, label, visible_type, hidden_type, manual_thickness_mm):
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

                thickness_mm, valid_thickness = get_insulation_thickness(diameter_mm, thermal_position)
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


# Ottieni tubazioni e raccordi selezionati
pipes, fittings = get_pipes_and_fittings_from_selection()

if not pipes and not fittings:
    forms.alert('Select at least one pipe or pipe fitting before running the tool.', exitscript=True)

# Ottieni i tipi di isolamento disponibili
insulation_types_dict = get_insulation_types()

if not insulation_types_dict:
    forms.alert('The project has no pipe insulation type.', exitscript=True)

config = script.get_config(CONFIG_NAME)
values = SetInsulationDPR412_ui.SetInsulationDPR412Window(
    script.get_bundle_file(XAML_FILE_NAME), sorted(insulation_types_dict.keys()),
    len(pipes), len(fittings), load_settings(config)).show()
if values is None:
    script.exit()
save_settings(config, values)

visible_insulation_type = insulation_types_dict[values["visible_type"]]
hidden_insulation_type = insulation_types_dict[values["hidden_type"]]
manual_thickness_mm = values["thickness"] if values["manual"] else None

# Esegui le operazioni in una transazione
with revit.Transaction("Set pipe and fitting insulation"):
    pipe_counts = process_elements(pipes, "Pipe", visible_insulation_type, hidden_insulation_type, manual_thickness_mm)
    fitting_counts = process_elements(fittings, "Fitting", visible_insulation_type, hidden_insulation_type, manual_thickness_mm)

# Riepilogo nella finestra di output
output.print_md("---")
output.print_md("## Summary")
if manual_thickness_mm is not None:
    output.print_md("- Thickness: **{} mm** (fixed)".format(manual_thickness_mm))
else:
    output.print_md("- Thickness: **DPR 412/93 table**")
output.print_table(
    table_data=[["Pipes", pipe_counts["added"], pipe_counts["modified"], pipe_counts["skipped"], pipe_counts["error"]],
                ["Fittings", fitting_counts["added"], fitting_counts["modified"], fitting_counts["skipped"], fitting_counts["error"]]],
    columns=["", "Insulated", "Updated", "Skipped", "Failed"])
if pipe_counts["skipped"] or fitting_counts["skipped"]:
    output.print_md("Skipped elements are listed above, with the reason.")
