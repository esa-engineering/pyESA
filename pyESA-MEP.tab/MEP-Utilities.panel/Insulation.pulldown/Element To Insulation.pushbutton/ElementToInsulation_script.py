# -*- coding: utf-8 -*-
__title__   = "Element To\nInsulation"
__doc__     = """Version = 1.1
Date    = 14.10.2026
________________________________________________________________
Copies the values of the chosen parameters from pipes and ducts
to their insulation.

- Parameter names separated by ; (e.g. Comments;Mark)
- Choose the categories: pipe insulation, duct insulation
- Work on the selected elements only, or on the whole project
- Text, number, Yes/No and element parameters are supported
________________________________________________________________
Author(s):
Andrea Patti, Tommaso Lorenzi
"""

#REFERENCES
from pyrevit import revit, script, DB
from Autodesk.Revit.DB import ElementId, StorageType, FilteredElementCollector, BuiltInCategory

import ElementToInsulation_ui

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

XAML_FILE_NAME = "ElementToInsulation_form.xaml"
CONFIG_NAME = "ESA_ElementToInsulation"
DEFAULTS = {"parameters": "", "pipe": True, "duct": True}

# Categorie gestite: chiave del form -> (categoria, etichetta nel report)
# Raccordi tubazioni (OST_PipeFittingInsulation) e raccordi canali (OST_DuctFittingInsulation)
# erano gia' esclusi nella versione con FlexForm.
CATEGORIES = [
    ("pipe", BuiltInCategory.OST_PipeInsulations, "Pipe insulation"),
    ("duct", BuiltInCategory.OST_DuctInsulations, "Duct insulation"),
]


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
        settings["parameters"] = config.get_option("parameters", DEFAULTS["parameters"]) or ""
        settings["pipe"] = to_bool(config.get_option("pipe", DEFAULTS["pipe"]))
        settings["duct"] = to_bool(config.get_option("duct", DEFAULTS["duct"]))
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


# Ottieni gli elementi selezionati
selection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]

config = script.get_config(CONFIG_NAME)
values = ElementToInsulation_ui.ElementToInsulationWindow(
    script.get_bundle_file(XAML_FILE_NAME), len(selection), load_settings(config)).show()
if values is None:
    script.exit()
save_settings(config, values)

parameters_split = values["names"]
use_selection = values["use_selection"]
selected_categories = [(cat, label) for key, cat, label in CATEGORIES if values[key]]

# Dizionario per tenere traccia delle categorie
category_counts = dict((cat, 0) for cat, label in selected_categories)

# Lista per tutti gli isolamenti da processare
matching_insulations = []

# Ottieni gli ID degli elementi host selezionati
selected_host_ids = [elem.Id for elem in selection]

for category, label in selected_categories:
    insulations = FilteredElementCollector(doc)\
                .OfCategory(category)\
                .WhereElementIsNotElementType()\
                .ToElements()
    for insulation in insulations:
        host_id = insulation.HostElementId
        if use_selection:
            # Modalita' elementi selezionati: solo isolamenti con host nella selezione
            if host_id not in selected_host_ids:
                continue
        elif host_id == ElementId.InvalidElementId:
            continue
        matching_insulations.append(insulation)
        category_counts[category] += 1

# Stampa il conteggio per categoria
output.print_md("## Insulation found")
for category, label in selected_categories:
    output.print_md("- {}: {}".format(label, category_counts[category]))

# Contatori per il report finale
updated_count = 0
errors_count = 0

with revit.Transaction("Set Insulation Parameter"):
    # Ciclo for per ogni isolante che corrisponde a un host selezionato
    for insulation in matching_insulations:
        host = doc.GetElement(insulation.HostElementId)

        # Verifica che l'host sia valido
        if not host:
            continue

        # Flag per verificare se almeno un parametro e' stato copiato
        param_copied = False

        # Ciclo for per estrarre i valori dagli elementi host e copiarli sugli isolanti
        for param_name in parameters_split:
            # Recupera il parametro dall'elemento Host
            host_param = host.LookupParameter(param_name)
            if not host_param:
                continue
            # Recupera il parametro nell'insulation
            insulation_param = insulation.LookupParameter(param_name)
            if not insulation_param:
                continue
            # Verifica che il parametro sia scrivibile
            if insulation_param.IsReadOnly:
                continue
            try:
                # Copia il valore in base al tipo di parametro
                storage_type = host_param.StorageType
                if storage_type == StorageType.String:
                    host_value = host_param.AsString()
                    if host_value:  # Verifica che il valore non sia vuoto
                        insulation_param.Set(host_value)
                        param_copied = True
                elif storage_type == StorageType.Integer:
                    insulation_param.Set(host_param.AsInteger())
                    param_copied = True
                elif storage_type == StorageType.Double:
                    insulation_param.Set(host_param.AsDouble())
                    param_copied = True
                elif storage_type == StorageType.ElementId:
                    insulation_param.Set(host_param.AsElementId())
                    param_copied = True
            except Exception as e:
                errors_count += 1
                output.print_md("Error copying parameter '{}' to element {} ({}): {}".format(
                    param_name, output.linkify(insulation.Id), get_element_id_value(insulation.Id), str(e)))
        if param_copied:
            updated_count += 1

output.print_md("---")
output.print_md("**Summary:**")
output.print_md("- Scope: {}".format("selected elements" if use_selection else "all elements in the project"))
output.print_md("- Parameters: {}".format(", ".join(parameters_split)))
if use_selection:
    output.print_md("- Selected elements: {}".format(len(selection)))
output.print_md("- Insulation elements found: {}".format(len(matching_insulations)))
output.print_md("- Insulation elements updated: {}".format(updated_count))
if errors_count > 0:
    output.print_md("- Errors: {}".format(errors_count))
if matching_insulations and not updated_count:
    output.print_md("No value was copied: check that the parameters exist on both the host and the "
                    "insulation, that they are not read-only and that the host values are not empty.")
output.print_md("---")
