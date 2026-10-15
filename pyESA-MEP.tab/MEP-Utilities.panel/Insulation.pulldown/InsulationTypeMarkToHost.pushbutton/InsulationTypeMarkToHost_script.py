# -*- coding: utf-8 -*-
__title__ = "Insulation TypeMark\nTo Host"
__doc__ = """Version = 2.1
Date = 14.10.2026
________________________________________________________________
Copies the Type Mark of the insulation to an instance parameter
of its pipe or duct.

- Choose pipes, ducts or both
- Work on the selected elements only, or on the whole project
- Elements without insulation are skipped
- The target is any writable text instance parameter
________________________________________________________________
Author: Andrea Patti
"""

from pyrevit import revit, DB, forms, script

import InsulationTypeMarkToHost_ui

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

XAML_FILE_NAME = "InsulationTypeMarkToHost_form.xaml"
CONFIG_NAME = "ESA_InsulationTypeMarkToHost"
DEFAULTS = {"pipe": True, "duct": True, "target_param": ""}
# Elementi per categoria letti per proporre i parametri di destinazione
PARAM_SAMPLE_SIZE = 10

# ---------------------------------------------------------------------------
# Compatibilita ElementId: Revit <= 2025 usa .IntegerValue, Revit >= 2026 usa .Value
# ---------------------------------------------------------------------------
def get_element_id_value(eid):
    """Restituisce il valore numerico di un ElementId, compatibile con tutte le versioni."""
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
        settings["pipe"] = to_bool(config.get_option("pipe", DEFAULTS["pipe"]))
        settings["duct"] = to_bool(config.get_option("duct", DEFAULTS["duct"]))
        settings["target_param"] = config.get_option("target_param", DEFAULTS["target_param"]) or ""
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

def get_all_of(category):
    """Ottiene tutte le istanze di una categoria del modello"""
    return list(DB.FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .ToElements())

def get_selected_of(category):
    """Ottiene gli elementi selezionati di una categoria"""
    selection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
    return [elem for elem in selection
            if elem and elem.Category and get_element_id_value(elem.Category.Id) == int(category)]

def get_instance_text_parameters(elements):
    """Ottiene tutti i parametri di istanza di tipo testo scrivibili"""
    param_names = set()
    for element in elements:
        for param in element.Parameters:
            if (not param.IsReadOnly and
                param.StorageType == DB.StorageType.String):
                param_names.add(param.Definition.Name)
    return sorted(list(param_names))

def get_type_name(insulation_type):
    param = insulation_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
    return param.AsString() if param else ""

def write_type_marks(elements, insulation_dict, label, selected_param, counts):
    """Scrive il Type Mark dell'isolamento sugli elementi; aggiorna counts (success, no_typemark, error)."""
    for element in elements:
        try:
            # Ottieni il tipo di isolamento
            insulation_type = doc.GetElement(insulation_dict[element.Id].GetTypeId())
            if not insulation_type:
                counts["error"] += 1
                output.print_md('✗ {} {}: insulation type not found'.format(label, output.linkify(element.Id)))
                continue

            # Ottieni il Type Mark
            type_mark_param = insulation_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK)
            type_mark = type_mark_param.AsString() if type_mark_param and type_mark_param.HasValue else None
            if not (type_mark and type_mark.strip()):
                counts["no_typemark"] += 1
                output.print_md('⚠ {} {}: empty Type Mark on insulation **{}**'.format(
                    label, output.linkify(element.Id), get_type_name(insulation_type)))
                continue

            # Trova il parametro nell'elemento
            target = element.LookupParameter(selected_param)
            if target and not target.IsReadOnly:
                target.Set(type_mark)
                counts["success"] += 1
                output.print_md('✓ {} {}: Type Mark = **{}** (insulation: {})'.format(
                    label, output.linkify(element.Id), type_mark, get_type_name(insulation_type)))
            else:
                counts["error"] += 1
                output.print_md('✗ {} {}: parameter not found or read-only'.format(
                    label, output.linkify(element.Id)))
        except Exception as e:
            counts["error"] += 1
            output.print_md('✗ {} {}: error - {}'.format(label, output.linkify(element.Id), str(e)))

def main():
    # Ottieni gli elementi selezionati (se presenti)
    selected_pipes = get_selected_of(DB.BuiltInCategory.OST_PipeCurves)
    selected_ducts = get_selected_of(DB.BuiltInCategory.OST_DuctCurves)
    has_selection = bool(selected_pipes or selected_ducts)

    # Parametri proposti: letti da un campione di tubazioni e canali (selezionati, altrimenti tutti)
    if has_selection:
        sample = selected_pipes[:PARAM_SAMPLE_SIZE] + selected_ducts[:PARAM_SAMPLE_SIZE]
    else:
        sample = (get_all_of(DB.BuiltInCategory.OST_PipeCurves)[:PARAM_SAMPLE_SIZE] +
                  get_all_of(DB.BuiltInCategory.OST_DuctCurves)[:PARAM_SAMPLE_SIZE])
    available_params = get_instance_text_parameters(sample)

    config = script.get_config(CONFIG_NAME)
    values = InsulationTypeMarkToHost_ui.TypeMarkToHostWindow(
        script.get_bundle_file(XAML_FILE_NAME), len(selected_pipes), len(selected_ducts),
        available_params, load_settings(config)).show()
    if values is None:
        script.exit()
    save_settings(config, values)

    process_pipes = values["pipe"]
    process_ducts = values["duct"]
    use_selection = values["use_selection"]
    selected_param = values["target_param"]

    # Ottieni gli elementi da processare
    pipes_to_process = []
    ducts_to_process = []
    if process_pipes:
        pipes_to_process = selected_pipes if use_selection else get_all_of(DB.BuiltInCategory.OST_PipeCurves)
    if process_ducts:
        ducts_to_process = selected_ducts if use_selection else get_all_of(DB.BuiltInCategory.OST_DuctCurves)

    output.print_md('## Insulation Type Mark to host')
    output.print_md('Target parameter: **{}**'.format(selected_param))
    output.print_md('---')
    if process_pipes:
        output.print_md('Pipes to check: **{}**'.format(len(pipes_to_process)))
    if process_ducts:
        output.print_md('Ducts to check: **{}**'.format(len(ducts_to_process)))

    if not pipes_to_process and not ducts_to_process:
        output.print_md('Nothing to process: {}.'.format(
            'the selection contains no element of the chosen categories' if use_selection
            else 'the model contains no element of the chosen categories'))
        script.exit()

    # Crea dizionari: element_id -> insulation
    pipe_insulation_dict = {}
    duct_insulation_dict = {}
    for wanted, category, insulation_dict in (
            (process_pipes, DB.BuiltInCategory.OST_PipeInsulations, pipe_insulation_dict),
            (process_ducts, DB.BuiltInCategory.OST_DuctInsulations, duct_insulation_dict)):
        if not wanted:
            continue
        for insulation in get_all_of(category):
            host_id = insulation.HostElementId
            if host_id and host_id != DB.ElementId.InvalidElementId:
                insulation_dict[host_id] = insulation

    # Filtra gli elementi che hanno isolamento
    pipes_with_insulation = [p for p in pipes_to_process if p.Id in pipe_insulation_dict]
    ducts_with_insulation = [d for d in ducts_to_process if d.Id in duct_insulation_dict]

    output.print_md('---')
    if process_pipes:
        output.print_md('**Pipes:** with insulation **{}**, without insulation **{}**'.format(
            len(pipes_with_insulation), len(pipes_to_process) - len(pipes_with_insulation)))
    if process_ducts:
        output.print_md('**Ducts:** with insulation **{}**, without insulation **{}**'.format(
            len(ducts_with_insulation), len(ducts_to_process) - len(ducts_with_insulation)))
    output.print_md('---')

    elements_with_insulation = pipes_with_insulation + ducts_with_insulation
    if not elements_with_insulation:
        output.print_md('Nothing to write: none of the elements checked has insulation.')
        script.exit()

    counts = {"success": 0, "no_typemark": 0, "error": 0}

    try:
        with revit.Transaction('Copy Insulation Type Mark'):
            if pipes_with_insulation:
                output.print_md('### Pipes')
                write_type_marks(pipes_with_insulation, pipe_insulation_dict, 'Pipe', selected_param, counts)
            if ducts_with_insulation:
                output.print_md('---')
                output.print_md('### Ducts')
                write_type_marks(ducts_with_insulation, duct_insulation_dict, 'Duct', selected_param, counts)
    except Exception as e:
        forms.alert('Error while processing, no change was made:\n{}'.format(str(e)), exitscript=True)

    # Stampa riepilogo
    output.print_md('---')
    output.print_md('## Summary')
    output.print_md('- Scope: **{}**'.format('selected elements' if use_selection else 'all elements in the project'))
    output.print_md('- Categories: **{}**'.format(
        ', '.join([c for c in ['Pipes' if process_pipes else None,
                                'Ducts' if process_ducts else None] if c])))
    if use_selection:
        if selected_pipes:
            output.print_md('- Selected pipes: **{}**'.format(len(selected_pipes)))
        if selected_ducts:
            output.print_md('- Selected ducts: **{}**'.format(len(selected_ducts)))
    output.print_md('- Elements with insulation: **{}**'.format(len(elements_with_insulation)))
    output.print_md('- ✓ **Written**: {}'.format(counts["success"]))
    output.print_md('- ⚠ **Empty Type Mark**: {}'.format(counts["no_typemark"]))
    output.print_md('- ✗ **Errors**: {}'.format(counts["error"]))
    output.print_md('---')

if __name__ == '__main__':
    main()
