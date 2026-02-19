# -*- coding: utf-8 -*-
"""Gestione avanzata colori filtri - Copia colori tra linee e superfici.

Questo script permette di:
1. Selezionare uno o piu View Template dal progetto
2. Selezionare uno o piu filtri applicati ai template
3. Scegliere la direzione della copia (linea->superficie o superficie->linea)
4. Scegliere il pattern di superficie da applicare
5. Gestire il colore delle linee (mantenere, resettare, nuovo colore)
6. Gestire il pattern delle linee (mantenere, resettare)
7. Salvare e caricare preset di impostazioni
8. Visualizzare anteprima dei colori prima di applicare
"""

__title__ = "Filter\nColors\nManager"
__author__ = "Andrea Patti"

import traceback
import json
import os
from pyrevit import revit, DB, forms, script

# Ottiene il documento corrente
doc = revit.doc
output = script.get_output()

# Percorso per salvare i preset
PRESET_FOLDER = os.path.join(os.path.dirname(__file__), "presets")
if not os.path.exists(PRESET_FOLDER):
    try:
        os.makedirs(PRESET_FOLDER)
    except:
        PRESET_FOLDER = None


# ============================================================
# COMPATIBILITA VERSIONI REVIT
# ============================================================

def get_element_id_value(element_id):
    """Ottiene il valore numerico di un ElementId in modo compatibile.
    Revit <= 2023 usa .IntegerValue (int)
    Revit >= 2024 usa .Value (long)
    """
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


# ============================================================
# FUNZIONI UTILITA
# ============================================================

def get_all_view_templates():
    """Recupera tutti i View Template dal progetto."""
    collector = DB.FilteredElementCollector(doc)\
        .OfClass(DB.View)\
        .WhereElementIsNotElementType()
    
    templates = []
    for view in collector:
        if view.IsTemplate:
            templates.append(view)
    
    return templates


def get_filters_from_templates(view_templates):
    """Recupera tutti i filtri applicati a una lista di View Template."""
    filters_dict = {}
    
    for template in view_templates:
        filter_ids = template.GetFilters()
        for filter_id in filter_ids:
            filter_element = doc.GetElement(filter_id)
            if filter_element:
                filter_name = filter_element.Name
                if filter_name not in filters_dict:
                    filters_dict[filter_name] = {
                        "id": filter_id,
                        "templates": []
                    }
                filters_dict[filter_name]["templates"].append(template)
    
    return filters_dict


def get_all_fill_patterns():
    """Recupera tutti i pattern di riempimento dal progetto."""
    collector = DB.FilteredElementCollector(doc)\
        .OfClass(DB.FillPatternElement)
    
    patterns_dict = {}
    solid_pattern_name = None
    
    for pattern in collector:
        fill_pattern = pattern.GetFillPattern()
        if fill_pattern.Target == DB.FillPatternTarget.Drafting:
            pattern_name = pattern.Name
            patterns_dict[pattern_name] = pattern.Id
            if fill_pattern.IsSolidFill:
                solid_pattern_name = pattern_name
    
    return patterns_dict, solid_pattern_name


def get_line_color_from_overrides(overrides):
    """Estrae il colore della linea dalle override."""
    try:
        line_color = overrides.ProjectionLineColor
        if line_color.IsValid:
            return line_color, "ProjectionLineColor"
    except:
        pass
    
    try:
        line_color = overrides.CutLineColor
        if line_color.IsValid:
            return line_color, "CutLineColor"
    except:
        pass
    
    return None, None


def get_surface_color_from_overrides(overrides):
    """Estrae il colore della superficie dalle override."""
    try:
        surface_color = overrides.SurfaceForegroundPatternColor
        if surface_color.IsValid:
            return surface_color, "SurfaceForegroundPatternColor"
    except:
        pass
    
    try:
        surface_color = overrides.CutForegroundPatternColor
        if surface_color.IsValid:
            return surface_color, "CutForegroundPatternColor"
    except:
        pass
    
    return None, None


def parse_rgb_input(rgb_string):
    """Converte una stringa RGB in un oggetto Color di Revit."""
    try:
        rgb_string = rgb_string.strip().replace(" ", "")
        
        for separator in [",", ";", " "]:
            if separator in rgb_string:
                parts = rgb_string.split(separator)
                break
        else:
            parts = rgb_string.split()
        
        if len(parts) != 3:
            return None
        
        r = int(parts[0].strip())
        g = int(parts[1].strip())
        b = int(parts[2].strip())
        
        if all(0 <= v <= 255 for v in [r, g, b]):
            return DB.Color(r, g, b)
        else:
            return None
    except:
        return None


def color_to_hex(color):
    """Converte un colore Revit in formato HEX per HTML."""
    if color and color.IsValid:
        r = color.Red
        g = color.Green
        b = color.Blue
        return "#{:02X}{:02X}{:02X}".format(r, g, b)
    return None


def color_to_rgb_string(color):
    """Converte un colore Revit in stringa RGB."""
    if color and color.IsValid:
        return "RGB({},{},{})".format(color.Red, color.Green, color.Blue)
    return "Non definito"


# ============================================================
# FUNZIONI PRESET
# ============================================================

def get_preset_files():
    """Recupera la lista dei preset salvati."""
    if not PRESET_FOLDER or not os.path.exists(PRESET_FOLDER):
        return []
    
    presets = []
    for f in os.listdir(PRESET_FOLDER):
        if f.endswith(".json"):
            presets.append(f.replace(".json", ""))
    return presets


def save_preset(name, settings):
    """Salva un preset su file."""
    if not PRESET_FOLDER:
        return False
    
    try:
        filepath = os.path.join(PRESET_FOLDER, "{}.json".format(name))
        with open(filepath, "w") as f:
            json.dump(settings, f, indent=2)
        return True
    except:
        return False


def load_preset(name):
    """Carica un preset da file."""
    if not PRESET_FOLDER:
        return None
    
    try:
        filepath = os.path.join(PRESET_FOLDER, "{}.json".format(name))
        with open(filepath, "r") as f:
            return json.load(f)
    except:
        return None


def delete_preset(name):
    """Elimina un preset."""
    if not PRESET_FOLDER:
        return False
    
    try:
        filepath = os.path.join(PRESET_FOLDER, "{}.json".format(name))
        if os.path.exists(filepath):
            os.remove(filepath)
            return True
    except:
        pass
    return False


# ============================================================
# FUNZIONI ANTEPRIMA
# ============================================================

def show_color_preview(filters_dict, selected_filter_names, templates):
    """Mostra anteprima dei colori dei filtri selezionati."""
    
    output.print_md("# Anteprima Colori Filtri")
    output.print_md("---")
    
    for filter_name in selected_filter_names:
        filter_info = filters_dict[filter_name]
        filter_id = filter_info["id"]
        
        output.print_md("## Filtro: {}".format(filter_name))
        
        for template in filter_info["templates"]:
            if template not in templates:
                continue
                
            overrides = template.GetFilterOverrides(filter_id)
            
            # Colore linea
            line_color, line_source = get_line_color_from_overrides(overrides)
            line_rgb = color_to_rgb_string(line_color)
            
            # Colore superficie
            surface_color, surface_source = get_surface_color_from_overrides(overrides)
            surface_rgb = color_to_rgb_string(surface_color)
            
            output.print_md("**Template: {}**".format(template.Name))
            
            # Costruisci tabella HTML con colori RGB inline
            html_table = '<table style="border-collapse:collapse; margin:10px 0;">'
            
            # Riga colore linea
            if line_color and line_color.IsValid:
                r, g, b = line_color.Red, line_color.Green, line_color.Blue
                html_table += '<tr>'
                html_table += '<td style="padding:5px;"><strong>Colore Linea:</strong></td>'
                html_table += '<td style="width:50px; height:20px; border:1px solid #000; background-color:rgb({},{},{});"></td>'.format(r, g, b)
                html_table += '<td style="padding:5px;">{}</td>'.format(line_rgb)
                html_table += '</tr>'
            else:
                html_table += '<tr>'
                html_table += '<td style="padding:5px;"><strong>Colore Linea:</strong></td>'
                html_table += '<td colspan="2" style="padding:5px;">Non definito</td>'
                html_table += '</tr>'
            
            # Riga colore superficie
            if surface_color and surface_color.IsValid:
                r, g, b = surface_color.Red, surface_color.Green, surface_color.Blue
                html_table += '<tr>'
                html_table += '<td style="padding:5px;"><strong>Colore Superficie:</strong></td>'
                html_table += '<td style="width:50px; height:20px; border:1px solid #000; background-color:rgb({},{},{});"></td>'.format(r, g, b)
                html_table += '<td style="padding:5px;">{}</td>'.format(surface_rgb)
                html_table += '</tr>'
            else:
                html_table += '<tr>'
                html_table += '<td style="padding:5px;"><strong>Colore Superficie:</strong></td>'
                html_table += '<td colspan="2" style="padding:5px;">Non definito</td>'
                html_table += '</tr>'
            
            html_table += '</table>'
            
            output.print_html(html_table)
            output.print_md("")
    
    output.print_md("---")


# ============================================================
# FUNZIONI DIAGNOSTICA
# ============================================================

def get_filter_override_info(overrides):
    """Raccoglie informazioni diagnostiche sulle override di un filtro."""
    info = {}
    
    try:
        color = overrides.ProjectionLineColor
        info["ProjectionLineColor"] = color_to_rgb_string(color)
    except Exception as e:
        info["ProjectionLineColor"] = "Errore: {}".format(str(e))
    
    try:
        color = overrides.CutLineColor
        info["CutLineColor"] = color_to_rgb_string(color)
    except Exception as e:
        info["CutLineColor"] = "Errore: {}".format(str(e))
    
    try:
        pattern_id = overrides.ProjectionLinePatternId
        if pattern_id != DB.ElementId.InvalidElementId:
            pattern_elem = doc.GetElement(pattern_id)
            info["ProjectionLinePattern"] = pattern_elem.Name if pattern_elem else "ID: {}".format(get_element_id_value(pattern_id))
        else:
            info["ProjectionLinePattern"] = "Default"
    except Exception as e:
        info["ProjectionLinePattern"] = "Errore: {}".format(str(e))
    
    try:
        pattern_id = overrides.CutLinePatternId
        if pattern_id != DB.ElementId.InvalidElementId:
            pattern_elem = doc.GetElement(pattern_id)
            info["CutLinePattern"] = pattern_elem.Name if pattern_elem else "ID: {}".format(get_element_id_value(pattern_id))
        else:
            info["CutLinePattern"] = "Default"
    except Exception as e:
        info["CutLinePattern"] = "Errore: {}".format(str(e))
    
    try:
        pattern_id = overrides.SurfaceForegroundPatternId
        if pattern_id != DB.ElementId.InvalidElementId:
            pattern_elem = doc.GetElement(pattern_id)
            info["SurfaceForegroundPattern"] = pattern_elem.Name if pattern_elem else "ID: {}".format(get_element_id_value(pattern_id))
        else:
            info["SurfaceForegroundPattern"] = "Non definito"
    except Exception as e:
        info["SurfaceForegroundPattern"] = "Errore: {}".format(str(e))
    
    try:
        color = overrides.SurfaceForegroundPatternColor
        info["SurfaceForegroundColor"] = color_to_rgb_string(color)
    except Exception as e:
        info["SurfaceForegroundColor"] = "Errore: {}".format(str(e))
    
    return info


def print_error_report(errors_list):
    """Stampa un report dettagliato degli errori."""
    
    if not errors_list:
        return
    
    output.print_md("---")
    output.print_md("# Report Dettagliato Errori")
    output.print_md("")
    
    for i, error in enumerate(errors_list, 1):
        output.print_md("## Errore {}: {}".format(i, error["filter_name"]))
        output.print_md("- **Template:** {}".format(error["template_name"]))
        output.print_md("- **Filter ID:** {}".format(error["filter_id"]))
        output.print_md("- **Tipo errore:** {}".format(error["error_type"] or "N/A"))
        output.print_md("- **Messaggio:** {}".format(error["message"]))
        
        if error["override_info"]:
            output.print_md("")
            output.print_md("### Stato Override del Filtro:")
            for key, value in error["override_info"].items():
                output.print_md("- **{}:** {}".format(key, value))
        
        if error["traceback"]:
            output.print_md("")
            output.print_md("### Traceback:")
            output.print_md("```")
            output.print_md(error["traceback"])
            output.print_md("```")
        
        output.print_md("")


# ============================================================
# FUNZIONI ELABORAZIONE
# ============================================================

def process_filter(view_template, filter_id, filter_name, settings):
    """Elabora un singolo filtro secondo le impostazioni."""
    
    error_details = {
        "filter_name": filter_name,
        "template_name": view_template.Name,
        "filter_id": get_element_id_value(filter_id),
        "success": False,
        "message": "",
        "error_type": None,
        "traceback": None,
        "override_info": None
    }
    
    try:
        overrides = view_template.GetFilterOverrides(filter_id)
        error_details["override_info"] = get_filter_override_info(overrides)
        
        # Determina direzione copia
        if settings["direction"] == "line_to_surface":
            source_color, color_source = get_line_color_from_overrides(overrides)
            if source_color is None:
                error_details["message"] = "Nessun colore linea definito"
                error_details["error_type"] = "NO_SOURCE_COLOR"
                return error_details
        else:  # surface_to_line
            source_color, color_source = get_surface_color_from_overrides(overrides)
            if source_color is None:
                error_details["message"] = "Nessun colore superficie definito"
                error_details["error_type"] = "NO_SOURCE_COLOR"
                return error_details
        
        new_overrides = DB.OverrideGraphicSettings(overrides)
        
        # Applica colore in base alla direzione
        if settings["direction"] == "line_to_surface":
            # Linea -> Superficie
            if settings["surface_pattern_id"]:
                new_overrides.SetSurfaceForegroundPatternId(settings["surface_pattern_id"])
                new_overrides.SetSurfaceForegroundPatternColor(source_color)
                new_overrides.SetCutForegroundPatternId(settings["surface_pattern_id"])
                new_overrides.SetCutForegroundPatternColor(source_color)
            
            # Gestione colore linea
            if settings["line_color_option"] == "reset":
                invalid_color = DB.Color.InvalidColorValue
                new_overrides.SetProjectionLineColor(invalid_color)
                new_overrides.SetCutLineColor(invalid_color)
            elif settings["line_color_option"] == "new_color" and settings["new_line_color"]:
                new_overrides.SetProjectionLineColor(settings["new_line_color"])
                new_overrides.SetCutLineColor(settings["new_line_color"])
            
            # Gestione pattern linea
            if settings["line_pattern_option"] == "reset":
                invalid_pattern = DB.ElementId.InvalidElementId
                new_overrides.SetProjectionLinePatternId(invalid_pattern)
                new_overrides.SetCutLinePatternId(invalid_pattern)
        
        else:  # surface_to_line
            # Superficie -> Linea
            new_overrides.SetProjectionLineColor(source_color)
            new_overrides.SetCutLineColor(source_color)
            
            # Gestione colore superficie
            if settings["surface_color_option"] == "reset":
                invalid_color = DB.Color.InvalidColorValue
                new_overrides.SetSurfaceForegroundPatternColor(invalid_color)
                new_overrides.SetCutForegroundPatternColor(invalid_color)
            elif settings["surface_color_option"] == "new_color" and settings["new_surface_color"]:
                new_overrides.SetSurfaceForegroundPatternColor(settings["new_surface_color"])
                new_overrides.SetCutForegroundPatternColor(settings["new_surface_color"])
            
            # Gestione pattern superficie
            if settings["surface_pattern_option"] == "reset":
                invalid_pattern = DB.ElementId.InvalidElementId
                new_overrides.SetSurfaceForegroundPatternId(invalid_pattern)
                new_overrides.SetCutForegroundPatternId(invalid_pattern)
        
        view_template.SetFilterOverrides(filter_id, new_overrides)
        
        error_details["success"] = True
        error_details["message"] = "{} da {}".format(color_to_rgb_string(source_color), color_source)
        return error_details
        
    except Exception as e:
        error_details["message"] = str(e)
        error_details["error_type"] = type(e).__name__
        error_details["traceback"] = traceback.format_exc()
        return error_details


# ============================================================
# INTERFACCIA UTENTE
# ============================================================

def select_preset_action():
    """Permette all'utente di scegliere se usare un preset."""
    
    presets = get_preset_files()
    
    options = ["Nuova configurazione"]
    if presets:
        options.append("Carica preset esistente")
        options.append("Elimina preset")
    
    selected = forms.SelectFromList.show(
        options,
        title="Gestione Preset",
        button_name="Continua",
        multiselect=False
    )
    
    if not selected:
        return None, None
    
    if selected == "Nuova configurazione":
        return "new", None
    elif selected == "Carica preset esistente":
        preset_name = forms.SelectFromList.show(
            presets,
            title="Seleziona Preset",
            button_name="Carica",
            multiselect=False
        )
        if preset_name:
            return "load", preset_name
        return None, None
    elif selected == "Elimina preset":
        preset_name = forms.SelectFromList.show(
            presets,
            title="Seleziona Preset da Eliminare",
            button_name="Elimina",
            multiselect=False
        )
        if preset_name:
            if forms.alert("Sei sicuro di voler eliminare il preset '{}'?".format(preset_name), yes=True, no=True):
                delete_preset(preset_name)
                forms.alert("Preset eliminato.")
        return None, None
    
    return None, None


def get_user_settings(loaded_preset=None):
    """Raccoglie tutte le impostazioni dall'utente."""
    
    settings = {}
    
    # 1. Selezione View Template (multipla)
    templates = get_all_view_templates()
    if not templates:
        forms.alert("Nessun View Template trovato nel progetto.", exitscript=True)
        return None
    
    templates_dict = {t.Name: t for t in templates}
    
    default_templates = []
    if loaded_preset and "template_names" in loaded_preset:
        default_templates = [n for n in loaded_preset["template_names"] if n in templates_dict]
    
    selected_template_names = forms.SelectFromList.show(
        sorted(templates_dict.keys()),
        title="Seleziona View Template (selezione multipla)",
        button_name="Seleziona",
        multiselect=True
    )
    
    if not selected_template_names:
        return None
    
    selected_templates = [templates_dict[name] for name in selected_template_names]
    settings["templates"] = selected_templates
    settings["template_names"] = selected_template_names
    
    # 2. Recupera filtri dai template selezionati
    filters_dict = get_filters_from_templates(selected_templates)
    
    if not filters_dict:
        forms.alert("I View Template selezionati non hanno filtri applicati.", exitscript=True)
        return None
    
    # 3. Selezione filtri
    selected_filter_names = forms.SelectFromList.show(
        sorted(filters_dict.keys()),
        title="Seleziona Filtri",
        button_name="Continua",
        multiselect=True
    )
    
    if not selected_filter_names:
        return None
    
    settings["filters_dict"] = filters_dict
    settings["filter_names"] = selected_filter_names
    
    # 4. Mostra anteprima colori?
    show_preview = forms.alert(
        "Vuoi visualizzare l'anteprima dei colori attuali?",
        yes=True, no=True
    )
    
    if show_preview:
        show_color_preview(filters_dict, selected_filter_names, selected_templates)
    
    # 5. Direzione copia
    direction_options = [
        "Linea -> Superficie (copia colore linea alla superficie)",
        "Superficie -> Linea (copia colore superficie alla linea)"
    ]
    
    default_direction = 0
    if loaded_preset and loaded_preset.get("direction") == "surface_to_line":
        default_direction = 1
    
    selected_direction = forms.SelectFromList.show(
        direction_options,
        title="Direzione Copia Colore",
        button_name="Continua",
        multiselect=False
    )
    
    if not selected_direction:
        return None
    
    settings["direction"] = "line_to_surface" if selected_direction == direction_options[0] else "surface_to_line"
    
    # Opzioni diverse in base alla direzione
    if settings["direction"] == "line_to_surface":
        # 6a. Pattern superficie
        patterns_dict, solid_pattern_name = get_all_fill_patterns()
        
        if not patterns_dict:
            forms.alert("Nessun pattern di riempimento trovato nel progetto.", exitscript=True)
            return None
        
        pattern_names = sorted(patterns_dict.keys())
        if solid_pattern_name and solid_pattern_name in pattern_names:
            pattern_names.remove(solid_pattern_name)
            pattern_names.insert(0, solid_pattern_name + " (Solido)")
            patterns_dict[solid_pattern_name + " (Solido)"] = patterns_dict.pop(solid_pattern_name)
        
        selected_pattern_name = forms.SelectFromList.show(
            pattern_names,
            title="Seleziona Pattern Superficie",
            button_name="Continua",
            multiselect=False
        )
        
        if not selected_pattern_name:
            return None
        
        settings["surface_pattern_id"] = patterns_dict[selected_pattern_name]
        settings["surface_pattern_name"] = selected_pattern_name
        
        # 7a. Opzioni colore linea
        line_color_options = [
            "Mantieni colore linea originale",
            "Resetta colore linea (default)",
            "Applica nuovo colore RGB alla linea"
        ]
        
        selected_line_color_option = forms.SelectFromList.show(
            line_color_options,
            title="Opzioni Colore Linea",
            button_name="Continua",
            multiselect=False
        )
        
        if not selected_line_color_option:
            return None
        
        settings["new_line_color"] = None
        if selected_line_color_option == line_color_options[0]:
            settings["line_color_option"] = "keep"
        elif selected_line_color_option == line_color_options[1]:
            settings["line_color_option"] = "reset"
        else:
            settings["line_color_option"] = "new_color"
            rgb_input = forms.ask_for_string(
                prompt="Inserisci il colore RGB (es: 255,0,0 per rosso):",
                title="Nuovo Colore RGB Linea",
                default="128,128,128"
            )
            
            if not rgb_input:
                return None
            
            settings["new_line_color"] = parse_rgb_input(rgb_input)
            settings["new_line_color_str"] = rgb_input
            
            if settings["new_line_color"] is None:
                forms.alert("Formato RGB non valido.", exitscript=True)
                return None
        
        # 8a. Opzioni pattern linea
        line_pattern_options = [
            "Mantieni pattern linea originale",
            "Resetta pattern linea (default)"
        ]
        
        selected_line_pattern_option = forms.SelectFromList.show(
            line_pattern_options,
            title="Opzioni Pattern Linea",
            button_name="Continua",
            multiselect=False
        )
        
        if not selected_line_pattern_option:
            return None
        
        settings["line_pattern_option"] = "keep" if selected_line_pattern_option == line_pattern_options[0] else "reset"
        
        # Non usati in questa direzione
        settings["surface_color_option"] = "keep"
        settings["surface_pattern_option"] = "keep"
        settings["new_surface_color"] = None
    
    else:  # surface_to_line
        # 6b. Opzioni colore superficie
        surface_color_options = [
            "Mantieni colore superficie originale",
            "Resetta colore superficie (default)",
            "Applica nuovo colore RGB alla superficie"
        ]
        
        selected_surface_color_option = forms.SelectFromList.show(
            surface_color_options,
            title="Opzioni Colore Superficie",
            button_name="Continua",
            multiselect=False
        )
        
        if not selected_surface_color_option:
            return None
        
        settings["new_surface_color"] = None
        if selected_surface_color_option == surface_color_options[0]:
            settings["surface_color_option"] = "keep"
        elif selected_surface_color_option == surface_color_options[1]:
            settings["surface_color_option"] = "reset"
        else:
            settings["surface_color_option"] = "new_color"
            rgb_input = forms.ask_for_string(
                prompt="Inserisci il colore RGB (es: 255,0,0 per rosso):",
                title="Nuovo Colore RGB Superficie",
                default="128,128,128"
            )
            
            if not rgb_input:
                return None
            
            settings["new_surface_color"] = parse_rgb_input(rgb_input)
            settings["new_surface_color_str"] = rgb_input
            
            if settings["new_surface_color"] is None:
                forms.alert("Formato RGB non valido.", exitscript=True)
                return None
        
        # 7b. Opzioni pattern superficie
        surface_pattern_options = [
            "Mantieni pattern superficie originale",
            "Resetta pattern superficie (default)"
        ]
        
        selected_surface_pattern_option = forms.SelectFromList.show(
            surface_pattern_options,
            title="Opzioni Pattern Superficie",
            button_name="Continua",
            multiselect=False
        )
        
        if not selected_surface_pattern_option:
            return None
        
        settings["surface_pattern_option"] = "keep" if selected_surface_pattern_option == surface_pattern_options[0] else "reset"
        
        # Non usati in questa direzione
        settings["line_color_option"] = "keep"
        settings["line_pattern_option"] = "keep"
        settings["new_line_color"] = None
        settings["surface_pattern_id"] = None
        settings["surface_pattern_name"] = None
    
    return settings


def ask_save_preset(settings):
    """Chiede all'utente se vuole salvare le impostazioni come preset."""
    
    if not PRESET_FOLDER:
        return
    
    save = forms.alert(
        "Vuoi salvare queste impostazioni come preset per uso futuro?",
        yes=True, no=True
    )
    
    if not save:
        return
    
    preset_name = forms.ask_for_string(
        prompt="Nome del preset:",
        title="Salva Preset",
        default="MioPreset"
    )
    
    if not preset_name:
        return
    
    # Prepara dati per il salvataggio (senza oggetti Revit)
    preset_data = {
        "direction": settings["direction"],
        "line_color_option": settings.get("line_color_option"),
        "line_pattern_option": settings.get("line_pattern_option"),
        "surface_color_option": settings.get("surface_color_option"),
        "surface_pattern_option": settings.get("surface_pattern_option"),
        "surface_pattern_name": settings.get("surface_pattern_name"),
        "new_line_color_str": settings.get("new_line_color_str"),
        "new_surface_color_str": settings.get("new_surface_color_str")
    }
    
    if save_preset(preset_name, preset_data):
        forms.alert("Preset '{}' salvato con successo!".format(preset_name))
    else:
        forms.alert("Errore nel salvataggio del preset.")


# ============================================================
# MAIN
# ============================================================

def main():
    """Funzione principale dello script."""
    
    # Gestione preset
    action, preset_name = select_preset_action()
    
    if action is None:
        script.exit()
    
    loaded_preset = None
    if action == "load" and preset_name:
        loaded_preset = load_preset(preset_name)
        if loaded_preset:
            forms.alert("Preset '{}' caricato. Alcune impostazioni saranno precompilate.".format(preset_name))
    
    # Raccogli impostazioni
    settings = get_user_settings(loaded_preset)
    
    if not settings:
        script.exit()
    
    # Stampa riepilogo
    output.print_md("# Elaborazione Filtri")
    output.print_md("**View Templates:** {}".format(", ".join(settings["template_names"])))
    output.print_md("**Filtri selezionati:** {}".format(len(settings["filter_names"])))
    output.print_md("**Direzione:** {}".format(
        "Linea -> Superficie" if settings["direction"] == "line_to_surface" else "Superficie -> Linea"))
    
    if settings["direction"] == "line_to_surface":
        output.print_md("**Pattern superficie:** {}".format(settings.get("surface_pattern_name", "N/A")))
        output.print_md("**Opzione colore linea:** {}".format(settings["line_color_option"]))
        output.print_md("**Opzione pattern linea:** {}".format(settings["line_pattern_option"]))
    else:
        output.print_md("**Opzione colore superficie:** {}".format(settings["surface_color_option"]))
        output.print_md("**Opzione pattern superficie:** {}".format(settings["surface_pattern_option"]))
    
    output.print_md("---")
    
    # Elaborazione
    success_count = 0
    error_count = 0
    errors_list = []
    
    with revit.Transaction("Gestione colori filtri"):
        for filter_name in settings["filter_names"]:
            filter_info = settings["filters_dict"][filter_name]
            filter_id = filter_info["id"]
            
            for template in filter_info["templates"]:
                if template not in settings["templates"]:
                    continue
                
                result = process_filter(template, filter_id, filter_name, settings)
                
                if result["success"]:
                    output.print_md("OK **{}** [{}]: {}".format(
                        filter_name, template.Name, result["message"]))
                    success_count += 1
                else:
                    output.print_md("ERRORE **{}** [{}]: {}".format(
                        filter_name, template.Name, result["message"]))
                    errors_list.append(result)
                    error_count += 1
    
    # Riepilogo
    output.print_md("---")
    output.print_md("## Riepilogo")
    output.print_md("- Operazioni completate con successo: **{}**".format(success_count))
    if error_count > 0:
        output.print_md("- Operazioni con errori: **{}**".format(error_count))
    
    # Report errori
    if errors_list:
        print_error_report(errors_list)
    
    # Chiedi se salvare preset
    ask_save_preset(settings)
    
    # Messaggio finale
    if error_count > 0:
        forms.alert(
            "Elaborazione completata con errori.\n\n"
            "Operazioni riuscite: {}\n"
            "Errori: {}\n\n"
            "Consulta il report per i dettagli.".format(success_count, error_count),
            title="Completato con Errori"
        )
    else:
        forms.alert(
            "Elaborazione completata!\n\n"
            "Operazioni riuscite: {}".format(success_count),
            title="Completato"
        )


if __name__ == "__main__":
    main()