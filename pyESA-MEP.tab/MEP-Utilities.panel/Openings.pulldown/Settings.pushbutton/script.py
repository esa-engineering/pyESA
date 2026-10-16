# -*- coding: utf-8 -*-
__title__ = "Settings"
__doc__ = """Version = 2.0
Date    = 15.10.2026
________________________________________________________________
Description:
Settings of Openings Place and Check&Report, saved per user: the four
families (wall / floor, rectangular / round), the names of the status,
related element Id and related element category parameters, clearance
per side, rounding, merge distance, extra depth beyond each face,
separate MEC and ELE openings, To Be Refreshed status in Check&Report.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"
__context__ = "zero-doc"

import os.path as op

from pyrevit import revit, script, DB

import Settings_ui

XAML_FILE_NAME = "Settings_form.xaml"

# =================================================================================================================
# Blocco comune a Place, Check&Report e Settings (nomi, parametri, impostazioni predefinite): tenerlo allineato
# =================================================================================================================
CONFIG_NAME = "ESA_Openings"
DEFAULTS = {"fam_wall_rect": "e_GM.PH.VD_Void Demand_Wall_Rect", "fam_wall_round": "e_GM.PH.VD_Void Demand_Wall_Circ",
            "fam_floor_rect": "e_GM.PH.VD_Void Demand_Floor_Rect",
            "fam_floor_round": "e_GM.PH.VD_Void Demand_Floor_Circ",
            "status_param": "e_GEN_Void_Status", "related_id_param": "e_GEN_RelatedElement_ID",
            "related_cat_param": "e_GEN_RelatedElement_Cat",
            "clearance_mm": 25.0, "rounding_mm": 50.0, "merge_mm": 100.0, "extra_depth_mm": 0.0, "layer_gap_mm": 200.0,
            "min_angle_deg": 15.0,
            "separate": False, "misaligned": True}
NUMBER_KEYS = ("clearance_mm", "rounding_mm", "merge_mm", "extra_depth_mm", "layer_gap_mm", "min_angle_deg")
BOOL_KEYS = ("separate", "misaligned")
TEXT_KEYS = ("status_param", "related_id_param", "related_cat_param")  # vuoti: tornano al predefinito
PARAM_WIDTH = "e_DIM_Width_1"
PARAM_HEIGHT = "e_DIM_Height_1"
PARAM_DEPTH = "e_DIM_Depth_1"
PARAM_DIAMETER = "e_DIM_Diameter_1"
# stati scritti in e_GEN_Void_Status, dal migliore al peggiore
VALID, TO_REFRESH, NOT_ASSOCIATED = "Valid", "To Be Refreshed", "Not Associated"
# =================================================================================================================

doc = revit.doc


def load_settings(config):
    settings = dict(DEFAULTS)
    try:
        for key in DEFAULTS:
            value = config.get_option(key, DEFAULTS[key])
            if key in NUMBER_KEYS:
                value = float(value)
            elif key in BOOL_KEYS and not isinstance(value, bool):
                value = str(value).strip().lower() in ("true", "1", "yes")
            elif key in TEXT_KEYS:
                value = str(value or "").strip() or DEFAULTS[key]
            settings[key] = value
    except Exception:
        pass
    return settings


def generic_model_families():
    if doc is None:
        return []
    names = set()
    for symbol in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(
            DB.BuiltInCategory.OST_GenericModel):
        names.add(symbol.Family.Name)
    return sorted(names)


def main():
    config = script.get_config(CONFIG_NAME)
    settings = load_settings(config)
    params_text = ("Size parameters (instance, length, fixed names): {}, {}, {}, {}. Wall, round: cylinder along X, "
                   "Width = length through the wall. Floor, round: cylinder along Z, Height = length through the "
                   "floor. Rectangular: Depth = length through.\n"
                   "Status: {} (elements in place), {} (elements moved but still inside the opening), {} (elements "
                   "removed or moved out, opening without elements, element without opening).\n"
                   "Opening: related Ids = Ids of its MEP elements, related categories = Pipes, Ducts, Cable Trays, "
                   "Conduits.\nMEP element: related Ids = Ids of its openings, related categories = Generic Models, "
                   "status = the worst of its openings.").format(
        PARAM_WIDTH, PARAM_HEIGHT, PARAM_DEPTH, PARAM_DIAMETER, VALID, TO_REFRESH, NOT_ASSOCIATED)
    xaml = script.get_bundle_file(XAML_FILE_NAME) or op.join(op.dirname(__file__), XAML_FILE_NAME)
    result = Settings_ui.SettingsWindow(xaml, settings, DEFAULTS, generic_model_families(), params_text).show()
    if result is None:
        script.exit()
    try:
        for key in DEFAULTS:
            setattr(config, key, result[key])
        script.save_config()
    except Exception:
        pass


main()
