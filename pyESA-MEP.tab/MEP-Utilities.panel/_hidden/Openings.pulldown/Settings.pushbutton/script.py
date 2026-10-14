# -*- coding: utf-8 -*-
__title__ = "Settings"
__doc__ = """Version = 1.0
Date    = 13.10.2026
________________________________________________________________
Description:
Settings of Openings Place and Refresh, saved per user: the four
families (wall / floor, rectangular / round), clearance per side,
rounding, merge distance, extra depth beyond each face, separate MEC
and ELE openings, Misaligned status in Refresh.
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
# Blocco comune a Place, Refresh e Settings (nomi, parametri, impostazioni predefinite): tenerlo allineato
# =================================================================================================================
CONFIG_NAME = "ESA_Openings"
DEFAULTS = {"fam_wall_rect": "e_GM.PH.PFV_Wall_Rect", "fam_wall_round": "e_GM.PH.PFV_Wall_Round",
            "fam_floor_rect": "e_GM.PH.PFV_Floor_Rect", "fam_floor_round": "e_GM.PH.PFV_Floor_Round",
            "clearance_mm": 25.0, "rounding_mm": 50.0, "merge_mm": 100.0, "extra_depth_mm": 0.0, "layer_gap_mm": 200.0,
            "min_angle_deg": 15.0,
            "separate": False, "misaligned": True}
NUMBER_KEYS = ("clearance_mm", "rounding_mm", "merge_mm", "extra_depth_mm", "layer_gap_mm", "min_angle_deg")
BOOL_KEYS = ("separate", "misaligned")
PARAM_WIDTH = "e_DIM_Width_1"
PARAM_HEIGHT = "e_DIM_Height_1"
PARAM_DEPTH = "e_DIM_Depth_1"
PARAM_DIAMETER = "e_DIM_Diameter_1"
STATUS_PARAM = "e_GEN_PFV_Status"
ID_PARAM = "e_GEN_PFV_ID"
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
    params_text = ("Size parameters (instance, length): {}, {}, {}, {}.\nOpening: {} = Ids of its MEP elements, "
                   "{} = <MEC|ELE|MEC+ELE> - <New|OK|Split|Joined|Misaligned|Changed|Obsolete>.\nMEP element: {} = codes of "
                   "its openings (PFV + Id of the opening), {} = <MEC|ELE> - <status>.").format(
        PARAM_WIDTH, PARAM_HEIGHT, PARAM_DEPTH, PARAM_DIAMETER, ID_PARAM, STATUS_PARAM, ID_PARAM, STATUS_PARAM)
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
