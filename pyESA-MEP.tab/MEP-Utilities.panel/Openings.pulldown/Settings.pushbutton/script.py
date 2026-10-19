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
import sys

from pyrevit import revit, script, DB

import Settings_ui

# nomi, parametri e impostazioni predefinite: una sola copia, in openings_modules/openings_core.py
MODULES = op.join(op.dirname(op.dirname(op.abspath(__file__))), "openings_modules")
if MODULES not in sys.path:
    sys.path.insert(0, MODULES)
from openings_core import CONFIG_NAME, DEFAULTS, PARAM_WIDTH, PARAM_HEIGHT, PARAM_DEPTH, PARAM_DIAMETER, VALID, \
    TO_REFRESH, NOT_ASSOCIATED, load_settings  # noqa: E402

XAML_FILE_NAME = "Settings_form.xaml"

doc = revit.doc


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
                   "removed or moved out, opening without elements inside).\n"
                   "Opening: related Ids = its MEP elements, those of this model first, then one group per linked "
                   "model: \"111,222; PRJ_MEC: 333,444\"; related categories = Pipes, Ducts, Cable Trays, Conduits.\n"
                   "MEP elements are not written: the link is on the opening.").format(
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
