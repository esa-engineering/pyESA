# -*- coding: utf-8 -*-
__title__ = "Settings"
__doc__ = """Version = 3.0
Date    = 19.10.2026
________________________________________________________________
Description:
Settings of Openings, saved per user: clearance, rounding, merge
distance, extra depth, layers joined across a gap, separate MEC and
ELE openings. Shows whether the four opening families are loaded.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"
__context__ = "zero-doc"

import os.path as op
import sys

from pyrevit import revit, script, DB

import Settings_ui

# nomi e impostazioni predefinite: una sola copia, in openings_modules/openings_core.py
MODULES = op.join(op.dirname(op.dirname(op.abspath(__file__))), "openings_modules")
if MODULES not in sys.path:
    sys.path.insert(0, MODULES)
from openings_core import CONFIG_NAME, DEFAULTS, NUMBER_KEYS, BOOL_KEYS, load_settings  # noqa: E402

XAML_FILE_NAME = "Settings_form.xaml"
FAMILY_CASES = (("Wall, rectangular", "fam_wall_rect"), ("Wall, round", "fam_wall_round"),
                ("Floor, rectangular", "fam_floor_rect"), ("Floor, round", "fam_floor_round"))

doc = revit.doc


def family_status():
    """[(caso, famiglia, caricata)]: caricata None se non c'e' un modello aperto."""
    loaded = None
    if doc is not None:
        loaded = set(symbol.Family.Name for symbol in DB.FilteredElementCollector(doc).OfClass(
            DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_GenericModel))
    return [(case, DEFAULTS[key], None if loaded is None else DEFAULTS[key] in loaded)
            for case, key in FAMILY_CASES]


def main():
    config = script.get_config(CONFIG_NAME)
    settings = load_settings(config)
    xaml = script.get_bundle_file(XAML_FILE_NAME) or op.join(op.dirname(__file__), XAML_FILE_NAME)
    result = Settings_ui.SettingsWindow(xaml, settings, family_status()).show()
    if result is None:
        script.exit()
    try:
        for key in NUMBER_KEYS + BOOL_KEYS:
            setattr(config, key, result[key])
        script.save_config()
    except Exception:
        pass


main()
