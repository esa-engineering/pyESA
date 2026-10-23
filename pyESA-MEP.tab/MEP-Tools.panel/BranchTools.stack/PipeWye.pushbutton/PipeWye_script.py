# -*- coding: utf-8 -*-
__title__ = "Pipe\nWye"
__doc__ = """Version = 0.14.0
Date    = 23.10.2026
Places a wye on a pipe, or connects another pipe to it from above or from the side with a 45 degree wye.
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

# ======================= STORICO REVISIONI =======================
# Versione | Data       | Tipo | Descrizione                                   | PR
# 0.14.0   | 23.10.2026 | FEAT | Diviso da Fitting Tool 0.13.0 (sezione Pipe); storico precedente in lib/BranchTools_DeveloperNote.md | -
# Testato su Revit: non ancora testato
# =================================================================

from Autodesk.Revit.DB import TransactionGroup
from Autodesk.Revit.DB.Plumbing import Pipe

from pyrevit import revit, script, forms

import BranchTools_Pipe as fpp
import BranchTools_Select as sel
import PipeWye_Form_ui as ui

XAML_FILE_NAME = "PipeWye_Form.xaml"
CONFIG_NAME = "ESA_PipeWye"
TITLE = "Pipe Wye"
NUMBERS = {"wye_roll": ("wye_roll", ui.DEFAULT_SETTINGS["wye_roll"]),
           "wye_angle": ("wye_angle", ui.DEFAULT_SETTINGS["wye_angle"])}
TEXTS = {"wye_family": "wye_family"}

doc = revit.doc
uidoc = revit.uidoc


def is_pipe(element):
    return isinstance(element, Pipe)


def main():
    pipe, point = sel.get_element(uidoc, "Select the main pipe", is_pipe)
    if pipe is None:
        script.exit()
    main_pipe = fpp.MainPipe(doc, uidoc, pipe, point)
    if not main_pipe.ok:
        forms.alert("This pipe cannot be used: {}".format(main_pipe.reason), title=TITLE, exitscript=True)
    config = script.get_config(CONFIG_NAME)
    settings = sel.load_settings(config, NUMBERS, TEXTS)
    if not settings["wye_family"]:
        # prima volta: una famiglia con Part Type Wye, se c'e'; altrimenti restano le routing preferences
        symbol = fpp.wye_symbol(doc, pipe)
        if symbol is not None:
            settings["wye_family"] = fpp.symbol_label(symbol)

    window = ui.PipeWyeWindow(script.get_bundle_file(XAML_FILE_NAME), main_pipe,
                              fpp.pipe_fitting_symbols(doc), settings)
    # quello che si fa nella finestra e' una sola voce del menu Annulla; Cancel annulla tutto
    group = TransactionGroup(doc, TITLE)
    group.Start()
    try:
        result = window.show()
        if result is None:
            group.RollBack()
            script.exit()
        group.Assimilate()
    finally:
        if group.HasStarted() and not group.HasEnded():
            group.RollBack()
    sel.save_settings(script, config, result, NUMBERS, TEXTS)

    # collegamenti: l'altro tubo si clicca dopo la chiusura della finestra
    if result["action"] in ("connect_above", "connect_side"):
        other, picked = sel.pick(uidoc, "Pick the pipe to connect, near its open end", is_pipe, pipe.Id)
        if other is not None:
            error = main_pipe.connect(other, picked or main_pipe.point, result["action"] == "connect_above")
            if error:
                forms.alert("The pipe was not connected.\n{}".format(error), title=TITLE)


main()
