# -*- coding: utf-8 -*-
__title__ = "Fitting\nRotate"
__doc__ = """Version = 0.14.0
Date    = 23.10.2026
Rotates a fitting and sets its branch angle with sliders; aligns and connects the open end of a pipe or fitting.
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

# ======================= STORICO REVISIONI =======================
# Versione | Data       | Tipo | Descrizione                                   | PR
# 0.14.0   | 23.10.2026 | FEAT | Diviso da Fitting Tool 0.13.0 (Rotate, Branch angle, Align); storico precedente in lib/BranchTools_DeveloperNote.md | -
# Testato su Revit: non ancora testato
# =================================================================

from Autodesk.Revit.DB import TransactionGroup

from pyrevit import revit, script, forms

import BranchTools_Align as fa
import BranchTools_Rotate as fr
import BranchTools_Select as sel
import FittingRotate_Form_ui as ui

XAML_FILE_NAME = "FittingRotate_Form.xaml"
CONFIG_NAME = "ESA_FittingRotate"
TITLE = "Fitting Rotate"
NUMBERS = {"rotate": ("step", ui.DEFAULT_SETTINGS["rotate"]),
           "angle": ("angle_step", ui.DEFAULT_SETTINGS["angle"])}

doc = revit.doc
uidoc = revit.uidoc


def main():
    element, _ = sel.get_element(uidoc, "Select one fitting, or a pipe to align", fa.is_alignable)
    if element is None:
        script.exit()
    fittings = [element] if fr.is_fitting(element) else []
    rotator = fr.Rotator(doc, uidoc, fittings)
    angle_setter = fr.AngleSetter(doc, uidoc, fittings)
    aligner = fa.Aligner(doc, uidoc, element)
    config = script.get_config(CONFIG_NAME)

    window = ui.FittingRotateWindow(script.get_bundle_file(XAML_FILE_NAME), rotator, angle_setter, aligner,
                                    sel.load_settings(config, NUMBERS))
    # rotazioni e cambi di angolo sono una sola voce del menu Annulla; Cancel li annulla tutti
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
    sel.save_settings(script, config, result, NUMBERS)

    # allineamento: l'elemento a cui allinearsi si clicca dopo la chiusura della finestra
    if result["action"] == "align":
        target, _ = sel.pick(uidoc, "Pick the pipe or fitting to line up with", fa.is_alignable, element.Id)
        if target is not None:
            error = aligner.run_with(target)
            if error:
                forms.alert(error, title=TITLE)


main()
