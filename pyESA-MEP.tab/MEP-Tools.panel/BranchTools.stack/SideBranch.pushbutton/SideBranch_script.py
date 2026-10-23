# -*- coding: utf-8 -*-
__title__ = "Side\nBranch"
__doc__ = """Version = 0.14.0
Date    = 23.10.2026
Makes a side branch from a wye at a set invert offset, repeats it along the main and changes its length.
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

# ======================= STORICO REVISIONI =======================
# Versione | Data       | Tipo | Descrizione                                   | PR
# 0.14.0   | 23.10.2026 | FEAT | Diviso da Fitting Tool 0.13.0 (Side branch, Place copies, Branch length); storico precedente in lib/BranchTools_DeveloperNote.md | -
# Testato su Revit: non ancora testato
# =================================================================

from Autodesk.Revit.DB import TransactionGroup

from pyrevit import revit, script, forms

import BranchTools_Length as fl
import BranchTools_Place as fp
import BranchTools_Rotate as fr
import BranchTools_Select as sel
import BranchTools_SideBranch as fs
import SideBranch_Form_ui as ui

XAML_FILE_NAME = "SideBranch_Form.xaml"
CONFIG_NAME = "ESA_SideBranch"
TITLE = "Side Branch"
# chiave nella finestra -> (opzione nel config, default); "elbow" e' storica: pendenza del tubo inclinato
NUMBERS = dict((key, (option, ui.DEFAULT_SETTINGS[key])) for key, option in (
    ("offset", "branch_offset_mm"), ("elbow", "branch_elbow_mode"), ("length", "branch_length_mm"),
    ("direction", "branch_direction"), ("keep_level", "copy_keep_level"), ("sloped", "sloped_step")))

doc = revit.doc
uidoc = revit.uidoc


def main():
    fitting, _ = sel.get_element(uidoc, "Select the wye on the main", fr.is_pipe_fitting)
    if fitting is None:
        script.exit()
    rotator = fr.Rotator(doc, uidoc, [fitting])
    side_branch = fs.SideBranch(doc, uidoc, rotator)
    branch_length = fl.BranchLength(doc, uidoc, side_branch)
    config = script.get_config(CONFIG_NAME)

    window = ui.SideBranchWindow(script.get_bundle_file(XAML_FILE_NAME), rotator, side_branch, branch_length,
                                 sel.load_settings(config, NUMBERS))
    # stacco, lunghezza e rotazioni sono una sola voce del menu Annulla; Cancel annulla tutto
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

    # copie: una per clic sul collettore, ognuna con la sua voce nel menu Annulla
    if result["action"] == "place":
        placed, failures = fp.place_copies(doc, uidoc, side_branch, result["keep_level"] > 0)
        if failures:
            forms.alert("{} branch(es) not placed, {} placed.\n{}".format(
                len(failures), placed, failures[-1]), title=TITLE)


main()
