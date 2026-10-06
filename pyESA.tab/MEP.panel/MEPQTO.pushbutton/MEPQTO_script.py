# -*- coding: utf-8 -*-
__title__ = "MEP\nQTO"

__doc__ = """Version = 1.3
Date    = 06.10.2026
_____________________________________________________________________
Quantity takeoff of the MEP model, driven by the Type Mark.

Counted categories (1 instance = 1 piece): Air Terminals,
Communication / Data / Fire Alarm / Nurse Call / Security Devices,
Conduit Fittings, Duct and Pipe Accessories, Electrical Equipment and
Fixtures, Lighting Devices and Fixtures, Mechanical Equipment,
Plumbing Fixtures, Sprinklers, Specialty Equipment.

Before reading, a window asks which Revit links to read besides the
open model (every placed link instance is counted), which worksets
to exclude (by name, in every model read) and which categories.
Category and workset selections can be saved as named sets in the
pyRevit settings. A link is read in the phase with the same name; a
link without it is skipped and listed in the Issues tab.

Measured categories: Ducts, Flex Ducts, Pipes, Flex Pipes (m, mq or
kg, from the unit of the price code), Cable Trays and Conduits
(always m), Duct and Pipe Insulation (mq or m). Formulas, sheet
weights, pipe densities and the allowance for fittings come from the
company QTO sheets and can be edited in the Rules tab.

Price codes: every category reads instance parameters (default
e_DAT_PriceCode_i_1..10). Piece categories also read type parameters
(default e_DAT_PriceCode_1..10, instance parameters with the same name
as a fallback) and count both sets; linear categories read only the
instance ones, because the same type changes price item with its
size. A Yes/No parameter leaves elements out of the bill:
e_DAT_BOQ_i on the instances of every category (only an explicit No
excludes). All of them can be remapped with
Parameters... and are stored in the project file.

Descriptions, units and unit prices are NOT stored in the model: they
come from a shared price list (.json, edited with the price list
editor - Edit... button - or .xlsx / .csv, read only) and from a
project file <Model>_MEPQTO.json next to the central model, which
keeps everything typed in the window.

Tabs: EPU (editable: chapter, subchapter, EPU item No.,
reference price book, short description, description, unit from a
drop-down list, unit price; the price book code is the code read from
the model), Bill of
quantities (filled automatically, grouped by WBS when WBS levels are
set; right click or Allowance Override... sets the allowance of the
selected items, whose quantity is then highlighted), Type Marks (type
and instance price codes 1..10 and their descriptions for each family
type), Rules (editable), Issues.

WBS Levels...: up to 15 levels, one parameter each (read on the
instance, then the type, then the host). Levels left empty are
skipped. Save writes the project file, Export Excel writes the same
five tabs to a workbook, one sheet each.

The model is only read, never modified.
_____________________________________________________________________
Author(s): Claude + Andrea Patti
"""

__author__ = "Claude + Andrea Patti"

from pyrevit import revit, script, forms

import mepqto_model as qm
from mepqto_ui import show_takeoff_window

doc = revit.doc


def main():
    if doc is None or doc.IsFamilyDocument:
        forms.alert("Open a project model to run the takeoff.", title="MEP QTO")
        script.exit()
    if not qm.available_rules():
        forms.alert("None of the takeoff categories exists in this Revit version.",
                    title="MEP QTO")
        script.exit()

    # Nessun report alla chiusura: computo, anomalie ed export stanno nella finestra.
    # Se l'utente annulla la scelta di modelli e categorie la finestra non si apre.
    show_takeoff_window(doc)


main()
