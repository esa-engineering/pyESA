# -*- coding: utf-8 -*-
__title__ = "Clash\nSetup"

__doc__ = """Version = 2.1
Date    = 22.10.2026
_____________________________________________________________________
Clash detection driven by saved tests, in the spirit of Navisworks.

A test compares Set A with Set B. Each set is defined by categories
(by default) and, optionally, by rules on parameter values. Tests are
saved in a personal library and can be exported to / imported from
JSON files to share them.

The window stays open while you work in Revit. Click a saved
test to load it, then edit it for Set A and Set B:
  Sources       - current Revit selection, active view, whole
                  model or links only, plus loaded links (chosen
                  from a drop-down list, saved by file name)
  Categories    - ticked categories, with the number of elements
                  that enter the set
  Model filters - view filters of the model, ANDed with categories
  Custom rules  - rules on parameter values
Run executes the test on screen; Save as stores it with a name.
The selection is read again at every refresh and when the run
starts.

Clashes are found as in Revit's Interference Check (solid
intersection) and listed in a report window as an expandable list:
clicking a row selects and zooms to its elements, optionally fitting
the section box of the active 3D view.

Each set can also take elements from loaded Revit links
(open model vs link and link vs link).
_____________________________________________________________________
Author(s): Claude + Matteo Maiorano
"""

__author__ = "Claude + Matteo Maiorano"

# ======================= STORICO REVISIONI =======================
# Versione | Data       | Tipo | Descrizione                                              | PR
# 2.1.0    | 22.10.2026 | UI   | Link scelti da un menu a discesa; riga del test con      |
#          |            |      | "N links"; etichette "Links" e "Links only"              |
# Testato su Revit: (da compilare dopo le prove)
# =================================================================

# Le finestre restano aperte dopo la fine dello script: il motore IronPython
# deve sopravvivere per gestirne gli eventi.
__persistentengine__ = True

from pyrevit import revit, forms

doc = revit.doc

if doc is None:
    forms.alert("Open a model first.", exitscript=True)
if doc.IsFamilyDocument:
    forms.alert("This command is not available in a family document.", exitscript=True)
if int(doc.Application.VersionNumber) < 2022:
    forms.alert("Clash Setup needs Revit 2022 or later.", exitscript=True)

from clashfinder_bridge import RevitBridge
import clashfinder_ui

bridge = RevitBridge("Clash Setup")
try:
    clashfinder_ui.show_main_window(doc, bridge, __revit__.MainWindowHandle)  # noqa: F821
except Exception:
    bridge.dispose()
    raise
