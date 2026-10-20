# -*- coding: utf-8 -*-
__title__ = "Quick\nClash"

__doc__ = """Version = 1.0
Date    = 20.10.2026
_____________________________________________________________________
Compact version of Clash Setup: only the Run part.

Click a saved test, check or change where Set A and Set B come from
(current Revit selection, active view, whole model or nothing, plus
loaded links) for this run, then run. The report is the same as Clash
Setup's.

The current selection is read again at every Refresh and when the run
starts, so you can change it in Revit with the window open.

Tests and settings are shared with Clash Setup, where the tests are
created and edited.
_____________________________________________________________________
Author(s): Claude + Matteo Maiorano
"""

__author__ = "Claude + Matteo Maiorano"

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
    forms.alert("Quick Clash needs Revit 2022 or later.", exitscript=True)

# Moduli condivisi con Clash Setup in ClashFinder.pulldown/lib (nel percorso di ricerca
# di tutti i pulsanti dello stack).
from clashfinder_bridge import RevitBridge
import clashfinder_quick_ui

bridge = RevitBridge("Quick Clash")
try:
    clashfinder_quick_ui.show_quick_window(doc, bridge, __revit__.MainWindowHandle)  # noqa: F821
except Exception:
    bridge.dispose()
    raise
