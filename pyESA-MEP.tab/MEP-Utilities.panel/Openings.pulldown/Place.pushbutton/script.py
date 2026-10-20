# -*- coding: utf-8 -*-
__title__ = "Place&Check"
__doc__ = """Version = 5.0
Date    = 19.10.2026
________________________________________________________________
Description:
Openings (void demand): Place&Check, New Report and Last Report share
the same script; only MODE changes.

PLACE&CHECK: openings where pipes, ducts, cable trays and conduits (this
model and RVT links) cross walls, floors and structural elements. Walls
and structures: any angle; floors: only elements steeper than 85
degrees. Size = section with insulation + clearance, rounded up.
NEW REPORT: status of the existing openings (Valid, To Be Refreshed,
Not Associated) in a scope, written on the openings.
LAST REPORT: the last New Report again, no check.
Sizes and options: Settings button of this menu.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"
# la finestra del report e' non modale e lavora tramite ExternalEvent dopo la fine del comando: il motore
# IronPython deve restare vivo
__persistentengine__ = True

import os.path as op
import sys

MODE = "place"  # "place" | "new" | "last": unica differenza tra Place, New Report e Last Report, tenere allineati

# tutto il codice sta in openings_modules, accanto ai pulsanti (cartella senza suffisso: pyRevit non la carica come
# pulsante); vedi la docstring di openings_core
MODULES = op.join(op.dirname(op.dirname(op.abspath(__file__))), "openings_modules")
if MODULES not in sys.path:
    sys.path.insert(0, MODULES)

import openings_core  # noqa: E402

openings_core.start(MODE)
if MODE == "place":
    import openings_place  # noqa: E402
    openings_place.place_main()
else:
    import openings_report  # noqa: E402
    if MODE == "last":
        openings_report.last_main()
    else:
        openings_report.check_main()
