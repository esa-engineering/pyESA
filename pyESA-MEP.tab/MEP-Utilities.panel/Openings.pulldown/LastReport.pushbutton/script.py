# -*- coding: utf-8 -*-
__title__ = "Last Report"
__doc__ = """Version = 4.0
Date    = 15.10.2026
________________________________________________________________
Description:
Openings (void demand). Place, New Report and Last Report share the same script; only
MODE changes. Settings (families, parameters, clearance, rounding,
merge distance, extra depth, MEC/ELE, To Be Refreshed) come from the
Settings button of this menu.

PLACE: with elements already selected, those already associated with an
opening (openings, MEP elements listing an existing opening) are
checked; the other MEP elements, walls, floors and structural elements
get openings. With nothing selected, a window asks what to select: walls
and floors (host model or RVT links), MEP elements, or an area (rectangle
dragged in a plan view, or the section box of a 3D view). Every straight pipe, duct, cable tray or conduit of the host
model that crosses them gets a Generic Model opening at the centre of
the crossing, sized with the outer section (insulation included) plus
the clearance, rounded up; the depth is the thickness plus the extra
depth on each side. Crossings of the same wall or floor closer than the
merge distance share one rectangular opening (only within the same
discipline, if MEC and ELE are kept separate).
- Opening: e_GEN_RelatedElement_ID = Ids of its MEP elements (111,222);
  e_GEN_RelatedElement_Cat = their categories (Pipes, Ducts);
  e_GEN_Void_Status = Valid
- MEP elements: e_GEN_RelatedElement_ID = Ids of their openings;
  e_GEN_RelatedElement_Cat = Generic Models; e_GEN_Void_Status = Valid
Parameter names can be changed in Settings. Comments and Mark are not
touched. Crossings that already have an opening are skipped.

NEW REPORT: choose the scope (whole model, MEP selection, ARC
selection, area). Every opening is compared with its MEP elements: still
there, still crossing, in the same place and size? Statuses: Valid,
To Be Refreshed (the elements moved but are still inside the opening;
optional), Not Associated (elements removed or moved out, no element
left; MEP elements left without an opening). Statuses and Ids are
written, stale Ids (deleted elements or openings) removed; the MEP
elements of the report are checked for crossings without an opening.
LAST REPORT: no check; the last New Report again, same scope, with the
statuses in the model.
A summary comes first, then a Text Report (export to Excel) or an
Interactive Report to select the elements, create, recreate, move or
delete openings.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"
# la finestra del report e' non modale e lavora tramite ExternalEvent dopo la fine del comando: il motore
# IronPython deve restare vivo
__persistentengine__ = True

import os.path as op
import sys

MODE = "last"  # "place" | "new" | "last": unica differenza tra Place, New Report e Last Report, tenere allineati

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
