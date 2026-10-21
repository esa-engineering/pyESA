# -*- coding: utf-8 -*-
__title__ = "Last\nReport"

__doc__ = """Version = 1.1
Date    = 20.10.2026
_____________________________________________________________________
Opens again the last Clash Finder report made on this model, with the
same navigation (colours, section box, isolation, export).

Every run of Clash Setup or Quick Clash saves its report in the
report folder of the model (asked the first time, changed with Export
JSON): one file per clash test, earlier ones moved to _old.

CLICK: most recent report in the folder (a file to import if none).
SHIFT + CLICK: import a report file (also from _old).

Elements are found again by id: if the model changed after the run,
some rows may point to elements that no longer exist.
_____________________________________________________________________
Author(s): Claude + Matteo Maiorano
"""

__author__ = "Claude + Matteo Maiorano"

# La finestra del report resta aperta dopo la fine dello script.
__persistentengine__ = True

import os
import traceback

import clr
clr.AddReference('PresentationFramework')
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage

from pyrevit import revit, forms, script

doc = revit.doc

if doc is None:
    forms.alert("Open a model first.", exitscript=True)
if doc.IsFamilyDocument:
    forms.alert("This command is not available in a family document.", exitscript=True)
if int(doc.Application.VersionNumber) < 2022:
    forms.alert("Last Report needs Revit 2022 or later.", exitscript=True)


def _error(text):
    MessageBox.Show(text, "Last Report", MessageBoxButton.OK, MessageBoxImage.Error)


# Con il motore persistente un errore puo' non arrivare al report di pyRevit:
# ogni passo mostra il proprio errore in una MessageBox.
try:
    # Moduli condivisi con Clash Setup in ClashFinder.pulldown/lib.
    import clashfinder_engine as ce
    import clashfinder_rules as cr
    import clashfinder_report_ui as report_ui
except Exception:
    _error(u"Clash Finder modules could not be loaded:\n\n{}".format(traceback.format_exc()))
    script.exit()

# Cartella dei report del modello (scelta al primo run o con Export JSON).
folder = None
try:
    folder = ce.report_folder(doc)
except Exception:
    folder = None
if folder and not os.path.isdir(folder):
    folder = None

path = None
if not __shiftclick__ and folder:  # noqa: F821
    path = ce.latest_report(folder)
if path is None:
    title = ("Import a clash report" if __shiftclick__  # noqa: F821
             else "No report saved for this model: import a clash report")
    path = forms.pick_file(file_ext='json', init_dir=folder or '', title=title)
if not path:
    script.exit()

try:
    result = ce.load_report(path)
except Exception as error:
    _error(u"The file could not be read as a Clash Finder report:\n{}\n\n{}".format(
        path, error))
    script.exit()

try:
    form = report_ui.show_report(doc, result, cr.category_display_names(doc),
                                 __revit__.MainWindowHandle)  # noqa: F821
    form.Activate()
except Exception:
    _error(u"The report could not be opened:\n{}\n\n{}".format(path, traceback.format_exc()))
