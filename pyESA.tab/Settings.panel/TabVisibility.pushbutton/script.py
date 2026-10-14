# -*- coding: utf-8 -*-
__title__ = "Tab\nVisibility"
__doc__ = """Version = 1.1
Date    = 14.10.2026
________________________________________________________________
Description:
Shows or hides the tabs of the Revit ribbon with a checkbox for each
tab (checked = shown). The list shows the tabs as they are now.

The pyESA tab is never listed: hiding it would leave this tool out
of reach. Revit starts again with all tabs shown: run this tool again
after a restart to hide them again.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import clr
clr.AddReference("AdWindows")
clr.AddReference("RevitAPIUI")
from Autodesk.Windows import ComponentManager

from pyrevit import script

import Tab_ui

XAML_FILE_NAME = "Tab_form.xaml"
OWN_TAB = "pyESA"  # mai elencata: nasconderla lascerebbe senza accesso a questo strumento

output = script.get_output()


def tab_key(tab):
    return str(tab.Id)


def tab_title(tab):
    try:
        return tab.Title or ""
    except Exception:
        return ""


def main():
    try:
        tabs = list(ComponentManager.Ribbon.Tabs)
    except Exception as ex:
        output.print_md("**The ribbon could not be read:** {}".format(ex))
        return
    listed = [t for t in tabs if tab_title(t) and tab_title(t) != OWN_TAB]
    items = [(tab_key(t), tab_title(t), bool(t.IsVisible)) for t in listed]
    if not items:
        output.print_md("No tab can be listed.")
        return

    result = Tab_ui.TabWindow(script.get_bundle_file(XAML_FILE_NAME), items).show()
    if result is None:
        script.exit()

    changed, failed = 0, []
    for tab in listed:
        wanted = result.get(tab_key(tab))
        if wanted is None or wanted == bool(tab.IsVisible):
            continue
        try:
            tab.IsVisible = wanted
            changed += 1
        except Exception as ex:
            failed.append("{}: {}".format(tab_title(tab), ex))
    if failed:
        output.print_md("**Not changed:** " + "; ".join(failed))
    if changed == 0 and not failed:
        output.print_md("No change.")


main()
