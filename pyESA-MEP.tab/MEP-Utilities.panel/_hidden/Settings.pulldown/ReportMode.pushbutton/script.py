# -*- coding: utf-8 -*-
__title__ = "Report\nMode"
__context__ = "zero-doc"
__doc__ = """Version = 0.1
Date    = 08.10.2026
________________________________________________________________
Description:
Chooses when the pyESA tools of the MEP-Routing and MEP-Tools panels open their full
report at the end of the command.

Modes:
- Only on problems (default): the full report opens only if something
  was left unchanged or failed; otherwise a short notification shows
  the result
- Always: the full report opens every time
- Never: only the short notification, even with problems

The choice is saved and shared by all the tools of both panels.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

from pyrevit import script, forms

REPORT_CONFIG = "ESA_Report"
MODES = (
    ("problems", "Only on problems (default)"),
    ("always", "Always show the full report"),
    ("never", "Never: short notification only"),
)


def main():
    config = script.get_config(REPORT_CONFIG)
    try:
        current = config.get_option("mode", "problems")
    except Exception:
        current = "problems"
    labels = {}
    for key, label in MODES:
        labels[label + ("   (current)" if key == current else "")] = key
    choice = forms.CommandSwitchWindow.show([label + ("   (current)" if key == current else "") for key, label in MODES],
                                            message="When should the tools open their full report?")
    if not choice:
        script.exit()
    config.mode = labels[choice]
    script.save_config()
    try:
        script.toast("Report mode: {}".format(choice.replace("   (current)", "")), title="pyESA")
    except Exception:
        pass


main()
