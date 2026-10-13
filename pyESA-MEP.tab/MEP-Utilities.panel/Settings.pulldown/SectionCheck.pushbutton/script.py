# -*- coding: utf-8 -*-
__title__ = "Section\nCheck"
__doc__ = """Version = 0.1
Date    = 12.10.2026
________________________________________________________________
Description:
Read-only. Lists the section views whose crop and section symbol in the
active plan do not line up by more than a threshold (default 300 mm).
The crop line is taken from the section crop box (the reference). The
symbol is measured by the centre of its plan bounding box, which also
includes the arrows and tags, so the value is an approximation. Nothing
in the model is changed. A section whose symbol is not visible in the
active plan is listed separately, not measured.

Fix the listed sections by hand: drag the section line in plan.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from pyrevit import revit, script, DB, forms

doc = revit.doc
output = script.get_output()
MM_PER_FT = 304.8
DEFAULT_THRESHOLD_MM = 300.0


def crop_line_middle(view):
    """Centro della linea di sezione, ricavato dal ritaglio (piano di taglio = Max.Z)."""
    box = view.CropBox
    return box.Transform.OfPoint(DB.XYZ((box.Min.X + box.Max.X) / 2.0, 0, box.Max.Z))


def symbol_centre(view, plan):
    """Centro del riquadro del simbolo di sezione (categoria Viewers, stesso nome della vista) visto in pianta.
    None se il simbolo non si vede in questa pianta."""
    for marker in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Viewers) \
            .WhereElementIsNotElementType():
        if marker.Name != view.Name:
            continue
        box = marker.get_BoundingBox(plan)
        if box is not None:
            return DB.XYZ((box.Min.X + box.Max.X) / 2.0, (box.Min.Y + box.Max.Y) / 2.0, 0)
    return None


def plan_distance_mm(a, b):
    return ((a.X - b.X) ** 2 + (a.Y - b.Y) ** 2) ** 0.5 * MM_PER_FT


def main():
    plan = doc.ActiveView
    if not isinstance(plan, DB.ViewPlan) or plan.IsTemplate:
        forms.alert("Open a plan view first: the section symbols are read in it.", exitscript=True)
    text = forms.ask_for_string(default=str(int(DEFAULT_THRESHOLD_MM)),
                                prompt="Threshold in mm: list the sections that differ by more than this",
                                title="Section Check")
    try:
        threshold = float(text.replace(",", "."))
    except (AttributeError, ValueError):
        script.exit()
    if threshold < 0:
        script.exit()

    sections = sorted((v for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewSection) if not v.IsTemplate),
                      key=lambda v: v.Name)
    flagged, not_visible, checked = [], [], 0
    for view in sections:
        centre = symbol_centre(view, plan)
        if centre is None:
            not_visible.append(view)
            continue
        checked += 1
        difference = plan_distance_mm(crop_line_middle(view), centre)
        if difference > threshold:
            flagged.append((view, difference))

    output.print_md("# Section check (threshold {:g} mm)".format(threshold))
    output.print_md("Sections in the model: **{}** | measured in this plan: **{}** | not visible here: **{}** | "
                    "over the threshold: **{}**".format(len(sections), checked, len(not_visible), len(flagged)))
    if flagged:
        rows = [[view.Name, "{:.0f}".format(difference), output.linkify(view.Id)]
                for view, difference in sorted(flagged, key=lambda item: -item[1])]
        output.print_table(table_data=rows, columns=["Section", "Difference [mm]", "View"])
    else:
        output.print_md("No section is over the threshold.")
    if not_visible:
        output.print_md("**Not measured (symbol not visible in this plan):** " +
                        ", ".join(v.Name for v in not_visible))
    output.print_md("*Approximate: the symbol is measured by the centre of its bounding box, which includes the "
                    "arrows and tags. The crop line is the reference.*")


main()
