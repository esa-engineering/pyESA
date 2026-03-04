__title__ = "Template\nCleanup"
__doc__ = """Version = 1.0
Date = 03.03.2026
________________________________________________________________
Cleans up template-related entities from the active document:
- View Templates
- View Filters
- Schedules
- Legends

Entities are filtered by discipline tag in their name.
Flag the disciplines / items you want to KEEP.
________________________________________________________________
Author(s):
Antonio Miano
"""

# REFERENCES
from pyrevit import revit, DB, script
from rpw.ui.forms import FlexForm, Label, CheckBox, Separator, Button

doc = revit.doc

# FORM
components = [
    Label('Flag disciplines / items to KEEP:'),
    Separator(),
    CheckBox('keep_arc',   'ARC'),
    CheckBox('keep_mep',   'MEP'),
    CheckBox('keep_str',   'STR'),
    Separator(),
    CheckBox('keep_items', 'Elements (Placeholders, Groups, Text)'),
    Separator(),
    Button('Run Cleanup')
]

flex_form = FlexForm('Template Cleanup', components)
flex_form.show()

if not flex_form.values:
    script.exit()

keep_arc   = flex_form.values['keep_arc']
keep_mep   = flex_form.values['keep_mep']
keep_str   = flex_form.values['keep_str']
keep_items = flex_form.values['keep_items']

# COLLECT CANDIDATES
views     = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()
templates = [v for v in views if v.IsTemplate]
filters   = DB.FilteredElementCollector(doc).OfClass(DB.FilterElement).ToElements()
schedules = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule).ToElements()
legends   = [v for v in views if v.ViewType.ToString() == "Legend"]
materials   = DB.FilteredElementCollector(doc).OfClass(DB.Material).ToElements()
lines_cat   = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
line_styles = [cat for cat in lines_cat.SubCategories]

candidates = list(templates) + list(filters) + list(schedules) + list(legends) + list(materials) + list(line_styles)

# DISCIPLINE TAG GROUPS
MEP_TAGS   = ("ELE", "FIR", "LIT", "LHT", "MEC", "MEP", "PLU", "MPF", "SYS", "STR", "_DS", "_CO", "_CT", "_PS")
ARC_OTHERS = ("COO", "ELE", "FIR", "GEN", "LIT", "LHT", "MEC", "MEP", "PLU", "MPF", "SYS", "STR", "_DS", "_CO", "_CT", "_PS")
MEP_OTHERS = ("ARC", "COO", "GEN", "STR")
STR_OTHERS = ("ARC", "COO", "ELE", "FIR", "GEN", "LIT", "LHT", "MEC", "MEP", "PLU", "MPF", "SYS", "_DS", "_CO", "_CT", "_PS")

# DETERMINE ELEMENTS TO DELETE
ids_to_delete = set()
elems_map     = {}

for elem in candidates:
    try:
        name = elem.Name
        eid  = elem.Id
        if not keep_arc:
            if "ARC" in name and not any(t in name for t in ARC_OTHERS):
                ids_to_delete.add(eid)
                elems_map[eid] = elem
        if not keep_mep:
            if any(t in name for t in MEP_TAGS) and not any(t in name for t in MEP_OTHERS):
                ids_to_delete.add(eid)
                elems_map[eid] = elem
        if not keep_str:
            if "STR" in name and not any(t in name for t in STR_OTHERS):
                ids_to_delete.add(eid)
                elems_map[eid] = elem
    except:
        pass

if not keep_items:
    FAMILY_TAGS = ("e_GM.MT_Placeholder", "e_AN.TG.MT_Materials", "e_DT_LibraryContainer")
    fams = DB.FilteredElementCollector(doc).OfClass(DB.Family).ToElements()
    for fam in fams:
        if any(tag in fam.Name for tag in FAMILY_TAGS):
            ids_to_delete.add(fam.Id)
            elems_map[fam.Id] = fam

    groups = DB.FilteredElementCollector(doc).OfClass(DB.GroupType).ToElements()
    for group in groups:
        param = group.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if param and param.AsString() == "Header_8000":
            ids_to_delete.add(group.Id)
            elems_map[group.Id] = group

    text_elems = DB.FilteredElementCollector(doc).OfClass(DB.ModelText).ToElements()
    for t in text_elems:
        ids_to_delete.add(t.Id)
        elems_map[t.Id] = t

# DELETE
errors  = []
deleted = 0

with revit.Transaction('Template Cleanup'):
    for eid in ids_to_delete:
        try:
            doc.Delete(eid)
            deleted += 1
        except Exception as e:
            elem = elems_map.get(eid)
            name = elem.Name if elem else repr(eid)
            errors.append('{} - {}'.format(name, repr(e)))

"""
# OUTPUT
out = script.get_output()
out.print_md('## Template Cleanup - Results')
out.print_md('**Deleted:** {} elements'.format(deleted))
if errors:
    out.print_md('**Errors ({}):**'.format(len(errors)))
    for err in errors:
        out.print_md('- ' + err)
"""
