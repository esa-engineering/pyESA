__title__ = "Categories\nVisibility"

__doc__ = """Set the visibility of the filtered Categories
and SubCategories on the Active View
---
SHIFT-CLICK to apply the visibility to a list of views"""

__author__ = "bimdifferent"

# REFERENCES
import pyrevit
from pyrevit import revit, script, DB
from pyrevit import forms
from rpw.ui.forms import FlexForm, Label, TextBox, ComboBox, Separator, Button

# DEFINITIONS
l_tolist = lambda x : x if hasattr(x, '__iter__') else [x]
doc = revit.doc

# INPUT

## Collect Categories and SubCategories in the document
cats_subcats = []
for cat in doc.Settings.Categories:
	cats_subcats.append(cat)
	cats_subcats.extend(cat.SubCategories)

## Collect view/views
if __shiftclick__:
	selected_views = l_tolist(forms.select_views())
else:
	selected_views = l_tolist(doc.ActiveView)

## Form
cbx_toggle_dict = {'Toggle OFF': True, 'Toggle ON': False}

components = [
	Label('Category/SubCategory contains?'),
	TextBox('txt_cats'),
	ComboBox('cbx_toggle', cbx_toggle_dict),
	Separator(),
	Button('OK')
]

flex_form = FlexForm('Categories Visibility', components)
flex_form.show()

if not flex_form.values:	script.exit()

## Get FlexForm values
search = flex_form.values['txt_cats']
if not search:	script.exit()

toggle = flex_form.values['cbx_toggle']

# CODE
csc_filtered = []
for csc in cats_subcats:
	if search.lower() in csc.Name.lower():
		csc_filtered.append(csc)

with revit.Transaction('Categories Visibility'):
	for view in selected_views:
		for cscf in csc_filtered:
			if view.CanCategoryBeHidden(cscf.Id):
				try:	view.SetCategoryHidden(cscf.Id, toggle)
				except:	pass
