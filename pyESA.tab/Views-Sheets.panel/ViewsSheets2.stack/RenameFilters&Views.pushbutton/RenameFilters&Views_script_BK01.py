#REFERENCES
import pyrevit
from pyrevit import revit, script, DB
from pyrevit import forms
from rpw.ui.forms import TaskDialog, FlexForm, Label, TextBox, Separator, Button


#CODE
doc = revit.doc
if __shiftclick__:
	elements = forms.select_viewtemplates()
else:
	elements = forms.select_views(use_selection=True)

if elements:
	components = [
		Label('Add Prefix'),
		TextBox('prefix'),
		Label('Add Suffix'),
		TextBox('suffix'),
		Label('Find'),
		TextBox('find'),
		Label('Replace'),
		TextBox('replace'),
		Separator(),
		Button('OK')
	]

	form = FlexForm('Rename Views & Sheets', components)
	form.show()

	if form.values:
		with revit.Transaction('Rename Views&Sheets'):
			try:
				for elem in elements:
					orig_name = revit.query.get_name(elem)
					new_name = orig_name.replace(form.values['find'],form.values['replace']) if len(form.values['find'])>0 else orig_name
					revit.update.set_name(elem,'{0}{1}{2}'.format(form.values['prefix'],new_name,form.values['suffix']))
			except Exception, e:
				print(str(e))
	else:
		script.exit()
else:
	script.exit()