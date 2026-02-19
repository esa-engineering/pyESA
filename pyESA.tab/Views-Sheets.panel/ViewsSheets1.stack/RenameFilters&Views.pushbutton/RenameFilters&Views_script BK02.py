#REFERENCES
import pyrevit
from pyrevit import revit, script, DB
from pyrevit import forms
from rpw.ui.forms import TaskDialog, FlexForm, Label, TextBox, CheckBox, Separator, Button

#DEFINITIONS
l_tolist = lambda x : x if hasattr(x, '__iter__') else [x]

#CODE
doc = revit.doc

if __shiftclick__:
	viewtemplates = revit.query.get_all_view_templates(doc=doc)
	viewtemplates.sort(key=lambda x:x.Name)

	rulefilters = revit.query.get_rule_filters(doc=doc)
	rulefilters.sort(key=lambda x:x.Name)

	options = {'ViewTemplates': viewtemplates, 'Filters': rulefilters}

	elements = forms.SelectFromList.show(
		options,
		multiselect=True,
		name_attr='Name',
		group_selector_title='Filters | ViewTemplates',
		button_name='Select'
	)

else:
	user_selection = l_tolist(revit.get_selection())

	if len(user_selection) > 0:
		elements = forms.select_views(use_selection=True)

	else:
		views = revit.query.get_all_views(doc=doc)
		views.sort(key=lambda x:x.Name)

		sheets = revit.query.get_sheets(doc=doc)
		sheets.sort(key=lambda x:x.Name)

		options = {'Views': views, 'Sheets': sheets}

		elements = forms.SelectFromList.show(
			options,
			multiselect=True,
			name_attr='Name',
			group_selector_title='Sheets | Views',
			button_name='Select'
		)

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

	form = FlexForm('Rename Filters&Views', components)
	form.show()

	if form.values:
		with revit.Transaction('Rename Filters&Views'):
			try:
				for elem in elements:
					orig_name = revit.query.get_name(elem)
					new_name = orig_name.replace(form.values['find'],form.values['replace']) if len(form.values['find'])>0 else orig_name
					revit.update.set_name(elem,'{0}{1}{2}'.format(form.values['prefix'],new_name,form.values['suffix']))
			except Exception as e:
				print(str(e))
	else:
		script.exit()
else:
	script.exit()