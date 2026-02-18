#REFERENCES
import pyrevit
import System

from pyrevit import revit, DB, UI, script
from pyrevit import forms
from rpw.ui.forms import TaskDialog, FlexForm, Label, ComboBox, TextBox, CheckBox, Separator, Button

from System.Collections.Generic import List

output = script.get_output()

#DEFINITIONS
l_tolist = lambda x : x if hasattr(x, '__iter__') else [x]

def f_set_par_value(p_read, p_write, conv, success, override):
	count_ok = 0
	count_no = 0
	exc_msg = success
	if p_read.HasValue:
		if p_read.StorageType == DB.StorageType.String:
			p_read_value = p_read.AsString()
			p_read_value_str = p_read.AsString()
		elif p_read.StorageType == DB.StorageType.ElementId:
			p_read_value = p_read.AsElementId()
			p_read_value_str = p_read.AsValueString()
		elif p_read.StorageType == DB.StorageType.Double:
			p_read_value = p_read.AsDouble()
			p_read_value_str = p_read.AsValueString()
		elif p_read.StorageType == DB.StorageType.Integer:
			p_read_value = p_read.AsInteger()
			p_read_value_str = p_read.AsValueString()
		
		if p_write.StorageType == DB.StorageType.String:
			p_write_value_str = p_write.AsString()
		elif p_write.StorageType == DB.StorageType.ElementId:
			p_write_value_str = p_write.AsValueString()
		elif p_write.StorageType == DB.StorageType.Double:
			p_write_value_str = p_write.AsValueString()
		elif p_write.StorageType == DB.StorageType.Integer:
			p_write_value_str = p_write.AsValueString()

		if override:
			try:
				if conv:
					p_write.Set(p_read_value_str)
					# p_write.Set.Overloads[[DB.StorageType.String]](p_read_value_str)
				else:
					p_write.Set(p_read_value)
				count_ok = 1
			except Exception as e:
				exc_msg = str(e)
				count_no = 1
		else:
			if p_write.HasValue:
				if len(p_write_value_str)>0:
					exc_msg = 'Parameter already filled'
					count_no = 1
			else:	
				if conv:
					p_write.Set(p_read_value_str)
					# p_write.Set.Overloads[[DB.StorageType.String]](p_read_value_str)
				else:
					p_write.Set(p_read_value)
				count_ok = 1
	else:
		count_ok = 0
	return count_ok, count_no, exc_msg

#Requires f_set_par_value()
def f_transfer_param(el, param_read, param_write, param_text, param_override):
	elem_count_ok = 0
	elem_count_no = 0
	elem_skipped = ("","","")
	if param_text:
		if 'String' not in str(param_write.StorageType):
			elem_skipped = (output.linkify(el.Id), el.Category.Name, msg_storage_type)
			elem_count_no = 1
		else:
			set_par_value = f_set_par_value(param_read, param_write, param_text, msg_transfer, param_override)
			elem_count_ok = set_par_value[0]
			elem_count_no = set_par_value[1]
			if set_par_value[1]>0:
				elem_skipped = (output.linkify(el.Id), el.Category.Name, set_par_value[2])
	else:
		if (
			param_read.StorageType != param_write.StorageType
			and not (
				'String' in str(param_read.StorageType)
				and 'String' in str(param_write.StorageType)
			)
		):
			elem_skipped = (output.linkify(el.Id), el.Category.Name, msg_storage_type)
			elem_count_no = 1						
		else:
			set_par_value = f_set_par_value(param_read, param_write, param_text, msg_transfer, param_override)
			elem_count_ok = set_par_value[0]
			elem_count_no = set_par_value[1]
			if set_par_value[1]>0:
				elem_skipped = (output.linkify(el.Id), el.Category.Name, set_par_value[2])
	return elem_skipped, elem_count_ok, elem_count_no

#INPUTS
doc = revit.doc
BIC = DB.BuiltInCategory
label_read = 'READ - Parameter Name:'
label_write = 'WRITE - Parameter Name:'

if __shiftclick__:
	elements_filtered = revit.get_selection()
	if not elements_filtered:
	##Create output message
		msg = 'Select elements first!'
		dialog = TaskDialog('Transfer Parameters', content = msg, buttons = ['OK'], footer = '', show_close = True)
		dialog.show(exit = True)
		script.exit()
	##Create inputs for the flexform and show it
	components = [
	CheckBox('cb_override', 'Override Value?', default=True),
	CheckBox('cb_text', 'Convert to Text? (mandatory for worksets)'),
	CheckBox('cb_skipped', 'List skipped elements?', default=False),
	Separator(),
	Label(label_read),
	TextBox('tbr_name'),
	Label(label_write),
	TextBox('tbw_name'),
	Separator(),
	Button('OK')	
	]
	flex_form = FlexForm('Transfer Parameter', components)
	flex_form.show()
	if not flex_form.values:
		script.exit()

	##Extract inputs from flexform
	par_text = flex_form.values['cb_text']
	par_override = flex_form.values['cb_override']	
	lst_skipped = flex_form.values['cb_skipped']
	par_read_name = flex_form.values['tbr_name']
	par_write_name = flex_form.values['tbw_name']

else:
	##Collect the Categories that will be shown in the flexform
	categories = doc.Settings.Categories
	model_categories = [cat for cat in categories if cat.CategoryType == DB.CategoryType.Model]

	##Get BuiltInCategories from Categories and add some extra BuiltinCategories
	bicategories = [revit.query.get_builtincategory(mc.Id) for mc in model_categories if revit.query.get_builtincategory(mc.Id) != None]
	add_bics = [
		BIC.OST_Levels,
		BIC.OST_Grids,
		BIC.OST_Views,
		BIC.OST_Sheets,
		BIC.OST_Schedules
	]
	bicategories.extend(add_bics)
	bicategories_name  = [revit.query.get_category(bicat).Name for bicat in bicategories]
	bicategories_list = List[BIC](bicategories)

	##Add the option to select all categories
	bicategories_name.insert(0, '- All Categories')
	bicategories.insert(0, None)

	##Create inputs for the flexform and show it
	cats_dic = dict(zip(bicategories_name,bicategories))
	components = [
		Label('Select Category:'),
		ComboBox('cb_cats',cats_dic),
		Separator(),
		CheckBox('cb_override', 'Override Value?', default=True),
		CheckBox('cb_text', 'Convert to Text? (mandatory for worksets)'),
		CheckBox('cb_skipped', 'List skipped elements?', default=False),
		Separator(),
		Label(label_read),
		TextBox('tbr_name'),
		Label(label_write),
		TextBox('tbw_name'),
		Separator(),
		Button('OK')	
	]
	flex_form = FlexForm('Transfer Parameter', components)
	flex_form.show()
	if not flex_form.values:
		script.exit()

	##Extract inputs from flexform
	bicategory = flex_form.values['cb_cats']
	par_text = flex_form.values['cb_text']
	lst_skipped = flex_form.values['cb_skipped']
	par_override = flex_form.values['cb_override']
	par_read_name = flex_form.values['tbr_name']
	par_write_name = flex_form.values['tbw_name']

	##Collect elements to be processed
	if bicategory == None:
		multi_cat_filter = DB.ElementMulticategoryFilter(bicategories_list)
		elements_filtered = 	DB.FilteredElementCollector(doc)\
									.WherePasses(multi_cat_filter)\
									.WhereElementIsNotElementType()\
									.ToElements()
	else:
		elements_filtered =	DB.FilteredElementCollector(doc)\
									.OfCategory(bicategory)\
									.WhereElementIsNotElementType()\
									.ToElements()

#CODE
##Define messages
msg_storage_type = 'Parameters are of different types'
msg_instance_type = 'Instance Parameter cannot be transferred to Type Parameter'
msg_read_notfound = 'Parameter to READ not found'
msg_write_notfound = 'Parameter to WRITE not found'
msg_transfer = 'Parameter transferred'
msg_error = 'Error'

##Transfer the parameter's values between the specified parameters
elems_count_ok = 0
elems_count_no = 0
elems_skipped = []
with revit.Transaction(name='Transfer Parameter', doc=doc):
	for elem in elements_filtered:
		par_read = elem.LookupParameter(par_read_name)
		par_write = elem.LookupParameter(par_write_name)
		if par_read:
			if par_write:
				transfer_param = f_transfer_param(elem, par_read, par_write, par_text, par_override)
				elems_count_ok += transfer_param[1]
				elems_count_no += transfer_param[2]
				if lst_skipped:
					elems_skipped.append(transfer_param[0])
			else:
				elem_type = elem.Document.GetElement(elem.GetTypeId())
				if elem_type:
					par_write = elem_type.LookupParameter(par_write_name)
					if par_write:
						elems_count_no += 1
						if lst_skipped:
							elems_skipped.append((output.linkify(elem.Id), elem.Category.Name, msg_instance_type))
						# transfer_param = f_transfer_param(elem, par_read, par_write, par_text, par_override)
						# elems_skipped.append(transfer_param[0])
						# elems_count_ok += transfer_param[1]
						# elems_count_no += transfer_param[2]
					else:
						elems_count_no += 1
						if lst_skipped:
							elems_skipped.append((output.linkify(elem.Id), elem.Category.Name, msg_write_notfound))

		else:
			elem_type = elem.Document.GetElement(elem.GetTypeId())
			if elem_type:
				par_read = elem_type.LookupParameter(par_read_name)
				if par_read:
					par_write = elem.LookupParameter(par_write_name)
					if par_write:
						transfer_param = f_transfer_param(elem, par_read, par_write, par_text, par_override)
						elems_count_ok += transfer_param[1]
						elems_count_no += transfer_param[2]
						if lst_skipped:
							elems_skipped.append(transfer_param[0])
					else:
						par_write = elem_type.LookupParameter(par_write_name)
						if par_write:
							transfer_param = f_transfer_param(elem, par_read, par_write, par_text, par_override)
							elems_count_ok += transfer_param[1]
							elems_count_no += transfer_param[2]
							if lst_skipped:
								elems_skipped.append(transfer_param[0])					
						else:
							elems_count_no += 1
							if lst_skipped:
								elems_skipped.append((output.linkify(elem.Id), elem.Category.Name, msg_write_notfound))
				else:
					elems_count_no += 1
					if lst_skipped:	
						elems_skipped.append((output.linkify(elem.Id), elem.Category.Name, msg_read_notfound))
						
			else:
				elems_count_no += 1
				if lst_skipped:
					elems_skipped.append((output.linkify(elem.Id), elem.Category.Name, msg_read_notfound))

table1_headers = [label_read.replace(':',''), label_write.replace(':', ''),'Override Value?', 'Convert to Text?']
table1_body = [[par_read_name, par_write_name, par_override, par_text]]
table2_headers = ['Total elements', 'Nr of elements modified', 'Nr of elements skipped']
table2_body = [[str(len(elements_filtered)), str(elems_count_ok), str(elems_count_no)]]
table3_headers = ['Id', 'Category', 'Message']

output.print_table(
	table_data = table1_body,
	title = 'TRANSFER PARAMETER',
	columns = table1_headers
)

output.print_table(
	table_data = table2_body,
	title = 'PROCESS REPORT',
	columns = table2_headers
)

if lst_skipped:
	if any([item != ("","","") for item in elems_skipped]):
		output.print_table(
			table_data = elems_skipped,
			title = 'SKIPPED ELEMENT(s)',
			columns = table3_headers
		)
