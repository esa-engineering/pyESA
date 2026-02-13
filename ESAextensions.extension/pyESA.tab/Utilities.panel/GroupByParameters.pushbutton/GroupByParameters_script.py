__doc__ = 'Group Elements of selected Category by combining\n'\
		'the values of a list of Parameters in a new Parameter\n'\
		'(that can be used in Revit Schedules)\n\n'\
		'SHIFT-CLICK for ESA Door-Schedule Parameters'
__title__ = 'Group by\nParameters'

#REFERENCES
import clr
import System
from pyrevit import revit, DB, UI
from rpw.ui.forms import (TaskDialog, FlexForm, Label, ComboBox, TextBox, Separator, Button, CheckBox)

doc = revit.doc

#DEFINITIONS
##Return the value of the parameter as a string
def f_pValue_string(parameter):
	if str(parameter.StorageType) == 'String':
		pValue = parameter.AsString()
	elif str(parameter.StorageType) == 'ElementId':
		pValue = parameter.AsValueString()
	elif str(parameter.StorageType) == 'Double':
		pValue = parameter.AsValueString()
	elif str(parameter.StorageType) == 'Integer':
		pValue = parameter.AsValueString()
	else:
		pValue = 'Check Parameter Storage Type'
	return pValue

##Return the value of the element's parameter as a string (requires f_pValue_string)
def f_elem_pValue_string(elem,parName):
	elem_pValue = None
	param = elem.LookupParameter(parName)
	if param == None:
		try:
			elemTypeId = elem.GetTypeId()
			elemType = elem.Document.GetElement(elemTypeId)
			param = elemType.LookupParameter(parName)
			if param == None:
				try:
					elemSysTypeId = elem.LookupParameter('System Type').AsElementId()
					elemSysType = elem.Document.GetElement(elemSysTypeId)
					param = elemSysType.LookupParameter(parName)
					if param != None:
						elem_pValue = f_pValue_string(param)
				except Exception, e1:
					elem_pValue = str(e1)
			else:
				elem_pValue = f_pValue_string(param)
		except Exception, e2:
			elem_pValue = str(e2)
	else:
		elem_pValue = f_pValue_string(param)
	return elem_pValue

##Replace None objects
def f_replaceNone(val):
	if val == None:
		val = ''
	return val

#CODE
categories = doc.Settings.Categories
cats_names = [cat.Name for cat in categories]
bics_from_cats = [System.Enum.ToObject(DB.BuiltInCategory, cat.Id.IntegerValue) for cat in categories]
cats_dic = dict(zip(cats_names,bics_from_cats))

##Create form for inputs
if __shiftclick__: 
	components = 	[\
					Label('Select Category:'),\
					ComboBox('combobox1',cats_dic),\
					Label('Parameters to combine separated by \';\' eg: aaa;bbb;...'),\
					TextBox('textbox01', Text=	'e_DIM_DR_Width_1;e_DIM_DR_Width_2;e_DIM_DR_Width_Clear_1;e_DIM_DR_Height_1;'\
														'e_DAT_DW_ComponentDims_1;e_DAT_DR_Supplier_1;e_DAT_DR_nLeaves_1;'\
														'e_DAT_DR_Egress_1;e_IDA_AcousticRating_1;e_FPR_FireRating_1;'\
														'e_DAT_ResistenceClass_1;e_DAT_DR_Requirements_1;e_DAT_DR_Material_1;'\
														'e_DAT_DR_FinishA_1;e_DAT_DR_FinishB_1;e_DAT_DR_HandleA_1;e_DAT_DR_HandleB_1;'\
														'e_DAT_DR_AccessControl_1;e_DAT_DR_AlarmSystem_1;e_DAT_DR_Closer_1;'\
														'e_DAT_DR_Hinges_1;e_DAT_DR_Lockset_1;e_DAT_DR_Stop_1;e_DAT_Notes_1'),\
					Label('Parameter to write'),\
					TextBox('textbox16', Text='e_OTH_Temp_1_i'),\
					Separator(),\
					Button('OK')\
					]
	form = FlexForm('Group by Parameters', components)
	form.show()
else:
	components = 	[\
					Label('Select Category:'),\
					ComboBox('combobox1',cats_dic),\
					Label('Parameters to combine separated by \';\' eg: aaa;bbb;...'),\
					TextBox('textbox01'),\
					Label('Parameter to write'),\
					TextBox('textbox16'),\
					Separator(),\
					Button('OK')\
					]
	form = FlexForm('Group by Parameters', components)
	form.show()

try:
	cat_selected = form.values['combobox1']
	##Collect elements of selected category and group them by combining the specified parameters
	elems = DB.FilteredElementCollector(doc)\
			.OfCategory(cat_selected)\
			.WhereElementIsNotElementType().ToElements()

	param_names0 = form.values['textbox01']
	param_names0.replace('; ',';').replace(' ;',';')
	param_names1 = param_names0.split(';')

	param_names = [pn1 for pn1 in param_names1 if (pn1 != None and pn1 != '')]
	sep = ''

	marks = []

	for elem in elems:
		marks.append([f_replaceNone(f_elem_pValue_string(elem,p_name)) for p_name in param_names])

	marks_keys = [sep.join(mark) for mark in marks]
	elems_dic = {k: [elems[i] for i in [j for j,x in enumerate(marks_keys) if x==k]] for k in set(marks_keys)}
	n_dic = len(elems_dic)
	temp_keys = range(len(elems_dic))
	temp_keys = [str(tk).zfill(2)+'_Group' for tk in temp_keys]

	for old_key,temp_key in zip(elems_dic.keys(),temp_keys):
		elems_dic[temp_key] = elems_dic.pop(old_key)

	n_elems = len(elems)

	n_modif = 0
	t = DB.Transaction(doc, 'GroupByParameter')
	t.Start()
	for elem_groups,elem_key in zip(elems_dic.values(),elems_dic.keys()):
		for elem in elem_groups:
			try:
				elem.LookupParameter(form.values['textbox16']).Set(elem_key)
				n_modif += 1
			except: pass

	t.Commit()

	##Create output message
	msg = str(n_dic) + ' Groups created' + '\n\n' + str(n_elems) + ' Elements Selected' + '\n' + str(n_modif) + ' Elements Grouped'
	dialog = TaskDialog('Group by Parameters', content = msg, buttons = ['OK'], footer = '', show_close = True)
	dialog.show(exit = True)
except:	pass