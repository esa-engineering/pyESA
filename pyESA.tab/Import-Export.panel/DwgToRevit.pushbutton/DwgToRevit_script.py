#REFERENCES
import clr
import csv

from pyrevit import revit, DB, UI,script
from pyrevit import forms
from rpw.ui.forms import TaskDialog

from System.Collections.Generic import *

doc = revit.doc

#DEFINITIONS
def f_flatten(x):
	result = []
	for el in x:
		if hasattr(el, "__iter__") and not isinstance(el, basestring):
			result.extend(f_flatten(el))
		else:
			result.append(el)
	return result

##Transpose a list
def f_transpose(lista):
	return map(list, zip(*lista))

def f_integers(lst):
	colors0 = []
	for item in lst:
		if ',' in item:
			colors0.append(item.split(','))
		else:
			colors0.append(item)
	colors1 = []
	for col0 in colors0:
		if isinstance(col0, list):
			col1 = [int(c0) for c0 in col0]
		else:
			col1 = col0
		colors1.append(col1)
	return colors1

def f_getColors(item):
	return [item.LineColor.Red,item.LineColor.Green,item.LineColor.Blue]

def f_matchcolor(colors0,colors1):
	indices = []
	for col0 in colors0:
		ind = 'NotFound'
		for i,col1 in enumerate(colors1):
			if col0 == col1:
				ind = i
		indices.append(ind)
	return indices

#INPUTS
##Get CSV file
path = script.get_bundle_file("ESA_CTB1_Arc.csv")

if __shiftclick__:
	path = forms.pick_file(file_ext='csv')

##Get dwgs in the active document
dwgs = 	DB.FilteredElementCollector(doc)\
			.OfClass(DB.ImportInstance)\
			.WhereElementIsNotElementType().ToElements()
if dwgs:
	dwgs_dic = dict((dwg.Category.Name,dwg) for dwg in dwgs if dwg.Category != None)
	try:
		##Set UI input window
		dwgs_names_sel = forms.SelectFromList.show(dwgs_dic.keys(), title='Select DWGs', multiselect=True)
	except:
		script.exit()
else:
	##Create output message
	msg0 = 'No DWGs found'
	dialog0 = TaskDialog('Dwg To Revit', content = msg0, buttons = ['OK'], footer = '', show_close = True)
	dialog0.show(exit = True)
	script.exit()

#CODE
##Get information from UI input
if dwgs_names_sel:
	dwgs_sel = [dwgs_dic[k] for k in dwgs_names_sel]
	sub_cats = f_flatten([dwg_sel.Category.SubCategories for dwg_sel in dwgs_sel])
	sub_cats_color = [f_getColors(sc) for sc in sub_cats]

	##Read ;delimited CSV file data
	with open(path,'rb') as csv_file:
		csv_reader = csv.reader(csv_file, delimiter=';')
		csv_headers = next(csv_reader)
		lines = f_transpose([line for line in csv_reader])
		color_rgb =  f_integers([str(item) for item in lines[1]])
		color_plot =  f_integers([str(item) for item in lines[2]])
		weight_rvt =  f_integers([str(item) for item in lines[4]])

		##Match the items between excel and revit
		indices = f_matchcolor(sub_cats_color,color_rgb)

		##Change object styles
		n_mod = 0
		ok1 = False
		ok2 = False
		with revit.Transaction('DwgToRevit'):
			for ind,sc in zip(indices,sub_cats):
				try:
					sc.LineColor = DB.Color(color_plot[ind][0],color_plot[ind][1],color_plot[ind][2])
					ok1 = True
				except: pass
				try:
					sc.SetLineWeight(weight_rvt[ind], DB.GraphicsStyleType.Projection)
					ok2 = True
				except:
					sc.SetLineWeight(1, DB.GraphicsStyleType.Projection)
					ok2 = True
				if any([ok1,ok2]):
					n_mod += 1

			##Refresh active view
			v = UI.UIDocument(doc).ActiveView
			for ds in dwgs_sel:
				vId = List[DB.ElementId]( )
				vId.Add(ds.Id)
				v.HideElements(vId)
				v.UnhideElements(vId)

		##Create output message
		msg1 = '{0} Object Styles changed'.format(n_mod)
		dialog1 = TaskDialog('Dwg To Revit', content = msg1, buttons = ['OK'], footer = '', show_close = True)
		dialog1.show(exit = True)
else:
	script.exit()
