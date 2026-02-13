__author__ = 'Antonio Miano'
# Doesn't work for curved or multisegment grids

#REFERENCES
import clr

import System
from System.Collections.Generic import List

import pyrevit
from pyrevit import revit, DB, UI, HOST_APP, script, PyRevitException
from pyrevit import forms

from rpw.ui.forms import SelectFromList, CommandLink, TaskDialog

doc = revit.doc
active_view = doc.ActiveView
script_output = script.get_output()
BIC = DB.BuiltInCategory

#DEFINITIONS
l_tolist = lambda x : x if hasattr(x, '__iter__') else [x]

def f_intersection(crv1, crv2):
	results = clr.Reference[DB.IntersectionResultArray]()
	result = crv1.Intersect(crv2, results)
	if result != DB.SetComparisonResult.Overlap:
		return None
	intersection = results.Item[0]
	return intersection.XYZPoint

def f_new_crvs(dat_crv, ref_crv):
	dc_pt_ends = (dat_crv.GetEndPoint(0), dat_crv.GetEndPoint(1))
	dc_pt_ends_new = list(dc_pt_ends)
	dc_clone = DB.Line.CreateBound(dc_pt_ends[0], dc_pt_ends[1])
	dc_clone.MakeUnbound()
	dc_pt_inters = f_intersection(dc_clone, ref_crv)
	if dc_pt_inters != None:
		distances = (dc_pt_inters.DistanceTo(dc_pt_ends[0]), dc_pt_inters.DistanceTo(dc_pt_ends[1]))
		index_max = distances.index(min(distances))
		dc_pt_ends_new[index_max] = dc_pt_inters
		dat_crv_new = DB.Line.CreateBound(dc_pt_ends_new[0],dc_pt_ends_new[1])
		return dat_crv_new
	else:
		return None

def f_get_set_crvs(item, ref_item, view, dim):
	if dim == '3D':
		try:
			item_crv = item.GetCurvesInView(DB.DatumExtentType.Model, view)[0]
			ref_item_crv = ref_item.GetCurvesInView(DB.DatumExtentType.Model, view)[0]
			item_crv_new = f_new_crvs(item_crv, ref_item_crv)
			if item_crv_new:
				item.SetCurveInView(DB.DatumExtentType.Model, view, item_crv_new)
		except Exception as e:
			out_item = script_output.linkify(item.Id)
			out_exception = str(e)
			return [out_item, out_exception]
	else:
		try:
			item_crv = item.GetCurvesInView(DB.DatumExtentType.ViewSpecific, view)[0]
			ref_item_crv = ref_item.GetCurvesInView(DB.DatumExtentType.ViewSpecific, view)[0]
			item_crv_new = f_new_crvs(item_crv, ref_item_crv)
			if item_crv_new:
				item.SetCurveInView(DB.DatumExtentType.ViewSpecific, view, item_crv_new)
		except Exception as e:
			out_item = script_output.linkify(item.Id)
			out_exception = str(e)
			return [out_item, out_exception]

#Define a selection filter for Datums Categories
class DatumSelectionFilter(UI.Selection.ISelectionFilter):
	def __init__(self):
		pass
	
	def AllowElement(self, element):
		# if element.Category.Name == "Grids" or element.Category.Name == "Reference Planes" or element.Category.Name == "Levels":
		if element.Category.Name in (
			"Grids",
			"Reference Planes",
			"Levels"
		):
			return True
		else:
			return False

	def AllowReference(self, element):
		return False


#INPUTS
sel_filter = DatumSelectionFilter()

##Form to specify if the modification will affect the 2D or 3D
dims = [CommandLink('2D', return_value='2D'), CommandLink('3D', return_value='3D')]
dims_dialog = TaskDialog(	'Datum Extents',
									content='Specify in which dimension the extents will be modified',
									commands=dims,
									show_close=True)
dimension = dims_dialog.show()
if not dimension:	script.exit()

##Select the elements to Trim/Extend 
with forms.WarningBar(title='Select Elements to Trim/Extend'):
	try:
		# datums = HOST_APP.uidoc.Selection.PickElementsByRectangle(sel_filter)
		datums_reference = HOST_APP.uidoc.Selection.PickObjects(UI.Selection.ObjectType.Element, sel_filter)
		datums_ids = [dr.ElementId for dr in datums_reference]
		datums = [doc.GetElement(di) for di in datums_ids]
	except:
		script.exit()

##Select the element to be considered as reference
with forms.WarningBar(title='Select Reference Element'):
	try:
		ref_element = revit.pick_element_by_category(BIC.OST_CLines, "Select Reference Element")
	except:
		script.exit()

##If 2D is chosen, select in which views the modifications will be applied
if dimension == '2D':
	views_collection = forms.select_views(	title='Select Views where modifications will be applied')
	if not views_collection:	script.exit()
else:
	views_collection = [active_view]

#CODE
table_body = []
with revit.Transaction('Datum_Extents'):
	for view in views_collection:
		for datum in datums:
			result = f_get_set_crvs(datum, ref_element, view, dimension)
			if result:
				table_body.append(result)

##Print the output
table_headers = ['Element', 'Error']

if len(table_body)>0:
	script_output.print_table(
		table_data = table_body,
		title = 'Datum Extents',
		columns = table_headers
	)
