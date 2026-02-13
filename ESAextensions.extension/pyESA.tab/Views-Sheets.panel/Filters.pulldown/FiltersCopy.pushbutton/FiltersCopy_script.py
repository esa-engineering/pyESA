__title__ = "Filters\nCopy"

__doc__ ="""Copy one or more View Filters
from a view/template to multiple ones.
---
SHIFT-CLICK to copy View Filters from
other opened Document"""

__author__ = "Macro4BIM\nbimdifferent"

# REFERENCES
import sys
from pyrevit import revit, DB, script
from pyrevit import PyRevitException, PyRevitIOError
from pyrevit import forms

# DEFINITIONS
doc = revit.doc
collector = DB.FilteredElementCollector


if __shiftclick__:
	source_doc = forms.select_open_docs(title='Select Source Document', multiple=False)
	if not source_doc:
		sys.exit(0)
	
	doc_filters = DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement).ToElements()
	doc_filters_name = [dflt.Name for dflt in doc_filters]
	source_filters = DB.FilteredElementCollector(source_doc).OfClass(DB.ParameterFilterElement).ToElements()
	selected_filters = forms.SelectFromList.show(source_filters,
						multiselect=True,
						name_attr='Name',
						button_name='Select Filters')
	if not selected_filters: script.exit()


	with revit.Transaction('Copy/Paste Filters'):
		for flt in selected_filters:
			flt_name = flt.Name
			flt_categories = flt.GetCategories()
			flt_elemfilter = flt.GetElementFilter()
			if flt_name not in doc_filters_name:
				try:
					DB.ParameterFilterElement.Create(doc, flt.Name, flt_categories, flt_elemfilter)
				except Exception as ex:
					print(str(ex))
			else:	pass
				
else:
	# INPUTS
	allViewClass = collector(doc).OfClass(DB.View)
	_views = []
	_templates = []
	for v in allViewClass:
		if not v.IsTemplate and v.ViewTemplateId == DB.ElementId.InvalidElementId:
			_views.append(v)
		elif v.IsTemplate:
			_templates.append(v)

	## MAKE FORMS
	### Views / Viewtemplates
	ops = {"Views":_views, "ViewTemplates":_templates}
	selected_view = forms.SelectFromList.show(ops,
						multiselect=False,
						name_attr='Name',
						button_name='Select Views/ViewTemplate')

	if not selected_view: script.exit()

	## CHOOSE BETWEEN FILTERS
	filters = [doc.GetElement(i) for i in selected_view.GetFilters()]
	selected_filters = forms.SelectFromList.show(filters,
						multiselect=True,
						name_attr='Name',
						button_name='Select Filters')

	if not selected_filters: script.exit()

	## DESTINATION VIEWS
	target_view = forms.SelectFromList.show(ops,
						multiselect=True,
						name_attr='Name',
						button_name='Select Views/ViewTemplate')

	if not target_view: script.exit()


	## APPLY
	nrOfWarning = 0

	msg = "Some views haven't been edit:\n"
	if forms.alert("Edit {} views/templates?".format(len(target_view)),
					yes = True, no = True):
		with revit.Transaction('Copy/Paste view filters'):
			for flt in selected_filters:
				ogs = selected_view.GetFilterOverrides(flt.Id)
				vis = selected_view.GetFilterVisibility(flt.Id)
				for vw in target_view:
					try:
						vw.SetFilterOverrides(flt.Id, ogs)
						vw.SetFilterVisibility(flt.Id, vis)
					except Exception as ex:
						nrOfWarning += 1
						msg += '{}. {} ({})\n'.format(nrOfWarning, vw.Name, flt.Name)
						descr += "{}. {}\n".format(nrOfWarning, ex)

	if nrOfWarning > 0:
		forms.alert(msg, expanded=descr)