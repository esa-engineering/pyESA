__title__ = 'IFC(s)\nExport'
__context__ = 'zero-doc'
__doc__ = 'Export IFC(s) from selected RVT(s) by specifying\nthe View Name and the Json File to be used.'
__author__ = 'Antonio Miano'

#REFERENCES
import pyrevit
import json

from pyrevit import revit, DB, UI, script, output
from pyrevit import PyRevitException, PyRevitIOError
from pyrevit import forms

from rpw.ui.forms import TaskDialog, CheckBox, FlexForm, Label, TextBox, Separator, Button

#DEFINITIONS
l_tolist = lambda x: x if hasattr(x, '__iter__') else [x]

#INPUTS
##Collect RVT files
rvt_files = l_tolist(forms.pick_file(files_filter=	'Revit Files |*.rvt',
									multi_file=True,
									title='Select Revit File(s)'))
if not rvt_files[0]: script.exit()

##Create form
components = [
Label('View Name contains:'),
TextBox('txt_viewname'),
Separator(),
CheckBox('cb_pass', 'Export IFC even if view is not found'),
CheckBox('json_pass', 'Select JSON file', default=True),
Separator(),
Button('Continue')	
]
flex_form = FlexForm('IFC(s) Export', components)
flex_form.show()
if not flex_form.values.items(): script.exit()

##Collect JSON file
json_name = 'Not specified!'
if flex_form.values['json_pass']:
	json_path = forms.pick_file(files_filter='Json Files |*.json', multi_file=False, title='Select Json File')
	json_name = json_path.split('\\')[-1]
	if not json_path: script.exit()
	with open(json_path) as json_file:
		json_dict = json.load(json_file)

#CODE
# app = __revit__.Application
script_output = script.get_output()

##Define IFC export options based on json settings
ifc_options = DB.IFCExportOptions()
if flex_form.values['json_pass']:
	for item in json_dict.items():
		ifc_options.AddOption(item[0],str(item[1]))

##Loop through RVT files
out_rows = []
for rvt_file in rvt_files:
	temp_folder = '\\'.join(rvt_files[0].split('\\')[:-1])
	temp_name = rvt_file.split('\\')[-1].replace('.rvt','.ifc')
	rvt_file_info = revit.files.get_file_info(rvt_file)

	###Specify options when opening the original RVT file
	open_opt = DB.OpenOptions()
	if rvt_file_info.IsWorkshared:
		####Add opening options for Workshared RVT file
		open_config = DB.WorksetConfiguration(DB.WorksetConfigurationOption.OpenAllWorksets)
		open_opt = DB.OpenOptions()
		open_opt.DetachFromCentralOption = DB.DetachFromCentralOption.DetachAndPreserveWorksets
		open_opt.SetOpenWorksetsConfiguration(open_config)
	
	###Open the original RVT file
	model_path = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(rvt_file)
	temp_doc = __revit__.Application.OpenDocumentFile(model_path, open_opt)

	###Collect the view for IFC export (if not found, default view will be used)
	temp_doc_3Dviews_1 = DB.FilteredElementCollector(temp_doc).OfClass(DB.View3D).ToElements()
	temp_doc_3Dviews_2 = [view for view in temp_doc_3Dviews_1 if not view.IsTemplate]
	user_view_name = flex_form.values['txt_viewname']
	view_found = False
	if len(user_view_name)>0:
		ifc_views = [view for view in temp_doc_3Dviews_2 if user_view_name in view.Name]
		if len(ifc_views)>0:
			view_found = True
			ifc_view = ifc_views[0]
			ifc_view_name = ifc_view.Name
			ifc_options.FilterViewId = ifc_view.Id
		else:
			ifc_view_name = 'View for export not found!'
			if not flex_form.values['cb_pass']: continue
	else:
		ifc_view_name = 'View for export not specified!'

	###Export IFC
	out_path = 'None'
	with revit.Transaction(name='IFC(s) Export', doc=temp_doc, swallow_errors=True, clear_after_rollback=True):
		if view_found:
			ifc_view.IsSectionBoxActive = False
			# rvt_links_cat = DB.Category.GetCategory(temp_doc, DB.BuiltInCategory.OST_RvtLinks)
			# ifc_view.SetCategoryHidden(rvt_links_cat.Id, True)
		export_test = temp_doc.Export(temp_folder,temp_name,ifc_options)
		if export_test:
			out_path = temp_folder + '\\' + temp_name

	out_rows.append([out_path, ifc_view_name])
	temp_doc.Close(False)
	# temp_doc.Dispose()

##Print the output
table_headers = ['Saved File Path', 'IFC Export View']
table_body = out_rows

script_output.print_table(
	table_data = table_body,
	title = 'JSON file: ' + json_name,
	columns = table_headers
)
