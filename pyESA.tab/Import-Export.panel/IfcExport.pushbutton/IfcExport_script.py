# -*- coding: utf-8 -*-
__title__ = 'IFC(s)\nExport'
__context__ = 'zero-doc'
__doc__ = 'Export IFC(s) from selected RVT(s) by specifying\nthe View Name and the Json File to be used.'
__author__ = 'Antonio Miano'

#REFERENCES
import pyrevit
import json
import System

from pyrevit import revit, DB, UI, script, output
from pyrevit import PyRevitException, PyRevitIOError
from pyrevit import forms

from rpw.ui.forms import TaskDialog, CheckBox, FlexForm, Label, TextBox, Separator, Button

#DEFINITIONS
l_tolist = lambda x: x if hasattr(x, '__iter__') else [x]

def to_bool(value):
	# Il JSON dell'add-in IFC usa true/false, ma accettiamo anche stringhe
	if isinstance(value, basestring):
		return value.strip().lower() in ('true', '1', 'yes')
	return bool(value)

def get_ifc_version(json_dict):
	# Converte il valore 'IFCVersion' del JSON nell'enum DB.IFCVersion.
	# Restituisce (enum o None, messaggio per il report)
	raw = json_dict.get('IFCVersion')
	if raw is None:
		return None, 'IFCVersion not in JSON, Revit default used'
	try:
		if isinstance(raw, basestring) and not raw.strip().isdigit():
			version = System.Enum.Parse(DB.IFCVersion, raw.strip())
		else:
			version = System.Enum.ToObject(DB.IFCVersion, int(raw))
			if not System.Enum.IsDefined(DB.IFCVersion, version):
				raise ValueError()
	except Exception:
		return None, 'IFCVersion "{}" not supported by this Revit version, Revit default used'.format(raw)
	return version, str(version)

def build_ifc_options(json_dict, ifc_version):
	# Proprieta' tipizzate (non lette da AddOption) + tutte le chiavi come opzioni testuali,
	# come fa IFCExportConfiguration.UpdateOptions dell'add-in IFC
	options = DB.IFCExportOptions()
	if json_dict is None:
		return options
	if ifc_version is not None:
		options.FileVersion = ifc_version
	if 'SpaceBoundaries' in json_dict:
		try:
			options.SpaceBoundaryLevel = int(json_dict['SpaceBoundaries'])
		except Exception:
			pass
	if 'ExportBaseQuantities' in json_dict:
		options.ExportBaseQuantities = to_bool(json_dict['ExportBaseQuantities'])
	if 'SplitWallsAndColumns' in json_dict:
		options.WallAndColumnSplitting = to_bool(json_dict['SplitWallsAndColumns'])
	for key, value in json_dict.items():
		if value is None:
			continue
		if isinstance(value, (dict, list)):
			# I valori annidati (ProjectAddress, ClassificationSettings...) vanno passati come JSON, non come repr Python
			value = json.dumps(value)
		elif isinstance(value, bool):
			value = 'true' if value else 'false'
		elif not isinstance(value, basestring):
			value = str(value)
		options.AddOption(key, value)
	return options

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
json_dict = None
ifc_version = None
ifc_version_label = 'Revit default (no JSON)'
if flex_form.values['json_pass']:
	json_path = forms.pick_file(files_filter='Json Files |*.json', multi_file=False, title='Select Json File')
	if not json_path: script.exit()
	json_name = json_path.split('\\')[-1]
	with open(json_path) as json_file:
		json_dict = json.load(json_file)
	ifc_version, ifc_version_label = get_ifc_version(json_dict)

#CODE
# app = __revit__.Application
script_output = script.get_output()

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

	###Build IFC export options for each file (FilterViewId must not carry over between documents)
	ifc_options = build_ifc_options(json_dict, ifc_version)

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
			if not flex_form.values['cb_pass']:
				out_rows.append(['Skipped: ' + rvt_file, ifc_view_name])
				temp_doc.Close(False)
				continue
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
	title = 'JSON file: ' + json_name + ' | IFC version: ' + ifc_version_label,
	columns = table_headers
)
