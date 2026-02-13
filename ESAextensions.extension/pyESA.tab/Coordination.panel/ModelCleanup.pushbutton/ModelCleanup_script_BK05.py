__context__ = "zero-doc"

# REFERENCES
import shutil
import time
import pyrevit
import System

from pyrevit import revit, DB, UI, script, forms
from pyrevit import PyRevitException, PyRevitIOError

from rpw.ui.forms import TaskDialog, FlexForm, Label, ComboBox, TextBox, CheckBox, Separator, Button

from System.Collections.Generic import List

import clr
import sys

hostapp = pyrevit._HostApplication()
app = hostapp.app
rvt_version = int(hostapp.version)

transmit_path_1 = "C:\\Program Files\\Autodesk\\eTransmit for Revit " + str(rvt_version)
sys.path.append(transmit_path_1)
clr.AddReference ('eTransmitForRevitDB')
from eTransmitForRevitDB import *

# DEFINITIONS
l_tolist = lambda x: x if hasattr(x, '__iter__') else [x]
l_string_clean = lambda x: '' if x == ' ' else x
l_list_elem_id = lambda lst: [x.Id.IntegerValue for x in lst]

def f_param_value_list(entry_string):
	clean_string = l_string_clean(entry_string)
	if len(clean_string)>0:
		list_string = clean_string.replace('; ',';').replace(' ;',';').split(';')
		return l_tolist(list_string)
	else:
		return None

def f_param_check(elem, param, values):
	"""
	Return True ANY value of the checked parameter matches
	with the provided value, False in any other cases
	"""
	param_elem = revit.query.get_param(elem, param)
	if param_elem:
		try:
			param_value = revit.query.get_param_value(param_elem)
			if param_value:
				return any([item in param_value for item in values])
		except:
			return False

def f_id_to_elem(doc, id):
	return doc.GetElement(DB.ElementId(id))

def f_get_all_views_on_sheet(doc, sheet, elemID = True):
	"""
	Return Viewsand Schedules (OR their IDs as Integer Values
 	if 'elemID' is set to True) placed on the Sheet
	"""
	sheet_views_all = []
	sheet_views = sheet.GetAllPlacedViews()
	if elemID:
		sheet_views_all = [item.IntegerValue for item in sheet_views]
	else:
		sheet_views_all = [doc.GetElement(item) for item in sheet_views]
	
	# Get Schedules placed on the sheet (Revision Schedule is included)
	sheet_schedules_placed = DB.FilteredElementCollector(doc, sheet.Id).OfClass(DB.ScheduleSheetInstance).ToElements()

	if elemID:
		sheet_schedules = [item.ScheduleId.IntegerValue for item in sheet_schedules_placed]
	else:
		sheet_schedules = [doc.GetElement(item.ScheduleId) for item in sheet_schedules_placed]

	sheet_views_all.extend(sheet_schedules)
	return sheet_views_all

# INPUTS
rvt_files = l_tolist(forms.pick_file(files_filter=	'Revit files (*.rvt)|*.rvt',
									multi_file=True,
									title='Select Revit File(s)'))

if rvt_files[0] != None:

	# Make form
	components = [
		Label('WORKSETS - to delete (empty for RVT < 2023)'),
		Label('Workset name contains (use \';\' as separator):'),
		TextBox('txt_worksets'),
		Separator(),
		Label('VIEWS - to keep'),
		Label('View parameter name:'),
		TextBox('txt_view_param'),
		Label('Parameter value contains (use \';\' as separator):'),
		TextBox('txt_view_contains'),
		Separator(),
		Label('SHEETS - to keep'),
		Label('Sheet parameter name:'),
		TextBox('txt_sheet_param'),
		Label('Parameter value contains (use \';\' as separator):'),
		TextBox('txt_sheet_contains'),
		Separator(),
		# CheckBox('cb_links', 'Include XRefs'),
		CheckBox('cb_purge', 'Purge Unused'),
		CheckBox('cb_detach', 'Create Transmit'),
		Separator(),
		Button('OK')			
		]

	flex_form = FlexForm('Model(s) Cleanup & OverWrite', components)
	flex_form.show()

	# CODE
	script_output = script.get_output()
	out_rows = []

	if	len(flex_form.values.items())>0:
		
		# Get inputs from form
		worksets_name = f_param_value_list(flex_form.values['txt_worksets'])
		views_param_name = l_string_clean(flex_form.values['txt_view_param'])
		views_param_values = f_param_value_list(flex_form.values['txt_view_contains'])
		sheets_param_name = l_string_clean(flex_form.values['txt_sheet_param'])
		sheets_param_values = f_param_value_list(flex_form.values['txt_sheet_contains'])

		# Iterate over each selected RVT file
		for rvt_file in rvt_files:
			
			start = time.time()
			with forms.ProgressBar(title=rvt_file.split('\\')[-1]) as pb:

				temp_name = rvt_file.split('\\')[-1]
				rvt_file_info = revit.files.get_file_info(rvt_file)

				pb.update_progress(5, 100)

				# Specify options when opening the original RVT file
				open_opt = DB.OpenOptions()

				# Add opening options for Workshared RVT file
				if rvt_file_info.IsWorkshared:					
					open_config = DB.WorksetConfiguration(DB.WorksetConfigurationOption.CloseAllWorksets)
					open_opt = DB.OpenOptions()
					open_opt.DetachFromCentralOption = DB.DetachFromCentralOption.DetachAndPreserveWorksets
					open_opt.SetOpenWorksetsConfiguration(open_config)
	
				# Open the original RVT file
				model_path = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(rvt_file)
				temp_doc = __revit__.Application.OpenDocumentFile(model_path, open_opt)

				# Instanciate the containers for Views and Sheets
				sheets_keep_set, sheets_delete_set = set(), set()
				views_keep_set, views_delete_set = set(), set()

				# Find sheets to delete/keep
				sheets_all = revit.query.get_sheets(doc=temp_doc)
				sheets_all_ids = [item.Id.IntegerValue for item in sheets_all]
				sheets_all_set = set(sheets_all_ids)

				if sheets_param_name:
					if sheets_param_values:
						for elemid in sheets_all_set:
							if f_param_check(f_id_to_elem(temp_doc, elemid), sheets_param_name, sheets_param_values):
								sheets_keep_set.add(elemid)
					else:
						sheets_delete_set = sheets_all_set
				else:
					sheets_keep_set = sheets_all_set
				sheets_delete_set = sheets_all_set - sheets_keep_set

				# Get the ID of views placed on Sheets and add them to the views_keep_set
				for elemid in sheets_keep_set:
					views_on_sheet = f_get_all_views_on_sheet(temp_doc, f_id_to_elem(temp_doc, elemid), elemID=True)
					views_on_sheet_set = set(views_on_sheet)
					views_keep_set.update(views_on_sheet_set)

				pb.update_progress(10, 100)

				# Find views to delete/keep
				views_all = revit.query.get_all_views(doc=temp_doc)
				views_all = [va for va in views_all if va.GetType().ToString() != 'Autodesk.Revit.DB.ViewSheet']
				views_all_ids = [item.Id.IntegerValue for item in views_all]
				views_all_set = set(views_all_ids)

				if views_param_name:
					if views_param_values:
						for elemid in views_all_set:
							if f_param_check(f_id_to_elem(temp_doc, elemid), views_param_name, views_param_values):
								views_keep_set.add(elemid)
					else:
						views_delete_set = views_all_set - views_keep_set
				else:
					views_keep_set = views_all_set
				views_delete_set = views_all_set - views_keep_set

				pb.update_progress(15, 100)

				# Find ViewTemplates and ParameterElementFitler to delete/keep
				viewtemplates_delete_set, paramfilters_delete_set = set(), set()

				# Get all ViewTemplates
				viewtemplates_all = revit.query.get_all_view_templates(doc=temp_doc)
				viewtemplates_all_set = set()

				if len(viewtemplates_all) > 0:
					viewtemplates_all_ids = [item.Id.IntegerValue for item in viewtemplates_all]
					viewtemplates_all_set = set(viewtemplates_all_ids)

					# Get ViewTemplates applied to Views to keep
					viewtemplates_keep_ids = [f_id_to_elem(temp_doc, elemid).ViewTemplateId.IntegerValue for elemid in views_keep_set]
					viewtemplates_keep_set = set(viewtemplates_keep_ids)

					# Get ViewTemplates to delete
					viewtemplates_delete_set = viewtemplates_all_set - viewtemplates_keep_set

				# Get all ParameterFilterElements
				paramfilters_all = revit.query.get_rule_filters(doc=temp_doc)
				paramfilters_all_set = set()
				
				if len(paramfilters_all) > 0:
					paramfilters_all_ids = [item.Id.IntegerValue for item in paramfilters_all]
					paramfilters_all_set = set(paramfilters_all_ids)

					# Get ParameterFilterElements applied to Views to keep
					paramfilters_keep_set = set()

					for elemid in views_keep_set:
						if f_id_to_elem(temp_doc, elemid).AreGraphicsOverridesAllowed():
							view_filters = f_id_to_elem(temp_doc, elemid).GetFilters()
							for vf in view_filters:
								paramfilters_keep_set.add(vf.IntegerValue)

					# Get ParameterFilterElements to delete
					paramfilters_delete_set = paramfilters_all_set - paramfilters_keep_set
				
				else:
					viewtemplates_delete_set = viewtemplates_all_set
					paramfilters_delete_set = paramfilters_all_set

				pb.update_progress(20, 100)

				# Find Worksets to delete/keep or Elements on Worksets (for RVT < 2023)
				worksets_delete = None
				if temp_doc.IsWorkshared:

					if worksets_name:

						# Collect user-created worksets only and filter according to provided inputs
						user_worksets = DB.FilteredWorksetCollector(temp_doc).OfKind(DB.WorksetKind.UserWorkset).ToWorksets()
						worksets_delete = [uw for uw in user_worksets if any(wn in uw.Name for wn in worksets_name)]

						# If RVT version is > 2022 the Workset (and its elements) will be deleted, otherwise only the Workset's elements will be deleted
						if rvt_version < 2023:

							# Construct MultiCategory Filter #1
							categories_filter_1 = []
							for cat in temp_doc.Settings.Categories:
								if (
									cat.CategoryType == DB.CategoryType.Model
									or cat.CategoryType == DB.CategoryType.Annotation
								):
									categories_filter_1.append(cat)
							categories_ids_1 = List[DB.ElementId]()
							for fc1 in categories_filter_1:
								categories_ids_1.Add(fc1.Id)
							categories_filter_1 = DB.ElementMulticategoryFilter(categories_ids_1)

							# Collect elements on Worksets
							elements_delete = []
							for wtd in worksets_delete:
								worksets_filter = DB.ElementWorksetFilter(wtd.Id)
								composed_filter_1 = DB.LogicalAndFilter(categories_filter_1, worksets_filter)
								elems_worksets = DB.FilteredElementCollector(temp_doc).WhereElementIsNotElementType().WherePasses(composed_filter_1)
								elements_delete.extend(elems_worksets)

				pb.update_progress(25, 100)

				# Delete Sheets, Views, ViewTemplates, ParameterFilterElements and Worksets (Elements on Worksets for RVT < 2023)
				error_elems = []

				with revit.Transaction(name='CleanupModel', doc=temp_doc, swallow_errors=False):

					# Delete Sheets
					if sheets_delete_set:
						for elemid in sheets_delete_set:
							item = f_id_to_elem(temp_doc, elemid)
							if item.IsValidObject:
								try:	temp_doc.Delete(item.Id)
								except:	error_elems.append((item.Id.IntegerValue, item.Category.Name, item.Name))

					# Delete Views
					if views_delete_set:
						for elemid in views_delete_set:
							item = f_id_to_elem(temp_doc, elemid)
							if item:
								if item.IsValidObject:
									try:	temp_doc.Delete(item.Id)
									except:	error_elems.append((item.Id.IntegerValue, item.Category.Name, item.Name))

					# Delete ViewTemplates
					if viewtemplates_delete_set:
						for elemid in viewtemplates_delete_set:
							item = f_id_to_elem(temp_doc, elemid)
							if item:
								if item.IsValidObject:
									try:	temp_doc.Delete(item.Id)
									except:	error_elems.append((item.Id.IntegerValue, item.Category.Name, item.Name))

					# Delete ParameterFilterElements
					if paramfilters_delete_set:
						for elemid in paramfilters_delete_set:
							item = f_id_to_elem(temp_doc, elemid)
							if item:
								if item.IsValidObject:
									try:	temp_doc.Delete(item.Id)
									except:	error_elems.append((item.Id.IntegerValue, item.Category.Name, item.Name))

					# Delete Worksets (Elements on Worksets for RVT < 2023)
					if worksets_delete:
						if rvt_version < 2023:
							for item in elements_delete:
								if item:
									if item.IsValidObject:
										try:	temp_doc.Delete(item.Id)
										except:	error_elems.append((item.Id.IntegerValue, item.Category.Name, item.Name))

						else:
							dws = DB.DeleteWorksetSettings()
							for item in worksets_delete:
								try:	DB.WorksetTable.DeleteWorkset(temp_doc, item.Id, dws)
								except: error_elems.append((item.Name, item.Id.IntegerValue))

				pb.update_progress(35, 100)

				# Specify options when saving and overwrite the RVT file
				save_opt = DB.SaveAsOptions()
				save_opt.Compact = True
				save_opt.OverwriteExistingFile = True

				# Add saving options for Workshared RVT file
				if temp_doc.IsWorkshared:
					worksharing_save_opt = DB.WorksharingSaveAsOptions()
					worksharing_save_opt.SaveAsCentral = True
					save_opt.SetWorksharingOptions(worksharing_save_opt)
					relinquish_opt = DB.RelinquishOptions(True)
					transact_opts = DB.TransactWithCentralOptions()
					DB.WorksharingUtils.RelinquishOwnership(temp_doc, relinquish_opt, transact_opts)

				# Check if there is at least one view in the project, otherwise create an empty drafting view
				views_check = revit.query.get_all_views(doc=temp_doc)
				if not views_check:
					view_fam_types = DB.FilteredElementCollector(temp_doc).OfClass(DB.ViewFamilyType).ToElements()
					view_draft_type = [item for item in view_fam_types if item.ViewFamily == DB.ViewFamily.Drafting][0]
					with revit.Transaction(name='CreateView', doc=temp_doc, swallow_errors=True):
						view_draft = DB.ViewDrafting.Create(temp_doc,view_draft_type.Id)
						view_draft.Name = 'Empty Drafting View'

				pb.update_progress(40, 100)

				# Purge Document if selected by using eTransmit for Revit add-in API
				purge_result = None
				if flex_form.values['cb_purge']:
					progress = 40
					purge_result_01 = []
					etuom = eTransmitUpgradeOMatic(app)
					tran_opt = TransmissionOptions()
					open_file = etuom.openUpgradeSave(model_path, tran_opt)
					
					purge_terations = 1
					for i in range(purge_terations):
						uft = etuom.purgeUnused(temp_doc)
						purge_result_01.append(str(uft))
						progress += int(45/purge_terations)
						pb.update_progress(progress, 100)
					
					purge_result = list(set(purge_result_01))[0]
	
				temp_doc.SaveAs(rvt_file, save_opt)

				# Save as detached if selected
				if flex_form.values['cb_detach'] and rvt_file_info.IsWorkshared:
					tr_data = DB.TransmissionData.ReadTransmissionData(model_path)
					tr_data.IsTransmitted = True
					DB.TransmissionData.WriteTransmissionData(model_path, tr_data)

				temp_doc.Close(False)
				temp_doc.Dispose()

				pb.update_progress(100, 100)
				
				end = time.time()
				exec_time = end - start

			out_rows.append((rvt_file, purge_result, exec_time, error_elems))
			
			# Delete backup folder for Workshared RVT file
			bk_folder_path = rvt_file.replace('.rvt', '_backup')
			shutil.rmtree(bk_folder_path)

		__revit__.Application.Dispose()

		## Print the output
		table_headers = ['Saved File Path', 'Purged Unused', 'Execution Time [s]', 'Error Elements']

		script_output.print_table(
			table_data = out_rows,
			title = 'Model(s) Cleanup & OverWrite',
			columns = table_headers
		)

	else:	script.exit()
else:	script.exit()