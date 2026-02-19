__context__ = "zero-doc"

# REFERENCES
import shutil
import time
import pyrevit
import System

from pyrevit import revit, DB, UI, script, forms
from pyrevit import PyRevitException, PyRevitIOError

from rpw.ui.forms import TaskDialog, FlexForm, Label, ComboBox, TextBox, CheckBox, Separator, Button

from System.Collections.Generic import List, HashSet

import clr
import sys

hostapp = pyrevit._HostApplication()
app = hostapp.app
rvt_version = int(hostapp.version)

# ============================================================================
# COMPATIBILITY LAYER FOR REVIT 2024+ (ElementId changes from Int32 to Int64)
# ============================================================================

def get_element_id_value(element_id):
	"""
	Get the integer value from an ElementId.
	Handles both old (.IntegerValue) and new (.Value) API.
	- Revit < 2024: Uses IntegerValue (Int32)
	- Revit >= 2024: Uses Value (Int64)
	"""
	if rvt_version >= 2024:
		return element_id.Value
	else:
		return element_id.IntegerValue

def create_element_id(id_value):
	"""
	Create an ElementId from an integer value.
	Handles both old (Int32) and new (Int64) constructor.
	- Revit < 2024: ElementId(int)
	- Revit >= 2024: ElementId(long)
	"""
	if rvt_version >= 2024:
		# Use Int64 (long) for Revit 2024+
		return DB.ElementId(System.Int64(id_value))
	else:
		# Use Int32 (int) for older versions
		return DB.ElementId(int(id_value))

# ============================================================================
# PURGE METHODS (Multiple fallback options)
# ============================================================================

def purge_using_native_api(doc, iterations=3):
	"""
	Purge unused elements using the native GetUnusedElements API.
	Available in Revit 2024+.
	Returns tuple: (success, message)
	"""
	if rvt_version < 2024:
		return (False, "Native API not available (Revit < 2024)")
	
	try:
		total_deleted = 0
		for i in range(iterations):
			# Get all unused elements (empty HashSet means all categories)
			unused_ids = doc.GetUnusedElements(HashSet[DB.ElementId]())
			if unused_ids and len(unused_ids) > 0:
				# Convert to list for deletion
				ids_to_delete = list(unused_ids)
				doc.Delete(List[DB.ElementId](ids_to_delete))
				total_deleted += len(ids_to_delete)
			else:
				break
		
		return (True, "Purged {} elements (Native API)".format(total_deleted))
	except Exception as e:
		return (False, "Native API Error: " + str(e))

def purge_using_performance_adviser(doc, iterations=3):
	"""
	Purge unused elements using PerformanceAdviser.
	Available in Revit 2017+.
	Returns tuple: (success, message)
	"""
	try:
		# GUID for "Project contains unused families and types" rule
		purge_guid = "e8c63650-70b7-435a-9010-ec97660c1bda"
		
		# Find the rule
		all_rule_ids = DB.PerformanceAdviser.GetPerformanceAdviser().GetAllRuleIds()
		purge_rule_id = None
		
		for rule_id in all_rule_ids:
			if str(rule_id.Guid) == purge_guid:
				purge_rule_id = rule_id
				break
		
		if purge_rule_id is None:
			return (False, "PerformanceAdviser purge rule not found")
		
		total_deleted = 0
		for i in range(iterations):
			# Execute the purge rule
			rule_ids = List[DB.PerformanceAdviserRuleId]()
			rule_ids.Add(purge_rule_id)
			
			failure_messages = DB.PerformanceAdviser.GetPerformanceAdviser().ExecuteRules(doc, rule_ids)
			
			if failure_messages and failure_messages.Count > 0:
				# Get purgeable elements from the failure message
				purgeable_ids = failure_messages[0].GetFailingElements()
				if purgeable_ids and purgeable_ids.Count > 0:
					doc.Delete(purgeable_ids)
					total_deleted += purgeable_ids.Count
				else:
					break
			else:
				break
		
		return (True, "Purged {} elements (PerformanceAdviser)".format(total_deleted))
	except Exception as e:
		return (False, "PerformanceAdviser Error: " + str(e))

def purge_using_etransmit(doc, app_instance, model_path):
	"""
	Purge unused elements using eTransmit API.
	Fallback method - may not work in Revit 2025+ due to .NET 8.
	Returns tuple: (success, message)
	"""
	# Try to load eTransmit
	etransmit_paths = [
		"C:\\Program Files\\Autodesk\\eTransmit for Revit " + str(rvt_version),
		"C:\\Program Files\\Autodesk\\eTransmit for Revit\\" + str(rvt_version),
		"C:\\Program Files (x86)\\Autodesk\\eTransmit for Revit " + str(rvt_version),
	]
	
	for transmit_path in etransmit_paths:
		try:
			if transmit_path not in sys.path:
				sys.path.append(transmit_path)
			clr.AddReference('eTransmitForRevitDB')
			from eTransmitForRevitDB import eTransmitUpgradeOMatic, TransmissionOptions, UpgradeFailureType
			
			etuom = eTransmitUpgradeOMatic(app_instance)
			tran_opt = TransmissionOptions()
			open_file = etuom.openUpgradeSave(model_path, tran_opt)
			
			result = etuom.purgeUnused(doc)
			
			if str(result) == "UpgradeSucceeded":
				return (True, "Purged (eTransmit)")
			else:
				return (False, "eTransmit result: " + str(result))
		except:
			continue
	
	return (False, "eTransmit not available")

def purge_document(doc, app_instance=None, model_path=None, iterations=3):
	"""
	Main purge function that tries multiple methods in order of preference:
	1. Native API (Revit 2024+)
	2. PerformanceAdviser (Revit 2017+)
	3. eTransmit (fallback)
	
	Returns tuple: (success, message)
	"""
	# Method 1: Try Native API first (Revit 2024+)
	if rvt_version >= 2024:
		success, message = purge_using_native_api(doc, iterations)
		if success:
			return (success, message)
	
	# Method 2: Try PerformanceAdviser
	success, message = purge_using_performance_adviser(doc, iterations)
	if success:
		return (success, message)
	
	# Method 3: Try eTransmit as last resort
	if app_instance and model_path:
		success, message = purge_using_etransmit(doc, app_instance, model_path)
		if success:
			return (success, message)
	
	return (False, "All purge methods failed")

# ============================================================================
# HELPER DEFINITIONS (Updated for ElementId compatibility)
# ============================================================================

l_tolist = lambda x: x if hasattr(x, '__iter__') else [x]
l_string_clean = lambda x: '' if x == ' ' else x

# Updated lambda to use compatibility function
l_list_elem_id = lambda lst: [get_element_id_value(x.Id) for x in lst]

def f_param_value_list(entry_string):
	clean_string = l_string_clean(entry_string)
	if len(clean_string) > 0:
		list_string = clean_string.replace('; ', ';').replace(' ;', ';').split(';')
		return l_tolist(list_string)
	else:
		return None

def f_param_check(elem, param, values):
	"""
	Return True if ANY value of the checked parameter matches
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
	"""
	Convert an integer ID to a Revit Element.
	Uses compatibility layer for ElementId creation.
	"""
	return doc.GetElement(create_element_id(id))

def f_get_all_views_on_sheet(doc, sheet, elemID=True):
	"""
	Return Views and Schedules (OR their IDs as Integer Values
	if 'elemID' is set to True) placed on the Sheet.
	Updated to use compatibility functions.
	"""
	sheet_views_all = []
	sheet_views = sheet.GetAllPlacedViews()
	
	if elemID:
		sheet_views_all = [get_element_id_value(item) for item in sheet_views]
	else:
		sheet_views_all = [doc.GetElement(item) for item in sheet_views]
	
	# Get Schedules placed on the sheet (Revision Schedule is included)
	sheet_schedules_placed = DB.FilteredElementCollector(doc, sheet.Id).OfClass(DB.ScheduleSheetInstance).ToElements()

	if elemID:
		sheet_schedules = [get_element_id_value(item.ScheduleId) for item in sheet_schedules_placed]
	else:
		sheet_schedules = [doc.GetElement(item.ScheduleId) for item in sheet_schedules_placed]

	sheet_views_all.extend(sheet_schedules)
	return sheet_views_all

# ============================================================================
# INPUTS
# ============================================================================

rvt_files = l_tolist(forms.pick_file(files_filter='Revit files (*.rvt)|*.rvt',
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

	if len(flex_form.values.items()) > 0:
		
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

				# Instantiate the containers for Views and Sheets
				sheets_keep_set, sheets_delete_set = set(), set()
				views_keep_set, views_delete_set = set(), set()

				# Find sheets to delete/keep
				sheets_all = revit.query.get_sheets(doc=temp_doc)
				sheets_all_ids = [get_element_id_value(item.Id) for item in sheets_all]
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
				views_all_ids = [get_element_id_value(item.Id) for item in views_all]
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

				# Find ViewTemplates and ParameterFilterElements to delete/keep
				viewtemplates_delete_set, paramfilters_delete_set = set(), set()

				# Get all ViewTemplates
				viewtemplates_all = revit.query.get_all_view_templates(doc=temp_doc)
				viewtemplates_all_set = set()

				if len(viewtemplates_all) > 0:
					viewtemplates_all_ids = [get_element_id_value(item.Id) for item in viewtemplates_all]
					viewtemplates_all_set = set(viewtemplates_all_ids)

					# Get ViewTemplates applied to Views to keep
					viewtemplates_keep_ids = [get_element_id_value(f_id_to_elem(temp_doc, elemid).ViewTemplateId) for elemid in views_keep_set]
					viewtemplates_keep_set = set(viewtemplates_keep_ids)

					# Get ViewTemplates to delete
					viewtemplates_delete_set = viewtemplates_all_set - viewtemplates_keep_set

				# Get all ParameterFilterElements
				paramfilters_all = revit.query.get_rule_filters(doc=temp_doc)
				paramfilters_all_set = set()
				
				if len(paramfilters_all) > 0:
					paramfilters_all_ids = [get_element_id_value(item.Id) for item in paramfilters_all]
					paramfilters_all_set = set(paramfilters_all_ids)

					# Get ParameterFilterElements applied to Views to keep
					paramfilters_keep_set = set()

					for elemid in views_keep_set:
						if f_id_to_elem(temp_doc, elemid).AreGraphicsOverridesAllowed():
							view_filters = f_id_to_elem(temp_doc, elemid).GetFilters()
							for vf in view_filters:
								paramfilters_keep_set.add(get_element_id_value(vf))

					# Get ParameterFilterElements to delete
					paramfilters_delete_set = paramfilters_all_set - paramfilters_keep_set
				
				else:
					viewtemplates_delete_set = viewtemplates_all_set
					paramfilters_delete_set = paramfilters_all_set

				pb.update_progress(20, 100)

				# Find Worksets to delete/keep or Elements on Worksets (for RVT < 2023)
				worksets_delete = None
				elements_delete = []
				
				if temp_doc.IsWorkshared:

					if worksets_name:

						# Collect user-created worksets only and filter according to provided inputs
						user_worksets = DB.FilteredWorksetCollector(temp_doc).OfKind(DB.WorksetKind.UserWorkset).ToWorksets()
						worksets_delete = [uw for uw in user_worksets if any(wn in uw.Name for wn in worksets_name)]

						# If RVT version is > 2022 the Workset (and its elements) will be deleted,
						# otherwise only the Workset's elements will be deleted
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
							for wtd in worksets_delete:
								worksets_filter = DB.ElementWorksetFilter(wtd.Id)
								composed_filter_1 = DB.LogicalAndFilter(categories_filter_1, worksets_filter)
								elems_worksets = DB.FilteredElementCollector(temp_doc).WhereElementIsNotElementType().WherePasses(composed_filter_1)
								elements_delete.extend(elems_worksets)

				pb.update_progress(25, 100)

				# Delete Sheets, Views, ViewTemplates, ParameterFilterElements and Worksets
				# (Elements on Worksets for RVT < 2023)
				error_elems = []

				with revit.Transaction(name='CleanupModel', doc=temp_doc, swallow_errors=False):

					# Delete Sheets
					if sheets_delete_set:
						for elemid in sheets_delete_set:
							item = f_id_to_elem(temp_doc, elemid)
							if item.IsValidObject:
								try:
									temp_doc.Delete(item.Id)
								except:
									error_elems.append((get_element_id_value(item.Id), item.Category.Name, item.Name))

					# Delete Views
					if views_delete_set:
						for elemid in views_delete_set:
							item = f_id_to_elem(temp_doc, elemid)
							if item:
								if item.IsValidObject:
									try:
										temp_doc.Delete(item.Id)
									except:
										error_elems.append((get_element_id_value(item.Id), item.Category.Name, item.Name))

					# Delete ViewTemplates
					if viewtemplates_delete_set:
						for elemid in viewtemplates_delete_set:
							item = f_id_to_elem(temp_doc, elemid)
							if item:
								if item.IsValidObject:
									try:
										temp_doc.Delete(item.Id)
									except:
										error_elems.append((get_element_id_value(item.Id), item.Category.Name, item.Name))

					# Delete ParameterFilterElements
					if paramfilters_delete_set:
						for elemid in paramfilters_delete_set:
							item = f_id_to_elem(temp_doc, elemid)
							if item:
								if item.IsValidObject:
									try:
										temp_doc.Delete(item.Id)
									except:
										error_elems.append((get_element_id_value(item.Id), item.Category.Name, item.Name))

					# Delete Worksets (Elements on Worksets for RVT < 2023)
					if worksets_delete:
						if rvt_version < 2023:
							for item in elements_delete:
								if item:
									if item.IsValidObject:
										try:
											temp_doc.Delete(item.Id)
										except:
											error_elems.append((get_element_id_value(item.Id), item.Category.Name, item.Name))

						else:
							dws = DB.DeleteWorksetSettings()
							for item in worksets_delete:
								try:
									DB.WorksetTable.DeleteWorkset(temp_doc, item.Id, dws)
								except:
									error_elems.append((item.Name, get_element_id_value(item.Id)))

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

				# Check if there is at least one view in the project,
				# otherwise create an empty drafting view
				views_check = revit.query.get_all_views(doc=temp_doc)
				if not views_check:
					view_fam_types = DB.FilteredElementCollector(temp_doc).OfClass(DB.ViewFamilyType).ToElements()
					view_draft_type = [item for item in view_fam_types if item.ViewFamily == DB.ViewFamily.Drafting][0]
					with revit.Transaction(name='CreateView', doc=temp_doc, swallow_errors=True):
						view_draft = DB.ViewDrafting.Create(temp_doc, view_draft_type.Id)
						view_draft.Name = 'Empty Drafting View'

				pb.update_progress(40, 100)

				# Purge Document if selected
				purge_result = None
				if flex_form.values.get('cb_purge', False):
					try:
						with revit.Transaction(name='PurgeUnused', doc=temp_doc, swallow_errors=False):
							success, purge_result = purge_document(
								temp_doc, 
								app_instance=app, 
								model_path=model_path,
								iterations=3
							)
					except Exception as purge_error:
						purge_result = "Error: " + str(purge_error)

				pb.update_progress(85, 100)
	
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
			try:
				shutil.rmtree(bk_folder_path)
			except:
				pass  # Backup folder may not exist

		__revit__.Application.Dispose()

		# Print the output
		table_headers = ['Saved File Path', 'Purged Unused', 'Execution Time [s]', 'Error Elements']

		script_output.print_table(
			table_data=out_rows,
			title='Model(s) Cleanup & OverWrite',
			columns=table_headers
		)

	else:
		script.exit()
else:
	script.exit()