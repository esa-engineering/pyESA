__context__ = "zero-doc"

#REFERENCES
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


#DEFINITIONS
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

def f_sheets_check(elem, par_name, par_values):
	sc_delete = []
	sc_keep = []
	#Fetch the parameter
	par_elem = revit.query.get_param(elem, par_name)
	if par_elem:
		#Get parameter's value
		par_val = str(revit.query.get_param_value(par_elem))
		if par_values != None:
			#Collect the sheet and all its placed views in the sc_delete list
			if not any(pv in par_val for pv in par_values):
				sc_delete.append(elem)
				##Get all placed views
				for i in elem.GetAllPlacedViews():
					sc_delete.append(elem.Document.GetElement(i))
				##Get all placed schedules
				schedule_instances = DB.FilteredElementCollector(elem.Document, elem.Id).OfClass(DB.ScheduleSheetInstance).ToElements()
				schedules = [elem.Document.GetElement(si.ScheduleId) for si in schedule_instances]
				sc_delete.extend(schedules)
			else:
				sc_keep.append(elem)
				##Get all placed views
				for i in elem.GetAllPlacedViews():
					sc_keep.append(elem.Document.GetElement(i))
				##Get all placed schedules
				schedule_instances = DB.FilteredElementCollector(elem.Document, elem.Id).OfClass(DB.ScheduleSheetInstance).ToElements()
				schedules = [elem.Document.GetElement(si.ScheduleId) for si in schedule_instances]
				sc_keep.extend(schedules)
		else:
			sc_delete.append(elem)
			for i in elem.GetAllPlacedViews():
				sc_delete.append(elem.Document.GetElement(i))
	return sc_delete,sc_keep

def f_views_check(elem, par_name, par_values):
	vc_delete = None
	par_elem = revit.query.get_param(elem, par_name)
	if par_elem:
		par_val = str(revit.query.get_param_value(par_elem))
		if par_values != None:
			if not any(pv in par_val for pv in par_values):
				vc_delete = elem
		else:
				vc_delete = elem
	return vc_delete

#INPUTS
rvt_files = l_tolist(forms.pick_file(files_filter=	'Revit Files |*.rvt;*.rte;*.rfa|'
														'Revit Model |*.rvt|'
														'Revit Template |*.rte|'
														'Revit Family |*.rfa',
									multi_file=True,
									title='Select Revit File(s)'))

if rvt_files[0] != None:

	if __shiftclick__:
		components = [
		Label('WORKSETS - to delete (empty for RVT < 2023)'),
		Label('Workset name contains (use \';\' as separator):'),
		TextBox('txt_worksets', default='WIP'),
		Separator(),
		Label('VIEWS - to keep'),
		Label('View parameter name:'),
		TextBox('txt_view_param', default='u_ViewChapter1'),
		Label('Parameter value contains (use \';\' as separator):'),
		TextBox('txt_view_contains', default='200'),
		Separator(),
		Label('SHEETS - to keep'),
		Label('Sheet parameter name:'),
		TextBox('txt_sheet_param', default='u_ViewChapter1'),
		Label('Parameter value contains (use \';\' as separator):'),
		TextBox('txt_sheet_contains', default='00'),
		Separator(),
		CheckBox('cb_purge', 'Purge Unused', default=True),
		CheckBox('cb_detach', 'Create Transmit', default=True),
		Separator(),
		Button('OK')			
		]
	else:
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

	#CODE
	script_output = script.get_output()
	out_rows = []

	if	len(flex_form.values.items())>0:
		for rvt_file in rvt_files:
			
			with forms.ProgressBar(title=rvt_file.split('\\')[-1]) as pb:

				temp_name = rvt_file.split('\\')[-1]
				rvt_file_info = revit.files.get_file_info(rvt_file)

				pb.update_progress(5, 100)

				##Specify options when opening the original RVT file
				open_opt = DB.OpenOptions()
				##Add opening options for Workshared RVT file
				if rvt_file_info.IsWorkshared:					
					open_config = DB.WorksetConfiguration(DB.WorksetConfigurationOption.CloseAllWorksets)
					open_opt = DB.OpenOptions()
					open_opt.DetachFromCentralOption = DB.DetachFromCentralOption.DetachAndPreserveWorksets
					open_opt.SetOpenWorksetsConfiguration(open_config)				
				##Open the original RVT file
				model_path = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(rvt_file)
				temp_doc = __revit__.Application.OpenDocumentFile(model_path, open_opt)

				##Find sheets (with related placed views) to delete and keep
				sheets_views_todelete = None
				sheets_param_name = l_string_clean(flex_form.values['txt_sheet_param'])
				sheets_param_values = f_param_value_list(flex_form.values['txt_sheet_contains'])
				sheets_views_tokeep_ids = []
				if len(sheets_param_name)>0:
					sheets_views_todelete = []
					sheets_views_tokeep = []
					sheets_all = revit.query.get_sheets(doc=temp_doc)
					if len(sheets_all)>0:
						for sa in sheets_all:
							sheets_views_todelete.extend(f_sheets_check(sa, sheets_param_name, sheets_param_values)[0])
							sheets_views_tokeep.extend(f_sheets_check(sa, sheets_param_name, sheets_param_values)[1])
						sheets_views_tokeep_ids = [svtk.Id.IntegerValue for svtk in sheets_views_tokeep]

				pb.update_progress(10, 100)

				##Find views to delete
				views_todelete = None
				views_param_name = l_string_clean(flex_form.values['txt_view_param'])
				views_param_values = f_param_value_list(flex_form.values['txt_view_contains'])
				if len(views_param_name)>0:
					views_all = revit.query.get_all_views(doc=temp_doc)
					# views_all.extend(revit.query.get_all_schedules(doc=temp_doc))
					views_filtered = [va for va in views_all if va.GetType().ToString() != 'Autodesk.Revit.DB.ViewSheet']
					views_filtered = [vf for vf in views_filtered if vf.Id.IntegerValue not in sheets_views_tokeep_ids]
					views_todelete = [vf for vf in views_filtered if f_views_check(vf, views_param_name, views_param_values) != None]

				pb.update_progress(15, 100)

				##Find elements on worksets to empty (only for Workshared RVT file)
				elements_todelete = None
				worksets_todelete = None
				if temp_doc.IsWorkshared:
					worksets_name = f_param_value_list(flex_form.values['txt_worksets'])
					if worksets_name:
						###Collect user-created worksets only
						user_worksets = DB.FilteredWorksetCollector(temp_doc).OfKind(DB.WorksetKind.UserWorkset).ToWorksets()
						###If RVT version is > 2022 the Workset (and its elements) will be deleted, otherwise only the Workset's elements will be deleted
						worksets_todelete = [uw for uw in user_worksets if any(wn in uw.Name for wn in worksets_name)]

						if rvt_version < 2023:
							###Construct MultiCategory Filter #1
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

							###Collect elements on Worksets
							elements_todelete = []
							for wtd in worksets_todelete:
								worksets_filter = DB.ElementWorksetFilter(wtd.Id)
								composed_filter_1 = DB.LogicalAndFilter(categories_filter_1,worksets_filter)
								elems_worksets = DB.FilteredElementCollector(temp_doc).WhereElementIsNotElementType().WherePasses(composed_filter_1)
								elements_todelete.extend(elems_worksets)

				pb.update_progress(20, 100)

				##Delete Sheets, Views and elements on Worksets (or Worksets if RVT version is > 2022)
				error_elems = []
				with revit.Transaction(name='CleanupModel', doc=temp_doc, swallow_errors=False):
					if sheets_views_todelete:
						for item in sheets_views_todelete:
							if item.IsValidObject:
								try:	temp_doc.Delete(item.Id)
								except:	error_elems.append((item.Name, item.Id.IntegerValue))

					if views_todelete:
						for item in views_todelete:
							if item.IsValidObject:
								try:	temp_doc.Delete(item.Id)
								except:	error_elems.append((item.Name, item.Id.IntegerValue))

					if worksets_todelete:
						if rvt_version > 2022:
							dws = DB.DeleteWorksetSettings()
							for wtd in worksets_todelete:
								DB.WorksetTable.DeleteWorkset(temp_doc, wtd.Id, dws)
						else:
							if elements_todelete:
								for item in elements_todelete:
									if item.IsValidObject:
										try:	temp_doc.Delete(item.Id)
										except:	error_elems.append((item.Name, item.Id.IntegerValue))

				pb.update_progress(30, 100)

				##Specify options when saving and overwrite the RVT file
				save_opt = DB.SaveAsOptions()
				save_opt.Compact = True
				save_opt.OverwriteExistingFile = True

				if temp_doc.IsWorkshared:
					###Add saving options for Workshared RVT file
					worksharing_save_opt = DB.WorksharingSaveAsOptions()
					worksharing_save_opt.SaveAsCentral = True
					save_opt.SetWorksharingOptions(worksharing_save_opt)
					relinquish_opt = DB.RelinquishOptions(True)
					transact_opts = DB.TransactWithCentralOptions()
					DB.WorksharingUtils.RelinquishOwnership(temp_doc, relinquish_opt, transact_opts)

				##Check if there is at least one view in the project, otherwise create an empty drafting view
				views_check = revit.query.get_all_views(doc=temp_doc)
				if not views_check:
					view_fam_types = DB.FilteredElementCollector(temp_doc).OfClass(DB.ViewFamilyType).ToElements()
					view_draft_type = [item for item in view_fam_types if item.ViewFamily == DB.ViewFamily.Drafting][0]
					with revit.Transaction(name='CreateView', doc=temp_doc, swallow_errors=True):
						view_draft = DB.ViewDrafting.Create(temp_doc,view_draft_type.Id)
						view_draft.Name = 'Empty Drafting View'

				pb.update_progress(40, 100)

				##Purge Document if selected by using eTransmit for Revit add-in API
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

				##Save as detached if selected
				if flex_form.values['cb_detach'] and rvt_file_info.IsWorkshared:
					tr_data = DB.TransmissionData.ReadTransmissionData(model_path)
					tr_data.IsTransmitted = True
					DB.TransmissionData.WriteTransmissionData(model_path, tr_data)

				temp_doc.Close(False)
				temp_doc.Dispose()

				pb.update_progress(100, 100)

			out_rows.append((rvt_file, purge_result, error_elems))

		__revit__.Application.Dispose()
		
		##Print the output
		table_headers = ['Saved File Path', 'Purged Unused', 'Error Elements']

		script_output.print_table(
			table_data = out_rows,
			title = 'Model(s) Cleanup & OverWrite',
			columns = table_headers
		)

	else:	script.exit()
else:	script.exit()