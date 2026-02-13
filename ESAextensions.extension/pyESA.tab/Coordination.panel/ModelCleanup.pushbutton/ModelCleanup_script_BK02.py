__context__ = "zero-doc"

#REFERENCES
import pyrevit
import System

from pyrevit import revit, DB, UI, script
from pyrevit import PyRevitException, PyRevitIOError
from pyrevit import forms

from rpw.ui.forms import TaskDialog, FlexForm, Label, ComboBox, TextBox, CheckBox, Separator, Button

from System.Collections.Generic import List

#DEFINITIONS
l_tolist = lambda x: x if hasattr(x, '__iter__') else [x]
l_string_clean = lambda x: '' if x == ' ' else x
l_list_elem_id = lambda lst: [x.Id.IntegerValue for x in lst]

def f_purgeable_elems(doc, rule_id_list):
	failure_messages = DB.PerformanceAdviser.GetPerformanceAdviser().ExecuteRules(doc, rule_id_list)
	if failure_messages.Count > 0:
		purge_elem_ids = failure_messages[0].GetFailingElements()
		return purge_elem_ids

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

	components = [
	Label('WORKSETS - to empty'),
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

	new_paths = []
	error_elements = []

	#CODE
	if	len(flex_form.values.items())>0:
		for rvt_file in rvt_files:

			with forms.ProgressBar(title=rvt_file.split('\\')[-1]) as pb:

				temp_name = rvt_file.split('\\')[-1]
				rvt_file_info = revit.files.get_file_info(rvt_file)

				pb.update_progress(5, 100)

				##Specify options when opening the original RVT file
				open_opt = DB.OpenOptions()
				if rvt_file_info.IsWorkshared:
					##Add opening options for Workshared RVT file
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
				elements_todelete = False
				if temp_doc.IsWorkshared:
					worksets_name = f_param_value_list(flex_form.values['txt_worksets'])
					if worksets_name:

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
						if temp_doc.IsWorkshared:
							user_worksets = DB.FilteredWorksetCollector(temp_doc).OfKind(DB.WorksetKind.UserWorkset).ToWorksets()
							worksets_todelete = [uw for uw in user_worksets if any(wn in uw.Name for wn in worksets_name)]
							for wtd in worksets_todelete:
								worksets_filter = DB.ElementWorksetFilter(wtd.Id)
								composed_filter_1 = DB.LogicalAndFilter(categories_filter_1,worksets_filter)
								elems_worksets = DB.FilteredElementCollector(temp_doc).WhereElementIsNotElementType().WherePasses(composed_filter_1)
								elements_todelete.extend(elems_worksets)

				pb.update_progress(20, 100)

				##Delete Sheets, Views and elements on Worksets
				error_elems = []
				with revit.Transaction(name='CleanupModel', doc=temp_doc, swallow_errors=True):
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
					if elements_todelete:
						for item in elements_todelete:
							if item.IsValidObject:
								try:	temp_doc.Delete(item.Id)
								except:	error_elems.append((item.Name, item.Id.IntegerValue))

				pb.update_progress(30, 100)

				##Find purgeable elements using the PerformanceAdviser
				purge_elem_ids = None
				purge_elems = None
				if flex_form.values['cb_purge']:
					for iteration in range(4):
						purge_guid = "e8c63650-70b7-435a-9010-ec97660c1bda"
						rule_id_list = List[DB.PerformanceAdviserRuleId]()

						for rule_id in DB.PerformanceAdviser.GetPerformanceAdviser().GetAllRuleIds():
							if str(rule_id.Guid) == purge_guid:
								rule_id_list.Add(rule_id)
								break
						purge_elem_ids = f_purgeable_elems(temp_doc, rule_id_list)
						if purge_elem_ids:
							purge_elems = [temp_doc.GetElement(pei) for pei in purge_elem_ids]

						###Find materials to delete
						mats_todelete_ids = None
						mats_todelete = None
						####Find materials ids to keep first
						elem_types = DB.FilteredElementCollector(temp_doc).WhereElementIsElementType().ToElements()
						mats_to_keep = []
						for elem in elem_types:
							mats_ids=[]
							for mat_id in elem.GetMaterialIds(False):
								mats_ids.append(mat_id)
							mats_to_keep.extend(mats_ids)

						compound = DB.FilteredElementCollector(temp_doc).OfClass(DB.HostObjAttributes).WhereElementIsElementType().ToElements()
						compound_structure = [comp.GetCompoundStructure() for comp in compound]
						for cs in compound_structure:
							if cs != None:
								n_layer = cs.LayerCount
								for j in range (0, n_layer):
									if cs.GetMaterialId(j) != DB.ElementId(-1):
										mats_to_keep.append(cs.GetMaterialId(j))

						####Finad all materials and assets ids
						all_mats_ids = DB.FilteredElementCollector(temp_doc).OfClass(DB.Material).ToElementIds()
						all_assets_ids = DB.FilteredElementCollector(temp_doc).OfClass(DB.AppearanceAssetElement).ToElementIds()
						all_pset_ids = DB.FilteredElementCollector(temp_doc).OfClass(DB.PropertySetElement).ToElementIds()
						thermal_asset_ids = [temp_doc.GetElement(mat_id).ThermalAssetId for mat_id in set(mats_to_keep)]
						structural_asset_ids  =[temp_doc.GetElement(mat_id).StructuralAssetId for mat_id in set(mats_to_keep)]
						all_appearance_asset_ids=[temp_doc.GetElement(mat_id).AppearanceAssetId for mat_id in set(mats_to_keep)]
						pset_ids = [e for e in all_pset_ids if e not in thermal_asset_ids and e not in structural_asset_ids]
						appearance_asset_ids = [e for e in all_assets_ids if e not in all_appearance_asset_ids]

						####Finally, exclude materials to keep from all materials
						mats_todelete_ids = [m for m in all_mats_ids if not m in mats_to_keep]
						mats_todelete_ids.extend(appearance_asset_ids)
						mats_todelete_ids.extend(pset_ids)
						if mats_todelete_ids:
							mats_todelete = [temp_doc.GetElement(mti) for mti in mats_todelete_ids]

						##Delete prugeable elements and materials
						with revit.Transaction(name='PurgeElements', doc=temp_doc, swallow_errors=True):
							if purge_elems:
								for item in purge_elems:
									if item.IsValidObject:
										try:	temp_doc.Delete(item.Id)
										except:	error_elems.append((item.Name, item.Id.IntegerValue))

							if mats_todelete:
								for item in mats_todelete:
									if item.IsValidObject:
										try:	temp_doc.Delete(item.Id)
										except:	error_elems.append((item.Name, item.Id.IntegerValue))

							pb.update_progress(30+15*(iteration+1), 100)

				error_elements.append(error_elems)

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

				views_check = revit.query.get_all_views(doc=temp_doc)
				if not views_check:
					view_fam_types = DB.FilteredElementCollector(temp_doc).OfClass(DB.ViewFamilyType).ToElements()
					view_draft_type = [item for item in view_fam_types if item.ViewFamily == DB.ViewFamily.Drafting][0]
					with revit.Transaction(name='CreateView', doc=temp_doc, swallow_errors=True):
						view_draft = DB.ViewDrafting.Create(temp_doc,view_draft_type.Id)
						view_draft.Name = 'Empty Drafting View'

				pb.update_progress(90, 100)

				temp_doc.SaveAs(rvt_file, save_opt)

				##Create detached if selected
				if flex_form.values['cb_detach'] and rvt_file_info.IsWorkshared:
					tr_data = DB.TransmissionData.ReadTransmissionData(model_path)
					tr_data.IsTransmitted = True
					DB.TransmissionData.WriteTransmissionData(model_path, tr_data)

				temp_doc.Close(False)
				new_paths.append('SAVED PATH:\n' + rvt_file)

				__revit__.Application.Dispose()

				pb.update_progress(100, 100)

		##Print the output
		for np, error_elem in zip(new_paths, error_elements):
			print('----------')
			print(np)
			if len(error_elem)>0:
				print('NON DELETED ELEMENTS:')
				for ee in error_elem:
					print(ee)
		print('----------')

	else:	script.exit()
else:	script.exit()