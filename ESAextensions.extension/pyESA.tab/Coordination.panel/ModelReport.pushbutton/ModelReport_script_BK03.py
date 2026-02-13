__doc__ =	'Generate the Model Report for the active document\n'\
				'---\n'\
				'SHIFT-CLICK Generate the Model Report for the selected link'
__title__ = 'Model\nReport'

#REFERENCES
import csv
import os
import math
import datetime
import System

import pyrevit
from pyrevit import script, revit, DB, UI
#from pyrevit import output
from pyrevit import forms

from System.Collections.Generic import List

uidoc = __revit__.ActiveUIDocument
time_start = datetime.datetime.now()

#DEFINITIONS
l_tolist = lambda x: x if hasattr(x, '__iter__') else [x]
l_safe_str = lambda x: str(x).encode('ascii','replace').replace('\r',' ').replace('\n',' ')
l_bool_str = lambda x: 'True' if x == True else 'False'
l_nest = lambda fx, lst: [fx(i) if not isinstance(i, list) else l_nest(fx, i) for i in lst]

##Get Revit document path
def f_doc_path(docum, sclick):
	if sclick:
		doc_path = l_safe_str(docum.PathName)
	else:
		if docum.IsWorkshared:
			doc_path = l_safe_str(DB.BasicFileInfo.Extract(docum.PathName).CentralPath)
		else:
			doc_path = l_safe_str(docum.PathName)
	return doc_path

#Convert file size from bytes
def f_convert_size(size_bytes):
	if not size_bytes:
		return "N/A"
	size_unit = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
	i = int(math.floor(math.log(size_bytes, 1024)))
	p = math.pow(1024, i)
	size = round(size_bytes / p, 2)
	return "{} {}".format(size, size_unit[i])

##Convert measurement from internal units to the users selected ones
def f_convert_units(measure,um):
	conversion = 'ConversionError'
	try:
		if um == 'm':
			conversion = DB.UnitUtils.ConvertFromInternalUnits(measure,DB.DisplayUnitType.DUT_METERS)
		elif um == 'mq':
			conversion = DB.UnitUtils.ConvertFromInternalUnits(measure,DB.DisplayUnitType.DUT_SQUARE_METERS)
		elif um == 'mc':
			conversion = DB.UnitUtils.ConvertFromInternalUnits(measure,DB.DisplayUnitType.DUT_CUBIC_METERS)
		elif um == 'rad':
			conversion = DB.UnitUtils.ConvertFromInternalUnits(measure,DB.DisplayUnitType.DUT_DECIMAL_DEGREES)
	except:	pass
	return conversion

##Convert coordinates of point in meters
def f_convert_point(point):
	point_m = 'ConversionError'
	try:
		point_m = (\
		DB.UnitUtils.ConvertFromInternalUnits(point.X,DB.DisplayUnitType.DUT_METERS),\
		DB.UnitUtils.ConvertFromInternalUnits(point.Y,DB.DisplayUnitType.DUT_METERS),\
		DB.UnitUtils.ConvertFromInternalUnits(point.Z,DB.DisplayUnitType.DUT_METERS))
	except:	pass
	return point_m

##Get the starting view of the document
def f_start_view(docum):
	return docum.GetElement(DB.StartingViewSettings.GetStartingViewSettings(docum).ViewId)

##Get Project Shared Parameters' names and GUIDs
def f_proj_shared_params(docum):
	par_names = []
	par_guids = []
	par_iterator = docum.ParameterBindings.ForwardIterator()
	while par_iterator.MoveNext():
		par_elem = docum.GetElement(par_iterator.Key.Id)
		if par_elem.GetType().ToString() == 'Autodesk.Revit.DB.SharedParameterElement':
			par_names.append(l_safe_str(par_iterator.Key.Name))
			par_guids.append(l_safe_str(par_elem.GuidValue))
	return par_names, par_guids

##Get Rooms/Areas/Spaces classification
def f_spatialelems_classification(spatiaelem):
	se_option = DB.SpatialElementBoundaryOptions()
	se_classification = 'Placed'
	if (spatiaelem.Area == 0 and spatiaelem.Location == None):
		se_classification = 'Not Placed'
	elif (spatiaelem.Area == 0 and len(spatiaelem.GetBoundarySegments(se_option))>0):
		se_classification = 'Redundant'
	elif (spatiaelem.Area == 0 and len(spatiaelem.GetBoundarySegments(se_option))==0):
		se_classification = 'Not Enclosed'
	return se_classification

##Attempt to find the host level name of an element
def f_elem_level(elem):
	elem_level = 'Undefined'
	try:	elem_level = elem.Document.GetElement(elem.LevelId).Name
	except:
		try:
			if elem.Host.GetType().ToString() == 'Autodesk.Revit.DB.Level':
				elem_level = elem.Host.Name
		except:
			try:
				level_id = elem.get_Parameter(BIP.STAIRS_BASE_LEVEL_PARAM).AsElementId()
				elem_level = elem.Document.GetElement(level_id).Name
			except:
				try:
					host_level = elem.Document.GetElement(elem.OwnerViewId).GenLevel
					if host_level.GetType().ToString() == 'Autodesk.Revit.DB.Level':
						elem_level = host_level.Name
				except:	pass
	return elem_level

#INPUTS
##Chapters of the Model Report
chapters =	(
	'01 GENERAL INFO',
	'02 PROJECT INFO',
	'03 COORDINATION',
	'04 WORKSETS',
	'05 EXTERNAL REFERENCES',
	'06 WARNINGS',
	'07 MODEL OBJECTS',
	'08 SPATIAL ELEMENTS',
	'09 DOCUMENTATION',
	'10 RECORD'
	)

BIC = DB.BuiltInCategory
BIP = DB.BuiltInParameter

tolerance = int(5)

##Select document
if __shiftclick__:
	with forms.WarningBar(title='Select RVT Link'):
		try:
			link_inst = revit.pick_element_by_category(BIC.OST_RvtLinks,message='Select RVT Link')
		except:
			script.exit()
	if link_inst:
		rvtdoc = link_inst.GetLinkDocument()
		shift = True
	else:
		script.exit()
else:
	rvtdoc = revit.doc
	shift = False


#CODE
doc_path = f_doc_path(rvtdoc,shift)
##Headers of CSV file
csv_list = []
csv_headers = (
	'Chapter', 'Paragraph', 'Category', 'Id',
	'Text1_Title', 'Text1_Value', 'Text2_Title', 'Text2_Value', 'Text3_Title', 'Text3_Value',
	'Number1_Title', 'Number1_Value', 'Number2_Title', 'Number2_Value', 'Number3_Title', 'Number3_Value')
csv_list.append(csv_headers)

##Specify output CSV file
tday = datetime.date.today()
csv_path = forms.save_file(file_ext='csv',default_name='{0}_ESA-ModelReport_{1}.csv'.format(str(tday),str(doc_path.split('\\')[-1].replace('.rvt',''))))

if csv_path:

	##Construct MultiCategory Filter #1
	categories_filter_1 = []
	for cat in rvtdoc.Settings.Categories:
		if (
			cat.CategoryType == DB.CategoryType.Model
			or cat.CategoryType == DB.CategoryType.Annotation
		):
			categories_filter_1.append(cat)
	categories_ids_1 = List[DB.ElementId]()
	for fc1 in categories_filter_1:
		categories_ids_1.Add(fc1.Id)
	categories_filter_1 = DB.ElementMulticategoryFilter(categories_ids_1)

	##Construct MultiCategory Filter #2
	categories_filter_2 = []
	for cat in rvtdoc.Settings.Categories:
		if (
			cat.CategoryType == DB.CategoryType.Model
			or cat.CategoryType == DB.CategoryType.Annotation
			or cat.CategoryType == DB.CategoryType.Internal
		):
			if (
				cat.Id.IntegerValue != BIC.OST_RvtLinks.value__
				and cat.Id.IntegerValue != BIC.OST_ProjectInformation.value__
				and cat.Id.IntegerValue != BIC.OST_PipingSystem.value__
				and cat.Id.IntegerValue != BIC.OST_DuctSystem.value__
				and 'dwg' not in cat.Name
			):
				categories_filter_2.append(cat)
	categories_ids_2 = List[DB.ElementId]()
	for fc2 in categories_filter_2:
		categories_ids_2.Add(fc2.Id)
	categories_filter_2 = DB.ElementMulticategoryFilter(categories_ids_2)

	chapter_10 = chapters[9]


	##		CHAPTER 01
	chapter_01 = chapters[0]

	###	01_01 File Name
	row_01_01 = (chapter_01, 'File Name', '',	'',
	'', doc_path.split('\\')[-1])
	csv_list.append(row_01_01)

	###	01_02 File Size
	row_01_02 = (chapter_01, 'File Size', '',	'',
	'', f_convert_size(os.path.getsize(doc_path)))
	csv_list.append(row_01_02)
	###	10_XX Record
	row_01_02 = (chapter_10, 'File Size', '',	'',
	'', '', '', '', '', '',
	'', str(f_convert_size(os.path.getsize(doc_path)).split(' ')[0]))
	csv_list.append(row_01_02)

	###	01_03 Workshared
	row_01_03 = (chapter_01, 'Workshared', '', '',
	'', l_safe_str(rvtdoc.IsWorkshared))
	csv_list.append(row_01_03)

	###	01_04 Starting View
	starting_view = f_start_view(rvtdoc)
	if starting_view:
		starting_view_name = l_safe_str(starting_view.Name)
		starting_view_category = l_safe_str(starting_view.Category.Name)
		starting_view_id = starting_view.Id.IntegerValue
	else:
		starting_view_name = ''
		starting_view_category = ''
		starting_view_id = ''
	row_01_04 = (chapter_01, 'Starting View', starting_view_category, starting_view_id,
	'Name', starting_view_name)	
	csv_list.append(row_01_04)


	##		CHAPTER 02
	chapter_02 = chapters[1]
	proj_info = DB.FilteredElementCollector(rvtdoc).OfCategory(BIC.OST_ProjectInformation).ToElements()[0]
	proj_info_params_name = ('Project Name', 'Project Number', 'Project Address', 'Project Status', 'Client Name', 'Author')
	proj_info_biparams = (BIP.PROJECT_NAME, BIP.PROJECT_NUMBER, BIP.PROJECT_ADDRESS, BIP.PROJECT_STATUS, BIP.CLIENT_NAME, BIP.PROJECT_AUTHOR)
	row_02_XX = []
	for pipn,pibp in zip(proj_info_params_name,proj_info_biparams):
		row_02_XX.append((chapter_02,	pipn,	l_safe_str(proj_info.Category.Name), proj_info.Id.IntegerValue,
		'', l_safe_str(proj_info.get_Parameter(pibp).AsString())))
	csv_list.extend(row_02_XX)


	##		CHAPTER 03
	chapter_03 = chapters[2]

	###	03_01 Project Shared Parameters
	row_03_01 = []
	ps_params_name = f_proj_shared_params(rvtdoc)[0]
	ps_params_guid = f_proj_shared_params(rvtdoc)[1]
	if ps_params_name:
		for ppn,ppg in zip(ps_params_name, ps_params_guid):
			row_03_01.append((chapter_03,'Project Shared Parameters', '', ppg,
			ppn))
	else:
		row_03_01.append((chapter_03, 'Project Shared Parameters'))
	csv_list.extend(row_03_01)

	###	03_02 Levels & Grids
	row_03_02 = []

	levels = DB.FilteredElementCollector(rvtdoc).OfCategory(BIC.OST_Levels).WhereElementIsNotElementType().ToElements()
	if len(levels)>0:
		lev_elev = [round(f_convert_units(lev.get_Parameter(BIP.LEVEL_ELEV).AsDouble(),'m'),3) for lev in levels]
		for lev, le in zip(levels, lev_elev):
			row_03_02.append((chapter_03, 'Levels & Grids', l_safe_str(lev.Category.Name), lev.Id.IntegerValue,
			'Name', l_safe_str(lev.Name), 'IsMonitoringLinkElement', lev.IsMonitoringLinkElement(), '', '',
			'Elevation', le))
	else:
		row_03_02.append((chapter_03, 'Levels & Grids'))

	grids = DB.FilteredElementCollector(rvtdoc).OfCategory(BIC.OST_Grids).WhereElementIsNotElementType().ToElements()
	if len(grids)>0:
		for g in grids:
			row_03_02.append((chapter_03, 'Levels & Grids', l_safe_str(g.Category.Name), g.Id.IntegerValue,
			'Name', l_safe_str(g.Name), 'IsMonitoringLinkElement', g.IsMonitoringLinkElement()))
	else:
		if len(row_03_02) == 0:
			row_03_02.append((chapter_03, 'Levels & Grids'))
	csv_list.extend(row_03_02)


	###	03_03 Coordinates
	row_03_03 = []
	site_location = rvtdoc.SiteLocation
	site_location_values = [f_convert_units(site_location.Elevation,'m')]
	site_location_props = ('Elevation', 'Latitude', 'Longitude')
	for slp in site_location_props[1:]:
		site_location_values.append(f_convert_units(getattr(site_location, slp),'rad'))
	for slp,slv in zip(site_location_props, site_location_values):
		row_03_03.append((chapter_03, 'Coordinates', l_safe_str(site_location.Category.Name), site_location.Id.IntegerValue,
		'', '', '', '', '', '',
		l_safe_str(slp),l_safe_str(slv)))

	project_position = rvtdoc.ActiveProjectLocation.GetProjectPosition(DB.XYZ(0,0,0))
	survey_point = DB.FilteredElementCollector(rvtdoc).OfCategory(BIC.OST_SharedBasePoint).ToElements()[0]
	project_point = DB.FilteredElementCollector(rvtdoc).OfCategory(BIC.OST_ProjectBasePoint).ToElements()[0]
	if shift:
		survey_point_location = survey_point.SharedPosition
		project_point_location = project_point.SharedPosition
		project_point_angle = abs(2*math.pi - project_position.Angle)
	else:
		survey_point_location = DB.XYZ(
			survey_point.get_Parameter(BIP.BASEPOINT_EASTWEST_PARAM).AsDouble(),
			survey_point.get_Parameter(BIP.BASEPOINT_NORTHSOUTH_PARAM).AsDouble(),
			survey_point.get_Parameter(BIP.BASEPOINT_ELEVATION_PARAM).AsDouble()
		)
		project_point_location = DB.XYZ(
			project_point.get_Parameter(BIP.BASEPOINT_EASTWEST_PARAM).AsDouble(),
			project_point.get_Parameter(BIP.BASEPOINT_NORTHSOUTH_PARAM).AsDouble(),
			project_point.get_Parameter(BIP.BASEPOINT_ELEVATION_PARAM).AsDouble()
		)
		project_point_angle = project_point.get_Parameter(BIP.BASEPOINT_ANGLETON_PARAM).AsDouble()

	survey_point_params = ('E/W', 'N/S', 'Elev')
	survey_point_values = (
		f_convert_units(survey_point_location.X, 'm'),
		f_convert_units(survey_point_location.Y, 'm'),
		f_convert_units(survey_point_location.Z, 'm')
		)
	for spp,spv in zip(survey_point_params,survey_point_values):
		row_03_03.append((chapter_03, 'Coordinates', l_safe_str(survey_point.Category.Name), survey_point.Id.IntegerValue,
		'', '', '', '', '', '',
		spp, spv))

	project_point_params = ('E/W', 'N/S', 'Elev', 'Angle to True North')
	project_point_values = (
		f_convert_units(project_point_location.X, 'm'),
		f_convert_units(project_point_location.Y, 'm'),
		f_convert_units(project_point_location.Z, 'm'),
		f_convert_units(project_point_angle, 'rad'),
	)
	for ppp,ppv in zip(project_point_params,project_point_values):
		row_03_03.append((chapter_03, 'Coordinates', l_safe_str(project_point.Category.Name), project_point.Id.IntegerValue,
		'', '', '', '', '', '',
		ppp, ppv))
	csv_list.extend(row_03_03)

	###	03_04	Design Options
	row_03_04 = []
	design_options = DB.FilteredElementCollector(rvtdoc).OfCategory(BIC.OST_DesignOptions).ToElements()
	if len(design_options)>0:
		for do in design_options:
			row_03_04.append((chapter_03, 'Design Options', l_safe_str(do.Category.Name), do.Id.IntegerValue,
			'Name', l_safe_str(do.Name)))
	else:
		row_03_04.append((chapter_03, 'Design Options'))
	csv_list.extend(row_03_04)

	###	03_05 Phases
	row_03_05 = []
	phases = DB.FilteredElementCollector(rvtdoc).OfClass(DB.Phase).ToElements()
	for phase in phases:
		row_03_05.append((chapter_03, 'Phases', l_safe_str(phase.Category.Name), phase.Id.IntegerValue,
		'Name', l_safe_str(phase.Name)))
	phase_filters = DB.FilteredElementCollector(rvtdoc).OfClass(DB.PhaseFilter).ToElements()
	for pfilter in phase_filters:
		row_03_05.append((chapter_03, 'Phase Filters', '', pfilter.Id.IntegerValue,
		'Name', l_safe_str(pfilter.Name)))
	csv_list.extend(row_03_05)

	###	03_06 Model Extents
	###	WIP


	##		CHAPTER 04
	chapter_04 = chapters[3]

	###	04_01 Categories per Workset
	row_04_01 = []
	if rvtdoc.IsWorkshared:
		user_worksets = DB.FilteredWorksetCollector(rvtdoc).OfKind(DB.WorksetKind.UserWorkset).ToWorksets()
		worksets_categories = []
		for uw in user_worksets:
			user_worksets_filter = DB.ElementWorksetFilter(uw.Id)
			composed_filter_1 = DB.LogicalAndFilter(categories_filter_1,user_worksets_filter)	
			elems_worksets = DB.FilteredElementCollector(rvtdoc).WhereElementIsNotElementType().WherePasses(composed_filter_1)
			elems_category_name = []
			for ew in elems_worksets:
				if ew.Category.Name not in elems_category_name:
					elems_category_name.append(l_safe_str(ew.Category.Name))
					ew.Dispose()
			worksets_categories.append(elems_category_name)

		for uw,w_categories in zip(user_worksets, worksets_categories):
			for wc in w_categories:
				row_04_01.append((chapter_04, 'Categories per Workset', '', uw.Id.IntegerValue,
				'Workset Name',  l_safe_str(uw.Name), 'Category Name', wc))
	else:
		row_04_01.append((chapter_04, 'Categories per Workset'))
	csv_list.extend(row_04_01)


	##		CHAPTER 05
	chapter_05 = chapters[4]

	###	05_01 Revit Links
	row_05_01 = []
	rvtlinks_type = DB.FilteredElementCollector(rvtdoc).OfClass(DB.RevitLinkType).ToElements()
	if rvtlinks_type:
		rvtlinks_instances = [rlt.GetDependentElements(DB.ElementClassFilter(DB.RevitLinkInstance)) for rlt in rvtlinks_type]
		rvtlinks_instances_len = [len(rli) for rli in rvtlinks_instances]
		rvtlinks_instances_pin = []
		for rvtlinks_inst in rvtlinks_instances:
			rlip = [rvtdoc.GetElement(rli).Pinned for rli in rvtlinks_inst]
			rvtlinks_instances_pin.append(l_bool_str(all(rlip)))
		for rlt,rlil,rlip in zip(rvtlinks_type, rvtlinks_instances_len, rvtlinks_instances_pin):
			row_05_01.append((chapter_05, 'RVT Links', l_safe_str(rlt.get_Parameter(BIP.SYMBOL_NAME_PARAM).AsString()), rlt.Id.IntegerValue,
			'', '', 'All Pinned', rlip, '', '',
			'Nr Instances', rlil))
	else:
		row_05_01.append((chapter_05, 'RVT Links', '','',
		'', '', 'AllPinned', '', '', '',
		'Nr Instances'))
	csv_list.extend(row_05_01)

	### 05_02 CAD References
	row_05_02 = []
	cadlinks_type = DB.FilteredElementCollector(rvtdoc).OfClass(DB.CADLinkType).ToElements()
	if cadlinks_type:
		cadlinks_instances = [clt.GetDependentElements(DB.ElementClassFilter(DB.ImportInstance)) for clt in cadlinks_type]
		cadlinks_instances_len = [len(cli) for cli in cadlinks_instances]
		cadlinks_instances_vs, cadlinks_instances_pin = [], []
		for cadlinks_inst in cadlinks_instances:
			cli_vs = [rvtdoc.GetElement(cli).ViewSpecific for cli in cadlinks_inst]
			clip = [rvtdoc.GetElement(cli).Pinned for cli in cadlinks_inst]
			cadlinks_instances_vs.append(l_bool_str(all(cli_vs)))
			cadlinks_instances_pin.append(l_bool_str(all(clip)))
		cadlinks_extref = [clt.IsExternalFileReference() for clt in cadlinks_type]
		for clt,clil,cle,clivs,clip in zip(cadlinks_type, cadlinks_instances_len, cadlinks_extref, cadlinks_instances_vs,cadlinks_instances_pin):
			row_05_02.append((chapter_05, 'CAD References', l_safe_str(clt.get_Parameter(BIP.SYMBOL_NAME_PARAM).AsString()), clt.Id.IntegerValue,
			'Is Linked', cle, 'All Pinned', clip,'Is View Specific', clivs,
			'Nr Instances', clil))
	else:
		row_05_02.append((chapter_05, 'CAD References', '', '',
		'Is Linked', '', 'All Pinned', '', 'Is View Specific', '',
		'Nr Instances'))
	csv_list.extend(row_05_02)


	##		CHAPTER 06
	chapter_06 = chapters[5]

	###	06_01 Warnings Classification
	row_06_01 = []
	script_folder = '\\'.join(os.path.realpath(__file__).split('\\')[:-1])
	warnings_path = os.path.join(script_folder, 'ClassifiedWarnings.csv')	
	with open(warnings_path,'rb') as csv_warnings:
		warnings_reader = csv.reader(csv_warnings, delimiter=';')
		warnings_headers = next(warnings_reader)
		warnings_data = zip(*[wr for wr in warnings_reader])

	wd_description = [wd for wd in warnings_data[0]]

	warnings = rvtdoc.GetWarnings()
	wrvt_description = []
	wrvt_score = []
	wrvt_elements = []
	wrvt_levels = []
	for wrn in warnings:
		wrvt_description.append(wrn.GetDescriptionText())
		if wrn.GetDescriptionText() in wd_description:
			wrvt_score.append(warnings_data[1][wd_description.index(wrn.GetDescriptionText())])
		else:
			wrvt_score.append('04_Unclassified')
		wrvt_elements.append(';'.join([str(x.IntegerValue) for x in wrn.GetFailingElements()]))
		if len(wrn.GetFailingElements())>0:
			wrvt_levels.append(l_safe_str(f_elem_level(rvtdoc.GetElement(wrn.GetFailingElements()[0]))))
		else:
			wrvt_levels.append('Undefined')

	if warnings:
		for wd, ws, we, wl in zip(wrvt_description, wrvt_score, wrvt_elements, wrvt_levels):
			row_06_01.append((chapter_06, 'Warnings', l_safe_str(wd), '',
			'Classification', ws, 'Elements', we, 'Level', wl))
		###	10_XX Record
		row_06_01.append((chapter_10, 'TOT Warnings', '', '',
		'', '', '', '', '', '',
		'', len(warnings)))
	else:
		row_06_01.append((chapter_06, 'Warnings'))
		row_06_01.append((chapter_10, 'TOT Warnings','','','','','','','','','',0))
	csv_list.extend(row_06_01)


	##		CHAPTER 07
	chapter_07 = chapters[6]


	###	07_01 Instances per Category
	row_07_01 = []
	elems_filtered_2 = DB.FilteredElementCollector(rvtdoc).WhereElementIsNotElementType().WherePasses(categories_filter_2).ToElements()
	ef2_cats_name = [l_safe_str(ef2.Category.Name) for ef2 in elems_filtered_2]
	ef2_elems_name = []
	for ef2 in elems_filtered_2:
		try:	ef2_elems_name.append(l_safe_str(ef2.Name))
		except:	ef2_elems_name.append('Not Found')

	###	10_XX Record
	row_07_01.append((chapter_10, 'TOT instances (x100)', '', '',
	'', '', '', '', '', '',
	'', float(len(elems_filtered_2))/100))

	for ef2,ecn, een in zip(elems_filtered_2, ef2_cats_name, ef2_elems_name):
		row_07_01.append((chapter_07, 'Instances per Category', ecn, ef2.Id.IntegerValue,
		'Names', een, 'Level', l_safe_str(f_elem_level(ef2))))
		ef2.Dispose()
	csv_list.extend(row_07_01)

	###	07_02 In-Place Families
	row_07_02 = []
	families_instance = DB.FilteredElementCollector(rvtdoc).OfClass(DB.FamilyInstance).WhereElementIsNotElementType().ToElements()
	families_inplace = [fi for fi in families_instance if fi.Symbol.Family.IsInPlace]
	if len(families_inplace)>0:
		for fip in families_inplace:
			row_07_02.append((chapter_07, 'In-Place Families', l_safe_str(fip.Symbol.Family.FamilyCategory.Name), fip.Id.IntegerValue,
			'Name', l_safe_str(fip.Symbol.Family.Name)))
		###	10_XX Record
		row_07_02.append((chapter_10, 'In-Place Families', '', '',
		'', '', '', '', '', '',
		'', len(families_inplace)))
	else:
		row_07_02.append((chapter_07,'In-Place Families'))
		row_07_02.append((chapter_10,'In-Place Families','','','','','','','','','',0))
	csv_list.extend(row_07_02)

	###	07_03 Purgeable Elements
	row_07_03 = []
	purge_guid = 'e8c63650-70b7-435a-9010-ec97660c1bda'
	purge_elems_ids = []
	performance_adviser = DB.PerformanceAdviser.GetPerformanceAdviser()
	rule_guid = System.Guid(purge_guid)
	rule_id = None
	all_rules_ids = performance_adviser.GetAllRuleIds()
	for ari in all_rules_ids:
		if str(ari.Guid) == purge_guid:
			rule_id = ari
	rule_ids = List[DB.PerformanceAdviserRuleId]([rule_id])
	fail_message = performance_adviser.ExecuteRules(rvtdoc, rule_ids)
	if fail_message.Count>0:
		purge_elems_ids.extend(fail_message[0].GetFailingElements())
	if purge_elems_ids:
		purge_elems = []
		for pei in purge_elems_ids:
			if (
				rvtdoc.GetElement(pei).Category != None
				and rvtdoc.GetElement(pei).Category.Id.IntegerValue != BIC.OST_RvtLinks.value__):
				purge_elems.append(rvtdoc.GetElement(pei))
		if len(purge_elems)>0:
			purge_elems_name = []
			for pe in purge_elems:
				try:	purge_elems_name.append(l_safe_str(pe.get_Parameter(BIP.SYMBOL_NAME_PARAM).AsString()))
				except:	purge_elems_name.append('Name Error')
			for pe,pen in zip(purge_elems, purge_elems_name):
				row_07_03.append((chapter_07, 'Purgeable Elements', l_safe_str(pe.Category.Name), pe.Id.IntegerValue,
				'Name', pen))
				pe.Dispose()
			### 10_XX Record
			row_07_03.append((chapter_10, 'TOT Purgeable Elements','', '',
			'', '', '', '', '', '',
			'', len(purge_elems)))
		else:
			row_07_03.append((chapter_07, 'Purgeable Elements'))
			row_07_03.append((chapter_10, 'TOT Purgeable Elements','','','','','','','','','',0))
	csv_list.extend(row_07_03)


	##		CHAPTER 08
	chapter_08 = chapters[7]

	###	08_01 Areas
	row_08_01 = []
	areas = DB.FilteredElementCollector(rvtdoc).OfCategory(BIC.OST_Areas).WhereElementIsNotElementType().ToElements()
	if len(areas)>0:
		for a in areas:
			row_08_01.append((chapter_08, 'Areas', l_safe_str(a.Category.Name), a.Id.IntegerValue,
			'Property', f_spatialelems_classification(a), 'Level', l_safe_str(f_elem_level(a))))
			a.Dispose()
	else:
		row_08_01.append((chapter_08,'Areas'))
	csv_list.extend(row_08_01)

	###	08_02 Rooms
	row_08_02 = []
	rooms = DB.FilteredElementCollector(rvtdoc).OfCategory(BIC.OST_Rooms).WhereElementIsNotElementType().ToElements()
	if len(rooms)>0:
		for r in rooms:
			row_08_02.append((chapter_08, 'Rooms', l_safe_str(r.Category.Name), r.Id.IntegerValue,
			'Property', f_spatialelems_classification(r), 'Level', l_safe_str(f_elem_level(r))))
			r.Dispose()
	else:
		row_08_02.append((chapter_08,'Rooms'))
	csv_list.extend(row_08_02)

	###	08_03 Spaces
	row_08_03 = []
	spaces = DB.FilteredElementCollector(rvtdoc).OfCategory(BIC.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
	if len(spaces)>0:
		for s in spaces:
			row_08_03.append((chapter_08, 'Spaces', l_safe_str(s.Category.Name), s.Id.IntegerValue,
			'Property', f_spatialelems_classification(s), 'Level', l_safe_str(f_elem_level(s))))
			s.Dispose()
	else:
		row_08_03.append((chapter_08,'Spaces'))
	csv_list.extend(row_08_03)


	##		CHAPTER 09
	chapter_09 = chapters[8]

	###	09_01 Views on Sheets
	row_09_01 = []
	views_all = DB.FilteredElementCollector(rvtdoc).OfCategory(BIC.OST_Views).WhereElementIsNotElementType().ToElements()
	views = [va for va in views_all if not va.IsTemplate]
	sheets = DB.FilteredElementCollector(rvtdoc).OfCategory(BIC.OST_Sheets).WhereElementIsNotElementType().ToElements()
	###	10_XX Record
	if len(sheets)>0:
		row_09_01.append((chapter_10, 'TOT Sheets', '', '', '', '', '', '', '', '', '', len(sheets)))
	else:
		row_09_01.append((chapter_10, 'TOT Sheets', '', '', '', '', '', '', '', '', '', 0))	
	
	views_sheets_ids = []
	if sheets:
		for sht in sheets:
			views_sheets_ids.extend(sht.GetAllPlacedViews())
			sht.Dispose()
	views_onsheets = [v for v in views if v.Id in views_sheets_ids]
	if views_onsheets:
		for vs in views_onsheets:
			row_09_01.append((chapter_09, 'Views on Sheets', l_safe_str(vs.Category.Name), vs.Id.IntegerValue,
			'Name', l_safe_str(vs.Name), 'Has View Template', l_bool_str(vs.ViewTemplateId.IntegerValue>0)))
	else:
		row_09_01.append((chapter_09, 'Views on Sheets'))
	csv_list.extend(row_09_01)

	###	09_02 Schedules on Sheets
	row_09_02 = []
	schedules_all = DB.FilteredElementCollector(rvtdoc).OfCategory(BIC.OST_Schedules).WhereElementIsNotElementType().ToElements()
	schedules = [sa for sa in schedules_all if not sa.IsTemplate]
	schedules_sheets = DB.FilteredElementCollector(rvtdoc).OfClass(DB.ScheduleSheetInstance).ToElements()
	schedules_sheets_ids = [ss.ScheduleId for ss in schedules_sheets]
	schedules_onsheet = [s for s in schedules if s.Id in schedules_sheets_ids]
	if schedules_onsheet:
		for ss in schedules_onsheet:
			row_09_02.append((chapter_09, 'Schedules on Sheets', l_safe_str(ss.Category.Name), ss.Id.IntegerValue,
			'Name', l_safe_str(ss.Name), 'Has View Template', l_bool_str(ss.ViewTemplateId.IntegerValue>0)))
	else:
		row_09_02.append((chapter_09, 'Schedules on Sheets'))
	csv_list.extend(row_09_02)

	###	09_03 View Templates
	row_09_03 = []
	views_template_ids = [va.Id.IntegerValue for va in views_all if va.IsTemplate]
	schedule_template_ids = [sa.Id.IntegerValue for sa in schedules_all if sa.IsTemplate]
	templates_ids_all = views_template_ids + schedule_template_ids
	templates_all = [rvtdoc.GetElement(DB.ElementId(tia)) for tia in templates_ids_all]
	views_schedules = views + schedules
	templates_ids_used = set()
	for vs in views_schedules:
		if vs.ViewTemplateId.IntegerValue>0:
			templates_ids_used.add(vs.ViewTemplateId.IntegerValue)
	if templates_all:
		for ta, tia in zip(templates_all, templates_ids_all):
			row_09_03.append((chapter_09, 'View Templates', l_safe_str(ta.Category.Name), tia,
			'Name', l_safe_str(ta.Name), 'Used', l_bool_str(tia in templates_ids_used)))
	else:
		row_09_03.append((chapter_09, 'View Templates'))
	###	10_XX Record
	if len(views_schedules)>0:
		row_09_03.append((chapter_10, 'TOT Views', '', '', '', '', '', '', '', '', '', len(views_schedules)))
	else:
		row_09_03.append((chapter_10, 'TOT Views', '', '', '', '', '', '', '', '', '', 0))		
	csv_list.extend(row_09_03)

	###	09_04 Filters
	row_09_04 = []
	filters = DB.FilteredElementCollector(rvtdoc).OfClass(DB.ParameterFilterElement).ToElements()
	filters_ids_all = [f.Id.IntegerValue for f in filters]
	filters_ids_used = set()
	for v in views:
		if v.AreGraphicsOverridesAllowed():
			view_filters = v.GetFilters()
			for vf in view_filters:
				filters_ids_used.add(vf.IntegerValue)
	if filters:
		for f, fia in zip(filters, filters_ids_all):
			row_09_04.append((chapter_09, 'Filters', '', fia, 'Name', l_safe_str(f.Name), 'Used', l_bool_str(fia in filters_ids_used)))
			f.Dispose()
	else:
		row_09_04.append((chapter_09, 'Filters'))
	csv_list.extend(row_09_04)

	for vs in views_schedules:
		vs.Dispose()

	###	09_05 Orphaned Tags
	row_09_05 = []
	tags = DB.FilteredElementCollector(rvtdoc).OfClass(DB.IndependentTag).WhereElementIsNotElementType().ToElements()
	tags_orphaned = [t for t in tags if t.IsOrphaned]
	if tags_orphaned:
		for tag in tags_orphaned:
			row_09_05.append((chapter_09, 'Orphaned Tags', l_safe_str(tag.Category.Name), tag.Id.IntegerValue,
			'Name', l_safe_str(tag.Name), 'Owner View', l_safe_str(rvtdoc.GetElement(tag.OwnerViewId).Name)))
			tag.Dispose()
	else:
		row_09_05.append((chapter_09, 'Tags'))
	csv_list.extend(row_09_05)

	with open(csv_path, 'wb') as csv_file:
		csv_writer = csv.writer(csv_file, delimiter=';')
		for cl in csv_list:
			csv_writer.writerow(cl)
else:
	script.exit()
