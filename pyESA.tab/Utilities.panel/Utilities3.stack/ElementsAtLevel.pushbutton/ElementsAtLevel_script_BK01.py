__author__ = 'Antonio Miano'

#REFERENCES
from pyrevit import revit, script, DB, UI
from pyrevit import PyRevitException, PyRevitIOError
from pyrevit import forms

from rpw.ui.forms import SelectFromList, CommandLink, TaskDialog

from Autodesk.Revit.UI.Selection import Selection

doc = revit.doc
uidoc = revit.uidoc
BIC = DB.BuiltInCategory
BIP = DB.BuiltInParameter

#DEFINITIONS
def f_change_level(bip_level, bip_offset, future_level):
	try:
		##Get
		current_level = item.Parameter[bip_level].AsElementId()
		current_offset = item.Parameter[bip_offset].AsDouble()
		current_elevation = doc.GetElement(current_level).Elevation + current_offset
		future_offset = current_elevation - lev.Elevation
		##Set
		item.Parameter[bip_level].Set(future_level.Id)
		item.Parameter[bip_offset].Set(future_offset)
	except:
		pass

#INPUTS
##Ask for Level
lev = forms.select_levels(title='Select Levels', button_name='Select', width=500, multiple=False, filterfunc=None, doc=None, use_selection=False)
if not lev:	script.exit()

##Ask for Category
family_level_param = [
	BIC.OST_DuctTerminal,
	BIC.OST_DuctFitting,
	BIC.OST_CommunicationDevices,
	BIC.OST_DataDevices,
	BIC.OST_DuctAccessory,
	BIC.OST_ElectricalEquipment,
	BIC.OST_ElectricalFixtures,
	BIC.OST_FireAlarmDevices,
	BIC.OST_LightingDevices,
	BIC.OST_LightingFixtures,
	BIC.OST_MechanicalEquipment,
	BIC.OST_NurseCallDevices,
	BIC.OST_PipeAccessory,
	BIC.OST_PipeFitting,
	BIC.OST_PlumbingFixtures,
	BIC.OST_SecurityDevices,
	BIC.OST_Sprinklers,
	BIC.OST_TelephoneDevices,
	BIC.OST_Furniture,
	BIC.OST_Casework
]

rbs_start_level_param = [
	BIC.OST_DuctCurves,
	BIC.OST_FlexDuctCurves,
	BIC.OST_PipeCurves,
	BIC.OST_FlexPipeCurves,
	BIC.OST_CableTray,
	BIC.OST_Conduit
]

arc_str_bics = [
	BIC.OST_Walls,
	BIC.OST_Floors,
	BIC.OST_StructuralColumns
]

allowed_bics = []
allowed_bics.extend(family_level_param)
allowed_bics.extend(rbs_start_level_param)
allowed_bics.extend(arc_str_bics)

allowed_bics_name  = [revit.query.get_category(item).Name for item in allowed_bics]
allowed_bics_dict = dict(zip(allowed_bics_name, allowed_bics))

cat = SelectFromList('Select Category', allowed_bics_dict)
cat_name = revit.query.get_category(cat).Name

#CODE
items = revit.pick_elements_by_category(cat, "Select {} to edit".format(cat_name))
if not items:	script.exit()

with revit.Transaction("ElementsAtLevel"):

	for item in items:

		if cat in family_level_param:
			f_change_level(BIP.FAMILY_LEVEL_PARAM, BIP.INSTANCE_ELEVATION_PARAM, lev)

		elif cat in rbs_start_level_param:
			f_change_level(BIP.RBS_START_LEVEL_PARAM, BIP.RBS_CTC_TOP_ELEVATION, lev)
		
		elif cat == BIC.OST_Walls:
			f_change_level(BIP.WALL_BASE_CONSTRAINT, BIP.WALL_BASE_OFFSET, lev)		

		elif cat == BIC.OST_Floors:
			f_change_level(BIP.LEVEL_PARAM, BIP.FLOOR_HEIGHTABOVELEVEL_PARAM, lev)	

		elif cat == BIC.OST_StructuralColumns:
			f_change_level(BIP.FAMILY_BASE_LEVEL_PARAM, BIP.FAMILY_BASE_LEVEL_OFFSET_PARAM, lev)	

		else:
			script.exit()
