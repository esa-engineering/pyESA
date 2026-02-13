__title__ = "Elements\nat Level"

__doc__ ="""Edit the host level of all the picked elements of
a selected category by keeping the same position."""

__author__ = "bimdifferent"

#REFERENCES
from pyrevit import revit, script, DB, UI
from pyrevit import PyRevitException, PyRevitIOError
from pyrevit import forms

from rpw.ui.forms import TaskDialog, CheckBox, FlexForm, Label, TextBox, Separator, Button, ComboBox

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
		future_offset = current_elevation - future_level.Elevation
		##Set
		item.Parameter[bip_level].Set(future_level.Id)
		item.Parameter[bip_offset].Set(future_offset)
	except:	pass

#INPUTS
##Get Levels
levels = revit.query.get_elements_by_categories([DB.BuiltInCategory.OST_Levels])
levels_name = [lev.Name for lev in levels]
levels_dict = dict(zip(levels_name, levels))

##Get Categories
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
	BIC.OST_Casework,
	BIC.OST_Doors,
	BIC.OST_Windows
]

rbs_start_level_param = [
	BIC.OST_DuctCurves,
	BIC.OST_FlexDuctCurves,
	BIC.OST_PipeCurves,
	BIC.OST_FlexPipeCurves,
	BIC.OST_CableTray,
	BIC.OST_Conduit
]

base_top_bics = [
	BIC.OST_Walls,
	BIC.OST_Floors,
	BIC.OST_StructuralColumns,
	BIC.OST_Columns,
	BIC.OST_Ceilings,
	BIC.OST_Roofs
]

allowed_bics = []
allowed_bics.extend(family_level_param)
allowed_bics.extend(rbs_start_level_param)
allowed_bics.extend(base_top_bics)

allowed_bics_name  = [revit.query.get_category(item).Name for item in allowed_bics]
allowed_bics_dict = dict(zip(allowed_bics_name, allowed_bics))

##Create form
components = [
Label('Select Level'),
ComboBox('cmb_levs', levels_dict),
Label('Select Category'),
ComboBox('cmb_cats', allowed_bics_dict),
Separator(),
Label('Select Option (Only for Walls and Columns)'),
CheckBox('ckb_base', 'Base'),
CheckBox('ckb_top', 'Top'),
Separator(),
Button('OK')
]
flex_form = FlexForm('Elements at Level', components)
flex_form.show()
if not flex_form.values.items(): script.exit()

##Get Selected Category
cat = flex_form.values['cmb_cats']
cat_name = revit.query.get_category(cat).Name

##Get Selected level
lev = flex_form.values['cmb_levs']

#CODE
items = revit.pick_elements_by_category(cat, "Select {} to edit".format(cat_name))
if not items:	script.exit()

with revit.Transaction("ElementsAtLevel"):

	for item in items:

		if cat in family_level_param:
			f_change_level(BIP.FAMILY_LEVEL_PARAM, BIP.INSTANCE_ELEVATION_PARAM, lev)

		elif cat in rbs_start_level_param:
			f_change_level(BIP.RBS_START_LEVEL_PARAM, BIP.RBS_OFFSET_PARAM, lev)
		
		elif cat == BIC.OST_Walls:
			if item.Parameter[BIP.WALL_BASE_CONSTRAINT] == None:
				f_change_level(BIP.FACEROOF_LEVEL_PARAM, BIP.FACEROOF_OFFSET_PARAM, lev)
			else:	
				if flex_form.values['ckb_base']:
					f_change_level(BIP.WALL_BASE_CONSTRAINT, BIP.WALL_BASE_OFFSET, lev)
				else:	pass
				if flex_form.values['ckb_top']:
					if doc.GetElement(item.Parameter[BIP.WALL_HEIGHT_TYPE].AsElementId()) is not None:
						f_change_level(BIP.WALL_HEIGHT_TYPE, BIP.WALL_TOP_OFFSET, lev)
					else:
						try:
							##Get
							current_base_level = item.Parameter[BIP.WALL_BASE_CONSTRAINT].AsElementId()
							current_base_offset = item.Parameter[BIP.WALL_BASE_OFFSET].AsDouble()
							current_height = item.Parameter[BIP.WALL_USER_HEIGHT_PARAM].AsDouble()
							current_elevation = doc.GetElement(current_base_level).Elevation + current_base_offset + current_height
							future_offset = current_elevation - lev.Elevation
							##Set
							item.Parameter[BIP.WALL_HEIGHT_TYPE].Set(lev.Id)
							item.Parameter[BIP.WALL_TOP_OFFSET].Set(future_offset)
						except:	pass
				else:	pass

		elif cat == BIC.OST_Roofs:
			if item.Parameter[BIP.ROOF_BASE_LEVEL_PARAM] != None:
				f_change_level(BIP.ROOF_BASE_LEVEL_PARAM, BIP.ROOF_LEVEL_OFFSET_PARAM, lev)
			if item.Parameter[BIP.ROOF_CONSTRAINT_LEVEL_PARAM] != None:
				f_change_level(BIP.ROOF_CONSTRAINT_LEVEL_PARAM, BIP.ROOF_CONSTRAINT_OFFSET_PARAM, lev)
			if item.Parameter[BIP.FACEROOF_LEVEL_PARAM] != None:
				f_change_level(BIP.FACEROOF_LEVEL_PARAM, BIP.FACEROOF_OFFSET_PARAM, lev)

		elif cat == BIC.OST_Floors:
			f_change_level(BIP.LEVEL_PARAM, BIP.FLOOR_HEIGHTABOVELEVEL_PARAM, lev)

		elif cat == BIC.OST_Ceilings:
			f_change_level(BIP.LEVEL_PARAM, BIP.CEILING_HEIGHTABOVELEVEL_PARAM, lev)	

		elif cat == BIC.OST_StructuralColumns or cat == BIC.OST_Columns:
			if flex_form.values['ckb_base']:
				f_change_level(BIP.FAMILY_BASE_LEVEL_PARAM, BIP.FAMILY_BASE_LEVEL_OFFSET_PARAM, lev)
			else:	pass
			if flex_form.values['ckb_top']:
					f_change_level(BIP.FAMILY_TOP_LEVEL_PARAM, BIP.FAMILY_TOP_LEVEL_OFFSET_PARAM, lev)
			else:	pass

		else:
			script.exit()

