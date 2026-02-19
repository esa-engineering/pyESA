#REFERENCES
import pyrevit
from pyrevit import revit, DB, UI, script
from pyrevit import forms

output = script.get_output()
doc = revit.doc

#DEFINITIONS
def f_value_from_parameter(par):
	value = 0.0
	if par != None:
		if 'Length' in par.Definition.Name:
			par_length = item.LookupParameter('Length')
			if par_length != None:
				value = par_length.AsDouble()
			else:
				par_length = item.LookupParameter('Total Length')
				if par_length != None:
					value = par_length.AsDouble()
		elif 'Area' in par.Definition.Name:
			par_area = item.LookupParameter('Area')
			if par_area != None:
				value = par_area.AsDouble()

		elif 'Volume' in par.Definition.Name:
			par_volume = item.LookupParameter('Volume')
			if par_volume != None:
				value = par_volume.AsDouble()
	return value

def f_convert_from_internal_units(doc, unit_type, value):
	if int(doc.Application.VersionNumber) < 2022:
		if unit_type == 'mm':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.DisplayUnitType.DUT_METERS)
			return value_conv * 1000
		elif unit_type == 'cm':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.DisplayUnitType.DUT_METERS)
			return value_conv * 100
		elif unit_type == 'm':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.DisplayUnitType.DUT_METERS)
			return value_conv
		elif unit_type == 'mm2':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.DisplayUnitType.DUT_SQUARE_METERS)
			return value_conv * 1000000
		elif unit_type == 'cm2':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.DisplayUnitType.DUT_SQUARE_METERS)
			return value_conv * 10000
		elif unit_type == 'm2':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.DisplayUnitType.DUT_SQUARE_METERS)
			return value_conv
		elif unit_type == 'mm3':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.DisplayUnitType.DUT_CUBIC_METERS)
			return value_conv * 1000000000
		elif unit_type == 'cm3':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.DisplayUnitType.DUT_CUBIC_METERS)
			return value_conv * 1000000
		elif unit_type == 'm3':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.DisplayUnitType.DUT_CUBIC_METERS)
			return value_conv	
		else:
			return 'select correct unit type'
	else:
		if unit_type == 'mm':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.UnitTypeId.Meters)
			return value_conv * 1000
		elif unit_type == 'cm':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.UnitTypeId.Meters)
			return value_conv * 100
		elif unit_type == 'm':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.UnitTypeId.Meters)
			return value_conv
		elif unit_type == 'mm2':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.UnitTypeId.SquareMeters)
			return value_conv * 1000000
		elif unit_type == 'cm2':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.UnitTypeId.SquareMeters)
			return value_conv * 10000
		elif unit_type == 'm2':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.UnitTypeId.SquareMeters)
			return value_conv
		elif unit_type == 'mm3':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.UnitTypeId.CubicMeters)
			return value_conv * 1000000000
		elif unit_type == 'cm3':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.UnitTypeId.CubicMeters)
			return value_conv * 1000000
		elif unit_type == 'm3':
			value_conv = DB.UnitUtils.ConvertFromInternalUnits(value, DB.UnitTypeId.CubicMeters)
			return value_conv	
		else:
			return 'select correct unit type'

#CODE
##Get selected items
selection = revit.get_selection()

##Calculate totals
itmLengths,itmAreas,itmVolumes = [],[],[]
msgLength = 'Selected objects do not have any length!'
msgAreas = 'Selected objects do not have any area!'
msgVolumes = 'Selected objects do not have any volume!'

###Total lenghts
nl = 0
for item in selection:
	try:
		length = f_value_from_parameter(item.LookupParameter('Length'))
		if length > 0:
			nl += 1
		itmLengths.append(length)
		totLenghts = sum(itmLengths)
		if totLenghts > 0.0:
			msgLength_mm = "{:e}".format(f_convert_from_internal_units(doc, 'mm', totLenghts))
			msgLength_cm = "{:e}".format(f_convert_from_internal_units(doc, 'cm', totLenghts))
			msgLength_m = str(f_convert_from_internal_units(doc, 'm', totLenghts))
		else:
			msgLength_mm = msgLength
			msgLength_cm = msgLength 
			msgLength_m = msgLength
	except Exception as e:
		msgLength_mm = str(e)
		msgLength_cm = str(e)
		msgLength_m = str(e)

###Total areas
na = 0
for item in selection:
	try:
		area = f_value_from_parameter(item.LookupParameter('Area'))
		if area > 0:
			na += 1
		itmAreas.append(area)
		totAreas = sum(itmAreas)
		if totAreas > 0.0:
			msgAreas_mm = "{:e}".format(f_convert_from_internal_units(doc, 'mm2', totAreas))
			msgAreas_cm = "{:e}".format(f_convert_from_internal_units(doc, 'cm2', totAreas))
			msgAreas_m = str(f_convert_from_internal_units(doc, 'm2', totAreas))
		else:
			msgAreas_mm = msgAreas
			msgAreas_cm = msgAreas
			msgAreas_m = msgAreas
	except Exception as e:
		msgAreas_mm = str(e)
		msgAreas_cm = str(e)
		msgAreas_m = str(e)

###Total volumes
nv = 0
for item in selection:
	try:
		volume = f_value_from_parameter(item.LookupParameter('Volume'))
		if volume > 0:
			nv += 1
		itmVolumes.append(volume)
		totVolumes = sum(itmVolumes)
		if totVolumes > 0.0:
			msgVolumes_mm = "{:e}".format(f_convert_from_internal_units(doc, 'mm3', totVolumes))
			msgVolumes_cm = "{:e}".format(f_convert_from_internal_units(doc, 'cm3', totVolumes))
			msgVolumes_m = str(f_convert_from_internal_units(doc, 'm3', totVolumes))
		else:
			msgVolumes_mm = msgVolumes
			msgVolumes_cm = msgVolumes
			msgVolumes_m = msgVolumes
	except Exception as e:
		msgVolumes_mm = str(e)
		msgVolumes_cm = str(e)
		msgVolumes_m = str(e)

##Print output
if msgLength_mm != msgLength:
	tab_lengths_headers = ['Total elements', '[mm]', '[cm]', '[m]']
	tab_lengths_body = [[str(nl), msgLength_mm, msgLength_cm, msgLength_m]]

	output.print_table(
		table_data = tab_lengths_body,
		title = 'LENGHTS',
		columns = tab_lengths_headers
	)

if msgAreas_mm != msgAreas:
	tab_areas_headers = ['Total elements', '[mm2]', '[cm2]', '[m2]']
	tab_areas_body = [[str(na), msgAreas_mm, msgAreas_cm, msgAreas_m]]

	output.print_table(
		table_data = tab_areas_body,
		title = 'AREAS',
		columns = tab_areas_headers
	)

if msgVolumes_mm != msgVolumes:
	tab_volumes_headers = ['Total elements', '[mm3]', '[cm3]', '[m3]']
	tab_volumes_body = [[str(nv), msgVolumes_mm, msgVolumes_cm, msgVolumes_m]]

	output.print_table(
		table_data = tab_volumes_body,
		title = 'VOLUMES',
		columns = tab_volumes_headers
	)
