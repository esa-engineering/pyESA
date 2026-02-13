#REFERENCES
import pyrevit
from pyrevit import revit, DB
from rpw.ui.forms import TaskDialog

doc = __revit__.ActiveUIDocument.Document

#Get Units Format in the current document
FormatOptionsLength = doc.GetUnits().GetFormatOptions(DB.UnitType.UT_Length)
folDispUnits = FormatOptionsLength.DisplayUnits
folUnitSym = FormatOptionsLength.UnitSymbol
lengthUnitNames = DB.LabelUtils.GetLabelFor(folDispUnits)

FormatOptionsArea = doc.GetUnits().GetFormatOptions(DB.UnitType.UT_Area)
foaDispUnits = FormatOptionsArea.DisplayUnits
foaUnitSym = FormatOptionsArea.UnitSymbol
areaUnitNames = DB.LabelUtils.GetLabelFor(foaDispUnits)

FormatOptionsVolume = doc.GetUnits().GetFormatOptions(DB.UnitType.UT_Volume)
fovDispUnits = FormatOptionsVolume.DisplayUnits
fovUnitSym = FormatOptionsVolume.UnitSymbol
volumeUnitNames = DB.LabelUtils.GetLabelFor(fovDispUnits)

#Get selected items
selection = revit.get_selection()

#Calculate totals
itmLengths,itmAreas,itmVolumes = [],[],[]

##Total lenghts
nl = 0
for item in selection:
	try:
		p_length = item.LookupParameter('Length')
		if p_length != None:
			length = p_length.AsDouble()
			nl += 1
		else:
			p_length = item.LookupParameter('Total Length')
			if p_length != None:
				length = p_length.AsDouble()
				nl += 1
			else:
				length = 0.0
		itmLengths.append(length)
		totLenghts = sum(itmLengths)
		if totLenghts > 0.0:
			msgLength = str(DB.UnitUtils.ConvertFromInternalUnits(totLenghts,folDispUnits)) + " " + lengthUnitNames
		else:
			msgLength = 'Selected objects do not have any length!'
	except:
		msgLength = 'Selected objects do not have any length!'

##Total areas
na = 0
for item in selection:
	try:
		p_area = item.LookupParameter('Area')
		if p_area != None:
			area = p_area.AsDouble()
			na += 1
		else:
			p_area = item.LookupParameter('Surface Area')
			if p_area != None:
				area = p_area.AsDouble()
				na += 1
			else:
				area = 0.0
		itmAreas.append(area)
		totAreas = sum(itmAreas)
		if totAreas > 0.0:
			msgAreas = str(DB.UnitUtils.ConvertFromInternalUnits(totAreas,foaDispUnits)) + " " + areaUnitNames
		else:
			msgAreas = 'Selected objects do not have any area!'
	except:
		msgAreas = 'Selected objects do not have any area!'

##Total Volume
nv = 0
for item in selection:
	try:
		p_volume = item.LookupParameter('Volume')
		if p_volume != None:
			volume = p_volume.AsDouble()
			nv += 1
		itmVolumes.append(volume)
		totVolumes = sum(itmVolumes)
		if totVolumes > 0.0:
			msgVolumes = str(DB.UnitUtils.ConvertFromInternalUnits(totVolumes,fovDispUnits)) + " " + volumeUnitNames
		else:
			msgVolumes = 'Selected objects do not have any volume!'
	except:
		msgVolumes = 'Selected objects do not have any volume!'

#Create output message
msg =	'[{}] Total Lenghts: {}\n\
		[{}] Total Areas: {}\n\
		[{}] Total Volumes: {}\n'\
		.format(str(nl),msgLength,str(na),msgAreas,str(nv),msgVolumes)
#msg = '(' + str(nl) + ') Total Lenghts: '+ msgLength + '\n\n' + '(' + str(na) + ') Total Areas: ' + msgAreas + '\n\n' +  '(' + str(nv) + ') Total Volumes: ' + msgVolumes
dialog = TaskDialog('Measure Dimensions', content = msg, buttons = ['OK'], footer = '', show_close = True)
dialog.show(exit = True)