# -*- coding: utf-8 -*-
__title__ = 'Cable Trays\nConnect 30°/45°/60°'
__doc__ =	"""Version = 1.1
Date = 08.05.2025
________________________________________________________________
Duct slope: Default 90 degrees\n
Instruction:\n
- Select in order the ducts to connect.\n
- Press ESC to end.
\n
Differently from the Trim/Extend, it can also connect
ductes at different elevation and aligned each other.\n
When the selected ducts are not aligned, the tool will
create a new duct so to connect the two closest extremity of the ducts.
_______________________________________________________________
Author:\n 
ESA Engineering Srl
"""

import math

import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct
from Autodesk.Revit.DB.Electrical import CableTray
from pyrevit import revit, DB
from pyrevit import PyRevitException, PyRevitIOError
from pyrevit import forms

doc = revit.doc

### FUNCTIONS

def copyTrayDimensions(sourceTray, targetTray):
	"""Copia le dimensioni (larghezza, altezza) dalla tray sorgente a quella destinazione"""
	try:
		# Parametri di dimensione per le cable tray
		width_param = sourceTray.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
		height_param = sourceTray.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
		
		if width_param and not width_param.IsReadOnly:
			target_width = targetTray.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
			if target_width and not target_width.IsReadOnly:
				target_width.Set(width_param.AsDouble())
		
		if height_param and not height_param.IsReadOnly:
			target_height = targetTray.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
			if target_height and not target_height.IsReadOnly:
				target_height.Set(height_param.AsDouble())
				
		# Copia anche altri parametri se necessario (spessore, materiale, ecc.)
		thickness_param = sourceTray.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_THICKNESS_PARAM)
		if thickness_param and not thickness_param.IsReadOnly:
			target_thickness = targetTray.get_Parameter(DB.BuiltInParameter.RBS_CABLETRAY_THICKNESS_PARAM)
			if target_thickness and not target_thickness.IsReadOnly:
				target_thickness.Set(thickness_param.AsDouble())
				
	except Exception as e:
		print("Errore nella copia delle dimensioni: {}".format(str(e)))

def placeElbow(p1, p2):
	try:
		temp = []
		for c1 in list(p1.ConnectorManager.Connectors):
			for c2 in list(p2.ConnectorManager.Connectors):
				temp.append( [c1.Origin.DistanceTo(c2.Origin), c1, c2] )
		temp = sorted(temp)
		e = doc.Create.NewElbowFitting(temp[0][1], temp[0][2])
		return	e
	except:
		return	None

def	connectTraysWithTray(doc, p1, p2):
	unused1 = list(p1.ConnectorManager.UnusedConnectors)
	unused2 = list(p2.ConnectorManager.UnusedConnectors)
	temp = []	
	for u1 in unused1:
		for u2 in unused2:
			temp.append( [u1.Origin.DistanceTo(u2.Origin), u1, u2] )
	
	conn01 = sorted(temp)[0][1]
	conn02 = sorted(temp)[0][2]

	startpoint = conn01.Origin
	endpoint = conn02.Origin
	
	levId = p1.Parameter[DB.BuiltInParameter.RBS_START_LEVEL_PARAM].AsElementId()
	newTray = CableTray.Create(doc, p1.GetTypeId(), startpoint, endpoint, levId)
	
	# Copia le dimensioni dalla tray originale
	copyTrayDimensions(p1, newTray)
	
	return	newTray

def createTray(doc, p1, pt1, pt2):
	unused1 = list(p1.ConnectorManager.UnusedConnectors)
	dists = sorted([[u.Origin.DistanceTo(pt1), u] for u in unused1])
	conn01 = dists[0][1]
	pconn01 = conn01.Origin

	levId = p1.Parameter[DB.BuiltInParameter.RBS_START_LEVEL_PARAM].AsElementId()
	newTray = CableTray.Create(doc, p1.GetTypeId(), pconn01, pt2, levId)
	
	# Copia le dimensioni dalla tray originale
	copyTrayDimensions(p1, newTray)
	
	return	newTray

## LINE CONNECTING DUCTS
def get_connectingLine(ln1, ln2):
	pts1 = (ln1.GetEndPoint(0), ln1.GetEndPoint(1))
	pts2 = (ln2.GetEndPoint(0), ln2.GetEndPoint(1))
	dists = []
	for p1 in pts1:
		for p2 in pts2:	dists.append([p1.DistanceTo(p2), p1, p2])
	pts = sorted(dists)[0][1:]
	pt1 = pts[0]
	pt2 = DB.XYZ(pts[1].X, pts[1].Y, pts[0].Z)
	return	DB.Line.CreateBound(pt1, pt2)

## LINE INTERSECTIONS
projAtZero = lambda p: DB.XYZ(p.X, p.Y, 0)

def NotTouch(crv1, crv2):
	clone01 = crv1.Clone()
	clone01 = DB.Line.CreateBound(projAtZero(clone01.GetEndPoint(0)), projAtZero(clone01.GetEndPoint(1)))
	clone01.MakeUnbound()
	clone02 = crv2.Clone()
	clone02 = DB.Line.CreateBound(projAtZero(clone02.GetEndPoint(0)), projAtZero(clone02.GetEndPoint(1)))
	#clone02.MakeUnbound()
	if clone01.Intersect(clone02) == DB.SetComparisonResult.Disjoint:
		return	True
	else:
		return False

def testFloat(s):
	try:	
		return	float(s) 
	except:
		return	None

# INPUTS
slopeDef = forms.CommandSwitchWindow.show([30,45,60,90], message='Cable Tray connector slope [degrees]:')
if slopeDef:	slopeDef = 90-testFloat(slopeDef)
else:	slopeDef = 0

BIC = DB.BuiltInCategory
nr = 0

with forms.ProgressBar(title='select Cable Trays to connect - press Esc to stop', cancellable=True) as pb:

	selected_trays = []
	prog = 0
	pb.update_progress(prog, 100)
	for tray in revit.get_picked_elements_by_category(BIC.OST_CableTray, "Select Cable Tray element"):
		prog += 50
		pb.update_progress(prog, 100)
		selected_trays.append(tray)
		if tray == None:
			break
		elif prog == 100:
			prog = 0
			tray1 = selected_trays[0]
			tray2 = selected_trays[1]
			selected_trays = []

			ln1 = tray1.Location.Curve
			ln2 = tray2.Location.Curve

			baseNewLine = get_connectingLine(ln1, ln2)

			# DEFINE VALUE FOR DUCT INCLINATION
			deltaH = math.fabs(ln2.GetEndPoint(0).Z - baseNewLine.GetEndPoint(1).Z)
			Co = math.tan(math.radians(slopeDef)) * deltaH
			oppDir = baseNewLine.Direction.Multiply(Co)

			nr += 1
			with revit.Transaction('M4B - Connessione {}'.format(nr)):
				# CHECK ALIGNEMENT
				if ln1.Direction.IsAlmostEqualTo(baseNewLine.Direction)\
					and ln1.Direction.IsAlmostEqualTo(ln2.Direction)\
					and deltaH==0:
					baseNewLine	= DB.Line.CreateBound(ln1.GetEndPoint(0), ln2.GetEndPoint(1))
					tray1.Location.Curve = baseNewLine
					doc.Delete(tray2.Id)

				elif ln1.Direction.IsAlmostEqualTo(baseNewLine.Direction):
					if baseNewLine.GetEndPoint(0).IsAlmostEqualTo(ln1.GetEndPoint(0)):
						baseNewLine = DB.Line.CreateBound(ln1.GetEndPoint(1),
									baseNewLine.GetEndPoint(1).Subtract(oppDir))
					else:	
						baseNewLine = DB.Line.CreateBound(ln1.GetEndPoint(0),
									baseNewLine.GetEndPoint(1).Subtract(oppDir))
					tray1.Location.Curve = baseNewLine
					tray3 = connectTraysWithTray(doc, tray1, tray2)
					# CONNECT
					placeElbow(tray1, tray3)
					placeElbow(tray3, tray2)

				else:
					# CREATE NEW DUCT
					tray3 = createTray(doc, tray1, 
							baseNewLine.GetEndPoint(0),
							baseNewLine.GetEndPoint(1).Subtract(oppDir))
					
					# CONNECT THEM ALL
					placeElbow(tray1, tray3)
					
					try:
						tray4 = connectTraysWithTray(doc, tray3, tray2)
						placeElbow(tray3, tray4)
						placeElbow(tray4, tray2)
					except:
						# in case duct3 is at same elevation of duct2
						placeElbow(tray3, tray2)
					
				baseNewLine.Dispose()