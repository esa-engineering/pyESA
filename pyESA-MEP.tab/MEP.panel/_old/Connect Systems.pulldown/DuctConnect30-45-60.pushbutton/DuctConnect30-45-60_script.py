# -*- coding: utf-8 -*-
__title__ = 'Duct Connect 30°/45°/60°'
__doc__ =	"""Version = 1.0
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

from pyrevit import revit, DB
from pyrevit import PyRevitException, PyRevitIOError
from pyrevit import forms

doc = revit.doc
### FUNCTIONS

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

def	connectDuctsWithDuct(doc, p1, p2):
	unused1 = list(p1.ConnectorManager.UnusedConnectors)
	unused2 = list(p2.ConnectorManager.UnusedConnectors)
	temp = []	
	for u1 in unused1:
		for u2 in unused2:
			temp.append( [u1.Origin.DistanceTo(u2.Origin), u1, u2] )
	
	conn01 = sorted(temp)[0][1]
	conn02 = sorted(temp)[0][2]
	
	levId = p1.Parameter[DB.BuiltInParameter.RBS_START_LEVEL_PARAM].AsElementId()
	newDuct = Duct.Create(doc, duct1.GetTypeId(), levId, conn01, conn02)
	return	newDuct

def createDuct(doc, p1, pt1, pt2):
	unused1 = list(p1.ConnectorManager.UnusedConnectors)
	dists = sorted([[u.Origin.DistanceTo(pt1), u] for u in unused1])
	conn01 = dists[0][1]

	levId = p1.Parameter[DB.BuiltInParameter.RBS_START_LEVEL_PARAM].AsElementId()
	newDuct = Duct.Create(doc, duct1.GetTypeId(), levId, conn01, pt2)
	return	newDuct

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
slopeDef = forms.CommandSwitchWindow.show([30,45,60,90], message='Duct connector slope [degrees]:')
if slopeDef:	slopeDef = 90-testFloat(slopeDef)
else:	slopeDef = 0

BIC = DB.BuiltInCategory
nr = 0

with forms.ProgressBar(title='select Ducts to connect - press Esc to stop', cancellable=True) as pb:

	selected_ducts = []
	prog = 0
	pb.update_progress(prog, 100)
	for duct in revit.get_picked_elements_by_category(BIC.OST_DuctCurves, "Select Duct element"):
		prog += 50
		pb.update_progress(prog, 100)
		selected_ducts.append(duct)
		if duct == None:
			break
		elif prog == 100:
			prog = 0
			duct1 = selected_ducts[0]
			duct2 = selected_ducts[1]
			selected_ducts = []

			ln1 = duct1.Location.Curve
			ln2 = duct2.Location.Curve

			# EXTEND/TRIM FIRST DUCT
			# pts1 = (ln1.GetEndPoint(0), ln1.GetEndPoint(1))
			# pts2 = (ln2.GetEndPoint(0), ln2.GetEndPoint(1))
			# dists = []
			# for p1 in pts1:
			# 	for p2 in pts2:	dists.append([p1.DistanceTo(p2), p1, p2])
			# pts = sorted(dists)[0][1:]
			# pt1 = pts[0]
			# pt2 = DB.XYZ(pts[1].X, pts[1].Y, pts[0].Z)
			# baseNewLine = DB.Line.CreateBound(pt1, pt2)
			baseNewLine = get_connectingLine(ln1, ln2)

			# DEFINE VALUE FOR DUCT INCLINATION
			# deltaH = math.fabs(pt1.Z - pts[1].Z)
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
					duct1.Location.Curve = baseNewLine
					doc.Delete(duct2.Id)

				elif ln1.Direction.IsAlmostEqualTo(baseNewLine.Direction):
					if baseNewLine.GetEndPoint(0).IsAlmostEqualTo(ln1.GetEndPoint(0)):
						baseNewLine = DB.Line.CreateBound(ln1.GetEndPoint(1),
									baseNewLine.GetEndPoint(1).Subtract(oppDir))
					else:	
						baseNewLine = DB.Line.CreateBound(ln1.GetEndPoint(0),
									baseNewLine.GetEndPoint(1).Subtract(oppDir))
					duct1.Location.Curve = baseNewLine
					duct3 = connectDuctsWithDuct(doc, duct1, duct2)
					# CONNECT
					placeElbow(duct1, duct3)
					placeElbow(duct3, duct2)

				else:
					# CREATE NEW DUCT
					duct3 = createDuct(doc, duct1, 
							baseNewLine.GetEndPoint(0),
							baseNewLine.GetEndPoint(1).Subtract(oppDir))
					
					# CONNECT THEM ALL
					placeElbow(duct1, duct3)
					
					try:
						duct4 = connectDuctsWithDuct(doc, duct3, duct2)
						placeElbow(duct3, duct4)
						placeElbow(duct4, duct2)
					except:
						# in case duct3 is at same elevation of duct2
						placeElbow(duct3, duct2)
					
				baseNewLine.Dispose()


