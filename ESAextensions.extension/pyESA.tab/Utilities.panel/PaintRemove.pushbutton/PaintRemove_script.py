#REFERENCES
import pyrevit
from pyrevit import revit, DB
from rpw.ui.forms import TaskDialog

#from Autodesk.Revit import DB
#from Autodesk.Revit.DB import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

#DEFINITIONS
def tolist(obj1):
	if hasattr(obj1,"__iter__"): return obj1
	else: return [obj1]

def flatten(x):
	result = []
	for el in x:
		if hasattr(el, "__iter__") and not isinstance(el, basestring):
			result.extend(flatten(el))
		else:
			result.append(el)
	return result

#INPUTS
elements = revit.get_selection()

#CODE
opt = DB.Options()
opt.ComputeReferences = True

##Find faces from element solid geometry
elem_geoms = []
elements2 = []
for elem in elements:
	if elem.get_Geometry(opt) != None:
		elem_geoms.append(elem.get_Geometry(opt))
		elements2.append(elem)

elem_faces = []
for elem_geom in elem_geoms:
	e_faces = []
	for e_geom in elem_geom:
		if isinstance(e_geom, DB.Solid):
			for f in e_geom.Faces:
				e_faces.append(f)
	elem_faces.append(e_faces)

##Add regions if any
reg_faces = []
for elem_face in elem_faces:
	reg_face = [e_face.GetRegions() for e_face in elem_face]
	reg_faces.append(flatten(reg_face))

for elem_face,reg_face in zip(elem_faces,reg_faces):
	elem_face.extend(reg_face)
	
##Remove Paint from faces
with revit.Transaction('RemovePaint'):
	n_paints = []
	try:
		for elem2,elem_face in zip(elements2,elem_faces):
			n_paint = 0
			for face in elem_face:
				if doc.IsPainted(elem2.Id,face):
					n_paint += 1
					doc.RemovePaint(elem2.Id,face)
			n_paints.append(n_paint)
		msg = 'Paint removed from {} faces'.format(str(sum(n_paints)))

	except Exception, e:
		msg = str(e)

uidoc.RefreshActiveView()

##Display results
dialog = TaskDialog('Remove Paint', content = msg, buttons = ['OK'], show_close = True)
dialog.show(exit = True)