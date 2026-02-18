#REFERENCES
import pyrevit
from pyrevit import script, revit, DB, UI
from pyrevit import forms
from rpw.ui.forms import TextInput

from System.Collections.Generic import List

doc = revit.doc
uidoc = __revit__.ActiveUIDocument
username = str(__revit__.Application.Username)
BIC = DB.BuiltInCategory

#DEFINITIONS
tolist = lambda x : x if hasattr(x, '__iter__') else [x]

def f_bb_min(boundingboxes):
	min_x = min([bb.Min.X for bb in boundingboxes])
	min_y = min([bb.Min.Y for bb in boundingboxes])
	min_z = min([bb.Min.Z for bb in boundingboxes])
	return DB.XYZ(min_x, min_y, min_z)

def f_bb_max(boundingboxes):
	max_x = max([bb.Max.X for bb in boundingboxes])
	max_y = max([bb.Max.Y for bb in boundingboxes])
	max_z = max([bb.Max.Z for bb in boundingboxes])
	return DB.XYZ(max_x, max_y, max_z)

#INPUTS
view_name =	'3D_SBRL_{user}'.format(user = username)
offset = 0.02 #scale percentage

##Select linked elements
bboxes = []
if __shiftclick__:
	with forms.WarningBar(title='Select RVT Link'):
		try:
			link_inst = revit.pick_element_by_category(BIC.OST_RvtLinks,message='Select RVT Link')
		except:
			script.exit()
	if link_inst:
		form_text = TextInput(title='Linked Elements\' ID(s)',description='Use \';\' as separator (eg: 111;222;...)')
	else:
		script.exit()
	if len(form_text)>0:
		link_inst_transform = link_inst.GetTotalTransform()
		link_doc = link_inst.GetLinkDocument()
		form_text_ids = form_text.replace('; ',';').replace(' ;',';').split(';')
		link_elem_ids = []
		for fti in form_text_ids:
			try:	link_elem_ids.append(int(fti))
			except:	pass
		if len(link_elem_ids)>0:
			for lei in link_elem_ids:
				link_elem = link_doc.GetElement(DB.ElementId(lei))
				try:
					link_elem_geom = link_elem.get_Geometry(DB.Options()).GetTransformed(link_inst_transform)
					bb = link_elem_geom.GetBoundingBox()
					link_elem_geom.Dispose()
					bboxes.append(bb)
				except:
					script.exit()
		else:
			script.exit()
else:
	with forms.WarningBar(title='Select linked elements and then press Finish'):
		try:
			elem_refs = tolist(uidoc.Selection.PickObjects(UI.Selection.ObjectType.LinkedElement, "Select linked elements"))
		except:
			script.exit()
	for er in elem_refs:
		link_inst = doc.GetElement(er.ElementId)
		link_inst_transform = link_inst.GetTotalTransform()
		link_doc = link_inst.GetLinkDocument()
		link_elem = link_doc.GetElement(er.LinkedElementId)
		link_elem_geom = link_elem.get_Geometry(DB.Options()).GetTransformed(link_inst_transform)
		bb = link_elem_geom.GetBoundingBox()
		link_elem_geom.Dispose()
		bboxes.append(bb)

##Find BoundingBox for selected linked elements
bbox_min = f_bb_min(bboxes)
bbox_max = f_bb_max(bboxes)
new_outline = DB.Outline(bbox_min,bbox_max)
new_outline.Scale(1+offset)
no_min = new_outline.MinimumPoint
no_max = new_outline.MaximumPoint

new_bbox = DB.BoundingBoxXYZ()
new_bbox.Min = no_min
new_bbox.Max = no_max

new_outline.Dispose()

##Create View3D with SectionBox
views_3D = DB.FilteredElementCollector(doc).OfClass(DB.View3D).ToElements()
view_bbox = [v3d for v3d in views_3D if v3d.Name == view_name]

view_types = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
for vt in view_types:
	if vt.ViewFamily == DB.ViewFamily.ThreeDimensional:
		view_type = vt
	else:
		continue

with revit.Transaction('SectionBox_RevitLink'):

	if type(doc.ActiveView) == DB.View3D:
		view_sectionbox = doc.ActiveView

	elif len(view_bbox) == 1:
		view_sectionbox = view_bbox[0]
#		doc.Delete(view_bbox[0].Id)
	else:
		view_sectionbox = DB.View3D.CreateIsometric(doc,view_type.Id)
		view_sectionbox.Name = view_name
	view_sectionbox.SetSectionBox(new_bbox)

#Switch to 3D view and zoom to fit
uidoc.RefreshActiveView()
uidoc.RequestViewChange(view_sectionbox)

for uiv in uidoc.GetOpenUIViews():
	if uiv.ViewId == view_sectionbox.Id:
		uiv.ZoomToFit()
		break
