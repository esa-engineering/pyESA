# -*- coding: utf-8 -*-
__title__ = "Section\nFrom Lines"

__doc__ = """Creates Section Views perpendicular to selected model lines.
Pick a reference point to define the viewing direction."""

__author__ = "Antonio Miano"

# REFERENCES
from pyrevit import revit, script, DB, UI
from pyrevit import forms
from rpw.ui.forms import FlexForm, Label, TextBox, ComboBox, Separator, Button

doc = revit.doc
uidoc = revit.uidoc

# DEFINITIONS
class CurveSelectionFilter(UI.Selection.ISelectionFilter):
	def AllowElement(self, element):
		try:
			return element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Lines)
		except:
			return False
	def AllowReference(self, _reference, _position):
		return False

# INPUTS
## Collect section view family types
def get_element_name(element, fallback_param):
	param = element.get_Parameter(fallback_param)
	if param and param.AsString():
		return param.AsString()
	return str(element.Id.IntegerValue)

section_types_dict = {}
for vft in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
	if vft.ViewFamily != DB.ViewFamily.Section:
		continue
	name = get_element_name(vft, DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
	section_types_dict[name] = vft
if not section_types_dict:
	forms.alert('No Section View Family Type found in this document.', exitscript=True)

## Collect view templates
view_templates_dict = {'< None >': None}
for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
	if not v.IsTemplate:
		continue
	name = get_element_name(v, DB.BuiltInParameter.VIEW_NAME)
	view_templates_dict[name] = v

## Form for section dimensions
components = [
	Label('Section Height [mm]'),
	TextBox('tb_height', '3000'),
	Label('Section Depth [mm]'),
	TextBox('tb_depth', '3000'),
	Separator(),
	Label('Section Type'),
	ComboBox('cb_section_type', section_types_dict),
	Label('View Template'),
	ComboBox('cb_view_template', view_templates_dict),
	Separator(),
	Button('OK')
]
flex_form = FlexForm('Create Sections From Lines', components)
flex_form.show()
if not flex_form.values:
	script.exit()

try:
	section_height_mm = int(flex_form.values['tb_height'])
	section_depth_mm = int(flex_form.values['tb_depth'])
except (ValueError, KeyError):
	forms.alert('Please enter valid integer values for height and depth.', exitscript=True)

selected_section_type = flex_form.values['cb_section_type']
selected_template     = flex_form.values['cb_view_template']

## Convert mm to internal units
if int(doc.Application.VersionNumber) < 2022:
	section_height = DB.UnitUtils.ConvertToInternalUnits(section_height_mm, DB.DisplayUnitType.DUT_MILLIMETERS)
	section_depth  = DB.UnitUtils.ConvertToInternalUnits(section_depth_mm,  DB.DisplayUnitType.DUT_MILLIMETERS)
else:
	section_height = DB.UnitUtils.ConvertToInternalUnits(section_height_mm, DB.UnitTypeId.Millimeters)
	section_depth  = DB.UnitUtils.ConvertToInternalUnits(section_depth_mm,  DB.UnitTypeId.Millimeters)

## Pick model lines
selection_filter = CurveSelectionFilter()
with forms.WarningBar(title='Select model lines, then press Finish'):
	try:
		curve_refs = list(uidoc.Selection.PickObjects(
			UI.Selection.ObjectType.Element,
			selection_filter,
			'Select model lines'
		))
	except:
		script.exit()

if not curve_refs:
	script.exit()

curves = [doc.GetElement(cr.ElementId) for cr in curve_refs]

## Pick reference point for viewing direction
with forms.WarningBar(title='Pick a point to define the viewing direction'):
	try:
		point = uidoc.Selection.PickPoint('Pick a point to define the viewing direction')
	except:
		script.exit()

# CODE
errors = []
with revit.Transaction('SectionLines_CreateFromLines'):
	for curve in curves:
		try:
			# Geometry from selected curve
			crv = curve.GeometryCurve
			crv_start = crv.GetEndPoint(0)
			crv_end   = crv.GetEndPoint(1)

			# Project the picked point onto the (infinite) curve line
			line_inf  = DB.Line.CreateUnbound(crv_start, crv_end)
			pt_on_crv = line_inf.Project(point).XYZPoint
			line_dist = point - pt_on_crv

			# Curve direction vectors and length
			crv_dist_1   = crv_end - crv_start
			crv_dist_2   = crv_start - crv_end
			crv_dist_len = crv_dist_1.GetLength()

			# Build local coordinate system
			new_x    = crv_dist_1.Normalize()
			new_y    = DB.XYZ.BasisZ
			new_z    = new_x.CrossProduct(new_y)
			sec_start = crv_start

			# Flip direction if picked point is on the opposite side
			if new_z.Normalize().DotProduct(line_dist.Normalize()) < 0:
				new_x     = crv_dist_2.Normalize()
				new_z     = new_x.CrossProduct(new_y)
				sec_start = crv_end

			transf          = DB.Transform.Identity
			transf.Origin   = sec_start
			transf.BasisX   = new_x
			transf.BasisY   = new_y
			transf.BasisZ   = new_z

			# BoundingBox for the new section view
			bbox           = DB.BoundingBoxXYZ()
			bbox.Min       = DB.XYZ(0, 0, 0)
			bbox.Max       = DB.XYZ(crv_dist_len, section_height, section_depth)
			bbox.Transform = transf

			new_view = DB.ViewSection.CreateSection(doc, selected_section_type.Id, bbox)
			if selected_template:
				new_view.ViewTemplateId = selected_template.Id

		except Exception as e:
			errors.append(str(e))

if errors:
	forms.alert(
		'Some sections could not be created:\n\n' + '\n'.join(errors),
		title='SectionLines — Errors'
	)
