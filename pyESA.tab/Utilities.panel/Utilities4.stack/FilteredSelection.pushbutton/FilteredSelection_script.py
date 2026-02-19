__title__ = "Filtered\nSelection"

__doc__ = """Select elements in the document filtering by:
- Categoies
- Levels
- Loadable Families
- Materials
- Worksets
---
If some elements are already selected, it will apply the
filtering rules inside the selection"""

__author__ = "Antonio Miano"


#REFERENCES
import pyrevit
import System

from pyrevit import revit, script
# import DB
from pyrevit import DB
from pyrevit import forms

from System.Collections.Generic import List, IList


#DEFINITIONS
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
BIC = DB.BuiltInCategory
BIP = DB.BuiltInParameter
collector = DB.FilteredElementCollector

def f_flatten(x):
	result = []
	if x != None:
		for el in x:
			if hasattr(el, "__iter__") and not isinstance(el, basestring):
				result.extend(f_flatten(el))
			else:
				result.append(el)
		return result
	else:
		return None

def f_get_category(item):
	item_category = None
	try:
		item_category = item.Category
		return item_category.Id
	except:
		return item_category

def f_get_level(item):
	item_level = None
	try:
		item_level = item.Document.GetElement(item.LevelId)
		if item_level:	return item_level.Id
		else:
			try:
				item_level = item.ReferenceLevel
				if item_level:	return item_level.Id
			except:
				try:
					item_level_id = item.Parameter[BIP.STAIRS_BASE_LEVEL_PARAM].AsElementId()
					if item_level_id:	return item_level_id
				except:	return item_level
	except: return item_level

def f_get_family(item):
	item_family = None
	try:
		item_family = item.Symbol.Family
		return item_family.Id
	except:
		return item_family

def f_get_materials_01(item):
	item_materials = None
	try:
		item_materials_ids = []
		item_materials_ids.extend(list(item.GetMaterialIds(False)))
		item_materials_ids.extend(list(item.GetMaterialIds(True)))
		if len(item_materials_ids) == 0:
			try:
				item_sys_type = item.Document.GetElement(item.MEPSystem.GetTypeId())
				item_materials_ids = item_sys_type.MaterialId
				return item_materials_ids
			except:
				return item_materials
		else:
			return item_materials_ids
	except:
		return item_materials

def f_get_materials_02(item):
	item_materials_ids = []
	##First attempt: get material directly from the element (fast)
	try:
		item_materials_ids.extend(list(item.GetMaterialIds(False)))
		item_materials_ids.extend(list(item.GetMaterialIds(True)))
		if len(item_materials_ids) > 0:
			return list(set(item_materials_ids))
		else:
			##Second attempt: for MEP system get material from its system type (fast)
			try:
				item_sys_type = item.Document.GetElement(item.MEPSystem.GetTypeId())
				item_materials_ids.append(item_sys_type.MaterialId)
				if len(item_materials_ids) > 0:
					return list(set(item_materials_ids))
			except:
				###Third attempt: If neither of the previous two methods worked get materials directly from the geometry of the element (slow)
				try:
					item_geometries = []
					item_geometries.extend(item.get_Geometry(DB.Options()))
					###Add a speciic condition for railings, to get also the geometry of the top rail
					if type(item) == DB.Architecture.Railing:
						item_top = item.Document.GetElement(item.TopRail)
						if item_top != None:
							item_geometries.extend(item_top.get_Geometry(DB.Options()))
					###Iterate through the geometries of the element, searching for the solid's faces to get material
					for item_geom in item_geometries:
						if type(item_geom) == DB.Solid:
							for face in item_geom.Faces:
								item_materials_ids.append(face.MaterialElementId)
								face.Dispose()
						if type(item_geom) == DB.GeometryInstance:
							solids = [sld for sld in item_geom.GetInstanceGeometry() if type(sld) == DB.Solid]
							for solid in solids:
								for face in solid.Faces:
									item_materials_ids.append(face.MaterialElementId)
									face.Dispose()
								solid.Dispose()
						# else:
						# 	item_materials_ids.append(None)
						item_geom.Dispose()
					if len(item_materials_ids) > 0:
						return list(set(item_materials_ids))
					else:	return None
				except:	return None
	except:	return None

def f_get_workset(item):
	item_workset = None
	if item.Document.IsWorkshared:
		try:
			item_workset = item.Document.GetWorksetTable().GetWorkset(item.WorksetId)
			return item_workset.Id
		except:
			return item_workset
	else:
		return item_workset


#INPUTS
##Collect Categories
categories = doc.Settings.Categories
filtered_categories = []
for cat in categories:
	 if (
		cat.CategoryType == DB.CategoryType.Model
		or cat.CategoryType == DB.CategoryType.Annotation
	 ):
		 filtered_categories.append(cat)
filtered_categories.sort(key=lambda x:x.Name)

##Collect Levels
levels = list(collector(doc).OfClass(DB.Level).ToElements())
levels.sort(key=lambda x:x.Elevation)

##Collect Families
families = list(collector(doc).OfClass(DB.Family).ToElements())
families.sort(key=lambda x:x.Name)

##Collect Materials
materials = list(collector(doc).OfClass(DB.Material).ToElements())
materials.sort(key=lambda x:x.Name)

##Collect Worksets
if doc.IsWorkshared:
	worksets = list(DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset))
	worksets.sort(key=lambda x:x.Name)
else:
	worksets = []

##Prepare Form
options = {
	'Categories': filtered_categories,
	'Levels': levels,
	'Loadable Families': families,
	'Materials': materials,
	'Worksets': worksets
}

rules = forms.SelectFromList.show(
	options,
	multiselect=True,
	name_attr='Name',
	group_selector_title='Filter by',
	button_name='Select'
)

user_selection = revit.get_selection()
user_selection_test = False
if len(user_selection)>0:
	user_selection_test = True

#CODE
if not rules:
	script.exit()

##Collect all the instances in the document that belong to model or annotation categories or only elements selected by the user
category_ids = List[DB.ElementId]()
for cat in filtered_categories:
	category_ids.Add(cat.Id)
category_filter = DB.ElementMulticategoryFilter(category_ids)
if user_selection_test:
	instances = user_selection
else:
	instances = collector(doc).WherePasses(category_filter).WhereElementIsNotElementType().ToElements()

##Create a dictionary collecting all the instances' properties necessary for filtering
instances_dict = {instance.Id: f_flatten((
	f_get_category(instance),
	f_get_level(instance),
	f_get_family(instance),
	f_get_materials_02(instance),
	f_get_workset(instance)
)) for instance in instances}

##Filter the instances according to criteria
instances_selection_ids = []
for id, properties in instances_dict.items():
	for rule in rules:
		if rule.Id in properties:
			instances_selection_ids.append(id)

selection_ids = List[DB.ElementId]()
for item in instances_selection_ids:
	selection_ids.Add(item)

##Select the filtered instances in the document
uidoc.Selection.SetElementIds(selection_ids)