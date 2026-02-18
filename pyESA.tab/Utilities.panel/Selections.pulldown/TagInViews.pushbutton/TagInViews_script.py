__title__ = "Tag\ninViews"

__doc__ ="""Select all instances of selected Tag in specified views.
"""

__author__ = "Antonio Miano"

# REFERENCES
import pyrevit
import System

from pyrevit import script, revit, DB, UI, forms
from System.Collections.Generic import List

# DEFINITIONS
doc = revit.doc
uidoc = __revit__.ActiveUIDocument
l_tolist = lambda x : x if hasattr(x, '__iter__') else [x]


class CustomISelectionFilter(UI.Selection.ISelectionFilter):
	def AllowElement(self, element):
		if isinstance(element, DB.IndependentTag):
			return True
		return False

# INPUTS
tag_ref = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element, CustomISelectionFilter(), 'Select Tag')
if not tag_ref:	script.exit()

views_all = revit.query.get_all_views()
views_all.sort(key=lambda x:x.Name)
views = forms.SelectFromList.show(
	views_all,
	title='Select Views',
	multiselect = True,
	name_attr = 'Name',
	button_name = 'Select'
)
if not views:	script.exit()

# CODE
tag_type_id = doc.GetElement(tag_ref.ElementId).GetTypeId()
tag_elems_ids = doc.GetElement(tag_type_id).GetDependentElements(DB.ElementClassFilter(DB.IndependentTag))

views_ids = [view.Id.IntegerValue for view in views]
tag_elems = [doc.GetElement(item) for item in tag_elems_ids]
tag_views = [item.OwnerViewId for item in tag_elems]
tag_views_ids = [item.IntegerValue for item in tag_views]

tag_selection = List[DB.ElementId]()
for tag, tag_view_id in zip(tag_elems, tag_views_ids):
	if tag_view_id in views_ids:
		tag_selection.Add(tag.Id)

uidoc.Selection.SetElementIds(tag_selection)