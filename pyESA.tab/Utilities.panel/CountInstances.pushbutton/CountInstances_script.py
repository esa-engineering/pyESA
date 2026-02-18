#REFERENCES
import pyrevit
from pyrevit import revit, DB, script
from pyrevit import forms

from System.Collections.Generic import List

uidoc = __revit__.ActiveUIDocument

#DEFINITIONS
tolist = lambda x : x if hasattr(x, '__iter__') else [x]

#INPUTS
if __shiftclick__: 
	with forms.WarningBar(title='Select RVT Link'):
		try:
			link_inst = revit.pick_element_by_category(DB.BuiltInCategory.OST_RvtLinks,message='Select RVT Link')
			doc = link_inst.GetLinkDocument()
			doc_name = str(doc.Title)+'.rvt'
		except:
			script.exit()
else:
	doc = revit.doc
	try:
		if doc.IsWorkshared:
			doc_name = DB.BasicFileInfo.Extract(doc.PathName).CentralPath.encode('string-escape').split('\\')[-1]
		else:
			doc_name = str(doc.Title)+'.rvt'
	except:
		doc_name = str(doc.Title) + '.rvt'

#CODE
##Collect all model categories except for specified ones
cats = doc.Settings.Categories
cats_mod = []
for c in cats:
	if (
		c.CategoryType == DB.CategoryType.Model 
		and c.Name != 'Project Information'
		and c.Name != 'Lines'
#		and c.Name != 'Materials'
		and c.Name != 'Analysis Display Style'
		and c.Name != 'RVT Links'
#		and c.Name != 'Sheets'
		and 'dwg' not in c.Name
		):
		cats_mod.append(c)

for c in cats:
	if (
		c.Name == 'Levels'
		or c.Name == 'Grids'
	):
		cats_mod.append(c)

##Create an ElementMulticategoryFilter
cats_mod_ids = List[DB.ElementId]()
for cm in tolist(cats_mod):
	cats_mod_ids.Add(cm.Id)
cats_filter = DB.ElementMulticategoryFilter(cats_mod_ids)

##Collect all elements of filtered categories
elems = DB.FilteredElementCollector(doc).WherePasses(cats_filter).WhereElementIsNotElementType().ToElements()

##Group elements per category and count
el_cat_name = [el.Category.Name+': ' for el in elems]
elems_dict = {k: [elems[i] for i in [j for j,x in enumerate(el_cat_name) if x==k]] for k in set(el_cat_name)}
elems_dict_len = {k: len(elem) for k,elem in elems_dict.items()}
cats_count = [k+str(i) for k,i in elems_dict_len.items()]
total_count = sum(elems_dict_len.values())
cats_count.sort()
cats_count_print = [str(cc) + '\n' for cc in cats_count]


##Create output
space = '----------'
#print 'TOTAL INSTANCES GROUPED BY CATEGORIES'
#print '---'
print 'Document Title: {0}'.format(doc_name)
print space
for cc in cats_count:
	print cc
print space
print 'GRAND TOTAL: {0}'.format(total_count)