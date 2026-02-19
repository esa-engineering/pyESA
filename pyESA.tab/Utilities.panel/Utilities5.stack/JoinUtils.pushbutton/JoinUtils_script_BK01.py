#REFERENCES
import pyrevit
from pyrevit import revit, DB, script
from rpw.ui.forms import TaskDialog
from itertools import combinations

doc = __revit__.ActiveUIDocument.Document

#CODE
##Get selected items and their combination
elements = list(revit.get_selection())
comb_elements = list(combinations(elements, 2))


##Switch the join order of joined elements
nStart = len(elements)
eElems = 0

with revit.Transaction('JoinUtils'):
	jgu = DB.JoinGeometryUtils
	for couple in comb_elements:
		try:
			if jgu.AreElementsJoined(doc,couple[0],couple[1]):
				if __shiftclick__:
					jgu.UnjoinGeometry(doc,couple[0],couple[1])
				else:
					jgu.SwitchJoinOrder(doc,couple[0],couple[1])			
				eElems += 1
		except:	pass

##Create output message
if __shiftclick__:
	msg = '{} Elements Selected\n{} Elements Unjoined'.format(str(nStart),str(eElems))
else:	
	msg = '{} Elements Selected\n{} Joins Order Switched'.format(str(nStart),str(eElems))
dialog = TaskDialog('Join Switch', content = msg, buttons = ['OK'], footer = '', show_close = True)
dialog.show(exit = True)
