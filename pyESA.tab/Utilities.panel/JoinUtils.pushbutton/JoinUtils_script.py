#REFERENCES
import pyrevit
from pyrevit import revit, DB, script
from pyrevit import forms
from rpw.ui.forms import TaskDialog, SelectFromList
from itertools import combinations

doc = __revit__.ActiveUIDocument.Document

#CODE
##Get selected items and their combination
elements = list(revit.get_selection())
comb_elements = list(combinations(elements, 2))


sfl_join = 'Join'
sfl_unjoin = 'Unjoin'
sfl_switch = 'Switch Join Order'

##Create inputs for the flexform and show it
operation = SelectFromList(
	'Join Utils',
	[sfl_join, sfl_unjoin, sfl_switch]
)

##Switch the join order of joined elements
nStart = len(elements)
eElems = 0

with revit.Transaction('JoinUtils', swallow_errors=True):
	jgu = DB.JoinGeometryUtils
	for couple in comb_elements:
		if operation == sfl_join:
			try:
				if not jgu.AreElementsJoined(doc,couple[0],couple[1]):
					jgu.JoinGeometry(doc,couple[0],couple[1])
					eElems += 1
			except:	pass
		elif operation == sfl_unjoin:
			try:
				if jgu.AreElementsJoined(doc,couple[0],couple[1]):
					jgu.UnjoinGeometry(doc,couple[0],couple[1])
					eElems += 1
			except:	pass
		else:
			try:
				if jgu.AreElementsJoined(doc,couple[0],couple[1]):
					jgu.SwitchJoinOrder(doc,couple[0],couple[1])
					eElems += 1
			except:	pass

##Create output message
if operation == sfl_join:
	msg = '{} Elements Selected\n{} Elements Joined'.format(str(nStart),'[]')
elif operation == sfl_unjoin:
	msg = '{} Elements Selected\n{} Elements Unjoined'.format(str(nStart),str(eElems))
elif operation == sfl_switch:	
	msg = '{} Elements Selected\n{} Joins Order Switched'.format(str(nStart),str(eElems))

dialog = TaskDialog('Join Utils', content = msg, buttons = ['OK'], footer = '', show_close = True)
dialog.show(exit = True)
