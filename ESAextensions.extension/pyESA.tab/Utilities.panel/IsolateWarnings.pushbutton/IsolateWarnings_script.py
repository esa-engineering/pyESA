#REFERENCES
from pyrevit import revit, DB, UI, script
from pyrevit import PyRevitException, PyRevitIOError

from pyrevit import forms

from System.Collections.Generic import List

#CODE
doc = revit.doc
allWarnings = doc.GetWarnings()

##Create list from which select Warnings
allNames = [w.GetDescriptionText() for w in allWarnings]
NameSet = list(set(allNames))
for i in range(len(NameSet)):
	nr = str( allNames.count(NameSet[i]) ).zfill(3)
	NameSet[i] = "[{}] ".format(nr) + NameSet[i]

NameSet = sorted(NameSet)[::-1]
NameSet = ["[{}] ALL".format(str(len(allNames)).zfill(3))]+NameSet

i = forms.SelectFromList.show(NameSet, title='Select Warning to Isolate')
if not i:	script.exit()
i = i[i.index(" ")+1:]

##Get selected warning
if i != "ALL":	intWarnings = [w for w in allWarnings if w.GetDescriptionText() == i]
else:	intWarnings = allWarnings

##Get all failing elements
allElems = []
for w in intWarnings:
	allElems.extend(w.GetFailingElements())

##Temporary isolate failing elements in the active view
with revit.Transaction('IsolateWarnings'):
	vw = doc.ActiveView
	vw.TemporaryViewModes.DeactivateMode(DB.TemporaryViewMode.TemporaryHideIsolate)
	vw.IsolateElementsTemporary(List[DB.ElementId](allElems))

