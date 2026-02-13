#REFERENCES
import datetime
from pyrevit import revit, DB, UI, forms
from rpw.ui.forms import TaskDialog
#DEFINITIONS
tolist = lambda x : x if hasattr(x, '__iter__') else [x]

#CODE
doc = revit.doc
tday = datetime.date.today().strftime('%y%m%d')

if __shiftclick__:
	schedules = tolist(forms.select_schedules())
else:
	schedules = tolist(doc.ActiveView)

try:
	s_names = ['{0}_{1}.txt'.format(str(tday),s.Name) for s in schedules]
	folder_path = forms.pick_folder()
	vseo = DB.ViewScheduleExportOptions()
	ok = 0
	no = 0
	for s,sn in zip(schedules,s_names):
		try:
			s.Export(folder_path,sn,vseo)
			ok += 1
		except:
			no += 1

	msg = '{0} Schedules Exported\n{1} Schedules NOT Exported'.format(str(ok),str(no))
	dialog = TaskDialog('Schedule Export', content=msg, buttons=['OK'], footer='', show_close=True)
	dialog.show(exit=True)
except:	pass