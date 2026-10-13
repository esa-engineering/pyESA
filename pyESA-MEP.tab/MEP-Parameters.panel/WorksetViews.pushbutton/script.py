# -*- coding: utf-8 -*-
__title__ = "Workset\nViews"
__doc__ = """Version = 0.1
Date    = 07.10.2026
________________________________________________________________
Description:
Creates one 3D view per user workset, named
102_WRK_Workset_<last part of the workset name> Control
(e.g. 04_MEP_Electrical -> 102_WRK_Workset_Electrical Control),
showing only that workset, to check which elements belong to it.

- CLICK: choose the worksets; SHIFT + CLICK: all worksets
- Existing views with the same name are updated, not duplicated
- You choose a 3D view template (or none) for the views
- If the views have the parameters u_OTH_ViewChapter1 and
  u_OTH_ViewChapter2, they are set to 100_CON_Control and
  102_WRK_Worksets
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

from pyrevit import revit, script, forms, DB

doc = revit.doc
output = script.get_output()


REPORT_CONFIG = "ESA_Report"  # impostazione condivisa da tutti i pulsanti, cambiata col pulsante Report


def want_report(total, problems, noun):
    """True se va stampato il report completo. Modalita' "problems" (predefinita): solo se qualcosa non e' andato a
    buon fine; "always": sempre; "never": mai, resta la notifica. Altrimenti notifica breve (e nessuna finestra)."""
    try:
        mode = script.get_config(REPORT_CONFIG).get_option("mode", "problems")
    except Exception:
        mode = "problems"
    if mode == "always" or (mode == "problems" and problems):
        return True
    if not total:
        return False
    text = "{} {}".format(total - problems, noun) + (", {} not done".format(problems) if problems else "")
    try:
        script.toast(text, title="pyESA")
    except Exception:
        return problems > 0  # notifica non disponibile: con dei problemi meglio il report che niente
    return False

VIEW_PREFIX = "102_WRK_Workset_"
VIEW_SUFFIX = " Control"
VIEW_PARAMETERS = (
    ("u_OTH_ViewChapter1", "100_CON_Control"),
    ("u_OTH_ViewChapter2", "102_WRK_Worksets"),
)
NO_TEMPLATE = "<No view template>"
# caratteri non ammessi nei nomi delle viste Revit
INVALID_NAME_CHARS = '\\:{}[]|;<>?`~'


def get_element_id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


def view_name_for(workset):
    """Es. 04_MEP_Electrical -> 102_WRK_Workset_Electrical Control (senza underscore: nome intero)."""
    name = workset.Name.rsplit("_", 1)[-1].strip() or workset.Name
    for ch in INVALID_NAME_CHARS:
        name = name.replace(ch, "-")
    return "{}{}{}".format(VIEW_PREFIX, name, VIEW_SUFFIX)


def get_user_worksets():
    worksets = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset).ToWorksets()
    return sorted(worksets, key=lambda w: w.Name)


def get_3d_view_type_id():
    for vft in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
        if vft.ViewFamily == DB.ViewFamily.ThreeDimensional:
            return vft.Id
    return None


def get_existing_3d_views():
    """Viste 3D (non modelli) per nome."""
    views = {}
    for view in DB.FilteredElementCollector(doc).OfClass(DB.View3D):
        if not view.IsTemplate:
            views[view.Name] = view
    return views


def get_3d_templates():
    templates = [v for v in DB.FilteredElementCollector(doc).OfClass(DB.View3D) if v.IsTemplate]
    return sorted(templates, key=lambda v: v.Name)


def template_controls_worksets(template):
    """True se il modello di vista controlla la visibilita' dei workset."""
    ws_param = get_element_id_value(DB.ElementId(DB.BuiltInParameter.VIS_GRAPHICS_WORKSETS))
    controlled = set(get_element_id_value(i) for i in template.GetTemplateParameterIds())
    not_controlled = set(get_element_id_value(i) for i in template.GetNonControlledTemplateParameterIds())
    return ws_param in controlled and ws_param not in not_controlled


def choose_worksets(worksets):
    """CLICK: l'utente sceglie i workset. Esce se annulla o non sceglie nulla."""
    names = [ws.Name for ws in worksets]
    chosen = forms.SelectFromList.show(names, title="Worksets to create views for",
                                       button_name="Create views", multiselect=True)
    if not chosen:
        script.exit()
    return [ws for ws in worksets if ws.Name in chosen]


def choose_template():
    """Ritorna il modello scelto, None per nessun modello. Esce se l'utente annulla."""
    templates = get_3d_templates()
    if not templates:
        return None
    names = [NO_TEMPLATE] + [t.Name for t in templates]
    choice = forms.SelectFromList.show(names, title="View template for the workset views",
                                       button_name="Use this template", multiselect=False)
    if choice is None:
        script.exit()
    if choice == NO_TEMPLATE:
        return None
    return templates[names.index(choice) - 1]


def set_view_parameters(view):
    """Imposta i parametri dei capitoli se esistono. Ritorna (impostati, mancanti)."""
    done, missing = [], []
    for name, value in VIEW_PARAMETERS:
        param = view.LookupParameter(name)
        if param is None or param.IsReadOnly or param.StorageType != DB.StorageType.String:
            missing.append(name)
            continue
        param.Set(value)
        done.append(name)
    return done, missing


def apply_visibility(view, target, worksets):
    for ws in worksets:
        visibility = DB.WorksetVisibility.Visible if ws.Id == target.Id else DB.WorksetVisibility.Hidden
        view.SetWorksetVisibility(ws.Id, visibility)


def print_report(rows, template, template_warning):
    output.print_md("# Workset views report")
    output.print_md("View template: **{}**".format(template.Name if template else "none"))
    if template_warning:
        output.print_md("**Warning:** the template controls workset visibility, so the per-view "
                        "settings are overridden by the template. Check the views or use no template.")
    created = len([r for r in rows if r[2] == "Created"])
    updated = len([r for r in rows if r[2] == "Updated"])
    failed = len(rows) - created - updated
    output.print_md("Worksets: **{}** | views created: **{}** | updated: **{}** | failed: **{}**".format(
        len(rows), created, updated, failed))
    output.print_table(table_data=rows,
                       columns=["Workset", "View", "Action", "Parameters set", "Parameters missing / notes"])


def main():
    if not doc.IsWorkshared:
        output.print_md("# Workset views report")
        output.print_md("The model is not workshared: there are no worksets to create views for.")
        script.exit()

    worksets = get_user_worksets()
    if __shiftclick__:  # noqa: F821
        selected = worksets
    else:
        selected = choose_worksets(worksets)
    view_type_id = get_3d_view_type_id()
    if view_type_id is None:
        output.print_md("# Workset views report")
        output.print_md("No 3D view type found in the model: the views cannot be created.")
        script.exit()

    template = choose_template()
    template_warning = template is not None and template_controls_worksets(template)

    existing = get_existing_3d_views()
    rows = []
    used_names = {}
    with revit.Transaction("Workset 3D Views"):
        for ws in selected:
            name = view_name_for(ws)
            # due workset con la stessa parte finale darebbero la stessa vista: il secondo non viene elaborato
            if name in used_names:
                rows.append([ws.Name, name, "Failed", "-",
                             "Same view name as workset '{}'".format(used_names[name])])
                continue
            used_names[name] = ws.Name
            sub = DB.SubTransaction(doc)
            sub.Start()
            try:
                view = existing.get(name)
                if view is None:
                    view = DB.View3D.CreateIsometric(doc, view_type_id)
                    view.Name = name
                    action = "Created"
                else:
                    action = "Updated"
                view.ViewTemplateId = template.Id if template else DB.ElementId.InvalidElementId
                apply_visibility(view, ws, worksets)
                done, missing = set_view_parameters(view)
                sub.Commit()
                rows.append([ws.Name, output.linkify(view.Id), action,
                             ", ".join(done) or "-", ", ".join(missing) or "-"])
            except Exception as ex:
                sub.RollBack()
                rows.append([ws.Name, name, "Failed", "-", str(ex)])

    problems = len([r for r in rows if r[2] not in ("Created", "Updated")])
    # l avviso sul template va sempre letto: in quel caso il report si mostra comunque
    if template_warning or want_report(len(rows), problems, "views created or updated"):
        print_report(rows, template, template_warning)


main()
