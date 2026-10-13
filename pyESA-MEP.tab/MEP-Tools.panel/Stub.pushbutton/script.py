# -*- coding: utf-8 -*-
__title__ = "Spring"
__doc__ = """Version = 0.1
Date    = 07.10.2026
________________________________________________________________
Description:
Creates short pipe / duct stubs that start from the free mechanical
connectors of equipment (any element with piping or HVAC connectors).

Instruction:
- CLICK: a window sets the standard stub length, whether each stub
  gets a different length (+50 mm per connector), whether to use all
  connectors or choose them for each object, and the pipe / duct type
- Click each element, press ESC to finish: a report lists the stubs
- SHIFT + CLICK: no window; stubs on all free connectors of each
  clicked element, starting from the last length and adding 50 mm for
  each connector
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB.Plumbing import Pipe, PipeType
from Autodesk.Revit.DB.Mechanical import Duct, DuctType
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, forms, DB

import Stub_ui

XAML_FILE_NAME = "Stub_form.xaml"
CONFIG_NAME = "ESA_Stub"
DEFAULTS = {"length_mm": 200.0, "different": False, "choose": False, "pipe_type": "", "duct_type": ""}

doc = revit.doc
uidoc = revit.uidoc
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

MM_PER_FT = 304.8
MIN_LENGTH_FT = 0.01  # l'API rifiuta tratti piu' corti di 1/10 pollice
INCREMENT_MM = 50.0

MECHANICAL_DOMAINS = (DB.Domain.DomainPiping, DB.Domain.DomainHvac)
PHYSICAL_TYPES = (DB.ConnectorType.End, DB.ConnectorType.Curve, DB.ConnectorType.Physical)


def to_bool(value):
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ("true", "1", "yes")


def type_name(element_type):
    try:
        return element_type.Name
    except Exception:
        param = element_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
        return param.AsString() if param else "?"


def get_types(cls):
    """{nome: ElementId} dei tipi di una classe (PipeType, DuctType)."""
    types = {}
    for element_type in DB.FilteredElementCollector(doc).OfClass(cls):
        types[type_name(element_type)] = element_type.Id
    return types


def get_mechanical_connectors(element):
    """Connettori fisici di tubazioni o canali, ordinati per Id."""
    if not isinstance(element, DB.FamilyInstance) or element.MEPModel is None:
        return []
    manager = element.MEPModel.ConnectorManager
    if manager is None:
        return []
    found = []
    for conn in manager.Connectors:
        try:
            if conn.ConnectorType in PHYSICAL_TYPES and conn.Domain in MECHANICAL_DOMAINS:
                found.append(conn)
        except Exception:
            continue
    return sorted(found, key=lambda c: c.Id)


class MechanicalConnectorFilter(ISelectionFilter):
    """Accetta solo elementi con connettori meccanici."""

    def AllowElement(self, element):
        return len(get_mechanical_connectors(element)) > 0

    def AllowReference(self, reference, point):
        return False


def axis_label(vector):
    comps = (("+X", vector.X), ("-X", -vector.X), ("+Y", vector.Y), ("-Y", -vector.Y),
             ("up", vector.Z), ("down", -vector.Z))
    label, value = max(comps, key=lambda c: c[1])
    return label if value > 0.9 else "inclined"


def connector_size_text(conn):
    try:
        if conn.Domain == DB.Domain.DomainPiping or conn.Shape == DB.ConnectorProfileType.Round:
            return "{:.0f} mm".format(conn.Radius * 2 * MM_PER_FT)
        return "{:.0f} x {:.0f} mm".format(conn.Width * MM_PER_FT, conn.Height * MM_PER_FT)
    except Exception:
        return "size n/a"


def connector_system_text(conn):
    try:
        if conn.Domain == DB.Domain.DomainPiping:
            return str(conn.PipeSystemType)
        return str(conn.DuctSystemType)
    except Exception:
        return "no system"


def connector_label(conn, number):
    return "#{} | {} | {} | {} | {}".format(
        number, "Piping" if conn.Domain == DB.Domain.DomainPiping else "HVAC",
        connector_system_text(conn), connector_size_text(conn), axis_label(conn.CoordinateSystem.BasisZ))


def choose_connectors(element, free):
    """Elenco di scelta dei connettori liberi. Ritorna i connettori scelti (lista vuota se annulla)."""
    labels = [connector_label(conn, i) for i, conn in free]
    chosen = forms.SelectFromList.show(
        labels, title="Connectors of {}".format(element.Name if hasattr(element, "Name") else element.Id),
        button_name="Create stubs", multiselect=True)
    if not chosen:
        return []
    return [free[i][1] for i, label in enumerate(labels) if label in chosen]


def get_level_id(element, origin):
    if element.LevelId != DB.ElementId.InvalidElementId:
        return element.LevelId
    levels = sorted(DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements(),
                    key=lambda lvl: getattr(lvl, "ProjectElevation", lvl.Elevation))
    if not levels:
        return None
    below = [lvl for lvl in levels if getattr(lvl, "ProjectElevation", lvl.Elevation) <= origin.Z + 1e-6]
    return (below[-1] if below else levels[0]).Id


def create_stub(conn, length_ft, type_id, level_id):
    end_point = conn.Origin.Add(conn.CoordinateSystem.BasisZ.Multiply(length_ft))
    if conn.Domain == DB.Domain.DomainPiping:
        return Pipe.Create(doc, type_id, level_id, conn, end_point)
    return Duct.Create(doc, type_id, level_id, conn, end_point)


def process_element(element, settings, type_ids, base_length_mm):
    """Crea gli stub di un elemento. Ritorna una lista di dict (una riga di report per connettore)."""
    free = [(i + 1, c) for i, c in enumerate(get_mechanical_connectors(element)) if not c.IsConnected]
    rows = []
    if not free:
        return [{"element": element, "connector": "-", "length_mm": None, "new": None,
                 "status": "No free mechanical connectors"}]

    if settings["choose"]:
        selected = choose_connectors(element, free)
        if not selected:
            return [{"element": element, "connector": "-", "length_mm": None, "new": None,
                     "status": "Skipped: no connector chosen"}]
        numbers = dict((c.Id, n) for n, c in free)
        chosen = [(numbers[c.Id], c) for c in selected]
    else:
        chosen = free

    transaction = DB.Transaction(doc, "Spring")
    transaction.Start()
    try:
        for k, (number, conn) in enumerate(chosen):
            length_mm = base_length_mm + (INCREMENT_MM * k if settings["different"] else 0.0)
            row = {"element": element, "connector": connector_label(conn, number),
                   "length_mm": length_mm, "new": None, "status": ""}
            rows.append(row)
            domain_key = "pipe" if conn.Domain == DB.Domain.DomainPiping else "duct"
            type_id = type_ids.get(domain_key)
            if type_id is None:
                row["status"] = "No {} type selected or available".format(domain_key)
                continue
            level_id = get_level_id(element, conn.Origin)
            if level_id is None:
                row["status"] = "No level in the model"
                continue
            if length_mm / MM_PER_FT < MIN_LENGTH_FT:
                row["status"] = "Length too short"
                continue
            sub = DB.SubTransaction(doc)
            sub.Start()
            try:
                row["new"] = create_stub(conn, length_mm / MM_PER_FT, type_id, level_id)
                sub.Commit()
                row["status"] = "Created"
            except Exception as ex:
                sub.RollBack()
                row["status"] = "Not created: {}".format(ex)
        transaction.Commit()
    except Exception as ex:
        transaction.RollBack()
        for row in rows:
            row["new"] = None
            row["status"] = "Not created: {}".format(ex)
    return rows


def resolve_type_ids(settings, pipe_types, duct_types):
    def pick(types, wanted):
        if wanted in types:
            return types[wanted]
        return types[sorted(types)[0]] if types else None
    return {"pipe": pick(pipe_types, settings["pipe_type"]), "duct": pick(duct_types, settings["duct_type"])}


def load_settings(config):
    settings = dict(DEFAULTS)
    try:
        settings["length_mm"] = float(config.get_option("length_mm", DEFAULTS["length_mm"]))
        settings["different"] = to_bool(config.get_option("different", DEFAULTS["different"]))
        settings["choose"] = to_bool(config.get_option("choose", DEFAULTS["choose"]))
        settings["pipe_type"] = config.get_option("pipe_type", DEFAULTS["pipe_type"])
        settings["duct_type"] = config.get_option("duct_type", DEFAULTS["duct_type"])
    except Exception:
        pass
    return settings


def save_settings(config, settings):
    try:
        for key in DEFAULTS:
            setattr(config, key, settings[key])
        script.save_config()
    except Exception:
        pass


def fmt(value):
    return "" if value is None else "{:.0f}".format(value)


def print_report(results, settings):
    output.print_md("# Stub report")
    output.print_md("Length: **{:g} mm**{} | connectors: **{}**".format(
        settings["length_mm"], " + 50 mm per connector" if settings["different"] else "",
        "chosen per object" if settings["choose"] else "all free"))
    created = len([r for r in results if r["new"] is not None])
    output.print_md("Elements picked: **{}** | stubs created: **{}** | not created: **{}**".format(
        len(set(r["element"].Id for r in results)), created, len(results) - created))
    if not results:
        output.print_md("No element was picked.")
        return
    rows = []
    for i, r in enumerate(results, 1):
        rows.append([i, output.linkify(r["element"].Id), r["connector"], fmt(r["length_mm"]),
                     output.linkify(r["new"].Id) if r["new"] is not None else "", r["status"]])
    output.print_table(table_data=rows,
                       columns=["#", "Element", "Connector", "Length [mm]", "New element", "Status"])


def main():
    config = script.get_config(CONFIG_NAME)
    settings = load_settings(config)
    pipe_types, duct_types = get_types(PipeType), get_types(DuctType)

    if __shiftclick__:  # noqa: F821
        # automatico: tutti i connettori liberi, ultima lunghezza + 50 mm per ogni connettore
        settings["different"] = True
        settings["choose"] = False
    else:
        xaml_path = script.get_bundle_file(XAML_FILE_NAME)
        values = Stub_ui.StubWindow(xaml_path, sorted(pipe_types), sorted(duct_types), settings).show()
        if values is None:
            script.exit()
        settings = values
        save_settings(config, settings)

    type_ids = resolve_type_ids(settings, pipe_types, duct_types)
    results = []
    selection_filter = MechanicalConnectorFilter()
    prompt = "Spring ({:g} mm{}): click an element with mechanical connectors (ESC to finish)".format(
        settings["length_mm"], " +50 each" if settings["different"] else "")
    while True:
        try:
            reference = uidoc.Selection.PickObject(ObjectType.Element, selection_filter, prompt)
        except OperationCanceledException:
            break
        element = doc.GetElement(reference.ElementId)
        results.extend(process_element(element, settings, type_ids, settings["length_mm"]))

    problems = len([r for r in results if r["new"] is None])
    if want_report(len(results), problems, "stubs created"):
        print_report(results, settings)


main()
