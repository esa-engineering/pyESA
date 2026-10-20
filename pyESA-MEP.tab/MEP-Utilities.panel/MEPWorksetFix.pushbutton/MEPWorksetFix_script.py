# -*- coding: utf-8 -*-
__title__ = "MEP\nWorkset Fix"

__doc__ = """Version = 1.4
Date    = 20.10.2026
_____________________________________________________________________
Moves MEP elements to the workset of their discipline, chosen in a
dialog for four sections: mechanical, plumbing, fire protection and
electrical.

  Cable Trays, Cable Tray Fittings,
  Conduits, Conduit Fittings,
  Data / Fire Alarm / Lighting /
  Nurse Call / Security / Telephone
  Devices, Electrical Equipment,
  Electrical Fixtures                   -> electrical, always
  Air Terminals, Ducts, Duct Fittings,
  Duct Accessories, Flex Ducts
  (System Type)                         -> DS.FF fire protection
                                           DS.HV mechanical
                                           DS.OT mechanical
  Pipes, Pipe Fittings, Flex Pipes
  (System Type)                         -> PS.FF fire protection
                                           PS.PD mechanical
                                           PS.SS plumbing
                                           PS.GA plumbing
  Sprinklers                            -> fire protection, always

"Contains" match on the System Type name, case-insensitive. Elements
matching no rule, or rules of two different sections, are left where
they are and listed in the report. Elements in the worksets chosen as
"left untouched" (e.g. temporary elements) are never moved. Insulation
and lining follow their host. Model groups are never changed: elements inside a group (or
whose insulation / lining is inside one), elements nested in families,
owned by another user or not updated from central are skipped and
listed.

A summary is shown before writing: answering No prints the same
report as a preview, without touching the model.
_____________________________________________________________________
Author(s): Andrea Patti
"""

__author__ = "Andrea Patti"

from collections import OrderedDict

from System.Collections.Generic import List

from pyrevit import revit, script, DB, forms

import MEPWorksetFix_ui as ui

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

BIC = DB.BuiltInCategory
BIP = DB.BuiltInParameter
ST = DB.StorageType

MAX_DETAIL_ROWS = 300
PROGRESS_STEP = 100

# =============================================================================
# REGOLE
# =============================================================================

MECHANICAL = 'mechanical'
PLUMBING = 'plumbing'
FIRE = 'fire'
ELECTRICAL = 'electrical'

SECTION_LABELS = OrderedDict([
    (MECHANICAL, "Mechanical"),
    (PLUMBING, "Plumbing"),
    (FIRE, "Fire protection"),
    (ELECTRICAL, "Electrical"),
])

# (testo cercato nel nome del System Type, sezione di destinazione)
DUCT_RULES = (
    ("DS.FF", FIRE),
    ("DS.HV", MECHANICAL),
    ("DS.OT", MECHANICAL),
)
PIPE_RULES = (
    ("PS.FF", FIRE),
    ("PS.PD", MECHANICAL),
    ("PS.SS", PLUMBING),
    ("PS.GA", PLUMBING),
)


class CategoryRule(object):
    """Come si decide la sezione di una categoria.

    fixed_section valorizzata: la categoria va sempre li', senza leggere il sistema.
    Altrimenti system_param e domain dicono dove leggere il System Type e rules
    lo traducono in sezione.
    """

    def __init__(self, bic, label, fixed_section=None,
                 system_param=None, domain=None, rules=()):
        self.bic = bic
        self.label = label
        self.fixed_section = fixed_section
        self.system_param = system_param
        self.domain = domain
        self.rules = rules


def _duct_rule(bic, label):
    return CategoryRule(bic, label, system_param=BIP.RBS_DUCT_SYSTEM_TYPE_PARAM,
                        domain=DB.Domain.DomainHvac, rules=DUCT_RULES)


def _pipe_rule(bic, label):
    return CategoryRule(bic, label, system_param=BIP.RBS_PIPING_SYSTEM_TYPE_PARAM,
                        domain=DB.Domain.DomainPiping, rules=PIPE_RULES)


def _electrical_rule(bic, label):
    return CategoryRule(bic, label, fixed_section=ELECTRICAL)


CATEGORY_RULES = (
    _duct_rule(BIC.OST_DuctTerminal, "Air Terminals"),
    _duct_rule(BIC.OST_DuctCurves, "Ducts"),
    _duct_rule(BIC.OST_DuctFitting, "Duct Fittings"),
    _duct_rule(BIC.OST_DuctAccessory, "Duct Accessories"),
    _duct_rule(BIC.OST_FlexDuctCurves, "Flex Ducts"),
    _pipe_rule(BIC.OST_PipeCurves, "Pipes"),
    _pipe_rule(BIC.OST_PipeFitting, "Pipe Fittings"),
    _pipe_rule(BIC.OST_FlexPipeCurves, "Flex Pipes"),
    CategoryRule(BIC.OST_Sprinklers, "Sprinklers", fixed_section=FIRE),
    _electrical_rule(BIC.OST_CableTray, "Cable Trays"),
    _electrical_rule(BIC.OST_CableTrayFitting, "Cable Tray Fittings"),
    _electrical_rule(BIC.OST_Conduit, "Conduits"),
    _electrical_rule(BIC.OST_ConduitFitting, "Conduit Fittings"),
    _electrical_rule(BIC.OST_DataDevices, "Data Devices"),
    _electrical_rule(BIC.OST_ElectricalEquipment, "Electrical Equipment"),
    # Electrical Fixtures e Fire Alarm Devices sono spesso famiglie annidate
    # condivise: quelle istanze si saltano (NESTED), il workset segue l'host.
    _electrical_rule(BIC.OST_ElectricalFixtures, "Electrical Fixtures"),
    _electrical_rule(BIC.OST_FireAlarmDevices, "Fire Alarm Devices"),
    _electrical_rule(BIC.OST_LightingDevices, "Lighting Devices"),
    _electrical_rule(BIC.OST_NurseCallDevices, "Nurse Call Devices"),
    _electrical_rule(BIC.OST_SecurityDevices, "Security Devices"),
    _electrical_rule(BIC.OST_TelephoneDevices, "Telephone Devices"),
)
RULE_BY_CATEGORY = dict((int(rule.bic), rule) for rule in CATEGORY_RULES)

# Esiti. TO_MOVE e' solo il marcatore interno prima della scrittura; gli altri
# finiscono nel report e sono in inglese.
TO_MOVE = "TO MOVE"
MOVED = "MOVED"
ALREADY_OK = "ALREADY CORRECT"
NO_RULE = "NO MATCHING SYSTEM TYPE"
CONFLICT = "CONFLICTING SYSTEM TYPES"
NOT_SET = "SECTION SET TO DO NOT CHANGE"
EXCLUDED_WORKSET = "IN A WORKSET LEFT UNTOUCHED"
IN_GROUP = "IN A GROUP"
NESTED = "NESTED IN A FAMILY"
READ_ONLY = "WORKSET READ-ONLY"
BORROWED = "OWNED BY ANOTHER USER"
NOT_LATEST = "NOT UPDATED FROM CENTRAL"
TARGET_OWNED = "TARGET WORKSET OWNED BY ANOTHER USER"
FAILED = "ERROR"

SKIP_STATUSES = (EXCLUDED_WORKSET, NOT_SET, IN_GROUP, NESTED, READ_ONLY, BORROWED, NOT_LATEST,
                 TARGET_OWNED, FAILED)


# =============================================================================
# UTILITY
# =============================================================================

def get_element_id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


def element_name(element):
    """Nome di un elemento: el.Name fallisce in IronPython su molte sottoclassi."""
    try:
        return DB.Element.Name.GetValue(element)
    except Exception:
        pass
    try:
        return element.Name
    except Exception:
        pass
    try:
        param = element.get_Parameter(BIP.SYMBOL_NAME_PARAM)
        if param is not None and param.AsString():
            return param.AsString()
    except Exception:
        pass
    return "<id {}>".format(get_element_id_value(element.Id))


class WorksetInfo(object):
    """Nomi e proprietari dei workset, letti una volta per esecuzione."""

    def __init__(self, document):
        self.table = document.GetWorksetTable()
        self.user = document.Application.Username
        self._names = {}

    def name(self, workset_int):
        if workset_int not in self._names:
            try:
                self._names[workset_int] = self.table.GetWorkset(
                    DB.WorksetId(workset_int)).Name
            except Exception:
                self._names[workset_int] = "<workset {}>".format(workset_int)
        return self._names[workset_int]

    def other_owner(self, workset_int):
        """Proprietario del workset se e' un altro utente, altrimenti None."""
        try:
            owner = self.table.GetWorkset(DB.WorksetId(workset_int)).Owner
        except Exception:
            return None
        if owner and owner != self.user:
            return owner
        return None


# =============================================================================
# LETTURA DEL SYSTEM TYPE
# =============================================================================

def element_connectors(element):
    manager = getattr(element, "ConnectorManager", None)      # MEPCurve
    if manager is None:
        mep_model = getattr(element, "MEPModel", None)         # FamilyInstance
        if mep_model is not None:
            manager = mep_model.ConnectorManager
    if manager is None:
        return []
    try:
        return list(manager.Connectors)
    except Exception:
        return []


def connector_system_type_names(element, domain):
    """System Type dei sistemi collegati ai connettori del dominio indicato."""
    names = []
    for connector in element_connectors(element):
        try:
            if connector.Domain != domain:
                continue
            system = connector.MEPSystem
        except Exception:
            continue
        if system is None:
            continue
        system_type = doc.GetElement(system.GetTypeId())
        if system_type is None:
            continue
        name = element_name(system_type)
        if name not in names:
            names.append(name)
    return names


def system_type_names(element, rule):
    """Nomi dei System Type dell'istanza: prima il parametro, poi i connettori.

    I raccordi non collegati hanno il parametro vuoto: in quel caso si guarda ai
    sistemi dei connettori, che di solito sono vuoti anch'essi.
    """
    names = []
    param = element.get_Parameter(rule.system_param)
    if param is not None and param.HasValue:
        try:
            if param.StorageType == ST.ElementId:
                system_type = doc.GetElement(param.AsElementId())
                if system_type is not None:
                    names.append(element_name(system_type))
            else:
                text = param.AsValueString() or param.AsString()
                if text:
                    names.append(text)
        except Exception:
            pass
    if not names:
        names = connector_system_type_names(element, rule.domain)
    return names


def resolve_section(element, rule):
    """(sezione o None, testo del System Type, esito se non risolta)."""
    if rule.fixed_section:
        return rule.fixed_section, "-", None

    names = system_type_names(element, rule)
    system_text = ", ".join(names) if names else "<none>"
    matched = []
    for name in names:
        upper = name.upper()
        for pattern, section in rule.rules:
            if pattern.upper() in upper and section not in matched:
                matched.append(section)
    if not matched:
        return None, system_text, NO_RULE
    if len(matched) > 1:
        return None, system_text, CONFLICT
    return matched[0], system_text, None


# =============================================================================
# PIANO
# =============================================================================

class PlanItem(object):
    __slots__ = ('element_id', 'category', 'system', 'section', 'source',
                 'target', 'status', 'note', 'linked_ids')

    def __init__(self, element_id, category):
        self.element_id = element_id
        self.category = category
        self.system = ""
        self.section = None
        self.source = None
        self.target = None
        self.status = None
        self.note = ""
        self.linked_ids = []


def collect_elements(config):
    bics = List[BIC]([rule.bic for rule in CATEGORY_RULES])
    category_filter = DB.ElementMulticategoryFilter(bics)
    if config.scope == ui.SCOPE_SELECTION:
        ids = uidoc.Selection.GetElementIds()
        if ids.Count == 0:
            return []
        collector = DB.FilteredElementCollector(doc, ids)
    elif config.scope == ui.SCOPE_VIEW:
        collector = DB.FilteredElementCollector(doc, doc.ActiveView.Id)
    else:
        collector = DB.FilteredElementCollector(doc)
    elements = collector.WherePasses(category_filter)\
        .WhereElementIsNotElementType()\
        .ToElements()
    # I gruppi non hanno nessuna delle categorie gestite, ma il workset di un
    # gruppo non va mai toccato: la guardia esplicita resta anche se le regole
    # cambiano.
    return [element for element in elements if not isinstance(element, DB.Group)]


def group_of(element):
    """Gruppo di modello che contiene l'elemento, oppure None."""
    group_id = element.GroupId
    if group_id is None or group_id == DB.ElementId.InvalidElementId:
        return None
    return doc.GetElement(group_id)


def insulation_and_lining_ids(element, rule):
    """Isolamenti e rivestimenti dell'elemento.

    Solo canali e tubi ne hanno: le categorie a sezione fissa (elettriche,
    sprinkler) no.
    """
    if rule.fixed_section:
        return []
    ids = []
    try:
        ids.extend(DB.InsulationLiningBase.GetInsulationIds(doc, element.Id))
    except Exception:
        pass
    try:
        ids.extend(DB.InsulationLiningBase.GetLiningIds(doc, element.Id))
    except Exception:
        pass
    return ids


def ownership_problem(element_id):
    """(esito, nota) se l'elemento non puo' essere modificato da questo utente.

    Va chiamata fuori transazione: dentro una transazione il valore e' cache
    locale e non affidabile (documentazione di GetCheckoutStatus).
    """
    try:
        status = DB.WorksharingUtils.GetCheckoutStatus(doc, element_id)
        if status == DB.CheckoutStatus.OwnedByOtherUser:
            owner = ""
            try:
                owner = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, element_id).Owner
            except Exception:
                pass
            return BORROWED, owner
    except Exception:
        pass
    try:
        updates = DB.WorksharingUtils.GetModelUpdatesStatus(doc, element_id)
        if updates in (DB.ModelUpdatesStatus.UpdatedInCentral,
                       DB.ModelUpdatesStatus.DeletedInCentral):
            return NOT_LATEST, "reload latest"
    except Exception:
        pass
    return None, ""


def plan_item(element, config, worksets):
    rule = RULE_BY_CATEGORY.get(int(get_element_id_value(element.Category.Id)))
    item = PlanItem(element.Id, rule.label)
    item.source = element.WorksetId.IntegerValue

    section, item.system, unresolved = resolve_section(element, rule)
    item.section = section
    # Workset da non toccare (elementi temporanei): prima di ogni altra regola.
    if item.source in config.excluded:
        item.status = EXCLUDED_WORKSET
        return item
    if unresolved:
        item.status = unresolved
        return item

    target = config.targets.get(section)
    if target is None:
        item.status = NOT_SET
        return item
    item.target = target
    if item.source == target:
        item.status = ALREADY_OK
        return item

    # Da qui l'elemento andrebbe spostato: si verifica che si possa farlo.
    # I membri dei gruppi di modello si saltano sempre: il workset si cambia
    # solo modificando il gruppo, e lo strumento non tocca i gruppi.
    group = group_of(element)
    if group is not None:
        item.status, item.note = IN_GROUP, element_name(group)
        return item
    # Istanze annidate condivise (es. Electrical Fixtures / Fire Alarm Devices
    # dentro un'altra famiglia): il workset segue la famiglia ospite e non si
    # puo' cambiare. La nota dice quale famiglia le contiene.
    host = getattr(element, "SuperComponent", None)
    if host is not None:
        item.status = NESTED
        item.note = "inside {} ({})".format(element_name(host),
                                            get_element_id_value(host.Id))
        return item
    param = element.get_Parameter(BIP.ELEM_PARTITION_PARAM)
    if param is None or param.IsReadOnly:
        item.status = READ_ONLY
        return item
    status, note = ownership_problem(element.Id)
    if status:
        item.status, item.note = status, note
        return item
    owner = worksets.other_owner(target)
    if owner:
        item.status, item.note = TARGET_OWNED, owner
        return item

    item.linked_ids = insulation_and_lining_ids(element, rule)
    for linked_id in item.linked_ids:
        linked = doc.GetElement(linked_id)
        linked_group = group_of(linked) if linked is not None else None
        if linked_group is not None:
            item.status = IN_GROUP
            item.note = "insulation/lining {} in group {}".format(
                get_element_id_value(linked_id), element_name(linked_group))
            return item
        status, note = ownership_problem(linked_id)
        if status:
            item.status = status
            item.note = "insulation/lining {}{}".format(
                get_element_id_value(linked_id), ", " + note if note else "")
            return item

    item.status = TO_MOVE
    return item


def build_plan(elements, config, worksets):
    plan = []
    total = len(elements)
    with forms.ProgressBar(title="MEP Workset Fix - analysing elements",
                           cancellable=False) as progress:
        for index, element in enumerate(elements):
            try:
                plan.append(plan_item(element, config, worksets))
            except Exception as ex:
                item = PlanItem(element.Id, "?")
                item.status, item.note = FAILED, str(ex)
                plan.append(item)
            if index % PROGRESS_STEP == 0:
                progress.update_progress(index, total)
    return plan


# =============================================================================
# SCRITTURA
# =============================================================================

def apply_plan(plan):
    """Sposta gli elementi TO_MOVE; ritorna (isolamenti spostati, rimasti fuori)."""
    to_move = [item for item in plan if item.status == TO_MOVE]
    total = len(to_move)
    with revit.Transaction("MEP Workset Fix"):
        with forms.ProgressBar(title="MEP Workset Fix - moving elements",
                               cancellable=False) as progress:
            for index, item in enumerate(to_move):
                try:
                    element = doc.GetElement(item.element_id)
                    element.get_Parameter(BIP.ELEM_PARTITION_PARAM).Set(item.target)
                    # Di norma isolamento e rivestimento seguono l'host da soli:
                    # si allineano a mano solo quelli rimasti indietro.
                    for linked_id in item.linked_ids:
                        linked = doc.GetElement(linked_id)
                        if linked is None or linked.WorksetId.IntegerValue == item.target:
                            continue
                        linked_param = linked.get_Parameter(BIP.ELEM_PARTITION_PARAM)
                        if linked_param is not None and not linked_param.IsReadOnly:
                            linked_param.Set(item.target)
                except Exception as ex:
                    item.status, item.note = FAILED, str(ex)
                if index % PROGRESS_STEP == 0:
                    progress.update_progress(index, total)

    # Verifica a transazione chiusa: conta quello che e' davvero cambiato.
    linked_moved = 0
    linked_left = 0
    for item in to_move:
        if item.status != TO_MOVE:
            continue
        element = doc.GetElement(item.element_id)
        if element is not None and element.WorksetId.IntegerValue == item.target:
            item.status = MOVED
        else:
            item.status, item.note = FAILED, "workset unchanged after commit"
            continue
        for linked_id in item.linked_ids:
            linked = doc.GetElement(linked_id)
            if linked is None:
                continue
            if linked.WorksetId.IntegerValue == item.target:
                linked_moved += 1
            else:
                linked_left += 1
    return linked_moved, linked_left


# =============================================================================
# REPORT
# =============================================================================

def select_link(ids):
    return output.linkify(ids, title="{} element(s)".format(len(ids)))


def scope_label(config):
    if config.scope == ui.SCOPE_SELECTION:
        return "current selection"
    if config.scope == ui.SCOPE_VIEW:
        return "elements visible in view '{}'".format(doc.ActiveView.Name)
    return "whole model"


def print_header(config, preview):
    output.print_md("# MEP Workset Fix")
    if preview:
        output.print_md("**Preview only: no changes were made to the model.**")
    output.print_md("Model: **{}** - scope: **{}**".format(doc.Title, scope_label(config)))
    rows = []
    for key, label in SECTION_LABELS.items():
        rows.append([label, config.target_names.get(key) or ui.NOT_SET_LABEL])
    output.print_table(rows, columns=["Section", "Target workset"])
    if config.excluded_names:
        output.print_md("Worksets left untouched: **{}**".format(
            "**, **".join(config.excluded_names)))


def print_report(plan, config, worksets, preview, linked_counts=None):
    output.close_others()
    print_header(config, preview)

    if not plan:
        output.print_md(
            "No element of the handled categories was found in the chosen scope "
            "({}).".format(", ".join(rule.label for rule in CATEGORY_RULES)))
        return

    moved_status = TO_MOVE if preview else MOVED
    moved_label = "To move" if preview else "Moved"

    # Riepilogo per categoria.
    summary = OrderedDict((rule.label, [0, 0, 0, 0, 0]) for rule in CATEGORY_RULES)
    for item in plan:
        counts = summary.setdefault(item.category, [0, 0, 0, 0, 0])
        counts[0] += 1
        if item.status == moved_status:
            counts[1] += 1
        elif item.status == ALREADY_OK:
            counts[2] += 1
        elif item.status in (NO_RULE, CONFLICT):
            counts[3] += 1
        else:
            counts[4] += 1
    rows = [[label] + counts for label, counts in summary.items() if counts[0]]
    totals = [sum(row[i] for row in rows) for i in range(1, 6)]
    rows.append(["Total"] + totals)
    output.print_md("## Summary")
    output.print_table(rows, columns=["Category", "Elements", moved_label,
                                      "Already correct", "No rule matched", "Skipped"])

    # Spostamenti raggruppati.
    groups = OrderedDict()
    for item in plan:
        if item.status != moved_status:
            continue
        key = (SECTION_LABELS[item.section], worksets.name(item.source),
               worksets.name(item.target), item.category)
        groups.setdefault(key, []).append(item.element_id)
    output.print_md("## {}".format(moved_label))
    if groups:
        rows = [[section, source, target, category, select_link(ids)]
                for (section, source, target, category), ids in sorted(groups.items())]
        output.print_table(rows, columns=["Section", "From workset", "To workset",
                                          "Category", "Select"])
    else:
        output.print_md("Nothing to move: every element is already in its workset "
                        "or was left unchanged for one of the reasons below.")
    if linked_counts is not None:
        linked_moved, linked_left = linked_counts
        if linked_moved or linked_left:
            output.print_md("Insulation and lining moved with their host: **{}**.".format(
                linked_moved))
        if linked_left:
            output.print_md(
                "**{}** insulation / lining element(s) are still on another workset: "
                "run *Set Insulation Workset*.".format(linked_left))

    # System Type senza regola, raggruppati per nome.
    unmatched = OrderedDict()
    for item in plan:
        if item.status in (NO_RULE, CONFLICT):
            key = (item.status, item.category, item.system)
            unmatched.setdefault(key, []).append(item.element_id)
    if unmatched:
        output.print_md("## Left unchanged: System Type not recognised")
        output.print_md("Rules: duct elements DS.FF / DS.HV / DS.OT, pipe elements "
                        "PS.FF / PS.PD / PS.SS / PS.GA.")
        rows = [[status, category, system, select_link(ids)]
                for (status, category, system), ids in sorted(unmatched.items())]
        output.print_table(rows, columns=["Reason", "Category", "System Type", "Select"])

    # Elementi saltati, uno per riga.
    skipped = [item for item in plan if item.status in SKIP_STATUSES]
    if skipped:
        output.print_md("## Skipped")
        reasons = OrderedDict()
        for item in skipped:
            reasons.setdefault(item.status, []).append(item.element_id)
        output.print_table([[reason, select_link(ids)] for reason, ids in reasons.items()],
                           columns=["Reason", "Select"])
        rows = []
        for item in skipped[:MAX_DETAIL_ROWS]:
            section = SECTION_LABELS.get(item.section, "-")
            rows.append([output.linkify(item.element_id), item.category, item.system,
                         section, worksets.name(item.source), item.status, item.note])
        output.print_table(rows, columns=["Element", "Category", "System Type", "Section",
                                          "Current workset", "Reason", "Note"])
        if len(skipped) > MAX_DETAIL_ROWS:
            output.print_md("... and {} more (use the Select links above).".format(
                len(skipped) - MAX_DETAIL_ROWS))


# =============================================================================
# MAIN
# =============================================================================

def section_hints():
    """Testo sotto ogni sezione della finestra, ricavato dalle regole."""
    hints = OrderedDict((key, []) for key in SECTION_LABELS)
    fixed = OrderedDict((key, []) for key in SECTION_LABELS)
    by_rules = OrderedDict()
    for rule in CATEGORY_RULES:
        if rule.fixed_section:
            fixed[rule.fixed_section].append(rule.label)
        else:
            by_rules.setdefault(rule.rules, []).append(rule.label)
    for key in SECTION_LABELS:
        if fixed[key]:
            hints[key].append("{}: always.".format(", ".join(fixed[key])))
        for rules, labels in by_rules.items():
            patterns = [pattern for pattern, section in rules if section == key]
            if patterns:
                hints[key].append("{}: System Type containing {}.".format(
                    ", ".join(labels), " or ".join(patterns)))
    return dict((key, "\n".join(lines)) for key, lines in hints.items())


def main():
    if doc is None or doc.IsFamilyDocument:
        forms.alert("Open a project model to run this tool.", title="MEP Workset Fix")
        script.exit()
    if not doc.IsWorkshared:
        forms.alert("The model is not workshared: there are no worksets to assign.",
                    title="MEP Workset Fix")
        script.exit()

    selection_count = 0
    try:
        selection_count = uidoc.Selection.GetElementIds().Count
    except Exception:
        pass

    config = ui.show_config_form(doc, section_hints(), selection_count)
    if config is None:
        script.exit()

    worksets = WorksetInfo(doc)
    elements = collect_elements(config)
    plan = build_plan(elements, config, worksets)

    to_move = [item for item in plan if item.status == TO_MOVE]
    if not to_move:
        print_report(plan, config, worksets, preview=False)
        return

    lines = []
    for key, label in SECTION_LABELS.items():
        count = len([item for item in to_move if item.section == key])
        if count:
            lines.append("  {}  ->  {}: {}".format(label, config.target_names[key], count))
    unchanged = len(plan) - len(to_move)
    message = ("{} of {} elements will be moved:\n\n{}\n\n"
               "{} elements stay where they are (already correct, no rule matched "
               "or skipped: see the report).\n\n"
               "Move them now? Choose No for a preview report only.").format(
        len(to_move), len(plan), "\n".join(lines), unchanged)
    if not forms.alert(message, title="MEP Workset Fix", yes=True, no=True):
        print_report(plan, config, worksets, preview=True)
        return

    linked_counts = apply_plan(plan)
    print_report(plan, config, worksets, preview=False, linked_counts=linked_counts)


main()
