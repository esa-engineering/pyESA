# -*- coding: utf-8 -*-
__title__ = "Elements\nat Level"

__doc__ = """Version = 2.1
Date    = 10.09.2026
_____________________________________________________________________
CLICK
Reassigns the host level of the selected elements while keeping them
in their current position (the offset is recalculated). Clearing the
option makes the elements move together with the new level.

The selection is free: you no longer pick a category, the script
recognises the category of each element and applies the right method.
For dual-level elements (walls, columns, stairs, roofs) base level and
top level can be assigned in a single run, even to two different
levels.

SHIFT + CLICK
The elements keep their level: the levels are moved instead. After the
element selection, pick the levels in a section / elevation / 3D view
and enter an offset (mm, + or -) for each of them. The selected
elements hosted on the moved levels get their offset recalculated and
stay where they are; every other element follows its level.
A selection made before launching the command is used as is: the
levels it contains are the levels to move, the rest are the elements.
_____________________________________________________________________
Author(s): bimdifferent, ESA Engineering
"""

__author__ = "bimdifferent, ESA Engineering"

from collections import OrderedDict

from pyrevit import revit, script, DB, UI, forms

from rpw.ui.forms import CheckBox, FlexForm, Label, Separator, Button, ComboBox, TextBox

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

BIC = DB.BuiltInCategory
BIP = DB.BuiltInParameter

MAX_SKIPPED_ROWS = 200
FEET_TO_MM = 304.8
# La finestra rpw usa SizeToContent: si dimensiona sul controllo piu' largo.
# Le Label WPF non vanno a capo: il testo oltre la larghezza viene tagliato.
FORM_WIDTH = 270
MOVE_FORM_WIDTH = 380   # form Move Levels: righe 'nome (quota) - n selected' piu' lunghe


# ---------------------------------------------------------------- helpers

def element_id_value(eid):
    """ElementId.IntegerValue e' stato rimosso in Revit 2026 (sostituito da .Value)."""
    if hasattr(eid, "Value"):
        return eid.Value
    return eid.IntegerValue


def category_id(bic_name):
    """Id intero di una BuiltInCategory, None se l'enum non esiste in questa versione."""
    member = getattr(BIC, bic_name, None)
    if member is None:
        return None
    return int(member)


def find_param(element, bip_names, param_names=None, writable=True):
    """Primo parametro esistente (e scrivibile) fra i BIP indicati, poi fra i nomi.

    I BuiltInParameter sono referenziati per nome: alcuni enum non esistono in
    tutte le versioni di Revit e un attributo mancante farebbe fallire l'import.
    """
    for name in bip_names or []:
        bip = getattr(BIP, name, None)
        if bip is None:
            continue
        try:
            param = element.get_Parameter(bip)
        except Exception:
            param = None
        if param is not None and (not writable or not param.IsReadOnly):
            return param
    for name in param_names or []:
        try:
            param = element.LookupParameter(name)
        except Exception:
            param = None
        if param is not None and (not writable or not param.IsReadOnly):
            return param
    return None


def format_elevation(value_feet):
    return u'{:+.0f} mm'.format(value_feet * FEET_TO_MM)


def separator(width=FORM_WIDTH):
    """rpw Separator e' l'unico controllo senza __init__ proprio: i kwargs finiscono
    a object.__new__ e sollevano TypeError. I valori si impostano dopo la costruzione,
    altrimenti restano i default di rpw (300 x 25): 300 e' piu' largo di FORM_WIDTH e
    la finestra, che usa SizeToContent, si allargherebbe sul separatore."""
    control = Separator()
    control.Width = width
    control.Height = 5
    return control


# ------------------------------------------------------------ regole livelli

def P(level_bip, offset_bips, mode='first'):
    """Coppia livello/offset. mode 'all' applica il delta a tutti gli offset."""
    return (level_bip, offset_bips, mode)


def S(role, pairs, special=None, lnames=None, onames=None):
    """Slot di livello: 'base' o 'top'."""
    return {'role': role, 'pairs': pairs, 'special': special,
            'lnames': lnames, 'onames': onames}


LEVEL_NAMES = [u'Level', u'Livello', u'Reference Level', u'Livello di riferimento',
               u'Base Level', u'Livello di base', u'Schedule Level', u'Livello di abaco']
OFFSET_NAMES = [u'Offset from Level', u'Elevation from Level', u'Quota dal livello',
                u'Offset dal livello', u'Offset', u'Elevazione dal livello']

# Famiglie posizionate su un livello: parametro Livello + Quota dal livello.
FAMILY_PAIRS = [
    P('FAMILY_LEVEL_PARAM', ['INSTANCE_ELEVATION_PARAM', 'INSTANCE_FREE_HOST_OFFSET_PARAM']),
    P('SCHEDULE_LEVEL_PARAM', ['INSTANCE_FREE_HOST_OFFSET_PARAM', 'INSTANCE_ELEVATION_PARAM']),
    P('FAMILY_BASE_LEVEL_PARAM', ['FAMILY_BASE_LEVEL_OFFSET_PARAM']),
]

FAMILY_CATS = [
    'OST_DuctTerminal', 'OST_DuctFitting', 'OST_CommunicationDevices',
    'OST_DataDevices', 'OST_DuctAccessory', 'OST_ElectricalEquipment',
    'OST_ElectricalFixtures', 'OST_FireAlarmDevices', 'OST_LightingDevices',
    'OST_LightingFixtures', 'OST_MechanicalEquipment', 'OST_NurseCallDevices',
    'OST_PipeAccessory', 'OST_PipeFitting', 'OST_PlumbingFixtures',
    'OST_SecurityDevices', 'OST_Sprinklers', 'OST_TelephoneDevices',
    'OST_Furniture', 'OST_Casework', 'OST_Doors', 'OST_Windows',
    'OST_GenericModel', 'OST_SpecialityEquipment',
]

MEP_CURVE_CATS = [
    'OST_DuctCurves', 'OST_FlexDuctCurves', 'OST_PipeCurves',
    'OST_FlexPipeCurves', 'OST_CableTray', 'OST_Conduit',
]

RULES = {}


def add_rule(bic_names, slots):
    for bic_name in bic_names:
        cid = category_id(bic_name)
        if cid is not None:
            RULES[cid] = slots


add_rule(FAMILY_CATS, [
    S('base', FAMILY_PAIRS, lnames=LEVEL_NAMES, onames=OFFSET_NAMES),
])

add_rule(MEP_CURVE_CATS, [
    S('base', [P('RBS_START_LEVEL_PARAM', ['RBS_OFFSET_PARAM'])],
      lnames=[u'Reference Level', u'Livello di riferimento'],
      onames=[u'Offset', u'Offset from Level', u'Quota dal livello']),
])

add_rule(['OST_Walls'], [
    S('base', [P('WALL_BASE_CONSTRAINT', ['WALL_BASE_OFFSET']),
               P('FACEROOF_LEVEL_PARAM', ['FACEROOF_OFFSET_PARAM'])]),
    S('top', [P('WALL_HEIGHT_TYPE', ['WALL_TOP_OFFSET'])], special='wall_top'),
])

add_rule(['OST_StructuralColumns', 'OST_Columns'], [
    S('base', [P('FAMILY_BASE_LEVEL_PARAM', ['FAMILY_BASE_LEVEL_OFFSET_PARAM'])]),
    S('top', [P('FAMILY_TOP_LEVEL_PARAM', ['FAMILY_TOP_LEVEL_OFFSET_PARAM'])]),
])

add_rule(['OST_Floors'], [
    S('base', [P('LEVEL_PARAM', ['FLOOR_HEIGHTABOVELEVEL_PARAM'])]),
])

add_rule(['OST_Ceilings'], [
    S('base', [P('LEVEL_PARAM', ['CEILING_HEIGHTABOVELEVEL_PARAM'])]),
])

add_rule(['OST_Roofs'], [
    S('base', [P('ROOF_BASE_LEVEL_PARAM', ['ROOF_LEVEL_OFFSET_PARAM']),
               P('FACEROOF_LEVEL_PARAM', ['FACEROOF_OFFSET_PARAM'])]),
    S('top', [P('ROOF_CONSTRAINT_LEVEL_PARAM', ['ROOF_CONSTRAINT_OFFSET_PARAM'])]),
])

add_rule(['OST_Stairs'], [
    S('base', [P('STAIRS_BASE_LEVEL_PARAM', ['STAIRS_BASE_OFFSET'])],
      lnames=[u'Base Level', u'Livello di base'],
      onames=[u'Base Offset', u'Offset di base', u'Offset livello di base']),
    S('top', [P('STAIRS_TOP_LEVEL_PARAM', ['STAIRS_TOP_OFFSET'])],
      lnames=[u'Top Level', u'Livello superiore'],
      onames=[u'Top Offset', u'Offset superiore', u'Offset livello superiore']),
])

add_rule(['OST_StairsRailing', 'OST_Railings'], [
    S('base', [P('STAIRS_RAILING_BASE_LEVEL_PARAM', ['STAIRS_RAILING_HEIGHT_OFFSET'])]
      + FAMILY_PAIRS,
      lnames=[u'Base Level', u'Livello di base'] + LEVEL_NAMES,
      onames=[u'Base Offset', u'Offset di base'] + OFFSET_NAMES),
])

# Travi: un solo livello di riferimento, due offset di estremita' da traslare insieme.
add_rule(['OST_StructuralFraming'], [
    S('base', [P('INSTANCE_REFERENCE_LEVEL_PARAM',
                 ['STRUCTURAL_BEAM_END0_ELEVATION', 'STRUCTURAL_BEAM_END1_ELEVATION'],
                 mode='all')]
      + FAMILY_PAIRS,
      lnames=[u'Reference Level', u'Livello di riferimento'],
      onames=[u'Start Level Offset', u'End Level Offset',
              u'Offset livello iniziale', u'Offset livello finale', u'z Offset Value']),
])

add_rule(['OST_StructuralTruss'], [
    S('base', [P('TRUSS_ELEMENT_REFERENCE_LEVEL_PARAM',
                 ['TRUSS_ELEMENT_REFERENCE_LEVEL_ELEVATION', 'Z_OFFSET_VALUE']),
               P('INSTANCE_REFERENCE_LEVEL_PARAM',
                 ['INSTANCE_ELEVATION_PARAM', 'Z_OFFSET_VALUE'])]
      + FAMILY_PAIRS,
      lnames=[u'Reference Level', u'Livello di riferimento'],
      onames=OFFSET_NAMES),
])

add_rule(['OST_Cornices', 'OST_Reveals'], [
    S('base', FAMILY_PAIRS, lnames=LEVEL_NAMES, onames=OFFSET_NAMES),
])

GENERIC_SLOTS = [
    S('base', FAMILY_PAIRS
      + [P('INSTANCE_REFERENCE_LEVEL_PARAM', ['INSTANCE_ELEVATION_PARAM', 'Z_OFFSET_VALUE']),
         P('LEVEL_PARAM', ['INSTANCE_ELEVATION_PARAM'])],
      lnames=LEVEL_NAMES, onames=OFFSET_NAMES),
]

# Categorie i cui elementi non hanno un livello proprio: si risale all'host.
REDIRECT_CATS = set()
for _name in ['OST_StairsRuns', 'OST_StairsLandings', 'OST_StairsSupports',
              'OST_StairsTrisers', 'OST_StairsRailing', 'OST_Railings',
              'OST_Cornices', 'OST_Reveals']:
    _cid = category_id(_name)
    if _cid is not None:
        REDIRECT_CATS.add(_cid)


# --------------------------------------------------------------- selezione

def host_of(element):
    """Elemento padre di un sotto-elemento (rampa/pianerottolo, ringhiera, sweep)."""
    stairs_id = getattr(element, 'StairsId', None)
    if isinstance(stairs_id, DB.ElementId) and element_id_value(stairs_id) > 0:
        return doc.GetElement(stairs_id)

    if getattr(element, 'HasHost', False):
        host_id = getattr(element, 'HostId', None)
        if isinstance(host_id, DB.ElementId) and element_id_value(host_id) > 0:
            return doc.GetElement(host_id)

    get_hosts = getattr(element, 'GetHostIds', None)
    if get_hosts is not None:
        try:
            host_ids = list(get_hosts())
        except Exception:
            host_ids = []
        for host_id in host_ids:
            if element_id_value(host_id) > 0:
                return doc.GetElement(host_id)
    return None


def preselected_elements():
    """Elementi (non tipi) gia' selezionati prima del lancio del comando."""
    elements = []
    try:
        pre_ids = list(uidoc.Selection.GetElementIds())
    except Exception:
        pre_ids = []
    for eid in pre_ids:
        element = doc.GetElement(eid)
        if element is not None and not isinstance(element, DB.ElementType):
            elements.append(element)
    return elements


def pick_elements():
    """Selezione a video degli elementi. Esce se l'utente annulla."""
    elements = []
    with forms.WarningBar(title='Select the elements to reassign, then press Finish'):
        try:
            references = list(uidoc.Selection.PickObjects(
                UI.Selection.ObjectType.Element, 'Select the elements'))
        except Exception:
            references = []
    if not references:
        script.exit()
    for reference in references:
        element = doc.GetElement(reference.ElementId)
        if element is not None and not isinstance(element, DB.ElementType):
            elements.append(element)
    return elements


def resolve_targets(elements):
    """Risale agli host dove serve e deduplica. Ritorna (elementi, n_redirect)."""
    resolved = OrderedDict()
    redirected = 0
    for element in elements:
        target = element
        try:
            cid = element_id_value(element.Category.Id)
        except Exception:
            cid = None
        if cid in REDIRECT_CATS:
            host = host_of(element)
            if host is not None:
                target = host
                redirected += 1
        key = element_id_value(target.Id)
        if key not in resolved:
            resolved[key] = target
    return list(resolved.values()), redirected


# ---------------------------------------------------------------- modifica

APPLIED = 'applied'
ALREADY = 'already'
SKIPPED = 'skipped'


def resolve_pair(element, pair, slot, level_writable=True):
    """Parametro livello + parametri offset per una coppia di regole.

    level_writable=False accetta anche un parametro livello in sola lettura: serve
    alla modalita' Move Levels, che il livello non lo scrive mai.
    """
    level_bip, offset_bips, mode = pair
    level_param = find_param(element, [level_bip], writable=level_writable)
    if level_param is None or level_param.StorageType != DB.StorageType.ElementId:
        return None, []
    offset_params = []
    for offset_bip in offset_bips:
        offset_param = find_param(element, [offset_bip])
        if offset_param is not None and offset_param.StorageType == DB.StorageType.Double:
            offset_params.append(offset_param)
            if mode != 'all':
                break
    if not offset_params and slot.get('onames'):
        offset_param = find_param(element, [], slot['onames'])
        if offset_param is not None and offset_param.StorageType == DB.StorageType.Double:
            offset_params.append(offset_param)
    return level_param, offset_params


def resolve_by_name(element, slot, level_writable=True):
    """Ultima spiaggia: parametri cercati per nome (IT/EN)."""
    if not slot.get('lnames'):
        return None, []
    level_param = find_param(element, [], slot['lnames'], writable=level_writable)
    if level_param is None or level_param.StorageType != DB.StorageType.ElementId:
        return None, []
    offset_params = []
    offset_param = find_param(element, [], slot.get('onames'))
    if offset_param is not None and offset_param.StorageType == DB.StorageType.Double:
        offset_params.append(offset_param)
    return level_param, offset_params


def apply_wall_top_unconnected(element, level_param, offset_params, target_level, keep):
    """Muro con top 'Unconnected': la quota del top si ricava da base + altezza."""
    if keep and not offset_params:
        return SKIPPED, u'top offset is not writable'
    base_param = find_param(element, ['WALL_BASE_CONSTRAINT'], writable=False)
    base_level = None
    if base_param is not None:
        base_level = doc.GetElement(base_param.AsElementId())
    if keep and base_level is None:
        return SKIPPED, u'base level cannot be resolved'
    if keep:
        base_offset = find_param(element, ['WALL_BASE_OFFSET'], writable=False)
        height = find_param(element, ['WALL_USER_HEIGHT_PARAM'], writable=False)
        if height is None:
            return SKIPPED, u'wall height cannot be read'
        top_elevation = base_level.Elevation + height.AsDouble()
        if base_offset is not None:
            top_elevation += base_offset.AsDouble()
    level_param.Set(target_level.Id)
    if keep:
        offset_params[0].Set(top_elevation - target_level.Elevation)
    return APPLIED, u''


def slot_candidates(element, slot, level_writable=True):
    """Coppie (parametro livello, parametri offset) di uno slot, in ordine di regola."""
    candidates = []
    for pair in slot['pairs']:
        level_param, offset_params = resolve_pair(element, pair, slot, level_writable)
        if level_param is not None:
            candidates.append((level_param, offset_params))
    by_name = resolve_by_name(element, slot, level_writable)
    if by_name[0] is not None:
        candidates.append(by_name)
    return candidates


def current_slot_level(element, slot):
    """Livello attuale di uno slot e relativi parametri offset, in sola lettura.

    Vince il primo candidato con un livello assegnato, come in apply_slot: cosi'
    un parametro offset condiviso da due coppie (INSTANCE_ELEVATION_PARAM compare
    due volte in FAMILY_PAIRS) non viene compensato due volte.
    Ritorna (Level, offset_params) oppure None (slot assente, livello non
    assegnato, muro con top 'Unconnected').
    """
    for level_param, offset_params in slot_candidates(element, slot, level_writable=False):
        level = doc.GetElement(level_param.AsElementId())
        if isinstance(level, DB.Level):
            return level, offset_params
    return None


def apply_slot(element, slot, target_level, keep):
    """Applica un livello a uno slot dell'elemento. Ritorna (stato, motivo)."""
    reason = u'level parameter missing or read-only'
    candidates = slot_candidates(element, slot)

    for level_param, offset_params in candidates:
        current_level = doc.GetElement(level_param.AsElementId())

        if current_level is None:
            if slot.get('special') == 'wall_top':
                return apply_wall_top_unconnected(
                    element, level_param, offset_params, target_level, keep)
            reason = u'current level not assigned'
            continue

        if element_id_value(current_level.Id) == element_id_value(target_level.Id):
            return ALREADY, u''

        if keep and not offset_params:
            reason = u'no writable offset parameter'
            continue

        delta = current_level.Elevation - target_level.Elevation
        current_offsets = [param.AsDouble() for param in offset_params]
        level_param.Set(target_level.Id)
        if keep:
            for param, value in zip(offset_params, current_offsets):
                param.Set(value + delta)
        return APPLIED, u''

    return SKIPPED, reason


def slots_for(element):
    try:
        cid = element_id_value(element.Category.Id)
    except Exception:
        return None
    return RULES.get(cid, GENERIC_SLOTS)


def category_name(element):
    try:
        return element.Category.Name
    except Exception:
        return u'-'


def bump(stats, name, index):
    """Incrementa la colonna index della riga di categoria name."""
    row = stats.get(name)
    if row is None:
        row = [0, 0, 0]
        stats[name] = row
    row[index] += 1


def print_skipped(skipped):
    if not skipped:
        return
    output.print_md(u'## Skipped elements ({})'.format(len(skipped)))
    shown = skipped[:MAX_SKIPPED_ROWS]
    detail = [[output.linkify(element.Id), category_name(element), reason]
              for element, reason in shown]
    output.print_table(table_data=detail, title='',
                       columns=[u'Element', u'Category', u'Reason'])
    remaining = len(skipped) - len(shown)
    if remaining > 0:
        output.print_md(u'_...and {} more skipped elements not listed._'.format(remaining))


# ------------------------------------------------- SHIFT + CLICK: move levels

LEVEL_PICK_VIEWS = (DB.ViewType.Section, DB.ViewType.Elevation, DB.ViewType.ThreeD)


class FailureCollector(DB.IFailuresPreprocessor):
    """Raccoglie i failure di Revit al commit.

    Gli avvisi restano (Continue) e vengono elencati nel report; gli errori
    mandano la transazione in rollback (ProceedWithRollBack) e il loro testo
    finisce nel report invece di sparire in un rollback silenzioso.
    """

    def __init__(self):
        self.warnings = []
        self.errors = []

    def PreprocessFailures(self, accessor):
        for message in accessor.GetFailureMessages():
            try:
                text = message.GetDescriptionText()
            except Exception:
                text = u'(no description)'
            if message.GetSeverity() == DB.FailureSeverity.Warning:
                if text not in self.warnings:
                    self.warnings.append(text)
            elif text not in self.errors:
                self.errors.append(text)
        if self.errors:
            return DB.FailureProcessingResult.ProceedWithRollBack
        return DB.FailureProcessingResult.Continue


class LevelSelectionFilter(UI.Selection.ISelectionFilter):
    """Filtro di selezione: accetta solo livelli."""

    def AllowElement(self, element):
        return isinstance(element, DB.Level)

    def AllowReference(self, reference, point):
        return False


def pick_levels():
    """Livelli scelti a video, dal piu' alto al piu' basso. Esce se l'utente annulla.

    I livelli sono selezionabili solo in sezione, prospetto e 3D: in pianta il
    filtro non accetterebbe mai nulla e l'utente resterebbe bloccato nel pick.
    """
    try:
        view_type = doc.ActiveView.ViewType
    except Exception:
        view_type = None
    if view_type not in LEVEL_PICK_VIEWS:
        forms.alert(u'Levels can be picked only in a section, elevation or 3D view.',
                    title=u'Elements at Level', exitscript=True)

    with forms.WarningBar(title='Select the levels to move, then press Finish'):
        try:
            references = list(uidoc.Selection.PickObjects(
                UI.Selection.ObjectType.Element, LevelSelectionFilter(),
                'Select the levels to move'))
        except Exception:
            references = []

    picked = OrderedDict()
    for reference in references:
        level = doc.GetElement(reference.ElementId)
        if isinstance(level, DB.Level):
            picked[element_id_value(level.Id)] = level
    if not picked:
        script.exit()
    return sorted(picked.values(), key=lambda lev: lev.Elevation, reverse=True)


def parse_mm(text):
    """float da un testo utente (virgola ammessa, vuoto = 0); None se non e' un numero."""
    cleaned = (text or u'').strip().replace(u',', u'.')
    if not cleaned:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return None


def ask_level_offsets(picked_levels, usage, n_elements):
    """FlexForm con un campo offset (mm) per ogni livello scelto.

    Ritorna la lista [(Level, mm)] dei soli offset diversi da zero; esce se
    l'utente annulla. FlexForm non ha un aggancio di validazione e non si puo'
    riaprire una volta chiusa: in caso di errore si costruisce una nuova finestra
    con i testi gia' digitati.
    """
    typed = [u'0'] * len(picked_levels)
    while True:
        components = [
            Label('Offset per level in mm: + up, - down.', Width=MOVE_FORM_WIDTH),
            Label('0 = level not moved.', Width=MOVE_FORM_WIDTH),
            separator(MOVE_FORM_WIDTH),
        ]
        for index, level in enumerate(picked_levels):
            count = usage.get(element_id_value(level.Id), 0)
            hosted = u'{} selected'.format(count) if count else u'none selected'
            components.append(Label(u'{}   ({})   -   {}'.format(
                level.Name, format_elevation(level.Elevation), hosted),
                Width=MOVE_FORM_WIDTH))
            components.append(TextBox('lvl_{}'.format(index), default=typed[index],
                                      Width=MOVE_FORM_WIDTH))
        components.append(separator(MOVE_FORM_WIDTH))
        components.append(Button('OK', Width=MOVE_FORM_WIDTH))

        flex_form = FlexForm(u'Move Levels  -  {} levels, {} elements'.format(
            len(picked_levels), n_elements), components)
        flex_form.show()
        if not flex_form.values:
            script.exit()

        typed = [flex_form.values['lvl_{}'.format(index)]
                 for index in range(len(picked_levels))]
        offsets = []
        error = None
        for level, text in zip(picked_levels, typed):
            value = parse_mm(text)
            if value is None:
                error = u"The offset of level '{}' is not a number.".format(level.Name)
                break
            if value:
                offsets.append((level, value))
        if error is None and not offsets:
            error = u'Enter a non-zero offset for at least one level.'
        if error is None:
            return offsets
        forms.alert(error, title=u'Elements at Level')


def run_move_levels(preselected):
    """Modalita' SHIFT + CLICK: sposta i livelli scelti e compensa gli offset degli
    elementi selezionati perche' restino fermi. Il resto del modello segue i livelli.

    La preselezione vale per entrambi: i livelli che contiene sono i livelli da
    spostare, il resto sono gli elementi. Cio' che manca viene chiesto a video.
    """
    picked_levels = [element for element in preselected if isinstance(element, DB.Level)]
    elements = [element for element in preselected if not isinstance(element, DB.Level)]
    if not elements:
        elements = pick_elements()
    elements, redirected = resolve_targets(elements)

    if picked_levels:
        picked_levels.sort(key=lambda lev: lev.Elevation, reverse=True)
    else:
        picked_levels = pick_levels()

    # Dry run, prima di toccare il modello: livello attuale di ogni slot.
    found_slots = []    # (elemento, [(ruolo, id livello, parametri offset)])
    usage = {}          # id livello -> n. elementi selezionati ospitati
    for element in elements:
        slots_found = []
        seen = set()
        for slot in slots_for(element) or []:
            try:
                found = current_slot_level(element, slot)
            except Exception:
                found = None
            if found is None:
                continue
            level, offset_params = found
            lid = element_id_value(level.Id)
            slots_found.append((slot['role'], lid, offset_params))
            if lid not in seen:
                seen.add(lid)
                usage[lid] = usage.get(lid, 0) + 1
        found_slots.append((element, slots_found))

    offsets = ask_level_offsets(picked_levels, usage, len(elements))

    moved = OrderedDict()   # id livello -> [Level, delta piedi, mm, quota iniziale]
    for level, mm in offsets:
        moved[element_id_value(level.Id)] = [level, mm / FEET_TO_MM, mm, level.Elevation]

    stats = OrderedDict()   # categoria -> [base, top, non interessati]
    skipped = []            # (elemento, motivo)
    writes = []             # (elemento, categoria, ruolo, [(parametro, nuovo valore)])
    for element, slots_found in found_slots:
        name = category_name(element)
        affected = False
        for role, lid, offset_params in slots_found:
            if lid not in moved:
                continue
            affected = True
            if not offset_params:
                skipped.append((element, u'{}: no writable offset parameter, '
                                         u'the element will move with the level'.format(role)))
                continue
            delta_ft = moved[lid][1]
            try:
                values = [(param, param.AsDouble() - delta_ft) for param in offset_params]
            except Exception as error:
                skipped.append((element, u'{}: Revit error: {}'.format(role, error)))
                continue
            writes.append((element, name, role, values))
        if not affected:
            bump(stats, name, 2)

    # Transazione esplicita, non revit.Transaction: quella di pyRevit ignora lo
    # stato del Commit, e con la gestione non modale dei failure un errore di
    # Revit manda la transazione in rollback in silenzio (il livello "torna
    # indietro" a comando finito). Qui gli errori vengono raccolti e riportati.
    # Prima i livelli, poi gli offset. I livelli pinnati vengono sbloccati per
    # lo spostamento e ripinnati nella stessa transazione.
    failures = FailureCollector()
    transaction = DB.Transaction(doc, 'ElementsAtLevel - Move Levels')
    options = transaction.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(failures)
    options.SetForcedModalHandling(False)
    transaction.SetFailureHandlingOptions(options)

    unpinned = []
    failing_level = None
    status = None
    transaction.Start()
    try:
        for level, delta_ft, _, _ in moved.values():
            failing_level = level
            if level.Pinned:
                level.Pinned = False
                unpinned.append(level)
            level.Elevation = level.Elevation + delta_ft
        failing_level = None
        for element, name, role, values in writes:
            try:
                for param, value in values:
                    param.Set(value)
                bump(stats, name, 0 if role == 'base' else 1)
            except Exception as error:
                skipped.append((element, u'{}: Revit error: {}'.format(role, error)))
        for level in unpinned:
            level.Pinned = True
        status = transaction.Commit()
    except Exception as error:
        transaction.RollBack()
        what = (u"Level '{}' could not be moved".format(failing_level.Name)
                if failing_level is not None else u'Revit error')
        forms.alert(u'{}: {}\n\nNothing was changed.'.format(what, error),
                    title=u'Elements at Level', exitscript=True)
    finally:
        transaction.Dispose()

    committed = status == DB.TransactionStatus.Committed and not failures.errors

    # Report
    output.print_md(u'# Elements at Level - Move Levels')
    if not committed:
        output.print_md(u'## Revit rejected the changes - nothing was modified')
        output.print_md(u'Transaction status: **{}**'.format(status))
        for text in failures.errors:
            output.print_md(u'- {}'.format(text))
        forms.alert(u'Revit rejected the changes and rolled back the transaction.\n'
                    u'See the output panel for the Revit error messages.',
                    title=u'Elements at Level')

    output.print_md(u'Processed elements: **{}**{}  \n'
                    u'Moved levels: **{}**'.format(
                        len(elements),
                        u' ({} of them redirected to their host)'.format(redirected) if redirected else u'',
                        len(moved) if committed else 0))

    # Quote rilette dal modello dopo il commit, non calcolate: cosi' il report
    # dice la verita' anche se Revit ha annullato lo spostamento.
    level_rows = [[level.Name, format_elevation(old_elevation),
                   format_elevation(level.Elevation), format_elevation(delta_ft),
                   u'yes' if level in unpinned else u'']
                  for level, delta_ft, _, old_elevation in moved.values()]
    output.print_table(table_data=level_rows, title='',
                       columns=[u'Level', u'Old elevation', u'Current elevation',
                                u'Requested offset', u'Was pinned'])

    if failures.warnings:
        output.print_md(u'## Revit warnings ({})'.format(len(failures.warnings)))
        for text in failures.warnings:
            output.print_md(u'- {}'.format(text))

    if not committed:
        return

    rows = []
    totals = [0, 0, 0]
    for name, row in stats.items():
        rows.append([name, row[0], row[1], row[2]])
        for i in range(3):
            totals[i] += row[i]
    rows.append([u'**Total**', totals[0], totals[1], totals[2]])
    output.print_table(table_data=rows, title='',
                       columns=[u'Category', u'Base offset recalculated',
                                u'Top offset recalculated', u'Not affected'])

    print_skipped(skipped)


# ------------------------------------------------------------------- input

levels = list(DB.FilteredElementCollector(doc)
              .OfClass(DB.Level)
              .WhereElementIsNotElementType()
              .ToElements())
if not levels:
    forms.alert(u'The model does not contain any level.', title=u'Elements at Level', exitscript=True)
levels.sort(key=lambda lev: lev.Elevation)

# I livelli in elenco vanno dal piu' alto al piu' basso: OrderedDict + sort=False
# perche' rpw ordina alfabeticamente le chiavi di un dizionario normale.
levels_dict = OrderedDict()
for level in reversed(levels):
    levels_dict[u'{}   ({})'.format(level.Name, format_elevation(level.Elevation))] = level

preselected = preselected_elements()

if __shiftclick__:  # noqa: F821
    run_move_levels(preselected)
    script.exit()

elements = preselected or pick_elements()
if not elements:
    script.exit()

KEEP_TOOLTIP = (u"Checked: the offset is recalculated, the elements do not move.\n"
                u"Unchecked: the offset stays unchanged and the elements move "
                u"together with the new level.")

components = [
    Label('Base / reference level', Width=FORM_WIDTH),
    CheckBox('ckb_base', 'Apply the base level', default=True, Width=FORM_WIDTH),
    ComboBox('cmb_base', levels_dict, sort=False, Width=FORM_WIDTH),
    separator(),
    Label('Top level (dual-level elements only)', Width=FORM_WIDTH),
    CheckBox('ckb_top', 'Apply the top level', default=False, Width=FORM_WIDTH),
    ComboBox('cmb_top', levels_dict, sort=False, Width=FORM_WIDTH),
    separator(),
    CheckBox('ckb_keep', 'Keep the elements in their current position',
             default=True, Width=FORM_WIDTH, ToolTip=KEEP_TOOLTIP),
    separator(),
    Button('OK', Width=FORM_WIDTH),
]
flex_form = FlexForm('Elements at Level  -  {} elements'.format(len(elements)), components)
flex_form.show()
if not flex_form.values:
    script.exit()

do_base = bool(flex_form.values['ckb_base'])
do_top = bool(flex_form.values['ckb_top'])
keep_elevation = bool(flex_form.values['ckb_keep'])
base_level = flex_form.values['cmb_base']
top_level = flex_form.values['cmb_top']

if not do_base and not do_top:
    forms.alert(u'Select at least one between base level and top level.',
                title=u'Elements at Level', exitscript=True)


# ------------------------------------------------------------------ lavoro

elements, redirected = resolve_targets(elements)

stats = OrderedDict()   # categoria -> [base, top, gia' corretti]
skipped = []            # (elemento, motivo)

with revit.Transaction('ElementsAtLevel'):
    for element in elements:
        name = category_name(element)
        slots = slots_for(element)
        reasons = []
        touched = False
        already = False

        for slot in slots:
            if slot['role'] == 'base':
                if not do_base:
                    continue
                target_level = base_level
            else:
                if not do_top:
                    continue
                target_level = top_level

            try:
                status, reason = apply_slot(element, slot, target_level, keep_elevation)
            except Exception as error:
                status, reason = SKIPPED, u'Revit error: {}'.format(error)

            if status == APPLIED:
                touched = True
                bump(stats, name, 0 if slot['role'] == 'base' else 1)
            elif status == ALREADY:
                already = True
            else:
                reasons.append(u'{}: {}'.format(slot['role'], reason))

        if touched:
            if reasons:
                skipped.append((element, u'partial - {}'.format(u'; '.join(reasons))))
            continue
        if already and not reasons:
            bump(stats, name, 2)
            continue
        if not reasons:
            reasons.append(u'no level applicable with the selected options')
        skipped.append((element, u'; '.join(reasons)))


# ------------------------------------------------------------------ report

output.print_md(u'# Elements at Level')
output.print_md(u'Processed elements: **{}**{}  \n'
                u'Base level: **{}**  \n'
                u'Top level: **{}**  \n'
                u'Elements kept in their current position: **{}**'.format(
                    len(elements),
                    u' ({} of them redirected to their host)'.format(redirected) if redirected else u'',
                    base_level.Name if do_base else u'not applied',
                    top_level.Name if do_top else u'not applied',
                    u'yes' if keep_elevation else u'no'))

if stats:
    rows = []
    totals = [0, 0, 0]
    for name, row in stats.items():
        rows.append([name, row[0], row[1], row[2]])
        for i in range(3):
            totals[i] += row[i]
    rows.append([u'**Total**', totals[0], totals[1], totals[2]])
    output.print_table(table_data=rows, title='',
                       columns=[u'Category', u'Base level', u'Top level',
                                u'Already correct'])
else:
    output.print_md(u'_No element modified._')

print_skipped(skipped)
