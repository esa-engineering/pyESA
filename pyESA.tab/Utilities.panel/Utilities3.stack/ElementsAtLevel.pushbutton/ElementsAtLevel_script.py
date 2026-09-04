# -*- coding: utf-8 -*-
__title__ = "Elements\nat Level"

__doc__ = """Version = 2.0
Date    = 04.09.2026
_____________________________________________________________________
Riassegna il livello di host degli elementi selezionati lasciandoli
nella posizione attuale (l'offset viene ricalcolato). Togliendo la
spunta all'opzione, gli elementi si spostano insieme al nuovo livello.

La selezione e' libera: non si sceglie piu' una categoria, e' lo script
che riconosce la categoria di ogni elemento e applica il metodo giusto.
Per gli elementi a doppio livello (muri, pilastri, scale, coperture) si
possono assegnare livello di base e livello superiore in una sola
operazione, anche a due livelli diversi.
_____________________________________________________________________
Author(s): bimdifferent, ESA Engineering
"""

__author__ = "bimdifferent, ESA Engineering"

from collections import OrderedDict

from pyrevit import revit, script, DB, UI, forms

from rpw.ui.forms import CheckBox, FlexForm, Label, Separator, Button, ComboBox

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

BIC = DB.BuiltInCategory
BIP = DB.BuiltInParameter

MAX_SKIPPED_ROWS = 200
FEET_TO_MM = 304.8
# La finestra rpw usa SizeToContent: si dimensiona sul controllo piu' largo.
FORM_WIDTH = 270


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


def collect_elements():
    elements = []
    try:
        pre_ids = list(uidoc.Selection.GetElementIds())
    except Exception:
        pre_ids = []
    for eid in pre_ids:
        element = doc.GetElement(eid)
        if element is not None and not isinstance(element, DB.ElementType):
            elements.append(element)
    if elements:
        return elements

    with forms.WarningBar(title='Seleziona gli elementi da riassegnare, poi premi Finish'):
        try:
            references = list(uidoc.Selection.PickObjects(
                UI.Selection.ObjectType.Element, 'Seleziona gli elementi'))
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


def resolve_pair(element, pair, slot):
    """Parametro livello + parametri offset per una coppia di regole."""
    level_bip, offset_bips, mode = pair
    level_param = find_param(element, [level_bip])
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


def resolve_by_name(element, slot):
    """Ultima spiaggia: parametri cercati per nome (IT/EN)."""
    if not slot.get('lnames'):
        return None, []
    level_param = find_param(element, [], slot['lnames'])
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
        return SKIPPED, u'offset superiore non scrivibile'
    base_param = find_param(element, ['WALL_BASE_CONSTRAINT'], writable=False)
    base_level = None
    if base_param is not None:
        base_level = doc.GetElement(base_param.AsElementId())
    if keep and base_level is None:
        return SKIPPED, u'livello di base non risolvibile'
    if keep:
        base_offset = find_param(element, ['WALL_BASE_OFFSET'], writable=False)
        height = find_param(element, ['WALL_USER_HEIGHT_PARAM'], writable=False)
        if height is None:
            return SKIPPED, u'altezza muro non leggibile'
        top_elevation = base_level.Elevation + height.AsDouble()
        if base_offset is not None:
            top_elevation += base_offset.AsDouble()
    level_param.Set(target_level.Id)
    if keep:
        offset_params[0].Set(top_elevation - target_level.Elevation)
    return APPLIED, u''


def apply_slot(element, slot, target_level, keep):
    """Applica un livello a uno slot dell'elemento. Ritorna (stato, motivo)."""
    reason = u'parametro di livello assente o di sola lettura'
    candidates = []
    for pair in slot['pairs']:
        level_param, offset_params = resolve_pair(element, pair, slot)
        if level_param is not None:
            candidates.append((level_param, offset_params))
    by_name = resolve_by_name(element, slot)
    if by_name[0] is not None:
        candidates.append(by_name)

    for level_param, offset_params in candidates:
        current_level = doc.GetElement(level_param.AsElementId())

        if current_level is None:
            if slot.get('special') == 'wall_top':
                return apply_wall_top_unconnected(
                    element, level_param, offset_params, target_level, keep)
            reason = u'livello corrente non assegnato'
            continue

        if element_id_value(current_level.Id) == element_id_value(target_level.Id):
            return ALREADY, u''

        if keep and not offset_params:
            reason = u'nessun parametro di offset scrivibile'
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


# ------------------------------------------------------------------- input

levels = list(DB.FilteredElementCollector(doc)
              .OfClass(DB.Level)
              .WhereElementIsNotElementType()
              .ToElements())
if not levels:
    forms.alert(u'Il modello non contiene livelli.', title=u'Elements at Level', exitscript=True)
levels.sort(key=lambda lev: lev.Elevation)

# I livelli in elenco vanno dal piu' alto al piu' basso: OrderedDict + sort=False
# perche' rpw ordina alfabeticamente le chiavi di un dizionario normale.
levels_dict = OrderedDict()
for level in reversed(levels):
    levels_dict[u'{}   ({})'.format(level.Name, format_elevation(level.Elevation))] = level

elements = collect_elements()
if not elements:
    script.exit()

KEEP_TOOLTIP = (u"Spuntato: l'offset viene ricalcolato, gli elementi non si muovono.\n"
                u"Non spuntato: l'offset resta invariato e gli elementi si spostano "
                u"insieme al nuovo livello.")

components = [
    Label('Livello di base / riferimento', Width=FORM_WIDTH),
    CheckBox('ckb_base', 'Applica il livello di base', default=True, Width=FORM_WIDTH),
    ComboBox('cmb_base', levels_dict, sort=False, Width=FORM_WIDTH),
    Separator(Width=FORM_WIDTH),
    Label('Livello superiore (solo elementi a doppio livello)', Width=FORM_WIDTH),
    CheckBox('ckb_top', 'Applica il livello superiore', default=False, Width=FORM_WIDTH),
    ComboBox('cmb_top', levels_dict, sort=False, Width=FORM_WIDTH),
    Separator(Width=FORM_WIDTH),
    CheckBox('ckb_keep', 'Mantieni gli elementi nella posizione attuale',
             default=True, Width=FORM_WIDTH, ToolTip=KEEP_TOOLTIP),
    Separator(Width=FORM_WIDTH),
    Button('OK', Width=FORM_WIDTH),
]
flex_form = FlexForm('Elements at Level  -  {} elementi'.format(len(elements)), components)
flex_form.show()
if not flex_form.values:
    script.exit()

do_base = bool(flex_form.values['ckb_base'])
do_top = bool(flex_form.values['ckb_top'])
keep_elevation = bool(flex_form.values['ckb_keep'])
base_level = flex_form.values['cmb_base']
top_level = flex_form.values['cmb_top']

if not do_base and not do_top:
    forms.alert(u'Seleziona almeno uno fra livello di base e livello superiore.',
                title=u'Elements at Level', exitscript=True)


# ------------------------------------------------------------------ lavoro

elements, redirected = resolve_targets(elements)

stats = OrderedDict()   # categoria -> [base, top, gia' corretti]
skipped = []            # (elemento, motivo)


def bump(name, index):
    row = stats.get(name)
    if row is None:
        row = [0, 0, 0]
        stats[name] = row
    row[index] += 1


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
                status, reason = SKIPPED, u'errore Revit: {}'.format(error)

            if status == APPLIED:
                touched = True
                bump(name, 0 if slot['role'] == 'base' else 1)
            elif status == ALREADY:
                already = True
            else:
                reasons.append(u'{}: {}'.format(slot['role'], reason))

        if touched:
            if reasons:
                skipped.append((element, u'parziale - {}'.format(u'; '.join(reasons))))
            continue
        if already and not reasons:
            bump(name, 2)
            continue
        if not reasons:
            reasons.append(u'nessun livello applicabile con le opzioni scelte')
        skipped.append((element, u'; '.join(reasons)))


# ------------------------------------------------------------------ report

output.print_md(u'# Elements at Level')
output.print_md(u'Elementi elaborati: **{}**{}  \n'
                u'Livello di base: **{}**  \n'
                u'Livello superiore: **{}**  \n'
                u'Elementi lasciati nella posizione attuale: **{}**'.format(
                    len(elements),
                    u' (di cui {} ricondotti al proprio host)'.format(redirected) if redirected else u'',
                    base_level.Name if do_base else u'non applicato',
                    top_level.Name if do_top else u'non applicato',
                    u'si' if keep_elevation else u'no'))

if stats:
    rows = []
    totals = [0, 0, 0]
    for name, row in stats.items():
        rows.append([name, row[0], row[1], row[2]])
        for i in range(3):
            totals[i] += row[i]
    rows.append([u'**Totale**', totals[0], totals[1], totals[2]])
    output.print_table(table_data=rows, title='',
                       columns=[u'Categoria', u'Livello base', u'Livello superiore',
                                u'Gia\' corretti'])
else:
    output.print_md(u'_Nessun elemento modificato._')

if skipped:
    output.print_md(u'## Elementi ignorati ({})'.format(len(skipped)))
    shown = skipped[:MAX_SKIPPED_ROWS]
    detail = [[output.linkify(element.Id), category_name(element), reason]
              for element, reason in shown]
    output.print_table(table_data=detail, title='',
                       columns=[u'Elemento', u'Categoria', u'Motivo'])
    remaining = len(skipped) - len(shown)
    if remaining > 0:
        output.print_md(u'_...e altri {} elementi ignorati non elencati._'.format(remaining))
