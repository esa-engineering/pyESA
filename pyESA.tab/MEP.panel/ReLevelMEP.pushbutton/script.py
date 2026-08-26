# -*- coding: utf-8 -*-
"""Riassegnazione del livello di riferimento agli elementi MEP.

Riassegna il livello di riferimento agli elementi MEP mantenendo invariata la
loro quota assoluta nel modello.

Questo strumento unifica in un unico comando le quattro varianti realizzate
originariamente in Dynamo:

    LevelReassignment_OnSelectedLevel_OnSelectedElements
    LevelReassignment_OnSelectedLevel_OnActiveViewElements
    LevelReassignment_Auto_OnSelectedElements
    LevelReassignment_Auto_OnActiveViewElements

Le due variabili che distinguevano i quattro grafi (fonte degli elementi e
metodo di determinazione del livello) sono diventate opzioni della finestra
di dialogo definita in LevelReassignmentWindow.xaml.

--------------------------------------------------------------------------
LOGICA APPLICATA (identica ai grafi Dynamo originali)
--------------------------------------------------------------------------

Gli elementi vengono classificati in tre gruppi in base alla categoria,
perche' ciascun gruppo richiede un trattamento diverso dei parametri Revit:

  GRUPPO A - elementi puntuali (terminali, apparecchi, quadri, dispositivi)
      imposta "Level" al livello di destinazione e ricalcola
      "Elevation from Level" come (quota Z originale - quota del nuovo
      livello), in modo che la quota assoluta dell'elemento resti identica.

  GRUPPO B - elementi lineari a percorso (canali, tubi, cavidotti, condotti)
      imposta solo "Reference Level". Nessun ricalcolo di quota: la
      geometria e' definita da coordinate reali e Revit aggiorna da se'
      l'offset rispetto al nuovo livello.

  GRUPPO C - raccordi e accessori (fitting e accessori di canali/tubi)
      imposta solo "Level". Nessun ricalcolo di quota.

Con il metodo automatico, il livello di destinazione viene determinato per
ogni singolo elemento cercando, tra i livelli di progetto ordinati per
quota, il primo livello immediatamente sottostante il punto di riferimento
dell'elemento (punto di inserimento, oppure punto medio della curva per gli
elementi lineari). Se nessun livello risulta sottostante, viene usato il
livello piu' basso del progetto. Lo spessore del pacchetto di finitura viene
sottratto alla quota del livello prima del confronto, e si applica a tutti
e tre i gruppi (come nei grafi originali).

--------------------------------------------------------------------------
NOTE DI CONVERSIONE DA DYNAMO
--------------------------------------------------------------------------

1. CATEGORIE. I grafi Dynamo classificavano gli elementi confrontando il
   NOME testuale della categoria (es. "Air Terminals"). Questo dipende dalla
   lingua dell'interfaccia di Revit e dalla versione, ed era infatti fonte di
   incongruenze tra i quattro file ("Flex Duct" / "Flex Ducts",
   "Speciality Equipment" / "Specialty Equipment"). Qui la classificazione
   usa gli enumeratori BuiltInCategory, indipendenti dalla lingua.

2. PARAMETRI. Analogamente, i parametri vengono cercati prima tramite
   BuiltInParameter e solo in subordine per nome (inglese e italiano),
   cosi' lo strumento funziona anche su installazioni Revit localizzate.

3. UNITA' DI MISURA. I grafi Dynamo mescolavano lo spessore in millimetri
   con quote lette in unita' di progetto: funzionavano correttamente solo
   nei progetti con unita' di lunghezza in millimetri. Qui tutti i calcoli
   avvengono nelle unita' interne di Revit (piedi decimali) e la conversione
   dai millimetri inseriti dall'utente e' esplicita, quindi lo strumento e'
   indipendente dalle unita' di progetto.

4. QUOTA DI RIFERIMENTO DEL LIVELLO. I grafi usavano Level.ProjectElevation.
   Qui si usa Level.Elevation, che e' la quota espressa nello stesso sistema
   di coordinate del punto di inserimento degli elementi: e' la grandezza
   corretta per calcolare l'offset. Le due coincidono in tutti i progetti in
   cui la base delle quote non e' traslata. Per tornare al comportamento
   Dynamo e' sufficiente modificare la funzione level_reference_elevation().

5. CATEGORIA "PIPE ACCESSORIES". Nei quattro file originali era classificata
   in modo incoerente: gruppo A in un file, gruppo C negli altri tre. E'
   stata mantenuta nel gruppo C (solo "Level", nessun ricalcolo di quota),
   come da indicazione ricevuta e coerentemente con la maggioranza dei file.

6. CATEGORIE "GENERIC MODELS" e "LIGHTING FIXTURES". Erano presenti nel
   gruppo A di uno solo dei quattro file. Sono state mantenute nell'elenco
   unificato. Per escluderle e' sufficiente rimuoverle da
   GROUP_A_CATEGORY_NAMES qui sotto.

--------------------------------------------------------------------------
MOTORE PYTHON
--------------------------------------------------------------------------

Lo script e' scritto per il motore IronPython di pyRevit, che e' quello
predefinito: il file NON deve contenere la direttiva "#! python3". Il
modulo pyrevit.forms, usato qui per caricare la finestra XAML, si appoggia
a wpf.LoadComponent, disponibile solo in IronPython. Sotto il motore
CPython3 le finestre WPF di pyRevit non sono attualmente supportate
(pyRevit issue #3033).
"""

from System.Collections.Generic import List

from pyrevit import revit, DB, forms, script


doc = revit.doc
output = script.get_output()
logger = script.get_logger()


# =========================================================================
# CONFIGURAZIONE
# =========================================================================
# Le categorie sono dichiarate come nomi di BuiltInCategory. I nomi non
# riconosciuti dalla versione di Revit in uso vengono semplicemente
# ignorati, cosi' lo script resta compatibile con piu' versioni.

# Gruppo A: elementi puntuali. Level + ricalcolo di "Elevation from Level".
GROUP_A_CATEGORY_NAMES = [
    'OST_DuctTerminal',            # Air Terminals
    'OST_CommunicationDevices',
    'OST_DataDevices',
    'OST_ElectricalEquipment',
    'OST_ElectricalFixtures',
    'OST_FireAlarmDevices',
    'OST_GenericModel',            # vedi nota 6
    'OST_LightingDevices',
    'OST_LightingFixtures',        # vedi nota 6
    'OST_MechanicalEquipment',
    'OST_NurseCallDevices',
    'OST_PlumbingFixtures',
    'OST_SecurityDevices',
    'OST_SpecialityEquipment',     # "Specialty Equipment" nelle UI recenti
    'OST_Sprinklers',
    'OST_TelephoneDevices',
]

# Gruppo B: elementi lineari a percorso. Solo "Reference Level".
GROUP_B_CATEGORY_NAMES = [
    'OST_CableTray',
    'OST_Conduit',
    'OST_DuctCurves',
    'OST_FlexDuctCurves',
    'OST_FlexPipeCurves',
    'OST_PipeCurves',
]

# Gruppo C: raccordi e accessori. Solo "Level".
GROUP_C_CATEGORY_NAMES = [
    'OST_CableTrayFitting',
    'OST_ConduitFitting',
    'OST_DuctAccessory',
    'OST_DuctFitting',
    'OST_PipeAccessory',           # vedi nota 5
    'OST_PipeFitting',
]

# Parametri: prima i BuiltInParameter (indipendenti dalla lingua), poi i
# nomi come ricerca di riserva. Viene usato il primo candidato presente
# sull'elemento e non di sola lettura.
LEVEL_PARAM_BIPS = [
    'FAMILY_LEVEL_PARAM',
    'SCHEDULE_LEVEL_PARAM',
    'RBS_START_LEVEL_PARAM',
    'FAMILY_BASE_LEVEL_PARAM',
]
LEVEL_PARAM_NAMES = [
    'Level',
    'Livello',
]

ELEVATION_PARAM_BIPS = [
    'INSTANCE_ELEVATION_PARAM',
]
ELEVATION_PARAM_NAMES = [
    'Elevation from Level',
    'Offset from Level',
    'Quota dal livello',
    'Offset dal livello',
]

REFERENCE_LEVEL_PARAM_BIPS = [
    'RBS_START_LEVEL_PARAM',
]
REFERENCE_LEVEL_PARAM_NAMES = [
    'Reference Level',
    'Livello di riferimento',
]

XAML_FILE_NAME = 'LevelReassignmentWindow.xaml'

GROUP_A = 'A'
GROUP_B = 'B'
GROUP_C = 'C'

GROUP_LABELS = {
    GROUP_A: u'Gruppo A - elementi puntuali (Level + Elevation from Level)',
    GROUP_B: u'Gruppo B - elementi lineari (Reference Level)',
    GROUP_C: u'Gruppo C - raccordi e accessori (Level)',
}

MAX_SKIPPED_DETAIL_ROWS = 200


# =========================================================================
# COMPATIBILITA' TRA VERSIONI DI REVIT
# =========================================================================

def element_id_value(element_id):
    """Valore numerico di un ElementId.

    ElementId.Value esiste da Revit 2024, IntegerValue nelle versioni
    precedenti (e deprecato in quelle successive).
    """
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def mm_to_internal(value_mm):
    """Converte millimetri nelle unita' interne di Revit (piedi decimali)."""
    try:
        # Revit 2021 e successivi
        return DB.UnitUtils.ConvertToInternalUnits(
            value_mm, DB.UnitTypeId.Millimeters)
    except AttributeError:
        # Revit 2020 e precedenti
        return DB.UnitUtils.ConvertToInternalUnits(
            value_mm, DB.DisplayUnitType.DUT_MILLIMETERS)


def internal_to_mm(value_internal):
    """Converte dalle unita' interne di Revit a millimetri."""
    try:
        return DB.UnitUtils.ConvertFromInternalUnits(
            value_internal, DB.UnitTypeId.Millimeters)
    except AttributeError:
        return DB.UnitUtils.ConvertFromInternalUnits(
            value_internal, DB.DisplayUnitType.DUT_MILLIMETERS)


def format_length(value_internal):
    """Formatta una lunghezza secondo le unita' di progetto.

    Se la formattazione nativa non e' disponibile nella versione in uso,
    ricade su una rappresentazione in millimetri.
    """
    try:
        # Revit 2021 e successivi
        return DB.UnitFormatUtils.Format(
            doc.GetUnits(), DB.SpecTypeId.Length, value_internal, False)
    except Exception:
        pass
    try:
        # Revit 2020 e precedenti
        return DB.UnitFormatUtils.Format(
            doc.GetUnits(), DB.UnitType.UT_Length, value_internal, False, False)
    except Exception:
        pass
    return u'{:+.0f} mm'.format(internal_to_mm(value_internal))


def resolve_categories(category_names):
    """Risolve i nomi di BuiltInCategory validi nella versione in uso."""
    resolved = []
    for name in category_names:
        built_in_category = getattr(DB.BuiltInCategory, name, None)
        if built_in_category is None:
            logger.debug('BuiltInCategory non disponibile: %s', name)
            continue
        resolved.append(built_in_category)
    return resolved


GROUP_A_CATEGORIES = resolve_categories(GROUP_A_CATEGORY_NAMES)
GROUP_B_CATEGORIES = resolve_categories(GROUP_B_CATEGORY_NAMES)
GROUP_C_CATEGORIES = resolve_categories(GROUP_C_CATEGORY_NAMES)
ALL_CATEGORIES = GROUP_A_CATEGORIES + GROUP_B_CATEGORIES + GROUP_C_CATEGORIES


def category_keys(built_in_categories):
    return set(
        element_id_value(DB.ElementId(bic)) for bic in built_in_categories
    )


GROUP_KEYS = [
    (GROUP_A, category_keys(GROUP_A_CATEGORIES)),
    (GROUP_B, category_keys(GROUP_B_CATEGORIES)),
    (GROUP_C, category_keys(GROUP_C_CATEGORIES)),
]


def build_category_filter():
    """Filtro multicategoria con tutte le categorie gestite."""
    category_list = List[DB.BuiltInCategory]()
    for built_in_category in ALL_CATEGORIES:
        category_list.Add(built_in_category)
    return DB.ElementMulticategoryFilter(category_list)


# =========================================================================
# CLASSIFICAZIONE E LETTURA DEGLI ELEMENTI
# =========================================================================

def classify_element(element):
    """Ritorna 'A', 'B', 'C' oppure None se la categoria non e' gestita."""
    category = element.Category
    if category is None:
        return None
    key = element_id_value(category.Id)
    for group, keys in GROUP_KEYS:
        if key in keys:
            return group
    return None


def level_reference_elevation(level):
    """Quota del livello nel sistema di coordinate dei punti di inserimento.

    Si usa Level.Elevation e non Level.ProjectElevation: vedi nota 4 nella
    documentazione in testa al file.
    """
    return level.Elevation


def get_reference_z(element):
    """Quota Z del punto di riferimento dell'elemento.

    Elementi puntuali: quota del punto di inserimento.
    Elementi lineari:  quota del punto medio della curva.
    Ritorna None se l'elemento non ha una posizione utilizzabile (tipico
    delle famiglie host based, workplane based e in place).
    """
    try:
        location = element.Location
    except Exception:
        return None

    if location is None:
        return None

    if isinstance(location, DB.LocationPoint):
        return location.Point.Z

    if isinstance(location, DB.LocationCurve):
        curve = location.Curve
        if curve is None:
            return None
        return curve.Evaluate(0.5, True).Z

    return None


def get_writable_parameter(element, bip_names, parameter_names):
    """Primo parametro scrivibile tra i candidati indicati."""
    for bip_name in bip_names:
        built_in_parameter = getattr(DB.BuiltInParameter, bip_name, None)
        if built_in_parameter is None:
            continue
        try:
            parameter = element.get_Parameter(built_in_parameter)
        except Exception:
            parameter = None
        if parameter is not None and not parameter.IsReadOnly:
            return parameter

    for parameter_name in parameter_names:
        parameter = element.LookupParameter(parameter_name)
        if parameter is not None and not parameter.IsReadOnly:
            return parameter

    return None


def get_current_level_id(parameter):
    """ElementId attualmente contenuto in un parametro di tipo livello."""
    try:
        return parameter.AsElementId()
    except Exception:
        return None


# =========================================================================
# CALCOLO DEL LIVELLO DI DESTINAZIONE
# =========================================================================

def get_project_levels():
    """Livelli del progetto ordinati per quota crescente."""
    levels = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.Level)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    levels.sort(key=level_reference_elevation)
    return levels


def find_level_below(elevation, levels_ascending, offset):
    """Primo livello immediatamente sottostante la quota indicata.

    Replica il comportamento dei grafi Dynamo: si contano i livelli la cui
    quota (diminuita dello spessore di finitura) e' inferiore alla quota
    dell'elemento e si prende l'ultimo di essi. Se nessun livello soddisfa
    la condizione, si ricade sul livello piu' basso del progetto.
    """
    chosen = levels_ascending[0]
    for level in levels_ascending:
        if elevation > (level_reference_elevation(level) - offset):
            chosen = level
        else:
            break
    return chosen


# =========================================================================
# APPLICAZIONE DELLE MODIFICHE
# =========================================================================

class ElementOutcome(object):
    """Esito dell'elaborazione di un singolo elemento."""

    APPLIED = 'applied'
    ALREADY_OK = 'already_ok'
    SKIPPED = 'skipped'

    def __init__(self, status, reason=None):
        self.status = status
        self.reason = reason


def apply_to_element(element, group, target_level, dry_run):
    """Applica la riassegnazione di livello a un singolo elemento."""

    if group == GROUP_A:
        reference_z = get_reference_z(element)
        if reference_z is None:
            return ElementOutcome(
                ElementOutcome.SKIPPED,
                u'nessuna posizione utilizzabile: elemento non basato su livello')

        level_parameter = get_writable_parameter(
            element, LEVEL_PARAM_BIPS, LEVEL_PARAM_NAMES)
        if level_parameter is None:
            return ElementOutcome(
                ElementOutcome.SKIPPED,
                u'parametro del livello assente o di sola lettura')

        elevation_parameter = get_writable_parameter(
            element, ELEVATION_PARAM_BIPS, ELEVATION_PARAM_NAMES)
        if elevation_parameter is None:
            return ElementOutcome(
                ElementOutcome.SKIPPED,
                u'parametro "Elevation from Level" assente o di sola lettura')

        current_level_id = get_current_level_id(level_parameter)
        if current_level_id is not None \
                and element_id_value(current_level_id) == element_id_value(target_level.Id):
            # Livello gia' corretto: l'offset attuale e' per definizione
            # quello giusto, quindi non serve riscrivere nulla.
            return ElementOutcome(ElementOutcome.ALREADY_OK)

        if not dry_run:
            level_parameter.Set(target_level.Id)
            new_offset = reference_z - level_reference_elevation(target_level)
            elevation_parameter.Set(new_offset)

        return ElementOutcome(ElementOutcome.APPLIED)

    if group == GROUP_B:
        reference_parameter = get_writable_parameter(
            element, REFERENCE_LEVEL_PARAM_BIPS, REFERENCE_LEVEL_PARAM_NAMES)
        if reference_parameter is None:
            return ElementOutcome(
                ElementOutcome.SKIPPED,
                u'parametro "Reference Level" assente o di sola lettura')

        current_level_id = get_current_level_id(reference_parameter)
        if current_level_id is not None \
                and element_id_value(current_level_id) == element_id_value(target_level.Id):
            return ElementOutcome(ElementOutcome.ALREADY_OK)

        if not dry_run:
            reference_parameter.Set(target_level.Id)

        return ElementOutcome(ElementOutcome.APPLIED)

    if group == GROUP_C:
        level_parameter = get_writable_parameter(
            element, LEVEL_PARAM_BIPS, LEVEL_PARAM_NAMES)
        if level_parameter is None:
            return ElementOutcome(
                ElementOutcome.SKIPPED,
                u'parametro del livello assente o di sola lettura')

        current_level_id = get_current_level_id(level_parameter)
        if current_level_id is not None \
                and element_id_value(current_level_id) == element_id_value(target_level.Id):
            return ElementOutcome(ElementOutcome.ALREADY_OK)

        if not dry_run:
            level_parameter.Set(target_level.Id)

        return ElementOutcome(ElementOutcome.APPLIED)

    return ElementOutcome(ElementOutcome.SKIPPED, u'categoria non gestita')


def resolve_target_level(element, group, options, levels_ascending):
    """Livello di destinazione per l'elemento indicato."""
    if not options.auto_level:
        return options.manual_level, None

    reference_z = get_reference_z(element)
    if reference_z is None:
        return None, u'nessuna posizione utilizzabile: impossibile calcolare il livello'

    level = find_level_below(
        reference_z, levels_ascending, options.thickness_internal)
    return level, None


# =========================================================================
# RACCOLTA DEGLI ELEMENTI
# =========================================================================

def collect_selected_elements():
    """Elementi attualmente selezionati in Revit, escluse le definizioni di tipo."""
    elements = []
    for element in revit.get_selection():
        if element is None:
            continue
        if isinstance(element, DB.ElementType):
            continue
        elements.append(element)
    return elements


def collect_active_view_elements():
    """Elementi MEP gestiti, visibili nella vista attiva.

    Ritorna None se la vista attiva non ammette la raccolta di elementi
    (per esempio un abaco o una vista non grafica).
    """
    active_view = doc.ActiveView
    if active_view is None:
        return None
    try:
        collector = DB.FilteredElementCollector(doc, active_view.Id)
        return list(
            collector
            .WhereElementIsNotElementType()
            .WherePasses(build_category_filter())
            .ToElements()
        )
    except Exception as error:
        logger.debug('Raccolta dalla vista attiva non possibile: %s', error)
        return None


# =========================================================================
# FINESTRA DI DIALOGO
# =========================================================================

class ReassignmentOptions(object):
    """Scelte effettuate dall'utente nella finestra di dialogo."""

    def __init__(self, use_selection, auto_level, manual_level,
                 thickness_mm, dry_run):
        self.use_selection = use_selection
        self.auto_level = auto_level
        self.manual_level = manual_level
        self.thickness_mm = thickness_mm
        self.thickness_internal = mm_to_internal(thickness_mm)
        self.dry_run = dry_run


class LevelReassignmentWindow(forms.WPFWindow):
    """Finestra di dialogo definita in LevelReassignmentWindow.xaml."""

    def __init__(self, xaml_file, levels, selected_count, view_count,
                 view_name, preselect_level=None):
        forms.WPFWindow.__init__(self, xaml_file)

        self.options = None
        self._combo_levels = []

        self._setup_sources(selected_count, view_count, view_name)
        self._setup_levels(levels, preselect_level)

    # ------------------------------------------------------------------
    # inizializzazione dei controlli
    # ------------------------------------------------------------------

    def _setup_sources(self, selected_count, view_count, view_name):
        if selected_count:
            self.tb_selection_info.Text = \
                u'{} elementi attualmente selezionati.'.format(selected_count)
        else:
            self.tb_selection_info.Text = \
                u'Nessun elemento selezionato in Revit.'
            self.rb_source_selection.IsEnabled = False

        if view_count is None:
            self.tb_view_info.Text = \
                u'La vista attiva non consente la raccolta degli elementi.'
            self.rb_source_view.IsEnabled = False
        else:
            self.tb_view_info.Text = \
                u'{} elementi MEP gestiti nella vista "{}".'.format(
                    view_count, view_name)

        if selected_count:
            self.rb_source_selection.IsChecked = True
        elif view_count:
            self.rb_source_view.IsChecked = True

    def _setup_levels(self, levels, preselect_level):
        # In elenco i livelli sono mostrati dal piu' alto al piu' basso,
        # come si leggono in una sezione.
        self._combo_levels = list(reversed(levels))
        for level in self._combo_levels:
            self.cb_levels.Items.Add(
                u'{}   ({})'.format(
                    level.Name,
                    format_length(level_reference_elevation(level))))

        # Se la vista attiva e' associata a un livello, quello viene
        # proposto come default. Altrimenti nessuna voce e' preselezionata,
        # cosi' la scelta del livello resta un gesto esplicito.
        selected_index = -1
        if preselect_level is not None:
            target = element_id_value(preselect_level.Id)
            for index, level in enumerate(self._combo_levels):
                if element_id_value(level.Id) == target:
                    selected_index = index
                    break
        self.cb_levels.SelectedIndex = selected_index

        self.rb_level_manual.IsChecked = True

    # ------------------------------------------------------------------
    # gestori degli eventi
    # ------------------------------------------------------------------

    def on_run(self, sender, args):
        use_selection = bool(self.rb_source_selection.IsChecked)
        use_view = bool(self.rb_source_view.IsChecked)
        if not use_selection and not use_view:
            forms.alert(u'Indicare quali elementi elaborare.', title=u'Dati mancanti')
            return

        auto_level = bool(self.rb_level_auto.IsChecked)
        manual_level = None

        if not auto_level:
            index = self.cb_levels.SelectedIndex
            if index is None or index < 0:
                forms.alert(u'Selezionare il livello di destinazione.',
                            title=u'Dati mancanti')
                return
            manual_level = self._combo_levels[index]

        thickness_mm = 0.0
        if auto_level:
            raw_value = (self.tb_thickness.Text or u'').strip().replace(u',', u'.')
            if not raw_value:
                raw_value = u'0'
            try:
                thickness_mm = float(raw_value)
            except ValueError:
                forms.alert(
                    u'Lo spessore del pacchetto di finitura deve essere un '
                    u'numero espresso in millimetri.',
                    title=u'Valore non valido')
                return
            if thickness_mm < 0:
                forms.alert(
                    u'Lo spessore del pacchetto di finitura non puo\' essere '
                    u'negativo.',
                    title=u'Valore non valido')
                return

        self.options = ReassignmentOptions(
            use_selection=use_selection,
            auto_level=auto_level,
            manual_level=manual_level,
            thickness_mm=thickness_mm,
            dry_run=bool(self.chk_dryrun.IsChecked),
        )
        self.Close()

    def on_cancel(self, sender, args):
        self.options = None
        self.Close()


# =========================================================================
# RESOCONTO
# =========================================================================

def print_report(options, applied, already_ok, skipped, unmanaged_count,
                 total_count):
    if options.dry_run:
        output.print_md(u'# Riassegnazione livello - simulazione')
        output.print_md(
            u'**Nessuna modifica e\' stata applicata al modello.** '
            u'I valori seguenti sono una stima di quanto verrebbe modificato.')
    else:
        output.print_md(u'# Riassegnazione livello - resoconto')

    if options.auto_level:
        level_description = \
            u'automatico (primo livello sottostante), spessore di finitura ' \
            u'{:.0f} mm'.format(options.thickness_mm)
    else:
        level_description = u'manuale, livello "{}"'.format(
            options.manual_level.Name)

    source_description = u'elementi selezionati' if options.use_selection \
        else u'elementi della vista attiva'

    output.print_md(
        u'- Fonte: {}\n'
        u'- Livello di destinazione: {}\n'
        u'- Elementi esaminati: {}'.format(
            source_description, level_description, total_count))

    output.print_md(u'## Elementi elaborati')
    rows = []
    for group in (GROUP_A, GROUP_B, GROUP_C):
        rows.append([
            GROUP_LABELS[group],
            str(applied.get(group, 0)),
            str(already_ok.get(group, 0)),
        ])
    rows.append([
        u'**Totale**',
        str(sum(applied.values())),
        str(sum(already_ok.values())),
    ])
    output.print_table(
        table_data=rows,
        title='',
        columns=[u'Gruppo', u'Modificati', u'Gia\' corretti'])

    output.print_md(
        u'- Elementi con categoria non gestita: {}\n'
        u'- Elementi ignorati: {}'.format(unmanaged_count, len(skipped)))

    if not skipped:
        return

    output.print_md(u'## Elementi ignorati')
    output.print_md(
        u'Tipicamente famiglie host based, workplane based o in place, '
        u'oppure elementi i cui parametri di livello sono di sola lettura.')

    shown = skipped[:MAX_SKIPPED_DETAIL_ROWS]
    detail_rows = []
    for element, reason in shown:
        try:
            category_name = element.Category.Name
        except Exception:
            category_name = u'-'
        detail_rows.append([
            output.linkify(element.Id),
            category_name,
            reason,
        ])
    output.print_table(
        table_data=detail_rows,
        title='',
        columns=[u'Elemento', u'Categoria', u'Motivo'])

    remaining = len(skipped) - len(shown)
    if remaining > 0:
        output.print_md(
            u'_...e altri {} elementi ignorati non elencati._'.format(remaining))


# =========================================================================
# PROGRAMMA PRINCIPALE
# =========================================================================

def main():
    if not ALL_CATEGORIES:
        forms.alert(
            u'Nessuna delle categorie configurate e\' disponibile in questa '
            u'versione di Revit.',
            exitscript=True)

    levels_ascending = get_project_levels()
    if not levels_ascending:
        forms.alert(u'Nel progetto non sono presenti livelli.', exitscript=True)

    selected_elements = collect_selected_elements()
    view_elements = collect_active_view_elements()

    if not selected_elements and not view_elements:
        forms.alert(
            u'Non ci sono elementi da elaborare: nessun elemento selezionato '
            u'e nessun elemento MEP gestito nella vista attiva.',
            exitscript=True)

    active_view = doc.ActiveView
    active_view_name = active_view.Name if active_view is not None else u'-'

    # Livello associato alla vista attiva, se esiste (viste in pianta):
    # viene proposto come livello di destinazione predefinito.
    view_level = None
    if active_view is not None:
        try:
            view_level = active_view.GenLevel
        except Exception:
            view_level = None

    window = LevelReassignmentWindow(
        script.get_bundle_file(XAML_FILE_NAME),
        levels_ascending,
        len(selected_elements),
        None if view_elements is None else len(view_elements),
        active_view_name,
        view_level,
    )
    window.ShowDialog()

    options = window.options
    if options is None:
        script.exit()

    elements = selected_elements if options.use_selection else view_elements
    if not elements:
        forms.alert(u'Non ci sono elementi da elaborare.', exitscript=True)

    applied = {GROUP_A: 0, GROUP_B: 0, GROUP_C: 0}
    already_ok = {GROUP_A: 0, GROUP_B: 0, GROUP_C: 0}
    skipped = []
    unmanaged = []

    def process_all():
        """Elabora tutti gli elementi. Chiamata dentro o fuori transazione."""
        for element in elements:
            group = classify_element(element)
            if group is None:
                # Nella modalita' "vista attiva" il filtro esclude a monte le
                # categorie non gestite; qui contano solo le selezioni manuali.
                unmanaged.append(element)
                continue

            target_level, error = resolve_target_level(
                element, group, options, levels_ascending)
            if target_level is None:
                skipped.append((element, error))
                continue

            try:
                outcome = apply_to_element(
                    element, group, target_level, options.dry_run)
            except Exception as exception:
                skipped.append((element, u'errore Revit: {}'.format(exception)))
                continue

            if outcome.status == ElementOutcome.APPLIED:
                applied[group] += 1
            elif outcome.status == ElementOutcome.ALREADY_OK:
                already_ok[group] += 1
            else:
                skipped.append((element, outcome.reason))

    if options.dry_run:
        process_all()
    else:
        with revit.Transaction('Riassegnazione livello elementi MEP'):
            process_all()

    print_report(options, applied, already_ok, skipped, len(unmanaged),
                 len(elements))


if __name__ == '__main__':
    main()
