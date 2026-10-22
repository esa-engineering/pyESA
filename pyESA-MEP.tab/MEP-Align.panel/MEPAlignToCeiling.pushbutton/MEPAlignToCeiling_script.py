# -*- coding: utf-8 -*-
# --------------------------------------------------------------------------
# Allineamento in quota degli elementi MEP al controsoffitto o, dove manca,
# all'intradosso del solaio che chiude il piano di riferimento.
#
# Stessa impostazione di MEPAlignToWall, ruotata di novanta gradi: per ogni
# elemento si traccia una RETTA DI ANALISI lungo l'asse Z del modello,
# passante per il punto di inserimento e lunga due volte la tolleranza
# indicata dall'utente, con il punto medio sull'origine della famiglia (le
# due estremita' sono quindi equidistanti dal punto). Un controsoffitto o un
# solaio e' un bersaglio valido se la retta interseca la sua SUPERFICIE
# INFERIORE. L'elemento viene poi traslato solo in Z fino a portare il punto
# di inserimento su quella superficie: la posizione in pianta e la rotazione
# non vengono mai toccate.
#
# Screening: le famiglie possono essere montate a parete o a soffitto, quindi
# vengono trattate solo le istanze con il parametro di istanza Yes/No
# "u_VIS_Ceiling_installation" = Yes. Le altre compaiono nel resoconto con il
# loro motivo.
#
# Il bersaglio non lo sceglie l'utente: se sulla verticale del dispositivo
# c'e' un controsoffitto nel piano si va li', altrimenti all'intradosso del
# solaio. Fanno eccezione i rilevatori da controsoffitto (family name con
# "e_FA.DE" e "False Ceiling"): vanno SEMPRE all'intradosso (il solaio, o un
# controsoffitto incollato al solaio come finitura o isolante), e il loro
# parametro dz_Buzzer riceve la distanza, sempre misurata, fino al
# controsoffitto staccato sottostante. Senza quel controsoffitto il rilevatore
# e' segnalato come non adeguato e non si sposta.
#
# Esito: una finestra NON MODALE con i totali e la tabella interrogabile dei
# soli elementi non allineati; il pulsante di una riga seleziona e inquadra
# l'elemento, e Revit resta libero per modificarlo. Il resoconto dettagliato
# nel pannello di output e' il "report avanzato", stampato su richiesta.
#
# La logica completa (piano di riferimento, pacchetti di piu' istanze anche
# da modelli collegati, precedenza del controsoffitto, limiti noti) e'
# descritta nel README.md della cartella: va tenuto allineato al codice.
# --------------------------------------------------------------------------
__title__ = "MEP\nAlign to Ceiling"
__doc__ = """Version = 1.4
Date    = 22.10.2026
________________________________________________________________
Description:
Moves the selected MEP devices vertically so that their insertion point
lies on the bottom face of the ceiling above them or, where there is no
ceiling, on the soffit of the floor slab that closes their level.
Only instances with the Yes/No instance parameter
u_VIS_Ceiling_installation = Yes are processed.
False ceiling detectors (family name with "e_FA.DE" and "False Ceiling")
always go to the soffit, and dz_Buzzer gets the measured distance down to
the false ceiling; without a false ceiling they are reported as unsuitable.
Plan position and rotation are never changed.
________________________________________________________________
Author(s): ESA Engineering
"""
__author__ = "ESA Engineering"

import os.path as op

import clr
from System.Collections.Generic import List

from pyrevit import revit, DB, UI, forms, script

import MEPAlignToCeiling_ui as ceiling_ui
import MEPAlignToCeiling_result_ui as result_ui


doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
logger = script.get_logger()


# =========================================================================
# CONFIGURAZIONE
# =========================================================================
# Le categorie sono nomi di BuiltInCategory: quelli non riconosciuti dalla
# versione di Revit in uso vengono ignorati. Stesso elenco di MEPAlignToWall,
# perche' le famiglie sono le stesse: e' il parametro di screening, non la
# categoria, a dire se un'istanza e' montata a soffitto.

MEP_CATEGORY_NAMES = [
    'OST_DuctTerminal',            # Air Terminals
    'OST_CommunicationDevices',
    'OST_DataDevices',
    'OST_ElectricalFixtures',
    'OST_FireAlarmDevices',
    'OST_LightingDevices',
    'OST_LightingFixtures',
    'OST_NurseCallDevices',
    'OST_SecurityDevices',
]

# Parametro di istanza Yes/No che distingue il montaggio a soffitto da quello
# a parete. Solo il valore Yes ammette l'istanza.
INSTALLATION_PARAM = 'u_VIS_Ceiling_installation'

# Rilevatori da controsoffitto: il family name deve contenere TUTTE queste
# stringhe (confronto senza distinzione fra maiuscole e minuscole). Vanno
# sempre all'intradosso del solaio, e BUZZER_PARAM (parametro di istanza di
# tipo lunghezza) riceve la distanza fra l'intradosso e il controsoffitto, a
# cui la famiglia porta la sua parte visibile.
BUZZER_FAMILY_TOKENS = ('e_FA.DE', 'False Ceiling')
BUZZER_PARAM = 'dz_Buzzer'

CEILING_CATEGORY_NAME = 'OST_Ceilings'
FLOOR_CATEGORY_NAME = 'OST_Floors'

XAML_FILE_NAME = 'MEPAlignToCeiling_form.xaml'
RESULT_XAML_FILE_NAME = 'MEPAlignToCeiling_result.xaml'
CONFIG_SECTION = 'ESA_MEPAlignToCeiling'

TRANSACTION_NAME = u'Align MEP to ceilings'

DEFAULT_TOLERANCE_MM = 500.0    # soglia proposta nella finestra (in mm)
MAX_TOLERANCE_MM = 5000.0      # oltre e' quasi certamente un errore di battitura

# La retta di analisi e' lunga LINE_FACTOR volte la tolleranza ed e'
# centrata sul punto di inserimento: si estende di una tolleranza sopra e di
# una sotto. E' la richiesta, e coincide con "superficie entro la tolleranza
# dal punto, sopra o sotto".
LINE_FACTOR = 2.0

POSITION_TOL_MM = 1.0           # sotto questo scarto l'elemento e' gia' a posto
LEVEL_TOL_MM = 1.0              # margine sul confronto con la quota del livello

# Due istanze impilate (solaio strutturale e finitura, controsoffitto in due
# strati, anche da modelli collegati diversi) separate da meno di questo
# scarto formano un pacchetto unico, e il bersaglio e' la faccia inferiore
# del pacchetto. Piu' largo dei 2 mm dei muri: fra due modelli collegati di
# discipline diverse il contatto esatto non e' garantito, mentre un vuoto
# vero fra controsoffitto e solaio e' sempre di decine di centimetri.
STACK_GAP_MM = 5.0

# Raggio verticale della ricerca oltre la retta di analisi. Serve per due
# motivi: trovare il solaio di chiusura anche quando e' oltre la tolleranza
# (per riportarne la distanza) e raccogliere il solaio su cui poggia il piano,
# cosi' un pavimento sopraelevato appoggiato sopra si fonde con lui in un
# pacchetto che sta sotto il livello e non diventa un falso intradosso.
SEARCH_MARGIN_MM = 2000.0

PROBE_PAD_MM = 100.0            # la sonda verticale sporge dal bounding box
XY_EPS_MM = 1.0                 # margine sul test in pianta delle facce
AMBIGUITY_TOL_MM = 20.0         # due controsoffitti entro questo scarto: segnalato

VERTICAL_FACE_TOL = 1.0e-4      # |normale Z| sotto: faccia verticale, scartata
LINK_AXIS_TOL = 1.0e-6          # scarto ammesso sull'asse Z di un collegamento

MAX_MOVED_ROWS = 300
MAX_SKIPPED_ROWS = 500
MAX_OVER_TOLERANCE_ROWS = 100

# Tipo di bersaglio, per la colonna omonima del resoconto.
K_CEILING = u'Ceiling'
K_FLOOR = u'Floor soffit'

# Vocabolario chiuso dei motivi di esclusione.
R_NO_POINT = u'element has no insertion point'
R_PARAM_MISSING = u'parameter "{}" not found on the instance'.format(INSTALLATION_PARAM)
R_PARAM_TYPE = u'"{}" is a type parameter: an instance parameter is expected'.format(INSTALLATION_PARAM)
R_PARAM_NOT_YESNO = u'"{}" is not a Yes/No parameter'.format(INSTALLATION_PARAM)
R_PARAM_EMPTY = u'"{}" has no value'.format(INSTALLATION_PARAM)
R_PARAM_NO = u'wall installation ("{}" = No)'.format(INSTALLATION_PARAM)
R_HOSTED_CEILING = u'hosted by ceiling {}: already on its host face'
R_HOSTED_LINK = u'hosted by a face of a linked model'
R_HOSTED_OTHER = u'hosted by {} {}'
R_PINNED = u'element is pinned'
R_GROUP = u'element inside group "{}"'
R_SUBCOMPONENT = u'sub-component of a nested family'
R_DESIGN_OPTION = u'inactive design option'
R_BORROWED = u'borrowed by another user'
R_CONNECTED = u'connected to other elements ({} connectors)'
R_CATEGORY_NOT_HANDLED = u'category not handled by the tool'
R_CATEGORY_NOT_SELECTED = u'category cleared in the dialog'
R_NO_SURFACE = u'no ceiling and no floor slab above the element in its storey'
R_BUZZER_NO_SLAB = u'false ceiling detector: no floor slab above the element in its storey'
# Il rilevatore da controsoffitto ha senso solo sotto un controsoffitto
# staccato dal solaio. Senza, e' la famiglia sbagliata: errore, non si sposta.
R_BUZZER_NO_CEILING = u'Unsuitable detector: no false ceiling below it (no ceiling detached ' \
                      u'from the slab under the device)'
R_BUZZER_MISSING = u'false ceiling detector: parameter "{}" not found on the instance'.format(BUZZER_PARAM)
R_BUZZER_TYPE = u'false ceiling detector: "{}" is a type parameter, an instance parameter is ' \
                u'expected'.format(BUZZER_PARAM)
R_BUZZER_NOT_LENGTH = u'false ceiling detector: "{}" is not a length parameter'.format(BUZZER_PARAM)
R_BUZZER_READ_ONLY = u'false ceiling detector: "{}" is read-only (driven by a formula?)'.format(BUZZER_PARAM)
R_ANALYSIS_ERROR = u'analysis error: {}'

# Motivi di scarto di un controsoffitto o di un solaio.
S_NO_BBOX = u'extents cannot be read'
S_NO_GEOMETRY = u'geometry cannot be read'
S_NO_FACE = u'no usable face'
S_DEMOLISHED = u'demolished in phase "{}"'

L_NOT_LOADED = u'link not loaded'
L_TILTED = u'link tilted or flipped: its vertical axis is not the model vertical'
L_NO_TRANSFORM = u'link transform cannot be read'
L_SCALED = u'scaled link: distances would not be comparable'


# =========================================================================
# COMPATIBILITA' TRA VERSIONI DI REVIT E UNITA'
# =========================================================================

def element_id_value(element_id):
    """Valore numerico di un ElementId (Value da Revit 2024, IntegerValue prima)."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def mm_to_internal(value_mm):
    """Converte millimetri nelle unita' interne di Revit (piedi decimali)."""
    return DB.UnitUtils.ConvertToInternalUnits(value_mm, DB.UnitTypeId.Millimeters)


def internal_to_mm(value_internal):
    """Converte dalle unita' interne di Revit a millimetri."""
    return DB.UnitUtils.ConvertFromInternalUnits(value_internal, DB.UnitTypeId.Millimeters)


def format_mm(value_internal):
    return u'{:.0f} mm'.format(internal_to_mm(value_internal))


def format_signed_mm(value_internal):
    return u'{:+.0f} mm'.format(internal_to_mm(value_internal))


def resolve_categories(category_names):
    """[(nome, BuiltInCategory)] validi nella versione in uso."""
    resolved = []
    for name in category_names:
        built_in_category = getattr(DB.BuiltInCategory, name, None)
        if built_in_category is None:
            logger.debug('BuiltInCategory non disponibile: %s', name)
            continue
        resolved.append((name, built_in_category))
    return resolved


def category_display_name(built_in_category):
    """Nome della categoria nella lingua dell'interfaccia di Revit."""
    try:
        category = DB.Category.GetCategory(doc, built_in_category)
        if category is not None:
            return category.Name
    except Exception:
        pass
    return u'{}'.format(built_in_category)


def revit_version():
    try:
        return int(doc.Application.VersionNumber)
    except Exception:
        return 0


def bic_key(built_in_category):
    return element_id_value(DB.ElementId(built_in_category))


MEP_CATEGORIES = resolve_categories(MEP_CATEGORY_NAMES)
MEP_CATEGORY_KEYS = set(bic_key(bic) for _name, bic in MEP_CATEGORIES)
CEILING_CATEGORY = getattr(DB.BuiltInCategory, CEILING_CATEGORY_NAME)
FLOOR_CATEGORY = getattr(DB.BuiltInCategory, FLOOR_CATEGORY_NAME)
CEILING_KEY = bic_key(CEILING_CATEGORY)

POSITION_TOL = mm_to_internal(POSITION_TOL_MM)
LEVEL_TOL = mm_to_internal(LEVEL_TOL_MM)
STACK_GAP = mm_to_internal(STACK_GAP_MM)
SEARCH_MARGIN = mm_to_internal(SEARCH_MARGIN_MM)
PROBE_PAD = mm_to_internal(PROBE_PAD_MM)
XY_EPS = mm_to_internal(XY_EPS_MM)
AMBIGUITY_TOL = mm_to_internal(AMBIGUITY_TOL_MM)


# =========================================================================
# AVVISI
# =========================================================================

WARNINGS = []


def warn(message):
    """Registra un avviso globale, senza duplicati."""
    if message not in WARNINGS:
        WARNINGS.append(message)


# =========================================================================
# ETICHETTE
# =========================================================================

def element_name(element):
    try:
        return DB.Element.Name.GetValue(element)
    except Exception:
        pass
    try:
        return element.Name
    except Exception:
        return u'id {}'.format(element_id_value(element.Id))


def category_of(element):
    """(nome localizzato, chiave numerica) della categoria di un elemento."""
    try:
        category = element.Category
        if category is not None:
            return category.Name, element_id_value(category.Id)
    except Exception:
        pass
    return u'-', None


def type_name_of(element):
    """"Famiglia: Tipo" di un'istanza."""
    try:
        symbol = element.Symbol
        if symbol is not None:
            family = symbol.Family
            family_name = DB.Element.Name.GetValue(family) if family else u'?'
            return u'{}: {}'.format(family_name, DB.Element.Name.GetValue(symbol))
    except Exception:
        pass
    return element_name(element)


def insertion_point(element):
    """Punto di inserimento, None se l'elemento non e' puntuale."""
    try:
        location = element.Location
    except Exception:
        return None
    if isinstance(location, DB.LocationPoint):
        try:
            return location.Point
        except Exception:
            return None
    return None


# =========================================================================
# SCREENING: MONTAGGIO A SOFFITTO
# =========================================================================

def is_yes_no(parameter):
    """True se il parametro e' di tipo Yes/No.

    GetDataType() esiste da Revit 2022; ParameterType resta come riserva per
    le versioni in cui il primo non risponde.
    """
    definition = parameter.Definition
    try:
        return definition.GetDataType() == DB.SpecTypeId.Boolean.YesNo
    except Exception:
        pass
    try:
        return definition.ParameterType == DB.ParameterType.YesNo
    except Exception:
        return parameter.StorageType == DB.StorageType.Integer


def screening_reason(element):
    """None se l'istanza e' montata a soffitto, altrimenti il motivo.

    Il parametro e' cercato SOLO sull'istanza: un parametro di tipo con lo
    stesso nome renderebbe uguali tutte le istanze del tipo, mentre la
    scelta parete / soffitto e' per singola istanza. Il caso viene distinto
    nel motivo, perche' e' un errore di impostazione della famiglia.
    """
    parameter = None
    try:
        parameter = element.LookupParameter(INSTALLATION_PARAM)
    except Exception:
        parameter = None

    if parameter is None:
        try:
            symbol = doc.GetElement(element.GetTypeId())
            if symbol is not None and symbol.LookupParameter(INSTALLATION_PARAM) is not None:
                return R_PARAM_TYPE
        except Exception:
            pass
        return R_PARAM_MISSING

    if parameter.StorageType != DB.StorageType.Integer or not is_yes_no(parameter):
        return R_PARAM_NOT_YESNO
    try:
        if not parameter.HasValue:
            return R_PARAM_EMPTY
    except Exception:
        pass
    if parameter.AsInteger() != 1:
        return R_PARAM_NO
    return None


# =========================================================================
# PIANO DI RIFERIMENTO
# =========================================================================
# "Solaio di chiusura del piano di riferimento": valgono solo le superfici
# che stanno SOPRA il livello dell'elemento. Senza questo limite la retta,
# che si estende anche verso il basso, prenderebbe per buono l'intradosso del
# solaio su cui il piano poggia quando un dispositivo e' modellato a quota
# zero, e lo porterebbe sotto il pavimento.

class Storey(object):
    """Livello di riferimento dell'elemento e quota minima delle superfici."""

    def __init__(self, bound, level, note):
        self.bound = bound          # None: nessun limite inferiore
        self.level = level
        self.note = note

    @property
    def level_name(self):
        if self.level is None:
            return u'-'
        return element_name(self.level)


def level_elevation(level):
    """Quota del livello in coordinate interne.

    ProjectElevation e non Elevation: Elevation puo' essere riferita al
    Survey Point, mentre la geometria e' sempre in coordinate interne (vedi
    CLAUDE.md e TagLinkedRooms_NOTE-SVILUPPO.md).
    """
    try:
        return level.ProjectElevation
    except Exception:
        pass
    try:
        return level.Elevation
    except Exception:
        return None


def host_levels():
    """[(quota, livello)] del documento corrente, dal basso verso l'alto."""
    levels = []
    try:
        for level in DB.FilteredElementCollector(doc).OfClass(DB.Level):
            elevation = level_elevation(level)
            if elevation is not None:
                levels.append((elevation, level))
    except Exception:
        pass
    levels.sort(key=lambda row: row[0])
    return levels


def reference_level(element):
    """Livello di riferimento dell'istanza, None se non ne ha uno."""
    level_id = None
    try:
        level_id = element.LevelId
    except Exception:
        level_id = None
    if level_id is None or level_id == DB.ElementId.InvalidElementId:
        level_id = None
        for name in ('FAMILY_LEVEL_PARAM', 'INSTANCE_REFERENCE_LEVEL_PARAM',
                     'INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM', 'SCHEDULE_LEVEL_PARAM'):
            bip = getattr(DB.BuiltInParameter, name, None)
            if bip is None:
                continue
            try:
                parameter = element.get_Parameter(bip)
                if parameter is not None and parameter.StorageType == DB.StorageType.ElementId:
                    candidate = parameter.AsElementId()
                    if candidate is not None and candidate != DB.ElementId.InvalidElementId:
                        level_id = candidate
                        break
            except Exception:
                continue
    if level_id is None:
        return None
    level = doc.GetElement(level_id)
    return level if isinstance(level, DB.Level) else None


def level_below(z, levels):
    """(quota, livello) piu' alto alla quota z o sotto, None se non c'e'."""
    found = None
    for elevation, level in levels:
        if elevation <= z + LEVEL_TOL:
            found = (elevation, level)
        else:
            break
    return found


def storey_of(element, point, levels):
    """Piano di riferimento dell'elemento.

    Il livello e' quello dell'istanza. Se pero' il punto di inserimento sta
    SOTTO il proprio livello, l'elemento e' riferito al livello superiore
    con uno scostamento negativo (pratica diffusa per i dispositivi a
    soffitto): il piano di riferimento e' allora quello del livello piu' alto
    sotto il punto, e la nota lo dichiara.
    """
    level = reference_level(element)
    if level is not None:
        elevation = level_elevation(level)
        if elevation is not None and point.Z >= elevation - LEVEL_TOL:
            return Storey(elevation, level, None)
        fallback = level_below(point.Z, levels)
        if fallback is not None:
            return Storey(fallback[0], fallback[1],
                          u'below its level "{}": storey of "{}" used'.format(
                              element_name(level), element_name(fallback[1])))
        return Storey(None, level, u'below its level "{}" and no level under it: '
                                   u'no lower limit'.format(element_name(level)))

    fallback = level_below(point.Z, levels)
    if fallback is not None:
        return Storey(fallback[0], fallback[1],
                      u'no reference level: storey of "{}" used'.format(element_name(fallback[1])))
    return Storey(None, None, u'no level found: no lower limit')


# =========================================================================
# SORGENTI: DOCUMENTO CORRENTE E MODELLI COLLEGATI
# =========================================================================

class SurfaceSource(object):
    """Documento in cui cercare controsoffitti e solai.

    Per il documento host la trasformazione e' l'identita'. Per un modello
    collegato porta dalle coordinate del link a quelle dell'host; il suo asse
    Z e' garantito verticale da link_transform_problem(), quindi una retta
    verticale resta verticale e una quota si converte con una traslazione.
    """

    def __init__(self, key, document, transform, link_name):
        self.key = key
        self.document = document
        self.transform = transform
        self.inverse = transform.Inverse if transform is not None else None
        self.link_name = link_name
        self.z_offset = transform.Origin.Z if transform is not None else 0.0
        self.surfaces = []

    @property
    def is_linked(self):
        return self.link_name is not None

    def to_local(self, point):
        if self.inverse is None:
            return point
        try:
            return self.inverse.OfPoint(point)
        except Exception:
            return None


def link_display_name(link):
    for getter in (lambda: DB.Element.Name.GetValue(link),
                   lambda: DB.Element.Name.GetValue(doc.GetElement(link.GetTypeId()))):
        try:
            name = getter()
            if name:
                return name
        except Exception:
            continue
    return u'link {}'.format(element_id_value(link.Id))


def link_transform_problem(transform):
    """Motivo per cui un collegamento non e' utilizzabile, None se va bene.

    A differenza di MEPAlignToWall un collegamento SPECULARE va bene: la
    riflessione avviene nel piano orizzontale e le quote non cambiano.
    Restano esclusi i collegamenti inclinati, dove la verticale del link non
    e' quella dell'host, e quelli in scala.
    """
    try:
        basis_z = transform.BasisZ
    except Exception:
        return L_NO_TRANSFORM
    if abs(basis_z.X) > LINK_AXIS_TOL or abs(basis_z.Y) > LINK_AXIS_TOL or basis_z.Z <= 0.0:
        return L_TILTED
    try:
        if abs(transform.Scale - 1.0) > 1.0e-9:
            return L_SCALED
    except Exception:
        pass
    return None


def build_category_filter(built_in_categories):
    category_list = List[DB.BuiltInCategory]()
    for built_in_category in built_in_categories:
        category_list.Add(built_in_category)
    return DB.ElementMulticategoryFilter(category_list)


def local_outline(source, host_min, host_max):
    """Volume di ricerca in coordinate della sorgente.

    Gli otto vertici del volume host vengono portati nel collegamento e se ne
    prende l'involucro allineato agli assi: un collegamento ruotato in pianta
    ha assi diversi da quelli dell'host.
    """
    corners = []
    for x in (host_min.X, host_max.X):
        for y in (host_min.Y, host_max.Y):
            for z in (host_min.Z, host_max.Z):
                local = source.to_local(DB.XYZ(x, y, z))
                if local is not None:
                    corners.append(local)
    if not corners:
        return None
    return DB.Outline(
        DB.XYZ(min(p.X for p in corners), min(p.Y for p in corners), min(p.Z for p in corners)),
        DB.XYZ(max(p.X for p in corners), max(p.Y for p in corners), max(p.Z for p in corners)))


# =========================================================================
# GEOMETRIA DI CONTROSOFFITTI E SOLAI
# =========================================================================

class SurfaceInfo(object):
    """Un controsoffitto o un solaio, con le facce lette una volta sola.

    Il bounding box e' letto alla raccolta e serve al filtro in pianta per
    ogni elemento; le facce si estraggono alla prima richiesta e restano in
    cache, perche' lo stesso solaio ricorre per tutti i dispositivi del
    locale.
    """

    def __init__(self, element, source, kind, bbox):
        self.element = element
        self.element_id = element.Id
        self.source = source
        self.kind = kind
        self.label = surface_label(element, source, kind)
        self.min_x = bbox.Min.X
        self.min_y = bbox.Min.Y
        self.min_z = bbox.Min.Z
        self.max_x = bbox.Max.X
        self.max_y = bbox.Max.Y
        self.max_z = bbox.Max.Z
        self.faces = None
        self.problem = None


def surface_label(element, source, kind):
    """"Ceiling: tipo" o "Floor: tipo", con il collegamento fra parentesi."""
    type_name = u'?'
    try:
        element_type = source.document.GetElement(element.GetTypeId())
        if element_type is not None:
            type_name = DB.Element.Name.GetValue(element_type)
        else:
            type_name = element_name(element)
    except Exception:
        pass
    prefix = u'Ceiling' if kind == K_CEILING else u'Floor'
    label = u'{}: {}'.format(prefix, type_name)
    if source.is_linked:
        label = u'[{}] {}'.format(source.link_name, label)
    return label


def demolished_reason(element, document):
    """Motivo di scarto se l'elemento e' demolito in una fase, altrimenti None."""
    try:
        phase_id = element.DemolishedPhaseId
    except Exception:
        return None
    if phase_id is None or phase_id == DB.ElementId.InvalidElementId:
        return None
    phase = document.GetElement(phase_id)
    return S_DEMOLISHED.format(element_name(phase) if phase is not None else u'?')


def iter_solids(geometry):
    """Solidi di un GeometryElement, istanze annidate comprese (famiglie locali)."""
    if geometry is None:
        return
    for item in geometry:
        if isinstance(item, DB.Solid):
            yield item
        elif isinstance(item, DB.GeometryInstance):
            try:
                nested = item.GetInstanceGeometry()
            except Exception:
                nested = None
            for solid in iter_solids(nested):
                yield solid


def face_xy_bounds(face):
    """(min x, min y, max x, max y) del contorno della faccia, None se ignoto.

    Serve solo a evitare Face.Intersect sulle facce che stanno altrove in
    pianta: il bordo tassellato racchiude la faccia per le facce piane, che
    sono la quasi totalita' su solai e controsoffitti.
    """
    xs = []
    ys = []
    try:
        for loop in face.EdgeLoops:
            for edge in loop:
                for point in edge.Tessellate():
                    xs.append(point.X)
                    ys.append(point.Y)
    except Exception:
        return None
    if not xs:
        return None
    return min(xs), min(ys), max(xs), max(ys)


def surface_faces(info):
    """Facce utilizzabili della superficie, [(faccia, limiti in pianta)].

    Si leggono TUTTE le facce non verticali, non solo quelle rivolte verso il
    basso: la superficie inferiore e' quella della quota piu' bassa colpita
    dalla retta, e cosi' non serve fidarsi del verso della normale (un
    controsoffitto Basic senza spessore puo' esporre una sola faccia).
    """
    if info.faces is not None or info.problem is not None:
        return info.faces

    options = DB.Options()
    options.ComputeReferences = False
    options.IncludeNonVisibleObjects = False
    options.DetailLevel = DB.ViewDetailLevel.Fine

    try:
        geometry = info.element.get_Geometry(options)
    except Exception:
        geometry = None
    if geometry is None:
        info.problem = S_NO_GEOMETRY
        return None

    faces = []
    for solid in iter_solids(geometry):
        try:
            solid_faces = solid.Faces
        except Exception:
            continue
        for face in solid_faces:
            if isinstance(face, DB.PlanarFace):
                try:
                    if abs(face.FaceNormal.Z) < VERTICAL_FACE_TOL:
                        continue
                except Exception:
                    pass
            faces.append((face, face_xy_bounds(face)))

    if not faces:
        info.problem = S_NO_FACE
        return None
    info.faces = faces
    return faces


def face_hits_z(face, line):
    """Quote dei punti in cui la retta verticale interseca la faccia."""
    results = clr.Reference[DB.IntersectionResultArray]()
    try:
        outcome = face.Intersect(line, results)
    except Exception:
        return []
    if outcome != DB.SetComparisonResult.Overlap:
        return []
    array = results.Value
    if array is None:
        return []
    found = []
    try:
        for index in range(array.Size):
            found.append(array.get_Item(index).XYZPoint.Z)
    except Exception:
        pass
    return found


class Layer(object):
    """Tratto della retta verticale occupato da una superficie, in quote host."""

    def __init__(self, bottom, top, info):
        self.bottom = bottom
        self.top = top
        self.info = info


def surface_layer(info, local_point):
    """Layer della superficie sulla verticale del punto, None se non c'e'.

    La sonda e' una retta verticale per il punto che attraversa tutto il
    bounding box della superficie: cosi' si leggono sia la faccia inferiore
    sia quella superiore, che serve a sapere se due istanze sono a contatto.
    Se la verticale cade in un'apertura (cavedio, botola) non c'e' nessuna
    intersezione, ed e' giusto cosi'.
    """
    faces = surface_faces(info)
    if not faces:
        return None
    x = local_point.X
    y = local_point.Y
    try:
        line = DB.Line.CreateBound(DB.XYZ(x, y, info.min_z - PROBE_PAD),
                                   DB.XYZ(x, y, info.max_z + PROBE_PAD))
    except Exception:
        return None

    heights = []
    for face, bounds in faces:
        if bounds is not None:
            if (x < bounds[0] - XY_EPS or x > bounds[2] + XY_EPS
                    or y < bounds[1] - XY_EPS or y > bounds[3] + XY_EPS):
                continue
        heights.extend(face_hits_z(face, line))
    if not heights:
        return None
    offset = info.source.z_offset
    return Layer(min(heights) + offset, max(heights) + offset, info)


class Package(object):
    """Piu' superfici impilate a contatto: un solo intradosso, il piu' basso."""

    def __init__(self, layer):
        self.layers = [layer]
        self.bottom = layer.bottom
        self.top = layer.top

    def add(self, layer):
        self.layers.append(layer)
        self.bottom = min(self.bottom, layer.bottom)
        self.top = max(self.top, layer.top)

    @property
    def lowest(self):
        return min(self.layers, key=lambda layer: layer.bottom)


def merge_layers(layers):
    """Pacchetti di layer a contatto (scarto <= STACK_GAP), dal basso."""
    packages = []
    for layer in sorted(layers, key=lambda item: item.bottom):
        if packages and layer.bottom <= packages[-1].top + STACK_GAP:
            packages[-1].add(layer)
        else:
            packages.append(Package(layer))
    return packages


# =========================================================================
# RACCOLTA DELLE SUPERFICI
# =========================================================================

def collect_in(source, host_min, host_max, context):
    """Controsoffitti e solai della sorgente nel volume indicato."""
    outline = local_outline(source, host_min, host_max)
    if outline is None:
        return
    categories = [CEILING_CATEGORY, FLOOR_CATEGORY]
    label = u'"{}"'.format(source.link_name) if source.is_linked else u'this model'
    try:
        elements = DB.FilteredElementCollector(source.document)\
            .WherePasses(build_category_filter(categories))\
            .WhereElementIsNotElementType()\
            .WherePasses(DB.BoundingBoxIntersectsFilter(outline))\
            .ToElements()
    except Exception as error:
        warn(u'Collecting ceilings and floors in {} failed: {}'.format(label, error))
        return

    for element in elements:
        _name, key = category_of(element)
        kind = K_CEILING if key == CEILING_KEY else K_FLOOR
        reason = demolished_reason(element, source.document)
        if reason is not None:
            context.note_problem(source, element, kind, reason)
            continue
        try:
            bbox = element.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is None:
            context.note_problem(source, element, kind, S_NO_BBOX)
            continue
        source.surfaces.append(SurfaceInfo(element, source, kind, bbox))


def collect_surface_sources(host_min, host_max, options, context):
    """Documento host piu' i modelli collegati utilizzabili.

    Una sola raccolta per documento sul volume occupato dalla selezione; il
    filtro per singolo elemento avviene poi in Python sui bounding box gia'
    letti. Controsoffitti e solai si raccolgono sempre insieme: i solai
    delimitano il piano (un controsoffitto sopra il solaio di chiusura
    appartiene al piano superiore) e i controsoffitti decidono il bersaglio.
    """
    sources = []
    host = SurfaceSource(0, doc, None, None)
    collect_in(host, host_min, host_max, context)
    sources.append(host)

    if not options.include_links:
        return sources

    try:
        links = list(DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).ToElements())
    except Exception as error:
        warn(u'The list of linked models cannot be read: {}'.format(error))
        return sources

    for link in links:
        name = link_display_name(link)
        try:
            link_doc = link.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            warn(u'Link "{}" ignored: {}.'.format(name, L_NOT_LOADED))
            continue
        try:
            transform = link.GetTotalTransform()
        except Exception:
            transform = None
        if transform is None:
            warn(u'Link "{}" ignored: {}.'.format(name, L_NO_TRANSFORM))
            continue
        problem = link_transform_problem(transform)
        if problem is not None:
            warn(u'Link "{}" ignored: {}.'.format(name, problem))
            continue

        source = SurfaceSource(element_id_value(link.Id), link_doc, transform, name)
        collect_in(source, host_min, host_max, context)
        if source.surfaces:
            sources.append(source)

    return sources


# =========================================================================
# ANALISI (nessuna transazione)
# =========================================================================

class AlignOptions(object):
    """Scelte effettuate dall'utente nella finestra di dialogo."""

    def __init__(self, category_keys, tolerance_mm, include_links, skip_connected, dry_run):
        self.category_keys = category_keys
        self.tolerance_mm = tolerance_mm
        self.tolerance_internal = mm_to_internal(tolerance_mm)
        # Semilunghezza della retta di analisi: la retta e' lunga
        # LINE_FACTOR volte la tolleranza e centrata sul punto.
        self.half_line = self.tolerance_internal * LINE_FACTOR / 2.0
        self.include_links = include_links
        self.skip_connected = skip_connected
        self.dry_run = dry_run


class Target(object):
    """Superficie scelta per un elemento, con la distanza misurata."""

    def __init__(self, kind, package, point_z, half_line):
        self.kind = kind
        self.package = package
        self.layer = package.lowest
        self.z = package.bottom
        # Positiva se il punto sta sopra la superficie inferiore.
        self.distance = point_z - package.bottom
        self.inside = package.bottom - POSITION_TOL <= point_z <= package.top + POSITION_TOL
        self.above = point_z > package.top + POSITION_TOL
        # La retta interseca la superficie inferiore se questa dista dal
        # punto non piu' di mezza retta. Un punto dentro lo spessore vale
        # sempre, come nei muri: una superficie piu' spessa della tolleranza
        # non deve lasciare il dispositivo sepolto.
        self.on_line = self.inside or abs(self.distance) <= half_line

    @property
    def label(self):
        return self.layer.info.label

    @property
    def surface_id(self):
        return self.layer.info.element_id

    @property
    def is_linked(self):
        return self.layer.info.source.is_linked


class TargetSearch(object):
    """Esito della ricerca del bersaglio per un elemento."""

    def __init__(self):
        self.target = None
        self.ceiling = None             # controsoffitto del piano piu' vicino, se c'e'
        self.second = None              # secondo controsoffitto candidato
        self.closing = None             # pacchetto del solaio di chiusura
        self.attached = 0               # rilevatori: controsoffitti incollati al solaio
        self.reason = None              # motivo se target e' None


class PlannedMove(object):
    """Uno spostamento gia' calcolato: la transazione legge solo questi campi."""

    PLANNED = 'planned'
    ALREADY_OK = 'already_ok'

    def __init__(self):
        self.element = None
        self.element_id = None
        self.category_name = u'-'
        self.category_key = None
        self.type_name = u'-'
        self.level_name = u'-'
        self.level_z = None
        self.kind = u'-'
        self.surface_label = u'-'
        self.surface_id = None
        self.surface_is_linked = False
        self.point_before = None
        self.target_z = None
        self.distance = 0.0
        self.connected = 0
        self.note = None
        self.status = PlannedMove.PLANNED
        self.needs_move = False
        self.applied = False
        self.z_after = None
        self.error = None

        # Rilevatori da controsoffitto: valore da scrivere in dz_Buzzer,
        # valore letto prima, e se e' misurato o preso dalla finestra.
        self.is_buzzer = False
        self.buzzer_value = None
        self.buzzer_before = None
        self.buzzer_measured = False
        self.needs_param = False
        self.buzzer_after = None

    @property
    def move(self):
        return self.target_z - self.point_before.Z


class SkippedElement(object):
    """Un elemento escluso, con il motivo e, se nota, la superficie piu' vicina."""

    def __init__(self, element, category_name, category_key, reason, screened=False):
        self.element = element
        self.element_id = element.Id if element is not None else None
        self.category_name = category_name
        self.category_key = category_key
        self.reason = reason
        self.screened = screened
        self.target = None


class OverTolerance(object):
    """Un elemento con una superficie bersaglio oltre la tolleranza."""

    def __init__(self, element, category_name, category_key, target, tolerance, note,
                 nearest_distance=None):
        self.element = element
        self.element_id = element.Id
        self.category_name = category_name
        self.category_key = category_key
        self.target = target
        # Per i rilevatori da controsoffitto la superficie piu' vicina puo'
        # essere il controsoffitto e non il bersaglio (il solaio).
        if nearest_distance is None:
            nearest_distance = abs(target.distance)
        self.excess = nearest_distance - tolerance
        self.note = note


class PlanResult(object):

    def __init__(self):
        self.planned = []
        self.already_ok = []
        self.skipped = []
        self.over_tolerance_items = []
        self.per_category = {}
        self.surface_problems = []


class PlanContext(object):
    """Stato condiviso dalla fase di analisi."""

    def __init__(self, active_design_option_id):
        self.sources = []
        self.storeys = {}
        self.active_design_option_id = active_design_option_id
        # Indicizzati per (sorgente, id): una superficie scartata ricorre per
        # molti elementi e nel resoconto deve comparire una volta sola.
        self.surface_problems = {}

    def note_problem(self, source, element, kind, reason):
        key = (source.key, element_id_value(element.Id))
        self.surface_problems[key] = (source, element, kind, reason)

    def layers_at(self, point, low_z, high_z):
        """(layer dei controsoffitti, layer dei solai) sulla verticale del punto.

        low_z / high_z delimitano in quote host la fascia in cui cercare; i
        bounding box sono nelle coordinate della propria sorgente, quindi la
        fascia viene traslata della quota del collegamento.
        """
        ceilings = []
        floors = []
        for source in self.sources:
            if not source.surfaces:
                continue
            local = source.to_local(point)
            if local is None:
                continue
            low = low_z - source.z_offset
            high = high_z - source.z_offset
            for info in source.surfaces:
                if (local.X < info.min_x - XY_EPS or local.X > info.max_x + XY_EPS
                        or local.Y < info.min_y - XY_EPS or local.Y > info.max_y + XY_EPS
                        or info.max_z < low or info.min_z > high):
                    continue
                layer = surface_layer(info, local)
                if info.problem is not None:
                    self.note_problem(source, info.element, info.kind, info.problem)
                if layer is None:
                    continue
                if info.kind == K_CEILING:
                    ceilings.append(layer)
                else:
                    floors.append(layer)
        return ceilings, floors


def search_window(point_z, storey, half_line):
    """Fascia verticale in cui cercare le superfici per un elemento."""
    low = storey.bound if storey.bound is not None else point_z - half_line
    return low - SEARCH_MARGIN, point_z + half_line + SEARCH_MARGIN


def above_storey_floor(package, storey):
    """True se il pacchetto sta sopra il livello di riferimento."""
    if storey.bound is None:
        return True
    return package.bottom > storey.bound + LEVEL_TOL


def soffit_package(closing, ceiling_packages):
    """(intradosso, controsoffitti incollati) per un rilevatore da controsoffitto.

    Si parte dal solaio di chiusura e si scende finche' c'e' un pacchetto di
    controsoffitti la cui faccia superiore coincide (scarto <= STACK_GAP) con
    la faccia inferiore raggiunta: e' una finitura o un isolante modellato
    come Ceiling, e l'intradosso diventa la sua faccia inferiore.
    """
    soffit = Package(closing.layers[0])
    for layer in closing.layers[1:]:
        soffit.add(layer)
    attached = []
    for package in sorted(ceiling_packages, key=lambda p: -p.top):
        if package.top >= soffit.bottom - STACK_GAP:
            for layer in package.layers:
                soffit.add(layer)
            attached.append(package)
    return soffit, attached


def find_target(point, storey, options, context, buzzer=False):
    """Superficie bersaglio di un elemento.

    1. Si leggono, sulla verticale del punto, tutti i solai e i
       controsoffitti della fascia di ricerca, e si fondono in pacchetti le
       istanze a contatto: vale per i solai e per i controsoffitti fatti di
       piu' istanze, e il bersaglio e' sempre la faccia inferiore del
       pacchetto.
    2. Il solaio di chiusura e' il pacchetto di solai piu' basso fra quelli
       sopra il livello di riferimento.
    3. I controsoffitti del piano sono quelli sopra il livello e sotto il
       solaio di chiusura: uno piu' in alto appartiene al piano superiore.
    4. Se nel piano c'e' un controsoffitto, il bersaglio e' il controsoffitto
       (il piu' vicino al punto) ANCHE SE e' fuori tolleranza: in quel caso
       l'elemento viene saltato, non portato all'intradosso del solaio,
       perche' finirebbe nel plenum sopra il controsoffitto.
    5. Altrimenti il bersaglio e' il solaio di chiusura.

    Un rilevatore da controsoffitto (buzzer=True) va invece sempre
    all'INTRADOSSO: il solaio di chiusura piu' i controsoffitti a contatto con
    la sua faccia inferiore (una finitura, uno strato di isolante modellati
    come Ceiling), quindi la faccia inferiore del controsoffitto quando questo
    e' incollato al solaio. Il controsoffitto vero, quello che giustifica il
    rilevatore, e' uno STACCATO dall'intradosso: va in search.ceiling e serve
    a misurare dz_Buzzer. Se non c'e', search.ceiling resta None e lo script
    segnala il rilevatore come non adeguato.
    """
    search = TargetSearch()
    low_z, high_z = search_window(point.Z, storey, options.half_line)
    ceiling_layers, floor_layers = context.layers_at(point, low_z, high_z)

    floor_packages = [p for p in merge_layers(floor_layers) if above_storey_floor(p, storey)]
    closing = floor_packages[0] if floor_packages else None
    search.closing = closing

    ceiling_packages = [p for p in merge_layers(ceiling_layers)
                        if above_storey_floor(p, storey)
                        and (closing is None or p.bottom < closing.bottom)]

    if buzzer:
        if closing is None:
            search.reason = R_BUZZER_NO_SLAB
            return search
        soffit, attached = soffit_package(closing, ceiling_packages)
        search.attached = len(attached)
        search.target = Target(K_CEILING if attached else K_FLOOR, soffit, point.Z,
                               options.half_line)
        ceiling_packages = [p for p in ceiling_packages if not any(p is a for a in attached)]

    ceilings = [Target(K_CEILING, p, point.Z, options.half_line) for p in ceiling_packages]
    ceilings.sort(key=lambda t: (0.0 if t.inside else abs(t.distance), t.z))
    if ceilings:
        search.ceiling = ceilings[0]
    if len(ceilings) > 1:
        search.second = ceilings[1]
    if buzzer:
        return search

    if search.ceiling is not None:
        search.target = search.ceiling
    elif closing is not None:
        search.target = Target(K_FLOOR, closing, point.Z, options.half_line)
    else:
        search.reason = R_NO_SURFACE
    return search


def is_buzzer_family(element):
    """True se il family name contiene tutte le BUZZER_FAMILY_TOKENS."""
    try:
        family_name = DB.Element.Name.GetValue(element.Symbol.Family)
    except Exception:
        return False
    lowered = (family_name or u'').lower()
    return all(token.lower() in lowered for token in BUZZER_FAMILY_TOKENS)


def is_length(parameter):
    """True se il parametro e' una lunghezza (Revit 2022+, con riserva)."""
    if parameter.StorageType != DB.StorageType.Double:
        return False
    try:
        return parameter.Definition.GetDataType() == DB.SpecTypeId.Length
    except Exception:
        return True


def buzzer_parameter(element):
    """(parametro dz_Buzzer, None) oppure (None, motivo per cui non si usa).

    Cercato solo sull'istanza, come lo screening: la distanza varia da un
    dispositivo all'altro. Un parametro di sola lettura (pilotato da una
    formula) va segnalato prima, non scoperto a transazione aperta.
    """
    try:
        parameter = element.LookupParameter(BUZZER_PARAM)
    except Exception:
        parameter = None
    if parameter is None:
        try:
            symbol = doc.GetElement(element.GetTypeId())
            if symbol is not None and symbol.LookupParameter(BUZZER_PARAM) is not None:
                return None, R_BUZZER_TYPE
        except Exception:
            pass
        return None, R_BUZZER_MISSING
    if not is_length(parameter):
        return None, R_BUZZER_NOT_LENGTH
    if parameter.IsReadOnly:
        return None, R_BUZZER_READ_ONLY
    return parameter, None


def connected_connectors(element):
    """Numero di connettori collegati, 0 se la famiglia non ne ha."""
    try:
        manager = element.MEPModel.ConnectorManager
    except Exception:
        return 0
    if manager is None:
        return 0
    count = 0
    try:
        for connector in manager.Connectors:
            try:
                if connector.IsConnected:
                    count += 1
            except Exception:
                continue
    except Exception:
        return count
    return count


def host_skip_reason(element):
    """Motivo di esclusione legato all'host, None se l'elemento e' libero.

    Un livello come host non conta: e' il caso normale delle famiglie basate
    sul livello, che si spostano in quota cambiando lo scostamento.
    """
    try:
        host = element.Host
    except Exception:
        return None
    if host is None or isinstance(host, DB.Level):
        return None
    if isinstance(host, DB.RevitLinkInstance):
        return R_HOSTED_LINK
    host_category, host_key = category_of(host)
    if host_key == CEILING_KEY:
        return R_HOSTED_CEILING.format(element_id_value(host.Id))
    return R_HOSTED_OTHER.format(host_category, element_id_value(host.Id))


def edit_skip_reason(element, active_design_option_id):
    """Motivo di esclusione legato allo stato di modificabilita'.

    I membri dei gruppi vengono SALTATI, mai modificati: spostarne uno
    cambierebbe tutte le istanze del gruppo.
    """
    try:
        if element.Pinned:
            return R_PINNED
    except Exception:
        pass
    try:
        group_id = element.GroupId
        if group_id is not None and group_id != DB.ElementId.InvalidElementId:
            group = doc.GetElement(group_id)
            group_name = element_name(group) if group is not None else u'?'
            return R_GROUP.format(group_name)
    except Exception:
        pass
    try:
        if element.SuperComponent is not None:
            return R_SUBCOMPONENT
    except Exception:
        pass
    try:
        design_option = element.DesignOption
        if design_option is not None:
            if element_id_value(design_option.Id) != element_id_value(active_design_option_id):
                return R_DESIGN_OPTION
    except Exception:
        pass
    # GetCheckoutStatus va chiamato fuori transazione.
    try:
        if doc.IsWorkshared:
            status = DB.WorksharingUtils.GetCheckoutStatus(doc, element.Id)
            if status == DB.CheckoutStatus.OwnedByOtherUser:
                return R_BORROWED
    except Exception:
        pass
    return None


def skipped_with_target(element, category_name, category_key, reason, search):
    skipped = SkippedElement(element, category_name, category_key, reason)
    if search is not None:
        skipped.target = search.target
    return skipped


def stack_note(package, kind):
    """Nota per un intradosso formato da piu' istanze impilate."""
    if len(package.layers) < 2:
        return None
    kinds = set(layer.info.kind for layer in package.layers)
    if len(kinds) > 1:
        noun = u'floors and ceilings'      # intradosso di un rilevatore con finitura
    else:
        noun = u'ceilings' if kind == K_CEILING else u'floors'
    linked = len(set(layer.info.source.key for layer in package.layers))
    note = u'soffit of {} stacked {} (lowest: {})'.format(
        len(package.layers), noun, package.lowest.info.label)
    if linked > 1:
        note += u', from {} models'.format(linked)
    return note


def plan_element(element, options, context):
    """Decisione completa per un singolo elemento indicato dall'utente."""
    category_name, category_key = category_of(element)

    point = insertion_point(element)
    if point is None:
        return SkippedElement(element, category_name, category_key, R_NO_POINT)

    # Screening per primo: un dispositivo montato a parete non e' affare di
    # questo strumento, e non vale la pena di leggerne la geometria intorno.
    reason = screening_reason(element)
    if reason is not None:
        return SkippedElement(element, category_name, category_key, reason, screened=True)

    storey = context.storeys.get(element_id_value(element.Id))
    if storey is None:
        storey = Storey(None, None, u'no level found: no lower limit')

    buzzer = is_buzzer_family(element)

    # Il bersaglio si calcola prima delle guardie, cosi' anche un elemento
    # saltato (bloccato, in un gruppo) porta nel resoconto la distanza
    # misurata: uno a 12 mm vale la pena di sbloccarlo, uno a 400 no.
    search = find_target(point, storey, options, context, buzzer)

    for reason in (host_skip_reason(element),
                   edit_skip_reason(element, context.active_design_option_id)):
        if reason is not None:
            return skipped_with_target(element, category_name, category_key, reason, search)

    connected = connected_connectors(element)
    if options.skip_connected and connected > 0:
        return skipped_with_target(element, category_name, category_key,
                                   R_CONNECTED.format(connected), search)

    buzzer_param = None
    if buzzer:
        if search.target is None:
            return SkippedElement(element, category_name, category_key, search.reason)
        # Il rilevatore da controsoffitto si mette solo sotto un controsoffitto
        # staccato dal solaio: senza, la famiglia e' quella sbagliata.
        if search.ceiling is None:
            return skipped_with_target(element, category_name, category_key,
                                       R_BUZZER_NO_CEILING, search)
        # Senza un dz_Buzzer scrivibile il rilevatore non si sposta: portato
        # al solaio, la sua parte visibile resterebbe lassu' invece che sul
        # controsoffitto.
        buzzer_param, reason = buzzer_parameter(element)
        if reason is not None:
            return skipped_with_target(element, category_name, category_key, reason, search)

    target = search.target
    if target is None:
        return SkippedElement(element, category_name, category_key, search.reason)

    notes = []
    if storey.note:
        notes.append(storey.note)

    ceiling = search.ceiling
    if buzzer:
        # Il rilevatore va all'intradosso, ma di solito e' modellato
        # all'altezza del controsoffitto: la retta vale se interseca l'uno o
        # l'altro.
        if not (target.on_line or ceiling.on_line):
            nearest = min(abs(target.distance), abs(ceiling.distance))
            notes.append(u'false ceiling detector: neither the soffit nor the false ceiling is '
                         u'within tolerance')
            return OverTolerance(element, category_name, category_key, target,
                                 options.half_line, u'; '.join(notes), nearest)
    elif not target.on_line:
        if target.kind == K_CEILING and search.closing is not None:
            slab = Target(K_FLOOR, search.closing, point.Z, options.half_line)
            if slab.on_line:
                notes.append(u'slab soffit within tolerance ({}) not used: the ceiling '
                             u'comes first'.format(format_signed_mm(-slab.distance)))
        return OverTolerance(element, category_name, category_key, target,
                             options.half_line, u'; '.join(notes) if notes else None)

    record = PlannedMove()
    record.element = element
    record.element_id = element.Id
    record.category_name = category_name
    record.category_key = category_key
    record.type_name = type_name_of(element)
    record.level_name = storey.level_name
    record.level_z = storey.bound
    record.kind = target.kind
    record.surface_label = target.label
    record.surface_id = target.surface_id
    record.surface_is_linked = target.is_linked
    record.point_before = point
    record.target_z = target.z
    record.distance = target.distance
    record.connected = connected
    record.needs_move = abs(target.distance) > POSITION_TOL

    if buzzer:
        # dz_Buzzer e' sempre misurato: dall'intradosso alla faccia
        # inferiore del controsoffitto staccato.
        record.is_buzzer = True
        record.buzzer_before = buzzer_param.AsDouble()
        record.buzzer_value = target.z - ceiling.z
        record.buzzer_measured = True
        notes.append(u'false ceiling detector: dz_Buzzer measured to {}'.format(ceiling.label))
        if search.attached:
            notes.append(u'soffit includes a ceiling in contact with the slab (finish or '
                         u'insulation): aligned to its bottom face')
        note = stack_note(ceiling.package, K_CEILING)
        if note:
            notes.append(u'false ceiling ' + note)
        record.needs_param = abs(record.buzzer_value - record.buzzer_before) > POSITION_TOL

    if not record.needs_move and not record.needs_param:
        record.status = PlannedMove.ALREADY_OK

    noun = u'ceiling' if target.kind == K_CEILING else u'slab'
    if target.inside and abs(target.distance) > POSITION_TOL:
        notes.append(u'insertion point was inside the {}'.format(noun))
    elif target.above:
        notes.append(u'element was above the {}: moved down through it'.format(noun))

    note = stack_note(target.package, target.kind)
    if note:
        notes.append(note)

    # Due controsoffitti quasi equidistanti: conta sia per il bersaglio sia
    # per la misura di dz_Buzzer, che usa lo stesso controsoffitto.
    second = search.second
    if ceiling is not None and second is not None and second.on_line:
        gap = abs(second.distance) - abs(ceiling.distance)
        if gap <= AMBIGUITY_TOL:
            notes.append(u'ambiguous: second ceiling at {}'.format(
                format_signed_mm(-second.distance)))

    record.note = u'; '.join(notes) if notes else None
    return record


def build_plan(elements, options, context):
    """Analisi di tutti gli elementi indicati. Nessuna transazione.

    Ogni elemento selezionato finisce in esattamente una delle quattro liste
    del risultato: nessuna esclusione silenziosa.
    """
    result = PlanResult()
    selected_keys = set(options.category_keys)

    with forms.ProgressBar(title='Analysing... ({value} of {max_value})',
                           cancellable=True) as progress:
        total = len(elements)
        for index, element in enumerate(elements):
            if progress.cancelled:
                raise UserWarning(u'cancelled')
            progress.update_progress(index + 1, total)

            category_name, category_key = category_of(element)
            if category_key is None or category_key not in MEP_CATEGORY_KEYS:
                result.skipped.append(SkippedElement(
                    element, category_name, category_key, R_CATEGORY_NOT_HANDLED))
                continue
            if category_key not in selected_keys:
                result.skipped.append(SkippedElement(
                    element, category_name, category_key, R_CATEGORY_NOT_SELECTED))
                continue

            try:
                outcome = plan_element(element, options, context)
            except Exception as error:
                result.skipped.append(SkippedElement(
                    element, category_name, category_key,
                    R_ANALYSIS_ERROR.format(u'{}'.format(error)[:160])))
                continue

            if isinstance(outcome, SkippedElement):
                result.skipped.append(outcome)
            elif isinstance(outcome, OverTolerance):
                result.over_tolerance_items.append(outcome)
            elif outcome.status == PlannedMove.ALREADY_OK:
                result.already_ok.append(outcome)
            else:
                result.planned.append(outcome)

    result.surface_problems = list(context.surface_problems.values())
    result.per_category = summarise_per_category(result, options)
    return result


def summarise_per_category(result, options):
    """Conteggi per categoria, per la tabella di riepilogo."""
    summary = {}

    def bucket(key, name):
        if key not in summary:
            summary[key] = {'name': name, 'selected': 0, 'moved': 0, 'ok': 0,
                            'screened': 0, 'skipped': 0, 'over': 0}
        return summary[key]

    for record in result.planned:
        bucket(record.category_key, record.category_name)['moved'] += 1
    for record in result.already_ok:
        bucket(record.category_key, record.category_name)['ok'] += 1
    for record in result.skipped:
        row = bucket(record.category_key, record.category_name)
        row['screened' if record.screened else 'skipped'] += 1
    for record in result.over_tolerance_items:
        bucket(record.category_key, record.category_name)['over'] += 1
    for row in summary.values():
        row['selected'] = row['moved'] + row['ok'] + row['screened'] + row['skipped'] + row['over']
    return summary


# =========================================================================
# GESTIONE DEI FAILURE DI REVIT
# =========================================================================

class AlignFailurePreprocessor(DB.IFailuresPreprocessor):
    """Registra gli avvisi di Revit senza applicare nessuna risoluzione.

    Non si usa il FailureSwallower di pyRevit: la sua lista RESOLUTION_TYPES
    include UnlockConstraints e DeleteElements, quindi sbloccherebbe in
    silenzio quote e allineamenti dell'utente. Qui vengono eliminati soltanto
    i warning privi di risoluzioni (aprirebbero un popup e fermerebbero il
    lotto) e si chiede il rollback in presenza di errori.
    """

    def __init__(self):
        self.messages = []

    def PreprocessFailures(self, failuresAccessor):
        has_error = False
        try:
            failures = failuresAccessor.GetFailureMessages()
        except Exception:
            return DB.FailureProcessingResult.Continue

        for failure in failures:
            description = u'-'
            try:
                description = failure.GetDescriptionText()
            except Exception:
                pass
            severity = None
            try:
                severity = failure.GetSeverity()
            except Exception:
                pass
            ids = []
            try:
                for element_id in failure.GetFailingElementIds():
                    ids.append(element_id)
            except Exception:
                pass

            is_error = severity == DB.FailureSeverity.Error
            if is_error:
                has_error = True
            record = (description, ids, u'error' if is_error else u'warning')
            if record not in self.messages:
                self.messages.append(record)

            try:
                has_resolutions = failure.HasResolutions()
            except Exception:
                has_resolutions = True
            if severity == DB.FailureSeverity.Warning and not has_resolutions:
                try:
                    failuresAccessor.DeleteWarning(failure)
                except Exception:
                    pass

        if has_error:
            return DB.FailureProcessingResult.ProceedWithRollBack
        return DB.FailureProcessingResult.Continue


# =========================================================================
# APPLICAZIONE DELLE MODIFICHE (transazione)
# =========================================================================

def apply_record(record):
    """Traslazione verticale verso la quota bersaglio ASSOLUTA.

    Il vettore si ricalcola dal punto corrente invece di riusare quello
    dell'analisi: un elemento mosso nel frattempo (per esempio trascinato da
    un elemento connesso) finisce comunque alla quota giusta. X e Y del
    vettore sono sempre zero. Per un rilevatore da controsoffitto dz_Buzzer
    si scrive nella stessa SubTransaction: spostamento e parametro riescono
    o falliscono insieme. Una SubTransaction per elemento: un errore lascia
    intatto quell'elemento senza annullare gli altri.
    """
    sub = DB.SubTransaction(doc)
    sub.Start()
    try:
        current = insertion_point(record.element)
        if current is None:
            current = record.point_before
        delta = record.target_z - current.Z
        if abs(delta) > POSITION_TOL:
            DB.ElementTransformUtils.MoveElement(doc, record.element_id, DB.XYZ(0.0, 0.0, delta))
        if record.is_buzzer and record.needs_param:
            parameter = record.element.LookupParameter(BUZZER_PARAM)
            if parameter is None or not parameter.Set(record.buzzer_value):
                raise Exception(u'"{}" cannot be set'.format(BUZZER_PARAM))
        sub.Commit()
        record.applied = True
        return True
    except Exception as error:
        try:
            sub.RollBack()
        except Exception:
            pass
        record.applied = False
        record.error = u'{}'.format(error)[:180]
        return False


def apply_all(planned):
    """Applica tutti i movimenti pianificati. Ritorna (riusciti, falliti)."""
    ok = 0
    failed = 0
    with forms.ProgressBar(title='Aligning... ({value} of {max_value})',
                           cancellable=True) as progress:
        total = len(planned)
        for index, record in enumerate(planned):
            if progress.cancelled:
                raise UserWarning(u'cancelled')
            progress.update_progress(index + 1, total)
            if apply_record(record):
                ok += 1
            else:
                failed += 1
    return ok, failed


def verify_applied(planned):
    """Rilegge la quota dopo il commit, quando il modello e' rigenerato.

    Alcune famiglie (vincolate al piano di lavoro, con parametri di quota
    bloccati) accettano MoveElement senza spostarsi davvero: il resoconto lo
    deve dire, invece di dichiararle allineate.
    """
    mismatched = 0
    for record in planned:
        if not record.applied:
            continue
        if record.is_buzzer:
            try:
                record.buzzer_after = record.element.LookupParameter(BUZZER_PARAM).AsDouble()
            except Exception:
                record.buzzer_after = None
        point = insertion_point(record.element)
        if point is not None:
            record.z_after = point.Z
        if record_problem(record) is not None:
            mismatched += 1
    return mismatched


def record_problem(record):
    """Testo del problema di un elemento applicato male, None se e' a posto."""
    if record.error:
        return u'error: {}'.format(record.error)
    if record.z_after is not None and abs(record.z_after - record.target_z) > POSITION_TOL:
        if record.level_z is not None:
            elevation = offset_cell(record.z_after, record.level_z)
        else:
            elevation = format_mm(record.z_after)
        return u'elevation after the move is {}: check the family'.format(elevation)
    if record.is_buzzer and record.buzzer_after is not None \
            and abs(record.buzzer_after - record.buzzer_value) > POSITION_TOL:
        return u'{} after the edit is {}: check the family'.format(
            BUZZER_PARAM, format_mm(record.buzzer_after))
    return None


# =========================================================================
# SELEZIONE E FINESTRA DI DIALOGO
# =========================================================================

class MEPSelectionFilter(UI.Selection.ISelectionFilter):
    """Ammette solo le categorie MEP gestite, confrontate per id."""

    def AllowElement(self, element):
        try:
            if element.Category is None:
                return False
            return element_id_value(element.Category.Id) in MEP_CATEGORY_KEYS
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return False


def elements_from_selection():
    """Elementi nella selezione corrente, definizioni di tipo escluse.

    Non si filtra per categoria: gli scarti avvengono nel piano, ognuno con
    il suo motivo nel resoconto.
    """
    elements = []
    try:
        for element in revit.get_selection():
            if element is None or isinstance(element, DB.ElementType):
                continue
            elements.append(element)
    except Exception:
        pass
    return elements


def pick_elements():
    """Selezione grafica degli elementi MEP. Vuota se l'utente preme Esc."""
    with forms.WarningBar(title='Select the MEP elements to align to the ceiling, then press Finish'):
        try:
            references = uidoc.Selection.PickObjects(
                UI.Selection.ObjectType.Element, MEPSelectionFilter(),
                'Select the MEP elements to align to the ceiling')
        except Exception:
            return []

    elements = []
    ids = List[DB.ElementId]()
    for reference in references:
        element = doc.GetElement(reference.ElementId)
        if element is None:
            continue
        elements.append(element)
        ids.Add(element.Id)
    # Gli elementi restano selezionati: rilanciando il comando dopo una
    # simulazione non serve ripetere la selezione grafica.
    if ids.Count:
        try:
            uidoc.Selection.SetElementIds(ids)
        except Exception:
            pass
    return elements


def resolve_elements():
    """(elementi, etichetta della fonte). Interrompe se non ce ne sono."""
    elements = elements_from_selection()
    if elements:
        return elements, u'from the current selection'
    elements = pick_elements()
    if not elements:
        script.exit()
    return elements, u'picked graphically'


def resolve_xaml_path(file_name):
    path = None
    try:
        path = script.get_bundle_file(file_name)
    except Exception:
        path = None
    if not path:
        path = op.join(op.dirname(__file__), file_name)
    if not op.isfile(path):
        forms.alert(u'Layout file not found:\n\n{}\n\nIt must sit in the same folder as the '
                    u'script.'.format(path), title=u'Missing layout', exitscript=True)
    return path


def load_settings():
    settings = {
        'tolerance_mm': DEFAULT_TOLERANCE_MM,
        'include_links': True,
        'skip_connected': False,
        'excluded_categories': u'',
    }
    try:
        config = script.get_config(CONFIG_SECTION)
        # La tolleranza era in cm fino alla versione 1.1 (chiave tolerance_cm).
        tolerance = config.get_option('tolerance_mm', None)
        if tolerance is None:
            tolerance = float(config.get_option('tolerance_cm', DEFAULT_TOLERANCE_MM / 10.0)) * 10.0
        tolerance = float(tolerance)
        if 0.0 < tolerance <= MAX_TOLERANCE_MM:
            settings['tolerance_mm'] = tolerance
        for key in ('include_links', 'skip_connected'):
            settings[key] = bool(config.get_option(key, settings[key]))
        settings['excluded_categories'] = u'{}'.format(config.get_option('excluded_categories', u''))
    except Exception:
        pass
    return settings


def save_settings(settings):
    try:
        config = script.get_config(CONFIG_SECTION)
        for key, value in settings.items():
            setattr(config, key, value)
        script.save_config()
    except Exception:
        pass


def buzzer_count(elements):
    """Rilevatori da controsoffitto nella selezione: (totale, di cui con Yes)."""
    total = 0
    flagged = 0
    for element in elements:
        _name, key = category_of(element)
        if key not in MEP_CATEGORY_KEYS or not is_buzzer_family(element):
            continue
        total += 1
        if screening_reason(element) is None:
            flagged += 1
    return total, flagged


def count_per_category(elements):
    """[(nome BIC, etichetta, conteggio, montati a soffitto)] per la finestra.

    Compaiono solo le categorie presenti nella selezione: una riga a zero non
    e' una scelta che l'utente possa fare.
    """
    counts = {}
    flagged = {}
    for element in elements:
        _name, key = category_of(element)
        if key is None:
            continue
        counts[key] = counts.get(key, 0) + 1
        if key in MEP_CATEGORY_KEYS and screening_reason(element) is None:
            flagged[key] = flagged.get(key, 0) + 1

    info = []
    for name, built_in_category in MEP_CATEGORIES:
        key = bic_key(built_in_category)
        count = counts.get(key, 0)
        if not count:
            continue
        info.append((name, category_display_name(built_in_category), count, flagged.get(key, 0)))
    info.sort(key=lambda row: row[1])
    return info


def out_of_scope_count(elements):
    total = 0
    for element in elements:
        _name, key = category_of(element)
        if key is None or key not in MEP_CATEGORY_KEYS:
            total += 1
    return total


def ask_options(elements, source_label, category_info):
    """Mostra la finestra e ritorna AlignOptions, None se annullata."""
    settings = load_settings()
    excluded = set(name for name in settings['excluded_categories'].split(u'|') if name)

    flagged_total = sum(row[3] for row in category_info)
    out_of_scope = out_of_scope_count(elements)
    # Solo righe di stato brevi: le spiegazioni stanno nel README.
    detail = u'{} with {} = Yes: only those are aligned.'.format(flagged_total, INSTALLATION_PARAM)
    if out_of_scope:
        detail += u' {} outside the handled categories.'.format(out_of_scope)

    buzzers, buzzers_flagged = buzzer_count(elements)
    if buzzers:
        buzzer_text = u'{} in the selection, {} with {} = Yes.'.format(
            buzzers, buzzers_flagged, INSTALLATION_PARAM)
    else:
        buzzer_text = u'No false ceiling detector in the selection.'

    dialog_info = {
        'selection_text': u'{} elements selected ({}).'.format(len(elements), source_label),
        'detail_text': detail,
        'categories': [(name, label, count, flagged, name not in excluded)
                       for name, label, count, flagged in category_info],
        'tolerance_mm': settings['tolerance_mm'],
        'max_tolerance_mm': MAX_TOLERANCE_MM,
        'line_factor': LINE_FACTOR,
        'buzzer_text': buzzer_text,
        'include_links': settings['include_links'],
        'skip_connected': settings['skip_connected'],
    }
    chosen = ceiling_ui.show_dialog(resolve_xaml_path(XAML_FILE_NAME), dialog_info)
    if chosen is None:
        return None

    # Le categorie tolte restano escluse anche alla prossima apertura; quelle
    # non presenti in questa selezione mantengono lo stato precedente.
    shown = set(row[0] for row in category_info)
    excluded = (excluded - shown) | (shown - set(chosen['categories']))
    save_settings({
        'tolerance_mm': chosen['tolerance_mm'],
        'include_links': chosen['include_links'],
        'skip_connected': chosen['skip_connected'],
        'excluded_categories': u'|'.join(sorted(excluded)),
    })

    by_name = dict(MEP_CATEGORIES)
    keys = [bic_key(by_name[name]) for name in chosen['categories'] if name in by_name]
    return AlignOptions(keys, chosen['tolerance_mm'], chosen['include_links'],
                        chosen['skip_connected'], chosen['dry_run'])


# =========================================================================
# RESOCONTO
# =========================================================================

def surface_cell(surface_id, label, is_linked):
    """Cella "superficie": gli elementi collegati non sono linkify-abili."""
    if surface_id is None:
        return u'-'
    if is_linked:
        return u'{} (id {})'.format(label, element_id_value(surface_id))
    return u'{} {}'.format(output.linkify(surface_id), label)


def offset_cell(z, level_z):
    """Quota rispetto al livello di riferimento usato, in mm."""
    if z is None or level_z is None:
        return u'-'
    return format_signed_mm(z - level_z)


def format_buzzer(value):
    return u'-' if value is None else format_mm(value)


def print_header(options, elements, source_label, result):
    if options.dry_run:
        output.print_md(u'# MEP alignment to ceilings - simulation')
        output.print_md(u'**No change has been applied to the model.** '
                        u'The figures below are the result that would be produced.')
    else:
        output.print_md(u'# MEP alignment to ceilings - report')

    screened = sum(1 for r in result.skipped if r.screened)
    lines = [
        u'- Elements selected: **{}** ({})'.format(len(elements), source_label),
        u'- Screening: only instances with **{} = Yes**; excluded by it: **{}**'.format(
            INSTALLATION_PARAM, screened),
        u'- Target: the ceiling bottom face where the storey has a ceiling on the vertical of '
        u'the device, otherwise the floor slab soffit',
        u'- False ceiling detectors (family name with {}): always the soffit (slab, or a '
        u'ceiling in contact with it), {} measured from the soffit to the false ceiling below; '
        u'without a false ceiling the detector is reported as unsuitable'.format(
            u' and '.join(u'"{}"'.format(t) for t in BUZZER_FAMILY_TOKENS), BUZZER_PARAM),
        u'- Analysis line: vertical, **{:.0f} mm** long, centred on the insertion point '
        u'(tolerance {:.0f} mm above and below)'.format(
            options.tolerance_mm * LINE_FACTOR, options.tolerance_mm * LINE_FACTOR / 2.0),
        u'- Only surfaces above the reference level of the element count',
        u'- Stacked instances closer than {:.0f} mm form one package: the target is its '
        u'bottom face'.format(STACK_GAP_MM),
        u'- Elements with connected connectors: {}'.format(
            u'skipped' if options.skip_connected else u'processed'),
        u'- Linked models: {}'.format(
            u'included in the search' if options.include_links else u'excluded from the search'),
        u'- Plan position and rotation are never changed',
    ]
    if result.surface_problems:
        lines.append(u'- Ceilings and floors discarded: **{}** (table at the end)'.format(
            len(result.surface_problems)))
    output.print_md(u'\n'.join(lines))


def print_category_summary(result, options):
    output.print_md(u'## Summary by category')
    moved_header = u'To align' if options.dry_run else u'Aligned'
    rows = []
    totals = [0, 0, 0, 0, 0, 0]
    fields = ('selected', 'moved', 'ok', 'screened', 'skipped', 'over')
    for key in sorted(result.per_category.keys(), key=lambda k: result.per_category[k]['name']):
        row = result.per_category[key]
        rows.append([row['name']] + [str(row[f]) for f in fields])
        for index, field in enumerate(fields):
            totals[index] += row[field]
    rows.append([u'**Total**'] + [str(v) for v in totals])
    output.print_table(table_data=rows, title='',
                       columns=[u'Category', u'Selected', moved_header, u'Already aligned',
                                u'Not ceiling-installed', u'Skipped', u'Beyond tolerance'])


def print_moves_table(result, options):
    if not result.planned:
        return
    output.print_md(u'## Elements {}'.format(u'to align' if options.dry_run else u'aligned'))

    ordered = sorted(result.planned, key=lambda r: (r.level_name, r.kind, -abs(r.distance)))
    shown = ordered[:MAX_MOVED_ROWS]
    columns = [u'Element', u'Category', u'Type', u'Level', u'Target', u'Surface',
               u'Elev. before', u'Elev. target', u'Move', BUZZER_PARAM, u'Connectors', u'Notes']
    if not options.dry_run:
        columns.append(u'Outcome')

    rows = []
    for record in shown:
        if record.is_buzzer:
            buzzer = u'{} -> {}'.format(format_buzzer(record.buzzer_before),
                                       format_buzzer(record.buzzer_value))
        else:
            buzzer = u'-'
        row = [
            output.linkify(record.element_id),
            record.category_name,
            record.type_name,
            record.level_name,
            record.kind,
            surface_cell(record.surface_id, record.surface_label, record.surface_is_linked),
            offset_cell(record.point_before.Z, record.level_z),
            offset_cell(record.target_z, record.level_z),
            format_signed_mm(record.move),
            buzzer,
            str(record.connected) if record.connected else u'-',
            record.note or u'',
        ]
        if not options.dry_run:
            row.append(record_problem(record) or u'OK')
        rows.append(row)
    output.print_table(table_data=rows, title='', columns=columns)

    remaining = len(ordered) - len(shown)
    if remaining > 0:
        output.print_md(u'_...and {} more elements not listed. All of them were processed '
                        u'anyway._'.format(remaining))
    output.print_md(u'_Elevations are measured from the reference level in the Level column. '
                    u'Move: positive upwards._')


def print_already_ok(result):
    if not result.already_ok:
        return
    output.print_md(u'## Elements already aligned\n\n**{}** elements were already within {:.0f} mm '
                    u'of their target surface and were left untouched.'.format(
                        len(result.already_ok), POSITION_TOL_MM))


def print_screened_table(result):
    """Esclusi dallo screening, raggruppati per tipo e motivo.

    Con una selezione a finestra possono essere la maggioranza (i dispositivi
    a parete): una riga per elemento seppellirebbe il resto del resoconto.
    Il collegamento di ogni riga seleziona tutti gli elementi del gruppo.
    """
    screened = [r for r in result.skipped if r.screened]
    if not screened:
        return
    output.print_md(u'## Not flagged for ceiling installation')
    groups = {}
    for record in screened:
        key = (type_name_of(record.element), record.reason)
        groups.setdefault(key, []).append(record.element_id)
    rows = []
    for (type_name, reason), ids in sorted(groups.items()):
        rows.append([type_name, reason, str(len(ids)), output.linkify(ids, title=u'select')])
    output.print_table(table_data=rows, title='',
                       columns=[u'Type', u'Reason', u'Count', u'Elements'])


def print_skipped_table(result):
    skipped = [r for r in result.skipped if not r.screened]
    if not skipped:
        return
    output.print_md(u'## Skipped elements')
    shown = skipped[:MAX_SKIPPED_ROWS]
    rows = []
    for record in shown:
        target = record.target
        rows.append([
            output.linkify(record.element_id),
            record.category_name,
            record.reason,
            surface_cell(target.surface_id, target.label, target.is_linked) if target else u'-',
            format_signed_mm(-target.distance) if target else u'-',
        ])
    output.print_table(table_data=rows, title='',
                       columns=[u'Element', u'Category', u'Reason', u'Nearest target',
                                u'Move needed'])
    remaining = len(skipped) - len(shown)
    if remaining > 0:
        output.print_md(u'_...and {} more skipped elements not listed._'.format(remaining))


def print_over_tolerance_table(result, options):
    if not result.over_tolerance_items:
        return
    output.print_md(u'## Skipped as beyond the tolerance')
    ordered = sorted(result.over_tolerance_items, key=lambda r: abs(r.target.distance))
    shown = ordered[:MAX_OVER_TOLERANCE_ROWS]
    rows = []
    for record in shown:
        target = record.target
        rows.append([
            output.linkify(record.element_id),
            record.category_name,
            target.kind,
            surface_cell(target.surface_id, target.label, target.is_linked),
            format_signed_mm(-target.distance),
            u'+{}'.format(format_mm(record.excess)),
            record.note or u'',
        ])
    output.print_table(table_data=rows, title='',
                       columns=[u'Element', u'Category', u'Target', u'Surface', u'Move needed',
                                u'Overshoot', u'Notes'])
    remaining = len(ordered) - len(shown)
    if remaining > 0:
        output.print_md(u'_...and {} more elements not listed._'.format(remaining))
    suggestion = tolerance_suggestion(ordered, options)
    if suggestion:
        output.print_md(suggestion)


def tolerance_suggestion(ordered, options):
    """Riga che dice di quanto alzare la soglia e quanto si recupererebbe."""
    if not ordered:
        return None

    # Distanza dalla superficie piu' vicina: per i rilevatori da controsoffitto
    # puo' essere il controsoffitto e non il bersaglio, ed e' gia' in excess.
    def nearest(record):
        return record.excess + options.tolerance_internal

    ranked = sorted(ordered, key=nearest)
    index = min(int(len(ranked) * 0.8), len(ranked) - 1)
    target_mm = internal_to_mm(nearest(ranked[index]))
    step_mm = 50.0
    candidate_mm = (int(target_mm / step_mm) + 1) * step_mm
    if candidate_mm <= options.tolerance_mm:
        candidate_mm = options.tolerance_mm + step_mm
    if candidate_mm > MAX_TOLERANCE_MM:
        return None
    limit = mm_to_internal(candidate_mm)
    recovered = len([r for r in ranked if nearest(r) <= limit])
    if not recovered:
        return None
    return u'_Raising the tolerance to **{:.0f} mm** would bring in **{}** more elements. Check ' \
           u'first that the target is the right one: a large move often means the family or the ' \
           u'level is wrong._'.format(candidate_mm, recovered)


def print_surface_problems(result):
    if not result.surface_problems:
        return
    output.print_md(u'## Ceilings and floors discarded')
    rows = []
    for source, element, kind, reason in result.surface_problems:
        label = surface_label(element, source, kind)
        rows.append([surface_cell(element.Id, label, source.is_linked), reason])
    output.print_table(table_data=rows, title='', columns=[u'Element', u'Reason'])


def print_revit_failures(preprocessor):
    if preprocessor is None or not preprocessor.messages:
        return
    output.print_md(u'## Revit warnings raised while editing')
    rows = []
    for description, ids, kind in preprocessor.messages:
        links = u', '.join([output.linkify(i) for i in ids[:5]]) or u'-'
        if len(ids) > 5:
            links += u' ...'
        rows.append([kind, description, links])
    output.print_table(table_data=rows, title='', columns=[u'Kind', u'Description', u'Elements'])
    output.print_md(u'_No constraint was unlocked or deleted: the tool applies no automatic '
                    u'resolutions._')


def print_warnings():
    if not WARNINGS:
        return
    output.print_md(u'## Warnings')
    for message in WARNINGS:
        output.print_md(u'- {}'.format(message))


def print_report(options, elements, source_label, result, preprocessor,
                 applied_ok, applied_failed, mismatched, rolled_back):
    """Report avanzato: il resoconto completo nel pannello di output."""
    output.close_others()

    print_header(options, elements, source_label, result)
    print_category_summary(result, options)
    print_moves_table(result, options)
    print_already_ok(result)
    print_over_tolerance_table(result, options)
    print_skipped_table(result)
    print_screened_table(result)
    print_surface_problems(result)
    print_revit_failures(preprocessor)
    print_warnings()

    output.print_md(u'## Outcome')
    if rolled_back:
        output.print_md(u'**The edit was rolled back: the model was not touched.** The {} planned '
                        u'elements are still at their original elevation. The reason is in the '
                        u'warnings above.'.format(len(result.planned)))
    elif not result.planned:
        output.print_md(u'Nothing to align. The tables above say whether the problem is the '
                        u'screening parameter, the tolerance, the categories or the ceilings and '
                        u'floors found.')
    elif options.dry_run:
        output.print_md(u'**{}** elements would be aligned. Run the command again with '
                        u'"Simulation" cleared to apply the change.'.format(len(result.planned)))
    else:
        output.print_md(u'Elements aligned: **{}** of {} planned.'.format(
            applied_ok - mismatched, len(result.planned)))
        if applied_failed:
            output.print_md(u'Elements left unchanged because of a Revit error: **{}**. The reason '
                            u'is in the "Outcome" column.'.format(applied_failed))
        if mismatched:
            output.print_md(u'Elements whose elevation or {} after the edit differs from the '
                            u'target: **{}**. See the "Outcome" column.'.format(BUZZER_PARAM,
                                                                               mismatched))
        output.print_md(u'The change is grouped into a single undo step, named "{}".'.format(
            TRANSACTION_NAME))


# =========================================================================
# FINESTRA DI ESITO
# =========================================================================
# Report semplificato: in testa solo i TOTALI (allineati, gia' a posto, non
# allineati), sotto la tabella dei soli elementi NON allineati con le note. La
# finestra non conosce la Revit API: riceve righe gia' pronte, ciascuna con
# una chiave che lo script traduce in ElementId.
#
# La finestra e' NON MODALE: il comando termina e la lascia aperta, cosi'
# Revit e' libero e l'utente puo' spostare gli elementi e cambiarne i
# parametri. Il pulsante di una riga seleziona e inquadra l'elemento (come in
# RNWFlowCalculation, ma senza la PickObjects che bloccava le modifiche).
# Da una finestra non modale la Revit API si usa solo tramite ExternalEvent:
# ResultEventHandler esegue le richieste nel contesto API. Il motore
# IronPython deve restare vivo dopo il comando: "engine: persistent: true"
# nel bundle.yaml, come ColorSplasher e Keynotes di pyRevit.

STATUS_FAILED = u'Failed'
STATUS_NOT_ALIGNED = u'Not aligned'

# Finestre aperte e relativi ExternalEvent: un riferimento vivo evita che il
# garbage collector li raccolga mentre la finestra e' a video.
OPEN_RESULT_WINDOWS = []


def mark_rolled_back(result):
    """Transazione annullata: nessun elemento e' stato toccato."""
    for record in result.planned:
        record.applied = False
        record.z_after = None
        record.buzzer_after = None
        if not record.error:
            record.error = u'transaction rolled back'


def mm_value(value_internal):
    return None if value_internal is None else round(internal_to_mm(value_internal), 1)


def join_notes(*parts):
    return u'; '.join(part for part in parts if part)


def result_rows(options, result, context):
    """(righe per la finestra, {chiave: elemento}).

    Solo gli elementi NON allineati: falliti in applicazione, oltre
    tolleranza e saltati (screening compreso). Gli allineati e quelli gia' a
    posto compaiono solo nei totali; in simulazione anche gli "da allineare".
    """
    rows = []
    elements = {}

    def level_of(element):
        storey = context.storeys.get(element_id_value(element.Id))
        return storey.level_name if storey is not None else u'-'

    def add(element, status, category_name, level, kind, move, buzzer, notes):
        key = len(rows)
        elements[key] = element
        rows.append({
            'key': key,
            'status': status,
            'id': element_id_value(element.Id),
            'category': category_name,
            'type': type_name_of(element),
            'level': level,
            'target': kind,
            'move_mm': mm_value(move),
            'buzzer_mm': mm_value(buzzer),
            'notes': notes,
        })

    for record in failed_records(options, result):
        add(record.element, STATUS_FAILED, record.category_name, record.level_name, record.kind,
            record.move, record.buzzer_value if record.is_buzzer else None,
            join_notes(record_problem(record), record.note))

    for record in result.over_tolerance_items:
        target = record.target
        add(record.element, STATUS_NOT_ALIGNED, record.category_name, level_of(record.element),
            target.kind, -target.distance, None,
            join_notes(u'beyond the tolerance by {}'.format(format_mm(record.excess)), record.note))

    for record in result.skipped:
        target = record.target
        add(record.element, STATUS_NOT_ALIGNED, record.category_name, level_of(record.element),
            target.kind if target is not None else u'-',
            -target.distance if target is not None else None, None, record.reason)

    return rows, elements


def failed_records(options, result):
    """Pianificati che non risultano allineati: errore, rollback o quota sbagliata.

    In simulazione nessuno e' fallito: sono tutti "da allineare".
    """
    if options.dry_run:
        return []
    return [r for r in result.planned if not r.applied or record_problem(r) is not None]


def result_summary(options, result, rolled_back):
    """Totali in testa alla finestra: gli allineati non sono elencati."""
    failed_list = failed_records(options, result)
    failed = len(failed_list)
    done = len(result.planned) - failed
    buzzers = sum(1 for r in result.planned if r.is_buzzer and r not in failed_list)
    not_aligned = failed + len(result.skipped) + len(result.over_tolerance_items)
    total = len(result.planned) + len(result.already_ok) + len(result.skipped) \
        + len(result.over_tolerance_items)

    lines = []
    if options.dry_run:
        lines.append(u'Simulation: nothing was changed.')
    elif rolled_back:
        lines.append(u'The edit was rolled back: nothing was changed.')
    lines.append(u'Elements selected: {}'.format(total))
    if options.dry_run:
        lines.append(u'To align: {}'.format(done))
    else:
        lines.append(u'Aligned: {}'.format(done))
    if buzzers:
        lines.append(u'   of which false ceiling detectors (to the soffit, {} {}): {}'.format(
            BUZZER_PARAM, u'to set' if options.dry_run else u'set', buzzers))
    lines.append(u'Already aligned: {}'.format(len(result.already_ok)))
    if not_aligned:
        detail = u' ({} failed)'.format(failed) if failed else u''
        lines.append(u'Not aligned: {}{}, listed below'.format(not_aligned, detail))
    else:
        lines.append(u'Not aligned: 0')
    return u'\n'.join(lines)


class ResultEventHandler(UI.IExternalEventHandler):
    """Esegue nel contesto API di Revit le richieste della finestra non modale.

    La finestra imposta request (una funzione che riceve la UIApplication) e
    alza l'ExternalEvent; Revit chiama Execute appena e' libero. Un errore
    finisce nella riga di stato della finestra, non in un popup.
    """

    def __init__(self):
        self.request = None
        self.window = None

    def Execute(self, uiapp):
        request, self.request = self.request, None
        if request is None:
            return
        try:
            message = request(uiapp)
        except Exception as error:
            logger.debug('Richiesta della finestra di esito fallita: %s', error)
            message = u'The request failed: {}'.format(error)
        if self.window is not None and message:
            try:
                self.window.set_status(message)
            except Exception:
                pass

    def GetName(self):
        return u'MEP Align to Ceiling - result window'


def show_elements(uiapp, document, elements):
    """Seleziona e inquadra gli elementi nel documento attivo. Ritorna l'esito.

    Le chiamate partono da una finestra non modale, quindi a comando finito:
    il documento attivo puo' essere un altro, gli elementi possono essere
    stati cancellati nel frattempo. Si controlla tutto e lo si dice.
    """
    active = uiapp.ActiveUIDocument
    if active is None or not active.Document.Equals(document):
        return u'Switch back to the model "{}" to show its elements.'.format(document.Title)
    valid = [e for e in elements if e is not None and e.IsValidObject]
    if not valid:
        return u'The element no longer exists in the model.'
    element_ids = List[DB.ElementId]([e.Id for e in valid])
    active.Selection.SetElementIds(element_ids)
    try:
        active.ShowElements(element_ids)
    except Exception:
        return u'{} element(s) selected, but the active view cannot zoom to them: open a ' \
               u'plan, section or 3D view.'.format(len(valid))
    missing = len(elements) - len(valid)
    message = u'{} element(s) selected and shown: edit them in Revit, this window stays ' \
              u'open.'.format(len(valid))
    if missing:
        message += u' {} no longer exist.'.format(missing)
    return message


def show_result_window(options, result, context, rolled_back, advanced_report):
    """Apre la finestra di esito NON MODALE e ritorna subito.

    Il comando termina con la finestra a video: da li' il pulsante di una
    riga seleziona l'elemento e Revit resta libero per modificarlo.
    """
    rows, elements = result_rows(options, result, context)
    document = doc
    handler = ResultEventHandler()
    event = UI.ExternalEvent.Create(handler)

    def request(action):
        handler.request = action
        event.Raise()

    def show(keys):
        request(lambda uiapp: show_elements(uiapp, document, [elements.get(k) for k in keys]))

    def report():
        def run(uiapp):
            advanced_report()
            return u'Advanced report printed in the pyRevit output panel.'
        request(run)

    entry = {}

    def closed():
        if entry in OPEN_RESULT_WINDOWS:
            OPEN_RESULT_WINDOWS.remove(entry)
        try:
            event.Dispose()
        except Exception:
            pass

    info = {
        'title': u'MEP Align to Ceiling - simulation' if options.dry_run
        else u'MEP Align to Ceiling - result',
        'summary': result_summary(options, result, rolled_back),
        'buzzer_header': u'{} (mm)'.format(BUZZER_PARAM),
        'rows': rows,
        'show': show,
        'report': report,
        'closed': closed,
    }
    window = result_ui.ResultWindow(resolve_xaml_path(RESULT_XAML_FILE_NAME), info)
    handler.window = window
    entry.update({'window': window, 'handler': handler, 'event': event})
    OPEN_RESULT_WINDOWS.append(entry)
    owner = None
    try:
        owner = uidoc.Application.MainWindowHandle
    except Exception:
        owner = None
    window.show_modeless(owner)


# =========================================================================
# PROGRAMMA PRINCIPALE
# =========================================================================

def build_non_pickable_view_types():
    types = []
    for name in ('Schedule', 'ColumnSchedule', 'PanelSchedule', 'DrawingSheet', 'Legend',
                 'Report', 'ProjectBrowser', 'SystemBrowser', 'Undefined'):
        view_type = getattr(DB.ViewType, name, None)
        if view_type is not None:
            types.append(view_type)
    return types


def check_preconditions():
    """Guardie di apertura, prima di qualunque interfaccia."""
    if doc is None:
        forms.alert(u'No document is open.', exitscript=True)
    if doc.IsFamilyDocument:
        forms.alert(u'The command is not available in the family editor.', exitscript=True)
    version = revit_version()
    if version and version < 2022:
        forms.alert(u'The tool requires Revit 2022 or later. Version detected: {}.'.format(version),
                    exitscript=True)
    if not MEP_CATEGORIES:
        forms.alert(u'None of the configured categories is available in this version of Revit.',
                    exitscript=True)
    active_view = doc.ActiveView
    if active_view is None:
        forms.alert(u'No active view.', exitscript=True)
    if active_view.ViewType in build_non_pickable_view_types():
        forms.alert(u'Activate a graphical view (plan, section, 3D) before running the command: '
                    u'elements are picked in the view.', exitscript=True)


def stop_with_report(headline, alert):
    output.close_others()
    output.print_md(u'# MEP alignment to ceilings')
    output.print_md(headline)
    forms.alert(alert, title=u'Nothing to align', exitscript=True)


def active_design_option_id():
    try:
        return DB.DesignOption.GetActiveDesignOptionId(doc)
    except Exception:
        return DB.ElementId.InvalidElementId


def search_region(elements, storeys, options):
    """Volume host che contiene le fasce di ricerca di tutti gli elementi."""
    xs = []
    ys = []
    lows = []
    highs = []
    for element in elements:
        point = insertion_point(element)
        if point is None:
            continue
        storey = storeys.get(element_id_value(element.Id))
        low, high = search_window(point.Z, storey, options.half_line)
        xs.append(point.X)
        ys.append(point.Y)
        lows.append(low)
        highs.append(high)
    if not xs:
        return None, None
    margin = XY_EPS
    return (DB.XYZ(min(xs) - margin, min(ys) - margin, min(lows)),
            DB.XYZ(max(xs) + margin, max(ys) + margin, max(highs)))


def explain_empty_search(options, sources):
    """Avviso quando nessuna superficie e' stata trovata nella regione."""
    if any(source.surfaces for source in sources):
        return
    linked = u' nor in the linked models' if options.include_links else u''
    hint = u'' if options.include_links else \
        u' Searching in links is off: tick "Look in linked models as well" in the dialog.'
    warn(u'No ceiling or floor was found around the selected elements, neither in this '
         u'model{}.{}'.format(linked, hint))


def main():
    check_preconditions()

    elements, source_label = resolve_elements()

    if not any(insertion_point(element) is not None for element in elements):
        stop_with_report(
            u'None of the **{}** selected elements has an insertion point: they are all host '
            u'based, line based or in place. There is nothing to align.'.format(len(elements)),
            u'None of the selected elements has an insertion point.')

    category_info = count_per_category(elements)
    if not category_info:
        stop_with_report(
            u'None of the **{}** selected elements belongs to the {} categories this tool '
            u'handles.'.format(len(elements), len(MEP_CATEGORIES)),
            u'None of the selected elements belongs to a handled category.')

    options = ask_options(elements, source_label, category_info)
    if options is None:
        script.exit()

    context = PlanContext(active_design_option_id())

    # Il piano di riferimento si calcola prima della raccolta: delimita in
    # basso la fascia in cui cercare i solai.
    levels = host_levels()
    for element in elements:
        point = insertion_point(element)
        if point is not None:
            context.storeys[element_id_value(element.Id)] = storey_of(element, point, levels)

    host_min, host_max = search_region(elements, context.storeys, options)
    context.sources = collect_surface_sources(host_min, host_max, options, context)
    explain_empty_search(options, context.sources)

    linked_sources = [s for s in context.sources if s.is_linked]
    if linked_sources:
        warn(u'Ceilings and floors also searched in {} linked models: {}.'.format(
            len(linked_sources), u', '.join([s.link_name for s in linked_sources])))

    try:
        result = build_plan(elements, options, context)
    except UserWarning:
        forms.alert(u'Analysis cancelled: the model was not modified.', title=u'Cancelled',
                    exitscript=True)

    preprocessor = None
    applied_ok = 0
    applied_failed = 0
    mismatched = 0
    rolled_back = False

    if not options.dry_run and result.planned:
        preprocessor = AlignFailurePreprocessor()
        transaction = DB.Transaction(doc, TRANSACTION_NAME)
        handling = transaction.GetFailureHandlingOptions()
        handling = handling.SetForcedModalHandling(False)
        handling = handling.SetClearAfterRollback(True)
        handling = handling.SetFailuresPreprocessor(preprocessor)
        transaction.SetFailureHandlingOptions(handling)

        transaction.Start()
        try:
            applied_ok, applied_failed = apply_all(result.planned)
            status = transaction.Commit()
            if status == DB.TransactionStatus.RolledBack:
                warn(u'Revit rolled the transaction back on an error: no change was applied. The '
                     u'detail is in the "Revit warnings raised while editing" section.')
                applied_failed += applied_ok
                applied_ok = 0
                rolled_back = True
            else:
                mismatched = verify_applied(result.planned)
        except UserWarning:
            transaction.RollBack()
            warn(u'Operation cancelled by the user: no change applied.')
            applied_ok = 0
            applied_failed = 0
            rolled_back = True
        except Exception as error:
            transaction.RollBack()
            warn(u'Transaction rolled back on an error: {}'.format(error))
            applied_ok = 0
            applied_failed = 0
            rolled_back = True

    if rolled_back:
        mark_rolled_back(result)

    def advanced_report():
        print_report(options, elements, source_label, result, preprocessor,
                     applied_ok, applied_failed, mismatched, rolled_back)

    # Esito nella finestra non modale; il resoconto completo nel pannello di
    # output e' il report avanzato, su richiesta. Se la finestra non si apre
    # il resoconto si stampa subito: l'utente deve sempre vedere l'esito.
    try:
        show_result_window(options, result, context, rolled_back, advanced_report)
    except Exception as error:
        logger.debug('Finestra di esito non disponibile: %s', error)
        warn(u'The result window could not be opened ({}): the full report is shown '
             u'instead.'.format(error))
        advanced_report()


if __name__ == '__main__':
    main()
