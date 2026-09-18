# -*- coding: utf-8 -*-
"""Allineamento in pianta degli elementi MEP alle facce dei muri.

Dati uno o piu' muri selezionati, lo strumento raccoglie i dispositivi MEP
puntuali vicini e li porta a filo della faccia del muro piu' vicina, ruotandoli
di quanto basta perche' risultino perpendicolari alla parete.

La quota Z non viene MAI modificata: e' il motivo per cui il primo elemento
architettonico gestito e' il muro. L'allineamento in quota e' un problema
diverso (serve un livello o un soffitto come riferimento) ed e' fuori ambito.

--------------------------------------------------------------------------
LOGICA APPLICATA
--------------------------------------------------------------------------

1. MURI. Si parte dai muri nella selezione corrente; se non ce ne sono, viene
   chiesta una selezione grafica. I muri stacked vengono espansi nei loro
   sotto-muri. Muri tenda, muri inclinati e muri senza linea di
   posizionamento vengono scartati con il motivo nel resoconto.

2. CANDIDATI. Una sola passata di raccolta sul documento per le categorie
   scelte, poi un filtro di prossimita' per ogni muro sugli id gia' raccolti.
   La ricerca copre tutto il modello, ma e' limitata in verticale
   all'estensione dei muri: un dispositivo del piano superiore, allineato in
   pianta con il muro, non viene toccato.

3. DISTANZA. Misurata in pianta fra il punto di inserimento dell'elemento e la
   faccia del muro piu' vicina, non il suo asse. Un elemento che cade dentro
   lo spessore del muro ha distanza negativa e quindi rientra sempre nella
   soglia.

4. TRASLAZIONE. Perpendicolare al muro, tale da portare il punto di
   inserimento esattamente sulla faccia. La posizione lungo il muro resta
   invariata, la quota resta invariata.

5. ROTAZIONE. Attorno a un asse verticale passante per il punto di
   inserimento, dell'angolo piu' piccolo che rende il fronte dell'elemento
   perpendicolare al muro. Fra le due direzioni possibili (normale uscente e
   sua opposta) si scegle quella piu' vicina all'orientamento attuale, quindi
   l'angolo applicato non supera mai 90 gradi e un elemento gia' orientato
   bene non viene ribaltato. Conseguenza da conoscere: un elemento montato al
   contrario non viene raddrizzato, viene solo reso perpendicolare.

--------------------------------------------------------------------------
COME VIENE RICAVATO IL PIANO DI RIFERIMENTO DEL MURO
--------------------------------------------------------------------------

Wall.Location.Curve NON e' l'asse del muro: e' la linea di posizionamento,
che dipende dal parametro "Location Line" (WALL_KEY_REF_PARAM). Ignorare
questo dettaglio produce un errore fino a mezzo spessore di muro.

Lo scostamento firmato della linea di posizionamento rispetto alla mezzeria,
misurato lungo la normale esterna, vale:

    0 WallCenterline        0
    1 CoreCenterline        (d_int - d_ext) / 2
    2 FinishFaceExterior    + W/2
    3 FinishFaceInterior    - W/2
    4 CoreExterior          W/2 - d_ext
    5 CoreInterior          d_int - W/2

dove W e' lo spessore del muro e d_ext / d_int sono le somme degli spessori
degli strati fuori dal nucleo. CompoundStructure.GetLayers() e' ordinato
dall'esterno verso l'interno: GetFirstCoreLayerIndex() e' il lato esterno,
GetLastCoreLayerIndex() il lato interno. Con d_ext = d_int = 0 i tre casi
"core" collassano sui casi "finish" e "centerline", che e' il controllo di
coerenza della tabella.

Non serve mai costruire la curva di mezzeria: per rette e archi concentrici il
piede della perpendicolare sta sulla stessa direzione radiale, quindi basta lo
scalare di scostamento.

HostObjectUtils.GetSideFaces NON e' la strategia primaria, anche se a prima
vista sembrerebbe la piu' pulita. La documentazione di Face.Project dice che
restituisce null se il punto piu' vicino cade fuori dalla faccia, e una porta
o una finestra bucano la faccia laterale del muro: un dispositivo davanti al
vano di una porta proietterebbe dentro il buco e verrebbe scartato senza
motivo. In piu' null non permette di distinguere "40 mm oltre la testata" da
"tre stanze piu' in la'", informazione che serve al resoconto. Il calcolo
analitico risponde sempre e con una distanza firmata. GetSideFaces resta come
riserva per i muri senza linea di posizionamento.

--------------------------------------------------------------------------
NOTE SULLA GEOMETRIA
--------------------------------------------------------------------------

1. NORMALE ESTERNA. Sui muri rettilinei si usa Wall.Orientation: e' costante,
   autorevole e senza ambiguita' di segno. La formula
   BasisZ.CrossProduct(tangente), invertita su Wall.Flipped, serve solo per
   gli archi, dove la normale varia lungo il muro. Sui muri rettilinei viene
   comunque eseguito un autocontrollo fra le due, e una discordanza
   sistematica segnala che la convenzione del prodotto vettoriale va invertita.

   Il segno della normale pesa molto meno di quanto sembri. Invertendo n si
   inverte anche signed_center, e quindi side, e il punto bersaglio

       p + n * (side * half_width - signed_center)

   resta identico: il termine cambia segno due volte. Quando lo scostamento
   della linea di posizionamento e' nullo (Location Line = Wall Centerline,
   il caso di gran lunga piu' frequente) lo SPOSTAMENTO E' DUNQUE INVARIANTE
   rispetto al segno della normale. Un segno sbagliato altera solo
   l'etichetta "faccia esterna / interna" nel resoconto e la correzione di
   scostamento sui muri con Location Line diversa dalla mezzeria.

2. PARAMETRIZZAZIONE DEGLI ARCHI. Per un Arc il parametro grezzo e' quasi
   certamente l'angolo in radianti, ma la documentazione e' ambigua: dice che
   per rette e archi il parametro grezzo misura in piedi, e al tempo stesso
   Arc.Period vale 2*pi. Il codice non assume: verifica a runtime se
   span * raggio corrisponde alla lunghezza, e srotola atan2 (che vive in
   -pi..pi) sul dominio effettivo della curva.

3. TANGENTE. Non si usa Curve.ComputeDerivatives. Con normalized=False un
   valore normalizzato passato per errore e' un angolo perfettamente valido:
   nessuna eccezione, risultato sbagliato in silenzio. E i vettori restituiti
   non sono normalizzati. Si usa Line.Direction per le rette e
   Normal.CrossProduct(radiale) per gli archi, entrambi gia' unitari.

4. FILTRO VERTICALE. L'estensione verticale del muro si legge dal suo bounding
   box, non dai livelli. Confrontando geometria contro geometria, entrambe in
   coordinate interne, la trappola Level.Elevation contro Level.ProjectElevation
   documentata nel CLAUDE.md non si presenta affatto. Il bounding box tiene
   inoltre conto degli attacchi a tetto e dei profili modificati.

5. ARBITRAGGIO FRA MURI. Un elemento vicino a piu' muri selezionati viene
   assegnato a quello con il valore ASSOLUTO della distanza dalla faccia piu'
   piccolo. Usare il valore firmato sarebbe sbagliato: un elemento immerso in
   un muro spesso ha distanza molto negativa e vincerebbe sempre contro un
   muro adiacente a pochi millimetri.

--------------------------------------------------------------------------
GESTIONE DEGLI AVVISI DI REVIT
--------------------------------------------------------------------------

Lo strumento NON usa revit.Transaction(swallow_errors=True) ne'
revit.ErrorSwallower(). Entrambi si appoggiano al FailureSwallower di pyRevit,
la cui lista RESOLUTION_TYPES include UnlockConstraints e DeleteElements:
su un avviso "Constraints are not satisfied" sbloccherebbe in silenzio le
quote e gli allineamenti dell'utente. Su uno strumento che sposta geometria e'
esattamente il comportamento da non avere.

Al suo posto c'e' un IFailuresPreprocessor dedicato che registra ogni avviso
per il resoconto, elimina soltanto i warning privi di risoluzioni (quelli che
aprirebbero un popup e bloccherebbero il lotto), non applica nessuna
risoluzione e chiede il rollback in presenza di errori.

--------------------------------------------------------------------------
ORDINE DELLE OPERAZIONI
--------------------------------------------------------------------------

Prima la rotazione, poi la traslazione, e la traslazione viene ricalcolata dal
punto corrente verso un punto bersaglio ASSOLUTO memorizzato in fase di
analisi. Non e' garantito che ElementTransformUtils.RotateElement attorno a un
asse passante per il punto di inserimento lasci quel punto esattamente
invariato per ogni tipo di famiglia. Ricalcolando la traslazione dopo la
rotazione, qualunque deriva si autocorregge; la quota resta invariata per
costruzione, perche' la componente Z del vettore e' forzata a zero.

Ogni elemento viene elaborato in una SubTransaction propria: se la rotazione
riesce e la traslazione fallisce, l'elemento torna intatto invece di restare
ruotato e non spostato.

--------------------------------------------------------------------------
LIMITI NOTI
--------------------------------------------------------------------------

- Solo elementi con LocationPoint. Canali, tubi e passerelle hanno una
  LocationCurve e richiedono un algoritmo diverso.
- Il punto di inserimento finisce sulla faccia. Se una famiglia ha l'origine
  al centro del proprio ingombro invece che sul retro, l'elemento risulta per
  meta' dentro il muro. Le colonne "distanza prima" e "distanza dopo" del
  resoconto servono a rilevarlo sul primo modello reale.
- Muri di modelli collegati non sono gestiti.
- Lo strumento non cambia mai l'host di un elemento: se e' ospitato da un muro
  lo salta.

--------------------------------------------------------------------------
MOTORE PYTHON
--------------------------------------------------------------------------

Lo script e' scritto per il motore IronPython di pyRevit, che e' quello
predefinito: il file NON deve contenere la direttiva "#! python3". Il modulo
pyrevit.forms, usato qui per caricare la finestra XAML, si appoggia a
wpf.LoadComponent, disponibile solo in IronPython. Sotto il motore CPython3 le
finestre WPF di pyRevit non sono attualmente supportate (pyRevit issue #3033).
"""

import math
import os.path as op

import clr
# Gli assembly WPF servono per costruire in codice le checkbox delle categorie.
# Thickness sta in WindowsBase, Controls in PresentationFramework: vanno
# referenziati esplicitamente, perche' IronPython non li carica da solo e
# l'ordine degli import non garantisce che pyrevit.forms lo abbia gia' fatto.
for _asm in ("WindowsBase", "PresentationCore", "PresentationFramework"):
    try:
        clr.AddReference(_asm)
    except Exception:
        pass

from System.Collections.Generic import List
from System.Windows import Thickness
from System.Windows import Controls

from pyrevit import revit, DB, UI, forms, script


doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
logger = script.get_logger()


# =========================================================================
# CONFIGURAZIONE
# =========================================================================
# Le categorie sono dichiarate come nomi di BuiltInCategory. I nomi non
# riconosciuti dalla versione di Revit in uso vengono semplicemente
# ignorati, cosi' lo script resta compatibile con piu' versioni.

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

XAML_FILE_NAME = 'MEPAlignWindow.xaml'

TRANSACTION_NAME = u'Allineamento MEP ai muri'

DEFAULT_TOLERANCE_CM = 30.0     # soglia proposta nella finestra
MAX_TOLERANCE_CM = 500.0        # oltre e' quasi certamente un errore di battitura

POSITION_TOL_MM = 1.0           # sotto questo spostamento l'elemento e' gia' a posto
ANGLE_TOL_DEG = 0.1             # sotto questo angolo non si ruota
WALL_END_TOL_MM = 0.1           # sporgenza ammessa oltre l'estremita' del muro
Z_TOL_MM = 10.0                 # margine verticale sul test di contenimento
AMBIGUITY_TOL_MM = 20.0         # due muri entro questo scarto: caso segnalato

# Categorie che contano come partizione verticale architettonica a cui
# allineare. I muri tenda restano esclusi dalla categoria Walls perche'
# hanno spessore nullo e le due facce coincidono.
PARTITION_CATEGORY_NAMES = [
    'OST_Walls',
]

# Dilatazione del volume di ricerca oltre la tolleranza. Il filtro nativo
# confronta il BOUNDING BOX della partizione con il volume, mentre il
# criterio vero e' la distanza del punto di inserimento dalla FACCIA: il
# margine deve coprire lo spessore della partizione e le partizioni la cui
# geometria si estende ben oltre il tratto vicino all'elemento.
PARTITION_MARGIN_MM = 2000.0

VERTICAL_FACING_TOL = 0.087     # sin(5 gradi): sotto, il fronte e' verticale
GEOM_EPS = 1.0e-9

SLANT_ANGLE_TOL = 1.0e-6        # radianti: sotto, il muro e' verticale
SLANT_NORMAL_TOL = 0.02         # circa 1.1 gradi di inclinazione della faccia

# Alzare a True se in prova la rilettura del punto di inserimento subito dopo
# RotateElement risultasse non aggiornata (vedi "ORDINE DELLE OPERAZIONI").
FORCE_REGEN = False

MAX_MOVED_ROWS = 300
MAX_SKIPPED_ROWS = 200
MAX_OVER_TOLERANCE_ROWS = 50

# Valori dell'enumeratore WallLocationLine.
LOC_CENTERLINE = 0
LOC_CORE_CENTERLINE = 1
LOC_FINISH_EXTERIOR = 2
LOC_FINISH_INTERIOR = 3
LOC_CORE_EXTERIOR = 4
LOC_CORE_INTERIOR = 5

# Vocabolario chiuso dei motivi di esclusione: dichiarati in un unico punto
# perche' lo stesso motivo non finisca scritto in due modi diversi.
R_NO_POINT = u'elemento senza punto di inserimento'
R_HOSTED_WALL = u'ospitato dal muro {}'
R_HOSTED_OTHER = u'ospitato da {} {}'
R_PINNED = u'elemento bloccato (pin)'
R_GROUP = u'elemento nel gruppo "{}"'
R_SUBCOMPONENT = u'sotto-componente di famiglia annidata'
R_DESIGN_OPTION = u'opzione di progetto non attiva'
R_BORROWED = u'in prestito ad altro utente'
R_FACING_VERTICAL = u'fronte verticale (elemento a soffitto o a pavimento)'
R_OUT_OF_Z = u'ingombro fuori dall\'estensione verticale dei muri'
R_BEYOND_END = u'oltre l\'estremita\' del muro di {}'
R_NO_PROJECTION = u'proiezione sulla geometria del muro non calcolabile'
R_CATEGORY_NOT_HANDLED = u'categoria non gestita dallo strumento'
R_CATEGORY_NOT_SELECTED = u'categoria esclusa nella finestra'
R_NO_PARTITION = u'nessuna partizione verticale nel raggio di ricerca'
R_OVER_TOLERANCE = u"oltre la tolleranza: {} dalla faccia piu' vicina"
R_CONNECTED = u'collegato ad altri elementi ({} connettori)'
R_ANALYSIS_ERROR = u'errore in analisi: {}'

W_CURTAIN = u'muro tenda: spessore nullo, la distanza dalla faccia non e\' definita'
W_SLANTED = u'muro inclinato: la faccia non e\' verticale'
W_NO_CURVE = u'muro senza linea di posizionamento'
W_STACKED_PARENT = u'contenitore di muro sovrapposto: si usano i suoi membri'
W_NO_BBOX = u'estensione verticale del muro non leggibile'
W_NO_OFFSET = u'scostamento della linea di posizionamento indeterminato'


# =========================================================================
# COMPATIBILITA' TRA VERSIONI DI REVIT
# =========================================================================

def element_id_value(element_id):
    """Valore numerico di un ElementId.

    ElementId.Value esiste da Revit 2024, IntegerValue nelle versioni
    precedenti (e rimosso in Revit 2026).
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


def cm_to_internal(value_cm):
    return mm_to_internal(value_cm * 10.0)


def format_mm(value_internal):
    """Lunghezza in millimetri, con il segno se negativa."""
    return u'{:.0f} mm'.format(internal_to_mm(value_internal))


def format_deg(value_rad):
    return u'{:+.1f} deg'.format(math.degrees(value_rad))


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
    """Numero di versione di Revit, 0 se non leggibile."""
    try:
        return int(doc.Application.VersionNumber)
    except Exception:
        return 0


MEP_CATEGORIES = resolve_categories(MEP_CATEGORY_NAMES)
MEP_CATEGORY_KEYS = set(
    element_id_value(DB.ElementId(bic)) for bic in MEP_CATEGORIES)
PARTITION_CATEGORIES = resolve_categories(PARTITION_CATEGORY_NAMES)

POSITION_TOL = mm_to_internal(POSITION_TOL_MM)
WALL_END_TOL = mm_to_internal(WALL_END_TOL_MM)
Z_TOL = mm_to_internal(Z_TOL_MM)
AMBIGUITY_TOL = mm_to_internal(AMBIGUITY_TOL_MM)
PARTITION_MARGIN = mm_to_internal(PARTITION_MARGIN_MM)
ANGLE_TOL = math.radians(ANGLE_TOL_DEG)


# =========================================================================
# AVVISI
# =========================================================================
# Niente degradazione silenziosa: ogni valvola di sicurezza che scatta
# lascia una traccia. La deduplicazione non e' cosmetica, perche' alcune di
# queste funzioni vengono chiamate una volta per elemento candidato.

WARNINGS = []


def warn(message):
    """Registra un avviso globale, senza duplicati."""
    if message not in WARNINGS:
        WARNINGS.append(message)


# =========================================================================
# GEOMETRIA DEI MURI
# =========================================================================

class WallInfo(object):
    """Dati precalcolati di un muro, in coordinate interne."""

    def __init__(self, wall):
        self.wall = wall
        self.wall_id = wall.Id
        self.label = wall_label(wall)
        self.curve = None
        self.curve_z = 0.0
        self.is_straight = True
        self.orientation = None
        self.width = 0.0
        self.half_width = 0.0
        self.offset_locline = 0.0
        self.bbox = None
        self.z_min = 0.0
        self.z_max = 0.0


class WallHit(object):
    """Esito del test di un punto contro un muro."""

    def __init__(self, wall_info, foot, normal_ext, signed_center,
                 face_distance, side, beyond):
        self.wall_info = wall_info
        self.foot = foot
        self.normal_ext = normal_ext
        self.signed_center = signed_center
        self.face_distance = face_distance      # negativa se dentro lo spessore
        self.side = side                        # +1 faccia esterna, -1 interna
        self.beyond = beyond                    # sporgenza oltre la testata


def wall_label(wall):
    """Nome del tipo di muro, per il resoconto."""
    try:
        wall_type = wall.WallType
        if wall_type is not None:
            return DB.Element.Name.GetValue(wall_type)
    except Exception:
        pass
    return u'Muro'


def is_stacked_parent(wall):
    """True se e' il contenitore di un muro multistrato sovrapposto.

    Il padre non ha spessore ne' compound structure propri: i suoi membri
    sono elementi a se' stanti, gia' raccolti dal collector sulla categoria
    Walls, ognuno con il proprio spessore ed estensione verticale. Scartare
    il padre evita quindi di contare due volte la stessa parete.
    """
    try:
        return bool(wall.IsStackedWall)
    except Exception:
        return False


def wall_parameter(wall, name):
    """Parametro per nome di BuiltInParameter, None se non esiste."""
    bip = getattr(DB.BuiltInParameter, name, None)
    if bip is None:
        return None
    try:
        return wall.get_Parameter(bip)
    except Exception:
        return None


def side_face_max_normal_z(wall):
    """Componente Z massima, in valore assoluto, delle normali delle facce
    laterali del muro. None se la geometria non e' leggibile.

    Su una faccia verticale la normale e' orizzontale, quindi Z vale zero.
    E' una misura diretta sulla geometria, indipendente dalla versione di
    Revit e dalla semantica degli enumeratori.
    """
    references = []
    for layer_name in ('Exterior', 'Interior'):
        layer = getattr(DB.ShellLayerType, layer_name, None)
        if layer is None:
            continue
        try:
            found = DB.HostObjectUtils.GetSideFaces(wall, layer)
        except Exception:
            continue
        if found:
            references.extend(list(found))

    if not references:
        return None

    worst = None
    for reference in references:
        try:
            face = wall.GetGeometryObjectFromReference(reference)
            if face is None:
                continue
            box = face.GetBoundingBox()
            middle = DB.UV((box.Min.U + box.Max.U) / 2.0,
                           (box.Min.V + box.Max.V) / 2.0)
            normal_z = abs(face.ComputeNormal(middle).Z)
        except Exception:
            continue
        if worst is None or normal_z > worst:
            worst = normal_z
    return worst


def is_slanted(wall):
    """(True se il muro non e' verticale, dettaglio della misura).

    Su un muro inclinato la faccia non e' verticale, quindi "distanza in
    pianta dalla faccia" dipenderebbe dalla quota e l'intero modello a
    traslazione orizzontale pura non regge: il muro va scartato.

    La decisione si prende sulla GEOMETRIA, non sugli enumeratori. La prima
    versione di questa funzione leggeva WALL_CROSS_SECTION e scartava ogni
    muro di un progetto fatto di soli muri verticali: il valore intero di
    quel parametro non significa quello che sembra. L'angolo di inclinazione
    resta come scorciatoia, ma solo per confermare la verticalita' ed
    evitare il calcolo geometrico nel caso normale: da solo non basta mai a
    scartare un muro.
    """
    angle = wall_parameter(wall, 'WALL_SINGLE_SLANT_ANGLE_FROM_VERTICAL')
    if angle is not None:
        try:
            if abs(angle.AsDouble()) <= SLANT_ANGLE_TOL:
                return False, None
        except Exception:
            pass

    normal_z = side_face_max_normal_z(wall)
    if normal_z is None:
        # Geometria non leggibile: si assume verticale. Un falso negativo
        # produce un risultato sbagliato su quel muro, un falso positivo
        # renderebbe il comando inutilizzabile sull'intero progetto.
        return False, None

    if normal_z > SLANT_NORMAL_TOL:
        return True, u'componente Z della normale: {:.3f}'.format(normal_z)
    return False, None


def shell_thicknesses(wall):
    """Spessore del guscio esterno e di quello interno (piedi decimali).

    Ritorna (None, None) se la compound structure non e' leggibile.
    GetLayers() e' ordinato dall'esterno verso l'interno.
    """
    try:
        wall_type = wall.WallType
        compound = wall_type.GetCompoundStructure() if wall_type else None
    except Exception:
        compound = None
    if compound is None:
        return None, None

    try:
        count = compound.LayerCount
        first_core = compound.GetFirstCoreLayerIndex()
        last_core = compound.GetLastCoreLayerIndex()
    except Exception:
        return None, None

    if first_core < 0 or last_core < 0 or first_core > last_core \
            or last_core >= count:
        return None, None

    try:
        d_ext = 0.0
        for index in range(0, first_core):
            d_ext += compound.GetLayerWidth(index)
        d_int = 0.0
        for index in range(last_core + 1, count):
            d_int += compound.GetLayerWidth(index)
    except Exception:
        return None, None

    return d_ext, d_int


def location_line_offset(wall):
    """Scostamento firmato della linea di posizionamento dalla mezzeria.

    Misurato lungo la normale esterna, positivo verso l'esterno.
    Ritorna (offset, errore) con errore None se il calcolo e' esatto.
    """
    width = wall.Width
    kind = LOC_CENTERLINE
    try:
        param = wall.get_Parameter(DB.BuiltInParameter.WALL_KEY_REF_PARAM)
        if param is not None:
            kind = param.AsInteger()
    except Exception:
        kind = LOC_CENTERLINE

    if kind == LOC_CENTERLINE:
        return 0.0, None
    if kind == LOC_FINISH_EXTERIOR:
        return width / 2.0, None
    if kind == LOC_FINISH_INTERIOR:
        return -width / 2.0, None

    d_ext, d_int = shell_thicknesses(wall)
    if d_ext is None:
        # Indeterminato: l'errore potrebbe valere fino a mezzo spessore,
        # quindi non si approssima alla mezzeria.
        return None, W_NO_OFFSET

    if kind == LOC_CORE_EXTERIOR:
        return width / 2.0 - d_ext, None
    if kind == LOC_CORE_INTERIOR:
        return d_int - width / 2.0, None
    if kind == LOC_CORE_CENTERLINE:
        return (d_int - d_ext) / 2.0, None

    warn(u'Muro {}: valore "Location Line" non riconosciuto ({}), '
         u'usata la mezzeria.'.format(element_id_value(wall.Id), kind))
    return 0.0, None


def geometric_normal(tangent):
    """Normale orizzontale unitaria ricavata dalla tangente."""
    normal = DB.XYZ.BasisZ.CrossProduct(tangent)
    length = normal.GetLength()
    if length < GEOM_EPS:
        return None
    return normal.Normalize()


def exterior_normal_at(wall_info, tangent):
    """Normale esterna unitaria e orizzontale nel punto di tangente data.

    Sui muri rettilinei Wall.Orientation e' la verita' assoluta: costante e
    senza ambiguita' di segno. Sugli archi la normale varia lungo il muro e
    va ricavata dalla tangente, invertendola se il muro e' flipped.
    """
    if wall_info.is_straight and wall_info.orientation is not None:
        return wall_info.orientation

    normal = geometric_normal(tangent)
    if normal is None:
        return None
    try:
        if wall_info.wall.Flipped:
            normal = normal.Negate()
    except Exception:
        pass
    return normal


def build_wall_info(wall):
    """Precalcola i dati di un muro.

    Ritorna (WallInfo, None) oppure (None, motivo di scarto).
    """
    try:
        wall_type = wall.WallType
        if wall_type is not None and wall_type.Kind == DB.WallKind.Curtain:
            return None, W_CURTAIN
    except Exception:
        pass

    if is_stacked_parent(wall):
        return None, W_STACKED_PARENT

    try:
        width = wall.Width
    except Exception:
        width = 0.0
    if width is None or width < GEOM_EPS:
        return None, W_CURTAIN

    slanted, slant_detail = is_slanted(wall)
    if slanted:
        return None, u'{} ({})'.format(W_SLANTED, slant_detail)

    location = None
    try:
        location = wall.Location
    except Exception:
        location = None
    if not isinstance(location, DB.LocationCurve):
        return None, W_NO_CURVE
    curve = location.Curve
    if curve is None:
        return None, W_NO_CURVE

    bbox = None
    try:
        bbox = wall.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is None:
        return None, W_NO_BBOX

    offset, offset_error = location_line_offset(wall)
    if offset is None:
        return None, offset_error

    info = WallInfo(wall)
    info.curve = curve
    info.width = width
    info.half_width = width / 2.0
    info.offset_locline = offset
    info.is_straight = isinstance(curve, DB.Line)
    info.bbox = bbox
    info.z_min = bbox.Min.Z
    info.z_max = bbox.Max.Z

    if isinstance(curve, DB.Arc):
        info.curve_z = curve.Center.Z
    else:
        info.curve_z = curve.GetEndPoint(0).Z

    try:
        info.orientation = wall.Orientation
    except Exception:
        info.orientation = None

    # Autocontrollo permanente sui muri rettilinei: se la convenzione del
    # prodotto vettoriale fosse invertita, sugli archi tutti gli spostamenti
    # andrebbero dalla parte sbagliata e nessuno se ne accorgerebbe.
    if info.is_straight and info.orientation is not None:
        computed = geometric_normal(curve.Direction)
        if computed is not None:
            try:
                if info.wall.Flipped:
                    computed = computed.Negate()
            except Exception:
                pass
            if computed.DotProduct(info.orientation) < 0.99:
                warn(u'Muro {}: la normale calcolata dalla tangente non '
                     u'coincide con Wall.Orientation. Sui muri rettilinei si '
                     u'usa Wall.Orientation, ma sugli archi la convenzione '
                     u'del prodotto vettoriale va verificata.'.format(
                         element_id_value(info.wall_id)))

    return info, None


def project_on_curve(wall_info, point_flat):
    """Proiezione in pianta sulla curva ILLIMITATA del muro.

    Ritorna (foot, tangent, beyond) dove beyond e' la lunghezza di
    superamento dell'estremita' (0.0 se la proiezione cade dentro il muro),
    oppure None se la proiezione non e' calcolabile.

    Il test dei limiti e' fatto a mano e non con Curve.Project perche' la
    documentazione non chiarisce se Project tronchi alla curva limitata, e
    perche' su un arco un punto coincidente con il centro fa lanciare
    InvalidOperationException.
    """
    curve = wall_info.curve

    if isinstance(curve, DB.Line):
        start = curve.GetEndPoint(0)
        end = curve.GetEndPoint(1)
        vector = end - start
        length = vector.GetLength()
        if length < GEOM_EPS:
            return None
        ratio = (point_flat - start).DotProduct(vector) / (length * length)
        foot = start + vector.Multiply(ratio)
        tangent = vector.Normalize()
        if ratio < 0.0:
            beyond = -ratio * length
        elif ratio > 1.0:
            beyond = (ratio - 1.0) * length
        else:
            beyond = 0.0
        return foot, tangent, beyond

    if isinstance(curve, DB.Arc):
        center = curve.Center
        axis = curve.Normal.Normalize()
        radial = point_flat - center
        radial = radial - axis.Multiply(radial.DotProduct(axis))
        if radial.GetLength() < GEOM_EPS:
            # Punto sull'asse dell'arco: direzione radiale indefinita.
            return None
        radial_unit = radial.Normalize()
        foot = center + radial_unit.Multiply(curve.Radius)
        tangent = axis.CrossProduct(radial_unit)

        theta = math.atan2(radial.DotProduct(curve.YDirection),
                           radial.DotProduct(curve.XDirection))

        radius = curve.Radius
        try:
            start_param = curve.GetEndParameter(0)
            end_param = curve.GetEndParameter(1)
        except Exception:
            return foot, tangent, 0.0
        span = end_param - start_param

        # Il parametro grezzo di un Arc e' l'angolo o la lunghezza d'arco?
        # La documentazione e' ambigua: si verifica invece di assumere.
        if abs(span * radius - curve.Length) < 1.0e-6:
            param_is_angle = True
        elif abs(span - curve.Length) < 1.0e-6:
            param_is_angle = False
        else:
            param_is_angle = True
            warn(u'Muro {}: parametrizzazione dell\'arco non riconosciuta, '
                 u'assunto l\'angolo in radianti.'.format(
                     element_id_value(wall_info.wall_id)))

        if param_is_angle:
            value = theta
            period = 2.0 * math.pi
            to_length = radius
        else:
            value = theta * radius
            period = 2.0 * math.pi * radius
            to_length = 1.0

        # atan2 vive in -pi..pi, gli estremi dell'arco no: si srotola.
        half_period = period / 2.0
        guard = 0
        while value < start_param - half_period and guard < 8:
            value += period
            guard += 1
        guard = 0
        while value >= start_param + half_period and guard < 8:
            value -= period
            guard += 1

        if value < start_param:
            beyond = (start_param - value) * to_length
        elif value > end_param:
            beyond = (value - end_param) * to_length
        else:
            beyond = 0.0
        return foot, tangent, beyond

    # Altri tipi di curva (ellissi, spline): non gestiti.
    warn(u'Muro {}: tipo di linea di posizionamento non gestito ({}).'.format(
        element_id_value(wall_info.wall_id), type(wall_info.curve).__name__))
    return None


def test_point_against_wall(wall_info, point):
    """Distanza firmata di un punto dalla faccia del muro piu' vicina."""
    point_flat = DB.XYZ(point.X, point.Y, wall_info.curve_z)
    projected = project_on_curve(wall_info, point_flat)
    if projected is None:
        return None
    foot, tangent, beyond = projected

    normal_ext = exterior_normal_at(wall_info, tangent)
    if normal_ext is None:
        return None

    signed_center = (point_flat - foot).DotProduct(normal_ext) \
        + wall_info.offset_locline
    face_distance = abs(signed_center) - wall_info.half_width
    side = 1.0 if signed_center >= 0.0 else -1.0

    return WallHit(wall_info, foot, normal_ext, signed_center,
                   face_distance, side, beyond)


# =========================================================================
# RICERCA DELLE PARTIZIONI VERTICALI
# =========================================================================
# La direzione della ricerca e' invertita rispetto alla prima versione: si
# parte dagli elementi indicati dall'utente e per ciascuno si cercano le
# partizioni verticali vicine, invece di partire dai muri e raccogliere gli
# elementi attorno. Cambia solo la fase larga: la matematica di
# test_point_against_wall() resta identica.

# Cache dei WallInfo, indicizzata per valore di ElementId. Non e' una
# ottimizzazione accessoria: build_wall_info() puo' estrarre la geometria
# della partizione per il controllo di inclinazione, e la stessa partizione
# ricorre per tutti gli elementi che le stanno davanti. Senza cache quel
# costo verrebbe moltiplicato per il numero di elementi selezionati.
PARTITION_CACHE = {}


def build_category_filter(built_in_categories):
    """Filtro multicategoria con le categorie indicate."""
    category_list = List[DB.BuiltInCategory]()
    for built_in_category in built_in_categories:
        category_list.Add(built_in_category)
    return DB.ElementMulticategoryFilter(category_list)


def partition_info(partition):
    """(WallInfo, motivo di scarto) di una partizione, con cache."""
    key = element_id_value(partition.Id)
    cached = PARTITION_CACHE.get(key)
    if cached is not None:
        return cached
    try:
        info, reason = build_wall_info(partition)
    except Exception as error:
        info = None
        reason = u'errore in analisi: {}'.format(u'{}'.format(error)[:120])
    PARTITION_CACHE[key] = (info, reason)
    return info, reason


def points_outline(points, margin):
    """Volume che racchiude tutti i punti indicati, dilatato del margine."""
    if not points:
        return None
    xs = [p.X for p in points]
    ys = [p.Y for p in points]
    zs = [p.Z for p in points]
    return DB.Outline(
        DB.XYZ(min(xs) - margin, min(ys) - margin, min(zs) - margin),
        DB.XYZ(max(xs) + margin, max(ys) + margin, max(zs) + margin))


def point_outline(point, tolerance_internal):
    """Volume di ricerca attorno al punto di inserimento di un elemento."""
    margin = tolerance_internal + PARTITION_MARGIN
    return DB.Outline(
        DB.XYZ(point.X - margin, point.Y - margin, point.Z - margin),
        DB.XYZ(point.X + margin, point.Y + margin, point.Z + margin))


def collect_partition_ids(points, tolerance_internal):
    """Id delle partizioni verticali nella regione occupata dalla selezione.

    Una sola passata sul documento. La restrizione per singolo elemento
    avviene poi su questo insieme ridotto e non sull'intero modello: con
    duemila elementi selezionati la differenza e' fra duemila scansioni del
    documento e duemila filtri su qualche centinaio di id.
    """
    empty = List[DB.ElementId]()
    if not PARTITION_CATEGORIES:
        warn(u'Nessuna categoria di partizione verticale disponibile in '
             u'questa versione di Revit.')
        return empty

    outline = points_outline(points, tolerance_internal + PARTITION_MARGIN)
    if outline is None:
        return empty

    try:
        collector = DB.FilteredElementCollector(doc)            .WherePasses(build_category_filter(PARTITION_CATEGORIES))            .WhereElementIsNotElementType()            .WherePasses(DB.BoundingBoxIntersectsFilter(outline))
        return List[DB.ElementId](collector.ToElementIds())
    except Exception as error:
        warn(u'Raccolta delle partizioni verticali non riuscita: {}'.format(
            error))
        return empty


def partitions_near_point(point, tolerance_internal, region_ids):
    """Partizioni utilizzabili vicine a un punto.

    Ritorna (lista di WallInfo, lista di (partizione, motivo di scarto)).
    """
    if region_ids is None or region_ids.Count == 0:
        return [], []

    try:
        near = DB.FilteredElementCollector(doc, region_ids)            .WherePasses(DB.BoundingBoxIntersectsFilter(
                point_outline(point, tolerance_internal)))            .ToElements()
    except Exception as error:
        warn(u"Filtro di prossimita' non riuscito: {}".format(error))
        return [], []

    infos = []
    problems = []
    for partition in near:
        info, reason = partition_info(partition)
        if info is None:
            problems.append((partition, reason))
            continue
        infos.append(info)
    return infos, problems


def z_extent(element):
    """Estensione verticale di un elemento, in coordinate del modello."""
    try:
        bbox = element.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is not None:
        try:
            if bbox.Transform.IsIdentity:
                return bbox.Min.Z, bbox.Max.Z
        except Exception:
            return bbox.Min.Z, bbox.Max.Z
    try:
        location = element.Location
        if isinstance(location, DB.LocationPoint):
            elevation = location.Point.Z
            return elevation, elevation
    except Exception:
        pass
    return None


def z_contained(element_extent, wall_infos):
    """True se l'ingombro verticale ricade in quello di almeno una partizione.

    Il filtro nativo verifica l'INTERSEZIONE fra bounding box, mentre il
    requisito e' il CONTENIMENTO: serve quindi questo test esplicito a valle.
    """
    if element_extent is None:
        return False
    low, high = element_extent
    for wall_info in wall_infos:
        if low >= wall_info.z_min - Z_TOL and high <= wall_info.z_max + Z_TOL:
            return True
    return False


# =========================================================================
# COSTRUZIONE DEL PIANO (nessuna transazione)
# =========================================================================

class AlignOptions(object):
    """Scelte effettuate dall'utente nella finestra di dialogo."""

    def __init__(self, categories, tolerance_cm, apply_rotation,
                 skip_connected, dry_run):
        self.categories = categories
        self.tolerance_cm = tolerance_cm
        self.tolerance_internal = cm_to_internal(tolerance_cm)
        self.apply_rotation = apply_rotation
        self.skip_connected = skip_connected
        self.dry_run = dry_run


class PlannedMove(object):
    """Uno spostamento gia' calcolato.

    La fase di applicazione legge soltanto questi campi: nessun ricalcolo
    geometrico dentro la transazione.
    """

    PLANNED = 'planned'
    ALREADY_OK = 'already_ok'

    def __init__(self):
        self.element = None
        self.element_id = None
        self.category_name = u'-'
        self.category_key = None
        self.type_name = u'-'

        self.wall_id = None
        self.wall_label = u'-'
        self.face_side = u'-'

        self.point_before = None
        self.distance_before = 0.0

        # Punto bersaglio ASSOLUTO, non vettore: la traslazione viene
        # ricalcolata dopo la rotazione, cosi' ogni deriva si autocorregge.
        self.target_point = None
        self.distance_after = 0.0
        self.translation_length = 0.0

        self.rotation_rad = 0.0
        self.rotation_applicable = True

        self.needs_move = False
        self.needs_rotation = False
        self.status = PlannedMove.PLANNED

        self.connected = 0
        self.competing_walls = 1
        self.second_distance = None
        self.note = None

        self.applied_move = False
        self.applied_rotation = False
        self.error = None


class SkippedElement(object):
    """Un elemento escluso, con il motivo specifico e la distanza misurata."""

    def __init__(self, element, category_name, category_key, reason,
                 wall_id=None, distance=None):
        self.element = element
        self.element_id = element.Id if element is not None else None
        self.category_name = category_name
        self.category_key = category_key
        self.reason = reason
        self.wall_id = wall_id
        self.distance = distance


class OverTolerance(object):
    """Un elemento vicino ma oltre la soglia."""

    def __init__(self, element, category_name, category_key, wall_id,
                 wall_label, distance, excess):
        self.element = element
        self.element_id = element.Id
        self.category_name = category_name
        self.category_key = category_key
        self.wall_id = wall_id
        self.wall_label = wall_label
        self.distance = distance
        self.excess = excess


class PlanResult(object):

    def __init__(self):
        self.planned = []
        self.already_ok = []
        self.skipped = []
        self.over_tolerance_items = []
        self.per_category = {}
        self.candidate_count = 0
        self.wall_problems = []


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
            return u'{}: {}'.format(family_name,
                                    DB.Element.Name.GetValue(symbol))
    except Exception:
        pass
    try:
        return DB.Element.Name.GetValue(element)
    except Exception:
        return u'-'


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


def planar_facing(element):
    """Fronte dell'elemento proiettato in pianta e normalizzato.

    None se il fronte e' quasi verticale: un diffusore a controsoffitto non
    ha un verso in pianta da allineare a un muro, e traslarlo senza ruotarlo
    produrrebbe uno spostamento arbitrario.
    """
    try:
        facing = element.FacingOrientation
    except Exception:
        return None
    if facing is None:
        return None
    planar = DB.XYZ(facing.X, facing.Y, 0.0)
    if planar.GetLength() < VERTICAL_FACING_TOL:
        return None
    return planar.Normalize()


def connected_connectors(element):
    """Numero di connettori collegati, 0 se la famiglia non ne ha."""
    try:
        mep_model = element.MEPModel
    except Exception:
        return 0
    if mep_model is None:
        return 0
    try:
        manager = mep_model.ConnectorManager
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
    """Motivo di esclusione legato all'host, None se l'elemento e' libero."""
    try:
        host = element.Host
    except Exception:
        return None
    if host is None:
        return None
    if isinstance(host, DB.Wall):
        return R_HOSTED_WALL.format(element_id_value(host.Id))
    host_category = u'-'
    try:
        if host.Category is not None:
            host_category = host.Category.Name
    except Exception:
        pass
    return R_HOSTED_OTHER.format(host_category, element_id_value(host.Id))


def edit_skip_reason(element, active_design_option_id):
    """Motivo di esclusione legato allo stato di modificabilita'."""
    try:
        if element.Pinned:
            return R_PINNED
    except Exception:
        pass

    try:
        group_id = element.GroupId
        if group_id is not None and group_id != DB.ElementId.InvalidElementId:
            group = doc.GetElement(group_id)
            group_name = u'?'
            if group is not None:
                try:
                    group_name = DB.Element.Name.GetValue(group)
                except Exception:
                    group_name = u'{}'.format(element_id_value(group_id))
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
            if element_id_value(design_option.Id) != \
                    element_id_value(active_design_option_id):
                return R_DESIGN_OPTION
    except Exception:
        pass

    # GetCheckoutStatus va chiamato fuori transazione: la documentazione
    # avverte che il valore e' cache locale e non affidabile dentro una
    # transazione locale.
    try:
        if doc.IsWorkshared:
            status = DB.WorksharingUtils.GetCheckoutStatus(doc, element.Id)
            if status == DB.CheckoutStatus.OwnedByOtherUser:
                return R_BORROWED
    except Exception:
        pass

    return None


def signed_angle_about_z(vector_from, vector_to):
    """Angolo firmato attorno a +Z, positivo antiorario visto dall'alto."""
    cross_z = vector_from.X * vector_to.Y - vector_from.Y * vector_to.X
    dot = vector_from.X * vector_to.X + vector_from.Y * vector_to.Y
    return math.atan2(cross_z, dot)


def minimal_rotation(facing, target_normal):
    """Rotazione minima per rendere l'elemento perpendicolare al muro.

    Fra normale uscente e sua opposta si scegle quella piu' vicina
    all'orientamento attuale, quindi l'angolo resta in -90..+90 gradi e un
    elemento gia' orientato bene non viene ribaltato di 180 gradi.
    """
    if facing.DotProduct(target_normal) >= 0.0:
        target = target_normal
    else:
        target = target_normal.Negate()
    return signed_angle_about_z(facing, target), target


def all_wall_hits(wall_infos, point):
    """Test del punto contro tutti i muri, ordinati per vicinanza alla faccia.

    L'ordinamento usa il VALORE ASSOLUTO della distanza dalla faccia: un
    elemento immerso in un muro spesso ha distanza molto negativa e
    vincerebbe sempre contro un muro adiacente a pochi millimetri. L'id del
    muro come secondo criterio rende l'esito riproducibile in caso di parita'.

    Il test costa una proiezione per muro, quindi viene fatto UNA VOLTA per
    elemento e il risultato viene riusato da tutte le guardie.
    """
    hits = []
    for wall_info in wall_infos:
        hit = test_point_against_wall(wall_info, point)
        if hit is not None:
            hits.append(hit)
    hits.sort(key=lambda h: (abs(h.face_distance),
                             element_id_value(h.wall_info.wall_id)))
    return hits


def skipped_with_context(element, category_name, category_key, reason, hits):
    """Elemento escluso, con il muro piu' vicino e la distanza misurata.

    La distanza viene riportata anche quando il motivo non c'entra con la
    distanza: un elemento bloccato a 62 mm vale la pena di sbloccarlo, uno a
    290 mm probabilmente no.
    """
    nearest = hits[0] if hits else None
    return SkippedElement(
        element, category_name, category_key, reason,
        nearest.wall_info.wall_id if nearest else None,
        nearest.face_distance if nearest else None)


class PlanContext(object):
    """Stato condiviso dalla fase di analisi, calcolato una volta sola."""

    def __init__(self, region_ids, active_design_option_id):
        self.region_ids = region_ids
        self.active_design_option_id = active_design_option_id
        # Indicizzati per id: la stessa partizione scartata ricorre per molti
        # elementi e nel resoconto deve comparire una volta sola.
        self.partition_problems = {}

    def note_problem(self, partition, reason):
        self.partition_problems[element_id_value(partition.Id)] = (
            partition, reason)


def plan_element(element, options, context):
    """Decisione completa per un singolo elemento indicato dall'utente."""
    category_name, category_key = category_of(element)

    point = insertion_point(element)
    if point is None:
        return SkippedElement(element, category_name, category_key, R_NO_POINT)

    # La ricerca parte dall'ELEMENTO: si raccolgono le partizioni vicine a
    # lui. Nella prima versione l'elenco dei muri arrivava gia' fatto da
    # monte, perche' era l'utente a sceglierli.
    wall_infos, problems = partitions_near_point(
        point, options.tolerance_internal, context.region_ids)
    for partition, problem_reason in problems:
        context.note_problem(partition, problem_reason)

    # Le proiezioni si calcolano subito, prima delle guardie, cosi' ogni riga
    # del resoconto porta la distanza misurata anche quando il motivo di
    # scarto non c'entra con la distanza: un elemento bloccato a 62 mm vale
    # la pena di sbloccarlo, uno a 290 mm probabilmente no.
    hits = all_wall_hits(wall_infos, point)

    reason = host_skip_reason(element)
    if reason is not None:
        return skipped_with_context(element, category_name, category_key,
                                    reason, hits)

    reason = edit_skip_reason(element, context.active_design_option_id)
    if reason is not None:
        return skipped_with_context(element, category_name, category_key,
                                    reason, hits)

    if not wall_infos:
        return SkippedElement(element, category_name, category_key,
                              R_NO_PARTITION)

    # Secondo test verticale, preciso: la fase larga ha confrontato il
    # bounding box con l'outline, che e' intersezione, non contenimento.
    if not z_contained(z_extent(element), wall_infos):
        return skipped_with_context(element, category_name, category_key,
                                    R_OUT_OF_Z, hits)

    connected = connected_connectors(element)
    if options.skip_connected and connected > 0:
        return skipped_with_context(element, category_name, category_key,
                                    R_CONNECTED.format(connected), hits)

    facing = planar_facing(element)
    if facing is None:
        return skipped_with_context(element, category_name, category_key,
                                    R_FACING_VERTICAL, hits)

    if not hits:
        # Partizioni trovate, ma nessuna ha prodotto una proiezione
        # utilizzabile. Va riportato, non scartato in silenzio.
        return SkippedElement(element, category_name, category_key,
                              R_NO_PROJECTION)

    qualifying = [h for h in hits
                  if h.face_distance <= options.tolerance_internal
                  and h.beyond <= WALL_END_TOL]

    if not qualifying:
        in_range = [h for h in hits if h.beyond <= WALL_END_TOL]
        if not in_range:
            nearest = hits[0]
            return skipped_with_context(
                element, category_name, category_key,
                R_BEYOND_END.format(format_mm(nearest.beyond)), hits)

        # Oltre la tolleranza. Nella prima versione era informativo, perche'
        # era lo strumento a pescare gli elementi. Ora l'elemento lo ha
        # indicato l'utente, quindi e' uno SCARTO che deve vedere.
        nearest = in_range[0]
        return OverTolerance(
            element, category_name, category_key,
            nearest.wall_info.wall_id,
            nearest.wall_info.label,
            nearest.face_distance,
            nearest.face_distance - options.tolerance_internal)

    best = qualifying[0]
    second = qualifying[1] if len(qualifying) > 1 else None
    wall_info = best.wall_info
    normal_side = best.normal_ext.Multiply(best.side)

    # Punto bersaglio: sulla faccia del muro, stessa posizione lungo il muro,
    # stessa quota. La componente Z e' invariata per costruzione.
    target_signed = best.side * wall_info.half_width
    delta = target_signed - best.signed_center
    target_point = DB.XYZ(point.X + best.normal_ext.X * delta,
                          point.Y + best.normal_ext.Y * delta,
                          point.Z)

    record = PlannedMove()
    record.element = element
    record.element_id = element.Id
    record.category_name = category_name
    record.category_key = category_key
    record.type_name = type_name_of(element)

    record.wall_id = wall_info.wall_id
    record.wall_label = wall_info.label
    record.face_side = u'esterna' if best.side > 0 else u'interna'

    record.point_before = point
    record.distance_before = best.face_distance
    record.target_point = target_point
    record.distance_after = 0.0
    record.translation_length = abs(delta)
    record.connected = connected

    if options.apply_rotation:
        angle, _target = minimal_rotation(facing, normal_side)
        record.rotation_rad = angle
    else:
        record.rotation_rad = 0.0
    record.rotation_applicable = True

    record.needs_move = record.translation_length > POSITION_TOL
    record.needs_rotation = abs(record.rotation_rad) > ANGLE_TOL
    if not record.needs_move and not record.needs_rotation:
        record.status = PlannedMove.ALREADY_OK

    # Segnalazione di ambiguita': due muri praticamente equidistanti sono il
    # caso che l'utente controllera' per primo quando qualcosa sembrera'
    # sbagliato, quindi va dichiarato invece di lasciarlo scoprire.
    if second is not None:
        record.competing_walls = len(qualifying)
        record.second_distance = second.face_distance
        gap = abs(second.face_distance) - abs(best.face_distance)
        if gap <= AMBIGUITY_TOL:
            record.note = u'ambiguo: secondo muro a {}'.format(
                format_mm(second.face_distance))

    return record


def build_plan(elements, options, context):
    """Analisi completa di tutti gli elementi indicati. Nessuna transazione.

    Ogni elemento selezionato finisce in esattamente una delle quattro liste
    del risultato: nessuna esclusione silenziosa. L'utente ha puntato questi
    oggetti uno per uno, quindi un elemento che non compare nel resoconto
    sarebbe un difetto.
    """
    result = PlanResult()
    result.candidate_count = len(elements)

    selected_keys = set()
    for built_in_category in options.categories:
        selected_keys.add(element_id_value(DB.ElementId(built_in_category)))

    with forms.ProgressBar(title='Analisi... ({value} di {max_value})',
                           cancellable=True) as progress:
        total = len(elements)
        for index, element in enumerate(elements):
            if progress.cancelled:
                raise UserWarning(u'annullato')
            progress.update_progress(index + 1, total)

            category_name, category_key = category_of(element)

            if category_key is None or category_key not in MEP_CATEGORY_KEYS:
                result.skipped.append(SkippedElement(
                    element, category_name, category_key,
                    R_CATEGORY_NOT_HANDLED))
                continue

            if category_key not in selected_keys:
                result.skipped.append(SkippedElement(
                    element, category_name, category_key,
                    R_CATEGORY_NOT_SELECTED))
                continue

            try:
                outcome = plan_element(element, options, context)
            except Exception as error:
                result.skipped.append(SkippedElement(
                    element, category_name, category_key,
                    R_ANALYSIS_ERROR.format(u'{}'.format(error)[:160])))
                continue

            if outcome is None:
                continue
            if isinstance(outcome, SkippedElement):
                result.skipped.append(outcome)
            elif isinstance(outcome, OverTolerance):
                result.over_tolerance_items.append(outcome)
            elif outcome.status == PlannedMove.ALREADY_OK:
                result.already_ok.append(outcome)
            else:
                result.planned.append(outcome)

    result.wall_problems = list(context.partition_problems.values())
    result.per_category = summarise_per_category(result, options)
    return result


def summarise_per_category(result, options):
    """Conteggi per categoria, per la tabella di riepilogo."""
    summary = {}

    def bucket(key, name):
        if key not in summary:
            summary[key] = {'name': name, 'candidates': 0, 'moved': 0,
                            'ok': 0, 'skipped': 0, 'near': 0}
        return summary[key]

    for built_in_category in options.categories:
        key = element_id_value(DB.ElementId(built_in_category))
        bucket(key, category_display_name(built_in_category))

    for record in result.planned:
        row = bucket(record.category_key, record.category_name)
        row['moved'] += 1
        row['candidates'] += 1
    for record in result.already_ok:
        row = bucket(record.category_key, record.category_name)
        row['ok'] += 1
        row['candidates'] += 1
    for record in result.skipped:
        row = bucket(record.category_key, record.category_name)
        row['skipped'] += 1
        row['candidates'] += 1
    for record in result.over_tolerance_items:
        row = bucket(record.category_key, record.category_name)
        row['near'] += 1
        row['candidates'] += 1

    return summary


# =========================================================================
# GESTIONE DEI FAILURE DI REVIT
# =========================================================================

class AlignFailurePreprocessor(DB.IFailuresPreprocessor):
    """Registra gli avvisi di Revit senza applicare nessuna risoluzione.

    Non si usa il FailureSwallower di pyRevit: la sua lista RESOLUTION_TYPES
    include UnlockConstraints e DeleteElements, quindi su un avviso
    "Constraints are not satisfied" sbloccherebbe in silenzio le quote e gli
    allineamenti dell'utente. Qui vengono eliminati soltanto i warning privi
    di risoluzioni, che sono quelli che aprirebbero un popup e bloccherebbero
    il lotto a meta'.
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

            is_error = False
            try:
                is_error = (severity == DB.FailureSeverity.Error)
            except Exception:
                is_error = False
            if is_error:
                has_error = True

            record = (description, ids, u'errore' if is_error else u'avviso')
            if record not in self.messages:
                self.messages.append(record)

            has_resolutions = True
            try:
                has_resolutions = failure.HasResolutions()
            except Exception:
                has_resolutions = True

            is_warning = False
            try:
                is_warning = (severity == DB.FailureSeverity.Warning)
            except Exception:
                is_warning = False

            if is_warning and not has_resolutions:
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
    """Applica uno spostamento gia' calcolato.

    Prima la rotazione, con l'asse nel punto di inserimento originale che e'
    noto con certezza; poi la traslazione, ricalcolata dal punto CORRENTE
    verso il bersaglio assoluto, cosi' ogni deriva introdotta dalla rotazione
    si autocorregge. La componente Z del vettore e' sempre zero.
    """
    sub = DB.SubTransaction(doc)
    sub.Start()
    try:
        if record.needs_rotation:
            origin = record.point_before
            axis = DB.Line.CreateBound(
                origin,
                DB.XYZ(origin.X, origin.Y, origin.Z + 1.0))
            DB.ElementTransformUtils.RotateElement(
                doc, record.element_id, axis, record.rotation_rad)
            record.applied_rotation = True
            if FORCE_REGEN:
                doc.Regenerate()

        if record.needs_move:
            current = insertion_point(record.element)
            if current is None:
                current = record.point_before
            vector = DB.XYZ(record.target_point.X - current.X,
                            record.target_point.Y - current.Y,
                            0.0)
            if vector.GetLength() > POSITION_TOL:
                DB.ElementTransformUtils.MoveElement(
                    doc, record.element_id, vector)
            record.applied_move = True

        sub.Commit()
        return True
    except Exception as error:
        try:
            sub.RollBack()
        except Exception:
            pass
        record.applied_move = False
        record.applied_rotation = False
        record.error = u'{}'.format(error)[:180]
        return False


def apply_all(planned):
    """Applica tutti i movimenti pianificati. Ritorna (riusciti, falliti)."""
    ok = 0
    failed = 0
    with forms.ProgressBar(title='Allineamento... ({value} di {max_value})',
                           cancellable=True) as progress:
        total = len(planned)
        for index, record in enumerate(planned):
            if progress.cancelled:
                raise UserWarning(u'annullato')
            progress.update_progress(index + 1, total)
            if apply_record(record):
                ok += 1
            else:
                failed += 1
    return ok, failed


# =========================================================================
# SELEZIONE DEI MURI E FINESTRA DI DIALOGO
# =========================================================================

class MEPSelectionFilter(UI.Selection.ISelectionFilter):
    """Ammette solo le categorie MEP gestite.

    Confronto per id di categoria e non per nome, cosi' funziona anche sulle
    installazioni Revit localizzate. Il filtro agisce sulla selezione
    grafica: impedisce di indicare per sbaglio un muro o una porta, e quindi
    non c'e' niente da scartare a valle. La selezione fatta PRIMA di
    lanciare il comando non passa di qui, quindi viene filtrata nel piano,
    con il motivo scritto nel resoconto.
    """

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

    Non si filtra per categoria: l'utente ha selezionato questi oggetti, e
    scartarli in silenzio perche' fuori categoria sarebbe una sorpresa. Il
    filtro avviene nel piano, dove ogni scarto porta il suo motivo.
    """
    elements = []
    try:
        for element in revit.get_selection():
            if element is None:
                continue
            if isinstance(element, DB.ElementType):
                continue
            elements.append(element)
    except Exception:
        pass
    return elements


def pick_elements():
    """Selezione grafica degli elementi MEP. Vuota se l'utente preme Esc."""
    with forms.WarningBar(title='Seleziona gli elementi MEP da allineare, '
                                'poi premi Finish'):
        try:
            references = uidoc.Selection.PickObjects(
                UI.Selection.ObjectType.Element,
                MEPSelectionFilter(),
                'Seleziona gli elementi MEP da allineare')
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

    # Gli elementi scelti restano selezionati: rilanciando il comando, per
    # esempio per applicare dopo una simulazione, non serve ripetere la
    # selezione grafica.
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
        return elements, u'dalla selezione corrente'

    elements = pick_elements()
    if not elements:
        script.exit()
    return elements, u'scelti con la selezione grafica'


class MEPAlignWindow(forms.WPFWindow):
    """Finestra di dialogo definita in MEPAlignWindow.xaml."""

    # Attributo di classe: gli handler delle CheckBox scattano durante il
    # popolamento iniziale, prima che __init__ abbia finito.
    _ready = False

    def __init__(self, xaml_file, elements, source_label, category_info,
                 out_of_scope):
        forms.WPFWindow.__init__(self, xaml_file)

        self.options = None
        self._checks = []
        self._category_info = category_info

        self._setup_selection(elements, source_label, out_of_scope)
        self._setup_categories(category_info)

        self._ready = True
        self._refresh_count()

    # --- popolamento ---------------------------------------------------

    def _setup_selection(self, elements, source_label, out_of_scope):
        self.tb_walls_info.Text = u'{} elementi selezionati ({}).'.format(
            len(elements), source_label)

        detail = (u'Per ognuno lo strumento cerca la partizione verticale '
                  u'piu\' vicina. Gli elementi oltre la tolleranza vengono '
                  u'saltati e riportati nel resoconto.')
        if out_of_scope:
            detail += (u' {} elementi selezionati non appartengono alle '
                       u'categorie gestite e verranno saltati.'.format(
                           out_of_scope))
        self.tb_walls_extent.Text = detail

    def _setup_categories(self, category_info):
        for built_in_category, label, count in category_info:
            check = Controls.CheckBox()
            check.Content = u'{}  ({})'.format(label, count)
            check.Tag = built_in_category
            check.Margin = Thickness(2, 2, 2, 2)
            check.IsChecked = True
            check.Checked += self.on_category_toggled
            check.Unchecked += self.on_category_toggled
            self.sp_categories.Children.Add(check)
            self._checks.append(check)

    def _refresh_count(self):
        if not self._ready:
            return
        selected = [c for c in self._checks if c.IsChecked]
        candidates = 0
        for check in selected:
            for built_in_category, _label, count in self._category_info:
                if built_in_category == check.Tag:
                    candidates += count
                    break
        self.tb_cat_count.Text = u'{} categorie su {}, {} elementi ' \
                                 u'candidati'.format(len(selected),
                                                     len(self._checks),
                                                     candidates)

    # --- handler -------------------------------------------------------

    def on_category_toggled(self, sender, args):
        self._refresh_count()

    def on_select_all(self, sender, args):
        for check in self._checks:
            check.IsChecked = True
        self._refresh_count()

    def on_select_none(self, sender, args):
        for check in self._checks:
            check.IsChecked = False
        self._refresh_count()

    def _parse_tolerance(self):
        """(valore in cm, errore). Accetta virgola o punto."""
        raw = (self.tb_threshold.Text or u'').strip().replace(u',', u'.')
        if not raw:
            return None, u'Inserisci la distanza massima in centimetri.'
        try:
            value = float(raw)
        except ValueError:
            return None, u'La distanza massima non e\' un numero valido.'
        if value <= 0.0:
            return None, u'La distanza massima deve essere maggiore di zero.'
        if value > MAX_TOLERANCE_CM:
            return None, u'La distanza massima sembra fuori scala ' \
                         u'(oltre {:.0f} cm).'.format(MAX_TOLERANCE_CM)
        return value, None

    def on_run(self, sender, args):
        tolerance_cm, error = self._parse_tolerance()
        if error:
            forms.alert(error, title=u'Valore non valido')
            return

        categories = [c.Tag for c in self._checks if c.IsChecked]
        if not categories:
            forms.alert(u'Seleziona almeno una categoria da allineare.',
                        title=u'Nessuna categoria')
            return

        self.options = AlignOptions(
            categories,
            tolerance_cm,
            bool(self.chk_rotate.IsChecked),
            bool(self.chk_skip_connected.IsChecked),
            bool(self.chk_dryrun.IsChecked))
        self.Close()

    def on_cancel(self, sender, args):
        self.options = None
        self.Close()


def resolve_xaml_path():
    path = None
    try:
        path = script.get_bundle_file(XAML_FILE_NAME)
    except Exception:
        path = None
    if not path:
        path = op.join(op.dirname(__file__), XAML_FILE_NAME)
    if not op.isfile(path):
        forms.alert(
            u'File grafica non trovato:\n\n{}\n\nDeve stare nella stessa '
            u'cartella dello script.'.format(path),
            title=u'Grafica mancante', exitscript=True)
    return path


# =========================================================================
# RESOCONTO
# =========================================================================

def print_header(options, elements, source_label, result):
    if options.dry_run:
        output.print_md(u'# Allineamento MEP alle partizioni - simulazione')
        output.print_md(
            u'**Nessuna modifica e\' stata applicata al modello.** '
            u'I valori seguenti sono il risultato che verrebbe prodotto.')
    else:
        output.print_md(u'# Allineamento MEP alle partizioni - resoconto')

    lines = [
        u'- Elementi selezionati: **{}** ({})'.format(
            len(elements), source_label),
        u'- Tolleranza: **{:.0f} cm** dalla faccia della partizione'.format(
            options.tolerance_cm),
        u'- Oltre la tolleranza l\'elemento viene saltato, non spostato',
        u'- Posizione finale: punto di inserimento sulla faccia',
        u'- Rotazione: {}'.format(
            u'attiva, minima, mai oltre 90 gradi'
            if options.apply_rotation else u'disattivata'),
        u'- Elementi con connettori collegati: {}'.format(
            u'saltati' if options.skip_connected else u'elaborati'),
        u'- Categorie elaborate: **{}** su {}'.format(
            len(options.categories), len(MEP_CATEGORIES)),
        u'- La quota Z non viene modificata',
    ]
    if result.wall_problems:
        lines.append(
            u'- Partizioni scartate: **{}** (tabella in fondo)'.format(
                len(result.wall_problems)))
    output.print_md(u'\n'.join(lines))


def print_category_summary(result, options):
    output.print_md(u'## Riepilogo per categoria')
    moved_header = u'Da allineare' if options.dry_run else u'Allineati'

    rows = []
    totals = [0, 0, 0, 0, 0]
    for key in sorted(result.per_category.keys(),
                      key=lambda k: result.per_category[k]['name']):
        row = result.per_category[key]
        rows.append([row['name'],
                     str(row['candidates']),
                     str(row['moved']),
                     str(row['ok']),
                     str(row['skipped']),
                     str(row['near'])])
        totals[0] += row['candidates']
        totals[1] += row['moved']
        totals[2] += row['ok']
        totals[3] += row['skipped']
        totals[4] += row['near']

    rows.append([u'**Totale**'] + [str(v) for v in totals])
    output.print_table(
        table_data=rows,
        title='',
        columns=[u'Categoria', u'Selezionati', moved_header,
                 u'Gia\' allineati', u'Ignorati', u'Oltre tolleranza'])


def print_moves_table(result, options):
    if not result.planned:
        return
    output.print_md(u'## Elementi {}'.format(
        u'da allineare' if options.dry_run else u'allineati'))

    ordered = sorted(result.planned,
                     key=lambda r: (element_id_value(r.wall_id),
                                    -abs(r.distance_before)))
    shown = ordered[:MAX_MOVED_ROWS]

    columns = [u'Elemento', u'Categoria', u'Tipo', u'Muro', u'Faccia',
               u'Dist. prima', u'Dist. dopo', u'Spostamento', u'Rotazione',
               u'Connettori', u'Note']
    if not options.dry_run:
        columns.append(u'Esito')

    rows = []
    for record in shown:
        row = [
            output.linkify(record.element_id),
            record.category_name,
            record.type_name,
            u'{} {}'.format(output.linkify(record.wall_id), record.wall_label),
            record.face_side,
            format_mm(record.distance_before),
            format_mm(record.distance_after),
            format_mm(record.translation_length),
            format_deg(record.rotation_rad) if record.needs_rotation else u'-',
            str(record.connected) if record.connected else u'-',
            record.note or u'',
        ]
        if not options.dry_run:
            if record.error:
                outcome = u'errore: {}'.format(record.error)
            elif record.needs_rotation and not record.applied_rotation:
                outcome = u'spostato ma non ruotato'
            elif record.needs_move and not record.applied_move:
                outcome = u'ruotato ma non spostato'
            else:
                outcome = u'OK'
            row.append(outcome)
        rows.append(row)

    output.print_table(table_data=rows, title='', columns=columns)

    remaining = len(ordered) - len(shown)
    if remaining > 0:
        output.print_md(u'_...e altri {} elementi non elencati. '
                        u'Sono stati comunque elaborati tutti._'.format(
                            remaining))


def print_already_ok(result):
    if not result.already_ok:
        return
    output.print_md(
        u'## Elementi gia\' allineati\n\n'
        u'**{}** elementi risultavano gia\' a posto (entro {:.0f} mm dalla '
        u'faccia e {:.1f} gradi) e non sono stati toccati.'.format(
            len(result.already_ok), POSITION_TOL_MM, ANGLE_TOL_DEG))


def print_skipped_table(result):
    if not result.skipped:
        return
    output.print_md(u'## Elementi ignorati')

    shown = result.skipped[:MAX_SKIPPED_ROWS]
    rows = []
    for record in shown:
        rows.append([
            output.linkify(record.element_id),
            record.category_name,
            record.reason,
            output.linkify(record.wall_id) if record.wall_id else u'-',
            format_mm(record.distance) if record.distance is not None else u'-',
        ])
    output.print_table(
        table_data=rows,
        title='',
        columns=[u'Elemento', u'Categoria', u'Motivo', u'Muro piu\' vicino',
                 u'Distanza misurata'])

    remaining = len(result.skipped) - len(shown)
    if remaining > 0:
        output.print_md(u'_...e altri {} elementi ignorati non '
                        u'elencati._'.format(remaining))


def print_over_tolerance_table(result, options):
    if not result.over_tolerance_items:
        return
    output.print_md(u'## Saltati perche\' oltre la tolleranza')

    ordered = sorted(result.over_tolerance_items, key=lambda r: r.distance)
    shown = ordered[:MAX_OVER_TOLERANCE_ROWS]
    rows = []
    for record in shown:
        rows.append([
            output.linkify(record.element_id),
            record.category_name,
            u'{} {}'.format(output.linkify(record.wall_id), record.wall_label),
            format_mm(record.distance),
            u'+{}'.format(format_mm(record.excess)),
        ])
    output.print_table(
        table_data=rows,
        title='',
        columns=[u'Elemento', u'Categoria', u'Muro piu\' vicino',
                 u'Distanza', u'Eccedenza'])

    remaining = len(ordered) - len(shown)
    if remaining > 0:
        output.print_md(u'_...e altri {} elementi non elencati._'.format(
            remaining))

    suggestion = tolerance_suggestion(ordered, options)
    if suggestion:
        output.print_md(suggestion)


def tolerance_suggestion(ordered_over_tolerance_items, options):
    """Riga che dice di quanto alzare la soglia e quanto si recupererebbe."""
    if not ordered_over_tolerance_items:
        return None
    index = int(len(ordered_over_tolerance_items) * 0.8)
    if index >= len(ordered_over_tolerance_items):
        index = len(ordered_over_tolerance_items) - 1
    target_mm = internal_to_mm(ordered_over_tolerance_items[index].distance)
    step_cm = 5.0
    candidate_cm = math.ceil((target_mm / 10.0) / step_cm) * step_cm
    if candidate_cm <= options.tolerance_cm:
        candidate_cm = options.tolerance_cm + step_cm
    limit = cm_to_internal(candidate_cm)
    recovered = len([r for r in ordered_over_tolerance_items if r.distance <= limit])
    if not recovered:
        return None
    return u'_Portando la tolleranza a **{:.0f} cm** rientrerebbero altri ' \
           u'**{}** elementi._'.format(candidate_cm, recovered)


def print_wall_problems(result):
    if not result.wall_problems:
        return
    output.print_md(u'## Partizioni scartate')
    rows = []
    for wall, reason in result.wall_problems:
        rows.append([output.linkify(wall.Id), wall_label(wall), reason])
    output.print_table(
        table_data=rows,
        title='',
        columns=[u'Partizione', u'Tipo', u'Motivo'])


def print_revit_failures(preprocessor):
    if preprocessor is None or not preprocessor.messages:
        return
    output.print_md(u'## Avvisi di Revit durante la modifica')
    rows = []
    for description, ids, kind in preprocessor.messages:
        links = u', '.join([output.linkify(i) for i in ids[:5]]) or u'-'
        if len(ids) > 5:
            links += u' ...'
        rows.append([kind, description, links])
    output.print_table(
        table_data=rows,
        title='',
        columns=[u'Tipo', u'Descrizione', u'Elementi'])
    output.print_md(u'_Nessun vincolo e\' stato sbloccato o eliminato: lo '
                    u'strumento non applica risoluzioni automatiche._')


def print_warnings():
    if not WARNINGS:
        return
    output.print_md(u'## Avvisi')
    for message in WARNINGS:
        output.print_md(u'- {}'.format(message))


def print_report(options, elements, source_label, result, preprocessor,
                 applied_ok, applied_failed, rolled_back=False):
    """Resoconto completo. Emesso sempre, anche quando non c'e' nulla da fare."""
    output.close_others()

    # Se la transazione e' stata annullata, nessun elemento e' stato toccato:
    # la colonna "Esito" non deve dire il contrario.
    if rolled_back:
        for record in result.planned:
            record.applied_move = False
            record.applied_rotation = False
            if not record.error:
                record.error = u'transazione annullata'

    print_header(options, elements, source_label, result)
    print_category_summary(result, options)
    print_moves_table(result, options)
    print_already_ok(result)
    print_skipped_table(result)
    print_over_tolerance_table(result, options)
    print_wall_problems(result)
    print_revit_failures(preprocessor)
    print_warnings()

    output.print_md(u'## Esito')
    if rolled_back:
        output.print_md(
            u'**La modifica e\' stata annullata: il modello non e\' stato '
            u'toccato.** I {} elementi pianificati sono ancora nella loro '
            u'posizione originale. Il motivo e\' negli avvisi qui '
            u'sopra.'.format(len(result.planned)))
    elif not result.planned:
        output.print_md(
            u'Nessun elemento da allineare. Le tabelle "Elementi ignorati" e '
            u'"Saltati perche\' oltre la tolleranza" dicono se il problema '
            u'e\' la tolleranza, '
            u'le categorie, i muri scelti o gli elementi ospitati.')
    elif options.dry_run:
        output.print_md(
            u'**{}** elementi verrebbero allineati. Rieseguire il comando '
            u'senza la spunta "Simulazione" per applicare la '
            u'modifica.'.format(len(result.planned)))
    else:
        output.print_md(
            u'Elementi allineati: **{}** su {} pianificati.'.format(
                applied_ok, len(result.planned)))
        if applied_failed:
            output.print_md(
                u'Elementi non modificati per un errore di Revit: '
                u'**{}**. Il motivo e\' nella colonna "Esito".'.format(
                    applied_failed))
        output.print_md(
            u'La modifica e\' raggruppata in un unico passo di annullamento, '
            u'nominato "{}".'.format(TRANSACTION_NAME))


# =========================================================================
# PROGRAMMA PRINCIPALE
# =========================================================================

def build_non_pickable_view_types():
    types = []
    for name in ('Schedule', 'ColumnSchedule', 'PanelSchedule',
                 'DrawingSheet', 'Legend', 'Report', 'ProjectBrowser',
                 'SystemBrowser', 'Undefined'):
        view_type = getattr(DB.ViewType, name, None)
        if view_type is not None:
            types.append(view_type)
    return types


def check_preconditions():
    """Guardie di apertura, prima di qualunque interfaccia."""
    if doc is None:
        forms.alert(u'Nessun documento aperto.', exitscript=True)

    if doc.IsFamilyDocument:
        forms.alert(u'Comando non disponibile nell\'editor di famiglie.',
                    exitscript=True)

    version = revit_version()
    if version and version < 2022:
        forms.alert(u'Lo strumento richiede Revit 2022 o successivo. '
                    u'Versione rilevata: {}.'.format(version),
                    exitscript=True)

    if not MEP_CATEGORIES:
        forms.alert(u'Nessuna delle categorie configurate e\' disponibile '
                    u'in questa versione di Revit.', exitscript=True)

    active_view = doc.ActiveView
    if active_view is None:
        forms.alert(u'Nessuna vista attiva.', exitscript=True)
    if active_view.ViewType in build_non_pickable_view_types():
        forms.alert(
            u'Attivare una vista grafica (pianta, sezione, 3D) prima di '
            u'lanciare il comando: la scelta dei muri avviene nella vista.',
            exitscript=True)


def count_per_category(elements):
    """[(BuiltInCategory, nome, conteggio)] per le caselle della finestra.

    I conteggi vengono dalla SELEZIONE dell'utente, non da una ricerca nel
    modello: una categoria a zero dice subito che fra gli oggetti indicati
    non ce n'e' nessuno di quel tipo.
    """
    counts = {}
    for element in elements:
        _name, key = category_of(element)
        if key is None:
            continue
        counts[key] = counts.get(key, 0) + 1

    info = []
    for built_in_category in MEP_CATEGORIES:
        key = element_id_value(DB.ElementId(built_in_category))
        info.append((built_in_category,
                     category_display_name(built_in_category),
                     counts.get(key, 0)))
    info.sort(key=lambda row: row[1])
    return info


def out_of_scope_count(elements):
    """Quanti elementi selezionati stanno fuori dalle categorie gestite."""
    total = 0
    for element in elements:
        _name, key = category_of(element)
        if key is None or key not in MEP_CATEGORY_KEYS:
            total += 1
    return total


def active_design_option_id():
    try:
        return DB.DesignOption.GetActiveDesignOptionId(doc)
    except Exception:
        return DB.ElementId.InvalidElementId


def main():
    check_preconditions()

    elements, source_label = resolve_elements()

    points = []
    for element in elements:
        point = insertion_point(element)
        if point is not None:
            points.append(point)

    if not points:
        output.close_others()
        output.print_md(u'# Allineamento MEP ai muri')
        output.print_md(
            u'Nessuno dei **{}** elementi selezionati ha un punto di '
            u'inserimento: sono tutti basati su host, su piano di lavoro o '
            u'realizzati sul posto. Non c\'e\' niente da allineare in '
            u'pianta.'.format(len(elements)))
        forms.alert(
            u'Nessuno degli elementi selezionati ha un punto di '
            u'inserimento.',
            title=u'Niente da allineare', exitscript=True)

    window = MEPAlignWindow(resolve_xaml_path(), elements, source_label,
                            count_per_category(elements),
                            out_of_scope_count(elements))
    window.ShowDialog()

    options = window.options
    if options is None:
        script.exit()

    # Una sola passata sul documento per la regione occupata dalla
    # selezione; la restrizione per singolo elemento avviene poi su questo
    # insieme ridotto.
    region_ids = collect_partition_ids(points, options.tolerance_internal)
    context = PlanContext(region_ids, active_design_option_id())

    try:
        result = build_plan(elements, options, context)
    except UserWarning:
        forms.alert(
            u'Analisi annullata: il modello non e\' stato modificato.',
            title=u'Annullato', exitscript=True)

    preprocessor = None
    applied_ok = 0
    applied_failed = 0
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
            # Il preprocessor chiede ProceedWithRollBack in presenza di
            # errori: in quel caso il commit annulla tutto, e dichiarare
            # comunque N elementi allineati sarebbe falso.
            if status == DB.TransactionStatus.RolledBack:
                warn(u'Revit ha annullato la transazione per un errore: '
                     u'nessuna modifica e\' stata applicata. Il dettaglio e\' '
                     u'nella sezione "Avvisi di Revit durante la modifica".')
                applied_failed += applied_ok
                applied_ok = 0
                rolled_back = True
        except UserWarning:
            transaction.RollBack()
            warn(u'Operazione annullata dall\'utente: nessuna modifica '
                 u'applicata.')
            applied_ok = 0
            applied_failed = 0
            rolled_back = True
        except Exception as error:
            transaction.RollBack()
            warn(u'Transazione annullata per un errore: {}'.format(error))
            applied_ok = 0
            applied_failed = 0
            rolled_back = True

    print_report(options, elements, source_label, result, preprocessor,
                 applied_ok, applied_failed, rolled_back)

    if not result.planned:
        forms.alert(
            u'Nessun elemento da allineare.\n\n'
            u'Il resoconto nel pannello di output spiega il motivo elemento '
            u'per elemento.',
            title=u'Niente da fare')


if __name__ == '__main__':
    main()
