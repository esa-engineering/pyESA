# -*- coding: utf-8 -*-
"""Allineamento in pianta degli elementi MEP alle facce dei muri.

L'utente seleziona i dispositivi MEP da allineare. Per ognuno lo strumento
cerca la partizione verticale architettonica piu' vicina e porta il
dispositivo a filo di una delle sue due facce. Se la partizione piu' vicina e'
oltre la tolleranza indicata dall'utente, il dispositivo viene saltato e
riportato con la distanza misurata.

Quale delle due facce lo decide il FRONTE della famiglia, ricavato dal Room
Calculation Point: e' il motivo per cui lo spostamento puo' avvenire nei due
versi lungo la normale del muro, e non solo verso la faccia piu' vicina.

Sull'orientamento lo strumento interviene solo per RADDRIZZARE: porta il
fronte perpendicolare alla faccia bersaglio, e mai di piu' di
FRONT_MAX_ANGLE_DEG. Non e' un riorientamento, perche' una partizione viene
ammessa solo se il fronte le sta gia' entro quel cono: l'angolo da
recuperare non puo' superarlo. Un elemento senza Room Calculation Point non
viene ruotato affatto, perche' non c'e' nessun verso di cui fidarsi.

La quota Z non viene MAI modificata: e' il motivo per cui il primo elemento
architettonico gestito e' il muro. L'allineamento in quota e' un problema
diverso (serve un livello o un soffitto come riferimento) ed e' fuori ambito.

--------------------------------------------------------------------------
LOGICA APPLICATA
--------------------------------------------------------------------------

1. ELEMENTI. Si parte dagli elementi nella selezione corrente; se non ce ne
   sono, viene chiesta una selezione grafica limitata alle categorie gestite.
   Un elemento selezionato prima del comando e fuori da quelle categorie non
   viene ignorato in silenzio: compare fra i saltati con il suo motivo.

2. PARTIZIONI. Cercate nel documento host E nei modelli collegati. Una sola
   passata di raccolta per documento sulla regione occupata dalla selezione,
   poi un filtro di prossimita' per ogni elemento sugli id gia' raccolti.
   Host e collegamenti concorrono insieme: vince la partizione piu' vicina,
   da qualunque documento provenga. I WallInfo sono in cache per id, perche' la stessa
   partizione ricorre per tutti i dispositivi che le stanno davanti e il
   controllo di inclinazione puo' estrarne la geometria.
   Le partizioni che non coprono la quota del dispositivo vengono tolte dalla
   gara: un muretto basso non e' un bersaglio valido per un rilevatore in
   alto. Muri tenda, muri inclinati, contenitori di muri sovrapposti e muri
   senza linea di posizionamento sono scartati con il motivo nel resoconto.

3. TOLLERANZA. Distanza misurata in pianta fra il punto di inserimento
   dell'elemento e la faccia della partizione piu' vicina, non il suo asse.
   Un elemento che cade dentro lo spessore ha distanza negativa e quindi
   rientra sempre. Oltre la tolleranza l'elemento viene SALTATO: e' un esito
   che l'utente deve vedere, perche' quell'oggetto lo ha indicato lui.

4. FACCIA. Il muro e' scelto per distanza FRA QUELLI CHE IL DISPOSITIVO
   GUARDA, e quale delle sue due facce usare lo decide il fronte della
   famiglia. E se piu' istanze di muro sono a
   contatto, il bersaglio e' la faccia piu' esterna del pacchetto: vedi le
   due sezioni dedicate qui sotto.

5. TRASLAZIONE. Lungo la RETTA DI ANALISI, cioe' l'asse di vista del
   dispositivo, fino a portare il punto di inserimento sul piano della
   faccia bersaglio. Non lungo la normale del muro: un dispositivo che
   guarda la parete in obliquo scorre quindi anche lateralmente, e la sua
   posizione lungo il muro cambia. La quota resta invariata.

   Muovendosi in obliquo la corsa vale 1/cos della distanza da coprire, e
   questo e' il motivo per cui il fronte deve stare entro FRONT_MAX_ANGLE_DEG
   dalla perpendicolare: oltre, l'elemento scivolerebbe lungo il muro invece
   di avvicinarglisi. Con i 5 gradi attuali il termine 1/cos vale 1,004 e lo
   scivolamento laterale resta sotto il 9% della distanza: lo spostamento e'
   di fatto perpendicolare, e la retta pesa soprattutto sulla scelta del muro
   e della faccia. Senza un fronte leggibile non c'e' nessuna retta, e la
   traslazione torna a essere perpendicolare al muro.

--------------------------------------------------------------------------
QUALE DELLE DUE FACCE: IL FRONTE DAL ROOM CALCULATION POINT
--------------------------------------------------------------------------

Scegliere sempre la faccia piu' vicina al punto di inserimento e' corretto
solo finche' il dispositivo e' gia' modellato dalla parte giusta della
partizione. Non lo e' quando sta dentro lo spessore del muro, e non lo e'
quando e' stato inserito dal lato sbagliato: in quei casi il dispositivo
finisce a filo della faccia che ha alle spalle, cioe' dentro la stanza
sbagliata.

Serve quindi distinguere il fronte della famiglia dal suo retro, e
FacingOrientation non basta: dipende da come e' orientato il sistema di
riferimento con cui la famiglia e' stata autorata, che non e' una convenzione
rispettata da tutti i produttori.

Il Room Calculation Point e' il punto che Revit usa per stabilire in quale
locale sta il dispositivo, quindi per costruzione cade DAVANTI
all'apparecchio, dentro la stanza che serve. La direzione del fronte e' il
vettore che va dal punto di inserimento alla proiezione del punto di calcolo
sul piano orizzontale passante per il punto di inserimento. Si proietta
invece di usare il vettore nello spazio perche' su un dispositivo a parete il
punto di calcolo sta quasi sempre anche piu' in alto o piu' in basso
dell'origine, e la componente verticale falserebbe l'angolo rispetto alla
normale del muro.

Lungo quella direzione si traccia dal punto di inserimento una retta lunga
DUE VOLTE la tolleranza impostata dall'utente. Se la retta intercetta la
superficie del muro, il dispositivo guarda la partizione invece di
allontanarsene: va quindi portato sulla SECONDA faccia di finitura, non sulla
prima.

Il fronte non sceglie pero' solo la faccia: ESCLUDE le partizioni che
corrono parallele alla retta di analisi (faces_the_wall). Una parete
parallela alla retta e' una parete che il dispositivo ha di FIANCO, non
davanti: allinearcelo significherebbe spostarlo lungo una normale ortogonale
al suo asse di vista, cioe' proprio lo spostamento di traverso che la retta
serve a evitare.

L'esclusione e' secca, senza ripiego. Se dopo il filtro non resta nessuna
partizione, l'elemento viene SALTATO e riportato con R_NOT_FACED: meglio
lasciarlo dov'e' dicendolo, che allinearlo a un riferimento che non e' il
suo. Un ripiego "se non ne resta nessuna concorrono tutte" sembra prudente
ma riporta esattamente il difetto che il filtro doveva togliere, e lo fa
proprio nel caso peggiore, quello in cui l'unica parete vicina e' quella
sbagliata.

Fra le partizioni ammesse il criterio resta "vince la piu' vicina": il
fronte decide chi entra in gara, non chi la vince.

Il filtro guarda l'orientamento e non il verso, perche' la retta di analisi
e' una retta e non una semiretta: una parete davanti e una alle spalle sono
entrambe cose che il dispositivo "guarda", ed e' giusto cosi', visto che il
montaggio normale e' proprio con le spalle al muro.

Senza fronte leggibile non esiste nessuna retta, quindi non c'e' nulla a cui
una parete possa essere parallela: li' concorrono tutte, con il criterio
posizionale delle versioni precedenti.

Il test non usa la geometria del muro. Sotto una trasformazione rigida tutte
le grandezze in gioco sono scalari invarianti (lo scostamento firmato dalla
mezzeria, il semispessore e la componente del fronte lungo la normale),
quindi le partizioni collegate non richiedono nessuna conversione in piu'.
Vedi face_side_from_front().

La retta lunga il doppio della tolleranza copre da sola un cono di 60 gradi:
e' quindi abbondante rispetto al cono di FRONT_MAX_ANGLE_DEG entro cui una
parete conta come fronteggiata, e per un elemento fuori dal muro non e' mai
lei a decidere. Resta vincolante solo per un elemento sepolto dentro una
partizione piu' spessa della tolleranza, dove la retta puo' non raggiungere
la faccia di uscita.

Quando la famiglia non espone il Room Calculation Point, o quando l'autore
non lo ha spostato dall'origine, non c'e' nessuna retta: l'elemento viene
comunque allineato con il criterio posizionale, e la colonna "Fronte" del
resoconto dice con quale criterio.

--------------------------------------------------------------------------
PIU' ISTANZE DI MURO A CONTATTO
--------------------------------------------------------------------------

Una parete modellata come piu' istanze accostate - la muratura piu' il suo
rivestimento, due tramezzi affiancati, un contromuro tecnico - va trattata
come un pacchetto unico: il dispositivo si allinea alla faccia PIU' ESTERNA,
non alla prima che incontra. Senza questo, un apparecchio modellato dentro il
muro strutturale finirebbe a filo del suo rivestimento interno, cioe' dentro
la stratigrafia.

outermost_face() lavora su un asse orientato come la normale del muro
vincente, con l'origine nel punto di inserimento, su cui ogni partizione
parallela occupa un intervallo noto (span_along). Partendo dalla faccia
scelta il confine viene spinto in fuori finche' esiste una partizione che lo
contiene o lo sfiora e che si estende oltre.

Tre vincoli tengono la regola stretta:

  - solo partizioni PARALLELE entrano nel pacchetto (ADJACENT_PARALLEL_TOL);
  - solo partizioni su cui l'elemento si PROIETTA davvero, cioe' quelle con
    beyond entro tolleranza: una parete che finisce prima di arrivargli
    davanti non gli e' adiacente;
  - solo scarti sotto ADJACENT_GAP, due millimetri. Un'intercapedine vera e'
    molto piu' larga e non viene attraversata.

Il confine puo' solo essere spinto piu' in fuori: il muro vincente resta
quello scelto per distanza e la sua faccia resta quella scelta dal fronte.

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

5. ARBITRAGGIO FRA PARTIZIONI. Un elemento vicino a piu' partizioni viene
   assegnato a quella con il valore ASSOLUTO della distanza dalla faccia piu'
   piccolo. Usare il valore firmato sarebbe sbagliato: un elemento immerso in
   un muro spesso ha distanza molto negativa e vincerebbe sempre contro un
   muro adiacente a pochi millimetri. Con la ricerca automatica l'ambiguita'
   non e' piu' un caso raro: in un angolo due muri sono quasi equidistanti,
   e il resoconto la segnala nella colonna Note.

--------------------------------------------------------------------------
MODELLI COLLEGATI
--------------------------------------------------------------------------

Nei progetti MEP l'architettonico e' quasi sempre un collegamento, e un
FilteredElementCollector sul documento corrente non vede i suoi elementi.
Lo strumento raccoglie quindi le partizioni anche dai RevitLinkInstance
caricati.

Il punto chiave e' che sotto una trasformazione RIGIDA tutte le grandezze
scalari del calcolo sono INVARIANTI: la distanza dalla faccia, lo spessore
del muro, lo scostamento della linea di posizionamento e la sporgenza oltre
la testata valgono lo stesso nelle due terne. Quindi:

  - il punto di inserimento dell'elemento, che vive in coordinate host,
    viene portato nelle coordinate del collegamento con la trasformazione
    inversa;
  - la matematica gia' collaudata gira li' dentro senza sapere nulla dei
    link, e il muro non viene mai trasformato;
  - solo la NORMALE USCENTE viene riportata in coordinate host, perche' e'
    l'unica direzione che deve convivere con il fronte dell'elemento e con
    il punto bersaglio.

Un collegamento viene ignorato, con il motivo negli avvisi, quando e'
scaricato, quando la sua trasformazione non e' leggibile, quando e'
inclinato o capovolto (la normale riportata nell'host avrebbe una componente
verticale e lo spostamento cambierebbe la quota), quando e' speculare (i
versi esterno e interno risulterebbero invertiti) o quando e' in scala (le
distanze non sarebbero confrontabili con la tolleranza dell'utente).

I muri collegati sono usati SOLO come riferimento: la fase di applicazione
lavora esclusivamente sul documento host e non tocca mai un modello
collegato.

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

Prima il raddrizzamento, attorno a un asse verticale per il punto di
inserimento originale, che e' noto con certezza. Poi la traslazione,
ricalcolata dal punto CORRENTE verso un punto bersaglio ASSOLUTO memorizzato
in fase di analisi, invece di riusare il vettore calcolato allora: non e'
garantito che RotateElement lasci il punto di inserimento esattamente
invariato per ogni tipo di famiglia, e ricalcolando dopo la rotazione
qualunque deriva si autocorregge. La quota resta invariata per costruzione,
perche' la componente Z del vettore e' forzata a zero.

Se la rilettura del punto subito dopo RotateElement risultasse non
aggiornata, basta alzare FORCE_REGEN in testa allo script.

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
- I pannelli di facciata continua non sono ancora gestiti: il muro tenda ha
  spessore nullo e la faccia di riferimento andrebbe presa dal pannello.
- Lo strumento non cambia mai l'host di un elemento: se e' ospitato da un muro
  lo salta.
- Il raddrizzamento e' fine per costruzione, al massimo FRONT_MAX_ANGLE_DEG:
  un dispositivo montato davvero di traverso non viene raddrizzato, viene
  SALTATO prima, perche' nessuna partizione risulta fronteggiata. Le due cose
  sono la stessa regola vista da due lati.
- Si ruota il FRONTE, cioe' la direzione del Room Calculation Point, non
  FacingOrientation. La rotazione e' rigida e porta con se' anche il punto di
  calcolo, quindi il risultato non dipende dalla convenzione con cui la
  famiglia e' stata autorata. Il rovescio: se in una famiglia il punto e'
  autorato leggermente di sbieco, entro il cono, l'elemento viene ruotato di
  quel tanto anche se era montato bene.
- Lo spostamento puo' superare di molto la tolleranza, per tre motivi che si
  sommano: la faccia opposta aggiunge lo spessore del muro, il pacchetto
  murario aggiunge quello delle partizioni attraversate, e la corsa obliqua
  moltiplica tutto per 1/cos (con il cono a 5 gradi, al massimo 1,004).
  E' voluto: la tolleranza e' il
  raggio entro cui cercare la partizione, non un limite allo spostamento. La
  colonna "Spostamento" e le note del resoconto lo rendono visibile.
- Con la corsa lungo la retta la posizione del dispositivo LUNGO il muro non
  e' piu' invariante: un apparecchio che guarda la parete in obliquo scivola
  anche di lato. E' la conseguenza diretta di "spostamento esclusivamente
  lungo la retta calcolata", ma con il cono a 5 gradi lo scivolamento non
  supera il 9% della distanza da coprire.

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
from System.Windows import SystemParameters

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

XAML_FILE_NAME = 'MEPAlignToWallWindow.xaml'

TRANSACTION_NAME = u'Align MEP to walls'

DEFAULT_TOLERANCE_CM = 30.0     # soglia proposta nella finestra
MAX_TOLERANCE_CM = 500.0        # oltre e' quasi certamente un errore di battitura

POSITION_TOL_MM = 1.0           # sotto questo spostamento l'elemento e' gia' a posto
ANGLE_TOL_DEG = 0.1             # sotto questo angolo non si ruota
WALL_END_TOL_MM = 0.1           # sporgenza ammessa oltre l'estremita' del muro
Z_TOL_MM = 10.0                 # margine verticale sul test di contenimento
AMBIGUITY_TOL_MM = 20.0         # due muri entro questo scarto: caso segnalato

# Il vettore che va dal punto di inserimento al Room Calculation Point
# distingue il fronte della famiglia dal suo retro. Sotto questa lunghezza in
# pianta il punto di calcolo sta praticamente sull'origine della famiglia
# (l'autore non lo ha spostato) oppure gli sta esattamente sopra: in entrambi
# i casi non indica nessuna direzione utilizzabile.
FRONT_MIN_LENGTH_MM = 1.0

# Moltiplicatore della tolleranza che da' la lunghezza della retta di
# analisi tracciata lungo il fronte, come richiesto: lunga il doppio della
# tolleranza. Con il cono stretto qui sotto la retta e' abbondante per un
# elemento fuori dal muro, e resta vincolante solo per uno sepolto dentro
# una partizione piu' spessa della tolleranza.
FRONT_RAY_FACTOR = 2.0

# Apertura massima, in gradi, fra la retta di analisi e la PERPENDICOLARE
# al muro perche' quella partizione conti come fronteggiata. Oltre, la
# parete corre di fatto parallela alla retta e viene esclusa; se non ne
# resta nessuna l'elemento viene saltato e riportato.
#
# Non e' piu' derivata da FRONT_RAY_FACTOR come nella prima stesura. Una
# retta lunga il doppio della tolleranza copre da sola un cono di 60 gradi,
# ma un cono cosi' largo ammette scivolamenti laterali fino a 1,7 volte la
# distanza da coprire: troppo, per uno strumento che deve mettere a filo un
# dispositivo, non spostarlo di stanza. A 5 gradi la corsa supera la
# distanza dello 0,4% e lo scivolamento laterale resta sotto il 9%, quindi
# lo spostamento e' di fatto perpendicolare e la retta serve soprattutto a
# scegliere il muro e la faccia.
#
# La soglia decide anche SE agire, non solo come: e' il numero da rivedere
# se nel resoconto comparissero troppi elementi saltati per R_NOT_FACED.
FRONT_MAX_ANGLE_DEG = 5.0

# Due murature separate da meno di questo scarto lungo la normale contano
# come adiacenti: il dispositivo le attraversa e si allinea alla faccia piu'
# esterna del pacchetto. Non e' una tolleranza di modellazione generosa di
# proposito: un'intercapedine vera e' molto piu' larga di cosi' e non deve
# essere attraversata.
ADJACENT_GAP_MM = 2.0

# Coseno dell'angolo massimo fra le normali di due murature perche' contino
# come parallele, e quindi come parte dello stesso pacchetto: circa 2.5 gradi.
ADJACENT_PARALLEL_TOL = 0.999

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

GEOM_EPS = 1.0e-9

SLANT_ANGLE_TOL = 1.0e-6        # radianti: sotto, il muro e' verticale
SLANT_NORMAL_TOL = 0.02         # circa 1.1 gradi di inclinazione della faccia
LINK_AXIS_TOL = 1.0e-6          # scarto ammesso sull'asse Z di un collegamento

# Alzare a True se in prova la rilettura del punto di inserimento subito
# dopo RotateElement risultasse non aggiornata: la traslazione viene
# ricalcolata da quel punto, quindi leggerlo stantio sposterebbe male.
FORCE_REGEN = False

MAX_MOVED_ROWS = 300
MAX_SKIPPED_ROWS = 500
MAX_OVER_TOLERANCE_ROWS = 50

# Valori dell'enumeratore WallLocationLine.
LOC_CENTERLINE = 0
LOC_CORE_CENTERLINE = 1
LOC_FINISH_EXTERIOR = 2
LOC_FINISH_INTERIOR = 3
LOC_CORE_EXTERIOR = 4
LOC_CORE_INTERIOR = 5

# Esito del test del fronte, per la colonna omonima del resoconto. Dichiarati
# qui perche' PlannedMove li usa come valore iniziale.
F_TOWARDS = u'towards the wall'     # la retta intercetta: faccia opposta
F_AWAY = u'away from the wall'      # la retta si allontana: faccia vicina
F_PARALLEL = u'grazing the wall'    # la retta non arriva: faccia vicina
F_UNKNOWN = u'undefined'            # niente Room Calculation Point

# Vocabolario chiuso dei motivi di esclusione: dichiarati in un unico punto
# perche' lo stesso motivo non finisca scritto in due modi diversi.
R_NO_POINT = u'element has no insertion point'
R_HOSTED_WALL = u'hosted by wall {}'
R_HOSTED_OTHER = u'hosted by {} {}'
R_PINNED = u'element is pinned'
R_GROUP = u'element inside group "{}"'
R_SUBCOMPONENT = u'sub-component of a nested family'
R_DESIGN_OPTION = u'inactive design option'
R_BORROWED = u'borrowed by another user'
R_OUT_OF_Z = u'extents outside the vertical range of the walls'
R_BEYOND_END = u'past the wall end by {}'
R_NO_PROJECTION = u'projection onto the wall geometry cannot be computed'
R_CATEGORY_NOT_HANDLED = u'category not handled by the tool'
R_CATEGORY_NOT_SELECTED = u'category cleared in the dialog'
R_NO_PARTITION = u'no vertical partition within the search radius'
R_NOT_FACED = u'no partition is being faced: the nearby ones run ' \
              u'parallel to the analysis line'
R_OVER_TOLERANCE = u'beyond the tolerance: {} from the nearest face'
R_CONNECTED = u'connected to other elements ({} connectors)'
R_ANALYSIS_ERROR = u'analysis error: {}'

W_CURTAIN = u'curtain wall: zero thickness, the distance from the face is undefined'
W_SLANTED = u'slanted wall: the face is not vertical'
W_NO_CURVE = u'wall with no location line'
W_STACKED_PARENT = u'stacked wall container: its members are used instead'
L_NOT_LOADED = u'link not loaded'
L_TILTED = u'link tilted or flipped: the tool works in plan'
L_MIRRORED = u'mirrored link: face directions would be reversed'
L_NO_TRANSFORM = u'link transform cannot be read'
L_SCALED = u'scaled link: distances would not be comparable'
W_NO_BBOX = u'wall vertical extent cannot be read'
W_NO_OFFSET = u'location line offset is indeterminate'


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


def format_deg(value_rad):
    return u'{:+.1f} deg'.format(math.degrees(value_rad))


def format_mm(value_internal):
    """Lunghezza in millimetri, con il segno se negativa."""
    return u'{:.0f} mm'.format(internal_to_mm(value_internal))



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
ANGLE_TOL = math.radians(ANGLE_TOL_DEG)
WALL_END_TOL = mm_to_internal(WALL_END_TOL_MM)
Z_TOL = mm_to_internal(Z_TOL_MM)
AMBIGUITY_TOL = mm_to_internal(AMBIGUITY_TOL_MM)
PARTITION_MARGIN = mm_to_internal(PARTITION_MARGIN_MM)
FRONT_MIN_LENGTH = mm_to_internal(FRONT_MIN_LENGTH_MM)
ADJACENT_GAP = mm_to_internal(ADJACENT_GAP_MM)

# Componente minima del fronte lungo la normale del muro, cioe' il coseno
# dell'apertura ammessa. Sotto questa soglia la parete e' considerata
# parallela alla retta e non e' un riferimento valido per quel dispositivo.
FRONT_MIN_COS = math.cos(math.radians(FRONT_MAX_ANGLE_DEG))

# Rotazione massima applicabile. E' la stessa apertura del cono, e non per
# simmetria estetica: una partizione e' ammessa solo se il fronte le sta
# entro FRONT_MAX_ANGLE_DEG dalla perpendicolare, quindi l'angolo da
# recuperare per raddrizzare l'elemento NON PUO' superare quel valore. La
# guardia esiste lo stesso, perche' quell'invariante si regge su due
# funzioni diverse e un domani potrebbe rompersi in silenzio.
MAX_ROTATION = math.radians(FRONT_MAX_ANGLE_DEG)


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

    def __init__(self, wall, source):
        self.wall = wall
        self.wall_id = wall.Id
        self.source = source
        self.label = wall_label(wall)
        if source.is_linked:
            self.label = u'[{}] {}'.format(source.link_name, self.label)
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

    def __init__(self, wall_info, foot, normal_ext, normal_host,
                 signed_center, face_distance, side, beyond):
        self.wall_info = wall_info
        self.foot = foot
        self.normal_ext = normal_ext        # coordinate del documento origine
        self.normal_host = normal_host      # coordinate del documento host
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
    return u'Wall'


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
        return True, u'normal Z component: {:.3f}'.format(normal_z)
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

    warn(u'Wall {}: unrecognised "Location Line" value ({}), '
         u'centreline assumed.'.format(element_id_value(wall.Id), kind))
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


def build_wall_info(wall, source):
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

    info = WallInfo(wall, source)
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
                warn(u'Wall {}: the normal computed from the tangent does '
                     u'not match Wall.Orientation. Straight walls use '
                     u'Wall.Orientation, but on arcs the cross product '
                     u'convention needs checking.'.format(
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
            warn(u'Wall {}: unrecognised arc parametrisation, angle in '
                 u'radians assumed.'.format(
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
    warn(u'Wall {}: unhandled location line type ({}).'.format(
        element_id_value(wall_info.wall_id), type(wall_info.curve).__name__))
    return None


def test_point_against_wall(wall_info, point):
    """Distanza firmata di un punto dalla faccia del muro piu' vicina.

    Il punto deve essere gia' espresso nelle coordinate del documento in cui
    vive il muro: per un muro collegato lo converte all_wall_hits().
    """
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

    # Le grandezze scalari sono invarianti per trasformazione rigida: solo
    # la direzione va riportata nelle coordinate dell'host, dove vivono il
    # punto di inserimento e il fronte dell'elemento.
    normal_host = wall_info.source.to_host_vector(normal_ext)
    if normal_host is None:
        return None

    return WallHit(wall_info, foot, normal_ext, normal_host, signed_center,
                   face_distance, side, beyond)


# =========================================================================
# RICERCA DELLE PARTIZIONI VERTICALI
# =========================================================================
# La direzione della ricerca e' invertita rispetto alla prima versione: si
# parte dagli elementi indicati dall'utente e per ciascuno si cercano le
# partizioni verticali vicine, invece di partire dai muri e raccogliere gli
# elementi attorno. Cambia solo la fase larga: la matematica di
# test_point_against_wall() resta identica.

# Cache dei WallInfo, indicizzata per (sorgente, valore di ElementId). La
# sorgente entra nella chiave perche' documenti diversi possono contenere id
# uguali. Non e' una ottimizzazione accessoria: build_wall_info() puo'
# estrarre la geometria della partizione per il controllo di inclinazione, e
# la stessa partizione ricorre per tutti gli elementi che le stanno davanti.
PARTITION_CACHE = {}


class PartitionSource(object):
    """Documento in cui cercare le partizioni, con la sua trasformazione.

    Per il documento host la trasformazione e' l'identita' e non viene
    applicata. Per un modello collegato porta dalle coordinate del link a
    quelle dell'host.
    """

    def __init__(self, key, document, transform, link_name):
        self.key = key
        self.document = document
        self.transform = transform
        self.inverse = transform.Inverse if transform is not None else None
        self.link_name = link_name
        self.z_offset = transform.Origin.Z if transform is not None else 0.0
        self.ids = None

    @property
    def is_linked(self):
        return self.link_name is not None

    def to_local(self, point):
        """Porta un punto dalle coordinate host a quelle della sorgente."""
        if self.inverse is None:
            return point
        try:
            return self.inverse.OfPoint(point)
        except Exception:
            return None

    def to_host_vector(self, vector):
        """Porta una direzione dalle coordinate della sorgente all'host."""
        if self.transform is None:
            return vector
        try:
            return self.transform.OfVector(vector)
        except Exception:
            return None


def build_category_filter(built_in_categories):
    """Filtro multicategoria con le categorie indicate."""
    category_list = List[DB.BuiltInCategory]()
    for built_in_category in built_in_categories:
        category_list.Add(built_in_category)
    return DB.ElementMulticategoryFilter(category_list)


def link_display_name(link):
    """Nome leggibile di un'istanza di collegamento."""
    for getter in (lambda: DB.Element.Name.GetValue(link),
                   lambda: DB.Element.Name.GetValue(
                       doc.GetElement(link.GetTypeId()))):
        try:
            name = getter()
            if name:
                return name
        except Exception:
            continue
    return u'link {}'.format(element_id_value(link.Id))


def link_transform_problem(transform):
    """Motivo per cui un collegamento non e' utilizzabile, None se va bene.

    Lo strumento lavora in pianta e non tocca mai la quota. Se il
    collegamento fosse inclinato, la normale di un muro riportata nell'host
    avrebbe una componente verticale e lo spostamento cambierebbe la quota.
    Se fosse speculare, i versi esterno e interno delle facce risulterebbero
    invertiti e i dispositivi finirebbero sulla faccia sbagliata.
    """
    try:
        basis_z = transform.BasisZ
    except Exception:
        return L_NO_TRANSFORM
    if (abs(basis_z.X) > LINK_AXIS_TOL
            or abs(basis_z.Y) > LINK_AXIS_TOL
            or basis_z.Z <= 0.0):
        return L_TILTED
    try:
        if transform.HasReflection:
            return L_MIRRORED
    except Exception:
        pass
    # Con una scala diversa da 1 le distanze misurate nel documento
    # collegato non sarebbero confrontabili con la tolleranza dell'utente,
    # che e' espressa in coordinate host.
    try:
        if abs(transform.Scale - 1.0) > 1.0e-9:
            return L_SCALED
    except Exception:
        pass
    return None


def partition_info(source, partition):
    """(WallInfo, motivo di scarto) di una partizione, con cache."""
    key = (source.key, element_id_value(partition.Id))
    cached = PARTITION_CACHE.get(key)
    if cached is not None:
        return cached
    try:
        info, reason = build_wall_info(partition, source)
    except Exception as error:
        info = None
        reason = R_ANALYSIS_ERROR.format(u'{}'.format(error)[:120])
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


def collect_ids_in(document, points, margin, label):
    """Id delle partizioni di un documento nella regione data."""
    empty = List[DB.ElementId]()
    outline = points_outline(points, margin)
    if outline is None:
        return empty
    try:
        collector = DB.FilteredElementCollector(document)\
            .WherePasses(build_category_filter(PARTITION_CATEGORIES))\
            .WhereElementIsNotElementType()\
            .WherePasses(DB.BoundingBoxIntersectsFilter(outline))
        return List[DB.ElementId](collector.ToElementIds())
    except Exception as error:
        warn(u'Collecting partitions in {} failed: {}'.format(
            label, error))
        return empty


def collect_partition_sources(points, tolerance_internal, include_links):
    """Documento host piu' i modelli collegati utilizzabili.

    Una sola passata per documento sulla regione occupata dalla selezione.
    La restrizione per singolo elemento avviene poi su questi insiemi
    ridotti e non sull'intero modello.
    """
    if not PARTITION_CATEGORIES:
        warn(u'No vertical partition category is available in this '
             u'version of Revit.')
        return []

    margin = tolerance_internal + PARTITION_MARGIN
    sources = []

    host = PartitionSource(0, doc, None, None)
    host.ids = collect_ids_in(doc, points, margin, u'this model')
    sources.append(host)

    if not include_links:
        return sources

    try:
        links = list(DB.FilteredElementCollector(doc)
                     .OfClass(DB.RevitLinkInstance)
                     .ToElements())
    except Exception as error:
        warn(u'The list of linked models cannot be read: {}'.format(error))
        return sources

    for link in links:
        name = link_display_name(link)

        link_doc = None
        try:
            link_doc = link.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            warn(u'Link "{}" ignored: {}.'.format(name, L_NOT_LOADED))
            continue

        transform = None
        try:
            transform = link.GetTotalTransform()
        except Exception:
            transform = None
        if transform is None:
            warn(u'Link "{}" ignored: {}.'.format(
                name, L_NO_TRANSFORM))
            continue

        problem = link_transform_problem(transform)
        if problem is not None:
            warn(u'Link "{}" ignored: {}.'.format(name, problem))
            continue

        source = PartitionSource(element_id_value(link.Id), link_doc,
                                 transform, name)
        local_points = []
        for point in points:
            local = source.to_local(point)
            if local is not None:
                local_points.append(local)
        source.ids = collect_ids_in(link_doc, local_points, margin,
                                    u'"{}"'.format(name))
        if source.ids.Count:
            sources.append(source)

    return sources


def partitions_near_point(point, tolerance_internal, sources):
    """Partizioni utilizzabili vicine a un punto, host e collegamenti.

    Ritorna (lista di WallInfo, lista di (sorgente, partizione, motivo)).
    """
    infos = []
    problems = []

    for source in sources:
        if source.ids is None or source.ids.Count == 0:
            continue

        local_point = source.to_local(point)
        if local_point is None:
            continue

        try:
            near = DB.FilteredElementCollector(source.document, source.ids)\
                .WherePasses(DB.BoundingBoxIntersectsFilter(
                    point_outline(local_point, tolerance_internal)))\
                .ToElements()
        except Exception as error:
            warn(u'Proximity filter failed: {}'.format(error))
            continue

        for partition in near:
            info, reason = partition_info(source, partition)
            if info is None:
                problems.append((source, partition, reason))
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


def z_contains(element_extent, wall_info):
    """True se la partizione copre l'ingombro verticale dell'elemento.

    Il filtro nativo verifica l'INTERSEZIONE fra bounding box, mentre il
    requisito e' il CONTENIMENTO: serve quindi questo test esplicito a valle.
    """
    if element_extent is None:
        return False
    low, high = element_extent
    # z_min e z_max sono nelle coordinate del documento della partizione;
    # l'ingombro dell'elemento e' in coordinate host. Con l'asse Z del
    # collegamento verticale, garantito da link_transform_problem(), la
    # conversione e' una sola traslazione.
    offset = wall_info.source.z_offset
    return (low >= wall_info.z_min + offset - Z_TOL
            and high <= wall_info.z_max + offset + Z_TOL)


def partitions_at_height(element_extent, wall_infos):
    """Sottoinsieme delle partizioni che coprono la quota dell'elemento.

    E' un FILTRO SUI MURI, non una guardia sull'elemento. Un muro che alla
    quota del dispositivo non esiste non e' un bersaglio valido e va tolto
    dalla gara: se restasse, un dispositivo vicino a un muro alto e a un
    muretto basso potrebbe essere allineato al muretto, che alla sua quota
    non c'e'.
    """
    return [w for w in wall_infos if z_contains(element_extent, w)]


# =========================================================================
# COSTRUZIONE DEL PIANO (nessuna transazione)
# =========================================================================

class AlignOptions(object):
    """Scelte effettuate dall'utente nella finestra di dialogo."""

    def __init__(self, categories, tolerance_cm,
                 skip_connected, include_links, dry_run):
        self.categories = categories
        self.tolerance_cm = tolerance_cm
        self.tolerance_internal = cm_to_internal(tolerance_cm)
        # Lunghezza della retta di analisi del fronte: la tolleranza e'
        # l'unica misura che l'utente ha dichiarato, quindi la scala di
        # ricerca del fronte si aggancia a quella.
        self.front_ray_length = cm_to_internal(tolerance_cm) * FRONT_RAY_FACTOR
        self.skip_connected = skip_connected
        self.include_links = include_links
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
        self.wall_is_linked = False
        self.face_side = u'-'

        # Esito del test del fronte e faccia effettivamente scelta: senza
        # queste due colonne un dispositivo portato dall'altra parte del
        # muro sembrerebbe uno spostamento sbagliato invece di una
        # correzione voluta.
        self.front_outcome = F_UNKNOWN
        self.front_flipped = False

        # Corsa lungo la retta di analisi invece che lungo la normale, e
        # numero di partizioni attraversate oltre la prima: sono i due
        # motivi per cui uno spostamento puo' risultare molto piu' lungo
        # della distanza misurata.
        self.along_front = False
        self.crossed_partitions = 0

        # Raddrizzamento fine attorno all'asse verticale. Non e' un
        # riorientamento: per costruzione non supera FRONT_MAX_ANGLE_DEG.
        self.rotation_rad = 0.0
        self.rotation_over_limit = False
        self.needs_rotation = False
        self.applied_rotation = False

        self.point_before = None
        self.distance_before = 0.0

        # Punto bersaglio ASSOLUTO, non vettore: la traslazione viene
        # ricalcolata dal punto corrente al momento di applicarla, cosi' un
        # elemento mosso nel frattempo finisce comunque dove deve.
        self.target_point = None
        self.distance_after = 0.0
        self.translation_length = 0.0

        self.needs_move = False
        self.status = PlannedMove.PLANNED

        self.connected = 0
        self.competing_walls = 1
        self.second_distance = None
        self.note = None

        self.applied_move = False
        self.error = None


class SkippedElement(object):
    """Un elemento escluso, con il motivo specifico e la distanza misurata."""

    def __init__(self, element, category_name, category_key, reason,
                 wall_id=None, distance=None, wall_label=u'-',
                 wall_is_linked=False):
        self.element = element
        self.element_id = element.Id if element is not None else None
        self.category_name = category_name
        self.category_key = category_key
        self.reason = reason
        self.wall_id = wall_id
        self.wall_label = wall_label
        self.wall_is_linked = wall_is_linked
        self.distance = distance


class OverTolerance(object):
    """Un elemento vicino ma oltre la soglia."""

    def __init__(self, element, category_name, category_key, wall_id,
                 wall_label, distance, excess, wall_is_linked=False):
        self.element = element
        self.element_id = element.Id
        self.category_name = category_name
        self.category_key = category_key
        self.wall_id = wall_id
        self.wall_label = wall_label
        self.wall_is_linked = wall_is_linked
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


def room_calculation_point(element):
    """Room Calculation Point dell'istanza, in coordinate del modello.

    None quando la famiglia non lo espone. E' il punto che Revit usa per
    stabilire in quale locale sta il dispositivo, quindi per una famiglia
    autorata con criterio cade DAVANTI all'apparecchio, dentro la stanza:
    e' il dato che distingue il fronte dal retro senza dipendere da come e'
    orientato il sistema di riferimento della famiglia.

    GetSpatialElementCalculationPoint() solleva InvalidOperationException
    quando il punto non c'e', quindi la property va interrogata prima. Il
    metodo esiste dalla versione 2016 ed e' disponibile su tutte le versioni
    coperte dallo strumento.
    """
    try:
        if not element.HasSpatialElementCalculationPoint:
            return None
    except Exception:
        return None
    try:
        return element.GetSpatialElementCalculationPoint()
    except Exception:
        return None


def front_direction(element, point):
    """Direzione del fronte in pianta, ricavata dal Room Calculation Point.

    Il punto di calcolo viene proiettato sul piano orizzontale passante per
    il punto di inserimento, e la direzione e' quella che va dal secondo al
    primo. Proiettare invece di usare il vettore nello spazio serve perche'
    su un dispositivo a parete il punto di calcolo e' quasi sempre anche
    piu' in alto o piu' in basso dell'origine, e la componente verticale
    falserebbe l'angolo rispetto alla normale del muro.

    Ritorna None quando il punto di calcolo manca, quando l'autore della
    famiglia non lo ha spostato dall'origine e quando gli sta esattamente
    sopra: in tutti e tre i casi non c'e' nessuna direzione da leggere, e
    la faccia viene scelta con il criterio posizionale di sempre.
    """
    calculation_point = room_calculation_point(element)
    if calculation_point is None:
        return None
    planar = DB.XYZ(calculation_point.X - point.X,
                    calculation_point.Y - point.Y,
                    0.0)
    if planar.GetLength() < FRONT_MIN_LENGTH:
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


def face_side_from_front(hit, front, ray_length):
    """Su quale delle due facce del muro va portato il dispositivo.

    Il criterio posizionale da solo sbaglia ogni volta che il dispositivo e'
    modellato dalla parte sbagliata della partizione, o dentro il suo
    spessore: porta il punto di inserimento sulla faccia piu' vicina, che
    puo' essere quella alle spalle dell'apparecchio. Il fronte ricavato dal
    Room Calculation Point risolve il caso, perche' dice da che parte
    guarda il dispositivo.

    Il test e' quello chiesto: dal punto di inserimento si traccia una retta
    lunga ray_length lungo il fronte e si guarda se intercetta la superficie
    del muro. Non serve la geometria, perche' sotto una trasformazione
    rigida tutte le grandezze in gioco sono scalari invarianti: lo
    scostamento firmato dalla mezzeria, il semispessore e la componente del
    fronte lungo la normale. Lavorando in coordinate host anche il prodotto
    scalare fra fronte e normale e' invariante, quindi le partizioni
    collegate non richiedono nessuna conversione aggiuntiva.

    Lungo la normale il muro occupa l'intervallo [-half_width, +half_width]
    e il punto sta a signed_center. Se la retta lo intercetta, il
    dispositivo va sulla faccia dalla parte verso cui guarda, che e' la
    SECONDA faccia quando il fronte punta contro il muro. Se non lo
    intercetta, perche' se ne allontana o perche' e' troppo radente per
    arrivarci entro ray_length, resta il criterio posizionale.

    Ritorna (side, esito del test, componente del fronte lungo la normale).
    L'ultimo valore serve a chi deve poi muovere l'elemento LUNGO la retta:
    e' il coseno fra la retta e la normale, cioe' il fattore che lega la
    distanza da coprire alla corsa da percorrere.
    """
    if front is None:
        return hit.side, F_UNKNOWN, 0.0

    along_normal = front.DotProduct(hit.normal_host)
    if abs(along_normal) < FRONT_MIN_COS:
        # Fronte troppo radente: questo muro non e' cio' che il dispositivo
        # guarda, quindi non ha titolo per sceglierne la faccia.
        return hit.side, F_PARALLEL, along_normal

    signed_center = hit.signed_center
    half_width = hit.wall_info.half_width
    front_side = 1.0 if along_normal > 0.0 else -1.0

    if abs(signed_center) <= half_width:
        # Punto dentro lo spessore: la retta esce comunque, dalla faccia
        # verso cui guarda il dispositivo.
        crossing = (front_side * half_width - signed_center) / along_normal
    elif front_side * signed_center > 0.0:
        # La retta si allontana dal muro: nessuna intersezione possibile.
        return hit.side, F_AWAY, along_normal
    else:
        # La retta punta contro il muro: entra dalla faccia vicina.
        near_face = half_width if signed_center > 0.0 else -half_width
        crossing = (near_face - signed_center) / along_normal

    if crossing < 0.0 or crossing > ray_length:
        # La retta non arriva alla faccia entro la lunghezza di analisi.
        # Con il gate sul coseno qui sopra questo puo' accadere solo a un
        # elemento sepolto dentro un muro piu' spesso della tolleranza.
        return hit.side, F_PARALLEL, along_normal

    return front_side, F_TOWARDS, along_normal


def signed_angle_about_z(vector_from, vector_to):
    """Angolo firmato attorno a +Z, positivo antiorario visto dall'alto."""
    cross_z = vector_from.X * vector_to.Y - vector_from.Y * vector_to.X
    dot = vector_from.X * vector_to.X + vector_from.Y * vector_to.Y
    return math.atan2(cross_z, dot)


def faces_the_wall(hit, front):
    """True se il dispositivo sta guardando questa partizione.

    Guarda solo l'orientamento, non il verso: una parete davanti e una
    dietro sono entrambe cose che il dispositivo "guarda", perche' la retta
    di analisi e' una retta e non una semiretta. Cio' che esclude e' la
    parete di FIANCO, quella rispetto a cui la retta e' radente.
    """
    if front is None:
        return False
    return abs(front.DotProduct(hit.normal_host)) >= FRONT_MIN_COS


def same_partition(first, second):
    """True se due esiti riguardano la stessa partizione.

    L'id da solo non basta: documenti diversi possono contenere id uguali,
    quindi la sorgente entra nel confronto come entra nella chiave della
    cache.
    """
    return (first.wall_info.source.key == second.wall_info.source.key
            and element_id_value(first.wall_info.wall_id)
            == element_id_value(second.wall_info.wall_id))


def span_along(hit, normal):
    """Intervallo occupato da una partizione lungo `normal`, relativo al punto.

    L'asse ha origine nel punto di inserimento ed e' orientato come la
    normale del muro di riferimento. Ritorna None se la partizione non e'
    parallela a quella di riferimento, perche' allora non fa parte dello
    stesso pacchetto murario e un intervallo su questo asse non
    significherebbe nulla.

    La mezzeria della partizione sta a -signed_center lungo la PROPRIA
    normale. Se questa e' opposta a quella di riferimento il segno si
    ribalta, ed e' il solo motivo per cui serve il prodotto scalare.
    """
    alignment = hit.normal_host.DotProduct(normal)
    if abs(alignment) < ADJACENT_PARALLEL_TOL:
        return None
    center = -hit.signed_center if alignment > 0.0 else hit.signed_center
    half_width = hit.wall_info.half_width
    return center - half_width, center + half_width


def outermost_face(best, face_side, hits):
    """Faccia piu' esterna del pacchetto di murature adiacenti.

    Una parete modellata come piu' istanze a contatto - il caso tipico e' la
    muratura piu' il suo rivestimento, oppure due tramezzi accostati - va
    trattata come un pacchetto unico: il dispositivo si allinea alla faccia
    piu' esterna, non alla prima che incontra. Senza questo, un apparecchio
    modellato dentro il muro strutturale finirebbe a filo del suo
    rivestimento interno, cioe' dentro la stratigrafia.

    Si lavora sull'asse orientato come la normale del muro vincente, con
    l'origine nel punto di inserimento. Partendo dalla faccia scelta si
    cerca una partizione parallela che la contenga o la sfiori e che si
    estenda oltre; quando la si trova il confine avanza al suo bordo
    esterno, e la ricerca riparte da li'. Il ciclo termina da solo perche'
    ogni passo consuma una partizione.

    Il muro vincente resta quello scelto per distanza e la sua faccia resta
    quella scelta dal fronte: qui il confine puo' solo essere spinto piu' in
    fuori, mai tirato indietro e mai girato dall'altra parte.

    Ritorna (coordinata della faccia sull'asse, partizioni attraversate).
    """
    normal = best.normal_host
    boundary = face_side * best.wall_info.half_width - best.signed_center

    spans = []
    for hit in hits:
        if same_partition(hit, best):
            continue
        span = span_along(hit, normal)
        if span is not None:
            spans.append((span, hit))

    crossed = []
    while True:
        winner = None
        for (low, high), hit in spans:
            if any(hit is done for done in crossed):
                continue
            if face_side > 0.0:
                # Deve iniziare entro il confine (o sfiorarlo) e finire oltre.
                if low <= boundary + ADJACENT_GAP \
                        and high > boundary + GEOM_EPS \
                        and (winner is None or high > winner[0]):
                    winner = (high, hit)
            else:
                if high >= boundary - ADJACENT_GAP \
                        and low < boundary - GEOM_EPS \
                        and (winner is None or low < winner[0]):
                    winner = (low, hit)
        if winner is None:
            return boundary, crossed
        boundary = winner[0]
        crossed.append(winner[1])


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
        local_point = wall_info.source.to_local(point)
        if local_point is None:
            continue
        hit = test_point_against_wall(wall_info, local_point)
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
    if nearest is None:
        return SkippedElement(element, category_name, category_key, reason)
    return SkippedElement(
        element, category_name, category_key, reason,
        nearest.wall_info.wall_id,
        nearest.face_distance,
        nearest.wall_info.label,
        nearest.wall_info.source.is_linked)


class PlanContext(object):
    """Stato condiviso dalla fase di analisi, calcolato una volta sola."""

    def __init__(self, sources, active_design_option_id):
        self.sources = sources
        self.active_design_option_id = active_design_option_id
        # Indicizzati per (sorgente, id): la stessa partizione scartata
        # ricorre per molti elementi e nel resoconto deve comparire una
        # volta sola. La sorgente entra nella chiave perche' documenti
        # diversi possono contenere id uguali.
        self.partition_problems = {}

    def note_problem(self, source, partition, reason):
        key = (source.key, element_id_value(partition.Id))
        self.partition_problems[key] = (source, partition, reason)


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
        point, options.tolerance_internal, context.sources)
    for problem_source, partition, problem_reason in problems:
        context.note_problem(problem_source, partition, problem_reason)

    # Un muro che non copre la quota dell'elemento non e' un bersaglio
    # valido: si toglie dalla gara PRIMA di calcolare le proiezioni.
    element_extent = z_extent(element)
    at_height = partitions_at_height(element_extent, wall_infos)

    # Le proiezioni si calcolano subito, prima delle guardie, cosi' ogni riga
    # del resoconto porta la distanza misurata anche quando il motivo di
    # scarto non c'entra con la distanza: un elemento bloccato a 62 mm vale
    # la pena di sbloccarlo, uno a 290 mm probabilmente no.
    hits = all_wall_hits(at_height, point)
    context_hits = hits if hits else all_wall_hits(wall_infos, point)

    reason = host_skip_reason(element)
    if reason is not None:
        return skipped_with_context(element, category_name, category_key,
                                    reason, context_hits)

    reason = edit_skip_reason(element, context.active_design_option_id)
    if reason is not None:
        return skipped_with_context(element, category_name, category_key,
                                    reason, context_hits)

    if not wall_infos:
        return SkippedElement(element, category_name, category_key,
                              R_NO_PARTITION)

    if not at_height:
        return skipped_with_context(element, category_name, category_key,
                                    R_OUT_OF_Z, context_hits)

    connected = connected_connectors(element)
    if options.skip_connected and connected > 0:
        return skipped_with_context(element, category_name, category_key,
                                    R_CONNECTED.format(connected),
                                    context_hits)

    # Il fronte serve a scegliere QUALE delle due facce del muro usare, non
    # a decidere se l'elemento sia trattabile: quando manca si ricade sul
    # criterio posizionale e l'elemento viene comunque allineato.
    front = front_direction(element, point)

    if not hits:
        # Partizioni trovate, ma nessuna ha prodotto una proiezione
        # utilizzabile. Va riportato, non scartato in silenzio.
        return SkippedElement(element, category_name, category_key,
                              R_NO_PROJECTION)

    # Le partizioni su cui l'elemento si proietta davvero. Solo queste
    # possono far parte del suo pacchetto murario: una parete che finisce
    # prima di arrivargli davanti non gli sta adiacente.
    in_range = [h for h in hits if h.beyond <= WALL_END_TOL]

    if not in_range:
        nearest = hits[0]
        return skipped_with_context(
            element, category_name, category_key,
            R_BEYOND_END.format(format_mm(nearest.beyond)), hits)

    # Una partizione parallela alla retta di analisi e' ESCLUSA, non
    # sfavorita. Non e' il muro che il dispositivo guarda, e allinearcelo
    # significherebbe spostarlo lungo una normale ortogonale al suo asse di
    # vista: esattamente lo spostamento di traverso che la retta serve a
    # evitare. Se non ne resta nessuna l'elemento viene SALTATO e riportato,
    # non allineato a un riferimento sbagliato.
    #
    # Senza fronte leggibile non esiste nessuna retta, quindi non c'e' nulla
    # a cui una parete possa essere parallela: li' concorrono tutte, con il
    # criterio posizionale delle versioni precedenti.
    candidates = in_range
    if front is not None:
        candidates = [h for h in in_range if faces_the_wall(h, front)]
        if not candidates:
            return skipped_with_context(element, category_name, category_key,
                                        R_NOT_FACED, hits)

    qualifying = [h for h in candidates
                  if h.face_distance <= options.tolerance_internal]

    if not qualifying:
        # Oltre la tolleranza. Nella prima versione era informativo, perche'
        # era lo strumento a pescare gli elementi. Ora l'elemento lo ha
        # indicato l'utente, quindi e' uno SCARTO che deve vedere. La
        # distanza riportata e' quella della partizione fronteggiata piu'
        # vicina, non di una che il dispositivo non guarda: sarebbe un
        # suggerimento fuorviante su quanto alzare la tolleranza.
        nearest = candidates[0]
        return OverTolerance(
            element, category_name, category_key,
            nearest.wall_info.wall_id,
            nearest.wall_info.label,
            nearest.face_distance,
            nearest.face_distance - options.tolerance_internal,
            nearest.wall_info.source.is_linked)

    best = qualifying[0]
    second = qualifying[1] if len(qualifying) > 1 else None
    wall_info = best.wall_info

    # Quale delle due facce. Il muro e' scelto per distanza fra quelli che
    # il dispositivo guarda: qui si decide soltanto da che parte della
    # partizione portarlo, e la retta di analisi puo' mandarlo sulla faccia
    # opposta a quella piu' vicina.
    face_side, front_outcome, along_normal = face_side_from_front(
        best, front, options.front_ray_length)

    # Piu' istanze di muro a contatto sono un pacchetto unico: il bersaglio
    # e' la faccia piu' esterna, non quella della prima partizione.
    delta, crossed = outermost_face(best, face_side, in_range)

    # La corsa avviene LUNGO LA RETTA DI ANALISI, non lungo la normale del
    # muro: l'elemento scivola sul proprio asse di vista finche' il punto di
    # inserimento non raggiunge il piano della faccia bersaglio. Muovendosi
    # in obliquo deve percorrere 1/cos in piu' della distanza da coprire, ed
    # e' il motivo per cui l'apertura ammessa e' stretta (FRONT_MIN_COS).
    #
    # Senza un fronte leggibile non c'e' nessuna retta, e resta la normale:
    # e' il comportamento delle versioni precedenti.
    use_front = front is not None and front_outcome in (F_TOWARDS, F_AWAY)
    if use_front:
        travel = delta / along_normal
        direction = front
    else:
        travel = delta
        direction = best.normal_host

    # La quota resta invariata per costruzione: sia il fronte sia la normale
    # sono orizzontali, e la componente Z non viene comunque usata.
    target_point = DB.XYZ(point.X + direction.X * travel,
                          point.Y + direction.Y * travel,
                          point.Z)

    record = PlannedMove()
    record.element = element
    record.element_id = element.Id
    record.category_name = category_name
    record.category_key = category_key
    record.type_name = type_name_of(element)

    record.wall_id = wall_info.wall_id
    record.wall_label = wall_info.label
    record.wall_is_linked = wall_info.source.is_linked
    record.face_side = u'exterior' if face_side > 0 else u'interior'
    record.front_outcome = front_outcome
    record.front_flipped = face_side != best.side

    record.point_before = point
    record.distance_before = best.face_distance
    record.target_point = target_point
    # Distanza dalla faccia del muro in colonna, misurata dove l'elemento
    # va a finire. E' zero nel caso normale, ma non quando il pacchetto
    # murario ha spinto il bersaglio oltre la prima partizione: li' vale lo
    # spessore di quelle attraversate, e scriverci zero sarebbe una bugia.
    record.distance_after = abs(delta + best.signed_center) \
        - wall_info.half_width
    record.translation_length = abs(travel)
    record.along_front = use_front
    record.crossed_partitions = len(crossed)
    record.connected = connected

    # Raddrizzamento: il fronte viene portato esattamente perpendicolare
    # alla faccia bersaglio. E' una correzione fine, non un riorientamento,
    # perche' quella partizione e' stata ammessa solo se il fronte le stava
    # gia' entro FRONT_MAX_ANGLE_DEG dalla perpendicolare: l'angolo da
    # recuperare non puo' superare quel valore.
    #
    # Si ruota il fronte, non FacingOrientation: la rotazione e' rigida e
    # porta con se' anche il Room Calculation Point, quindi dopo il comando
    # il fronte e' perpendicolare al muro qualunque sia la convenzione con
    # cui la famiglia e' stata autorata. Senza fronte leggibile non c'e'
    # nessun verso di cui fidarsi, e l'elemento non viene ruotato.
    if use_front:
        target_direction = best.normal_host.Multiply(face_side)
        angle = signed_angle_about_z(front, target_direction)
        if abs(angle) <= MAX_ROTATION:
            record.rotation_rad = angle
        else:
            # Non puo' accadere finche' faces_the_wall filtra i candidati,
            # ma l'invariante si regge su due funzioni diverse: se un giorno
            # si rompe, deve comparire nel resoconto e non nel modello.
            record.rotation_over_limit = True
            record.rotation_rad = 0.0

    record.needs_move = record.translation_length > POSITION_TOL
    record.needs_rotation = abs(record.rotation_rad) > ANGLE_TOL
    if not record.needs_move and not record.needs_rotation:
        record.status = PlannedMove.ALREADY_OK

    notes = []
    if record.rotation_over_limit:
        notes.append(u'rotation beyond the {:.0f} degree limit: '
                     u'not applied'.format(FRONT_MAX_ANGLE_DEG))

    # Uno spostamento che attraversa il muro e' molto piu' lungo della
    # distanza misurata: senza questa riga sembrerebbe un errore.
    if record.front_flipped:
        notes.append(u'front faces the wall: moved to the opposite face')

    # Un muro piu' vicino escluso perche' parallelo alla retta va
    # dichiarato: e' la domanda che l'utente si fa per prima guardando il
    # risultato. Va detto solo quando quel muro sarebbe stato davvero un
    # contendente, cioe' quando rientrava nella tolleranza.
    closest = in_range[0]
    if closest is not best \
            and closest.face_distance <= options.tolerance_internal \
            and not faces_the_wall(closest, front):
        notes.append(
            u'nearer wall excluded, parallel to the line: {} at {}'
            .format(closest.wall_info.label,
                    format_mm(closest.face_distance)))

    # Stesso discorso per il pacchetto murario, che allunga la corsa di
    # tutto lo spessore delle partizioni attraversate.
    if crossed:
        notes.append(u'outer face of {} adjacent walls (last one: {})'
                     .format(len(crossed) + 1, crossed[-1].wall_info.label))

    # Due muri praticamente equidistanti sono il caso che l'utente
    # controllera' per primo quando qualcosa sembrera' sbagliato, quindi va
    # dichiarato invece di lasciarlo scoprire.
    if second is not None:
        record.competing_walls = len(qualifying)
        record.second_distance = second.face_distance
        gap = abs(second.face_distance) - abs(best.face_distance)
        if gap <= AMBIGUITY_TOL:
            notes.append(u'ambiguous: second wall at {}'.format(
                format_mm(second.face_distance)))

    record.note = u'; '.join(notes) if notes else None

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

            record = (description, ids, u'error' if is_error else u'warning')
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
    """Applica rotazione e spostamento gia' calcolati.

    Prima la rotazione, attorno a un asse verticale per il punto di
    inserimento originale, che e' noto con certezza. Poi la traslazione,
    ricalcolata dal punto CORRENTE verso il bersaglio assoluto memorizzato
    in analisi, invece di riusare il vettore calcolato allora: non e'
    garantito che RotateElement lasci il punto di inserimento esattamente
    invariato per ogni tipo di famiglia, e ricalcolando dopo la rotazione
    qualunque deriva si autocorregge. La componente Z e' sempre zero.
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
    with forms.WarningBar(title='Select the MEP elements to align, '
                                'then press Finish'):
        try:
            references = uidoc.Selection.PickObjects(
                UI.Selection.ObjectType.Element,
                MEPSelectionFilter(),
                'Select the MEP elements to align')
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
        return elements, u'from the current selection'

    elements = pick_elements()
    if not elements:
        script.exit()
    return elements, u'picked graphically'


class MEPAlignToWallWindow(forms.WPFWindow):
    """Finestra di dialogo definita in MEPAlignToWallWindow.xaml."""

    # Attributo di classe: gli handler delle CheckBox scattano durante il
    # popolamento iniziale, prima che __init__ abbia finito.
    _ready = False

    def __init__(self, xaml_file, elements, source_label, category_info,
                 out_of_scope):
        forms.WPFWindow.__init__(self, xaml_file)

        self.options = None
        self._checks = []
        self._category_info = category_info

        self._fit_to_screen()
        self._setup_selection(elements, source_label, out_of_scope)
        self._setup_categories(category_info)

        self._ready = True
        self._refresh_count()

    # --- popolamento ---------------------------------------------------

    def _fit_to_screen(self):
        """Impedisce alla finestra di superare l'area di lavoro.

        Lo XAML usa SizeToContent="Height": senza un limite la finestra
        cresce quanto il contenuto e su schermi piccoli, o con scalatura di
        Windows elevata, esce dal monitor portandosi fuori i pulsanti. Il
        limite si legge a runtime perche' dipende dal monitor e dalla
        scalatura, che in XAML non sono note.
        """
        try:
            available = SystemParameters.WorkArea.Height - 60
        except Exception:
            return
        if available > 200:
            self.MaxHeight = available


    def _setup_selection(self, elements, source_label, out_of_scope):
        self.tb_walls_info.Text = u'{} elements selected ({}).'.format(
            len(elements), source_label)

        detail = (u'For each of them the tool looks for the nearest '
                  u'vertical partition. Elements beyond the tolerance are '
                  u'skipped and reported.')
        if out_of_scope:
            detail += (u' {} selected elements do not belong to the '
                       u'handled categories and will be skipped.'.format(
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
        self.tb_cat_count.Text = u'{} of {} categories, {} candidate ' \
                                 u'elements'.format(len(selected),
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
            return None, u'Enter the maximum distance in centimetres.'
        try:
            value = float(raw)
        except ValueError:
            return None, u'The maximum distance is not a valid number.'
        if value <= 0.0:
            return None, u'The maximum distance must be greater than zero.'
        if value > MAX_TOLERANCE_CM:
            return None, u'The maximum distance looks out of scale ' \
                         u'(over {:.0f} cm).'.format(MAX_TOLERANCE_CM)
        return value, None

    def on_run(self, sender, args):
        tolerance_cm, error = self._parse_tolerance()
        if error:
            forms.alert(error, title=u'Invalid value')
            return

        categories = [c.Tag for c in self._checks if c.IsChecked]
        if not categories:
            forms.alert(u'Select at least one category to align.',
                        title=u'No category')
            return

        self.options = AlignOptions(
            categories,
            tolerance_cm,
            bool(self.chk_skip_connected.IsChecked),
            bool(self.chk_links.IsChecked),
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
            u'Layout file not found:\n\n{}\n\nIt must sit in the same '
            u'folder as the script.'.format(path),
            title=u'Missing layout', exitscript=True)
    return path


# =========================================================================
# RESOCONTO
# =========================================================================

def print_header(options, elements, source_label, result):
    if options.dry_run:
        output.print_md(u'# MEP alignment to partitions - simulation')
        output.print_md(
            u'**No change has been applied to the model.** '
            u'The figures below are the result that would be produced.')
    else:
        output.print_md(u'# MEP alignment to partitions - report')

    lines = [
        u'- Elements selected: **{}** ({})'.format(
            len(elements), source_label),
        u'- Tolerance: **{:.0f} cm** from the partition face'.format(
            options.tolerance_cm),
        u'- Beyond the tolerance the element is skipped, not moved',
        u'- Final position: insertion point on the outermost face of '
        u'the wall package',
        u'- Face chosen from the family Room Calculation Point, '
        u'analysis line **{:.0f} cm** long'.format(
            options.tolerance_cm * FRONT_RAY_FACTOR),
        u'- Travel along the analysis line, not along the wall normal',
        u'- Squaring up: the front is brought perpendicular to the face, '
        u'by at most **{:.0f} degrees**'.format(FRONT_MAX_ANGLE_DEG),
        u'- Elements with connected connectors: {}'.format(
            u'skipped' if options.skip_connected else u'processed'),
        u'- Linked models: {}'.format(
            u'included in the search' if options.include_links
            else u'excluded from the search'),
        u'- Categories processed: **{}** of {}'.format(
            len(options.categories), len(MEP_CATEGORIES)),
        u'- Elevation is never changed',
    ]
    if result.wall_problems:
        lines.append(
            u'- Partitions discarded: **{}** (table at the end)'.format(
                len(result.wall_problems)))

    # Due numeri che dicono quanto ha pesato il fronte su questo lotto. Il
    # primo e' il motivo per cui certi elementi si spostano molto piu' della
    # tolleranza; il secondo dice su quanti elementi il criterio non ha
    # potuto esprimersi, ed e' la prima cosa da guardare se il risultato non
    # convince.
    flipped = sum(1 for r in result.planned if r.front_flipped)
    no_front = sum(1 for r in result.planned
                   if r.front_outcome == F_UNKNOWN)
    if flipped:
        lines.append(
            u'- Moved to the opposite face because the front was facing '
            u'the partition: **{}**'.format(flipped))
    if no_front:
        lines.append(
            u'- Without a Room Calculation Point, face chosen by position '
            u'and travel along the normal: **{}**'.format(no_front))

    packaged = sum(1 for r in result.planned if r.crossed_partitions)
    if packaged:
        lines.append(
            u'- Aligned to the outer face of a package of several '
            u'adjacent walls: **{}**'.format(packaged))

    output.print_md(u'\n'.join(lines))


def print_category_summary(result, options):
    output.print_md(u'## Summary by category')
    moved_header = u'To align' if options.dry_run else u'Aligned'

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

    rows.append([u'**Total**'] + [str(v) for v in totals])
    output.print_table(
        table_data=rows,
        title='',
        columns=[u'Category', u'Selected', moved_header,
                 u'Already aligned', u'Skipped', u'Beyond tolerance'])


def print_moves_table(result, options):
    if not result.planned:
        return
    output.print_md(u'## Elements {}'.format(
        u'to align' if options.dry_run else u'aligned'))

    ordered = sorted(result.planned,
                     key=lambda r: (element_id_value(r.wall_id),
                                    -abs(r.distance_before)))
    shown = ordered[:MAX_MOVED_ROWS]

    columns = [u'Element', u'Category', u'Type', u'Wall', u'Face',
               u'Front', u'Travel', u'Dist. before', u'Dist. after',
               u'Move', u'Rotation', u'Connectors', u'Notes']
    if not options.dry_run:
        columns.append(u'Outcome')

    rows = []
    for record in shown:
        row = [
            output.linkify(record.element_id),
            record.category_name,
            record.type_name,
            partition_cell(record.wall_id, record.wall_label,
                           record.wall_is_linked),
            record.face_side,
            record.front_outcome,
            u'along the line' if record.along_front else u'along the normal',
            format_mm(record.distance_before),
            format_mm(record.distance_after),
            format_mm(record.translation_length),
            format_deg(record.rotation_rad) if record.needs_rotation else u'-',
            str(record.connected) if record.connected else u'-',
            record.note or u'',
        ]
        if not options.dry_run:
            if record.error:
                outcome = u'error: {}'.format(record.error)
            elif record.needs_rotation and not record.applied_rotation:
                outcome = u'moved but not squared up'
            elif record.needs_move and not record.applied_move:
                outcome = u'squared up but not moved'
            else:
                outcome = u'OK'
            row.append(outcome)
        rows.append(row)

    output.print_table(table_data=rows, title='', columns=columns)

    remaining = len(ordered) - len(shown)
    if remaining > 0:
        output.print_md(u'_...and {} more elements not listed. '
                        u'All of them were processed anyway._'.format(
                            remaining))


def print_already_ok(result):
    if not result.already_ok:
        return
    output.print_md(
        u'## Elements already aligned\n\n'
        u'**{}** elements were already in place (within {:.0f} mm of the '
        u'chosen face and {:.1f} degrees) and were left untouched.'.format(
            len(result.already_ok), POSITION_TOL_MM, ANGLE_TOL_DEG))


def print_skipped_table(result):
    if not result.skipped:
        return
    output.print_md(u'## Skipped elements')

    shown = result.skipped[:MAX_SKIPPED_ROWS]
    rows = []
    for record in shown:
        rows.append([
            output.linkify(record.element_id),
            record.category_name,
            record.reason,
            (partition_cell(record.wall_id, record.wall_label,
                            record.wall_is_linked)
             if record.wall_id else u'-'),
            format_mm(record.distance) if record.distance is not None else u'-',
        ])
    output.print_table(
        table_data=rows,
        title='',
        columns=[u'Element', u'Category', u'Reason', u'Nearest wall',
                 u'Measured distance'])

    remaining = len(result.skipped) - len(shown)
    if remaining > 0:
        output.print_md(u'_...and {} more skipped elements not '
                        u'listed._'.format(remaining))


def print_over_tolerance_table(result, options):
    if not result.over_tolerance_items:
        return
    output.print_md(u'## Skipped as beyond the tolerance')

    ordered = sorted(result.over_tolerance_items, key=lambda r: r.distance)
    shown = ordered[:MAX_OVER_TOLERANCE_ROWS]
    rows = []
    for record in shown:
        rows.append([
            output.linkify(record.element_id),
            record.category_name,
            partition_cell(record.wall_id, record.wall_label,
                           record.wall_is_linked),
            format_mm(record.distance),
            u'+{}'.format(format_mm(record.excess)),
        ])
    output.print_table(
        table_data=rows,
        title='',
        columns=[u'Element', u'Category', u'Nearest wall',
                 u'Distance', u'Overshoot'])

    remaining = len(ordered) - len(shown)
    if remaining > 0:
        output.print_md(u'_...and {} more elements not listed._'.format(
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
    return u'_Raising the tolerance to **{:.0f} cm** would bring in ' \
           u'**{}** more elements._'.format(candidate_cm, recovered)


def partition_cell(wall_id, label, is_linked):
    """Cella "partizione" del resoconto.

    Gli elementi di un modello collegato non sono selezionabili dall'host,
    quindi output.linkify() non produrrebbe un collegamento risolvibile: per
    quelli si stampa l'id nudo, che resta utilizzabile per una ricerca
    dentro al modello collegato.
    """
    if is_linked:
        return u'{} (id {})'.format(label, element_id_value(wall_id))
    return u'{} {}'.format(output.linkify(wall_id), label)


def print_wall_problems(result):
    if not result.wall_problems:
        return
    output.print_md(u'## Partitions discarded')
    rows = []
    for problem_source, wall, reason in result.wall_problems:
        label = wall_label(wall)
        if problem_source.is_linked:
            label = u'[{}] {}'.format(problem_source.link_name, label)
        rows.append([partition_cell(wall.Id, label, problem_source.is_linked),
                     reason])
    output.print_table(
        table_data=rows,
        title='',
        columns=[u'Partition', u'Reason'])


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
    output.print_table(
        table_data=rows,
        title='',
        columns=[u'Kind', u'Description', u'Elements'])
    output.print_md(u'_No constraint was unlocked or deleted: the tool '
                    u'applies no automatic resolutions._')


def print_warnings():
    if not WARNINGS:
        return
    output.print_md(u'## Warnings')
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
                record.error = u'transaction rolled back'

    print_header(options, elements, source_label, result)
    print_category_summary(result, options)
    print_moves_table(result, options)
    print_already_ok(result)
    print_skipped_table(result)
    print_over_tolerance_table(result, options)
    print_wall_problems(result)
    print_revit_failures(preprocessor)
    print_warnings()

    output.print_md(u'## Outcome')
    if rolled_back:
        output.print_md(
            u'**The edit was rolled back: the model was not touched.** '
            u'The {} planned elements are still in their original '
            u'position. The reason is in the warnings above.'.format(
                len(result.planned)))
    elif not result.planned:
        output.print_md(
            u'Nothing to align. The "Skipped elements" and "Skipped as '
            u'beyond the tolerance" tables say whether the problem is the '
            u'tolerance, the categories, the walls found or hosted '
            u'elements.')
    elif options.dry_run:
        output.print_md(
            u'**{}** elements would be aligned. Run the command again '
            u'with "Simulation" cleared to apply the change.'.format(
                len(result.planned)))
    else:
        output.print_md(
            u'Elements aligned: **{}** of {} planned.'.format(
                applied_ok, len(result.planned)))
        if applied_failed:
            output.print_md(
                u'Elements left unchanged because of a Revit error: '
                u'**{}**. The reason is in the "Outcome" column.'.format(
                    applied_failed))
        output.print_md(
            u'The change is grouped into a single undo step, named '
            u'"{}".'.format(TRANSACTION_NAME))


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
        forms.alert(u'No document is open.', exitscript=True)

    if doc.IsFamilyDocument:
        forms.alert(u'The command is not available in the family editor.',
                    exitscript=True)

    version = revit_version()
    if version and version < 2022:
        forms.alert(u'The tool requires Revit 2022 or later. '
                    u'Version detected: {}.'.format(version),
                    exitscript=True)

    if not MEP_CATEGORIES:
        forms.alert(u'None of the configured categories is available '
                    u'in this version of Revit.', exitscript=True)

    active_view = doc.ActiveView
    if active_view is None:
        forms.alert(u'No active view.', exitscript=True)
    if active_view.ViewType in build_non_pickable_view_types():
        forms.alert(
            u'Activate a graphical view (plan, section, 3D) before '
            u'running the command: elements are picked in the view.',
            exitscript=True)


def count_per_category(elements):
    """[(BuiltInCategory, nome, conteggio)] per le caselle della finestra.

    I conteggi vengono dalla SELEZIONE dell'utente, non da una ricerca nel
    modello. Le categorie che nella selezione non hanno nemmeno un elemento
    NON compaiono affatto: una riga a zero non e' una scelta che l'utente
    possa fare, e su una selezione stretta riempiva la lista di voci inerti
    fra cui cercare le due che contano davvero.

    Il filtro non cambia l'esito di nulla. Una categoria assente dalla
    selezione non ha elementi da spuntare o da togliere, quindi tenerla
    fuori dalle opzioni non sottrae niente all'utente e non fa sparire
    nessun elemento dal resoconto.
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
        count = counts.get(key, 0)
        if not count:
            continue
        info.append((built_in_category,
                     category_display_name(built_in_category),
                     count))
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


def document_has_walls():
    """True se il documento corrente contiene almeno un muro.

    Serve a distinguere due casi che l'utente vivrebbe allo stesso modo:
    "i dispositivi sono lontani dai muri" e "i muri non sono in questo
    modello". Nei progetti MEP l'architettonico e' spesso un collegamento, e
    un FilteredElementCollector sul documento corrente non vede gli elementi
    dei modelli collegati.
    """
    try:
        collector = DB.FilteredElementCollector(doc)\
            .WherePasses(build_category_filter(PARTITION_CATEGORIES))\
            .WhereElementIsNotElementType()
        for _ in collector:
            return True
    except Exception:
        return True
    return False


def report_no_partition_found(options):
    """Spiega perche' la ricerca non ha trovato nessuna partizione.

    Distingue il caso "nessun muro raggiungibile" dal caso "muri presenti ma
    lontani": il secondo non passa di qui, perche' i muri verrebbero comunque
    raccolti nella regione e ogni elemento finirebbe fra i saltati per
    tolleranza, che e' un messaggio del tutto diverso.
    """
    has_own = document_has_walls()
    link_count = count_loaded_links()

    output.close_others()
    output.print_md(u'# MEP alignment to partitions')

    if has_own:
        headline = (u'**No wall is near the selected elements.** This '
                    u'model does contain walls, but none falls in the '
                    u'region occupied by the selection.')
        hint = (u'Check that you selected the right elements, or raise '
                u'the tolerance.')
    elif options.include_links and link_count:
        headline = (u'**No wall found, neither in this model nor in the '
                    u'{} loaded links.**'.format(link_count))
        hint = (u'If the links were ignored, the reason is in the '
                u'warnings below: an unloaded, tilted or mirrored link is '
                u'not used.')
    elif link_count:
        headline = (u'**This model contains no walls**, but {} linked '
                    u'models are loaded.'.format(link_count))
        hint = (u'Searching in links is off: tick "Look for partitions '
                u'in linked models as well" in the command dialog.')
    else:
        headline = (u'**This model contains no walls and has no linked '
                    u'models loaded.**')
        hint = (u'Load the architectural link, or run the command in the '
                u'model that contains the walls.')

    output.print_md(u'{}\n\n{}'.format(headline, hint))
    print_warnings()

    forms.alert(u'No vertical partition found.\n\n'
                u'The report in the output panel explains why.',
                title=u'No partition found', exitscript=True)


def count_loaded_links():
    """Numero di modelli collegati effettivamente caricati."""
    total = 0
    try:
        for link in DB.FilteredElementCollector(doc)\
                .OfClass(DB.RevitLinkInstance).ToElements():
            try:
                if link.GetLinkDocument() is not None:
                    total += 1
            except Exception:
                continue
    except Exception:
        return 0
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
        output.print_md(u'# MEP alignment to walls')
        output.print_md(
            u'None of the **{}** selected elements has an insertion '
            u'point: they are all host based, work plane based or in '
            u'place. There is nothing to align in plan.'.format(
                len(elements)))
        forms.alert(
            u'None of the selected elements has an insertion point.',
            title=u'Nothing to align', exitscript=True)

    # Con la lista filtrata, una selezione priva di categorie gestite
    # aprirebbe una finestra senza nemmeno una casella: meglio dirlo.
    category_info = count_per_category(elements)
    if not category_info:
        output.close_others()
        output.print_md(u'# MEP alignment to walls')
        output.print_md(
            u'None of the **{}** selected elements belongs to the {} '
            u'categories this tool handles, so there is nothing it can '
            u'align.'.format(len(elements), len(MEP_CATEGORIES)))
        forms.alert(
            u'None of the selected elements belongs to a handled category.',
            title=u'Nothing to align', exitscript=True)

    window = MEPAlignToWallWindow(resolve_xaml_path(), elements, source_label,
                            category_info,
                            out_of_scope_count(elements))
    window.ShowDialog()

    options = window.options
    if options is None:
        script.exit()

    # Una sola passata sul documento per la regione occupata dalla
    # selezione; la restrizione per singolo elemento avviene poi su questo
    # insieme ridotto.
    sources = collect_partition_sources(points, options.tolerance_internal,
                                        options.include_links)
    found_any = False
    for source in sources:
        if source.ids is not None and source.ids.Count:
            found_any = True
            break

    if not found_any:
        # Due cause si presenterebbero all'utente allo stesso modo, un
        # resoconto di soli scarti generici: "i dispositivi sono lontani dai
        # muri" e "i muri non sono dove sto cercando". Vanno distinte.
        report_no_partition_found(options)

    linked_sources = [s for s in sources if s.is_linked]
    if linked_sources:
        warn(u'Partitions also searched in {} linked models: {}.'.format(
            len(linked_sources),
            u', '.join([s.link_name for s in linked_sources])))

    context = PlanContext(sources, active_design_option_id())

    try:
        result = build_plan(elements, options, context)
    except UserWarning:
        forms.alert(
            u'Analysis cancelled: the model was not modified.',
            title=u'Cancelled', exitscript=True)

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
                warn(u'Revit rolled the transaction back on an error: no '
                     u'change was applied. The detail is in the "Revit '
                     u'warnings raised while editing" section.')
                applied_failed += applied_ok
                applied_ok = 0
                rolled_back = True
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

    print_report(options, elements, source_label, result, preprocessor,
                 applied_ok, applied_failed, rolled_back)

    if not result.planned:
        forms.alert(
            u'Nothing to align.\n\n'
            u'The report in the output panel explains why, element by '
            u'element.',
            title=u'Nothing to do')


if __name__ == '__main__':
    main()
