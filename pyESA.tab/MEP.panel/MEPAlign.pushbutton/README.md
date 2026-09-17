# MEP Align — strumento pyRevit

Allinea in pianta i dispositivi MEP alle facce dei muri selezionati.

Ogni dispositivo entro una distanza indicata dalla faccia del muro viene portato **a filo
della faccia** e ruotato dell'angolo minimo che lo rende perpendicolare alla parete.

> **La quota Z non viene mai modificata.** È il motivo per cui il primo elemento
> architettonico gestito è il muro. L'allineamento in quota è un problema diverso, perché
> richiede un livello o un soffitto come riferimento, ed è fuori ambito per questa versione.

## Contenuto della cartella

| File | Ruolo |
| --- | --- |
| `MEPAlign_script.py` | Logica dello strumento (Revit API, motore IronPython) |
| `MEPAlignWindow.xaml` | Interfaccia grafica della finestra di dialogo |
| `bundle.yaml` | Etichetta e descrizione del pulsante nella barra pyRevit |
| `icon.png` / `icon.dark.png` | Icone per il tema chiaro e per quello scuro |
| `README.md` | Questo file |

## Uso

**1. Selezione dei muri.** Il comando parte dai muri già selezionati in Revit. Se la
selezione non contiene muri, chiede di indicarli nella vista (`Finish` per confermare,
`Esc` per annullare). I muri scelti graficamente restano selezionati, così rilanciando il
comando non serve ripetere la selezione.

La selezione avviene **prima** della finestra di dialogo: con un dialogo modale WPF aperto
la vista di Revit non accetta la selezione, e conoscendo i muri in anticipo la finestra può
mostrare quanti elementi candidati ci sono per ogni categoria.

**2. Categorie da allineare.** Le nove categorie gestite, tutte pre-spuntate. Fra parentesi
il numero di istanze che ricadono nel volume di ricerca calcolato con la distanza
predefinita di 30 cm: una categoria a `(0)` dice subito che da lì non arriverà nulla.

| Categorie gestite |
| --- |
| Air Terminals, Communication Devices, Data Devices, Electrical Fixtures, Fire Alarm Devices, Lighting Devices, Lighting Fixtures, Nurse Call Devices, Security Devices |

**3. Distanza massima.** Predefinita 30 cm, misurata in pianta fra il punto di inserimento
dell'elemento e la **faccia** del muro più vicina, non il suo asse. Gli elementi che cadono
dentro lo spessore del muro hanno distanza negativa e rientrano quindi sempre nella soglia.

**4. Opzioni.**

- *Ruota gli elementi*: attiva per default. Disattivandola viene applicata solo la
  traslazione.
- *Salta gli elementi collegati*: non attiva per default. Vedi "Connettori" più sotto.
- *Simulazione*: calcola e riporta il risultato senza aprire alcuna transazione, quindi
  senza modificare il modello. **Conviene sempre partire da qui.**

Al termine, il pannello di output riporta cosa è stato allineato, cosa era già a posto,
cosa è stato ignorato e perché, e cosa è rimasto appena fuori soglia, con collegamenti
cliccabili ai singoli elementi.

## Come viene ricavato il piano di riferimento del muro

`Wall.Location.Curve` **non è l'asse del muro**: è la linea di posizionamento, che dipende
dal parametro *Location Line*. Ignorare questo dettaglio produce un errore fino a mezzo
spessore di muro. Lo scostamento firmato della linea di posizionamento rispetto alla
mezzeria, misurato lungo la normale esterna, vale:

| Location Line | Scostamento |
| --- | --- |
| Wall Centerline | `0` |
| Core Centerline | `(d_int - d_ext) / 2` |
| Finish Face: Exterior | `+ W/2` |
| Finish Face: Interior | `- W/2` |
| Core Face: Exterior | `W/2 - d_ext` |
| Core Face: Interior | `d_int - W/2` |

dove `W` è lo spessore del muro e `d_ext` / `d_int` sono le somme degli spessori degli
strati fuori dal nucleo. `CompoundStructure.GetLayers()` è ordinato dall'esterno verso
l'interno. Con `d_ext = d_int = 0` i tre casi "core" collassano sui casi "finish" e
"centerline": è il controllo di coerenza della tabella.

### Perché non si usa `HostObjectUtils.GetSideFaces`

A prima vista sarebbe la strada più pulita, perché delegherebbe a Revit tutta la questione
della linea di posizionamento. È stata scartata per due motivi concreti:

1. `Face.Project` restituisce `null` se il punto più vicino cade **fuori dalla faccia**, e
   una porta o una finestra **bucano la faccia laterale del muro**. Un dispositivo davanti
   al vano di una porta proietterebbe dentro il buco e verrebbe scartato senza motivo.
2. `null` non permette di distinguere "40 mm oltre la testata del muro" da "tre stanze più
   in là", informazione che serve al resoconto.

Il calcolo analitico risponde sempre, e con una distanza **firmata**. `GetSideFaces` resta
come riserva per i muri privi di linea di posizionamento.

## La rotazione minima

Fra le due direzioni perpendicolari al muro (la normale uscente e la sua opposta) viene
scelta quella più vicina all'orientamento attuale dell'elemento. L'angolo applicato non
supera quindi mai 90 gradi, e un elemento già orientato bene non viene ribaltato.

> **Conseguenza da conoscere:** un elemento montato al contrario viene reso perpendicolare
> al muro ma **non raddrizzato**. È il prezzo della regola "verso conservato", che evita
> rotazioni di 180 gradi indesiderate su famiglie autorate con il fronte invertito.

## Elementi ignorati

Ogni esclusione compare nel resoconto con il **motivo specifico** e con la **distanza
misurata**, anche quando il motivo non c'entra con la distanza: un elemento bloccato a
62 mm vale la pena di sbloccarlo, uno a 290 mm probabilmente no.

| Motivo | Nota |
| --- | --- |
| ospitato dal muro *n* | un elemento wall-hosted è già vincolato alla faccia del suo host |
| ospitato da *categoria n* | ospitato da soffitto, pavimento o altra faccia |
| elemento bloccato (pin) | `MoveElement` lancia un'eccezione sugli elementi bloccati |
| elemento nel gruppo *nome* | spostare un membro desincronizzerebbe tutte le istanze del gruppo |
| sotto-componente di famiglia annidata | non spostabile da solo |
| opzione di progetto non attiva | non modificabile |
| in prestito ad altro utente | modello workshared |
| elemento senza punto di inserimento | host based, workplane based o in place |
| fronte verticale | diffusore a controsoffitto: non ha un verso in pianta da allineare |
| ingombro fuori dall'estensione verticale dei muri | elemento di un altro piano |
| oltre l'estremità del muro di *x* | la proiezione cade fuori dalla testata |

Un elemento vicino a più muri selezionati viene assegnato a quello con il **valore assoluto**
della distanza dalla faccia più piccolo, e se il secondo muro è a meno di 20 mm di scarto il
caso viene segnalato come ambiguo nella colonna Note.

## Connettori

Spostare un terminale aria collegato a un canale passa per lo stesso motore del comando di
spostamento dell'interfaccia: Revit prova a mantenere la connessione allungando il canale e,
se non riesce, **rompe la connessione con un avviso**, non con un'eccezione.

La casella *Salta gli elementi collegati* è **non spuntata per default**, perché gli Air
Terminals sono fra le categorie richieste e saltarli per default renderebbe il comando
inerte proprio dove serve. Il numero di connettori collegati compare in una colonna del
resoconto, così il rischio è visibile e controllabile. Da rivedere dopo il primo collaudo
su modelli reali.

## Avvisi di Revit

Lo strumento **non** usa `revit.Transaction(swallow_errors=True)` né
`revit.ErrorSwallower()`. Entrambi si appoggiano al `FailureSwallower` di pyRevit, la cui
lista `RESOLUTION_TYPES` include `UnlockConstraints` e `DeleteElements`: su un avviso
*"Constraints are not satisfied"* **sbloccherebbe in silenzio le quote e gli allineamenti
dell'utente**. Su uno strumento che sposta geometria è esattamente il comportamento da non
avere.

Al suo posto c'è un `IFailuresPreprocessor` dedicato che registra ogni avviso per il
resoconto, elimina soltanto i warning privi di risoluzioni (quelli che aprirebbero un popup
e bloccherebbero il lotto a metà), non applica nessuna risoluzione e chiede il rollback in
presenza di errori.

## Ordine delle operazioni

Prima la rotazione, poi la traslazione. La traslazione viene **ricalcolata dal punto
corrente** verso un punto bersaglio assoluto memorizzato in fase di analisi, perché non è
garantito che `RotateElement` attorno a un asse passante per il punto di inserimento lasci
quel punto esattamente invariato per ogni tipo di famiglia. Ricalcolandola dopo la
rotazione, qualunque deriva si autocorregge, e la quota resta invariata per costruzione
perché la componente Z del vettore è forzata a zero.

Ogni elemento viene elaborato in una `SubTransaction` propria: se la rotazione riesce e la
traslazione fallisce, l'elemento torna intatto invece di restare ruotato e non spostato.
L'intera operazione è comunque un unico passo di annullamento, nominato
*"Allineamento MEP ai muri"*.

## Motore Python

Lo script gira sul motore **IronPython**, che è quello predefinito di pyRevit: il file non
deve contenere la direttiva `#! python3`. Il modulo `pyrevit.forms`, usato per caricare la
finestra XAML, si appoggia a `wpf.LoadComponent`, disponibile solo in IronPython. Sotto il
motore CPython3 le finestre WPF di pyRevit non sono attualmente supportate: si veda la issue
[pyrevitlabs/pyRevit#3033](https://github.com/pyrevitlabs/pyRevit/issues/3033).

## Punti aperti e non verificati

Da risolvere in Revit. Il risultato va scritto qui, non lasciato nel codice come
assunzione tacita.

1. **Segno di `BasisZ.CrossProduct(tangente)` rispetto a `Wall.Orientation`.** Sui muri
   rettilinei si usa direttamente `Wall.Orientation`, che è costante e autorevole, quindi
   non è un problema. Sugli archi la normale varia lungo il muro e va ricavata dalla
   tangente. Lo script esegue un autocontrollo permanente su ogni muro rettilineo e
   registra un avviso in caso di discordanza.

   Il rischio è però più contenuto di quanto sembri. Invertendo `n` si inverte anche
   `signed_center`, e quindi `side`, e il punto bersaglio
   `p + n * (side * half_width - signed_center)` resta identico, perché il termine cambia
   segno due volte. **Con Location Line = Wall Centerline lo spostamento è quindi
   invariante rispetto al segno della normale.** Un segno sbagliato altera solo
   l'etichetta *faccia esterna / interna* nel resoconto e la correzione di scostamento sui
   muri la cui Location Line non è la mezzeria.
2. **Parametrizzazione degli `Arc`**, angolo in radianti o lunghezza d'arco. La
   documentazione è ambigua: dice che per rette e archi il parametro grezzo misura in
   piedi, e al tempo stesso `Arc.Period` vale `2*pi`. Lo script verifica a runtime quale
   delle due valga, quindi è corretto in entrambi i casi, ma va confermato quale ramo si
   attiva.
3. **Mappatura di `WALL_KEY_REF_PARAM` sull'enumeratore `WallLocationLine`** (0…5
   nell'ordine della tabella sopra). Su un valore inatteso lo script ripiega sulla mezzeria
   e registra un avviso. Il caso di prova che lo verifica è un muro con *Location Line*
   impostata su `Finish Face: Exterior`.
4. **`RotateElement` su un elemento bloccato**: lancia un'eccezione o falisce in silenzio?
   È documentato solo per `MoveElement`. Il pre-controllo su `Pinned` rende la domanda
   accademica, ma il pre-controllo non va rimosso.
5. **`ElementTransformUtils` su un membro di gruppo**: eccezione, failure o successo. Anche
   qui il pre-controllo precede la domanda.
6. **`LocationPoint.Point` è aggiornato subito dopo `RotateElement`** o serve un
   `doc.Regenerate()`? Se in prova la rilettura risultasse non aggiornata, basta alzare la
   costante `FORCE_REGEN` in testa allo script.
7. **Effetto reale di `MoveElement` su un diffusore collegato a un canale rigido.**

## Comportamenti noti e accettati

- Il punto di inserimento finisce **esattamente sulla faccia**. Se una famiglia ha
  l'origine al centro del proprio ingombro invece che sul retro, l'elemento risulta per
  metà dentro il muro. Le colonne *distanza prima* e *distanza dopo* del resoconto servono
  a rilevarlo sul primo modello reale. Se si rivelasse un problema diffuso, la versione
  successiva introdurrà un offset impostabile o l'allineamento del bordo dell'ingombro.
- Un elemento montato al contrario viene reso perpendicolare ma non raddrizzato.
- Un elemento di un altro piano, allineato in pianta con il muro, **non compare affatto nel
  resoconto**: viene escluso dal filtro verticale a monte. È corretto, perché elencare ogni
  elemento fuori portata del modello renderebbe la tabella inutilizzabile.
- Muri tenda, muri inclinati, muri senza linea di posizionamento e muri con estensione
  verticale non leggibile vengono scartati con il motivo nella tabella *Muri scartati*.
  I muri stacked non vengono scartati ma espansi nei loro sotto-muri.
- Curve di posizionamento diverse da retta e arco (ellissi, spline) non sono gestite e
  producono un avviso.
- Il volume di ricerca è dilatato di 1 m oltre la soglia, perché il filtro di prossimità
  confronta il *bounding box* dell'elemento mentre il criterio vero è la distanza del
  *punto di inserimento*: esistono famiglie con la geometria modellata lontano
  dall'origine. Il margine non allarga la soglia, che viene verificata a valle in modo
  esatto.

## Avvertenze operative

- Eseguire sempre prima in **Simulazione**, leggere il resoconto, e solo dopo applicare.
- L'operazione è annullabile con un solo Undo di Revit.
- Tag e quote agganciati agli elementi spostati li seguono, ma la loro posizione relativa
  può diventare illeggibile: non c'è nulla che lo strumento possa fare via API.
