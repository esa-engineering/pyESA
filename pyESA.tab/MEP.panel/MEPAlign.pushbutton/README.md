# MEP Align — strumento pyRevit

Allinea in pianta i dispositivi MEP selezionati alla partizione verticale architettonica
più vicina.

**Selezioni tu gli oggetti da allineare.** Per ognuno lo strumento cerca la partizione
verticale più vicina, porta il dispositivo **a filo della sua faccia** e lo ruota
dell'angolo minimo che lo rende perpendicolare alla parete. Se la partizione più vicina è
oltre la tolleranza che hai indicato, il dispositivo viene **saltato** e riportato con la
distanza misurata.

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

**1. Selezione degli elementi.** Il comando parte dagli elementi già selezionati in Revit.
Se non c'è selezione, chiede di indicarli nella vista (`Finish` per confermare, `Esc` per
annullare). Gli elementi scelti graficamente restano selezionati, così rilanciando il
comando — per esempio per applicare dopo una simulazione — non serve ripetere la selezione.

La selezione avviene **prima** della finestra di dialogo: con un dialogo modale WPF aperto
la vista di Revit non accetta la selezione, e conoscendo gli elementi in anticipo la
finestra può mostrare quanti ce ne sono per ogni categoria.

Le due strade si comportano diversamente, di proposito:

| Come selezioni | Cosa succede agli oggetti fuori dalle nove categorie |
| --- | --- |
| Selezione grafica dal comando | Il filtro non te li fa nemmeno indicare |
| Selezione fatta prima di lanciare | Vengono saltati e riportati con il motivo "categoria non gestita" |

**2. Categorie da allineare.** Le nove categorie gestite, tutte pre-spuntate. Fra parentesi
il numero di elementi di quella categoria **presenti nella tua selezione**: una categoria a
`(0)` dice subito che fra gli oggetti indicati non ce n'è nessuno di quel tipo. Togliendo
la spunta a una categoria i suoi elementi vengono saltati e riportati, non spariscono.

| Categorie gestite |
| --- |
| Air Terminals, Communication Devices, Data Devices, Electrical Fixtures, Fire Alarm Devices, Lighting Devices, Lighting Fixtures, Nurse Call Devices, Security Devices |

**3. Tolleranza.** Predefinita 30 cm, misurata in pianta fra il punto di inserimento
dell'elemento e la **faccia** della partizione più vicina, non il suo asse. Gli elementi che
cadono dentro lo spessore hanno distanza negativa e rientrano quindi sempre.

Oltre la tolleranza l'elemento **viene saltato**, e compare nel resoconto con la distanza
misurata e di quanto ha sforato. In fondo a quella tabella una riga dice di quanto alzare
la tolleranza e quanti elementi si recupererebbero: è il modo più rapido per capire se il
valore era troppo stretto, senza rilanciare il comando.

**4. Opzioni.**

- *Ruota gli elementi*: attiva per default. Disattivandola viene applicata solo la
  traslazione.
- *Salta gli elementi collegati*: non attiva per default. Vedi "Connettori" più sotto.
- *Simulazione*: calcola e riporta il risultato senza aprire alcuna transazione, quindi
  senza modificare il modello. **Conviene sempre partire da qui.**

Al termine, il pannello di output riporta cosa è stato allineato, cosa era già a posto,
cosa è stato saltato e perché, con collegamenti cliccabili ai singoli elementi.

> **Ogni elemento che hai selezionato compare nel resoconto**, in esattamente una delle
> quattro liste. Dato che gli oggetti li hai indicati uno per uno, un elemento che sparisce
> senza spiegazione sarebbe un difetto, non una semplificazione.

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
| categoria non gestita dallo strumento | selezionato prima di lanciare il comando, fuori dalle nove categorie |
| categoria esclusa nella finestra | la categoria esiste ma le hai tolto la spunta |
| nessuna partizione verticale nel raggio di ricerca | nessun muro vicino: controlla anche i modelli collegati |
| oltre la tolleranza: *x* dalla faccia più vicina | il caso richiesto esplicitamente; vedi la tabella dedicata |
| ospitato dal muro *n* | un elemento wall-hosted è già vincolato alla faccia del suo host |
| ospitato da *categoria n* | ospitato da soffitto, pavimento o altra faccia |
| elemento bloccato (pin) | `MoveElement` lancia un'eccezione sugli elementi bloccati |
| elemento nel gruppo *nome* | spostare un membro desincronizzerebbe tutte le istanze del gruppo |
| sotto-componente di famiglia annidata | non spostabile da solo |
| opzione di progetto non attiva | non modificabile |
| in prestito ad altro utente | modello workshared |
| elemento senza punto di inserimento | host based, workplane based o in place |
| fronte verticale | diffusore a controsoffitto: non ha un verso in pianta da ruotare. Scartato **solo se la rotazione è attiva**: con la sola traslazione l'elemento viene comunque spostato |
| ingombro fuori dall'estensione verticale dei muri | nessuna delle partizioni vicine arriva alla quota dell'elemento |
| oltre l'estremità del muro di *x* | la proiezione cade fuori dalla testata |

Un elemento vicino a più partizioni viene assegnato a quella con il **valore assoluto**
della distanza dalla faccia più piccolo, e se la seconda è a meno di 20 mm di scarto il caso
viene segnalato come ambiguo nella colonna Note.

Con la ricerca automatica l'ambiguità non è più un caso raro: in un angolo due muri sono
quasi equidistanti. Vale la pena leggere quella colonna, perché è il motivo principale per
cui un dispositivo può finire sulla faccia sbagliata.

Le partizioni che **non coprono la quota** del dispositivo vengono tolte dalla gara prima
del confronto: un muretto basso non è un bersaglio valido per un rilevatore montato in
alto, anche se in pianta gli sta più vicino di tutti.

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
- **L'inclinazione del muro si misura sulla geometria, non sui parametri.** La prima
  versione leggeva `BuiltInParameter.WALL_CROSS_SECTION` assumendo `0 = Vertical`, e su un
  progetto fatto di soli muri verticali scartava ogni muro: il valore intero di quel
  parametro non significa quello che sembra. Ora si legge la componente Z della normale
  delle facce laterali (`HostObjectUtils.GetSideFaces` più `Face.ComputeNormal`), che su
  una faccia verticale vale zero per definizione. L'angolo `Angle From Vertical` resta
  come scorciatoia per confermare la verticalità ed evitare il calcolo geometrico nel caso
  normale, ma **da solo non basta mai a scartare un muro**. Se la geometria non è
  leggibile il muro si assume verticale: un falso negativo sbaglia un muro, un falso
  positivo rende il comando inutilizzabile sull'intero progetto.
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

## Modelli collegati

Nei progetti MEP l'architettonico è quasi sempre un collegamento, e un
`FilteredElementCollector` sul documento corrente non vede i suoi elementi. Lo strumento
raccoglie quindi le partizioni **anche dai modelli collegati caricati**, controllato dalla
casella *Cerca le partizioni anche nei modelli collegati* (attiva per default).

Host e collegamenti concorrono insieme: vince la partizione più vicina, da qualunque
documento provenga.

### Come funziona

Sotto una trasformazione **rigida** tutte le grandezze scalari del calcolo sono
**invarianti**: la distanza dalla faccia, lo spessore del muro, lo scostamento della linea
di posizionamento e la sporgenza oltre la testata valgono lo stesso nelle due terne. Quindi
il muro collegato non viene mai trasformato. Si fa il contrario:

1. il punto di inserimento dell'elemento, che vive in coordinate host, viene portato nelle
   coordinate del collegamento con la trasformazione inversa;
2. la matematica già collaudata gira lì dentro senza sapere nulla dei link;
3. solo la **normale uscente** viene riportata in coordinate host, perché è l'unica
   direzione che deve convivere con il fronte dell'elemento e con il punto bersaglio.

È il motivo per cui il supporto ai collegamenti non ha richiesto di toccare il nucleo
geometrico.

### Quando un collegamento viene ignorato

Sempre con il motivo scritto negli avvisi del resoconto:

| Caso | Perché |
| --- | --- |
| collegamento non caricato | non c'è nessun documento da interrogare |
| trasformazione non leggibile | senza trasformazione non si può convertire nulla |
| collegamento inclinato o capovolto | la normale riportata nell'host avrebbe una componente verticale e lo spostamento cambierebbe la quota |
| collegamento speculare | i versi esterno e interno delle facce risulterebbero invertiti e i dispositivi finirebbero sulla faccia sbagliata |
| collegamento in scala | le distanze misurate nel collegamento non sarebbero confrontabili con la tolleranza, che è espressa in coordinate host |

La regola è la stessa applicata altrove nello strumento: meglio saltare un riferimento
dichiarandolo, che usarlo e spostare dispositivi nel punto sbagliato.

### Nel resoconto

Le partizioni collegate compaiono come `[Nome collegamento] Tipo di muro (id 123456)`.
L'id è stampato nudo e non come collegamento cliccabile: gli elementi di un modello
collegato non sono selezionabili dall'host, quindi un link non risolverebbe. L'id resta
utilizzabile per cercarli aprendo il modello collegato.

I muri collegati sono usati **solo come riferimento**: la fase di applicazione lavora
esclusivamente sul documento host e non tocca mai un modello collegato.
