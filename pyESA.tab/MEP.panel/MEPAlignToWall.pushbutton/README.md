# MEP Align — strumento pyRevit

Allinea in pianta i dispositivi MEP selezionati alla partizione verticale architettonica
più vicina.

**Selezioni tu gli oggetti da allineare.** Per ognuno lo strumento cerca la partizione
verticale più vicina e porta il dispositivo **a filo di una delle sue due facce**. Se la
partizione più vicina è oltre la tolleranza che hai indicato, il dispositivo viene
**saltato** e riportato con la distanza misurata.

**Quale delle due facce lo decide il fronte della famiglia**, letto dal suo Room
Calculation Point: un dispositivo che guarda la partizione invece di allontanarsene viene
portato dall'altra parte. Lo spostamento avviene quindi nei due versi lungo la normale del
muro, non solo verso la faccia più vicina. Vedi "Quale delle due facce" più sotto.

**Lo spostamento avviene lungo la retta di analisi**, cioè l'asse di vista del
dispositivo, non lungo la normale della parete: un apparecchio che guarda il muro in
obliquo scorre quindi anche lateralmente. E se più istanze di muro sono a contatto, il
bersaglio è la **faccia più esterna del pacchetto**, non quella della prima partizione
incontrata.

> **Sull'orientamento lo strumento interviene solo per raddrizzare**, portando il fronte
> perpendicolare alla faccia, e mai di più di 5 gradi. Non è un riorientamento: vedi
> "Il raddrizzamento" più sotto.

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

Lo stesso valore regola anche la **retta con cui si riconosce il fronte** della
famiglia, che è lunga il doppio della tolleranza: alzando la tolleranza si allarga sia
il raggio di ricerca sia il cono entro cui il fronte viene considerato rivolto alla
partizione. Vedi "Quale delle due facce".

**4. Opzioni.**

| Opzione | Default | Effetto |
| --- | --- | --- |
| *Cerca le partizioni anche nei modelli collegati* | attiva | Vedi "Modelli collegati" più sotto |
| *Salta gli elementi collegati ad altri (connettori)* | non attiva | Vedi "Connettori" più sotto |
| *Simulazione: non modificare il modello* | non attiva | Calcola e riporta senza aprire alcuna transazione. **Conviene sempre partire da qui** |

Ogni opzione ha un tooltip con la spiegazione estesa: passandoci sopra si legge il
perché, senza che la finestra debba contenerlo tutto.

Al termine, il pannello di output riporta cosa è stato allineato, cosa era già a posto,
cosa è stato saltato e perché, con collegamenti cliccabili ai singoli elementi. Fanno
eccezione le partizioni dei modelli collegati, che non sono selezionabili dall'host e per
cui viene stampato l'id nudo.

> **Ogni elemento che hai selezionato compare nel resoconto**, in esattamente una delle
> quattro liste. Dato che gli oggetti li hai indicati uno per uno, un elemento che sparisce
> senza spiegazione sarebbe un difetto, non una semplificazione.

## Più istanze di muro a contatto

Una parete modellata come più istanze accostate (la muratura più il suo rivestimento, due
tramezzi affiancati, un contromuro tecnico) viene trattata come **un pacchetto unico**: il
dispositivo si allinea alla faccia più esterna, non alla prima che incontra. Senza questa
regola un apparecchio modellato dentro il muro strutturale finirebbe a filo del suo
rivestimento interno, cioè dentro la stratigrafia.

Si lavora su un asse orientato come la normale del muro vincente, con l'origine nel punto
di inserimento, su cui ogni partizione parallela occupa un intervallo noto. Partendo dalla
faccia scelta il confine viene spinto in fuori finché esiste una partizione che lo contiene
o lo sfiora e che si estende oltre.

Tre vincoli tengono la regola stretta:

| Vincolo | Perché |
| --- | --- |
| solo partizioni **parallele**, entro circa 2,5 gradi | un muro che forma un angolo non è parte dello stesso pacchetto |
| solo partizioni su cui l'elemento **si proietta davvero** | una parete che finisce prima di arrivargli davanti non gli è adiacente |
| solo scarti sotto **2 mm** | un'intercapedine vera è molto più larga e non deve essere attraversata |

Il confine può solo essere **spinto più in fuori**: il muro vincente resta quello scelto
per distanza e la sua faccia resta quella scelta dal fronte. Quando il pacchetto entra in
gioco, la colonna *Note* del resoconto dice quante murature sono state attraversate e
qual è l'ultima.

## Il raddrizzamento

Dopo aver scelto muro e faccia, lo strumento porta il fronte **perpendicolare alla faccia
bersaglio**, ruotando l'elemento attorno a un asse verticale passante per il suo punto di
inserimento.

È una correzione fine, non un riorientamento, e lo è **per costruzione**: quella partizione
è stata ammessa solo perché il fronte le stava già entro 5 gradi dalla perpendicolare,
quindi l'angolo da recuperare non può superare i 5 gradi. Le due regole sono la stessa cosa
vista da due lati: un dispositivo montato davvero di traverso non viene raddrizzato, viene
*saltato* prima, perché per lui nessuna partizione risulta fronteggiata.

Nel codice la guardia sui 5 gradi è comunque scritta esplicitamente, e se scattasse lo
direbbe in colonna *Note*. Non perché possa scattare oggi, ma perché quell'invariante si
regge su due funzioni diverse e un domani potrebbe rompersi in silenzio.

Due cose da sapere:

- **Si ruota il fronte, cioè la direzione del Room Calculation Point, non
  `FacingOrientation`.** La rotazione è rigida e porta con sé anche il punto di calcolo,
  quindi dopo il comando il fronte è perpendicolare al muro qualunque sia la convenzione
  con cui la famiglia è stata autorata. È il motivo per cui il risultato non dipende da chi
  ha disegnato la famiglia.
- **Il rovescio della stessa medaglia:** se in una famiglia il punto di calcolo è autorato
  leggermente di sbieco, ma entro il cono, l'elemento viene ruotato di quel tanto anche se
  era montato bene. L'errore è limitato a 5 gradi e la colonna *Rotazione* lo rende
  visibile, ma se vedi un'intera famiglia ruotata sempre dello stesso angolo, il problema è
  dov'è il pallino in quella famiglia.

Un elemento senza Room Calculation Point **non viene ruotato affatto**: non c'è nessun verso
di cui fidarsi. Viene solo traslato, come nelle versioni precedenti.

## La direzione dello spostamento

Lo spostamento avviene **esclusivamente lungo la retta di analisi**: il dispositivo scivola
sul proprio asse di vista finché il punto di inserimento non raggiunge il piano della
faccia bersaglio. Non lungo la normale del muro.

Due conseguenze da conoscere:

- **La posizione lungo il muro non è più invariante.** Un apparecchio che guarda la parete
  in obliquo scivola anche di lato. Su un fronte perpendicolare non si vede affatto.
- **La corsa vale 1/cos della distanza da coprire.** È il motivo per cui il fronte deve
  stare entro **5 gradi** dalla perpendicolare al muro: oltre, l'elemento scivolerebbe
  lungo la parete invece di avvicinarglisi. Entro quel cono la corsa supera la distanza
  dello 0,4% e lo scivolamento laterale resta sotto il 9%, quindi in pratica lo
  spostamento è perpendicolare: la retta pesa soprattutto sulla **scelta** del muro e
  della faccia, non sulla direzione della corsa.

La colonna *Corsa* del resoconto dice, elemento per elemento, se lo spostamento è avvenuto
*lungo la retta* o *lungo la normale*. Il secondo caso è quello delle famiglie senza Room
Calculation Point e dei fronti troppo obliqui: senza una retta da seguire resta la
perpendicolare, che è il comportamento delle versioni precedenti.

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

## Quale delle due facce

Portare il dispositivo sempre sulla faccia più vicina è corretto solo finché è già
modellato dalla parte giusta della partizione. Non lo è quando sta **dentro lo spessore**
del muro, e non lo è quando è stato inserito **dal lato sbagliato**: in quei casi il
dispositivo finisce a filo della faccia che ha alle spalle, cioè dentro la stanza
sbagliata.

Serve quindi distinguere il fronte della famiglia dal suo retro, e `FacingOrientation` non
basta: dipende da come è orientato il sistema di riferimento con cui la famiglia è stata
autorata, che non è una convenzione rispettata da tutti i produttori.

Il **Room Calculation Point** è il punto che Revit usa per stabilire in quale locale sta il
dispositivo, quindi per costruzione cade *davanti* all'apparecchio, dentro la stanza che
serve. Il procedimento è questo:

1. si prende il Room Calculation Point dell'istanza e lo si **proietta sul piano
   orizzontale passante per il punto di inserimento**;
2. la direzione che va dal punto di inserimento a quella proiezione è il **fronte**;
3. lungo il fronte si traccia dal punto di inserimento una retta lunga **due volte la
   tolleranza** impostata all'avvio;
4. se la retta **intercetta la superficie del muro**, il dispositivo sta guardando la
   partizione invece di allontanarsene, e viene quindi portato sulla **seconda superficie
   di finitura** e non sulla prima.

Attenzione al punto 4: un dispositivo che si trova a meno della tolleranza da una parete
e la **guarda** viene portato dall'altra parte di quella parete. È la regola richiesta, ed
è corretta per l'apparecchio modellato dal lato sbagliato, ma su un dispositivo che sta
semplicemente davanti a un muro e lo fronteggia produce uno spostamento lungo tutto lo
spessore. Il montaggio normale, con le **spalle** al muro, non è toccato: lì la retta si
allontana e la faccia resta quella vicina.

### Le pareti parallele alla retta sono escluse

Non solo la faccia: il fronte **esclude le partizioni che corrono parallele alla retta di
analisi**. Una parete parallela alla retta è una parete che il dispositivo ha *di fianco*,
non davanti: allinearcelo significherebbe spostarlo lungo una normale ortogonale al suo
asse di vista, cioè proprio lo spostamento di traverso che la retta serve a evitare.

**L'esclusione è secca, senza ripiego.** Se dopo il filtro non resta nessuna partizione,
l'elemento viene **saltato** e compare fra gli ignorati con il motivo *nessuna partizione
fronteggiata*. Meglio lasciarlo dov'è dicendolo, che allinearlo a un riferimento che non è
il suo: un ripiego "se non ne resta nessuna concorrono tutte" sembra prudente, ma riporta
esattamente il difetto che il filtro doveva togliere, e proprio nel caso peggiore, quello
in cui l'unica parete vicina è quella sbagliata.

Fra le partizioni ammesse il criterio resta "vince la più vicina": il fronte decide chi
entra in gara, non chi la vince. Quando un muro più vicino viene escluso, la colonna *Note*
lo dichiara con il suo nome e la sua distanza.

Il filtro guarda l'orientamento e **non il verso**, perché la retta di analisi è una retta e
non una semiretta: una parete davanti e una alle spalle sono entrambe cose che il
dispositivo guarda, ed è giusto così, visto che il montaggio normale è proprio con le
spalle al muro.

La soglia è **5 gradi** (`FRONT_MAX_ANGLE_DEG`): una parete che si discosta di più di 5
gradi dalla perpendicolare al fronte è considerata parallela alla retta ed esclusa. È una
soglia stretta di proposito, perché un dispositivo a parete è per costruzione
perpendicolare al proprio muro: se non lo è entro quel margine, quel muro quasi certamente
non è il suo riferimento.

Senza Room Calculation Point non esiste nessuna retta, quindi non c'è nulla a cui una
parete possa essere parallela, e concorrono tutte come nelle versioni precedenti.

La proiezione al punto 1 non è un dettaglio: su un dispositivo a parete il punto di calcolo
sta quasi sempre anche più in alto o più in basso dell'origine, e la componente verticale
falserebbe l'angolo rispetto alla normale del muro.

Il test non tocca la geometria del muro. Sotto una trasformazione rigida tutte le grandezze
in gioco sono scalari invarianti (lo scostamento firmato dalla mezzeria, il semispessore e
la componente del fronte lungo la normale), quindi **le partizioni dei modelli collegati
non richiedono nessuna conversione in più**.

### Perché la retta ha una lunghezza finita

La retta lunga il doppio della tolleranza copre da sola un cono di 60 gradi, quindi è
abbondante rispetto al cono di 5 gradi entro cui una parete conta come fronteggiata: per un
dispositivo fuori dal muro non è mai la sua lunghezza a decidere. Resta vincolante solo per
un dispositivo **sepolto dentro** una partizione più spessa della tolleranza, dove la retta
può non raggiungere la faccia di uscita.

### Quando il fronte non si può leggere

L'elemento **viene comunque allineato** con il criterio posizionale di sempre, e la colonna
*Fronte* del resoconto dice con quale criterio, elemento per elemento:

| Valore | Significato | Faccia scelta | Corsa |
| --- | --- | --- | --- |
| *verso il muro* | la retta intercetta la partizione | quella verso cui guarda il dispositivo | lungo la retta |
| *opposto al muro* | la retta si allontana dalla partizione | la più vicina | lungo la retta |
| *radente al muro* | il fronte si discosta di oltre 5 gradi dalla perpendicolare, oppure la retta non arriva alla faccia | la più vicina | lungo la normale |
| *non definito* | la famiglia non espone il Room Calculation Point, oppure l'autore non lo ha spostato dall'origine | la più vicina | lungo la normale |

In testa al resoconto due righe contano quanti elementi sono stati portati sulla faccia
opposta e su quanti il fronte non era leggibile. La seconda è la prima cosa da guardare se
il risultato non convince: un lotto tutto a *non definito* vuol dire che quelle famiglie il
punto di calcolo non ce l'hanno, e il comando si è comportato come nella versione
precedente.

## Elementi ignorati

Ogni esclusione compare nel resoconto con il **motivo specifico** e con la **distanza
misurata**, anche quando il motivo non c'entra con la distanza: un elemento bloccato a
62 mm vale la pena di sbloccarlo, uno a 290 mm probabilmente no.

| Motivo | Nota |
| --- | --- |
| categoria non gestita dallo strumento | selezionato prima di lanciare il comando, fuori dalle nove categorie |
| categoria esclusa nella finestra | la categoria esiste ma le hai tolto la spunta |
| nessuna partizione verticale nel raggio di ricerca | nessun muro vicino: controlla anche i modelli collegati |
| nessuna partizione fronteggiata | ci sono muri vicini, ma corrono tutti paralleli alla retta di analisi: il dispositivo li ha di fianco, non davanti |
| oltre la tolleranza: *x* dalla faccia più vicina | il caso richiesto esplicitamente; vedi la tabella dedicata |
| ospitato dal muro *n* | un elemento wall-hosted è già vincolato alla faccia del suo host |
| ospitato da *categoria n* | ospitato da soffitto, pavimento o altra faccia |
| elemento bloccato (pin) | `MoveElement` lancia un'eccezione sugli elementi bloccati |
| elemento nel gruppo *nome* | spostare un membro desincronizzerebbe tutte le istanze del gruppo |
| sotto-componente di famiglia annidata | non spostabile da solo |
| opzione di progetto non attiva | non modificabile |
| in prestito ad altro utente | modello workshared |
| elemento senza punto di inserimento | host based, workplane based o in place |
| ingombro fuori dall'estensione verticale dei muri | nessuna delle partizioni vicine arriva alla quota dell'elemento |
| oltre l'estremità del muro di *x* | la proiezione cade fuori dalla testata |
| proiezione sulla geometria del muro non calcolabile | partizioni trovate, ma nessuna ha prodotto una proiezione utilizzabile |
| errore in analisi: *messaggio* | eccezione Revit su quel singolo elemento; il lotto prosegue |

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

## Applicazione

Prima il raddrizzamento, attorno a un asse verticale passante per il punto di inserimento
originale, che è noto con certezza. Poi la traslazione, **ricalcolata dal punto corrente**
verso un punto bersaglio assoluto memorizzato in fase di analisi, invece di riusare il
vettore calcolato allora: non è garantito che `RotateElement` lasci il punto di inserimento
esattamente invariato per ogni tipo di famiglia, e ricalcolando dopo la rotazione qualunque
deriva si autocorregge. La quota resta invariata per costruzione, perché la componente Z del
vettore è forzata a zero.

Ogni elemento viene elaborato in una `SubTransaction` propria: se la rotazione riesce e la
traslazione fallisce, l'elemento torna intatto invece di restare ruotato e non spostato. L'intera operazione è comunque un unico passo di annullamento,
nominato *"Allineamento MEP ai muri"*.

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
4. **`ElementTransformUtils` su un membro di gruppo**: eccezione, failure o successo. Il
   pre-controllo precede la domanda.
5. **Effetto reale di `MoveElement` su un diffusore collegato a un canale rigido.**
6. **Il Room Calculation Point di un dispositivo a parete cade davanti all'apparecchio.**
   È la convenzione, ed è il senso stesso del punto (Revit lo usa per assegnare il
   dispositivo a un locale, e il locale sta davanti), ma su famiglie di terze parti va
   verificato prima di fidarsi del risultato su un lotto grande. Il caso di prova è un
   dispositivo modellato **dentro lo spessore** di un muro: deve uscire dalla parte verso
   cui guarda. Se su una libreria il punto risultasse dietro, il sintomo è vistoso (tutti
   i dispositivi di quella famiglia finiscono nella stanza sbagliata) e la colonna
   *Fronte* del resoconto dice subito quale criterio ha deciso.
7. **`GetSpatialElementCalculationPoint()` restituisce il punto in coordinate del
   modello**, non in coordinate della famiglia. Il codice lo assume, ed è su questa
   assunzione che poggia la sottrazione con il punto di inserimento. Se fosse in
   coordinate locali il vettore risultante sarebbe privo di senso, e il sintomo sarebbe
   di nuovo dispositivi portati dalla parte sbagliata: verificarlo su un singolo elemento
   in simulazione **prima** di applicare.
8. **I 5 gradi di `FRONT_MAX_ANGLE_DEG`** sono la costante più delicata dello strumento,
   perché **decide se un elemento viene spostato o saltato**, non solo con quale criterio.
   Era inizialmente derivata dal fattore della retta (60 gradi) ed è stata poi stretta a 5:
   un cono largo ammetteva scivolamenti laterali fino a 1,7 volte la distanza da coprire.

   Va confermata su un lotto reale, leggendo quanti elementi finiscono fra gli ignorati con
   il motivo *nessuna partizione fronteggiata*. Se fossero dispositivi che guardano
   chiaramente una partizione, è la costante da alzare; se comparissero invece dispositivi
   spostati con la parete di fianco, da abbassare.

   **Il rischio concreto non è il disallineamento del dispositivo, è il Room Calculation
   Point.** Il punto non è garantito stare esattamente davanti all'apparecchio: è dove
   l'autore della famiglia lo ha trascinato, e molti lo spostano in diagonale per farlo
   cadere comodamente dentro la stanza. Un RCP in diagonale produce un fronte che si
   discosta di decine di gradi dalla perpendicolare anche su un dispositivo montato
   perfettamente, e con un cono di 5 gradi quelle famiglie vengono saltate in blocco. Se
   nel resoconto una famiglia intera risulta *nessuna partizione fronteggiata*, guardare
   dov'è il pallino in quella famiglia prima di toccare la costante.
9. **I 2 mm di `ADJACENT_GAP_MM`.** Due istanze di muro accostate in Revit dovrebbero
   combaciare esattamente, ma fra un modello collegato e l'host non è garantito. Se in
   prova un pacchetto evidente non venisse riconosciuto, è la costante da alzare; se
   invece venisse attraversata un'intercapedine, da abbassare. Il caso di prova è una
   muratura con contromuro modellato come istanza separata.
10. **Il pacchetto murario attraversa i confini fra documenti.** Un muro dell'host e uno
    di un collegamento che si toccano vengono uniti nello stesso pacchetto, perché il
    calcolo lavora su scalari invarianti in coordinate host. È la scelta fatta, coerente
    con "vince la partizione più vicina da qualunque documento provenga", ma non è stata
    verificata su un caso reale di contromuro impiantistico modellato nell'host davanti
    all'architettonico collegato.
11. **`RevitLinkInstance.GetTotalTransform()` è la trasformazione giusta**, cioè quella che
   include l'eventuale spostamento da coordinate condivise, e non `GetTransform()`. È la
   scelta fatta, ma non è stata verificata su un modello con il collegamento posizionato
   per coordinate condivise: se fosse sbagliata, i muri collegati risulterebbero spostati
   in blocco e **tutti** gli elementi finirebbero fuori tolleranza, che è un sintomo
   vistoso e quindi diagnosticabile in fretta.
12. **Il bounding box di un muro collegato, letto con `get_BoundingBox(None)` dentro al
    documento collegato, è nelle coordinate di quel documento** e non già trasformato. Il
    codice lo assume, ed è su questa assunzione che poggia la conversione della quota
    (`z_offset`). Se fosse già in coordinate host, il filtro verticale sui collegamenti
    sbaglierebbe di quanto vale l'offset del link.
13. **`Transform.Scale` vale 1 per un collegamento normale.** Il controllo esiste come
    valvola di sicurezza, ma non è mai scattato su un caso reale.
14. **`RotateElement` su un elemento bloccato**: lancia un'eccezione o fallisce in silenzio?
    È documentato solo per `MoveElement`. Il pre-controllo su `Pinned` rende la domanda
    accademica, ma il pre-controllo non va rimosso.
15. **`LocationPoint.Point` è aggiornato subito dopo `RotateElement`**, o serve un
    `doc.Regenerate()`? La traslazione viene ricalcolata proprio da quel punto, quindi
    leggerlo stantio sposterebbe male. Se in prova risultasse non aggiornato, basta alzare
    la costante `FORCE_REGEN` in testa allo script. Il sintomo sarebbe uno scarto residuo
    pari allo spostamento dovuto alla rotazione: piccolo, e visibile solo sulle famiglie la
    cui origine non sta sull'asse di rotazione.

## Comportamenti noti e accettati

- Il punto di inserimento finisce **esattamente sulla faccia**. Se una famiglia ha
  l'origine al centro del proprio ingombro invece che sul retro, l'elemento risulta per
  metà dentro il muro. Le colonne *distanza prima* e *distanza dopo* del resoconto servono
  a rilevarlo sul primo modello reale. Se si rivelasse un problema diffuso, la versione
  successiva introdurrà un offset impostabile o l'allineamento del bordo dell'ingombro.
- **Un dispositivo montato davvero di traverso non viene raddrizzato: viene saltato prima**,
  perché nessuna partizione risulta fronteggiata. Il raddrizzamento e l'esclusione sono la
  stessa regola vista da due lati.
- **Lo spostamento può superare di molto la tolleranza**, per tre motivi che si sommano: la
  faccia opposta aggiunge lo spessore del muro, il pacchetto murario aggiunge quello delle
  partizioni attraversate, e la corsa obliqua moltiplica tutto per 1/cos (al massimo 2). È
  voluto: la tolleranza è il raggio entro cui cercare la partizione, non un limite allo
  spostamento. La colonna *Spostamento* e le note lo rendono visibile.
- **La posizione del dispositivo lungo il muro non è più invariante**, perché la corsa
  segue la retta di analisi e non la normale.
- **Fra i muri che il dispositivo guarda vince ancora il più vicino.** Il fronte decide
  chi è ammesso alla gara, non chi la vince: in un angolo con due muri quasi equidistanti
  ed entrambi fronteggiati vince il più vicino, ed è il motivo per cui la colonna *Note*
  segnala l'ambiguità.
- Un elemento la cui quota non è coperta da nessuna partizione vicina **compare nel
  resoconto** fra gli ignorati, con il motivo *ingombro fuori dall'estensione verticale
  dei muri*. Nella versione precedente, quando era lo strumento a pescare gli elementi,
  spariva in silenzio ed era corretto così; ora che lo hai indicato tu deve dirti perché
  non lo ha toccato.
- **Il filtro verticale toglie di mezzo le partizioni, non l'elemento.** Una partizione che
  alla quota del dispositivo non arriva viene esclusa dal confronto prima ancora di
  calcolare le distanze. Senza questo, un dispositivo vicino sia a un muro pieno sia a un
  muretto basso poteva essere allineato al muretto, che alla sua quota non esiste.
- Muri tenda, muri inclinati, muri senza linea di posizionamento e muri con estensione
  verticale non leggibile vengono scartati con il motivo nella tabella *Partizioni
  scartate*.
- Il **contenitore** di un muro sovrapposto (stacked) viene scartato, non espanso: i suoi
  membri sono elementi a sé stanti, già raccolti dalla ricerca sulla categoria Walls, e
  ognuno ha spessore ed estensione verticale propri. Scartare il contenitore evita di
  contare due volte la stessa parete. Nella versione a selezione di muri veniva invece
  espanso, perché lì i muri arrivavano dalla selezione dell'utente e non da un collector.
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
- Il volume di ricerca attorno a ogni elemento è dilatato di **2 m** oltre la tolleranza
  (`PARTITION_MARGIN_MM`), perché il filtro nativo confronta il *bounding box della
  partizione* con quel volume, mentre il criterio vero è la distanza del *punto di
  inserimento dalla faccia*. Il margine deve coprire lo spessore della partizione e il
  fatto che il bounding box di un muro lungo si estende ben oltre il tratto vicino
  all'elemento. Il margine **non allarga la tolleranza**, che viene verificata a valle in
  modo esatto.
- I `WallInfo` sono **in cache** per coppia (documento, id). Non è un dettaglio di
  prestazioni: il controllo di inclinazione può estrarre la geometria della partizione, e
  la stessa parete ricorre per tutti i dispositivi che le stanno davanti. Senza cache quel
  costo verrebbe moltiplicato per il numero di elementi selezionati.

### Sui modelli collegati

- I muri collegati sono **solo letti**: la fase di applicazione lavora esclusivamente sul
  documento host e non apre mai una transazione su un collegamento.
- Host e collegamenti **concorrono insieme**. Se un dispositivo ha davanti un muro dell'host
  a 5 cm e uno collegato a 3 cm, vince quello collegato. È voluto: la regola è la distanza,
  non la provenienza.
- Un collegamento ignorato non blocca il comando: gli altri vengono comunque usati, e il
  motivo compare fra gli avvisi. Se però l'architettonico è l'unico collegamento e viene
  ignorato, tutti gli elementi finiranno fra gli ignorati per *nessuna partizione
  verticale nel raggio di ricerca*: vale la pena leggere sempre la sezione **Avvisi**.
- Le partizioni collegate compaiono come `[Nome collegamento] Tipo di muro (id 123456)`,
  con l'id non cliccabile.

## Avvertenze operative

- Eseguire sempre prima in **Simulazione**, leggere il resoconto, e solo dopo applicare.
- L'operazione è annullabile con un solo Undo di Revit.
- Tag e quote agganciati agli elementi spostati li seguono, ma la loro posizione relativa
  può diventare illeggibile: non c'è nulla che lo strumento possa fare via API. Con la
  corsa lungo la retta di analisi il fenomeno è più marcato di prima, perché gli elementi
  si spostano anche lateralmente.

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
