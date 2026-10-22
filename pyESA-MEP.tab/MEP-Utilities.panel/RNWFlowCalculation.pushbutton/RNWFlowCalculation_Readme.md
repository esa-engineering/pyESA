# RNW Flow Calculation

Verifica e dimensionamento delle reti di acque meteoriche. Per ogni tratto della rete
legge la portata di progetto, la confronta con le tabelle del manuale tecnico e propone il
diametro minimo adeguato. I risultati si vedono in una finestra dove l'utente sceglie quali
tratti modificare; solo quelli vengono cambiati nel modello, dopo conferma.

## Portata

La portata viene dalla rete stessa. Le famiglie dei bocchettoni scrivono la propria
portata [l/s] nelle unità di carico (Fixture Units) del connettore, e Revit la somma lungo
la rete. Lo script legge quindi il parametro **Fixture Units** di ogni tubo
(`RBS_PIPE_FIXTURE_UNITS_PARAM`) come portata in l/s.

Come la famiglia calcola la portata (intensità pluviometrica, superficie, coefficienti) è
definito dall'utente in Revit e non interessa lo script.

## Tubi verificati

Nella prima finestra si sceglie da una tendina il **tipo di sistema di tubazioni** degli
scarichi delle acque meteoriche; ogni voce riporta il numero di tubi. Senza impostazioni
salvate la tendina propone il primo tipo il cui nome contiene *meteor*, *pluv*, *bianc*,
*storm*, *rain* o *RNW*, altrimenti il tipo con più tubi dotati di Fixture Units; poi vale
l'ultima scelta. Se nessun tubo del tipo scelto ha Fixture Units, la finestra lo segnala in
rosso.

Sotto la tendina, l'elenco dei **sistemi** di quel tipo (istanze, per esempio SS.RNW 1,
SS.RNW 2), ciascuno con il numero di tubi e di tubi con Fixture Units, con ricerca e
Select All / Deselect All: il calcolo riguarda solo i sistemi spuntati (almeno uno). La
scelta si ricorda per nome: riaprendo, restano spuntati i sistemi scelti l'ultima volta, se
ci sono ancora; altrimenti tutti. Cambiando tipo l'elenco si ricarica.

Si calcolano sempre i **sistemi interi** spuntati, in tutto il modello: non c'è più la
scelta fra modello, vista attiva e tubi selezionati. **Se lo strumento parte con una
selezione** (tubi, raccordi, accessori o terminali), `selection_systems` legge i sistemi
degli elementi selezionati (`RBS_SYSTEM_NAME_PARAM`, più sistemi separati da virgola) e
tiene quelli che sono sistemi dei tubi del modello. La finestra seleziona il tipo con più
sistemi selezionati e spunta i suoi sistemi, al posto delle scelte salvate; una riga
azzurra lo dice, e cita i sistemi selezionati di un altro tipo, che non si calcolano.
L'utente può comunque cambiare la scelta. Se la selezione non tocca nessun sistema, la riga
diventa rossa e valgono le scelte salvate. Senza selezione la finestra parte dalle scelte
salvate, come prima. I sistemi di una partenza con selezione non diventano le scelte salvate
(`RunConfig.from_selection`): la partenza senza selezione resta quella dell'ultima scelta
manuale.

**La serie di tubo non si sceglie**: si legge per ogni tubo dal modello (`pipe_series`):
dal nome del materiale del suo segmento, se non basta dal nome del segmento, poi dal nome
del tipo di tubazione (`RNWFlowCalculation_Tables.series_from_text`: polietilene / PE / HDPE / PEAD,
polipropilene / PP, Triplus, Silere). La serie dà la colonna della tabella 6.2; se non è
riconosciuta si usa il valore più cautelativo fra le colonne che hanno quel De, con una
nota sulla riga. Il riepilogo delle serie riconosciute per segmento è nel testo in cima alla
finestra dei risultati e nel report.

## Verifica del singolo tubo

- **Verticali (colonne)**: tubi inclinati di più di 60° sull'orizzontale
  (`STACK_MIN_ANGLE_DEG`). Tabella 6.2, f = 0,33, colonna della serie del tubo (PE, PP,
  Triplus, Silere), letta dal segmento. I tratti a 45° sono **orizzontali**: si verificano con le tabelle 6.3 -
  6.5, alla riga 10 cm/m (la loro pendenza supera la tabella). La nota "Slope over 10 cm/m"
  compare solo per i tubi fra 10 cm/m e 30° (`RUN_45_MIN_ANGLE_DEG`), non per i tratti a 45°,
  dove sarebbe sempre presente.
- **Collettori orizzontali**: pendenza dalla geometria del tubo, in cm/m. Tabelle 6.3 / 6.4 /
  6.5 secondo il grado di riempimento scelto (0,5 / 0,6 / 0,7). Pendenza fra due righe: riga
  con pendenza inferiore. Sotto 0,5 cm/m il tratto non è verificabile; sopra 10 cm/m si usa
  la riga 10 cm/m.
- Il De del modello è il diametro esterno (`RBS_PIPE_OUTER_DIAMETER`) riconosciuto entro
  ±1 mm da un De di tabella.
- Il **diametro suggerito** è il minimo De fra le misure del tipo di tubazione del tubo
  (preferenze di instradamento, vedi "Misure disponibili") con portata di tabella ≥ Q (e,
  se è impostata una velocità massima, che la rispetta). Per le colonne conta anche che il
  De abbia un valore nella colonna della serie.

### Velocità e grado di riempimento

La v delle tabelle 6.3 - 6.5 vale al grado di riempimento della tabella, cioè alla portata
Q tabellare, non alla portata reale del tubo: un tubo che porta meno della sua capacità
lavora più vuoto e più lento. Lo script calcola quindi velocità e riempimento **alla
portata di progetto** (`RNWFlowCalculation_Tables.partial_state`):

1. le tabelle seguono la relazione di moto uniforme a riempimento parziale di un tubo
   circolare (Gauckler-Strickler / Manning): Q(f=0,7)/Q(f=0,5) vale 1,674 in teoria e 1,67
   in tabella per De 110 e 160;
2. per De e pendenza si ricavano dalle tre tabelle portata e velocità a tubo pieno (media
   delle tre stime, per smorzare l'arrotondamento a un decimale delle v);
3. alla portata reale si risolve il grado di riempimento y = h/D e la velocità
   v = v_pieno · (R(y)/R_pieno)^(2/3).

Alla portata di tabella il modello ritrova f entro 0,01 e v entro 0,06 m/s (arrotondamento
delle tabelle). Esempio: De 110 all'1 % con 1 l/s lavora al 31 % a 0,52 m/s, contro i
0,7 m/s letti in tabella. Un tubo con portata oltre il pelo libero (riempimento 100 %) non
ha velocità.

Nella finestra delle impostazioni si danno velocità minima (predefinita 0,6 m/s) e massima
(vuota = nessuna verifica), da adeguare alle prescrizioni del progetto. Stati
`LOW VELOCITY` / `HIGH VELOCITY` se il tubo, verificato per portata, è fuori dai limiti;
la nota dice se nessun De risolve (serve cambiare pendenza). Le colonne non hanno verifica
di velocità.

## Tratti

Una Section è la coppia **nome del sistema + parametro `Section`** del tubo. Il parametro
si cerca per nome (`Section`, vale anche per un parametro condiviso) e poi come parametro
di sistema `RBS_SECTION` (in un Revit italiano si chiama *Sezione*). I numeri di Section
ripartono in ogni sistema, per questo la chiave comprende il sistema.

Ogni Section si divide in **due righe**: la parte orizzontale (tubi orizzontali e a 45°) e
la parte verticale, verificate e dimensionate separatamente. Chiave della riga:
(sistema, Section, orientamento); `group_key` riunisce le righe della stessa Section. Un
tubo senza Section forma una riga da solo, con una nota.

Aggregazione di una riga in `RNWFlowCalculation_Check.py`:

- portata del tratto = la massima dei suoi tubi (se differiscono c'è una nota);
- De suggerito = il massimo fra quelli dei suoi tubi; nessuno se un tubo non è
  verificabile (pendenza sotto tabella, portata oltre tabella);
- stato = il più grave fra i tubi per gli errori; `OK` / `OVERSIZED` si decidono sul tratto,
  confrontando il De suggerito con il De del modello;
- velocità = la minima fra i tubi orizzontali, riempimento = il massimo;
- misure disponibili = quelle comuni ai tipi di tubazione dei suoi tubi (vedi sotto).

Stati: `UNDERSIZED`, `FLOW BEYOND TABLE`, `DE NOT IN TABLE`, `SLOPE UNDER 0.5 cm/m`,
`HIGH VELOCITY`, `LOW VELOCITY`, `OVERSIZED` (verificato ma più grande del necessario),
`OK`, `NO FLOW` (Fixture Units nullo).

### Misure disponibili

Le misure della tendina vengono dalle **preferenze di instradamento del tipo di
tubazione** (`PipeType.RoutingPreferenceManager`, regole *Segments*): ogni segmento
contribuisce le misure nel proprio campo di diametro nominale (`PrimarySizeCriterion`),
in ordine di priorità, e solo quelle abilitate negli elenchi delle dimensioni
(`MEPSize.UsedInSizeLists`). Se le preferenze non sono leggibili si usa il segmento attuale
del tubo. Ogni misura porta De (diametro esterno) e DN (diametro nominale).

## Finestra dei risultati

**Terminali e ordine della rete.** La finestra elenca anche i terminali (plumbing fixtures,
per esempio i bocchettoni) collegati ai tubi verificati, come righe verdi in corsivo
"Fixture" con la portata di ognuno (Fixture Units del connettore, `fixture_flow`),
l'attacco (DN del connettore), il codice elemento e la famiglia; non hanno misura nuova, il
pulsante 🔍 li seleziona. I loro parametri si modificano nella scheda **Fixtures**.

Le righe seguono la rete **dal punto di monte al punto di valle** (colonna `#`):

1. `network_topology` visita la rete dai tubi verificati attraverso i raccordi (fermandosi
   ai tubi dei sistemi non calcolati) e raccoglie i **nodi**, cioè i tubi che si incontrano in un
   raccordo o in una catena di raccordi (due in una curva, tre in una braga), e i
   terminali raggiunti;
2. in ogni nodo il tratto di valle è quello più carico (la portata cresce verso valle), a
   parità di portata il più basso; gli altri vi confluiscono, quindi i rami di una braga
   restano fratelli e non uno a monte dell'altro. I terminali sono sempre a monte;
3. `RNWFlowCalculation_Check.order_network` parte da ogni punto di valle (il più carico per primo) e
   risale i rami scrivendo prima ciò che sta a monte: ogni ramo compare dal suo terminale in
   giù, prima del tratto in cui confluisce, e lo sbocco è l'ultima riga.

La colonna **Fed by** dice da quali righe deriva ciascuna riga: le righe che vi confluiscono
direttamente, in forma breve: "Fixture <id>" per un terminale, numero di Section seguito
da V (parte verticale) o O (parte orizzontale) per i tratti, uniti da "+" quando sono più
d'uno, per esempio "1V+2O"; "-" se a monte non c'è nulla. Il
dato è il grafo da monte a valle di `order_network` (attributi `upstream` / `downstream`
di ogni riga) ed entra anche nella ricerca.

Con il filtro "Only sections that need attention" i terminali restano visibili, perché
danno il punto di partenza dei rami. Cliccando un'intestazione la griglia si riordina per
quella colonna; cliccando `#` si torna all'ordine della rete.


`RNWFlowCalculation_Results_ui.py` + `RNWFlowCalculation_Results.xaml`: tre schede, **General**,
**Sections** e **Fixtures**, griglie legate a `DataTable` come in MEPQTO. La finestra si
apre su Sections; dopo la navigazione o il ricalcolo torna sulla scheda da cui si era
partiti.

### Scheda General

Riepilogo del calcolo (`general_summary()`, Python puro, ricostruito a ogni apertura):

- **Total covered area** [m²]: somma delle superfici di raccolta (`u_PLU_Area_1`, istanza
  poi tipo) dei terminali collegati;
- **Rainfall intensity**, in l/(s·m²) e in mm/h (× 3600). L'indice pluviometrico lo usa la
  famiglia del bocchettone e lo script non lo legge: lo **ricava** da ogni terminale come
  Q / (A × C1 × C2), cioè supponendo che la famiglia calcoli Q = i × A × C1 × C2. Si usano
  i terminali con Q, A, C1 e C2 tutti maggiori di zero; se i valori differiscono di più
  dello 0,5 % (`INTENSITY_TOLERANCE`) si mostra l'intervallo e una nota invita a
  controllare le famiglie;
- **Total collected flow** [l/s]: somma delle portate dei terminali (Fixture Units del
  connettore);
- tabella dei **sistemi** calcolati (anche quelli senza terminali), con numero di
  terminali, superficie e portata.

Le note sotto i valori contano i terminali senza superficie o senza portata leggibile,
esclusi dai totali. Tutti i valori sono quelli del modello all'apertura della finestra:
le modifiche ai terminali non ancora scritte contano solo dopo Apply and Recalculate.

### Scheda Sections

Una riga per parte di Section (colonna *Run*: Horizontal / Vertical). Prima colonna **Show**
(🔍), poi `#`, System, **Element code** (`e_MEC_ElementCode_1` dei tubi della riga, istanza
poi tipo; valori diversi uniti da "/"), Section, Pipes, **Pipe Type** (Type Mark del tipo
di tubazione, `ALL_MODEL_TYPE_MARK`; più tipi nella stessa riga uniti da " - ", un tipo senza
Type Mark compare come "(no type mark)") e le altre. Pipe Type entra anche nella ricerca.
**Status** (stato attuale) sta subito prima di **New size**.

Le celle sono **su una riga sola**: le colonne di testo (System, Element code, Pipe Type,
Fed by, Status, Check new, Notes) hanno larghezza automatica e la griglia scorre in
orizzontale; le colonne numeriche restano strette, con l'intestazione su più righe. Con la
virtualizzazione delle righe una colonna automatica può allargarsi scorrendo verso righe
con testi più lunghi. Sotto la griglia, una **legenda** spiega i colori di righe, celle
della misura nuova e celle Status / Check new.

- Colonne del modello: DN, De, portata ammissibile, utilizzo, velocità e riempimento alla
  portata di progetto, De suggerito.
- **New size**: tendina con le misure del tipo di tubazione ("De 160 / DN 150"). A ogni
  scelta la riga ricalcola Qmax, v, riempimento ed esito della misura nuova
  (`SectionResult.evaluate_at`).
- **Non ci sono spunte**: Apply Changes modifica tutte le righe la cui misura nuova è
  diversa da quella del modello (colonna nascosta `Changed`, `is_change()`; per un tratto
  con tubi di misure diverse qualunque scelta è un cambio, perché li uniforma). Queste
  righe sono **azzurre e in grassetto**, con le celle della misura nuova in azzurro più
  scuro. Per questo la tendina parte dalla misura suggerita solo nei tratti da cambiare
  (`default_apply`: `UNDERSIZED`, `DE NOT IN TABLE`, `HIGH VELOCITY` con De suggerito fra le
  misure del tipo), altrimenti dalla misura del modello: un tratto `OVERSIZED` non si
  riduce se l'utente non lo sceglie. Si annulla un cambio riportando la tendina sulla misura
  del modello.
- **Sim. slope [%]** (dopo New size, solo per le righe orizzontali): pendenza simulata, una
  per ogni riga delle tabelle 6.3 - 6.5: da 0,5 a 5 cm/m ogni 0,5, poi da 6 a 10 cm/m ogni 1
  (`RNWFlowCalculation_Check.SIMULATED_SLOPES` = `RNWFlowCalculation_Tables.HORIZONTAL_SLOPES`); "Model" = pendenza del modello. Serve quando la portata non sta in tabella alla
  pendenza del modello (pendenza sotto 0,5 cm/m, portata oltre la tabella). Tutti i tubi
  orizzontali della riga si verificano a quella riga di tabella
  (`SectionResult.evaluate_at(de, slope)`), e De suggerito, Qmax, v, riempimento ed esito
  della misura nuova seguono la simulazione (`suggested_at`); le colonne del modello no.
  Una riga simulata è **gialla** e **non si modifica** (colonna `Simulated`, esclusa da
  `_collect_changes`): cambiare la pendenza nel modello sposterebbe i tubi e romperebbe
  quasi certamente i collegamenti, quindi la si cambia a mano e si rilancia lo strumento.
  La conferma di Apply Changes dice quante righe simulate restano escluse. La pendenza
  simulata sopravvive alla navigazione e ad Apply and Recalculate (`preset_slope`).
- Il colore dell'esito sta solo sulle celle **Status** (stato attuale) e **Check new**
  (misura nuova): rosso per `UNDERSIZED` / `FLOW BEYOND TABLE`, giallo per i tratti non
  verificabili, testo blu per `OVERSIZED`; testo grigio per le righe senza portata. La riga
  resta libera per il colore del cambio di misura.
- **Show** (🔍) seleziona in Revit gli elementi della riga e li inquadra. La finestra si
  chiude e lo script resta in una selezione (`PickObjects`, con gli elementi già selezionati
  e cliccabili solo quelli): si naviga la vista attiva (pan, zoom, orbita); con **Fine** o
  **Esc** la finestra si riapre con lo stesso `ResultsState` (stesse `DataTable`: misure,
  modifiche ai terminali, ricerca, filtro, scheda) e torna sulla riga di partenza. La
  finestra resta modale di proposito: una finestra non modale richiederebbe `ExternalEvent`
  e lascerebbe modificare il modello mentre la tabella mostra dati superati; durante la
  selezione invece il modello non si può modificare.
- Filtro "Only sections that need attention" (salvato; restano visibili i terminali e le
  righe che cambiano misura) e ricerca su sistema, codice elemento, Section, stato, note e
  Fed by.

### Scheda Fixtures

I terminali collegati ai tratti calcolati, nell'ordine della rete, con 🔍, portata e i
valori modificabili (celle gialle).

Prima dell'Id, la colonna **Level, offset [mm]** riporta il livello e l'offset impostato
nel formato `Livello, +150` (`fixture_level()` nello script): livello da `LevelId`, poi
dai parametri del livello di abaco / di famiglia; offset da *Elevation from Level*
(`INSTANCE_ELEVATION_PARAM`), poi *Offset from Host* (`INSTANCE_FREE_HOST_OFFSET_PARAM`,
per le famiglie ospitate è quindi la distanza dalla faccia ospite), altrimenti la distanza
del punto di inserimento dalla quota del livello (`ProjectElevation`). La colonna si ordina
per quota assoluta. Come Notes e System, ha larghezza automatica: la griglia scorre in
orizzontale invece di stringere le altre colonne.

| Colonna | Parametro |
| --- | --- |
| Family and type | tipo del terminale (`ChangeTypeId`), tendina con tutti i tipi di plumbing fixture del modello |
| Collection area [m²] | `u_PLU_Area_1` |
| Runoff coeff. C1 | `C1` |
| Risk coeff. C2 | `C2` |
| Element code | `e_MEC_ElementCode_1` |

Ogni valore si cerca prima sull'istanza, poi sul tipo; la colonna **from** dice dove sta
(`instance`, `type`, `missing`, `read-only ...`). Le celle di un parametro mancante o in sola
lettura (per esempio guidato da una formula) sono grigie e non si modificano.

- Lo script legge i valori e li passa già ridotti a testo (`read_fixture_data` ->
  `rres.FixtureData`: valori di istanza per terminale, valori di tipo per tipo). I numeri
  sono nelle unità di visualizzazione del parametro; un parametro di tipo Area si mostra e
  si scrive sempre in m² (`_display_unit`). Si scrive con virgola o punto (`parse_decimal`);
  un valore non numerico o negativo viene rifiutato e la cella torna com'era.
- **Un parametro di tipo cambia tutte le istanze di quel tipo nel modello**, come in Revit:
  modificandolo su una riga cambiano anche le altre righe dello stesso tipo
  (`ResultsState.pending_type`). Cambiando tipo, la riga mostra i valori di tipo del tipo
  nuovo e conserva i valori di istanza già modificati.
- Le righe con modifiche da scrivere sono azzurre (confronto numerico: "0,90" e "0.9" non
  sono una modifica). Il codice elemento modificato si vede anche nella scheda Sections e
  nell'export.
- **Apply and Recalculate** scrive le modifiche (`write_fixture_edits`, una `Transaction`
  "Edit rainwater fixtures": prima i cambi di tipo, poi i parametri di tipo, poi quelli di
  istanza), lascia a Revit l'aggiornamento delle portate, rilegge la rete (`analyze`) e
  riapre la finestra con i risultati nuovi. Restano le misure scelte a mano
  (`ResultsState.carry`, per chiave della Section: le proposte automatiche si ricalcolano
  con le portate nuove), la ricerca, il filtro e la scheda. I tubi non vengono ancora
  modificati.
- Chiudere la finestra con modifiche ai terminali non scritte chiede conferma.

### Applicazione

**Apply Changes** chiede conferma e chiude la finestra. La conferma riassume tubi da
ridimensionare e modifiche ai terminali, e segnala:

- che con modifiche ai terminali in sospeso le misure sono calcolate con le portate attuali
  (per verificarle con i valori nuovi si usa prima Apply and Recalculate);
- i tratti che con la misura scelta non superano ancora la verifica;
- le Section di cui cambia una parte sola, o le cui parti ricevono misure diverse: la curva
  fra la parte orizzontale e la verticale ha una sola misura e farebbe annullare la modifica.

Lo script scrive prima le modifiche ai terminali, poi ridimensiona i tubi. Non è il pulsante
predefinito: Invio nella griglia non modifica il modello. **Close** chiude senza modificare
nulla.

### Esportazione

**Export** nella finestra dei risultati apre la finestra **Export Calculation**
(`RNWFlowCalculation_Export_ui.py` + `RNWFlowCalculation_Export.xaml`), su tre colonne:

- **Format**: Excel (.xlsx) o PDF (.pdf); per il PDF il foglio, A4 o A3 orizzontale;
- **General**: colonne della tabella dei sistemi (riepilogo e impostazioni si esportano
  sempre);
- **Sections** e **Fixtures**: una casella per colonna, con *Default*, *Select All* e
  *Deselect All*. Ogni tabella deve avere almeno una colonna.

Le scelte si salvano in `ESA_RNWFlowCalculation` (`export_format`, `export_paper`,
`export_general_columns`, `export_section_columns`, `export_fixture_columns`: chiavi
separate da virgole; una chiave che non esiste più si ignora). Le colonne predefinite di
Sections sono quelle del **risultato del calcolo**, come nell'export Excel delle versioni
precedenti: #, System, Element code, Section, Pipes, Pipe Type, Run, Fed by, Min slope, Q,
New DN / New De (la tendina New size), Sim. slope, Qmax / Use / v / Fill con quella misura,
Check new, Notes. Le colonne dello stato attuale (Model DN / De, Qmax, Use, v e Fill del
modello, Suggested De, Status) si aggiungono a mano. Fixtures: tutte tranne le colonne
*from* e *Changes not written*.

Il contenuto è lo stesso nei due formati (`RNWFlowCalculation_Report.py`, Python puro:
cataloghi delle colonne, schede come blocchi di coppie, note e tabelle):

- **General**: superficie, indice pluviometrico, portata, note (compreso il numero di
  modifiche ai terminali non ancora scritte), tabella dei sistemi;
- **Sections**: tutte le righe, tratti e terminali, nell'ordine della rete (filtro e
  ricerca della griglia non contano). Status e Check new hanno i colori della finestra
  (rosso, giallo, testo blu per OVERSIZED, verde per OK), i terminali sono in corsivo su
  verde chiaro, le righe senza portata in grigio (solo PDF);
- **Fixtures**: i valori della griglia, **modifiche non ancora scritte comprese** (colonna
  *Changes not written*);
- impostazioni del calcolo (`settings_rows`): foglio **Settings** in Excel, in coda alla
  scheda General nel PDF.

Excel (`RNWFlowCalculation_Xlsx.py`, writer zip + XML senza Excel installato, parte
generica copiata da `MEPQTO_xlsx.py`): un foglio per scheda più Settings; nei fogli
Sections e Fixtures intestazione bloccata e filtro automatico; i numeri restano numeri.

PDF (`RNWFlowCalculation_Pdf.py`): scritto a mano, senza librerie né stampanti. PDF 1.4
con i font standard Helvetica (non incorporati) e codifica WinAnsi: i caratteri fuori da
cp1252 diventano un sostituto (`<=`, `->`) o `?`, le emoji si tolgono. Ogni scheda parte su
una pagina nuova; intestazione con nome dello strumento e scheda, piede con nome dello
strumento e numero di pagina (né modello né data); le tabelle ripetono l'intestazione a
ogni pagina. Le colonne prendono la larghezza del contenuto; se non ci stanno il carattere
scende da 7 a 5,5 pt, poi le colonne di testo si stringono fino alla parola più lunga e
vanno a capo.

Il nome proposto è `<Modello>_RainwaterCalculation_<aaaammgg>.xlsx` (o `.pdf`) nella
cartella del modello. Un file aperto in un altro programma non si sovrascrive: la finestra
lo segnala. Dopo l'esportazione il file si può aprire con il programma predefinito.

## Modifica del modello

Per ogni tubo dei tratti scelti si imposta il diametro nominale (`RBS_PIPE_DIAMETER_PARAM`)
della misura scelta; i tubi già a quella misura non si toccano. Revit sceglie il segmento
dalle preferenze di instradamento. Dopo `Regenerate()` lo script controlla che ogni tubo
abbia davvero il De scelto: se Revit ha assegnato un'altra misura, il tratto viene
annullato con un messaggio.

### Raccordi e accessori

**I tubi hanno sempre la misura del calcolo**: per far tornare un raccordo non si cambia
mai la misura di un tubo, non si sposta nessun tubo e non si cancella nessun raccordo
esistente. Lo script si comporta come Revit quando si cambia a mano la misura di un
raccordo: il raccordo prende la misura dei tubi modificati e, dove resta collegato a un tubo
che non cambia, Revit mette una riduzione.

Per ogni raccordo (`OST_PipeFitting`) o accessorio (`OST_PipeAccessory`) collegato a un
tubo modificato (`resize_fittings`), ogni tentativo in una `SubTransaction`:

1. se i raggi verso gli elementi già alla misura giusta coincidono (±0,5 mm), nulla;
2. **misura del connettore** impostata direttamente (`Connector.Radius`), come il campo
   "Diameter" della barra delle opzioni;
3. **ricollegamento** del connettore: collegando un tubo, Revit adegua il raccordo;
4. **parametro associato** alla misura del connettore (per esempio NR nelle famiglie
   e_PP.GG), trovato con `GetAssociateFamilyParameterId` o leggendo la famiglia
   (`scan_fitting_families`, prima delle transazioni), risalendo le formule semplici
   (`X`, `X / k`, `X * k`, `k * X`); un parametro di tipo si soddisfa cambiando tipo;
5. in 2, 3 e 4, se il raccordo cambiato non combacia più con un tubo che resta com'è (curva
   fra la parte modificata e una che non cambia, linea principale di una braga con un solo
   lato modificato), fra i due si inserisce una **riduzione** con `NewTransitionFitting` e
   la famiglia delle preferenze di instradamento (`_bridge`). `NewTransitionFitting` vuole
   come primo connettore quello del tubo ("The owner should be (flex) duct or pipe"); verso
   un altro raccordo una riduzione non si può aggiungere. La prova sul modello ha mostrato
   che, cambiando NR tramite API, Revit **non** inserisce la riduzione da solo (lo fa
   dall'interfaccia): la inserisce lo script; se Revit l'avesse già inserita, lo script la
   riconoscerebbe;
6. se il raccordo non accetta la misura nuova (per esempio la riduzione eccentrica bloccata
   verso un bocchettone, che dovrebbe allargare invece di restringere), resta com'è e si
   inserisce una riduzione fra il raccordo e il tubo modificato (`_try_bridge_wrong`);
7. **catena**: un raccordo attaccato direttamente a un raccordo appena cambiato entra a sua
   volta in coda.

Se nulla riesce, la Section viene annullata e il messaggio elenca i tentativi con l'errore
completo di Revit e i connettori del raccordo (misura, elemento collegato).

**Prova (SHIFT + clic).** Si sceglie un raccordo nella vista, un suo connettore e un DN.
Lo script prova, in una transazione che poi annulla (`run_fitting_test`): la misura del
connettore, il parametro associato, e il parametro associato seguito dalla riduzione verso
i tubi che restano com'erano (il percorso usato durante la modifica). Il report mostra per
ogni tentativo l'esito e cosa è collegato a ogni connettore dopo. Il modello non viene
modificato.

I connettori sono elencati e scelti per Id (`piping_connectors` li ordina così): Revit li
restituisce in ordine variabile, quindi "connettore 1" per posizione nell'elenco non sarebbe
sempre lo stesso.

Verifica finale (`verify_fittings`): ogni raccordo toccato deve avere lo stesso raggio di
**tutti** gli elementi a cui è collegato, anche quelli non modificati. Il report riporta per
riga quanti raccordi sono stati controllati e ridimensionati, e come (misura del
connettore o parametro associato, con il valore).

Non si toccano gli apparecchi sanitari (i bocchettoni): il loro attacco ha una misura
propria.

Una `Transaction` per **Section** (le righe orizzontale e verticale che cambiano insieme, così la
curva fra le due vede subito entrambe le misure) dentro un `TransactionGroup` ("Resize
rainwater sections", un solo passo di annulla). Il preprocessor dei failure è dedicato, come
in MEPAlignToWall, e non è il `FailureSwallower` di pyRevit, che può sbloccare vincoli o
cancellare elementi. Elimina solo i warning senza risoluzioni, registra tutti i messaggi e
chiede il rollback sugli errori. Una Section che non si riesce a ridimensionare (tubi o
raccordi) torna com'era e non annulla le altre.

### Report

Il report di pyRevit si apre **solo in uscita e solo se una Section o una modifica ai
terminali non si è potuta applicare**: avviso in testa con il numero di parti non
modificate, tabella dei parametri dei terminali non scritti (parametro mancante o in sola
lettura, valore rifiutato da Revit, transazione annullata), impostazioni, poi la tabella di
tutte le modifiche dei tubi con esito e messaggi. Gli errori di scrittura durante Apply and
Recalculate compaiono anche subito, in un avviso. Se tutto riesce, o se la finestra
si chiude senza applicare, non compare nessun report. I casi senza nulla da verificare
(nessun tubo nel modello, nessun tubo nei sistemi scelti) danno un avviso a finestra.

Dopo la modifica Revit ricalcola Fixture Units e Section: per verificare la rete con i nuovi
diametri si rilancia lo strumento.

## Moduli

| File | Ruolo |
| --- | --- |
| `RNWFlowCalculation_script.py` | lettura dei tubi, modifica del modello, report |
| `RNWFlowCalculation_Check.py` | verifica del tubo e aggregazione per tratto (Python puro) |
| `RNWFlowCalculation_Tables.py` | tabelle 6.2 - 6.5 e regole di lettura (Python puro) |
| `RNWFlowCalculation_Form_ui.py` + `RNWFlowCalculation_Form.xaml` | finestra delle impostazioni |
| `RNWFlowCalculation_Results_ui.py` + `RNWFlowCalculation_Results.xaml` | finestra dei risultati |
| `RNWFlowCalculation_Xlsx.py` | writer .xlsx del report (Python + .NET, nessun Revit) |
| `RNWFlowCalculation_Pdf.py` | writer PDF del report (Python + .NET, nessun Revit) |
| `RNWFlowCalculation_Report.py` | contenuto dell'esportazione e cataloghi delle colonne (Python puro) |
| `RNWFlowCalculation_Export_ui.py` + `RNWFlowCalculation_Export.xaml` | finestra di impostazione dell'esportazione |

Le impostazioni si salvano nella sezione `ESA_RNWFlowCalculation` della config di pyRevit.

## Limiti

- Lo script si fida del valore calcolato da Revit: se la rete è scollegata o il sistema non
  è configurato perché Revit sommi le Fixture Units, i tubi compaiono come `NO FLOW`.
- Non è controllato che il diametro non diminuisca nel verso del flusso.
- Il pulsante 🔍 lavora mentre la finestra è aperta (modale): la selezione e lo zoom sono da
  verificare in Revit.
- Apply and Recalculate si affida a Revit per aggiornare le Fixture Units dopo la scrittura
  dei parametri: se le impostazioni di calcolo MEP del progetto non ricalcolano i sistemi,
  le portate restano quelle di prima. Da verificare sul modello.
- Cambiare la famiglia di un terminale (non solo il tipo) può staccarlo dal tubo, se
  l'attacco della famiglia nuova non coincide: la rete si rilegge dopo il ricalcolo e un
  terminale staccato non compare più.
- La velocità è stimata dalla relazione di riempimento parziale tarata sulle tabelle, non
  dalla scabrezza o dal diametro interno del tubo del modello.
