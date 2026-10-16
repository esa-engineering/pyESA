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
rosso. L'ambito può essere l'intero modello, la vista attiva oppure i tubi selezionati.

## Verifica del singolo tubo

- **Verticali (colonne)**: tubi inclinati di più di 60° sull'orizzontale
  (`STACK_MIN_ANGLE_DEG`). Tabella 6.2, f = 0,33, colonna della serie scelta (PE, PP,
  Triplus, Silere). I tratti a 45° sono **orizzontali**: si verificano con le tabelle 6.3 -
  6.5, alla riga 10 cm/m (la loro pendenza supera la tabella), con una nota.
- **Collettori orizzontali**: pendenza dalla geometria del tubo, in cm/m. Tabelle 6.3 / 6.4 /
  6.5 secondo il grado di riempimento scelto (0,5 / 0,6 / 0,7). Pendenza fra due righe: riga
  con pendenza inferiore. Sotto 0,5 cm/m il tratto non è verificabile; sopra 10 cm/m si usa
  la riga 10 cm/m.
- Il De del modello è il diametro esterno (`RBS_PIPE_OUTER_DIAMETER`) riconosciuto entro
  ±1 mm da un De di tabella.
- Il **diametro suggerito** è il minimo De della serie con portata di tabella ≥ Q (e, se è
  impostata una velocità massima, che la rispetta). La serie limita anche i diametri
  suggeriti per gli orizzontali (per esempio PP arriva a 160, Silere usa 135 al posto di
  125).

### Velocità e grado di riempimento

La v delle tabelle 6.3 - 6.5 vale al grado di riempimento della tabella, cioè alla portata
Q tabellare, non alla portata reale del tubo: un tubo che porta meno della sua capacità
lavora più vuoto e più lento. Lo script calcola quindi velocità e riempimento **alla
portata di progetto** (`rnwflow_tables.partial_state`):

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

Aggregazione di una riga in `rnwflow_check.py`:

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

`rnwflow_results_ui.py` + `RNWFlowCalculation_results.xaml`: una riga per parte di Section
(colonna *Run*: Horizontal / Vertical), legata a una `DataTable` come in MEPQTO. Righe rosse per `UNDERSIZED` / `FLOW BEYOND TABLE`, gialle
per i tratti non verificabili, blu per `OVERSIZED`, grigie senza portata.

- Colonne del modello: DN, De, portata ammissibile, utilizzo, velocità e riempimento alla
  portata di progetto, De suggerito.
- **New size**: tendina con le misure del tipo di tubazione ("De 160 / DN 150"), proposta
  sulla misura suggerita. A ogni scelta la riga ricalcola, con sfondo azzurro, Qmax, v,
  riempimento ed esito della misura nuova (`SectionResult.evaluate_at`). Scegliere una
  misura diversa da quella del modello spunta il tratto.
- **Apply**: spunta il tratto da modificare. All'apertura sono spuntati i tratti
  `UNDERSIZED`, `DE NOT IN TABLE` e `HIGH VELOCITY` il cui De suggerito è una misura del
  tipo di tubazione.
- **Tick Sections to Resize** ripete la spunta automatica (e riporta la tendina sulla
  misura suggerita), **Untick All** toglie tutto.
- **Show** (🔍, una colonna per riga) e **Select in Model** (righe evidenziate)
  selezionano i tubi in Revit e li inquadrano. La finestra si chiude e lo script resta in
  una selezione (`PickObjects`, con i tubi già selezionati e cliccabili solo quelli):
  si naviga la vista attiva (pan, zoom, orbita); con **Fine** o **Esc** la finestra si
  riapre con lo stesso `ResultsState` (stessa `DataTable`: spunte, misure, ricerca,
  filtro) e torna sulla riga di partenza. I tubi restano selezionati. La finestra resta
  modale di proposito: una finestra non modale richiederebbe `ExternalEvent` e
  lascerebbe modificare il modello mentre la tabella mostra dati superati; durante la
  selezione invece il modello non si può modificare.
- Filtro "Only sections that need attention" (salvato) e ricerca su sistema, Section,
  stato e note.
- **Apply Changes** controlla tutti i tratti spuntati, chiede conferma e chiude la
  finestra. La conferma elenca i tratti che con la misura scelta non superano ancora la
  verifica e le Section di cui è spuntata una parte sola, o le cui parti ricevono misure
  diverse: la curva fra la parte orizzontale e la verticale ha una sola misura e farebbe
  annullare la modifica. Non è il
  pulsante predefinito: Invio nella griglia non modifica il modello. **Close** chiude senza
  modificare nulla.

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

Una `Transaction` per **Section** (le righe orizzontale e verticale spuntate insieme, così la
curva fra le due vede subito entrambe le misure) dentro un `TransactionGroup` ("Resize
rainwater sections", un solo passo di annulla). Il preprocessor dei failure è dedicato, come
in MEPAlignToWall, e non è il `FailureSwallower` di pyRevit, che può sbloccare vincoli o
cancellare elementi. Elimina solo i warning senza risoluzioni, registra tutti i messaggi e
chiede il rollback sugli errori. Una Section che non si riesce a ridimensionare (tubi o
raccordi) torna com'era e non annulla le altre.

### Report

Il report di pyRevit si apre **solo in uscita e solo se una Section non si è potuta
modificare**: avviso in testa con il numero di parti non modificate, impostazioni, poi
la tabella di tutte le modifiche con esito e messaggi. Se tutto riesce, o se la finestra
si chiude senza applicare, non compare nessun report. I casi senza nulla da verificare
(nessun tubo, nessun tubo del tipo scelto nell'ambito) danno un avviso a finestra.

Dopo la modifica Revit ricalcola Fixture Units e Section: per verificare la rete con i nuovi
diametri si rilancia lo strumento.

## Moduli

| File | Ruolo |
| --- | --- |
| `RNWFlowCalculation_script.py` | lettura dei tubi, modifica del modello, report |
| `rnwflow_check.py` | verifica del tubo e aggregazione per tratto (Python puro) |
| `rnwflow_tables.py` | tabelle 6.2 - 6.5 e regole di lettura (Python puro) |
| `rnwflow_ui.py` + `RNWFlowCalculation_form.xaml` | finestra delle impostazioni |
| `rnwflow_results_ui.py` + `RNWFlowCalculation_results.xaml` | finestra dei risultati |

Le impostazioni si salvano nella sezione `ESA_RNWFlowCalculation` della config di pyRevit.

## Limiti

- Lo script si fida del valore calcolato da Revit: se la rete è scollegata o il sistema non
  è configurato perché Revit sommi le Fixture Units, i tubi compaiono come `NO FLOW`.
- Non è controllato che il diametro non diminuisca nel verso del flusso.
- **Select in Model** lavora mentre la finestra è aperta (modale): la selezione e lo zoom
  sono da verificare in Revit.
- La velocità è stimata dalla relazione di riempimento parziale tarata sulle tabelle, non
  dalla scabrezza o dal diametro interno del tubo del modello.
