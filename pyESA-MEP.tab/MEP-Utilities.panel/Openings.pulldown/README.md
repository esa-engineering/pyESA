# Openings (forometrie, void demand)

Fori di predisposizione dove tubi, canali, passerelle e conduit del modello host attraversano muri e solai, del
modello host o di link RVT. Tutto passa da tre parametri di testo: niente Extensible Storage, niente commenti, Mark
non toccato.

## Struttura

| Pulsante | Cosa fa |
| --- | --- |
| Place | Selezione di muri e solai, di elementi MEP o di un'area, poi un foro per ogni attraversamento (accorpati se vicini) |
| New Report | Check di fori e MEP (tutto il modello, selezione MEP o ARC, area), pulizia degli Id vecchi, fori mancanti, riepilogo, Text Report (export Excel) o Interactive Report |
| Last Report | L'ultimo New Report del modello, senza controllo ne' domande |
| Settings | Impostazioni per utente (`script.get_config("ESA_Openings")`) |

- New Report e Last Report: cartelle `NewReport.pushbutton` e `LastReport.pushbutton` (prima un pulsante unico,
  chiamato Refresh e poi Check&Report).
- I tre `script.py` di Place, New Report e Last Report sono **identici tranne `MODE`** e `__title__`: si modifica
  quello di Place e si rigenerano gli altri due.
- Il blocco "comune" (nome della configurazione, predefiniti, nomi dei parametri, stati scritti) e' ripetuto negli
  script e in Settings. Una modifica va riportata ovunque.
- Il vecchio pulsante Families (creava le famiglie PFV dal template) e' in `_old/`: le famiglie VD_Void Demand
  vengono dalla libreria aziendale.

## Parametri

Nomi predefiniti; i tre `e_GEN` si possono cambiare in Settings (le misure no).

| Elemento | e_GEN_RelatedElement_ID | e_GEN_RelatedElement_Cat | e_GEN_Void_Status |
| --- | --- | --- | --- |
| Foro (Generic Model) | Id degli elementi MEP: `111,222` | `Pipes, Ducts, Cable Trays, Conduits` (solo quelle presenti) | stato del foro |
| Elemento MEP | Id dei fori: `123456,123999` | `Generic Models` | il peggiore dei suoi fori |

Le categorie sono nomi fissi in inglese, non `Category.Name` (che dipende dalla lingua di Revit). MEC = tubi e
canali, ELE = passerelle e conduit: la disciplina non si scrive piu', si ricava dalla classe dell'elemento.

Stati scritti, dal migliore al peggiore:

| Stato | Quando | Stato interno del New Report |
| --- | --- | --- |
| Valid | posato, ricreato, o elementi ancora al loro posto | New, OK |
| To Be Refreshed | elementi spostati ma ancora dentro il foro (il vecchio Misaligned); foro da dividere o unire secondo Settings; foro con almeno un elemento ancora dentro o in parte fuori (anche se altri sono tolti o usciti) | Misaligned, Regroup, Partial |
| Not Associated | elementi tolti o usciti e nessuno rimasto dentro o in parte fuori; foro senza elementi; elemento MEP senza piu' fori | Changed, Obsolete |

Se un foro ha elementi ancora dentro o in parte fuori vale To Be Refreshed (Partial; se To Be Refreshed e' spento
in Settings, Not Associated), anche con altri elementi tolti o usciti. Not Associated solo se non ne resta
nessuno. I fori Partial non entrano nelle unioni, come i Changed. Split e Joined
non esistono piu': un foro ricreato (anche diviso o unito) e' Valid. Lo stato interno serve solo al New Report e al
dettaglio del report; nel modello va solo lo stato scritto (`WRITTEN`).

Un foro e' "gestito" se ha lo stato compilato, qualunque sia la famiglia: si possono cambiare le famiglie in
Settings senza perdere i fori gia' posati. Rinominando il parametro di stato, invece, i fori posati con il nome
vecchio non vengono piu' trovati.

Famiglie predefinite (Generic Model non ospitate, origine al centro, parametri di istanza condivisi):

- `e_GM.PH.VD_Void Demand_Wall_Rect`: X `e_DIM_Width_1`, Y `e_DIM_Depth_1`, Z `e_DIM_Height_1`
- `e_GM.PH.VD_Void Demand_Wall_Circ`: cilindro lungo X, `e_DIM_Diameter_1`, lunghezza `e_DIM_Width_1`
  (Depth e Height non usati); la X della famiglia va lungo l'asse di attraversamento
- `e_GM.PH.VD_Void Demand_Floor_Rect`: X `e_DIM_Width_1`, Y `e_DIM_Height_1`, Z `e_DIM_Depth_1`
- `e_GM.PH.VD_Void Demand_Floor_Circ`: cilindro lungo Z, `e_DIM_Diameter_1`, lunghezza `e_DIM_Height_1`

Il parametro della lunghezza da parte a parte per ogni famiglia sta in `LENGTH_PARAMS`, usato sia per scrivere
(Place, ricreazione) sia per leggere il foro esistente (New Report).

Le famiglie rettangolari non sono centrate in Z: il solido poggia sul livello di riferimento, il punto di
inserimento e' il centro della faccia inferiore. `BASE_PARAMS` dice quale parametro misura la famiglia lungo Z
(muro: Height, solaio: Depth): Place inserisce il foro meta' di quella misura sotto il centro dell'attraversamento,
il New Report (`opening_center()`) ricava il centro alzando il punto di inserimento della stessa quantita'. Le
circolari restano centrate sull'origine. Una famiglia scelta in Settings deve seguire gli
stessi assi e gli stessi nomi di parametro.

## Settings (per utente)

Famiglie dei quattro casi; nomi dei tre parametri `e_GEN` (vuoto = predefinito); franco per lato (25 mm),
arrotondamento per eccesso (50 mm), accorpamento entro (100 mm), sporgenza oltre ogni faccia (0 mm); MEC ed ELE
separati (no); stato To Be Refreshed nel New Report (si'; spento, ogni spostamento rende il foro Not Associated).

## Place

1. `PickObjects(PointOnElement)`: muri e solai, host o link.
2. Solidi dell'elemento in coordinate host, MEP rettilinei vicini (bounding box), intersezione asse/solido
   (`IntersectWithCurve`): ogni segmento interno e' un attraversamento.
3. Muro: normale nel punto (vale anche per muri curvi). Solaio: verticale. Elemento MEP quasi parallelo (sotto
   l'angolo minimo di Settings, default 15 gradi = pendenza 27%): corre dentro, nessun foro. Esclude scarichi
   a pavimento e montanti dentro i muri. Nel New Report la stessa soglia rende Not Associated un elemento ruotato fino a
   correre dentro.
4. Sezione esterna con isolamento, dai solidi di elemento e isolamento sugli assi del connettore (vale per
   canali ruotati e passerelle di costa). La sezione all'ingresso e all'uscita, proiettata sul piano del foro,
   da' l'ingombro anche per gli attraversamenti obliqui.
5. Accorpamento: stesso elemento attraversato, piani paralleli, distanza <= accorpamento, stessa disciplina se
   separate. Tondo solo per un elemento tondo singolo e (quasi) perpendicolare.
6. Profondita' = spessore attraversato + 2 x sporgenza.
7. Gia' forato: il punto medio dell'attraversamento sta nel bounding box di un foro gestito.

## New Report e Last Report

Due pulsanti, stesso script (`MODE = "new"` / `"last"`):

- **New Report** (`check_main()`, classe `CheckSession`): finestra dell'ambito (`OpeningsCheck_scope.xaml`): tutto
  il modello; selezione MEP (fori associati agli elementi MEP o ai fori scelti, `linked_openings()`); selezione ARC
  (fori su muri, solai e strutture scelti, anche nei link, `openings_on_hosts()`). Se c'e' gia' una selezione la si
  usa e la finestra preseleziona l'opzione giusta, altrimenti si seleziona dopo; ultimo ambito in `ESA_Openings`
  (`check_scope`). Poi scansione degli attraversamenti (sotto), controllo di ogni foro (`check_opening()`) e
  **scrittura sempre**, in una transazione (un Undo): stati, Id, pulizia degli Id vecchi. Si scrivono solo i valori
  che cambiano (`set_text()`): nei modelli condivisi non si prendono in prestito elementi per niente. Il report
  (ambito, fori mancanti, dettaglio per elemento) si salva per il Last Report.
- **Last Report** (`last_main()`, classe `LastReport`): nessun controllo e nessuna domanda. Riapre l'ultimo New
  Report del modello fatto **su questo computer**: stesso ambito e fori mancanti dal file, stati e Id letti dal
  modello adesso (scritti dall'ultimo New Report, di chiunque), dettaglio per elemento dal file. Se lo stato nel
  modello e' diverso da quello salvato lo dice. Senza un New Report salvato: messaggio e basta.

Poi per entrambi il **riepilogo** (`OpeningsCheck_summary.xaml`, `show_report()`): totali per stato (riga *Report
of* con la data, nel Last Report) e due schede grandi:

- **Text Report** (`OpeningsCheck_text.xaml`, `TextReportWindow`, dati da `report_data()`): finestra non modale
  nello stile delle altre, non il pannello di pyRevit. Riepilogo, una sezione per stato (Missing openings, Not
  Associated, To Be Refreshed: ogni foro con note ed elementi a colonne; Valid in forma compatta; Id vecchi tolti
  dai MEP), dettagli (impostazioni, conteggi della scansione, azioni fatte). Oltre 300 voci per sezione il resto e'
  solo nell'Excel. **Export to Excel** (`OpeningsReport_xlsx.write_report()`, writer ridotto da `mepqto_xlsx`, niente
  Excel installato): foglio *Report* (una riga per elemento, foro ripetuto, filtro automatico) e foglio *Summary*;
  file aperto in Excel -> messaggio. Non tocca il modello: niente ExternalEvent.
- **Interactive Report** (`OpeningsCheck_report.xaml`): albero stato -> foro -> elementi, ricerca, dettaglio del
  nodo scelto. *Show in model* / *Select and close* sul nodo scelto (un foro con i suoi elementi, un gruppo intero);
  *Create / Recreate*, *Move onto their elements* e *Delete (Not Associated)* sui fori spuntati (Delete solo sui Not
  Associated: conferma con l'elenco, una transazione, poi New Report dell'ambito; gli Id dei fori cancellati si
  tolgono dai MEP tramite `CheckScope.removed`). Dopo ogni azione i fori si ricontrollano, si riscrive e l'albero si
  aggiorna (i fori ricreati entrano nell'ambito al posto dei vecchi). *Text report* apre il Text Report dello stato
  attuale. Da un Last Report la prima azione fa prima il New Report dell'ambito (`ReportController.check_again()`).

**Finestre non modali** (Text Report e Interactive Report, `Window.Show()`, owner = finestra di Revit):

- dai pulsanti di una finestra non modale non si chiama l'API: ogni azione dell'Interactive Report passa da
  `ReportController.call(nome, argomenti, fatto)`, che la mette in coda su un `ExternalEvent` (`ReportRunner`);
  Revit la esegue appena e' libero e `fatto(risultato, errore)` aggiorna la finestra (stesso thread). Intanto i
  pulsanti delle azioni sono spenti;
- lo script dichiara `__persistentengine__ = True`: il comando finisce, ma il motore IronPython resta vivo per
  le finestre;
- il modello puo' cambiare mentre la finestra e' aperta: *Refresh* rifa' il New Report dell'ambito; Create /
  Recreate, Move e Delete lo rifanno sempre prima di agire (Recreate poi mostra l'anteprima);
- le azioni girano solo se il modello attivo e' quello del report, altrimenti un messaggio chiede di tornarci;
- niente piu' pannello di output: le azioni fatte finiscono nella finestra (*Last action*) e nei dettagli del Text
  Report.

Le altre finestre (ambito, riepilogo, anteprima di Place, Settings) restano modali: scelte brevi, e dopo l'ambito
parte la selezione nel modello. Nel codice Python dei moduli UI niente emoji: stanno negli XAML (un `Run` fisso
piu' un `Run` con nome che il codice riempie).

**Scansione degli attraversamenti** (New Report e Refresh, non Last Report; `scan_crossings()`): prima di
controllare i fori si cercano gli attraversamenti **dei soli elementi MEP del report** (`scan_meps()`: quelli
elencati nei fori dell'ambito, piu' quelli scelti nella selezione MEP), qualunque sia l'ambito, con la stessa
rilevazione di Place sui muri, solai e strutture che toccano (`hosts_from_meps()`, anche nei link;
`find_crossings(..., keep_existing=True)` restituisce anche il foro che contiene l'attraversamento). Gli altri
elementi del modello non si guardano: per quelli c'e' Place.

- **elemento dentro un foro ma non associato** (`unlisted_crossings()`): entra nel check di quel foro
  (`check_opening(instance, extra)`), stato "crossing it, now associated" (o "partly outside"), e viene associato
  alla scrittura; il foro entra nell'ambito (`CheckScope.add()`). Lo stato del foro segue le regole solite: To Be
  Refreshed se il foro ideale cambia misura, Not Associated se l'elemento esce in parte. Gli elementi aggiunti non
  contano come "persi";
- **attraversamento senza foro** (`missing_openings()`): gruppo **Missing openings** nel riepilogo, nel report
  testuale e in cima all'Interactive Report, raggruppato come Place (`merge_overlapping()` + `clusters()`),
  chiavi stabili `new:<elementi attraversati>:<elementi MEP>` (una spunta data nel Last Report resta valida dopo il
  ricontrollo che precede Create). Sono esclusi gli attraversamenti che il check di un foro dell'elemento gia' copre
  (elemento uscito ma entro la portata del foro, `handled_by_opening()`), per non creare due fori. *Create /
  Recreate* li crea come Place (stato Valid, origine "new") insieme alle ricreazioni, in una transazione.
**Recreate** (`recreate_openings()`):

- prima della conferma, un'anteprima per foro (`recreate_plan()`, anche nel dettaglio del report): ricreato su N
  elementi, diviso in due (MEC / ELE), unito con quali fori (si ricrea **tutto il gruppo**, anche i fori non
  spuntati), quanti elementi perdono l'associazione;
- dopo: eventi in testo ("Opening 123 split into 456 (MEC) and 457 (ELE)", "Openings 123 + 124 joined into 458"),
  e ogni foro nuovo porta la sua origine nell'etichetta ("split from 123 (MEC)", "joined from 123 + 124",
  "recreated from 123"; anche nel Last Report);
- poi il check riscrive gli Id: i MEP prendono gli Id dei fori nuovi, perdono quelli cancellati; il messaggio dice
  su quanti MEP sono cambiati e quanti sono rimasti senza foro. I MEP che citano un foro cancellato si riscrivono
  anche se quel foro non e' piu' nell'ambito (`CheckScope.removed`).

**File dell'ultimo report**: `script.get_document_data_file("ESA_Openings_LastCheck", "json")` (cartella dati di
pyRevit dell'utente, un file per modello). `{"format": "ESA_Openings_LastCheck", "openings": {Id del foro: {status,
no_element, notes, move, origin, date, elements: [[Id, stato]]}}, "report": {date, scope: {kind, description, ids,
meps}, missing: [campi di `MISSING_SAVED`]}}`: ogni New Report sovrascrive i fori controllati, lascia gli altri, toglie
quelli cancellati, e sostituisce `report`. Errori di lettura o scrittura silenziosi: serve solo al Last Report.

Stato di ogni elemento elencato nel foro (`ELEMENT_*`): inside, partly outside (la sezione tocca ancora il foro),
moved out (non lo tocca piu': anche entro la distanza di unione, `overlaps_opening()`), no longer crosses, section
not readable, deleted.

**Associazioni e Id vecchi** (`plan_writes()`):

- Foro: si scrivono gli Id dei suoi elementi che **esistono ancora**, anche se usciti dal foro: restano associati
  (e visibili nel report) finche' il foro non si ricrea. Gli Id di elementi cancellati, o che non sono piu' MEP, si
  tolgono.
- Elemento MEP: la lista e' lo **specchio dei fori**: gli Id dei fori che lo citano, letti da tutti i fori gestiti
  del modello (quelli fuori ambito con il valore che hanno). Id di fori cancellati o che non lo citano piu' si
  tolgono; fori che lo citano ma mancavano si aggiungono. Stato = il peggiore di quei fori; nessun foro: Not
  Associated. Fuori dall'ambito si toccano solo i MEP legati ai fori controllati (o scelti).
- Ricreando un foro, gli elementi che non lo attraversano piu' (`outside_meps()`) non entrano nel foro nuovo, che
  e' Valid, e perdono l'associazione (Not Associated); restano pero' nell'ambito (`CheckScope.release()`), cosi' la
  scansione li riporta in *Missing openings* finche' Create non fa un foro attorno a loro.
- Delete di un foro Not Associated: allo stesso modo, i suoi elementi ancora esistenti restano nell'ambito (`CheckScope.released`,
  salvato nel Last Report), cosi' restano scansionati e compaiono in *Missing openings* se attraversano un muro o
  solaio senza foro. Nelle righe di *Missing openings* un elemento ancora elencato da fori presenti a modello mostra
  il suo stato in quei fori e i loro Id (es. "moved out (opening 3541063)"), e Show li seleziona; se i fori sono
  stati cancellati, "no opening (opening N deleted)" (`missing_status()`).
- Interactive Report: scegliendo un elemento sotto un foro, Show / Select selezionano anche i fori che lo elencano
  (`owners_of()`).

`OpeningsCheck_ui.py`, `OpeningsReport_xlsx.py` e i quattro XAML `OpeningsCheck_*` sono **copie identiche** nelle
cartelle di Place, NewReport e LastReport: Place lancia il check
sugli elementi gia' associati. Il modulo non legge Revit: riceve testi e Id interi (`OpeningEntry`) e un oggetto
`CheckActions` che lavora sul modello.

Il foro fa da muro (piano per il centro, perpendicolare all'asse di attraversamento, spessore = lunghezza del foro)
per capire se gli elementi sono ancora dentro. **Profondita' e centro lungo l'asse** invece vengono dalla scansione
(`scanned_plane()`): gli attraversamenti degli elementi del foro, paralleli al foro e che ne toccano lo spessore,
sono gia' allungati da `find_crossings` attraverso gli strati entro l'intercapedine di Settings, come in Place.
Un muro nuovo accanto (o uno tolto, o sporgenza e intercapedine cambiate in Settings) cambia la profondita': nota
"depth 250 -> 600", foro To Be Refreshed, e Recreate lo rifa' con la profondita' nuova. Senza dati di scansione
restano piano e profondita' attuali. Per ogni foro gestito, con gli Id di `e_GEN_RelatedElement_ID`:

- elemento sparito, non piu' attraversante, o uscito oltre meta' foro + accorpamento, oppure sezione (con
  isolamento) non piu' tutta dentro il foro: Partial (**To Be Refreshed**) se resta almeno un elemento dentro o in
  parte fuori; altrimenti Changed (**Not Associated**);
- tutti dentro ma il foro ideale si e' spostato (5 mm) o ha altre misure (1 mm): Misaligned (**To Be Refreshed**; Not
  Associated se l'opzione e' spenta);
- nessun elemento rimasto: Obsolete (**Not Associated**; il foro resta, non si ricrea).

Si ricreano dalla finestra del report i fori spuntati Not Associated o To Be Refreshed con elementi ancora
attraversanti. Un foro ricreato e' Valid e ha un nuovo Id.

## Limiti noti (da verificare in Revit)

- Muri e solai si rileggono solo attraverso la scansione degli elementi del report: un foro senza elementi
  (Not Associated, nessun elemento rimasto) non si confronta con il muro.
- La scansione degli attraversamenti riguarda solo gli elementi MEP del report: un elemento mai associato a un foro
  che attraversa un muro senza foro non viene segnalato (si usa Place). Si ripete dopo ogni azione dell'Interactive
  Report.
- I MEP mai passati da Place non hanno stato: Not Associated si scrive solo sui MEP che avevano un foro e lo hanno
  perso, non su tutti i MEP del modello che attraversano un muro senza foro.

## New Report: ambito e spunta MEC/ELE

- **Selezione ARC:** un foro appartiene a un elemento se una linea per il centro del suo bounding box, lungo X, Y
  o Z della famiglia, entra nel solido dell'elemento.
- **Spunta MEC/ELE cambiata dopo Place:**
  - separati: un foro misto e' To Be Refreshed (Regroup) e, ricreato, diventa due fori (uno per disciplina);
  - insieme: un foro MEC e uno ELE sullo stesso piano, a distanza <= accorpamento, sono To Be Refreshed e,
    ricreati, diventano un foro solo.
- **Fori avvicinati:** due o piu' fori sullo stesso piano i cui fori ricalcolati (franco compreso) distano al
  massimo la distanza di unione sono To Be Refreshed e, ricreati, diventano uno solo (Valid). Con MEC ed ELE
  separati si uniscono solo fori della stessa disciplina. I fori Not Associated (Changed: elementi tolti o usciti)
  **non entrano nelle unioni**: un elemento spostato vicino a un foro in ordine non deve renderlo Not Associated.
  Ricreato il foro Not Associated sulla nuova posizione, il check dopo propone l'unione dei due (To Be Refreshed). I
  gruppi hanno cosi' solo fori To Be Refreshed e si ricreano tutti insieme o per niente.

## Strutture composite (Place)

Un muro composto (un solo elemento a strati) e' gia' attraversato tutto dall'intersezione con il suo solido. Se
invece la struttura e' fatta di piu' elementi (es. muro strutturale nel link, intercapedine, controparete nel
modello host), l'attraversamento viene allungato strato per strato:

- strati cercati: muri (fori a muro) o solai (fori a solaio) del modello host e dei link RVT caricati, entro 1 m
  dall'elemento scelto;
- uno strato conta se, lungo l'asse dell'elemento MEP, sta entro l'intercapedine massima di Settings ("Join
  layers across", 200 mm) da quanto gia' attraversato; l'intercapedine entra nel foro;
- per i muri lo strato deve essere parallelo all'elemento scelto (entro circa 11 gradi): un muro a T o d'angolo
  non conta; e uno strato attraversato per piu' di 1 m e' scartato (l'elemento MEP ci corre dentro);
- si ripete finche' non si trovano altri strati: il foro attraversa tutta la struttura.

La chiave di accorpamento e' l'intera pila: scegliendo piu' strati della stessa struttura si ottiene un foro solo.
Il report conta gli attraversamenti allungati e, nella colonna "Crossed element", elenca gli strati aggiunti.

## Area (Place e New Report)

Terza scelta della finestra di Place (`PICK_AREA`) e quarto ambito del New Report (`SCOPE_AREA`). `pick_area()`
legge l'area dalla vista attiva:

- pianta (floor, ceiling, engineering plan): riquadro trascinato (`PickBox`), allineato agli assi dello schermo
  (riquadro di taglio ruotato compreso), quote dal fondo alla cima del view range (`view_range_z()`, come Quick
  SectionBox; view range illimitato: livello + 4000 mm);
- vista 3D: la section box attiva, senza trascinare; senza section box un messaggio;
- altre viste: messaggio.

L'area e' un dict di soli numeri (`make_area()`: asse destra, min / max in destra, alto, quota) salvato nel file del
Last Report (`scope.area`), cosi' Refresh e Last Report la ritrovano. Un attraversamento e' nell'area
(`crossing_in_area()`) se il punto medio sta nel rettangolo in pianta e il tratto tocca l'intervallo di quote: con il
view range consueto (fondo al livello) il solaio sotto il livello ha il centro fuori quota ma il tratto lo tocca.
Gli elementi MEP che escono dall'area contano, ma solo per gli attraversamenti dentro (`stats["outside_area"]`).

- **Place**: tutti i MEP rettilinei dell'host che toccano l'area (`meps_in_area()`) contro muri, solai e strutture
  dell'host e dei link che la toccano (`hosts_in_area()`), poi come la modalita' normale (anteprima, fori nuovi).
- **New Report**: fori gestiti nell'area (`opening_in_area()`, centro del bounding box) e scansione dei MEP dell'area;
  attraversamenti senza foro in *Missing openings*. A ogni controllo (Refresh, azioni) MEP e fori comparsi nell'area
  nel frattempo entrano nell'ambito; i fori usciti dall'area restano (per Id).

## Place: modalita' dalla selezione

- **Elementi gia' selezionati** al click (`split_selection()`), nessuna domanda:
  - fori gestiti ed elementi MEP che rimandano ad almeno un foro esistente (`e_GEN_RelatedElement_ID`): **check**
    sui fori legati a loro (`linked_openings()`), senza la finestra dell'ambito (scrittura come l'ultima scelta),
    poi riepilogo e report come nel New Report;
  - elementi MEP senza foro (anche Not Associated, o con Id di fori cancellati): modalita' inversa. Si cercano muri,
    solai e strutture (travi, pilastri, fondazioni) del modello host e dei link caricati che toccano il bounding box
    di almeno un elemento scelto, poi la stessa rilevazione della modalita' normale, limitata agli elementi scelti;
  - muri, solai e strutture del modello host: come se fossero scelti con la selezione (tutti i MEP che li
    attraversano). Gli elementi di un link non si possono preselezionare: si usa la finestra.
  - Prima il check, poi i fori nuovi (chiudere il riepilogo non ferma i fori nuovi). Gli altri elementi sono
    contati nel report come non usati.
- **Nessuna selezione**: la finestra `OpeningsPlace_pick.xaml` chiede se selezionare muri/solai/strutture o elementi
  MEP (ultima scelta ricordata in `ESA_Openings`, chiave `place_pick`).
- Prima di creare, l'anteprima `OpeningsPlace_confirm.xaml` (conteggi e problemi di lettura); il report resta nel
  pannello di output.
- `OpeningsPlace_ui.py` e i due XAML stanno solo nella cartella di Place: lo script li importa dentro `place_main()`,
  cosi' New Report (stesso script) non ne ha bisogno.
- **Strutture** (anche nella modalita' normale e nell'ambito del New Report): foro a solaio per elementi MEP verticali
  (oltre 45 gradi), a muro per gli altri, con asse perpendicolare alla trave o, per i pilastri, lungo l'elemento MEP;
  un attraversamento piu' lungo di 1 m e' scartato (l'elemento ci corre dentro). Niente allungamento per strati.

## Elementi sovrapposti (trave nel muro, pilastro nel muro...)

Prima dell'accorpamento, gli attraversamenti dello **stesso elemento MEP** attraverso elementi diversi che si
sovrappongono o distano al massimo l'intercapedine di Settings lungo l'elemento (stesso tipo di foro, assi
paralleli) diventano uno solo: intervallo complessivo, tutti gli elementi attraversati, sistema del foro preso da
muro o solaio se c'e'. Nell'accorpamento due attraversamenti stanno nello stesso foro se hanno almeno un elemento
attraversato in comune (non serve piu' la stessa identica pila).

## New Report: elementi fuori dal foro

- **Elementi fuori dal foro** (in parte fuori, o spostati oltre la distanza di unione ma ancora attraversanti con
  lo stesso orientamento): nella finestra del report, *Move onto their elements* sui fori spuntati
  (`alignment_vector()`; gli elementi MEP non si toccano, le misure del foro restano). Spostamento nel piano del
  foro: centrato sugli elementi se il foro e' di un solo elemento, tondo, o se non ci stanno comunque; altrimenti il
  minimo perche' tutti rientrino con il franco. E' una transazione a parte (un Undo); poi i fori si ricontrollano
  (un foro spostato ma ancora troppo piccolo risulta Not Associated).
