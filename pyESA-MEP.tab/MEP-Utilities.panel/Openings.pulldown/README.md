# Openings (forometrie, void demand)

Fori di predisposizione dove tubi, canali, passerelle e conduit del modello host attraversano muri e solai, del
modello host o di link RVT. Tutto passa da tre parametri di testo: niente Extensible Storage, niente commenti, Mark
non toccato.

## Struttura

| Pulsante | Cosa fa |
| --- | --- |
| Place&Check (cartella `Place.pushbutton`) | **Verifica dove servono fori nuovi** e li crea: selezione di muri e solai, di elementi MEP o di un'area, link da leggere, anteprima, un foro per ogni attraversamento (accorpati se vicini), esito in una finestra (anche "nessun foro da aggiungere") |
| New Report | **Solo lo stato dei fori presenti** (tutto il modello, selezione MEP o ARC, area): check di fori e MEP, pulizia degli Id vecchi, riepilogo, Text Report (export Excel) o Interactive Report. Non cerca fori nuovi |
| Last Report | L'ultimo New Report del modello, senza controllo ne' domande |
| Settings | Impostazioni per utente (`script.get_config("ESA_Openings")`) |

- New Report e Last Report: cartelle `NewReport.pushbutton` e `LastReport.pushbutton` (prima un pulsante unico,
  chiamato Refresh e poi Check&Report).
- **Tutto il codice sta in `openings_modules/`** (cartella senza suffisso: pyRevit non la carica come pulsante),
  una copia sola. I `script.py` dei pulsanti aggiungono la cartella a `sys.path`, chiamano
  `openings_core.start(MODE)` e lanciano il comando; quelli di Place, New Report e Last Report differiscono solo per
  `MODE` e `__title__`. Settings legge da `openings_core` predefiniti, nomi dei parametri e `load_settings()`.

  | Modulo | Contenuto |
  | --- | --- |
  | `openings_core.py` | costanti, impostazioni, stato del comando (`ctx`, `start()`), parametri `e_GEN` |
  | `openings_geometry.py` | sezioni MEP, famiglie e inserimento, area, ricerca degli attraversamenti, lettura dei fori |
  | `openings_pick.py` | selezioni nel modello e smistamento della selezione corrente |
  | `openings_check.py` | controllo dei fori, ambito (`CheckScope`), unioni e divisioni, Recreate, scrittura, Missing openings |
  | `openings_exchange.py` | file del report JSON: scrittura (escape non ASCII a mano, come MEPQTO), lettura, cartella e file ricordati per modello, trasformazioni |
  | `openings_report.py` | file dell'ultimo report, righe e dati del report, `CheckSession`, `LastReport`, Export, `ImportedReport`, finestra del report, `check_main()`, `last_main()` |
  | `openings_place.py` | `place_main()` |
  | `OpeningsCheck_ui.py`, `OpeningsPlace_ui.py`, `OpeningsReport_xlsx.py`, XAML | finestre ed export Excel (non leggono Revit) |

  Ogni modulo importa il precedente con `from ... import *` (core <- geometry <- pick <- check <- exchange <- report
  <- place):
  un modulo non usa funzioni di quelli che vengono dopo. **Rocket mode**: i moduli restano caricati tra un comando e
  l'altro, quindi niente documento, output o impostazioni a livello di modulo. Stanno in `ctx` (`ctx.doc`,
  `ctx.settings`, `ctx.status_param`...), riempito da `start()` a ogni comando; la finestra del report, non modale,
  ne salva una copia e la ripristina prima di ogni azione (`ReportController`). Dopo una modifica ai moduli serve
  pyRevit Reload.
- Il vecchio pulsante Families (creava le famiglie PFV dal template) e' in `_old/`: le famiglie VD_Void Demand
  vengono dalla libreria aziendale.

## Elementi MEP nei link (modello dei soli fori)

I MEP si leggono dal modello aperto **e da tutti i link RVT caricati**: lo stesso codice serve il modello MEP con i
fori (come prima) e il modello Openings con MEP e ARC linkati (piano in `Openings_NOTE-LINK.md`).

- Ogni elemento MEP e' un `MepRef` (`openings_geometry`): elemento, modello (`Source`), curva e sezione gia' in
  coordinate host (sezione misurata nel documento dell'elemento, assi portati con il transform del link).
  `mep_near()`, `meps_in_area()` e `hosts_from_meps()` cercano in tutti i modelli (`link_sources()`).
- **Riferimenti sul foro** (`e_GEN_RelatedElement_ID`): gli Id del modello aperto senza prefisso, poi un gruppo per
  link, in ordine di nome: `111,222; PRJ_MEC: 333,444; PRJ_ELE: 555` (`model_name()` del file linkato; `PRJ_MEC#2` se il
  link e' inserito piu' volte). Leggibile in un abaco, "contiene 333" trova il foro, gli Id di un gruppo si
  incollano in *Select by ID*. `format_refs()` scrive, `parse_refs()` legge (anche gli Id nudi dei modelli unici e
  il formato di prova `PRJ_MEC:333,PRJ_ELE:555`; alla prima scrittura passano al formato nuovo). Internamente ogni
  elemento e' `MepRef.text` = `"333"` o `"PRJ_MEC:333"`, risolto da `mep_by_ref()`.
- **Link letti da Place**: sezione *Linked models* della finestra di Place (spunta = leggi, per MEP e per muri,
  solai e strutture; il modello aperto si legge sempre). Ricordata per modello in `ESA_Openings`
  (`place_skipped_links`: `"<percorso del modello>::<chiave>|<chiave>"`, al massimo 50; `load_model_list()` /
  `save_model_list()` in `openings_geometry`); i link nuovi compaiono
  spuntati, quelli tolti e ora non caricati restano tolti. Con elementi gia' selezionati la finestra non c'e' e vale
  l'ultima scelta. `ctx.skipped_links` toglie quei link da `link_sources()`; anche le selezioni nel modello li
  rifiutano. New Report e Last Report leggono sempre tutti i link.
- **Fase** (`phase_name()`): quella della vista attiva; se la vista non ne ha (fogli, abachi), l'ultima del progetto.
  Ogni modello, aperto compreso, si legge nella fase con lo **stesso nome** (`Source.phase_id`): si cercano solo gli
  elementi che esistono in quella fase (`ElementPhaseStatusFilter`, Existing + New: niente demoliti ne' futuri). Un
  link senza una fase con quel nome non si cerca (`Source.searchable`), ma i riferimenti gia' scritti sui fori si
  risolvono lo stesso (`mep_by_ref` non filtra): niente "link not loaded" per questo.
- **Workset**: nessun filtro (deciso il 19.10.2026: un workset da ignorare del tutto e' raro, e semmai quegli
  elementi vanno in un modello a parte, non spuntato in Place&Check). Un filtro per workset escluderebbe fori senza
  segnale.
- Tutte le ricerche (`meps_in_outline()`, `hosts_from_meps()`, `hosts_in_area()`, `neighbours()`) passano da
  `searched_sources()` e `Source.collector()`, che applica la fase. `models_lines()` riporta la fase e i link senza
  quella fase.
- **Tempi** nell'esito di Place&Check (ricerca, creazione) e nei dettagli del New Report: servono a decidere se la
  lettura dei link va ottimizzata (lettura unica in record Python, fase 2 del piano).
- **Selezione MEP**: Revit non seleziona insieme elementi del modello e dei link. Se il modello aperto ha MEP propri si
  scelgono quelli, altrimenti quelli dei link (`pick_meps()`).
- **Sui MEP non si scrive niente**, ne' nei link ne' nel modello aperto (decisione del 19.10.2026): il legame sta
  solo sul foro. I parametri `e_GEN` rimasti sui MEP da versioni precedenti non si toccano piu'; `linked_openings()`
  li legge ancora per la selezione MEP.
- **Link non caricato**: i riferimenti a quel modello restano sul foro (non sono "deleted"), il foro mantiene lo
  stato scritto, non si ricrea ne' si unisce (`keep_status()`); riga dell'elemento "link not loaded".
- Report: Place e i dettagli del New Report elencano i modelli letti con il percorso (`models_lines()`), i link non
  caricati e i file diversi con lo stesso nome. Show / Select di un elemento di un link selezionano solo i fori.
- Link e riferimenti si rileggono a ogni controllo (`reset_models()`): un link ricaricato con la finestra del report
  aperta viene ripreso al Refresh.

## Export e Import del report (modello Openings -> modelli MEP / ARC)

Terzo pulsante del riepilogo (`OpeningsCheck_summary.xaml`, `ThirdAction`), che cambia con il modello:

- **Modello Openings** (New Report e Last Report): **Export report**. Si sceglie una cartella ogni volta (proposta
  l'ultima, `export_folder` per modello). Scrive `<modello>_Openings.json`, sovrascritto (quello che leggono i
  modelli MEP / ARC), e una copia datata `<modello>_Openings_AAAAMMGG-HHMM.json` come storico. Contiene **tutti i
  fori del modello**, non solo lo scope (`build_export()`): stato dal parametro del foro, dettaglio degli elementi
  dall'ultimo controllo su questo computer (il messaggio dice quanti elementi non sono mai stati controllati qui).
- **Modello MEP / ARC** (Last Report in un modello senza fori gestiti): **Import report**. Si sceglie il file;
  il riepilogo si riempie con il report (`ImportedReport`) e il file si ricorda (`report_file` per modello). Le
  volte dopo Last Report lo riapre da solo, **riletto dal disco**: un Export nuovo nella stessa cartella si vede
  subito. Senza report (o file non leggibile): solo il messaggio (`NoImportedReport`), schede spente, Import acceso.
- Report importato: i fori che citano questo modello (riconosciuto per nome, `model_name()`: per i modelli condivisi
  il nome del file centrale, perche' il titolo di una copia locale ha in coda il nome dell'utente; se il nome non
  c'e', per percorso del centrale), con stato e note dell'export; gli elementi di questo modello si cercano per Id e si confermano
  con lo UniqueId ("not in this model any more" se mancano o l'Id e' stato riusato). Riepilogo con *Report of*,
  *From*, *Exported by*, *File*.
- Interactive Report di un report importato: **sola lettura** (`ReportWindow(readonly=True)`, `ImportController`):
  niente spunte, Recreate, Move, Delete, Refresh. Show / Select: elementi di questo modello per Id; fori come
  riferimenti nel link Openings, se e' linkato qui (`Selection.SetReferences`, Revit 2023+; il foro si ritrova per
  UniqueId); zoom anche sul box del foro, portato nelle coordinate di questo modello con l'inverso del transform
  salvato nel report (`models.<modello>.transform`).
- Last Report nel modello Openings senza un New Report su questo computer: non si ferma piu', mostra tutto il
  modello con gli stati scritti (e l'Export).

Formato del file in `Openings_NOTE-LINK.md` ("Report JSON").

## Parametri

Nomi predefiniti; i tre `e_GEN` si possono cambiare in Settings (le misure no).

| Elemento | e_GEN_RelatedElement_ID | e_GEN_RelatedElement_Cat | e_GEN_Void_Status |
| --- | --- | --- | --- |
| Foro (Generic Model) | elementi MEP: `111,222; PRJ_MEC: 333` | `Pipes, Ducts, Cable Trays, Conduits` (solo quelle presenti) | stato del foro |

Sugli elementi MEP non si scrive niente.

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

Due pulsanti, stessi moduli (`MODE = "new"` / `"last"`):

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
- **attraversamento senza foro**: il New Report **non lo cerca piu'** (decisione del 19.10.2026: il report verifica
  solo lo stato dei fori presenti; i fori nuovi li cerca Place&Check). `CheckSession.missing_rows` e
  `LastReport.missing_rows` valgono `None`: niente gruppo *Missing openings* nel riepilogo, nel Text Report e
  nell'Interactive Report, e Recreate non crea fori nuovi. `missing_openings()`, `missing_entry()`,
  `saved_missing_rows()` e il ramo "create" di `CheckSession.recreate()` sono rimasti nel codice ma non si usano.

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
not readable, deleted, link not loaded; per gli elementi non elencati: crossing it, now associated (o partly
outside), inside, not associated (foro orfano).

**Associazioni e Id vecchi** (`CheckSession.run()`):

- Foro: si scrivono i riferimenti dei suoi elementi che **esistono ancora**, anche se usciti dal foro: restano
  associati (e visibili nel report) finche' il foro non si ricrea. Gli Id di elementi cancellati, o che non sono piu'
  MEP, si tolgono; quelli di link non caricati restano.
- Elementi MEP: non si scrive niente.
- **Foro orfano** (decisione del 19.10.2026): aveva elementi associati e nessuno e' piu' dentro, nemmeno in parte.
  Non si cancella mai da solo: e' Not Associated e lo decide l'utente (Delete). Se ora lo attraversano altri
  elementi (es. foro spostato a mano) **non si associano da soli**: righe "inside, not associated (Recreate
  associates it)" e nota in testa; il foro resta Not Associated. Per trovarli `check_opening()` cerca anche nel
  bounding box del foro, in tutti i modelli (la scansione vede solo i MEP gia' citati dai fori). Recreate lo rifa'
  su quegli elementi; quelli elencati prima perdono l'associazione. Move non c'e'. Un foro con almeno un elemento
  ancora dentro, invece, prende in automatico gli elementi nuovi che lo attraversano (come prima).
- Ricreando un foro, gli elementi che non lo attraversano piu' (`outside_meps()`) non entrano nel foro nuovo, che
  e' Valid, e perdono l'associazione (Not Associated); restano nell'ambito (`CheckScope.release()`) e vengono
  scansionati, ma un foro nuovo per loro lo trova Place&Check.
- Delete di un foro Not Associated: allo stesso modo, i suoi elementi ancora esistenti restano nell'ambito
  (`CheckScope.released`, salvato nel Last Report); un foro nuovo per loro lo trova Place&Check.
- Interactive Report: scegliendo un elemento sotto un foro, Show / Select selezionano anche i fori che lo elencano
  (`owners_of()`).

`OpeningsCheck_ui.py` (usato anche da Place, che lancia il check sugli elementi gia' associati) non legge Revit:
riceve testi e Id interi (`OpeningEntry`) e un oggetto actions (`ReportController`) che lavora sul modello.

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
  che attraversa un muro senza foro non viene segnalato (si usa Place&Check). Si ripete dopo ogni azione
  dell'Interactive Report.

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

- **Place&Check**: tutti i MEP rettilinei del modello e dei link letti che toccano l'area (`meps_in_area()`) contro
  muri, solai e strutture che la toccano (`hosts_in_area()`), poi come la modalita' normale (anteprima, fori nuovi).
- **New Report**: fori gestiti nell'area (`opening_in_area()`, centro del bounding box). I MEP dell'area si
  scansionano solo per gli elementi entrati in un foro senza esservi associati; niente fori nuovi. A ogni controllo
  (Refresh, azioni) i fori comparsi nell'area nel frattempo entrano nell'ambito; quelli usciti restano (per Id).

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
- Prima di creare, l'anteprima `OpeningsPlace_confirm.xaml` (conteggi e problemi di lettura).
- **Esito in una finestra**, mai nel pannello di output di pyRevit (`show_result()`, `OpeningsPlace_result.xaml`,
  `ResultWindow`): *Openings created* (riepilogo, elenco dei fori nuovi, parametri non scritti, modelli letti,
  *Select new openings* li seleziona e li mostra dopo la chiusura), *No new openings needed* (stessi conteggi, cosi'
  si vede perche') o *Nothing to do* (selezione senza elementi utili). Openings non apre piu' il pannello di output.
- `OpeningsPlace_ui.py` si importa dentro `place_main()`: New Report e Last Report non lo caricano.
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
