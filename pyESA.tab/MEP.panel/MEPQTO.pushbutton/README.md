# MEP QTO: strumento pyRevit

Estrae il computo metrico di un modello MEP partendo dal **Type Mark**. Le categorie
"a pezzo" contano le istanze (ogni istanza conta 1). Canali, tubi, flessibili, passerelle,
conduit e isolanti si misurano in m, mq o kg con le formule dei fogli QTO aziendali (vedi
"Regole di misura").

**Il modello viene solo letto, mai modificato.** Dal modello arrivano solo i codici
prezzario e il Sì/No di inclusione. Capitolo, n. articolo EPU, prezzario, descrizione,
unità e prezzo delle voci stanno nel listino comune e nel file di progetto (vedi "Dove
finiscono i dati").

## Categorie computate

**A pezzo:** Air Terminals, Communication Devices, Conduit Fittings, Data Devices, Duct
Accessories, Electrical Equipment, Electrical Fixtures, Fire Alarm Devices, Lighting
Devices, Lighting Fixtures, Mechanical Equipment, Nurse Call Devices, Pipe Accessories,
Plumbing Fixtures, Security Devices, Sprinklers, Specialty Equipment.

**Lineari:** Ducts, Flex Ducts, Pipes, Flex Pipes, Cable Trays, Conduits, Duct Insulation,
Pipe Insulation.

L'elenco è la tabella `CATEGORY_RULES` in `mepqto_model.py`, con il tipo di misura di
ogni categoria. Le formule stanno in `mepqto_rules.py`.

## Modelli e categorie da leggere

Prima di ogni lettura si apre la finestra **Models and categories** (modelli, workset da
leggere, categorie): all'avvio del tool
e dal pulsante **Models and Categories...** (riga Models nelle impostazioni). Annullando
la scelta iniziale il tool non si apre.

- **Modelli:** il modello aperto si legge sempre. In più si spuntano le istanze di link
  Revit da leggere. Ogni istanza posizionata si legge e si computa per conto suo: un file
  linkato due volte compare come `File.rvt [1]` e `File.rvt [2]`, e se si spuntano
  entrambe conta doppio. I link non caricati compaiono disattivati. I link annidati
  dentro i modelli linkati non si leggono.
- **Fase dei link:** un link si legge nella fase con lo stesso nome di quella scelta nella
  finestra del computo. Se non ce l'ha non si legge e compare nella scheda Issues come
  "Linked model not read", così il computo non prende elementi di fasi sbagliate. Lo
  stesso vale per un link scaricato dopo la scelta.
- **Workset da leggere:** l'elenco riunisce, per nome, i workset utente del modello
  aperto e dei link spuntati, e si aggiorna quando cambiano le spunte dei link (il
  tooltip di ogni riga dice in quali modelli c'è). Si leggono **solo gli elementi dei
  workset spuntati**, in tutti i modelli che hanno un workset con quel nome; quelli dei
  workset non spuntati si saltano. Serve almeno un workset spuntato.
  - **Workset nuovi:** si ricordano i workset spuntati e quelli già visti. Un workset mai
    visto prima (creato dopo, oppure di un link spuntato per la prima volta) compare
    **spuntato**, e il conteggio sotto l'elenco lo segnala ("New since last time,
    ticked: ..."): così un workset nuovo non sparisce dal computo senza che te ne
    accorga. La prima volta sono tutti spuntati.
  - La spunta di un workset che non è in elenco (link non spuntato) si conserva e torna
    visibile spuntando il link.
  - Decide il workset dell'elemento: un isolante su un workset spuntato si legge anche
    se il suo canale o tubo sta su un workset non spuntato, e viceversa.
  - Gli elementi saltati si contano: il totale è nelle note della finestra, il dettaglio
    per workset nel tooltip della riga Models. L'intestazione del foglio Bill of
    quantities elenca i workset non letti.
  - I modelli non condivisi non hanno workset: la sezione lo segnala e si legge tutto.
  - Fino alla versione precedente la spunta **escludeva**. Le esclusioni ricordate
    diventano, alla prima apertura, workset non spuntati; i set di workset salvati con
    quella logica (chiave `workset_sets`) non si riusano, perché il loro significato si
    capovolgerebbe: vanno ricreati.
- **Categorie:** si leggono solo quelle spuntate. Nella finestra del computo il pannello
  CATEGORIES mostra solo le categorie lette, tutte spuntate: togliendo una spunta la
  categoria esce dal computo senza rileggere i modelli.
- **Set di categorie e di workset:** funzionano allo stesso modo. Si scrive un nome nella
  casella *Category set* o *Workset set* e si preme **Save Set** per salvare le spunte
  correnti (un set con lo stesso nome si sostituisce, dopo conferma). Scegliendo un set
  dalla tendina se ne applicano le spunte; *Delete Set* cancella il set con il nome nella
  casella. Scrivere nella casella il nome di un set esistente non cambia le spunte: si
  applica solo scegliendolo dalla tendina. Un set di workset elenca i workset **da
  leggere** e si applica esattamente: si spuntano i suoi workset e solo quelli. Un set
  vuoto non è ammesso. I set stanno nella config pyRevit dell'utente (`ESA_MEPQTO`,
  chiavi `category_sets` e `workset_include_sets`), quindi valgono per tutte le commesse
  su quel PC ma non si condividono con i colleghi. I set di workset salvano i nomi,
  quindi valgono anche fra commesse diverse.
- **Cosa si ricorda:** le ultime categorie, gli ultimi workset spuntati e quelli già
  visti, gli ultimi set (`last_categories`, `last_included_worksets`,
  `last_known_worksets`, `last_category_set`, `last_workset_set`)
  e, per ogni modello, i link scelti l'ultima volta (`scope_links`, per UniqueId
  dell'istanza, ultime 50 associazioni).
- **Cambiare la scelta** rilegge i modelli, come un cambio di fase.

Il computo somma gli elementi di tutti i modelli letti: un Type Mark presente nel modello
aperto e in un link ha una sola riga nel Bill of quantities. Il riepilogo Type Marks invece
ha un gruppo per modello (colonna Model), perché lo stesso tipo può portare codici diversi
nei diversi modelli. Gli ElementId degli elementi dei link sono riportati con il nome del
link (`MEP_A.rvt: 12345`), perché modelli diversi possono avere gli stessi Id.

## Come si calcola

1. Si raccolgono le istanze delle categorie scelte, nel modello aperto e nei link scelti
   (vedi "Modelli e categorie da leggere"), che esistono nella fase scelta:
   - con stato `New`, solo quelle create nella fase;
   - con `New + Existing`, anche quelle che vengono dalle fasi precedenti.

   Le demolite non contano mai. Si computano solo il modello principale e le opzioni di
   progetto primarie: le istanze delle opzioni secondarie si saltano sempre (vedi
   "Opzioni di progetto"). Si tolgono poi gli elementi che il **Sì/No di inclusione**
   mette a No (vedi "Parametri del modello").
2. Per ogni istanza si leggono:
   - il Type Mark del tipo;
   - i codici delle voci, dai parametri impostati con **Parameters...** (default fra
     parentesi):
     - **tutte le categorie:** parametri di istanza (`e_DAT_PriceCode_i_1..10`);
     - **categorie a pezzo, in più:** parametri di tipo (`e_DAT_PriceCode_1..10`); se il
       tipo non ha il parametro, si legge quello d'istanza con lo stesso nome. I codici di
       tipo e quelli d'istanza **si sommano**: un elemento conta per tutti i suoi codici.
       Un codice presente sia sul tipo sia sull'istanza conta una volta sola;
     - **categorie lineari** (canali, tubi, flessibili, passerelle, conduit, isolanti):
       solo i codici d'istanza, perché lo stesso tipo cambia voce con il diametro o la
       sezione. Qui i parametri di tipo **non** vengono letti.
3. La quantità di un codice è la somma, sugli elementi delle categorie spuntate che lo
   portano, di:
   - 1 per le categorie a pezzo;
   - la misura nell'**unità della voce** (scheda EPU) per quelle lineari,
     maggiorata della percentuale della categoria per raccordi e sfridi.

   Gli elementi senza Type Mark sono esclusi. L'importo è la quantità per il prezzo
   unitario.

Le famiglie annidate condivise sono istanze vere e quindi si contano. Nel riepilogo Type
Mark hanno una riga separata (`Nested = Yes`), così un doppio conteggio padre/figlio si
vede subito.

## Opzioni di progetto

Oggi il computo comprende **solo il modello principale e le opzioni primarie**, senza
scelta nella finestra. Le istanze delle opzioni secondarie si saltano e si contano: le
note sotto le impostazioni riportano "N instances in secondary design options skipped".
Vale per il modello aperto e per i link.

**Come si riconoscono.** Per ogni elemento si legge `element.DesignOption`:

- `None`: l'elemento è nel modello principale e si legge;
- un'opzione con `IsPrimary = True`: si legge;
- un'opzione con `IsPrimary = False`: è un'opzione secondaria e si salta.

Il controllo sta in `_collect_source()` in `mepqto_model.py` ed è comandato da
`CollectOptions.primary_only`; la finestra passa sempre `PRIMARY_OPTIONS_ONLY = True`
(costante in `mepqto_ui.py`). La logica è rimasta intatta, quindi la funzione si può
riattivare.

**Prima versione (da ripristinare, se serve).** Nella sezione SETTINGS c'era la casella
*Main model and primary design options only* (`chk_primary_only`, spuntata di default).
Togliendo la spunta si passava `primary_only=False` e si leggevano **tutte** le opzioni,
primarie e secondarie insieme, rileggendo il modello come per un cambio di fase; lo stato
si ricordava nella config (`last_primary_only`). Per ripristinarla: rimettere la casella
in `MEPQTO_form.xaml`, collegare `Checked` / `Unchecked` a `OnCollectOptionsChanged` e
passare `bool(chk_primary_only.IsChecked)` al posto della costante in `_collect_model()`.

**Attenzione prima di riattivarla.** Leggere tutte le opzioni insieme somma alternative
che si escludono a vicenda: due opzioni dello stesso set con lo stesso impianto lo
contano due volte. Per un computo di una variante serve piuttosto scegliere **una
opzione per ogni set** (Design Option Set), per esempio nella finestra di scelta dei
modelli. `element.DesignOption` dà l'opzione dell'elemento e le opzioni del documento
si elencano con `FilteredElementCollector(doc).OfClass(DesignOption)`; il set a cui
appartiene un'opzione si legge da un suo parametro (probabilmente
`BuiltInParameter.OPTION_SET_ID`: **da verificare sull'API** prima di usarlo). Nei link
le opzioni appartengono al documento linkato, quindi la scelta va fatta per modello.

## Parametri del modello

Il pulsante **Parameters...** (riga Parameters nelle impostazioni) apre la mappatura dei
parametri che il tool legge dal modello. Come per la WBS, ogni campo propone i parametri
trovati su un campione di elementi, oppure se ne scrive il nome.

| Gruppo | Dove si legge | Default |
| --- | --- | --- |
| Price codes di tipo, categorie a pezzo (10) | tipo, poi istanza con lo stesso nome | `e_DAT_PriceCode_1..10` |
| Price codes d'istanza, tutte le categorie (10) | istanza | `e_DAT_PriceCode_i_1..10` |
| Include Yes/No, tutte le categorie | istanza | `e_DAT_BOQ_i` |

- La posizione conta: "Code N" di tipo alimenta la colonna "Type price code N" del
  riepilogo Type Marks, "Code N" d'istanza la colonna "Instance price code N". I campi
  vuoti si saltano.
- Lo stesso parametro non può comparire due volte nello stesso gruppo, né in entrambi i
  gruppi dei codici (sulle categorie a pezzo si leggerebbe due volte), né essere sia un
  codice sia un Sì/No. *Restore Defaults* rimette i valori della tabella.
- **Sì/No di inclusione:** un solo parametro d'istanza, letto su ogni elemento di tutte
  le categorie. Esclude solo il No esplicito. Gli elementi senza il parametro o senza
  valore sono computati, quindi un modello senza il parametro si computa per intero.
  Lasciando vuoto il campo, il controllo non si fa. Il parametro va assegnato come
  parametro d'istanza a tutte le categorie computate: dove manca, gli elementi si
  contano sempre, senza avvisi.
- **Elementi esclusi:** spariscono dal computo, dal riepilogo Type Marks e dall'elenco
  prezzi, ma compaiono nella scheda Issues come "Excluded from the bill", con il numero
  di elementi per tipo.
- **Salvataggio:** la mappa si salva nel file di progetto (`parameters`). Cambiandola, i
  valori si rileggono dal modello. Il foglio Rules dell'export Excel riporta la mappa
  usata.

## Le schede della finestra

| Scheda | Contenuto | Modificabile |
| --- | --- | --- |
| EPU | capitolo, sottocapitolo, codice prezzario, descrizione breve, descrizione, unità, prezzo unitario dei codici usati nel modello | sì: tutto tranne il codice prezzario (unità da tendina) |
| Bill of quantities | Type Mark, codice prezzario, descrizione, unità, quantità, prezzo unitario, importo, con il totale sotto; una riga per Type Mark e codice (per combinazione WBS, con le colonne dei livelli dopo Amount); voci raggruppate per i livelli scelti con Layout... (default capitolo › sottocapitolo), con il totale di ogni gruppo e gruppi da aprire e chiudere; livelli WBS anche sulle colonne, come una tabella pivot | no, si compila da sola; si può solo dare a una voce una maggiorazione propria |
| Manual items | voci non modellate: WBS, Type Mark, codice, descrizione, unità, quantità, prezzo unitario, importo | sì: tutto tranne l'importo; descrizione, unità e prezzo sono quelli della voce EPU |
| CME | computo metrico estimativo nella forma del foglio Bill of quantities dell'export Excel: colonne WBS in testa, una riga di totale in grassetto per ogni combinazione WBS seguita dalle sue voci, ordinate per Type Mark e codice | no, come Bill of quantities (override della maggiorazione compreso) |
| Type Marks | matrice: una riga di gruppo per tipo e modello (categoria, Type Mark, famiglia e tipo, annidata sì/no, modello), seguita da una riga per codice con il parametro che lo porta (`e_DAT_PriceCode_1 (T)`, `e_DAT_PriceCode_i_1 (I)`), codice prezzario e descrizione | no |
| Rules | maggiorazioni per categoria, kg/mq della lamiera dei canali, densità dei tubi | sì |
| Issues | anomalie (vedi sotto) | no |

- **Bill of quantities** ha una riga per ogni coppia Type Mark e codice. Un codice usato
  da più Type Mark compare una volta per ognuno, con la sua quantità, così si vede da
  quale tipo arriva. Le righe sono ordinate per codice, poi Type Mark, poi livelli WBS in
  successione (nell'export Excel: Type Mark, poi codice). Una voce che
  non si può misurare (unità o dimensioni mancanti) resta in elenco con quantità 0 e la
  sua anomalia. Il totale per codice resta quello della somma delle righe. Una voce può
  avere una maggiorazione propria (vedi "Override della maggiorazione").
- **CME** mostra lo stesso computo del Bill of quantities nella forma dell'export Excel
  (era la vista del Bill of quantities prima del raggruppamento): colonne WBS in testa,
  una riga `Total` in grassetto per ogni combinazione di valori WBS con il totale delle
  sue voci, poi le voci ordinate per Type Mark e codice. Senza livelli WBS è una lista
  semplice, ordinabile cliccando le intestazioni. Ha la sua ricerca e i suoi filtri sulle
  colonne, indipendenti da quelli del Bill of quantities; una combinazione resta visibile
  se ha almeno una voce visibile. Tasto destro e *Allowance Override...* agiscono sulle
  voci selezionate nella scheda aperta. Si aggiorna insieme al computo. Come nel Bill of
  quantities, le colonne si allargano trascinandone il bordo e la tabella scorre in
  orizzontale; restano ferme le colonne WBS, Type Mark e Price book code,
  così le righe `Total` si riconoscono dai valori WBS anche scorrendo.
- **Manual items** (voci non modellate): oggetti che nel modello non ci sono, scritti a
  mano. Ogni riga ha i valori dei livelli WBS attivi, Type Mark, codice prezzario,
  descrizione, unità, quantità e prezzo unitario; l'importo si calcola. *Add Row*,
  *Duplicate* (righe selezionate), *Remove Selected*.
  - **Descrizione, unità e prezzo sono quelli del codice**, come nella scheda EPU: scritti
    qui vanno nel file di progetto come correzioni della voce EPU (e viceversa), e un
    codice del listino porta i suoi valori. Un codice ha quindi sempre un solo prezzo in
    tutto il computo. Prima del codice non si possono scrivere.
  - Quantità e prezzo accettano la virgola decimale. Una riga senza codice o senza
    quantità è **arancione** e non entra nel computo.
  - **Nelle altre schede:** i codici delle voci manuali entrano nella scheda EPU (colonna
    **Source**: `Model`, `Manual`, `Model + Manual`) e le loro righe sono in **lavanda e
    corsivo**; il rosso del codice assente dal listino prevale sul fondo. Nel Bill of
    quantities e nella CME ogni voce manuale è una riga distinta, in lavanda e corsivo,
    anche quando ha lo stesso Type Mark e codice di una voce del modello; la quantità si
    somma ai totali così com'è, senza maggiorazione, e l'override della maggiorazione non
    vale per queste righe. Un codice manuale assente dal listino compare negli Issues
    ("Code not in the price list", con "Manual items" fra gli usi).
  - **WBS:** i valori si scrivono nelle colonne dei livelli attivi e si salvano per nome
    del parametro, così restano se i livelli cambiano ordine.
  - **Salvataggio:** nel file di progetto (`manual_items`), in blocco come regole e livelli
    WBS: se due utenti li cambiano insieme, vince l'ultimo che salva.
  - **Export Excel:** le quantità manuali sono nel foglio Bill of quantities con il fondo
    lavanda (e una nota in testa), e un foglio **Manual items** le elenca.
- **Disposizione del Bill of quantities** (riga *Layout* sopra la griglia, pulsante
  **Layout...**): come una tabella pivot, si scelgono i livelli sulle **righe** e i
  livelli WBS sulle **colonne**. La finestra ha tre elenchi: campi disponibili, Rows e
  Columns. Un campo si aggiunge in coda (*Add to Rows*, *Add to Columns*, doppio clic =
  righe): **l'ordine di inserimento è l'ordine di raggruppamento**, e *Up* / *Down* lo
  cambiano. Sulle righe vanno Chapter, Subchapter (quelli della scheda EPU), Type Mark,
  Unit e i livelli WBS; sulle colonne solo i livelli WBS. Default: righe Chapter ›
  Subchapter, nessuna colonna (*Restore Default*). La scelta si ricorda nella config
  pyRevit (`bill_group_by` per le righe, `bill_column_levels` per le colonne).
  - **Righe.** Ogni gruppo ha una riga in grassetto, rientrata per livello e con il fondo
    più scuro ai livelli alti, con il nome del gruppo nella colonna Description e il
    **totale dell'importo** di tutte le sue voci nella colonna Amount. Seguono i
    sottogruppi o le voci, ordinate per codice, poi Type Mark, poi livelli WBS. I valori
    mancanti stanno in fondo: `(no chapter)`, `(no subchapter)`, `(no unit)`, `(not set)`
    per la WBS. Senza livelli sulle righe la griglia è una lista semplice e ordinabile.
  - **Colonne.** Ogni combinazione di valori dei livelli WBS sulle colonne diventa una
    colonna, dopo le colonne WBS, nell'ordine dei livelli (con due livelli l'intestazione
    è su due righe, es. `GEN` / `P00`). Una voce non si divide più per quei livelli: ha una
    riga sola e in ogni colonna il suo **importo** in quella combinazione (quantità ×
    prezzo unitario; vuoto se la voce non ha prezzo), come le righe di gruppo. Quantity e
    Amount restano i totali, e la somma delle colonne di una riga dà il suo Amount. La colonna WBS di un livello sulle
    colonne si nasconde. Un livello non può stare sia sulle righe sia sulle colonne.
  - **Aprire e chiudere:** clic sul triangolino a sinistra o doppio clic sulla riga del
    gruppo; *Expand All* / *Collapse All* per tutti. Un gruppo chiuso resta chiuso
    quando il computo si ricalcola (cambio di prezzi, unità, regole); cambiando la
    disposizione si riapre tutto.
  - Con almeno un livello sulle righe le colonne non si riordinano, perché l'ordine delle
    righe è la struttura.
  - **Larghezza e scorrimento:** le colonne hanno larghezza fissa (Description 380 px) e si
    allargano trascinandone il bordo nell'intestazione; la tabella allora scorre in
    orizzontale invece di stringere le altre colonne. Il triangolino, Type Mark e Price
    book code restano fermi mentre si scorre. Le larghezze non si salvano:
    alla riapertura tornano quelle di default. Il nome di un gruppo sta nella colonna
    Description, quindi scorrendo molto a destra non si vede più.
  - I livelli WBS che non sono sulle colonne dividono le voci: una riga per ogni
    combinazione di valori, con i valori nelle colonne WBS, anche quando non sono fra i
    livelli delle righe.
  - Le colonne della disposizione non hanno l'imbuto dei filtri; i filtri sulle altre
    colonne valgono come sempre.
  - È solo la vista della scheda: l'export Excel mantiene la struttura per combinazioni
    WBS descritta in "WBS", la stessa della scheda CME.
- **EPU** (elenco prezzi unitari; in precedenza la scheda si chiamava *Price list*) mostra tutti i codici usati nel modello nella fase scelta, anche quelli
  delle categorie non spuntate. Il computo invece segue le categorie spuntate. Le righe
  sono ordinate per capitolo, sottocapitolo e codice; i codici senza capitolo stanno in
  fondo. Una modifica di capitolo non riordina subito la griglia: l'ordine si aggiorna al
  ricalcolo successivo (cambio di fase o categorie, riapertura). Cliccando un'intestazione
  si può comunque riordinare per quella colonna. La ricerca guarda capitolo,
  sottocapitolo, codice, descrizione breve e descrizione.
- **Filtri sulle colonne** (schede EPU, Bill of quantities e CME), in stile Excel:
  l'imbuto nell'intestazione apre l'elenco dei valori della colonna, con caselle di
  spunta, una ricerca, *(Select All)* e *(Blanks)* per le celle vuote. L'elenco mostra
  solo i valori delle righe che passano la ricerca e gli altri filtri; con una ricerca
  nell'elenco, OK applica i valori trovati. Spuntare tutto toglie il filtro; *Clear
  Filter* lo toglie dalla colonna, **Clear Filters** accanto alla ricerca li toglie
  tutti. Una colonna filtrata ha l'imbuto blu e il titolo in grassetto.
  - I filtri si combinano fra loro e con la ricerca e restano attivi quando il computo si
    ricalcola (cambio di categorie, unità, prezzi). Un filtro su un livello WBS che si
    spegne si toglie.
  - Nel Bill of quantities si filtrano le voci. Una riga di gruppo resta visibile se ha
    almeno una voce visibile e mostra comunque il totale di tutte le sue voci; lo stesso
    vale per la ricerca. Un gruppo chiuso resta chiuso anche con i filtri.
  - I filtri cambiano solo cosa si vede: il totale in fondo alla finestra e l'export Excel
    comprendono tutte le voci. I filtri non si salvano.
  - Il codice sta in `mepqto_grid_filter.py` (`GridFilters`), riusabile su qualsiasi
    griglia legata a una `DataTable`.
- **Colonne della scheda EPU** (intestazioni in inglese nella finestra). N. articolo EPU
  (`epu_item`) e prezzario di riferimento (`price_book`) non sono più colonne delle
  schede EPU, Bill of quantities e CME: restano campi della voce, quindi si leggono dal
  listino (anche con *Columns...*), si modificano nell'editor del listino, si conservano
  nel file di progetto e compaiono nell'export Excel.
  - `Price book code`: il codice della voce nel prezzario, cioè il valore dei parametri
    `e_DAT_PriceCode_*` del modello. È la chiave della voce e non si modifica qui;
  - `Short Description`: descrizione breve della voce, accanto a quella estesa
    (`Description`). Si comporta come la descrizione: arriva dal listino, si corregge qui
    e la correzione va nel file di progetto. È facoltativa: l'anomalia "Code without
    description" guarda solo la descrizione estesa.
- **Ricerca nel listino e righe rosse.** Ogni codice della colonna `Price book code` si
  cerca nella colonna A del listino (come un CERCA.VERT di Excel) e le altre colonne si
  riempiono con i dati della riga trovata (vedi "Listino Excel / CSV"). Le righe il cui
  codice **non c'è** nel listino sono in **rosso** (riempimento rosso chiaro, testo rosso
  scuro) e compaiono nella scheda Issues come "Code not in the price list". La casella
  *Only codes not in the price list* mostra solo quelle. Una descrizione scritta a mano
  nella scheda non toglie il rosso: il codice resta assente dal listino. Senza un listino
  caricato nessuna riga è rossa.
- **L'unità** si sceglie da una tendina con `cad`, `m`, `kg`, `mq`, `mc`. La voce vuota
  torna al valore del listino.
- **Le unità lette dal listino** vengono ricondotte ai valori della tendina quando sono
  varianti note: `n`, `nr`, `n°`, `cad.`, `pz` → `cad` (negli EPU `n` e `cad` compaiono
  entrambe ma indicano la stessa unità); `ml` → `m`; `m2`, `m²` → `mq`; `m3`, `m³` → `mc`.
  Maiuscole, punti e spazi non contano (`N.`, `Nr.` valgono come `n`, `nr`). Le altre
  restano come sono e compaiono in fondo alla tendina.
- **Il riepilogo Type Marks** è una matrice. Ogni tipo ha una riga di gruppo in
  grassetto con categoria, Type Mark, famiglia e tipo, annidata sì/no e modello (aperto o
  link: lo stesso tipo in due modelli ha due gruppi), seguita da una
  riga per ogni codice valorizzato, con la sua descrizione. Prima vengono i codici di
  tipo, poi quelli d'istanza. Un tipo senza codici ha solo la riga di gruppo.
  - **Parameter** (prima *Code slot*): il nome del parametro che porta il codice, secondo
    la mappa di Parameters..., seguito da `(T)` se è un parametro di tipo e da `(I)` se è
    d'istanza. I codici di tipo valgono solo per le categorie a pezzo; quando il tipo non
    ha il parametro e il codice si legge sul parametro d'istanza con lo stesso nome, la
    riga riporta `(I)`. Un buco (codice 1 vuoto, codice 3 pieno) si vede dai nomi. Il
    foglio Type Marks dell'export ha la stessa colonna.
  - Un tipo le cui istanze portano codici d'istanza diversi ha un gruppo per ogni
    combinazione (ad esempio uno per diametro).
  - Gli elementi senza Type Mark compaiono in fondo alla loro categoria.
  - Le colonne non si riordinano, perché l'ordine delle righe è la struttura.
  - **Ricerca:** guarda categoria, Type Mark, famiglia e tipo, modello e codici. Cercando un Type
    Mark si vede il gruppo con tutti i suoi codici; cercando un codice si vedono le
    righe di quel codice e quelle dei tipi che lo portano.
  - La descrizione segue le modifiche della scheda EPU senza ricalcolo.

## WBS

Il pulsante **WBS Levels...** (riga WBS nelle impostazioni) apre una finestra con 15
livelli: per ognuno si sceglie il parametro che lo rappresenta, dalla lista dei parametri
trovati su un campione di elementi del modello oppure scrivendone il nome. I livelli
lasciati vuoti si saltano, quindi si possono usare i livelli 1, 3 e 4 senza il 2. Lo
stesso parametro non può stare su due livelli.

- **Lettura dei valori:** per ogni elemento il valore si cerca sull'istanza, poi sul
  tipo, poi sul "genitore" e sul suo tipo. Il genitore è la famiglia che contiene
  un'annidata condivisa, oppure il canale o tubo rivestito da un isolante. Cambiando i
  livelli, i valori si rileggono dal modello.
- **Nella scheda Bill of quantities:** dopo la colonna Amount c'è una colonna per ogni
  livello WBS attivo, con il nome del parametro come intestazione; le voci hanno una riga per ogni
  combinazione di valori. Ogni livello WBS si può mettere sulle righe o sulle colonne
  (*Layout...*, vedi "Le schede della finestra").
  - La combinazione di un elemento arriva fino al suo **ultimo livello valorizzato**: se
    ha i livelli 1 e 2 ma non il 3, la colonna del livello 3 resta vuota.
  - Un livello vuoto in mezzo (1 e 3 pieni, 2 vuoto) compare come `(not set)`, e così
    anche un elemento senza alcun valore WBS.
  - La ricerca vale anche sui valori WBS.
- **Salvataggio:** i livelli si salvano nel file di progetto (`wbs`), quindi valgono per
  tutta la commessa. Se due utenti li cambiano insieme, vince l'ultimo che salva.
- **Excel:** il foglio Bill of quantities ha colonne WBS in testa e, per ogni
  combinazione di valori, **una riga in grassetto** con il totale (`Total`) seguita dalle
  sue voci (il raggruppamento scelto nella scheda non cambia l'export). Le righe di
  totale e il totale generale usano `SUBTOTAL(9, ...)`, che non
  conta due volte le righe di combinazione. I valori WBS sulle voci permettono di filtrare
  il foglio o di metterlo in pivot.

## Regole di misura

Le formule vengono dai fogli `XXXXEN_QTO_Ducts`, `_Pipes`, `_Flex Ducts`, `_Cable Trays` e
`_Conduits`. L = lunghezza dell'elemento; le dimensioni sono quelle del tipo/istanza in
Revit.

| Categoria | m | mq | kg |
| --- | --- | --- | --- |
| Ducts | L | perimetro × L | kg/mq della lamiera × mq |
| Flex Ducts | L | perimetro × L | n/a |
| Pipes | L | π·De × L | densità × π/4·(De² − Di²) × L |
| Flex Pipes | L | π·De × L | n/a |
| Cable Trays, Conduits | L (sempre) | n/a | n/a |
| Duct Insulation | L | π·(D + 2s) × L, rettangolari 2·(W + H + 4s) × L | n/a |
| Pipe Insulation | L | π·(De + 2s) × L | n/a |

- **Perimetro dei canali:** π·D per i circolari, 2·(W + H) per rettangolari e ovali.
  Gli ovali usano la formula rettangolare come nel foglio, quindi il perimetro risulta
  un po' più grande.
- **kg/mq della lamiera:** si sceglie per fascia di lato maggiore (il diametro per i
  circolari), con tabelle diverse per circolari e rettangolari. La fascia vale quando il
  lato è **minore** del limite. Default dal foglio Ducts:

  | Lato maggiore [mm] | < 300 | < 750 | < 1200 | < 2000 | ≥ 2000 |
  | --- | --- | --- | --- | --- | --- |
  | Rettangolari / ovali | 5,1 | 6,7 | 8,2 | 9,8 | 12 |
  | Circolari | 4,8 | 6,4 | 8 | 9,6 | 12 |

- **Densità dei tubi:** il foglio Pipes calcola il peso solo per i Type Mark CS01..CS04
  (acciaio nero, 7820 / 7950 / 7850 / 7900 kg/mc) e GS01..GS04 (zincato, 8300 / 8100 /
  8050 / 7900 kg/mc). Un tubo al kg con un Type Mark che non è in tabella finisce tra
  le anomalie. Il foglio usa 3,14 al posto di π: qui si usa π, quindi c'è uno scarto
  dello 0,05%.
- **mq dei tubi** (superficie esterna, ad esempio per la verniciatura) non c'è nel
  foglio Pipes: è un'aggiunta.
- **Isolanti:** si misurano con la sezione dell'elemento che rivestono. Gli isolanti su
  raccordi e accessori non hanno una sezione misurabile: come nei fogli, non si
  computano e li copre la maggiorazione. Compaiono comunque tra le anomalie.
- **Maggiorazione per raccordi e sfridi:** si applica a m, mq e kg. Default 30% per
  Ducts, Flex Ducts, Pipes, Flex Pipes e per i due isolanti; 0% per Cable Trays e
  Conduits, che nei fogli non hanno maggiorazione. I raccordi dei conduit (Conduit
  Fittings) sono contati a pezzo.

**L'unità della voce decide la misura.** Uno stesso canale può portare un codice al kg
(lamiera) e uno al mq (verniciatura, isolamento). Se la voce non ha un'unità ammessa per
la categoria (vuota, `cad`, `kg` su un flessibile...) il codice non viene computato e
compare tra le anomalie. Passerelle e conduit fanno eccezione: si computano sempre in m,
e un'unità diversa viene solo segnalata.

Le tre tabelle si modificano nella scheda **Rules** e si salvano nel file di progetto,
quindi valgono per tutta la commessa. *Restore defaults* rimette i valori dei fogli. Le
regole sono riportate anche nel foglio Rules dell'export Excel.

### Override della maggiorazione

Una singola voce del Bill of quantities (coppia Type Mark e codice) può avere una
maggiorazione diversa da quella della sua categoria:

- **Come si imposta:** si selezionano una o più voci nella scheda Bill of quantities e si
  usa il tasto destro (*Allowance override...*) oppure il pulsante **Allowance
  Override...** accanto alla ricerca. La finestra mostra le voci scelte con la
  maggiorazione della loro categoria e l'eventuale override, e chiede la percentuale
  (virgola o punto decimale, 0 = nessuna maggiorazione).
- **Come si toglie:** *Remove Override* nella stessa finestra, oppure *Remove allowance
  override* dal tasto destro. La voce torna alla maggiorazione della categoria.
- **Cosa fa:** la percentuale **sostituisce** quella della categoria sugli elementi
  lineari della voce. Le categorie a pezzo restano a 1 per istanza: una voce solo a pezzo
  non ammette override e, in una selezione mista, viene saltata.
- **WBS:** l'override vale per la voce in tutte le combinazioni WBS in cui compare, quindi
  non si perde cambiando i livelli.
- **Come si riconosce:** la cella Quantity della voce è colorata (giallo) e il tooltip
  riporta l'override e la maggiorazione di categoria. Nell'export Excel la quantità ha lo
  stesso colore e il foglio Rules elenca gli override applicati.
- **Salvataggio:** gli override vanno nel file di progetto (`allowance_overrides`), voce
  per voce come i campi del listino. Un override su una voce che non compare più nel
  computo resta nel file e torna attivo se la voce ricompare.

## Dove finiscono i dati

| Livello | File | Chi lo scrive |
| --- | --- | --- |
| Listino comune | `.json` su percorso di rete | il tool, con l'editor del listino (pulsante Edit...) |
| Listino comune (alternativa) | `.xlsx` / `.xlsm` / `.csv` su percorso di rete | chi gestisce l'elenco prezzi, in Excel. Il tool lo legge soltanto |
| Progetto | `<Modello>_MEPQTO.json` accanto al modello centrale | il tool, con il pulsante Save |

### Come si fondono listino e progetto

I due file non si scambiano dati: la finestra li legge entrambi e per ogni codice
costruisce la voce **campo per campo** (`merge_item` in `mepqto_store.py`):

1. si parte da una voce vuota;
2. si copiano i campi non vuoti del listino;
3. sopra si copiano i campi non vuoti del progetto.

```
Listino:  IM.01.001  description="Diffusore 600x600"  unit="cad"  price=125.5  epu_item=""
Progetto: IM.01.001  price=118.0  epu_item="12"
---------------------------------------------------------------------------------
Finestra: IM.01.001  description="Diffusore 600x600"  unit="cad"  price=118.0  epu_item="12"
```

- **Quali codici:** solo quelli usati nel modello. Le voci del listino che il modello
  non usa non compaiono; un codice del modello assente dal listino compare con i campi
  vuoti, in rosso, con l'anomalia "Code not in the price list".
- **Modifiche nella scheda EPU:** vanno solo nel file di progetto, mai nel
  listino comune, e valgono per quella commessa. Le modifiche al listino comune si fanno
  con l'editor (Edit...) e valgono per tutte le commesse che lo usano.
- **Campo svuotato:** torna al valore del listino; al salvataggio i campi vuoti non si
  scrivono. Di conseguenza non si può forzare un campo vuoto se il listino lo ha.
- **Un campo corretto nel progetto resta fisso:** se il listino comune viene aggiornato
  dopo, la commessa continua a usare il suo valore. Per riallinearla si svuota la cella.
  La finestra non indica quali campi vengono dal progetto.
- Le correzioni di codici che il modello non usa più restano nel file di progetto.

### Listino JSON e editor

Il pulsante **Edit...** accanto al listino apre l'editor:

- **Listino in uso:** se è un `.json` si apre direttamente; se è un Excel o un CSV
  viene importato in un listino nuovo, da salvare come `.json`.
- **Griglia:** Chapter, Subchapter, EPU item No., Reference price book, Price book code,
  Short Description, Description, Unit (tendina), Unit price, con ricerca, Add Row, Duplicate e Remove
  Selected. Il codice prezzario deve essere unico e non vuoto; il prezzo accetta la
  virgola decimale.
- **File:** New / Open / Save / Save As, nome e valuta del listino.
- **Import:** legge `.xlsx`, `.xlsm`, `.csv` o un altro listino `.json`. Se alcuni codici
  ci sono già, chiede se sovrascriverli o aggiungere solo i nuovi.
- **Dopo il salvataggio:** se il file è il listino in uso, la finestra del computo lo
  ricarica; altrimenti chiede se usarlo.

Struttura del file:

```json
{
  "format": "ESA_MEPQTO_PriceList",
  "version": 1,
  "name": "Elenco prezzi ESA 2026",
  "currency": "EUR",
  "updated_by": "andreapa",
  "updated_at": "2026-10-02 16:30:00",
  "items": {
    "IM.01.001": {
      "chapter": "HVAC",
      "subchapter": "Diffusione",
      "epu_item": "12",
      "price_book": "Prezzario Regione Lombardia 2026",
      "short_description": "Diffusore quadrato 600x600",
      "description": "Diffusore a soffitto quadrato 600x600",
      "unit": "cad",
      "price": 125.5
    }
  }
}
```

- `items` è indicizzato per codice prezzario, quindi un codice non può comparire due
  volte. Le voci hanno gli stessi campi di `items` nel file di progetto.
- `epu_item`, `price_book` e `short_description` sono facoltativi: un listino salvato
  prima della loro introduzione si legge senza modifiche, con quei campi vuoti.
- `price` è un numero (oppure `null`); `unit` è normalizzata come per l'Excel (`m2` →
  `mq`).
- Il tool scrive le lettere accentate come `à`; in lettura accetta anche UTF-8
  normale. Un file con un altro `format` viene rifiutato.

**Salvataggio concorrente:** il listino sta su un percorso di rete e può essere aperto da
più persone. Al salvataggio si scrivono solo i codici aggiunti, cambiati o cancellati.
Se intanto qualcun altro ha salvato, il file si rilegge, le modifiche locali si
riapplicano sopra e l'editor mostra il risultato. Su uno stesso codice vince l'ultimo che
salva.

### Listino Excel / CSV

È l'elenco prezzi unitari (EPU) della commessa, preparato in Excel. **Foglio e colonne si
scelgono** nella finestra **Price list columns**, che si apre:

- quando si sceglie un listino `.xlsx`, `.xlsm` o `.csv` con *Browse...* (annullando, il
  listino non cambia);
- dal pulsante **Columns...** accanto al listino, attivo solo per i listini Excel / CSV;
- nell'editor del listino, prima di importare un Excel o un CSV (dalla finestra del
  computo arriva la mappatura già scelta).

La finestra ha:

- **Sheet:** i fogli del file, nell'ordine; per un CSV non c'è scelta.
- **Columns:** una tendina per campo con le colonne del foglio, mostrate con la lettera e
  l'intestazione (es. `C - Descrizione`). Il codice prezzo è obbligatorio, gli altri campi
  possono restare `(none)`. Una colonna non può servire a due campi.

  | Campo | Va nella scheda EPU in |
  | --- | --- |
  | Price book code | la chiave: si confronta con `Price book code` |
  | Short Description, Description | `Short Description`, `Description` |
  | Unit | `Unit` (`n`, `nr`, `cad`... ricondotte alla tendina, vedi sopra) |
  | Unit price | `Unit price` (numero, oppure testo con virgola decimale o `€`, `1.234,50`) |
  | Chapter, Subchapter, EPU item No., Reference price book | le colonne omonime |

- **Preview:** le prime 40 righe del foglio scelto, con le lettere delle colonne.
- *Detect from Headers* mappa le colonne con un'intestazione nota (Codice, Descrizione,
  Descrizione breve, UM, Prezzo unitario, Capitolo, N. articolo EPU...); *A to E* rimette
  la disposizione fissa A codice, B descrizione sintetica, C descrizione, D unità, E
  prezzo (F..I capitolo, sottocapitolo, n. articolo EPU, prezzario).
- **Proposta iniziale:** la mappatura già salvata per quel file; altrimenti quella
  ricavata dalle intestazioni; altrimenti A..E. Cambiando foglio si ripropone.

**Dove si salva.** Foglio e colonne vanno nel **file di progetto** insieme al percorso
del listino (`price_list_layout`, colonne come lettere: `{"sheet": "EPU", "columns":
{"code": "B", "description": "D", ...}}`), quindi valgono per tutto il team della
commessa; si ricordano anche nella config pyRevit (`last_price_list_layout`) per i
modelli senza file di progetto. *Reload* rilegge il listino con la stessa mappatura. Un
file di progetto senza mappatura legge il primo foglio con le colonne A..E, come le
versioni precedenti.

- **Righe saltate:** quelle senza codice, la riga di intestazione (codice "Codice",
  "Code", "Articolo"...) e le righe con il solo codice e nessun altro campo mappato, come
  titoli e intestazioni di capitolo. Un'intestazione con un nome diverso resta una voce
  che nessun codice del modello trova.
- **Codice ripetuto:** vale la prima riga; le altre si contano come duplicati (le note
  della finestra ne riportano il numero).
- **Confronto dei codici:** esatto, maiuscole comprese. Gli spazi ai lati si tolgono. Un
  codice numerico scritto come numero in Excel si legge senza decimali (`115004`, non
  `115004.0`).
- **Prezzo:** una cella non numerica lascia la voce senza prezzo.
- **CSV:** il separatore (`;`, `,` o tab) si riconosce da solo e i prezzi accettano la
  virgola decimale (`1.234,56`).
- **Export:** il foglio EPU dell'export Excel ha le colonne nell'ordine A..E, poi F..I,
  quindi si rilegge con *A to E*.
- **Errori:** un foglio che non c'è più o una colonna del codice senza codici lasciano il
  listino non caricato, con il motivo nella scheda Issues ("Price list not loaded"):
  si corregge con *Columns...*.

### File di progetto

Per i modelli condivisi il file sta accanto al **centrale**, così tutti gli utenti della
commessa lavorano sullo stesso file. Per i modelli cloud (ACC / BIM 360) o mai salvati il
percorso si sceglie con *Change...* e viene ricordato nella config di pyRevit
(`ESA_MEPQTO`).

Contenuto:

- `price_list_path`: percorso del listino. Se manca, si usa l'ultimo listino ricordato
  nella config di pyRevit;
- `price_list_layout`: foglio e colonne del listino Excel / CSV (vedi "Listino Excel /
  CSV"); si salva con il percorso;
- `items`: voci corrette, solo i campi modificati (`chapter`, `subchapter`, `epu_item`,
  `price_book`, `short_description`, `description`, `unit`, `price`);
- `rules`: regole di misura, solo se modificate o salvate dalla scheda Rules;
- `wbs`: livelli WBS;
- `parameters`: mappa dei parametri (assente = default). Le chiavi sono `piece_codes`
  (codici di tipo), `linear_codes` (codici d'istanza, tutte le categorie) e
  `linear_include` (Sì/No d'istanza, tutte le categorie): i nomi restano quelli storici
  per non rompere i file già salvati;
- `manual_items`: voci non modellate della scheda Manual items
  (`{"type_mark": "MAN1", "code": "M.001", "quantity": 2.5, "wbs": {"Zona": "Z2"}}`);
- `allowance_overrides`: override della maggiorazione sulle voci del computo, una riga
  per voce (`{"type_mark": "CS01", "code": "IM.02.010", "allowance": 0.15}`, maggiorazione
  come frazione);
- `updated_by`, `updated_at`: utente e data dell'ultimo salvataggio.

**Salvataggio concorrente.** Il tool ricorda quali campi hai toccato. Se al salvataggio
il file è stato modificato da qualcun altro dopo che l'hai aperto, lo rilegge e riapplica
sopra solo le tue modifiche. Su una stessa voce e campo (o sullo stesso override di
maggiorazione) vince l'ultimo che salva; il resto non si perde. Regole, livelli WBS e
mappa dei parametri si salvano in blocco: se due utenti li cambiano, vince l'ultimo.

Un file di progetto che non è JSON valido **non viene mai sovrascritto**: il tool avvisa
e chiede di sceglierne un altro.

### File salvati con versioni precedenti

Listini e file di progetto già esistenti si aprono senza conversioni:

- un listino senza `epu_item`, `price_book` o `short_description` si legge con quei campi
  vuoti;
- **listini Excel / CSV preparati per la versione precedente**, che riconosceva le colonne
  dal nome dell'intestazione: ora le colonne si leggono per posizione (A codice, B
  descrizione sintetica, C descrizione, D unità, E prezzo). Un foglio con un altro ordine
  va riordinato, altrimenti i campi finiscono nelle colonne sbagliate;
- nella mappa dei parametri la chiave `piece_include` (vecchio Sì/No di tipo
  `e_DAT_BOQ_t`) non viene più letta. Resta nel file finché la mappa non si cambia con
  Parameters..., poi sparisce.

**Attenzione al Sì/No di tipo:** gli elementi delle categorie a pezzo esclusi con
`e_DAT_BOQ_t` = No sul tipo ora **si contano**, a meno che le loro istanze non abbiano
`e_DAT_BOQ_i` = No. Nei modelli che usavano il parametro di tipo va riportato il No sulle
istanze prima di fare il computo.

## Anomalie (scheda Issues)

| Anomalia | Significato |
| --- | --- |
| No Type Mark | istanze il cui tipo non ha Type Mark: **escluse dal computo** |
| No price code | Type Mark senza nessun codice (né di tipo né d'istanza), oppure elementi lineari senza alcun codice d'istanza (solo quelli, **non computati**) |
| Inconsistent price codes | istanze a pezzo con lo stesso Type Mark ma codici di tipo diversi: ogni codice conta solo sulle istanze che lo portano. I codici d'istanza non contano per questo controllo (possono cambiare da elemento a elemento), quindi non vale per le lineari |
| Code not in the price list | codice usato nel computo che non c'è nella colonna A del listino caricato: la riga della scheda EPU è rossa |
| Code without description | codice presente nel listino (o computato senza listino) che non ha descrizione né nel listino né nel progetto |
| Price code parameters not found | nessun elemento contato ha i parametri dei codici impostati con Parameters... |
| Excluded from the bill | elementi tolti dal Sì/No di inclusione (No sull'istanza, tutte le categorie) |
| Price list not loaded | il listino non si legge (file mancante, aperto e bloccato, nessun codice in colonna A) |
| Linked model not read | link scelto ma non letto: non caricato, oppure senza una fase con il nome di quella scelta |
| Unit not valid for the category | la voce non ha un'unità ammessa per la categoria (es. canale con unità vuota o `cad`): **non computato** |
| Unit replaced | passerella o conduit con una voce non in m: computato comunque in m |
| No pipe density | tubo con voce al kg e Type Mark assente dalla tabella densità: **non computato** |
| Dimensions missing | lunghezza, sezione o spessore isolante mancanti per l'unità richiesta: **non computato** |
| Insulation on fittings not measured | isolanti su raccordi o accessori, coperti dalla maggiorazione |

Gli ElementId degli elementi coinvolti stanno nel foglio Issues dell'export Excel, con il
nome del link davanti per gli elementi dei modelli linkati. Alla
chiusura della finestra non viene generato alcun report pyRevit.

## Export Excel

L'export scrive il file `.xlsx` direttamente, senza bisogno di Excel installato. Il nome
di default è `YYMMDD_HHMMSS_MEPQTO_<Modello>.xlsx`. I fogli sono gli stessi delle schede:

- **EPU** (in precedenza il foglio si chiamava *Price list*): codice prezzario,
  descrizione breve, descrizione, unità, prezzo unitario (colonne A..E, come il listino
  Excel), poi capitolo, sottocapitolo, n. articolo EPU, prezzario di riferimento.
- **Bill of quantities**: importo = `Quantità × Prezzo` come formula, totali per
  combinazione WBS e totale generale con `SUBTOTAL(9, ...)`, così il file resta corretto
  se si corregge un prezzo in Excel. L'intestazione elenca i modelli letti e i workset non letti
  esclusi. Stesse colonne della scheda: con la WBS attiva le
  prime sono i livelli, poi Type Mark, n. articolo EPU e codice prezzario. Le quantità
  delle voci con override della maggiorazione sono colorate, con una nota in testa al
  foglio.
- **Type Marks**: la stessa matrice della scheda, con le righe di gruppo in grassetto su
  fondo chiaro. Le righe dei codici lasciano vuote le colonne del gruppo, quindi il filtro
  automatico su Category o Type Mark mostra solo le righe di gruppo.
- **Rules**: maggiorazioni, override sulle voci, kg/mq e densità usati per il calcolo, e
  la mappa dei parametri del modello.
- **Issues**: le anomalie con gli ElementId.

## Contenuto della cartella

| File | Ruolo |
| --- | --- |
| `MEPQTO_script.py` | entry point: controlli sul documento e apertura della finestra |
| `mepqto_model.py` | categorie, elenco dei link, raccolta dal modello e dai link (geometria inclusa), aggregazione, computo, riepilogo Type Mark, anomalie |
| `mepqto_rules.py` | formule di misura delle categorie lineari, maggiorazioni, tabelle kg/mq e densità |
| `mepqto_store.py` | lettura del listino, unità, file di progetto, fusione delle voci |
| `mepqto_xlsx.py` | writer `.xlsx` e fogli del computo |
| `mepqto_ui.py` | finestra (griglie su `DataTable`, salvataggio, export) |
| `MEPQTO_form.xaml` | layout della finestra |
| `mepqto_grid_filter.py` | filtri per colonna in stile Excel sulle griglie legate a una `DataTable` |
| `mepqto_layout_ui.py` | finestra Layout del Bill of quantities: livelli sulle righe e livelli WBS sulle colonne |
| `MEPQTO_layout.xaml` | layout della finestra Layout |
| `mepqto_columns_ui.py` | finestra Price list columns: foglio e colonne di un listino Excel / CSV, con anteprima |
| `MEPQTO_columns.xaml` | layout della finestra Price list columns |
| `mepqto_scope_ui.py` | finestra di scelta dei modelli (link), dei workset e delle categorie da leggere, con i set di categorie e di workset |
| `MEPQTO_scope.xaml` | layout della finestra di scelta |
| `mepqto_wbs_ui.py` | finestra di scelta dei livelli WBS |
| `MEPQTO_wbs.xaml` | layout della finestra WBS |
| `mepqto_params_ui.py` | finestra di mappatura dei parametri del modello |
| `MEPQTO_params.xaml` | layout della finestra dei parametri |
| `mepqto_allowance_ui.py` | finestra dell'override della maggiorazione sulle voci del computo |
| `MEPQTO_allowance.xaml` | layout della finestra dell'override |
| `mepqto_pricelist_ui.py` | editor del listino JSON |
| `MEPQTO_pricelist.xaml` | layout dell'editor del listino |
