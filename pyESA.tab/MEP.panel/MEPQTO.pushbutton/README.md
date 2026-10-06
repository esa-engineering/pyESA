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

## Come si calcola

1. Si raccolgono le istanze delle categorie che esistono nella fase scelta:
   - con stato `New`, solo quelle create nella fase;
   - con `New + Existing`, anche quelle che vengono dalle fasi precedenti.

   Le demolite non contano mai. Con l'opzione attiva (default) si escludono anche le
   istanze delle opzioni di progetto secondarie. Si tolgono poi gli elementi che il
   **Sì/No di inclusione** mette a No (vedi "Parametri del modello").
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
   - la misura nell'**unità della voce** (scheda Price list) per quelle lineari,
     maggiorata della percentuale della categoria per raccordi e sfridi.

   Gli elementi senza Type Mark sono esclusi. L'importo è la quantità per il prezzo
   unitario.

Le famiglie annidate condivise sono istanze vere e quindi si contano. Nel riepilogo Type
Mark hanno una riga separata (`Nested = Yes`), così un doppio conteggio padre/figlio si
vede subito.

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
| Price list | capitolo, sottocapitolo, n. articolo EPU, prezzario di riferimento, codice prezzario, descrizione, unità, prezzo unitario dei codici usati nel modello | sì: tutto tranne il codice prezzario (unità da tendina) |
| Bill of quantities | Type Mark, n. articolo EPU, codice prezzario, descrizione, unità, quantità, prezzo unitario, importo, con il totale sotto; una riga per Type Mark e codice; con la WBS, colonne dei livelli in testa e una riga di totale per combinazione | no, si compila da sola; si può solo dare a una voce una maggiorazione propria |
| Type Marks | categoria, Type Mark, famiglia e tipo, annidata sì/no, poi le coppie codice / descrizione di tipo 1..10 e d'istanza 1..10 | no |
| Rules | maggiorazioni per categoria, kg/mq della lamiera dei canali, densità dei tubi | sì |
| Issues | anomalie (vedi sotto) | no |

- **Bill of quantities** ha una riga per ogni coppia Type Mark e codice. Un codice usato
  da più Type Mark compare una volta per ognuno, con la sua quantità, così si vede da
  quale tipo arriva. Le righe sono ordinate per Type Mark e poi per codice. Una voce che
  non si può misurare (unità o dimensioni mancanti) resta in elenco con quantità 0 e la
  sua anomalia. Il totale per codice resta quello della somma delle righe. Una voce può
  avere una maggiorazione propria (vedi "Override della maggiorazione").
- **Price list** mostra tutti i codici usati nel modello nella fase scelta, anche quelli
  delle categorie non spuntate. Il computo invece segue le categorie spuntate. Le righe
  sono ordinate per capitolo, sottocapitolo e codice; i codici senza capitolo stanno in
  fondo. Una modifica di capitolo non riordina subito la griglia: l'ordine si aggiorna al
  ricalcolo successivo (cambio di fase o categorie, riapertura). Cliccando un'intestazione
  si può comunque riordinare per quella colonna. La ricerca guarda capitolo,
  sottocapitolo, n. articolo EPU, prezzario, codice e descrizione.
- **Colonne della scheda Price list** (intestazioni in inglese nella finestra):
  - `EPU item No.`: numero dell'articolo nell'elenco prezzi unitari (EPU) di progetto;
  - `Reference price book`: prezzario da cui viene la voce (es. Prezzario Regione
    Lombardia 2026);
  - `Price book code`: il codice della voce nel prezzario, cioè il valore dei parametri
    `e_DAT_PriceCode_*` del modello. È la chiave della voce e non si modifica qui.
- **L'unità** si sceglie da una tendina con `cad`, `m`, `kg`, `mq`, `mc`. La voce vuota
  torna al valore del listino.
- **Le unità lette dal listino** vengono ricondotte ai valori della tendina quando sono
  varianti note: `cad.`, `nr`, `pz` → `cad`; `ml` → `m`; `m2`, `m²` → `mq`; `m3`, `m³` →
  `mc`. Le altre restano come sono e compaiono in fondo alla tendina.
- **Nel riepilogo Type Marks** la colonna "Type price code N" corrisponde all'N-esimo
  parametro di tipo impostato con Parameters... (solo categorie a pezzo), la colonna
  "Instance price code N" all'N-esimo parametro d'istanza (tutte le categorie). Un buco
  (codice 1 vuoto, codice 3 pieno) resta visibile. Un tipo le cui istanze portano codici
  d'istanza diversi compare una volta per ogni combinazione (ad esempio una riga per
  diametro). Le coppie che nessuna riga usa sono nascoste, separatamente per tipo e
  istanza: le colonne d'istanza compaiono solo se almeno una riga ha un codice d'istanza
  (di norma le categorie lineari). Gli elementi senza Type Mark compaiono in fondo alla
  loro categoria.

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
- **Raggruppamento:** nella scheda Bill of quantities le prime colonne sono i livelli
  WBS attivi, una per livello, con il nome del parametro come intestazione.
  - Ogni combinazione di valori ha **una sola riga in grassetto** con i valori dei
    livelli e il totale dell'importo (`Total`), seguita dalle sue voci. Anche le voci
    riportano i valori WBS.
  - La combinazione di un elemento arriva fino al suo **ultimo livello valorizzato**: se
    ha i livelli 1 e 2 ma non il 3, la colonna del livello 3 resta vuota.
  - Un livello vuoto in mezzo (1 e 3 pieni, 2 vuoto) compare come `(not set)`, e così
    anche un elemento senza alcun valore WBS.
  - Con la WBS attiva le colonne non si riordinano, perché l'ordine delle righe è la
    struttura.
  - La ricerca vale anche sui valori WBS e lascia visibili le righe di totale.
- **Salvataggio:** i livelli si salvano nel file di progetto (`wbs`), quindi valgono per
  tutta la commessa. Se due utenti li cambiano insieme, vince l'ultimo che salva.
- **Excel:** il foglio Bill of quantities ha la stessa struttura, colonne WBS in testa
  comprese. Le righe di totale e il totale generale usano `SUBTOTAL(9, ...)`, che non
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
  vuoti e l'anomalia "Code without description".
- **Modifiche nella scheda Price list:** vanno solo nel file di progetto, mai nel
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
  Description, Unit (tendina), Unit price, con ricerca, Add Row, Duplicate e Remove
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
      "description": "Diffusore a soffitto quadrato 600x600",
      "unit": "cad",
      "price": 125.5
    }
  }
}
```

- `items` è indicizzato per codice prezzario, quindi un codice non può comparire due
  volte. Le voci hanno gli stessi campi di `items` nel file di progetto.
- `epu_item` e `price_book` sono facoltativi: un listino salvato prima della loro
  introduzione si legge senza modifiche, con i due campi vuoti.
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

Si legge il primo foglio. Nelle prime 20 righe deve esserci una riga di intestazione che
contenga almeno la colonna del codice. Le intestazioni riconosciute (maiuscole, spazi e
punteggiatura non contano) sono:

| Campo | Intestazioni accettate |
| --- | --- |
| Codice prezzario | Code, Codice, Cod, Price code, Item code, Codice voce, Articolo, Tariffa, Codice prezzario, Price book code |
| N. articolo EPU | N. articolo EPU, Nr. articolo EPU, Numero articolo EPU, Articolo EPU, N. EPU, EPU, EPU item No. |
| Prezzario di riferimento | Prezzario di riferimento, Prezzario, Prezziario, Reference price book, Price book |
| Descrizione | Description, Descrizione, Desc, Designazione, Designazione dei lavori |
| Unità | Unit, UM, U.M., UoM, Unità di misura |
| Prezzo | Price, Prezzo, Unit price, Prezzo unitario, P.U., Euro |
| Capitolo | Chapter, Capitolo, WBS, Section, Sezione |
| Sottocapitolo | Subchapter, Sottocapitolo, Subcapitolo, Subsection, Sottosezione |

Capitolo, sottocapitolo, n. articolo EPU e prezzario sono facoltativi; le altre colonne
vengono ignorate. Se un codice è ripetuto, vale la prima riga e le altre si contano come
duplicati. Nei CSV il separatore (`;`, `,` o tab) si riconosce da solo e i prezzi
accettano la virgola decimale (`1.234,56`). Le intestazioni del foglio Price list
dell'export Excel sono riconosciute, quindi un export si può reimportare nell'editor.

### File di progetto

Per i modelli condivisi il file sta accanto al **centrale**, così tutti gli utenti della
commessa lavorano sullo stesso file. Per i modelli cloud (ACC / BIM 360) o mai salvati il
percorso si sceglie con *Change...* e viene ricordato nella config di pyRevit
(`ESA_MEPQTO`).

Contenuto:

- `price_list_path`: percorso del listino. Se manca, si usa l'ultimo listino ricordato
  nella config di pyRevit;
- `items`: voci corrette, solo i campi modificati (`chapter`, `subchapter`, `epu_item`,
  `price_book`, `description`, `unit`, `price`);
- `rules`: regole di misura, solo se modificate o salvate dalla scheda Rules;
- `wbs`: livelli WBS;
- `parameters`: mappa dei parametri (assente = default). Le chiavi sono `piece_codes`
  (codici di tipo), `linear_codes` (codici d'istanza, tutte le categorie) e
  `linear_include` (Sì/No d'istanza, tutte le categorie): i nomi restano quelli storici
  per non rompere i file già salvati;
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

- un listino senza `epu_item` e `price_book` si legge con i due campi vuoti;
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
| Code without description | codice usato nel computo che non ha descrizione né nel listino né nel progetto |
| Price code parameters not found | nessun elemento contato ha i parametri dei codici impostati con Parameters... |
| Excluded from the bill | elementi tolti dal Sì/No di inclusione (No sull'istanza, tutte le categorie) |
| Price list not loaded | il listino non si legge (file mancante, aperto e bloccato, intestazione non trovata) |
| Unit not valid for the category | la voce non ha un'unità ammessa per la categoria (es. canale con unità vuota o `cad`): **non computato** |
| Unit replaced | passerella o conduit con una voce non in m: computato comunque in m |
| No pipe density | tubo con voce al kg e Type Mark assente dalla tabella densità: **non computato** |
| Dimensions missing | lunghezza, sezione o spessore isolante mancanti per l'unità richiesta: **non computato** |
| Insulation on fittings not measured | isolanti su raccordi o accessori, coperti dalla maggiorazione |

Gli ElementId degli elementi coinvolti stanno nel foglio Issues dell'export Excel. Alla
chiusura della finestra non viene generato alcun report pyRevit.

## Export Excel

L'export scrive il file `.xlsx` direttamente, senza bisogno di Excel installato. Il nome
di default è `YYMMDD_HHMMSS_MEPQTO_<Modello>.xlsx`. I fogli sono gli stessi delle schede:

- **Price list**: capitolo, sottocapitolo, n. articolo EPU, prezzario di riferimento,
  codice prezzario, descrizione, unità, prezzo unitario.
- **Bill of quantities**: importo = `Quantità × Prezzo` come formula, totali per
  combinazione WBS e totale generale con `SUBTOTAL(9, ...)`, così il file resta corretto
  se si corregge un prezzo in Excel. Stesse colonne della scheda: con la WBS attiva le
  prime sono i livelli, poi Type Mark, n. articolo EPU e codice prezzario. Le quantità
  delle voci con override della maggiorazione sono colorate, con una nota in testa al
  foglio.
- **Type Marks**: le stesse colonne della scheda, fino all'ultima coppia usata di tipo e
  poi d'istanza.
- **Rules**: maggiorazioni, override sulle voci, kg/mq e densità usati per il calcolo, e
  la mappa dei parametri del modello.
- **Issues**: le anomalie con gli ElementId.

## Contenuto della cartella

| File | Ruolo |
| --- | --- |
| `MEPQTO_script.py` | entry point: controlli sul documento e apertura della finestra |
| `mepqto_model.py` | categorie, raccolta dal modello (geometria inclusa), aggregazione, computo, riepilogo Type Mark, anomalie |
| `mepqto_rules.py` | formule di misura delle categorie lineari, maggiorazioni, tabelle kg/mq e densità |
| `mepqto_store.py` | lettura del listino, unità, file di progetto, fusione delle voci |
| `mepqto_xlsx.py` | writer `.xlsx` e fogli del computo |
| `mepqto_ui.py` | finestra (griglie su `DataTable`, salvataggio, export) |
| `MEPQTO_form.xaml` | layout della finestra |
| `mepqto_wbs_ui.py` | finestra di scelta dei livelli WBS |
| `MEPQTO_wbs.xaml` | layout della finestra WBS |
| `mepqto_params_ui.py` | finestra di mappatura dei parametri del modello |
| `MEPQTO_params.xaml` | layout della finestra dei parametri |
| `mepqto_allowance_ui.py` | finestra dell'override della maggiorazione sulle voci del computo |
| `MEPQTO_allowance.xaml` | layout della finestra dell'override |
| `mepqto_pricelist_ui.py` | editor del listino JSON |
| `MEPQTO_pricelist.xaml` | layout dell'editor del listino |
