# Family Naming - note di sviluppo

Registro delle scelte fatte costruendo il tool, con il motivo di ognuna e le
alternative scartate. Serve a due cose: non rimettere in discussione fra sei
mesi decisioni gia' prese, e avere il contesto necessario a reingerire un
Excel modificato senza ricostruire tutto il ragionamento da capo.

**Cosa non contiene.** La fonte dei codici resta l'Excel: qui non si
duplicano tabelle. Il funzionamento del tool dal punto di vista di chi lo usa
sta in `README.md`, che e' un documento complementare a questo. Dove serve, i
due si rimandano a vicenda.

Lavoro svolto dal 15 al 17 settembre 2026. Ambiente di riferimento dichiarato dal
repo: Revit 2026.4 / pyRevit 5.3.1, IronPython 2.7.

---

## 1. Architettura

Sei file, piu' i tre `.xlsx` come sorgente documentale.

| File | Ruolo | Chi lo modifica |
|---|---|---|
| `FamilyNaming_script.py` | Revit: selezione, rilevamento, lettura parametri, transazione, report | sviluppatore |
| `FamilyNaming_map.py` | Mappa dichiarativa: schede, tabelle, blocchi, categorie, filtro note | anche non sviluppatore |
| `FamilyNaming_rules.py` | Motore: Type Mark, blocchi dimensionali, composizione, validazione | sviluppatore |
| `FamilyNaming_ui.py` | Finestra, campi costruiti a runtime | sviluppatore |
| `FamilyNaming.xaml` | Layout | sviluppatore |
| `FamilyNaming_data.py` | **Generato.** Tabelle e note congelate | nessuno, si rigenera |
| `gen_data.py` | Rigenera il modulo dati dagli Excel | si lancia, non si modifica |

### Perche' i moduli sono separati

Il repo non ha librerie condivise e ogni script e' autosufficiente. Qui la
separazione serve comunque, per un motivo preciso: `FamilyNaming_map.py` e'
pensato per essere corretto da chi conosce la nomenclatura ma non programma,
mentre il resto e' logica. Mescolarli avrebbe reso la mappa illeggibile.

`FamilyNaming_rules.py` e' il quinto file rispetto ai quattro inizialmente
concordati. Lo scostamento e' stato dichiarato e accettato: 200 righe di
logica di composizione dentro 600 righe di dati dichiarativi avrebbero
vanificato la separazione appena descritta.

---

## 2. I dati: congelati, non letti a runtime

**Deciso.** Le tabelle del foglio DV diventano dizionari Python in
`FamilyNaming_data.py`, generati da `gen_data.py`. Il tool non apre mai gli
`.xlsx`.

**Scartato:** lettura degli `.xlsx` a ogni avvio, con cache. Era la mia
proposta iniziale perche' i BIM manager avrebbero visto le modifiche senza
passare da uno sviluppatore. Scartata su indicazione esplicita.

**Conseguenza da ricordare.** Dopo ogni modifica agli Excel serve rigenerare,
e **il file generato va committato**: e' l'unico modo di avere un diff
leggibile, visto che un `.xlsx` e' uno zip binario che git non sa confrontare.

**Scelta tecnica del generatore.** Legge le Excel Table nominate
(`Tbl_DR_G1`, `Tbl_BW_G2`, `Tbl_STR_G1`...) tramite `openpyxl`, non le
coordinate delle celle. E' quello che rende gratuito aggiungere o togliere
righe dentro una tabella esistente. La convenzione dei nomi tabella e' quindi
un contratto: `Tbl_<CODICE>_Cat`, `_G1`, `_G2`.

---

## 3. Il Type Mark

La forma e' sempre `AA-XXX`: due lettere, trattino, tre cifre.

| Caso | Prefisso | Cifre | Esempio |
|---|---|---|---|
| Caricabili con TM Code (DR, WN, CP, CM, PK, EN, PL, MS) | codice categoria | cifra del Group2 + 2 di sequenziale | `DR-101` |
| Caricabili senza TM Code (FN, FS, CK, PF, SE, SI, GM) | codice **Group2** | 3 di sequenziale | `SE-001` |
| System e Structural | codice categoria | cifra del **Group1** + 2 di sequenziale | `BW-101` |
| Group1 = Other | codice materiale (`Tbl_Mat`) | 3 di sequenziale | `WD-001` |
| Structural Framing con Use | codice Use | 3 di sequenziale | `SK-001` |

### L'inversione Group1 / Group2

Sui caricabili la cifra viene dal **Group2**, su System e Structural viene dal
**Group1**. Non e' un errore di trascrizione: e' cosi' negli Excel, dove la
colonna `TM Code` sta accanto a tabelle diverse nei due casi. E' il punto in
cui un errore si propagherebbe ovunque, quindi va riverificato a ogni
modifica delle tabelle.

Cambia di conseguenza anche l'ambito di unicita' del sequenziale.

### La ricerca del sequenziale

**Deciso.** Si cerca **per prefisso sull'intero documento corrente**, si legge
dal parametro Type Mark dei tipi, compresi quelli non piazzati, e si prende il
massimo piu' uno.

**Scartato:** ricerca per categoria, e riempimento dei buchi nella sequenza.
La ricerca per categoria produrrebbe Type Mark duplicati sulle sette categorie
a Group1 condiviso, dove il codice di categoria non compare nel Type Mark.
Riempire i buchi invece crea collisioni con tipi cancellati e poi ripristinati.

### Conseguenza: nove prefissi sono condivisi

Poiche' sulle sette categorie a Group1 condiviso il codice categoria sparisce
dal Type Mark, e gli stessi codici Group2 ricorrono su categorie diverse, oggi
**nove prefissi pescano da una numerazione comune**:

```
OT-xxx  <-  7 categorie
ST-xxx  <-  Furniture Storage + Furniture Systems Storage + Casework Storage
            + Generic Models Structural
CN-xxx  <-  Counter su tre categorie
SE-xxx  <-  Seating su Furniture e Furniture Systems
LB, MC, PR, SP, SV
```

La numerazione per prefisso li gestisce senza mai duplicare, ma `ST-001` non
dice piu' se e' un armadio o un elemento strutturale. **E' un effetto della
regola, non un difetto del codice.** Se un giorno da' fastidio in abaco, la
correzione e' nella regola, non nell'implementazione, e va fatta prima che i
modelli siano nominati.

### Group1 = Other

Il Type Mark diventa interamente manuale, prefisso compreso. **La cifra `9xx`
che le tabelle riportano accanto a Other non viene usata**: e' vestigiale. Sui
caricabili con Group1 proprio la regola non si applica, perche' li' il Group1
non alimenta il Type Mark.

### Segnaposto XX

Rimosso dall'interfaccia su richiesta. `format_type_mark()` e
`sequential_overflow()` in `FamilyNaming_rules.py` lo supportano ancora e sono
coperti dai controlli: la capacita' esiste, non e' esposta. Se dovesse
servire, si riespone senza toccare il motore.

### Esaurimento

Con due cifre l'ambito finisce a 99. Oltre, il tool smette di proporre e
chiede di scrivere il Type Mark a mano invece di inventare una forma diversa.

---

## 4. Divergenza nota fra tool ed Excel

**Sulle sette categorie a Group1 condiviso i fogli riportano `FN-SE01`, il
tool genera `SE-001`.** La regola buona e' quella del tool, decisa
esplicitamente. **Gli Excel vanno allineati**: finche' non lo sono, chi
consulta il foglio e chi usa il pulsante ottengono due risposte diverse.

Tutte le forme di nome prodotte dal tool coincidono con gli esempi dei fogli,
verificato riga per riga.

### Definizione di "faccia finita"

Il foglio Istruzioni parla di **numero di finiture** e di facce finite, che
suggerisce un criterio basato sulla funzione dello strato. Il tool applica
invece un criterio **posizionale**: conta i lati che hanno almeno uno strato
fuori dal core, qualunque funzione abbia (sezione 7). Le due definizioni
coincidono sui casi ordinari ma non su tutti: un cappotto o una barriera al
vapore fuori dal core fanno faccia per il tool e non per la lettera del foglio.

La regola operativa e' quella del tool, decisa esplicitamente. **Conviene
allineare anche il testo del foglio Istruzioni**, altrimenti chi lo legge conta
in un modo e il tool in un altro.

---

## 5. Rilevamento della categoria

**Deciso.** Completamente automatico, nessun menu. La finestra dichiara come
testo la categoria Revit, il codice categoria e la condizione di elemento in
place.

**Scartato:** combo precompilata e correggibile. Il ragionamento e' che
categoria e famiglia di sistema non sono opinioni: sono proprieta'
dell'elemento, e lasciarle modificabili invitava a produrre nomi falsi.

Tre strategie, in ordine:

1. `FORCED_CAT_CODE` per le categorie Revit che fissano il codice da sole
2. `WallType.Kind` per Basic Wall contro Curtain Wall
3. confronto normalizzato di `ElementType.FamilyName` con le etichette della
   tabella, che assorbe spazi e trattini (`Non-Monolithic Run` di Revit contro
   `NonMonolithic Run` della tabella)

Se nessuna porta a casa, si assume la prima voce e **si dichiara
l'incertezza** in giallo, perche' l'utente non ha modo di correggere.

### Due errori trovati verificando, da non rifare

- `OST_StairsRailingBaseTop` e `OST_TopRails`, che avevo scritto a memoria,
  **non esistono**. Il nome vero e' `OST_RailingTopRail`. I Top Rails non
  sarebbero mai stati riconosciuti.
- SlabEdge non e' un sottocaso di `OST_Floors`: e' `OST_EdgeSlab`, una
  categoria Revit a se'. Con la mappa sbagliata un bordo di solaio sarebbe
  stato nominato come un solaio.

**Regola operativa: ogni nome di `BuiltInCategory` va validato contro l'enum
reale, non scritto a memoria.** Vedi la sezione 9.

### Una inferenza non supportata dagli Excel

`OST_StructuralStiffener` e' mappata su `STRUCTURAL:07_CN` perche' li' esiste
il Group2 `SF Stiffener`. **Non e' scritto nei fogli, l'ho dedotto io.** Se e'
sbagliato si toglie una riga da `CATEGORY_MAP`.

### Platea contro solaio

`_is_foundation_slab()` usa solo `IsFoundationSlab`. **Scartato** il controllo
su `FLOOR_PARAM_IS_STRUCTURAL`: quello dice che il solaio e' portante, e un
solaio portante resta un solaio. Usarlo avrebbe mandato ogni solaio
strutturale sulla scheda delle fondazioni.

---

## 6. Cosa scrive il tool

| Caso | Dove scrive |
|---|---|
| Caricabili | `Family.Name` e `FamilySymbol.Name`. Il file `.rfa` su disco **non** viene toccato: Revit non lo consente da dentro un progetto |
| Sistema | solo il nome del tipo, Type Mark compreso |
| In place | **un nome solo**, famiglia e tipo uniti da `" - "`, scritto sia sulla famiglia sia sul suo unico tipo. Sulle categorie di sistema resta il solo nome del tipo |
| Type Mark | nel parametro, spunta attiva di default; se esiste gia' un valore diverso l'anteprima lo segnala |
| FireRating | quando REI e' compilato; se esiste un valore diverso la sovrascrittura viene **chiesta** prima della transazione |

Il separatore dell'in place e' `" - "` e non `" : "` perche' i due punti sono
fra i caratteri che Revit rifiuta nei nomi.

La rinomina usa `elem.Name = ...`, che e' l'idiom gia' presente in tutto il
repo, con `DB.Element.Name.SetValue()` come ripiego in `except`: su alcune
classi IronPython risolve la proprieta' sbagliata e la dichiara in sola
lettura.

Tutto in una sola transazione.

---

## 7. Cosa e' calcolato e cosa si digita

**Criterio: quello che si legge dal modello non si digita.** Una misura
ricavata da un parametro di tipo compare in sola lettura su fondo grigio, con
la provenienza accanto. Il nome deve dire quello che l'elemento e', non quello
che si vorrebbe.

**Scartato:** campo sempre editabile con il valore letto come default. Era la
proposta iniziale, superata da indicazione esplicita.

**Deroga lasciata, ed e' una mia decisione.** Se il valore non si riesce a
leggere, il campo torna editabile. Senza questa uscita, un tipo anomalo
bloccherebbe l'utente senza alternative.

### nFinishings: contato per posizione, non per funzione

**Deciso (17.09.2026).** Una faccia conta come finita quando su quel lato
esiste **almeno uno strato fuori dal core**, cioe' nello shell esterno o in
quello interno. Si legge con `GetNumberOfShellLayers()`, con gli indici di
`GetFirstCoreLayerIndex()` / `GetLastCoreLayerIndex()` come ripiego. Quanti
strati ci siano per lato non cambia nulla: uno o dieci fanno una faccia.

La regola dello strato singolo non e' piu' un caso speciale nel codice, ne
discende da sola: un elemento a un solo strato ha quello strato dentro al core,
nessuno shell, quindi `0F`.

**Superata** la prima implementazione, che contava gli strati estremi con
funzione `Finish1` / `Finish2`. Ignorava tutto cio' che sta fuori dal core ma
non e' marcato come finitura, e sbagliava su casi frequenti:

| Stratigrafia | Vecchia | Nuova |
|---|---|---|
| cappotto + muratura + intonaco | 1F | **2F** |
| barriera + isolante fuori core su un lato | 0F | **1F** |
| massetto + soletta + intonaco | 1F | **2F** |

I controsoffitti restano limitati a `1F` dalla tabella `Tbl_Fin01`.

La descrizione **non viene sanitizzata**, gli spazi sono ammessi. L'unico
controllo bloccante sui testi liberi e' sui caratteri che Revit rifiuta.

---

## 8. Le note e il loro filtro

I fogli portano 559 righe di note, divise in due blocchi: le NOTE di
categoria, che dicono come classificare, e le REGOLE DI COMPILAZIONE, che
spiegano colonna per colonna cosa scrivere.

Il secondo blocco, 373 righe su 559, e' ridondante nel tool: ogni campo porta
gia' il suo suggerimento accanto, i due gruppi hanno la descrizione viva sotto
la tendina, categoria e misure sono mostrate come valori calcolati.

**Deciso.** Ne restano 58. Il filtro sta in `FamilyNaming_map.py`:

| Chiave | Effetto |
|---|---|
| `NOTES_SHOW_FIELD_RULES` | `False` nasconde in blocco tutte le REGOLE DI COMPILAZIONE |
| `NOTES_HIDDEN` | 68 intestazioni di note di categoria, nascoste ovunque |
| `NOTES_HIDDEN_ON` | eccezioni per scheda |

La selezione delle note di categoria e' stata fatta a mano dal dipartimento su
`NOTES-review.xlsx`, che resta nella cartella come traccia della decisione: 52
SI e 75 NO su 127 gruppi.

**Il filtro sta nella mappa e non nel modulo dati** perche' quello e'
generato: toglierle di la' significherebbe vederle tornare alla prima
rigenerazione.

**Scartato:** lista di inclusione al posto della lista di esclusione. Con una
keep-list una nota nuova aggiunta agli Excel resterebbe invisibile finche'
qualcuno non aggiorna il codice. Con la drop-list scelta, **una nota nuova
compare di default** e va tolta esplicitamente se non serve. Un contenuto
scritto apposta non deve sparire in silenzio.

**Assunzione da confermare.** Il foglio `Field rules` di `NOTES-review.xlsx` e'
tornato **interamente vuoto**, 30 righe senza segno. Ho assunto che valgano
tutte NO. Se l'intenzione era diversa, basta rimettere
`NOTES_SHOW_FIELD_RULES = True` e poi togliere le singole.

Nota collaterale: `Top rail / handrail` e' l'unico campo dell'interfaccia
senza suggerimento inline.

---

## 9. Vincoli tecnici scoperti sul campo

Cose imparate facendo, che costerebbero un giro di prova in Revit a
riscoprirle.

**IronPython non lascia attaccare attributi Python a un oggetto .NET.**
`box._watermark = testo` su un `TextBox` solleva `AttributeError`. Lo stato
accessorio dei controlli va tenuto in dizionari della finestra, indicizzati
per chiave di campo. E' stato il primo crash in Revit.

**`OfClass()` vuole un `System.Type`.** `type(element_type)` restituisce il
wrapper Python: serve `element_type.GetType()`.

**`Visibility` va usato come enum.** Passare `1` da' `Hidden`, che lascia lo
spazio occupato: serve `Visibility.Collapsed`.

**Il XAML non ha BOM ne' dichiarazione XML**, come gli altri file del repo, e
con quella forma le emoji funzionano. Verificato confrontando i primi byte con
`ModelReportForm.xaml` e `WorksetCreateForm.xaml`.

**Le icone del repo hanno sfondo trasparente**, fra il 51% e il 79% di pixel a
alpha zero, e quasi tutte sono in bianco e nero. Le due icone sono state fatte
disegnando la sagoma in grande come maschera e usandone la luminosita' come
canale alpha: bordi morbidi e buchi veri, non toppe del colore di fondo.

**Lingua.** Tutto quello che l'utente vede e' in inglese, commenti e
documentazione in italiano. E' la regola del repo, vale anche per `bundle.yaml`
e per i tooltip.

---

## 10. La regola di composizione mostrata a schermo

La sezione NAMING RULE dichiara il pattern della categoria prima di chiedere i
campi.

**Deciso.** E' **ricavata dalla definizione di scheda** (`blocks` e `dim` in
`FamilyNaming_map.py`), la stessa che guida la composizione.

**Scartato:** mostrare i pattern gia' scritti in cima ai fogli Excel, che il
generatore estrae comunque in `PATTERNS`. Sono in italiano misto a inglese, ma
soprattutto sarebbero stati una seconda fonte di verita': il giorno che si
aggiunge un blocco al nome, la finestra continuerebbe a dichiarare la regola
vecchia senza che nessuno se ne accorga.

L'autore compare come `e_`, preso da `AUTHOR_CODE`, perche' non e' un campo da
compilare ma una costante.

---

## 11. Reingerire un Excel modificato

### Procedura

1. Sostituisci il `.xlsx` nella cartella del bundle
2. `py gen_data.py` (serve `openpyxl`, CPython, non IronPython)
3. `git diff FamilyNaming_data.py` -> **questo e' il changelog**
4. Dal diff si capisce a quale livello siamo

### I tre livelli di impatto

**Livello 1, nessun codice.** Righe aggiunte, tolte o corrette dentro una
tabella che esiste gia': nuovo Group1, nuovo Group2, etichetta o TM Code
cambiato. Funziona subito, menu e Type Mark compresi. Misurato: tre modifiche
all'Excel producono un diff di tre righe e girano senza toccare nulla.

**Livello 2, una quindicina di righe dichiarative.** Una scheda generatore
nuova. Il generatore ne estrae gia' tabelle, note e pattern, ma il tool non ci
arriva finche' non si aggiungono due voci in `FamilyNaming_map.py`, una in
`SHEETS` e una in `CATEGORY_MAP`. Servono quattro informazioni che dall'Excel
non si deducono: se la scheda genera il nome famiglia o solo quello del tipo,
quale blocco dimensionale usa, quali blocchi opzionali ha, quale categoria
Revit ci porta.

**Livello 3, codice vero.** Un tipo di campo che oggi non esiste, una forma di
blocco dimensionale mai vista, una regola condizionale nuova, o un cambio alla
legge del Type Mark. Si tocca il costruttore del campo nella UI, il traduttore
del blocco in `FamilyNaming_rules.py`, ed eventualmente il formattatore
dimensionale.

### Vincoli da tenere presenti

- La cifra del Type Mark ha dieci valori: **un Group che alimenta il Type Mark
  non puo' superare le dieci voci.** Oggi `8xx` e' libera quasi ovunque.
- Aggiungere un codice Group2 a una delle sette categorie a Group1 condiviso
  significa **creare o entrare in una numerazione condivisa**, perche' quel
  codice diventa il prefisso del Type Mark. Vedi la sezione 3.
- Le tabelle G1 e G2 possono essere per sottogruppo e non per scheda: Walls,
  Roofs e Stairs hanno liste distinte per famiglia di sistema, e gli Handrails
  usano la tabella Group1 dei Top Rails.
- **Nessun avviso oggi segnala una scheda nuova senza voce nella mappa**: il
  tool semplicemente non la raggiungera' mai, in silenzio. Un controllo di
  coerenza in `gen_data.py` e' stato proposto e non ancora fatto.

---

## 12. Controlli da rieseguire dopo ogni modifica

Sono tutti eseguibili senza Revit e hanno gia' trovato difetti veri.

| Controllo | Cosa verifica | Cosa ha trovato |
|---|---|---|
| Sintassi di ogni modulo | `compile()` su tutti i `.py` | - |
| XAML ben formato | parsing XML | - |
| Riferimenti tabella | ogni `cat_table`, `g1`, `g2` della mappa esiste in `TABLES` | - |
| Forma del Type Mark | tutte le schede producono `AA-XXX` | - |
| Regola contro nome | la regola mostrata ha gli stessi blocchi, nello stesso ordine, del nome composto con tutti i campi pieni | - |
| Nomi `BuiltInCategory` | ogni nome nella mappa esiste nell'enum, letto per reflection da `RevitAPI.dll` | **due nomi inesistenti** |
| Nomi `x:Name` | quelli del XAML e quelli cercati dalla UI coincidono nei due sensi | un `lbl_sheet` rimasto a meta' |
| Attributi su oggetti .NET | nessuna assegnazione di attributo Python a un controllo WPF | **il crash del watermark** |
| Icone | sola scala di grigi, canale alpha identico fra chiara e scura | - |

Il controllo sui nomi di categoria si fa caricando `RevitAPI.dll` in
reflection-only. La 2026 non si carica da Windows PowerShell 5.1 perche'
dipende da .NET 10: si usa la **2024**, che basta.

---

## 13. Stato di verifica

**Il tool non e' mai stato eseguito con successo fino in fondo in Revit.** Due
avvii hanno prodotto crash, entrambi corretti: il watermark e, trovato
rileggendo, `OfClass(type(...))`.

Tutto il resto e' verificato offline: motore di composizione su tutte e 36 le
schede, Type Mark nelle cinque forme, sequenziali, riconoscimento della
categoria contro i nomi reali che Revit assegna alle famiglie di sistema,
coerenza dei dati, nomi dell'interfaccia.

Restano non verificati, in ordine di rischio:

1. la lettura dei parametri dimensionali oltre `WallType.Width` e
   `CompoundStructure.GetWidth()`, in particolare i parametri cercati per nome
   su Furniture e simili, che quasi certamente non risponderanno
2. `BuiltInParameter.RAILING_HEIGHT` per l'altezza dei parapetti
3. la rinomina vera in transazione, e il comportamento di un elemento in place
   quando si scrive lo stesso nome su famiglia e simbolo
4. `detect_cat_code` su una libreria con famiglie di sistema rinominate

---

## 14. Punti aperti

- **Gli Excel vanno allineati** sulla regola `SE-001` (sezione 4)
- **Il foglio `Field rules` non e' stato compilato**: assunzione tutte NO
  (sezione 8)
- `OST_StructuralStiffener` su `07_CN` e' una mia inferenza (sezione 5)
- `Top rail / handrail` e' l'unico campo senza suggerimento inline
- Il controllo di coerenza in `gen_data.py` non e' stato fatto (sezione 11)
- La persistenza via `script.get_config` non e' stata implementata: con le
  tabelle congelate e l'autore fisso non restava niente di variabile da
  ricordare
- Nessuna funzione su SHIFT + CLICK, esclusa esplicitamente. Il candidato
  naturale resterebbe un audit dei nomi non conformi
