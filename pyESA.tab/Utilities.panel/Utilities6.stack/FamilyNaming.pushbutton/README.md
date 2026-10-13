# Family Naming

Rinomina guidata di una famiglia e di un tipo secondo la classificazione ESA.
Si seleziona un elemento nel modello, il tool riconosce la categoria, apre la
scheda che le corrisponde e compone il nome.

> Questo documento spiega **come funziona** il tool.
> Il perche' delle scelte fatte, le alternative scartate, i punti aperti e la
> procedura per reingerire un Excel modificato stanno in
> [FamilyNaming_NOTE-SVILUPPO.md](FamilyNaming_NOTE-SVILUPPO.md).

## I file

| File | Cosa contiene |
|---|---|
| `FamilyNaming_script.py` | Entry point. Selezione, rilevamento categoria, lettura delle misure dai parametri, transazione, report. E' l'unico che tocca le API di Revit. |
| `FamilyNaming_map.py` | Mappa dichiarativa: quali schede esistono, quali tabelle DV le alimentano, quali blocchi compongono il nome, quale categoria Revit porta a quale scheda. Si corregge senza saper programmare. |
| `FamilyNaming_rules.py` | Motore di composizione: Type Mark, blocco dimensionale, nomi, validazione. Logica pura, nessuna API. |
| `FamilyNaming_ui.py` | Finestra. Costruisce a runtime i campi che cambiano da una categoria all'altra. |
| `FamilyNaming.xaml` | Layout della finestra. |
| `FamilyNaming_data.py` | **Generato.** Le tabelle dei tre Excel congelate in dizionari Python. Non si modifica a mano. |
| `gen_data.py` | Rigenera `FamilyNaming_data.py` dagli Excel. Non fa parte del comando. |
| `Naming classification-*.xlsx` | I tre file di classificazione, sorgente documentale. Il tool **non** li legge a runtime. |
| `FamilyNaming_NOTE-SVILUPPO.md` | Registro delle scelte, divergenze note, procedura di reingestione. |
| `NOTES-review.xlsx` | Traccia della selezione delle note fatta dal dipartimento. Il tool non lo legge. |

## Modificare le tabelle di classificazione

Le tabelle non si leggono dagli Excel a ogni avvio: sono congelate in
`FamilyNaming_data.py`. Dopo una modifica a uno dei tre file serve rigenerare.

```powershell
cd <questa cartella>
py -m pip install openpyxl     # solo la prima volta
py gen_data.py
```

Il generatore legge le Excel Table nominate (`Tbl_DR_G1`, `Tbl_BW_G2`,
`Tbl_STR_G1`...), non le coordinate delle celle: aggiungere o togliere righe
dentro una tabella esistente funziona senza altri interventi. Aggiungere una
**categoria** nuova richiede invece anche una voce in `FamilyNaming_map.py`,
sia in `SHEETS` sia in `CATEGORY_MAP`.

## La regola di composizione a schermo

Subito sotto DETECTED la finestra dichiara la regola della categoria, in
grassetto blu, prima di chiedere qualunque campo: cosi' si sa dove andra' a
finire quello che si sta scrivendo senza doverlo dedurre dall'anteprima.

```
Doors     Family   e_[(I)]Cat.G1.G2_L{n}_[CW]_[EN]_[REI]_[Manufacturer]_[Brand]_[Description]
          Type     TypeMark_{W}x{H}_[Description]

Walls     Type     e_TypeMark_[(I)]Cat.G1.G2_nF.{T}_[REI]_[WI]_[Description]
```

Blocchi nudi sempre presenti, `[ ]` solo se compilati, `{ }` un valore, `|` una
forma o l'altra. L'autore compare per quello che e', la lettera che finisce
nel nome, e viene da `AUTHOR_CODE` nella mappa: se cambia, la regola lo segue. La riga Family sparisce sulle schede di sistema.

La regola non e' copiata dai pattern scritti negli Excel: e' **ricavata dalla
stessa definizione di scheda che guida la composizione** (`blocks` e `dim` in
`FamilyNaming_map.py`). Aggiungere un blocco al nome lo fa comparire anche
nella regola, senza allineamenti manuali, e non e' possibile che la finestra
dichiari una regola diversa da quella che applica.

## Il Type Mark

Sempre nella forma `AA-XXX`: due lettere, trattino, tre cifre.

| Caso | Prefisso | Cifre | Esempio |
|---|---|---|---|
| Caricabili con TM Code (DR, WN, CP, CM, PK, EN, PL, MS) | codice categoria | cifra del Group2 + 2 di sequenziale | `DR-101` |
| Caricabili senza TM Code (FN, FS, CK, PF, SE, SI, GM) | codice **Group2** | 3 di sequenziale | `SE-001` |
| System e Structural | codice categoria | cifra del Group1 + 2 di sequenziale | `BW-101` |
| Group1 = Other | codice materiale | 3 di sequenziale | `WD-001` |
| Structural Framing con Use | codice Use | 3 di sequenziale | `SK-001` |

Il primo sequenziale libero viene sempre proposto, senza opzioni da attivare, e
si puo' sovrascrivere scrivendoci sopra: finche' nel campo c'e' la proposta il
tool la aggiorna, appena c'e' un valore digitato lo lascia stare. La ricerca e'
**per prefisso sull'intero documento**, prendendo il massimo in uso piu' uno, e si legge dal parametro Type Mark dei
tipi, compresi quelli non piazzati. La ricerca non e' per categoria perche'
sulle sette categorie a Group1 condiviso il codice di categoria non compare nel
Type Mark e gli stessi codici Group2 ricorrono altrove: `SE` e' Seating sia su
Furniture sia su Furniture Systems, e numerare per categoria produrrebbe due
`SE-001` diversi.

> **Attenzione: divergenza nota dagli Excel.** Sulle sette categorie a Group1
> condiviso i fogli riportano ancora la forma vecchia `FN-SE01`, mentre il tool
> genera `SE-001`. La regola buona e' quella del tool. Gli Excel vanno allineati.

## Cosa scrive il tool

- **Famiglie caricabili**: nome della famiglia e nome del tipo. Il file `.rfa`
  su disco non viene toccato, Revit non lo consente da dentro un progetto.
- **Famiglie di sistema**: il solo nome del tipo, Type Mark compreso.
- **Elementi in place**: un nome solo, famiglia e tipo uniti da `" - "`, scritto
  sia sulla famiglia sia sul suo unico tipo. Sulle categorie di sistema, che una
  regola sul nome famiglia non ce l'hanno, resta il solo nome del tipo.
- **Parametro Type Mark**, salvo togliere la spunta. Se il tipo ne ha gia' uno
  diverso l'anteprima lo segnala prima di sovrascriverlo.
- **FireRating**, quando il campo REI e' compilato. Se il parametro ha gia' un
  valore diverso, la sovrascrittura viene chiesta prima di applicare.

Tutto avviene in una sola transazione.

## Cosa viene dedotto e cosa si digita

Il criterio e' che **quello che si legge dal modello non si digita**. Una misura
ricavata da un parametro di tipo compare in sola lettura su fondo grigio, con
l'indicazione della provenienza: spessore delle pareti da `WallType.Width`,
spessore dei pacchetti stratificati dalla `CompoundStructure`, larghezza e
altezza di porte e finestre dai rispettivi parametri. Il nome deve dire quello
che l'elemento e', non quello che si vorrebbe. Dove un parametro standard non
esiste il campo resta editabile, vuoto, con un esempio in grigio che non viene
mai salvato come valore.

Il numero di facce finite (`0F`, `1F`, `2F`) segue la stessa regola e si mostra
in sola lettura. Una faccia conta come finita quando su quel lato esiste
**almeno uno strato fuori dal core**, cioe' nello shell esterno o in quello
interno: non conta a che funzione sia assegnato lo strato, conta che ci sia. Uno
strato per lato o dieci fanno lo stesso una faccia. Da qui discende da sola la
regola dello strato singolo, che sta tutto dentro al core e quindi e' `0F` anche
quando e' proprio la finitura. I controsoffitti restano limitati a `1F`. Torna a essere una scelta solo se la
stratigrafia non si riesce a leggere, altrimenti non si potrebbe procedere.

La categoria, il codice della famiglia di sistema e la condizione di elemento
modellato in place sono anch'essi rilevati e mostrati come testo, senza menu:
non sono decisioni dell'utente. Se il riconoscimento della famiglia di sistema
non riesce, il tool assume la prima voce e lo dichiara in giallo invece di
tacere.

## Le note di categoria

I fogli generatore portano in coda 559 righe di note, divise in due blocchi: le
NOTE della categoria, che dicono come classificare, e le REGOLE DI COMPILAZIONE,
che spiegano colonna per colonna cosa scrivere. Il secondo blocco serve dentro
l'Excel, dove non c'e' nessuno a spiegarti le colonne, ma nel tool e' ridondante
perche' ogni campo porta gia' il suo suggerimento accanto.

Il pannello ne mostra **58 su 559**, selezionate a mano dal dipartimento. Il
filtro sta in `FamilyNaming_map.py` e non in `FamilyNaming_data.py`, perche'
quello e' generato e le note tornerebbero alla prima rigenerazione:

| Chiave | Cosa fa |
|---|---|
| `NOTES_SHOW_FIELD_RULES` | `False` nasconde in blocco tutte le REGOLE DI COMPILAZIONE |
| `NOTES_HIDDEN` | intestazioni di note di categoria da non mostrare, su tutte le schede |
| `NOTES_HIDDEN_ON` | eccezioni: intestazione da nascondere solo su alcune schede |

Una nota aggiunta agli Excel in futuro **compare di default**: per toglierla si
aggiunge la sua intestazione a `NOTES_HIDDEN`. La scelta e' voluta, cosi' un
contenuto nuovo scritto apposta non sparisce in silenzio.

Su Entourage, Mass, Ramps e Toposolid non resta nessuna nota e il riquadro
sparisce del tutto invece di restare vuoto.

## Categorie non coperte

Le categorie fuori dalle tre classificazioni (tutto l'MEP, annotazioni,
elementi di dettaglio) fanno comparire un avviso e il tool si ferma senza
rinominare nulla.

Il foglio Toposolid si applica da Revit 2024 in avanti.

## Limiti noti

- Una esecuzione rinomina **una famiglia e un tipo**. Non e' una rinomina di
  massa.
- Con il sequenziale a due cifre l'ambito si esaurisce a 99 tipi: oltre quella
  soglia il tool smette di suggerire e chiede di scrivere il Type Mark a mano.
- Il rilevamento della famiglia di sistema (Basic Wall contro Curtain Wall,
  platea contro plinto) non si puo' correggere dall'interfaccia. Nei casi
  verificati funziona, ma su una libreria con famiglie di sistema rinominate
  potrebbe sbagliare: in quel caso il tool lo segnala e la correzione va fatta
  sulla mappa.
- Non testato in Revit. Tutte le verifiche fatte finora sono offline sul motore
  di composizione e sulla coerenza delle tabelle.
