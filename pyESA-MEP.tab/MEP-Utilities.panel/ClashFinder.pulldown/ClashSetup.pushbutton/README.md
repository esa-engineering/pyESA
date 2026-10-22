# Clash Finder

Clash detection a test salvati, sul modello di Navisworks. Sostituisce l'uso diretto
dell'Interference Check di Revit con test riutilizzabili e un report navigabile.

Pulldown **Clash Finder** (`ClashFinder.pulldown`, pannello MEP-Utilities) con tre comandi:
**Clash Setup** (finestra principale: crea, modifica ed esegue i test), **Quick Clash**
(esegue un test salvato) e **Last Report** (riapre l'ultimo report).

## Concetti

- **Test**: confronta **Set A** con **Set B** e ha un nome. Sempre, senza spunte (tolte:
  `normalize_test` forza `ignore_connected` / `ignore_hosted` a true anche nei file che li
  hanno a false):
  - le coppie collegate da connettori MEP non sono clash (un canale e il suo raccordo si
    toccano sempre);
  - le coppie elemento / host non sono clash (porta e muro). Un elemento ospitato che sfonda
    davvero il suo host non viene quindi segnalato.
- **Set**: sorgenti, categorie, filtri di vista e regole, tutto salvato nel test.
  - *From this model*: *Current selection*, vista attiva, intero modello o *Links only* (nulla dal modello aperto, solo i link). La
    selezione corrente di Revit e' riletta a ogni aggiornamento del riepilogo e all'avvio del
    run (`_before_run`): si puo' cambiarla in Revit con la finestra aperta. Default per i test
    salvati prima di questa versione: A dalla selezione, B dal modello.
  - *Links*: menu a discesa (un bottone che apre un `Popup` con l'elenco a caselle, alto fino
    a 260 px con scroll interno; si chiude cliccando fuori). Il testo del bottone dice quanti
    link sono spuntati: "None (click to choose)", il nome se ce n'e' uno, "N links" da due in
    su (`_update_links_button`); il tooltip elenca i nomi. Un clic sul bottone mentre il popup
    e' aperto lo chiude e non lo riapre (guardia di 250 ms su `Closed`). Link salvati per
    **nome del file** (`Document.Title`),
    validi per tutte le istanze caricate di quel file e in ogni progetto che carica gli stessi
    link. Un link del test non caricato resta in elenco come "(not loaded)" e compare nelle
    note. Gli elementi linkati vengono dall'intero link; con sorgente *Active view* da quanto
    del link e' visibile nella vista (`FilteredElementCollector(doc, view, link)`, Revit
    2024+; prima l'intero link, con una nota).
  - *Categories*: nessuna categoria spuntata = qualsiasi categoria di modello. Accanto a ogni
    categoria spuntata il numero di elementi che entrano nel set; *Only categories with
    elements* guarda il modello aperto e i link caricati.
  - *Model filters*: filtri di vista (`ParameterFilterElement`) del modello aperto, salvati per
    nome. Con filtri spuntati l'elemento deve passarne almeno uno (sue categorie e sue
    regole, `GetElementFilter().PassesFilter`), in AND con le categorie spuntate. Sugli
    elementi linkati valgono i parametri di sistema; quelli di progetto e condivisi hanno id
    diversi nel link e non trovano corrispondenza. Un filtro salvato ma assente dal progetto
    compare come "(not in this model)" e svuota il set se e' l'unico.
  - *Custom rules*: combinate con *All rules* (AND) o *Any rule* (OR). Una regola legge il
    parametro d'istanza, poi quello del tipo, come testo mostrato da Revit (`AsString` /
    `AsValueString`), senza distinguere maiuscole e minuscole. `is greater than` /
    `is less than` confrontano il primo numero del testo (`300 mm` = 300). Parametro assente:
    soddisfa solo `does not equal`, `does not contain`, `has no value`. Pseudo-parametri
    indipendenti dalla lingua: `[Family]`, `[Type]`, `[Workset]`. I nomi proposti vengono da
    un campione di 10 elementi per categoria spuntata, nel modello aperto e in ogni link.
- Le categorie sono salvate come nomi `BuiltInCategory` (`OST_DuctCurves`): un test funziona
  con Revit in qualsiasi lingua. Una categoria salvata ma assente dal modello compare in cima
  all'elenco come "(not in this model)".

## Finestra principale (non modale)

Resta aperta mentre si lavora in Revit. Ogni accesso al modello passa da un ExternalEvent
(`clashfinder_bridge.RevitBridge`): da una finestra non modale non si e' in contesto API.

Una sola pagina:

- a sinistra la libreria: un clic carica il test nell'editor (sotto il nome, in grigio,
  le sorgenti di A e B: `describe_scope()`, con due o piu' link ridotti a "N links"; i nomi
  restano nel tooltip della riga); *New* (chiede il nome nella finestra di
  salvataggio e aggiunge subito alla libreria un test vuoto: nessuna categoria, link o filtro,
  A e B dalla selezione corrente), *▲* / *▼* (sposta il test
  selezionato; l'ordine si salva subito nella libreria e vale anche per Quick Clash), *Delete*, *Import JSON*, *Export JSON* (tutta la libreria);
- a destra l'editor (il nome del test caricato sta nel titolo della finestra, con `*` se
  modificato): le colonne **Set A** e **Set B** con la sorgente nel modello aperto sulla riga del titolo, i link, le
  categorie (conteggio in una colonna a destra), i filtri di vista (con ricerca per nome) e
  le regole. Negli elenchi di categorie e filtri SHIFT + clic porta
  allo stato della casella cliccata tutte quelle visibili fra lei e l'ultima cliccata
  (`_shift_range`). Un insieme vuoto e' segnalato in rosso nella riga di stato;
- in fondo la riga di stato ("Unsaved changes", elementi selezionati in Revit, note), *Refresh*, *Save*,
  *Save as...*, *Close*, *Run*.

*Run* esegue il test a schermo, anche con modifiche non salvate. *Save* sovrascrive il test
caricato (per un test mai salvato fa come *Save as...*). *Save as...* chiede il
nome in una finestra a parte (`ClashFinder_save.xaml`, proposto quello del test caricato):
lo stesso nome sovrascrive il test, un nome gia' usato da un altro test chiede conferma, un
nome nuovo aggiunge un test. Cambiando test, creandone uno nuovo o chiudendo con modifiche
non salvate: Yes = Save (sul test caricato, o con nome se mai salvato), No = scarta,
Cancel = resta.

Ogni modifica che cambia il contenuto dei set (sorgenti, link, categorie, filtri, regole:
queste ultime all'uscita dal campo) ricalcola i conteggi via bridge (`ce.summarize`,
`Summary.by_category`). Impostazioni utente: solo `last_test`, il test caricato alla
riapertura.

## Ricerca delle interferenze

**Fasi.** Ogni sorgente (selezione, vista attiva, intero modello, link) passa dalla fase e dal
filtro di fase della **vista attiva** (`ce.PhaseScope`, in `SetResolver`): un elemento entra
se il filtro della vista lo mostra (`PhaseFilter.GetPhaseStatusPresentation` diverso da
`DontShow`); con filtro `<none>` entrano tutti tranne gli stati *Past* e *Future*; gli
elementi senza fasi entrano sempre. Nei link si usa la fase con lo **stesso nome** (senza
distinzione di maiuscole): un link che non ce l'ha resta fuori, con una nota. Se la vista
attiva non ha i parametri Phase / Phase Filter (tavola, abaco, legenda...) Run e il Refresh
del report sono bloccati con un messaggio; il riepilogo lo segnala in rosso. Fase e filtro
usati sono nella riga di stato, nell'intestazione del report e nel file (`"phase"`).

Stesso criterio dell'Interference Check: intersezione dei solidi. Ogni elemento e' un `Item`
con la sua `ModelSource` (documento, istanza e trasformazione del link). Per ogni elemento
di A, prima il filtro sul bounding box, poi la prova sui solidi:

| A | B | Bounding box | Solidi |
| --- | --- | --- | --- |
| host | host | `BoundingBoxIntersectsFilter` | `ElementIntersectsElementFilter` |
| link | host | `BoundingBoxIntersectsFilter` (box del link trasformato) | `ElementIntersectsSolidFilter` con i solidi del link trasformati |
| host | link | griglia `BoxGrid` in Python | `ElementIntersectsSolidFilter` sul solo elemento host |
| link | link | griglia `BoxGrid` | `BooleanOperationsUtils` Intersect, volume > ~0.3 cm³ |

I solidi dei link sono letti a dettaglio Fine e portati nelle coordinate del modello aperto
con `GetTotalTransform()`; sono calcolati solo per i candidati che passano il bounding box.
Gli id di documenti diversi si sovrappongono: chiavi, cache e deduplicazione usano
(indice della sorgente, id). *Ignore connected* / *Ignore host* valgono solo fra elementi
dello stesso documento.

Gli elementi senza solidi utilizzabili (host: `ElementIntersectsFilter.IsElementSupported`
falso; link: nessun solido) sono contati come "skipped". Una coppia presente in entrambi gli
insiemi e' contata una volta. Barra di avanzamento annullabile: il risultato parziale viene
comunque mostrato.

Diagnostica: per ogni test il report indica quanti elementi linkati c'erano in A e in B,
quante coppie con un elemento linkato sono state provate sui solidi e quanti errori hanno
dato le prove geometriche (con il primo messaggio). Se `ElementIntersectsSolidFilter`
rifiuta un solido, i candidati vengono provati con l'intersezione booleana.

Limiti attuali: link annidati non letti, solo hard clash (nessuna tolleranza / clearance),
nessuno stato delle interferenze fra un'esecuzione e l'altra. Le operazioni booleane fra
solidi linkati possono fallire su geometrie difficili: la coppia viene allora scartata.

## Report (non modale)

Elenco espandibile con un gruppo per categoria dell'elemento di Set A ("Duct Fittings",
linea spessa sopra; `_category_groups`) e una riga per interferenza con davanti la categoria
dell'altro elemento ("Walls   Clash #12"), in ordine di altra categoria e numero. Il menu
accanto alla ricerca (*Group by Set A* / *Group by Set B*) sceglie l'insieme dei gruppi; i
numeri non cambiano. Nel gruppo delle risolte la riga riporta le due categorie;
espandendola, i due elementi con livello, categoria, famiglia : tipo e id, preceduti da un
quadratino del colore del loro insieme (A verde, B rosso) e dal nome del link se linkati. I
numeri seguono l'ordine dell'elenco e sono gli stessi dell'export CSV. Ricerca per modello,
categoria, famiglia, tipo, livello o id. Con piu' test (non piu' possibile dalla finestra
principale, che ne esegue uno) c'e' un livello in piu' per test.

Un clic su una riga (gruppo, interferenza o elemento), via bridge e in una sola transazione
"Clash Finder view":

- **elementi linkati disegnati**: Revit non ha override ne' "nascondi" via API sul singolo
  elemento di un link (`View.HideElements` non accetta elementi dei link), quindi i solidi
  dei linkati della riga, gia' trasformati (`ce.element_solids`), vengono **disegnati** con
  DirectContext3D (`pyrevit.revit.dc3dserver`: `Mesh.from_solid`, verde A / rosso B). Non
  si crea nessun elemento nel modello: niente workset, abachi o sincronizzazione. Il server
  (`_make_drawing_server`) disegna solo nella vista attiva al momento del clic
  (`CanExecute` sull'id della vista), viene creato al primo uso e rimosso alla chiusura. Un
  linkato senza mesh (nessun solido), o pyRevit senza `dc3dserver`, ricade sull'istanza del
  link intera. Le DirectShape `ESA_ClashFinder_Preview` delle versioni precedenti vengono
  ancora cancellate a ogni clic e alla chiusura;
- **colori**: override di elemento nella vista attiva sugli host, Set A verde (0,170,0), Set B
  rosso (220,0,0), linee e retino pieno di superficie e di taglio. Gli host colorati prima, in
  quella vista, tornano all'override originale (salvato prima di toccarli);
- **contenitore trasparente**: quando la riga e' una singola clash, dei due elementi quello
  con il bounding box piu' grande (`container_key`) e' reso trasparente
  (`CONTAINER_TRANSPARENCY`, 0-100 come in Revit, ora 40): override di trasparenza della
  superficie se host, canale alfa (0-255) della mesh DirectContext3D se linkato;
- **section box**: solo se la vista attiva e' 3D, attorno agli elementi con 10 cm di margine;
- **isolamento** (*Isolate selected*, **attivo all'apertura**): `IsolateElementsTemporary`
  sugli host della riga, cosi' link e resto spariscono e restano gli host e i linkati
  disegnati: **si vedono solo i due elementi**. Una riga di soli linkati nasconde
  temporaneamente tutto il resto della vista (`HideElementsTemporary`). Spegnendo il
  pulsante l'isolamento viene tolto (i linkati disegnati si sovrappongono allora al link);
- **zoom**: `ShowElements` sugli host se la riga non ha linkati, altrimenti
  `UIView.ZoomAndCenterRectangle` sul box memorizzato.

*Hide overlay* / *Show overlay* (`overlay_on`, attivo all'apertura) spegne colori e linkati
disegnati insieme; la navigazione resta (section box, isolamento con il link intero, zoom).

Ogni clic e' una transazione "Clash Finder view" (compare nell'annulla).

La selezione di Revit non viene toccata: l'evidenziazione blu coprirebbe i colori. *Copy ID*
copia negli appunti gli id della riga selezionata (separati da virgola; per un elemento
linkato l'id nel link). Alla chiusura colori e isolamento vengono tolti. Con un view
template che controlla Visibility/Graphics i colori non si applicano (nota nel report).
Export CSV (`;`, UTF-8 con BOM, con il modello di ciascun elemento) ed Export JSON (il
report intero, riapribile con Last Report; vedi sotto).

### Refresh

*Refresh* ricontrolla le coppie A/B gia' presenti nel report sul modello com'e' adesso,
senza rilanciare i test e senza usare la vista attiva o la selezione corrente. Confronta
con il **report di partenza** (quello aperto, non l'ultimo refresh: `self.baseline`).
Un'interferenza e' la coppia (istanza di link o host, id) dei due elementi
(`ce.pair_key`):

- quelle ancora presenti tengono il loro numero;
- quelle non piu' trovate finiscono nel gruppo grigio *✔ Resolved since the report*, con il
  loro numero, ancora navigabili.

Il refresh controlla solo clash gia' contenute nel report: non scopre nuove coppie.
Il refresh **non scrive** il file: Export JSON salva le interferenze di adesso (il file
precedente va in `_old`, perche' non e' piu' lo stesso report). Export CSV ha la colonna
*Status* (`resolved`) e le risolte in coda.

## Cartella dei report e Last Report

Ogni modello ha una **cartella dei report**, ricordata nella config
(`ESA_ClashFinder` / `report_folders`, voci `"<model_key>::<cartella>"`, al massimo 50;
`model_key` e' il percorso del centrale, poi del file, poi il titolo). Viene chiesta al
primo run (o se non esiste piu') e cambiata con *Export JSON...*, che chiede sempre la
cartella e la fa diventare quella del modello.

Dopo ogni run (`report_ui.save_run_reports`) e con Export JSON (`ce.save_reports`) si
scrive **un file per clash test**, `<nome del test>.json` (caratteri vietati da Windows
sostituiti con `_`). Se il file esiste gia' viene spostato in `_old\<test>_<aaaammgg_hhmmss>.json`
(data del file spostato); un file scritto per lo stesso report (export dopo il
salvataggio automatico, o di un report riaperto da quel file) viene solo riscritto. Se si
rinuncia a scegliere la cartella o la scrittura fallisce, il report si apre comunque con
una nota nell'intestazione.

Last Report e' il terzo comando del pulldown (`LastReport.pushbutton`): CLICK apre il
report piu' recente della cartella del modello (`ce.latest_report`, per data del file,
escluso `_old`); se non ce n'e' chiede un file. SHIFT + CLICK importa un file qualsiasi,
partendo dalla cartella del modello (anche da `_old`). Il report riaperto ha la stessa
navigazione; nell'intestazione la data del run e, se il file viene da un altro modello,
un avviso. Gli elementi si ritrovano per id (le istanze di link per UniqueId): se il
modello e' cambiato, alcune righe possono non trovare piu' l'elemento.

Formato: `{"format": "ESA_ClashFinder_Report", "version": 1, "model", "model_key",
"created", "seconds", "cancelled", "scope", "link_labels", "selection_ids", "elements": [...], "tests":
[{"name", "test", contatori e diagnostica, "clashes": [[indice A, indice B], ...]}]}`;
ogni elemento (`ElementInfo.to_dict()`) e' scritto una volta e richiamato per indice. La
coppia piu' vicina della diagnostica non viene salvata.

## Dati

| Dato | Dove |
| --- | --- |
| Libreria test | `pyRevit_ESA_ClashFinder_tests.json` nella cartella dati di pyRevit (`%APPDATA%\pyRevit`), creata al primo salvataggio; alla prima apertura contiene i test predefiniti (`cr.DEFAULT_TESTS`), vedi sotto |
| Import / export | stesso formato: `{"format": "ESA_ClashFinder_Tests", "version": 1, "tests": [...]}`; si accettano anche una lista di test o un test singolo |
| Impostazioni utente | `script.get_config('ESA_ClashFinder')`: `last_test` (test caricato alla riapertura). Sorgenti e link stanno nei test; il report non ha opzioni |

### Test predefiniti

| Test | Set A | Set B | Sorgente (A e B) |
| --- | --- | --- | --- |
| All vs All | tutte le categorie | tutte le categorie | Current selection |
| Pipes vs Ducts | tubi, raccordi, tubi flessibili (senza accessori) | canali, raccordi, canali flessibili (senza accessori) | Current selection |
| Mec vs Ele | distribuzione meccanica (canali e tubi con raccordi, accessori, flessibili) | passerelle, tubi protettivi e loro raccordi | Current selection |
| Ele vs Structures | passerelle, tubi protettivi e loro raccordi | muri, solai, tetti, pilastri, travi, fondazioni | Whole model |
| Openings vs Structures | Generic Models con *[Family] contains "Void Demand"* **oppure** *[Type] contains "VoidDemand"* (famiglie `e_GM.PH.VD_Void Demand_*`, tipo `PLH_VoidDemand` di `e_GM.MT_Placeholder`; `OPENING_RULES`) | travi, pilastri, solai | Whole model |
| Equipments vs Distribution | apparecchiature meccaniche ed elettriche | distribuzione meccanica ed elettrica | Whole model |

In una libreria gia'
esistente i predefiniti mancanti (per nome) si aggiungono una volta sola per
`DEFAULTS_VERSION` (config `ESA_ClashFinder` / `defaults_version`, `cr.pending_defaults`):
un test con lo stesso nome non viene toccato e un predefinito cancellato non torna, finche'
la versione non sale.

## Quick Clash

Secondo comando del pulldown (`QuickClash.pushbutton`), per eseguire in fretta. A sinistra
l'elenco dei test (un clic sceglie quello da eseguire), a destra sorgente nel modello aperto
e link di Set A e Set B: vengono dal test e si possono cambiare per l'esecuzione, senza
salvarli. Categorie, filtri e regole restano quelli del test. Nessuna tabella: una riga di
stato riporta conteggi e note. Libreria e `last_test` condivisi con Clash Setup; il report
e' lo stesso.

`QuickClashForm` (`clashfinder_quick_ui.py`) eredita da `ClashFinderForm` e cambia
`XAML_FILE`, `_find_controls`, `_init_controls`, `_load_editor`, `_read_editor_raw` (test
della libreria con sorgenti e link a schermo) e `_show_summary`; `_mark_dirty` non fa nulla.

## Moduli

Pulsanti e moduli stanno in `ClashFinder.pulldown`. I moduli e gli XAML sono in `ClashFinder.pulldown/lib`:
pyRevit mette la cartella `lib` di qualsiasi contenitore (qui il pulldown) nel percorso di
ricerca di tutti i suoi pulsanti, cosi' i tre comandi usano lo stesso codice senza copie. Gli XAML si risolvono da
`os.path.dirname(__file__)`, cioe' da `lib`.

| Modulo | Ruolo |
| --- | --- |
| `ClashSetup.pushbutton/ClashSetup_script.py` | controlli sul documento, crea il bridge e apre la finestra completa; `__persistentengine__ = True` |
| `QuickClash.pushbutton/QuickClash_script.py` | idem per la finestra rapida |
| `LastReport.pushbutton/LastReport_script.py` | riapre l'ultimo report del modello o ne importa uno (`ce.load_report`) |
| `lib/clashfinder_rules.py` | modello dati dei test, JSON, valutazione di categorie e regole sugli elementi |
| `lib/clashfinder_engine.py` | risoluzione degli insiemi sulle sorgenti, ricerca delle interferenze, `ElementInfo` |
| `lib/clashfinder_bridge.py` | coda di funzioni eseguite in un ExternalEvent |
| `lib/clashfinder_ui.py` + `ClashFinder_form.xaml`, `ClashFinder_save.xaml` | finestra completa e finestra del nome (Save as) |
| `lib/clashfinder_quick_ui.py` + `ClashFinder_quick.xaml` | finestra rapida (sottoclasse della completa) |
| `lib/clashfinder_report_ui.py` + `ClashFinder_report.xaml` | report |
