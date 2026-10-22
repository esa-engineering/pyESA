# Clash Finder: note di sviluppo

Registro dei tentativi fatti sul codice condiviso di `ClashFinder.pulldown` (Clash Setup,
Quick Clash, Last Report), per non ripetere soluzioni gia' provate. Non e' il README e non e'
lo storico revisioni degli script. Voci in ordine cronologico inverso.

## 22.10.2026, versioni Setup 2.1 / Quick 1.1 / Last Report 1.2

**Richiesta.** Quattro ritocchi alla finestra Clash Setup e a Quick Clash: riga del test piu'
corta quando i link sono molti, scelta dei link piu' comoda, etichette piu' chiare, guida per
l'utente inesperto.

**1. Riga del test con molti link: funziona.** `describe_sources(set_data, compact=False)` in
`clashfinder_rules.py`: con `compact=True` due o piu' link diventano "N links". Lo usa solo
`describe_scope()`, cioe' la riga grigia sotto il nome del test; il tooltip della riga
(`describe_set`) e il riepilogo di Quick Clash (`describe_sources` senza `compact`) mostrano
ancora i nomi. Soglia a due link: con uno solo il nome serve a riconoscere il test.

**2. Scelta dei link da un elenco a discesa: funziona, da provare in Revit.**

- Prima: `ListBox` alto 60 px, tre link visibili alla volta, scroll per gli altri.
- Ora: `Button` (`btn_links_<a|b>`) + `Popup` (`pop_links_<a|b>`) con la stessa `ListBox` di
  caselle (`lst_links_<a|b>`) dentro. Il testo del bottone riassume la scelta
  (`_update_links_button`).
- Il `Popup` ha `StaysOpen="False"`: un clic sul bottone con il popup aperto lo chiude al
  mouse down e subito dopo il `Click` lo riaprirebbe. Soluzione: `popup.Closed` registra
  l'ora, `Click` non riapre se sono passati meno di 250 ms.
- Non scelto: `ComboBox` con `CheckBox` come elementi. Il `ComboBox` non mostra bene piu'
  voci spuntate insieme (il testo di selezione e' quello di un solo elemento) e andrebbe
  forzato con trucchi sul template. `ToggleButton` con `IsChecked` legato a `IsOpen`: stesso
  problema di riapertura, senza una soluzione piu' semplice della guardia a tempo.
- Quick Clash ha un suo XAML (`ClashFinder_quick.xaml`) ma usa `_fill_link_lists`, che ora
  aggiorna il bottone: la prima versione di questa modifica lasciava la `ListBox` semplice
  in Quick e avrebbe dato `AttributeError` all'apertura (trovato rileggendo la nota, prima
  di una prova in Revit). Corretto: stesso bottone + popup anche in Quick, e il collegamento
  del popup sta in `_wire_links_dropdown(side)`, chiamato da `_wire_side` (Setup) e da
  `_init_controls` di Quick.

**3. Etichette: funziona.** "Nothing (links only)" diventa "Links only"
(`SOURCE_LABELS` in `clashfinder_engine.py`, condivisa da Setup e Quick); "Plus elements of
these links:" diventa "Links" (form e quick). I test salvati non cambiano: nel JSON c'e' il
valore interno della sorgente, non il testo.

**3b. Riga di stato di Quick Clash su tre righe: da provare in Revit.** Prima un solo
`TextBlock` rosso con nome del test, elementi selezionati, fase, Set A, Set B e note: tutto
rosso anche quando i conteggi erano a posto, e il colore era uno per tutto il testo. Ora
`clashfinder_ui.set_status_lines(text_block, [(testo, pennello)])` riempie le `Inlines`
con una `Run` colorata per riga: 1) fase della vista (rosso solo se `phase_problem`, grigio
verde se c'e'), 2) Set A e 3) Set B (verde se host + link > 0, grigio se 0). Tolti dalla riga
il nome del test (e' evidenziato nell'elenco) e il conteggio della selezione di Revit (e'
nel conteggio del Set con sorgente selezione). Le note del riepilogo restano, in grigio, ma senza "Set A is empty" / "Set B is empty" (le
produce `set_note` in `clashfinder_engine.py` unite con "; " in un'unica nota per test, e
ripetono il conteggio a zero gia' in grigio): il filtro e' in `_show_summary` di Quick, non
nel motore, perche' Clash Setup e il report le usano ancora.
Il messaggio "No test selected." non e' piu' rosso (non e' un errore). Clash Setup ha ancora
la sua riga di stato a un colore: non allineata.

**3c. Revisione per l'ottimizzazione: solo la misura dei tempi, nessun cambio di
comportamento.** Aggiunto `log_timing` in `clashfinder_engine.py`: scrive nel log di debug di
pyRevit (si vede solo in modalita' debug, Ctrl + clic sul pulsante) la durata di `summarize`,
della risoluzione dei Set nel Run, di ogni test (con elementi di A, B host, B link e numero di
clash) e del totale. Chiamata dentro `try/except`: un errore di log non ferma mai il comando.
Candidati **non applicati** perche' cambiano il comportamento o hanno rischio, da decidere con i
numeri alla mano:

- Controllo host contro host: un `FilteredElementCollector(doc, host_ids)` con tutti gli id di
  B per ogni elemento di A, costo come |A| x |B|. Idea: riusare `BoxGrid` anche per B host e
  passare al collector solo i candidati vicini. Rischio medio (parita' del box in Python con il
  filtro nativo).
- Riepilogo ricalcolato da zero a ogni clic (`summarize` risolve entrambi i Set): cache per
  sorgente / link / categoria, oppure saltare in `CompiledSet.matches` i controlli ripetuti di
  `is_model_element` quando la categoria e' gia' filtrata in modo nativo
  (`WhereElementIsViewIndependent` al posto di `OwnerViewId`).
- Memoria: `Item.solids()` tiene in cache i solidi trasformati di tutto B linkato per tutto il
  run; limitare o liberare la cache.
- Diagnostica (`box_gap`, `track_nearest`) calcolata anche per i candidati scartati: lasciata
  com'e', alimenta le note del report.

**3d. Riquadro "How sets are built" di Clash Setup: ora solo al passaggio del mouse.** Il
riquadro grigio in fondo all'editor occupava spazio fisso. Al suo posto una riga piccola
"ℹ️ How sets are built", allineata a destra, con lo stesso testo in un `ToolTip` (larghezza
massima 520, ritardo 200 ms, resta visibile fino a 60 s per poterlo leggere). Il testo e'
quello di prima, con una riga vuota fra i quattro paragrafi. Da provare in Revit.

**3e. Info in alto a destra e piu' spazio ai filtri (Clash Setup).** La riga "How sets are
built" in fondo e' sparita: nella riga del titolo c'e' ora una "ℹ️" a destra, con un `ToolTip`
(560 px, 200 ms di ritardo, visibile fino a 2 minuti) che spiega prima il flusso di lavoro in
sei passi (scegli o crea un test, sorgenti e link, categorie, Refresh, Run, Save) e poi, sotto,
"How sets are built" com'era. Le due `ListBox` dei filtri di vista passano da `Height` 100 a
135 (`MinHeight` 100; la prima versione era 190 ma con la finestra a 900 faceva comparire lo scroll dell'editor, circa 50 px di troppo): sono le due da cui si scorre di piu' con molti filtri. L'altezza della
finestra resta 900: con filtri alti il contenuto dell'editor supera di poco lo spazio su
schermi piccoli, e li' lavora lo `ScrollViewer` dell'editor, gia' presente. Da provare in
Revit (verificare che a 125% non serva scorrere per arrivare ai pulsanti).

**3f. "ℹ️" anche in Quick Clash e nel report; tooltip dei pulsanti riscritti.** Stesso schema di
Clash Setup (riga del titolo in un `Grid`, icona a destra, `ToolTip` con ritardo 200 ms): in
Quick Clash spiega i cinque passi e che categorie, filtri e regole si cambiano in Clash Setup;
nel report spiega come leggerlo (gruppi, clic su una riga, Isolate selected / Hide overlay,
ricerca, Refresh con NEW e resolved, Export JSON / CSV). Last Report riapre la stessa finestra
del report, quindi ha la stessa icona. I tre `bundle.yaml` erano molto piu' lunghi di quanto
chiede CLAUDE.md (una frase, al massimo quattro righe, niente note tecniche): riscritti con
`|-`, `Requires: a model view with a phase.` per Setup e Quick (il Run si blocca senza fase),
`CLICK` / `SHIFT + CLICK` solo per Last Report, che ha davvero due modi. Setup e Quick non
hanno `SHIFT + CLICK`, quindi non lo citano. Il contenuto tecnico tolto resta nel README.

**4. Guida per l'utente inesperto: messa da parte, non e' in questa versione.**
Tentativi, in ordine:

- Pulsante "?" che apre il `README.md` con `os.startfile`. Scartato: il README e' tecnico e in
  italiano, scritto per chi legge il codice; si apre nel Blocco note o in un editor di codice.
- Finestra di aiuto statica in inglese, tre colonne (pick a test, define the sets, run and
  read). Funzionava sulla carta, scartata prima di provarla in Revit quando e' arrivata la
  presentazione gia' fatta.
- Guida a slide: 17 slide da una presentazione esistente, esportate in PNG 1920x1080 con
  Edge headless e mostrate in una finestra con Previous / Next. Costruita e compilata
  (IronPython), non provata in Revit. Messa da parte su richiesta: i file
  (`clashfinder_guide_ui.py`, `ClashFinder_guide.xaml`, cartella `ClashFinder_guide/`) non sono
  nel commit e il "?" e' stato tolto da Setup, Quick e report. Da decidere: lingua delle slide
  (oggi italiano, il resto della UI e' inglese), screenshot da rifare con le nuove etichette,
  eventuale riduzione dei PNG a 1280x720.
- Note tecniche utili se si riprende: l'azione `read` con `paths` non scarica gli asset (serve
  `path` singolo per id); `x-icon` non si renderizza fuori dal runtime delle slide; per
  Edge serve `box-sizing: border-box` su tutti gli elementi, altrimenti le card si sovrappongono.

**Verifica.** Compilazione IronPython 2.7 di tutti i `.py` e controllo XML degli XAML. Nessuna
prova in Revit.
