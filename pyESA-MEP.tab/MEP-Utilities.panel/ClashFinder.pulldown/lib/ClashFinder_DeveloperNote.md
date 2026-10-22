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
