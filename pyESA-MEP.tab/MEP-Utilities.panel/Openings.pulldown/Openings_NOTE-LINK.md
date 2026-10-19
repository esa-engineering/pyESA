# Openings: modello di sole forometrie (piano, non implementato)

Piano di lavoro per il passaggio a un modello Openings separato, con MEP e ARC solo come link RVT. Il riquadro
(Area) e' gia' implementato sull'architettura attuale e va adattato, non rifatto.

## Flusso di lavoro

1. **Modello Openings** (contiene solo i fori; MEP e ARC sono link). *Place* con tre opzioni: **Area**, **MEP
   selection**, **ARC selection**. I fori si creano come oggi; a ogni Place si aggiorna il report JSON di lavoro.
2. Quando uno o piu' link vengono aggiornati (Reload del link in Revit): **New Report** sul modello Openings.
   Controlla i fori contro i MEP e gli ARC dei link, scrive gli stati sui fori, aggiorna il JSON di lavoro.
3. Sistemate le forometrie: **Export** del report JSON in una cartella condivisa.
4. **Modelli MEP / ARC**: *Last Report* carica il JSON esportato e mostra fori ed elementi del modello aperto.
   *Select* / *Show* li trovano subito. Opzionale: *Write status* in un parametro.

## Decisioni prese

| Tema | Decisione |
| --- | --- |
| Modello misto MEP + fori | **No**: modelli sempre separati. Nel modello Openings non ci sono MEP propri |
| Parametri sui MEP | **Nessuno** (tolti `write_mep()`, `mep_candidates()` e la parte MEP di `plan_writes()`). Il legame sta sul foro e nel JSON |
| MEP e ARC da quali link | Ogni link puo' dare sia MEP sia muri, solai e strutture. **Place**: sezione "Linked models" nella finestra, spunta = leggi (MEP e ARC), ricordata per modello (`place_skipped_links`), link nuovi spuntati; con elementi gia' selezionati vale l'ultima scelta (decisione del 19.10.2026, prima era "tutti, nessuna scelta"). **New Report / Last Report**: nessuna scelta, si leggono tutti i link caricati e si controllano i fori dello scope |
| JSON | Uno **di lavoro per utente** (`script.get_document_data_file`), aggiornato a ogni Place e New Report, mai salvato altrove. **Export** / import: la cartella si sceglie ogni volta (proposta l'ultima usata) |
| Chi cerca i fori nuovi | **Place&Check** (rinominato da Place): verifica dove servono fori nuovi e li crea; esito sempre in una finestra XAML, anche "no new openings needed". **New Report** verifica solo lo stato dei fori presenti: niente *Missing openings* (19.10.2026) |
| Pannello di output di pyRevit | Non si apre se non serve: Openings mostra tutto in finestre XAML |
| Foro reale nei modelli ARC | **Non interessa**: i fori restano non ospitati e non tagliano i link; il foro vero si approva dopo |
| Fori senza piu' elementi dentro | Si **tengono**: con un Id MEP associato ma nessun elemento al loro interno risultano *Not Associated* nel report |
| Revit minimo | **2023** (niente 2022): `Selection.SetReferences` utilizzabile, da verificare sui link |
| Percorso dei modelli | Quando si selezionano elementi di un link, il report mostra modello e percorso del file |
| Fasi e workset dei link | Come MEPQTO (vedi "Link: fasi e workset") |
| Fori nei modelli MEP / ARC | Il modello Openings e' **linkato** li': Show zooma sul foro vero nel link. Senza link, ripiego sul bounding box salvato nel JSON |
| Scrittura nei modelli MEP / ARC | Solo visualizzazione; *Write status* **opzionale** (pulsante a parte, scrive in un parametro se c'e') |

## Riferimenti agli elementi dei link

Gli Id di documenti diversi si sovrappongono: ogni riferimento porta il modello.

- **Chiave del modello** = titolo del file linkato (`link_doc.Title`, senza estensione). E' stabile tra sessioni ed e'
  la stessa nei modelli MEP / ARC, che cercano se stessi nel JSON con `doc.Title`. Il percorso (`PathName`, o
  quello del centrale) va nel JSON e nel report: due link con lo stesso titolo da cartelle diverse si segnalano.
- **Sul foro**, `e_GEN_RelatedElement_ID` diventa `<modello>:<id>,<modello>:<id>` (es. `PRJ_MEC:123456`). Nel
  modello Openings i fori sono scrivibili, quindi i parametri restano sui fori. Un valore senza `:` (fori di prima)
  non trova il suo elemento: nessuna migrazione, vale la regola dei fori orfani qui sotto.

### Fori orfani (decisione 19.10.2026)

Un foro con Id associati che non trova piu' i suoi elementi (cancellati, Id senza modello, foro spostato) **non si
cancella mai da solo**: resta, e' *Not Associated* nel report, l'utente sceglie se cancellarlo.

- Se alla verifica dentro il foro ci sono **altri elementi MEP** (es. foro spostato a mano), resta comunque *Not
  Associated*, con una nota che elenca gli elementi trovati dentro. Il legame **non** si aggiorna da solo.
- **Recreate** su quel foro riscrive gli Id associati (`<modello>:<id>`) con gli elementi trovati dentro e
  ricalcola dimensioni e stato. E' l'unico modo per riassociare, e vale anche come migrazione dei fori vecchi.

**Differenza con il codice di allora** (fatto in fase 4, 19.10.2026; in piu': ricerca nel bounding box del foro,
perche' la scansione non vede i MEP mai associati). `check_opening()` associa in automatico gli elementi
non elencati che attraversano il foro (`ELEMENT_ADDED`, "crossing it, now associated"): un foro orfano spostato su
altri tubi diventa *Partial* / *To Be Refreshed*, non *Not Associated*. `recreate_openings()` salta i fori senza
`items` (`skipped`), quindi oggi Recreate non puo' riassociare un orfano. Da fare:

- nessuno degli elementi elencati e' piu' presente: gli `extra` **non** si aggiungono, stato `OBSOLETE`, nota con
  gli elementi trovati dentro (nuovo stato di riga, es. `ELEMENT_FOUND_INSIDE`, non scritto sul foro);
- Recreate di un orfano usa quegli elementi trovati dentro invece di saltarlo; `recreate_plan()` lo dice.
- **Nel JSON** anche l'`UniqueId`: se nel modello MEP un elemento viene cancellato e ricreato, l'Id cambia e lo si
  vede subito.
- **Piu' istanze dello stesso link** (copie del link): la chiave del modello non basta, serve l'istanza. Proposta:
  `<modello>#<n>` solo quando le istanze sono piu' di una, `n` in ordine di `UniqueId` dell'istanza.

## Cosa cambia nel codice del modello Openings

1. **`MepRef`**: riferimento a un MEP di un link (documento, transform del link, elemento, chiave). Espone gia' in
   coordinate host: curva, direzione, sezione con isolamento (`GetInsulationIds(link_doc, ...)`), disciplina,
   categoria. Tutte le funzioni che oggi usano `mep.Location.Curve`, `mep.Id`, `mep_section(mep)`, `discipline(mep)`
   passano da `MepRef`. E' il grosso del lavoro: `find_crossings`, `crossing_item`, `merge_overlapping`,
   `recross`, `check_opening`, `scanned_plane`, `outside_meps`, `handled_by_opening`, `missing_openings`,
   `opening_kind`, le righe del report.
2. **Ricerca dei MEP nei link**: `mep_near()`, `meps_in_area()` e `hosts_from_meps()` girano su `model_sources()`
   senza l'host (il modello Openings non ha MEP); outline trasformato come in `hosts_in_area()`.
3. **Selezione MEP** in Place / New Report: `PickObjects(ObjectType.LinkedElement)` con un filtro per tubi,
   canali, passerelle e conduit. **ARC selection**: com'e' oggi (`PointOnElement`, gia' compatibile con i link).
4. **Fine della scrittura sui MEP**: `plan_writes()` tiene solo la parte dei fori. Lo stato di un MEP (il peggiore
   dei suoi fori, o *Not Associated*) si calcola e va solo nel report e nel JSON.
5. **Interactive Report** nel modello Openings: Show / Select su un MEP linkato selezionano il foro e zoomano sul
   bounding box del MEP (trasformato). Da verificare se `Selection.SetReferences` (Revit 2023+) seleziona elementi dei
   link in modo affidabile; se si', lo si usa dove disponibile.
6. **Link non caricato**: i fori che citano un modello assente non diventano *Not Associated*. Nuovo stato interno
   *Link not loaded*: stato scritto invariato, nota nel report.

## Report JSON

File di lavoro: solo dove sta oggi (`script.get_document_data_file`, per utente), nessuna copia condivisa. Lo stato
che conta tra colleghi e' quello scritto sui fori (parametri), che il New Report rilegge. Export: cartella scelta
ogni volta, `<Openings>_Openings.json` sovrascritto (quello che leggono i modelli MEP / ARC) piu' una copia datata
`<Openings>_Openings_AAAAMMGG-HHMM.json` come storico. Scrittura via `.tmp` come in MEPQTO.

```json
{
  "format": "ESA_Openings_Report", "version": 1,
  "openings_model": "PRJ_OPN", "date": "16.10.2026 15:30", "user": "...", "settings": {"clearance_mm": 25},
  "models": {"PRJ_MEC": {"path": "...", "kind_seen": ["MEP"], "phase": "New Construction",
                         "transform": [[1, 0, 0], [0, 1, 0], [0, 0, 1], [0, 0, 0]]},
             "PRJ_ARC": {"path": "...", "kind_seen": ["ARC"], "phase": "...", "transform": "..."}},
  "openings": [{
    "code": "123456", "unique_id": "...", "status": "Valid", "notes": [],
    "kind": "wall", "shape": "rect", "size": "300 x 200 x 250",
    "box": {"min": [0, 0, 0], "max": [0, 0, 0]},
    "elements": [{"model": "PRJ_MEC", "id": 98765, "unique_id": "...", "category": "Pipes", "state": "inside"}],
    "hosts": [{"model": "PRJ_ARC", "id": 4321, "category": "Walls"}]
  }],
  "missing": [{"elements": ["PRJ_MEC:98766"], "hosts": ["PRJ_ARC:4321"], "size": "..."}]
}
```

- `box` e' in coordinate interne del **modello Openings**. Nel modello MEP / ARC si porta in coordinate locali con
  il transform del link Openings. Senza link Openings: con l'inverso di `models.<modello>.transform` (BasisX, BasisY,
  BasisZ, Origin del link nel modello Openings, in piedi), gia' nella versione 1 dello schema.

## Link: fasi e workset (come MEPQTO)

- **Fase.** Un elemento del link ha le fasi del *suo* documento, che non sono quelle del modello Openings. Per ogni
  link si cerca la fase con lo **stesso nome** di quella scelta nell'host; si leggono solo gli elementi esistenti in
  quella fase (creati prima o in quella fase, non demoliti). Se il link non ha una fase con quel nome, il link si
  salta e il report lo dice (non si indovina).
- **Workset.** Si escludono per **nome**, non per Id: gli Id dei workset cambiano da documento a documento, i nomi
  no. Nella finestra di New Report l'elenco dei nomi dei workset di tutti i link, spunta = leggi. I nomi mai visti
  prima compaiono spuntati e segnalati, cosi' un workset nuovo (es. "MEP_Provvisorio") non sparisce in silenzio.
  Config `ESA_Openings`: `last_included_worksets`, `last_known_worksets`.
- Le chiavi sono in inglese, gli stati sono i testi scritti (`Valid`, `To Be Refreshed`, `Not Associated`).

## Last Report nei modelli MEP / ARC

Last Report riconosce il caso: nessun foro gestito nel modello aperto, e il modello compare in `models` del JSON.

- Il file da caricare: l'ultimo usato per quel modello (config `ESA_Openings`, `report_files` come
  `"<doc key>::<path>"`, come i `project_files` di MEPQTO), altrimenti `forms.pick_file`.
- Mostra solo i fori che citano il modello aperto, piu' i *Missing openings* sui suoi elementi; riepilogo e Interactive
  Report di sempre, con la data dell'export in evidenza.
- **Select / Show**: gli elementi del modello aperto si selezionano direttamente (sono host qui); i fori tramite il
  link Openings (selezione del link + zoom sul `box`, o `SetReferences` se affidabile).
- Spenti: Create / Recreate, Move, Delete, Refresh (il modello Openings non e' quello aperto).
- **Write status** (opzionale, pulsante a parte o nella finestra): scrive lo stato del MEP in un parametro di testo
  (nome in Settings), una transazione. Elementi del JSON non piu' presenti: elenco nel report.

## Fasi

0. **Fatto (19.10.2026):** risolto il merge `79cb08f` committato con i marker di conflitto; tenuta ovunque la
   versione con Area (commit `152f181`).
   **Fase 1 fatta e provata in Revit sul modello MEP + fori (commit `f545a96`).** Moduli in `openings_modules/`,
   stato del comando in `openings_core.ctx`.
   **Fase 3 scritta (19.10.2026), da provare in Revit:** `MepRef` in tutto il codice (Place, New Report, Last
   Report), MEP cercati anche nei link, riferimenti `<modello>:<id>`, selezione MEP nei link, modelli e percorsi nel
   report, link non caricato = riferimento tenuto e stato invariato. Provata in Revit, commit `adeb102` (con la
   scelta dei link in Place&Check e la finestra dell'esito). Fase 2 (fasi, workset, lettura unica) ancora da fare.
   **Fase 4 scritta (19.10.2026), da provare in Revit:** formato dei riferimenti
   `111,222; PRJ_MEC: 333,444` (`parse_refs` / `format_refs`, legge anche i formati vecchi), niente scrittura sui MEP
   (tolti `write_mep`, `mep_candidates`, `plan_writes` e le righe "Stale Ids" del report), fori orfani come deciso
   sotto (ricerca nel bounding box, Recreate sugli elementi trovati dentro). Verifiche fatte fuori da Revit: compilazione con l'IronPython 2.7 di pyRevit
   (`bin\netfx\engines\IPY2712PR`), nessun nome globale indefinito nella catena di `import *`, nessun nome definito
   in due moduli, AST di 181 blocchi identico all'originale (con `ctx.x` riportato al nome globale); diversi solo
   `load_settings`, `bundle_path`, `revit_owner`, `family_line`, `place_main`, `ReportController` (voluti).
1. **Divisione in moduli** come MEPQTO: un solo `openings_*.py` per tema (es. `openings_geometry.py`,
   `openings_check.py`, `openings_report.py`) in `Openings.pulldown/openings_modules/` (senza suffisso, pyRevit la
   ignora), aggiunta a `sys.path` dai tre `script.py`, che restano entry point con `MODE`. Ci finiscono anche
   `OpeningsCheck_ui`, `OpeningsReport_xlsx` e le XAML oggi copiati tre volte. Nessun modulo fuori da Openings la
   importa; eccezione documentata in CLAUDE.md, la regola "no shared library" resta. Refactor puro, nessun cambio di
   comportamento: si prova che fa esattamente quello di oggi, e che Reload ricarica i moduli della cartella.
2. **Lettura dei link una volta sola**: record Python (curva, sezione, categoria, chiave) per MEP e host di tutti i
   link, con fase e workset come sopra; poi i calcoli girano sui record. Misurare i tempi su un progetto vero.
3. **`MepRef`** + Place (Area, MEP selection, ARC selection) nel modello Openings. Verifica: fori identici a quelli
   fatti oggi con i MEP nell'host.
4. **New Report** e Interactive Report con i MEP linkati; niente scrittura sui MEP; stati *Link not loaded* e *Not
   Associated* (fori orfani, mantenuti, con nota se dentro ci sono altri elementi; Recreate li riassocia).
5. **JSON** di lavoro + Export (schema v1 con `path`, `phase`, `transform`).
6. **Last Report** nei modelli MEP / ARC (import, Select / Show con `SetReferences`, ripiego zoom + box).
7. Write status (opzionale).

Ogni fase si chiude e si prova in Revit prima della successiva.

## Punti aperti

- **`SetReferences` sui link** (Revit 2023+): da verificare che selezioni davvero gli elementi del link.

## Decisi il 19.10.2026

- **Moduli condivisi tra i tre pulsanti**: cartella interna `openings_modules/` + `sys.path` (fase 1). Scartate una
  `lib/` di estensione (cambierebbe la regola "no shared library") e le copie identiche (tre modifiche per ogni
  cambio, divergenze silenziose con Rocket mode).
- **Caso misto** (foro con alcuni elementi originali + uno nuovo che lo attraversa): resta l'associazione automatica
  di oggi, *To Be Refreshed*.
- **Export concorrente** dalla stessa cartella da due modelli Openings diversi: nomi distinti per modello, nessun
  conflitto; dallo stesso modello: last writer wins (lo storico datato copre).
- Se un link e' sia MEP sia ARC (es. modello unico di disciplina), non cambia nulla: "tutti i link" lo copre.
