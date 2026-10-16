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
| MEP e ARC da quali link | **Tutti i link caricati**: ogni link puo' dare sia MEP sia muri, solai e strutture. Nessuna scelta nella finestra |
| JSON | Uno **di lavoro**, aggiornato a ogni Place e New Report, piu' l'**Export** in una cartella condivisa scelta dall'utente (ricordata) |
| Fori nei modelli MEP / ARC | Il modello Openings e' **linkato** li': Show zooma sul foro vero nel link. Senza link, ripiego sul bounding box salvato nel JSON |
| Scrittura nei modelli MEP / ARC | Solo visualizzazione; *Write status* **opzionale** (pulsante a parte, scrive in un parametro se c'e') |

## Riferimenti agli elementi dei link

Gli Id di documenti diversi si sovrappongono: ogni riferimento porta il modello.

- **Chiave del modello** = titolo del file linkato (`link_doc.Title`, senza estensione). E' stabile tra sessioni ed e'
  la stessa nei modelli MEP / ARC, che cercano se stessi nel JSON con `doc.Title`.
- **Sul foro**, `e_GEN_RelatedElement_ID` diventa `<modello>:<id>,<modello>:<id>` (es. `PRJ_MEC:123456`). Nel
  modello Openings i fori sono scrivibili, quindi i parametri restano sui fori. Un valore senza `:` (fori di prima)
  viene letto come elemento dell'host e risulta *deleted* alla prima New Report: va deciso se migrarli (vedi Punti
  aperti).
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

File di lavoro: dove sta oggi (`script.get_document_data_file`, per utente) **piu'** una copia accanto al centrale
del modello Openings, cosi' il New Report di un collega riparte dallo stesso stato. Export: `<Openings>_Openings.json`
sovrascritto nella cartella scelta (quello che leggono i modelli MEP / ARC) piu' una copia datata
`<Openings>_Openings_AAAAMMGG-HHMM.json` come storico. Scrittura via `.tmp` come in MEPQTO.

```json
{
  "format": "ESA_Openings_Report", "version": 1,
  "openings_model": "PRJ_OPN", "date": "16.10.2026 15:30", "user": "...", "settings": {"clearance_mm": 25},
  "models": {"PRJ_MEC": {"path": "...", "kind_seen": ["MEP"]}, "PRJ_ARC": {"path": "...", "kind_seen": ["ARC"]}},
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
  il transform del link Openings. Senza link Openings: con l'inverso del transform con cui il modello MEP / ARC e'
  linkato nel modello Openings, salvato in `models` (da aggiungere: `transform`).
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

## Fasi proposte

1. `MepRef` + ricerca MEP nei link + Place (Area, MEP selection, ARC selection) nel modello Openings. Verifica:
   fori identici a quelli fatti oggi con i MEP nell'host.
2. New Report e Interactive Report con i MEP linkati; niente scrittura sui MEP; stato *Link not loaded*.
3. JSON di lavoro + Export.
4. Last Report nei modelli MEP / ARC (lettura, Select / Show).
5. Write status (opzionale).

Ogni fase si chiude su un branch e si prova in Revit prima della successiva. Dato il volume, valutare in fase 1 la
divisione di `script.py` (oltre 2.500 righe) in moduli come MEPQTO (`openings_geometry.py`, `openings_check.py`,
`openings_report.py`). Si toglierebbe anche la regola delle tre copie dello script.

## Punti aperti

- **Fori gia' posati** con Id senza modello: migrazione una tantum (pulsante o prima New Report che chiede il
  modello sorgente) oppure si ripartono i progetti da zero?
- **Revit minimo** per `Selection.SetReferences` sui link: se serve anche Revit 2022, resta il ripiego zoom + box.
- **Export concorrente** dalla stessa cartella da due modelli Openings diversi: nomi distinti per modello, nessun
  conflitto; dallo stesso modello: last writer wins (lo storico datato copre).
- Se un link e' sia MEP sia ARC (es. modello unico di disciplina), non cambia nulla: "tutti i link" lo copre.
