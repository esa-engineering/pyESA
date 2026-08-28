# TagLinkedRooms - Note di sviluppo v3 → v4.1

Registro della sessione di analisi, riprogettazione e debug dello script
`TagLinkedRooms_script.py`.

| | |
|---|---|
| **Script** | `pyESA.tab/Utilities.panel/Utilities4.stack/Tag.pulldown/TagLinkedRooms_new.pushbutton/TagLinkedRooms_script.py` |
| **Versioni** | da 3.0 a 4.1 |
| **Date** | 27-28 agosto 2026 |
| **Ambiente di test** | Revit 2026.4, pyRevit 5.3.1.25308 |
| **Autori script** | Tommaso Lorenzi, Andrea Patti |

> **Nota sui dati.** In questo documento i modelli usati per i test sono
> indicati con etichette neutre (`Modello A`, `Modello B`) invece dei codici
> commessa reali, perché il file vive in un repository git. Tutti i valori
> tecnici, le quote e i numeri dei report sono quelli effettivi e non sono
> stati alterati.

---

## Indice

1. [Verifica compatibilità Revit 2026](#1-verifica-compatibilità-revit-2026)
2. [Come funzionava la v3](#2-come-funzionava-la-v3)
3. [Limiti individuati nella v3](#3-limiti-individuati-nella-v3)
4. [L'approccio alternativo proposto](#4-lapproccio-alternativo-proposto)
5. [La riscrittura v4.0](#5-la-riscrittura-v40)
6. [Scelta utente fra L1 e L2](#6-scelta-utente-fra-l1-e-l2)
7. [Debug: tutte le rooms escluse](#7-debug-tutte-le-rooms-escluse)
8. [La causa reale: Elevation vs ProjectElevation](#8-la-causa-reale-elevation-vs-projectelevation)
9. [v4.1: fascia verticale](#9-v41-fascia-verticale)
10. [Punti aperti e non verificati](#10-punti-aperti-e-non-verificati)
11. [Riferimento rapido](#11-riferimento-rapido)

---

## 1. Verifica compatibilità Revit 2026

**Domanda iniziale:** lo script v3 è compatibile con Revit 2026?

**Esito: sostanzialmente sì**, nessun uso diretto di API rimosse.

### Già a posto

`ElementId.IntegerValue`, la rottura principale del 2026, era già gestita
dall'helper con switch su `RVT_VER >= 2026`. Verificato con grep: nessun
accesso diretto a `.IntegerValue` altrove nel file, tutti i punti d'uso
passavano dall'helper. Stessa convenzione già presente in
`TagInViews_script.py`.

Il valore di ritorno in 2026 è `System.Int64` invece di `Int32`, ma viene usato
solo per formattazione stringa e come chiave di tuple e set: in IronPython
`long` e `int` hanno hash e uguaglianza compatibili, quindi il set delle chiavi
dei tag esistenti continua a funzionare senza modifiche.

Altre API usate e ancora valide in 2026: `GetCropRegionShapeManager()` /
`GetCropShape()`, `PlanViewRange` con le proprietà statiche
`Unlimited`/`Current`/`LevelAbove`/`LevelBelow`, `RoomTag.TaggedRoomId`,
`LinkElementId`, `ElementTransformUtils.RotateElement`,
`BuiltInParameter.SYMBOL_NAME_PARAM`, `FamilySymbol.IsActive` / `Activate()`,
`XYZ.AngleOnPlaneTo`.

### Punti segnalati

1. **`doc.Create.NewRoomTag`** - fa parte di `Autodesk.Revit.Creation.Document`,
   famiglia che Autodesk sta progressivamente deprecando in favore delle
   factory statiche. Risulta ancora presente in 2026 e non obsoleta, ma è il
   punto da testare per primo. Non esiste un equivalente `IndependentTag.Create`
   che accetti una `LinkElementId`, quindi non è sostituibile.
   **Confermato funzionante** nei test successivi su Revit 2026.4.
2. **Versione pyRevit** - Revit 2026 gira su .NET 8 e richiede pyRevit 5.x.
   Confermato: l'ambiente di test usa pyRevit 5.3.1.
3. **`int(app.VersionNumber)`** - solleverebbe `ValueError` a import time,
   prima di qualsiasi try/except, se il valore non fosse numerico. Rischio
   basso, hardening applicato in v4.0.

### Osservazioni non legate al 2026

- Il test `if selected_options is None` gestiva l'annullamento della dialog, ma
  una conferma con zero voci selezionate restituisce una lista vuota, non
  `None`: tutte le opzioni risultavano disattivate invece di applicare i
  default. Corretto in v4.0 con `if not selected_options`.
- `script.exit()` solleva `SystemExit`, che discende da `BaseException` e non
  viene intercettato dall'`except Exception` finale. Comportamento corretto
  (l'annullamento non compare come errore critico), segnalato come dipendenza
  sottile da non rompere.

---

## 2. Come funzionava la v3

Architettura a due tempi: quattro dialog di setup, un pre-calcolo geometrico
globale, poi un doppio loop viste × rooms dentro una singola transazione.

La scelta più importante era il pre-calcolo: la geometria delle rooms veniva
risolta una volta sola e riusata per ogni vista. Con 500 rooms × 30 viste sono
15.000 iterazioni di aritmetica pura invece di 15.000 letture di parametri.

### Selezione della sorgente

Collector di `RevitLinkInstance`, filtrati per `GetLinkDocument() is not None`
(link caricati). Dalla scelta derivavano tre variabili:

| variabile | modello host | modello linkato |
|---|---|---|
| `room_doc` | `doc` | `link.GetLinkDocument()` |
| `link_transform` | `Transform.Identity` | `link.GetTotalTransform()` |
| `link_id_key` | `-1` | ID istanza del link |

### Selezione delle rooms candidate

Filtro di piazzamento: `room.Area > 0 and room.Location is not None`. Copre
tre stati problematici:

- **Unplaced**: `Location` è None → esclusa
- **Not Enclosed**: `Area` è 0 → esclusa
- **Redundant**: entrambe hanno `Area > 0` → **entrambe passano**, producendo
  due tag sovrapposti. Unico caso che sfuggiva.

### Geometria di ogni room

Tre numeri: un punto XY e due quote.

**Punto XY** (`get_room_point`): `LocationPoint.Point`, con fallback sul centro
del bounding box. Il fallback per una room a L può cadere fuori dalla room, ma
in pratica non si attiva mai.

**Estensione verticale** (`get_room_z_span`), ricostruita dai parametri:

```
z_min = elevazione(Level) + ROOM_LOWER_OFFSET          # Base Offset
z_max = elevazione(ROOM_UPPER_LEVEL) + ROOM_UPPER_OFFSET   # Upper Limit + Limit Offset
```

Con fallback su `ROOM_HEIGHT` e infine su `DEFAULT_ROOM_HEIGHT = 8.0` piedi se
lo span risultava degenere.

**Trasformazione in coordinate host** - dettaglio fatto bene: le quote non
venivano traslate sommando un offset, ma **trasformando due punti reali** e
leggendone la Z, con swap finale per il caso di link ribaltato.

### Caratterizzazione della vista

**View range** - la parte più delicata. `PlanViewRange.GetLevelId(plane)` non
restituisce sempre un livello reale: può restituire uno di quattro ElementId
sentinella con valori negativi. Risoluzione:

| sentinella | risultato |
|---|---|
| `Unlimited` | `None`, nessun limite su quel lato |
| `Current` o `InvalidElementId` | `GenLevel.Elevation + offset` |
| `LevelAbove` | primo livello sopra `GenLevel` + offset |
| `LevelBelow` | ultimo livello sotto `GenLevel` + offset |
| altro | elevazione del livello vero + offset |

Il limite inferiore effettivo era `low = min(bottom, depth)`: la View Depth può
scendere sotto il Bottom, e in quel caso comanda lei.

**Test orizzontale** (`build_xy_test`) - quattro strategie in cascata, e qui
stava la novità dichiarata della v3:

1. **Crop shape reale** (se `CropBoxActive`): `GetCropShape()` restituisce
   `CurveLoop` **già in coordinate modello**, tessellati in un poligono e
   testati con ray casting. Punto chiave: la rotazione da scope box è gestita
   implicitamente, senza conoscerne l'angolo. Funziona anche con crop non
   rettangolari.
2. **Crop box con Transform** (fallback): `BoundingBoxXYZ.Min`/`Max` sono in
   coordinate **locali**. Il bug classico è confrontarli con un punto mondo.
   La v3 faceva la cosa giusta: `Transform.Inverse.OfPoint()` e **poi**
   confronto con Min/Max.
3. **Scope box con crop disattiva**: `VIEWER_VOLUME_OF_INTEREST_CROP`.
4. Nessun limite.

Asimmetria notata: la tolleranza `XY_TOL` di 10 mm era applicata solo nei rami
2 e 3, non nel test poligonale del ramo 1, che è quello che si attiva più
spesso.

**Rotazione** - angolo fra `XYZ.BasisX` e `view.RightDirection` misurato
**attorno a** `view.ViewDirection`, normalizzato in `-pi..pi`. Usare
`ViewDirection` sia come asse di misura sia come asse di rotazione mantiene il
segno coerente fra Floor Plan e Ceiling Plan, che hanno `ViewDirection`
opposte.

**Indice dei tag esistenti** - set di chiavi `(link_id, room_id)` con
convenzione `-1` per l'host, letto da `TaggedRoomId`. Lookup O(1) invece di
riscansionare i tag per ogni room.

### Ordine dei filtri per room

Dal più economico al più costoso, con `continue` immediato:

1. sopra la vista: `z_min > top + Z_TOL`
2. sotto la vista: `z_max < low - Z_TOL`
3. cut plane (opzionale): richiede `z_min <= cut <= z_max`
4. dentro la crop (opzionale)
5. già taggata (opzionale)

I test 1 e 2 sono di **sovrapposizione fra intervalli**, non di contenimento:
basta che l'estensione della room intersechi il view range. Comportamento
corretto per replicare la visibilità di Revit.

La rotazione era differita a fine vista con un solo `doc.Regenerate()` per
vista invece di uno per tag.

---

## 3. Limiti individuati nella v3

Non bug, compromessi, ma da conoscere:

1. **Una room è ridotta a un punto.** Tutti i test XY usavano solo il punto di
   inserimento. Una room grande a cavallo del bordo della crop veniva scartata
   se il suo punto cadeva fuori, anche se metà era visibile.
2. **Le fasi non erano considerate.** Nessun filtro su `Phase`. Per le rooms
   dell'host è la causa principale degli errori di creazione: `NewRoomTag` su
   una room non visibile nella fase della vista fallisce.
3. **La visibilità effettiva non era verificata.** Categoria nascosta,
   `RevitLinkInstance` invisibile, view filter, elementi nascosti
   singolarmente, design option: tutto ignorato.
4. **Degradazione silenziosa sui sentinella del view range.** Se
   `PlanViewRange.Unlimited` non fosse accessibile, `SPECIAL_PLAN_IDS`
   resterebbe vuoto e un view range "Unlimited" verrebbe trattato come
   limitato, senza alcuna diagnostica.
5. **Lo span verticale era nominale, non geometrico.**

### Il limite strutturale

Lo script **reimplementava il motore di visibilità di Revit**. Ogni bug futuro
sarebbe stato una divergenza fra la sua approssimazione e il comportamento
reale.

---

## 4. L'approccio alternativo proposto

Principio: **delegare la visibilità a Revit dove possibile, e dove non è
possibile testare aree invece di punti.**

### Livello 0 - rooms dell'host: collector view-scoped

```python
FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Rooms)
```

Restituisce solo gli elementi effettivamente visibili in quella vista. Revit
applica da sé: crop region (rotata o non rettangolare, indifferente), view
range con tutti i casi speciali, filtro di fase, visibilità di categoria, view
filter, design option attiva, elementi nascosti, temporary hide/isolate.

Elimina circa 155 righe: `get_view_range_info`, `_resolve_plane_elevation`,
`_special_plan_level_ids`, `build_xy_test`, `_box_test_factory`,
`_polygon_from_crop_shape`, `_point_in_polygon`, `get_scope_box`,
`get_room_z_span`, `get_sorted_levels`.

**Ostacolo**: la categoria Rooms è quasi sempre spenta nelle viste dove si
taggano le rooms, e un collector view-scoped su categoria nascosta restituisce
zero elementi. Soluzione: riattivarla in una transazione di sola analisi poi
annullata. Gli `ElementId` raccolti sopravvivono al rollback, la modifica alla
vista no.

**Costo**: il primo collector view-scoped su una vista forza a Revit un calcolo
di visibilità completo. Una volta per vista, non per room.

### Livello 1 - rooms linkate: filtro a solido

Il collector view-scoped non funziona sui link: la vista appartiene all'host,
il collector dovrebbe girare sul documento linkato.

```python
crop_loops = view.GetCropRegionShapeManager().GetCropShape()   # coord. modello
z_low, z_top = resolve_view_extent(view)
base = flatten_loops_to_z(crop_loops, z_low)
view_solid = GeometryCreationUtilities.CreateExtrusionGeometry(base, XYZ.BasisZ, z_top - z_low)
view_solid_link = SolidUtils.CreateTransformed(view_solid, link_transform.Inverse)

rooms = FilteredElementCollector(link_doc)\
    .OfCategory(BuiltInCategory.OST_Rooms)\
    .WhereElementIsNotElementType()\
    .WherePasses(ElementIntersectsSolidFilter(view_solid_link))\
    .ToElements()
```

Risolve: sovrapposizione parziale, rooms concave, crop ruotate e non
rettangolari, estensione verticale reale, link ruotati o traslati.

**Prerequisiti**: volumi calcolati nel modello linkato (rilevabile con
`AreaVolumeSettings.GetAreaVolumeSettings(link_doc).ComputeVolumes`) e
`ElementIntersectsSolidFilter` è uno slow filter, da mitigare con un
`BoundingBoxIntersectsFilter` come broad-phase.

### Livello 2 - fallback: test ad area

`Room.GetBoundarySegments()` → poligono → test di sovrapposizione fra il
poligono della room e quello della crop, **in entrambi i versi** (un vertice di
A in B, oppure un vertice di B in A). Il doppio verso copre il caso "crop
piccola interamente dentro una room grande".

Limite residuo dichiarato: due rettangoli che si incrociano a croce senza
nessun vertice dell'uno dentro l'altro sfuggono a entrambi i test. Chiuso in
v4.0 aggiungendo il test di intersezione segmento-segmento.

### Separare identificazione e creazione

Considerato più importante di qualsiasi raffinamento geometrico:

```python
plan = []                       # FASE 1: nessuna transazione
show_preview(plan)              # FASE 2: l'utente decide
with revit.Transaction(...):    # FASE 3: transazione corta
    for item in plan: create_tag(item)
```

Il guadagno principale è il **dry-run**. Per uno strumento che può creare
migliaia di tag, poter vedere cosa succederà prima che succeda vale più di
qualunque precisione geometrica aggiuntiva. Prima, l'unico modo di scoprire che
i filtri erano sbagliati era creare i tag e fare Undo.

### Diagnostica azionabile

`Fuori crop/scope box: 37` non è azionabile: non sai quali, non sai se sono
esclusioni corrette o falsi negativi. Una tabella per vista con motivo
dell'esclusione e link cliccabili trasforma quel numero in qualcosa di
verificabile in trenta secondi.

### Cosa non cambiare

- pre-calcolo delle geometrie riusato su tutte le viste
- set di chiavi `(link_id, room_id)` per il dedup
- rotazione differita a fine vista con un solo `Regenerate()`

---

## 5. La riscrittura v4.0

Da 770 a 1670 righe. Struttura:

- **Dispatcher unico** `resolve_visible_rooms`: ogni strategia risponde solo a
  "quali rooms", il posizionamento del tag è logica condivisa a valle.
- **Verifica incrociata**: se la strategia primaria restituisce zero rooms, il
  dispatcher ricontrolla col modo geometrico. Se il geometrico ne trova, usa
  quello e stampa un avviso con la discrepanza. Il costo si paga solo nel caso
  a zero, ed è l'unico modo di distinguere "in questa vista non ci sono rooms"
  da "la strategia primaria non ha funzionato".
- **`polygons_overlap`** in tre passaggi: AABB, vertici nei due versi,
  intersezione segmento-segmento. Valvola su `MAX_SEGMENT_TESTS` che registra
  un avviso invece di tacere.
- **`pick_insertion_point`**: vincolo forte, il punto sta sempre dentro la
  room, altrimenti `NewRoomTag` non associa il tag. Se il punto di inserimento
  cade fuori dalla crop cerca in cascata centroide, midpoint dei lati, vertici
  tirati verso il centroide, vertici della crop tirati verso il centroide della
  room, e in ultima istanza una griglia 8×8 sull'AABB comune. Se nessun
  candidato sta in entrambi, crea il tag dentro la room e lo marca
  `in_crop=False`, contato e dichiarato in anteprima.
- **Transazione di probe** aperta solo quando serve a L0: se la sorgente è un
  link o il modo geometrico è forzato, l'analisi è lettura pura.

---

## 6. Scelta utente fra L1 e L2

Richiesta: quando L1 è disponibile, far scegliere all'utente se usarlo o
ricorrere a L2.

Il dialogo compare **solo quando la scelta è reale**:

```python
l1_available = (not source_is_host) and volumes_ok and opt_crop and not force_geo
```

Quattro precondizioni: sorgente linkata (per l'host vince comunque L0), volumi
calcolati, limite XY attivo (senza crop non esiste un volume finito da
estrudere), modo geometrico non già forzato.

Quando la scelta non è disponibile il motivo viene **dichiarato**: `L2
obbligato: volumi non calcolati nel modello linkato` oppure `L2 obbligato:
limite XY disattivato dalle opzioni`.

Il testo del dialogo espone il compromesso reale, non solo i nomi: L1 è esatto
ma costa una passata geometrica per vista, L2 è molto più rapido ma approssima
l'estensione verticale e ignora fasi, filtri e design option.

Innesto difensivo: `forms.alert(..., options=[...])` non è disponibile in tutte
le versioni di pyRevit. Se mancasse, lo script morirebbe su questo dialogo a
**ogni** esecuzione con modello linkato, che è il caso d'uso principale. La
chiamata è isolata in `ask_link_strategy()` con fallback su
`forms.CommandSwitchWindow.show`.

---

## 7. Debug: tutte le rooms escluse

### Primo caso - Modello A

Sintomo: zero tag creati.

```
SORGENTE ROOMS : Modello A (link architettonico)
ROOMS          : 111 totali, 92 posizionate
VOLUMI ROOMS   : calcolati
STRATEGIA LINK : scelta dall'utente -> Modo geometrico (L2)

| Vista | Strategia | Da creare | Escluse | Note |
| ...L1 | geometrico | 0 | 92 | fuori view range 92 |
```

**Fatti ricavati:** tutte le 92 rooms respinte dal test verticale, nessuna dal
test XY, che non era mai stato raggiunto. Nessuna sezione Avvisi, quindi i
sentinella di `PlanViewRange` si erano risolti e la crop shape era leggibile.
Le rooms erano visibili nella vista.

**Difetto della v4.0 emerso qui:** aveva perso la riga `View range ->
Top/Cut/Bottom/Depth` che la v3 stampava. Senza quei numeri l'esclusione non
era interpretabile. Regressione di diagnostica, corretta.

**Prima ipotesi, sbagliata:** la quota di base delle rooms. Il codice faceva
`base_level = room_doc.GetElement(room.LevelId) if room.LevelId else None`, con
due problemi: `room.LevelId` è un oggetto .NET quindi sempre truthy, e se
`GetElement` restituisce None `base_elev` diventa 0.0 in silenzio.

Interventi fatti in risposta, comunque utili e mantenuti:

- `get_room_base_elevation` con quattro tentativi in cascata: `Room.Level` →
  `LevelId` → `ROOM_LEVEL_ID` → **Z del punto di inserimento** (per una room
  quel punto sta sul piano del suo livello, quindi non può fallire).
- conteggio dei metodi di risoluzione stampato nel report
- diagnostica per vista con view range, estensione verticale complessiva delle
  rooms, e i primi 5 esclusi con quote e motivo specifico

**Risultato del giro successivo:** `quota di base risolta con Room.Level: 92
rooms`. **Ipotesi smentita dai dati.**

### Numeri del secondo report

```
TRASF. LINK : origine (-183.06, -58.78, -30.78) | traslazione Z = -30.78 piedi
View range  -> Top: 31.62 | Cut: 25.81 | Bottom: 21.88 | Depth: 21.88
Rooms       -> estensione verticale da -54.40 a -3.19 (coordinate host)
```

| grandezza | piedi | metri |
|---|---:|---:|
| View range Bottom (= livello) | 21.88 | +6.67 |
| View range Cut | 25.81 | +7.87 (livello +1.20) |
| View range Top | 31.62 | +9.64 (livello +2.97) |
| Rooms, estensione totale | -54.40 → -3.19 | -16.58 → -0.97 |

Il view range era **corretto**: cut plane a 1.20 m dal livello e Top a 2.97 m
sono valori canonici. Anche i livelli del link erano coerenti: le basi
ricostruite (-54.40, -44.56, -34.29) spaziate di 9.84 ft = esattamente 3.00 m.

Due sistemi di quote entrambi sensati, separati da 76.28 ft = 23.25 m. Non un
errore di calcolo. Interventi aggiunti per discriminare:

- **estensione verticale dal bounding box reale**
  (`get_room_z_span_host_from_bbox`), fonte primaria, con la ricostruzione dai
  parametri come fallback e **entrambi i valori stampati affiancati**
- controllo sulle istanze multiple dello stesso link, con la loro traslazione Z
- confronto fra le fonti per ogni room esclusa: quota del livello nel link, Z
  del punto di inserimento nel link e in host, span da bbox e da parametri

Su questo modello il problema si è risolto: la sorgente selezionata non
corrispondeva alle rooms visibili nella vista.

---

## 8. La causa reale: Elevation vs ProjectElevation

### Secondo caso - Modello B

```
SORGENTE ROOMS : Modello B (Areas and Rooms, linkato)
ROOMS          : 184 totali, 184 posizionate
VOLUMI ROOMS   : NON calcolati
STRATEGIA LINK : L2 obbligato: volumi non calcolati nel modello linkato
TRASF. LINK    : origine (0.00, 0.00, 0.00) | traslazione Z = 0.00 piedi
ISTANZA USATA  : id 2888798 | istanze dello stesso documento: 1

View range -> Top: 542.45 | Cut: 538.84 | Bottom: 534.90 | Depth: 534.90
Rooms      -> estensione verticale da -8.41 a 39.86 (coordinate host)

[escluso] room 1992039: z da 17.84 a 27.68 (bounding box)
    bbox host 17.84 .. 27.68 | parametri host 561.15 .. 570.99
    | livello link 561.15 | punto ins. link 17.84 -> host 17.84
```

La trasformazione del link era l'**identità**, quindi nessuna trasformazione in
gioco. Eppure per la stessa room:

- `Level.Elevation` → il livello è a **561.15** ft
- la geometria reale (punto di inserimento e bounding box) → **17.84** ft

**Differenza: 543.31 ft = 165.60 m.** Non una quota di edificio: l'**altitudine
del sito sul livello del mare**.

Conferma numerica incrociata: `534.90 − 543.31 = −8.41`, ed è esattamente il
minimo riportato in `Rooms -> estensione verticale da -8.41`.

### La causa

`Level.Elevation` restituisce la quota riferita al **Survey Point** quando il
parametro *Elevation Base* del livello è impostato così. La geometria invece è
**sempre** in coordinate interne: bounding box, punti di inserimento,
`Transform` dei link.

Lo script confrontava il view range in scala survey (~535 ft) con l'estensione
delle rooms in scala interna (~18 ft). Nessuna room poteva sovrapporsi: da qui
l'esclusione di **tutte**, 184 su 184. Sul Modello A funzionava perché là i due
sistemi coincidevano.

Le 5 rooms campionate stavano su un piano superiore e da sole non dicevano
niente: è stato il confronto fra `livello link` e `punto ins. link` sulla stessa
riga a rendere visibile lo scarto. Il bounding box, introdotto al giro
precedente per un motivo diverso, si è rivelato decisivo proprio perché essendo
geometria pura era già in coordinate interne.

### La correzione (v4.1)

```python
def level_internal_elevation(level):
    """Quota del livello nel sistema di coordinate INTERNO del suo documento."""
    if level is None:
        return None
    try:
        return level.ProjectElevation
    except Exception:
        pass
    try:
        return level.Elevation
    except Exception:
        return None
```

Applicata in cinque punti: ordinamento dei livelli, risoluzione dei quattro
piani del view range, sentinella Level Above / Level Below, quota di base della
room, Upper Limit della room.

**`check_level_geometry_consistency`** confronta, su un campione di rooms, la
quota interna del livello con la Z del punto di inserimento. Se lo scarto
supera un piede emette un avviso con il valore in metri. Questo controllo, da
solo, avrebbe intercettato il problema alla prima esecuzione invece di lasciarlo
comparire come "zero tag creati".

**Diagnostica del livello di vista**: `Elevation`, `ProjectElevation` e il
valore usato, stampati affiancati, per rendere verificabile la correzione
stessa.

Esito: **funzionante**.

---

## 9. v4.1: fascia verticale

Richiesta: considerare solo i primi 2 metri sopra il livello della vista, per
non intercettare rooms del piano superiore.

```python
def get_vertical_limits(view_range, use_band):
    low = view_range["low"]
    top = view_range["top"]
    if not use_band:
        return low, top
    base = view_range["base"]
    band_low = base
    band_top = base + LEVEL_BAND_FT
    if low is not None and low > band_low:
        band_low = low
    if top is not None and top < band_top:
        band_top = top
    if band_top < band_low:
        band_top = band_low
    return band_low, band_top
```

L'intersezione va in una sola direzione: la fascia può solo **restringere** il
view range, mai allargarlo. Se una vista ha il Top a 1.8 m resta 1.8 m; se ce
l'ha illimitato diventa 2 m. È proprio il caso illimitato quello che cattura il
piano superiore, perché lì il test verticale non ha alcun limite alto.

### Applicata a tutte e tre le strategie

| strategia | come |
|---|---|
| **L2 geometrico** | i limiti passano da `get_vertical_limits`, test con `z_overlaps` |
| **L1 filtro a solido** | la fascia entra in `view_solid_z_extent`: il solido estruso è già alto 2 m, quindi il filtro lavora su meno volume |
| **L0 visibilità Revit** | la selezione la fa Revit col suo view range e non gli si può passare un limite più restrittivo, quindi `filter_rooms_by_band` come post-filtro |

### Attribuzione delle esclusioni

Se una room sarebbe passata col view range grezzo ma non con la fascia,
l'esclusione viene attribuita alla fascia. Nella colonna Note si legge quindi
`fuori view range 12 | fuori fascia 47`, e si sa quante rooms ha tolto la nuova
regola. Più una riga per vista con i limiti effettivi:

```
Fascia -> da -8.41 a -1.85 (primi 2.0 m sopra il livello, intersecati col view range)
```

### Semantica

**Sovrapposizione, non appartenenza.** Una room è inclusa se il suo intervallo
verticale interseca la fascia, anche parzialmente.

- Conseguenza voluta: una room del piano di sotto a doppia altezza che sale
  nella fascia viene inclusa, ed è corretto perché in quella vista si vede.
- Conseguenza da tenere presente: una room del piano superiore con un Base
  Offset negativo che la fa scendere sotto i 2 m verrebbe ancora inclusa.

La regola più stretta ("la base della room deve stare dentro la fascia") è una
modifica di due righe, ma escluderebbe anche i volumi a doppia altezza
legittimi.

---

## 10. Punti aperti e non verificati

### Non verificati sul campo

1. **`ProjectElevation` è davvero la quota interna?** I dati provano che
   `Level.Elevation` non lo è (561.15 contro 17.84 di geometria). Non provano
   che `ProjectElevation` sia quella giusta: quella parte è un'inferenza sulla
   semantica dell'API, non una misura. Coincide con la quota interna quando il
   Project Base Point è a quota zero, che è il caso normale ma non garantito.
   Se un giorno `ProjectElevation` risultasse anch'esso in scala survey, la
   strada è leggere l'offset da `DB.BasePoint` (`Position` contro
   `SharedPosition`). Il controllo di coerenza lo segnalerebbe comunque.
2. **Collector view-scoped su `OST_Rooms`** (L0): pattern noto, ma è il perno
   della strategia per le rooms dell'host e non è stato testato in isolamento.
   La verifica incrociata automatica lo segnalerebbe.
3. **`SetCategoryHidden` dentro la transazione di probe** seguito da
   `Regenerate`: è sufficiente perché il collector veda le rooms nella stessa
   transazione?
4. **`ElementIntersectsSolidFilter`** sulle rooms di un documento linkato con i
   volumi attivi. Il controllo su `ComputeVolumes` intercetta il caso senza
   volumi, ma il percorso con volumi non è stato esercitato.

### Comportamenti noti e accettati

- Rooms **Redundant** producono due tag sovrapposti: entrambe hanno `Area > 0`
  e passano il filtro di piazzamento.
- Con `salta duplicati` disattivato, rieseguire lo script sulla stessa vista
  produce tag sovrapposti.
- Il test segmento-segmento in `polygons_overlap` viene saltato oltre
  `MAX_SEGMENT_TESTS`, con avviso esplicito.
- Le rooms linkate non sono selezionabili dall'host, quindi nel dettaglio
  dell'anteprima non hanno link cliccabile: viene stampato l'ID nudo.

### Non toccato

- `bundle.yaml` risultava già modificato in working tree all'inizio della
  sessione e non è stato allineato.

---

## 11. Riferimento rapido

### Costanti configurabili

| costante | valore | significato |
|---|---|---|
| `XY_TOL` | 10 mm in piedi | tolleranza sul test XY |
| `Z_TOL` | 10 mm in piedi | tolleranza sul test Z |
| `MIN_SEG` | 0.01 ft (~3 mm) | sopra la short curve tolerance di Revit |
| `DEFAULT_ROOM_HEIGHT` | 8.0 ft | fallback se l'estensione verticale non è calcolabile |
| `HUGE_Z` | 3280 ft (~1 km) | sostituto finito di un view range illimitato |
| `MAX_SEGMENT_TESTS` | 40000 | valvola sul test segmento-segmento |
| `PREVIEW_ROW_LIMIT` | 250 | righe di dettaglio in anteprima |
| `LEVEL_BAND_METERS` | 2.0 | altezza della fascia sopra il livello della vista |

### Opzioni

| opzione | default | effetto |
|---|:---:|---|
| Limita alle rooms dentro la crop region / scope box | on | attiva il test XY |
| Allinea i tag alla rotazione della vista | on | rotazione differita a fine vista |
| Salta le rooms che hanno già un tag nella vista | on | dedup su `(link_id, room_id)` |
| Crea i tag senza leader | on | `HasLeader = False` |
| Richiedi che il cut plane attraversi la room | off | solo modo geometrico, più restrittivo |
| Riposiziona i tag nella porzione visibile | on | `pick_insertion_point` cerca un punto in room ∩ crop |
| Considera solo i primi 2.0 m sopra il livello | on | fascia verticale |
| Forza il modo geometrico | off | diagnostica e confronto fra strategie |
| Solo anteprima | off | dry-run, nessuna modifica |

### Strategie

| livello | nome | quando | fonte della verità |
|---|---|---|---|
| **L0** | visibilità Revit | sorgente = host | collector view-scoped |
| **L1** | filtro a solido | sorgente = link, volumi calcolati, crop attiva | `ElementIntersectsSolidFilter` |
| **L2** | geometrico | fallback, o scelta dell'utente | aritmetica su poligoni e intervalli |

### Come confrontare due strategie sullo stesso modello

Eseguire due volte sulle stesse viste con `Solo anteprima` attivo, la prima con
`Forza il modo geometrico` e la seconda senza. Le due anteprime sono
direttamente confrontabili e la differenza nei conteggi quantifica quanto le due
strategie divergono, senza toccare il modello.
