# Openings (forometrie)

Fori di predisposizione ("provision for void") dove tubi, canali, passerelle e conduit attraversano
muri, solai e strutture, nel modello host e nei link RVT e IFC.

Ordine d'uso per progetto: **Links** → **Setup** → **Place** → **Refresh** (ripetuto a ogni revisione).

## Struttura

| Pulsante | Script | Cosa fa |
| --- | --- | --- |
| Links | `Links.pushbutton/script.py` | Inventario dei tipi attraversabili (host + link caricati), spunta e REI forzato per tipo |
| Setup | `Setup.pushbutton/script.py` | Configurazione del progetto (finestra `Setup_form.xaml` / `Setup_ui.py`) |
| Place | `Place.pushbutton/script.py` | Crea i fori nuovi, con anteprima |
| Refresh | `Refresh.pushbutton/script.py` | Confronto con lo stato attuale e applicazione selettiva |

Convenzione del repo: nessun modulo condiviso.

- **Blocco di configurazione:** il blocco tra `Configurazione di progetto` e `fine del blocco comune` è
  **identico nei quattro script**. Una correzione va riportata in tutti e quattro.
- **Place e Refresh:** sono identici dalla riga `corpo comune` in giù, tranne `MODE`.

## Dati salvati nel modello (Extensible Storage)

- **Configurazione:** schema `ESA_OpeningsConfig` (GUID `5d0a2f6e-…`) su Informazioni progetto, un campo
  `Json`. Contiene famiglie, parametri, regole REI, regole di dimensionamento, inventario dei tipi e revisione
  (`rev`, alzata a ogni salvataggio).
- **Identità di ogni foro:** schema `ESA_OpeningData` (GUID `a3e7c9d2-…`) sull'istanza, con:
  - modello, istanza di link ed elemento attraversato;
  - UniqueId degli elementi MEP;
  - famiglia, centro, dimensioni, valori dei parametri, stato e date.

  Non si aggiunge nessun parametro al progetto.
- **Riferimento all'elemento attraversato:** `IfcGUID` se presente (resta stabile anche ricaricando un
  IFC), altrimenti `UniqueId`.
- **Inventario dei tipi:** chiave `modello|categoria|Famiglia: Tipo`, cioè per nome e non per Id, che negli
  IFC ricollegati può cambiare.

## Rilevazione

1. Elementi MEP rettilinei (categorie scelte in Setup) nell'ambito: selezione, vista o modello.
2. Sezione esterna con isolamento:
   - tonda: diametro esterno;
   - rettangolare: letta dal solido, quindi vale anche per canali ruotati e passerelle di costa.

   La sezione viene estrusa lungo l'asse (`mep_prism`), così si ottiene un solido pulito.
3. Elementi attraversabili dei tipi spuntati: solidi portati in coordinate host con
   `SolidUtils.CreateTransformed` e messi in cache. Filtro rapido con il bounding box, poi
   `BooleanOperationsUtils` (Intersect). Se l'operazione booleana fallisce, il ripiego è l'asse dentro il solido
   allargato alla sezione; il report lo indica come "approximate geometry".
4. **Orientamento dato dall'elemento MEP:**
   - inclinazione oltre 45° → foro "floor" (verticale): solai, coperture, fondazioni, travi;
   - altrimenti → foro "wall": muri, pilastri, travi.

   Altre combinazioni sono escluse (montanti dentro i muri, tubi nel massetto). Così come una profondità
   oltre 3 m.
5. **Terna del foro:**
   - wall: u lungo l'elemento (tangente della curva di posizione, o perpendicolare al tubo), v verticale,
     w attraverso;
   - floor: u dalla sezione rettangolare, oppure dalla fila di montanti, oppure X; w verticale.
6. **Classe:**
   - Structural: categorie `OST_Structural*`, per nome interno, indipendente dalla lingua;
   - altrimenti REI: forzato in Links oppure regole di Setup;
   - altrimenti Normal.
7. **Accorpamento:** solo sullo stesso elemento attraversato, rettangoli nel piano (u, v) a distanza minore o
   uguale a "Merge within" della classe.
8. **Dimensioni:**
   - larghezza e altezza: ingombro più il franco per lato, arrotondati per eccesso al passo;
   - diametro: lato maggiore più il franco, arrotondato;
   - profondità: spessore più la sporgenza.

## Inserimento

- **Inserimento:** `NewFamilyInstance(punto, simbolo, livello, NonStructural)` al centro, poi rotazione
  attorno a Z. Servono famiglie Generic Model **non ospitate** con origine al centro del foro.
  - Wall: X larghezza, Y profondità, Z altezza.
  - Floor: X e Y in pianta, Z profondità.
- **Livello:**
  - wall: quello sotto il centro;
  - floor: il più vicino. Si usa `ProjectElevation`, vedi la nota sulle coordinate in CLAUDE.md.
- **Dimensioni:** se la famiglia ha il parametro di diametro (di istanza), si scrivono diametro e profondità;
  altrimenti larghezza, altezza e profondità. I parametri non trovati come parametri di istanza vengono
  saltati, e se non ne viene scritto nessuno il foro è solo inserito e orientato. Il report lo dice.
- **Workset** del Setup, se il modello è condiviso e il workset esiste (Setup → crea).

## Refresh

- **Abbinamento:** un foro calcolato corrisponde a quello esistente sullo stesso elemento attraversato con
  più elementi MEP in comune.
- **Differenze rilevate:**
  - famiglia o classe;
  - posizione (5 mm) e rotazione (0,5°);
  - dimensioni (1 mm);
  - elementi MEP;
  - parametri;
  - foro di nuovo necessario dopo essere stato marcato obsoleto.
- **Non più necessari:** fori senza corrispondenza. Nell'ambito vista contano solo quelli che riguardano
  elementi MEP della vista, oppure i cui elementi MEP non esistono più.
- **Applicazione:**
  - TransactionGroup, una transazione per foro con `RevitErrors`;
  - cambio famiglia con `ChangeTypeId`; se non è possibile, sostituzione del foro;
  - obsoleti marcati (parametro Status + dati interni) oppure cancellati, secondo il Setup.

## Setup → workset, abaco, filtri

- **Workset:** creato se il modello è condiviso.
- **Abaco Generic Models:** contiene Famiglia e tipo, i parametri del Setup trovati come campi schedulabili
  (solo parametri condivisi o di progetto) e Conteggio. È filtrato su Status con "ha un valore" quando Status
  è tra i campi.
- **Filtri di vista** "Openings - New / Modified / Obsolete" su Status: servono un parametro condiviso o di
  progetto. Vengono creati ma non applicati alle viste.

## Da verificare in Revit (non provato fuori da Revit)

- Operazioni booleane su DirectShape IFC e su geometrie complesse: frequenza del ripiego "approximate".
- Inserimento e quota delle famiglie level based con punto fuori dal livello: lo script sposta comunque al
  centro dopo l'inserimento.
- `ChangeTypeId` tra famiglie diverse della stessa categoria.
- `ScheduleFilterType.HasValue` e `CreateEqualsRule` nelle diverse versioni (2022-2026: c'è il ripiego con
  il parametro di maiuscole/minuscole).
- Prestazioni su modelli grandi: confronto dei bounding box in Python, migliaia di elementi.
