# Aggiungi Parametri alle Schedule (pyRevit pushbutton)

Aggiunge uno o più parametri come **campi** in una o più schedule. Sono
selezionabili sia i **parametri di progetto e condivisi** sia i **parametri
nativi di Revit**. I parametri non disponibili per la categoria della singola
schedule vengono **saltati** senza interrompere l'elaborazione e riportati nel
**report finale** con il motivo dello scarto.

## Contenuto del bundle

```
AddParamsToSchedules.pushbutton/
├── script.py             # logica Revit API (IronPython 2.7 / CPython3)
├── AddParamsWindow.xaml  # finestra di selezione schedule + parametri
├── ReportWindow.xaml     # finestra di report esiti/errori
├── bundle.yaml           # metadata del pulsante (titolo, tooltip, autore)
├── icon.png              # icona tema chiaro (96x96)
└── icon.dark.png         # icona tema scuro (96x96)
```

## Installazione

Copia la cartella `AddParamsToSchedules.pushbutton` dentro un panel della tua
extension pyRevit, ad esempio:

```
MiaExtension.extension/
└── ESA.tab/
    └── Schedule.panel/
        └── AddParamsToSchedules.pushbutton/
```

Poi `pyRevit > Reload` (o riavvia Revit). Se la extension non è ancora
registrata: `pyRevit > Settings > Custom Extension Directories`.

## Come funziona

1. **Selezione schedule** – vengono elencate tutte le `ViewSchedule` del
   progetto escluse le view template e le revision schedule dei cartigli.
   Per ogni schedule sono mostrati categoria, tipo (key schedule / material
   takeoff) e numero di campi attuali.
2. **Indice dei campi schedulabili** – all'avvio lo script scorre le schedule e
   costruisce l'unione dei loro `Definition.GetSchedulableFields()`, con barra di
   progresso annullabile. Questo indice ha due usi: è la sorgente dei parametri
   nativi ed è ciò che permette di dire in quante schedule ogni parametro è
   realmente utilizzabile.
3. **Selezione parametri** – la lista unisce due sorgenti:
   - **progetto e condivisi**, letti da `doc.ParameterBindings`, con binding
     (Istanza/Tipo), tipo dato e numero di categorie associate;
   - **nativi di Revit**, presi dall'indice del punto 2, con il tipo di campo
     (`ScheduleFieldType`: Istanza, Tipo, Conteggio, Da vista, ...).

   La deduplica avviene per `ParameterId` e, in fallback, per nome: un parametro
   di progetto che compare anche tra i campi schedulabili resta una sola voce.
   Filtri disponibili: sorgente (Tutti / Progetto e condivisi / Nativi Revit),
   ricerca multi-token e **"solo parametri disponibili in tutte le schedule
   selezionate"**, che azzera il rischio di scarti prima di eseguire.
4. **Opzione campi nascosti** – i campi vengono aggiunti con `IsHidden = True`,
   utile quando servono solo per filtri e ordinamenti.
5. **Esecuzione** – per ogni schedula lo script rilegge i
   `GetSchedulableFields()` e cerca il parametro **prima per `ParameterId`, poi
   per nome**. Questo è il controllo autoritativo di disponibilità: se il
   parametro non compare tra i campi schedulabili, viene saltato.
6. **Report** – tabella con Schedule / Parametro / Esito / Dettaglio, filtro
   "solo scarti ed errori", copia negli appunti ed esportazione CSV.

### Esiti possibili

| Esito | Significato |
|---|---|
| `Aggiunto` | campo inserito nella schedule |
| `Già presente` | il campo esisteva già (nessuna modifica) |
| `Non disponibile` | parametro non schedulabile per la categoria della schedule |
| `Errore` | eccezione dell'API Revit durante l'aggiunta |

## Gestione delle transazioni

- Un `TransactionGroup` racchiude l'intera operazione.
- Una `Transaction` **per singola schedule**: se non viene aggiunto nulla la
  transazione è annullata (nessuna voce inutile nell'undo di Revit); se una
  schedule solleva un'eccezione non gestita viene annullata solo quella, le
  altre restano valide.
- Un singolo `AddField` fallito non blocca gli altri parametri della stessa
  schedule.

## Note tecniche

- **Engine**: scritto per IronPython 2.7 (engine di default pyRevit); la sintassi
  è compatibile anche con CPython3.
- **Compatibilità API**: `ElementId.Value` / `IntegerValue` e
  `Definition.GetDataType()` / `ParameterType` sono gestiti con fallback, quindi
  funziona da Revit 2019 alle versioni recenti.
- **Databinding**: `IsChecked` non è in binding. Lo stato viene ripristinato
  dall'evento `Loaded` della CheckBox e scritto dagli eventi
  `Checked`/`Unchecked`, così non dipende da `INotifyPropertyChanged` (non
  disponibile su oggetti Python puri) né dall'engine in uso. Le liste sono
  virtualizzate in modalità `Standard` (container non riciclati), condizione
  necessaria perché `Loaded` rientri a ogni realizzazione di riga.
- **Dimensione delle liste**: con i nativi inclusi l'elenco parametri può
  contare migliaia di voci, perché una schedule multi-categoria espone i campi
  schedulabili di tutte le categorie. Da qui la virtualizzazione e i filtri.
- **Icone**: 96x96 PNG con canale alpha. `icon.dark.png` viene usato
  automaticamente da pyRevit quando Revit è in tema scuro.

## Limiti noti

- L'elenco dei nativi è l'unione sui campi schedulabili delle schedule
  **presenti nel progetto**: un parametro nativo di una categoria che non ha
  nessuna schedule non compare. È una conseguenza voluta della sorgente scelta,
  che in cambio garantisce che ogni voce sia realmente aggiungibile in almeno
  una schedule.
- La scansione iniziale costa una chiamata `GetSchedulableFields()` per
  schedule. Su progetti con centinaia di schedule richiede qualche secondo; la
  barra di progresso è annullabile e, se interrotta, l'elenco dei nativi è
  parziale (la finestra lo segnala).
- Le *panel schedule* elettriche (`PanelScheduleView`) non sono `ViewSchedule` e
  non compaiono nell'elenco.
- Lo script non modifica ordine colonne, formattazione, filtri o ordinamenti
  esistenti: aggiunge i campi in coda.
