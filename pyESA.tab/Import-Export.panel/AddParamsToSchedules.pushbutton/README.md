# Aggiungi Parametri alle Schedule (pyRevit pushbutton)

Aggiunge uno o più parametri di progetto/condivisi come **campi** in una o più
schedule. I parametri non disponibili per la categoria della singola schedule
vengono **saltati** senza interrompere l'elaborazione e riportati nel **report
finale** con il motivo dello scarto.

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
2. **Selezione parametri** – vengono letti i parametri presenti in
   `doc.ParameterBindings`, quindi **parametri di progetto e condivisi**, con
   indicazione di tipo (Progetto/Condiviso), binding (Istanza/Tipo), tipo dato e
   numero di categorie associate. Entrambe le liste hanno casella di ricerca
   multi-token e selezione/deselezione rapida.
3. **Opzione campi nascosti** – i campi vengono aggiunti con `IsHidden = True`,
   utile quando servono solo per filtri e ordinamenti.
4. **Esecuzione** – per ogni schedula lo script legge i
   `Definition.GetSchedulableFields()` e cerca il parametro **prima per
   `ParameterId`, poi per nome**. Questo è il controllo autoritativo di
   disponibilità: se il parametro non compare tra i campi schedulabili, viene
   saltato.
5. **Report** – tabella con Schedule / Parametro / Esito / Dettaglio, filtro
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
- **Databinding**: le checkbox usano binding `OneWay` più eventi
  `Checked`/`Unchecked`, così lo stato non dipende da `INotifyPropertyChanged`
  (non disponibile su oggetti Python puri).
- **Icone**: 96x96 PNG con canale alpha. `icon.dark.png` viene usato
  automaticamente da pyRevit quando Revit è in tema scuro.

## Limiti noti

- I parametri **built-in** (es. Volume, Livello) non sono elencati, perché la
  sorgente scelta sono i parametri di progetto/condivisi. Per includerli
  basterebbe popolare la lista anche dai `GetSchedulableFields()` delle schedule
  selezionate.
- Le *panel schedule* elettriche (`PanelScheduleView`) non sono `ViewSchedule` e
  non compaiono nell'elenco.
- Lo script non modifica ordine colonne, formattazione, filtri o ordinamenti
  esistenti: aggiunge i campi in coda.
