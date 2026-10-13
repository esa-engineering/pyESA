# Spring (cartella Stub)

## Logica di funzionamento

Il comando **Stub** crea uno stub (tubo/cassetta corto) da verticale (riser) collegandolo con una giunzione appropriata. La logica è strutturata in tre fasi principali:

### 1. Selezione verticale
- Il sistema legge l'elenco di verticali selezionati (riser)
- Verifica che sia un elemento valido (Pipe, Duct o CableTray)
- Calcola l'asse e la lunghezza totale della verticale

### 2. Creazione dello stub
- L'utente specifica:
  - **Livello**: livello di posizionamento dello stub
  - **Lunghezza**: dello stub (solitamente 500mm - 2000mm)
  - **Posizione**: sopra/sotto il punto di connessione
- Lo stub è creato con:
  - **Tipo** e **Dimensioni** ereditati dal riser
  - **Isolamento** se presente sul riser
- La giunzione dipende dal tipo di elemento:
  - Pipe → Tee (default)
  - Duct → Tap (default)
  - CableTray → Tee (obbligatorio)

### 3. Applicazione parametri
- Lo stub mantiene:
  - **Type Mark** del riser
  - **Size** del riser
  - **Isolamento** se presente
- Se la giunzione non entra, lo stub non viene creato
- Viene generato un report di errore

---

## Programmazione (schema a blocchi)

```
┌─────────────────────────────────────────┐
│  Inizio comando Stub                     │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  1. Selezione verticale (riser)          │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  2. Calcolo asse e lunghezza verticale    │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  3. Apertura dialogo configurazione      │
│  - Livello, Lunghezza, Posizione         │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  4. Verifica e creazione giunzione       │
│  - Tipo giunzione (Tee/Tap)               │
│  - Prova alternativa se fallisce      │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  5. Creazione elemento stub              │
│  - Tipo e dimensioni dal riser            │
│  - Isolamento se presente               │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  6. Generazione report di errore/fallito │
└─────────────────────────────────────────┘
```

---

## Commento al codice (RevitAPI/PyRevit)

### Importazioni principali
```python
# Importazioni RevitAPI per tipi MEP
from Autodesk.Revit.DB.Plumbing import Pipe, PipeInsulation
from Autodesk.Revit.DB.Mechanical import Duct, DuctInsulation
from Autodesk.Revit.DB.Electrical import CableTray

# Importazioni PyRevit per utilities
from pyrevit import revit, script, DB, forms

# Importazioni .NET
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
```

### Selezione verticale
```python
# Filtra gli elementi selezionati per tipo
selection = uidoc.Selection.PickObjects(ObjectType.Element)
verticals = []
for elem_id in selection:
    elem = doc.GetElement(elem_id)
    if isinstance(elem, Pipe) or isinstance(elem, Duct) or isinstance(elem, CableTray):
        verticals.append(elem)
```

### Calcolo asse verticale
```python
# Calcolo dell'asse per linee verticali
curve = elem.Location.Curve
p1, p2 = curve.GetEndPoint(0), curve.GetEndPoint(1)
verticality = abs(p1.X - p2.X) + abs(p1.Y - p2.Y) < tolerance
```

### Creazione stub
```python
# Crea stub con tipo e dimensioni dal riser
riser_symbol = riser.Symbol
riser_size = riser.Symbol.GetShape()

# Crea stub
stub = MechanicalUtils.CreateStub(
    doc=doc,
    riser=riser,
    level=level,
    length=length,
    symbol=riser_symbol,
    size=riser_size
)
```

### Copia parametri ereditati
```python
# Copia Type Mark e Size dal riser
new_stub.Symbol = riser.Symbol
new_stub.GetParameter("Type Mark").Set(riser.GetParameter("Type Mark").Value)

# Copia isolamento se presente
if isinstance(riser, Pipe) and riser.GetParameter("PipeInsulation"):
    new_stub.PipeInsulation = riser.PipeInsulation
```

### Generazione report
```python
# Report dettagliato per ogni errore
report_lines = []
for riser in risers:
    if stub_failed:
        report_lines.append(
            f"Stub per {riser.Name}: {reason} - Stub non creato"
        )
output.print_md("\n".join(report_lines))
```

---

## Parametri configurabili

| Parametro | Descrizione | Valori possibili |
|------------|-------------|------------------|
| **Level** | Livello di posizionamento dello stub | Qualsiasi livello del modello |
| **Length** | Lunghezza dello stub | 0.01 - 2000+ m (valore reale) |
| **Position** | Posizione sopra/sotto il punto di connessione | Above, Below |
| **Junction Type** | Tipo di giunzione | Tee, Tap (dipende dal tipo elemento) |

---

## Limitazioni e note

- **Pipe**: usa Tee come default, Tap come alternativa
- **Duct**: usa Tap come default, Tee come alternativa
- **CableTray**: usa Tee (obbligatorio)
- Se la giunzione non entra, lo stub non viene creato
- Lo stub mantiene i parametri del riser (Type Mark, Size, Isolamento)

---

## Autori

- Matteo Maiorano
