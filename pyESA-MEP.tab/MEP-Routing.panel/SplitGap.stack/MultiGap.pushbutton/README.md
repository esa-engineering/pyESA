# SplitGap

## Logica di funzionamento

Il comando **SplitGap** crea una giunzione (tee, tap o curve) su una verticale per eliminare uno spazio vuoto o un gap. La logica è strutturata in tre fasi principali:

### 1. Selezione verticale e gap
- Il sistema legge l'elenco di verticali selezionati (riser)
- Verifica che sia un elemento valido (Pipe, Duct o CableTray)
- Calcola l'asse e la lunghezza totale della verticale

### 2. Identificazione gap
- Il sistema calcola la posizione del gap:
  - Gap = punto di connessione desiderato - punto di connessione esistente
- Il gap può essere:
  - **Negativo**: spazio vuoto da aggiungere
  - **Positivo**: spazio in eccesso da rimuovere
- Viene calcolata la distanza dal punto di connessione all'asse

### 3. Creazione giunzione
- L'utente specifica:
  - **Tipo di giunzione**: Tee, Tap o Curve
  - **Posizione**: sopra/sotto il punto di connessione
  - **Livello**: livello di posizionamento
- La giunzione dipende dal tipo di elemento:
  - Pipe → Tee (default) o Curve
  - Duct → Tap (default) o Curve
  - CableTray → Tee (obbligatorio) o Curve
- Se la giunzione non entra, la verticale resta invariata
- Viene generato un report di errore

---

## Programmazione (schema a blocchi)

```
┌─────────────────────────────────────────┐
│  Inizio comando SplitGap                  │
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
│  3. Calcolo posizione gap                │
│  - Posizione desiderata - Posizione attua │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  4. Apertura dialogo configurazione      │
│  - Tipo giunzione, Posizione, Livello     │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  5. Verifica e creazione giunzione       │
│  - Tipo giunzione (Tee/Tap/Curve)        │
│  - Prova alternativa se fallisce        │
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

### Calcolo posizione gap
```python
# Calcola la posizione del gap
desired_point = level.GetPointXYZ(x, y)
current_point = riser.GetConnectionPoint()
gap = desired_point.Distance(current_point)

# Verifica se il gap è compatibile con la giunzione
if gap < min_gap_size:
    # Gap troppo piccolo per la giunzione
    raise GapTooSmallException
```

### Creazione giunzione
```python
# Dipende dal tipo di elemento
if isinstance(elem, Pipe):
    # Pipe → Tee (default) o Curve
    if junction_type == "Tee":
        fitting = MechanicalUtils.CreateTee(doc, riser, desired_point)
    else:
        fitting = MechanicalUtils.CreateCurve(doc, riser, desired_point)
elif isinstance(elem, Duct):
    # Duct → Tap (default) o Curve
    if junction_type == "Tap":
        fitting = MechanicalUtils.CreateTap(doc, riser, desired_point)
    else:
        fitting = MechanicalUtils.CreateCurve(doc, riser, desired_point)
elif isinstance(elem, CableTray):
    # CableTray → Tee (obbligatorio) o Curve
    fitting = MechanicalUtils.CreateTee(doc, riser, desired_point)
```

### Generazione report
```python
# Report dettagliato per ogni errore
report_lines = []
for riser in risers:
    if gap_too_small:
        report_lines.append(
            f"Gap per {riser.Name}: {gap}mm - Troppo piccolo per giunzione"
        )
output.print_md("\n".join(report_lines))
```

---

## Parametri configurabili

| Parametro | Descrizione | Valori possibili |
|------------|-------------|------------------|
| **Junction Type** | Tipo di giunzione | Tee, Tap, Curve |
| **Position** | Posizione sopra/sotto il punto di connessione | Above, Below |
| **Level** | Livello di posizionamento | Qualsiasi livello del modello |
| **Gap Distance** | Distanza del gap | Valore reale (mm) |

---

## Limitazioni e note

- **Pipe**: usa Tee come default, Curve come alternativa
- **Duct**: usa Tap come default, Curve come alternativa
- **CableTray**: usa Tee (obbligatorio), Curve come alternativa
- Se il gap è troppo piccolo per la giunzione, la verticale resta invariata
- Si possono creare curve 3D per ottenere giunzioni non standard

---

## Autori

- Matteo Maiorano
