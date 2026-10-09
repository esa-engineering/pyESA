# CreateSimilar

## Logica di funzionamento

Il comando **CreateSimilar** crea nuovi elementi MEP (Pipe, Duct o CableTray) simili a quelli esistenti, mantenendo:

### 1. Selezione elemento base
- L'utente seleziona un elemento di base (riser)
- Il sistema identifica:
  - **Tipo**: Pipe, Duct o CableTray
  - **Element ID**: identificativo univoco
  - **Livello**: a cui appartiene
  - **Dimensioni**: Size e Type Mark

### 2. Configurazione nuovo elemento
- L'utente specifica:
  - **Nome**: del nuovo elemento
  - **Tipo**: (se diverso dal base)
  - **Dimensioni**: (se diverso dal base)
  - **Posizione**: livello e punto di connessione
- Il nuovo elemento mantiene:
  - **Isolamento** se presente sul base
  - **Parametri** specifici (se configurati)
  - **Tipo** e **Dimensioni** (se non configurati)

### 3. Creazione elemento
- Viene creato un nuovo elemento MEP
- Viene applicata la giunzione appropriata
- Viene generato un report di completamento

---

## Programmazione (schema a blocchi)

```
┌─────────────────────────────────────────┐
│  Inizio comando CreateSimilar           │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  1. Selezione elemento base               │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  2. Analisi elemento base                │
│  - Tipo, Dimensioni, Isolamento            │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  3. Apertura dialogo configurazione         │
│  - Nome, Tipo, Dimensioni, Posizione    │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  4. Creazione nuovo elemento               │
│  - Nuovo Pipe/Duct/CableTray              │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  5. Applicazione parametri                │
│  - Isolamento, Parametri specifici        │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  6. Generazione report di completamento │
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
clr.AddReference("RevitAPIUI)
```

### Selezione elemento base
```python
# Seleziona elemento base
uidoc = revit.uidoc
doc = revit.doc

# Filtra per tipo MEP
categories = [
    BuiltInCategory.OST_Pipe,
    BuiltInCategory.OST_Duct,
    BuiltInCategory.OST_CableTray
]

elements = FilteredElementCollector(doc).OfCategories(categories).ToElements()

# Seleziona elemento
selection = uidoc.Selection.PickObject(ObjectType.Element, "Select base element")
base_elem = doc.GetElement(selection)
```

### Analisi elemento base
```python
# Identifica tipo e dimensioni
if isinstance(base_elem, Pipe):
    base_type = base_elem.Symbol
    base_size = base_elem.Symbol.GetShape()
    base_insulation = base_elem.PipeInsulation
elif isinstance(base_elem, Duct):
    base_type = base_elem.Symbol
    base_size = base_elem.Symbol.GetShape()
    base_insulation = base_elem.DuctInsulation
elif isinstance(base_elem, CableTray):
    base_type = base_elem.Symbol
    base_size = base_elem.Symbol.GetShape()
```

### Creazione nuovo elemento
```python
# Crea nuovo elemento con parametri base
new_elem = MechanicalUtils.CreatePipe(
    doc=doc,
    level=level,
    point=point,
    symbol=base_type,
    size=base_size,
    insulation=base_insulation
)
```

### Applicazione parametri
```python
# Copia isolamento se presente
if base_insulation:
    new_elem.PipeInsulation = base_insulation

# Applica parametri specifici
if new_name:
    new_elem.Name = new_name

# Applica parametri configurati
if new_type:
    new_elem.Symbol = new_type
```

### Generazione report
```python
# Report di completamento
report_lines = []
report_lines.append("# CreateSimilar Report")
report_lines.append(f"\nElemento base: {base_elem.Name}")
report_lines.append(f"Nuovo elemento: {new_elem.Name}")
report_lines.append(f"Tipo: {new_elem.GetType().Name}")

output.print_md("\n".join(report_lines))
```

---

## Parametri configurabili

| Parametro | Descrizione | Valori possibili |
|------------|-------------|------------------|
| **Name** | Nome del nuovo elemento | Qualsiasi stringa |
| **Type** | Tipo (se diverso dal base) | Pipe, Duct, CableTray |
| **Size** | Dimensioni (se diverso dal base) | Valore reale |
| **Position** | Livello e punto di connessione | Qualsiasi livello del modello |

---

## Limitazioni e note

- Il nuovo elemento mantiene i parametri base (Tipo, Dimensioni, Isolamento)
- I parametri specifici possono essere configurati durante la creazione
- Se un parametro non è configurato, viene usato quello del base
- Il comando è utile per replicare verticale simili

---

## Autori

- Matteo Maiorano
