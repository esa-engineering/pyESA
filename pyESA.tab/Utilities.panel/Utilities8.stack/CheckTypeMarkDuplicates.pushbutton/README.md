# CheckTypeMarkDuplicates

## Logica di funzionamento

Il comando **CheckTypeMarkDuplicates** verifica se esistono duplicati nel campo **Type Mark** di tutti gli elementi MEP del modello. Questo è utile per:

### 1. Identificazione duplicati
- Il sistema scandisce tutti gli elementi MEP del modello:
  - **Pipe**: tutti i tubi
  - **Duct**: tutti i canali
  - **CableTray**: tutti i cavi
- Per ogni elemento, recupera il valore di **Type Mark**
- Raggruppa gli elementi per valore di Type Mark

### 2. Rilevamento duplicati
- Se un Type Mark ha **2 o più elementi** associati:
  - Viene segnalato come duplicato
  - Viene generato un report dettagliato
- Se un Type Mark ha **solo 1 elemento**:
  - Viene considerato valido
- Se un Type Mark ha **0 elementi**:
  - Viene considerato vuoto

### 3. Report dettagliato
- Genera un report con:
  - **Statistiche generali**: numero totale di elementi, Type Mark unici, duplicati
  - **Elenco duplicati**: per ogni Type Mark duplicato, lista degli elementi
  - **Elementi singoli**: per ogni Type Mark unico, lista degli elementi
  - **Elementi vuoti**: Type Mark senza elementi

---

## Programmazione (schema a blocchi)

```
┌─────────────────────────────────────────┐
│  Inizio comando CheckTypeMarkDuplicates │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  1. Recupera tutti gli elementi MEP      │
│  - Pipe, Duct, CableTray                  │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  2. Estrae Type Mark per ogni elemento     │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  3. Raggruppa per Type Mark                │
│  - Dict: Type Mark → [Elementi]           │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  4. Identifica duplicati                    │
│  - Type Mark con > 1 elemento            │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  5. Generazione report                    │
│  - Statistiche, Elenco duplicati, Elementi singoli │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  6. Stampa report in output panel         │
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

### Recupero elementi MEP
```python
# Recupera tutti gli elementi MEP del modello
doc = revit.doc
mech_elements = []

# Pipe
pipes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Pipe).ToElements()
for pipe in pipes:
    mech_elements.append(pipe)

# Duct
ducts = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Duct).ToElements()
for duct in ducts:
    mech_elements.append(duct)

# CableTray
cable_trays = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).ToElements()
for cable_tray in cable_trays:
    mech_elements.append(cable_tray)
```

### Estrazione Type Mark
```python
# Estrae Type Mark per ogni elemento
def get_type_mark(elem):
    """Recupera valore Type Mark"""
    param = elem.GetParameter("Type Mark")
    if param and param.StorageType == StorageType.String:
        return param.AsString()
    return "None"

# Raggruppa per Type Mark
type_mark_dict = {}
for elem in mech_elements:
    type_mark = get_type_mark(elem)
    if type_mark not in type_mark_dict:
        type_mark_dict[type_mark] = []
    type_mark_dict[type_mark].append(elem)
```

### Identifica duplicati
```python
# Identifica duplicati
duplicates = {}
single_elements = []
empty_type_marks = []

for type_mark, elements in type_mark_dict.items():
    if len(elements) > 1:
        duplicates[type_mark] = elements
    elif len(elements) == 1:
        single_elements.append(elements[0])
    else:
        empty_type_marks.append(type_mark)
```

### Generazione report
```python
# Report in formato Markdown
report_lines = []
report_lines.append("# Type Mark Duplicates Check")
report_lines.append(f"\n## Statistiche")
report_lines.append(f"- Total elements: {len(mech_elements)}")
report_lines.append(f"- Unique Type Marks: {len(type_mark_dict)}")
report_lines.append(f"- Duplicates: {len(duplicates)}")
report_lines.append(f"- Single elements: {len(single_elements)}")

# Elenco duplicati
if duplicates:
    report_lines.append(f"\n## Duplicates")
    for type_mark, elements in duplicates.items():
        report_lines.append(f"\n### {type_mark}")
        for elem in elements:
            report_lines.append(f"- {elem.Name} ({elem.GetType().Name})")

# Elementi singoli
if single_elements:
    report_lines.append(f"\n## Single Elements")
    for elem in single_elements:
        report_lines.append(
            f"- {elem.Name} ({elem.GetType().Name}): Type Mark = {get_type_mark(elem)}"
        )

# Elementi vuoti
if empty_type_marks:
    report_lines.append(f"\n## Empty Type Marks")
    for type_mark in empty_type_marks:
        report_lines.append(f"- {type_mark}")

output.print_md("\n".join(report_lines))
```

---

## Parametri configurabili

Nessun parametro configurabile - il comando verifica automaticamente tutti gli elementi.

---

## Limitazioni e note

- Il report include tutti gli elementi MEP del modello
- Type Mark è estratto dal parametro "Type Mark" (se presente)
- Se un elemento non ha Type Mark, viene considerato "None"
- Il comando è utile per mantenere consistenza nella nomenclatura

---

## Autori

- Matteo Maiorano
