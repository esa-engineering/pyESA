# WorksetViews

## Logica di funzionamento

Il comando **WorksetViews** verifica e gestisce gli elementi visibili nel modello, con focus particolare su:

### 1. Identificazione elementi visibili
- Il sistema scandisce tutti gli elementi del modello
- Filtra per categorie MEP:
  - **Pipe**: tutti i tubi
  - **Duct**: tutti i canali
  - **CableTray**: tutti i cavi
- Verifica se ogni elemento è **visibile** nel piano corrente
  - Visibilità in base al livello
  - Visibilità in base al Workset
  - Visibilità in base al View Visibility

### 2. Verifica Workset
- Per ogni elemento, verifica:
  - **Workset**: a cui appartiene
  - **Is in Workset**: se appartiene a un workset attivo
  - **Workset Name**: nome del workset
- Identifica elementi in:
  - Workset attivo
  - Workset inattivo
  - Workset specifico

### 3. Verifica View Visibility
- Verifica se l'elemento è visibile nel view corrente:
  - **View Visibility**: parametro di visibilità
  - **Hidden**: se nascosto
  - **Visible**: se visibile
- Identifica elementi:
  - Visibili nel view
  - Nascosti nel view

### 4. Report dettagliato
- Genera un report con:
  - **Statistiche generali**: numero totale di elementi, visibili, nascosti
  - **Elementi visibili**: lista per Type Mark
  - **Elementi nascosti**: lista per Type Mark
  - **Elementi per Workset**: distribuzione per workset
  - **Elementi per View Visibility**: distribuzione per visibilità

---

## Programmazione (schema a blocchi)

```
┌─────────────────────────────────────────┐
│  Inizio comando WorksetViews           │
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
│  2. Verifica visibilità del view          │
│  - Elemento visibile/nascosto            │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  3. Verifica Workset                    │
│  - Workset attivo/inattivo                │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  4. Raggruppa per Type Mark               │
│  - Dict: Type Mark → [Elementi]           │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  5. Generazione report                    │
│  - Statistiche, Elementi visibili, Nascosti │
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

### Verifica visibilità del view
```python
# Verifica se elemento è visibile nel view corrente
def is_visible_in_view(elem):
    """Verifica se elemento è visibile nel view"""
    # Verifica se elemento è nascosto
    if elem.IsHiddenInView(doc.ActiveView):
        return False
    return True

# Raggruppa per visibilità
visible_elements = []
hidden_elements = []

for elem in mech_elements:
    if is_visible_in_view(elem):
        visible_elements.append(elem)
    else:
        hidden_elements.append(elem)
```

### Verifica Workset
```python
# Verifica Workset di ogni elemento
def get_workset(elem):
    """Recupera Workset di elemento"""
    workset = elem.Document.GetWorkset(elem.Id)
    return workset.Name if workset else "None"

# Raggruppa per Workset
workset_dict = {}
for elem in mech_elements:
    workset = get_workset(elem)
    if workset not in workset_dict:
        workset_dict[workset] = []
    workset_dict[workset].append(elem)
```

### Raggruppa per Type Mark
```python
# Raggruppa per Type Mark
def get_type_mark(elem):
    """Recupera valore Type Mark"""
    param = elem.GetParameter("Type Mark")
    if param and param.StorageType == StorageType.String:
        return param.AsString()
    return "None"

type_mark_dict = {}
for elem in mech_elements:
    type_mark = get_type_mark(elem)
    if type_mark not in type_mark_dict:
        type_mark_dict[type_mark] = []
    type_mark_dict[type_mark].append(elem)
```

### Generazione report
```python
# Report in formato Markdown
report_lines = []
report_lines.append("# Workset Views Check")
report_lines.append(f"\n## Statistiche Generali")
report_lines.append(f"- Total elements: {len(mech_elements)}")
report_lines.append(f"- Visible: {len(visible_elements)}")
report_lines.append(f"- Hidden: {len(hidden_elements)}")
report_lines.append(f"- Unique Type Marks: {len(type_mark_dict)}")

# Elementi visibili
if visible_elements:
    report_lines.append(f"\n## Visible Elements")
    for type_mark, elements in type_mark_dict.items():
        visible_in_type = [e for e in elements if e in visible_elements]
        if visible_in_type:
            report_lines.append(f"### {type_mark}")
            for elem in visible_in_type:
                workset = get_workset(elem)
                report_lines.append(
                    f"- {elem.Name} ({elem.GetType().Name}): Workset = {workset}"
                )

# Elementi nascosti
if hidden_elements:
    report_lines.append(f"\n## Hidden Elements")
    for type_mark, elements in type_mark_dict.items():
        hidden_in_type = [e for e in elements if e in hidden_elements]
        if hidden_in_type:
            report_lines.append(f"### {type_mark}")
            for elem in hidden_in_type:
                workset = get_workset(elem)
                report_lines.append(
                    f"- {elem.Name} ({elem.GetType().Name}): Workset = {workset}"
                )

# Elementi per Workset
if workset_dict:
    report_lines.append(f"\n## Elements per Workset")
    for workset, elements in workset_dict.items():
        visible_count = sum(1 for e in elements if e in visible_elements)
        hidden_count = len(elements) - visible_count
        report_lines.append(
            f"- {workset}: {len(elements)} total, "
            f"{visible_count} visible, {hidden_count} hidden"
        )

output.print_md("\n".join(report_lines))
```

---

## Parametri configurabili

Nessun parametro configurabile - il comando verifica automaticamente tutti gli elementi.

---

## Limitazioni e note

- Il report include tutti gli elementi MEP del modello
- Visibilità è verificata nel view corrente
- Workset è verificato in base al documento
- Il comando è utile per verificare consistenza nella distribuzione degli elementi

---

## Autori

- Matteo Maiorano
