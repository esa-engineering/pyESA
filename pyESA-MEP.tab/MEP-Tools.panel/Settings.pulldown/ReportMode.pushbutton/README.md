# ReportMode

## Logica di funzionamento

Il comando **ReportMode** genera un report dettagliato su tutte le verticale (riser) del modello, elencando:

### 1. Elementi identifcati
- **Tipo**: Pipe, Duct o CableTray
- **Element ID**: identificativo univoco
- **Livello**: a cui appartiene
- **Dimensioni**: Size e Type Mark
- **Isolamento**: PipeInsulation o DuctInsulation se presente

### 2. Informazioni geometriche
- **Lunghezza totale**: della verticale
- **Asse**: (X, Y, Z) del punto di connessione
- **Posizione**: sopra/sotto il punto di connessione
- **Numero di connessioni**: con altri elementi

### 3. Stati
- **Connesso**: se ha almeno una connessione
- **Non connesso**: se non ha connessioni
- **Gap**: se ci sono spazi vuoti
- **Giunzioni**: tipo di giunzione usata

### 4. Report dettagliato
- Genera un report in formato Markdown con:
  - Tabella con tutte le verticale
  - Statistiche generali (numero totale, connessi, non connessi)
  - Elenco elementi per tipo
  - Elementi con problemi (gap, isolamento mancante, etc.)

---

## Programmazione (schema a blocchi)

```
┌─────────────────────────────────────────┐
│  Inizio comando ReportMode                │
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
│  2. Filtra verticale (riser)             │
│  - Linee quasi verticali                  │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  3. Calcola informazioni per ogni verticale │
│  - Asse, lunghezza, connessioni           │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  4. Aggrega statistiche                │
│  - Totale, connessi, non connessi        │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  5. Generazione report Markdown           │
│  - Tabella elementi                       │
│  - Statistiche                           │
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
from Autodesk.Revit.DB.Plumbing import Pipe, PipeInsulation, PipingSystemType, PlumbingUtils
from Autodesk.Revit.DB.Mechanical import Duct, DuctInsulation, DuctLining, MechanicalSystemType, MechanicalUtils
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

### Filtra verticale (riser)
```python
# Identifica verticale basate su linee quasi verticali
def is_vertical_line(curve):
    """Verifica se una curva è quasi verticale"""
    p1, p2 = curve.GetEndPoint(0), curve.GetEndPoint(1)
    verticality = abs(p1.X - p2.X) + abs(p1.Y - p2.Y)
    return verticality < tolerance

risers = []
for elem in mech_elements:
    curve = elem.Location.Curve
    if is_vertical_line(curve):
        risers.append(elem)
```

### Calcolo informazioni verticale
```python
# Calcola informazioni per ogni verticale
def get_riser_info(riser):
    """Recupera informazioni geometriche"""
    curve = riser.Location.Curve
    p1, p2 = curve.GetEndPoint(0), curve.GetEndPoint(1)
    
    # Asse
    axis_point = Vector3(p2.X - p1.X, p2.Y - p1.Y, p2.Z - p1.Z)
    
    # Lunghezza
    length = p1.DistanceTo(p2)
    
    # Connessioni
    connections = get_connected_elements(riser)
    
    return {
        "element": riser,
        "axis": axis_point,
        "length": length,
        "connections": connections
    }
```

### Generazione report
```python
# Report in formato Markdown
report_lines = []
report_lines.append("# MEP Verticals Report")
report_lines.append(f"\n## Statistiche Generali")
report_lines.append(f"- Total verticals: {len(risers)}")
report_lines.append(f"- Connected: {connected_count}")
report_lines.append(f"- Disconnected: {disconnected_count}")

# Tabella elementi
report_lines.append(f"\n## Elementi per tipo")
report_lines.append(f"- Pipes: {pipe_count}")
report_lines.append(f"- Ducts: {duct_count}")
report_lines.append(f"- Cable Trays: {cable_tray_count}")

# Elenco verticale
report_lines.append(f"\n## Verticale List")
for riser in risers:
    info = get_riser_info(riser)
    report_lines.append(
        f"- {riser.Name} ({riser.GetType().Name}): "
        f"{info['length']:.2f}m, {len(info['connections'])} conn"
    )

output.print_md("\n".join(report_lines))
```

---

## Parametri configurabili

Nessun parametro configurabile - il comando genera un report automatico.

---

## Limitazioni e note

- Il report include tutti gli elementi MEP del modello
- Verticale sono identificate come linee quasi verticali
- Connessioni sono calcolate in base alla geometria
- Isolamento è segnato se presente

---

## Autori

- Matteo Maiorano
