# Branch

## Logica di funzionamento

Il comando **Branch** crea un ramo orizzontale partendo da verticali di tubo, condotto o passerella, con una finestra per categoria (tubi: anche direzione libera con un clic nella vista; passerelle: solo dal lato stretto, ramo di costa; nessun angolo di inclinazione). **La modalità SHIFT+CLICK (braga Y) è stata spostata nel pulsante Branch Y: le sezioni di questo file che la descrivono sono storiche.** La logica è strutturata in tre fasi principali:

### 1. Selezione e preparazione
- Il sistema legge l'elenco di verticali selezionati (riser)
- Per ogni verticale, calcola l'asse (axis) e la lunghezza totale
- Verifica che sia un elemento valido (Pipe, Duct o CableTray)

### 2. Creazione ramo orizzontale
- L'utente specifica:
  - **Lato**: destra/sinistra rispetto all'asse
  - **Livello**: livello + offset
  - **Lunghezza**: del ramo orizzontale
  - **Angolo**: 15°, 30°, 45° o 60° (solo SHIFT+CLICK)
- Viene creata una curva 3D che definisce il percorso del ramo
- La curve è vincolata nel piano orizzontale specificato

### 3. Connettori e giunzioni
- **Tipo di giunzione** dipende dal tipo di elemento:
  - Pipe → Tee (default)
  - Duct → Tap (default)
  - CableTray → Tee (obbligatorio)
- Se il primo tentativo fallisce, prova l'opzione alternativa (es. se default è Tee, prova Tap)
- Se il fitto non entra, il riser resta invariato e viene segnalato nel report

### 4. Parametri ereditati
- Il nuovo ramo mantiene:
  - **Tipo** (Type Mark) del riser
  - **Dimensioni** (Size) del riser
  - **Isolamento** del riser (PipeInsulation / DuctInsulation)
- Se il fitto non entra, il riser resta invariato e il report spiega perché

### 5. Modalità SHIFT+CLICK (Y branch)
- Quando si usa SHIFT+CLICK:
  - Creazione di un ramo Y su verticale di tubo
  - Passo 1: ramo orizzontale (obbligatorio)
  - Passo 2: ramo inclinato con gombo
  - Passo 3: ramo Y sul riser
  - Il punto Y è posizionato sotto il ramo orizzontale

---

## Programmazione (schema a blocchi)

```
┌─────────────────────────────────────────┐
│  Inizio comando Branch                  │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  1. Selezione verticale (riser)              │
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
│  - Lato, Livello, Lunghezza, Angolo        │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  4. Creazione curve 3D del ramo          │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  5. Verifica e creazione giunzione         │
│  - Tipo giunzione (Tee/Tap)                │
│  - Prova alternativa se fallisce          │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  6. Applicazione parametri ereditati       │
│  - Tipo e dimensioni riser                │
│  - Isolamento                             │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  7. Generazione report di errore/fallito  │
└─────────────────────────────────────────┘
```

**Per SHIFT+CLICK (Y branch):**
```
┌─────────────────────────────────────────┐
│  Inizio Y Branch                         │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  1. Seleziona verticale + SHIFT+CLICK   │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  2. Dialogo configurazione Y branch         │
│  - Lato, Livello, Lunghezza, Angolo       │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  3. Creazione ramo orizzontale           │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  4. Creazione ramo inclinato             │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  5. Creazione ramo Y sul riser             │
│  - Posizionato sotto ramo orizzontale      │
└────────────────┬──────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│  6. Generazione report                    │
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

# Importazioni .NET per XAML
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
# Per Pipe, Duct, CableTray: usa LocationCurve per ottenere la curva
# Verifica che la curva sia quasi verticale (punti quasi allineati verticalmente)
curve = elem.Location.Curve
p1, p2 = curve.GetEndPoint(0), curve.GetEndPoint(1)
verticality = abs(p1.X - p2.X) + abs(p1.Y - p2.Y) < tolerance
```

### Creazione curve 3D del ramo
```python
# Curve 3D definite nel piano orizzontale specificato
# Usiamo CurveCreator per creare curve spline
# La curve è vincolata nel livello specificato
curve_creator = CurveCreator.CreateArc(
    doc=doc,
    level=level,
    point=start_point,
    length=length,
    angle=angle,
    side=side
)
```

### Creazione fitto (Tee/Tap)
```python
# Dipende dal tipo di elemento
if isinstance(elem, Pipe):
    # Pipe → Tee (default) o Tap
    fitting = MechanicalUtils.CreateTee(doc, riser, branch)
    if not fitting:
        fitting = MechanicalUtils.CreateTap(doc, riser, branch)
elif isinstance(elem, Duct):
    # Duct → Tap (default) o Tee
    fitting = MechanicalUtils.CreateTap(doc, riser, branch)
    if not fitting:
        fitting = MechanicalUtils.CreateTee(doc, riser, branch)
elif isinstance(elem, CableTray):
    # CableTray → Tee (obbligatorio)
    fitting = MechanicalUtils.CreateTee(doc, riser, branch)
```

### Applicazione parametri ereditati
```python
# Copia tipo e dimensioni dal riser
riser_type = riser.Symbol
riser_size = riser.Symbol.GetShape()

# Copia isolamento se presente
if isinstance(riser, Pipe) and riser.GetParameter("PipeInsulation"):
    new_pipe.PipeInsulation = riser.PipeInsulation

# Applica parametri al nuovo ramo
new_branch.Symbol = riser_type
new_branch.Size = riser_size
```

### Generazione report
```python
# Report dettagliato per ogni errore o fallimento
report_lines = []
for riser in risers:
    if fitting_failed:
        report_lines.append(
            f"Riser {riser.Id}: {reason} - No branch created"
        )
output.print_md("\n".join(report_lines))
```

---

## Parametri configurabili

| Parametro | Descrizione | Valori possibili |
|------------|-------------|------------------|
| **Side** | Lato del ramo rispetto all'asse | Left, Right |
| **Level** | Livello di posizionamento | Qualsiasi livello del modello |
| **Length** | Lunghezza del ramo | 0.01 - 1000+ m (valore reale) |
| **Angle** | Angolo di inclinazione | 15°, 30°, 45°, 60° (solo SHIFT+CLICK) |
| **Junction Type** | Tipo di giunzione | Tee, Tap (dipende dal tipo elemento) |
| **Offset** | Offset dal livello | Valore reale o "None" per livello base |

---

## Limitazioni e note

- **CableTray**: solo creazione da lato stretto
- **Pipe**: usa Tee come default, Tap come alternativa
- **Duct**: usa Tap come default, Tee come alternativa
- **PipeInsulation/DuctInsulation**: copiato dal riser se presente
- Se il fitto non entra, il riser resta invariato
- SHIFT+CLICK crea ramo Y con tre segmenti: orizzontale → inclinato → Y

---

## Autori

- Matteo Maiorano
