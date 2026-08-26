# Riassegna Livello — strumento pyRevit

Riassegna il livello di riferimento agli elementi MEP mantenendo invariata la loro
quota assoluta nel modello.

Unifica in un unico comando le quattro varianti realizzate originariamente in Dynamo:

| File Dynamo originale | Corrispondenza nel nuovo strumento |
| --- | --- |
| `LevelReassignment_OnSelectedLevel_OnSelectedElements` | Elementi selezionati + Livello manuale |
| `LevelReassignment_OnSelectedLevel_OnActiveViewElements` | Vista attiva + Livello manuale |
| `LevelReassignment_Auto_OnSelectedElements` | Elementi selezionati + Livello automatico |
| `LevelReassignment_Auto_OnActiveViewElements` | Vista attiva + Livello automatico |

## Contenuto della cartella

| File | Ruolo |
| --- | --- |
| `script.py` | Logica dello strumento (Revit API, motore IronPython) |
| `LevelReassignmentWindow.xaml` | Interfaccia grafica della finestra di dialogo |
| `bundle.yaml` | Etichetta e descrizione del pulsante nella barra pyRevit |
| `icon.png` | Icona del pulsante |
| `README.md` | Questo file |

## Installazione

1. Individuare la cartella della propria estensione pyRevit, per esempio
   `%APPDATA%\pyRevit\Extensions\ESA.extension\ESA.tab\MEP.panel\`.
   Se non esiste ancora, si può crearne una nuova estensione dal pannello
   pyRevit → Settings → Custom Extension Directories.
2. Copiare l'intera cartella `LevelReassignment.pushbutton` all'interno del
   `.panel` desiderato, mantenendo tutti i file insieme.
3. In Revit, premere **pyRevit → Reload**. Il pulsante *Riassegna Livello*
   compare nel pannello.

La struttura risultante deve essere:

```
ESA.extension\
  ESA.tab\
    MEP.panel\
      LevelReassignment.pushbutton\
        script.py
        LevelReassignmentWindow.xaml
        bundle.yaml
        icon.png
```

## Uso

Il comando apre una sola finestra con tutte le scelte.

**1. Elementi da elaborare**

- *Elementi selezionati in Revit*: opera sulla selezione corrente. La voce è
  disabilitata se non c'è nulla di selezionato.
- *Elementi MEP della vista attiva*: opera su tutti gli elementi delle categorie
  gestite visibili nella vista. La voce è disabilitata sulle viste che non
  ammettono la raccolta di elementi (per esempio gli abachi).

Il conteggio degli elementi disponibili è mostrato sotto ciascuna voce.

**2. Livello di destinazione**

- *Livello scelto manualmente*: lo stesso livello per tutti gli elementi. Se la
  vista attiva è associata a un livello, quello è proposto come predefinito.
- *Livello automatico*: per ogni singolo elemento viene cercato il primo livello
  di progetto immediatamente sottostante il suo punto di riferimento (punto di
  inserimento, oppure punto medio della linea per gli elementi lineari). Se
  nessun livello risulta sottostante, si usa il livello più basso del progetto.

Con il metodo automatico si può indicare lo **spessore del pacchetto di
finitura** in millimetri: viene sottratto alla quota del livello prima del
confronto, così un elemento appoggiato sopra un pavimento finito viene comunque
attribuito al livello sottostante. Il valore si applica a tutti e tre i gruppi
di categorie, come nei grafi Dynamo originali.

**3. Opzioni**

- *Simulazione*: calcola e riporta il risultato senza aprire alcuna transazione,
  quindi senza modificare il modello. Utile per verificare l'esito prima di
  applicarlo.

Al termine, la finestra di output di pyRevit riporta quanti elementi sono stati
modificati per gruppo, quanti erano già sul livello corretto, quanti sono stati
ignorati e per quale motivo, con collegamenti cliccabili ai singoli elementi.

## Logica applicata per categoria

| Gruppo | Categorie | Parametri modificati |
| --- | --- | --- |
| **A** — elementi puntuali | Air Terminals, Communication / Data / Security / Nurse Call / Telephone / Fire Alarm / Lighting Devices, Electrical Equipment e Fixtures, Lighting Fixtures, Mechanical Equipment, Plumbing Fixtures, Sprinklers, Specialty Equipment, Generic Models | `Level` + ricalcolo di `Elevation from Level` |
| **B** — elementi lineari | Cable Trays, Conduits, Ducts, Flex Ducts, Flex Pipes, Pipes | solo `Reference Level` |
| **C** — raccordi e accessori | Cable Tray Fittings, Conduit Fittings, Duct Accessories, Duct Fittings, Pipe Accessories, Pipe Fittings | solo `Level` |

Per il gruppo A la quota assoluta resta invariata perché il nuovo offset è
calcolato come *(quota Z originale − quota del nuovo livello)*. Per i gruppi B e
C la geometria è definita da coordinate reali e Revit aggiorna da sé l'offset,
quindi non serve alcun ricalcolo.

Gli elementi non basati su livello (host based, workplane based, famiglie in
place) e quelli con i parametri di livello in sola lettura vengono ignorati e
riportati nel resoconto.

## Modifiche rispetto ai grafi Dynamo

1. **Categorie per enumeratore.** I grafi confrontavano il nome testuale della
   categoria, che dipende dalla lingua di Revit e che era fonte di incongruenze
   tra i quattro file (`Flex Duct` / `Flex Ducts`, `Speciality Equipment` /
   `Specialty Equipment`). Ora si usano gli enumeratori `BuiltInCategory`.
2. **Parametri per enumeratore.** I parametri vengono cercati prima tramite
   `BuiltInParameter` e solo in subordine per nome, in inglese e in italiano.
   Lo strumento funziona quindi anche su installazioni Revit localizzate.
3. **Unità di misura.** I grafi mescolavano lo spessore in millimetri con quote
   lette in unità di progetto: funzionavano correttamente solo nei progetti con
   unità di lunghezza in millimetri. Ora tutti i calcoli avvengono nelle unità
   interne di Revit e la conversione dai millimetri è esplicita, quindi il
   risultato è indipendente dalle unità di progetto.
4. **Quota di riferimento del livello.** I grafi usavano `Level.ProjectElevation`;
   lo script usa `Level.Elevation`, che è espressa nello stesso sistema di
   coordinate dei punti di inserimento ed è quindi la grandezza corretta per il
   calcolo dell'offset. Le due coincidono in tutti i progetti in cui la base
   delle quote non è traslata. Per tornare al comportamento Dynamo è sufficiente
   modificare la funzione `level_reference_elevation()`.
5. **`Pipe Accessories`.** Nei quattro file originali era classificata in modo
   incoerente: gruppo A in un file, gruppo C negli altri tre. È stata mantenuta
   nel gruppo C.
6. **`Generic Models` e `Lighting Fixtures`.** Erano presenti nel gruppo A di uno
   solo dei quattro file. Sono state mantenute nell'elenco unificato; per
   escluderle basta rimuoverle da `GROUP_A_CATEGORY_NAMES` in `script.py`.

Tutte le liste di categorie e di parametri sono raccolte in un blocco di
configurazione in testa a `script.py`, per poterle adattare senza toccare la
logica.

## Motore Python

Lo script gira sul motore **IronPython**, che è quello predefinito di pyRevit: il
file non deve contenere la direttiva `#! python3`.

Il modulo `pyrevit.forms`, usato per caricare la finestra XAML, si appoggia a
`wpf.LoadComponent`, disponibile solo in IronPython. Sotto il motore CPython3 le
finestre WPF di pyRevit non sono attualmente supportate: si veda la issue
[pyrevitlabs/pyRevit#3033](https://github.com/pyrevitlabs/pyRevit/issues/3033).

## Avvertenze

- Con il livello automatico, prestare particolare attenzione in presenza di doppi
  volumi, mezzanini o piani sfalzati: il criterio "primo livello sottostante" può
  non produrre il risultato atteso.
- Conviene provare l'esecuzione in *Simulazione* su un sottoinsieme di elementi
  prima di applicare la modifica a un'intera vista.
