# Set Pipe Insulation By System: note di sviluppo

Registro dei tentativi fatti sullo strumento, per non ripetere soluzioni già provate e
scartate. Non è il Readme (come funziona lo strumento oggi) né lo storico revisioni nello
script (cosa cambia in ogni versione). Voci dalla più recente.

Le versioni fino alla 2.1 non hanno note: risultano solo dai commit (#35, #36, #37, #46).
Nel #46 sono state tolte dalla cartella del pulsante le varianti `SetInsulation_script_v1.py`
e `SetInsulation_script_v2.py`.

---

## 22.10.2026, versione 2.2

**Richiesta.** 1) Portare nello strumento il caso dello script *Set Pipe Insulation -
DPR412-93*: una spunta prima del caricamento della tabella txt; se attiva, la barra del
percorso diventa grigia e si usa la tabella scritta nel codice. 2) Mostrare nella finestra
un'anteprima, con colonne e righe, della tabella txt appena importata.

**Fatto.**

- Nuovo modulo `SetPipeInsulation_Table.py` (nessun import Revit o WPF): `THERMAL_ZONES`,
  `DPR412_TABLE` (copiata dallo script DPR412-93, valori identici al template txt),
  `builtin_table()` (restituisce copie: con Rocket mode il modulo resta in memoria e la
  tabella non deve essere modificabile da un comando all'altro), `read_table_file()`.
- `read_table_file()` sostituisce `load_insulation_table()` dello script. Non mostra alert:
  solleva `TableError` con il messaggio per l'utente e conta le celle non intere invece di
  stamparle nel pannello di output. Aggiunta la tolleranza al BOM UTF-8 e agli spazi nelle
  intestazioni: prima un file salvato con BOM perdeva la colonna `Diameter`.
- La lettura e il controllo del file si spostano nella finestra: l'errore si vede subito,
  nella riga di stato e all'OK con un `MessageBox`, e la finestra resta aperta. Prima il
  file veniva letto dopo l'OK e un errore chiudeva lo strumento. La finestra restituisce la
  tabella già letta in `values["table"]` (non salvata in config: non è in `DEFAULTS`).
- Nuova chiave di configurazione `builtin_table` (default `False`, così chi usa già lo
  strumento non vede cambiare nulla).
- Finestra su due colonne (940 x 640, minimo 800 x 520) senza `ScrollViewer`: a sinistra
  sistema, tipi di isolamento, spessore fisso e nota informativa; a destra la sezione della
  tabella con spunta, percorso, hint, `DataGrid` di anteprima e riga di stato. Con
  l'anteprima la colonna unica sarebbe diventata troppo lunga.
- Lo spessore fisso è passato in una sezione propria (*FIXED THICKNESS*); il suo
  comportamento non cambia e non disattiva la sezione della tabella.

**Tentativi e correzioni.**

- *Colonne della griglia generate in automatico* (`AutoGenerateColumns`, come in
  `MEPQTO_columns_ui`): scartato senza provarlo. I nomi di colonna con lo spazio
  (`Non riscaldato`) nei `Binding` sono fragili. Si usano nomi neutri `C0..Cn` nel
  `DataTable` e colonne `DataGridTextColumn` create in codice, con l'intestazione vera.
- *Tutte le colonne a larghezza `*`*: le intestazioni "Diameter (mm)" e "Non riscaldato"
  risultavano troncate (verificato con la finestra renderizzata fuori da Revit). Ora le
  colonne sono `Auto` con `MinWidth` 60 e solo l'ultima è `*`.
- *Altezza 600*: la griglia mostrava 13 righe e mezza della tabella DPR, con barra di
  scorrimento. Portata a 640: le 14 righe stanno senza scorrere.
- *Campo percorso solo disattivato* (`IsEnabled = False`): con il tema WPF predefinito il
  campo resta quasi bianco, non "grigio" come richiesto. Aggiunto uno stile locale al
  `TextBox` con un trigger su `IsEnabled = False` che imposta lo sfondo `#E6E6E6`.
- Riga di stato: corretto il plurale ("1 cell is..." / "N cells are...").

**Esito.** Funziona fuori da Revit: sintassi dei tre file `.py` verificata; lettore provato
in CPython sul template (risultato identico a `DPR412_TABLE`), su un file con BOM e cella
non intera, su un file senza `Diameter`; finestra istanziata con il motore IronPython 2.7 di
pyRevit e provata con spunta DPR, percorso vuoto, file non valido, file con BOM e spessore
fisso. **Non provato in Revit**: la parte che modifica il modello non è cambiata, ma va
verificata con Reload.

**Decisione.** Il pulsante *Set Pipe Insulation - DPR412-93* resta nel pulldown finché il
proprietario non decide di toglierlo. Differenza da ricordare: quello lavora sulla sola
selezione senza filtro per tipo di sistema, questo richiede sempre un tipo di sistema.
