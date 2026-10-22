# Set Pipe Insulation By System

Aggiunge o aggiorna l'isolamento delle tubazioni e dei raccordi di un tipo di sistema di
tubazioni. Il tipo di isolamento dipende dall'esposizione dell'elemento (a vista / non a
vista); lo spessore viene da una tabella per diametro nominale e zona termica (la tabella
DPR 412/93 interna o un file di testo), oppure è un valore fisso.

## Finestra

La finestra è su due colonne. A sinistra:

- **Piping system**: il tipo di sistema (`RBS_PIPING_SYSTEM_TYPE_PARAM`, valore mostrato da
  Revit) e l'ambito: tutte le tubazioni e i raccordi del tipo di sistema in tutto il
  modello, oppure solo quelli selezionati che appartengono a quel tipo di sistema. La
  seconda scelta è disattivata se lo strumento parte senza selezione.
- **Insulation types**: tipo di isolamento per gli elementi a vista e per quelli non a vista.
- **Fixed thickness**: se spuntata, lo spessore scelto (6…60 mm) vale per tutti gli
  elementi e la tabella non viene usata.

A destra, **Thickness table**:

- spunta **Use the DPR 412/93 table built into the tool**: si usa la tabella scritta nel
  codice (`SetPipeInsulation_Table.DPR412_TABLE`, gli stessi valori dello strumento *Set
  Pipe Insulation - DPR412-93*). Il campo del percorso e *Browse...* si disattivano (grigi)
  e il percorso non viene controllato;
- senza la spunta, **Table file**: il file di testo con la tabella (vedi "Formato del
  file"). Il file viene letto quando lo si sceglie con *Browse...* o quando si esce dal
  campo dopo averlo scritto a mano, e di nuovo all'OK se il percorso è cambiato;
- **anteprima**: la griglia mostra la tabella che verrà usata, una riga per diametro e una
  colonna per zona termica, già ordinata per diametro. Sotto, una riga di stato: numero di
  righe, zone termiche trovate, celle ignorate perché non intere, oppure il motivo per cui
  il file non è utilizzabile. Con la spunta DPR la griglia mostra la tabella interna.

All'OK la finestra controlla i dati e resta aperta se qualcosa manca: tipo di sistema, i
due tipi di isolamento e, se lo spessore non è fisso e la spunta DPR è spenta, un file
esistente con almeno una riga valida.

Le scelte (tipo di sistema, ambito, tipi di isolamento, spunta DPR, percorso del file,
spessore fisso) sono salvate in `script.get_config('ESA_SetPipeInsulation')` e riproposte
all'apertura successiva. La spunta DPR al primo uso è spenta. Lo strumento non distingue
CLICK e SHIFT + CLICK: la finestra si apre sempre.

## Formato del file

Testo separato da tabulazioni, prima riga di intestazione. Il modello è
`insulation_table_template.txt` nella cartella del pulsante (ALT + CLICK sul pulsante apre
la cartella); contiene gli stessi valori della tabella DPR 412/93.

```
Diameter	Interno	Non riscaldato	Interrato	Esterno
10	6	13	13	25
15	9	13	13	25
...
```

- La colonna `Diameter` (diametro nominale in mm) è obbligatoria.
- Le colonne delle zone termiche sono `Interno`, `Non riscaldato`, `Interrato`, `Esterno`
  (maiuscole e spazi come scritti). Ne serve almeno una; le altre possono mancare. Altre
  colonne sono ignorate.
- Valori interi, in mm. Una cella non intera viene ignorata e contata nella riga di stato;
  una riga senza diametro intero viene scartata. Le righe con meno celle dell'intestazione
  (comprese le righe vuote) vengono scartate.
- Sono accettati il BOM UTF-8 (file salvati da Blocco note) e gli spazi attorno alle
  intestazioni.

La lettura è in `SetPipeInsulation_Table.read_table_file()`, senza codice Revit né WPF:
la usano sia la finestra (anteprima e controllo) sia, tramite i valori restituiti dalla
finestra, lo script.

## Calcolo per ogni elemento

1. **Esposizione**: parametro di istanza `e_EAN_IsVisible_1` (Sì/No). Sì = tipo *Exposed*;
   No, parametro assente o di tipo diverso = tipo *Concealed*.
2. **Spessore fisso**: se scelto, vale per tutti e i passi 3-5 si saltano.
3. **Diametro**: per le tubazioni `RBS_PIPE_DIAMETER_PARAM`; per i raccordi il primo
   parametro fra `Diametro`, `Diameter`, `Nominal Diameter`, `Diametro nominale`, e se non
   c'è il diametro del primo connettore circolare. In mm, arrotondato all'intero. Senza
   diametro l'elemento viene saltato.
4. **Zona termica**: parametro testo `e_EAN_ThermalPosition_1`, che deve valere esattamente
   una delle quattro zone. Altrimenti l'elemento viene saltato.
5. **Spessore**: riga della tabella con il diametro **più vicino** a quello dell'elemento
   (a pari distanza vince il diametro minore), colonna della zona termica. Se la zona non è
   fra le colonne del file, l'elemento viene saltato.
6. **Isolamento**: se l'elemento non ne ha, viene creato (`PipeInsulation.Create`). Se ne ha
   già uno, si scrive lo spessore nel primo parametro il cui nome contiene *thickness* o
   *spessore* e si cambia il tipo; se nessuna delle due cose riesce, l'isolamento viene
   eliminato e ricreato.

Tutto avviene in una sola transazione (*Set pipe and fitting insulation*, un solo passo di
annullamento).

## Esito

Nel pannello di output: una riga per elemento (aggiunto, aggiornato, ricreato, saltato con
il motivo, errore) e un riepilogo con tipo di sistema, ambito, origine dello spessore (fisso,
DPR 412/93 o nome del file) e la tabella dei conteggi per tubazioni e raccordi.

## Limiti noti

- Il diametro si confronta con quello **più vicino** in tabella, non con la classe
  superiore: un DN fuori dai valori della tabella prende lo spessore del diametro più
  vicino, anche se è più piccolo.
- La ricerca dell'isolamento esistente scorre tutti gli isolamenti del modello per ogni
  elemento: su modelli grandi lo strumento può essere lento.
- L'ambito *tutto il sistema* lavora per **tipo** di sistema, non per singola istanza di
  sistema (SS.1, SS.2...).
- Il riconoscimento del parametro di spessore dell'isolamento esistente si basa sul nome
  (inglese o italiano).
