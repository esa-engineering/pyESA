# Relink Models

Controlla i link Revit (e, a scelta, CAD) di uno o piu' modelli **chiusi** e li riaggancia ai file in una nuova cartella.

Il pulsante fa due cose:

- **Check**: per ogni link dice se il file esiste al percorso salvato e in che stato era il link al salvataggio del modello. Non scrive nulla.
- **Apply**: sostituisce il percorso di tutti i link spuntati, anche quelli che oggi funzionano.

## Come si usa

1. **Add files...**: scegli i modelli `.rvt`. Devono essere chiusi in Revit.
2. **Folder** (facoltativa per il Check): cartella dove ora si trovano i file collegati. `Include subfolders` la esplora per intero.
3. **Options**: tipi di link (Revit, CAD), copia di sicurezza prima della scrittura, visibilita' dei link non presenti nella nuova cartella.
4. **Check**: compila la tabella.
5. Controlla la tabella. Puoi togliere la spunta, correggere a mano `New path` o usare `Pick file for selected...`.
6. **Apply**: scrive i nuovi percorsi nei modelli.

## Organizzazione della tabella

Una riga per **link** (coppia modello che contiene il link, file linkato), non per file: un file linkato da tre modelli compare tre volte, una sotto ogni modello.

Le righe sono raggruppate per modello. La prima colonna, `Model / Link`, mostra sull'intestazione del gruppo (in grassetto) il nome del modello con il numero dei suoi link e, sotto, il nome di ogni link rientrato. Le altre colonne dell'intestazione riportano i totali: file mancanti, link non caricati, link trovati nella nuova cartella, esito di Apply. Doppio clic sull'intestazione: chiude o apre il gruppo. Spunta sull'intestazione: spunta o toglie tutti i link del modello che hanno un nuovo percorso. Le colonne non si ordinano, per non rompere i gruppi.

Un modello che non si legge ha solo l'intestazione, con il motivo in `Result`.

La ℹ️ in alto a destra della finestra mostra, al passaggio del mouse, la guida rapida (HOW TO USE e lettura della tabella), come in Clash Setup.

## Colonne

| Colonna | Contenuto |
| --- | --- |
| `File exists` | `Yes` / `No`: il file esiste ora al percorso salvato nel modello |
| `Saved state` | stato del link all'ultimo salvataggio: `Loaded`, `Unloaded`, `Unloaded (locally)`, `Not found`, `In closed workset`, ... |
| `In new folder` | solo con una cartella: `Match`, `Already correct`, `Ambiguous (n files)`, `Not in folder`, `Picked by hand` |
| `Result` | esito di Apply: `Written`, `Failed: <motivo>`, `No new path`; `Cannot read: <motivo>` se il modello non si legge |

Il riepilogo sotto la tabella conta file mancanti, link non caricati, righe spuntate e righe nascoste.

## Logica

- Non apre i modelli: usa `TransmissionData` (`ReadTransmissionData` / `WriteTransmissionData`), la stessa tabella dei riferimenti esterni usata da Transmit Model. Per questo e' veloce anche su molti modelli.
- `Saved state` viene dalla tabella dei riferimenti, quindi e' lo stato al momento dell'ultimo salvataggio, non uno stato "live".
- `File exists` interroga il disco o la rete una volta per percorso.
- Il file collegato si cerca per **nome file** (senza distinzione tra maiuscole e minuscole). Se lo stesso nome compare in piu' cartelle vince quella che condivide piu' cartelle finali con il vecchio percorso; a parita' il caso e' `Ambiguous` e il percorso resta vuoto.
- **Sostituzione in ogni caso**: ogni link con un file trovato nella nuova cartella e' spuntato, anche se oggi funziona o punta gia' li' (`Already correct`). Restano senza spunta solo `Ambiguous` e `Not in folder`.
- I link `Not in folder` (file non spostati, ad esempio i modelli dei consulenti rimasti sul server) sono nascosti quando c'e' una cartella; `Show links not in the new folder` li mostra. Senza cartella si vedono tutti.
- L'elenco contiene solo i link diretti di ogni modello: i link annidati non sono nella tabella dei riferimenti del modello ospite.
- Un link scaricato resta scaricato. Un link con file non trovato viene riscritto come caricato.
- I percorsi scritti sono assoluti, anche se il link originale era relativo (`Relative to Central Model`).
- Dopo la scrittura il modello viene riletto e il percorso confrontato; se non coincide, `Result` riporta cosa e' stato letto.
- Il modello viene marcato come "trasmesso" (`IsTransmitted`): senza questo flag Revit ignora i nuovi percorsi all'apertura. I nuovi percorsi diventano effettivi quando il modello viene aperto in Revit e salvato.
- Finche' non succede, il Check mostra ancora il vecchio percorso in `Current path` e in `Result` scrive `Rewritten, not yet opened in Revit` (il nuovo percorso e' quello in `New path`). Aprendo a mano un modello condiviso, Revit chiede se salvarlo come nuova central.

## Opzione `Open and save as new central after writing`

Fa in automatico il passaggio che altrimenti si fa a mano in Revit. Dopo la scrittura, ogni modello viene aperto in background: Revit applica i nuovi percorsi, il tool lo salva come nuova central nella stessa posizione (sovrascrivendo il file), rilascia la proprieta' degli elementi e lo chiude. Un modello non condiviso viene solo aperto e salvato. Alla fine il tool rilegge il file: il nuovo percorso deve risultare quello salvato.

- Esito in `Result`: `Written, saved as new central` oppure `Written, but saving as new central failed: <motivo>` (i percorsi restano scritti: si puo' aprire il modello a mano).
- Lento: minuti per modello, perche' Revit apre il modello e carica i link. Revit resta occupato fino alla fine.
- Una nuova central rende inutilizzabili i file locali della vecchia: vanno ricreati. Da usare quando nessuno lavora sui modelli.
- L'opzione e' spenta a ogni apertura della finestra.

## Input e output

- Input: modelli `.rvt` chiusi, cartella di ricerca (facoltativa per il Check).
- Output: i modelli riscritti sul posto. Con `Back up each model before writing` una copia va in `_RelinkBackup\<data_ora>_<nome>.rvt` accanto al modello (quella cartella e' ignorata nella ricerca).
- Nessun report finale: lo stato di ogni riga e' nella tabella.
- Impostazioni salvate in `ESA_RelinkModels` (modelli, cartella, opzioni).

## Limiti

- I riferimenti cloud (ACC) sono saltati.
- Un modello aperto in Revit non si puo' scrivere: la riga risulta `Failed`.
- Non c'e' SHIFT + CLICK: l'operazione riscrive file e richiede sempre l'anteprima.
- Da verificare in Revit (vedi DeveloperNote): se Revit applica il nuovo percorso all'apertura, comportamento su central workshared e su file di versione Revit diversa.
