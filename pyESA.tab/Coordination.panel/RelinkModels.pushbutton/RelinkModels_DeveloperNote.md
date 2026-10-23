# Relink Models - note di sviluppo

Registro dei tentativi. Piu' recente in alto.

## 23.10.2026 - v1.6.0 - pulizia della finestra (con AI)

**Richiesta** (da screenshot dell'utente): togliere la colonna Model, vuota sulle righe dei link, unendola a Link; togliere da `Result` il percorso ridondante; opzioni su due colonne; guida "how to use" come in Clash Setup.
**Fatto**:

- Colonna unica `Model / Link`: sull'intestazione `▼ modello.rvt  (n links)`, sotto il nome del link rientrato (`LINK_INDENT`). Il conteggio dei link, prima nella colonna Link dell'intestazione, e' ora accanto al nome del modello. Il nome interno della colonna e' `Name`, perche' una `/` nel nome di una colonna di DataTable rompe il binding delle colonne generate (in un Path WPF la `/` indica il "current item"); il titolo mostrato si imposta in `OnAutoColumn`.
- `Result` scrive solo `Rewritten, not yet opened in Revit`: il percorso e' gia' in `New path`.
- Opzioni su una griglia a due colonne: a sinistra cosa leggere e mostrare (Revit links, CAD links, Show links not in the new folder), a destra come scrivere (backup, nuova central). Etichetta accorciata in `Open and save as new central`, il "dopo la scrittura" e' nel tooltip.
- Guida: la ℹ️ nel titolo con tooltip HOW TO USE / READING THE TABLE, copiata da `ClashFinder_form.xaml`. Non il pulsante "?" di Clash Setup: apre una guida a slide, che qui non esiste.

**Stato**: compilato e XAML caricata in WPF, non provato in Revit.

## 23.10.2026 - v1.5.1 - spunta Apply (con AI)

**Problema** (utente, in Revit): la spunta degli elementi "e' un po' buggata".
**Causa**: la colonna `Apply` era la `DataGridCheckBoxColumn` generata da `AutoGenerateColumns`. Comportamento noto del DataGrid WPF: primo clic seleziona la riga, secondo entra in modifica e cambia il valore, che viene confermato solo all'uscita dalla riga (`CommitEdit`). Quindi spunte che "non prendono", e la spunta di un'intestazione propagata ai link solo dopo un clic altrove (rischio gia' annotato in v1.4.0).
**Scartato**: lasciare la colonna generata e forzare `CommitEdit` in `CurrentCellChanged` o simili: rincorre il DataGrid, e la riga resta in modifica mentre `_refresh_headers()` riscrive le intestazioni.
**Fatto**: in `OnAutoColumn` la colonna `Apply` e' sostituita da una `DataGridTemplateColumn` (sola lettura per il DataGrid) con una `CheckBox` legata in `OneWay` ad `Apply`. Il clic arriva a `OnApplyClick` (handler di `ButtonBase.ClickEvent` sul DataGrid), che scrive direttamente nel `DataRow` (quindi `ColumnChanged` propaga a gruppo e riepilogo subito) e riallinea la checkbox con `UpdateTarget()`. Una riga di link senza nuovo percorso non si spunta. Il doppio clic sulla checkbox non chiude il gruppo (`_inside_checkbox`).
**Verificato fuori da Revit** (PowerShell STA, finestra WPF con dispatcher): la checkbox segue il dato quando lo cambia il codice, il clic accettato resta, quello rifiutato torna indietro, il binding sopravvive ai clic. Senza dispatcher il binding non si aggiorna: falso negativo del test, non del codice.
**Stato**: da riprovare in Revit.

## 23.10.2026 - v1.5.0 - opzione nuova central (con AI)

**Esito del test v1.4.1** (utente, in Revit): **funziona**. Con `IsTransmitted = True` aprendo il modello i link puntano al nuovo percorso, e Revit chiede se salvarlo come nuova central.
**Richiesta**: un'opzione per ricreare i file come central, per rendere definitivi i percorsi.
**Valutato e scartato**: aprire ogni modello (Detach), `RevitLinkType.LoadFrom` su ogni link e salvare. Funzionerebbe senza TransmissionData, ma duplica un metodo appena verificato e apre domande non provate (LoadFrom dentro/fuori transazione, link in workset chiusi, CAD).
**Scelto**: riusare la scrittura TransmissionData e automatizzare la domanda di Revit. Firme verificate per reflection su RevitAPI.dll 2024: `DetachFromCentralOption.ClearTransmittedSaveAsNewCentral`, `WorksharingSaveAsOptions.SaveAsCentral` / `ClearTransmitted`, `SaveAsOptions.OverwriteExistingFile`, `WorksharingUtils.RelinquishOwnership`, `BasicFileInfo.IsWorkshared`. `save_as_new_central()` in `RelinkModels_Core.py`: apre con quell'opzione (solo se workshared), fa `SaveAs` come central sullo stesso percorso se l'apertura non l'ha gia' fatto (controllo su `GetWorksharingCentralModelPath()`), rilascia la proprieta', chiude; un modello non condiviso viene aperto e salvato. Verifica: il nuovo percorso deve risultare `LastSaved` nel file.
**Scelte**: opzione non ricordata tra un uso e l'altro (sovrascrive delle central); cursore di attesa durante Apply, nessuna barra di avanzamento.
**Stato**: non provato. Da verificare: (a) se `ClearTransmittedSaveAsNewCentral` salva gia' all'apertura o serve il `SaveAs` (il codice gestisce entrambi i casi); (b) il `SaveAs` sullo stesso percorso del file aperto; (c) dialoghi di Revit durante l'apertura in background (link non trovati, avvisi), che potrebbero bloccare; (d) tempi sul modello da 500 MB.

## 23.10.2026 - v1.4.1 - IsTransmitted (con AI) - FUNZIONA (test utente)

**Test dell'utente** (v1.4.0): Apply su E-PN0001 e M-PN0001, poi chiusura e riapertura: "non funziona".
**Verifica sui file**: entrambi riscritti alle 10:47 con `DesiredPath` = nuovo percorso (9 e 8 link), `LastSavedPath` ancora il vecchio, header `<TransmissionData isTransmitted="false">`. Quindi la scrittura c'e', ma con `isTransmitted="false"` Revit non applica il `DesiredPath` all'apertura. Conferma il punto 3 lasciato aperto in v1.0.0. Inoltre il Check rilegge il `LastSaved`, quindi riaprendo il tool si vedono ancora i vecchi percorsi.
**Fatto**:

- `td.IsTransmitted = True` prima di `WriteTransmissionData`.
- Il Check legge anche il dato desired (`GetPath()`, non `GetAbsolutePath()`) e, se diverso dal last saved, scrive in `Result` `Rewritten, not yet opened in Revit: <percorso>`; l'intestazione del modello li conta.

**Stato**: non provato. Da verificare aprendo in Revit un modello riscritto con la v1.4.1: (a) i link puntano al nuovo percorso; (b) che messaggi da' Revit su un modello workshared "trasmesso" (atteso: avviso di modello trasmesso, eventuale richiesta di salvarlo come nuova central); (c) se il flag si azzera dopo il salvataggio. Se (a) fallisce anche con il flag, resta l'alternativa: apertura in background (`OpenDocumentFile` + `RevitLinkType.LoadFrom` + salvataggio).
**Nota**: i due modelli scritti con la v1.4.0 vanno riscritti (Apply di nuovo), perche' hanno il flag a false.

## 23.10.2026 - v1.4.0 - tabella raggruppata per modello (con AI)

**Richiesta**: con 11 modelli nella cartella l'utente vedeva 24 righe e non capiva come leggerle.
**Analisi** (lettura diretta della TransmissionData degli 11 file): 35 link Revit in tutto, 22 verso file presenti nella cartella. Le righe sono coppie modello -> link, quindi lo stesso file compare una volta per ogni modello che lo linka (E-PN0001 e M-PN0001 da soli fanno 19 righe). I link vengono da 9 cartelle diverse (P:\ e tre server), vari link sono Unloaded. Le 2 righe in piu' rispetto ai 22 sono probabilmente intestazioni di modelli illeggibili (righe `Cannot read`, mai filtrate). Non verificato.
**Alternative proposte**: vista per file linkato (una riga per file, "usato da N modelli"), vista per modello raggruppata, selettore tra le due. Scelta dell'utente: **raggruppata per modello**.
**Fatto**: outline in un solo DataTable come la distinta di MEPQTO (non WPF grouping su DataView): righe `Kind = H` (intestazione con totali) e `Kind = L` (link), colonne nascoste `Group` e `Hidden`; doppio clic su un'intestazione = apri/chiudi; spunta sull'intestazione propagata ai link via `DataTable.ColumnChanged`, con il flag `_syncing` per non rientrare quando le scritture vengono dal codice. Ordinamento colonne disattivato (`CanUserSortColumns="False"`). Stile intestazione con `DataTrigger` sulla colonna `Kind` nella XAML.
**Stato**: sintassi verificata, non provato in Revit. Da controllare: il `DataTrigger` su una colonna non generata, e la spunta delle checkbox nel DataGrid (il valore si conferma solo all'uscita dalla riga).

## 23.10.2026 - v1.3.0 - Check + sostituzione in ogni caso (con AI)

**Richiesta**: il pulsante fa due cose: (1) check se i link sono presenti e caricati; (2) sostituzione del percorso in ogni caso. Chiarito con l'utente: il check = file esiste al percorso salvato + stato del link; "in ogni caso" = riscrivere anche i link che oggi funzionano (non un trova/sostituisci del prefisso).
**Fatto**:

- `Scan` diventa `Check`; la cartella nuova e' facoltativa (senza, la tabella mostra solo il controllo).
- Nuove colonne `File exists` (`os.path.isfile`, con cache per percorso: un server irraggiungibile si interroga una volta) e `Saved state` (`ExternalFileReference.GetLinkedFileStatus()` del dato "last saved", testo in `LOAD_STATE_TEXT` di `RelinkModels_Core.py`).
- `Status` diviso in `In new folder` (esito della ricerca) e `Result` (esito di Apply). L'esito "file non nella cartella" si chiama ora `Not in folder`, per non confonderlo con lo stato `Not found` del link.
- Spunta automatica su ogni riga con un nuovo percorso, `Already correct` compreso (prima solo `Match`).

**Limite noto**: `Saved state` e' lo stato all'ultimo salvataggio, non quello che Revit vedrebbe aprendo il modello ora. Per lo stato "live" serve `File exists`.
**Stato**: sintassi verificata, non provato in Revit. Resta aperto il punto decisivo di v1.2.0 (Revit applica il `DesiredPath` all'apertura?).

## 22.10.2026 - v1.2.0 - verifica corretta, link non trovati nascosti (con AI)

**Esito del test v1.1.0** (4 modelli workshared copiati sul Desktop, cartella `LinkPatcher_Test\BIM Models`):

- Ancora 71 link: il filtro su `GetReferencingId()` non toglieva nulla.
- Le 8 righe spuntate: `Failed: not confirmed on re-read`.

**Analisi fatta leggendo il file .rvt** (lo stream TransmissionData e' XML UTF-16 in chiaro, `<TransmissionData isTransmitted="false" version="5">` con un `<ExternalFileReference>` per riferimento: `ElementId`, `LastSavedPath`, `LastSavedAbsolutePath`, `LastSavedPathType`, `LastSavedLoadState`, `DesiredPath`, `DesiredPathType`, `DesiredLoadState`):

- **La scrittura funziona**: nel modello Lighting i due link M-6001 / M-6002 hanno `DesiredPath` = nuovo percorso, `DesiredPathType` = Absolute; nel backup no. Era sbagliata la verifica: sul dato "desired" `GetAbsolutePath()` non restituisce il percorso scritto. Ora si confrontano `GetPath()` e `GetAbsolutePath()` e, se non coincidono, lo Status riporta cosa e' stato letto.
- **Non c'erano link annidati**: l'XML non ha alcuna informazione di annidamento e i 18 link Revit del Lighting sono tutti diretti (consulenti su `\\10.10.2.157\...\02-CONSULTANTS`). Ipotesi v1.1.0 sbagliata, filtro `GetReferencingId()` rimosso. Il "rumore" erano i link ai file non spostati (`Not found`): ora nascosti di default, con la casella `Show links not found in the folder`.
- **Tipo di percorso**: molti link originali sono `Relative to Central Model` (`..\02-CONSULTANTS\...`, o solo il nome file per i modelli nella stessa cartella). Il tool li riscrive `Absolute`: cambia il comportamento a un futuro spostamento. Da decidere se preservare il tipo relativo.

**Ancora da verificare in Revit (punto decisivo)**: con `isTransmitted="false"` Revit applica il `DesiredPath` all'apertura? Aprire il modello Lighting (Detach, preserve worksets) e controllare Manage Links. Se i link puntano ancora al server, provare `td.IsTransmitted = True` prima di `WriteTransmissionData` (attenzione: un modello workshared "trasmesso" all'apertura chiede di diventare nuova central).

## 22.10.2026 - v1.1.0 - link annidati nascosti (con AI) - IPOTESI SBAGLIATA, vedi v1.2.0

**Richiesta**: l'elenco mostrava 71 link da 4 modelli: oltre ai link diretti c'erano i link dei link (annidati). Servono solo i diretti.
**Verifica API** (reflection su RevitAPI.dll 2022): le firme di `TransmissionData` usate in v1.0.0 sono tutte corrette. `ExternalFileReference` non ha una proprieta' "annidato"; ha `GetReferencingId()`.
**Tentativo**: un riferimento e' considerato annidato se il suo `GetReferencingId()` e' l'id di un altro riferimento della stessa tabella. Ipotesi non verificata: non so cosa restituisca `GetReferencingId()` per i link diretti. I valori sono scritti con `logger.debug` (pyRevit in debug mode, Ctrl + click sul pulsante) come `id=... referencing=... status=...`.
**Stato**: da provare in Revit. Se i link annidati restano o sparisce anche qualche link diretto, servono quei log per correggere il criterio.
**Aperto**: 8 link spuntati risultavano non scritti. Motivo non ancora letto (colonna Status fuori vista); ipotesi: erano link annidati, per cui `SetDesiredReferenceData` puo' fallire. Colonna Status spostata prima dei percorsi per leggerlo senza scorrere.

## 22.10.2026 - v1.0.1 - fix scrittura colonne in sola lettura (con AI)

**Problema** (provato in Revit): Apply falliva con `ReadOnlyException: Column 'Status' is read only`.
**Tentativo scartato**: `DataTable.Columns[x].ReadOnly = True` per rendere non modificabili le colonne dell'anteprima. Non va: la proprieta' vale anche per le scritture da codice (`row[COL_STATUS] = ...`), quindi bloccava il tool stesso.
**Soluzione**: nessun `ReadOnly` sul DataTable; la sola lettura e' impostata sulle colonne generate del DataGrid (`args.Column.IsReadOnly = True` in `OnAutoColumn`), che blocca solo la modifica a mano. Stato: da riprovare in Revit.

## 22.10.2026 - v1.0.0 - prima versione (con AI)

**Richiesta**: funzione che riagganci i link dopo lo spostamento dei file in un'altra cartella, anche a modelli chiusi e su piu' modelli in un colpo.

**Scelta**: `TransmissionData` invece dell'apertura in background (`OpenDocumentFile` + `RevitLinkType.LoadFrom` + salvataggio). Motivo: non apre il modello, e' molto piu' veloce, non tocca worksets e non richiede sincronizzazione. L'apertura in background resta il ripiego se `TransmissionData` non funzionasse sui modelli reali.

**Stato: scritto, sintassi verificata con IronPython 2.7, NON ancora provato in Revit.** Il matching (`RelinkModels_Match.py`, Python puro) e' stato provato fuori da Revit con cartelle finte: caso unico, ambiguo, non trovato, gia' corretto, cartella `_RelinkBackup` ignorata.

**Punti da verificare in Revit (firme scritte a memoria, non confrontate con la documentazione):**

1. `SetDesiredReferenceData(ElementId, ModelPath, PathType, bool)` e `GetDesiredReferenceData`: nomi e firme.
2. `ExternalFileReference.GetLinkedFileStatus()` e `LinkedFileStatus`: se mancano, `_should_load` ricade su `True` (un link scaricato verrebbe quindi caricato).
3. Se la sola `WriteTransmissionData` basta perche' Revit usi il nuovo percorso all'apertura, o se serve `td.IsTransmitted = True`. Non impostato in questa versione. Se il link non viene riagganciato all'apertura, provare questo per primo.
4. Comportamento su central workshared (non so se la scrittura e' ammessa) e su modelli salvati con una versione di Revit diversa da quella in uso.
5. Percorsi di rete UNC e percorsi relativi nei modelli: `GetAbsolutePath()` dovrebbe risolverli, da controllare.
6. DataGrid con `AutoGenerateColumns` su `DataTable`: la colonna `Apply` bool deve comparire come casella e `Idx` deve restare nascosta (`AutoGeneratingColumn`).

**Decisioni di progetto:**

- Niente SHIFT + CLICK: il tool riscrive file, quindi l'anteprima non si puo' saltare (regola di CLAUDE.md: SHIFT + CLICK non salta la conferma di un'operazione distruttiva).
- Una sola finestra: scan, anteprima e Apply senza chiudere, cosi' gli esiti restano visibili riga per riga.
- Backup attivo di default in `_RelinkBackup` accanto al modello.
- Ambito della v1.0.0: link Revit e CAD, riferimenti cloud saltati. Immagini, nuvole di punti e altri riferimenti esterni non sono trattati.
