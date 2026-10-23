# Fitting Rotate

Pulsante dello stack Branch Tools. La logica e' descritta in `lib/BranchTools_Readme.md`, sezioni "Rotazione", "Angolo della derivazione" e "Allineamento (Align)".

## Uso

- Selezione: un raccordo (tubazioni, canali, passerelle, tubi protettivi), oppure un tubo se serve solo Align. Preselezionato se e' l'unico elemento selezionato, altrimenti cliccato.
- Finestra, una colonna:
  - **Rotate**: slider assoluto dell'angolo della derivazione attorno all'asse (0 = orizzontale, 90 = in alto); spento con un tubo selezionato.
  - **Branch angle**: slider del parametro `Angle` della famiglia; spento se il parametro non c'e'.
  - **Align and connect**: conferma come OK, chiude la finestra e chiede di cliccare l'elemento a cui allinearsi; poi allinea e collega.
- OK conferma rotazioni e angolo (una voce `Fitting Rotate` nel menu Annulla), Cancel li annulla. Allineamento e collegamento, fatti dopo la chiusura, sono due voci a parte.

## Impostazioni (`ESA_FittingRotate`)

| Chiave | Contenuto |
| --- | --- |
| `step` | passo dello slider Rotate, gradi |
| `angle_step` | passo dello slider Branch angle, gradi |
