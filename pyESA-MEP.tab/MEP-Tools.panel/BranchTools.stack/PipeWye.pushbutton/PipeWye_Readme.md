# Pipe Wye

Pulsante dello stack Branch Tools. La logica (inserimento della braga, collegamenti, verso di flusso) e' descritta in `lib/BranchTools_Readme.md`, sezione "Tubo (Pipe Wye)".

## Uso

- Selezione: un tubo, il collettore (preselezionato se e' l'unico elemento selezionato, altrimenti cliccato; il punto cliccato e' dove va la braga, con la preselezione la mezzeria).
- Finestra:
  - **Place wye**: braga nel punto, con **rotation** (0 = derivazione orizzontale a destra guardando verso valle, 90 = in alto, 180 = a sinistra) e **angle** (tra derivazione e collettore verso monte) dai due campi. Dopo l'inserimento la braga si lavora con Side Branch e Fitting Rotate.
  - **Connect from above** / **Connect from the side**: confermano come OK, chiudono la finestra e chiedono di cliccare l'altro tubo vicino alla sua estremita' libera.
  - **Reverse flow** inverte il verso di valle; la riga di stato dice quale verso si usa.
  - **Wye**: famiglia della braga tra i raccordi a T, a croce o a braga del progetto, o `(routing preferences)`.
- OK conferma, Cancel annulla quello che e' stato fatto nella finestra (una voce `Pipe Wye` nel menu Annulla). I collegamenti, fatti dopo la chiusura, sono una voce `Connect Pipe` a parte.

## Impostazioni (`ESA_PipeWye`)

| Chiave | Contenuto |
| --- | --- |
| `wye_roll`, `wye_angle` | rotazione e angolo della braga da inserire, in gradi |
| `wye_family` | etichetta "Famiglia : Tipo" della braga; vuota = routing preferences |
