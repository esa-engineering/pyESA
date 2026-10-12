# Pulsanti nascosti

pyRevit non legge le cartelle senza suffisso riconosciuto (`.pushbutton`, `.pulldown`, `.stack`...), quindi
quello che sta qui non compare nella barra.

- `Openings.pulldown`: forometrie (Links, Setup, Place, Refresh), in sviluppo.
- `CreateSimilar.pushbutton`: Create Similar.

Per rimetterli: spostare la cartella in `MEP-Routing.panel` e aggiungere il nome (`Openings`,
`CreateSimilar`) al `layout` di `MEP-Routing.panel/bundle.yaml`.
