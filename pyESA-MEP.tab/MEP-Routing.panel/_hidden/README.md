# Pulsanti nascosti

pyRevit non legge le cartelle senza suffisso riconosciuto (`.pushbutton`, `.pulldown`, `.stack`...), quindi
quello che sta qui non compare nella barra.

- (`Openings.pulldown` e' stato spostato in `MEP-Utilities.panel`.)
- `CreateSimilar.pushbutton`: Create Similar.
- `Rise&Turn.pushbutton`: Rise&Turn (pettine di tubi contro una parete), sostituito nella barra da Quick Extend.

Per rimetterli: spostare la cartella in `MEP-Routing.panel` e aggiungere il nome (`Openings`,
`CreateSimilar`) al `layout` di `MEP-Routing.panel/bundle.yaml`. `Rise&Turn.pushbutton` torna invece in
`Rise&Turn.stack`, con il nome `Rise&Turn` nel `layout` di `Rise&Turn.stack/bundle.yaml`.
