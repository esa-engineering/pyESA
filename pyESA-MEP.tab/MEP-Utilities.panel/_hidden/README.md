# Pulsanti nascosti

pyRevit non legge le cartelle senza suffisso riconosciuto (`.pushbutton`, `.pulldown`, `.stack`...), quindi
quello che sta qui non compare nella barra.

- `Openings.pulldown`: Openings (Place, Refresh, Families, Settings).
- `Settings.pulldown`: Report Mode e Section Check. Finche' e' nascosto, la modalita' del report resta
  quella salvata (predefinita: report solo in caso di problemi).

Per rimetterli: spostare la cartella in `MEP-Utilities.panel` e aggiungere il nome (`Openings`, `Settings`) al
`layout` di `MEP-Utilities.panel/bundle.yaml`.
