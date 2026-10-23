# Side Branch

Pulsante dello stack Branch Tools. La logica e' descritta in `lib/BranchTools_Readme.md`, sezioni "Stacco laterale", "Copie dello stacco (Place copies)" e "Lunghezza dello stacco (Branch length)".

## Uso

- Selezione: una braga (raccordo di tubazione) sul collettore. Preselezionata se e' l'unico elemento selezionato, altrimenti cliccata.
- Finestra, due colonne:
  - **Side branch** (sinistra): Invert offset, Branch slope, Horizontal run, Horizontal length, `Create branch`; `Place copies` conferma come OK, chiude la finestra e ripete lo stacco a ogni clic sullo stesso collettore (Esc termina); **Copies keep this elevation**.
  - **Branch length** (destra): per una braga con lo stacco gia' collegato (anche appena creato): slider della lunghezza L del tubo inclinato e casella dello scorrimento H.
- Branch slope `From rotation` usa la rotazione attuale della braga: si imposta prima con Fitting Rotate.
- OK conferma (una voce `Side Branch` nel menu Annulla), Cancel annulla stacco, rotazioni e lunghezze. Le copie, fatte dopo la chiusura, sono una voce `Place Side Branch` ciascuna.

## Impostazioni (`ESA_SideBranch`)

| Chiave | Contenuto |
| --- | --- |
| `branch_offset_mm`, `branch_length_mm` | Invert offset e Horizontal length |
| `branch_elbow_mode` | Branch slope in gradi, 0 = From rotation |
| `branch_direction` | 1 = perpendicolare al collettore, 0 = come la derivazione |
| `copy_keep_level` | 1 = le copie tengono la quota dell'originale |
| `sloped_step` | passo dello slider Branch length, mm |
