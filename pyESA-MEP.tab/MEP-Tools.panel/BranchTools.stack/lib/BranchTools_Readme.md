# Branch Tools

Stack di tre comandi per braghe e stacchi delle tubazioni (scarichi), nato da Fitting Tool 0.13.0 (vedi `BranchTools_DeveloperNote.md` per la storia). Questo file descrive la logica comune, nei moduli di questa cartella; i Readme dei pulsanti descrivono solo le loro finestre.

| Pulsante | Funzioni | Moduli |
| --- | --- | --- |
| **Pipe Wye** | braga su un tubo, collegamento di un altro tubo da sopra o di lato | `BranchTools_Pipe` |
| **Fitting Rotate** | rotazione di un raccordo, parametro Angle, allineamento e collegamento di un'estremita' libera | `BranchTools_Rotate`, `BranchTools_Align` |
| **Side Branch** | stacco laterale da una braga, copie lungo il collettore, lunghezza del tubo inclinato | `BranchTools_SideBranch`, `BranchTools_Place`, `BranchTools_Length` |

Comuni: `BranchTools_Controls` (slider con casella, righe di stato, lettura dei numeri), `BranchTools_Select` (selezione di un solo elemento, impostazioni), `BranchTools_Rotate` (anche `run_step`, la transazione silenziosa usata da tutti i passi). Tutti i moduli hanno il prefisso `BranchTools_` (Rocket mode: nomi unici in tutta l'estensione) e nessun documento a livello di modulo.

## Uso comune

- Un solo elemento: se ne e' preselezionato esattamente uno adatto al comando si usa quello, altrimenti il comando chiede di cliccarne uno (Esc annulla). Il punto cliccato su un tubo e' dove Pipe Wye mette la braga; con la preselezione si usa la mezzeria del tubo.
- I pulsanti che richiedono di cliccare nel modello (Connect from above / from the side, Place copies, Align and connect) confermano le modifiche come OK e chiudono la finestra, poi chiedono il clic: con la finestra modale aperta non si puo' cliccare nel modello. Se qualcosa non riesce compare un solo avviso breve.
- Nessun SHIFT + CLICK: CLICK apre sempre la finestra.
- **OK** conferma; **Cancel** (o Esc) riporta tutto com'era. Tutte le modifiche fatte nella finestra sono una sola voce nel menu Annulla, con il nome del comando.
- Gli slider funzionano allo stesso modo: pulsanti − e + per un passo, casella per il valore esatto (Invio o uscita dalla casella, virgola decimale accettata), `Back to start` per il valore iniziale.
- Ogni passo e' una transazione dentro il gruppo: gli avvisi di Revit durante l'anteprima sono soppressi; se Revit rifiuta il passo, quel passo e' annullato, lo slider torna all'ultimo valore riuscito e la riga di stato dice `Last change refused` con il motivo (nessuna finestra di errore, che trascinando comparirebbe a ogni passo).
- Impostazioni: una sezione di configurazione per pulsante (`ESA_PipeWye`, `ESA_FittingRotate`, `ESA_SideBranch`), elencate nei Readme dei pulsanti. Quelle di `ESA_FittingTool` non sono piu' lette.

## Tubo (Pipe Wye)

Con un tubo selezionato, che fa da collettore.

- **Wye**: la famiglia della braga usata da Place wye e dai due Connect, scelta tra i raccordi di tubazione del progetto con Part Type Tee, Cross, Wye, LateralTee o LateralCross ("Famiglia : Tipo", si puo' digitare per cercare; ricordata per nome in `wye_family`). Una croce si usa come una braga: la seconda derivazione resta libera. Con una famiglia scelta il tool la inserisce e orienta direttamente. Con `(routing preferences)` prova prima la giunzione delle routing preferences e, se Revit la rifiuta, una famiglia con Part Type Wye; senza nessuna, l'avviso chiede di scegliere la famiglia. La prima volta e' proposta una famiglia con Part Type Wye, se c'e'. La famiglia scelta deve avere due connettori in linea (collettore) e uno di lato (derivazione), altrimenti l'avviso dice che non e' una braga.
- **Verso valle**: in discesa se il collettore e' in pendenza, altrimenti il verso del tubo (dal primo al secondo estremo); **Reverse flow** lo inverte. La riga di stato dice quale verso si usa.
- **Place wye**: inserisce una braga nel punto cliccato (o in mezzeria). Rotazione e angolo dai due campi della finestra: rotazione 0 = derivazione orizzontale a destra guardando verso valle, 90 = in alto, 180 = a sinistra; angolo = angolo tra derivazione e collettore verso monte. Inserimento in due tentativi: prima la giunzione delle routing preferences del tipo di tubo (`NewTeeFitting` con un tubo di derivazione provvisorio, eliminato subito: la derivazione resta libera); se Revit la rifiuta (con le braghe degli scarichi succede: "angle ... too small or too large"), la **famiglia della braga** inserita direttamente: la giunzione delle routing preferences se ha Part Type Wye, altrimenti la prima famiglia di raccordo con Part Type Wye del progetto. La famiglia riceve Angle e diametri (dei connettori; se una misura e' guidata da una formula resta quella della famiglia), viene orientata con il collettore sull'asse del tubo e la derivazione nella direzione voluta, e il collettore viene spezzato e collegato come in Place copies. Diametro della derivazione = diametro del collettore. Transazione `Place Wye`. Lo stacco sulla braga nuova si crea poi con Side Branch.
- **Connect from above**: dopo la chiusura si clicca un tubo che passa sopra il collettore, vicino alla sua estremita' libera. Il tubo viene allungato o accorciato fino al punto sopra l'asse del collettore; da li' un tubo scende a 45° rispetto al collettore, nel suo piano verticale, verso valle, e si innesta in una braga rivolta verso l'alto. Errore se in pianta il tubo e' parallelo al collettore o se non e' sopra di esso.
- **Connect from the side**: dopo la chiusura si clicca un tubo che finisce a lato del collettore. Dalla sua estremita' un tubo entra in una braga a 45° rispetto al collettore verso valle, con l'inclinazione che ne risulta: il collettore non cambia quota.
- In entrambi i collegamenti la braga sta a valle della proiezione del punto di partenza, di quanto quel punto dista dall'asse del collettore (tan 45° = 1): l'angolo con il collettore e' esattamente 45°. Il tubo nuovo ha tipo, sistema e diametro dell'altro tubo; la braga si inserisce come in Place wye (routing preferences, poi famiglia), la curva dalle routing preferences. Transazione `Connect Pipe`. Errore se la braga cadrebbe fuori dal collettore.
- Mentre la finestra e' aperta non si puo' navigare nella vista (finestra modale).

## Rotazione

- Asse, fissato all'apertura della finestra:
  - T, Y e raccordi diritti: l'asse del collettore (i due connettori allineati e opposti), anche se e' collegata una sola estremita' o nessuna;
  - curve e altri raccordi: l'estremita' collegata, o il primo connettore se il raccordo e' libero.
- Lo slider e' **assoluto**: legge l'angolo attuale della derivazione (o dell'estremita' libera) attorno all'asse. 0 = orizzontale, 90 = verso l'alto, -90 = verso il basso, ±180 = orizzontale sul lato opposto. Lo 0 e' sul lato verso cui punta la derivazione all'apertura della finestra, quindi all'apertura il valore e' tra -90 e 90. Esempio: uno stacco a 33° si legge 33; scrivendo 90 il tool ruota di 57°.
- Con l'asse verticale (colonna) l'orizzontale non esiste: l'angolo e' misurato in pianta dalla direzione Y del progetto.
- Non ruotabili (contati nella riga di stato con il motivo): raccordi bloccati con il pin (`pinned`), senza connettori, o collegati fuori dall'asse, per esempio una curva collegata alle due estremita' o un T con la derivazione collegata (`connected on more than one axis`).
- Passi 1, 5, 15, 45, 90. Transazione `Rotate Fitting`.

## Angolo della derivazione

- Parametro: `Angle` (o `Angolo`), di istanza, numerico e modificabile. I raccordi senza questo parametro sono ignorati; se nessuno ce l'ha la sezione e' disattivata.
- Lo slider mostra il valore **assoluto** del parametro (0-180 gradi).
- Passi 1, 2.5, 5, 15, 45: con 2.5 si arriva a 67.5 e 87.5. Transazione `Set Fitting Angle`.
- Lo slider mostra il valore effettivo dopo la rigenerazione: se la famiglia lo arrotonda o lo limita, si vede subito. I valori che la famiglia non accetta vengono rifiutati da Revit (vedi sopra).
- Cambiare l'angolo non cambia la lettura della rotazione: la derivazione resta sullo stesso piano attorno al collettore.

## Stacco laterale

Disponibile con una braga selezionata (raccordo di tubazione con una derivazione libera, Y o T), collegata a un tubo collettore. Altrimenti la sezione e' disattivata e la riga di stato dice perche'.

- Campi: **Invert offset** (mm, + sopra il collettore, - sotto), **Branch slope**, **Horizontal run**, **Horizontal length** (mm, dalla curva).
- **Branch slope** (pendenza del tubo inclinato):
  - `From rotation` (predefinito): la braga resta con la rotazione e l'angolo che ha (si impostano con Fitting Rotate); la pendenza e' quella della derivazione. Esempio: Y a 45° ruotata di 45° su un collettore orizzontale = pendenza 30°;
  - 15, 22.5, 30, 45: la rotazione viene ricalcolata perche' la derivazione salga (o scenda) proprio di quell'angolo, sul lato verso cui punta all'apertura, e applicata alla braga prima di creare lo stacco. Con una Y a 45° e pendenza 45° la rotazione e' 90° (circa 78,6° con un collettore in pendenza del 2%): la derivazione sta nel piano verticale del collettore.
- **Horizontal run** (direzione in pianta dopo la curva):
  - `Perpendicular to main` (predefinito): perpendicolare al collettore, sul lato verso cui va la derivazione. Non funziona se la derivazione corre lungo il collettore (rotazione 90°): la riga di stato lo dice;
  - `Along the branch`: nella direzione in pianta della derivazione.
- La curva prende l'angolo tra il tubo inclinato e il tratto orizzontale, mostrato nella riga di stato della sezione e aggiornato a ogni movimento degli slider. Esempio: Y a 45° ruotata di 45°, tratto perpendicolare = curva di 60°; tratto lungo la derivazione = curva uguale alla pendenza (30°). La curva puo' avere un angolo non commerciale.
- `Create branch`:
  1. calcola la quota dell'asse del tratto orizzontale: fondo interno del collettore in corrispondenza della braga + offset + raggio interno dello stacco. I raggi interni vengono dal parametro Inside Diameter: quello dello stacco da un tubo di prova del tipo del collettore con il diametro della derivazione, creato e annullato subito;
  2. solo con un angolo di curva scelto: ruota la braga come sopra;
  3. crea un tubo dalla derivazione lungo il suo asse fino al piano di quella quota, una curva in quel punto (dalle routing preferences del tipo di tubo) e un tubo orizzontale nella direzione scelta. La curva sta nel punto in cui l'asse incontra il piano, quindi la quota del tratto orizzontale e' esatta qualunque sia la misura dei raccordi.
- Con `From rotation` la derivazione deve puntare verso la quota richiesta: verso l'alto con offset positivo, verso il basso con offset negativo; altrimenti la riga di stato lo dice.
- Tubi: tipo, sistema e livello del collettore, diametro della derivazione. Transazione `Create Side Branch`.
- Dopo la creazione la derivazione e' collegata: rotazione, angolo e stacco si disattivano. OK lo conferma, Cancel lo toglie.
- Errori (nella riga di stato, la finestra resta aperta):
  - curva piu' ripida di quanto la derivazione possa salire: con una Y a 45° su un collettore orizzontale la curva massima e' 45°; con una curva da 45° il tratto orizzontale corre parallelo al collettore;
  - offset troppo piccolo: la derivazione e' gia' oltre la quota richiesta, o il tubo inclinato e' troppo corto per la curva (Revit rifiuta la curva);
  - offset 0 o con l'asse dello stacco alla quota del collettore: non serve la curva.

## Copie dello stacco (Place copies)

Ripete uno stacco identico in altri punti dello stesso collettore.

- Disponibile quando la braga selezionata ha gia' uno stacco collegato alla derivazione: subito dopo `Create branch`, oppure aprendo il tool su una braga con uno stacco fatto in precedenza (anche a mano).
- `Place copies` chiude la finestra confermando le modifiche come OK (la selezione di punti non si puo' fare con la finestra aperta), poi chiede di cliccare sul collettore. Ogni clic crea una copia; Esc termina.
- Selezionabili solo i tubi sullo stesso asse del collettore della braga (anche gli spezzoni creati dalle copie precedenti).
- La copia e' una traslazione **lungo l'asse del collettore** di braga e stacco (tutti gli elementi collegati alla derivazione, al massimo 30), con il centro della braga sul punto cliccato. Rotazione, pendenza e direzioni restano identiche.
- **Copies keep this elevation** (spuntata di default, ricordata in `copy_keep_level`): con il collettore in pendenza la braga copiata sta piu' in alto o in basso dell'originale; curva e tratti successivi vengono spostati lungo il tubo inclinato di quanto serve a riportare lo stacco **alla stessa quota dell'originale**. Il tubo inclinato si allunga o si accorcia, il tratto orizzontale si sposta un poco in pianta. Errore (copia non fatta) se il tubo inclinato diventerebbe piu' corto di circa 15 mm, se la derivazione e' orizzontale o se lo stacco non comincia con un tubo.
- Casella tolta: la copia tiene l'offset dal collettore e ne segue la pendenza (quota diversa a ogni copia).
- Limite geometrico con la casella spuntata: una copia piu' a monte (collettore piu' alto) deve avere il tubo inclinato piu' corto, una piu' a valle piu' lungo. Oltre una certa distanza a monte il tubo inclinato non basta e la copia non si fa; l'avviso dice di quanti mm andrebbe accorciato e quanto e' lungo. Rimedi: cliccare piu' vicino, allungare prima lo stacco con Branch length, o togliere la spunta.
- Il tubo cliccato viene spezzato alle due estremita' della nuova braga (`PlumbingUtils.BreakCurve`), il pezzo in mezzo eliminato e le due parti collegate alla braga.
- Ogni copia e' una transazione `Place Side Branch`, quindi una voce nel menu Annulla.
- Errori: una copia che non riesce non lascia nulla (per esempio clic troppo vicino alla fine del tubo o a un raccordo). Il comando continua; a fine comando un solo avviso dice quante copie non sono riuscite e l'ultimo motivo. Se e' andato tutto bene non compare nulla.
- Se la derivazione porta a una rete (piu' di 30 elementi, o collegata ad altri tratti) la copia non si fa: si copia solo uno stacco libero.

## Allineamento (Align)

- L'elemento selezionato (tubo, canale, passerella, tubo protettivo o raccordo) deve avere un'estremita' libera.
- `Align and connect` chiude la finestra (come OK), poi si clicca l'elemento a cui allinearsi (tubo o raccordo, anche lontano). Tra le estremita' libere dei due elementi si prende la coppia piu' vicina dello stesso dominio (tubazioni, canali, ...), parallele e rivolte una verso l'altra; se non c'e', un avviso lo dice.
- `Align and connect`, due passi:
  1. **sposta l'elemento selezionato** perpendicolarmente all'asse dell'altra estremita', finche' le due sono in asse (transazione `Align Ends`);
  2. chiude la distanza lungo l'asse e collega (transazione `Connect Ends`): se l'elemento selezionato e' un tubo (canale, passerella, tubo protettivo) la sua estremita' si allunga fino all'altra; altrimenti, se l'altro e' un tubo, si allunga l'altro; se sono due raccordi, il raccordo selezionato scorre lungo l'asse. Tra due tubi di misura diversa Revit inserisce una riduzione (routing preferences); con raccordi di misura diversa non si collega.
- Se il collegamento non riesce, l'allineamento resta e l'avviso dice `Aligned, not connected` con il motivo. Due estremita' gia' in asse ma staccate si collegano con lo stesso pulsante.
- I due passi avvengono dopo la chiusura della finestra: sono due voci nel menu Annulla, fuori dal gruppo `Fitting Rotate`. Se va tutto bene non compare nulla.
- Un elemento collegato ad altri li trascina come fa Revit spostandolo a mano; se Revit rifiuta lo spostamento l'avviso lo dice.
- Estremita' non parallele: non vengono considerate (ruotare prima).

## Lunghezza dello stacco (Branch length)

Per una braga che ha gia' uno stacco collegato, anche fatto a mano, o subito dopo `Create branch`. Lo stacco deve cominciare con un tubo (il tubo inclinato).

- **Sloped pipe (L)**: slider e casella per la lunghezza del tubo inclinato, in mm (da 20 a 2000, o fino a 3 volte la lunghezza iniziale). Passi 1, 5, 10, 50, 100 mm, ricordati in `sloped_step`.
- Cambiando L, curva e tratti successivi scorrono lungo il tubo inclinato (stesso movimento usato da Copies keep this elevation): la quota dello stacco cambia, braga e collettore restano fermi, pendenza e direzioni restano uguali.
- **Invert offset (H)**: scorrimento attuale dello stacco, fondo interno del tubo dopo la curva meno fondo interno del collettore in corrispondenza della braga. Scrivendo un valore (Invio o uscita dalla casella) il tool calcola L: `L = L attuale + (H voluto - H attuale) / pendenza`. Non funziona con la derivazione orizzontale.
- Transazione `Set Branch Length` per ogni passo, dentro il gruppo della finestra (Cancel annulla).

## Testi nella finestra

- Le spiegazioni stanno nel tooltip dell'icona `i` in alto a destra ("How to use", come in Clash Setup e FlexConnect). Se cambia il comportamento di un comando va aggiornato anche quel testo nel suo `<Tool>_Form.xaml`.
- Le righe grigie di stato compaiono solo quando c'e' qualcosa da dire: raccordi non ruotabili, parametro Angle mancante o diverso tra i raccordi, motivo per cui una sezione e' spenta, angolo della curva calcolato, errori (`Last change refused`, `Not created`).

## Limiti

- Il tratto orizzontale e' orizzontale: nessuna pendenza.
- Isolamento e parametri del collettore non vengono copiati sui tubi nuovi.
- Lo stacco non inserisce la braga: deve gia' esserci sul collettore.
