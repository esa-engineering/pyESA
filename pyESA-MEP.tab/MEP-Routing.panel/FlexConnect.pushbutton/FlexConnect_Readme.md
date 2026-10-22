# FlexConnect

Collega un tubo a un'attrezzatura meccanica o idraulica con un flexpipe, e un condotto a un terminale d'aria o a un'attrezzatura meccanica con un flexduct. Il tipo di collegamento si riconosce dall'elemento cliccato.

## Uso

- **CLICK**: esegue con le ultime impostazioni. Al primo uso (o se un tipo salvato non esiste piu' nel modello) apre la finestra.
- **SHIFT + CLICK**: apre la finestra delle impostazioni, poi esegue. Eccezione voluta alla regola comune dei tool ESA (CLICK = finestra, SHIFT + CLICK = esecuzione immediata); l'inversione della regola per tutti i tool e' proposta in `docs/Claude_Suggestions.md`.
- L'ordine di selezione e' libero: si parte dal tubo o condotto, oppure dall'attrezzatura o terminale (anche preselezionando un solo tubo o condotto). Il clic sul condotto e' un punto, non solo un elemento.
- Il comando resta attivo: dopo ogni collegamento riparte la selezione. Con un tubo si riparte da capo (tubo o attrezzatura, l'ordine e' libero); la preselezione vale solo per il primo collegamento. Sui condotti, dopo ogni collegamento il comando chiede il terminale successivo e il punto sullo stesso condotto.
- Esc in una qualsiasi selezione termina il comando (sui condotti, per cambiare condotto si rilancia il comando).
- Se durante l'attivazione e' stata inserita almeno una riduzione (tubi con misura diversa, condotti con transizione all'estremita'), a fine comando compare un solo messaggio breve con il numero di collegamenti coinvolti. Se non ce ne sono, non compare nulla.

## Impostazioni (`ESA_FlexConnect`)

L'icona `i` in alto a destra della finestra mostra, al passaggio del mouse, il tooltip "How to use" (definito in `FlexConnect_Form.xaml`, come in Clash Setup). Se cambia il comportamento del tool va aggiornato anche quel testo.

I campi sono raggruppati per categoria. Il parametro "Flex Pattern" non e' piu' gestito: il flex prende il valore predefinito del suo tipo.

| Sezione | Campo | Significato |
| --- | --- | --- |
| Pipes | Flex pipe type | tipo di flexpipe da creare (ricordato per nome) |
| Ducts | Flex duct type | tipo di flexduct da creare (ricordato per nome) |
| Ducts | Duct tap fitting | raccordo di derivazione (tap, spud) usato per la presa sul fianco di un condotto; la prima voce usa le preferenze di instradamento del tipo di condotto |
| General | Reducer offset | distanza in mm della riduzione dal connettore dell'attrezzatura (tubi) o dall'estremita' del condotto (transizione); usata solo se serve |
| General | Choose the connector if several are free | se l'attrezzatura ha piu' connettori liberi candidati, apre una finestra per sceglierlo (default: il piu' vicino) |

Le chiavi di configurazione `flex_pipe_pattern` e `flex_duct_pattern` salvate dalle versioni precedenti restano nel file di configurazione e vengono ignorate.

## Logica comune

1. Si considerano i connettori liberi del dominio giusto (tubazioni o canali) dell'attrezzatura, dei terminali d'aria e dell'elemento di partenza.
2. Si sceglie il connettore dello stesso sistema (`PipeSystemType` o `DuctSystemType`) a distanza minima. Solo se non ce n'e' nessuno si accettano i connettori `Global` o `Undefined`.
   Con la spunta attiva e piu' connettori candidati una finestra li elenca ordinati per distanza con **Connector Description** (parametro della famiglia, `Connector.Description`; "(no description)" se vuoto), sistema, misura e distanza; il piu' vicino e' preselezionato, doppio clic conferma, Annulla salta quel collegamento. Anche questa finestra ha l'icona `i` con un tooltip breve.
3. Il flex prende sistema e livello dell'elemento di partenza.
4. Se l'elemento di partenza e' isolato, il flex (e il tratto corto, se c'e') riceve un isolamento dello stesso tipo e spessore. La riduzione e la presa non sono isolate.
5. Il flex parte e arriva con le tangenti dei due connettori, cosi' esce perpendicolare.

Tutto avviene in una transazione per collegamento (`FlexConnect` nel menu Annulla): se qualcosa fallisce, quel collegamento non lascia niente.

## Tubi all'estremita'

- Flex della misura dell'estremita' del tubo.
- **Stessa misura**: il flex va dall'estremita' al connettore dell'attrezzatura.
- **Misura diversa**: sull'attrezzatura si crea una riduzione (`NewTransitionFitting`), poi un tratto rigido corto della misura del tubo e lungo quanto il suo diametro, dello stesso tipo del tubo; il flex va dall'estremita' al tratto corto.

## Condotti all'estremita'

Si ottiene cliccando la faccia terminale di un'estremita' libera (clic a meno di 100 mm, lungo l'asse, dal piano del connettore). Un clic sul fianco, anche vicino all'estremita', fa la presa.

- Il flex ha **sempre la misura del connettore del terminale** (che deve essere tondo).
- **Estremita' dello stesso diametro**: il flex si collega direttamente al condotto.
- **Estremita' di forma o misura diversa** (anche rettangolare): il flex parte a "Reducer offset" mm dall'estremita' e una transizione (`NewTransitionFitting`) lo collega al condotto. Non ci sono tratti rigidi in piu'.

## Condotti: presa sul fianco

Quando il clic non cade sulla faccia terminale di un'estremita' libera (vedi sopra).

- Come all'estremita', il flex ha la **misura del connettore del terminale**. Dove c'e' la presa non serve nessuna riduzione.
- Il punto cliccato viene proiettato sull'asse del condotto (quota centrale). Lo stacco e' sul fianco verticale, perpendicolare all'asse e in orizzontale, dal lato del terminale. Il flex parte dal connettore del terminale e arriva a 50 mm dalla superficie del fianco; `NewTakeoffFitting` crea la presa sul condotto.
- L'isolamento del flex e' quello del condotto principale.
- Il raccordo della presa e' quello scelto in "Duct tap fitting". Per usarlo il comando mette la famiglia in testa alle regole Junctions del tipo di condotto, crea la presa e toglie subito la regola, nella stessa transazione: il tipo di condotto non cambia. Con la prima voce decide Revit, in base alle preferenze di instradamento.

## Limiti

- Solo elementi del modello aperto, non dei link.
- L'elemento di partenza deve avere un tipo di sistema assegnato; per l'estremita' deve avere un'estremita' libera.
- Sui condotti il connettore del terminale deve essere tondo: il flex e' tondo e non c'e' una misura da usare.
- Se estremita' e connettore distano meno di circa 15 mm il flex non si crea.
- La riduzione dipende dalle famiglie di raccordo del tipo di tubo o condotto: se le preferenze di routing non ne hanno una, Revit rifiuta il raccordo e il comando si ferma con un avviso.
- Non e' ancora stato provato in Revit: vedi il DeveloperNote.
