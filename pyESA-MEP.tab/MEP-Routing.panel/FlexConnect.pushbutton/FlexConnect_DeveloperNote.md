# FlexConnect: note di sviluppo

Registro dei tentativi, il piu' recente in alto.

## 22.10.2026, versione 0.8.0: Flex Pattern (con AI)

**Richiesta** (screenshot del parametro "Flex Pattern = Single Line"): sceglierlo nelle impostazioni, con valore diverso per tubi e condotti.

**Fatto**: `BuiltInParameter.RBS_FLEX_PATTERN_PARAM` (esiste in `RevitAPI.xml` 2026). Due combo nella finestra (`PipePatternCombo`, `DuctPatternCombo`), chiavi di config `flex_pipe_pattern` e `flex_duct_pattern` (valore intero come testo, vuoto = default del tipo). `apply_pattern` lo imposta sul flex appena creato in tutti e tre i percorsi (tubo, estremita' condotto, presa).

**Come si scoprono le voci**: l'API non le elenca. `probe_flex_patterns` crea un flex temporaneo (primo livello e primo tipo di sistema del progetto, due punti a (0,0,0) e (10,0,0)), prova i valori 0..7 con `param.Set`, legge `AsValueString` e annulla la transazione con `RollBack`. Si tengono le voci con testo non vuoto, non numerico e non ripetuto, e solo se `AsInteger` conferma il valore. Si salva il valore intero e non il nome, cosi' la scelta vale anche se Revit e' in un'altra lingua.

**Scartato**: leggere le voci da flex gia' presenti nel modello (potrebbero non esserci); scrivere i nomi a mano o con una lista fissa (dipendono dalla lingua e dalla versione di Revit).

**Non provato in Revit**: che il parametro sia di tipo intero (se e' diverso la lista resta vuota e il combo e' disabilitato, il flex prende il default del tipo); che `AsValueString` restituisca il testo delle voci; che il flex temporaneo si annulli senza lasciare tracce. Se la finestra impiega troppo ad aprirsi, la scoperta si puo' fare una volta sola e salvare l'elenco.

**Esito**: da provare.

## 22.10.2026, versione 0.7.0: condotti, misura sempre dal connettore (con AI)

**Richiesta** (screenshot dopo il test): per i condotti la misura del flex e' sempre quella del connettore del terminale; dove c'e' lo stacco takeoff non serve nessuna riduzione.

**Cosa era successo**: nello screenshot il clic era caduto vicino all'estremita' libera del condotto, quindi partiva il collegamento all'estremita' come per i tubi (tratto rigido tondo con la riduzione sul terminale, evidenziato in blu), non la presa.

**Fatto**:
- Estremita' o fianco si distinguono ora con la distanza **lungo l'asse** del clic dal piano del connettore (soglia 100 mm), non con la distanza radiale (prima 300 mm): un clic sul fianco vicino all'estremita' resta una presa; l'estremita' si sceglie cliccando la faccia terminale.
- `FlexConnect_Duct.connect_end`: flex sempre della misura del connettore del terminale. Se l'estremita' del condotto ha forma o misura diversa il flex parte a "Reducer offset" mm dall'estremita' e `NewTransitionFitting(connettore del flex, connettore del condotto)` fa la transizione; niente tratto rigido. Questo rende possibile anche l'estremita' rettangolare (prima rifiutata). Il ramo e' stato tolto da `FlexConnect_Build.connect_end`, che ora vale solo per i tubi.
- La presa non ha riduzioni, come richiesto (il flex aveva gia' la misura del terminale).

**Scartato**: tratto rigido tondo all'estremita' del condotto rettangolare, come per i tubi: servirebbe un tipo di condotto tondo scelto a caso (il tipo del condotto di partenza e' rettangolare).

**Non provato in Revit**: che `NewTransitionFitting` accetti come primo connettore il flex ancora senza tratto rigido, con tangente iniziale lungo l'asse del condotto; il comportamento con la soglia di 100 mm sulle facce terminali con flange.

**Esito**: da provare.

## 22.10.2026, versione 0.6.0: scelta del raccordo di presa (con AI)

**Problema** (screenshot): la presa usciva con la famiglia "Rectangular Air Duct Shoe Branch Slip Joint Takeoff Tap", scelta da Revit in base alle preferenze di instradamento del tipo di condotto. Richiesta: poter scegliere nelle impostazioni quale raccordo usare.

**Fatto**: nuovo combo "Duct tap fitting" (chiave di config `duct_tap`, vuota = default). L'elenco e' composto dai tipi di famiglia della categoria raccordi di canale con `PartType` TapAdjustable, TapPerpendicular, SpudAdjustable o SpudPerpendicular. `NewTakeoffFitting` non accetta un parametro per la famiglia, quindi `_with_tap_family` aggiunge una regola temporanea in testa al gruppo `Junctions` del `RoutingPreferenceManager` del tipo di condotto, crea la presa e la rimuove nel `finally`. Tutto nella transazione del collegamento. `RoutingPreferenceManager.AddRule/RemoveRule` e `RoutingPreferenceRule(ElementId, string)` verificati in `RevitAPI.xml` 2026.

**Scartato**: cambiare il tipo del raccordo dopo la creazione (`FamilyInstance.Symbol`): Revit permette solo tipi della stessa famiglia, quindi non si passa da shoe a tap circolare. Scartato anche modificare in modo permanente le regole del tipo di condotto: effetto collaterale sul modello di tutti.

**Non provato in Revit** (da verificare):
- Che `NewTakeoffFitting` rispetti la regola messa in testa. Se la `PreferredJunctionType` del tipo di condotto e' Tee potrebbe ignorare le regole di tipo tap: in quel caso impostarla temporaneamente a Tap con la stessa tecnica.
- Che `AddRule` accetti il tipo scelto per il gruppo Junctions senza criteri di misura.
- Che il filtro per `PartType` includa il raccordo cercato (famiglie con `Part Type` diverso, ad esempio generiche, non compaiono nell'elenco).
- Dopo l'aggiornamento chi ha impostazioni salvate non vede la finestra: il nuovo campo si imposta con SHIFT + CLICK.

**Esito**: da provare.

## 22.10.2026, versione 0.5.2: presa su faccia inclinata (con AI)

**Problema** (screenshot dell'utente dopo il test): la presa usciva con un raccordo rettangolare "strano", inclinato, sulla faccia inferiore del condotto. Atteso: stacchi circolari sul fianco verticale, a quota centrale rispetto al condotto grande.

**Causa**: la 0.5.0 faceva arrivare il flex lungo la componente laterale del vettore asse-terminale, quindi in 3D: con un terminale piu' basso o piu' alto del condotto la direzione era diagonale e `NewTakeoffFitting` agganciava una faccia obliqua.

**Fatto**: per un condotto non verticale la direzione dello stacco e' ora orizzontale, perpendicolare all'asse (`asse x Z`), con il segno dal lato del terminale; il punto finale e' sull'asse (quindi a quota centrale) spostato di mezza larghezza (tondo: raggio) piu' 50 mm. Per un condotto verticale resta la componente laterale, come prima.

**Non verificato**: se il raccordo risulta ancora rettangolare, la scelta della famiglia dipende dalle preferenze di instradamento del tipo di condotto (Junctions: Taps): serve una famiglia di derivazione circolare. Non si controlla dall'API di `NewTakeoffFitting`. Assunzione: larghezza del condotto rettangolare in orizzontale (condotto non ruotato).

**Esito**: da provare.

## 22.10.2026, versione 0.5.1: selezione bloccata (con AI)

**Problema**: dopo la 0.5.0 il cursore mostrava il simbolo di divieto e non si poteva selezionare niente, nemmeno con un elemento preselezionato.

**Causa probabile**: il filtro `PredicateFilter` restituiva sempre `False` da `AllowReference`, come nelle versioni 0.1-0.4 dove la selezione era `ObjectType.Element` (li' `AllowReference` non viene chiamato). Con `ObjectType.PointOnElement` Revit chiede il consenso anche sul riferimento (faccia, spigolo), quindi tutto veniva rifiutato.

**Fatto**: `AllowReference` ora restituisce `True`; il filtro sull'elemento resta in `AllowElement`.

**Non verificato in Revit**. Se il divieto resta, l'alternativa e' scegliere gli elementi con `ObjectType.Element` e chiedere il punto sul condotto con un secondo `PickPoint`, oppure un filtro che controlli l'elemento del riferimento anche in `AllowReference`.

**Esito**: da provare.

## 22.10.2026, versione 0.5.0: condotti flessibili (con AI)

**Richiesta**: stesso pulsante anche per i condotti: clic su un tratto di condotto, poi su un terminale d'aria o attrezzatura meccanica, collegamento al connettore piu' vicino con la spunta per i connettori liberi; dimensioni e isolamento come per i tubi. Caso nuovo: collegare al fianco del condotto, non solo all'estremita' (screenshot con tre derivazioni a presa). Idea dell'utente: partire dal terminale e poi cliccare sul condotto, un clic per ogni terminale.

**Fatto**:
- `FlexConnect_Build.py` ora e' generico: `Kind` (PIPE / DUCT) raccoglie classi e parametri che cambiano (Pipe/Duct, FlexPipe/FlexDuct, isolamento, parametro del diametro, parametro del tipo di sistema). `select_target` sostituisce la ricerca della coppia (vale anche per la presa, dove il lato condotto e' un punto).
- Selezione libera: `pick_elements` accetta come primo clic tubo, condotto, terminale o attrezzatura, con `PointOnElement` per avere il punto cliccato; il secondo filtro dipende dal primo. Un solo tubo/condotto preselezionato vale come primo clic (per il condotto si chiede poi un punto).
- `FlexConnect_Duct.py`: `plan_connection` decide estremita' (clic entro 300 mm da un'estremita' libera, stessa logica dei tubi) o presa; `connect_tap` crea il flex per punti con le tangenti e poi `NewTakeoffFitting(connettore finale del flex, condotto)`. Firma e requisiti verificati in `RevitAPI.xml` 2026: il proprietario del connettore deve essere un flex duct o pipe.
- Ciclo: sui condotti, dopo ogni collegamento si chiede il terminale successivo e il punto sullo stesso condotto, fino a Esc.
- Finestra: aggiunto il combo del tipo di flexduct; chiave di config `flex_duct_type`.

**Scelte discutibili da rivedere dopo il test**:
- La misura del flex nella presa e' quella del connettore del terminale, non quella del condotto: il rettangolare non ha un diametro e su un tondo la derivazione e' piu' piccola. La richiesta diceva "come per i tubi"; lo scostamento e' motivato ma va confermato.
- Soglia di 300 mm per distinguere estremita' e fianco: numero scelto a mano.
- Distanza della fine del flex dalla superficie del condotto: 50 mm piu' mezza sezione (raggio, o mezzo lato maggiore per i rettangolari). La presa dovrebbe assorbire l'errore, ma non e' verificato.

**Non provato in Revit**: tutto il ramo condotti. In particolare che `NewTakeoffFitting` accetti il connettore finale di un flex non ancora collegato e la posizione in cui mette la presa; il comportamento delle tangenti con i terminali a soffitto (`BasisZ` verso il basso); `DuctInsulation.Create` su un flexduct.

**Scartato**: estremita' di condotto rettangolare con flex tondo. Servirebbe una transizione rettangolare-tonda subito all'estremita', senza spazio per il raccordo: il comando avvisa e rimanda alla presa.

**Esito**: da provare. Controllata solo la sintassi.

## 22.10.2026, versione 0.4.0: tubo corto lungo quanto il DN (con AI)

**Richiesta** (dopo il test, con screenshot): il tubo tra riduzione e flex era troppo lungo (100 mm fissi); ne basta uno minimo, pari al DN del tubo.

**Fatto**: la lunghezza del tubo corto e' ora il diametro del tubo di partenza (`pipe_diameter`), al posto della costante `STUB_LENGTH_MM`, rimossa. Il campo "Reducer offset" resta com'e': regola la distanza della riduzione dall'attrezzatura, non la lunghezza del tubo corto. Se il tratto risulta ancora lungo, il valore da ridurre e' quell'offset (default 150 mm).

**Non provato in Revit**: che Revit accetti un tubo cosi' corto su tutte le misure (la lunghezza minima dipende dalle preferenze di routing e dal raccordo).

**Esito**: da provare.

## 22.10.2026, versione 0.3.0: scelta del connettore (con AI)

**Richiesta**: spunta nella finestra delle opzioni per aprire una finestra di selezione quando ci sono piu' connettori dello stesso tipo.

**Fatto**: `find_connector_pair` costruisce un candidato per connettore dell'attrezzatura (con l'estremita' del tubo piu' vicina), nella stessa fascia gia' usata (sistema esatto, altrimenti jolly). Con la spunta e piu' di un candidato chiama un `chooser` passato dallo script, che apre `FlexConnect_Connector.xaml` (lista ordinata per distanza, primo preselezionato). Annulla solleva `FlexConnectCancelled`, che lo script trasforma in uscita silenziosa. La spunta e' salvata in `ESA_FlexConnect` come `ask_connector` ("1"/"0": evita ambiguita' di tipo nella lettura della config); le config vecchie senza la chiave valgono "spenta".

**Limite noto**: la finestra non evidenzia il connettore in Revit, quindi l'utente si orienta con id, misura e distanza. Se serve, si potrebbe selezionare/zoomare l'attrezzatura o mostrare un'anteprima: non fatto, per tenere la finestra semplice.

**Esito**: da provare in Revit. Controllata solo la sintassi.

## 22.10.2026, versione 0.2.0: isolamento e finestra solo con SHIFT + CLICK (con AI)

**Richiesta** (dopo il primo test, "funziona abbastanza bene"): copiare l'isolamento del tubo di partenza; non aprire la finestra a ogni uso, solo con SHIFT + CLICK.

**Fatto**:
- Isolamento: `InsulationLiningBase.GetInsulationIds(doc, pipe.Id)` sul tubo di partenza, poi `PipeInsulation.Create(doc, id, tipo, spessore)` su flex e tubo corto, dopo i `ConnectTo`. Firme verificate in `RevitAPI.xml` 2022 e 2026. Si copia il primo isolamento trovato. La riduzione non e' isolata. Un errore dell'isolamento annulla tutta la transazione (scelta: meglio niente che un collegamento a meta').
- Finestra: logica invertita rispetto alla regola comune di CLAUDE.md. Resta il primo uso senza impostazioni valide, dove la finestra si apre comunque. Documentato come eccezione in tooltip e Readme.

**Non provato in Revit**: che `PipeInsulation.Create` accetti un flexpipe come host (la documentazione non lo esclude) e il comportamento con tubi senza isolamento.

**Esito**: da provare.

## 22.10.2026, versione 0.1.0: prima stesura (con AI)

**Richiesta**: da un tubo selezionato, scegliere un'attrezzatura meccanica o idraulica e creare un collegamento in flexpipe con il connettore dello stesso tipo di sistema piu' vicino, ereditando la misura del tubo e mettendo una riduzione se serve.

**Verificato offline** (documentazione `RevitAPI.xml` 2022 e 2026):
- `FlexPipe.Create(doc, systemTypeId, flexTypeId, levelId, IList<XYZ>)` e l'overload con `startTangent` / `endTangent` esistono in entrambe le versioni (dal 2014).
- `Document.NewTransitionFitting(c1, c2)`: il proprietario di `c1` deve essere un tubo (o flex): per questo il primo connettore e' sempre quello del tubo corto, mai quello dell'attrezzatura.
- `Connector.PipeSystemType` esiste.

**Scelte**:
- Pulsante autonomo nel pannello MEP-Routing: lo stack `Connect` ha gia' 3 pulsanti (massimo).
- Flex creato per punti, non per connettori: non esiste un overload di `FlexPipe.Create` con connettori, quindi i due `ConnectTo` si fanno dopo aver impostato il diametro e rigenerato.
- Riduzione tra attrezzatura e un tubo corto, non tra flex e attrezzatura: `NewTransitionFitting` vuole due connettori coassiali e affacciati, cosa che un flex libero non garantisce.
- Sistema: prima corrispondenza esatta, poi jolly (`Global` / `Undefined`), altrimenti errore.

**Non provato in Revit** (da verificare al primo test):
- Segno delle tangenti: si assume che `StartTangent` sia la direzione di uscita dal connettore del tubo e `EndTangent` quella di ingresso nell'attrezzatura (quindi `-BasisZ`). Se il flex esce con anelli, invertire o togliere le tangenti (c'e' gia' il fallback senza tangenti solo in caso di eccezione).
- Che `NewTransitionFitting` accetti il connettore di un'attrezzatura come secondo connettore.
- Che `ConnectTo` tra flex e connettore di tubo/attrezzatura non protesti per piccole differenze di misura dopo `Regenerate`.
- Revit consigliato per il primo test: 2026 e almeno una versione piu' vecchia (2022).

**Esito**: da provare. Nessun tentativo scartato finora.
