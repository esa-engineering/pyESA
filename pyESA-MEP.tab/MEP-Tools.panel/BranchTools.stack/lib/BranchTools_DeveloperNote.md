# Branch Tools: note di sviluppo

Registro dei tentativi, il piu' recente in alto. Fino alla 0.13.0 il tool era un solo pulsante, Fitting Tool (`MEP-Tools.panel/FittingTool.pushbutton`, file `FittingTool_*`): le voci precedenti alla 0.14.0 usano quei nomi.

## 23.10.2026, versione 0.14.0: divisione in tre comandi, stack Branch Tools (con AI)

**Richiesta** (screenshot della finestra con tre riquadri: sezione Pipe; Rotate + Branch angle + Align; Side branch + Branch length): dividere il tool in 3 comandi separati in uno stack "Branch Tools"; nella scelta della famiglia lasciare solo raccordi a croce o a T.

**Fatto**:
- `MEP-Tools.panel/BranchTools.stack` con `PipeWye`, `FittingRotate`, `SideBranch` (ordine nel `bundle.yaml` dello stack); nel layout del pannello `BranchTools` al posto di `FittingTool`. Cartella `FittingTool.pushbutton` eliminata (mai committata: nessuna perdita in git).
- `lib`: i moduli di Fitting Tool rinominati `BranchTools_*` (Rocket mode: nomi unici) con gli import aggiornati; nuovi `BranchTools_Controls` (SliderSection, set_status, on_enter_or_leave, parse_number, step_text, prima nella UI di Fitting Tool) e `BranchTools_Select` (selezione di un solo elemento con un predicato, load/save delle impostazioni numeriche e di testo). Readme e DeveloperNote di Fitting Tool diventano quelli di `lib`.
- Pipe Wye: la braga da inserire non prende piu' rotazione e angolo dagli slider (che stanno in Fitting Rotate) ma da due campi; dopo Place wye la finestra non passa piu' alla braga nuova (era `_bind` + `factory`, tolti): lo stacco si fa con Side Branch.
- Side Branch: con una pendenza scelta la braga si ruota con `Rotator.rotate_to` direttamente (prima passava dallo slider Rotate per mostrare il valore).
- Fitting Rotate: selezione di un raccordo o di un tubo (per Align); con un tubo Rotate e Branch angle sono spenti con "Not available: select a fitting.".
- Lista Wye: solo `FamilySymbol` di `OST_PipeFitting` con Part Type Tee, Cross, Wye, LateralTee o LateralCross (`BRANCH_PART_TYPES`). Wye e i Lateral* aggiunti alla richiesta "croce o T": sono braghe a tutti gli effetti.
- Una sezione di configurazione per pulsante (`ESA_PipeWye`, `ESA_FittingRotate`, `ESA_SideBranch`, stesse chiavi di prima); `ESA_FittingTool` non piu' letto.
- Icone: Fitting Rotate riusa quella di Fitting Tool; Pipe Wye e Side Branch disegnate con System.Drawing (unione delle forme riempita, poi interno svuotato: contorno senza riempimento), chiara 40 e scura 230.
- Finestre misurate con `Measure` in WPF, tutte le righe di stato visibili: Pipe Wye 262 su 300, Fitting Rotate 563 (finestra portata a 580), Side Branch 401 su 420.

**Verifiche fuori da Revit**: sintassi IronPython di tutti i file; nomi dei controlli XAML usati dal codice; controllo statico in CPython dei nomi non definiti e dei riferimenti `alias.nome` ai moduli di `lib` (nessuno mancante).

**Non provato in Revit**: tutto. In particolare che pyRevit metta `BranchTools.stack/lib` nel percorso dei tre pulsanti (come per `ClashFinder.pulldown/lib`).

**Esito**: da provare in Revit.

## 23.10.2026, versione 0.13.0: scelta della famiglia della braga (con AI)

**Richiesta** (due screenshot): tutte le opzioni della sezione Pipe dicono "No wye family (Part Type Wye) is loaded in the project"; "fammela scegliere se possibile".

**Causa**: nel progetto nessuna famiglia di raccordo di tubazione ha Part Type Wye (`FAMILY_CONTENT_PART_TYPE`): la braga ESA `e_PP.GG.JN_Y Branch` e' probabilmente definita come Tee o Undefined. Non verificato quale.

**Fatto**:
- Sezione Pipe: combo `WyeFamilyCombo` (modificabile, ricerca per testo) sulla stessa riga dei pulsanti, per restare a 700 di altezza: prima voce `(routing preferences)`, poi tutti i `FamilySymbol` di `OST_PipeFitting` come "Famiglia : Tipo" (`pipe_fitting_symbols`, `symbol_label`). Scelta ricordata per etichetta in `wye_family` (prima impostazione di testo: `TEXT_OPTIONS` nello script, le altre restano numeri). Default alla prima apertura: una famiglia con Part Type Wye se c'e', altrimenti routing preferences.
- `MainPipe.symbol_id`: con una famiglia scelta `insert_junction(..., symbol)` la inserisce e orienta direttamente, senza routing preferences; con `(routing preferences)` resta il percorso della 0.12.2 (routing preferences, poi Part Type Wye) e l'errore finale chiede di scegliere la famiglia.
- Una famiglia scelta che non ha collettore e derivazione riconoscibili (`build_frame` senza `main_ids` o `branch_id`): errore "is not a wye".

**Non provato in Revit**: tutto il percorso con la famiglia scelta (vedi le note della 0.12.2).

**Esito**: sintassi verificata, da provare in Revit.

## 23.10.2026, versione 0.12.2: braga inserita come famiglia (con AI)

**Richiesta** (screenshot): Connect from above ancora rifiutato con lo stesso messaggio dopo la 0.12.1: "Revit could not place the junction at this angle (Fitting cannot be created between the input connectors because the angle between them is too small or too large)".

**Esito della 0.12.1**: **fallisce**. Provare i due ordini dei connettori del collettore in `NewTeeFitting` non basta: Revit rifiuta la derivazione a 45 gradi in entrambi. Causa non verificata: la giunzione delle routing preferences del tipo di tubo degli scarichi non e' una braga, oppure `NewTeeFitting` non usa la famiglia a Y a 45 gradi. Non riprovare la strada dell'ordine dei connettori.

**Fatto**:
- `insert_junction(doc, main, point, flow, branch_dir, angle, branch_diameter)` in due tentativi. Primo: `_junction_from_preferences`, la strada della 0.12.1 dentro una `SubTransaction` (se fallisce si annulla e il modello torna com'era). Secondo: `_junction_from_family`:
  - tipo con `wye_symbol()`: giunzione delle routing preferences (`RoutingPreferenceManager.GetRule(Junctions, i).MEPPartId`) se Part Type Wye (`FAMILY_CONTENT_PART_TYPE`), altrimenti la prima famiglia di raccordo del progetto con Part Type Wye;
  - `NewFamilyInstance(point, symbol, NonStructural)`, Angle impostato, diametri con `Connector.Radius` (ignorato se la misura e' guidata da una formula);
  - orientamento: rotazione che porta l'asse del collettore della famiglia (nel verso per cui la derivazione sta verso monte) sul verso di flusso, poi rotazione attorno al flusso finche' la derivazione e' in `branch_dir`, poi spostamento del centro (punto dell'asse del collettore piu' vicino all'asse della derivazione) sul punto;
  - `FittingTool_Place.split_and_connect()` (estratta da `BranchCopier.place`, stessa logica che funziona in Place copies) spezza il collettore e collega.
- Connect from above / from the side: la braga ora ha la derivazione libera in entrambi i casi; il tubo nuovo parte dal connettore della derivazione (trovato con `build_frame`) fino all'angolo, `ConnectTo` sulla braga e curva sull'altro tubo.

**Verifiche fuori da Revit**: sintassi; test numerico in CPython (`orient_test.py` nella cartella temporanea della sessione) dell'orientamento, con famiglia con origine fuori dal centro e derivazione da un lato o dall'altro, collettore piano e in pendenza, due versi, rotazioni 0/45/90/135: 64 casi su 64 con asse, derivazione e centro esatti.

**Non provato in Revit**: `Connector.Radius` sulla famiglia `e_PP.GG.JN_Y Branch` (NR1 / NR2); `NewFamilyInstance` senza livello per un raccordo di tubazione; che `RotateElement` su un raccordo non collegato non sia vincolato al piano di lavoro.

**Esito**: sintassi e orientamento verificati, da provare in Revit.

## 23.10.2026, versione 0.12.1: Place wye rifiutato da Revit (con AI)

**Richiesta** (screenshot): Place wye non funziona, riga di stato "Not placed: Fitting cannot be created between the input connectors because the angle between them is too small or too large. Flow: downhill."

**Causa probabile (non verificata)**: `NewTeeFitting(main1, main2, branch)` misura l'angolo della derivazione rispetto a uno dei due connettori del collettore. La derivazione e' a 45 gradi verso monte: rispetto a un connettore sono 45, rispetto all'altro 135. L'ordine dei due connettori veniva da `BreakCurve` (tubo originale, poi pezzo nuovo), quindi dal verso in cui era disegnato il tubo, non dal verso di flusso.

**Fatto**: `insert_junction` prova `NewTeeFitting(first, second, branch)` e, se Revit solleva un'eccezione, `NewTeeFitting(second, first, branch)` nella stessa transazione. Se falliscono entrambi, `PipeError` con il messaggio di Revit e l'indicazione di controllare che la giunzione delle routing preferences sia una braga. Vale anche per Connect from above / from the side, che usano la stessa funzione.

**Se fallisce ancora**: le routing preferences del tipo di tubo hanno probabilmente un T a 90 gradi o una presa come giunzione. Strada alternativa da provare: inserire direttamente il `FamilySymbol` della braga (`NewFamilyInstance`), orientarlo e collegare i connettori a mano, senza passare dalle routing preferences.

**Esito**: sintassi verificata, da provare in Revit.

## 23.10.2026, versione 0.12.0: selezione singola, sezione Pipe, Align sull'elemento cliccato (con AI)

**Richiesta**: in Align scegliere a quale tubo o raccordo allinearsi; togliere la selezione multipla iniziale; con un tubo selezionato, una sezione in alto per: posizionare un raccordo a 45 con le impostazioni delle sezioni sotto; collegare un altro tubo "da sopra calando poi a 45 verso la direzione di discesa" o "di lato a 45 ma scendendo con l'inclinazione necessaria a non modificare la quota del collettore".

**Scelte dell'utente** (domande prima del codice): il raccordo a 45 e' una braga inserita nel tubo nel punto cliccato, con rotazione e angolo dalle sezioni Rotate e Branch angle, e poi la finestra lavora sulla braga nuova (scartato: curva a 45 all'estremita' libera); verso valle = discesa se il collettore e' in pendenza, altrimenti il verso del tubo con la casella Reverse flow (scartato: casella sempre).

**Fatto**:
- Selezione: un solo elemento (`get_element`): la preselezione se e' esattamente uno, altrimenti `PickObject(ObjectType.Element)`; il `GlobalPoint` del clic e' il punto di inserimento della braga (preselezione: mezzeria). Rotate e Branch angle non gestiscono piu' piu' raccordi insieme.
- Azioni che richiedono un clic nel modello (Connect from above / side, Place copies, Align and connect): la finestra si chiude come OK con `action` nel risultato e lo script esegue dopo la chiusura (`run_action`), un solo `forms.alert` se qualcosa non riesce. Stesso schema di Place copies: con la finestra modale aperta non si puo' cliccare nel modello. Scartato: nascondere la finestra e riaprirla, perche' lo stato degli slider e il `TransactionGroup` aperto durante la selezione erano incerti.
- `FittingTool_Pipe.py` (nuovo), `MainPipe`:
  - `directions()`: f verso valle, U alto perpendicolare a f, R = f x U destra guardando verso valle;
  - `place_wye(roll, angle)`: derivazione `-f cos(angle) + (R cos(roll) + U sin(roll)) sin(angle)`; tubo provvisorio dal punto in quella direzione, `PlumbingUtils.BreakCurve` del collettore nel punto, `NewTeeFitting` (braga dalle routing preferences), tubo provvisorio eliminato;
  - `connect(other, point, from_above)`: angolo C = estremita' libera di other (di lato) oppure il punto dell'asse di other che in pianta sta sull'asse del collettore (da sopra, sistema 2x2 in XY; other allungato o accorciato fino a C). Braga a valle della proiezione di C di quanto C dista dall'asse: l'angolo tra derivazione e collettore e' 45 esatti in 3D, in qualunque posizione e con qualunque pendenza. Tubo nuovo da braga a C con tipo e diametro di other, `NewTeeFitting`, `NewElbowFitting` (o `ConnectTo` se gia' in linea).
- `build_frame` / `Rotator` con `zero_side` facoltativo: per la braga appena inserita lo 0 della rotazione resta a destra guardando verso valle (come in Place wye), invece che sul lato della derivazione.
- `Aligner(element)` e `run_with(target)`: candidati solo l'elemento cliccato, senza limite di distanza; senza target resta la ricerca entro 2 m (non piu' usata dalla finestra).
- UI: `_bind(fittings, zero_side)` ricollega Rotate, Branch angle, Side branch e Branch length (gli oggetti li crea `factory` dello script); in modalita' tubo gli slider tengono solo i valori della braga (`wye_roll`, `wye_angle` in config). Sezione Pipe in alto su tutta la larghezza; passi sulla stessa riga dei valori per recuperare altezza. Finestra 860 x 700: contenuto misurato 654 con tutte le righe di stato visibili, piu' 39 di cornice.

**Verifiche fuori da Revit**: sintassi IronPython; nomi dei controlli; test numerico in CPython (`connect_test.py` nella cartella temporanea della sessione) per collettore piano e in pendenza del 2%, da sopra e di lato, alla stessa quota e piu' in alto: angolo con il collettore 45,0000, braga sull'asse e a valle.

**Non provato in Revit**: `NewTeeFitting` con una derivazione a 45 (se le routing preferences hanno un T a 90 o una presa, Revit puo' rifiutare); `PickObject(ObjectType.Element).GlobalPoint` per il punto della braga (se fosse vuoto si usa la mezzeria solo con la preselezione: da verificare); `NewElbowFitting` con angoli non commerciali; collegamento "da sopra" con l'altro tubo inclinato.

**Esito**: sintassi e geometria verificate, da provare in Revit.

## 23.10.2026, versione 0.11.0: Align collega le estremita' (con AI)

**Richiesta**: "Si, fallo attaccare in automatico" (risposta alla proposta di collegare le estremita' dopo l'allineamento).

**Fatto**:
- `Aligner.connect()` dopo `align()`, in una seconda `run_step` "Connect Ends": cosi' se il collegamento fallisce l'allineamento resta (scelta: con un solo passo un errore di collegamento avrebbe annullato anche l'allineamento).
- Chiusura della distanza lungo l'asse: tubo selezionato allungato (`stretch`: nuova `Line` con l'estremita' sul connettore spostata), altrimenti tubo dell'altro elemento allungato, altrimenti (due raccordi) `MoveElement` del raccordo selezionato della distanza. Poi `ConnectTo`.
- Misure diverse (`same_size`: forma, raggio o larghezza e altezza): tra due `MEPCurve` con distanza > 0 `NewTransitionFitting` (riduzione dalle routing preferences); altrimenti errore "different sizes", nessun collegamento.
- La coppia trovata da `find()` e' salvata per Id (`pair`: connettore mio, elemento e connettore dell'altro) e riletta con `_pair_connectors()`: dopo lo spostamento gli oggetti `Connector` precedenti non vanno usati.
- `can_run()`: il pulsante resta attivo anche con le estremita' gia' in asse (solo da collegare). Pulsante rinominato `Align and connect`; How to use aggiornato.

**Non provato in Revit**: `ConnectTo` tra tubo e raccordo dopo `stretch`; `NewTransitionFitting` con tubi in asse ma distanti (Revit potrebbe chiedere una distanza minima); `MoveElement` di un raccordo collegato dall'altra parte (trascina i tubi collegati).

**Esito**: sintassi verificata, da provare in Revit.

## 23.10.2026, versione 0.10.0: Align e copie che non tengono la quota (con AI)

**Richiesta** (screenshot della finestra con lo spazio sotto Branch angle indicato e "ALIGN" scritto dentro, e lo spazio vuoto in fondo indicato): "Il place copies si rompe di nuovo su alcuni tratti dopo che modifico la lunghezza"; nuova sezione Allinea: selezionato un tubo o un raccordo, allineare le due estremita' aperte piu' vicine; metterla nello spazio indicato e togliere lo spazio vuoto.

**Place copies, causa probabile (testo dell'avviso non riportato, non verificata)**: con Copies keep this elevation una copia a monte sul collettore in pendenza nasce piu' in alto e il tubo inclinato deve accorciarsi di `drop / d.Z`. Dopo aver accorciato lo stacco con Branch length (L = 224,6 mm nello screenshot) sui tratti lontani a monte il tubo non basta e `slide_branch` alzava "The sloped pipe would be too short", annullando la copia: coerente con "su alcuni tratti" e "dopo che modifico la lunghezza". E' un limite geometrico; corretto il messaggio, ora in `_keep_level` prima dello scorrimento: di quanti mm andrebbe accorciato, quanto e' lungo, e i tre rimedi. Se il testo dell'avviso dell'utente e' un altro, la causa e' un'altra: chiedere il testo (con la fase tra parentesi).

**Align, fatto**:
- `FittingTool_Align.py` (nuovo): `Aligner` con un solo elemento selezionato (`MEPCurve` o raccordo). Estremita' libere (`ConnectorType.End`, non collegate) dell'elemento; candidati con `ElementMulticategoryFilter` (tubi, canali, passerelle, tubi protettivi, i loro raccordi, accessori di tubi e canali) AND `BoundingBoxIntersectsFilter` sul riquadro dell'elemento allargato di 2 m. Coppia piu' vicina tra estremita' dello stesso `Domain`, parallele (`|n1 x n2| < 1e-3`) e rivolte una verso l'altra (`n1 . n2 < 0` e l'altra davanti alla mia). Spostamento = componente di (o2 - o1) perpendicolare all'asse dell'altra estremita'; `MoveElement` dell'elemento selezionato in una `run_step` "Align Ends". Dopo lo spostamento la ricerca si rifa' e la riga dice che sono allineate.
- Scelta: si sposta l'elemento selezionato (quello che l'utente indica), non l'altro; nessun collegamento automatico, la richiesta era allineare. Possibile aggiunta: casella "connect" se le estremita' coincidono dopo l'allineamento.
- Selezione iniziale: `get_elements()` accetta anche i `MEPCurve` (`is_alignable`); `Rotator`, `AngleSetter`, `SideBranch` ricevono solo i raccordi.
- Finestra: sezione Align sotto Branch angle (colonna sinistra), voce ALIGN nel How to use. Altezza misurata con `Measure` in WPF: il contenuto richiede 587 (anche con tutte le righe di stato), piu' circa 39 di cornice: 640 resta, `MinHeight` 630. Lo spazio vuoto dello screenshot era la colonna sinistra piu' corta della destra; ora la sezione Align la riempie.
- Nota per il futuro: alla 0.5.0 avevo attribuito il taglio in Revit a testi piu' grandi; con la misura di oggi sembra piuttosto che l'altezza della finestra non contasse la cornice. Misurare con `Measure` e aggiungere la cornice.

**Non provato in Revit**: Align con un tubo collegato a un raccordo all'altra estremita' (Revit trascina il raccordo?); `BoundingBoxIntersectsFilter` con `Outline` allargato; estremita' di dominio diverso (es. tubo e accessorio con connettore di dominio Undefined).

**Esito**: sintassi verificata, da provare in Revit.

## 23.10.2026, versione 0.9.0: Branch length, How to use, meno testo grigio (con AI)

**Richiesta** (due screenshot: la finestra con lo spazio vuoto sotto Side branch evidenziato e un testo grigio barrato; uno stacco con le quote L e H): selezionando una braga con uno stacco gia' collegato, cambiare con uno slider la lunghezza L del tubo inclinato spostando di quota (H) il tratto collegato, perche' farlo dai parametri sposta anche il collettore. Aggiungere il pulsante "How to use" e ridurre al minimo le scritte grigie, come in Clash Setup.

**Fatto**:
- `FittingTool_Place.slide_branch()`: il movimento di `_keep_level` (curva e tratti successivi spostati lungo la derivazione, tubo inclinato allungato, ripiego se Revit non lo allunga) estratto in una funzione, usata da `_keep_level` e dalla nuova sezione.
- `FittingTool_Length.py` (nuovo): `BranchLength` con `load()` (braga adatta, derivazione collegata a un `Pipe`, catena con `branch_chain`), `value()` = lunghezza del tubo inclinato, `offset()` = H (fondo interno del primo tubo dopo la curva, quello con l'estremita' piu' vicina all'angolo, meno fondo interno del collettore), `length_for_offset()` = `L + (H' - H) / d.Z`, `set_to()` in una `run_step` "Set Branch Length". Si ricarica dopo `Create branch`.
- UI: sezione Branch length sotto Side branch (lo spazio indicato nello screenshot), controlli con prefisso `Sloped` perche' `LengthBox` esiste gia' (lunghezza orizzontale). Casella `HeightBox` per H. `SliderSection.reset()` per riattivare lo slider dopo Create branch; `on_enter_or_leave()` comune a caselle e slider.
- Testi: icona `i` con tooltip How to use (stesso `Border` di FlexConnect e Clash Setup) con le spiegazioni di tutte le sezioni; tolti i testi grigi fissi ("Angle of the branch around the run...", "Cancel puts the fittings back..."); le righe di stato sono `Collapsed` di default e compaiono con `set_status()` solo se c'e' un testo (non rotabili, Angle mancante o diverso, motivo di una sezione spenta, curva calcolata, errori). Tolto "Now: ...": il valore e' gia' nella casella.
- Finestra 860 x 640: nel rendering WPF con tutte le righe di stato visibili il contenuto occupa 531 su circa 555; in Revit i testi occupano di piu' (vedi 0.5.0), quindi margine.

**Non provato in Revit**: lo scorrimento con una curva fuori piano (tratto perpendicolare); H su uno stacco fatto a mano con piu' tubi dopo la curva (si prende il tubo con l'estremita' piu' vicina alla curva).

**Esito**: sintassi verificata, da provare in Revit.

## 23.10.2026, versione 0.8.0: le copie tengono la quota dell'originale (con AI)

**Richiesta** (screenshot di tre copie con il tratto finale a quote diverse, "Lower End Bottom Elevation `<varies>`"): Place copies funziona, ma le copie hanno tutte altezze diverse; mantenere la quota dello stacco copiato.

**Causa**: voluta dalla 0.7.0. La copia e' una traslazione lungo l'asse del collettore, che e' in pendenza: la braga copiata scende con il collettore e lo stacco tiene l'offset da esso, non la quota assoluta.

**Fatto**: `BranchCopier._keep_level()`. La braga copiata sta `drop = shift.Z` piu' in alto o in basso; tutti gli elementi copiati dopo il primo tubo dello stacco (curva, tratto orizzontale, eventuali altri, solo quelli con connettori: niente isolamento) sono spostati con `MoveElements` lungo la direzione d della derivazione di `-drop / d.Z`, cosi' la loro quota torna quella dell'originale e il tubo inclinato cambia solo lunghezza. Dopo `Regenerate`, se Revit non ha allungato il tubo inclinato insieme alla curva, lo si fa con `Location.Curve = Line(...)` e si ricollega il connettore coincidente. Errori: derivazione orizzontale (d.Z circa 0), stacco che non comincia con un `Pipe`, tubo inclinato piu' corto di 0,05 piedi.
- Scartato: tenere la braga sulla traslazione e spostare tutto lo stacco in verticale. Il tubo inclinato cambierebbe pendenza e la curva angolo, quindi la copia non sarebbe piu' uguale all'originale.
- Casella `Copies keep this elevation` (predefinita spuntata, config `copy_keep_level`); tolta, il comportamento della 0.7.
- Nuova fase `level` nei messaggi d'errore.

**Non provato in Revit**: che `MoveElements` di curva e tratto orizzontale allunghi il tubo inclinato collegato (c'e' il ripiego) e che i collegamenti restino.

**Esito**: sintassi verificata, da provare in Revit.

## 23.10.2026, versione 0.7.1: Place copies falliva sempre (con AI)

**Richiesta** (screenshot): copia di uno stacco in un punto del collettore, avviso "1 branch(es) not placed, 0 placed. The referenced object is not valid, possibly because it has been deleted from the database, or its creation was undone." Stesso errore in tutti i punti provati.

**Causa**: in `BranchCopier.place()`, dopo `doc.Delete(middle[0].Id)` la lista degli spezzoni veniva filtrata con `p.Id != middle[0].Id`: leggere `Id` di un elemento eliminato solleva `InvalidObjectException`. L'errore arrivava dopo copia e tagli, quindi a ogni clic; `run_step` annullava la transazione e non restava nulla nel modello.

**Fatto**: gli spezzoni sono gestiti per `ElementId` e riletti con `doc.GetElement`; lo spezzone eliminato e' escluso confrontando il valore dell'Id salvato prima di `Delete`; scartati gli oggetti non piu' validi (`IsValidObject`), anche in `_contains` / `_has_end` nel caso `BreakCurve` sostituisca il tubo originale. Regola: dopo `doc.Delete` non toccare piu' l'oggetto Python dell'elemento eliminato.
- Diagnostica: `BranchCopier.stage` (check, copy, split, connect) aggiunta al motivo di ogni copia non riuscita, cosi' l'avviso dice in che fase si e' fermata.

**Esito**: sintassi verificata, da riprovare in Revit.

## 23.10.2026, versione 0.7.0: Place copies (con AI)

**Richiesta**: "un pulsante copy o place, che mi fa cliccare sul tratto dello stesso collettore e mi fa creare nuovi stacchi uguali".

**Fatto**:
- `FittingTool_Place.py` (nuovo). `branch_chain()` raccoglie gli elementi collegati alla derivazione (visita dei connettori fisici, braga esclusa, massimo 30: oltre, la derivazione porta a una rete). `BranchCopier.place()`: `ElementTransformUtils.CopyElements` di braga + stacco con una traslazione lungo l'asse del collettore (centro della braga = proiezione del punto cliccato), braga copiata riconosciuta dalla `LocationPoint` traslata, collegamenti braga-stacco rifatti sui connettori coincidenti se la copia non li ha tenuti; poi `PlumbingUtils.BreakCurve` alle due estremita' della braga, eliminato lo spezzone con entrambe le estremita' sulla braga, `ConnectTo` delle due parti. Controllo preventivo: le due estremita' della nuova braga devono stare dentro il tubo cliccato con almeno 0,01 piedi di margine.
- Scelta: copia e traslazione invece di inserire una braga nuova dalle routing preferences e rifare il calcolo. Cosi' la copia ha esattamente tipo, Angle, rotazione e stacco dell'originale, anche se fatto a mano, e non dipende dalle routing preferences.
- `MainPipeFilter`: solo `Pipe` con entrambe le estremita' sull'asse del collettore (tolleranza circa 3 mm), cosi' si possono cliccare anche gli spezzoni lasciati dalle copie precedenti. `AllowReference` True, necessario per `ObjectType.PointOnElement`.
- `build_frame` ora calcola asse e riferimenti anche quando il raccordo non e' ruotabile (pinned o collegato fuori asse) e salva gli Id dei connettori del collettore (`main_ids`): una braga con lo stacco collegato non e' ruotabile ma va copiata.
- `SideBranch`: `reason` (braga adatta: Y o T su un tubo), `create_reason` (derivazione libera e braga ruotabile), `place_reason()` (derivazione collegata). Prima un'unica `reason` escludeva la braga con la derivazione collegata.
- UI: pulsante `Place copies` accanto a `Create branch`, attivo se c'e' gia' uno stacco o dopo averlo creato. Chiude la finestra con il risultato di OK piu' `place`: la selezione dei punti non e' possibile con la finestra modale aperta. Lo script assimila il gruppo e poi esegue il ciclo dei clic; ogni copia e' una `run_step` (`Place Side Branch`), quindi una voce del menu Annulla per copia. Un solo `forms.alert` a fine comando solo se qualche copia non e' riuscita.

**Non provato in Revit**: che `CopyElements` mantenga i collegamenti tra gli elementi copiati insieme (c'e' il ripiego `_connect_copy`) e copi anche l'isolamento dei tubi; `BreakCurve` su un collettore in pendenza e su un tubo gia' collegato ad altri raccordi alle estremita'; `ConnectTo` tra spezzoni e braga copiata; clic su un punto vicino a un'altra braga (il controllo di margine guarda solo le estremita' del tubo cliccato).

**Esito**: sintassi verificata, da provare in Revit.

## 23.10.2026, versione 0.6.0: tratto orizzontale perpendicolare al collettore (con AI)

**Richiesta** (pianta con lo stacco: braga a 45°, tubo che prosegue a 45° barrato, freccia perpendicolare al collettore): "puoi mettere che la curva vada ortogonale dal collettore?".

**Fatto**:
- `SideBranch.run_direction()`: tratto orizzontale perpendicolare al collettore in pianta (`Z x asse`, verso scelto dal lato in cui va la derivazione) oppure come la derivazione (comportamento fino alla 0.5). Errore se la derivazione corre lungo il collettore (componente laterale < 0,1% della componente in pianta).
- La curva sta sempre all'intersezione degli assi (tubo inclinato e tratto orizzontale partono dallo stesso punto), quindi la quota resta esatta anche con il tratto perpendicolare; l'angolo della curva e' quello tra le due direzioni (`elbow_angle()`), non piu' uguale alla pendenza.
- UI: nuova combo Horizontal run (predefinito Perpendicular to main, config `branch_direction`). Il campo Elbow diventa Branch slope (`SlopeCombo`): il valore era gia' la pendenza del tubo inclinato, che con il tratto perpendicolare non coincide piu' con la curva. La chiave di config resta `branch_elbow_mode` per non perdere la scelta salvata. Riga di stato: pendenza e curva calcolate dalla posizione attuale.
- Messaggi di `target_angle` parlano di slope invece che di elbow.

**Non provato in Revit**: `NewElbowFitting` con angoli fuori piano (tubo inclinato e tratto orizzontale non complanari con il collettore) e angoli non commerciali (es. 60°) con le famiglie di curva degli scarichi.

**Esito**: sintassi verificata, da provare in Revit.

## 23.10.2026, versione 0.5.0: lo stacco tiene la rotazione impostata (con AI)

**Richiesta** (due screenshot): braga impostata con rotazione 45° e Angle 45°, Elbow 45; dopo `Create branch` la rotazione e' diventata 78,58° e il tratto orizzontale e' venuto parallelo al collettore. "Incoerenza nella creazione delle branch quando inclinate."

**Causa**: non un errore di calcolo ma di interpretazione. La 0.4.0, con un angolo di curva scelto, ricalcola la rotazione perche' la derivazione salga di quell'angolo: Y 45° + curva 45° richiede rotazione 90°, e con un collettore in pendenza del 2% `(1 - a.Z) / u.Z` da' sin(phi) = 0,980, cioe' 78,58°; con la derivazione nel piano verticale del collettore il tratto orizzontale e' parallelo al collettore. L'utente invece si aspettava che la rotazione impostata restasse: era anche la sua prima risposta ("con le impostazioni di rotazioni impostate"), che nella 0.4.0 avevo letto come "partendo dalla braga gia' sul collettore".

**Fatto**:
- Elbow ha la voce `From rotation`, predefinita: nessuna rotazione, la curva prende l'inclinazione attuale della derivazione (`SideBranch.elevation()`, asin di BasisZ.Z del connettore). La riga di stato della sezione mostra quell'angolo e si aggiorna dopo ogni passo di Rotate e Branch angle (`_then_show_elbow` avvolge gli `apply` degli slider).
- Le voci 15 / 22.5 / 30 / 45 restano con il comportamento della 0.4.0 (rotazione ricalcolata), per chi vuole una curva commerciale.
- `create()` controlla che la derivazione punti verso la quota richiesta (su con offset +, giu' con offset -) prima di calcolare la distanza, con un messaggio esplicito; prima l'errore era il generico "already past the branch level".
- Chiave di config rinominata da `branch_elbow` a `branch_elbow_mode` (0 = From rotation): con la vecchia chiave chi aveva salvato 45 avrebbe ritrovato il comportamento che ha causato il problema.
- Finestra 860 x 540: nello screenshot in Revit la riga di stato di Branch angle era tagliata in basso con l'altezza 470, mentre nel rendering WPF fuori da Revit entrava. In Revit i testi occupano piu' spazio: tenere margine.

**Esito**: sintassi verificata, da provare in Revit.

## 23.10.2026, versione 0.4.0: rotazione assoluta e stacco laterale (con AI)

**Richiesta** (screenshot di uno stacco: braga ruotata, tubo inclinato, curva, tubo orizzontale): 1) "l'angolo funziona, ma non ruota piu' in maniera corretta"; 2) se il raccordo e' gia' ruotato il comando non dice di quanto (es. stacco a 33°, portarlo a 90°); 3) un comando per creare uno stacco laterale come nell'immagine con un offset di quota preciso.

**Scelte dell'utente** (domande prima del codice): lo stacco si crea dal collettore con le impostazioni di rotazione impostate, cioe' dalla braga gia' sul collettore dentro Fitting Tool (scartato: pulsante separato che inserisce anche la braga, interpretazione mia della risposta libera); si inserisce l'angolo della curva e il tool calcola la rotazione (scartato: rotazione inserita e curva calcolata, che puo' dare curve non commerciali); offset come scorrimento, fondo tubo (scartati: asse-asse, quota assoluta).

**1. Rotazione sbagliata, causa probabile (non verificata in Revit)**: dalla 0.3.0 l'asse era ricalcolato a ogni passo sul primo connettore collegato, nell'ordine di `ConnectorManager.Connectors`. Con un T o una Y collegati alle due estremita' del collettore quel connettore poteva cambiare tra un passo e l'altro, invertendo il verso dell'asse: lo stesso delta ruotava una volta in un senso e una nell'altro. Un'altra possibilita': Y collegata solo dalla derivazione, ruotata attorno alla derivazione invece che al collettore. L'utente non ha descritto il sintomo; l'immagine sembra lo stacco desiderato, non l'errore.

**Fatto**:
- `FittingFrame` / `build_frame` in `FittingTool_Rotate.py`: asse fissato all'apertura. Priorita': coppia di connettori allineati e opposti (`_main_pair`: collettore di T e Y, raccordi diritti), poi estremita' collegata, poi primo connettore. Connettori ordinati per `Connector.Id` per avere sempre lo stesso ordine. Se un connettore collegato e' fuori asse: non ruotabile (ora vale anche per una Y collegata solo dalla derivazione).
- Lettura assoluta: terna destra (h, u, asse) con u = parte di Z normale all'asse (Y del progetto se l'asse e' verticale), h = u x asse; angolo = atan2(r.u, r.h) con r la direzione del connettore di riferimento (derivazione, o estremita' non sull'asse, o un asse del raccordo non parallelo). All'apertura se r.h < 0 si gira l'asse, cosi' il lato attuale e' lo 0 e l'angolo iniziale e' tra -90 e 90 (= inclinazione sull'orizzontale per un collettore orizzontale). `RotateElement` attorno all'asse con angolo positivo fa crescere la lettura (mano destra, h x u = asse).
- `Rotator.rotate_to` ruota ogni raccordo della differenza tra il suo angolo e il valore richiesto (portata in [-180, 180)), quindi piu' raccordi finiscono allo stesso angolo assoluto. Slider di rotazione che parte dall'angolo letto.
- Riferimento di ripiego: il primo tra BasisX / Y / Z del raccordo non parallelo all'asse. Con BasisX parallelo (manicotto) la lettura sarebbe rimasta 0 e ogni passo avrebbe ruotato di nuovo di tutto il valore.
- `FittingTool_SideBranch.py` (nuovo): `SideBranch` controlla le condizioni (un solo raccordo di tubazione ruotabile, con derivazione libera, collegato a un `Pipe`), calcola la quota (`target_z`, raggio interno dello stacco da un tubo di prova in una transazione annullata) e la rotazione (`target_angle`: con `d(phi) = along * asse + across * (cos(phi) * h + sin(phi) * u)` e `sin(elev) = d.Z` si ricava `sin(phi)`), poi `create` in una `run_step`: tubo dalla derivazione all'intersezione dell'asse con il piano z = quota, tubo orizzontale nella direzione in pianta, `ConnectTo` sulla derivazione, `NewElbowFitting` all'angolo. La curva sta sull'intersezione degli assi, quindi la quota e' esatta senza conoscere le misure dei raccordi.
- UI: due colonne (Rotate e Branch angle a sinistra, Side branch a destra), 860 x 470; `SliderSection.set_enabled`. Il pulsante passa la rotazione allo slider (`go_to`), controlla che sia arrivata, poi crea; dopo la creazione disattiva tutto il resto. Errori nella riga di stato della sezione.

**Verifiche fuori da Revit**: sintassi IronPython; test numerico in CPython (`roll_test.py` nella cartella temporanea della sessione, non nel repo) della terna, del verso dell'asse, della lettura, della rotazione a 90 e di `target_angle`: collettore orizzontale e in pendenza 2%, derivazione a sinistra e a destra, asse nei due versi; stacco a 33° letto 33, portato a 90 esatto, curva 30° -> inclinazione 30.0000. Finestra disegnata con WPF: entra senza scroll.

**Non provato in Revit**: che la rotazione ora sia corretta nel caso dell'utente; `ConnectTo` del tubo sulla derivazione e `NewElbowFitting` con le routing preferences degli scarichi; la transazione di prova del diametro interno dentro il `TransactionGroup`; la scelta del livello (`ReferenceLevel` del collettore); cosa succede con una Y collegata a un tubo diverso da `Pipe` (flessibile).

**Esito**: sintassi e matematica verificate, da provare in Revit.

## 23.10.2026, versione 0.3.1: errore all'apertura della finestra (con AI)

**Richiesta**: traceback dopo la selezione di un raccordo, `ValueError: invalid literal for float(): 2,5` in `SliderSection.step()`.

**Causa**: in Revit `"{:g}".format(2.5)` di IronPython segue le impostazioni regionali di Windows e con l'italiano scrive `2,5`; il testo delle voci del passo veniva poi riletto con `float()`. Il controllo fuori da Revit (motore IronPython caricato da PowerShell) non lo mostrava: li' la formattazione usciva con il punto. Attenzione quindi: un testo prodotto con `format` puo' avere la virgola, e va riletto solo con `parse_number()`.

**Fatto**: `step()` rilegge la voce con `parse_number()` (virgola o punto); anche la lettura dei passi salvati in config accetta la virgola. Le caselle dei valori erano gia' rilette con `parse_number()`.

**Esito**: sintassi verificata, da riprovare in Revit.

## 23.10.2026, versione 0.3.0: slider del parametro Angle (con AI)

**Richiesta** (screenshot della famiglia `e_PP.GG.JN_Y Branch`, parametro di istanza `Angle` = 45.00°, modificabile): il tool sara' usato per gli scarichi, i raccordi hanno un parametro `Angle` che definisce la direzione della derivazione; aggiungere uno slider anche per quello.

**Fatto**:
- `FittingTool_Rotate.py`: `run_step()` raccoglie la transazione silenziosa comune (preprocessore, commit, refresh della vista); `Rotator` la usa e ricalcola l'asse a ogni passo (prima era calcolato una volta all'apertura: con la derivazione collegata e l'angolo cambiato l'asse sarebbe stato vecchio). Nuovi `angle_parameter()` (nomi `Angle`, `Angolo`; istanza, `StorageType.Double`, non in sola lettura) e `AngleSetter` (valore assoluto in gradi, `Set` in radianti, rilegge il valore dopo la rigenerazione).
- `FittingTool_Form_ui.py`: `SliderSection` raccoglie il comportamento comune ai due slider (casella, passo, − / +, Back to start, stato); i controlli nel XAML hanno il prefisso `Rotate` o `Angle`.
- Errori: non piu' `MessageBox` ma il testo `Last change refused: ...` nella riga di stato, per entrambi gli slider, perche' trascinando ne sarebbe comparso uno per passo.
- Passi dell'angolo 1, 2.5, 5, 15, 45 (2.5 per arrivare a 67.5 e 87.5, angoli tipici degli scarichi); chiave di config `angle_step`. La chiave `step` della rotazione ora e' salvata come numero con decimali, legge anche il vecchio intero.
- Finestra: due sezioni una sotto l'altra, 460 x 530 (verificato con un rendering WPF fuori Revit che tutto entra senza scroll).

**Scelte da confermare con l'utente**: range 0-180 dello slider (la famiglia puo' accettarne meno); stesso valore su tutti i raccordi selezionati; nomi del parametro `Angle` / `Angolo` fissi nel codice.

**Non provato in Revit**: comportamento con la derivazione collegata (rifiuto o tubo spostato); che `Set` di un valore fuori dai limiti della famiglia dia un errore pulito e non un raccordo deformato; che `Angle` sia in radianti interni (`StorageType.Double` di un parametro di tipo angolo: atteso si').

**Esito**: sintassi IronPython verificata, da provare in Revit.

## 23.10.2026, versione 0.2.0: tolto il blocco dei parametri (con AI)

**Richiesta**: "Non funziona il lock, togliamolo da qui questa impostazione."

**Esito della 0.1.0 in Revit**: il blocco dei parametri (Extensible Storage + `IUpdater` su richiesta, vedi la voce sotto) **non funziona** secondo la prova dell'utente. Il sintomo preciso non e' stato riportato: non si sa se l'updater non scattava, se il valore veniva rimesso e poi resettato di nuovo da Revit, o se l'errore sul tratto compariva prima dell'updater. Chi riprova questa strada deve prima capire quale dei tre casi si verifica.

**Fatto**: eliminato `FittingTool_Lock.py`; tolta la sezione Lock parameters dalla finestra (ora una colonna, 440 x 350); `is_fitting` e le categorie spostati in `FittingTool_Rotate.py`; tolto `__persistentengine__` dallo script (serviva solo all'updater). Tooltip e Readme parlano solo della rotazione.

**Residui nei modelli di prova**: i raccordi bloccati con la 0.1.0 conservano l'entity dello schema `ESA_FittingLock` (GUID `6f1c2a4e-8d3b-4c57-9e21-3a7b5d0c9f14`). E' inerte: senza updater nessuno la legge. L'updater registrato in una sessione di Revit sparisce alla chiusura di Revit.

**Esito**: sintassi IronPython verificata, rotazione da provare in Revit.

## 23.10.2026, versione 0.1.0: prima versione (con AI)

**Richiesta**: un pulsante per i raccordi delle varie categorie con due azioni: ruotare un raccordo con uno slider; impedire che Revit rigeneri i raccordi quando si collega o cambia il tratto, perche' un parametro di istanza di misura modificato a mano, se resettato, impedisce la creazione del tratto.

**Scelte dell'utente** (domande prima di scrivere il codice):
- un solo pulsante con una finestra (scartato: pulldown con due pulsanti e `lib`);
- blocco con updater attivato su richiesta, per la sessione di Revit (scartati: updater sempre attivo registrato in `startup.py`, che avrebbe girato per tutti gli utenti; solo ripristino manuale, che non evita l'errore al momento della connessione);
- finestra modale con anteprima live (scartata: finestra non modale con il bridge `ExternalEvent` di ClashFinder, piu' codice e piu' fragile);
- caso reale: parametro di istanza resettato (non sostituzione del tipo, non dimensione del connettore).

**Fatto**:
- Rotazione: asse dal connettore collegato (`find_axis` in `FittingTool_Rotate.py`), `ElementTransformUtils.RotateElement` per ogni raccordo, una transazione per passo dello slider dentro un `TransactionGroup` (OK = `Assimilate`, Cancel = `RollBack`). Preprocessore che cancella gli avvisi e annulla il passo in caso di errore. `uidoc.RefreshActiveView()` dopo ogni passo.
- Blocco: nessuna API di Revit blocca la rigenerazione di un singolo elemento. Si salva un'istantanea dei parametri scelti nell'elemento (Extensible Storage, un campo stringa "valore<TAB>nome", valori in unita' interne con `repr`) e un `IUpdater` la rimette dopo ogni modifica. Trigger: `ElementMulticategoryFilter` delle quattro categorie AND `ExtensibleStorageFilter(SCHEMA_GUID)`, `Element.GetChangeTypeAny()`: l'updater vede solo i raccordi bloccati. `DeleteEntity` quando non resta nulla di bloccato, cosi' l'elemento esce dal filtro.
- Parametri bloccabili: solo non built-in (`Id` > 0), `StorageType.Double`, non in sola lettura. Scartato "tutti i parametri modificabili": built-in come quota e offset cambiano legittimamente quando si sposta il raccordo e il ripristino annullerebbe lo spostamento; quelli associati al connettore (diametro nominale) bloccherebbero il ridimensionamento del tratto. Per questo l'utente sceglie quali bloccare.
- Updater: registrato con `isOptional = True`, applicazione intera. Lo stato e' letto da `UpdaterRegistry.IsUpdaterRegistered`, non da una variabile di modulo. Ogni attivazione deregistra e registra un oggetto nuovo, per sostituire quello di un motore chiuso dopo un Reload. `__persistentengine__ = True` nello script, come ClashFinder, perche' l'updater e' un oggetto IronPython chiamato da Revit dopo la fine del comando.
- Ripristino manuale all'OK: un blocco gia' esistente tiene il valore salvato e lo rimette se diverso (`plan_changes` / `apply_plan`).
- Nessun SHIFT + CLICK: non c'e' un'esecuzione "con le ultime impostazioni" che abbia senso (l'angolo dipende dal raccordo). Da rivedere se serve, per esempio "ruota del passo salvato".
- `startup.py` non e' stato toccato.

**Non provato in Revit** (da verificare prima del merge):
- che l'updater con `GetChangeTypeAny` scatti quando Revit resetta il parametro durante la connessione o il cambio di misura del tratto, e che rimettere il valore nella stessa transazione eviti l'errore. Se Revit fallisce prima che l'updater giri, il blocco automatico non basta e servono altre strade (per esempio intercettare il failure con `FailuresProcessing` o rimettere il valore e ricollegare).
- se l'updater, rimettendo il valore, faccia ripartire una rigenerazione che lo resetta di nuovo (ciclo): la tolleranza `VALUE_TOLERANCE` evita scritture inutili ma non un conflitto con Revit;
- `Entity.Get[String]` / `Entity.Set[String]` da IronPython;
- `UpdaterRegistry.IsUpdaterRegistered` dopo un Reload di pyRevit e comportamento dell'updater orfano;
- l'asse dei raccordi non collegati (primo connettore) e di passerelle e tubi protettivi;
- la fluidita' dell'anteprima trascinando lo slider con passo 1 su molti raccordi;
- la finestra a 125% e 150% di scala (800 x 440).

**Esito**: sintassi IronPython verificata, da provare in Revit.
