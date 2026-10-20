# MEP Workset Fix

Riassegna il workset agli elementi MEP in base alla disciplina. L'utente sceglie in una
finestra il workset di destinazione per quattro sezioni (impianti meccanici, idraulici,
antincendio, elettrici); lo script decide la sezione di ogni elemento dalla categoria e dal
nome del suo System Type, mostra un riepilogo e, dopo conferma, sposta gli elementi.

## Contenuto della cartella

| File | Ruolo |
| --- | --- |
| `MEPWorksetFix_script.py` | regole, lettura del modello, piano, scrittura, report |
| `MEPWorksetFix_ui.py` + `MEPWorksetFix_form.xaml` | finestra delle impostazioni |
| `bundle.yaml` | etichetta e tooltip del pulsante |
| `icon.png` / `icon.dark.png` | icone per il tema chiaro e scuro |
| `README.md` | questo file |

## Uso

Il comando funziona solo su un modello di progetto con worksharing attivo; altrimenti un
avviso spiega perché e il comando termina.

**Workset delle quattro sezioni.** Ogni tendina elenca i workset utente del modello
(`WorksetKind.UserWorkset`, in ordine alfabetico), preceduti da **(do not change)**: una
sezione lasciata così non sposta nessuno dei suoi elementi, che compaiono nel report come
`SECTION SET TO DO NOT CHANGE`. Serve almeno una sezione impostata. Le quattro sezioni
stanno in un'unica tabella *Target worksets*, una riga ciascuna: nome della sezione,
tendina e, a destra, un testo grigio che dice quali categorie e quali System Type
finiscono in quella sezione. Il testo è generato dalle regole dello script, quindi resta
allineato al codice. In testa alla finestra il riquadro *How it works* riassume il
funzionamento.

Alla prima apertura, se non c'è una scelta salvata, la finestra propone il primo workset il
cui nome contiene una parola chiave della sezione (`MECH`, `MECC`, `HVAC`; `PLUMB`, `IDRA`,
`IDRI`, `SANIT`; `FIRE`, `ANTINC`, `SPRINK`; `ELEC`, `ELET`). È solo una proposta, visibile
prima di OK.

**Workset da non toccare.** Nel modello possono esserci workset di elementi temporanei
(provvisori, di cantiere, di studio) che non vanno spostati. La sezione *Worksets to leave
untouched* parte con una tendina (`(none)` = nessuno); **+ Add workset** aggiunge un'altra
riga, il pulsante ✕ di una riga la toglie (sull'ultima riga rimasta la riporta a
`(none)`). Le righe vuote e i doppioni si ignorano. Un elemento che **si trova già** in uno
di questi workset non viene spostato, qualunque sia il suo System Type, e compare nel
report come `IN A WORKSET LEFT UNTOUCHED`. Il controllo viene prima di tutte le altre
regole. Un workset non può essere insieme destinazione di una sezione e workset da non
toccare: la finestra lo segnala e resta aperta.

**Elementi da elaborare.**

- *All elements in the model*: tutto il modello. È l'unico ambito in cui si elaborano
  anche i **Duct Systems** e i **Piping Systems**: con vista attiva o selezione vengono
  ignorati e il report lo dice in testa.
- *Only elements visible in the active view*: disabilitata se la vista attiva non è una
  vista grafica del modello (tavole, abachi, modelli di vista).
- *Only the current selection*: disabilitata se non c'è nulla di selezionato. Se al lancio
  c'è una selezione, è l'opzione proposta.

**Conferma.** Dopo l'analisi un avviso riporta quanti elementi verranno spostati, sezione
per sezione, e quanti restano dove sono. *Sì* applica le modifiche; *No* stampa lo stesso
report come anteprima ("Preview only") senza toccare il modello. Se non c'è nulla da
spostare il report esce subito, senza domanda.

## Regole di assegnazione

| Categorie | System Type contiene | Sezione |
| --- | --- | --- |
| Cable Trays, Cable Tray Fittings, Conduits, Conduit Fittings, Data Devices, Electrical Equipment, Electrical Fixtures, Fire Alarm Devices, Lighting Devices, Nurse Call Devices, Security Devices, Telephone Devices | (non si legge) | elettrici, sempre |
| Air Terminals, Ducts, Duct Fittings, Duct Accessories, Flex Ducts, Duct Systems (solo intero modello) | `DS.FF` | antincendio |
| | `DS.HV` | meccanici |
| | `DS.OT` | meccanici |
| Pipes, Pipe Fittings, Flex Pipes, Piping Systems (solo intero modello) | `PS.FF` | antincendio |
| | `PS.PD` | meccanici |
| | `PS.SS` | idraulici |
| | `PS.GA` | idraulici |
| Sprinklers | (non si legge) | antincendio, sempre |

Il confronto è "contiene" e non distingue maiuscole e minuscole (`ds.hv` vale come
`DS.HV`). Le regole sono in testa allo script (`DUCT_RULES`, `PIPE_RULES`,
`CATEGORY_RULES`): per aggiungere un codice basta una coppia `("XX.YY", sezione)`, per
aggiungere una categoria una riga in `CATEGORY_RULES` con la regola dei canali o dei tubi
(`_duct_rule`, `_pipe_rule`) oppure una sezione fissa.

### Lettura del System Type

1. Si legge il parametro di istanza `RBS_DUCT_SYSTEM_TYPE_PARAM` (elementi canale) o
   `RBS_PIPING_SYSTEM_TYPE_PARAM` (elementi tubo). Se contiene un `ElementId` si prende il
   nome del tipo di sistema (`Element.Name.GetValue`, perché `el.Name` in IronPython
   fallisce su alcune sottoclassi), altrimenti il testo del parametro.
2. Se il parametro è vuoto (tipicamente un raccordo non collegato) si guardano i sistemi
   dei connettori del dominio giusto (`DomainHvac` / `DomainPiping`) e si prendono i nomi
   dei loro tipi.

I **Duct Systems** (`OST_DuctSystem`, regole dei canali) e i **Piping Systems**
(`OST_PipingSystem`, regole dei tubi) sono essi stessi sistemi: il System Type è il
loro tipo (`GetTypeId()`), letto direttamente senza parametro né connettori
(`from_element_type` in `CategoryRule`). Non hanno isolamenti. La regola è marcata
`model_only`: `active_rules()` la esclude quando l'ambito è la vista attiva o la
selezione, perché i sistemi non sono elementi visibili in una vista e di rado si
selezionano.

Se nessun nome corrisponde a una regola l'elemento resta dov'è (`NO MATCHING SYSTEM TYPE`).
Se i nomi corrispondono a regole di **due sezioni diverse** (per esempio un raccordo fra un
sistema DS.FF e uno DS.HV) l'elemento resta dov'è (`CONFLICTING SYSTEM TYPES`): lo script
non sceglie al posto dell'utente.

## Elementi non spostati

Ogni elemento delle categorie gestite compare nel report con un esito. Oltre a quelli già
nel workset giusto (`ALREADY CORRECT`) e a quelli senza regola, restano invariati:

| Esito | Motivo |
| --- | --- |
| `IN A WORKSET LEFT UNTOUCHED` | l'elemento è in uno dei workset da non toccare scelti nella finestra |
| `SECTION SET TO DO NOT CHANGE` | la sezione dell'elemento è impostata su (do not change) |
| `IN A GROUP` | l'elemento, o un suo isolamento / rivestimento, è membro di un gruppo di modello (nota: nome del gruppo) |
| `NESTED IN A FAMILY` | istanza di famiglia annidata condivisa (`SuperComponent`), tipicamente Electrical Fixtures o Fire Alarm Devices dentro un'altra famiglia: il workset segue la famiglia ospite e non si può cambiare (nota: nome e id della famiglia ospite) |
| `WORKSET READ-ONLY` | il parametro `ELEM_PARTITION_PARAM` manca o è in sola lettura |
| `OWNED BY ANOTHER USER` | l'elemento, o un suo isolamento / rivestimento, è preso da un altro utente (nota: nome dell'utente) |
| `NOT UPDATED FROM CENTRAL` | l'elemento è cambiato o è stato cancellato nel centrale: serve *Reload Latest* |
| `TARGET WORKSET OWNED BY ANOTHER USER` | il workset di destinazione è preso da un altro utente |
| `ERROR` | eccezione durante l'analisi o la scrittura, oppure workset invariato dopo il commit |

**Gruppi di modello.** Lo strumento non cambia mai il workset di un gruppo né dei suoi
membri. I gruppi non appartengono alle categorie gestite e, per sicurezza, vengono comunque
scartati in modo esplicito dopo la raccolta (`collect_elements`), anche se in futuro si
aggiungessero categorie. Gli elementi che stanno dentro un gruppo (`Element.GroupId`
valido, `group_of()`) restano dove sono e compaiono come `IN A GROUP` con il nome del
gruppo. Lo stesso vale per un host che non è in un gruppo ma ha un isolamento o un
rivestimento dentro un gruppo: l'host viene saltato per non lasciarli separati. Per
correggere quegli elementi si modifica il gruppo in Revit (Edit Group) e si rilancia lo
strumento sugli elementi esplosi, oppure si corregge il workset dentro il gruppo a mano.
Se un elemento di un gruppo è già nel workset giusto, l'esito è `ALREADY CORRECT`.

I controlli di proprietà (`WorksharingUtils.GetCheckoutStatus`, `GetModelUpdatesStatus`)
si fanno **prima** della transazione: dentro una transazione il valore è cache locale e non
affidabile (lo dice la documentazione dell'API). Escludere in anticipo gli elementi non
modificabili evita che un solo elemento preso da un collega faccia fallire il commit di
tutti gli altri.

## Isolamenti e rivestimenti

Per ogni elemento da spostare si leggono isolamenti e rivestimenti
(`InsulationLiningBase.GetInsulationIds` / `GetLiningIds`). Se uno di essi è preso da un
altro utente, anche l'host viene saltato. Dopo aver spostato l'host, gli isolamenti e i
rivestimenti rimasti su un altro workset vengono spostati a mano, se il loro parametro è
scrivibile. Il report riporta quanti ne sono stati spostati; se qualcuno resta comunque su
un altro workset lo segnala e rimanda a *Set Insulation Workset*.

## Scrittura

Una sola transazione ("MEP Workset Fix", un passo di annulla) imposta
`ELEM_PARTITION_PARAM` con l'id del workset di destinazione. Un errore su un elemento lo
segna `ERROR` e non ferma gli altri. A transazione chiusa ogni elemento viene riletto: è
`MOVED` solo se il suo workset è davvero quello di destinazione.

Il modello non viene modificato se si annulla la finestra, se si risponde *No* alla
conferma o se non c'è nulla da spostare.

## Report

Il report di pyRevit esce sempre, anche quando non c'è nulla da fare, e contiene:

- modello, ambito, workset scelto per ogni sezione e workset lasciati invariati;
- **Summary**: per categoria, elementi trovati, spostati (o da spostare, in anteprima), già
  corretti, senza regola, saltati, con la riga dei totali;
- **Moved / To move**: gruppi per sezione, workset di partenza, workset di arrivo e
  categoria, con un link che seleziona gli elementi del gruppo;
- **Left unchanged: System Type not recognised**: raggruppati per esito, categoria e nome
  del System Type, con link di selezione; è la tabella da guardare per scoprire codici di
  sistema non previsti;
- **Skipped**: un link di selezione per motivo, poi il dettaglio elemento per elemento
  (al massimo 300 righe, `MAX_DETAIL_ROWS`).

## Impostazioni salvate

Nella sezione `ESA_MEPWorksetFix` della config di pyRevit: il nome del workset scelto per
ogni sezione (`target_mechanical`, `target_plumbing`, `target_fire`, `target_electrical`),
l'elenco dei workset da non toccare (`excluded_worksets`; alla riapertura si ricrea una riga
per ogni nome ancora presente nel modello) e se l'ultimo ambito era la vista attiva
(`last_scope_view`). I workset si ricordano **per
nome**, così la scelta vale anche su altri progetti con la stessa nomenclatura. La
selezione non si ricorda, perché dipende dal momento.

## Motore Python

IronPython 2.7, compatibile con Rocket mode: il modulo `MEPWorksetFix_ui` non tiene
documento, vista o finestra di output a livello di modulo e non ha cache. Dopo una modifica
al modulo serve **pyRevit → Reload**. Compatibile con Revit 2022 - 2026 (`ElementId.Value`
in 2026, `IntegerValue` prima; `WorksetId.IntegerValue` esiste in tutte le versioni).

## Limiti e punti da verificare

- Le categorie gestite sono solo quelle della tabella. Communication Devices, Generic
  Models, Lighting Fixtures, Mechanical Equipment, Pipe Accessories, Plumbing Fixtures e
  Specialty Equipment restano fuori per scelta: il loro workset si assegna con un abaco
  multicategoria (Multi-Category Schedule), e la finestra lo ricorda nel riquadro *How it
  works*.
- Si lavora solo sul modello aperto, non sui collegamenti.
- Non è verificato in Revit se isolamenti e rivestimenti seguano l'host da soli quando
  cambia il workset: lo script gestisce entrambi i casi.
- Il blocco sul workset di destinazione preso da un altro utente è prudenziale: non è
  verificato che Revit rifiuti davvero lo spostamento in quel caso.
- Elementi nelle opzioni di progetto non primarie vengono trattati come gli altri.

## Autori

- Andrea Patti
