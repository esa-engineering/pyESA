# MEP Align to Ceiling — strumento pyRevit

Allinea **in quota** i dispositivi MEP selezionati alla **faccia inferiore del
controsoffitto** che hanno sopra o, dove il controsoffitto manca, all'**intradosso del
solaio** che chiude il loro piano di riferimento.

È il gemello verticale di [MEPAlignToWall](../MEPAlignToWall.pushbutton/README.md): stessa
retta di analisi, stessa ricerca in host e collegamenti, stesso resoconto (qui come report
avanzato, dietro una finestra di esito con i totali e i non allineati), stessa gestione
degli avvisi di Revit.

> **Lo spostamento è solo verticale.** Posizione in pianta e rotazione non vengono mai
> toccate: il punto di inserimento scorre lungo l'asse Z del modello fino alla superficie
> bersaglio.

## Contenuto della cartella

| File | Ruolo |
| --- | --- |
| `MEPAlignToCeiling_script.py` | Logica dello strumento (Revit API, motore IronPython) |
| `MEPAlignToCeiling_ui.py` | Finestra di avvio (`XamlReader`, nessuna Revit API) |
| `MEPAlignToCeiling_form.xaml` | Layout della finestra di avvio |
| `MEPAlignToCeiling_result_ui.py` | Finestra di esito: tabella interrogabile (nessuna Revit API) |
| `MEPAlignToCeiling_result.xaml` | Layout della finestra di esito |
| `MEPAlignToCeiling_grid_filter.py` | Filtri per colonna in stile Excel: copia di `MEPQTO/mepqto_grid_filter.py` con nome proprio (Rocket mode). Le correzioni vanno riportate su entrambe le copie |
| `bundle.yaml` | Etichetta e descrizione del pulsante |
| `icon.png` / `icon.dark.png` | Icone per il tema chiaro e per quello scuro |
| `README.md` | Questo file |

## Uso

1. **Selezione.** Il comando parte dagli elementi già selezionati; se non ce ne sono chiede di
   indicarli nella vista (`Finish` per confermare). Gli elementi scelti restano selezionati,
   così si può rilanciare dopo una simulazione.
2. **Finestra di dialogo.**

   | Sezione | Contenuto |
   | --- | --- |
   | Elements | Quanti elementi sono selezionati e quanti hanno `u_VIS_Ceiling_installation = Yes` |
   | Categories | Solo le categorie presenti nella selezione, con i due conteggi. Ricerca, Select All / Deselect All |
   | False ceiling detectors | Quanti rilevatori da controsoffitto ci sono nella selezione e il valore di riserva di `dz_Buzzer` (cm, predefinito 50) |
   | Tolerance | Predefinita 50 cm, massimo 300 cm. Accetta virgola o punto |
   | Options | Cerca nei modelli collegati (attiva), salta gli elementi connessi (non attiva), simulazione (non attiva) |

   Il bersaglio non si sceglie: lo strumento riconosce da solo se sulla verticale del
   dispositivo il piano ha un controsoffitto (vedi "Controsoffitto prima del solaio").
   Le scelte (tranne la simulazione) restano memorizzate nella sezione `ESA_MEPAlignToCeiling`
   della configurazione pyRevit. Le categorie tolte restano tolte alla prossima apertura.
3. **Finestra di esito**: i totali e la tabella degli elementi non allineati, con le note e un
   pulsante per vederli nel modello. Il resoconto completo nel pannello di output è il **report avanzato**, stampato
   con il pulsante *Advanced Report* (vedi "Esito").

Le categorie gestite sono le stesse di MEPAlignToWall: Air Terminals, Communication Devices,
Data Devices, Electrical Fixtures, Fire Alarm Devices, Lighting Devices, Lighting Fixtures,
Nurse Call Devices, Security Devices.

## Screening: `u_VIS_Ceiling_installation`

Le famiglie possono essere montate a parete o a soffitto, quindi la categoria non basta.
Vengono trattate **solo** le istanze con il parametro **di istanza** Yes/No
`u_VIS_Ceiling_installation` = **Yes**. Le altre finiscono nella tabella *Not flagged for
ceiling installation*, raggruppate per tipo e motivo, con un collegamento che le seleziona
tutte:

| Motivo | Significato |
| --- | --- |
| `wall installation (... = No)` | Istanza a parete: è il caso normale |
| `... has no value` | Il parametro esiste ma non è valorizzato |
| `parameter ... not found on the instance` | La famiglia non ha il parametro |
| `... is a type parameter` | Il parametro esiste, ma come parametro di tipo: va corretto nella famiglia |
| `... is not a Yes/No parameter` | Esiste con quel nome ma di altro tipo |

Lo screening viene **prima** di tutto il resto: per un dispositivo a parete non si legge
nessuna geometria.

## La retta di analisi

Per ogni elemento si considera una retta **verticale**, passante per il punto di
inserimento, lunga **due volte la tolleranza** e con il punto medio sull'origine della
famiglia: si estende quindi di una tolleranza sopra e di una sotto il punto.

Un controsoffitto o un solaio è un bersaglio valido quando la retta **interseca la sua
superficie inferiore**, oppure quando il punto di inserimento cade **dentro il suo
spessore** (come per i muri: una superficie più spessa della tolleranza non deve lasciare il
dispositivo sepolto).

La retta va anche verso il basso di proposito: un dispositivo modellato nel plenum sopra il
controsoffitto viene riportato giù, alla faccia inferiore, attraversando il controsoffitto.

### Come si legge la superficie inferiore

Per ogni controsoffitto o solaio candidato si traccia una sonda verticale per il punto (in
pianta) che attraversa tutto il suo bounding box, e la si interseca (`Face.Intersect`) con
tutte le facce non verticali della geometria. La quota più bassa colpita è la superficie
inferiore, la più alta quella superiore. Non serve fidarsi del verso delle normali, e un
controsoffitto *Basic* senza spessore funziona come gli altri.

Se la verticale cade in un'apertura (cavedio, botola, foro), non c'è intersezione e quella
superficie in quel punto non esiste: è il comportamento corretto.

La lettura è **geometrica e indipendente dalle viste**: elementi nascosti, sezioni di
ritaglio o filtri di vista non cambiano il risultato.

## Il piano di riferimento

Valgono solo le superfici **sopra il livello di riferimento** dell'elemento
(`Level.ProjectElevation`, vedi la nota sulle coordinate in `CLAUDE.md`). Senza questo limite
la retta, che si estende anche verso il basso, prenderebbe per buono l'intradosso del solaio
su cui il piano poggia quando un dispositivo è modellato a quota zero.

| Caso | Livello usato |
| --- | --- |
| Punto di inserimento sopra il proprio livello | Il livello dell'istanza |
| Punto **sotto** il proprio livello (riferito al livello superiore con scostamento negativo) | Il livello più alto sotto il punto, con nota |
| Istanza senza livello | Il livello più alto sotto il punto, con nota |
| Nessun livello sotto il punto | Nessun limite inferiore, con nota |

Il **solaio di chiusura** è il pacchetto di solai più basso fra quelli sopra il livello.
I **controsoffitti del piano** sono quelli sopra il livello e sotto il solaio di chiusura: un
controsoffitto più in alto appartiene al piano superiore.

## Più istanze impilate, anche da modelli diversi

Un solaio è spesso fatto di più istanze: la soletta strutturale nel modello strutturale, il
massetto e l'intonaco nell'architettonico. Le istanze che, sulla verticale del punto, si
toccano o distano meno di **5 mm** (`STACK_GAP_MM`) formano un **pacchetto unico**, e il
bersaglio è la faccia inferiore del pacchetto, cioè quella dell'istanza più bassa. Lo stesso
vale per i controsoffitti in più strati. I controsoffitti e i solai non si fondono fra loro.

Lo scarto è più largo dei 2 mm dei muri perché fra due modelli collegati di discipline
diverse il contatto esatto non è garantito; un vuoto vero (plenum) è sempre di decine di
centimetri. La nota del resoconto dice quante istanze formano il pacchetto e da quanti
modelli vengono.

La fusione serve anche al pavimento sopraelevato: appoggiato sul solaio del piano si fonde con
lui in un pacchetto che sta sotto il livello, e non viene scambiato per un intradosso.

## Controsoffitto prima del solaio

Il bersaglio è automatico, non c'è nessuna scelta nella finestra:

| Situazione sulla verticale del punto | Bersaglio |
| --- | --- |
| C'è un controsoffitto nel piano | Il controsoffitto più vicino al punto (faccia inferiore del suo pacchetto) |
| Nessun controsoffitto | Il solaio di chiusura |
| Nessun controsoffitto e nessun solaio | Elemento saltato |
| Rilevatore da controsoffitto | Sempre il solaio di chiusura (sezione seguente) |

Se nel piano c'è un controsoffitto, il bersaglio resta il controsoffitto **anche quando è
fuori tolleranza**: l'elemento viene saltato e riportato, non portato all'intradosso del
solaio, perché finirebbe nel plenum sopra il controsoffitto. La nota del resoconto segnala se
il solaio sarebbe stato entro tolleranza.

Con due controsoffitti nel piano sulla stessa verticale vince il più vicino al punto; se il
secondo è entro 20 mm di differenza la nota lo segnala come ambiguo.

## Rilevatori da controsoffitto (`dz_Buzzer`)

Le istanze il cui **family name** contiene sia `e_FA.DE` sia `False Ceiling` (confronto senza
distinzione fra maiuscole e minuscole, `BUZZER_FAMILY_TOKENS`) sono rilevatori installati nel
plenum, con la parte visibile sul controsoffitto:

- passano lo **stesso screening** delle altre famiglie (`u_VIS_Ceiling_installation` = Yes);
- vanno **sempre all'intradosso del solaio di chiusura**, anche quando il piano ha un
  controsoffitto;
- il parametro di istanza `dz_Buzzer` (lunghezza) riceve la distanza **intradosso solaio -
  faccia inferiore del controsoffitto**, misurata sulla verticale del dispositivo con le stesse
  regole del bersaglio (controsoffitto del piano più vicino al punto, pacchetti di più istanze
  sulla faccia inferiore);
- dove sotto il dispositivo **non c'è controsoffitto**, `dz_Buzzer` prende il valore della
  finestra di avvio, e la nota lo dichiara.

La retta di analisi vale se interseca **il solaio oppure il controsoffitto**: questi rilevatori
sono di solito modellati all'altezza del controsoffitto, e misurare solo dal solaio li
manderebbe quasi tutti oltre tolleranza.

Senza un `dz_Buzzer` scrivibile il rilevatore viene **saltato**, non spostato: portato al
solaio, la sua parte visibile resterebbe lassù. Motivi: parametro mancante, di tipo invece
che di istanza, non di tipo lunghezza, in sola lettura (pilotato da una formula). Senza
solaio di chiusura sopra il dispositivo, viene saltato anch'esso.

Spostamento e scrittura di `dz_Buzzer` avvengono nella stessa `SubTransaction`: riescono o
falliscono insieme. Dopo il commit anche `dz_Buzzer` viene riletto.

## Modelli collegati

Controsoffitti e solai si cercano nel documento corrente e nei `RevitLinkInstance` caricati:
una sola raccolta per documento sul volume occupato dalla selezione, poi un filtro in pianta
per elemento sui bounding box già letti. La geometria di ogni superficie si estrae una volta
sola e resta in cache.

Il punto di inserimento viene portato nelle coordinate del collegamento con la trasformazione
inversa; le quote trovate tornano nell'host con la quota dell'origine del collegamento.

| Collegamento | Esito |
| --- | --- |
| scaricato | ignorato, negli avvisi |
| inclinato o capovolto | ignorato: la sua verticale non è quella dell'host |
| in scala | ignorato: le distanze non sarebbero confrontabili |
| speculare | **usato**: la riflessione è nel piano orizzontale e non cambia le quote (a differenza di MEPAlignToWall) |

Gli elementi collegati sono usati solo come riferimento: lo strumento non apre mai una
transazione su un collegamento.

Superfici **demolite** in una fase vengono scartate e listate in *Ceilings and floors
discarded*, insieme a quelle con geometria o estensione non leggibile.

## Esclusioni

| Motivo | Perché |
| --- | --- |
| hosted by ceiling *n* | È già sulla faccia del suo host |
| hosted by a face of a linked model | Famiglia basata su faccia: lo spostamento in Z la staccherebbe |
| hosted by *categoria n* | Ospitato da muro, pavimento o altro |
| element is pinned | Va sbloccato a mano |
| element inside group | I gruppi non si toccano: spostare un membro cambierebbe tutte le istanze |
| sub-component of a nested family | Si muove con il padre |
| inactive design option | Fuori dall'opzione attiva |
| borrowed by another user | Workset o elemento preso da un altro utente |
| connected to other elements | Solo con l'opzione attiva |
| element has no insertion point | Famiglie basate su linea o locali |

Un livello come host **non** è un'esclusione: è il caso normale delle famiglie basate sul
livello, che si spostano in quota cambiando lo scostamento.

Per gli elementi saltati per un motivo diverso dallo screening il resoconto riporta comunque
la superficie bersaglio e lo spostamento che servirebbe: uno a 12 mm vale la pena di
sbloccarlo, uno a 400 mm probabilmente no.

## Applicazione

- Una sola transazione, `Align MEP to ceilings`, quindi un solo passo di annullamento.
- Una `SubTransaction` per elemento: un errore lascia intatto quell'elemento.
- La traslazione è ricalcolata dalla quota corrente verso la quota bersaglio **assoluta**; X
  e Y del vettore sono sempre zero.
- Dopo il commit la quota viene **riletta**: se Revit ha accettato lo spostamento ma
  l'elemento non è alla quota bersaglio (famiglia vincolata al piano di lavoro, quota
  bloccata), la colonna *Outcome* lo dice invece di dichiararlo allineato.
- Niente `swallow_errors` di pyRevit: un `IFailuresPreprocessor` dedicato registra gli avvisi
  per il resoconto, elimina solo i warning senza risoluzioni e chiede il rollback sugli
  errori. Nessun vincolo dell'utente viene sbloccato.

## Esito

### Finestra di esito

È il report semplificato. In testa solo i **totali**: elementi selezionati, allineati (in
simulazione: da allineare), di cui rilevatori da controsoffitto con `dz_Buzzer`, già a posto,
non allineati. Gli elementi allineati **non sono elencati**.

Sotto, la tabella dei soli **elementi non allineati**: falliti in applicazione, oltre
tolleranza e saltati, screening compreso. Se non ce n'è nessuno la tabella lo dice.

| Colonna | Contenuto |
| --- | --- |
| Show | Pulsante 🔍: seleziona l'elemento nel modello e lascia navigare la vista (sotto) |
| Status | `Failed` (errore, rollback, quota o `dz_Buzzer` diversi dal bersaglio dopo il commit) o `Not aligned` |
| Id, Category, Family: Type, Level | Identificazione dell'elemento; il livello è quello del piano di riferimento usato |
| Target | `Ceiling` o `Floor soffit`, se un bersaglio c'era |
| Move (mm) | Spostamento che servirebbe, positivo verso l'alto |
| dz_Buzzer (mm) | Valore che si è tentato di scrivere sui rilevatori falliti |
| Notes | Motivo |

La tabella è **interrogabile**: ricerca su tutte le colonne, filtro per valore dall'imbuto di
ogni intestazione (come in Excel), ordinamento cliccando l'intestazione, *Clear Filters*.
Le righe `Not aligned` sono in arancio chiaro, le `Failed` in rosso chiaro. La griglia usa
lo schema di MEPQTO: `DataTable` legata alla `DataGrid`, ricerca con `DataView.RowFilter` e
`LIKE` con escape, filtri per colonna da `GridFilters`.

**Navigazione nel modello**, con lo schema di RNWFlowCalculation: il pulsante 🔍 di una riga
chiude la finestra, seleziona l'elemento e lo inquadra (`ShowElements`), poi lascia navigare
la vista attiva (pan, zoom, orbita) con una `PickObjects` in cui si possono cliccare solo gli
elementi della riga. Con **Fine** o **Esc** la finestra si riapre com'era: ricerca, filtri per
colonna e riga di partenza (`ResultState`). *🔍 Show All Listed* fa lo stesso con tutte le
righe visibili con la ricerca e i filtri correnti. La selezione resta sugli elementi anche
dopo il ritorno alla tabella.

La finestra è modale: una finestra non modale richiederebbe un `ExternalEvent` per ogni
chiamata alla Revit API.

### Report avanzato

Il pulsante **Advanced Report** chiude la finestra e stampa nel pannello di output il
resoconto completo, con le sezioni elencate qui sotto. Se la finestra di esito non si apre
(per esempio per un errore nel layout), il report avanzato viene stampato comunque.

| Sezione | Contenuto |
| --- | --- |
| Intestazione | Impostazioni usate, quanti esclusi dallo screening |
| Summary by category | Selezionati, allineati, già a posto, non a soffitto, saltati, oltre tolleranza |
| Elements aligned / to align | Livello, bersaglio, superficie, quota prima e dopo rispetto al livello, spostamento (positivo verso l'alto), `dz_Buzzer` prima -> dopo, note, esito |
| Elements already aligned | Già entro 1 mm dalla superficie |
| Skipped as beyond the tolerance | Superficie, spostamento necessario, sforamento, e una riga che dice di quanto alzare la tolleranza |
| Skipped elements | Motivo, superficie bersaglio, spostamento necessario |
| Not flagged for ceiling installation | Raggruppati per tipo e motivo |
| Ceilings and floors discarded | Superfici scartate e perché |
| Revit warnings / Warnings | Avvisi di Revit e avvisi dello strumento (link ignorati, nessuna superficie trovata) |

Le superfici dei modelli collegati non sono selezionabili dall'host: per loro si stampa l'id.

## Limiti noti

- **Il punto di inserimento finisce sulla superficie.** Se una famiglia ha l'origine al centro
  del proprio ingombro invece che sulla faccia di montaggio, l'elemento risulta per metà
  dentro il controsoffitto. Le colonne di quota del resoconto servono a vederlo sul primo
  modello reale.
- Solo elementi con `LocationPoint`.
- Solo categoria **Floors** per i solai: un tetto usato come ultimo solaio non viene letto.
  Per includerlo basta aggiungere `OST_Roofs` alla raccolta in `collect_in()`.
- La superficie inferiore è la quota più bassa colpita sulla verticale **del punto di
  inserimento**: su un intradosso inclinato o a gradini l'ingombro della famiglia può
  comunque sporgere.
- Un controsoffitto in un locale adiacente non conta: deve trovarsi sulla verticale del punto.
- Le fasi sono gestite solo scartando le superfici demolite; una superficie creata in una fase
  successiva a quella della vista viene comunque usata.
- Le opzioni di progetto dei modelli collegati non vengono filtrate.

## Da verificare in Revit

Lo strumento è stato provato fuori da Revit (sintassi, le due finestre con il motore
IronPython di pyRevit, logica di decisione e righe della finestra di esito con superfici e
record finti). Va ancora provato su un modello reale:

1. `Face.Intersect` con la sonda verticale su solai con aperture e su controsoffitti *Basic*;
2. spostamento in Z di famiglie basate sul livello e basate su piano di lavoro;
3. pacchetto solaio strutturale (link STR) + finiture (link ARC), e controsoffitto in più
   istanze;
4. dispositivo riferito al livello superiore con scostamento negativo;
5. rilevatori `e_FA.DE` / `False Ceiling`: scrittura di `dz_Buzzer` e posizione della parte
   visibile sul controsoffitto;
6. pulsante 🔍: selezione, inquadratura, navigazione con `PickObjects` e ritorno alla
   finestra con lo stato di prima;
7. tempi su una selezione grande in un modello con molti controsoffitti.
