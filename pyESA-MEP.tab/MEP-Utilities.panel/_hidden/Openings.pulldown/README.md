# Openings (forometrie, PFV)

Fori di predisposizione ("provision for void") dove tubi, canali, passerelle e conduit del modello host
attraversano muri e solai, del modello host o di link RVT. Tutto passa da due parametri di testo:
niente Extensible Storage, niente commenti, Mark non toccato.

## Struttura

| Pulsante | Cosa fa |
| --- | --- |
| Place | Selezione di muri e solai, poi un foro per ogni attraversamento (accorpati se vicini) |
| Refresh | Controllo veloce di fori e MEP con lo stato compilato, ricreazione a scelta |
| Settings | Impostazioni per utente (`script.get_config("ESA_Openings")`) |
| Families | Crea le quattro famiglie predefinite dal template Generic Model, le salva in .rfa e le carica |

- I due `script.py` di Place e Refresh sono **identici tranne `MODE`** e `__title__`.
- Il blocco "comune" (nome della configurazione, predefiniti, nomi dei parametri) e' ripetuto in Place, Refresh e
  Settings; nomi di famiglie e parametri anche in Families. Una modifica va riportata ovunque.

## Parametri

| Elemento | e_GEN_PFV_ID | e_GEN_PFV_Status |
| --- | --- | --- |
| Foro (Generic Model) | Id degli elementi MEP: `111,222` | `MEC - New`, `ELE - OK`, `MEC+ELE - Misaligned`... |
| Elemento MEP | codici dei fori: `PFV123456,PFV123999` (PFV + Id del foro) | `MEC - New`, `ELE - Changed`... |

MEC = tubi e canali, ELE = passerelle e conduit. Stati: New, OK, Split, Joined, Misaligned, Changed, Obsolete.
Split e Joined: fori ricreati dal Refresh dividendo un foro misto per disciplina, o unendo un foro MEC e uno ELE.
Un foro e' "gestito" se ha `e_GEN_PFV_Status` compilato, qualunque sia la famiglia: si possono cambiare le
famiglie in Settings senza perdere i fori gia' posati.

Famiglie predefinite (Generic Model non ospitate, origine al centro, parametri di istanza condivisi):

- `e_GM.PH.PFV_Wall_Rect`: X `e_DIM_Width_1`, Y `e_DIM_Depth_1`, Z `e_DIM_Height_1`
- `e_GM.PH.PFV_Wall_Round`: cilindro lungo Y, `e_DIM_Diameter_1`, `e_DIM_Depth_1`
- `e_GM.PH.PFV_Floor_Rect`: X `e_DIM_Width_1`, Y `e_DIM_Height_1`, Z `e_DIM_Depth_1`
- `e_GM.PH.PFV_Floor_Round`: cilindro lungo Z, `e_DIM_Diameter_1`, `e_DIM_Depth_1`

Una famiglia scelta in Settings deve seguire gli stessi assi e gli stessi nomi di parametro.

## Settings (per utente)

Famiglie dei quattro casi; franco per lato (25 mm), arrotondamento per eccesso (50 mm), accorpamento entro
(100 mm), sporgenza oltre ogni faccia (0 mm); MEC ed ELE separati (no); stato Misaligned nel Refresh (si').

## Place

1. `PickObjects(PointOnElement)`: muri e solai, host o link.
2. Solidi dell'elemento in coordinate host, MEP rettilinei vicini (bounding box), intersezione asse/solido
   (`IntersectWithCurve`): ogni segmento interno e' un attraversamento.
3. Muro: normale nel punto (vale anche per muri curvi). Solaio: verticale. Elemento MEP quasi parallelo (sotto
   l'angolo minimo di Settings, default 15 gradi = pendenza 27%): corre dentro, nessun foro. Esclude scarichi
   a pavimento e montanti dentro i muri. Nel Refresh la stessa soglia rende Changed un elemento ruotato fino a
   correre dentro.
4. Sezione esterna con isolamento, dai solidi di elemento e isolamento sugli assi del connettore (vale per
   canali ruotati e passerelle di costa). La sezione all'ingresso e all'uscita, proiettata sul piano del foro,
   da' l'ingombro anche per gli attraversamenti obliqui.
5. Accorpamento: stesso elemento attraversato, piani paralleli, distanza <= accorpamento, stessa disciplina se
   separate. Tondo solo per un elemento tondo singolo e (quasi) perpendicolare.
6. Profondita' = spessore attraversato + 2 x sporgenza.
7. Gia' forato: il punto medio dell'attraversamento sta nel bounding box di un foro gestito.

## Refresh

Il muro non viene riletto: il foro stesso fa da muro (piano per il centro, perpendicolare all'asse di
attraversamento, spessore = `e_DIM_Depth_1`). Per ogni foro gestito, con gli Id di `e_GEN_PFV_ID`:

- elemento sparito, non piu' attraversante, o uscito oltre meta' foro + accorpamento: **Changed**;
- sezione (con isolamento) non piu' tutta dentro il foro: **Changed**;
- tutti dentro ma il foro ideale si e' spostato (5 mm) o ha altre misure (1 mm): **Misaligned** (o Changed, se
  l'opzione e' spenta);
- nessun elemento rimasto: **Obsolete** (il foro resta).

Scelta: ricreare i Changed, i Changed e i Misaligned, o aggiornare solo gli stati. Un foro ricreato ha un nuovo
Id (quindi un nuovo codice PFV): i codici sui MEP vengono riscritti. I MEP senza piu' fori hanno i campi PFV
svuotati.

## Limiti noti (da verificare in Revit)

- Lo spostamento o la modifica del muro nel link non e' rilevata dal Refresh.
- Nuovi MEP che attraversano un foro esistente non vengono aggiunti dal Refresh: si rilancia Place.
- Families: quote sulle facce dell'estrusione e quota di diametro sull'arco; la prova di flessione nel report dice
  se un parametro non muove la geometria.

## Refresh: ambito e spunta MEC/ELE

- **Ambito:** tutti i fori, oppure quelli dei muri e solai selezionati. Un foro appartiene a un elemento se una
  linea per il suo centro, lungo Y o Z della famiglia, entra nel solido dell'elemento. Con l'ambito ridotto i MEP
  tengono i codici dei fori fuori ambito (se esistono ancora) e il loro stato.
- **Spunta MEC/ELE cambiata dopo Place:**
  - separati: un foro misto e' Changed e, ricreato, diventa due fori (uno per disciplina);
  - insieme: un foro MEC e uno ELE sullo stesso piano, a distanza <= accorpamento, sono Changed e, ricreati,
    diventano un foro solo.
- **Fori avvicinati:** due o piu' fori sullo stesso piano i cui fori ricalcolati (franco compreso) distano al
  massimo la distanza di unione sono Changed e, ricreati, diventano uno solo con stato Joined. Con MEC ed ELE
  separati si uniscono solo fori della stessa disciplina.

## Strutture composite (Place)

Un muro composto (un solo elemento a strati) e' gia' attraversato tutto dall'intersezione con il suo solido. Se
invece la struttura e' fatta di piu' elementi (es. muro strutturale nel link, intercapedine, controparete nel
modello host), l'attraversamento viene allungato strato per strato:

- strati cercati: muri (fori a muro) o solai (fori a solaio) del modello host e dei link RVT caricati, entro 1 m
  dall'elemento scelto;
- uno strato conta se, lungo l'asse dell'elemento MEP, sta entro l'intercapedine massima di Settings ("Join
  layers across", 200 mm) da quanto gia' attraversato; l'intercapedine entra nel foro;
- per i muri lo strato deve essere parallelo all'elemento scelto (entro circa 11 gradi): un muro a T o d'angolo
  non conta; e uno strato attraversato per piu' di 1 m e' scartato (l'elemento MEP ci corre dentro);
- si ripete finche' non si trovano altri strati: il foro attraversa tutta la struttura.

La chiave di accorpamento e' l'intera pila: scegliendo piu' strati della stessa struttura si ottiene un foro solo.
Il report conta gli attraversamenti allungati e, nella colonna "Crossed element", elenca gli strati aggiunti.

## Place: modalita' dalla selezione

- **Elementi MEP gia' selezionati** al click: modalita' inversa. Si cercano muri, solai e strutture (travi,
  pilastri, fondazioni) del modello host e dei link caricati che toccano il bounding box di almeno un elemento
  scelto, poi la stessa rilevazione della modalita' normale, limitata agli elementi scelti.
- **Nessuna selezione**: una finestra chiede se selezionare muri/solai/strutture o elementi MEP.
- **Strutture** (anche nella modalita' normale e nell'ambito del Refresh): foro a solaio per elementi MEP verticali
  (oltre 45 gradi), a muro per gli altri, con asse perpendicolare alla trave o, per i pilastri, lungo l'elemento MEP;
  un attraversamento piu' lungo di 1 m e' scartato (l'elemento ci corre dentro). Niente allungamento per strati.

## Elementi sovrapposti (trave nel muro, pilastro nel muro...)

Prima dell'accorpamento, gli attraversamenti dello **stesso elemento MEP** attraverso elementi diversi che si
sovrappongono o distano al massimo l'intercapedine di Settings lungo l'elemento (stesso tipo di foro, assi
paralleli) diventano uno solo: intervallo complessivo, tutti gli elementi attraversati, sistema del foro preso da
muro o solaio se c'e'. Nell'accorpamento due attraversamenti stanno nello stesso foro se hanno almeno un elemento
attraversato in comune (non serve piu' la stessa identica pila).

## Refresh: selezione ed elementi fuori dal foro

- **Ambito dalla selezione:** con elementi MEP o fori gia' selezionati si controllano solo i fori legati a quelli
  (codici PFV sul MEP, Id del MEP nel foro, o il foro stesso). Senza selezione si chiede come prima.
- **Elementi fuori dal foro** (in parte fuori, o spostati oltre la distanza di unione ma ancora attraversanti con
  lo stesso orientamento): lista con spunta dei fori interessati, si scelgono quelli da **spostare sui loro
  elementi** (gli elementi MEP non si toccano, le misure del foro restano). Spostamento nel piano del foro:
  centrato sugli elementi se il foro e' di un solo elemento, tondo, o se non ci stanno comunque; altrimenti il
  minimo perche' tutti rientrino con il franco. E' una transazione a parte (un Undo); poi i fori si ricontrollano
  (un foro spostato ma ancora troppo piccolo risulta Changed).
