# Match Top / Match Middle / Match Bottom

I tre pulsanti dello stack `MatchElevation` allineano in quota tubi, canali, passerelle e conduit
a un elemento di riferimento. Gli script sono identici tranne la costante `MATCH_MODE`
(`"top"`, `"middle"`, `"bottom"`): una correzione va riportata in tutte e tre le copie.

## Uso

- **CLICK**: clic sull'elemento di riferimento, poi selezione (clic o finestra) degli elementi da
  allineare e **Finish**.
- **SHIFT + CLICK**: impostazioni condivise dai tre pulsanti (config `ESA_MatchElevation`):
  - *Overwrite the reference level* (predefinito: si): gli elementi prendono il Reference Level
    del primo;
  - *Remove the slope* (predefinito: no): gli elementi inclinati diventano orizzontali.

## Quote allineate

Quote assolute in coordinate interne, calcolate sulla geometria (non sui parametri di Revit, che
dipendono dal livello):

| Pulsante | Punto allineato |
| --- | --- |
| Match Top | estremo piu' alto dell'asse + meta' ingombro verticale della sezione |
| Match Middle | orizzontale: quota dell'asse; inclinato: Upper o Lower End Centerline Elevation, scelta nel popup; verticale: centro dell'asse |
| Match Bottom | estremo piu' basso dell'asse - meta' ingombro verticale della sezione |

Ingombro della sezione: diametro esterno per tubi, conduit e canali tondi (moltiplicato per il
coseno della pendenza); per sezioni rettangolari/ovali larghezza e altezza del connettore
proiettate sulla verticale, scegliendo l'abbinamento con gli assi del connettore piu' vicino al
bounding box. L'isolamento non e' compreso (come nei parametri Top/Bottom Elevation di Revit).

Match Middle: se il riferimento o uno degli elementi selezionati e' inclinato, un popup
(`forms.CommandSwitchWindow`) chiede se allineare la quota d'asse dell'estremo alto o di quello
basso (`MIDDLE_END`). Se e' inclinato il riferimento la scelta si fa prima della selezione, cosi'
il prompt mostra gia' la quota giusta. La scelta vale per il riferimento e per tutti gli elementi.

## Elementi verticali (scostamento dall'asse Z entro 5 gradi)

- Come riferimento: ingombro della sezione nullo, quindi top = estremo superiore, middle = centro,
  bottom = estremo inferiore.
- Come elementi da allineare: spostati in blocco; l'opzione *Remove the slope* non si applica
  (renderli orizzontali li annullerebbe).

## Applicazione

1. Per ogni elemento si calcola un piano (nuovo livello, nuovo asse in coordinate assolute) senza
   toccare il modello; quelli gia' allineati non si modificano.
2. Tutti i piani si applicano in un'unica transazione: prima i livelli, poi le geometrie, senza
   rigenerare in mezzo. Cosi' i tratti selezionati e collegati tra loro (es. con un manicotto) si
   spostano insieme invece di trascinarsi a vicenda.
3. Se Revit rifiuta l'insieme, si riprova un elemento alla volta (ricalcolando il piano), e il
   report lo segnala.
4. Dopo il commit si rilegge la quota: se un elemento e' finito a piu' di 1 mm dal riferimento
   (trascinato da un elemento collegato), il report lo segnala con CHECK.

Gli errori di Revit al commit non aprono finestre: la transazione si annulla e il messaggio va
nel report (`RevitErrors`, come in Branch e RiserTurn). Il report segue il pulsante Report Mode.

## Punti da verificare in Revit

- Abbinamento larghezza/altezza agli assi del connettore per canali e passerelle ruotati.
- Comportamento del cambio di Reference Level: lo script imposta comunque la posizione assoluta
  dopo il cambio, quindi il risultato non dipende dal fatto che Revit mantenga la quota o l'offset.
