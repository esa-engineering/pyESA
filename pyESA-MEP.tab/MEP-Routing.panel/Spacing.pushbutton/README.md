# Spacing

Imposta la stessa distanza tra tubi, canali, passerelle e conduit paralleli, isolamento compreso,
e controlla che gli accessori collegati non vadano in conflitto.

## Uso

- **CLICK**: selezione (anche prima) di almeno due elementi, poi finestra con distanza, modo di
  misura e linea che resta ferma. Config `ESA_Spacing`.
- **SHIFT + CLICK**: copia le distanze da una serie di riferimento a un'altra (vedi sotto).
- Vista attiva: pianta (elementi affiancati) oppure sezione / prospetto (elementi uno sopra
  l'altro o affiancati). In 3D il comando si rifiuta.

## Logica

1. **Parallelismo**: gli elementi devono stare entro 10 gradi dalla direzione media; gli altri
   sono esclusi e nominati nel report.
2. **Asse di spostamento u**: nel piano della vista e perpendicolare agli elementi, cioe'
   `ViewDirection x direzione media`, orientato verso destra (o l'alto se verticale). Se gli
   elementi sono quasi perpendicolari alla vista (visti come punti) u e' la direzione tra i due
   punti piu' lontani.
3. **Linee**: tratti con i centri a meno di 1 mm lungo u sono la stessa linea (es. due tratti ai
   lati di una valvola) e si muovono insieme.
4. **Ingombro lungo u**: tubi, conduit e canali tondi: raggio esterno; sezioni rettangolari: vertici
   del solido proiettati su u (esatto per i box); piu' lo spessore dell'isolamento
   (`InsulationLiningBase.GetInsulationIds`, `.Thickness`) su entrambi i lati. Con misura sugli assi
   l'isolamento e' ignorato.
5. **Spostamenti**: le linee si ordinano lungo u e si posizionano in sequenza partendo da quella
   ferma (prima, ultima o centro del gruppo). Si sposta solo lungo u, le altre componenti restano.
6. **Transazione unica** con `RevitErrors`: se Revit rifiuta, non cambia nulla e il messaggio va nel
   report.
7. **Accessori** (valvole e serrande, categorie OST_PipeAccessory e OST_DuctAccessory) collegati
   direttamente agli elementi: dopo lo spostamento, se non hanno seguito il tratto, vengono spostati
   dello stesso vettore. Poi `ElementIntersectsElementFilter` li confronta con gli altri elementi
   spostati, il loro isolamento e gli altri accessori (esclusi i propri tratti e il loro isolamento).
   I conflitti si segnalano nel report e si selezionano, ma lo spostamento non si annulla.
8. **Verifica**: le distanze finali sono rilette dalla geometria; con scarto oltre 1 mm il report
   scrive CHECK.

## Copia dello spacing (SHIFT + CLICK)

1. Selezione della serie di riferimento, poi della serie da modificare (due serie separate).
2. Entrambe passano da `analyse()` (parallelismo, asse u, linee). Se il numero di linee e' diverso
   non cambia nulla e il report lo dice.
3. Le distanze tra linee consecutive del riferimento (`copied_gaps`, nella misura salvata con il
   CLICK: facce con isolamento o assi) diventano le distanze volute della seconda serie, una per
   coppia (`plan_shifts` accetta una lista). La linea ferma e' quella salvata con il CLICK.
4. Le linee si abbinano in ordine lungo u (sinistra-destra, basso-alto nella vista): se la seconda
   serie e' girata di 180 gradi rispetto alla prima l'ordine delle distanze si inverte.
5. Resto invariato: transazione unica, accessori, controllo dei conflitti, verifica delle distanze.

Con misura sulle facce le distanze dipendono dalle dimensioni: serie con diametri diversi ottengono
lo stesso vuoto ma passi diversi. Con misura sugli assi si copia il passo.

## Punti da verificare in Revit

- Se `Location.Move` su un tratto collegato a valvole e raccordi li trascini correttamente: lo
  script prova a recuperare gli accessori, ma i raccordi (gomiti, T) non sono gestiti.
- Lettura dei vertici del solido per canali rettangolari ruotati e passerelle.
- Falsi positivi dei conflitti con geometria a contatto (distanza 0 con isolamento).
