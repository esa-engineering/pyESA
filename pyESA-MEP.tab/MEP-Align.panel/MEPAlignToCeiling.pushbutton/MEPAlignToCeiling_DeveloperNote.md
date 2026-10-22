# MEP Align to Ceiling: note di sviluppo

Registro dei tentativi fatti sullo strumento, per non ripetere soluzioni già provate e
scartate. Non è il README (come funziona lo strumento oggi) né lo storico revisioni nello
script (cosa cambia in ogni versione). Voci dalla più recente.

Le voci da 1.0 a 1.3 sono ricostruite a posteriori dai commit e dalle versioni del README:
riportano le scelte che risultano dal codice, non i tentativi intermedi non registrati.

---

## 22.10.2026, versione 1.4

**Richiesta.** Allineare la finestra di avvio alle regole del CLAUDE.md: niente barra di
scorrimento verticale, al massimo 2 o 3 hint brevi, niente spiegazioni del metodo nella
finestra.

**Situazione di partenza.** Una sola colonna (500 x 760) dentro uno `ScrollViewer`, una lista
delle categorie con maniglia di ridimensionamento (`Thumb`), una sezione informativa *How it
works* e hint di più righe sotto categorie, rilevatori e tolleranza. Anche i tooltip delle
opzioni spiegavano il metodo.

**Fatto.**

- Due colonne (860 x 560, minimo 760 x 500): a sinistra *Elements* e *Categories*, a destra
  *Tolerance*, *False ceiling detectors* e *Options*. Tolto lo `ScrollViewer`.
- La lista delle categorie occupa l'altezza della colonna: tolti il `Thumb` e il gestore
  `_on_resize_list`. Le categorie gestite sono al massimo nove, quindi la maniglia non serve.
- Tolta la sezione *How it works*; il contenuto era già nel README.
- Hint e testi di stato ridotti a una riga (`ask_options()` nello script,
  `_refresh_tolerance_hint()` nel modulo della finestra): la tolleranza mostra la lunghezza
  attuale della retta di analisi, o chiede un valore maggiore di zero.
- Tooltip delle opzioni accorciati; tolti quelli sulla ricerca e sul pulsante di
  cancellazione (il CLAUDE.md non li vuole).

**Esito.** Sintassi verificata fuori da Revit. **Non provato in Revit**: da controllare che
tutto stia nella finestra senza scorrimento al 100%, 125% e 150% di scala dello schermo.

**Decisione.** Tenuta la disposizione su due colonne. Il metodo resta solo nel README.

---

## 21.10.2026, versione 1.3

**Richiesta.** Rivedere il trattamento dei rilevatori da controsoffitto, rendere la finestra
di esito utilizzabile mentre si lavora in Revit, esprimere la tolleranza in mm.

**Scartato: valore di riserva di `dz_Buzzer` digitato nella finestra (versione 1.1).** Dove
sotto il dispositivo non c'era controsoffitto, `dz_Buzzer` prendeva un valore inserito
dall'utente (cm, predefinito 50). Scartato: un rilevatore da controsoffitto senza
controsoffitto sotto è la famiglia sbagliata, e un valore inventato nasconde l'errore. Ora
`dz_Buzzer` riceve solo il valore misurato e il rilevatore senza false ceiling viene saltato
come *Unsuitable detector*.

**Scartato: intradosso = solo il solaio di chiusura (versione 1.1).** Una finitura o uno
strato di isolante modellati come Ceiling e incollati al solaio venivano presi per il
controsoffitto vero. Ora l'intradosso comprende i Ceiling a contatto con la faccia inferiore
del solaio (entro `STACK_GAP_MM`, `soffit_package()`), e il false ceiling è solo un
controsoffitto staccato dall'intradosso.

**Scartato: finestra di esito modale con navigazione via `PickObjects` (versione 1.1).** Il
pulsante 🔍 chiudeva la finestra, selezionava l'elemento e lasciava navigare la vista con una
`PickObjects`, poi riapriva la finestra con lo stato salvato (`ResultState`). Limite: non si
poteva modificare l'elemento con la tabella aperta. Sostituita da una finestra non
modale: `ExternalEvent` con `ResultEventHandler`, `engine: persistent: true` nel
`bundle.yaml`, riferimenti tenuti in `OPEN_RESULT_WINDOWS`, owner sulla finestra di Revit ed
`EnableModelessKeyboardInterop` per la tastiera. Limite accettato: la tabella non si aggiorna
dopo le modifiche a mano, si rilancia il comando.

**Tolleranza.** Da cm (predefinita 50, massimo 300) a mm (predefinita 500, massimo 5000); il
valore salvato in cm viene convertito.

**Esito.** Nessuna prova in Revit registrata; i punti da verificare sono nel README (*Da
verificare in Revit*).

---

## 20.10.2026, versione 1.1

**Scartato: scelta del bersaglio nella finestra (versione 1.0).** La finestra chiedeva
*Ceilings* e/o *Floor slab soffit*: con *Ceilings* tolto si andava al solaio anche sotto un
controsoffitto, con *Floor slab soffit* tolto e nessun controsoffitto l'elemento veniva
saltato. Scartata in favore del bersaglio automatico: il controsoffitto del piano se c'è,
altrimenti l'intradosso del solaio di chiusura. Il motivo preciso dell'abbandono non è
registrato.

**Aggiunto.** Rilevatori da controsoffitto (`e_FA.DE` e `False Ceiling` nel family name) con
scrittura di `dz_Buzzer`; finestra di esito con i totali e la tabella interrogabile dei non
allineati (filtri per colonna copiati da MEPQTO in `MEPAlignToCeiling_grid_filter.py`); il
resoconto nel pannello di output diventa il report avanzato, su richiesta.

**Retta di analisi dei rilevatori.** Vale se interseca il solaio oppure il controsoffitto:
misurare solo dal solaio manderebbe oltre tolleranza quasi tutti i rilevatori, che di solito
sono modellati all'altezza del controsoffitto.

---

## 20.10.2026, versione 1.0

Prima versione, con la stessa impostazione di MEPAlignToWall ruotata di novanta gradi: retta
di analisi verticale lunga due volte la tolleranza, screening su
`u_VIS_Ceiling_installation`, ricerca in host e modelli collegati, spostamento solo in Z.
Bersaglio scelto dall'utente nella finestra (vedi 1.1).
