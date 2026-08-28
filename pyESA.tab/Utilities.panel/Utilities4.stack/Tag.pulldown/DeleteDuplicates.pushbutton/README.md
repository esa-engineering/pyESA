# Delete Duplicate Tags — strumento pyRevit

Analizza le viste scelte dall'utente, segnala i **tag duplicati** (piu di un tag che
punta allo stesso elemento nella stessa vista) e cancella solo quelli confermati.

## Contenuto della cartella

| File | Ruolo |
| --- | --- |
| `DeleteDuplicates_script.py` | Logica dello strumento (Revit API, motore IronPython) |
| `DeleteDuplicatesUI.xaml` | Interfaccia grafica della finestra |
| `bundle.yaml` | Etichetta e descrizione del pulsante |
| `icon.png` / `icon.dark.png` | Icone tema chiaro e tema scuro (96x96) |
| `README.md` | Questo file |

## Come funziona

1. Alla partenza lo script raccoglie in una sola passata **tutti** i tag del modello
   (tag standard che derivano da `IndependentTag`, piu tag di locale, di spazio MEP e
   di area) e li raggruppa per vista di appartenenza (`OwnerViewId`). Nella finestra
   compaiono quindi solo le viste che contengono almeno un tag, con il conteggio.
2. L'utente seleziona le viste e imposta il criterio di duplicazione.
3. **Analizza** confronta i tag vista per vista. Il confronto non modifica il modello.
4. I gruppi trovati compaiono nell'elenco dei risultati, tutti preselezionati. Si
   possono deselezionare i gruppi da lasciare intatti.
5. **Elimina selezionati** chiede conferma, chiude la finestra ed esegue la
   cancellazione in una singola transazione (annullabile con Undo di Revit).
6. Il report completo viene stampato nel pannello di output di pyRevit, con i link
   cliccabili ai tag mantenuti.

Se la finestra viene chiusa senza eliminare nulla ma l'analisi e' stata eseguita, il
report viene stampato lo stesso: serve a documentare i duplicati senza toccare il modello.

## Criteri di duplicazione

| Criterio | Effetto |
| --- | --- |
| Stesso elemento taggato **e** stesso tipo di tag (default) | Due tag sono duplicati solo se sono della stessa famiglia e tipo |
| Stesso elemento taggato, **qualsiasi** tipo di tag | Piu aggressivo: due tag di famiglie diverse sullo stesso elemento (per esempio codice e dimensione) risultano duplicati |

Opzioni:

- **Includi i tag di elementi linkati** (default attivo): se disattivata, i tag che
  puntano a elementi di un modello linkato non vengono nemmeno confrontati.
- **Includi i tag bloccati (pin)** (default spento): se attivata, i tag pinnati vengono
  sbloccati prima di essere eliminati.
- **Solo se sovrapposti entro N mm** (default spento): restringe il duplicato ai tag la
  cui testa e' praticamente nella stessa posizione. Utile per i doppioni nati da un
  copia/incolla in posto, evita di toccare tag volutamente ripetuti in punti diversi
  della vista. Un tag di cui non si riesce a leggere la posizione della testa resta
  isolato e non viene eliminato: il report lo dichiara.
- **Tag da mantenere**: il piu vecchio (Id piu basso, default) oppure il piu recente.

Un tag multi-riferimento (Revit 2022+) e' duplicato di un altro solo se punta
esattamente allo stesso insieme di elementi.

## Note tecniche

- I tag senza elemento associato (host cancellato) non vengono mai eliminati: sono
  contati a parte nel report.
- Gli `ElementId` dei tag da eliminare vengono memorizzati **prima** della transazione,
  perche dopo la cancellazione gli oggetti `Element` non sono piu interrogabili.
- L'eliminazione parte come `Document.Delete` in blocco; se fallisce (per esempio un
  elemento posseduto da un altro utente in un modello condiviso) lo script ripiega su
  una cancellazione elemento per elemento e riporta in tabella quelli non eliminabili
  con il relativo motivo.
- Compatibilita 2022-2026: il valore degli `ElementId` viene letto con l'helper
  `get_element_id_value`, che usa `.Value` su Revit 2026 e `.IntegerValue` sulle
  versioni precedenti.
