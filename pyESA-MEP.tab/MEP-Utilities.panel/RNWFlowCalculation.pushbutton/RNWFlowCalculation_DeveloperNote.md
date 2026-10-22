# RNW Flow Calculation: note di sviluppo

Registro dei tentativi fatti sullo strumento, per non ripetere soluzioni già provate e
scartate. Non è il Readme (come funziona lo strumento oggi) né lo storico revisioni nello
script (cosa cambia in ogni versione). Voci dalla più recente.

Le versioni fino alla 4.1 sono precedenti a questo file e non hanno voci: i tentativi fatti
allora non sono registrati.

---

## 22.10.2026, versione 4.2

**Richiesta.** Il pulsante *Export Excel* diventa *Export* e apre una finestra di
impostazione: formato Excel o PDF, le tre schede della finestra dei risultati (General,
Sections, Fixtures) nel file, scelta delle colonne da esportare.

**Fatto.**
- Contenuto separato dai formati: `RNWFlowCalculation_Report.py` (Python puro) descrive le
  schede come blocchi (coppie, note, tabelle) e tiene i cataloghi delle colonne; la
  finestra dei risultati fornisce righe complete `{chiave: valore}` e `build_table()`
  tiene le colonne scelte. Così Excel e PDF hanno lo stesso contenuto e una colonna nuova
  si aggiunge in un punto solo (catalogo + valore in `_section_export_rows` /
  `_fixture_export_rows`).
- `RNWFlowCalculation_Xlsx.export_report` sostituisce `export_calculation` (un foglio
  Calculation): un foglio per scheda più Settings. Aggiunto lo stile `STYLE_INFO` (testo
  blu, OVERSIZED).
- `RNWFlowCalculation_Pdf.py`: writer PDF scritto a mano (vedi sotto le alternative).
- Finestra `RNWFlowCalculation_Export.xaml` + `_Export_ui.py`, tre colonne senza
  scorrimento; caselle create in codice con la chiave in `Tag`. Scelte salvate nella config.
- Le note della scheda General ora vengono da `general_notes()`, usata dalla finestra e
  dall'export (prima erano costruite dentro `_build_general`).
- Apertura del file esportato con `ProcessStartInfo(UseShellExecute=True)`: in .NET 8
  (Revit 2025+) `Process.Start(path)` non apre i documenti; prima l'errore era ignorato in
  silenzio.

**Scartato (PDF).**
- Stampante "Microsoft Print to PDF" (`System.Drawing.Printing`, `PrintToFile`): dipende da
  una stampante che l'IT può togliere, e l'impaginazione GDI va comunque scritta a mano.
- `XpsDocumentWriter` / `FixedDocument` di WPF: produce XPS, non PDF; la conversione
  richiede di nuovo una stampante.
- Librerie PDF .NET (PdfSharp ecc.): una DLL in più nel repo, da tenere compatibile con
  .NET Framework e .NET 8. Il writer a mano usa solo i font standard e non ha dipendenze,
  come il writer .xlsx.

**Problemi trovati provando.**
- Testo spezzato nelle colonne (`ME.RW.00` / `0`) anche quando la tabella stava in
  pagina: la colonna è larga quanto il testo più lungo più i margini, e togliendo i
  margini l'arrotondamento in virgola mobile lo mandava a capo. Risolto con una tolleranza
  di 0,01 pt in `wrap()`.
- Il primo algoritmo di restringimento divideva lo spazio in parti uguali partendo dalle
  colonne più strette, quindi spezzava i codici e lasciava larghe Status e Notes.
  Sostituito: minimo = parola più lunga (al massimo `MAX_WORD_WIDTH`), lo spazio che resta
  in proporzione a quanto manca a ogni colonna.
- Caselle di Sections in un `UniformGrid` a due colonne: l'ordine si leggeva a zig-zag.
  Ora due `StackPanel` (prima metà del catalogo a sinistra).
- Valori del riepilogo General passati come testo: in Excel restavano testo ("1446.00").
  Ora sono numeri (`Report.pair_text()` per il PDF).

**Esito.** Provato fuori da Revit con il motore IronPython 2.7 di pyRevit e dati finti:
finestra di esportazione (render, salvataggio e rilettura delle scelte), report delle tre
schede, .xlsx aperto in Excel via COM (fogli, intestazioni, numeri, filtro), PDF A4 e A3
aperto con il motore PDF di Windows (`Windows.Data.Pdf`), file bloccato riconosciuto. Non
provato in Revit, né con .NET 8 (Revit 2025+).

---

## 22.10.2026, versione 4.1 (invariata)

**Richiesta.** Allineare i nomi dei file dello strumento alle specifiche di `CLAUDE.md`
("Files of a button folder", "Naming").

**Fatto.** Rinominati con `git mv`, così la storia dei file resta collegata:

| Prima | Dopo |
| --- | --- |
| `README.md` | `RNWFlowCalculation_Readme.md` |
| `RNWFlowCalculation_form.xaml` | `RNWFlowCalculation_Form.xaml` |
| `RNWFlowCalculation_results.xaml` | `RNWFlowCalculation_Results.xaml` |
| `rnwflow_ui.py` | `RNWFlowCalculation_Form_ui.py` |
| `rnwflow_results_ui.py` | `RNWFlowCalculation_Results_ui.py` |
| `rnwflow_check.py` | `RNWFlowCalculation_Check.py` |
| `rnwflow_tables.py` | `RNWFlowCalculation_Tables.py` |
| `rnwflow_xlsx.py` | `RNWFlowCalculation_Xlsx.py` |

Aggiornati import, costanti `XAML_FILE_NAME`, docstring, commenti e Readme. Gli alias degli
import (`rt`, `rc`, `rui`, `rres`, `rx`) restano quelli di prima, quindi il corpo del codice
non cambia. Creato questo file. Invariati: cartella del pulsante, `bundle.yaml`, icone e
sezione di configurazione `ESA_RNWFlowCalculation` (già conforme: le impostazioni salvate
degli utenti restano valide).

**Nota Rocket mode.** I vecchi moduli `rnwflow_*` possono restare in `sys.modules` in una
sessione già aperta: dopo l'aggiornamento serve un Reload di pyRevit (o il riavvio di Revit).

**Esito.** Sintassi verificata fuori da Revit su tutti i `.py`. Non provato in Revit.
