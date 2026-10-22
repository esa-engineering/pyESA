# Suggerimenti per CLAUDE.md

Cose emerse lavorando con Claude Code che oggi **non sono scritte in `CLAUDE.md`**, raccolte qui perche' il proprietario decida se promuoverle a regola. Non sono regole in vigore finche' non entrano in `CLAUDE.md`.

Voce aperta il 22.10.2026, durante lo sviluppo di FlexConnect.

## 1. "How to use" nelle finestre di impostazioni

**Cosa manca**: `CLAUDE.md` dice che la finestra serve a compilare campi e che metodo e criteri vanno nel Readme e nella wiki (sezione "Notes inside windows"). Non dice come offrire all'utente un aiuto breve dentro la finestra. Clash Setup ha un'icona `i` con un tooltip "How to use": e' un modello gia' in uso, ma non e' descritto da nessuna parte.

**Regola proposta** (per le finestre con piu' di qualche campo, o con un comportamento non ovvio):

- **Icona**: un `Border` 22 x 22 con una `i` in grassetto (Georgia, 14), bordo `#444444` spessore 1.5, `CornerRadius="3"`, `Cursor="Help"`, in alto a destra della testata della finestra (stessa riga del titolo, titolo centrato).
- **Contenuto**: un `ToolTip` ricco al passaggio del mouse, non una finestra separata ne' un pulsante `?` (la guida a slide di Clash Finder e' un'altra cosa, vedi la nota di progetto sulla guida in attesa).
- **Tempi**: `ToolTipService.InitialShowDelay="0"` e `ToolTipService.ShowDuration="120000"`, perche' il testo e' lungo e deve restare leggibile finche' il mouse e' sopra.
- **Larghezza**: `MaxWidth="560"` sul `ToolTip`, testo con `TextWrapping="Wrap"`.
- **Titoli**: `TextBlock` in grassetto con `Foreground="#2D5A8A"` (lo stesso blu di `SectionHeader`), maiuscoli.
- **Testo**: inglese, frasi brevi. Prima "HOW TO USE" con i passi numerati (cosa si clicca e in che ordine, cosa fa SHIFT + CLICK, come si esce), poi una sezione breve per ogni comportamento che sorprende, poi "SETTINGS" con una riga per campo non ovvio.
- **Niente scroll**: il tooltip deve stare sullo schermo; se non ci sta si accorcia, o si manda il resto nel Readme.

**Manutenzione**: il testo del tooltip descrive il comportamento di oggi. Chi cambia il comportamento di un tool lo aggiorna nella stessa PR, come il Readme.

**Come si concilia con le regole esistenti**: non e' un `HintText` (quelli restano al massimo 2 o 3, sotto il campo) ne' un tooltip di campo (una frase di 100 o 120 caratteri); e' un elemento a parte. Il tooltip dell'icona puo' ripetere in sintesi il Readme, ma non lo sostituisce.

**Modello da copiare**: la testata di `FlexConnect.pushbutton/FlexConnect_Form.xaml` (riga con titolo centrato e icona a destra).

## 2. Da decidere

- Le finestre senza campi da compilare (elenchi di scelta, report) hanno bisogno dell'icona? Proposta: no, solo le finestre di impostazioni. 
- Il tooltip va tradotto in italiano nella wiki (come le altre guide)? Proposta: si, la wiki resta in italiano e il tooltip in inglese.
- Il "?" e la guida a slide di Clash Finder e l'icona `i` sono due strumenti diversi: tenerli entrambi o scegliere?
