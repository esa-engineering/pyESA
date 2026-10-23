# -*- coding: utf-8 -*-
"""Controlli comuni delle finestre di Branch Tools: lettura dei numeri, righe di stato, slider con casella.
Nessun documento a livello di modulo (Rocket mode)."""
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Visibility
from System.Windows.Input import Key


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


def step_text(value):
    return "{:g}".format(value)


def set_status(text_block, text):
    """Riga di stato visibile solo se c'e' qualcosa da dire: niente testo grigio fisso nella finestra."""
    text_block.Text = text or ""
    text_block.Visibility = Visibility.Visible if text else Visibility.Collapsed


def on_enter_or_leave(box, apply):
    """apply() quando si preme Invio nella casella (senza premere OK) o quando la casella perde il fuoco."""
    def on_key(sender, args):
        if args.Key in (Key.Enter, Key.Return):
            args.Handled = True
            apply()
    box.PreviewKeyDown += on_key
    box.LostFocus += lambda s, a: apply()


class SliderSection(object):
    """Slider con casella del valore, passo, pulsanti - / + / ritorno e riga di stato.
    prefix: prefisso dei nomi dei controlli nel XAML (Rotate, Angle, Sloped);
    apply(valore) cambia il modello e ritorna (valore effettivo, errore o None);
    describe(valore) ritorna il testo della riga di stato ("" = riga nascosta)."""

    def __init__(self, window, prefix, steps, step, start, apply, describe, enabled):
        self.apply = apply
        self.describe = describe
        self.start = start
        self._updating = False      # True mentre il codice imposta slider e casella: niente modifica al modello
        find = lambda name: window.FindName(prefix + name)
        self.slider = find("Slider")
        self.box = find("Box")
        self.step_combo = find("StepCombo")
        self.status = find("Status")
        for value in steps:
            self.step_combo.Items.Add(step_text(value))
        self.step_combo.SelectedItem = step_text(step if step in steps else steps[1])
        self.slider.TickFrequency = self.step()
        self.slider.ValueChanged += self._on_slider
        self.step_combo.SelectionChanged += self._on_step
        on_enter_or_leave(self.box, self._apply_box)
        find("MinusButton").Click += lambda s, a: self.go_to(self.slider.Value - self.step())
        find("PlusButton").Click += lambda s, a: self.go_to(self.slider.Value + self.step())
        find("ResetButton").Click += lambda s, a: self.go_to(self.start)
        self.controls = [find(name) for name in
                         ("Slider", "Box", "StepCombo", "MinusButton", "PlusButton", "ResetButton")]
        self.reset(start, enabled)

    def reset(self, start, enabled):
        """Nuovo valore iniziale (es. dopo Create branch) e stato attivo / disattivo."""
        self.start = start
        self.set_enabled(enabled)
        self._set_controls(start, None)

    def set_enabled(self, enabled):
        for control in self.controls:
            control.IsEnabled = enabled

    def value(self):
        return self.slider.Value

    def step(self):
        # in Revit "{:g}" segue le impostazioni regionali: con l'italiano il passo e' "2,5"
        ok, value = parse_number(self.step_combo.SelectedItem or "")
        return value if ok and value > 0 else 5.0

    def _set_controls(self, value, error):
        self._updating = True
        try:
            self.slider.Value = value
            self.box.Text = "{:g}".format(round(value, 2))
        finally:
            self._updating = False
        text = self.describe(value)
        if error:
            # niente MessageBox: trascinando lo slider ne comparirebbe uno a ogni passo
            text = (text + " " if text else "") + "Last change refused: {}".format(error)
        set_status(self.status, text)

    def go_to(self, value):
        """Porta il modello al valore (entro i limiti dello slider) e allinea i controlli al valore effettivo."""
        value = max(self.slider.Minimum, min(self.slider.Maximum, value))
        applied, error = self.apply(value)
        self._set_controls(applied, error)

    def show_error(self, error):
        self._set_controls(self.slider.Value, error)

    def _on_slider(self, sender, args):
        if not self._updating:
            self.go_to(args.NewValue)

    def _on_step(self, sender, args):
        self.slider.TickFrequency = self.step()

    def _apply_box(self):
        ok, value = parse_number(self.box.Text)
        if not ok:
            self.show_error("not a number.")
            return
        if abs(value - self.slider.Value) > 1e-6:
            # lo slider aggancia le tacche del passo: la casella imposta il valore esatto
            self.go_to(value)



