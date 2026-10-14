# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox
from System.Windows.Markup import XamlReader

# angoli dall'orizzontale proposti nella finestra (nome del RadioButton); 0 = stacco di lato, senza cambio di quota
ANGLES = ((0.0, "Ang0"), (15.0, "Ang15"), (22.5, "Ang22"), (30.0, "Ang30"), (45.0, "Ang45"), (90.0, "Ang90"))
SIZE_MODES = ("same", "rect", "round")  # misura dello stacco dei canali
MIN_DUCT_MM = 50.0  # misura minima di uno stacco di canale


def parse_number(text):
    """(ok, valore): accetta anche la virgola decimale."""
    try:
        return True, float(text.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return False, None


def parse_rect(text):
    """"LxH" (mm) salvato in config -> (larghezza, altezza) in mm, o None se vuoto o non leggibile."""
    try:
        width, height = [float(v) for v in str(text).lower().replace(",", ".").split("x")]
        return (width, height) if width > 0 and height > 0 else None
    except (ValueError, AttributeError):
        return None


def parse_round(text):
    """Diametro (mm) salvato in config -> numero in mm, o None se vuoto o non leggibile."""
    ok, value = parse_number(str(text))
    return value if ok and value > 0 else None


class HorizontalWindow(object):
    """Unica finestra di Horizontal Branch (a ogni CLICK): angolo (0 = di lato, 15, 22.5, 30, 45, 90), ritorno alla
    quota di partenza, offset, lunghezza dell'orizzontale e, per i canali, innesto (tap o T) e misura dello stacco
    (come il canale, rettangolare o circolare). show() ritorna {length_mm, drop_mm, angle_deg, return_back,
    duct_junction, duct_size_mode, duct_rect ("LxH" o ""), duct_round ("D" o "")} o None."""

    def __init__(self, xaml_path, settings):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        find = self.window.FindName
        self.length_box = find("LengthBox")
        self.offset_box = find("OffsetBox")
        self.angle_radios = [(angle, find(name)) for angle, name in ANGLES]
        self.return_check = find("ReturnCheck")
        self.tap_radio = find("DuctTap")
        self.tee_radio = find("DuctTee")
        self.size_radios = {"same": find("SizeSame"), "rect": find("SizeRect"), "round": find("SizeRound")}
        self.width_box = find("DuctWidthBox")
        self.height_box = find("DuctHeightBox")
        self.diameter_box = find("DuctDiameterBox")

        self.return_check.IsChecked = bool(settings.get("return_back", False))
        self.length_box.Text = "{:g}".format(settings["length_mm"])
        self.offset_box.Text = "{:g}".format(settings["drop_mm"])
        # angolo salvato non piu' proposto (es. il vecchio "custom"): il piu' vicino tra quelli della finestra
        saved = float(settings["angle_deg"])
        min(self.angle_radios, key=lambda item: abs(item[0] - saved))[1].IsChecked = True
        if settings.get("duct_junction") == "tee":
            self.tee_radio.IsChecked = True
        else:
            self.tap_radio.IsChecked = True
        mode = settings.get("duct_size_mode")
        self.size_radios[mode if mode in SIZE_MODES else "same"].IsChecked = True
        rect = parse_rect(settings.get("duct_rect", ""))
        if rect:
            self.width_box.Text, self.height_box.Text = "{:g}".format(rect[0]), "{:g}".format(rect[1])
        diameter = parse_round(settings.get("duct_round", ""))
        if diameter:
            self.diameter_box.Text = "{:g}".format(diameter)

        for _, radio in self.angle_radios:
            radio.Checked += self._update_enabled
        for radio in self.size_radios.values():
            radio.Checked += self._update_enabled
        self._update_enabled(None, None)
        find("OkButton").Click += self._on_ok
        self.result = None

    def _angle(self):
        for angle, radio in self.angle_radios:
            if radio.IsChecked:
                return angle
        return 90.0

    def _size_mode(self):
        for mode in SIZE_MODES:
            if self.size_radios[mode].IsChecked:
                return mode
        return "same"

    def _update_enabled(self, sender, args):
        """A 0 gradi lo stacco e' di lato: offset e ritorno non servono. Le misure solo per la forma scelta."""
        side = self._angle() == 0.0
        self.offset_box.IsEnabled = not side
        self.return_check.IsEnabled = not side
        mode = self._size_mode()
        self.width_box.IsEnabled = self.height_box.IsEnabled = mode == "rect"
        self.diameter_box.IsEnabled = mode == "round"

    def _validate(self):
        """Ritorna (ok, messaggio, valori)."""
        ok, length = parse_number(self.length_box.Text)
        if not ok or length < 50:
            return False, "The horizontal length must be a number of at least 50 mm.", None
        angle = self._angle()
        if angle == 0.0:
            # di lato: offset e ritorno non si usano (lo script li ignora), ma restano salvati per la prossima volta
            ok, offset = parse_number(self.offset_box.Text)
            offset = offset if ok and (offset == 0 or abs(offset) >= 50) else 0.0
            return_back = bool(self.return_check.IsChecked) and offset != 0
        else:
            ok, offset = parse_number(self.offset_box.Text)
            if not ok or (offset != 0 and abs(offset) < 50):
                return False, ("The offset must be 0, or a number of at least 50 mm in absolute value "
                               "(+ up, - down)."), None
            return_back = bool(self.return_check.IsChecked)
            if return_back and offset == 0:
                return False, "The return to the starting level needs a vertical offset other than 0.", None
        mode = self._size_mode()
        rect, round_ = "", ""
        if mode == "rect":
            ok_w, width = parse_number(self.width_box.Text)
            ok_h, height = parse_number(self.height_box.Text)
            if not (ok_w and ok_h) or width < MIN_DUCT_MM or height < MIN_DUCT_MM:
                return False, "Rectangular branch: width and height must be at least {:g} mm.".format(
                    MIN_DUCT_MM), None
            rect = "{:g}x{:g}".format(width, height)
        elif mode == "round":
            ok, diameter = parse_number(self.diameter_box.Text)
            if not ok or diameter < MIN_DUCT_MM:
                return False, "Circular branch: the diameter must be at least {:g} mm.".format(MIN_DUCT_MM), None
            round_ = "{:g}".format(diameter)
        else:
            # misure non usate ma ricordate per la prossima volta
            rect = "{}x{}".format(self.width_box.Text.strip(), self.height_box.Text.strip())
            rect = rect if parse_rect(rect) else ""
            round_ = self.diameter_box.Text.strip() if parse_round(self.diameter_box.Text) else ""
        return True, "", {"length_mm": length, "drop_mm": offset, "angle_deg": angle, "return_back": return_back,
                          "duct_junction": "tee" if self.tee_radio.IsChecked else "tap",
                          "duct_size_mode": mode, "duct_rect": rect, "duct_round": round_}

    def _on_ok(self, sender, args):
        ok, message, values = self._validate()
        if not ok:
            MessageBox.Show(message, "Horizontal Branch")
            return
        self.result = values
        self.window.DialogResult = True

    def show(self):
        self.window.ShowDialog()
        return self.result
