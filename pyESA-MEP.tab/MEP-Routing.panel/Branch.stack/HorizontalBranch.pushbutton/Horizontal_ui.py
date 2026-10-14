# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows import MessageBox, Visibility
from System.Windows.Markup import XamlReader

# angoli dall'orizzontale proposti nella finestra (nome del RadioButton); 0 = stacco di lato, senza cambio di quota
ANGLES = ((0.0, "Ang0"), (15.0, "Ang15"), (22.5, "Ang22"), (30.0, "Ang30"), (45.0, "Ang45"), (90.0, "Ang90"))
DUCT_SIZE_MODES = ("same", "rect", "round")  # misura dello stacco dei canali
MIN_SIZE_MM = 15.0  # misura minima di uno stacco (DN 15 per i tubi)
BASE_HEIGHT, SECTION_HEIGHT, MAX_HEIGHT = 430, 210, 900  # altezza della finestra secondo le sezioni visibili


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


def read_rect(width_box, height_box, label):
    """(ok, messaggio, "LxH") dai due campi."""
    ok_w, width = parse_number(width_box.Text)
    ok_h, height = parse_number(height_box.Text)
    if not (ok_w and ok_h) or width < MIN_SIZE_MM or height < MIN_SIZE_MM:
        return False, "{}: width and height must be at least {:g} mm.".format(label, MIN_SIZE_MM), ""
    return True, "", "{:g}x{:g}".format(width, height)


def read_round(box, label):
    """(ok, messaggio, "D") dal campo."""
    ok, value = parse_number(box.Text)
    if not ok or value < MIN_SIZE_MM:
        return False, "{}: the diameter must be at least {:g} mm.".format(label, MIN_SIZE_MM), ""
    return True, "", "{:g}".format(value)


def kept_rect(width_box, height_box):
    """Misura scritta ma non usata: si ricorda se leggibile."""
    text = "{}x{}".format(width_box.Text.strip(), height_box.Text.strip())
    return text if parse_rect(text) else ""


class HorizontalWindow(object):
    """Unica finestra di Horizontal Branch (a ogni CLICK). Sempre: angolo (0 = di lato, 15, 22.5, 30, 45, 90),
    ritorno alla quota di partenza, offset, lunghezza dell'orizzontale. Poi una sezione per ogni categoria presente
    nella selezione fatta prima del clic (settings["kinds"]: "duct", "pipe", "tray"), con innesto e misura dello
    stacco; le categorie non selezionate non compaiono e le loro impostazioni salvate restano come sono.
    show() ritorna il dizionario delle impostazioni aggiornato, o None."""

    def __init__(self, xaml_path, settings):
        self.settings = dict(settings)
        self.kinds = set(settings.get("kinds", ()))
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        find = self.window.FindName
        self.length_box = find("LengthBox")
        self.offset_box = find("OffsetBox")
        self.angle_radios = [(angle, find(name)) for angle, name in ANGLES]
        self.return_check = find("ReturnCheck")
        # canali
        self.duct_tap, self.duct_tee = find("DuctTap"), find("DuctTee")
        self.duct_size = {"same": find("SizeSame"), "rect": find("SizeRect"), "round": find("SizeRound")}
        self.duct_w, self.duct_h, self.duct_d = find("DuctWidthBox"), find("DuctHeightBox"), find("DuctDiameterBox")
        # tubi
        self.pipe_tee, self.pipe_tap = find("PipeTee"), find("PipeTap")
        self.pipe_same, self.pipe_custom, self.pipe_d = find("PipeSizeSame"), find("PipeSizeCustom"), find(
            "PipeDiameterBox")
        # passerelle
        self.tray_same, self.tray_custom = find("TraySizeSame"), find("TraySizeCustom")
        self.tray_w, self.tray_h = find("TrayWidthBox"), find("TrayHeightBox")

        self.return_check.IsChecked = bool(settings.get("return_back", False))
        self.length_box.Text = "{:g}".format(settings["length_mm"])
        self.offset_box.Text = "{:g}".format(settings["drop_mm"])
        # angolo salvato non piu' proposto (es. il vecchio "custom"): il piu' vicino tra quelli della finestra
        saved = float(settings["angle_deg"])
        min(self.angle_radios, key=lambda item: abs(item[0] - saved))[1].IsChecked = True

        (self.duct_tee if settings.get("duct_junction") == "tee" else self.duct_tap).IsChecked = True
        mode = settings.get("duct_size_mode")
        self.duct_size[mode if mode in DUCT_SIZE_MODES else "same"].IsChecked = True
        rect = parse_rect(settings.get("duct_rect", ""))
        if rect:
            self.duct_w.Text, self.duct_h.Text = "{:g}".format(rect[0]), "{:g}".format(rect[1])
        if parse_round(settings.get("duct_round", "")):
            self.duct_d.Text = "{:g}".format(parse_round(settings["duct_round"]))

        (self.pipe_tap if settings.get("pipe_junction") == "tap" else self.pipe_tee).IsChecked = True
        (self.pipe_custom if settings.get("pipe_size_mode") == "custom" else self.pipe_same).IsChecked = True
        if parse_round(settings.get("pipe_diameter", "")):
            self.pipe_d.Text = "{:g}".format(parse_round(settings["pipe_diameter"]))

        (self.tray_custom if settings.get("tray_size_mode") == "custom" else self.tray_same).IsChecked = True
        rect = parse_rect(settings.get("tray_rect", ""))
        if rect:
            self.tray_w.Text, self.tray_h.Text = "{:g}".format(rect[0]), "{:g}".format(rect[1])

        # una sezione per categoria selezionata, nessuna se non c'e' selezione
        sections = (("duct", "DuctSection"), ("pipe", "PipeSection"), ("tray", "TraySection"))
        for kind, name in sections:
            find(name).Visibility = Visibility.Visible if kind in self.kinds else Visibility.Collapsed
        find("NoSelectionHint").Visibility = Visibility.Collapsed if self.kinds else Visibility.Visible
        self.window.Height = min(MAX_HEIGHT, BASE_HEIGHT + SECTION_HEIGHT * len(self.kinds))

        radios = [r for _, r in self.angle_radios] + list(self.duct_size.values()) + [
            self.pipe_same, self.pipe_custom, self.tray_same, self.tray_custom]
        for radio in radios:
            radio.Checked += self._update_enabled
        self._update_enabled(None, None)
        find("OkButton").Click += self._on_ok
        self.result = None

    def _angle(self):
        for angle, radio in self.angle_radios:
            if radio.IsChecked:
                return angle
        return 90.0

    def _duct_mode(self):
        for mode in DUCT_SIZE_MODES:
            if self.duct_size[mode].IsChecked:
                return mode
        return "same"

    def _update_enabled(self, sender, args):
        """A 0 gradi lo stacco e' di lato: offset e ritorno non servono. Le misure solo per la scelta fatta."""
        side = self._angle() == 0.0
        self.offset_box.IsEnabled = not side
        self.return_check.IsEnabled = not side
        mode = self._duct_mode()
        self.duct_w.IsEnabled = self.duct_h.IsEnabled = mode == "rect"
        self.duct_d.IsEnabled = mode == "round"
        self.pipe_d.IsEnabled = bool(self.pipe_custom.IsChecked)
        self.tray_w.IsEnabled = self.tray_h.IsEnabled = bool(self.tray_custom.IsChecked)

    def _validate(self):
        """Ritorna (ok, messaggio, impostazioni)."""
        values = dict(self.settings)
        ok, length = parse_number(self.length_box.Text)
        if not ok or length < 50:
            return False, "The horizontal length must be a number of at least 50 mm.", None
        angle = self._angle()
        ok, offset = parse_number(self.offset_box.Text)
        if angle == 0.0:
            # di lato: offset e ritorno non si usano (lo script li ignora), ma restano salvati per la prossima volta
            offset = offset if ok and (offset == 0 or abs(offset) >= 50) else 0.0
            return_back = bool(self.return_check.IsChecked) and offset != 0
        else:
            if not ok or (offset != 0 and abs(offset) < 50):
                return False, ("The offset must be 0, or a number of at least 50 mm in absolute value "
                               "(+ up, - down)."), None
            return_back = bool(self.return_check.IsChecked)
            if return_back and offset == 0:
                return False, "The return to the starting level needs a vertical offset other than 0.", None
        values.update({"length_mm": length, "drop_mm": offset, "angle_deg": angle, "return_back": return_back})

        # solo le categorie visibili: le altre restano come salvate
        if "duct" in self.kinds:
            mode = self._duct_mode()
            rect, round_ = kept_rect(self.duct_w, self.duct_h), self.duct_d.Text.strip()
            round_ = round_ if parse_round(round_) else ""
            if mode == "rect":
                ok, message, rect = read_rect(self.duct_w, self.duct_h, "Rectangular duct branch")
                if not ok:
                    return False, message, None
            elif mode == "round":
                ok, message, round_ = read_round(self.duct_d, "Circular duct branch")
                if not ok:
                    return False, message, None
            values.update({"duct_junction": "tee" if self.duct_tee.IsChecked else "tap",
                           "duct_size_mode": mode, "duct_rect": rect, "duct_round": round_})
        if "pipe" in self.kinds:
            mode = "custom" if self.pipe_custom.IsChecked else "same"
            diameter = self.pipe_d.Text.strip() if parse_round(self.pipe_d.Text) else ""
            if mode == "custom":
                ok, message, diameter = read_round(self.pipe_d, "Pipe branch")
                if not ok:
                    return False, message, None
            values.update({"pipe_junction": "tap" if self.pipe_tap.IsChecked else "tee",
                           "pipe_size_mode": mode, "pipe_diameter": diameter})
        if "tray" in self.kinds:
            mode = "custom" if self.tray_custom.IsChecked else "same"
            rect = kept_rect(self.tray_w, self.tray_h)
            if mode == "custom":
                ok, message, rect = read_rect(self.tray_w, self.tray_h, "Cable tray branch")
                if not ok:
                    return False, message, None
            values.update({"tray_size_mode": mode, "tray_rect": rect})
        return True, "", values

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
