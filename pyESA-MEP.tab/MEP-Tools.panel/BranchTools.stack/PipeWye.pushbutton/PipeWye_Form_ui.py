# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows.Markup import XamlReader

from BranchTools_Controls import parse_number, set_status

ROUTING_LABEL = "(routing preferences)"
DEFAULT_SETTINGS = {"wye_roll": 0.0, "wye_angle": 45.0}


class PipeWyeWindow(object):
    """Finestra di Pipe Wye.
    main_pipe: BranchTools_Pipe.MainPipe; wye_symbols: [(etichetta "Famiglia : Tipo", FamilySymbol)] dei raccordi
    a T, a croce o a braga; settings: {"wye_roll", "wye_angle": gradi della braga da inserire; "wye_family":
    etichetta della famiglia, "" = routing preferences}.
    show() ritorna lo stesso dict aggiornato con "action" (None, "connect_above", "connect_side") per quello che
    lo script deve fare dopo la chiusura, o None se annullata."""

    def __init__(self, xaml_path, main_pipe, wye_symbols, settings):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        find = self.window.FindName
        self.result = None
        self.main_pipe = main_pipe
        self.settings = settings

        self.roll_box = find("RollBox")
        self.angle_box = find("AngleBox")
        self.roll_box.Text = "{:g}".format(settings["wye_roll"])
        self.angle_box.Text = "{:g}".format(settings["wye_angle"])
        self.place_button = find("PlaceWyeButton")
        self.place_button.Click += self._on_place_wye
        find("ConnectAboveButton").Click += lambda s, a: self._close_for("connect_above")
        find("ConnectSideButton").Click += lambda s, a: self._close_for("connect_side")
        self.reverse_box = find("ReverseFlowBox")
        self.reverse_box.Checked += self._on_reverse
        self.reverse_box.Unchecked += self._on_reverse
        self.status = find("PipeStatus")

        # famiglia della braga: la prima voce lascia scegliere alle routing preferences
        self.wye_combo = find("WyeFamilyCombo")
        self.wye_ids = {}
        self.wye_combo.Items.Add(ROUTING_LABEL)
        for label, symbol in wye_symbols:
            self.wye_combo.Items.Add(label)
            self.wye_ids[label] = symbol.Id
        saved = settings.get("wye_family", "")
        self.wye_combo.SelectedItem = saved if saved in self.wye_ids else ROUTING_LABEL
        self.wye_combo.SelectionChanged += self._on_wye_family
        self._on_wye_family(None, None)
        self._show_flow()

        find("OkButton").Click += self._on_ok
        find("CancelButton").Click += self._on_cancel

    # -------------------------------------------------------------- tubo

    def _wye_label(self):
        label = self.wye_combo.SelectedItem
        return label if label in self.wye_ids else ""

    def _on_wye_family(self, sender, args):
        label = self._wye_label()
        self.main_pipe.symbol_id = self.wye_ids[label] if label else None

    def _show_flow(self, text=None):
        flow = "Flow: {}.".format(self.main_pipe.directions()[3])
        set_status(self.status, "{} {}".format(text, flow) if text else flow)

    def _on_reverse(self, sender, args):
        self.main_pipe.reverse = bool(self.reverse_box.IsChecked)
        self._show_flow()

    def _wye_values(self):
        """(ok, rotazione, angolo) della braga da inserire."""
        ok_roll, roll = parse_number(self.roll_box.Text)
        ok_angle, angle = parse_number(self.angle_box.Text)
        if not ok_roll or not -180 <= roll <= 180:
            self._show_flow("Rotation: a number from -180 to 180.")
            return False, None, None
        if not ok_angle or not 0 < angle < 180:
            self._show_flow("Angle: a number between 0 and 180.")
            return False, None, None
        return True, roll, angle

    def _on_place_wye(self, sender, args):
        ok, roll, angle = self._wye_values()
        if not ok:
            return
        fitting, error = self.main_pipe.place_wye(roll, angle)
        if error:
            self._show_flow("Not placed: {}".format(error))
            return
        self.place_button.IsEnabled = False
        self._show_flow("Wye placed. Use Side Branch on it to make the branch.")

    # -------------------------------------------------------------- OK / Cancel

    def _settings(self):
        settings = dict(self.settings)
        ok, roll, angle = self._wye_values()
        if ok:
            settings["wye_roll"] = roll
            settings["wye_angle"] = angle
        settings["wye_family"] = self._wye_label()
        return settings

    def _close_for(self, action):
        """Chiude come OK e lascia allo script un'azione che richiede di cliccare nel modello
        (non si puo' con la finestra modale aperta)."""
        self.result = self._settings()
        self.result["action"] = action
        self.window.DialogResult = True

    def _on_ok(self, sender, args):
        self.result = self._settings()
        self.result["action"] = None
        self.window.DialogResult = True

    def _on_cancel(self, sender, args):
        self.window.DialogResult = False

    def show(self):
        self.window.ShowDialog()
        return self.result
