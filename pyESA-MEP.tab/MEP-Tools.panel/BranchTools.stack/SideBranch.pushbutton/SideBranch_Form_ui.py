# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows.Markup import XamlReader

from BranchTools_Controls import SliderSection, on_enter_or_leave, parse_number, set_status, step_text
from BranchTools_Length import MIN_LENGTH_MM
from BranchTools_Rotate import wrap_degrees
from BranchTools_SideBranch import ELBOW_ANGLES, FROM_ROTATION

SLOPED_STEPS = (1, 5, 10, 50, 100)
FROM_ROTATION_LABEL = "From rotation"
PERPENDICULAR_LABEL = "Perpendicular to main"
ALONG_BRANCH_LABEL = "Along the branch"
DEFAULT_SETTINGS = {"offset": 100.0, "elbow": FROM_ROTATION, "length": 500.0, "direction": 1.0,
                    "keep_level": 1.0, "sloped": 10.0}


class SideBranchWindow(object):
    """Finestra di Side Branch.
    rotator: BranchTools_Rotate.Rotator della braga (serve con una pendenza scelta); side_branch:
    BranchTools_SideBranch.SideBranch; branch_length: BranchTools_Length.BranchLength;
    settings: {"offset", "length": mm; "elbow": pendenza in gradi (0 = dalla rotazione); "direction": 1 = tratto
    orizzontale perpendicolare al collettore, 0 = come la derivazione; "keep_level": 1 = le copie tengono la quota
    dello stacco originale; "sloped": passo in mm}.
    show() ritorna lo stesso dict aggiornato con "action" (None o "place"), o None se annullata."""

    def __init__(self, xaml_path, rotator, side_branch, branch_length, settings):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        find = self.window.FindName
        self.result = None
        self.rotator = rotator
        self.side_branch = side_branch
        self.branch_length = branch_length

        # stacco laterale
        self.offset_box = find("OffsetBox")
        self.length_box = find("LengthBox")
        self.slope_combo = find("SlopeCombo")
        self.direction_combo = find("DirectionCombo")
        self.create_button = find("CreateBranchButton")
        self.branch_status = find("BranchStatus")
        self.offset_box.Text = "{:g}".format(settings["offset"])
        self.length_box.Text = "{:g}".format(settings["length"])
        # pendenza del tubo inclinato: la chiave "elbow" e' storica (fino alla 0.5 era anche l'angolo della curva)
        self.slope_combo.Items.Add(FROM_ROTATION_LABEL)
        for value in ELBOW_ANGLES:
            self.slope_combo.Items.Add(step_text(value))
        self.slope_combo.SelectedItem = (step_text(settings["elbow"]) if settings["elbow"] in ELBOW_ANGLES
                                         else FROM_ROTATION_LABEL)
        for label in (PERPENDICULAR_LABEL, ALONG_BRANCH_LABEL):
            self.direction_combo.Items.Add(label)
        self.direction_combo.SelectedItem = PERPENDICULAR_LABEL if settings["direction"] else ALONG_BRANCH_LABEL
        self.slope_combo.SelectionChanged += lambda s, a: self._show_elbow()
        self.direction_combo.SelectionChanged += lambda s, a: self._show_elbow()
        self.create_button.Click += self._on_create_branch
        self.branch_controls = (self.offset_box, self.length_box, self.slope_combo, self.direction_combo,
                                self.create_button)
        for control in self.branch_controls:
            control.IsEnabled = side_branch.ok
        self.place_button = find("PlaceCopiesButton")
        self.place_button.Click += self._on_place
        self.place_button.IsEnabled = side_branch.reason is None and side_branch.place_reason() is None
        self.keep_level_box = find("KeepLevelBox")
        self.keep_level_box.IsChecked = settings["keep_level"] > 0
        self.keep_level_box.IsEnabled = side_branch.reason is None
        self._show_elbow()

        # lunghezza del tubo inclinato di uno stacco esistente
        self.height_box = find("HeightBox")
        sloped_slider = find("SlopedSlider")
        sloped_slider.Minimum = MIN_LENGTH_MM
        if branch_length.ok:
            sloped_slider.Maximum = branch_length.max_value()
        self.sloped = SliderSection(self.window, "Sloped", SLOPED_STEPS, settings["sloped"],
                                    branch_length.start if branch_length.ok else MIN_LENGTH_MM,
                                    self._then_show_height(branch_length.set_to), self._describe_length,
                                    branch_length.ok)
        on_enter_or_leave(self.height_box, self._apply_height)
        self._show_height()

        find("OkButton").Click += self._on_ok
        find("CancelButton").Click += self._on_cancel

    # -------------------------------------------------------------- lunghezza dello stacco

    def _describe_length(self, value):
        if self.branch_length.ok:
            return ""
        return "Not available: {}".format(self.branch_length.reason)

    def _then_show_height(self, apply):
        def wrapped(value):
            result = apply(value)
            self._show_height()
            return result
        return wrapped

    def _show_height(self):
        self.height_box.IsEnabled = self.branch_length.ok
        if self.branch_length.ok:
            self.height_box.Text = "{:g}".format(round(self.branch_length.offset(), 1))
        else:
            self.height_box.Text = ""

    def _apply_height(self):
        if not self.branch_length.ok:
            return
        ok, offset = parse_number(self.height_box.Text)
        if not ok:
            self.sloped.show_error("the invert offset is not a number.")
            self._show_height()
            return
        if abs(offset - self.branch_length.offset()) < 0.05:
            return
        try:
            length = self.branch_length.length_for_offset(offset)
        except Exception as err:
            self.sloped.show_error(str(err))
            self._show_height()
            return
        self.sloped.go_to(length)

    # -------------------------------------------------------------- stacco laterale

    def _slope_choice(self):
        """Pendenza scelta del tubo inclinato in gradi, o None per "From rotation"."""
        ok, slope = parse_number(self.slope_combo.SelectedItem or "")
        return slope if ok else None

    def _perpendicular(self):
        return self.direction_combo.SelectedItem != ALONG_BRANCH_LABEL

    def _show_elbow(self):
        # riga di stato solo con un'informazione utile: il motivo per cui la sezione e' spenta, o la curva calcolata
        if self.side_branch.reason is not None:
            text = "Not available: {}".format(self.side_branch.reason)
        elif not self.side_branch.ok or self._slope_choice() is not None:
            text = ""
        else:
            try:
                text = "Elbow {:g}°.".format(round(self.side_branch.elbow_angle(self._perpendicular()), 2))
            except Exception as err:
                text = str(err)
        set_status(self.branch_status, text)

    def _branch_values(self):
        """(ok, offset mm, lunghezza mm) letti dalla sezione Side branch."""
        ok_offset, offset = parse_number(self.offset_box.Text)
        ok_length, length = parse_number(self.length_box.Text)
        if not ok_offset or offset == 0:
            set_status(self.branch_status, "Invert offset: a number of mm other than 0 (+ up, - down).")
            return False, None, None
        if not ok_length or length <= 0:
            set_status(self.branch_status, "Horizontal length: a number of mm greater than 0.")
            return False, None, None
        return True, offset, length

    def _on_create_branch(self, sender, args):
        ok, offset, length = self._branch_values()
        if not ok:
            return
        slope = self._slope_choice()
        perpendicular = self._perpendicular()
        if slope is not None:
            try:
                target = self.side_branch.target_angle(offset, slope)
            except Exception as err:
                set_status(self.branch_status, "Not created: {}".format(err))
                return
            applied, error = self.rotator.rotate_to(target)
            if error or abs(wrap_degrees(applied - target)) > 0.01:
                set_status(self.branch_status, "Not created: the wye could not be rotated to {:g}°. {}".format(
                    round(target, 2), error or ""))
                return
        try:
            elbow_used = self.side_branch.elbow_angle(perpendicular)
        except Exception as err:
            set_status(self.branch_status, "Not created: {}".format(err))
            return
        error = self.side_branch.create(offset, length, perpendicular)
        if error:
            set_status(self.branch_status, "Not created: {}".format(error))
            return
        for control in self.branch_controls:
            control.IsEnabled = False
        self.place_button.IsEnabled = True
        set_status(self.branch_status, "Created, elbow {:g}°.".format(round(elbow_used, 2)))
        # lo stacco appena creato si puo' allungare o accorciare
        if self.branch_length.load() is None:
            self.sloped.slider.Maximum = self.branch_length.max_value()
            self.sloped.reset(self.branch_length.start, True)
            self._show_height()

    # -------------------------------------------------------------- OK / Cancel

    def _settings(self):
        settings = {"sloped": self.sloped.step()}
        ok_offset, offset = parse_number(self.offset_box.Text)
        ok_length, length = parse_number(self.length_box.Text)
        slope = self._slope_choice()
        if ok_offset:
            settings["offset"] = offset
        if ok_length and length > 0:
            settings["length"] = length
        settings["elbow"] = FROM_ROTATION if slope is None else slope
        settings["direction"] = 1.0 if self._perpendicular() else 0.0
        settings["keep_level"] = 1.0 if self.keep_level_box.IsChecked else 0.0
        return settings

    def _on_place(self, sender, args):
        # cliccare nel modello non si puo' con la finestra modale aperta: si chiude come OK
        self.result = self._settings()
        self.result["action"] = "place"
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
