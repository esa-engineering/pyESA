# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import File
from System.Windows.Markup import XamlReader

from BranchTools_Controls import SliderSection, set_status

ROTATE_STEPS = (1, 5, 15, 45, 90)
ANGLE_STEPS = (1, 2.5, 5, 15, 45)
DEFAULT_SETTINGS = {"rotate": 5.0, "angle": 2.5}


class FittingRotateWindow(object):
    """Finestra di Fitting Rotate.
    rotator / angle_setter: BranchTools_Rotate (vuoti se l'elemento e' un tubo); aligner: BranchTools_Align.Aligner;
    settings: {"rotate", "angle": passi in gradi}.
    show() ritorna lo stesso dict aggiornato con "action" (None o "align"), o None se annullata."""

    def __init__(self, xaml_path, rotator, angle_setter, aligner, settings):
        self.window = XamlReader.Parse(File.ReadAllText(xaml_path))
        find = self.window.FindName
        self.result = None
        skipped = rotator.skipped
        is_fitting = bool(rotator.frames)

        def describe_rotation(value):
            if not is_fitting:
                return "Not available: select a fitting."
            if not skipped:
                return ""
            return "Not rotatable: {}.".format(", ".join(sorted(set(skipped))))

        count = len(angle_setter.fittings)

        def describe_angle(value):
            if not is_fitting:
                return "Not available: select a fitting."
            return "" if count else "No editable Angle parameter."

        self.rotate = SliderSection(self.window, "Rotate", ROTATE_STEPS, settings["rotate"], rotator.start,
                                    rotator.rotate_to, describe_rotation, len(rotator.rotatable) > 0)
        self.angle = SliderSection(self.window, "Angle", ANGLE_STEPS, settings["angle"], angle_setter.start,
                                   angle_setter.set_to, describe_angle, count > 0)

        # allineamento: l'elemento a cui allinearsi si sceglie dopo la chiusura della finestra
        align_button = find("AlignButton")
        align_button.Click += self._on_align
        align_button.IsEnabled = aligner.ok
        set_status(find("AlignStatus"), aligner.describe())

        find("OkButton").Click += self._on_ok
        find("CancelButton").Click += self._on_cancel

    def _settings(self):
        return {"rotate": self.rotate.step(), "angle": self.angle.step()}

    def _on_align(self, sender, args):
        # cliccare nel modello non si puo' con la finestra modale aperta: si chiude come OK
        self.result = self._settings()
        self.result["action"] = "align"
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
