# -*- coding: utf-8 -*-
"""Finestra del tool di nomenclatura.

Carica il XAML con XamlReader e costruisce a runtime i campi che cambiano da
una categoria all'altra: il blocco dimensionale, i criteri di codifica
specifici, le note. Tutto quello che l'utente vede e' in inglese; i commenti
restano in italiano come nel resto dell'estensione.

La finestra non tocca Revit: riceve dallo script un contesto gia' pronto
(misure lette dai parametri, Type Mark gia' in uso, nomi gia' assegnati) e
restituisce il risultato in self.result.
"""

import os

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.IO import StreamReader
from System.Windows import Window, Thickness, GridLength, GridUnitType
from System.Windows import (
    HorizontalAlignment, VerticalAlignment, TextWrapping, FontWeights, Visibility
)
from System.Windows.Controls import (
    Grid, ColumnDefinition, TextBlock, TextBox, ComboBox, CheckBox
)
from System.Windows.Markup import XamlReader
from System.Windows.Media import Brushes, SolidColorBrush, Color

import FamilyNaming_data as DATA
import FamilyNaming_map as MAP
import FamilyNaming_rules as RULES

XAML_FILE = "FamilyNaming.xaml"

GRAY = Brushes.Gray
LABEL_COL = 150


def _brush(r, g, b):
    return SolidColorBrush(Color.FromRgb(r, g, b))


ACCENT = _brush(0x2D, 0x5A, 0x8A)
WARNING = _brush(0xB0, 0x60, 0x00)
ERROR = _brush(0xC0, 0x00, 0x00)


class NamingWindow(Window):
    """Finestra unica del tool. self.result e' None se l'utente annulla."""

    def __init__(self, ctx):
        self.ctx = ctx
        self.sheet = ctx["sheet"]
        self.result = None

        # controlli costruiti a runtime, per chiave logica
        self._ctl = {}
        # testo di esempio e suo stato, per chiave: vedi la nota sul watermark
        self._wm_text = {}
        self._wm_on = {}
        # sospende il ricalcolo mentre si ripopolano i menu a cascata
        self._loading = True
        # ricostruire le misure a ogni tasto cancellerebbe quello che si sta
        # scrivendo: si rifanno solo quando la sezione passa da tonda a non
        # tonda, che e' l'unica cosa che ne cambia l'elenco
        self._dim_round = None

        self._load_xaml()
        self._build()
        self._loading = False
        self._refresh()

    # -- caricamento --------------------------------------------------------

    def _load_xaml(self):
        xaml_path = os.path.join(os.path.dirname(__file__), XAML_FILE)

        Window.__init__(self)
        reader = StreamReader(xaml_path)
        try:
            root = XamlReader.Load(reader.BaseStream)
        finally:
            reader.Close()

        self.Content = root.Content
        self.Title = root.Title
        self.Height = root.Height
        self.Width = root.Width
        self.MinHeight = root.MinHeight
        self.MinWidth = root.MinWidth
        self.WindowStartupLocation = root.WindowStartupLocation
        self.ResizeMode = root.ResizeMode
        self.ShowInTaskbar = root.ShowInTaskbar

        names = (
            "lbl_element", "lbl_sheet",
            "cmb_category", "cmb_group1", "cmb_group2",
            "lbl_group1", "lbl_group2", "lbl_group1_desc", "lbl_group2_desc",
            "chk_inplace",
            "txt_typemark", "lbl_tm_hint", "chk_tm_auto", "chk_tm_xx",
            "chk_tm_write", "lbl_tm_existing",
            "brd_specific", "pnl_specific",
            "brd_dimensions", "pnl_dimensions", "lbl_dim_hint",
            "lbl_desc_family", "txt_desc_family", "txt_desc_type",
            "lbl_notes_header", "pnl_notes",
            "grd_preview_family", "txt_preview_family", "txt_preview_type",
            "lbl_validation",
            "btn_cancel", "btn_apply",
        )
        for name in names:
            setattr(self, name, root.FindName(name))

    # -- costruzione --------------------------------------------------------

    def _build(self):
        ctx = self.ctx
        sheet = self.sheet

        self.lbl_element.Text = ctx["element_label"]
        self.lbl_sheet.Text = u"Naming sheet: {0}  ({1})".format(
            sheet["label"], ctx["sheet_id"].replace(u":", u" / "))

        self.chk_inplace.IsChecked = bool(ctx.get("in_place"))

        # una scheda di sistema non compone un nome famiglia: l'anteprima
        # mostra il solo nome del tipo
        if sheet.get("schema") != "loadable":
            self.grd_preview_family.Visibility = Visibility.Collapsed
            self.lbl_desc_family.Text = u"Description"
            self.txt_desc_type.IsEnabled = False
            self.txt_desc_type.Text = u""

        self._fill_categories()
        self._fill_groups()
        self._build_specific()
        self._build_dimensions()
        self._build_notes()

        self.cmb_category.SelectionChanged += self._on_category_changed
        self.cmb_group1.SelectionChanged += self._on_group_changed
        self.cmb_group2.SelectionChanged += self._on_group_changed
        self.txt_typemark.TextChanged += self._on_changed
        self.chk_tm_auto.Click += self._on_changed
        self.chk_tm_xx.Click += self._on_changed
        self.chk_tm_write.Click += self._on_changed
        self.txt_desc_family.TextChanged += self._on_changed
        self.txt_desc_type.TextChanged += self._on_changed
        self.btn_cancel.Click += self._on_cancel
        self.btn_apply.Click += self._on_apply

    # -- menu a tendina -----------------------------------------------------

    def _fill_categories(self):
        rows = RULES.table(self.sheet["cat_table"])
        self.cmb_category.Items.Clear()
        selected = 0
        for i, row in enumerate(rows):
            self.cmb_category.Items.Add(self._item(row))
            if row[0] == self.ctx.get("cat_code"):
                selected = i
        # una scheda con una sola famiglia di sistema non ha nulla da scegliere
        self.cmb_category.IsEnabled = len(rows) > 1
        if rows:
            self.cmb_category.SelectedIndex = selected

    def _fill_groups(self):
        """Ripopola Group1 e Group2 in base alla famiglia di sistema scelta.

        Su Walls, Roofs e Stairs le due liste cambiano con il sottogruppo,
        quindi vanno ricostruite a ogni cambio di categoria.
        """
        cat = self._code(self.cmb_category)

        prev_g1 = self._code(self.cmb_group1)
        prev_g2 = self._code(self.cmb_group2)

        g1_rows = RULES.table(RULES.resolve_table(self.sheet.get("g1"), cat))
        g2_rows = RULES.table(RULES.resolve_table(self.sheet.get("g2"), cat))

        self._reload_combo(self.cmb_group1, g1_rows, prev_g1)
        self._reload_combo(self.cmb_group2, g2_rows, prev_g2)

        # sulle sette categorie caricabili a Group1 condiviso il primo gruppo
        # non descrive com'e' fatta la famiglia ma da dove viene
        shared = self.sheet.get("g1") == MAP.TBL_SHARED_G1
        self.lbl_group1.Text = u"Group 1 (origin)" if shared else u"Group 1"

    def _reload_combo(self, combo, rows, keep_code):
        combo.Items.Clear()
        selected = 0
        for i, row in enumerate(rows):
            combo.Items.Add(self._item(row))
            if row[0] == keep_code:
                selected = i
        combo.IsEnabled = len(rows) > 0
        if rows:
            combo.SelectedIndex = selected

    def _item(self, row):
        code = row[0]
        label = row[1] if len(row) > 1 else code
        return u"{0}  -  {1}".format(code, label)

    def _code(self, combo):
        """Codice della voce selezionata, vuoto se non c'e' selezione."""
        item = combo.SelectedItem
        if not item:
            return u""
        return item.split(u"  -  ")[0].strip()

    # -- criteri di codifica specifici della categoria ----------------------

    def _build_specific(self):
        """Costruisce i campi previsti dai blocchi della scheda.

        L'ordine e' quello dei blocchi nel nome, cosi' il form si legge dalla
        stessa direzione in cui si legge il risultato.
        """
        panel = self.pnl_specific
        panel.Children.Clear()
        self._ctl = {}
        self._wm_text = {}
        self._wm_on = {}

        blocks = self.sheet.get("blocks", ())
        built = 0

        # il Type Mark manuale serve dove la cifra viene dal Group1: con
        # Group1 = Other il prefisso lo sceglie l'utente fra i materiali
        if self.sheet.get("tm_from") == "g1_digit":
            self._add_combo(
                panel, "tm_prefix", u"Manual TM prefix",
                RULES.table(MAP.TBL_MATERIAL),
                hint=u"Only with Group 1 = Other. Replaces the whole Type Mark prefix.")
            built += 1

        for block in blocks:
            if block == "description":
                continue                       # ha una sezione sua
            built += self._add_block(panel, block)

        self.brd_specific.Visibility = (
            Visibility.Visible if built else Visibility.Collapsed)

    def _add_block(self, panel, block):
        if block == "leaves":
            values = [(u"{0}".format(n), u"{0} leaf".format(n) if n == 1
                       else u"{0} leaves".format(n)) for n in range(0, 10)]
            self._add_combo(
                panel, "leaves", u"Leaves", values, required=True, preset=u"1",
                hint=u"Always filled. A free opening (Group 2 = OP) is L0. "
                     u"Two leaves are a different family, not a different type.")
            return 1
        if block == "cw_hosted":
            self._add_check(
                panel, "cw_hosted", u"Curtain-hosted",
                hint=u"Nested in a curtain wall panel instead of hosted by a wall.")
            return 1
        if block == "entrance":
            self._add_check(
                panel, "entrance", u"Entrance",
                hint=u"Marks the main entrances. Allowed only with Group 2 = EX.")
            return 1
        if block == "rei":
            self._add_text(
                panel, "rei", u"REI", watermark=u"REI60, EI90, RE60",
                hint=u"Optional. Full designation as certified. "
                     u"It is written into the FireRating parameter.")
            return 1
        if block == "wi":
            self._add_check(panel, "waterproof", u"W  waterproofing included")
            self._add_check(panel, "insulation", u"I  insulation included")
            return 1
        if block == "rail_top":
            self._add_combo(
                panel, "rail_top", u"Top rail / handrail",
                RULES.table(MAP.TBL_RAIL_TOP), required=True)
            return 1
        if block == "underside":
            self._add_combo(
                panel, "underside", u"Underside surface",
                RULES.table(MAP.TBL_UNDERSIDE),
                hint=u"Monolithic runs only.")
            return 1
        if block == "ln_shape":
            self._add_check(
                panel, "ln_shape", u"Same as Run (SR)",
                hint=u"The landing inherits the thickness from the run. "
                     u"It replaces the whole dimensional block.")
            return 1
        if block == "rm_shape":
            self._add_check(
                panel, "rm_shape", u"Solid (SL)",
                hint=u"The ramp is solid down to the ground. "
                     u"It replaces the whole dimensional block.")
            return 1
        if block == "use":
            self._add_combo(
                panel, "use", u"Use", RULES.table(MAP.TBL_USE),
                hint=u"Non structural use of the category. Requires Group 1 = "
                     u"Other and takes over the Type Mark prefix.")
            return 1
        if block == "role":
            self._add_combo(
                panel, "role", u"Family role", RULES.table(MAP.TBL_ROLE),
                hint=u"Leave empty for a complete, standalone family.")
            return 1
        if block == "custom":
            self._add_check(
                panel, "custom", u"Custom (CT)",
                hint=u"Drawn to measure. Not compatible with Manufacturer.")
            return 1
        if block == "manufacturer":
            self._add_text(
                panel, "manufacturer", u"Manufacturer", watermark=u"e.g. Schuco",
                hint=u"Optional. The company that makes the object.")
            return 1
        if block == "brand":
            self._add_text(
                panel, "brand", u"Brand", watermark=u"e.g. Dior",
                hint=u"Optional. The CLIENT brand this element is a standard of, "
                     u"not the manufacturer product line.")
            return 1
        return 0

    # -- blocco dimensionale ------------------------------------------------

    def _build_dimensions(self):
        panel = self.pnl_dimensions
        panel.Children.Clear()

        kind = self.sheet.get("dim")
        if kind == "none":
            self.brd_dimensions.Visibility = Visibility.Collapsed
            return
        self.brd_dimensions.Visibility = Visibility.Visible

        hints = {
            "wxh": u"Width x height of the opening, in mm.",
            "wxdxh": u"Plan footprint in mm. Height only when it tells two types apart.",
            "wxl": u"Bay size in mm. 250x500 standard, 320x500 accessible.",
            "wxd_or_d": u"Section in mm. A round section is quoted by diameter only.",
            "bxh_or_d": u"Profile section in mm. A round profile is quoted by diameter only.",
            "wxt": u"Width x thickness in mm.",
            "t": u"Thickness in mm.",
            "h": u"Height in mm.",
            "nfin_t": u"Finished faces and total thickness: 2F.125. A single layer "
                      u"element is always 0F, even when that layer is the finishing.",
            "free": self.sheet.get("dim_hint", u"")
                    or u"Use the standard designation when one exists.",
        }
        self.lbl_dim_hint.Text = hints.get(kind, u"")

        if kind == "nfin_t":
            self._add_combo(
                panel, "nfin", u"Finished faces",
                RULES.table(self.sheet.get("nfin_table") or MAP.TBL_NFIN),
                required=True, preset=self.ctx.get("nfin_default"))

        self._rebuild_dim_fields(panel, kind, force=True)

    def _rebuild_dim_fields(self, panel, kind, force=False):
        """Rifa' le righe delle misure. Serve solo quando cambia il loro elenco.

        L'elenco dipende da una cosa sola: se il Group1 dichiara una sezione
        tonda si chiede il diametro, altrimenti le due misure della sezione.
        Fuori da quel passaggio le righe si lasciano stare, altrimenti si
        cancellerebbe quello che l'utente sta scrivendo.
        """
        round_section = self._code(self.cmb_group1) in (self.sheet.get("round_g1") or ())
        if not force and round_section == self._dim_round:
            return
        self._dim_round = round_section

        for key in list(self._ctl.keys()):
            if key.startswith("dim:"):
                del self._ctl[key]
                self._wm_text.pop(key, None)
                self._wm_on.pop(key, None)
        # tolgo solo le righe dimensionali, la combo nFinishings resta
        keep = []
        for child in panel.Children:
            if getattr(child, "Tag", None) != "dimrow":
                keep.append(child)
        panel.Children.Clear()
        for child in keep:
            panel.Children.Add(child)

        defaults = self.ctx.get("dim_defaults", {})

        for key, label, required in RULES.dim_fields(kind, round_section):
            if kind == "free":
                row = self._add_text(
                    panel, "dim:" + key, label,
                    watermark=self.sheet.get("dim_hint", u""), required=required,
                    tag="dimrow")
            else:
                row = self._add_text(
                    panel, "dim:" + key, label, unit=u"mm", narrow=True,
                    required=required, tag="dimrow")
            preset = defaults.get(key)
            if preset:
                box = self._ctl["dim:" + key]
                self._clear_watermark("dim:" + key, box)
                box.Text = preset

    # -- note ---------------------------------------------------------------

    def _build_notes(self):
        panel = self.pnl_notes
        panel.Children.Clear()

        notes = DATA.NOTES.get(self.ctx["sheet_id"], ())
        if not notes:
            self.lbl_notes_header.Visibility = Visibility.Collapsed
            return

        for kind, head, body in notes:
            if kind == "section":
                block = TextBlock()
                block.Text = head
                block.FontWeight = FontWeights.Bold
                block.Foreground = ACCENT
                block.Margin = Thickness(0, 10, 0, 4)
                block.TextWrapping = TextWrapping.Wrap
                panel.Children.Add(block)
                continue
            if head:
                block = TextBlock()
                block.Text = head
                block.FontWeight = FontWeights.Bold
                block.FontSize = 11
                block.Margin = Thickness(0, 6, 0, 1)
                block.TextWrapping = TextWrapping.Wrap
                panel.Children.Add(block)
            if body:
                block = TextBlock()
                block.Text = body
                block.FontSize = 11
                block.Foreground = GRAY
                block.TextWrapping = TextWrapping.Wrap
                panel.Children.Add(block)

    # -- costruttori di righe -----------------------------------------------

    def _row(self, panel, label_text, control, unit=None, required=False, tag=None):
        grid = Grid()
        grid.Margin = Thickness(0, 5, 0, 0)
        if tag:
            grid.Tag = tag

        col_label = ColumnDefinition()
        col_label.Width = GridLength(LABEL_COL, GridUnitType.Pixel)
        col_ctl = ColumnDefinition()
        col_ctl.Width = GridLength(1, GridUnitType.Star)
        col_unit = ColumnDefinition()
        col_unit.Width = GridLength(1, GridUnitType.Auto)
        grid.ColumnDefinitions.Add(col_label)
        grid.ColumnDefinitions.Add(col_ctl)
        grid.ColumnDefinitions.Add(col_unit)

        label = TextBlock()
        label.Text = label_text if required else u"{0}  (optional)".format(label_text)
        label.VerticalAlignment = VerticalAlignment.Center
        label.TextWrapping = TextWrapping.Wrap
        label.Margin = Thickness(0, 0, 10, 0)
        label.SetValue(Grid.ColumnProperty, 0)
        grid.Children.Add(label)

        control.SetValue(Grid.ColumnProperty, 1)
        grid.Children.Add(control)

        if unit:
            unit_block = TextBlock()
            unit_block.Text = unit
            unit_block.FontSize = 11
            unit_block.Foreground = GRAY
            unit_block.VerticalAlignment = VerticalAlignment.Center
            unit_block.Margin = Thickness(6, 0, 0, 0)
            unit_block.SetValue(Grid.ColumnProperty, 2)
            grid.Children.Add(unit_block)

        panel.Children.Add(grid)
        return grid

    def _hint(self, panel, text, tag=None):
        if not text:
            return
        block = TextBlock()
        block.Text = text
        block.FontSize = 11
        block.Foreground = GRAY
        block.TextWrapping = TextWrapping.Wrap
        block.Margin = Thickness(LABEL_COL, 1, 0, 0)
        if tag:
            block.Tag = tag
        panel.Children.Add(block)

    def _add_combo(self, panel, key, label, rows, required=False, hint=None,
                   preset=None, tag=None):
        combo = ComboBox()
        combo.Padding = Thickness(4, 2, 4, 2)
        combo.HorizontalAlignment = HorizontalAlignment.Stretch
        if not required:
            combo.Items.Add(u"")                 # la cella vuota e' il caso normale
        for row in rows:
            combo.Items.Add(self._item(row))
        combo.SelectedIndex = 0
        if preset:
            for i in range(combo.Items.Count):
                if str(combo.Items[i]).startswith(preset):
                    combo.SelectedIndex = i
                    break
        combo.SelectionChanged += self._on_changed
        self._ctl[key] = combo
        self._row(panel, label, combo, required=required, tag=tag)
        self._hint(panel, hint, tag=tag)
        return combo

    def _add_text(self, panel, key, label, watermark=None, unit=None,
                  narrow=False, required=False, hint=None, tag=None):
        box = TextBox()
        box.Padding = Thickness(3, 2, 3, 2)
        if narrow:
            box.Width = 70
            box.HorizontalAlignment = HorizontalAlignment.Left
        else:
            box.HorizontalAlignment = HorizontalAlignment.Stretch
        if watermark:
            self._set_watermark(key, box, watermark)
        box.TextChanged += self._on_changed
        self._ctl[key] = box
        self._row(panel, label, box, unit=unit, required=required, tag=tag)
        self._hint(panel, hint, tag=tag)
        return box

    def _add_check(self, panel, key, label, hint=None, tag=None):
        check = CheckBox()
        check.Content = label
        check.Margin = Thickness(LABEL_COL, 8, 0, 0)
        if tag:
            check.Tag = tag
        check.Click += self._on_changed
        self._ctl[key] = check
        panel.Children.Add(check)
        self._hint(panel, hint, tag=tag)
        return check

    # -- watermark ----------------------------------------------------------
    #
    # Il testo di esempio e' grigio, sparisce appena si scrive e non viene mai
    # restituito come valore: un campo lasciato intatto resta vuoto.
    #
    # Lo stato sta in due dizionari della finestra e non sul controllo, perche'
    # IronPython non lascia attaccare attributi Python a un oggetto .NET: un
    # box._watermark = ... su un TextBox solleva AttributeError.

    def _set_watermark(self, key, box, text):
        self._wm_text[key] = text
        self._wm_on[key] = True
        box.Text = text
        box.Foreground = GRAY

        def on_got_focus(sender, args, k=key):
            if self._wm_on.get(k):
                self._wm_on[k] = False
                sender.Text = u""
                sender.Foreground = Brushes.Black

        def on_lost_focus(sender, args, k=key):
            if not sender.Text:
                self._wm_on[k] = True
                sender.Text = self._wm_text.get(k, u"")
                sender.Foreground = GRAY

        box.GotFocus += on_got_focus
        box.LostFocus += on_lost_focus

    def _clear_watermark(self, key, box):
        self._wm_on[key] = False
        box.Foreground = Brushes.Black

    def _text_of(self, key):
        box = self._ctl.get(key)
        if box is None:
            return u""
        if self._wm_on.get(key):
            return u""
        return (box.Text or u"").strip()

    def _combo_of(self, key):
        combo = self._ctl.get(key)
        if combo is None or not combo.IsEnabled:
            return u""
        item = combo.SelectedItem
        if not item:
            return u""
        text = str(item)
        if not text:
            return u""
        return text.split(u"  -  ")[0].strip()

    def _check_of(self, key):
        check = self._ctl.get(key)
        if check is None or not check.IsEnabled:
            return False
        return bool(check.IsChecked)

    # -- eventi -------------------------------------------------------------

    def _on_category_changed(self, sender, args):
        if self._loading:
            return
        self._loading = True
        self._fill_groups()
        self._loading = False
        self._refresh()

    def _on_group_changed(self, sender, args):
        if self._loading:
            return
        self._refresh()

    def _on_changed(self, sender, args):
        if self._loading:
            return
        self._refresh()

    # -- stato e anteprima --------------------------------------------------

    def _apply_conditional_rules(self, cat, g1, g2):
        """Accende e spegne i campi secondo le regole della scheda.

        Un menu che si apre vuoto o un campo grigio vogliono dire che quel
        campo non si applica al caso scelto, esattamente come nei fogli Excel.
        """
        nested = g1 in MAP.NESTED_G1
        manual_tm = g1 == MAP.MANUAL_TM_G1

        # un componente nidificato o un simbolo non hanno ante
        if "leaves" in self._ctl:
            self._ctl["leaves"].IsEnabled = not nested

        # Entrance esiste solo sugli ingressi esterni
        if "entrance" in self._ctl:
            allowed = g2 == u"EX"
            self._ctl["entrance"].IsEnabled = allowed
            if not allowed:
                self._ctl["entrance"].IsChecked = False

        # i suffissi W e I valgono solo sui Group2 che li ammettono: altrove
        # impermeabilizzazione e isolamento sono impliciti
        allowed_wi = self.sheet.get("wi_allowed_g2")
        if allowed_wi is not None:
            ok = g2 in allowed_wi
            for key in ("waterproof", "insulation"):
                if key in self._ctl:
                    self._ctl[key].IsEnabled = ok
                    if not ok:
                        self._ctl[key].IsChecked = False

        # la superficie d'intradosso si dichiara solo sulle rampe monolitiche
        if "underside" in self._ctl:
            cats = self.sheet.get("underside_cat")
            ok = cats is None or cat in cats
            self._ctl["underside"].IsEnabled = ok
            if not ok:
                self._ctl["underside"].SelectedIndex = 0

        # il campo Use si accende solo con Group1 = Other
        if "use" in self._ctl:
            self._ctl["use"].IsEnabled = manual_tm
            if not manual_tm:
                self._ctl["use"].SelectedIndex = 0

        # il prefisso manuale serve solo quando il Type Mark lo diventa
        if "tm_prefix" in self._ctl:
            use_code = self._combo_of("use")
            self._ctl["tm_prefix"].IsEnabled = manual_tm and not use_code

        # Manufacturer e Custom non convivono
        if "manufacturer" in self._ctl:
            custom = self._check_of("custom") or g1 == u"CT"
            self._ctl["manufacturer"].IsEnabled = not custom

        # due casi spengono l'intero blocco dimensionale: SR e SL, che ne
        # prendono il posto nel nome, e il pannello di sistema vuoto, che uno
        # spessore non ce l'ha perche' ospita una famiglia caricabile
        substitute = self._check_of("ln_shape") or self._check_of("rm_shape")
        no_dim = g1 in (self.sheet.get("dim_optional_g1") or ())
        enabled = not (substitute or no_dim)

        for key in self._ctl:
            if not key.startswith("dim:"):
                continue
            box = self._ctl[key]
            box.IsEnabled = enabled
            if not enabled:
                box.Text = u""

    def _collect(self):
        cat = self._code(self.cmb_category)
        g1 = self._code(self.cmb_group1)
        g2 = self._code(self.cmb_group2)

        values = {
            "cat": cat,
            "g1": g1,
            "g2": g2,
            "in_place": bool(self.chk_inplace.IsChecked),
            "leaves": self._combo_of("leaves"),
            "cw_hosted": self._check_of("cw_hosted"),
            "entrance": self._check_of("entrance"),
            "rei": self._text_of("rei"),
            "waterproof": self._check_of("waterproof"),
            "insulation": self._check_of("insulation"),
            "rail_top": self._combo_of("rail_top"),
            "underside": self._combo_of("underside"),
            "ln_shape": self._check_of("ln_shape"),
            "rm_shape": self._check_of("rm_shape"),
            "use": self._combo_of("use"),
            "role": self._combo_of("role"),
            "custom": self._check_of("custom"),
            "manufacturer": self._text_of("manufacturer"),
            "brand": self._text_of("brand"),
            "description": (self.txt_desc_family.Text or u"").strip(),
            "type_description": (self.txt_desc_type.Text or u"").strip(),
        }

        dim = {}
        for key in self._ctl:
            if key.startswith("dim:"):
                dim[key[4:]] = self._text_of(key)
        if "nfin" in self._ctl:
            dim["nfin"] = self._combo_of("nfin")
        values["dim"] = dim

        return values

    def _compute_type_mark(self, values):
        """Calcola il Type Mark suggerito e lo scrive nel campo se in automatico."""
        sheet = self.sheet
        use_code = values.get("use")

        # Con Group1 = Other il Type Mark e' interamente manuale, prefisso
        # compreso: il 9xx che la tabella riporta accanto a Other non si usa.
        # Il prefisso viene dal campo Use se compilato, altrimenti dal
        # materiale, e finche' non se ne sceglie uno non c'e' nulla da
        # suggerire.
        manual = None
        needs_manual = (values.get("g1") == MAP.MANUAL_TM_G1
                        and sheet.get("tm_from") == "g1_digit")
        if needs_manual:
            manual = use_code or self._combo_of("tm_prefix")

        prefix, digits = RULES.type_mark_prefix(
            sheet, values["cat"], values["g1"], values["g2"], manual)

        incomplete = (needs_manual and not manual) or \
            (prefix.endswith(u"-") and digits == 2)

        self.lbl_tm_hint.Text = u"" if incomplete else \
            u"prefix {0} + {1} digits".format(prefix, digits)

        use_xx = bool(self.chk_tm_xx.IsChecked)
        auto = bool(self.chk_tm_auto.IsChecked)
        self.txt_typemark.IsReadOnly = auto

        if not auto:
            return (self.txt_typemark.Text or u"").strip(), None

        if incomplete:
            self.txt_typemark.Text = u""
            return u"", u"Pick the manual Type Mark prefix to get a suggestion."

        sequential = None
        warning = None
        if not use_xx:
            sequential = RULES.next_sequential(
                self.ctx.get("existing_marks", ()), prefix, digits)
            if RULES.sequential_overflow(sequential, digits):
                warning = (u"All {0}-digit sequentials for prefix {1} are taken. "
                           u"Uncheck the automatic suggestion and set the Type "
                           u"Mark by hand.").format(digits, prefix)
                sequential = None

        mark = RULES.format_type_mark(prefix, digits, sequential, use_xx)
        self.txt_typemark.Text = mark
        return mark, warning

    def _refresh(self):
        if self._loading:
            return
        self._loading = True
        try:
            cat = self._code(self.cmb_category)
            g1 = self._code(self.cmb_group1)
            g2 = self._code(self.cmb_group2)

            self._apply_conditional_rules(cat, g1, g2)
            self._rebuild_dim_fields(self.pnl_dimensions, self.sheet.get("dim"))
            self._apply_conditional_rules(cat, g1, g2)

            g1_rows = RULES.table(RULES.resolve_table(self.sheet.get("g1"), cat))
            g2_rows = RULES.table(RULES.resolve_table(self.sheet.get("g2"), cat))
            self.lbl_group1_desc.Text = RULES.description_of(g1_rows, g1)
            self.lbl_group2_desc.Text = RULES.description_of(g2_rows, g2)

            values = self._collect()
            mark, tm_warning = self._compute_type_mark(values)
            values["type_mark"] = mark

            self._show_existing_mark(mark)

            family_name = u""
            if self.sheet.get("schema") == "loadable":
                family_name = RULES.compose_family_name(self.sheet, values)
            type_name = RULES.compose_type_name(self.sheet, values)

            self.txt_preview_family.Text = family_name
            self.txt_preview_type.Text = type_name

            problems = RULES.validate(self.sheet, values)
            problems.extend(self._collision_problems(family_name, type_name))
            if tm_warning:
                problems.insert(0, tm_warning)

            self.lbl_validation.Text = u"\n".join(problems)
            self.btn_apply.IsEnabled = not problems

            self._values = values
            self._family_name = family_name
            self._type_name = type_name
        finally:
            self._loading = False

    def _show_existing_mark(self, mark):
        current = self.ctx.get("current_type_mark") or u""
        if not self.chk_tm_write.IsChecked:
            self.lbl_tm_existing.Text = u""
            return
        if current and current != mark:
            self.lbl_tm_existing.Text = (
                u"The type already carries Type Mark '{0}'. "
                u"Applying will overwrite it.").format(current)
        else:
            self.lbl_tm_existing.Text = u""

    def _collision_problems(self, family_name, type_name):
        """Un nome gia' in uso fa fallire la rinomina: meglio fermarsi prima."""
        problems = []
        ctx = self.ctx

        if self.sheet.get("schema") == "loadable" and family_name:
            taken = ctx.get("existing_family_names") or set()
            if family_name != ctx.get("current_family") and family_name in taken:
                problems.append(
                    u"A family named '{0}' already exists in the project.".format(
                        family_name))

        if type_name:
            taken = ctx.get("existing_type_names") or set()
            if type_name != ctx.get("current_type") and type_name in taken:
                problems.append(
                    u"A type named '{0}' already exists in this family.".format(
                        type_name))

        return problems

    # -- uscita -------------------------------------------------------------

    def _on_cancel(self, sender, args):
        self.result = None
        self.Close()

    def _on_apply(self, sender, args):
        values = getattr(self, "_values", None)
        if values is None:
            return
        self.result = {
            "family_name": self._family_name,
            "type_name": self._type_name,
            "type_mark": values.get("type_mark", u""),
            "write_type_mark": bool(self.chk_tm_write.IsChecked),
            "fire_rating": values.get("rei", u""),
            "in_place": values.get("in_place"),
            "schema": self.sheet.get("schema"),
        }
        self.Close()
