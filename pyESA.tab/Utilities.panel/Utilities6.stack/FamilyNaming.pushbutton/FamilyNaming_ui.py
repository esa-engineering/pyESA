# -*- coding: utf-8 -*-
"""Finestra del tool di nomenclatura.

Carica il XAML con XamlReader e costruisce a runtime i campi che cambiano da
una categoria all'altra: il blocco dimensionale, i criteri di codifica
specifici, le note. Tutto quello che l'utente vede e' in inglese; i commenti
restano in italiano come nel resto dell'estensione.

La finestra non tocca Revit: riceve dallo script un contesto gia' pronto
(categoria riconosciuta, misure lette dai parametri, Type Mark gia' in uso,
nomi gia' assegnati) e restituisce il risultato in self.result.

Cosa l'utente NON sceglie, perche' lo decide lo script leggendo il modello:
la categoria e il suo codice, la famiglia di sistema, se l'elemento e'
modellato in place, e le misure che si riescono a leggere dai parametri di
tipo. Sono tutti mostrati come testo, non come campi.
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
COMPUTED_BG = _brush(0xF0, 0xF0, 0xF0)

VISIBLE = Visibility.Visible
COLLAPSED = Visibility.Collapsed


class NamingWindow(Window):
    """Finestra unica del tool. self.result e' None se l'utente annulla."""

    def __init__(self, ctx):
        self.ctx = ctx
        self.sheet = ctx["sheet"]
        self.cat = ctx.get("cat_code") or u""
        self.result = None

        # controlli costruiti a runtime, per chiave logica
        self._ctl = {}
        # elementi che compongono la riga di un controllo, per poterla
        # nascondere quando il campo non si applica al caso scelto
        self._row_parts = {}
        # testo di esempio e suo stato: vedi la nota sul watermark
        self._wm_text = {}
        self._wm_on = {}
        # ultimo Type Mark proposto, per non sovrascrivere quello che
        # l'utente ha eventualmente scritto a mano
        self._tm_suggested = None

        self._loading = True
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
            "lbl_element", "lbl_revit_category", "lbl_category",
            "grd_inplace", "lbl_inplace", "lbl_detect_warning",
            "cmb_group1", "cmb_group2",
            "lbl_group1", "lbl_group2", "lbl_group1_desc", "lbl_group2_desc",
            "txt_typemark", "lbl_tm_hint", "chk_tm_write", "lbl_tm_existing",
            "brd_specific", "pnl_specific",
            "brd_dimensions", "pnl_dimensions", "lbl_dim_hint",
            "lbl_desc_family", "txt_desc_family",
            "grd_desc_type", "lbl_desc_type", "txt_desc_type",
            "brd_notes", "lbl_notes_header", "pnl_notes",
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
        is_loadable = sheet.get("schema") == "loadable"

        self._show_detected()

        # una scheda di sistema lavora direttamente sul tipo: non compone un
        # nome famiglia, quindi ne' l'anteprima ne' la descrizione del tipo
        # hanno un senso da mostrare
        if not is_loadable:
            self.grd_preview_family.Visibility = COLLAPSED
            self.grd_desc_type.Visibility = COLLAPSED
            self.lbl_desc_family.Text = u"Description"
            self.txt_desc_type.Text = u""
        elif sheet.get("dim") != "none":
            # dove il blocco dimensionale entra nel nome del tipo, il tipo si
            # distingue gia' dalle misure e la descrizione diventa un di piu'
            self.lbl_desc_type.Text = u"Type description  (optional)"

        self._fill_groups()
        self._build_specific()
        self._build_dimensions()
        self._build_notes()

        self.cmb_group1.SelectionChanged += self._on_group_changed
        self.cmb_group2.SelectionChanged += self._on_group_changed
        self.txt_typemark.TextChanged += self._on_changed
        self.chk_tm_write.Click += self._on_changed
        self.txt_desc_family.TextChanged += self._on_changed
        self.txt_desc_type.TextChanged += self._on_changed
        self.btn_cancel.Click += self._on_cancel
        self.btn_apply.Click += self._on_apply

    def _show_detected(self):
        """Tutto quello che lo script ha riconosciuto da solo, in sola lettura."""
        ctx = self.ctx
        sheet = self.sheet

        self.lbl_element.Text = ctx["element_label"]
        self.lbl_revit_category.Text = ctx.get("revit_category") or u"-"

        cat_rows = RULES.table(sheet["cat_table"])
        label = RULES.label_of(cat_rows, self.cat)
        self.lbl_category.Text = u"{0}  -  {1}".format(self.cat, label)

        # la riga dell'in place compare solo quando serve dirlo
        if ctx.get("in_place"):
            self.grd_inplace.Visibility = VISIBLE
        else:
            self.grd_inplace.Visibility = COLLAPSED

        # se la famiglia di sistema non e' stata riconosciuta con certezza
        # conviene dirlo, perche' l'utente non ha un menu per correggerla
        if ctx.get("cat_uncertain"):
            self.lbl_detect_warning.Text = (
                u"The system family could not be matched against the "
                u"classification table, so the first entry was assumed. "
                u"Check the category code before applying.")
        else:
            self.lbl_detect_warning.Visibility = COLLAPSED

    # -- menu a tendina -----------------------------------------------------

    def _fill_groups(self):
        """Popola Group1 e Group2 per la categoria riconosciuta.

        Su Walls, Roofs e Stairs le liste dipendono dalla famiglia di sistema,
        che pero' e' fissata dal rilevamento: si risolvono una volta sola.
        """
        g1_rows = RULES.table(RULES.resolve_table(self.sheet.get("g1"), self.cat))
        g2_rows = RULES.table(RULES.resolve_table(self.sheet.get("g2"), self.cat))

        self._reload_combo(self.cmb_group1, g1_rows)
        self._reload_combo(self.cmb_group2, g2_rows)

        # sulle sette categorie caricabili a Group1 condiviso il primo gruppo
        # non descrive com'e' fatta la famiglia ma da dove viene
        shared = self.sheet.get("g1") == MAP.TBL_SHARED_G1
        self.lbl_group1.Text = u"Group 1 (origin)" if shared else u"Group 1"

    def _reload_combo(self, combo, rows):
        combo.Items.Clear()
        for row in rows:
            combo.Items.Add(self._item(row))
        combo.IsEnabled = len(rows) > 0
        if rows:
            combo.SelectedIndex = 0

    def _item(self, row):
        code = row[0]
        label = row[1] if len(row) > 1 else code
        return u"{0}  -  {1}".format(code, label)

    def _code(self, combo):
        item = combo.SelectedItem
        if not item:
            return u""
        return str(item).split(u"  -  ")[0].strip()

    # -- criteri di codifica specifici della categoria ----------------------

    def _build_specific(self):
        panel = self.pnl_specific
        panel.Children.Clear()
        self._ctl = {}
        self._row_parts = {}
        self._wm_text = {}
        self._wm_on = {}

        built = 0

        # Il prefisso manuale serve solo dove la cifra del Type Mark verrebbe
        # dal Group1 e il Group1 vale Other. La riga resta nascosta fino ad
        # allora, e quando compare e' obbligatoria.
        if self.sheet.get("tm_from") == "g1_digit":
            self._add_combo(
                panel, "tm_prefix", u"Manual TM prefix",
                RULES.table(MAP.TBL_MATERIAL), required=True,
                hint=u"Group 1 is Other, so the whole Type Mark prefix is "
                     u"manual. Pick the material family.")
            self._set_row_visible("tm_prefix", False)
            built += 1

        for block in self.sheet.get("blocks", ()):
            if block == "description":
                continue                       # ha una sezione sua
            built += self._add_block(panel, block)

        self.brd_specific.Visibility = VISIBLE if built else COLLAPSED

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
                panel, "cw_hosted", u"CW hosted",
                u"nested in a curtain wall panel",
                hint=u"The panel itself is named on the Curtain Panels sheet.")
            return 1
        if block == "entrance":
            self._add_check(
                panel, "entrance", u"Entrance", u"main entrance",
                hint=u"Allowed only with Group 2 = EX.")
            return 1
        if block == "rei":
            self._add_text(
                panel, "rei", u"REI", watermark=u"REI60, EI90, RE60",
                hint=u"Full designation as certified. It is written into the "
                     u"FireRating parameter.")
            return 1
        if block == "wi":
            # W e I non sono sotto-opzioni del REI: sono due blocchi del nome
            # con una regola loro, quindi portano la propria etichetta a
            # sinistra come tutti gli altri campi. La regola che li abilita
            # e' pero' una sola e vale per entrambi, percio' la nota sta in
            # fondo e li nomina tutti e due invece di sembrare dell'ultimo.
            self._add_check(
                panel, "waterproof", u"W  waterproof", u"included in the element")
            self._add_check(
                panel, "insulation", u"I  insulation", u"included in the element")
            self._hint(panel, "insulation", self._wi_note())
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
                panel, "ln_shape", u"Thickness mode", u"same as run (SR)",
                hint=u"The landing inherits the thickness from the run, and SR "
                     u"replaces the whole dimensional block.")
            return 1
        if block == "rm_shape":
            self._add_check(
                panel, "rm_shape", u"Shape", u"solid (SL)",
                hint=u"The ramp is solid down to the ground, and SL replaces "
                     u"the whole dimensional block.")
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
                panel, "custom", u"Custom", u"drawn to measure (CT)",
                hint=u"Not compatible with Manufacturer.")
            return 1
        if block == "manufacturer":
            self._add_text(
                panel, "manufacturer", u"Manufacturer", watermark=u"e.g. Schuco",
                hint=u"The company that makes the object.")
            return 1
        if block == "brand":
            self._add_text(
                panel, "brand", u"Brand", watermark=u"e.g. Dior",
                hint=u"The CLIENT brand this element is a standard of, not the "
                     u"manufacturer product line.")
            return 1
        return 0

    def _wi_note(self):
        """Nota comune ai suffissi W e I, con i codici che li ammettono.

        L'elenco lo dichiara la scheda della categoria: sulle pareti sono tre
        codici, sui solai e sui controsoffitti uno solo. Altrove i due strati
        sono impliciti e non si dichiarano.
        """
        allowed = self.sheet.get("wi_allowed_g2") or ()
        if not allowed:
            return (u"W and I are allowed only on the Group 2 codes where they "
                    u"are not already implicit.")
        if len(allowed) == 1:
            codes = allowed[0]
        else:
            codes = u"{0} or {1}".format(u", ".join(allowed[:-1]), allowed[-1])
        return (u"W and I are allowed only with Group 2 = {0}. On the other "
                u"codes waterproofing and insulation are implicit and are not "
                u"declared in the name.").format(codes)

    # -- blocco dimensionale ------------------------------------------------

    DIM_HINTS = {
        "wxh": u"Width x height of the opening, in mm.",
        "wxdxh": u"Plan footprint in mm. Height only when it tells two types apart.",
        "wxl": u"Bay size in mm. 250x500 standard, 320x500 accessible.",
        "wxd_or_d": u"Section in mm. A round section is quoted by diameter only.",
        "bxh_or_d": u"Profile section in mm. A round profile is quoted by diameter only.",
        "wxt": u"Width x thickness in mm.",
        "t": u"Thickness in mm.",
        "h": u"Height in mm.",
        "nfin_t": u"Number of finished faces and total thickness. A single "
                  u"layer element is always 0F, even when that layer is the "
                  u"finishing.",
    }

    def _build_dimensions(self):
        panel = self.pnl_dimensions
        panel.Children.Clear()

        kind = self.sheet.get("dim")
        if kind == "none":
            self.brd_dimensions.Visibility = COLLAPSED
            return
        self.brd_dimensions.Visibility = VISIBLE

        hint = self.DIM_HINTS.get(kind)
        if kind == "free":
            hint = self.sheet.get("dim_hint") \
                or u"Use the standard designation when one exists."
        self.lbl_dim_hint.Text = hint or u""

        # Le facce finite si contano dalla stratigrafia, non si scelgono: il
        # campo mostra il risultato. Solo se la deduzione non riesce si
        # riapre la scelta, altrimenti non si potrebbe procedere.
        if kind == "nfin_t":
            deduced = self.ctx.get("nfin_default")
            if deduced:
                self._add_computed(panel, "nfin", u"Finished faces", deduced,
                                   note=u"counted from the layers")
            else:
                self._add_combo(
                    panel, "nfin", u"Finished faces",
                    RULES.table(self.sheet.get("nfin_table") or MAP.TBL_NFIN),
                    required=True,
                    hint=u"The layer structure could not be read, so the value "
                         u"has to be picked by hand.")

        self._rebuild_dim_fields(panel, kind, force=True)

    def _rebuild_dim_fields(self, panel, kind, force=False):
        """Rifa' le righe delle misure. Serve solo quando cambia il loro elenco.

        L'elenco dipende da una cosa sola: se il Group1 dichiara una sezione
        tonda si chiede il diametro, altrimenti le due misure della sezione.
        """
        round_section = self._code(self.cmb_group1) in (self.sheet.get("round_g1") or ())
        if not force and round_section == self._dim_round:
            return
        self._dim_round = round_section

        for key in list(self._ctl.keys()):
            if key.startswith("dim:"):
                del self._ctl[key]
                self._row_parts.pop(key, None)
                self._wm_text.pop(key, None)
                self._wm_on.pop(key, None)

        keep = [c for c in panel.Children if getattr(c, "Tag", None) != "dimrow"]
        panel.Children.Clear()
        for child in keep:
            panel.Children.Add(child)

        defaults = self.ctx.get("dim_defaults", {})

        for key, label, required in RULES.dim_fields(kind, round_section):
            ckey = "dim:" + key
            value = defaults.get(key)

            # una misura che si legge dal tipo non si digita: il nome deve
            # dire quello che l'elemento e', non quello che si vorrebbe
            if value:
                self._add_computed(panel, ckey, label, value, unit=u"mm",
                                   note=u"read from the type", tag="dimrow")
                continue

            if kind == "free":
                self._add_text(panel, ckey, label,
                               watermark=self.sheet.get("dim_hint", u""),
                               required=required, tag="dimrow")
            else:
                self._add_text(panel, ckey, label, unit=u"mm", narrow=True,
                               required=required, tag="dimrow",
                               hint=u"No type parameter carries this measure, "
                                    u"so it has to be typed.")

    # -- note ---------------------------------------------------------------

    def _build_notes(self):
        panel = self.pnl_notes
        panel.Children.Clear()

        notes = RULES.visible_notes(self.ctx["sheet_id"])
        if not notes:
            # su alcune categorie non resta nessuna nota utile: il riquadro
            # sparisce del tutto invece di restare vuoto
            self.brd_notes.Visibility = COLLAPSED
            return

        for kind, head, body in notes:
            if kind == "section":
                panel.Children.Add(self._note_block(head, bold=True,
                                                    color=ACCENT, top=10))
                continue
            if head:
                panel.Children.Add(self._note_block(head, bold=True, top=6))
            if body:
                panel.Children.Add(self._note_block(body, color=GRAY))

    def _note_block(self, text, bold=False, color=None, top=0):
        block = TextBlock()
        block.Text = text
        block.FontSize = 11
        block.TextWrapping = TextWrapping.Wrap
        block.Margin = Thickness(0, top, 0, 1)
        if bold:
            block.FontWeight = FontWeights.Bold
        if color is not None:
            block.Foreground = color
        return block

    # -- costruttori di righe -----------------------------------------------

    def _row(self, panel, key, label_text, control, unit=None, required=False,
             tag=None, tight=False):
        """Una riga etichetta + controllo + unita' di misura.

        Con tight la colonna del controllo si stringe su quanto serve e lo
        spazio avanzato finisce in una colonna vuota in coda: cosi' la nota
        resta attaccata alla casella invece di essere spinta al bordo destro.
        Serve sui campi stretti, cioe' le misure e i valori calcolati.
        """
        grid = Grid()
        grid.Margin = Thickness(0, 5, 0, 0)
        if tag:
            grid.Tag = tag

        widths = [GridLength(LABEL_COL, GridUnitType.Pixel)]
        if tight:
            widths.append(GridLength(1, GridUnitType.Auto))    # controllo
            widths.append(GridLength(1, GridUnitType.Auto))    # unita' e nota
            widths.append(GridLength(1, GridUnitType.Star))    # spazio residuo
        else:
            widths.append(GridLength(1, GridUnitType.Star))    # controllo
            widths.append(GridLength(1, GridUnitType.Auto))    # unita' e nota

        for width in widths:
            column = ColumnDefinition()
            column.Width = width
            grid.ColumnDefinitions.Add(column)

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
        self._row_parts.setdefault(key, []).append(grid)
        return grid

    def _hint(self, panel, key, text, tag=None):
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
        self._row_parts.setdefault(key, []).append(block)

    def _set_row_visible(self, key, visible):
        state = VISIBLE if visible else COLLAPSED
        for part in self._row_parts.get(key, ()):
            part.Visibility = state

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
        self._row(panel, key, label, combo, required=required, tag=tag)
        self._hint(panel, key, hint, tag=tag)
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
        self._row(panel, key, label, box, unit=unit, required=required, tag=tag,
                  tight=narrow)
        self._hint(panel, key, hint, tag=tag)
        return box

    def _add_computed(self, panel, key, label, value, unit=None, note=None,
                      tag=None):
        """Un valore che viene dal modello: si mostra, non si modifica."""
        box = TextBox()
        box.Text = value
        box.IsReadOnly = True
        box.Width = 70
        box.HorizontalAlignment = HorizontalAlignment.Left
        box.Padding = Thickness(3, 2, 3, 2)
        box.Background = COMPUTED_BG
        box.Foreground = GRAY
        self._ctl[key] = box
        unit_text = unit
        if note:
            unit_text = u"{0} ({1})".format(unit, note) if unit else u"({0})".format(note)
        self._row(panel, key, label, box, unit=unit_text, required=True,
                  tag=tag, tight=True)
        return box

    def _add_check(self, panel, key, label, caption, hint=None, tag=None):
        """Casella con la sua etichetta a sinistra, come gli altri campi.

        Una casella nuda indentata sotto un campo etichettato si legge come
        una sua sotto-opzione, che e' esattamente quello che W e I non sono.
        """
        check = CheckBox()
        check.Content = caption
        check.VerticalAlignment = VerticalAlignment.Center
        check.Click += self._on_changed
        self._ctl[key] = check
        self._row(panel, key, label, check, required=True, tag=tag)
        self._hint(panel, key, hint, tag=tag)
        return check

    # -- watermark ----------------------------------------------------------
    #
    # Il testo di esempio e' grigio, sparisce appena si scrive e non viene mai
    # restituito come valore: un campo lasciato intatto resta vuoto.
    #
    # Lo stato sta in due dizionari della finestra e non sul controllo, perche'
    # IronPython non lascia attaccare attributi Python a un oggetto .NET.

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

    def _value_of(self, key):
        """Testo di un campo, sia editabile sia calcolato."""
        control = self._ctl.get(key)
        if control is None:
            return u""
        if isinstance(control, ComboBox):
            return self._combo_of(key)
        return self._text_of(key)

    def _check_of(self, key):
        check = self._ctl.get(key)
        if check is None or not check.IsEnabled:
            return False
        return bool(check.IsChecked)

    # -- eventi -------------------------------------------------------------

    def _on_group_changed(self, sender, args):
        if self._loading:
            return
        self._refresh()

    def _on_changed(self, sender, args):
        if self._loading:
            return
        self._refresh()

    # -- stato e anteprima --------------------------------------------------

    def _apply_conditional_rules(self, g1, g2):
        """Accende, spegne e nasconde i campi secondo le regole della scheda."""
        nested = g1 in MAP.NESTED_G1
        manual_tm = g1 == MAP.MANUAL_TM_G1
        needs_manual = manual_tm and self.sheet.get("tm_from") == "g1_digit"

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
            ok = cats is None or self.cat in cats
            self._ctl["underside"].IsEnabled = ok
            if not ok:
                self._ctl["underside"].SelectedIndex = 0

        # il campo Use si accende solo con Group1 = Other
        if "use" in self._ctl:
            self._ctl["use"].IsEnabled = manual_tm
            if not manual_tm:
                self._ctl["use"].SelectedIndex = 0

        # il prefisso manuale compare solo quando serve davvero, e allora
        # e' obbligatorio: il campo Use, se compilato, lo rimpiazza
        if "tm_prefix" in self._ctl:
            show = needs_manual and not self._combo_of("use")
            self._set_row_visible("tm_prefix", show)

        # Manufacturer e Custom non convivono
        if "manufacturer" in self._ctl:
            custom = self._check_of("custom") or g1 == u"CT"
            self._ctl["manufacturer"].IsEnabled = not custom

        # due casi spengono l'intero blocco dimensionale: SR e SL, che ne
        # prendono il posto nel nome, e il pannello di sistema vuoto
        substitute = self._check_of("ln_shape") or self._check_of("rm_shape")
        no_dim = g1 in (self.sheet.get("dim_optional_g1") or ())
        for key in self._ctl:
            if not key.startswith("dim:"):
                continue
            self._set_row_visible(key, not (substitute or no_dim))

    def _collect(self):
        values = {
            "cat": self.cat,
            "g1": self._code(self.cmb_group1),
            "g2": self._code(self.cmb_group2),
            "in_place": bool(self.ctx.get("in_place")),
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
            "type_description": u"" if self.sheet.get("schema") != "loadable"
            else (self.txt_desc_type.Text or u"").strip(),
        }

        dim = {}
        for key in self._ctl:
            if key.startswith("dim:"):
                dim[key[4:]] = self._value_of(key)
        if "nfin" in self._ctl:
            dim["nfin"] = self._value_of("nfin")
        values["dim"] = dim

        return values

    def _compute_type_mark(self, values):
        """Propone il prossimo sequenziale libero senza calpestare l'utente.

        La proposta si riscrive solo se nel campo c'e' ancora la proposta
        precedente: se l'utente ha scritto qualcosa di suo, resta il suo.
        """
        sheet = self.sheet
        use_code = values.get("use")

        # Con Group1 = Other il Type Mark e' interamente manuale, prefisso
        # compreso: il 9xx riportato in tabella accanto a Other non si usa.
        manual = None
        needs_manual = (values.get("g1") == MAP.MANUAL_TM_G1
                        and sheet.get("tm_from") == "g1_digit")
        if needs_manual:
            manual = use_code or self._combo_of("tm_prefix")

        prefix, digits = RULES.type_mark_prefix(
            sheet, values["cat"], values["g1"], values["g2"], manual)

        incomplete = (needs_manual and not manual) or \
            (prefix.endswith(u"-") and digits == 2)

        current = (self.txt_typemark.Text or u"").strip()

        if incomplete:
            self.lbl_tm_hint.Text = u""
            return current, u"Pick the manual Type Mark prefix."

        self.lbl_tm_hint.Text = u"prefix {0} + {1} digits".format(prefix, digits)

        warning = None
        sequential = RULES.next_sequential(
            self.ctx.get("existing_marks", ()), prefix, digits)
        if RULES.sequential_overflow(sequential, digits):
            warning = (u"All {0}-digit sequentials for prefix {1} are taken. "
                       u"Set the Type Mark by hand.").format(digits, prefix)
            suggestion = u""
        else:
            suggestion = RULES.format_type_mark(prefix, digits, sequential)

        # si sovrascrive solo la proposta precedente, mai un valore digitato
        if not current or current == self._tm_suggested:
            self.txt_typemark.Text = suggestion
            self._tm_suggested = suggestion
            return suggestion, warning

        return current, warning

    def _refresh(self):
        if self._loading:
            return
        self._loading = True
        try:
            g1 = self._code(self.cmb_group1)
            g2 = self._code(self.cmb_group2)

            self._rebuild_dim_fields(self.pnl_dimensions, self.sheet.get("dim"))
            self._apply_conditional_rules(g1, g2)

            g1_rows = RULES.table(RULES.resolve_table(self.sheet.get("g1"), self.cat))
            g2_rows = RULES.table(RULES.resolve_table(self.sheet.get("g2"), self.cat))
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
        if self.chk_tm_write.IsChecked and current and current != mark:
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
