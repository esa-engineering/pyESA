# -*- coding: utf-8 -*-
"""
clashfinder_quick_ui.py - Quick Clash: esecuzione rapida dei test di Clash Finder.

A sinistra l'elenco dei test (un clic sceglie quello da eseguire), a destra sorgente
nel modello aperto e link di Set A e Set B: vengono dal test e si possono cambiare
qui per l'esecuzione, senza salvarli. Categorie, filtri e regole restano quelli del
test; si modificano in Clash Finder. Una riga di stato riporta conteggi e note.

Libreria dei test e impostazioni sono quelle di Clash Finder (sezione ESA_ClashFinder).
"""

import clashfinder_rules as cr
import clashfinder_engine as ce
import clashfinder_ui as cu


class QuickClashForm(cu.ClashFinderForm):

    XAML_FILE = 'ClashFinder_quick.xaml'

    def _categories_in_use(self):
        # Nessun elenco di categorie in questa finestra.
        return set()

    def _find_controls(self):
        for name in ("lst_tests", "txt_summary_state", "txt_links_hint",
                     "btn_refresh", "btn_close", "btn_run", "btn_help",
                     "cbo_source_a", "cbo_source_b", "lst_links_a", "lst_links_b",
                     "btn_links_a", "btn_links_b", "pop_links_a", "pop_links_b"):
            setattr(self, name, self._root.FindName(name))

    def _init_controls(self):
        for side in cu.SIDES:
            combo = getattr(self, "cbo_source_" + side)
            for _, label in ce.SOURCE_LABELS:
                combo.Items.Add(label)
            combo.SelectedIndex = 0
            combo.SelectionChanged += lambda s, e: self._changed()
            self._wire_links_dropdown(side)
        self.lst_tests.SelectionChanged += self.OnTestSelected
        self.btn_refresh.Click += lambda s, e: self._request_summary()
        self.btn_close.Click += lambda s, e: self.Close()
        self.btn_help.Click += self.OnHelp
        self.btn_run.Click += self.OnRun
        self.Closing += self.OnWindowClosing
        self.Closed += self.OnWindowClosed

    # Nessun editor: le modifiche di sorgenti e link valgono solo per l'esecuzione.

    def _mark_dirty(self):
        pass

    def _update_editor_state(self):
        pass

    def _request_param_names(self, side):
        pass

    def _load_editor(self, test, name):
        self._editor_loading = True
        try:
            self.edit_test = test
            self.edit_name = name
            for side in cu.SIDES:
                set_data = test[cu.SET_KEYS[side]]
                getattr(self, "cbo_source_" + side).SelectedIndex = \
                    cu.SOURCE_KEYS.index(set_data["source"])
                self.edit_links[side] = set(set_data["links"])
            self._fill_link_lists()
        finally:
            self._editor_loading = False
        self.dirty = False
        self._request_summary()

    def _read_editor_raw(self):
        test = cr.copy_test(self.edit_test)
        for side in cu.SIDES:
            set_data = test[cu.SET_KEYS[side]]
            set_data["source"] = self._source(side)
            set_data["links"] = sorted(self.edit_links[side], key=lambda n: n.lower())
        return test

    # ------------------------------------------------------------- riepilogo

    def _show_summary(self, test, summary):
        """Nessuna tabella: tre righe di stato, fase, Set A, Set B.

        Verde: fase presente e insieme con elementi; grigio: insieme vuoto; rosso: solo la
        vista senza fase.
        """
        if self.edit_name is None:
            self.txt_summary_state.Foreground = cu.GRAY_BRUSH
            self.txt_summary_state.Text = u"No test selected."
            return
        lines = [(summary.phase_text,
                  cu.RED_BRUSH if summary.phase_problem else cu.GREEN_BRUSH)]
        for side in cu.SIDES:
            total = summary.host[side] + summary.linked[side]
            lines.append((u"Set {}: {} from {}.".format(
                side.upper(), self._count_text(summary, side),
                cr.describe_sources(test[cu.SET_KEYS[side]])),
                cu.GREEN_BRUSH if total else cu.GRAY_BRUSH))
        # "Set A is empty" ripete il conteggio gia' mostrato sopra: si toglie, le altre note restano.
        for note in summary.notes:
            parts = [p for p in note.split(u": ", 1)[-1].split(u"; ")
                     if p not in (u"Set A is empty", u"Set B is empty")]
            if parts:
                lines.append((u"; ".join(parts), cu.GRAY_BRUSH))
        cu.set_status_lines(self.txt_summary_state, lines)

    @staticmethod
    def _count_text(summary, side):
        host = summary.host[side]
        linked = summary.linked[side]
        if linked:
            return u"{} ({} in links)".format(cu._plural(host + linked, u"element"), linked)
        return cu._plural(host, u"element")

    def OnRun(self, sender, args):
        if self.edit_name is None:
            cu._info(u"Select a test first.")
            return
        cu.ClashFinderForm.OnRun(self, sender, args)


def show_quick_window(doc, bridge, owner_handle=None):
    form = QuickClashForm(doc, bridge)
    if owner_handle is not None:
        try:
            from System.Windows.Interop import WindowInteropHelper
            WindowInteropHelper(form).Owner = owner_handle
        except Exception:
            pass
    form.Show()
    return form
