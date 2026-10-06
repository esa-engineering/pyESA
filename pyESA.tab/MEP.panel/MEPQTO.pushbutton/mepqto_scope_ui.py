# -*- coding: utf-8 -*-
"""
mepqto_scope_ui.py - scelta dei modelli, dei workset e delle categorie da leggere.

Si apre prima di ogni lettura del modello: all'avvio del tool e dal pulsante Models
and Categories... della finestra del computo. Il modello aperto si legge sempre; si
scelgono le istanze di link da leggere in piu', i workset i cui elementi si escludono
e le categorie.

I workset si escludono per nome, in tutti i modelli letti: l'elenco riunisce i workset
utente del modello aperto e dei link spuntati, e si aggiorna quando cambiano i link.
Una spunta esclude: un workset nuovo, mai visto, si legge.

Le spunte delle categorie e quelle dei workset si possono salvare come set con nome
(_SetPicker): i set stanno nella config pyRevit dell'utente e si scrivono subito,
tramite le funzioni di salvataggio passate dal chiamante in ScopeSettings.

La finestra restituisce uno ScopeChoice, oppure None se annullata.
"""

import os
from collections import OrderedDict

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System import Action
from System.IO import FileStream, FileMode
from System.Windows import (Window, MessageBox, MessageBoxButton, MessageBoxImage,
                            MessageBoxResult, WindowStartupLocation)
from System.Windows.Controls import CheckBox, TextBlock
from System.Windows.Markup import XamlReader
from System.Windows.Threading import DispatcherPriority

from pyrevit import script

import mepqto_rules as qr

XAML_FILE_NAME = 'MEPQTO_scope.xaml'
TITLE = "Models and categories"
# Separatore fra nome del set e valori nella config: non ammesso nei nomi dei set.
SET_NAME_SEPARATOR = u"::"


def _set_text(box, text):
    """Testo di una casella in un TextBlock: come stringa, WPF userebbe il primo '_'
    come tasto di scelta rapida e lo toglierebbe ("MEP_A.rvt" -> "MEPA.rvt")."""
    block = TextBlock()
    block.Text = text
    box.Content = block


class ScopeSettings(object):
    """Dati iniziali della finestra e funzioni del chiamante.

    links: LinkSource del modello aperto; checked_link_ids: UniqueId spuntati;
    rules: (chiave, etichetta, tipo di misura); checked_keys: categorie spuntate;
    list_worksets(links) -> {nome workset: [modelli]} per il modello aperto e i link;
    excluded_worksets: nomi spuntati (esclusi); *_sets: {nome set: [valori]};
    *_set_name: nome mostrato nella casella; save_*_sets(sets) salva nella config.
    """

    def __init__(self, host_label=u"", links=(), checked_link_ids=(), rules=(),
                 checked_keys=(), category_sets=None, category_set_name=u"",
                 save_category_sets=None, list_worksets=None, excluded_worksets=(),
                 workset_sets=None, workset_set_name=u"", save_workset_sets=None):
        self.host_label = host_label
        self.links = list(links)
        self.checked_link_ids = list(checked_link_ids or [])
        self.rules = list(rules)
        self.checked_keys = list(checked_keys or [])
        self.category_sets = OrderedDict(category_sets or [])
        self.category_set_name = category_set_name or u""
        self.save_category_sets = save_category_sets or (lambda sets: None)
        self.list_worksets = list_worksets or (lambda links: OrderedDict())
        self.excluded_worksets = list(excluded_worksets or [])
        self.workset_sets = OrderedDict(workset_sets or [])
        self.workset_set_name = workset_set_name or u""
        self.save_workset_sets = save_workset_sets or (lambda sets: None)


class ScopeChoice(object):
    """Esito della finestra: link da leggere, categorie, workset esclusi e i nomi dei
    set rimasti nelle caselle."""

    def __init__(self, links, category_keys, set_name, excluded_worksets=(),
                 workset_set_name=u""):
        self.links = list(links)
        self.category_keys = list(category_keys)
        self.set_name = set_name
        self.excluded_worksets = list(excluded_worksets)
        self.workset_set_name = workset_set_name


class _SetPicker(object):
    """Set con nome di una lista di spunte: ComboBox modificabile + Save Set / Delete Set.

    Scegliere un set dalla tendina lo applica; dopo, la selezione si toglie lasciando
    il nome nella casella, cosi' riscegliere lo stesso set lo riapplica. La ricerca
    testuale della tendina e' spenta nell'XAML: scrivere il nome di un set esistente
    (per salvarne uno nuovo) non cambia le spunte.
    """

    def __init__(self, owner, combo, save_button, delete_button, noun, sets, set_name,
                 checked, apply, save, allow_empty):
        self.owner = owner
        self.combo = combo
        self.noun = noun
        self.sets = sets
        self._checked = checked
        self._apply = apply
        self._save = save
        self._allow_empty = allow_empty
        self._loading = False
        combo.SelectionChanged += self.OnChosen
        save_button.Click += self.OnSave
        delete_button.Click += self.OnDelete
        self.fill(set_name if set_name in sets else u"")

    @property
    def name(self):
        return (self.combo.Text or u"").strip()

    def fill(self, text):
        self._loading = True
        try:
            self.combo.ItemsSource = list(self.sets.keys())
        finally:
            self._loading = False
        self.clear_selection(text)

    def clear_selection(self, text=None):
        if text is None:
            text = self.combo.Text
        self._loading = True
        try:
            self.combo.SelectedIndex = -1
            self.combo.Text = text
        finally:
            self._loading = False

    def _persist(self):
        try:
            self._save(self.sets)
        except Exception as error:
            MessageBox.Show(u"The {} sets could not be saved:\n{}".format(self.noun, error),
                            TITLE, MessageBoxButton.OK, MessageBoxImage.Warning)

    def OnChosen(self, sender, args):
        if self._loading:
            return
        name = self.combo.SelectedItem
        if name is not None and name in self.sets:
            self._apply(self.sets[name])
            # Dopo che la ComboBox ha finito di aggiornare il testo.
            self.owner.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                                              Action(lambda: self.clear_selection(name)))

    def OnSave(self, sender, args):
        name = self.name
        values = self._checked()
        if not name:
            MessageBox.Show(u"Type a name for the {0} set in {1} set.".format(
                self.noun, self.noun.capitalize()), TITLE)
            return
        if SET_NAME_SEPARATOR in name:
            MessageBox.Show(u"The name of a {} set cannot contain '{}'.".format(
                self.noun, SET_NAME_SEPARATOR), TITLE)
            return
        if not values and not self._allow_empty:
            MessageBox.Show(u"Tick at least one {} before saving the set.".format(self.noun),
                            TITLE)
            return
        if name in self.sets:
            answer = MessageBox.Show(
                u"The {} set '{}' already exists.\n\nReplace it with the ticked "
                u"{}s?".format(self.noun, name, self.noun), TITLE, MessageBoxButton.YesNo,
                MessageBoxImage.Question)
            if answer != MessageBoxResult.Yes:
                return
        self.sets[name] = list(values)
        self._persist()
        self.fill(name)

    def OnDelete(self, sender, args):
        name = self.name
        if name not in self.sets:
            MessageBox.Show(u"There is no {} set named '{}'.".format(self.noun, name), TITLE)
            return
        answer = MessageBox.Show(u"Delete the {} set '{}'?".format(self.noun, name), TITLE,
                                 MessageBoxButton.YesNo, MessageBoxImage.Question)
        if answer != MessageBoxResult.Yes:
            return
        del self.sets[name]
        self._persist()
        self.fill(u"")


class ScopeForm(Window):

    def __init__(self, settings):
        self.result = None
        self._loading = True
        self._list_worksets = settings.list_worksets
        # Nomi dei workset esclusi, anche di quelli non in elenco (link non spuntato):
        # rispuntando il link tornano spuntati.
        self._excluded = set(settings.excluded_worksets)
        self._workset_names = OrderedDict()
        self._workset_boxes = []
        self._load_xaml()

        self.txt_host.Text = u"The open model is always read: {}".format(settings.host_label)
        checked_links = set(settings.checked_link_ids)
        self._link_boxes = []
        for link in settings.links:
            box = CheckBox()
            box.Tag = link
            if link.loaded:
                _set_text(box, link.label)
                box.IsChecked = link.unique_id in checked_links
            else:
                _set_text(box, u"{}  (not loaded)".format(link.label))
                box.IsChecked = False
                box.IsEnabled = False
            box.ToolTip = link.instance_name
            box.Checked += self.OnLinkChanged
            box.Unchecked += self.OnLinkChanged
            self._link_boxes.append(box)

        checked = set(settings.checked_keys)
        self._category_boxes = []
        for key, label, kind in settings.rules:
            box = CheckBox()
            box.Tag = key
            _set_text(box, label)
            box.IsChecked = key in checked
            box.ToolTip = u"Counted: one instance = one piece" if kind == qr.KIND_COUNT \
                else u"Measured in m, mq or kg according to the unit of the price code"
            box.Checked += self.OnCategoryChanged
            box.Unchecked += self.OnCategoryChanged
            self._category_boxes.append(box)

        self._fill_list(self.lst_links, self._link_boxes, u"")
        self._fill_list(self.lst_categories, self._category_boxes, u"")
        self._reload_worksets()
        self._category_sets = _SetPicker(
            self, self.cbo_sets, self.btn_set_save, self.btn_set_delete, u"category",
            settings.category_sets, settings.category_set_name, self._checked_keys,
            self._apply_category_set, settings.save_category_sets, allow_empty=False)
        # Un set di workset vuoto ("non escludere niente") e' ammesso.
        self._workset_sets = _SetPicker(
            self, self.cbo_workset_sets, self.btn_workset_set_save,
            self.btn_workset_set_delete, u"workset", settings.workset_sets,
            settings.workset_set_name, self._checked_worksets, self._apply_workset_set,
            settings.save_workset_sets, allow_empty=True)
        if not settings.links:
            self.txt_link_count.Text = u"This model has no Revit links."
        self._update_counts()
        self._loading = False

    # ------------------------------------------------------------------ setup

    def _load_xaml(self):
        xaml_path = script.get_bundle_file(XAML_FILE_NAME)
        if not xaml_path or not os.path.exists(xaml_path):
            xaml_path = os.path.join(os.path.dirname(__file__), XAML_FILE_NAME)

        Window.__init__(self)

        stream = FileStream(xaml_path, FileMode.Open)
        try:
            root = XamlReader.Load(stream)
        finally:
            stream.Close()

        self.Content = root.Content
        self.Title = root.Title
        self.Height = root.Height
        self.Width = root.Width
        self.MinHeight = root.MinHeight
        self.MinWidth = root.MinWidth
        self.WindowStartupLocation = root.WindowStartupLocation
        self.ResizeMode = root.ResizeMode
        self.ShowInTaskbar = root.ShowInTaskbar

        for name in ("txt_host", "txt_search_links", "btn_clear_links", "lst_links",
                     "thumb_links", "btn_links_all", "btn_links_none", "txt_link_count",
                     "cbo_workset_sets", "btn_workset_set_save", "btn_workset_set_delete",
                     "txt_search_worksets", "btn_clear_worksets", "lst_worksets",
                     "thumb_worksets", "btn_worksets_all", "btn_worksets_none",
                     "txt_workset_count",
                     "cbo_sets", "btn_set_save", "btn_set_delete",
                     "txt_search_categories", "btn_clear_categories", "lst_categories",
                     "thumb_categories", "btn_categories_all", "btn_categories_none",
                     "txt_category_count", "btn_ok", "btn_cancel"):
            setattr(self, name, root.FindName(name))

        self.txt_search_links.TextChanged += self.OnLinkSearchChanged
        self.btn_clear_links.Click += self.OnClearLinkSearch
        self.thumb_links.DragDelta += self.OnResizeLinks
        self.btn_links_all.Click += self.OnLinksAll
        self.btn_links_none.Click += self.OnLinksNone
        self.txt_search_worksets.TextChanged += self.OnWorksetSearchChanged
        self.btn_clear_worksets.Click += self.OnClearWorksetSearch
        self.thumb_worksets.DragDelta += self.OnResizeWorksets
        self.btn_worksets_all.Click += self.OnWorksetsAll
        self.btn_worksets_none.Click += self.OnWorksetsNone
        self.txt_search_categories.TextChanged += self.OnCategorySearchChanged
        self.btn_clear_categories.Click += self.OnClearCategorySearch
        self.thumb_categories.DragDelta += self.OnResizeCategories
        self.btn_categories_all.Click += self.OnCategoriesAll
        self.btn_categories_none.Click += self.OnCategoriesNone
        self.btn_ok.Click += self.OnOK
        self.btn_cancel.Click += self.OnCancel

    # ------------------------------------------------------------------ liste

    def _fill_list(self, list_box, boxes, search_text):
        """Mostra solo le caselle che contengono il testo cercato; le spunte nascoste
        restano valide."""
        needle = (search_text or u"").strip().lower()
        list_box.Items.Clear()
        for box in boxes:
            if not needle or needle in box.Content.Text.lower():
                list_box.Items.Add(box)

    def _visible(self, list_box):
        return [box for box in list_box.Items]

    def _checked_links(self):
        return [box.Tag for box in self._link_boxes if box.IsEnabled and box.IsChecked]

    def _checked_keys(self):
        return [box.Tag for box in self._category_boxes if box.IsChecked]

    def _checked_worksets(self):
        """Workset esclusi fra quelli in elenco (modello aperto e link spuntati)."""
        return [name for name in self._workset_names if name in self._excluded]

    def _update_counts(self):
        loaded = [box for box in self._link_boxes if box.IsEnabled]
        if self._link_boxes:
            text = u"{} of {} linked models ticked.".format(len(self._checked_links()),
                                                          len(loaded))
            not_loaded = len(self._link_boxes) - len(loaded)
            if not_loaded:
                text += u" {} not loaded.".format(not_loaded)
            self.txt_link_count.Text = text
        if self._workset_names:
            self.txt_workset_count.Text = u"{} of {} worksets excluded.".format(
                len(self._checked_worksets()), len(self._workset_names))
        else:
            self.txt_workset_count.Text = (u"The models to read are not workshared: "
                                           u"there are no worksets to exclude.")
        self.txt_category_count.Text = u"{} of {} categories ticked.".format(
            len(self._checked_keys()), len(self._category_boxes))

    def _set_checks(self, boxes, value):
        self._loading = True
        try:
            for box in boxes:
                if box.IsEnabled:
                    box.IsChecked = value
                    if box in self._workset_boxes:
                        self._mark_workset(box.Tag, value)
        finally:
            self._loading = False
        self._update_counts()

    # ------------------------------------------------------------------ workset

    def _mark_workset(self, name, excluded):
        if excluded:
            self._excluded.add(name)
        else:
            self._excluded.discard(name)

    def _reload_worksets(self):
        """Elenco dei workset del modello aperto e dei link spuntati; le spunte si
        conservano per nome."""
        try:
            self._workset_names = self._list_worksets(self._checked_links())
        except Exception:
            self._workset_names = OrderedDict()
        loading = self._loading
        self._loading = True
        try:
            self._workset_boxes = []
            for name, models in self._workset_names.items():
                box = CheckBox()
                box.Tag = name
                _set_text(box, name)
                box.IsChecked = name in self._excluded
                box.ToolTip = u"Workset of: {}".format(u", ".join(models))
                box.Checked += self.OnWorksetChanged
                box.Unchecked += self.OnWorksetChanged
                self._workset_boxes.append(box)
            self._fill_list(self.lst_worksets, self._workset_boxes,
                            self.txt_search_worksets.Text)
        finally:
            self._loading = loading

    def _apply_workset_set(self, names):
        self._excluded = set(names)
        self._loading = True
        try:
            for box in self._workset_boxes:
                box.IsChecked = box.Tag in self._excluded
        finally:
            self._loading = False
        self._update_counts()

    def _apply_category_set(self, keys):
        keys = set(keys)
        self._loading = True
        try:
            for box in self._category_boxes:
                box.IsChecked = box.Tag in keys
        finally:
            self._loading = False
        self._update_counts()

    # ------------------------------------------------------------------ eventi

    def OnLinkChanged(self, sender, args):
        if self._loading:
            return
        self._reload_worksets()
        self._update_counts()

    def OnWorksetChanged(self, sender, args):
        if self._loading:
            return
        self._mark_workset(sender.Tag, bool(sender.IsChecked))
        self._update_counts()

    def OnCategoryChanged(self, sender, args):
        if not self._loading:
            self._update_counts()

    def OnLinkSearchChanged(self, sender, args):
        self._fill_list(self.lst_links, self._link_boxes, self.txt_search_links.Text)

    def OnClearLinkSearch(self, sender, args):
        self.txt_search_links.Text = u""

    def OnWorksetSearchChanged(self, sender, args):
        self._fill_list(self.lst_worksets, self._workset_boxes, self.txt_search_worksets.Text)

    def OnClearWorksetSearch(self, sender, args):
        self.txt_search_worksets.Text = u""

    def OnCategorySearchChanged(self, sender, args):
        self._fill_list(self.lst_categories, self._category_boxes,
                        self.txt_search_categories.Text)

    def OnClearCategorySearch(self, sender, args):
        self.txt_search_categories.Text = u""

    # Select All / Deselect All valgono sulle righe visibili (filtrate dalla ricerca).
    def OnLinksAll(self, sender, args):
        self._set_checks(self._visible(self.lst_links), True)
        self._reload_worksets()
        self._update_counts()

    def OnLinksNone(self, sender, args):
        self._set_checks(self._visible(self.lst_links), False)
        self._reload_worksets()
        self._update_counts()

    def OnWorksetsAll(self, sender, args):
        self._set_checks(self._visible(self.lst_worksets), True)

    def OnWorksetsNone(self, sender, args):
        self._set_checks(self._visible(self.lst_worksets), False)

    def OnCategoriesAll(self, sender, args):
        self._set_checks(self._visible(self.lst_categories), True)

    def OnCategoriesNone(self, sender, args):
        self._set_checks(self._visible(self.lst_categories), False)

    def _resize(self, list_box, change, low, high):
        height = list_box.Height + change
        if low <= height <= high:
            list_box.Height = height

    def OnResizeLinks(self, sender, args):
        self._resize(self.lst_links, args.VerticalChange, 80, 500)

    def OnResizeWorksets(self, sender, args):
        self._resize(self.lst_worksets, args.VerticalChange, 80, 500)

    def OnResizeCategories(self, sender, args):
        self._resize(self.lst_categories, args.VerticalChange, 120, 600)

    def OnOK(self, sender, args):
        keys = self._checked_keys()
        if not keys:
            MessageBox.Show(u"Tick at least one category to read.", TITLE)
            return
        self.result = ScopeChoice(self._checked_links(), keys, self._category_sets.name,
                                  self._checked_worksets(), self._workset_sets.name)
        self.Close()

    def OnCancel(self, sender, args):
        self.result = None
        self.Close()


def show_scope_dialog(owner, settings):
    """ScopeChoice con link, workset esclusi e categorie da leggere, oppure None se
    annullata. settings: ScopeSettings."""
    form = ScopeForm(settings)
    if owner is not None:
        form.Owner = owner
        form.WindowStartupLocation = WindowStartupLocation.CenterOwner
    form.ShowDialog()
    return form.result
