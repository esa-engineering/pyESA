# -*- coding: utf-8 -*-
"""Guida a slide di Clash Finder, condivisa dai comandi del pulldown.

Le slide sono immagini PNG (1920 x 1080) in lib/ClashFinder_guide, in ordine di nome:
ClashFinder_guide_01.png, _02.png... La finestra mostra una slide alla volta, con
Previous / Next e le frecce della tastiera. Per aggiornare la guida si sostituiscono
i PNG (stessi nomi, numerazione continua): nessun codice da toccare.

Modulo senza stato a livello di modulo: vale con Rocket mode.
"""
import os
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import Uri, UriKind
from System.IO import FileStream, FileMode
from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage
from System.Windows.Input import Key
from System.Windows.Markup import XamlReader
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption

XAML_FILE = 'ClashFinder_guide.xaml'
IMAGE_DIR = 'ClashFinder_guide'
IMAGE_PREFIX = 'ClashFinder_guide_'
TITLE = u"Clash Finder Guide"


def _slide_paths(folder):
    names = [n for n in os.listdir(folder)
             if n.startswith(IMAGE_PREFIX) and n.lower().endswith(".png")]
    return [os.path.join(folder, n) for n in sorted(names)]


class GuideWindow(Window):
    """Finestra modale con le slide della guida."""

    def __init__(self, owner, slides):
        Window.__init__(self)
        xaml_path = os.path.join(os.path.dirname(__file__), XAML_FILE)
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
        if owner is not None:
            self.Owner = owner

        self.slides = slides
        self.index = 0
        self.img_slide = root.FindName("img_slide")
        self.btn_prev = root.FindName("btn_prev")
        self.btn_next = root.FindName("btn_next")
        self.txt_page = root.FindName("txt_page")
        root.FindName("btn_close").Click += lambda s, e: self.Close()
        self.btn_prev.Click += lambda s, e: self._go(-1)
        self.btn_next.Click += lambda s, e: self._go(1)
        self.KeyDown += self._on_key
        self._show()

    def _go(self, step):
        index = self.index + step
        if 0 <= index < len(self.slides):
            self.index = index
            self._show()

    def _show(self):
        image = BitmapImage()
        image.BeginInit()
        image.CacheOption = BitmapCacheOption.OnLoad   # non blocca il file
        image.UriSource = Uri(self.slides[self.index], UriKind.Absolute)
        image.EndInit()
        self.img_slide.Source = image
        self.txt_page.Text = u"{} / {}".format(self.index + 1, len(self.slides))
        self.btn_prev.IsEnabled = self.index > 0
        self.btn_next.IsEnabled = self.index < len(self.slides) - 1

    def _on_key(self, sender, args):
        if args.Key in (Key.Right, Key.PageDown, Key.Space):
            self._go(1)
        elif args.Key in (Key.Left, Key.PageUp):
            self._go(-1)
        elif args.Key == Key.Home:
            self.index = 0
            self._show()
        elif args.Key == Key.End:
            self.index = len(self.slides) - 1
            self._show()
        elif args.Key == Key.Escape:
            self.Close()
        else:
            return
        args.Handled = True


def show_guide(owner=None):
    """Apre la guida sopra la finestra owner. Da chiamare dal thread della UI."""
    folder = os.path.join(os.path.dirname(__file__), IMAGE_DIR)
    slides = _slide_paths(folder) if os.path.isdir(folder) else []
    if not slides:
        MessageBox.Show(u"The guide slides were not found in:\n{}".format(folder), TITLE,
                        MessageBoxButton.OK, MessageBoxImage.Warning)
        return
    GuideWindow(owner, slides).ShowDialog()
