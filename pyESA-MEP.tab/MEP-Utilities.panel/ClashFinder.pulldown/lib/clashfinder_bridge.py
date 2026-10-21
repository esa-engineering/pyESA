# -*- coding: utf-8 -*-
"""
clashfinder_bridge.py - chiamate all'API di Revit da finestre non modali.

Le finestre di Clash Finder restano aperte mentre l'utente lavora in Revit (per
selezionare elementi e poi caricarli, o per navigare fra le interferenze). Da un
evento WPF di una finestra non modale non si e' in un contesto API valido: ogni
lettura o selezione passa da un ExternalEvent, che Revit esegue appena e' libero.

    bridge = RevitBridge("Clash Finder")     # da creare in contesto API (lo script)
    bridge.run(lambda uiapp: ...)            # asincrono: torna subito
    bridge.dispose()                         # alla chiusura della finestra

Le funzioni accodate ricevono UIApplication. Un errore viene mostrato in una
MessageBox con il traceback, senza interrompere le chiamate successive.
"""

import traceback
from collections import deque

import clr
clr.AddReference('PresentationFramework')
from System.Windows import MessageBox, MessageBoxImage, MessageBoxButton

from pyrevit import UI


class _QueueHandler(UI.IExternalEventHandler):

    def __init__(self, title):
        self.title = title
        self.queue = deque()

    def Execute(self, uiapp):
        while self.queue:
            func = self.queue.popleft()
            try:
                func(uiapp)
            except SystemExit:
                pass
            except Exception:
                MessageBox.Show(u"Unexpected error:\n\n{}".format(traceback.format_exc()),
                                self.title, MessageBoxButton.OK, MessageBoxImage.Error)

    def GetName(self):
        return self.title


class RevitBridge(object):

    def __init__(self, title):
        self._handler = _QueueHandler(title)
        self._event = UI.ExternalEvent.Create(self._handler)

    def run(self, func):
        if self._event is None:
            return
        self._handler.queue.append(func)
        self._event.Raise()

    def dispose(self):
        if self._event is not None:
            try:
                self._event.Dispose()
            except Exception:
                pass
            self._event = None
        self._handler.queue.clear()
