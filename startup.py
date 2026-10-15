# -*- coding: utf-8 -*-
"""Avvio dell'estensione: nasconde di nuovo i tab salvati da pyESA > Settings > Tab Visibility.

pyRevit esegue questo file a ogni avvio di Revit e a ogni Reload. I tab vengono nascosti al
primo Idling, quando tutti gli add-in hanno gia' creato i loro tab. Se l'utente non ha salvato
nulla non si iscrive a nessun evento. Nessun output: un errore qui non deve disturbare l'avvio.
Sezione e chiavi di configurazione sono le stesse di TabVisibility.pushbutton/script.py.
"""
import clr
clr.AddReference("AdWindows")
from Autodesk.Windows import ComponentManager
from Autodesk.Revit.UI.Events import IdlingEventArgs

from pyrevit import HOST_APP, framework, script

CONFIG_NAME = "ESA_TabVisibility"
SEPARATOR = "|"
OWN_TAB = "pyESA"  # mai nascosta: conterrebbe lo strumento per farla ricomparire


def _saved_hidden_titles():
    try:
        text = script.get_config(CONFIG_NAME).get_option("hidden_tabs", "") or ""
    except Exception:
        return set()
    titles = set(t for t in text.split(SEPARATOR) if t)
    titles.discard(OWN_TAB)
    return titles


BUILTIN_TAB_NAMES = set((
    "architecture", "structure", "steel", "precast", "systems", "insert", "annotate",
    "analyze", "massing", "massingsite", "collaborate", "view", "manage", "addins",
    "modify", "create", "extensions"))


def _norm(text):
    return "".join(c for c in str(text).lower() if c.isalnum())


def _is_builtin(tab):
    # stessa logica di TabVisibility.pushbutton/script.py
    for attr in ("Id", "Name", "Title"):
        try:
            value = _norm(getattr(tab, attr))
        except Exception:
            continue
        if value in BUILTIN_TAB_NAMES or value.startswith("modify"):
            return True
        if value.startswith("id") and value[2:] in BUILTIN_TAB_NAMES:
            return True
    return False


def _hide(titles):
    for tab in list(ComponentManager.Ribbon.Tabs):
        try:
            # i tab nativi di Revit non si toccano (salvataggi di versioni precedenti)
            if _is_builtin(tab):
                continue
            if tab.Title in titles and tab.IsVisible:
                tab.IsVisible = False
        except Exception:
            pass


def _on_idling(sender, args):
    # una volta sola: ci si disiscrive subito
    try:
        sender.Idling -= _HANDLER
    except Exception:
        pass
    try:
        _hide(_HIDDEN)
    except Exception:
        pass


_HIDDEN = _saved_hidden_titles()
_HANDLER = None
if _HIDDEN:
    _HANDLER = framework.EventHandler[IdlingEventArgs](_on_idling)
    try:
        HOST_APP.uiapp.Idling += _HANDLER
    except Exception:
        # evento non disponibile: almeno i tab gia' presenti
        try:
            _hide(_HIDDEN)
        except Exception:
            pass
