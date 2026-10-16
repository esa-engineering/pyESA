# -*- coding: utf-8 -*-
__title__ = "Tab\nVisibility"
__doc__ = """Version = 1.3
Date    = 15.10.2026
________________________________________________________________
Description:
Shows or hides the add-in tabs of the Revit ribbon with a checkbox for
each tab (checked = shown). The list shows the tabs as they are now.

The default Revit tabs (Architecture, Structure, Insert, View...) are
not listed: they are managed in File > Options > User Interface >
Configure.

With "Keep these tabs hidden when Revit restarts" ticked, the hidden
tabs are saved and hidden again at every Revit start. Tabs are saved
by name: a hidden tab of an add-in that is not loaded now stays saved.

The pyESA tab is never listed: hiding it would leave this tool out
of reach.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import clr
clr.AddReference("AdWindows")
clr.AddReference("RevitAPIUI")
from Autodesk.Windows import ComponentManager

from pyrevit import script

import Tab_ui

XAML_FILE_NAME = "Tab_form.xaml"
OWN_TAB = "pyESA"  # mai elencata: nasconderla lascerebbe senza accesso a questo strumento
# Stessa sezione e stesse chiavi lette da startup.py nella radice dell'estensione
CONFIG_NAME = "ESA_TabVisibility"
SEPARATOR = "|"

output = script.get_output()


# Id dei tab nativi di Revit (non localizzati): non si gestiscono da qui
BUILTIN_TAB_NAMES = set((
    "architecture", "structure", "steel", "precast", "systems", "insert", "annotate",
    "analyze", "massing", "massingsite", "collaborate", "view", "manage", "addins",
    "modify", "create", "extensions"))


def _norm(text):
    return "".join(c for c in str(text).lower() if c.isalnum())


def tab_key(tab):
    return str(tab.Id)


def is_builtin(tab):
    try:
        if tab.IsContextualTab:
            return True
    except Exception:
        pass
    # Id, Name e Title: l'Id nativo non coincide sempre con il titolo (anche con prefisso ID_)
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


def tab_title(tab):
    try:
        return tab.Title or ""
    except Exception:
        return ""


def to_bool(value):
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ("true", "1", "yes")


def load_saved(config):
    """(titoli nascosti salvati, ricorda al riavvio)."""
    try:
        text = config.get_option("hidden_tabs", "") or ""
        remember = to_bool(config.get_option("remember", True))
    except Exception:
        return set(), True
    return set(t for t in text.split(SEPARATOR) if t), remember


def save(config, hidden_titles, remember):
    try:
        config.hidden_tabs = SEPARATOR.join(sorted(hidden_titles)) if remember else ""
        config.remember = remember
        script.save_config()
        return True
    except Exception as ex:
        output.print_md("**The settings could not be saved:** {}".format(ex))
        return False


def main():
    try:
        tabs = list(ComponentManager.Ribbon.Tabs)
    except Exception as ex:
        output.print_md("**The ribbon could not be read:** {}".format(ex))
        return
    builtin_titles = set(tab_title(t) for t in tabs if is_builtin(t))
    listed = [t for t in tabs
              if tab_title(t) and tab_title(t) != OWN_TAB and not is_builtin(t)]
    items = [(tab_key(t), tab_title(t), bool(t.IsVisible)) for t in listed]
    if not items:
        output.print_md("No add-in tab can be listed. The default Revit tabs are managed in "
                        "**File > Options > User Interface > Configure**.")
        return

    config = script.get_config(CONFIG_NAME)
    saved_hidden, remember = load_saved(config)

    result = Tab_ui.TabWindow(script.get_bundle_file(XAML_FILE_NAME), items, remember).show()
    if result is None:
        script.exit()
    visible_by_key, remember = result

    changed, failed = 0, []
    for tab in listed:
        wanted = visible_by_key.get(tab_key(tab))
        if wanted is None or wanted == bool(tab.IsVisible):
            continue
        try:
            tab.IsVisible = wanted
            changed += 1
        except Exception as ex:
            failed.append("{}: {}".format(tab_title(tab), ex))

    # i tab elencati ora prendono la scelta della finestra; quelli salvati ma assenti
    # in questa sessione (add-in non caricato) restano salvati
    listed_titles = set(tab_title(t) for t in listed)
    hidden_now = set(tab_title(t) for t in listed if visible_by_key.get(tab_key(t)) is False)
    hidden_titles = (saved_hidden - listed_titles) | hidden_now
    hidden_titles.discard(OWN_TAB)
    hidden_titles -= builtin_titles  # salvataggi di versioni precedenti
    saved = save(config, hidden_titles, remember)

    if failed:
        output.print_md("**Not changed:** " + "; ".join(failed))
    if changed == 0 and not failed and saved and hidden_titles == saved_hidden:
        output.print_md("No change.")


main()
