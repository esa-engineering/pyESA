# -*- coding: utf-8 -*-
__title__ = "Flex\nConnect"
__doc__ = """Version = 1.0.0
Date    = 22.10.2026
Connects a pipe or duct to equipment or an air terminal with a flex pipe or flex duct.
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

# ======================= STORICO REVISIONI =======================
# Versione | Data       | Tipo | Descrizione                                   | PR
# 1.0.0    | 22.10.2026 | FEAT | Prima versione per la main: funzioni delle 0.1-0.11 (vedi sotto) | -
# 0.11.0   | 22.10.2026 | FEAT | Tolto Flex Pattern; impostazioni a sezioni (Pipes, Ducts, General); messaggio a fine comando se inserite riduzioni; Connector Description nell'elenco connettori; icona i anche nella scelta del connettore | -
# 0.10.1   | 22.10.2026 | UI   | Titoli del tooltip How to use in blu come in Clash Setup | -
# 0.10.0   | 22.10.2026 | UI   | Icona i nella finestra delle impostazioni con il tooltip How to use | -
# 0.9.0    | 22.10.2026 | FEAT | Il comando resta attivo dopo ogni collegamento, Esc lo termina | -
# 0.8.0    | 22.10.2026 | FEAT | Scelta del parametro Flex Pattern, separata per tubi e condotti | -
# 0.7.0    | 22.10.2026 | FEAT | Condotti: misura sempre dal connettore, estremita' con transizione, estremita' o fianco dalla distanza assiale | -
# 0.6.0    | 22.10.2026 | FEAT | Scelta nelle impostazioni del raccordo di derivazione (tap) per la presa | -
# 0.5.2    | 22.10.2026 | FIX  | Presa sul fianco verticale a quota centrale, non in diagonale | -
# 0.5.1    | 22.10.2026 | FIX  | Selezione bloccata: AllowReference del filtro ora True | -
# 0.5.0    | 22.10.2026 | FEAT | Condotti flessibili: estremita' e presa sul fianco, ordine di selezione libero | -
# 0.4.0    | 22.10.2026 | FEAT | Tubo corto della riduzione lungo quanto il DN | -
# 0.3.0    | 22.10.2026 | FEAT | Spunta per scegliere il connettore se ce ne sono piu' di uno | -
# 0.2.0    | 22.10.2026 | FEAT | Isolamento copiato dal tubo; finestra solo con SHIFT + CLICK | -
# 0.1.0    | 22.10.2026 | FEAT | Prima versione                                | -
# Testato su Revit: non ancora testato
# =================================================================

from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script, forms

import FlexConnect_Build as fb
import FlexConnect_Duct as fd
import FlexConnect_Form_ui as fui
import FlexConnect_Connector_ui as cui

XAML_FILE_NAME = "FlexConnect_Form.xaml"
CONNECTOR_XAML_FILE_NAME = "FlexConnect_Connector.xaml"
CONFIG_NAME = "ESA_FlexConnect"
DEFAULT_GAP_MM = 150.0

doc = revit.doc
uidoc = revit.uidoc


def load_settings(config, pipe_names, duct_names):
    """(impostazioni salvate, True se valide per questo modello)."""
    flex_type = config.get_option("flex_type", "")
    flex_duct_type = config.get_option("flex_duct_type", "")
    duct_tap = config.get_option("duct_tap", "")
    try:
        gap_mm = float(config.get_option("gap_mm", DEFAULT_GAP_MM))
    except (TypeError, ValueError):
        gap_mm = DEFAULT_GAP_MM
    ask_connector = str(config.get_option("ask_connector", "0")).lower() in ("1", "true")
    settings = {"flex_type": flex_type, "flex_duct_type": flex_duct_type, "duct_tap": duct_tap,
                "gap_mm": gap_mm, "ask_connector": ask_connector}
    valid = ((not pipe_names or flex_type in pipe_names) and
             (not duct_names or flex_duct_type in duct_names))
    return settings, valid


def save_settings(config, settings):
    try:
        config.set_option("flex_type", settings["flex_type"])
        config.set_option("flex_duct_type", settings["flex_duct_type"])
        config.set_option("duct_tap", settings["duct_tap"])
        config.set_option("gap_mm", settings["gap_mm"])
        config.set_option("ask_connector", "1" if settings["ask_connector"] else "0")
        script.save_config()
    except Exception:
        pass


def get_settings(flex_types, tap_names):
    pipe_names = [n for n, _ in flex_types[fb.PIPE]]
    duct_names = [n for n, _ in flex_types[fb.DUCT]]
    config = script.get_config(CONFIG_NAME)
    settings, valid = load_settings(config, pipe_names, duct_names)
    # eccezione alla regola comune: la finestra si apre solo con SHIFT + CLICK (o al primo uso)
    if valid and not __shiftclick__:  # noqa: F821
        return settings
    xaml_path = script.get_bundle_file(XAML_FILE_NAME)
    settings = fui.FlexConnectWindow(xaml_path, pipe_names, duct_names, tap_names, settings).show()
    if settings is None:
        script.exit()
    save_settings(config, settings)
    return settings


def choose_connector(candidates):
    """Finestra di scelta del connettore: ritorna il candidato scelto, o None se annullata."""
    labels = [fb.describe_candidate(c) for c in candidates]
    xaml_path = script.get_bundle_file(CONNECTOR_XAML_FILE_NAME)
    index = cui.ConnectorWindow(xaml_path, labels).show()
    return None if index is None else candidates[index]


def connect_once(trunk, point, equipment, settings, type_ids, tap_ids):
    """Un collegamento: True se e' stata inserita una riduzione.
    Solleva FlexConnectError / FlexConnectCancelled / OperationCanceledException."""
    kind = fb.kind_of(trunk)
    name = settings["flex_type"] if kind is fb.PIPE else settings["flex_duct_type"]
    flex_type_id = type_ids[kind].get(name)
    if flex_type_id is None:
        raise fb.FlexConnectError("The project has no flex {} type.".format(kind.name))
    chooser = choose_connector if settings["ask_connector"] else None

    if kind is fb.PIPE:
        pipe_conn, eq_conn = fb.find_connector_pair(trunk, equipment, chooser)
        with revit.Transaction("FlexConnect"):
            return fb.connect_end(doc, trunk, pipe_conn, eq_conn, flex_type_id, settings["gap_mm"])

    if point is None:
        point = fb.pick_point_on(uidoc, trunk)
    plan = fd.plan_connection(trunk, point, equipment, chooser)
    with revit.Transaction("FlexConnect"):
        return fd.build(doc, trunk, plan, flex_type_id, settings["gap_mm"], tap_ids.get(settings["duct_tap"]))


def try_connect(trunk, point, equipment, settings, type_ids, tap_ids, session):
    """Un collegamento con gli errori mostrati in un avviso. False se l'utente ha annullato con Esc.
    session: dict dell'attivazione, "reducers" conta le riduzioni inserite."""
    try:
        if connect_once(trunk, point, equipment, settings, type_ids, tap_ids):
            session["reducers"] += 1
    except OperationCanceledException:
        return False
    except fb.FlexConnectCancelled:
        pass    # finestra di scelta del connettore chiusa: si salta questo collegamento
    except fb.FlexConnectError as err:
        forms.alert(str(err), title="FlexConnect")
    except Exception as err:
        forms.alert("The connection could not be created.\n{}".format(err), title="FlexConnect")
    return True


def main():
    flex_types = {fb.PIPE: fb.collect_flex_types(doc, fb.PIPE),
                  fb.DUCT: fb.collect_flex_types(doc, fb.DUCT)}
    if not flex_types[fb.PIPE] and not flex_types[fb.DUCT]:
        forms.alert("The project has no flex pipe or flex duct type.", exitscript=True)
    type_ids = dict((kind, dict(types)) for kind, types in flex_types.items())
    tap_symbols = fd.collect_tap_symbols(doc)
    tap_ids = dict(tap_symbols)
    settings = get_settings(flex_types, [n for n, _ in tap_symbols])

    # il comando resta attivo: dopo ogni collegamento riparte la selezione, Esc lo termina.
    # Sui condotti si continua con lo stesso condotto (solo terminale e punto); Esc termina il comando.
    session = {"reducers": 0}
    first_round = True
    while True:
        try:
            trunk, point, equipment = fb.pick_elements(uidoc, use_preselection=first_round)
        except OperationCanceledException:
            break
        first_round = False
        keep_going = try_connect(trunk, point, equipment, settings, type_ids, tap_ids, session)
        if fb.is_duct(trunk):
            while keep_going:
                next_connection = fb.pick_next_duct_connection(uidoc, trunk)
                if next_connection is None:
                    break
                point, equipment = next_connection
                keep_going = try_connect(trunk, point, equipment, settings, type_ids, tap_ids, session)
            break
        if not keep_going:
            break

    # un solo messaggio per attivazione del comando, a fine comando
    if session["reducers"]:
        forms.alert("A reducer was added on {} connection(s). Check the fittings in the model.".format(
            session["reducers"]), title="FlexConnect")


main()
