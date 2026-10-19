# -*- coding: utf-8 -*-
"""Openings: costanti, impostazioni e stato del comando in corso, comuni a tutti i moduli di Openings.

I moduli di openings_modules si importano a catena, ciascuno con `from <precedente> import *`:
openings_core <- openings_geometry <- openings_pick <- openings_check <- openings_report <- openings_place.
Gli script dei pulsanti (Place, NewReport, LastReport, Settings) aggiungono questa cartella a sys.path.

Rocket mode: i moduli restano in sys.modules tra un comando e l'altro. Per questo qui non c'e' nessun documento,
output o impostazione a livello di modulo: stanno in `ctx`, riempito da start(MODE) all'inizio di ogni comando. La
finestra del report, non modale, ne salva una copia e la ripristina prima di ogni azione (ReportController).
Dopo aver modificato un modulo serve pyRevit Reload."""
import json
import math
import os
import os.path as op
from datetime import datetime

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter, PickBoxStyle
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Collections.Generic import List

from pyrevit import revit, script, DB, forms

MODULE_DIR = op.dirname(op.abspath(__file__))  # anche le XAML stanno qui
TITLES = {"place": "Openings Place&Check", "new": "Openings New Report", "last": "Openings Last Report"}


class Context(object):
    """Stato del comando in corso. Un solo oggetto (ctx), mai riassegnato: start() e restore() ne cambiano i campi,
    cosi' chi l'ha importato con `from openings_core import *` vede sempre quelli attuali."""
    FIELDS = ("mode", "title", "doc", "uidoc", "output", "settings", "status_param", "id_param", "cat_param",
              "min_cross_sin", "neighbour_cache", "sources", "mep_cache", "skipped_links")

    def __init__(self):
        for name in self.FIELDS:
            setattr(self, name, None)

    def snapshot(self):
        return dict((name, getattr(self, name)) for name in self.FIELDS)

    def restore(self, saved):
        for name in self.FIELDS:
            setattr(self, name, saved[name])


ctx = Context()


def start(mode):
    """Da chiamare all'inizio di ogni comando, prima di ogni altra funzione dei moduli: documento, output e
    impostazioni di questo comando (mode: "place" | "new" | "last")."""
    settings = load_settings()
    ctx.restore({"mode": mode, "title": TITLES[mode], "doc": revit.doc, "uidoc": revit.uidoc,
                 "output": script.get_output(), "settings": settings, "status_param": settings["status_param"],
                 "id_param": settings["related_id_param"], "cat_param": settings["related_cat_param"],
                 # sotto quest'angolo con il piano di muro o solaio l'elemento MEP corre dentro (scarichi a
                 # pavimento, montanti nel muro): nessun foro
                 "min_cross_sin": math.sin(math.radians(min(max(settings["min_angle_deg"], 0.0), 89.0))),
                 # (id del link o 0, id dell'elemento) -> solidi in coordinate host
                 "neighbour_cache": {},
                 # modelli letti (host e link caricati) e riferimenti MEP gia' risolti: vedi reset_models()
                 "sources": None, "mep_cache": {},
                 # chiavi dei link da non leggere (scelta di Place); New Report e Last Report li leggono tutti
                 "skipped_links": frozenset()})


def reset_models():
    """Rilegge i link e i riferimenti agli elementi MEP alla prossima richiesta: a ogni controllo, perche' mentre la
    finestra del report e' aperta un link puo' essere ricaricato (gli elementi letti prima non valgono piu')."""
    ctx.sources = None
    ctx.mep_cache = {}


# =================================================================================================================
# Nomi, parametri, impostazioni predefinite: una sola copia, letta anche da Settings
# =================================================================================================================
CONFIG_NAME = "ESA_Openings"
DEFAULTS = {"fam_wall_rect": "e_GM.PH.VD_Void Demand_Wall_Rect", "fam_wall_round": "e_GM.PH.VD_Void Demand_Wall_Circ",
            "fam_floor_rect": "e_GM.PH.VD_Void Demand_Floor_Rect",
            "fam_floor_round": "e_GM.PH.VD_Void Demand_Floor_Circ",
            "status_param": "e_GEN_Void_Status", "related_id_param": "e_GEN_RelatedElement_ID",
            "related_cat_param": "e_GEN_RelatedElement_Cat",
            "clearance_mm": 25.0, "rounding_mm": 50.0, "merge_mm": 100.0, "extra_depth_mm": 0.0, "layer_gap_mm": 200.0,
            "min_angle_deg": 15.0,
            "separate": False, "misaligned": True}
NUMBER_KEYS = ("clearance_mm", "rounding_mm", "merge_mm", "extra_depth_mm", "layer_gap_mm", "min_angle_deg")
BOOL_KEYS = ("separate", "misaligned")
TEXT_KEYS = ("status_param", "related_id_param", "related_cat_param")  # vuoti: tornano al predefinito
FAMILY_KEYS = {("wall", "rect"): "fam_wall_rect", ("wall", "round"): "fam_wall_round",
               ("floor", "rect"): "fam_floor_rect", ("floor", "round"): "fam_floor_round"}
# famiglie Generic Model NON ospitate, origine al centro, parametri di istanza (condivisi), lunghezza.
# Muro rettangolare: X larghezza, Y profondita', Z altezza. Muro circolare: cilindro lungo X, Width = lunghezza.
# Solaio rettangolare: X e Y in pianta, Z profondita'. Solaio circolare: cilindro lungo Z, Height = lunghezza
PARAM_WIDTH = "e_DIM_Width_1"
PARAM_HEIGHT = "e_DIM_Height_1"
PARAM_DEPTH = "e_DIM_Depth_1"
PARAM_DIAMETER = "e_DIM_Diameter_1"
# lunghezza del foro da parte a parte (lungo l'asse di attraversamento), per famiglia
LENGTH_PARAMS = {("wall", "rect"): PARAM_DEPTH, ("wall", "round"): PARAM_WIDTH,
                 ("floor", "rect"): PARAM_DEPTH, ("floor", "round"): PARAM_HEIGHT}
# famiglie rettangolari: il solido poggia sul livello di riferimento (punto di inserimento = centro della faccia
# inferiore), non e' centrato in Z. Parametro della misura lungo Z: il centro sta a meta' sopra l'inserimento
BASE_PARAMS = {("wall", "rect"): PARAM_HEIGHT, ("floor", "rect"): PARAM_DEPTH}
# stati scritti in e_GEN_Void_Status, dal migliore al peggiore
VALID, TO_REFRESH, NOT_ASSOCIATED = "Valid", "To Be Refreshed", "Not Associated"
# =================================================================================================================

# stati interni del controllo; nel modello si scrive solo lo stato corrispondente di WRITTEN.
# REGROUP: foro da dividere (MEC/ELE separati) o da unire a un vicino secondo Settings, elementi ancora dentro
# PARTIAL: almeno un elemento ancora dentro o in parte fuori, altri tolti o usciti: resta To Be Refreshed
NEW, OK, MISALIGNED, REGROUP, PARTIAL, CHANGED, OBSOLETE = "New", "OK", "Misaligned", "Regroup", "Partial", "Changed", \
    "Obsolete"
WRITTEN = {NEW: VALID, OK: VALID, MISALIGNED: TO_REFRESH, REGROUP: TO_REFRESH, PARTIAL: TO_REFRESH,
           CHANGED: NOT_ASSOCIATED, OBSOLETE: NOT_ASSOCIATED}
WRITTEN_RANK = {VALID: 0, TO_REFRESH: 1, NOT_ASSOCIATED: 2}

MOVE_TOL_MM = 5.0
SIZE_TOL_MM = 1.0
MM_PER_FT = 304.8
VERTICAL_SIN = math.sin(math.radians(45.0))   # oltre: elemento MEP verticale, foro a solaio
ROUND_COS = 0.995            # tondo solo se l'attraversamento e' (quasi) perpendicolare
Z = DB.XYZ.BasisZ
MEP_CLASSES = (Pipe, Duct, CableTray, Conduit)
# nomi fissi in inglese per e_GEN_RelatedElement_Cat (Category.Name dipende dalla lingua di Revit)
MEP_LABELS = ((Pipe, "Pipes"), (Duct, "Ducts"), (CableTray, "Cable Trays"), (Conduit, "Conduits"))
MEP_SINGULAR = ((Pipe, "Pipe"), (Duct, "Duct"), (CableTray, "Cable Tray"), (Conduit, "Conduit"))
# stato di ogni elemento elencato in un foro (report del Check)
ELEMENT_INSIDE = "inside"
ELEMENT_PARTLY_OUT = "partly outside"
ELEMENT_MOVED_OUT = "moved out"
ELEMENT_NOT_CROSSING = "no longer crosses"
ELEMENT_UNREADABLE = "section not readable"
ELEMENT_DELETED = "deleted"
ELEMENT_ADDED = "crossing it, now associated"
ELEMENT_ADDED_OUT = "crossing it partly outside, now associated"
ELEMENT_NO_OPENING = "no opening"  # attraversamento senza foro (Missing openings)
ELEMENT_LINK_NOT_LOADED = "link not loaded"  # elemento di un link non caricato: Id tenuto, stato del foro invariato
# attraversa un foro orfano (nessun elemento associato e' piu' dentro): non si associa da solo, lo fa Recreate
ELEMENT_FOUND_INSIDE = "inside, not associated (Recreate associates it)"
MISSING = "Missing openings"  # gruppo del report: attraversamenti senza foro
MEP_CATEGORIES = (DB.BuiltInCategory.OST_PipeCurves, DB.BuiltInCategory.OST_DuctCurves,
                  DB.BuiltInCategory.OST_CableTray, DB.BuiltInCategory.OST_Conduit)
# muri e solai: foro a muro / a solaio; strutture (travi, pilastri, fondazioni): secondo la direzione dell'elemento MEP
HOST_KINDS = {int(DB.BuiltInCategory.OST_Walls): "wall", int(DB.BuiltInCategory.OST_Floors): "floor",
              int(DB.BuiltInCategory.OST_StructuralFraming): "auto", int(DB.BuiltInCategory.OST_StructuralColumns): "auto",
              int(DB.BuiltInCategory.OST_StructuralFoundation): "auto"}
STATE_RANK = {OK: 0, NEW: 1, MISALIGNED: 2, REGROUP: 3, PARTIAL: 4, CHANGED: 5, OBSOLETE: 6}

# finestre di Place&Check (OpeningsPlace_ui.py, importato solo da place_main)
PICK_XAML = "OpeningsPlace_pick.xaml"
CONFIRM_XAML = "OpeningsPlace_confirm.xaml"
RESULT_XAML = "OpeningsPlace_result.xaml"
# finestre dei report (OpeningsCheck_ui.py e OpeningsReport_xlsx.py); moduli e XAML tutti in openings_modules
SCOPE_XAML = "OpeningsCheck_scope.xaml"
SUMMARY_XAML = "OpeningsCheck_summary.xaml"
REPORT_XAML = "OpeningsCheck_report.xaml"
TEXT_XAML = "OpeningsCheck_text.xaml"


def load_settings(config=None):
    """Impostazioni dell'utente, completate con i predefiniti. config: sezione gia' aperta (Settings la riusa per
    salvare), altrimenti si apre ESA_Openings."""
    settings = dict(DEFAULTS)
    try:
        if config is None:
            config = script.get_config(CONFIG_NAME)
        for key in DEFAULTS:
            value = config.get_option(key, DEFAULTS[key])
            if key in NUMBER_KEYS:
                value = float(value)
            elif key in BOOL_KEYS and not isinstance(value, bool):
                value = str(value).strip().lower() in ("true", "1", "yes")
            elif key in TEXT_KEYS:
                value = str(value or "").strip() or DEFAULTS[key]
            settings[key] = value
    except Exception:
        pass
    return settings


def id_value(eid):
    return eid.Value if hasattr(eid, "Value") else eid.IntegerValue  # Revit 2026+ / <= 2025


def mm(value_ft):
    return value_ft * MM_PER_FT


def round_up(value_mm):
    step = ctx.settings["rounding_mm"]
    if step <= 0:
        return value_mm
    return math.ceil(value_mm / step - 1e-9) * step


def level_elevation(level):
    # ProjectElevation e' coerente con la geometria (coordinate interne), Elevation puo' riferirsi al Survey Point
    value = getattr(level, "ProjectElevation", None)
    return level.Elevation if value is None else value


# ----------------------------------------------------------------------------------------------------------------
# Parametri e_GEN: stato, Id e categorie degli elementi collegati (foro <-> elementi MEP)
# ----------------------------------------------------------------------------------------------------------------
def get_text(element, name):
    """Valore di testo del parametro, "" se vuoto, None se il parametro non c'e'."""
    param = element.LookupParameter(name)
    if param is None:
        return None
    return param.AsString() or ""


def set_text(element, name, value, missing):
    """Scrive solo se il valore cambia (modelli condivisi: non si prende in prestito un elemento per niente)."""
    param = element.LookupParameter(name)
    if param is None or param.IsReadOnly:
        missing.add("{} on {}".format(name, element.Category.Name if element.Category else "element"))
        return
    if (param.AsString() or "") != (value or ""):
        param.Set(value)


def element_of(mep):
    """Elemento Revit di un MepRef (openings_geometry), o l'elemento stesso."""
    return getattr(mep, "element", mep)


def ref_text(model, number):
    """Riferimento a un elemento MEP scritto sul foro: "123456" nel modello host (model vuoto), "PRJ_MEC:123456" in
    un link. Gli Id di documenti diversi si sovrappongono: per i link serve il modello."""
    return "{}:{}".format(model, number) if model else str(number)


def split_ref(text):
    """(modello, numero) di un riferimento; modello "" per l'host. I nomi dei file non contengono ":"."""
    model, _, number = text.rpartition(":")
    return model, number


def parse_refs(value):
    """Riferimenti (ref_text) scritti nel parametro Id di un foro. Formato: elementi del modello aperto senza prefisso,
    poi un gruppo per link: "111,222; PRJ_MEC: 333,444; PRJ_ELE: 555". Si leggono anche gli Id nudi dei modelli
    unici e il formato di prova "PRJ_MEC:333,PRJ_ELE:555": in un gruppo, un "<modello>:" vale per quell'Id e per i
    seguenti; ";" torna al modello aperto."""
    refs = []
    for group in (value or "").split(";"):
        model = ""
        for token in group.split(","):
            token = token.strip()
            if ":" in token:
                model, _, token = token.rpartition(":")
                model, token = model.strip(), token.strip()
            if token:
                text = ref_text(model, token)
                if text not in refs:
                    refs.append(text)
    return refs


def format_refs(refs):
    """Valore del parametro Id di un foro (inverso di parse_refs): Id del modello aperto, poi "; <modello>: Id,Id"
    per ogni link, in ordine di nome. Leggibile in un abaco; gli Id di un modello si incollano in Select by ID."""
    groups = {}
    for text in refs:
        model, number = split_ref(text)
        if number not in groups.setdefault(model, []):
            groups[model].append(number)
    parts = [",".join(groups.pop(""))] if "" in groups else []
    parts += ["{}: {}".format(model, ",".join(groups[model])) for model in sorted(groups)]
    return "; ".join(parts)


def discipline(mep):
    return "MEC" if isinstance(element_of(mep), (Pipe, Duct)) else "ELE"


def disciplines_text(meps):
    found = set(discipline(m) for m in meps)
    return "+".join(d for d in ("MEC", "ELE") if d in found)


def categories_text(meps):
    """Categorie degli elementi MEP del foro, nomi fissi in inglese: "Pipes, Ducts"."""
    return ", ".join(label for cls, label in MEP_LABELS if any(isinstance(element_of(m), cls) for m in meps))


def worst(statuses):
    """Stato scritto peggiore (Valid < To Be Refreshed < Not Associated); testi sconosciuti contano come Valid."""
    return max(statuses or [VALID], key=lambda s: WRITTEN_RANK.get(s, 0))


def code_of(opening):
    """Codice del foro scritto sugli elementi MEP: il suo Id."""
    return str(id_value(opening.Id))


def write_opening(instance, meps, state, missing, kept=()):
    """Stato, riferimenti e categorie degli elementi MEP del foro (MepRef). kept: riferimenti da non perdere anche
    se non leggibili ora (elementi di link non caricati); le loro categorie restano quelle gia' scritte."""
    set_text(instance, ctx.status_param, WRITTEN[state], missing)
    set_text(instance, ctx.id_param, format_refs([m.text for m in meps] + list(kept)), missing)
    if not kept:
        set_text(instance, ctx.cat_param, categories_text(meps), missing)


def split_list(text):
    return [part.strip() for part in (text or "").split(",") if part.strip()]


def element_by_id(text):
    try:
        return ctx.doc.GetElement(DB.ElementId(int(text)))
    except Exception:
        return None


def size_text(data):
    if data["shape"] == "round":
        return "D{:.0f} x {:.0f}".format(data["diameter"], data["depth"])
    return "{:.0f} x {:.0f} x {:.0f}".format(data["width"], data["height"], data["depth"])


def settings_line():
    return ("clearance {clearance_mm:g} mm per side, rounding {rounding_mm:g} mm, merge within {merge_mm:g} mm, "
            "extra depth {extra_depth_mm:g} mm per side, layers joined across gaps up to {layer_gap_mm:g} mm, "
            "elements under {min_angle_deg:g} degrees to the wall/floor skipped, MEC/ELE {sep}").format(
        sep="separate" if ctx.settings["separate"] else "together", **ctx.settings)


def bundle_path(name):
    """XAML di Openings: stanno in openings_modules con i moduli."""
    return op.join(MODULE_DIR, name)


def last_pick():
    try:
        return script.get_config(CONFIG_NAME).get_option("place_pick", "hosts")
    except Exception:
        return "hosts"


def remember_pick(pick):
    try:
        config = script.get_config(CONFIG_NAME)
        config.place_pick = pick
        script.save_config()
    except Exception:
        pass


def family_line():
    settings = ctx.settings
    return "Families: wall {} / {}, floor {} / {}.".format(settings["fam_wall_rect"], settings["fam_wall_round"],
                                                           settings["fam_floor_rect"], settings["fam_floor_round"])


def mep_name(mep):
    for cls, label in MEP_SINGULAR:
        if isinstance(element_of(mep), cls):
            return label
    return "Element"


def last_scope_kind(default):
    try:
        return script.get_config(CONFIG_NAME).get_option("check_scope", default)
    except Exception:
        return default


def remember_scope_kind(kind):
    try:
        config = script.get_config(CONFIG_NAME)
        config.check_scope = kind
        script.save_config()
    except Exception:
        pass


def unique_meps(items):
    meps = []
    for item in items:
        if all(m != item["mep"] for m in meps):
            meps.append(item["mep"])
    return meps


def group_discipline(items):
    return disciplines_text(unique_meps(items)) or "-"


def missing_text(missing):
    """Parametri non trovati, raggruppati: "e_GEN_Void_Status (Pipes, Ducts); ..."; vuoto se non ne mancano."""
    by_param = {}
    for text in missing:
        name, _, where = text.partition(" on ")
        by_param.setdefault(name, set()).add(where)
    return "; ".join("{} ({})".format(name, ", ".join(sorted(where))) for name, where in sorted(by_param.items()))


def revit_owner():
    """Finestra principale di Revit, proprietaria delle finestre non modali (__revit__ c'e' solo nello script del
    comando, non nei moduli importati: si passa da pyrevit.HOST_APP)."""
    try:
        from pyrevit import HOST_APP
        return HOST_APP.uiapp.MainWindowHandle
    except Exception:
        return None
