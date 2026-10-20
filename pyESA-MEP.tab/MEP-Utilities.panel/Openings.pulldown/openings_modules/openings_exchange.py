# -*- coding: utf-8 -*-
"""Openings: file del report JSON tra il modello Openings e i modelli MEP / ARC: scrittura, lettura, cartella e file
ricordati per modello, trasformazioni. Export e Import (che usano le righe del report) stanno in openings_report.
Formato in Openings_NOTE-LINK.md ("Report JSON")."""
import io

from openings_check import *  # noqa: F401,F403

REPORT_FORMAT = "ESA_Openings_Report"
REPORT_VERSION = 1
EXPORT_FOLDER_OPTION = "export_folder"   # ultima cartella dell'Export, per modello Openings
REPORT_FILE_OPTION = "report_file"       # ultimo report importato, per modello MEP / ARC (riaperto da Last Report)


# ----------------------------------------------------------------------------------------------------------------
# File JSON
# ----------------------------------------------------------------------------------------------------------------
def write_json_file(path, data):
    """Scrive data come JSON indentato, tramite un file temporaneo (come mepqto_store). ensure_ascii=True in
    IronPython 2.7 fallisce sulle lettere accentate: si serializza senza escape e si fa l'escape \\uXXXX a mano."""
    text = json.dumps(data, ensure_ascii=False, indent=1, sort_keys=True)
    text = u"".join(c if ord(c) < 128 else u"\\u{:04x}".format(ord(c)) for c in text)
    temp_path = path + ".tmp"
    with io.open(temp_path, "w", encoding="utf-8") as handle:
        handle.write(unicode(text))  # noqa: F821 - IronPython 2.7
    if os.path.exists(path):
        os.remove(path)
    os.rename(temp_path, path)


def read_report_file(path):
    """Report esportato; ValueError se il file non e' un report di Openings (o di una versione piu' nuova)."""
    with io.open(path, "r", encoding="utf-8-sig") as handle:
        data = json.loads(handle.read())
    if not isinstance(data, dict) or data.get("format") != REPORT_FORMAT:
        raise ValueError("not an Openings report (format {})".format(REPORT_FORMAT))
    if int(data.get("version", 0)) > REPORT_VERSION:
        raise ValueError("written by a newer version of Openings (version {}): update pyRevit".format(
            data.get("version")))
    return data


def model_value(option):
    """Valore singolo ricordato per il modello aperto (percorso di cartella o file), "" se non c'e'."""
    values = load_model_list(option)
    return sorted(values)[0] if values else ""


def save_model_value(option, value):
    save_model_list(option, [value])


# ----------------------------------------------------------------------------------------------------------------
# Export (modello Openings)
# ----------------------------------------------------------------------------------------------------------------
def xyz_list(point):
    return [point.X, point.Y, point.Z]


def transform_list(transform):
    """[BasisX, BasisY, BasisZ, Origin] in piedi; identita' per il modello aperto."""
    transform = transform or DB.Transform.Identity
    return [xyz_list(transform.BasisX), xyz_list(transform.BasisY), xyz_list(transform.BasisZ),
            xyz_list(transform.Origin)]


def transform_of(values):
    """DB.Transform da transform_list; None se il dato non e' leggibile."""
    try:
        transform = DB.Transform.Identity
        transform.BasisX, transform.BasisY, transform.BasisZ, transform.Origin = [DB.XYZ(*v) for v in values]
        return transform
    except Exception:
        return None
