# -*- coding: utf-8 -*-
"""
RelinkModels_Core.py - lettura e riscrittura dei riferimenti esterni di modelli chiusi.

Usa TransmissionData: non apre il modello, legge e scrive solo la tabella dei
riferimenti esterni (link Revit, link CAD). Nessun documento a livello di modulo.
"""

import os
import shutil
import datetime

from pyrevit import DB

from RelinkModels_Match import BACKUP_DIR_NAME

KIND_REVIT = "Revit"
KIND_CAD = "CAD"


class LinkRecord(object):
    """Un riferimento esterno letto da un modello chiuso."""

    def __init__(self, model_path, ref_id, kind, old_path, should_load):
        self.model_path = model_path      # percorso del modello ospite (stringa)
        self.ref_id = ref_id              # ElementId del tipo di link
        self.kind = kind
        self.old_path = old_path          # percorso assoluto salvato nel modello
        self.should_load = should_load
        self.load_state = ""              # testo di LinkedFileStatus all'ultimo salvataggio
        self.pending_path = ""            # gia' riscritto, in attesa che Revit apra il modello

    @property
    def name(self):
        return os.path.basename(self.old_path.replace("/", "\\"))


def _to_model_path(user_path):
    return DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(user_path)


def _kind_of(ref):
    """KIND_REVIT, KIND_CAD o None per gli altri tipi di riferimento."""
    ref_type = ref.ExternalFileReferenceType
    if ref_type == DB.ExternalFileReferenceType.RevitLink:
        return KIND_REVIT
    if ref_type == DB.ExternalFileReferenceType.CADLink:
        return KIND_CAD
    return None


def _should_load(ref):
    """False solo per i link scaricati: i link NotFound erano caricati prima dello spostamento."""
    try:
        status = ref.GetLinkedFileStatus()
        return status not in (DB.LinkedFileStatus.Unloaded, DB.LinkedFileStatus.LocallyUnloaded)
    except Exception:
        return True


# testo mostrato per LinkedFileStatus (stato del link all'ultimo salvataggio del modello)
LOAD_STATE_TEXT = {
    "Loaded": "Loaded",
    "Unloaded": "Unloaded",
    "LocallyUnloaded": "Unloaded (locally)",
    "NotFound": "Not found",
    "InClosedWorkset": "In closed workset",
    "CanBeUpgraded": "Older version",
    "Imported": "Imported",
    "Invalid": "Invalid",
}


def _load_state(ref):
    try:
        name = str(ref.GetLinkedFileStatus())
    except Exception:
        return "?"
    return LOAD_STATE_TEXT.get(name, name)


def read_links(model_path, kinds):
    """Legge i link del modello chiuso. Ritorna (lista di LinkRecord, TransmissionData).

    Salta i riferimenti cloud (non si possono riagganciare a un file) e quelli
    di tipo diverso da `kinds`. Solleva l'eccezione di Revit se il file non si legge.
    """
    td = DB.TransmissionData.ReadTransmissionData(_to_model_path(model_path))
    if td is None:
        raise Exception("No transmission data in the file")
    records = []
    for ref_id in td.GetAllExternalFileReferenceIds():
        ref = td.GetLastSavedReferenceData(ref_id)
        kind = _kind_of(ref)
        if kind is None or kind not in kinds:
            continue
        mp = ref.GetAbsolutePath()
        if getattr(mp, "CloudPath", False):
            continue
        try:
            old_path = DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(mp)
        except Exception:
            continue
        rec = LinkRecord(model_path, ref_id, kind, old_path, _should_load(ref))
        rec.load_state = _load_state(ref)
        rec.pending_path = _pending_path(td, ref_id, ref)
        records.append(rec)
    return records, td


def _user_path(model_path):
    try:
        return DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)
    except Exception:
        return ""


def _pending_path(td, ref_id, saved_ref):
    """Percorso gia' scritto (desired) ma non ancora applicato da Revit, o "".

    Finche' Revit non apre e salva il modello, last saved resta il vecchio percorso;
    si confrontano i due GetPath() perche' GetAbsolutePath() sul dato desired non e'
    affidabile (v1.0.1).
    """
    try:
        desired = _user_path(td.GetDesiredReferenceData(ref_id).GetPath())
    except Exception:
        return ""
    saved = _user_path(saved_ref.GetPath())
    if desired and _norm(desired) != _norm(saved):
        return desired
    return ""


def backup_file(model_path):
    """Copia il modello in <cartella>\\_RelinkBackup\\<data_ora>_<nome>. Ritorna la copia."""
    folder = os.path.join(os.path.dirname(model_path), BACKUP_DIR_NAME)
    if not os.path.isdir(folder):
        os.makedirs(folder)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    target = os.path.join(folder, "{}_{}".format(stamp, os.path.basename(model_path)))
    shutil.copy2(model_path, target)
    return target


def write_links(model_path, td, changes, make_backup):
    """Scrive i nuovi percorsi in un modello chiuso e li rilegge per verifica.

    `changes` e' una lista di (LinkRecord, nuovo percorso). Ritorna un dizionario
    {id(LinkRecord): (scritto, dettaglio)}; solleva se il file non si puo'
    scrivere (aperto in Revit, sola lettura, ...).
    """
    if make_backup:
        backup_file(model_path)
    for rec, new_path in changes:
        td.SetDesiredReferenceData(rec.ref_id, _to_model_path(new_path),
                                   DB.PathType.Absolute, rec.should_load)
    # Senza questo flag Revit ignora i percorsi desired all'apertura (provato in v1.4.0:
    # DesiredPath scritto correttamente, link ancora sul vecchio percorso).
    td.IsTransmitted = True
    DB.TransmissionData.WriteTransmissionData(_to_model_path(model_path), td)

    # verifica: rilegge il file e confronta il percorso desiderato.
    # Sul dato "desired" GetAbsolutePath() non restituisce il percorso scritto
    # (provato in v1.0.1: tutte le righe risultavano non confermate mentre l'XML
    # nel file era corretto), quindi si confrontano GetPath() e GetAbsolutePath().
    check = DB.TransmissionData.ReadTransmissionData(_to_model_path(model_path))
    results = {}
    for rec, new_path in changes:
        seen = []
        try:
            ref = check.GetDesiredReferenceData(rec.ref_id)
            for getter in (ref.GetPath, ref.GetAbsolutePath):
                try:
                    seen.append(DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(getter()))
                except Exception:
                    pass
        except Exception as ex:
            results[id(rec)] = (False, str(ex).split("\n")[0])
            continue
        if any(_norm(s) == _norm(new_path) for s in seen):
            results[id(rec)] = (True, "")
        else:
            results[id(rec)] = (False, "re-read gives '{}'".format(seen[0] if seen else "nothing"))
    return results


def _norm(path):
    return os.path.normcase(os.path.normpath(path or ""))


def save_as_new_central(app, model_path, changes):
    """Apre un modello gia' riscritto (trasmesso), lo salva come nuova central e lo chiude.

    E' l'automazione della domanda che Revit fa aprendo a mano un modello trasmesso
    (provato in v1.4.1): all'apertura Revit applica i percorsi desired, il salvataggio
    li rende definitivi (last saved) e azzera il flag di trasmissione.
    Un modello non condiviso viene solo aperto e salvato. Ritorna
    {id(LinkRecord): (ok, dettaglio)} rileggendo il last saved dal file.
    """
    workshared = DB.BasicFileInfo.Extract(model_path).IsWorkshared
    opts = DB.OpenOptions()
    if workshared:
        opts.DetachFromCentralOption = DB.DetachFromCentralOption.ClearTransmittedSaveAsNewCentral
    doc = app.OpenDocumentFile(_to_model_path(model_path), opts)
    try:
        if workshared:
            # l'opzione di apertura potrebbe aver gia' salvato la nuova central: in quel
            # caso un SaveAs sulla central stessa fallirebbe
            try:
                central = _user_path(doc.GetWorksharingCentralModelPath())
            except Exception:
                central = ""
            if _norm(central) != _norm(model_path):
                wso = DB.WorksharingSaveAsOptions()
                wso.SaveAsCentral = True
                wso.ClearTransmitted = True
                so = DB.SaveAsOptions()
                so.OverwriteExistingFile = True
                so.SetWorksharingOptions(wso)
                doc.SaveAs(model_path, so)
            try:
                DB.WorksharingUtils.RelinquishOwnership(doc, DB.RelinquishOptions(True),
                                                        DB.TransactWithCentralOptions())
            except Exception:
                pass
        else:
            doc.Save()
    finally:
        doc.Close(False)

    # verifica: dopo il salvataggio il nuovo percorso deve essere quello "last saved"
    check = DB.TransmissionData.ReadTransmissionData(_to_model_path(model_path))
    results = {}
    for rec, new_path in changes:
        try:
            saved = _user_path(check.GetLastSavedReferenceData(rec.ref_id).GetAbsolutePath())
        except Exception as ex:
            results[id(rec)] = (False, str(ex).split("\n")[0])
            continue
        if _norm(saved) == _norm(new_path):
            results[id(rec)] = (True, "")
        else:
            results[id(rec)] = (False, "after saving the link points to '{}'".format(saved))
    return results
