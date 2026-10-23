# -*- coding: utf-8 -*-
"""
RelinkModels_Match.py - abbinamento dei link ai file trovati nella nuova cartella.

Modulo di Python puro (nessun import di Revit): si puo' provare anche fuori da Revit.
Compatibile con IronPython 2.7.
"""

import os

BACKUP_DIR_NAME = "_RelinkBackup"

STATUS_MATCH = "Match"
STATUS_SAME = "Already correct"
STATUS_AMBIGUOUS = "Ambiguous"
STATUS_NOT_FOUND = "Not in folder"


def scan_folder(root, include_subfolders, extensions):
    """Indice {nome file minuscolo: [percorsi completi]} dei file con l'estensione data.

    Le cartelle dei backup del tool sono saltate, altrimenti ogni copia di
    sicurezza diventerebbe un candidato in piu'.
    """
    index = {}
    if include_subfolders:
        walker = os.walk(root)
    else:
        try:
            walker = [(root, [], os.listdir(root))]
        except OSError:
            walker = []
    for folder, dirs, files in walker:
        if BACKUP_DIR_NAME in dirs:
            dirs.remove(BACKUP_DIR_NAME)
        for name in files:
            if os.path.splitext(name)[1].lower() in extensions:
                index.setdefault(name.lower(), []).append(os.path.join(folder, name))
    return index


def _parts(path):
    return [p for p in path.replace("/", "\\").lower().split("\\") if p]


def _tail_score(old_path, candidate):
    """Quanti elementi finali del percorso coincidono (nome file compreso)."""
    a = _parts(old_path)
    b = _parts(candidate)
    score = 0
    while score < len(a) and score < len(b) and a[-1 - score] == b[-1 - score]:
        score += 1
    return score


def _same_path(a, b):
    return os.path.normcase(os.path.normpath(a)) == os.path.normcase(os.path.normpath(b))


def pick_candidate(old_path, index):
    """Sceglie il nuovo percorso per un link.

    Ritorna (nuovo_percorso, stato, numero_candidati). Con piu' file dallo stesso
    nome vince quello che condivide piu' cartelle finali con il vecchio percorso;
    a parita' il caso e' ambiguo e il percorso resta vuoto.
    """
    name = os.path.basename(old_path.replace("/", "\\")).lower()
    candidates = index.get(name, [])
    if not candidates:
        return "", STATUS_NOT_FOUND, 0
    scored = sorted(((_tail_score(old_path, c), c) for c in candidates), reverse=True)
    best_score = scored[0][0]
    best = [c for s, c in scored if s == best_score]
    if len(best) > 1:
        return "", STATUS_AMBIGUOUS, len(candidates)
    if _same_path(old_path, best[0]):
        return best[0], STATUS_SAME, len(candidates)
    return best[0], STATUS_MATCH, len(candidates)
