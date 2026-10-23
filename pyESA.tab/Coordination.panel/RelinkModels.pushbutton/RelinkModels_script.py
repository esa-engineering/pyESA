# -*- coding: utf-8 -*-
__title__   = "Relink\nModels"
__doc__     = """Version = 1.5.0
Date    = 23.10.2026
Checks the Revit and CAD links of closed models and repoints them to a new folder.
Author(s): Matteo Maiorano
"""
__author__  = "Matteo Maiorano"
__context__ = "zero-doc"

# ======================= STORICO REVISIONI =======================
# Versione | Data       | Tipo | Descrizione                                   | PR
# 1.5.0    | 23.10.2026 | FEAT | Opzione: apre e salva come nuova central dopo la scrittura | -
# 1.4.1    | 23.10.2026 | FIX  | IsTransmitted = True in scrittura; Check segnala i link riscritti in attesa | -
# 1.4.0    | 23.10.2026 | UI   | Tabella raggruppata per modello (intestazioni con totali, apri/chiudi) | -
# 1.3.0    | 23.10.2026 | FEAT | Check (file esiste + stato link), Apply riscrive tutti i link trovati | -
# 1.2.0    | 22.10.2026 | UI   | Link non trovati nascosti (opzione), verifica in rilettura corretta | -
# 1.1.0    | 22.10.2026 | UI   | Nascosti i link annidati, colonna Status prima dei percorsi | -
# 1.0.1    | 22.10.2026 | FIX  | Colonne DataTable non piu' ReadOnly (bloccava Apply) | -
# 1.0.0    | 22.10.2026 | FEAT | Prima versione: TransmissionData su modelli chiusi | -
# Testato su Revit: non ancora provato in Revit
# =================================================================

from RelinkModels_Form_ui import show_relink_form

show_relink_form(__revit__.Application)  # noqa: F821
