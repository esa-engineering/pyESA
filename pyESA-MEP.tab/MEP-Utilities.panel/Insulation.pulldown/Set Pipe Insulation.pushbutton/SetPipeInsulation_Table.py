# -*- coding: utf-8 -*-
"""Tabelle degli spessori di isolamento: tabella DPR 412/93 interna e lettura del file txt.

Nessun import di Revit o WPF: lo usano sia la finestra (anteprima) sia lo script.
"""
import csv

# Zone termiche valide (intestazioni della tabella e valori di e_EAN_ThermalPosition_1)
THERMAL_ZONES = ["Interno", "Non riscaldato", "Interrato", "Esterno"]
DIAMETER_COLUMN = "Diameter"

# Tabella degli spessori di isolamento DPR 412/93, in base al diametro e alla zona termica
# (stessi valori dello strumento "Set Pipe Insulation - DPR412-93")
DPR412_TABLE = [
    {"Diameter": 10, "Interno": 6, "Non riscaldato": 13, "Interrato": 13, "Esterno": 25},
    {"Diameter": 15, "Interno": 9, "Non riscaldato": 13, "Interrato": 13, "Esterno": 25},
    {"Diameter": 20, "Interno": 9, "Non riscaldato": 19, "Interrato": 19, "Esterno": 32},
    {"Diameter": 25, "Interno": 9, "Non riscaldato": 19, "Interrato": 19, "Esterno": 32},
    {"Diameter": 32, "Interno": 9, "Non riscaldato": 19, "Interrato": 19, "Esterno": 32},
    {"Diameter": 40, "Interno": 13, "Non riscaldato": 25, "Interrato": 25, "Esterno": 40},
    {"Diameter": 50, "Interno": 13, "Non riscaldato": 25, "Interrato": 25, "Esterno": 40},
    {"Diameter": 65, "Interno": 19, "Non riscaldato": 25, "Interrato": 25, "Esterno": 50},
    {"Diameter": 80, "Interno": 19, "Non riscaldato": 32, "Interrato": 32, "Esterno": 55},
    {"Diameter": 100, "Interno": 19, "Non riscaldato": 32, "Interrato": 32, "Esterno": 60},
    {"Diameter": 125, "Interno": 19, "Non riscaldato": 32, "Interrato": 32, "Esterno": 60},
    {"Diameter": 150, "Interno": 19, "Non riscaldato": 32, "Interrato": 32, "Esterno": 60},
    {"Diameter": 200, "Interno": 19, "Non riscaldato": 32, "Interrato": 32, "Esterno": 60},
    {"Diameter": 250, "Interno": 19, "Non riscaldato": 32, "Interrato": 32, "Esterno": 60}
]


class TableError(Exception):
    """File letto ma non utilizzabile come tabella degli spessori (messaggio in inglese, per l'utente)."""
    pass


def builtin_table():
    """Ritorna (righe, colonne) della tabella DPR 412/93; le righe sono copie."""
    return [dict(row) for row in DPR412_TABLE], [DIAMETER_COLUMN] + list(THERMAL_ZONES)


def read_table_file(file_path):
    """Legge la tabella da un file di testo separato da tabulazioni.

    Ritorna (righe, colonne, celle_ignorate): righe ordinate per diametro, colonne nell'ordine
    Diameter + zone termiche presenti nel file, celle_ignorate = celle non intere saltate.
    Solleva TableError se mancano le colonne necessarie; IOError se il file non si apre.
    """
    rows = []
    skipped_cells = 0
    with open(file_path, 'r') as table_file:
        reader = csv.reader(table_file, delimiter='\t')
        try:
            headers = next(reader)
        except StopIteration:
            raise TableError("The table file is empty.")

        # BOM UTF-8 (file salvati da Blocco note) e spazi attorno alle intestazioni
        headers = [h.strip() for h in headers]
        if headers and headers[0].startswith("\xef\xbb\xbf"):
            headers[0] = headers[0][3:].strip()
        elif headers and headers[0].startswith(u"﻿"):
            headers[0] = headers[0][1:].strip()

        if DIAMETER_COLUMN not in headers:
            raise TableError('The table has no "Diameter" column. Check the file format '
                             '(tab-separated, see insulation_table_template.txt).')
        zones = [h for h in headers if h in THERMAL_ZONES]
        if not zones:
            raise TableError('The table has no thermal zone column ({}). Check the file format.'.format(
                ", ".join(THERMAL_ZONES)))

        for row in reader:
            if len(row) < len(headers):  # riga incompleta (anche le righe vuote finali)
                continue
            row_dict = {}
            for i, header in enumerate(headers):
                if header == DIAMETER_COLUMN or header in zones:
                    try:
                        row_dict[header] = int(row[i])
                    except ValueError:
                        skipped_cells += 1
            if DIAMETER_COLUMN in row_dict:
                rows.append(row_dict)

    rows.sort(key=lambda x: x[DIAMETER_COLUMN])
    return rows, [DIAMETER_COLUMN] + zones, skipped_cells
