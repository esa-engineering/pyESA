# -*- coding: utf-8 -*-
"""
MEP_IsVertical
==============
Conversione in IronPython 2 (pyRevit) dello script Dynamo "MEP_IsVertical.dyn",
con in aggiunta il report dei tratti quasi verticali (fuori piombo).

Logica del grafo originale
--------------------------
1. Raccoglie tutti gli elementi delle categorie OST_PipeCurves (tubazioni)
   e OST_DuctCurves (canali) e li unisce in una sola lista (List.Join).
2. Per ogni elemento ricava i punti estremi della location curve
   (Clockwork "Element.Location+" -> output "curveEndpoints").
3. Costruisce il vettore start -> end (Vector.ByTwoPoints).
4. Calcola l'altitudine del vettore (Clockwork "Vector.AltitudeAndAzimuth"),
   ne prende il valore assoluto e verifica se e' uguale a 90 gradi.
   Altitudine = angolo fra il vettore e il piano XY.
5. Filtra con FilterByBoolMask e scrive il parametro "u_CON_IsVertical":
   True sugli elementi verticali, False su tutti gli altri.

Integrazione rispetto al grafo Dynamo
-------------------------------------
La finestra di output elenca i tratti QUASI verticali, cioe' quelli la cui
inclinazione cade nella fascia [NEAR_VERTICAL_MIN_DEG, 90 gradi) esclusa la
verticale esatta. Sono i probabili errori di modellazione: nascono verticali
ma hanno uno scostamento planimetrico fra i due estremi. Per ciascuno vengono
riportati Id cliccabile, categoria e inclinazione.
Gli orizzontali e le diagonali intenzionali restano fuori dall'elenco.
La scrittura del parametro NON cambia: i quasi verticali ricevono False,
esattamente come nel grafo originale.

Note sulla conversione
----------------------
* Il confronto originale e' "abs(altitude) == 90", un'uguaglianza esatta su
  float. Qui si usa una tolleranza angolare (ANGLE_TOLERANCE_DEG) impostata a
  1e-6 gradi: comportamento praticamente identico ma senza il rischio che un
  elemento verticale venga scartato per errore di arrotondamento.
* Gli elementi senza location curve (es. tubazioni di raccordo anomale) in
  Dynamo generavano null; qui vengono saltati e riportati nel log.
* La scrittura del parametro rispetta lo StorageType effettivo
  (Yes/No -> 1/0, Integer -> 1/0, String -> "True"/"False").

Compatibilita': IronPython 2.7 (pyRevit), Revit 2019+.
"""

import math

from Autodesk.Revit.DB import (
    BuiltInCategory,
    FilteredElementCollector,
    LocationCurve,
    StorageType,
    Transaction,
)

# --------------------------------------------------------------------------
# CONFIGURAZIONE
# --------------------------------------------------------------------------

PARAM_NAME = "u_CON_IsVertical"

# Categorie da processare (equivalente dei due nodi Categories + List.Join)
TARGET_CATEGORIES = [
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_DuctCurves,
]

# Tolleranza sul confronto abs(altitudine) == 90 gradi.
ANGLE_TOLERANCE_DEG = 1e-6

# Se True scrive False sugli elementi non verticali (come nel grafo Dynamo).
# Se False lascia il parametro invariato sui non verticali.
WRITE_FALSE_ON_HORIZONTAL = True

# Soglia inferiore della fascia "quasi verticale" da elencare nel report.
# Un tratto con inclinazione fra questo valore e 90 gradi (escluso) viene
# segnalato come fuori piombo. Abbassare per allargare la ricerca.
NEAR_VERTICAL_MIN_DEG = 80.0

# Numero massimo di righe stampate in tabella (0 = nessun limite).
MAX_REPORT_ROWS = 300


# --------------------------------------------------------------------------
# DOCUMENTO
# --------------------------------------------------------------------------

doc = __revit__.ActiveUIDocument.Document  # noqa: F821  (fornito da pyRevit)


# --------------------------------------------------------------------------
# FUNZIONI DI SUPPORTO
# --------------------------------------------------------------------------

def collect_elements(document, categories):
    """Equivalente di ElementsOfCategory + List.Join."""
    elements = []
    for bic in categories:
        collector = (
            FilteredElementCollector(document)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
        )
        elements.extend(list(collector))
    return elements


def get_curve_endpoints(element):
    """Equivalente dell'output 'curveEndpoints' di Clockwork Element.Location+.

    Restituisce (start, end) come XYZ, oppure None se l'elemento non ha
    una location curve.
    """
    location = element.Location
    if not isinstance(location, LocationCurve):
        return None
    curve = location.Curve
    if curve is None:
        return None
    return curve.GetEndPoint(0), curve.GetEndPoint(1)


def vector_altitude_deg(start, end):
    """Equivalente di Vector.ByTwoPoints + Clockwork Vector.AltitudeAndAzimuth.

    L'altitudine e' l'angolo (in gradi) fra il vettore e il piano XY:
    +90 verso l'alto, -90 verso il basso, 0 orizzontale.
    Restituisce None per vettori di lunghezza nulla.
    """
    dx = end.X - start.X
    dy = end.Y - start.Y
    dz = end.Z - start.Z

    length = math.sqrt(dx * dx + dy * dy + dz * dz)
    if length < 1e-12:
        return None

    # asin del rapporto fra componente verticale e lunghezza del vettore
    ratio = dz / length
    # clamp per sicurezza numerica
    if ratio > 1.0:
        ratio = 1.0
    elif ratio < -1.0:
        ratio = -1.0

    return math.degrees(math.asin(ratio))


def classify(element):
    """Classifica un elemento in base all'inclinazione del suo asse.

    Restituisce (kind, altitude_abs) dove kind vale:
      "vertical"  -> verticale entro tolleranza (parametro True)
      "near"      -> quasi verticale, fuori piombo (parametro False, in report)
      "other"     -> orizzontale o diagonale intenzionale (parametro False)
      None        -> non valutabile (location curve assente o lunghezza nulla)
    """
    endpoints = get_curve_endpoints(element)
    if endpoints is None:
        return None, None

    altitude = vector_altitude_deg(endpoints[0], endpoints[1])
    if altitude is None:
        return None, None

    altitude_abs = abs(altitude)

    if abs(altitude_abs - 90.0) <= ANGLE_TOLERANCE_DEG:
        return "vertical", altitude_abs

    if altitude_abs >= NEAR_VERTICAL_MIN_DEG:
        return "near", altitude_abs

    return "other", altitude_abs


def set_bool_parameter(element, param_name, value):
    """Equivalente di Element.SetParameterByName con valore booleano.

    Restituisce (True, None) in caso di successo, (False, motivo) altrimenti.
    """
    param = element.LookupParameter(param_name)
    if param is None:
        return False, "parametro non trovato"
    if param.IsReadOnly:
        return False, "parametro in sola lettura"

    storage = param.StorageType
    try:
        if storage == StorageType.Integer:
            param.Set(1 if value else 0)
        elif storage == StorageType.Double:
            param.Set(1.0 if value else 0.0)
        elif storage == StorageType.String:
            param.Set("True" if value else "False")
        else:
            return False, "StorageType non gestito ({0})".format(storage)
    except Exception as ex:
        return False, "errore in Set(): {0}".format(ex)

    return True, None


def category_name(element):
    try:
        if element.Category is not None:
            return element.Category.Name
    except Exception:
        pass
    return "n/d"


# --------------------------------------------------------------------------
# ESECUZIONE
# --------------------------------------------------------------------------

def main():
    elements = collect_elements(doc, TARGET_CATEGORIES)

    vertical = []
    near_vertical = []   # lista di (element, altitudine_assoluta)
    other = []
    skipped = []

    for element in elements:
        kind, altitude_abs = classify(element)
        if kind is None:
            skipped.append((element, "location curve assente o lunghezza nulla"))
        elif kind == "vertical":
            vertical.append(element)
        elif kind == "near":
            near_vertical.append((element, altitude_abs))
            other.append(element)
        else:
            other.append(element)

    # i piu' vicini alla verticale per primi: sono i piu' sospetti
    near_vertical.sort(key=lambda pair: pair[1], reverse=True)

    updated = 0
    failures = []

    transaction = Transaction(doc, "MEP_IsVertical - scrittura " + PARAM_NAME)
    transaction.Start()
    try:
        for element in vertical:
            ok, reason = set_bool_parameter(element, PARAM_NAME, True)
            if ok:
                updated += 1
            else:
                failures.append((element, reason))

        if WRITE_FALSE_ON_HORIZONTAL:
            for element in other:
                ok, reason = set_bool_parameter(element, PARAM_NAME, False)
                if ok:
                    updated += 1
                else:
                    failures.append((element, reason))

        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    return {
        "total": len(elements),
        "vertical": len(vertical),
        "near_vertical": near_vertical,
        "other": len(other),
        "skipped": skipped,
        "updated": updated,
        "failures": failures,
    }


# --------------------------------------------------------------------------
# REPORT
# --------------------------------------------------------------------------

def _get_reporter():
    """Restituisce (printer, linkify, table_printer).

    Usa pyrevit.script se disponibile, altrimenti ricade su print.
    """
    try:
        from pyrevit import script

        output = script.get_output()
        return output.print_md, output.linkify, output.print_table
    except Exception:
        def printer(text):
            print(text)

        def linkify(element_id):
            return "Id {0}".format(element_id)

        def table_printer(table_data, columns=None, title=""):
            if title:
                print(title)
            if columns:
                print(" | ".join(columns))
            for row in table_data:
                print(" | ".join([str(cell) for cell in row]))

        return printer, linkify, table_printer


def report(res):
    printer, linkify, print_table = _get_reporter()

    near = res["near_vertical"]

    printer("# MEP_IsVertical")
    printer(
        "Parametro **{0}** scritto su **{1}** elementi.".format(
            PARAM_NAME, res["updated"]
        )
    )
    printer("")
    printer("| Esito | Conteggio |")
    printer("|---|---|")
    printer("| Elementi analizzati (tubazioni + canali) | {0} |".format(res["total"]))
    printer("| Verticali esatti | {0} |".format(res["vertical"]))
    printer("| Non verticali | {0} |".format(res["other"]))
    printer("| di cui quasi verticali (fuori piombo) | {0} |".format(len(near)))
    printer("| Saltati (senza location curve) | {0} |".format(len(res["skipped"])))
    printer("| Errori di scrittura parametro | {0} |".format(len(res["failures"])))
    printer("")

    # ----------------------------------------------------------------------
    # Elenco dei tratti quasi verticali
    # ----------------------------------------------------------------------
    printer(
        "## Tratti quasi verticali (fra {0:.1f} e 90 gradi)".format(
            NEAR_VERTICAL_MIN_DEG
        )
    )

    if not near:
        printer(
            "Nessun tratto fuori piombo nella fascia analizzata. "
            "Tutti i tratti a sviluppo verticale risultano perfettamente "
            "verticali."
        )
    else:
        printer(
            "Ordinati dal piu' vicino alla verticale. "
            "**Scostamento** = quanti gradi mancano ai 90 esatti. "
            "Clic sull'Id per selezionare l'elemento nel modello."
        )

        rows = near if MAX_REPORT_ROWS <= 0 else near[:MAX_REPORT_ROWS]

        table_data = []
        for element, altitude_abs in rows:
            table_data.append([
                linkify(element.Id),
                category_name(element),
                "{0:.4f}".format(altitude_abs),
                "{0:.4f}".format(90.0 - altitude_abs),
            ])

        # i valori sono gia' formattati come stringhe: nessun "formats"
        print_table(
            table_data=table_data,
            columns=["Id", "Categoria", "Inclinazione [gradi]", "Scostamento [gradi]"],
        )

        if MAX_REPORT_ROWS > 0 and len(near) > MAX_REPORT_ROWS:
            printer(
                "_Elencati i primi {0} di {1}. "
                "Aumenta MAX_REPORT_ROWS per vederli tutti._".format(
                    MAX_REPORT_ROWS, len(near)
                )
            )

    # ----------------------------------------------------------------------
    # Diagnostica
    # ----------------------------------------------------------------------
    if res["skipped"]:
        printer("## Elementi saltati")
        for element, reason in res["skipped"][:50]:
            printer("- {0} - {1}".format(linkify(element.Id), reason))
        if len(res["skipped"]) > 50:
            printer("- ... e altri {0}".format(len(res["skipped"]) - 50))

    if res["failures"]:
        printer("## Errori di scrittura parametro")
        for element, reason in res["failures"][:50]:
            printer("- {0} - {1}".format(linkify(element.Id), reason))
        if len(res["failures"]) > 50:
            printer("- ... e altri {0}".format(len(res["failures"]) - 50))


if __name__ == "__main__":
    report(main())