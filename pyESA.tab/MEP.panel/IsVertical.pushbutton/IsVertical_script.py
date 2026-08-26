# -*- coding: utf-8 -*-
"""
MEP_IsVertical
==============

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

Note sulla conversione
----------------------
* Il confronto originale e' "abs(altitude) == 90", un'uguaglianza esatta su
  float. Qui si usa una tolleranza angolare (ANGLE_TOLERANCE_DEG) impostata a
  1e-6 gradi: comportamento praticamente identico ma senza il rischio che un
  elemento verticale venga scartato per errore di arrotondamento.
  Portando la tolleranza a valori maggiori (es. 0.5) si intercettano anche
  elementi "quasi verticali", cosa che il grafo Dynamo non fa.
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


def is_vertical(element, tolerance_deg=ANGLE_TOLERANCE_DEG):
    """True / False / None (None = elemento non valutabile)."""
    endpoints = get_curve_endpoints(element)
    if endpoints is None:
        return None

    altitude = vector_altitude_deg(endpoints[0], endpoints[1])
    if altitude is None:
        return None

    return abs(abs(altitude) - 90.0) <= tolerance_deg


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


# --------------------------------------------------------------------------
# ESECUZIONE
# --------------------------------------------------------------------------

def main():
    elements = collect_elements(doc, TARGET_CATEGORIES)

    vertical = []
    horizontal = []
    skipped = []

    for element in elements:
        result = is_vertical(element)
        if result is None:
            skipped.append((element, "location curve assente o lunghezza nulla"))
        elif result:
            vertical.append(element)
        else:
            horizontal.append(element)

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
            for element in horizontal:
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
        "horizontal": len(horizontal),
        "skipped": skipped,
        "updated": updated,
        "failures": failures,
    }


def report(res):
    """Stampa il riepilogo. Usa pyrevit.script se disponibile."""
    try:
        from pyrevit import script

        output = script.get_output()
        printer = output.print_md
        linkify = output.linkify
    except Exception:
        def printer(text):
            print(text)

        def linkify(element_id):
            return "Id {0}".format(element_id)

    printer("## MEP_IsVertical")
    printer("- Elementi analizzati (tubazioni + canali): **{0}**".format(res["total"]))
    printer("- Verticali: **{0}**".format(res["vertical"]))
    printer("- Non verticali: **{0}**".format(res["horizontal"]))
    printer("- Parametri scritti: **{0}**".format(res["updated"]))
    printer("- Elementi saltati: **{0}**".format(len(res["skipped"])))
    printer("- Errori di scrittura: **{0}**".format(len(res["failures"])))

    if res["skipped"]:
        printer("### Elementi saltati")
        for element, reason in res["skipped"][:50]:
            printer("- {0} - {1}".format(linkify(element.Id), reason))
        if len(res["skipped"]) > 50:
            printer("- ... e altri {0}".format(len(res["skipped"]) - 50))

    if res["failures"]:
        printer("### Errori di scrittura")
        for element, reason in res["failures"][:50]:
            printer("- {0} - {1}".format(linkify(element.Id), reason))
        if len(res["failures"]) > 50:
            printer("- ... e altri {0}".format(len(res["failures"]) - 50))


if __name__ == "__main__":
    report(main())