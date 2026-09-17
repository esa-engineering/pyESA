# -*- coding: utf-8 -*-
"""Rinomina guidata di famiglie e tipi secondo la codifica ESA.

Si seleziona un elemento nel modello, il tool riconosce la categoria, apre la
scheda di classificazione che le corrisponde e compone il nome. Le tabelle
sono in FamilyNaming_data.py, generato dai tre file Excel affiancati; la mappa
fra categorie Revit e schede e' in FamilyNaming_map.py; le regole di
composizione in FamilyNaming_rules.py.
"""

__title__ = "Family\nNaming"
__author__ = "pyESA"
__doc__ = """Version = 1.0
Date    = 15.09.2026
_____________________________________________________________________
Renames a family and a type following the ESA naming classification.
Select one element, then run the tool.
_____________________________________________________________________
Author(s): pyESA
"""

import re
import sys
import os

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit import DB
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult

from System import Enum

from pyrevit import revit, script

# i moduli affiancati non sono sul path quando pyRevit esegue lo script
sys.path.append(os.path.dirname(__file__))

import FamilyNaming_map as MAP          # noqa: E402
import FamilyNaming_rules as RULES      # noqa: E402
from FamilyNaming_ui import NamingWindow    # noqa: E402

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

DIALOG_TITLE = "Family Naming"

MIN_REVIT = 2022
FEET_TO_MM = 304.8


# ---------------------------------------------------------------------------
# Compatibilita' fra versioni
# ---------------------------------------------------------------------------

def get_element_id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025


def revit_version():
    try:
        return int(doc.Application.VersionNumber)
    except Exception:
        return 0


def alert(message, title=DIALOG_TITLE):
    TaskDialog.Show(title, message)


def ask_yes_no(message, title=DIALOG_TITLE):
    dialog = TaskDialog(title)
    dialog.MainInstruction = message
    dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    dialog.DefaultButton = TaskDialogResult.No
    return dialog.Show() == TaskDialogResult.Yes


# ---------------------------------------------------------------------------
# Selezione e rilevamento della categoria
# ---------------------------------------------------------------------------

def get_selected_element():
    ids = list(uidoc.Selection.GetElementIds())
    if not ids:
        alert("Select one element in the model, then run the tool again.")
        script.exit()
    if len(ids) > 1:
        alert("Select a single element. The tool renames one family and one "
              "type per run, and {0} elements are selected.".format(len(ids)))
        script.exit()
    return doc.GetElement(ids[0])


def builtin_category_name(element):
    """Nome della BuiltInCategory, del tipo 'OST_Walls'."""
    category = element.Category
    if category is None:
        return None
    value = get_element_id_value(category.Id)
    try:
        bic = Enum.ToObject(clr.GetClrType(DB.BuiltInCategory), value)
        return bic.ToString()
    except Exception:
        return None


def resolve_sheet(element, element_type, bic_name):
    """Sceglie la scheda quando la categoria Revit ne ammette piu' di una.

    Tre casi reali di ambiguita':
      - i solai: un pavimento ordinario e' 02_FL, una platea e' 01_FS
      - le fondazioni: platea, trave rovescia e plinto stanno tutti in
        OST_StructuralFoundation, ma le prime due sono famiglie di sistema e
        il plinto e' caricabile
      - i pannelli di facciata: pannello di sistema oppure famiglia caricabile
        annidata nella griglia
    """
    candidates = MAP.CATEGORY_MAP.get(bic_name)
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]

    is_loadable = isinstance(element_type, DB.FamilySymbol)

    if bic_name == "OST_CurtainWallPanels":
        return u"LOADABLE:03_CP" if is_loadable else u"SYSTEM:13_CP"

    if bic_name == "OST_StructuralFoundation":
        if is_loadable:
            return u"STRUCTURAL:03_FI"
        try:
            if isinstance(element_type, DB.WallFoundationType):
                return u"STRUCTURAL:02_FW"
        except AttributeError:
            pass                      # classe assente su alcune versioni
        return u"STRUCTURAL:01_FS"

    return candidates[0]


def _normalize(text):
    return re.sub(r"[^a-z0-9]", "", (text or u"").lower())


def detect_cat_code(sheet, element_type, bic_name):
    """Riconosce la famiglia di sistema. Restituisce (codice, certo).

    Il codice non si chiede all'utente: si legge dal modello. Tre strade, in
    ordine di affidabilita'.

      1. Alcune categorie Revit fissano il codice da sole, perche' sono
         categorie distinte anche dove la scheda le tiene insieme: un bordo
         di solaio e' OST_EdgeSlab, non OST_Floors.
      2. Le pareti si distinguono su WallType.Kind, che e' una proprieta'
         esplicita e non un nome.
      3. Tutte le altre si riconoscono confrontando il nome della famiglia di
         sistema con le etichette della tabella, normalizzando spazi e
         trattini: 'Non-Monolithic Run' di Revit contro 'NonMonolithic Run'
         della tabella.

    Se nessuna strada porta a casa si assume la prima voce e si segnala
    l'incertezza, perche' l'utente non ha un menu con cui correggere.
    """
    rows = RULES.table(sheet["cat_table"])
    if not rows:
        return u"", True

    forced = MAP.FORCED_CAT_CODE.get(bic_name)
    if forced:
        return forced, True

    if isinstance(element_type, DB.WallType):
        try:
            return (u"CW" if element_type.Kind == DB.WallKind.Curtain
                    else u"BW"), True
        except Exception:
            pass

    # una scheda con una sola voce non ha nulla da riconoscere
    if len(rows) == 1:
        return rows[0][0], True

    family_name = u""
    try:
        family_name = element_type.FamilyName or u""
    except Exception:
        pass

    target = _normalize(family_name)
    if target:
        for row in rows:
            if len(row) > 1 and _normalize(row[1]) == target:
                return row[0], True
        for row in rows:
            if len(row) > 1 and target.startswith(_normalize(row[1])):
                return row[0], True

    return rows[0][0], False


def is_in_place(element, element_type):
    try:
        if isinstance(element_type, DB.FamilySymbol):
            return bool(element_type.Family.IsInPlace)
    except Exception:
        pass
    try:
        return bool(element.Symbol.Family.IsInPlace)
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Lettura delle misure dai parametri di tipo
# ---------------------------------------------------------------------------

def to_mm(internal_value):
    """Piedi interni verso millimetri interi arrotondati."""
    return int(round(internal_value * FEET_TO_MM))


def _param_length_mm(param):
    if param is None or not param.HasValue:
        return None
    if param.StorageType != DB.StorageType.Double:
        return None
    return to_mm(param.AsDouble())


def read_dim_source(element_type, sources):
    """Prova le sorgenti in ordine, restituisce la prima che risponde."""
    for kind, key in sources:
        value = None
        if kind == "bip":
            bip = getattr(DB.BuiltInParameter, key, None)
            if bip is not None:
                try:
                    value = _param_length_mm(element_type.get_Parameter(bip))
                except Exception:
                    value = None
        elif kind == "name":
            try:
                value = _param_length_mm(element_type.LookupParameter(key))
            except Exception:
                value = None
        elif kind == "prop":
            value = _read_special(element_type, key)
        if value:
            return str(value)
    return None


def _read_special(element_type, key):
    """Misure che non sono parametri di tipo ma proprieta' dell'API."""
    if key == "wall_width":
        try:
            return to_mm(element_type.Width)
        except Exception:
            return None
    if key == "compound_width":
        structure = _compound_structure(element_type)
        if structure is not None:
            try:
                return to_mm(structure.GetWidth())
            except Exception:
                return None
        return None
    if key == "railing_height":
        try:
            param = element_type.get_Parameter(
                DB.BuiltInParameter.RAILING_HEIGHT)
            return _param_length_mm(param)
        except Exception:
            return None
    return None


def _compound_structure(element_type):
    try:
        return element_type.GetCompoundStructure()
    except Exception:
        return None


def read_dimensions(sheet, element_type):
    values = {}
    for key, sources in (sheet.get("dim_src") or {}).items():
        found = read_dim_source(element_type, sources)
        if found:
            values[key] = found
    return values


def deduce_nfinishings(element_type, table_name):
    """Conta le facce finite dalla posizione degli strati rispetto al core.

    Una faccia si considera finita quando su quel lato esiste almeno uno
    strato fuori dal core, cioe' nello shell esterno o in quello interno.
    Non conta a che funzione sia assegnato lo strato: conta che ci sia.
    Uno strato per lato, o dieci, fanno lo stesso una faccia.

    Ne discende da sola la regola dello strato singolo: un elemento fatto di
    un solo strato ha quello strato dentro al core, nessuno shell, e quindi
    e' 0F anche quando quell'unico strato e' proprio la finitura.
    """
    structure = _compound_structure(element_type)
    if structure is None:
        return None

    faces = _count_shell_sides(structure)
    if faces is None:
        return None

    # i controsoffitti si fermano a una faccia finita
    allowed = [row[0] for row in RULES.table(table_name)]
    code = u"{0}F".format(faces)
    if allowed and code not in allowed:
        code = allowed[-1]
    return code


def _count_shell_sides(structure):
    """Quanti dei due lati hanno almeno uno strato fuori dal core."""
    try:
        exterior = structure.GetNumberOfShellLayers(
            DB.ShellLayerType.Exterior)
        interior = structure.GetNumberOfShellLayers(
            DB.ShellLayerType.Interior)
        return (1 if exterior > 0 else 0) + (1 if interior > 0 else 0)
    except Exception:
        pass

    # ripiego per i tipi che non espongono gli shell: gli indici del core
    # dicono la stessa cosa, cioe' se c'e' qualcosa prima o dopo di esso
    try:
        count = structure.LayerCount
        first = structure.GetFirstCoreLayerIndex()
        last = structure.GetLastCoreLayerIndex()
    except Exception:
        return None
    return (1 if first > 0 else 0) + (1 if last < count - 1 else 0)


# ---------------------------------------------------------------------------
# Lettura di quello che c'e' gia' nel progetto
# ---------------------------------------------------------------------------

def type_mark_param(element_type):
    return element_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK)


def fire_rating_param(element_type):
    param = element_type.get_Parameter(DB.BuiltInParameter.FIRE_RATING)
    if param is None:
        param = element_type.LookupParameter("Fire Rating")
    return param


def set_element_name(element, new_name):
    """Rinomina un elemento.

    L'assegnazione diretta e' l'idiom usato nel resto dell'estensione. Su
    alcune classi IronPython risolve pero' la proprieta' sbagliata e la
    dichiara in sola lettura: in quel caso si passa dal descrittore di
    Element, che e' lo stesso usato in lettura.
    """
    try:
        element.Name = new_name
    except Exception:
        DB.Element.Name.SetValue(element, new_name)


def param_string(param):
    if param is None:
        return u""
    try:
        return (param.AsString() or u"").strip()
    except Exception:
        return u""


def collect_existing_marks():
    """Tutti i Type Mark gia' assegnati nel documento corrente.

    Il primo sequenziale libero si cerca per prefisso su questo insieme, non
    per categoria: vedi la nota in testa a FamilyNaming_map.py.
    """
    marks = []
    collector = DB.FilteredElementCollector(doc).WhereElementIsElementType()
    for element_type in collector:
        try:
            value = param_string(type_mark_param(element_type))
        except Exception:
            continue
        if value:
            marks.append(value)
    return marks


def collect_sibling_type_names(element_type):
    """Nomi gia' usati fra i tipi che condividono lo stesso spazio dei nomi."""
    names = set()
    if isinstance(element_type, DB.FamilySymbol):
        try:
            for sid in element_type.Family.GetFamilySymbolIds():
                symbol = doc.GetElement(sid)
                if symbol is not None:
                    names.add(DB.Element.Name.GetValue(symbol))
        except Exception:
            pass
        return names

    # serve il System.Type concreto: type() restituirebbe il wrapper Python
    collector = DB.FilteredElementCollector(doc).OfClass(element_type.GetType())
    for other in collector:
        try:
            names.add(DB.Element.Name.GetValue(other))
        except Exception:
            pass
    return names


def collect_family_names():
    names = set()
    for family in DB.FilteredElementCollector(doc).OfClass(DB.Family):
        try:
            names.add(family.Name)
        except Exception:
            pass
    return names


# ---------------------------------------------------------------------------
# Applicazione
# ---------------------------------------------------------------------------

def apply_naming(element_type, result):
    """Scrive nomi e parametri. Restituisce la lista di cio' che ha fatto."""
    done = []
    in_place = result.get("in_place")
    is_loadable_schema = result.get("schema") == "loadable"
    symbol = element_type if isinstance(element_type, DB.FamilySymbol) else None

    if in_place and symbol is not None:
        # un elemento in place non ha due nomi indipendenti: si scrive la
        # stessa stringa sulla famiglia e sul suo unico tipo
        combined = RULES.compose_in_place_name(
            result["family_name"] if is_loadable_schema else u"",
            result["type_name"])
        family = symbol.Family
        if family.Name != combined:
            set_element_name(family, combined)
            done.append(u"Family renamed to **{0}**".format(combined))
        if DB.Element.Name.GetValue(symbol) != combined:
            set_element_name(symbol, combined)
            done.append(u"Type renamed to **{0}**".format(combined))
    else:
        if is_loadable_schema and symbol is not None and result["family_name"]:
            family = symbol.Family
            if family.Name != result["family_name"]:
                set_element_name(family, result["family_name"])
                done.append(u"Family renamed to **{0}**".format(
                    result["family_name"]))
        if result["type_name"]:
            if DB.Element.Name.GetValue(element_type) != result["type_name"]:
                set_element_name(element_type, result["type_name"])
                done.append(u"Type renamed to **{0}**".format(
                    result["type_name"]))

    if result.get("write_type_mark") and result.get("type_mark"):
        param = type_mark_param(element_type)
        if param is not None and not param.IsReadOnly:
            param.Set(result["type_mark"])
            done.append(u"Type Mark set to **{0}**".format(result["type_mark"]))
        else:
            done.append(u"Type Mark could not be written: the parameter is "
                        u"read only on this type.")

    if result.get("fire_rating"):
        param = fire_rating_param(element_type)
        if param is not None and not param.IsReadOnly:
            param.Set(result["fire_rating"])
            done.append(u"FireRating set to **{0}**".format(
                result["fire_rating"]))

    return done


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

def report(element_type, sheet_id, sheet, result, done):
    output.print_md(u"# Family Naming")
    output.print_md(u"Naming sheet: **{0}** ({1})".format(
        sheet["label"], sheet_id.replace(u":", u" / ")))

    if done:
        output.print_md(u"## Applied")
        for line in done:
            output.print_md(u"- {0}".format(line))
    else:
        output.print_md(u"## Nothing changed")
        output.print_md(
            u"The composed names are identical to the ones already in place, "
            u"and no parameter needed updating.")

    output.print_md(u"## Result")
    rows = []
    if result.get("schema") == "loadable" and not result.get("in_place"):
        rows.append([u"Family name", result.get("family_name") or u"-"])
    rows.append([u"Type name", result.get("type_name") or u"-"])
    rows.append([u"Type Mark", result.get("type_mark") or u"-"])
    rows.append([u"FireRating", result.get("fire_rating") or u"-"])
    output.print_table(rows, columns=[u"Field", u"Value"])

    try:
        output.print_md(u"Renamed element: {0}".format(
            output.linkify(element_type.Id)))
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if doc.IsFamilyDocument:
        alert("This tool works on a project, not inside the Family Editor.\n\n"
              "Open a project model and run it again.")
        return

    version = revit_version()
    if version and version < MIN_REVIT:
        alert("This tool requires Revit {0} or newer.".format(MIN_REVIT))
        return

    element = get_selected_element()
    element_type = doc.GetElement(element.GetTypeId())
    if element_type is None:
        # la selezione potrebbe gia' essere un tipo, non un'istanza
        if isinstance(element, DB.ElementType):
            element_type = element
        else:
            alert("The selected element has no type to rename.")
            return

    bic_name = builtin_category_name(element)
    category_label = element.Category.Name if element.Category else "unknown"

    sheet_id = resolve_sheet(element, element_type, bic_name)
    if sheet_id is None:
        alert(
            "Category '{0}' is not covered by the ESA naming classification.\n\n"
            "The three classification sheets cover architectural loadable "
            "families, architectural system families and structural families. "
            "MEP categories, annotation and detail items are out of scope.\n\n"
            "Nothing was renamed.".format(category_label))
        return

    sheet = MAP.SHEETS[sheet_id]

    min_revit = sheet.get("min_revit")
    if min_revit and version and version < min_revit:
        alert("The '{0}' sheet applies from Revit {1} onwards.".format(
            sheet["label"], min_revit))
        return

    current_type_name = DB.Element.Name.GetValue(element_type)
    current_family_name = u""
    if isinstance(element_type, DB.FamilySymbol):
        try:
            current_family_name = element_type.Family.Name
        except Exception:
            pass

    cat_code, cat_certain = detect_cat_code(sheet, element_type, bic_name)

    ctx = {
        "sheet_id": sheet_id,
        "sheet": sheet,
        "cat_code": cat_code,
        "cat_uncertain": not cat_certain,
        "revit_category": category_label,
        "in_place": is_in_place(element, element_type),
        "element_label": u"{0}  :  {1}".format(
            current_family_name or u"system family", current_type_name),
        "current_family": current_family_name,
        "current_type": current_type_name,
        "current_type_mark": param_string(type_mark_param(element_type)),
        "current_fire_rating": param_string(fire_rating_param(element_type)),
        "dim_defaults": read_dimensions(sheet, element_type),
        "nfin_default": deduce_nfinishings(
            element_type, sheet.get("nfin_table") or MAP.TBL_NFIN)
        if sheet.get("dim") == "nfin_t" else None,
        "existing_marks": collect_existing_marks(),
        "existing_type_names": collect_sibling_type_names(element_type),
        "existing_family_names": collect_family_names(),
    }

    window = NamingWindow(ctx)
    window.ShowDialog()

    result = window.result
    if result is None:
        return

    # la sovrascrittura di un FireRating gia' compilato si chiede sempre
    existing_rating = ctx["current_fire_rating"]
    if result.get("fire_rating") and existing_rating \
            and existing_rating != result["fire_rating"]:
        keep = not ask_yes_no(
            "This type already has FireRating '{0}'.\n\n"
            "Overwrite it with '{1}'?".format(
                existing_rating, result["fire_rating"]))
        if keep:
            result["fire_rating"] = u""

    with revit.Transaction("Rename family and type"):
        done = apply_naming(element_type, result)

    output.close_others()
    report(element_type, sheet_id, sheet, result, done)


main()
