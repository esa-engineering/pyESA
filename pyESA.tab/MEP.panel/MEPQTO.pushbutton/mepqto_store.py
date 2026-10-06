# -*- coding: utf-8 -*-
"""
mepqto_store.py - archivio delle voci di computo, fuori dal modello Revit.

Due livelli, fusi campo per campo (il progetto vince):

* listino comune: file .json (PriceListDocument, modificabile con l'editor del
  listino) oppure .xlsx / .xlsm / .csv mantenuto in Excel, che il tool legge
  soltanto.
* file di progetto <Modello>_MEPQTO.json accanto al modello centrale: descrizioni,
  unita' e prezzi aggiunti o corretti nella finestra, regole, WBS, parametri.

Salvataggio concorrente, sia sul file di progetto sia sul listino JSON: si
ricordano le modifiche locali; se al salvataggio il file su disco e' cambiato dopo
il caricamento, lo si rilegge e si riapplicano sopra solo quelle.
"""

import io
import json
import os
import re
import csv
from datetime import datetime

import clr
clr.AddReference('System.Xml')

from System.IO import StreamReader, File
from System.Text import Encoding, UTF8Encoding
from System.Xml import XmlDocument

from pyrevit import DB

PROJECT_FILE_SUFFIX = "_MEPQTO.json"
FILE_VERSION = 1

# epu_item: n. articolo dell'elenco prezzi unitari (EPU) di progetto.
# price_book: prezzario di riferimento della voce (es. Prezzario Regione Lombardia 2026).
# short_description: descrizione breve della voce, accanto a quella estesa.
# Il codice della voce (chiave, letto dal modello) e' il codice del prezzario.
FIELDS = ("chapter", "subchapter", "epu_item", "price_book", "short_description",
          "description", "unit", "price")

# Unita' di misura proposte dalla tendina dell'elenco prezzi.
UNITS = (u"cad", u"m", u"kg", u"mq", u"mc")
# Varianti frequenti nei listini, confrontate dopo _norm_unit (minuscolo, senza
# punti ne' spazi, apici 2/3 come cifre).
UNIT_ALIASES = {
    # "n" e "nr" sono la stessa unita' di "cad": negli EPU compaiono entrambe.
    u"cad": u"cad", u"cadauno": u"cad", u"cada": u"cad", u"n": u"cad", u"nr": u"cad",
    u"n\u00b0": u"cad", u"n\u00ba": u"cad", u"nro": u"cad", u"numero": u"cad",
    u"num": u"cad", u"pz": u"cad", u"pezzi": u"cad", u"pc": u"cad", u"pcs": u"cad",
    u"ea": u"cad", u"each": u"cad",
    u"m": u"m", u"ml": u"m", u"mt": u"m", u"m1": u"m", u"metri": u"m",
    u"kg": u"kg",
    u"mq": u"mq", u"m2": u"mq", u"sqm": u"mq",
    u"mc": u"mc", u"m3": u"mc", u"cum": u"mc",
}

ORIGIN_PRICE_LIST = "Price list"
ORIGIN_PROJECT = "Project"
ORIGIN_MISSING = "Missing"

# Listino Excel / CSV: colonne in posizione fissa, come un EPU (elenco prezzi unitari).
# A codice, B descrizione sintetica, C descrizione completa, D unita', E prezzo unitario;
# F..I facoltative (le scrive l'export, cosi' si puo' reimportare).
EPU_COLUMNS = ("code", "short_description", "description", "unit", "price",
               "chapter", "subchapter", "epu_item", "price_book")
# Intestazioni della colonna A riconosciute come riga di titolo (normalizzate con
# _norm_header: solo a-z0-9). Le altre righe senza un codice del modello sono innocue:
# la ricerca parte dai codici del modello.
CODE_HEADERS = ("code", "codice", "cod", "codici", "pricecode", "itemcode", "codicevoce",
                "articolo", "art", "tariffa", "codicetariffa", "codiceprezzo",
                "codiceprezzario", "codiceprezziario", "pricebookcode", "codiceepu")
HEADER_SCAN_ROWS = 20
EURO_SIGN = u"\u20ac"


# =============================================================================
# NUMERI
# =============================================================================

def parse_decimal(text):
    """float da testo utente o CSV italiano; None se vuoto, ValueError se non valido.

    "12,5" -> 12.5   "1.234,56" -> 1234.56   "12.5" -> 12.5
    """
    if text is None:
        return None
    if isinstance(text, (int, float)):
        return float(text)
    cleaned = text.replace(EURO_SIGN, u"").replace(u" ", u"").replace(u"\u00a0", u"").strip()
    if not cleaned:
        return None
    if u"," in cleaned:
        cleaned = cleaned.replace(u".", u"").replace(u",", u".")
    return float(cleaned)


def format_decimal(value):
    """Testo di un numero per la griglia: due decimali, di piu' solo se servono."""
    if value is None:
        return u""
    if round(value, 2) == value:
        return u"{:.2f}".format(value)
    return (u"{:.6f}".format(value)).rstrip(u"0")


def normalize_unit(text):
    """Unita' del listino ricondotta ai valori della tendina quando e' una variante
    nota ("cad." -> "cad", "m2" -> "mq"); altrimenti resta com'e'."""
    raw = (text or u"").strip()
    key = raw.lower().replace(u".", u"").replace(u" ", u"")
    key = key.replace(u"\u00b2", u"2").replace(u"\u00b3", u"3")
    return UNIT_ALIASES.get(key, raw)


# =============================================================================
# VOCI
# =============================================================================

class MergedItem(object):
    """Voce di computo vista dalla finestra: listino + override di progetto."""

    __slots__ = ("code", "chapter", "subchapter", "epu_item", "price_book",
                 "short_description", "description", "unit", "price", "origin",
                 "in_price_list")

    def __init__(self, code):
        self.code = code
        self.chapter = u""
        self.subchapter = u""
        self.epu_item = u""
        self.price_book = u""
        self.short_description = u""
        self.description = u""
        self.unit = u""
        self.price = None
        self.origin = ORIGIN_MISSING
        # True se il codice e' nel listino (colonna A dell'EPU Excel)
        self.in_price_list = False


def merge_item(code, price_list_items, project_items):
    item = MergedItem(code)
    base = price_list_items.get(code)
    item.in_price_list = code in price_list_items
    override = project_items.get(code)
    overridden = False
    for source in (base, override):
        if not source:
            continue
        for field in FIELDS:
            if source.get(field) not in (None, u""):
                setattr(item, field, source[field])
                overridden = overridden or source is override
    if overridden:
        item.origin = ORIGIN_PROJECT
    elif base:
        item.origin = ORIGIN_PRICE_LIST
    return item


def merge_items(codes, price_list_items, project_items):
    return dict((code, merge_item(code, price_list_items, project_items)) for code in codes)


# =============================================================================
# LISTINO (sola lettura)
# =============================================================================

class PriceListError(Exception):
    pass


class PriceList(object):
    def __init__(self, path=None):
        self.path = path
        self.items = {}
        self.duplicates = 0
        self.sheet_name = None


def _norm_header(text):
    return re.sub(r"[^a-z0-9]", "", (text or u"").lower())


def _parse_price(text, numeric_prices):
    """Prezzo di una cella: numero grezzo dell'xlsx ("12.5") oppure testo con virgola
    decimale o simbolo dell'euro ("12,50", "€ 1.234,56"); None se vuoto o non numerico."""
    if not text:
        return None
    if numeric_prices:
        try:
            return float(text)
        except ValueError:
            pass
    try:
        return parse_decimal(text)
    except ValueError:
        return None


def _rows_to_items(rows, price_list, numeric_prices):
    """Voci dalle righe del foglio, per posizione (EPU_COLUMNS: A codice, B descrizione
    sintetica, C descrizione, D unita', E prezzo, F..I facoltative).

    Si saltano: le righe senza codice in colonna A, la riga di intestazione (colonna A
    "Codice", "Code"...) e le righe con la sola colonna A piena (titoli, capitoli), che
    non hanno ne' descrizioni ne' unita' ne' prezzo. Un codice ripetuto: vale la prima
    riga.
    """
    for cells in rows:
        def cell(index):
            return (cells[index] or u"").strip() if index < len(cells) else u""

        code = cell(0)
        if not code or _norm_header(code) in CODE_HEADERS:
            continue
        if not any(cell(index) for index in range(1, 5)):
            continue
        if code in price_list.items:
            price_list.duplicates += 1
            continue
        values = dict((field, cell(index)) for index, field in enumerate(EPU_COLUMNS))
        values["unit"] = normalize_unit(values["unit"])
        values["price"] = _parse_price(values["price"], numeric_prices)
        del values["code"]
        price_list.items[code] = values
    if not price_list.items:
        raise PriceListError(
            "No price codes found in column A of the first worksheet. The price list must "
            "have the price codes in column A, the short description in B, the description "
            "in C, the unit in D and the unit price in E.")


# --- xlsx: zip + xml, senza Excel (stesso approccio di WorksetCreate) ---------

def _open_archive(path):
    clr.AddReference('System.IO.Compression')
    clr.AddReference('System.IO.Compression.FileSystem')
    from System.IO.Compression import ZipFile
    return ZipFile.OpenRead(path)


def _read_entry(archive, name):
    entry = archive.GetEntry(name)
    if entry is None:
        lname = name.lower()
        for candidate in archive.Entries:
            if candidate.FullName.lower() == lname:
                entry = candidate
                break
    if entry is None:
        return None
    stream = entry.Open()
    try:
        reader = StreamReader(stream, Encoding.UTF8)
        try:
            return reader.ReadToEnd()
        finally:
            reader.Close()
    finally:
        stream.Close()


def _children(node, local_name):
    out = []
    for child in node.ChildNodes:
        if child.NodeType.ToString() == 'Element' and child.LocalName == local_name:
            out.append(child)
    return out


def _collect_text(node, acc):
    for child in node.ChildNodes:
        if child.NodeType.ToString() != 'Element':
            continue
        if child.LocalName == 'rPh':
            continue
        if child.LocalName == 't':
            acc.append(child.InnerText)
        else:
            _collect_text(child, acc)


def _load_shared_strings(archive):
    xml = _read_entry(archive, 'xl/sharedStrings.xml')
    if not xml:
        return []
    document = XmlDocument()
    document.LoadXml(xml)
    strings = []
    for si in _children(document.DocumentElement, 'si'):
        acc = []
        _collect_text(si, acc)
        strings.append(u''.join(acc))
    return strings


def _first_sheet(archive):
    """(nome, percorso della parte xml) del primo foglio del workbook."""
    rels = {}
    rel_xml = _read_entry(archive, 'xl/_rels/workbook.xml.rels')
    if rel_xml:
        rdoc = XmlDocument()
        rdoc.LoadXml(rel_xml)
        for rel in _children(rdoc.DocumentElement, 'Relationship'):
            target = rel.GetAttribute('Target')
            if not target:
                continue
            if target.startswith('/'):
                target = target[1:]
            elif not target.startswith('xl/'):
                target = 'xl/' + target.replace('../', '')
            rels[rel.GetAttribute('Id')] = target

    wb_xml = _read_entry(archive, 'xl/workbook.xml')
    if wb_xml:
        wdoc = XmlDocument()
        wdoc.LoadXml(wb_xml)
        for sheets_node in _children(wdoc.DocumentElement, 'sheets'):
            for sheet in _children(sheets_node, 'sheet'):
                rid = sheet.GetAttribute('r:id') or sheet.GetAttribute('id')
                return sheet.GetAttribute('name'), rels.get(rid, 'xl/worksheets/sheet1.xml')
    return None, 'xl/worksheets/sheet1.xml'


def _column_index(ref):
    """'C12' -> 2."""
    index = 0
    for ch in ref:
        if not ch.isalpha():
            break
        index = index * 26 + (ord(ch.upper()) - ord('A') + 1)
    return index - 1


def _cell_value(cell, shared):
    ctype = cell.GetAttribute('t')
    values = _children(cell, 'v')
    if ctype == 's':
        if not values:
            return u''
        try:
            idx = int(values[0].InnerText)
        except ValueError:
            return u''
        return shared[idx] if 0 <= idx < len(shared) else u''
    if ctype == 'inlineStr' or not values:
        acc = []
        _collect_text(cell, acc)
        return u''.join(acc)
    raw = values[0].InnerText
    if ctype in ('', 'n'):
        # 100.0 deve leggersi "100": i codici numerici non devono prendere un ".0"
        try:
            fval = float(raw)
            if fval == int(fval) and 'E' not in raw.upper():
                return u"{}".format(int(fval))
        except (ValueError, OverflowError):
            pass
    return raw


def _read_xlsx_rows(path, price_list):
    archive = _open_archive(path)
    try:
        shared = _load_shared_strings(archive)
        price_list.sheet_name, part = _first_sheet(archive)
        xml = _read_entry(archive, part)
        if not xml:
            raise PriceListError("The first worksheet of the workbook is empty.")
        document = XmlDocument()
        document.LoadXml(xml)
        rows = []
        for sheet_data in _children(document.DocumentElement, 'sheetData'):
            for row in _children(sheet_data, 'row'):
                cells = []
                for position, cell in enumerate(_children(row, 'c')):
                    ref = cell.GetAttribute('r')
                    index = _column_index(ref) if ref else position
                    while len(cells) < index:
                        cells.append(u'')
                    cells.append(_cell_value(cell, shared))
                rows.append(cells)
        return rows
    finally:
        archive.Dispose()


def read_text_file(path):
    """Testo UTF-8, con ripiego sulla codepage di sistema (CSV salvati da Excel)."""
    data = File.ReadAllBytes(path)
    try:
        content = UTF8Encoding(False, True).GetString(data)
    except Exception:
        content = Encoding.Default.GetString(data)
    if content and content[0] == u'\ufeff':
        content = content[1:]
    return content


def _read_csv_rows(path):
    lines = read_text_file(path).splitlines()
    if not lines:
        return []
    sample = u"\n".join(lines[:HEADER_SCAN_ROWS])
    delimiter = max((u";", u",", u"\t"), key=sample.count)
    return [list(row) for row in csv.reader(lines, delimiter=str(delimiter))]


def load_price_list(path):
    """Legge il listino. Solleva PriceListError con un messaggio per l'utente."""
    price_list = PriceList(path)
    if not path:
        return price_list
    if not os.path.isfile(path):
        raise PriceListError("Price list not found: {}".format(path))
    extension = os.path.splitext(path)[1].lower()
    try:
        if extension == ".json":
            document = PriceListDocument(path)
            document.load()
            price_list.items = document.items
            price_list.sheet_name = document.name
        elif extension in (".xlsx", ".xlsm"):
            _rows_to_items(_read_xlsx_rows(path, price_list), price_list, True)
        elif extension in (".csv", ".txt"):
            _rows_to_items(_read_csv_rows(path), price_list, False)
        else:
            raise PriceListError(
                "Unsupported price list format '{}': use .json, .xlsx, .xlsm or .csv.".format(
                    extension))
    except PriceListError:
        raise
    except IOError as error:
        raise PriceListError("Cannot read the price list (is it open and locked?): {}".format(error))
    except Exception as error:
        raise PriceListError("Cannot read the price list: {}".format(error))
    return price_list


# =============================================================================
# LISTINO JSON (modificabile con l'editor)
# =============================================================================

PRICE_LIST_FORMAT = u"ESA_MEPQTO_PriceList"
PRICE_LIST_VERSION = 1


def clean_item(values):
    """Voce normalizzata: solo i campi noti, testi senza spazi ai lati, unita' ricondotta
    ai valori della tendina, prezzo numerico o None."""
    item = {}
    for field in FIELDS:
        value = values.get(field) if isinstance(values, dict) else None
        if field == "price":
            try:
                item[field] = parse_decimal(value) if value not in (None, u"") else None
            except ValueError:
                item[field] = None
        elif field == "unit":
            item[field] = normalize_unit(u"{}".format(value or u""))
        else:
            item[field] = (u"{}".format(value or u"")).strip()
    return item


class PriceListDocument(object):
    """Listino comune in formato JSON:

        {"format": "ESA_MEPQTO_PriceList", "version": 1, "name": ..., "currency": "EUR",
         "updated_by": ..., "updated_at": ...,
         "items": {codice: {chapter, subchapter, epu_item, price_book,
                            short_description, description, unit, price}}}

    Salvataggio concorrente come per il file di progetto: si salvano solo i codici
    toccati (aggiunti, cambiati, cancellati); se il file su disco e' cambiato dopo il
    caricamento, lo si rilegge e le modifiche locali si riapplicano sopra.
    """

    def __init__(self, path=None):
        self.path = path
        self.name = u""
        self.currency = u"EUR"
        self.items = {}
        self.updated_by = u""
        self.updated_at = u""
        self._loaded_mtime = None

    def load(self):
        self.items = {}
        self._loaded_mtime = None
        if not self.path or not os.path.isfile(self.path):
            return
        try:
            data = _read_json(self.path)
        except Exception as error:
            raise PriceListError(u"The price list is not valid JSON:\n{}\n\n{}".format(
                self.path, error))
        if data.get("format") not in (None, PRICE_LIST_FORMAT):
            raise PriceListError(u"Not a MEP QTO price list (format '{}'):\n{}".format(
                data.get("format"), self.path))
        self._apply(data)
        self._loaded_mtime = _mtime(self.path)

    def _apply(self, data):
        self.name = u"{}".format(data.get("name") or u"")
        self.currency = u"{}".format(data.get("currency") or u"EUR")
        self.updated_by = u"{}".format(data.get("updated_by") or u"")
        self.updated_at = u"{}".format(data.get("updated_at") or u"")
        items = data.get("items") or {}
        self.items = dict((u"{}".format(code).strip(), clean_item(values))
                          for code, values in items.items()
                          if u"{}".format(code).strip())

    def changed_since_load(self):
        current = _mtime(self.path) if self.path else None
        return current is not None and current != self._loaded_mtime

    def save(self, items, changed_codes, deleted_codes, name, currency, user_name,
             path=None):
        """Scrive il listino; True se le modifiche sono state fuse con quelle di un altro.

        items: tutte le voci correnti. changed_codes / deleted_codes: i codici toccati
        dall'utente, gli unici che si riapplicano se il file e' cambiato nel frattempo.
        path: nuovo percorso (Save As); in quel caso si scrive senza fondere.
        """
        merged = False
        if path is not None and path != self.path:
            self.path = path
            self._loaded_mtime = None
            final = dict(items)
        elif self.changed_since_load():
            disk = PriceListDocument(self.path)
            disk.load()
            final = disk.items
            for code in deleted_codes:
                final.pop(code, None)
            for code in changed_codes:
                if code in items:
                    final[code] = items[code]
            merged = True
        else:
            final = dict(items)

        self.items = dict((code, clean_item(values)) for code, values in final.items())
        self.name = name or u""
        self.currency = currency or u"EUR"
        self.updated_by = user_name or u""
        self.updated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        data = {
            "format": PRICE_LIST_FORMAT,
            "version": PRICE_LIST_VERSION,
            "name": self.name,
            "currency": self.currency,
            "updated_by": self.updated_by,
            "updated_at": self.updated_at,
            "items": self.items,
        }
        try:
            write_json_file(self.path, data)
        except (IOError, OSError) as error:
            raise PriceListError(u"Cannot write the price list:\n{}\n\n{}".format(
                self.path, error))
        self._loaded_mtime = _mtime(self.path)
        return merged


# =============================================================================
# FILE DI PROGETTO
# =============================================================================

def _is_local_path(path):
    if not path or "://" in path:
        return False
    folder = os.path.dirname(path)
    return bool(folder) and os.path.isdir(folder)


def default_project_file(doc):
    """<Modello>_MEPQTO.json accanto al centrale (o al file, se non condiviso).

    None per modelli cloud o mai salvati: in quel caso il percorso lo sceglie
    l'utente e lo si ricorda nella config.
    """
    path = None
    try:
        if doc.IsWorkshared:
            central = doc.GetWorksharingCentralModelPath()
            if central is not None:
                path = DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(central)
    except Exception:
        path = None
    if not _is_local_path(path):
        path = doc.PathName
    if not _is_local_path(path):
        return None
    folder, name = os.path.split(path)
    return os.path.join(folder, os.path.splitext(name)[0] + PROJECT_FILE_SUFFIX)


def document_key(doc):
    """Chiave stabile del documento per ricordare il file di progetto nella config."""
    try:
        if doc.IsWorkshared:
            central = doc.GetWorksharingCentralModelPath()
            if central is not None:
                return DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(central)
    except Exception:
        pass
    return doc.PathName or doc.Title


class ProjectStoreError(Exception):
    pass


def _mtime(path):
    try:
        return os.path.getmtime(path)
    except OSError:
        return None


def write_json_file(path, data):
    """Scrive data come JSON indentato, tramite un file temporaneo.

    ensure_ascii=True in IronPython 2.7 prova a decodificare come UTF-8 le stringhe (che
    sono gia' unicode) e fallisce sulle lettere accentate: si serializza senza escape e
    si fa l'escape \\uXXXX a mano. Solleva IOError / OSError.
    """
    text = json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True)
    text = u"".join(c if ord(c) < 128 else u"\\u{:04x}".format(ord(c)) for c in text)
    temp_path = path + ".tmp"
    with io.open(temp_path, "w", encoding="utf-8") as handle:
        handle.write(unicode(text))  # noqa: F821 - IronPython 2.7
    if os.path.exists(path):
        os.remove(path)
    os.rename(temp_path, path)


def _read_json(path):
    with io.open(path, "r", encoding="utf-8-sig") as handle:
        data = json.loads(handle.read())
    if not isinstance(data, dict):
        raise ValueError("not a JSON object")
    return data


def _read_overrides(data):
    """{(Type Mark, codice): frazione} dalla lista del file di progetto; le righe non
    valide si saltano."""
    overrides = {}
    if not isinstance(data, list):
        return overrides
    for entry in data:
        if not isinstance(entry, dict):
            continue
        type_mark = u"{}".format(entry.get("type_mark") or u"").strip()
        code = u"{}".format(entry.get("code") or u"").strip()
        try:
            fraction = float(entry.get("allowance"))
        except (TypeError, ValueError):
            continue
        if type_mark and code and fraction >= 0:
            overrides[(type_mark, code)] = fraction
    return overrides


def _write_overrides(overrides):
    """Lista ordinata per il JSON: le chiavi del dizionario sono tuple."""
    return [{"type_mark": type_mark, "code": code, "allowance": fraction}
            for (type_mark, code), fraction in sorted(overrides.items())]


class ProjectStore(object):
    """Override di progetto: voci, regole di misura, livelli WBS, percorso del listino e
    override della maggiorazione sulle voci del computo."""

    def __init__(self, path):
        self.path = path
        self.items = {}
        self.price_list_path = None
        # rules: dizionario di mepqto_rules.Rules.to_dict(), None = valori di default
        self.rules = None
        # wbs: nomi dei parametri dei livelli WBS (stringhe vuote per i livelli spenti)
        self.wbs = []
        # parameters: dizionario di mepqto_model.ParameterMap.to_dict(), None = default
        self.parameters = None
        # allowance_overrides: {(Type Mark, codice): frazione} che sostituisce la
        # maggiorazione della categoria su quella voce del computo
        self.allowance_overrides = {}
        self.updated_by = u""
        self.updated_at = u""
        self._loaded_mtime = None
        self._dirty_fields = set()
        self._dirty_overrides = set()
        self._dirty_price_list = False
        self._dirty_rules = False
        self._dirty_wbs = False
        self._dirty_parameters = False

    # ------------------------------------------------------------ lettura

    def load(self):
        """Carica il file, se esiste. Un file mancante e' un progetto vuoto."""
        self.items = {}
        self.price_list_path = None
        self.rules = None
        self.wbs = []
        self.parameters = None
        self.allowance_overrides = {}
        self._loaded_mtime = None
        if not self.path or not os.path.isfile(self.path):
            return
        try:
            data = _read_json(self.path)
        except Exception as error:
            raise ProjectStoreError(
                "The project file is not valid JSON and was not loaded:\n{}\n\n{}".format(
                    self.path, error))
        self._apply_data(data)
        self._loaded_mtime = _mtime(self.path)

    def _apply_data(self, data):
        items = data.get("items") or {}
        self.items = dict((code, dict(values)) for code, values in items.items()
                          if isinstance(values, dict))
        self.price_list_path = data.get("price_list_path") or None
        rules = data.get("rules")
        self.rules = rules if isinstance(rules, dict) else None
        wbs = data.get("wbs")
        self.wbs = [u"{}".format(name or u"") for name in wbs] if isinstance(wbs, list) else []
        parameters = data.get("parameters")
        self.parameters = parameters if isinstance(parameters, dict) else None
        self.allowance_overrides = _read_overrides(data.get("allowance_overrides"))
        self.updated_by = data.get("updated_by") or u""
        self.updated_at = data.get("updated_at") or u""

    # ------------------------------------------------------------ modifiche

    @property
    def is_dirty(self):
        return bool(self._dirty_fields or self._dirty_price_list or self._dirty_rules or
                    self._dirty_wbs or self._dirty_parameters or self._dirty_overrides)

    def set_item_field(self, code, field, value):
        self.items.setdefault(code, {})[field] = value
        self._dirty_fields.add((code, field))

    def set_allowance_override(self, type_mark, code, fraction):
        """Override della maggiorazione di una voce del computo; None lo toglie.
        Come i campi delle voci, si salva voce per voce."""
        key = (type_mark, code)
        if fraction is None:
            self.allowance_overrides.pop(key, None)
        else:
            self.allowance_overrides[key] = fraction
        self._dirty_overrides.add(key)

    def set_rules(self, rules_dict):
        """Le regole si salvano in blocco: su un salvataggio concorrente vince l'ultimo."""
        self.rules = rules_dict
        self._dirty_rules = True

    def set_wbs(self, names):
        """I livelli WBS si salvano in blocco, come le regole."""
        self.wbs = list(names)
        self._dirty_wbs = True

    def set_parameters(self, parameters_dict):
        """La mappa dei parametri si salva in blocco, come WBS e regole."""
        self.parameters = parameters_dict
        self._dirty_parameters = True

    def set_price_list_path(self, path):
        if path != self.price_list_path:
            self.price_list_path = path
            self._dirty_price_list = True

    # ------------------------------------------------------------ scrittura

    def _merge_with_disk(self):
        """Se qualcun altro ha salvato dopo il nostro caricamento, riparte dal disco
        e riapplica solo le modifiche locali."""
        current = _mtime(self.path)
        if current is None or current == self._loaded_mtime:
            return False
        try:
            data = _read_json(self.path)
        except Exception:
            return False

        local_items = self.items
        local_price_list = self.price_list_path
        local_rules = self.rules
        local_wbs = self.wbs
        local_parameters = self.parameters
        local_overrides = self.allowance_overrides
        self._apply_data(data)

        for code, field in self._dirty_fields:
            value = local_items.get(code, {}).get(field)
            self.items.setdefault(code, {})[field] = value
        for key in self._dirty_overrides:
            if key in local_overrides:
                self.allowance_overrides[key] = local_overrides[key]
            else:
                self.allowance_overrides.pop(key, None)
        if self._dirty_price_list:
            self.price_list_path = local_price_list
        if self._dirty_rules:
            self.rules = local_rules
        if self._dirty_wbs:
            self.wbs = local_wbs
        if self._dirty_parameters:
            self.parameters = local_parameters
        return True

    def save(self, user_name):
        """Scrive il file; True se nel frattempo era stato modificato da altri."""
        if not self.path:
            raise ProjectStoreError("No project file selected.")
        folder = os.path.dirname(self.path)
        if folder and not os.path.isdir(folder):
            raise ProjectStoreError("The folder does not exist:\n{}".format(folder))

        merged = self._merge_with_disk()
        self.updated_by = user_name or u""
        self.updated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Si salvano solo le voci con almeno un campo valorizzato.
        items = {}
        for code, values in self.items.items():
            kept = dict((field, values[field]) for field in FIELDS
                        if values.get(field) not in (None, u""))
            if kept:
                items[code] = kept

        data = {
            "version": FILE_VERSION,
            "price_list_path": self.price_list_path,
            "items": items,
            "rules": self.rules,
            "wbs": self.wbs,
            "parameters": self.parameters,
            "allowance_overrides": _write_overrides(self.allowance_overrides),
            "updated_by": self.updated_by,
            "updated_at": self.updated_at,
        }
        try:
            write_json_file(self.path, data)
        except (IOError, OSError) as error:
            raise ProjectStoreError("Cannot write the project file:\n{}\n\n{}".format(
                self.path, error))

        self._loaded_mtime = _mtime(self.path)
        self._dirty_fields.clear()
        self._dirty_overrides.clear()
        self._dirty_price_list = False
        self._dirty_rules = False
        self._dirty_wbs = False
        self._dirty_parameters = False
        return merged
