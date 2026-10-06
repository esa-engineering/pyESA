# -*- coding: utf-8 -*-
"""
mepqto_xlsx.py - writer .xlsx minimale (zip + xml), senza Excel installato.

Scrive solo cio' che serve al computo: celle di testo inline (niente
sharedStrings), numeri, formule con il valore gia' calcolato in cache, pochi
stili fissi, larghezza colonne, riquadro bloccato e filtro automatico.
fullCalcOnLoad fa ricalcolare le formule a Excel all'apertura.
"""

import re

import clr
clr.AddReference('System.IO.Compression')
clr.AddReference('System.IO.Compression.FileSystem')

from System.IO import File, IOException, StreamWriter
from System.IO.Compression import ZipFile, ZipArchiveMode, CompressionLevel
from System.Text import UTF8Encoding

import mepqto_model as qm
import mepqto_rules as qr

# Indici di cellXfs in styles.xml
STYLE_DEFAULT = 0
STYLE_HEADER = 1
STYLE_NUMBER = 2
STYLE_MONEY = 3
STYLE_BOLD = 4
STYLE_BOLD_NUMBER = 5
STYLE_BOLD_MONEY = 6
STYLE_TITLE = 7
STYLE_WRAP = 8
# quantita' di una voce con override della maggiorazione (stesso colore della scheda)
STYLE_NUMBER_OVERRIDE = 9
# riga di gruppo del riepilogo Type Mark: grassetto su fondo chiaro, a capo automatico
STYLE_GROUP = 10

_INVALID_XML = re.compile(u"[\\x00-\\x08\\x0b\\x0c\\x0e-\\x1f]")
MAX_CELL_TEXT = 32000


class FileLockedError(Exception):
    pass


class Cell(object):
    """Valore (testo o numero), stile e formula opzionale (senza '=')."""

    __slots__ = ("value", "style", "formula")

    def __init__(self, value=None, style=STYLE_DEFAULT, formula=None):
        self.value = value
        self.style = style
        self.formula = formula


class Sheet(object):
    def __init__(self, name):
        self.name = name[:31]
        self.rows = []
        self.widths = []
        self.freeze_rows = 0
        self.filter_row = None   # numero di riga (1-based) dell'intestazione filtrabile
        self.filter_columns = 0

    def add_row(self, cells=None):
        self.rows.append(list(cells or []))
        return len(self.rows)   # numero di riga Excel della riga appena aggiunta


def column_letter(index):
    """0 -> A, 25 -> Z, 26 -> AA."""
    letters = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _escape(text):
    text = _INVALID_XML.sub(u"", u"{}".format(text))[:MAX_CELL_TEXT]
    return text.replace(u"&", u"&amp;").replace(u"<", u"&lt;")\
        .replace(u">", u"&gt;").replace(u'"', u"&quot;")


def _number(value):
    return repr(float(value))


def _cell_xml(ref, cell):
    if not isinstance(cell, Cell):
        cell = Cell(cell)
    style = u' s="{}"'.format(cell.style) if cell.style else u""
    value = cell.value
    if cell.formula:
        cached = u"<v>{}</v>".format(_number(value)) if isinstance(value, (int, float)) else u""
        return u'<c r="{}"{}><f>{}</f>{}</c>'.format(ref, style, _escape(cell.formula), cached)
    if value is None or value == u"":
        return u'<c r="{}"{}/>'.format(ref, style) if style else u""
    if isinstance(value, bool):
        value = u"Yes" if value else u"No"
    if isinstance(value, (int, float)):
        return u'<c r="{}"{}><v>{}</v></c>'.format(ref, style, _number(value))
    return u'<c r="{}"{} t="inlineStr"><is><t xml:space="preserve">{}</t></is></c>'.format(
        ref, style, _escape(value))


def _sheet_xml(sheet):
    parts = [u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             u'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             u'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
    if sheet.freeze_rows:
        top = u"A{}".format(sheet.freeze_rows + 1)
        parts.append(u'<sheetViews><sheetView workbookViewId="0">'
                     u'<pane ySplit="{}" topLeftCell="{}" activePane="bottomLeft" state="frozen"/>'
                     u'</sheetView></sheetViews>'.format(sheet.freeze_rows, top))
    if sheet.widths:
        parts.append(u"<cols>")
        for index, width in enumerate(sheet.widths):
            parts.append(u'<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(
                index + 1, width))
        parts.append(u"</cols>")

    parts.append(u"<sheetData>")
    for row_index, cells in enumerate(sheet.rows):
        row_number = row_index + 1
        cell_parts = []
        for col_index, cell in enumerate(cells):
            xml = _cell_xml(u"{}{}".format(column_letter(col_index), row_number), cell)
            if xml:
                cell_parts.append(xml)
        parts.append(u'<row r="{}">{}</row>'.format(row_number, u"".join(cell_parts)))
    parts.append(u"</sheetData>")

    if sheet.filter_row and sheet.filter_columns and len(sheet.rows) >= sheet.filter_row:
        parts.append(u'<autoFilter ref="A{0}:{1}{2}"/>'.format(
            sheet.filter_row, column_letter(sheet.filter_columns - 1), len(sheet.rows)))
    parts.append(u"</worksheet>")
    return u"".join(parts)


_CONTENT_TYPES = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    u'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    u'<Default Extension="xml" ContentType="application/xml"/>'
    u'<Override PartName="/xl/workbook.xml" '
    u'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    u'<Override PartName="/xl/styles.xml" '
    u'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    u'{sheets}</Types>')

_ROOT_RELS = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    u'<Relationship Id="rId1" '
    u'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    u'Target="xl/workbook.xml"/></Relationships>')

# numFmt 164: quantita', 165: importi in euro.
_STYLES = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    u'<numFmts count="2">'
    u'<numFmt numFmtId="164" formatCode="#,##0.00"/>'
    u'<numFmt numFmtId="165" formatCode="&quot;\u20ac&quot; #,##0.00"/>'
    u'</numFmts>'
    u'<fonts count="3">'
    u'<font><sz val="10"/><name val="Arial"/></font>'
    u'<font><b/><sz val="10"/><name val="Arial"/></font>'
    u'<font><b/><sz val="13"/><name val="Arial"/></font>'
    u'</fonts>'
    u'<fills count="5">'
    u'<fill><patternFill patternType="none"/></fill>'
    u'<fill><patternFill patternType="gray125"/></fill>'
    u'<fill><patternFill patternType="solid"><fgColor rgb="FFD9E2EC"/><bgColor indexed="64"/></patternFill></fill>'
    u'<fill><patternFill patternType="solid"><fgColor rgb="FFFFE3A3"/><bgColor indexed="64"/></patternFill></fill>'
    u'<fill><patternFill patternType="solid"><fgColor rgb="FFEEF3F8"/><bgColor indexed="64"/></patternFill></fill>'
    u'</fills>'
    u'<borders count="2">'
    u'<border><left/><right/><top/><bottom/><diagonal/></border>'
    u'<border><left/><right/><top/><bottom style="thin"><color rgb="FF808080"/></bottom><diagonal/></border>'
    u'</borders>'
    u'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    u'<cellXfs count="11">'
    u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    u'<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
    u'<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    u'<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    u'<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    u'<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>'
    u'<xf numFmtId="165" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>'
    u'<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
    u'<alignment wrapText="1" vertical="top"/></xf>'
    u'<xf numFmtId="164" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>'
    u'<xf numFmtId="0" fontId="1" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1" '
    u'applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>'
    u'</cellXfs>'
    u'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    u'</styleSheet>')


def _workbook_xml(sheets):
    entries = u"".join(
        u'<sheet name="{}" sheetId="{}" r:id="rId{}"/>'.format(_escape(sheet.name), i + 1, i + 1)
        for i, sheet in enumerate(sheets))
    return (u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            u'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            u'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            u'<sheets>{}</sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>'.format(entries))


def _workbook_rels(sheets):
    entries = [u'<Relationship Id="rId{0}" '
               u'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
               u'Target="worksheets/sheet{0}.xml"/>'.format(i + 1) for i in range(len(sheets))]
    entries.append(u'<Relationship Id="rId{}" '
                   u'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
                   u'Target="styles.xml"/>'.format(len(sheets) + 1))
    return (u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            u'{}</Relationships>'.format(u"".join(entries)))


def _write_entry(archive, name, text):
    entry = archive.CreateEntry(name, CompressionLevel.Optimal)
    writer = StreamWriter(entry.Open(), UTF8Encoding(False))
    try:
        writer.Write(text)
    finally:
        writer.Close()


def write_workbook(path, sheets):
    """Scrive il workbook. FileLockedError se il file esiste ed e' aperto in Excel."""
    try:
        if File.Exists(path):
            File.Delete(path)
    except IOException:
        raise FileLockedError(path)

    archive = ZipFile.Open(path, ZipArchiveMode.Create)
    try:
        overrides = u"".join(
            u'<Override PartName="/xl/worksheets/sheet{}.xml" '
            u'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            .format(i + 1) for i in range(len(sheets)))
        _write_entry(archive, "[Content_Types].xml", _CONTENT_TYPES.format(sheets=overrides))
        _write_entry(archive, "_rels/.rels", _ROOT_RELS)
        _write_entry(archive, "xl/workbook.xml", _workbook_xml(sheets))
        _write_entry(archive, "xl/_rels/workbook.xml.rels", _workbook_rels(sheets))
        _write_entry(archive, "xl/styles.xml", _STYLES)
        for i, sheet in enumerate(sheets):
            _write_entry(archive, "xl/worksheets/sheet{}.xml".format(i + 1), _sheet_xml(sheet))
    finally:
        archive.Dispose()


# =============================================================================
# CONTENUTO DEL COMPUTO
# =============================================================================

MAX_IDS_IN_CELL = 500


def _ids_text(element_ids, id_value):
    values = [u"{}".format(id_value(eid)) for eid in element_ids[:MAX_IDS_IN_CELL]]
    text = u", ".join(values)
    if len(element_ids) > MAX_IDS_IN_CELL:
        text += u", ... (+{})".format(len(element_ids) - MAX_IDS_IN_CELL)
    return text


def _header(sheet, headers):
    sheet.add_row([Cell(h, STYLE_HEADER) for h in headers])
    sheet.freeze_rows = len(sheet.rows)
    sheet.filter_row = len(sheet.rows)
    sheet.filter_columns = len(headers)


def _price_list_sheet(session):
    # Stesso ordine dell'EPU Excel letto dal tool (A codice ... E prezzo, poi F..I):
    # l'export si puo' usare di nuovo come listino.
    sheet = Sheet("EPU")
    sheet.widths = [18, 36, 80, 8, 14, 18, 18, 14, 24]
    _header(sheet, (u"Price book code", u"Short Description", u"Description", u"Unit",
                    u"Unit price", u"Chapter", u"Subchapter", u"EPU item No.",
                    u"Reference price book"))
    for code in session.price_codes:
        item = session.items[code]
        sheet.add_row([code, Cell(item.short_description, STYLE_WRAP),
                       Cell(item.description, STYLE_WRAP), item.unit,
                       Cell(item.price, STYLE_MONEY), item.chapter, item.subchapter,
                       item.epu_item, item.price_book])
    return sheet


def _bill_sheet(session):
    sheet = Sheet("Bill of quantities")
    labels = list(session.wbs_labels or [])
    wbs_count = len(labels)
    sheet.widths = [16] * wbs_count + [14, 12, 18, 80, 8, 12, 14, 16]
    sheet.add_row([Cell(u"MEP Bill of Quantities - {}".format(session.model_name), STYLE_TITLE)])
    sheet.add_row([u"Phase: {}    Categories: {}    Generated: {}".format(
        session.phase_label, session.categories_label, session.generated_at)])
    sheet.add_row([u"Models: {}    Worksets not read: {}".format(
        getattr(session, "models_label", u"") or session.model_name,
        getattr(session, "worksets_label", u"") or u"none")])
    sheet.add_row([u"Price list: {}    Project file: {}".format(
        session.price_list_path or u"(none)", session.project_file or u"(none)")])
    overridden = session.bill.overridden if session.bill is not None else {}
    if overridden:
        sheet.add_row([Cell(u"Highlighted quantities have an allowance of their own "
                            u"(see the Rules sheet).", STYLE_NUMBER_OVERRIDE)])
    sheet.add_row([])
    # Come nella scheda: le colonne WBS in testa, una riga di totale per combinazione
    # (SUBTOTAL sulle sue voci) e le voci con i valori WBS ripetuti, cosi' il foglio si
    # filtra o si mette in pivot senza ricostruire la WBS.
    headers = labels + [u"Type Mark", u"EPU item No.", u"Price book code", u"Description",
                        u"Unit", u"Quantity", u"Unit price", u"Amount"]
    sheet.add_row([Cell(h, STYLE_HEADER) for h in headers])
    sheet.freeze_rows = len(sheet.rows)
    first = len(sheet.rows) + 1

    quantity_col = column_letter(wbs_count + 5)
    price_col = column_letter(wbs_count + 6)
    amount_col = column_letter(wbs_count + 7)

    outline = qm.bill_outline(session.bill, labels)
    rows = [first + index for index in range(len(outline))]
    amounts = []
    for entry in outline:
        item = session.items.get(entry.code) if entry.kind == "item" else None
        amounts.append(entry.quantity * item.price
                       if item is not None and item.price is not None else 0.0)

    for index, entry in enumerate(outline):
        row = rows[index]
        wbs = qm.wbs_cells(entry.key, wbs_count)
        if entry.kind == "group":
            end = qm.outline_group_end(outline, index)
            subtotal = sum(amounts[index + 1:end + 1])
            formula = u"SUBTOTAL(9,{0}{1}:{0}{2})".format(amount_col, row + 1, rows[end]) \
                if end > index else None
            sheet.add_row([Cell(value, STYLE_BOLD) for value in wbs] + [
                None, None, None, Cell(entry.text, STYLE_BOLD), None, None, None,
                Cell(subtotal, STYLE_BOLD_MONEY, formula)])
            continue
        item = session.items[entry.code]
        sheet.add_row(wbs + [
            entry.type_mark, item.epu_item, entry.code,
            Cell(item.description or u"", STYLE_WRAP),
            session.bill_units.get(entry.code, item.unit),
            Cell(entry.quantity, STYLE_NUMBER_OVERRIDE
                 if (entry.type_mark, entry.code) in overridden else STYLE_NUMBER),
            Cell(item.price, STYLE_MONEY),
            Cell(amounts[index], STYLE_MONEY,
                 u"{0}{2}*{1}{2}".format(quantity_col, price_col, row)),
        ])

    last = max(len(sheet.rows), first)
    sheet.add_row([])
    sheet.add_row([None] * wbs_count + [
        None, None, None, Cell(u"TOTAL", STYLE_BOLD), None, None, None,
        Cell(sum(amounts), STYLE_BOLD_MONEY,
             u"SUBTOTAL(9,{0}{1}:{0}{2})".format(amount_col, first, last))])
    return sheet


def _type_marks_sheet(session):
    """Come nella scheda: una riga di gruppo per tipo, poi una riga per codice."""
    sheet = Sheet("Type Marks")
    sheet.widths = [24, 16, 40, 9, 24, 12, 20, 80]
    headers = (u"Category", u"Type Mark", u"Family and Type", u"Nested", u"Model",
               u"Code slot", u"Price book code", u"Description")
    _header(sheet, headers)
    for row in session.type_rows:
        # Tutte le celle con lo stile del gruppo, cosi' il fondo copre la riga intera.
        sheet.add_row([Cell(value, STYLE_GROUP) for value in (
            row.category, row.type_mark, row.type_label,
            u"Yes" if row.nested else u"No", getattr(row, "model", u""), None, None, None)])
        for label, code in row.code_entries():
            item = session.items.get(code)
            sheet.add_row([None, None, None, None, None, label, code,
                           Cell(item.description if item else u"", STYLE_WRAP)])
    return sheet


def _issues_sheet(session, id_value):
    sheet = Sheet("Issues")
    sheet.widths = [28, 24, 80, 11, 60]
    _header(sheet, (u"Issue", u"Subject", u"Detail", u"Instances", u"Element Ids"))
    for issue in session.issues:
        sheet.add_row([
            issue.kind, issue.subject, Cell(issue.detail, STYLE_WRAP), issue.count,
            Cell(_ids_text(issue.element_ids, id_value), STYLE_WRAP),
        ])
    return sheet


def _rules_sheet(session, category_label, linear_keys):
    """Regole usate per il computo: maggiorazioni, kg/mq dei canali, densita' dei tubi."""
    rules = session.rules
    sheet = Sheet("Rules")
    sheet.widths = [44, 22, 14, 40]
    sheet.add_row([Cell(u"Allowance for fittings and waste", STYLE_BOLD)])
    sheet.add_row([Cell(u"Category", STYLE_HEADER), Cell(u"Allowance %", STYLE_HEADER)])
    for key in linear_keys:
        sheet.add_row([category_label(key), Cell(rules.allowance_for(key) * 100.0, STYLE_NUMBER)])
    sheet.add_row([])

    # Override sulle singole voci: solo quelli applicati a voci presenti nel computo.
    bill = getattr(session, "bill", None)
    overridden = bill.overridden if bill is not None else {}
    if overridden:
        sheet.add_row([Cell(u"Allowance overrides on bill items", STYLE_BOLD)])
        sheet.add_row([Cell(u"Type Mark", STYLE_HEADER), Cell(u"Price book code", STYLE_HEADER),
                       Cell(u"Allowance %", STYLE_HEADER), Cell(u"Category allowance", STYLE_HEADER)])
        for line_key in sorted(overridden):
            categories = u", ".join(
                u"{} {:g}%".format(category_label(key), rules.allowance_for(key) * 100.0)
                for key in bill.line_categories.get(line_key, []))
            sheet.add_row([line_key[0], line_key[1],
                           Cell(overridden[line_key] * 100.0, STYLE_NUMBER_OVERRIDE), categories])
        sheet.add_row([])
    sheet.add_row([Cell(u"Duct sheet weight", STYLE_BOLD)])
    sheet.add_row([Cell(u"Shape", STYLE_HEADER), Cell(u"Largest side below [mm]", STYLE_HEADER),
                   Cell(u"kg/mq", STYLE_HEADER)])
    for shape in sorted(rules.duct_weight):
        for limit, weight in rules.duct_weight[shape]:
            sheet.add_row([qr.SHAPE_LABELS.get(shape, shape), u"no limit" if limit is None else limit,
                           Cell(weight, STYLE_NUMBER)])
    sheet.add_row([])
    sheet.add_row([Cell(u"Pipe steel density", STYLE_BOLD)])
    sheet.add_row([Cell(u"Type Mark", STYLE_HEADER), Cell(u"kg/mc", STYLE_HEADER),
                   None, Cell(u"Note", STYLE_HEADER)])
    for mark, note, density in rules.pipe_density:
        sheet.add_row([mark, Cell(density, STYLE_NUMBER), None, note])

    param_map = getattr(session, "param_map", None)
    if param_map is not None:
        sheet.add_row([])
        sheet.add_row([Cell(u"Model parameters", STYLE_BOLD)])
        sheet.add_row([Cell(u"Use", STYLE_HEADER), Cell(u"Parameter", STYLE_HEADER)])
        for index, name in enumerate(param_map.piece_codes):
            if name:
                sheet.add_row([u"Price code {} - piece categories (type)".format(index + 1), name])
        for index, name in enumerate(param_map.linear_codes):
            if name:
                sheet.add_row([u"Price code {} - all categories (instance)".format(index + 1),
                               name])
        sheet.add_row([u"Include Yes/No - all categories (instance)",
                       param_map.include or u"(none)"])
    return sheet


def export_takeoff(path, session, id_value, category_label=None, linear_keys=()):
    """Stessi contenuti delle schede: elenco prezzi, computo, Type Mark, anomalie."""
    sheets = [_price_list_sheet(session), _bill_sheet(session), _type_marks_sheet(session)]
    if category_label is not None:
        sheets.append(_rules_sheet(session, category_label, linear_keys))
    sheets.append(_issues_sheet(session, id_value))
    write_workbook(path, sheets)
