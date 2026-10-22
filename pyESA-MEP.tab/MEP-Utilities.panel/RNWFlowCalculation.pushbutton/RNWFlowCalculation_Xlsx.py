# -*- coding: utf-8 -*-
"""
RNWFlowCalculation_Xlsx.py - writer .xlsx minimale (zip + xml), senza Excel installato, per
l'esportazione del calcolo di RNWFlowCalculation (export_report: un foglio per
scheda del report, RNWFlowCalculation_Report, piu' le impostazioni).

La parte generica (celle, fogli, zip) e' copiata da MEPQTO.pushbutton/
MEPQTO_xlsx.py, che importa i moduli di MEPQTO e quindi non si puo' usare da un
altro pulsante (gli helper sono duplicati per strumento, vedi CLAUDE.md). Gli
stili sono propri: intestazione, numeri a 0 / 1 / 2 decimali, titolo, testo a
capo, e i colori dell'esito (errore, avviso, OK, sovradimensionato) come nella
finestra.

Nessun import di Revit; compatibile con Rocket mode (nessuno stato di modulo).
"""

import re

import clr
clr.AddReference('System.IO.Compression')
clr.AddReference('System.IO.Compression.FileSystem')

from System.IO import File, IOException, StreamWriter
from System.IO.Compression import ZipFile, ZipArchiveMode, CompressionLevel
from System.Text import UTF8Encoding

import RNWFlowCalculation_Report as rp

# Indici di cellXfs in styles.xml
STYLE_DEFAULT = 0
STYLE_HEADER = 1
STYLE_NUM2 = 2
STYLE_NUM1 = 3
STYLE_NUM0 = 4
STYLE_TITLE = 5
STYLE_WRAP = 6
STYLE_BOLD = 7
STYLE_ERROR = 8
STYLE_WARNING = 9
STYLE_PASS = 10
STYLE_FIXTURE = 11
STYLE_INFO = 12

_INVALID_XML = re.compile(u"[\\x00-\\x08\\x0b\\x0c\\x0e-\\x1f]")
MAX_CELL_TEXT = 32000


class FileLockedError(Exception):
    pass


class Cell(object):
    """Valore (testo o numero) e stile."""

    __slots__ = ("value", "style")

    def __init__(self, value=None, style=STYLE_DEFAULT):
        self.value = value
        self.style = style


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
        return len(self.rows)


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

# numFmt 164: 0.00, 165: 0.0, 166: 0. Colori dell'esito come nella finestra.
_STYLES = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    u'<numFmts count="3">'
    u'<numFmt numFmtId="164" formatCode="0.00"/>'
    u'<numFmt numFmtId="165" formatCode="0.0"/>'
    u'<numFmt numFmtId="166" formatCode="0"/>'
    u'</numFmts>'
    u'<fonts count="7">'
    u'<font><sz val="10"/><name val="Arial"/></font>'
    u'<font><b/><sz val="10"/><name val="Arial"/></font>'
    u'<font><b/><sz val="13"/><name val="Arial"/></font>'
    u'<font><sz val="10"/><color rgb="FF9C0006"/><name val="Arial"/></font>'
    u'<font><sz val="10"/><color rgb="FF7A5B00"/><name val="Arial"/></font>'
    u'<font><i/><sz val="10"/><name val="Arial"/></font>'
    u'<font><sz val="10"/><color rgb="FF2D5A8A"/><name val="Arial"/></font>'
    u'</fonts>'
    u'<fills count="7">'
    u'<fill><patternFill patternType="none"/></fill>'
    u'<fill><patternFill patternType="gray125"/></fill>'
    u'<fill><patternFill patternType="solid"><fgColor rgb="FFD9E2EC"/><bgColor indexed="64"/></patternFill></fill>'
    u'<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>'
    u'<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/><bgColor indexed="64"/></patternFill></fill>'
    u'<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill>'
    u'<fill><patternFill patternType="solid"><fgColor rgb="FFEAF4EA"/><bgColor indexed="64"/></patternFill></fill>'
    u'</fills>'
    u'<borders count="2">'
    u'<border><left/><right/><top/><bottom/><diagonal/></border>'
    u'<border><left/><right/><top/><bottom style="thin"><color rgb="FF808080"/></bottom><diagonal/></border>'
    u'</borders>'
    u'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    u'<cellXfs count="13">'
    u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    u'<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" '
    u'applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="bottom"/></xf>'
    u'<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    u'<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    u'<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    u'<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
    u'<alignment wrapText="1" vertical="top"/></xf>'
    u'<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    u'<xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    u'<xf numFmtId="0" fontId="4" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    u'<xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>'
    u'<xf numFmtId="0" fontId="5" fillId="6" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    u'<xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
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
            u'<sheets>{}</sheets></workbook>'.format(entries))


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
# CONTENUTO DELL'ESPORTAZIONE
# =============================================================================

_NUMBER_STYLES = {0: STYLE_NUM0, 1: STYLE_NUM1, 2: STYLE_NUM2}
_SEVERITY_STYLES = {
    rp.SEV_ERROR: STYLE_ERROR,
    rp.SEV_WARNING: STYLE_WARNING,
    rp.SEV_PASS: STYLE_PASS,
    rp.SEV_INFO: STYLE_INFO,
}


def _table_cell(column, row):
    value = row.values.get(column.key)
    if isinstance(value, bool):
        value = u"Yes" if value else u""
    is_number = isinstance(value, (int, float))
    if is_number:
        style = _NUMBER_STYLES.get(column.digits, STYLE_NUM2)
    elif row.kind == rp.ROW_FIXTURE:
        style = STYLE_FIXTURE
    else:
        style = STYLE_DEFAULT
    severity = row.severities.get(column.key) if column.severity else None
    if severity in _SEVERITY_STYLES:
        style = _SEVERITY_STYLES[severity]
    return Cell(value, style)


def _add_table(sheet, table, filtered):
    """Intestazione e righe; con filtered la tabella e' l'unica del foglio:
    intestazione bloccata e filtro automatico."""
    if table.title and not filtered:
        sheet.add_row([Cell(table.title, STYLE_BOLD)])
    header_row = sheet.add_row([Cell(column.header, STYLE_HEADER) for column in table.columns])
    for row in table.rows:
        sheet.add_row([_table_cell(column, row) for column in table.columns])
    if filtered:
        sheet.freeze_rows = header_row
        sheet.filter_row = header_row
        sheet.filter_columns = len(table.columns)


def _tab_sheet(title, tab):
    sheet = Sheet(tab.name)
    sheet.add_row([Cell(title, STYLE_TITLE)])
    sheet.add_row([])
    single = len(tab.blocks) == 1 and isinstance(tab.blocks[0], rp.Table)
    widths = []
    for block in tab.blocks:
        if isinstance(block, rp.Pairs):
            if block.title:
                sheet.add_row([Cell(block.title, STYLE_BOLD)])
            for label, value in block.rows:
                is_number = isinstance(value, (int, float)) and not isinstance(value, bool)
                style = _NUMBER_STYLES.get(rp.PAIR_DIGITS, STYLE_NUM2) if is_number \
                    else STYLE_DEFAULT
                sheet.add_row([Cell(label), Cell(value, style)])
            widths = _merge_widths(widths, [36, 18])
        elif isinstance(block, rp.Text):
            # Senza a capo: il testo continua nelle celle vuote a destra
            sheet.add_row([Cell(block.text)])
        elif isinstance(block, rp.Table):
            _add_table(sheet, block, single)
            widths = _merge_widths(widths, [column.width for column in block.columns])
        sheet.add_row([])
    sheet.widths = widths
    return sheet


def _merge_widths(current, new):
    merged = list(current)
    for index, width in enumerate(new):
        if index < len(merged):
            merged[index] = max(merged[index], width)
        else:
            merged.append(width)
    return merged


def export_report(path, report):
    """Scrive il report (RNWFlowCalculation_Report.Report): un foglio per scheda
    (General, Sections, Fixtures) e il foglio Settings. FileLockedError se il
    file e' aperto in Excel."""
    sheets = [_tab_sheet(report.title, tab) for tab in report.tabs]

    settings = Sheet("Settings")
    settings.add_row([Cell(report.title, STYLE_TITLE)])
    settings.add_row([])
    for label, value in report.settings:
        settings.add_row([Cell(label, STYLE_BOLD), Cell(value, STYLE_WRAP)])
    settings.widths = [42, 110]
    sheets.append(settings)

    write_workbook(path, sheets)
