# -*- coding: utf-8 -*-
"""Export Excel del report di Openings: writer .xlsx minimale (zip + xml, niente Excel installato), ridotto da
mepqto_xlsx.py (MEPQTO): solo testo inline, intestazione, grassetto, a capo, riquadro bloccato e filtro automatico.
Nessun accesso a Revit: scrive i dati gia' pronti del report (vedi report_data() in openings_report).
Una sola copia, in openings_modules."""
import re

import clr
clr.AddReference("System.IO.Compression")
clr.AddReference("System.IO.Compression.FileSystem")

from System.IO import File, IOException, StreamWriter
from System.IO.Compression import ZipFile, ZipArchiveMode, CompressionLevel
from System.Text import UTF8Encoding

STYLE_DEFAULT, STYLE_HEADER, STYLE_BOLD, STYLE_TITLE, STYLE_WRAP = 0, 1, 2, 3, 4
_INVALID_XML = re.compile(u"[\\x00-\\x08\\x0b\\x0c\\x0e-\\x1f]")
MAX_CELL_TEXT = 32000
REPORT_COLUMNS = ("Status", "Opening", "Description", "Element", "Category", "Element status", "Notes")
REPORT_WIDTHS = (18, 12, 60, 12, 14, 34, 70)


class FileLockedError(Exception):
    """Il file esiste ed e' aperto (in Excel): non si puo' sovrascrivere."""


class Sheet(object):
    def __init__(self, name):
        self.name = name[:31]
        self.rows = []  # [[(valore, stile)]]
        self.widths = []
        self.freeze_rows = 0
        self.filter_columns = 0  # filtro automatico sulla riga 1


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
    return text.replace(u"&", u"&amp;").replace(u"<", u"&lt;").replace(u">", u"&gt;").replace(u'"', u"&quot;")


def _cell_xml(ref, value, style):
    style_text = u' s="{}"'.format(style) if style else u""
    if value is None or value == u"":
        return u'<c r="{}"{}/>'.format(ref, style_text) if style else u""
    return u'<c r="{}"{} t="inlineStr"><is><t xml:space="preserve">{}</t></is></c>'.format(ref, style_text,
                                                                                          _escape(value))


def _sheet_xml(sheet):
    parts = [u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             u'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             u'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
    if sheet.freeze_rows:
        parts.append(u'<sheetViews><sheetView workbookViewId="0"><pane ySplit="{0}" topLeftCell="A{1}" '
                     u'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'.format(
                         sheet.freeze_rows, sheet.freeze_rows + 1))
    if sheet.widths:
        parts.append(u"<cols>")
        for index, width in enumerate(sheet.widths):
            parts.append(u'<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(index + 1, width))
        parts.append(u"</cols>")
    parts.append(u"<sheetData>")
    for row_index, cells in enumerate(sheet.rows):
        xml = [_cell_xml(u"{}{}".format(column_letter(i), row_index + 1), value, style)
               for i, (value, style) in enumerate(cells)]
        parts.append(u'<row r="{}">{}</row>'.format(row_index + 1, u"".join(x for x in xml if x)))
    parts.append(u"</sheetData>")
    if sheet.filter_columns and len(sheet.rows) > 1:
        parts.append(u'<autoFilter ref="A1:{}{}"/>'.format(column_letter(sheet.filter_columns - 1), len(sheet.rows)))
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

# stili: 0 normale, 1 intestazione (grassetto su fondo azzurro, bordo sotto), 2 grassetto, 3 titolo, 4 a capo
_STYLES = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    u'<fonts count="3">'
    u'<font><sz val="10"/><name val="Arial"/></font>'
    u'<font><b/><sz val="10"/><name val="Arial"/></font>'
    u'<font><b/><sz val="13"/><name val="Arial"/></font>'
    u'</fonts>'
    u'<fills count="3">'
    u'<fill><patternFill patternType="none"/></fill>'
    u'<fill><patternFill patternType="gray125"/></fill>'
    u'<fill><patternFill patternType="solid"><fgColor rgb="FFD9E2EC"/><bgColor indexed="64"/></patternFill></fill>'
    u'</fills>'
    u'<borders count="2">'
    u'<border><left/><right/><top/><bottom/><diagonal/></border>'
    u'<border><left/><right/><top/><bottom style="thin"><color rgb="FF808080"/></bottom><diagonal/></border>'
    u'</borders>'
    u'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    u'<cellXfs count="5">'
    u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    u'<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
    u'<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    u'<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
    u'<alignment wrapText="1" vertical="top"/></xf>'
    u'</cellXfs>'
    u'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    u'</styleSheet>')


def _workbook_xml(sheets):
    entries = u"".join(u'<sheet name="{}" sheetId="{}" r:id="rId{}"/>'.format(_escape(s.name), i + 1, i + 1)
                       for i, s in enumerate(sheets))
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
            u'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(i + 1)
            for i in range(len(sheets)))
        _write_entry(archive, "[Content_Types].xml", _CONTENT_TYPES.format(sheets=overrides))
        _write_entry(archive, "_rels/.rels", _ROOT_RELS)
        _write_entry(archive, "xl/workbook.xml", _workbook_xml(sheets))
        _write_entry(archive, "xl/_rels/workbook.xml.rels", _workbook_rels(sheets))
        _write_entry(archive, "xl/styles.xml", _STYLES)
        for i, sheet in enumerate(sheets):
            _write_entry(archive, "xl/worksheets/sheet{}.xml".format(i + 1), _sheet_xml(sheet))
    finally:
        archive.Dispose()


def write_report(path, data):
    """Due fogli: Report (una riga per elemento di ogni foro, il foro ripetuto; filtro automatico) e Summary
    (totali, avvisi, dettagli e azioni). data: vedi report_data() nello script."""
    report = Sheet("Report")
    report.widths = list(REPORT_WIDTHS)
    report.freeze_rows = 1
    report.filter_columns = len(REPORT_COLUMNS)
    report.rows.append([(name, STYLE_HEADER) for name in REPORT_COLUMNS])
    for section in data["sections"]:
        for item in section["items"]:
            notes = u"; ".join(item["notes"])
            head = [(section["title"], STYLE_DEFAULT), (item["code"], STYLE_DEFAULT),
                    (item["description"], STYLE_DEFAULT)]
            if not item["rows"]:
                report.rows.append(head + [(u"", 0), (u"", 0), (u"", 0), (notes, STYLE_WRAP)])
            for element, category, state in item["rows"]:
                report.rows.append(head + [(element, 0), (category, 0), (state, 0), (notes, STYLE_WRAP)])
    summary = Sheet("Summary")
    summary.widths = [32, 90]
    summary.rows.append([(u"Openings {}".format(data["title"]), STYLE_TITLE)])
    summary.rows.append([(data.get("date", u""), STYLE_DEFAULT)])
    summary.rows.append([])
    for row in data["summary"]:
        if row:
            summary.rows.append([(row[0], STYLE_BOLD), (u"{}".format(row[1]), STYLE_DEFAULT)])
    for line in ([data["warning"]] if data["warning"] else []) + data["details"] + data["log"]:
        summary.rows.append([(u"", 0), (line, STYLE_WRAP)])
    write_workbook(path, [report, summary])
