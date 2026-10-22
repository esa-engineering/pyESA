# -*- coding: utf-8 -*-
"""
RNWFlowCalculation_Pdf.py - writer PDF minimale per l'esportazione del calcolo di
RNWFlowCalculation (export_report), senza librerie esterne ne' stampanti.

Come RNWFlowCalculation_Xlsx per Excel: il file si scrive a mano. PDF 1.4 con i
tre font standard Helvetica (normale, grassetto, corsivo), che ogni lettore PDF
ha gia' e che non si incorporano; codifica WinAnsi (cp1252), quindi i caratteri
fuori da cp1252 diventano un sostituto (_REPLACE) o "?". Le larghezze dei
caratteri (_WIDTHS) sono quelle delle metriche Adobe di Helvetica e servono per
andare a capo nelle celle e dimensionare le colonne. Il testo si scrive come
stringa esadecimale, quindi il file e' tutto ASCII e gli offset della tabella
xref sono la lunghezza del testo.

Pagine orizzontali A4 o A3. Ogni scheda del report (General, Sections,
Fixtures) comincia su una pagina nuova; le tabelle vanno a capo nelle celle e
ripetono l'intestazione a ogni pagina. Le colonne prendono la larghezza del
contenuto; se non ci stanno si riduce il carattere (fino a TABLE_MIN_SIZE), poi
si stringono le colonne di testo, che vanno a capo.

Nessun import di Revit; nessuno stato di modulo (Rocket mode).
"""

import clr  # noqa: F401  (assicura il caricamento di System in IronPython)

from System.IO import File, IOException
from System.Text import Encoding

import RNWFlowCalculation_Report as rp


class FileLockedError(Exception):
    pass


# Pagine orizzontali [pt]
PAPER_SIZES = {"A4": (842.0, 595.0), "A3": (1191.0, 842.0)}
PAPER_DEFAULT = "A4"

MARGIN = 28.0
HEADER_SIZE = 9.0
TITLE_SIZE = 13.0
BLOCK_TITLE_SIZE = 9.5
BODY_SIZE = 8.0
NOTE_SIZE = 7.5
FOOTER_SIZE = 7.0
TABLE_SIZES = (7.0, 6.5, 6.0, 5.5)   # il primo che fa stare la tabella in larghezza
CELL_PAD_X = 2.5
CELL_PAD_Y = 1.8
LEADING = 1.2
# Una colonna di testo stretta per farla stare resta larga almeno quanto la sua
# parola piu' lunga, fino a questa larghezza [pt]
MAX_WORD_WIDTH = 90.0

FONT_REGULAR = "F1"
FONT_BOLD = "F2"
FONT_ITALIC = "F3"

# Colori come nella finestra dei risultati (RGB 0-255)
BLACK = (0, 0, 0)
GRAY = (110, 110, 110)
BLUE = (45, 90, 138)
HEADER_FILL = (217, 226, 236)
RULE = (200, 200, 200)
LIGHT_RULE = (232, 232, 232)
FIXTURE_FILL = (234, 244, 234)
SEVERITY_COLORS = {                 # (sfondo o None, testo)
    rp.SEV_ERROR: ((255, 199, 206), (156, 0, 6)),
    rp.SEV_WARNING: ((255, 235, 156), (122, 91, 0)),
    rp.SEV_PASS: ((198, 239, 206), BLACK),
    rp.SEV_INFO: (None, BLUE),
}

# Larghezze Helvetica / Helvetica-Bold per i codici 32-126 [1/1000 em]
_HELVETICA = (
    278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278,
    556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556,
    1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778,
    667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556,
    333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556,
    556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584)
_HELVETICA_BOLD = (
    278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278,
    556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611,
    975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778,
    667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556,
    333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611,
    611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584)
# Codici 128-255: larghezza media (lettere accentate, simboli di unita')
_UPPER_WIDTHS = {0xb0: 400, 0xb2: 333, 0xb3: 333, 0xb7: 278, 0xd7: 584, 0x95: 350,
                 0x96: 556, 0x97: 1000, 0x91: 222, 0x92: 222, 0x93: 333, 0x94: 333}

# Caratteri Unicode fuori da Latin-1 che cp1252 ha nei codici 128-159
_CP1252 = {0x20ac: 0x80, 0x2026: 0x85, 0x2018: 0x91, 0x2019: 0x92, 0x201c: 0x93,
           0x201d: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97}
# Sostituti per i caratteri che WinAnsi non ha
_REPLACE = {0x2264: u"<=", 0x2265: u">=", 0x2192: u"->", 0x2190: u"<-", 0x2248: u"~",
            0x0394: u"D", 0x2212: u"-", 0x00a0: u" ", 0x2715: u"x", 0x2713: u"v"}


def _clean(text):
    """Testo -> stringa con soli codici WinAnsi (un carattere = un byte del PDF)."""
    if text is None:
        return u""
    out = []
    for ch in u"{0}".format(text):
        code = ord(ch)
        if 32 <= code < 127 or 160 <= code < 256:
            out.append(ch)
        elif code in _CP1252:
            out.append(unichr(_CP1252[code]))
        elif code in _REPLACE:
            out.append(_REPLACE[code])
        elif code in (9, 10, 13):
            out.append(u" ")
        elif 0xd800 <= code < 0xe000 or 0xfe00 <= code < 0xfe10 or code == 0x200d:
            continue                # emoji (coppie surrogate) e selettori di variante
        else:
            out.append(u"?")
    return u"".join(out)


def _char_width(code, bold):
    if 32 <= code < 127:
        return (_HELVETICA_BOLD if bold else _HELVETICA)[code - 32]
    return _UPPER_WIDTHS.get(code, 611 if bold else 556)


def text_width(text, size, font=FONT_REGULAR):
    """Larghezza [pt] di un testo gia' pulito (_clean)."""
    bold = font == FONT_BOLD
    return sum(_char_width(ord(ch), bold) for ch in text) * size / 1000.0


def wrap(text, width, size, font=FONT_REGULAR):
    """Righe del testo (gia' pulito) entro width [pt]: a capo fra le parole,
    parole troppo lunghe spezzate."""
    # Tolleranza: la colonna e' larga quanto il suo testo piu' lungo piu' i
    # margini, e togliendo i margini l'arrotondamento puo' mandarlo a capo.
    width += 0.01
    lines = []
    for paragraph in text.split(u"\n"):
        current = u""
        for word in paragraph.split(u" "):
            candidate = word if not current else current + u" " + word
            if text_width(candidate, size, font) <= width:
                current = candidate
                continue
            if current:
                lines.append(current)
            current = u""
            while text_width(word, size, font) > width and len(word) > 1:
                cut = len(word) - 1
                while cut > 1 and text_width(word[:cut], size, font) > width:
                    cut -= 1
                lines.append(word[:cut])
                word = word[cut:]
            current = word
        lines.append(current)
    return lines or [u""]


def _rgb(color):
    return u"{0:.3f} {1:.3f} {2:.3f}".format(color[0] / 255.0, color[1] / 255.0,
                                             color[2] / 255.0)


def _num(value):
    return u"{0:.2f}".format(value)


# =============================================================================
# DOCUMENTO
# =============================================================================

class PdfDocument(object):
    """Pagine come liste di operatori del content stream. Le coordinate dei
    metodi partono dall'angolo in alto a sinistra (y verso il basso)."""

    def __init__(self, width, height):
        self.width = width
        self.height = height
        self.pages = []
        self.current = -1           # pagina su cui scrivono text(), rect(), line()

    def new_page(self):
        self.pages.append([])
        self.current = len(self.pages) - 1

    def _ops(self):
        return self.pages[self.current]

    def text(self, x, y, text, size, font=FONT_REGULAR, color=BLACK):
        """Testo (gia' pulito) con la linea di base a y."""
        if not text:
            return
        hex_text = u"".join(u"{0:02X}".format(ord(ch)) for ch in text)
        self._ops().append(u"BT {0} rg /{1} {2} Tf {3} {4} Td <{5}> Tj ET".format(
            _rgb(color), font, _num(size), _num(x), _num(self.height - y), hex_text))

    def rect(self, x, y, width, height, fill):
        self._ops().append(u"{0} rg {1} {2} {3} {4} re f".format(
            _rgb(fill), _num(x), _num(self.height - y - height), _num(width), _num(height)))

    def line(self, x1, y1, x2, y2, color=RULE, width=0.5):
        self._ops().append(u"{0} RG {1} w {2} {3} m {4} {5} l S".format(
            _rgb(color), _num(width), _num(x1), _num(self.height - y1), _num(x2),
            _num(self.height - y2)))

    def to_text(self):
        """Il file PDF come testo ASCII."""
        objects = []                # testo di ogni oggetto, numerati da 1

        def add(body):
            objects.append(body)
            return len(objects)

        catalog = add(None)         # riempiti dopo, servono i numeri
        pages_id = add(None)
        fonts = {}
        for key, name in ((FONT_REGULAR, "Helvetica"), (FONT_BOLD, "Helvetica-Bold"),
                          (FONT_ITALIC, "Helvetica-Oblique")):
            fonts[key] = add(u"<< /Type /Font /Subtype /Type1 /BaseFont /{0} "
                             u"/Encoding /WinAnsiEncoding >>".format(name))
        font_refs = u" ".join(u"/{0} {1} 0 R".format(key, number)
                              for key, number in sorted(fonts.items()))
        page_ids = []
        for ops in self.pages:
            stream = u"\n".join(ops)
            content = add(u"<< /Length {0} >>\nstream\n{1}\nendstream".format(
                len(stream), stream))
            page_ids.append(add(
                u"<< /Type /Page /Parent {0} 0 R /MediaBox [0 0 {1} {2}] "
                u"/Resources << /Font << {3} >> >> /Contents {4} 0 R >>".format(
                    pages_id, _num(self.width), _num(self.height), font_refs, content)))
        objects[catalog - 1] = u"<< /Type /Catalog /Pages {0} 0 R >>".format(pages_id)
        objects[pages_id - 1] = u"<< /Type /Pages /Kids [{0}] /Count {1} >>".format(
            u" ".join(u"{0} 0 R".format(number) for number in page_ids), len(page_ids))

        parts = [u"%PDF-1.4\n"]
        offsets = []
        position = len(parts[0])
        for index, body in enumerate(objects):
            chunk = u"{0} 0 obj\n{1}\nendobj\n".format(index + 1, body)
            offsets.append(position)
            parts.append(chunk)
            position += len(chunk)
        xref = [u"xref\n0 {0}\n".format(len(objects) + 1), u"0000000000 65535 f \n"]
        xref.extend(u"{0:010d} 00000 n \n".format(offset) for offset in offsets)
        parts.extend(xref)
        parts.append(u"trailer\n<< /Size {0} /Root {1} 0 R >>\nstartxref\n{2}\n%%EOF\n".format(
            len(objects) + 1, catalog, position))
        return u"".join(parts)


def _write_file(path, text):
    data = Encoding.ASCII.GetBytes(text)
    try:
        File.WriteAllBytes(path, data)
    except IOException:
        raise FileLockedError(path)


# =============================================================================
# IMPAGINAZIONE DEL REPORT
# =============================================================================

class _Layout(object):
    """Scrive i blocchi del report una pagina dopo l'altra."""

    def __init__(self, doc, report):
        self.doc = doc
        self.report = report
        self.left = MARGIN
        self.right = doc.width - MARGIN
        self.top = MARGIN + HEADER_SIZE + 8.0
        self.bottom = doc.height - MARGIN - FOOTER_SIZE - 6.0
        self.y = self.top
        self.tab_name = u""

    @property
    def width(self):
        return self.right - self.left

    def new_page(self):
        self.doc.new_page()
        self.doc.text(self.left, MARGIN + HEADER_SIZE, _clean(self.report.header), HEADER_SIZE,
                      FONT_BOLD, BLUE)
        name = _clean(self.tab_name)
        self.doc.text(self.right - text_width(name, HEADER_SIZE, FONT_BOLD),
                      MARGIN + HEADER_SIZE, name, HEADER_SIZE, FONT_BOLD, GRAY)
        self.doc.line(self.left, MARGIN + HEADER_SIZE + 3.0, self.right,
                      MARGIN + HEADER_SIZE + 3.0, RULE, 0.6)
        self.y = self.top + 6.0

    def ensure(self, height):
        if self.y + height > self.bottom:
            self.new_page()

    def footers(self):
        """Riga in fondo a ogni pagina: report.footer e "Page n of N"."""
        total = len(self.doc.pages)
        y = self.doc.height - MARGIN
        footer = _clean(self.report.footer)
        for index in range(total):
            self.doc.current = index
            self.doc.text(self.left, y, footer, FOOTER_SIZE, FONT_REGULAR, GRAY)
            label = u"Page {0} of {1}".format(index + 1, total)
            self.doc.text(self.right - text_width(label, FOOTER_SIZE), y, label, FOOTER_SIZE,
                          FONT_REGULAR, GRAY)

    # ------------------------------------------------------------------ blocchi

    def tab(self, tab, extra_blocks=()):
        self.tab_name = tab.name
        self.new_page()
        self.doc.text(self.left, self.y + TITLE_SIZE, _clean(tab.name), TITLE_SIZE, FONT_BOLD,
                      BLACK)
        self.y += TITLE_SIZE + 8.0
        for block in list(tab.blocks) + list(extra_blocks):
            if isinstance(block, rp.Pairs):
                self.pairs(block)
            elif isinstance(block, rp.Text):
                self.note(block.text)
            elif isinstance(block, rp.Table):
                self.table(block)
            self.y += 8.0

    def block_title(self, title):
        if not title:
            return
        self.ensure(BLOCK_TITLE_SIZE * 3)
        self.doc.text(self.left, self.y + BLOCK_TITLE_SIZE, _clean(title).upper(),
                      BLOCK_TITLE_SIZE, FONT_BOLD, BLUE)
        self.y += BLOCK_TITLE_SIZE + 5.0

    def pairs(self, block):
        self.block_title(block.title)
        size = BODY_SIZE
        labels = [_clean(label) for label, _value in block.rows]
        label_width = min(max([text_width(label, size, FONT_BOLD) for label in labels] + [0.0])
                          + 12.0, self.width * 0.4)
        value_width = self.width - label_width
        for label, (_raw, value) in zip(labels, block.rows):
            label_lines = wrap(label, label_width - 12.0, size, FONT_BOLD)
            value_lines = wrap(_clean(rp.pair_text(value)), value_width, size)
            height = max(len(label_lines), len(value_lines)) * size * LEADING + 2.0
            self.ensure(height)
            self._lines(self.left, self.y, label_lines, size, FONT_BOLD, BLACK)
            self._lines(self.left + label_width, self.y, value_lines, size, FONT_REGULAR, BLACK)
            self.y += height

    def note(self, text):
        size = NOTE_SIZE
        for line in wrap(_clean(text), self.width, size):
            self.ensure(size * LEADING)
            self.doc.text(self.left, self.y + size, line, size, FONT_REGULAR, GRAY)
            self.y += size * LEADING

    def _lines(self, x, y, lines, size, font, color, align_right_width=None):
        for index, line in enumerate(lines):
            baseline = y + size + index * size * LEADING
            if align_right_width is not None:
                self.doc.text(x + align_right_width - text_width(line, size, font), baseline,
                              line, size, font, color)
            else:
                self.doc.text(x, baseline, line, size, font, color)

    # ------------------------------------------------------------------ tabelle

    def _column_widths(self, columns, headers, cells):
        """(dimensione del carattere, larghezze) che fanno stare la tabella.

        Prima si riduce il carattere. Al carattere piu' piccolo le colonne di
        testo si stringono fino alla parola piu' lunga (vanno a capo fra le
        parole) e lo spazio che resta si divide in proporzione a quanto manca a
        ciascuna per stare su una riga. Solo se nemmeno cosi' ci stanno le
        colonne di testo scendono sotto la parola piu' lunga (parole spezzate).
        """
        for size in TABLE_SIZES:
            natural, minimum = self._natural_widths(columns, headers, cells, size)
            if sum(natural) <= self.width:
                return size, natural
        flexible = [i for i, column in enumerate(columns) if column.kind == rp.TEXT]
        widths = list(natural)
        room = self.width - sum(natural[i] for i in range(len(columns)) if i not in flexible)
        least = sum(minimum[i] for i in flexible)
        if flexible and room >= least:
            extra = room - least
            wanted = sum(natural[i] - minimum[i] for i in flexible)
            for i in flexible:
                share = (natural[i] - minimum[i]) * extra / wanted if wanted else 0.0
                widths[i] = minimum[i] + share
        elif flexible and room > 0:
            for i in flexible:
                widths[i] = minimum[i] * room / least
        total = sum(widths)
        if total > self.width:
            # Ultima risorsa: tutto in proporzione
            widths = [w * self.width / total for w in widths]
        return size, widths

    @staticmethod
    def _natural_widths(columns, headers, cells, size):
        """Per colonna: larghezza su una riga e larghezza minima (parola piu'
        lunga, al massimo MAX_WORD_WIDTH); per i numeri coincidono."""
        natural = []
        minimum = []
        for index, column in enumerate(columns):
            header_word = max([text_width(word, size, FONT_BOLD)
                               for word in headers[index].split(u" ")] + [0.0])
            content = max([text_width(row[index], size) for row in cells] + [0.0])
            width = max(content, header_word) + 2 * CELL_PAD_X
            natural.append(width)
            if column.kind == rp.TEXT:
                longest = max([text_width(word, size) for row in cells
                               for word in row[index].split(u" ")] + [0.0])
                minimum.append(min(width, max(header_word, min(longest, MAX_WORD_WIDTH))
                                   + 2 * CELL_PAD_X))
            else:
                minimum.append(width)
        return natural, minimum

    def table(self, block):
        self.block_title(block.title)
        columns = block.columns
        if not columns:
            return
        headers = [_clean(column.header) for column in columns]
        cells = [[_clean(rp.cell_text(column, row.values.get(column.key))) for column in columns]
                 for row in block.rows]
        size, widths = self._column_widths(columns, headers, cells)
        lines_height = size * LEADING
        header_lines = [wrap(text, w - 2 * CELL_PAD_X, size, FONT_BOLD)
                        for text, w in zip(headers, widths)]
        header_height = max(len(lines) for lines in header_lines) * lines_height \
            + 2 * CELL_PAD_Y

        def draw_header():
            self.ensure(header_height + lines_height + 2 * CELL_PAD_Y)
            self.doc.rect(self.left, self.y, sum(widths), header_height, HEADER_FILL)
            x = self.left
            for column, lines, width in zip(columns, header_lines, widths):
                # Intestazioni allineate in basso, come nella griglia
                top = self.y + header_height - CELL_PAD_Y - len(lines) * lines_height
                self._lines(x + CELL_PAD_X, top, lines, size, FONT_BOLD, BLACK,
                            width - 2 * CELL_PAD_X if column.kind == rp.NUMBER else None)
                x += width
            self.y += header_height
            self.doc.line(self.left, self.y, self.left + sum(widths), self.y, GRAY, 0.6)

        draw_header()
        for row, texts in zip(block.rows, cells):
            fixture = row.kind == rp.ROW_FIXTURE
            font = FONT_ITALIC if fixture else FONT_REGULAR
            wrapped = [wrap(text, width - 2 * CELL_PAD_X, size, font) if column.kind == rp.TEXT
                       else [text] for column, text, width in zip(columns, texts, widths)]
            height = max(len(lines) for lines in wrapped) * lines_height + 2 * CELL_PAD_Y
            if self.y + height > self.bottom:
                self.new_page()
                draw_header()
            if fixture:
                self.doc.rect(self.left, self.y, sum(widths), height, FIXTURE_FILL)
            x = self.left
            for column, lines, width in zip(columns, wrapped, widths):
                color = GRAY if row.kind == rp.ROW_NO_FLOW else BLACK
                severity = row.severities.get(column.key) if column.severity else None
                if severity in SEVERITY_COLORS and lines != [u""]:
                    fill, color = SEVERITY_COLORS[severity]
                    if fill is not None:
                        self.doc.rect(x, self.y, width, height, fill)
                self._lines(x + CELL_PAD_X, self.y + CELL_PAD_Y, lines, size, font, color,
                            width - 2 * CELL_PAD_X if column.kind == rp.NUMBER else None)
                x += width
            self.y += height
            self.doc.line(self.left, self.y, self.left + sum(widths), self.y, LIGHT_RULE, 0.4)


def export_report(path, report, paper=PAPER_DEFAULT):
    """Scrive il report (RNWFlowCalculation_Report.Report) in PDF: una scheda per
    pagina nuova, le impostazioni in coda alla scheda General. FileLockedError
    se il file e' aperto in un altro programma."""
    width, height = PAPER_SIZES.get(paper, PAPER_SIZES[PAPER_DEFAULT])
    doc = PdfDocument(width, height)
    layout = _Layout(doc, report)
    settings = [rp.Pairs("Settings", report.settings)] if report.settings else []
    for tab in report.tabs:
        layout.tab(tab, settings if tab.name == rp.TAB_GENERAL else ())
    layout.footers()
    _write_file(path, doc.to_text())
