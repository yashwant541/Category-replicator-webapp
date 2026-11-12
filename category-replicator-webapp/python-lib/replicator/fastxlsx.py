"""A small streaming .xlsx writer for BIG outputs (millions of cells), able to use several worker processes.

Why not openpyxl: for the template this app writes - pasted values, formulas, a few styles - openpyxl spends most of its time
building cell/style objects.  Here every row is turned straight into XML text, in chunks, and the chunks are encoded by a pool
of processes while the parent streams them into the zip in order.  Small outputs still use openpyxl (see engine.write_workbook).

What is written: inline strings (no shared-string table), numbers, booleans, dates (as Excel serial numbers + a date style),
formulas (text only; Excel calculates on open), column widths, a frozen header, number formats, header fills.
"""

from __future__ import annotations

import datetime as dt
import io
import numbers
import re
import zipfile
from typing import Any, Dict, Iterable, Iterator, List, Optional, Tuple

from . import perf

_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
_EPOCH = dt.datetime(1899, 12, 30)
_LETTERS: Dict[int, str] = {}


def letters(c0: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA"""
    s = _LETTERS.get(c0)
    if s is None:
        n, s = c0 + 1, ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        _LETTERS[c0] = s
    return s


def esc(s: str) -> str:
    return _BAD.sub("", s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def esc_attr(s: str) -> str:
    return esc(s).replace('"', "&quot;")


# ----------------------------------------------------------------------------- formulas
def fill_formula(c: Tuple[List[str], List[int], int], row: int) -> str:
    """Compiled formula (see engine.compile_formula) for another row: Excel's fill-down."""
    texts, nums, base = c
    if not nums:
        return texts[0]
    d = row - base
    out = [texts[0]]
    for i, n in enumerate(nums):
        m = n + d
        out.append(str(m if m >= 1 else n))
        out.append(texts[i + 1])
    return "".join(out)


# ------------------------------------------------------------------------------- styles
_BUILTIN_FMT = {"General": 0, "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10, "@": 49}


class Styles:
    """The cell formats of one workbook: font / fill / number format / alignment combinations -> a style index."""

    FONTS = ['<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>',
             '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>',
             '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>']
    FILLS = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>',
             '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3A5F"/><bgColor indexed="64"/></patternFill></fill>',
             '<fill><patternFill patternType="solid"><fgColor rgb="FF2F6B4F"/><bgColor indexed="64"/></patternFill></fill>']
    PLAIN, BOLD, WHITE_BOLD = 0, 1, 2
    NO_FILL, FILL_INPUT, FILL_FORMULA = 0, 2, 3

    def __init__(self) -> None:
        self.fmt_ids: Dict[str, int] = {}
        self.xfs: List[Tuple[int, int, int, bool]] = [(0, 0, 0, False)]

    def fmt(self, code: Optional[str]) -> int:
        if not code or code == "General":
            return 0
        if code in _BUILTIN_FMT:
            return _BUILTIN_FMT[code]
        if code not in self.fmt_ids:
            self.fmt_ids[code] = 164 + len(self.fmt_ids)
        return self.fmt_ids[code]

    def xf(self, font: int = 0, fill: int = 0, fmt: Optional[str] = None, wrap: bool = False) -> int:
        key = (font, fill, self.fmt(fmt), wrap)
        if key not in self.xfs:
            self.xfs.append(key)
        return self.xfs.index(key)

    def xml(self) -> str:
        nf = "".join('<numFmt numFmtId="%d" formatCode="%s"/>' % (i, esc_attr(c)) for c, i in self.fmt_ids.items())
        xfs = []
        for font, fill, fmt, wrap in self.xfs:
            al = '<alignment vertical="center" wrapText="1"/>' if wrap else ""
            xfs.append('<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="0" xfId="0"%s%s%s%s>%s</xf>' % (
                fmt, font, fill, ' applyNumberFormat="1"' if fmt else "", ' applyFont="1"' if font else "",
                ' applyFill="1"' if fill else "", ' applyAlignment="1"' if wrap else "", al) if al else
                '<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="0" xfId="0"%s%s%s/>' % (
                    fmt, font, fill, ' applyNumberFormat="1"' if fmt else "", ' applyFont="1"' if font else "",
                    ' applyFill="1"' if fill else ""))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                + ('<numFmts count="%d">%s</numFmts>' % (len(self.fmt_ids), nf) if nf else "")
                + '<fonts count="%d">%s</fonts>' % (len(self.FONTS), "".join(self.FONTS))
                + '<fills count="%d">%s</fills>' % (len(self.FILLS), "".join(self.FILLS))
                + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
                + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                + '<cellXfs count="%d">%s</cellXfs>' % (len(xfs), "".join(xfs))
                + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
                + '</styleSheet>')


# ---------------------------------------------------------------------------- one cell
def _serial(v: dt.date) -> float:
    if isinstance(v, dt.datetime):
        if v.tzinfo is not None:
            v = v.replace(tzinfo=None)
        d = v - _EPOCH
    else:
        d = dt.datetime(v.year, v.month, v.day) - _EPOCH
    return d.days + (d.seconds + d.microseconds / 1e6) / 86400.0


def cell_xml(ref: str, v: Any, style: int, date_styles: Tuple[int, int] = (0, 0), formula: bool = False) -> str:
    """The <c> element for one value ('' for blank).  A string that starts with = is a formula only in formula columns."""
    if v is None:
        return ""
    s = ' s="%d"' % style if style else ""
    if isinstance(v, str):
        if formula:
            return '<c r="%s"%s><f>%s</f></c>' % (ref, s, esc(v[1:] if v.startswith("=") else v))
        if not v:
            return ""
        if v[0] == "=":
            v = "'" + v                                  # pasted text that merely starts with "=" stays text
        return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, esc(v[:32767]))
    if v != v:                                           # NaN / NaT
        return ""
    if isinstance(v, bool) or type(v).__name__ == "bool_":
        return '<c r="%s"%s t="b"><v>%d</v></c>' % (ref, s, 1 if v else 0)
    if isinstance(v, numbers.Integral):
        return '<c r="%s"%s><v>%d</v></c>' % (ref, s, int(v))
    if isinstance(v, numbers.Real):
        f = float(v)
        if f in (float("inf"), float("-inf")):
            return ""
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, repr(f))
    if isinstance(v, dt.date):                           # datetime is a date
        if not style:
            s = ' s="%d"' % (date_styles[1] if isinstance(v, dt.datetime) and (v.hour or v.minute or v.second or v.microsecond)
                             else date_styles[0])
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, repr(_serial(v)))
    return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, esc(str(v)[:32767]))


# --------------------------------------------------------------------- encoding the rows
# plan entry: (col0, kind, a, style)   kind "v": a = position in the row's value tuple; kind "f": a = compiled formula
def encode_rows(task: Tuple[int, List[tuple], list, Dict[int, Dict[int, tuple]], Tuple[int, int]]) -> bytes:
    """Rows first .. first+len(rows)-1 as sheet XML (run in a worker process)."""
    first, rows, plan, extras, date_styles = task
    cols = [(c0, letters(c0), kind, a, style) for c0, kind, a, style in plan]
    out: List[str] = []
    for k, vals in enumerate(rows):
        r = first + k
        cells = []
        for c0, L, kind, a, style in cols:
            if kind == "f":
                cells.append(cell_xml(L + str(r), fill_formula(a, r), style, date_styles, formula=True))
            else:
                cells.append(cell_xml(L + str(r), vals[a], style, date_styles))
        ex = extras.get(r) if extras else None
        if ex:                                           # a parameter cell that sits inside the data rows
            by_col = {c[0]: x for c, x in zip(cols, cells)}
            for c0, (v, st) in ex.items():
                by_col[c0] = cell_xml(letters(c0) + str(r), v, st, date_styles)
            cells = [by_col[c] for c in sorted(by_col)]
        out.append('<row r="%d">%s</row>' % (r, "".join(cells)))
    return "".join(out).encode("utf-8")


def _static_row(r: int, cells: Iterable[Tuple[int, Any, int]], date_styles: Tuple[int, int]) -> str:
    xs = [cell_xml(letters(c0) + str(r), v, st, date_styles) for c0, v, st in sorted(cells, key=lambda t: t[0])]
    xs = [x for x in xs if x]
    return '<row r="%d">%s</row>' % (r, "".join(xs)) if xs else ""


# ------------------------------------------------------------------------------ package
_HEAD = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')


def _sheet_prefix(spec: Dict[str, Any]) -> str:
    x = [_HEAD]
    if spec.get("freeze"):
        top = int(spec["freeze"])                       # first row that scrolls; the rows above it stay
        x.append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="%d" topLeftCell="A%d" activePane="bottomLeft" state="frozen"/>'
                 '<selection pane="bottomLeft" activeCell="A%d" sqref="A%d"/></sheetView></sheetViews>' % (top - 1, top, top, top))
    x.append('<sheetFormatPr defaultRowHeight="15"/>')
    if spec.get("widths"):
        x.append("<cols>" + "".join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (c + 1, c + 1, w)
                                    for c, w in sorted(spec["widths"].items())) + "</cols>")
    x.append("<sheetData>")
    return "".join(x)


def _write_sheet(f, spec: Dict[str, Any], pool, chunk_rows: int, date_styles: Tuple[int, int]) -> None:
    """spec: name, widths, freeze, then EITHER static (iterable of (row, [(col0, value, style)])) OR
    head {row: {col0: (value, style)}} + data {first, n, values: [column lists], plan}."""
    f.write(_sheet_prefix(spec).encode("utf-8"))
    if spec.get("static") is not None:
        for r, cells in spec["static"]:
            x = _static_row(r, cells, date_styles)
            if x:
                f.write(x.encode("utf-8"))
    else:
        head = spec.get("head") or {}
        data = spec.get("data")
        first = data["first"] if data else 10 ** 9
        for r in sorted(k for k in head if k < first):
            x = _static_row(r, [(c0, v, st) for c0, (v, st) in head[r].items()], date_styles)
            if x:
                f.write(x.encode("utf-8"))
        if data:
            n, cols, plan = data["n"], data["values"], data["plan"]
            extras = {r: cells for r, cells in head.items() if r >= first}

            def tasks() -> Iterator[tuple]:
                for a in range(0, n, chunk_rows):
                    b = min(n, a + chunk_rows)
                    rows = list(zip(*(c[a:b] for c in cols))) if cols else [()] * (b - a)
                    yield (first + a, rows, plan, {r: v for r, v in extras.items() if first + a <= r < first + b}, date_styles)
            for blob in perf.map_ordered(pool, encode_rows, tasks()):
                f.write(blob)
            for r in sorted(k for k in extras if k >= first + n):          # parameter cells below the last data row
                x = _static_row(r, [(c0, v, st) for c0, (v, st) in extras[r].items()], date_styles)
                if x:
                    f.write(x.encode("utf-8"))
    f.write(b"</sheetData></worksheet>")


def write_package(sheets: List[Dict[str, Any]], styles: Styles, date_styles: Tuple[int, int] = (0, 0),
                  pool=None, chunk_rows: int = perf.CHUNK_ROWS) -> bytes:
    """Assemble the workbook (every sheet spec as described in _write_sheet) -> .xlsx bytes."""
    buf = io.BytesIO()
    try:
        zf = zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, allowZip64=True, compresslevel=1)
    except TypeError:                                    # Python < 3.7
        zf = zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, allowZip64=True)
    with zf:
        n = len(sheets)
        zf.writestr("[Content_Types].xml", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            + "".join('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1) for i in range(n))
            + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '</Types>'))
        zf.writestr("_rels/.rels", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>'))
        zf.writestr("xl/workbook.xml", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<bookViews><workbookView/></bookViews><sheets>'
            + "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (esc_attr(s["name"]), i + 1, i + 1) for i, s in enumerate(sheets))
            + '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'))
        zf.writestr("xl/_rels/workbook.xml.rels", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1) for i in range(n))
            + '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (n + 1)
            + '</Relationships>'))
        for i, spec in enumerate(sheets):
            with zf.open("xl/worksheets/sheet%d.xml" % (i + 1), "w", force_zip64=True) as f:
                _write_sheet(f, spec, pool, chunk_rows, date_styles)
        zf.writestr("xl/styles.xml", styles.xml())
    return buf.getvalue()
