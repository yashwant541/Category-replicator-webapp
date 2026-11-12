"""A fast, values-only reader for .xlsx / .xlsm sheets - for the big NEW input files of a run.

openpyxl builds an element tree and a cell object per cell (tens of thousands of cells a second); this reads the sheet XML as text
in chunks and pulls the cells out with regular expressions (a few hundred thousand cells a second per process).  It returns the same
cell values openpyxl gives with data_only=True (numbers, text, booleans, dates as datetime).  Anything unusual raises ``Unsupported``
and the caller falls back to openpyxl, so a strange file is slower, never wrong.

Not for the sample / mapping files: those need formula text, which only openpyxl reads (see tables.read_workbook).
"""

from __future__ import annotations

import codecs
import datetime as dt
import html
import io
import re
import zipfile
from typing import Any, Dict, Iterator, List, Optional, Tuple


class Unsupported(Exception):
    pass


_ATTR = re.compile(r'([\w:]+)="([^"]*)"')
_ROW = re.compile(r"<row\b([^>]*?)(?:/>|>(.*?)</row>)", re.S)
_CELL = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
_V = re.compile(r"<v>(.*?)</v>", re.S)
_T = re.compile(r"<t(?:\s[^>]*)?>(.*?)</t>", re.S)
_PH = re.compile(r"<rPh\b.*?</rPh>", re.S)
_ESC = re.compile(r"_x([0-9A-Fa-f]{4})_")
_BUILTIN_DATE = set(range(14, 23)) | {27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58}
_EPOCH = dt.datetime(1899, 12, 30)
_COL: Dict[str, int] = {}


def _col(letters: str) -> int:
    n = _COL.get(letters)
    if n is None:
        n = 0
        for ch in letters:
            n = n * 26 + ord(ch) - 64
        _COL[letters] = n - 1
        n -= 1
    return n


def _text(s: str) -> str:
    if "&" in s:
        s = html.unescape(s)
    if "_x" in s:
        s = _ESC.sub(lambda m: chr(int(m.group(1), 16)), s)
    return s


def _is_date_code(code: str) -> bool:
    c = re.sub(r'\[[^\]]*\]|"[^"]*"|\\.|_.|\*.', "", code)
    return bool(re.search(r"[dmyhs]", c, re.I))


def _read_part(z: zipfile.ZipFile, name: str) -> Optional[str]:
    try:
        return z.read(name).decode("utf-8")
    except KeyError:
        return None


def _targets(z: zipfile.ZipFile) -> Tuple[List[Tuple[str, str]], bool]:
    """-> ([(sheet title, part path)], date1904)"""
    wb = _read_part(z, "xl/workbook.xml")
    rels = _read_part(z, "xl/_rels/workbook.xml.rels")
    if wb is None or rels is None:
        raise Unsupported("no workbook parts")
    rid = {}
    for m in re.finditer(r"<Relationship\b([^>]*?)/?>", rels):
        a = dict(_ATTR.findall(m.group(1)))
        if "Id" in a and "Target" in a:
            t = a["Target"]
            rid[a["Id"]] = t.lstrip("/") if t.startswith("/") else "xl/" + t
    out = []
    for m in re.finditer(r"<sheet\b([^>]*?)/?>", wb):
        a = dict(_ATTR.findall(m.group(1)))
        r = a.get("r:id") or next((v for k, v in a.items() if k.endswith(":id")), None)
        if "name" not in a or r not in rid:
            raise Unsupported("sheet without a part")
        out.append((_text(a["name"]), rid[r]))
    return out, bool(re.search(r'date1904="(1|true)"', wb))


def _shared_strings(z: zipfile.ZipFile) -> List[str]:
    xml = _read_part(z, "xl/sharedStrings.xml")
    if xml is None:
        return []
    out = []
    for m in re.finditer(r"<si\b[^>]*?(?:/>|>(.*?)</si>)", xml, re.S):
        body = m.group(1) or ""
        if "<rPh" in body:
            body = _PH.sub("", body)
        out.append(_text("".join(_T.findall(body))))
    return out


def _date_styles(z: zipfile.ZipFile) -> set:
    """Indexes of the cell formats (cellXfs) that show a date."""
    xml = _read_part(z, "xl/styles.xml")
    if xml is None:
        return set()
    custom = {}
    for m in re.finditer(r"<numFmt\b([^>]*?)/?>", xml):
        a = dict(_ATTR.findall(m.group(1)))
        if "numFmtId" in a and "formatCode" in a:
            custom[int(a["numFmtId"])] = _text(a["formatCode"])
    m = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", xml, re.S)
    if not m:
        return set()
    dates = set()
    for i, x in enumerate(re.finditer(r"<xf\b([^>]*?)/?>", m.group(1))):
        fid = int(dict(_ATTR.findall(x.group(1))).get("numFmtId", 0))
        if fid in _BUILTIN_DATE or (fid in custom and _is_date_code(custom[fid])):
            dates.add(i)
    return dates


def _chunks(f, size: int = 1 << 23) -> Iterator[str]:
    dec = codecs.getincrementaldecoder("utf-8")()
    while True:
        b = f.read(size)
        if not b:
            break
        yield dec.decode(b)
    tail = dec.decode(b"", final=True)
    if tail:
        yield tail


def _serial(v: float, d1904: bool) -> dt.datetime:
    base = dt.datetime(1904, 1, 1) if d1904 else _EPOCH
    return base + dt.timedelta(days=v)


def _rows(z: zipfile.ZipFile, part: str, shared: List[str], dates: set, d1904: bool) -> List[List[Any]]:
    rows: List[List[Any]] = []
    buf = ""
    in_data = False
    with z.open(part) as f:
        for chunk in _chunks(f):
            buf += chunk
            if not in_data:
                i = buf.find("<sheetData")
                if i < 0:
                    if "<x:sheetData" in buf:
                        raise Unsupported("prefixed xml")
                    buf = buf[-64:]
                    continue
                j = buf.find(">", i)
                if j < 0:
                    continue
                if buf[j - 1] == "/":                              # <sheetData/>: no cells
                    return rows
                buf = buf[j + 1:]
                in_data = True
            end = buf.rfind("</row>")
            cut = max(end + 6 if end >= 0 else 0, 0)
            block, buf = buf[:cut], buf[cut:]
            _parse_block(block, rows, shared, dates, d1904)
        if buf.strip():
            _parse_block(buf, rows, shared, dates, d1904)
    return rows


def _parse_block(block: str, rows: List[List[Any]], shared: List[str], dates: set, d1904: bool) -> None:
    for rm in _ROW.finditer(block):
        ra = dict(_ATTR.findall(rm.group(1)))
        rnum = int(ra["r"]) if "r" in ra else len(rows) + 1
        while len(rows) < rnum - 1:
            rows.append([])
        body = rm.group(2)
        row: List[Any] = []
        if body:
            for cm in _CELL.finditer(body):
                a = dict(_ATTR.findall(cm.group(1)))
                ref = a.get("r")
                if ref is None:
                    raise Unsupported("cell without a reference")
                k = 0
                while ref[k].isalpha():
                    k += 1
                c0 = _col(ref[:k])
                inner = cm.group(2)
                if not inner:
                    continue
                t = a.get("t", "n")
                v = None
                if t == "inlineStr":
                    v = _text("".join(_T.findall(inner)))
                else:
                    vm = _V.search(inner)
                    if vm is None:
                        continue
                    raw = vm.group(1)
                    if t == "s":
                        v = shared[int(raw)]
                    elif t == "str":
                        v = _text(raw)
                    elif t == "b":
                        v = raw.strip() in ("1", "true")
                    elif t == "e":
                        v = _text(raw)
                    elif t == "d":
                        try:
                            v = dt.datetime.fromisoformat(raw.strip().rstrip("Z"))
                        except ValueError:
                            v = raw
                    else:
                        if raw == "":
                            continue
                        try:
                            n = float(raw) if ("." in raw or "E" in raw or "e" in raw) else int(raw)
                        except ValueError:
                            raise Unsupported("odd number %r" % raw)
                        if "s" in a and int(a["s"]) in dates:
                            try:
                                v = _serial(float(n), d1904)
                            except (OverflowError, ValueError):
                                v = n
                        else:
                            v = n
                if v is None:
                    continue
                if len(row) <= c0:
                    row.extend([None] * (c0 + 1 - len(row)))
                row[c0] = v
        rows.append(row)


def read_sheets(file_bytes: bytes, only: Optional[set] = None) -> "Dict[str, List[List[Any]]]":
    """{sheet title: rows of cell values} for an .xlsx/.xlsm workbook, in workbook order (``only`` limits the sheets)."""
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            sheets, d1904 = _targets(z)
            shared, dates = _shared_strings(z), _date_styles(z)
            return {title: _rows(z, part, shared, dates, d1904) for title, part in sheets if only is None or title in only}
    except Unsupported:
        raise
    except (KeyError, ValueError, IndexError, zipfile.BadZipFile, UnicodeDecodeError) as exc:
        raise Unsupported(str(exc))


def read_sheet(file_bytes: bytes, title: str) -> List[List[Any]]:
    return read_sheets(file_bytes, {title})[title]
