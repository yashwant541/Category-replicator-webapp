"""Reading uploaded files into tidy DataFrames and (de)serialising mapping tables."""

from __future__ import annotations

import datetime as dt
import gzip
import html
import io
import math
import re
import uuid
import zipfile
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import openpyxl
import pandas as pd

from replicator_formulas.sample_parser import MAX_SCAN_ROWS, detect_header_row, _headers_from_row, _norm, is_toggle_sheet
from . import fastread, perf
from .rules import is_blank, nk


def jsonable(v: Any) -> Any:
    if v is None or v is pd.NaT:
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating, float)):
        f = float(v)
        return None if math.isnan(f) or math.isinf(f) else (int(f) if f.is_integer() and abs(f) < 1e15 else f)
    if isinstance(v, (np.bool_, bool)):
        return bool(v)
    if isinstance(v, pd.Timestamp):
        return v.isoformat()
    if isinstance(v, (dt.datetime, dt.date)):
        return v.isoformat()
    if isinstance(v, dt.time):
        return v.isoformat()
    return v if isinstance(v, (int, str)) else str(v)


EXCEL_EXT = (".xlsx", ".xlsm", ".xltx", ".xltm")          # formulas can be read
VALUES_EXT = (".xls", ".xlsb", ".ods")                     # values only
TEXT_EXT = (".csv", ".tsv", ".txt")
ALLOWED_EXT = EXCEL_EXT + VALUES_EXT + TEXT_EXT + (".parquet", ".gz")
_ENGINES = {".xls": ("xlrd", "xlrd"), ".xlsb": ("pyxlsb", "pyxlsb"), ".ods": ("odf", "odfpy")}


def file_ext(name: str) -> str:
    """'.csv' for 'a.CSV' and for 'a.csv.gz' - the extension that decides how a file is read."""
    low = (name or "").lower()
    if low.endswith(".gz"):
        low = low[:-3]
        return low[low.rfind("."):] if "." in low else ".csv"
    return low[low.rfind("."):] if "." in low else ""


def supported(name: str) -> bool:
    return file_ext(name) in EXCEL_EXT + VALUES_EXT + TEXT_EXT + (".parquet",)


def holds_formulas(name: str) -> bool:
    """Only the .xlsx family keeps formula text we can read; a sample OUTPUT must be one of these."""
    return file_ext(name) in EXCEL_EXT


def _blank_mask(raw: pd.DataFrame) -> np.ndarray:
    """True where a cell is empty: None / NaN / NaT / an empty or whitespace-only string (same rule as ``is_blank``)."""
    m = raw.isna().values.copy()
    for j in range(raw.shape[1]):
        col = raw.iloc[:, j]
        if col.dtype != object:
            continue
        isstr = col.map(type).eq(str).values
        if isstr.any():
            m[np.nonzero(isstr)[0], j] |= pd.Series(col.values[isstr]).str.strip().eq("").values
    return m


def _frame_to_df(raw: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    """A sheet read with no header (cells in rows/columns) -> a tidy table + the 0-based header row.  Vectorised, so a
    million-row sheet is a few seconds, not minutes."""
    if raw.shape[0] == 0 or raw.shape[1] == 0:
        return pd.DataFrame(), 0
    raw = raw.reset_index(drop=True)
    raw.columns = range(raw.shape[1])
    width = raw.shape[1]
    blank = _blank_mask(raw)
    if blank.all():
        return pd.DataFrame(), 0
    head = [[None if blank[i, j] else raw.iat[i, j] for j in range(width)] for i in range(min(len(raw), MAX_SCAN_ROWS + 2))]
    hr = detect_header_row(head)
    hdr = [None if blank[hr, j] else raw.iat[hr, j] for j in range(width)]
    body_blank = blank[hr + 1:]
    keep = [j for j in range(width) if _norm(hdr[j]) or (body_blank.shape[0] and not body_blank[:, j].all())]
    rowkeep = ~body_blank.all(axis=1) if body_blank.shape[0] else np.zeros(0, bool)
    body = raw.iloc[hr + 1:, keep]
    body = body[rowkeep]
    body = body.astype(object).where(body.notna(), None)
    body.columns = _headers_from_row([hdr[j] for j in keep])
    body = body.reset_index(drop=True)
    body.attrs["col_offset"] = int(min(keep)) if keep else 0
    body.attrs["col_idx"] = [int(k) for k in keep]
    return body, hr


def _rows_to_df(rows: List[List[Any]]) -> Tuple[pd.DataFrame, int]:
    if not rows:
        return pd.DataFrame(), 0
    return _frame_to_df(pd.DataFrame(rows, dtype=object))


def _sniff_sep(head: bytes) -> str:
    lines = [ln for ln in head.decode("utf-8", "replace").splitlines() if ln.strip()][:5]
    best, score = ",", 0
    for sep in (",", ";", "\t", "|"):
        counts = [ln.count(sep) for ln in lines]
        s = min(counts) if counts else 0                      # present on every sampled line
        if s > score:
            best, score = sep, s
    return best


def _read_text(file_bytes: bytes, filename: str) -> pd.DataFrame:
    if (filename or "").lower().endswith(".gz"):
        file_bytes = gzip.decompress(file_bytes)
    ext = file_ext(filename)
    sep = "\t" if ext == ".tsv" else _sniff_sep(file_bytes[:65536])
    kw = dict(sep=sep, header=None, dtype=object, keep_default_na=False, skip_blank_lines=False)
    try:
        file_bytes.decode("utf-8-sig")
        enc = "utf-8-sig"
    except UnicodeDecodeError:
        enc = "cp1252"                                        # Excel's "CSV (comma delimited)" on Windows
    try:
        raw = pd.read_csv(io.BytesIO(file_bytes), encoding=enc, encoding_errors="replace", **kw)
    except TypeError:                                         # pandas < 1.3 has no encoding_errors
        raw = pd.read_csv(io.BytesIO(file_bytes), encoding=enc, **kw)
    return raw.where(raw != "", None)                         # an empty field is an empty cell


_SHARED: Dict[str, bytes] = {}                                # file bytes shared with forked worker processes


def _sheet_titles(file_bytes: bytes) -> List[str]:
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        xml = z.read("xl/workbook.xml").decode("utf-8", "replace")
    return [html.unescape(m) for m in re.findall(r"<(?:\w+:)?sheet\b[^>]*?\bname=\"([^\"]*)\"", xml)]


def _sheet_rows_job(args: Tuple[str, Optional[bytes], str, bool, bool]) -> Tuple[str, List[List[Any]]]:
    key, data, title, data_only, fast = args
    if fast:
        return title, fastread.read_sheet(data if data is not None else _SHARED[key], title)
    wb = openpyxl.load_workbook(io.BytesIO(data if data is not None else _SHARED[key]), data_only=data_only, read_only=True)
    try:
        ws = wb[title]
        try:
            ws.reset_dimensions()
        except Exception:  # noqa: BLE001
            pass
        return title, [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def _read_sheets_raw(file_bytes: bytes, data_only: bool, workers: int, only: Optional[set] = None,
                     fast: bool = False) -> Dict[str, List[List[Any]]]:
    """Every sheet's cells.  With several workers and a big file the sheets are read by separate processes.
    ``fast`` = values only through fastread (raises fastread.Unsupported when it cannot read the file)."""
    if workers > 1 and len(file_bytes) >= perf.parallel_min_bytes():
        try:
            titles = [t for t in _sheet_titles(file_bytes) if only is None or t in only]
            if len(titles) > 1:
                pool = perf.make_pool(min(workers, len(titles)))
                if pool is not None:
                    key = uuid.uuid4().hex
                    shared = perf.start_method() == "fork"
                    if shared:
                        _SHARED[key] = file_bytes
                    try:
                        jobs = [(key, None if shared else file_bytes, t, data_only, fast) for t in titles]
                        got = dict(perf.map_ordered(pool, _sheet_rows_job, jobs, window=len(jobs)))
                        return {t: got[t] for t in titles}
                    finally:
                        _SHARED.pop(key, None)
                        pool.shutdown(wait=False)
        except Exception:  # noqa: BLE001 - fall back to reading the sheets one after another
            pass
    if fast:
        return fastread.read_sheets(file_bytes, only)
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=data_only, read_only=True)
    out: Dict[str, List[List[Any]]] = {}
    for ws in wb.worksheets:
        if only is not None and ws.title not in only:
            continue
        try:
            ws.reset_dimensions()          # some generators write a wrong <dimension>; recompute it
        except Exception:  # noqa: BLE001
            pass
        out[ws.title] = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return out


def read_workbook(file_bytes: bytes, filename: str, report: Optional[List[Dict[str, Any]]] = None,
                  workers: Optional[int] = None, values_only: bool = False) -> List[Dict[str, Any]]:
    """-> [{sheet, header_row, df, n_formulas}] for every non-empty sheet of an .xlsx/.xlsm/.xltx/.xltm/.xls/.xlsb/.ods
    workbook, a .csv/.tsv/.txt (optionally .gz) or a .parquet file.

    ``report`` (optional list) receives one {sheet, issue} entry per sheet that was skipped or looks wrong."""
    ext = file_ext(filename)
    out: List[Dict[str, Any]] = []
    nw = perf.workers() if workers is None else max(1, int(workers))
    if ext in TEXT_EXT:
        df, hr = _frame_to_df(_read_text(file_bytes, filename))
        if not df.empty:
            out.append({"sheet": "Sheet1", "header_row": hr, "df": df, "n_formulas": 0})
        return out
    if ext == ".parquet":
        try:
            df = pd.read_parquet(io.BytesIO(file_bytes)).astype(object)
        except ImportError:
            raise ValueError("%s: reading .parquet needs the 'pyarrow' Python package; save the data as .csv or .xlsx instead." % filename)
        df.columns = _headers_from_row(list(df.columns))
        df = df.where(df.notna(), None)
        df.attrs["col_offset"], df.attrs["col_idx"] = 0, list(range(df.shape[1]))
        if len(df):
            out.append({"sheet": "Sheet1", "header_row": 0, "df": df, "n_formulas": 0, "col_offset": 0,
                        "col_idx": df.attrs["col_idx"]})
        return out
    if ext in VALUES_EXT:
        eng, pkg = _ENGINES[ext]
        try:
            book = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None, dtype=object, engine=eng)
        except ImportError:
            raise ValueError("%s: reading %s files needs the '%s' Python package (pip install %s) - or save the file as .xlsx." % (filename, ext, pkg, pkg))
        for name, raw in book.items():
            df, hr = _frame_to_df(raw)
            if not df.empty:
                out.append({"sheet": name, "header_row": hr, "df": df, "n_formulas": 0})
                if ext == ".xlsb" and report is not None:
                    report.append({"sheet": name, "issue": "read as values only (a .xlsb keeps no readable formulas); dates in it arrive as plain "
                                   "numbers (Excel day counts) - lock those columns to Date under Data types or save as .xlsx"})
        return out
    # .xlsx family: read once in streaming mode.  Formulas come back as "=..." strings; only if the workbook has any do
    # we read it a second time for the saved results (so plain data files are parsed just once).
    raw = None
    fast_used = False
    if values_only:                                   # a run's input: the cell values are all that is needed -> the fast reader
        try:
            raw = _read_sheets_raw(file_bytes, True, nw, fast=True)
            fast_used = True
        except Exception:  # noqa: BLE001 - anything unusual: the careful (slower) reader below
            raw = None
    if raw is None:
        raw = _read_sheets_raw(file_bytes, False, nw)

    def has_formula(rows):
        return any(isinstance(c, str) and c.startswith("=") for r in rows for c in r)
    needs_values = set() if fast_used else {n for n, rows in raw.items() if has_formula(rows)}
    vals: Dict[str, List[List[Any]]] = _read_sheets_raw(file_bytes, True, nw, needs_values) if needs_values else {}
    for name, frows in raw.items():
        nf = unc = 0
        if name in needs_values:
            vrows = vals.get(name, [])
            for fr, vr in zip(frows, vrows):
                for c, v in zip(fr, vr):
                    if isinstance(c, str) and c.startswith("="):
                        nf += 1
                        unc += v is None
            rows = vrows
        else:
            rows = frows
        grid = None
        if is_toggle_sheet(name):            # name|Yes/No rows: keep the grid so cell addresses (B2) stay exact
            grid = [[jsonable(c) for c in r] for r in rows]
            while grid and not any(c is not None for c in grid[-1]):
                grid.pop()
            if grid and grid[0] and len(grid[0]) > 1 and _yn_like(grid[0][1]):     # no header row at all
                rows = [["Label", "Value"]] + [list(r[:2]) for r in rows]
        df, hr = _rows_to_df(rows)
        if df.empty:
            if report is not None:
                report.append({"sheet": name, "issue": (
                    "skipped: %d formula cell(s) have no saved result, so there are no values to read - open the "
                    "file in Excel, save it, and upload it again" % nf) if nf else "skipped: the sheet is empty"})
            continue
        if unc and report is not None:
            report.append({"sheet": name, "issue": "%d of %d formula cell(s) have no saved result and read as blank "
                           "- open the file in Excel and save it, then upload again" % (unc, nf)})
        fcells: Dict[str, Dict[str, Any]] = {}
        if name in needs_values:
            for j, k in enumerate(df.attrs.get("col_idx", [])):
                for ri in range(hr + 1, len(frows)):
                    v = frows[ri][k] if k < len(frows[ri]) else None
                    if isinstance(v, str) and v.startswith("="):
                        fcells[str(df.columns[j])] = {"raw": v, "row": ri + 1}      # verbatim text, 1-based Excel row
                        break
        out.append({"sheet": name, "header_row": hr, "df": df, "n_formulas": nf, "col_offset": df.attrs.get("col_offset", 0),
                    "col_idx": df.attrs.get("col_idx", []), "formula_cells": fcells, "grid": grid})
    return out


def _file_job(args: Tuple[str, Optional[bytes], str, bool]) -> List[Dict[str, Any]]:
    key, data, name, values_only = args
    return read_workbook(data if data is not None else _SHARED[key], name, workers=1, values_only=values_only)


def read_many(files: List[Tuple[str, bytes]], workers: Optional[int] = None, values_only: bool = False) -> List[List[Dict[str, Any]]]:
    """read_workbook for several files, the big ones at the same time (one process each)."""
    nw = perf.workers() if workers is None else max(1, int(workers))
    big = sum(len(d) >= perf.parallel_min_bytes() for _, d in files)
    if nw > 1 and len(files) > 1 and big:
        try:
            pool = perf.make_pool(min(nw, len(files)))
            if pool is not None:
                keys = [uuid.uuid4().hex for _ in files]
                shared = perf.start_method() == "fork"
                if shared:
                    for k, (_, d) in zip(keys, files):
                        _SHARED[k] = d
                try:
                    jobs = [(k, None if shared else d, n, values_only) for k, (n, d) in zip(keys, files)]
                    return list(perf.map_ordered(pool, _file_job, jobs, window=len(jobs)))
                finally:
                    for k in keys:
                        _SHARED.pop(k, None)
                    pool.shutdown(wait=False)
        except ValueError:
            raise
        except Exception:  # noqa: BLE001 - one file after the other
            pass
    return [read_workbook(d, n, workers=nw, values_only=values_only) for n, d in files]


def _yn_like(v: Any) -> bool:
    return str(v).strip().lower() in ("yes", "no", "y", "n", "true", "false", "1", "0")


def yes_no(raw: Any) -> Optional[str]:
    """Toggle cell value -> 'Yes' / 'No' (None when it isn't a switch value)."""
    s = str(raw).strip().lower() if raw is not None else ""
    if s in ("yes", "y", "true", "1"):
        return "Yes"
    if s in ("no", "n", "false", "0"):
        return "No"
    return None


def guess_role(filename: str, sheet: str, n_formulas: int, is_first: bool) -> str:
    if re.search(r"dictionary|readme|read_me|notes?\b|instructions|legend|glossary|about", sheet.lower()):
        return "ignore"
    s = f"{filename} {sheet}".lower()
    if re.search(r"\b(map|mapping|lookup|rate|rates|ref|reference|master|param|toggle)", s):
        return "mapping"
    return "input"


def df_to_spec(df: pd.DataFrame, header_row: int = 0, col_offset: int = 0, grid: Any = None) -> Dict[str, Any]:
    recs = [{c: jsonable(v) for c, v in zip(df.columns, row)}
            for row in df.itertuples(index=False, name=None)]
    return {"headers": [str(c) for c in df.columns], "records": recs, "header_row": int(header_row), "col_offset": int(col_offset),
            **({"grid": grid} if grid else {})}


def spec_to_df(spec: Dict[str, Any]) -> pd.DataFrame:
    return pd.DataFrame(spec["records"], columns=spec["headers"], dtype=object)


def summarize(df: pd.DataFrame, n_preview: int = 5) -> Dict[str, Any]:
    prev = [[jsonable(v) for v in row] for row in df.head(n_preview).itertuples(index=False, name=None)]
    return {"rows": int(len(df)), "cols": int(df.shape[1]),
            "headers": [str(c) for c in df.columns], "preview": prev}


def canonical_columns(df: pd.DataFrame, wanted: List[str]) -> pd.DataFrame:
    """Rename df columns to the stored names (case/space-insensitive)."""
    low = {nk(c): c for c in wanted}
    ren, seen = {}, set()
    for c in df.columns:
        t = low.get(nk(c))
        if t and t not in seen:
            ren[c] = t
            seen.add(t)
    return df.rename(columns=ren)


_SHEET_REF = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!")


def formula_sheet_refs(file_bytes: bytes, filename: str) -> Dict[str, Any]:
    """Which sheets hold formulas, and which sheets those formulas point at (mapping sheets)."""
    if not holds_formulas(filename):
        return {"with_formulas": set(), "referenced": set()}
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    withf, refd = set(), set()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    withf.add(ws.title)
                    for m in _SHEET_REF.finditer(c.value):
                        refd.add((m.group(1) or m.group(2)).replace("''", "'"))
    return {"with_formulas": withf, "referenced": refd - {ws.title for ws in wb.worksheets if ws.title in withf}}
