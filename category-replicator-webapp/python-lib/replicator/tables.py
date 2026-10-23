"""Reading uploaded files into tidy DataFrames and (de)serialising mapping tables."""

from __future__ import annotations

import datetime as dt
import io
import math
import re
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import openpyxl
import pandas as pd

from replicator_formulas.sample_parser import detect_header_row, _headers_from_row, _norm
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


def _rows_to_df(rows: List[List[Any]]) -> Tuple[pd.DataFrame, int]:
    rows = [[None if (isinstance(c, float) and math.isnan(c)) else c for c in r] for r in rows]
    if not any(any(not is_blank(c) for c in r) for r in rows):
        return pd.DataFrame(), 0
    hr = detect_header_row(rows)
    raw_hdr = rows[hr]
    body = rows[hr + 1:]
    width = max([len(raw_hdr)] + [len(r) for r in body[:2000]]) if body else len(raw_hdr)
    raw_hdr = list(raw_hdr) + [None] * (width - len(raw_hdr))
    keep = [j for j in range(width)
            if _norm(raw_hdr[j]) or any(j < len(r) and not is_blank(r[j]) for r in body)]
    headers = _headers_from_row([raw_hdr[j] for j in keep])
    data = [[(r[j] if j < len(r) else None) for j in keep] for r in body
            if any(not is_blank(c) for c in r)]
    df = pd.DataFrame(data, columns=headers, dtype=object)
    return df, hr


def read_workbook(file_bytes: bytes, filename: str) -> List[Dict[str, Any]]:
    """-> [{sheet, header_row, df, n_formulas}] for every non-empty sheet."""
    low = (filename or "").lower()
    out: List[Dict[str, Any]] = []
    if low.endswith((".csv", ".txt")):
        raw = pd.read_csv(io.BytesIO(file_bytes), header=None, dtype=object,
                          keep_default_na=False, skip_blank_lines=False, encoding_errors="replace")
        df, hr = _rows_to_df(raw.replace("", None).values.tolist())
        if not df.empty:
            out.append({"sheet": "Sheet1", "header_row": hr, "df": df, "n_formulas": 0})
        return out
    if low.endswith(".xls"):
        book = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None,
                             dtype=object, engine="xlrd")
        for name, raw in book.items():
            df, hr = _rows_to_df(raw.values.tolist())
            if not df.empty:
                out.append({"sheet": name, "header_row": hr, "df": df, "n_formulas": 0})
        return out
    wb_v = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    wb_f = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    for name in wb_v.sheetnames:
        rows = [list(r) for r in wb_v[name].iter_rows(values_only=True)]
        df, hr = _rows_to_df(rows)
        if df.empty:
            continue
        nf = sum(1 for r in wb_f[name].iter_rows()
                 for c in r if isinstance(c.value, str) and c.value.startswith("="))
        out.append({"sheet": name, "header_row": hr, "df": df, "n_formulas": nf})
    return out


def guess_role(filename: str, sheet: str, n_formulas: int, is_first: bool) -> str:
    s = f"{filename} {sheet}".lower()
    if re.search(r"\b(map|mapping|lookup|rate|rates|ref|reference|master|param|toggle)", s):
        return "mapping"
    return "input"


def df_to_spec(df: pd.DataFrame, header_row: int = 0) -> Dict[str, Any]:
    recs = [{c: jsonable(v) for c, v in zip(df.columns, row)}
            for row in df.itertuples(index=False, name=None)]
    return {"headers": [str(c) for c in df.columns], "records": recs, "header_row": int(header_row)}


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
