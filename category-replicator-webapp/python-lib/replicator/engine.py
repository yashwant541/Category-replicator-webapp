"""The whole product, simply: a *template* learned from the sample, replayed on new input.

For every output sheet of the sample we remember
  * where it starts (header row, first column) and its column order,
  * which columns are INPUT columns (pasted from the input rows that pass the sheet's filter), and
  * which columns are DERIVED (header + the sample's formula, stored verbatim and filled down),
plus the sheet filters (editable), the mapping / toggle sheets (copied along) and any parameter cell a formula points at.

Nothing is calculated or guessed at run time.  Dataiku-free; the backend only moves bytes in and out of folders.
"""

from __future__ import annotations

import datetime as dt
import io
import re
import uuid
import zipfile
from collections import OrderedDict
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from xml.sax.saxutils import escape

from . import rules as R
from .learn import align, build_tree
from .rules import cmp_mask, disp, is_blank, nk
from replicator_formulas.sample_parser import is_toggle_sheet
from .tables import canonical_columns, df_to_spec, jsonable, spec_to_df, yes_no


ENGINE_BUILD = "2026.10a"


SRC = "Source sheet"


_BAD_SHEET = re.compile(r"[\[\]:*?/\\]")


def conditions_mask(df: pd.DataFrame, conds: List[Dict[str, Any]], match: str = "all") -> pd.Series:
    if not conds:
        return pd.Series(True, index=df.index)
    masks = []
    for c in conds:
        col = R.resolve_col(df, c["col"])
        masks.append(cmp_mask(df[col], c["op"], c.get("value")))
    m = masks[0]
    for x in masks[1:]:
        m = (m & x) if match == "all" else (m | x)
    return m


_INV = {">=": "<", "<": ">=", ">": "<=", "<=": ">", "==": "!=", "!=": "=="}


def _tree_to_conditions(tree: Dict[str, Any], keep: str = "Yes") -> Optional[List[Dict[str, Any]]]:
    out: List[Dict[str, Any]] = []
    node = tree
    while "test" in node:
        t = dict(node["test"])
        a, b = node["then"], node["else"]

        def is_leaf(n, val):
            return "leaf" in n and nk(n["leaf"].get("const")) == nk(val)
        if is_leaf(b, "No") and not is_leaf(a, "No"):
            out.append(t)
            node = a
        elif is_leaf(a, "No") and not is_leaf(b, "No"):
            t["op"] = _INV[t["op"]]
            out.append(t)
            node = b
        else:
            return None
    return out if is_leaf_yes(node) else None


def is_leaf_yes(node) -> bool:
    return "leaf" in node and nk(node["leaf"].get("const")) == "yes"


def detect_pre_filter(X: pd.DataFrame, included: np.ndarray) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    if included.all():
        return [], None
    y = pd.Series(np.where(included, "Yes", "No"))
    tree = build_tree(X.reset_index(drop=True), y, max_depth=3, max_leaves=6)
    conds = _tree_to_conditions(tree) if tree else None
    if conds and conditions_mask(X.reset_index(drop=True), conds).eq(included).all():
        return conds, None
    return [], ("Some input rows are not in the output but no simple rule explains which; add an "
                "input filter manually if rows should be dropped.")


def detect_split(work: pd.DataFrame, sheet: pd.Series) -> Tuple[Dict[str, Any], Optional[str]]:
    sheets = list(dict.fromkeys(sheet))
    off = {"enabled": False, "mode": "column", "column": "", "map": {}, "new_values": "create",
           "sheets": [], "rest_sheet": ""}
    if len(sheets) < 2:
        return off, None
    for c in work.columns:
        if c == "__sheet__":
            continue
        keys = work[c].map(nk)
        modal: Dict[str, str] = {}
        ok = True
        for sh in sheets:
            vc = keys[(sheet == sh).values].value_counts()
            if vc.empty or vc.iloc[0] < 0.98 * vc.sum() or vc.index[0] == "":
                ok = False
                break
            modal[sh] = vc.index[0]
        if ok and len(set(modal.values())) == len(sheets):
            mp = {}
            for sh in sheets:
                row = int(np.argmax(((sheet == sh).values) & (keys == modal[sh]).values))
                mp[str(jsonable(work[c].iloc[row]) if not isinstance(work[c].iloc[row], str)
                       else work[c].iloc[row])] = sh
            return {"enabled": True, "mode": "column", "column": c, "map": mp,
                    "new_values": "create", "sheets": [], "rest_sheet": ""}, None
    cust = dict(off)
    cust.update(enabled=True, mode="custom",
                sheets=[{"name": sh, "match": "all", "conditions": []} for sh in sheets])
    return cust, ("The output has several sheets but no single column explains the split - define "
                  "each sheet's filter under Filters.")


def _route_conditions(full: pd.DataFrame, mask: np.ndarray, sheet: str) -> Optional[List[Dict[str, Any]]]:
    best = None
    for pos, c in enumerate(full.columns):
        ks = full[c].map(nk)
        inc, exc = set(ks[mask]), set(ks[~mask])
        if not inc or "" in inc or inc & exc:
            continue
        exact = inc == {nk(sheet)}
        score = (0 if exact else 1, len(inc), pos)
        if best is None or score < best[0]:
            orig = [next(v for v, k in zip(full[c], ks) if k == i) for i in sorted(inc)]
            cond = {"col": c, "op": "==", "value": jsonable(orig[0])} if len(orig) == 1 else \
                {"col": c, "op": "in", "value": [str(jsonable(o)) for o in orig]}
            best = (score, [cond])
    if best:
        return best[1]
    tree = build_tree(full, pd.Series(np.where(mask, "Yes", "No")), max_depth=3, max_leaves=6)
    conds = _tree_to_conditions(tree) if tree else None
    return conds


def _stem(fn: str) -> str:
    return re.sub(r"\.[^.]+$", "", fn or "")


def _unique_name(base: str, taken: set) -> str:
    n, i = base, 2
    while n in taken:
        n = f"{base}_{i}"
        i += 1
    return n


_REF = re.compile(r"(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_!])")


_LITERALS = re.compile(r"(\"(?:[^\"]|\"\")*\"|'(?:[^']|'')*')")


def shift_formula(f: str, from_row: int, to_row: int) -> str:
    """The sample's formula text for another row - exactly what Excel does when you fill a formula down:
    row numbers without a $ move by the row difference; $-anchored refs, ranges, sheet names and text stay."""
    d = to_row - from_row
    if d == 0 or not f:
        return f
    parts = _LITERALS.split(f)
    for i in range(0, len(parts), 2):          # odd parts are "text" / 'Sheet name' literals
        parts[i] = _REF.sub(lambda m: m.group(0) if m.group(3) == "$" or int(m.group(4)) + d < 1
                            else "%s%s%s%d" % (m.group(1), m.group(2), m.group(3), int(m.group(4)) + d), parts[i])
    return "".join(parts)


def _cell_addr(a1: str) -> Tuple[int, int]:
    m = re.match(r"^\$?([A-Za-z]{1,3})\$?(\d+)$", a1.strip())
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)) - 1, col - 1


def effective_toggles(system: Dict[str, Any], frames: Optional[Dict[str, pd.DataFrame]] = None,
                      overrides: Optional[Dict[str, str]] = None) -> List[Dict[str, Any]]:
    """Toggle values to use *now*: the Toggle sheet of the (stored or freshly attached) mapping wins over the
    value seen in the example; an explicit per-run override wins over everything."""
    frames = frames if frames is not None else tables_to_frames(system)
    out = []
    for t in system.get("toggles", []):
        t = dict(t)
        val = None
        if t.get("sheet"):
            tname = next((n for n in frames if nk(n) == nk(t["sheet"])), None)
            if tname is not None:
                df = frames[tname]
                base = re.sub(r"\s*\(\d+\)$", "", t["name"])
                if df.shape[1] >= 2:                       # label | value rows
                    hit = df[df.iloc[:, 0].map(nk).isin([nk(t["name"]), nk(base)])]
                    if len(hit):
                        val = yes_no(hit.iloc[0, 1])
                spec = system.get("tables", {}).get(tname) or {}
                if val is None and spec.get("grid") and t.get("cell"):    # exact cell of the stored grid
                    r, c = _cell_addr(t["cell"])
                    g = spec["grid"]
                    if r < len(g) and c < len(g[r]):
                        val = yes_no(g[r][c])
        if overrides and t["name"] in overrides:
            val = yes_no(overrides[t["name"]]) or val
        if val:
            t["value"] = val
        out.append(t)
    return out


def tables_to_frames(system: Dict[str, Any], override: Optional[Dict[str, pd.DataFrame]] = None
                     ) -> Dict[str, pd.DataFrame]:
    t = {n: spec_to_df(s) for n, s in system.get("tables", {}).items()}
    for n, df in (override or {}).items():
        t[n] = df
    return t


def prepare_input(system: Dict[str, Any], df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    wanted = system["input"]["columns"]
    df = canonical_columns(df.reset_index(drop=True), wanted)
    notes = []
    if SRC in wanted and SRC not in df.columns:
        df[SRC] = "Input"
        notes.append("The input had a single sheet, so no split by input sheet is possible (sheet name 'Input' used).")
    have = set(df.columns)
    missing_req = [c for c in system["input"].get("required", []) if c not in have]
    if missing_req:
        raise ValueError("Input is missing required column(s): " + ", ".join(missing_req) +
                         ". Expected columns: " + ", ".join(wanted))
    for c in wanted:
        if c not in have:
            df[c] = np.nan
            notes.append(f"Input column '{c}' not found - treated as blank.")
    extra = [c for c in df.columns if c not in set(wanted)]
    if extra:
        notes.append("Ignored extra input column(s): " + ", ".join(map(str, extra[:8])))
    return df, notes


def safe_sheet_name(name: str, taken: set) -> str:
    base = _BAD_SHEET.sub("-", str(name)).strip().strip("'") or "Sheet"
    base = base[:31]
    n, i = base, 2
    while n.lower() in {t.lower() for t in taken}:
        suf = f" ({i})"
        n = base[:31 - len(suf)] + suf
        i += 1
    taken.add(n)
    return n


def split_frames(system: Dict[str, Any], out: pd.DataFrame, work: pd.DataFrame
                 ) -> Tuple["OrderedDict[str, pd.DataFrame]", List[str]]:
    sp = system.get("filters", {}).get("split", {})
    notes: List[str] = []
    taken: set = set()
    res0: "OrderedDict[str, pd.DataFrame]" = OrderedDict()
    res = res0
    sc = system["output"].get("sheet_columns") or {}
    if not sp.get("enabled"):
        nm = system["output"].get("sheet_name") or "Output"
        res[safe_sheet_name(nm, taken)] = out[[c for c in sc[nm] if c in out.columns]] if nm in sc else out
        return res, notes
    common = [c for c in system["output"]["columns"] if all(c in v for v in sc.values())] if sc else []

    def pick(name: str, df: pd.DataFrame) -> pd.DataFrame:
        """Each sheet shows its own columns; sheets the system never saw get the shared ones."""
        cols = sc.get(name) or (common if sc and common else None)
        return df[[c for c in cols if c in df.columns]] if cols else df
    if sp.get("mode") == "column":
        col = R.resolve_col(work, sp["column"])
        keys = work[col].map(nk)
        mapping = {nk(k): v for k, v in sp.get("map", {}).items()}
        sheet_of = []
        unassigned = 0
        for v, k in zip(work[col], keys):
            if k == "":
                sheet_of.append(None)
                unassigned += 1
            elif k in mapping:
                sheet_of.append(mapping[k])
            elif sp.get("new_values", "create") == "create":
                sheet_of.append(disp(v))
            else:
                sheet_of.append(None)
                unassigned += 1
        s = pd.Series(sheet_of, index=out.index, dtype=object)
        order = [v for v in dict.fromkeys(mapping.values()) if (s == v).any()] + \
                [v for v in dict.fromkeys(x for x in s if x is not None) if v not in set(mapping.values())]
        for nm in order:
            res[safe_sheet_name(nm, taken)] = out[(s == nm).values]
        if unassigned:
            notes.append(f"{unassigned} row(s) had a blank/unknown '{sp['column']}' and were not placed in a split sheet"
                         + (" (put in 'Unassigned')." if sp.get("rest_sheet") else "."))
            if sp.get("rest_sheet"):
                res[safe_sheet_name(sp["rest_sheet"] or "Unassigned", taken)] = out[s.isna().values]
        if not res:
            res[safe_sheet_name("Output", taken)] = out
        return OrderedDict((k, pick(k, v)) for k, v in res.items()), notes
    covered = pd.Series(False, index=out.index)
    for sh in sp.get("sheets", []):
        m = conditions_mask(work, sh.get("conditions", []), sh.get("match", "all"))
        covered |= m.values
        res[safe_sheet_name(sh["name"], taken)] = out[m.values]
    if sp.get("rest_sheet"):
        res[safe_sheet_name(sp["rest_sheet"], taken)] = out[~covered.values]
    elif (~covered).any():
        notes.append(f"{int((~covered).sum())} row(s) matched no sheet filter and were left out.")
    if not res:
        res[safe_sheet_name("Output", taken)] = out
    return OrderedDict((k, pick(k, v)) for k, v in res.items()), notes


HDR_FILL = PatternFill("solid", fgColor="1F3A5F")


HDR_FONT = Font(bold=True, color="FFFFFF")


HELP_FILL = PatternFill("solid", fgColor="E8EDF3")


def xl_val(v: Any) -> Any:
    if v is None or v is pd.NaT:
        return None
    if isinstance(v, float) and np.isnan(v):
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return None if np.isnan(v) else float(v)
    if isinstance(v, np.bool_):
        return bool(v)
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if isinstance(v, str) and v.startswith("="):
        return "'" + v
    return v


def _autosize(ws, header_len: List[int]) -> None:
    for i, w in enumerate(header_len, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(8, min(48, w + 2))


def _like(orig: Any, yn: str) -> Any:
    """Write a Yes/No switch in the same style the cell had (TRUE/FALSE, true/false, Yes/No, 1/0)."""
    on = yn == "Yes"
    if isinstance(orig, bool):
        return on
    if isinstance(orig, (int, float)):
        return 1 if on else 0
    t = str(orig or "")
    word = ("TRUE", "FALSE") if t.lower() in ("true", "false") else (("Y", "N") if t.lower() in ("y", "n") else ("Yes", "No"))
    w = word[0] if on else word[1]
    return w.upper() if t.isupper() else (w.lower() if t.islower() else w)


def inject_cached(blob: bytes, sheet_names: List[str], cached: Dict[str, Dict[str, Any]]) -> bytes:
    """openpyxl writes formulas without their results; add them so the file shows numbers in any viewer
    (Excel recalculates on open anyway)."""
    import zipfile
    from xml.sax.saxutils import escape
    zin = zipfile.ZipFile(io.BytesIO(blob))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = re.match(r"xl/worksheets/sheet(\d+)\.xml$", item.filename)
            if m and int(m.group(1)) <= len(sheet_names):
                vals = cached.get(sheet_names[int(m.group(1)) - 1]) or {}

                def fix(mm):
                    coord, attrs, f = mm.group(1), mm.group(2), mm.group(3)
                    v = vals.get(coord)
                    if v is None or (isinstance(v, float) and np.isnan(v)) or isinstance(v, (dt.date, pd.Timestamp)):
                        return mm.group(0)
                    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
                    if isinstance(v, (bool, np.bool_)):
                        return '<c r="%s"%s t="b"><f>%s</f><v>%d</v></c>' % (coord, attrs, f, int(v))
                    if isinstance(v, (int, float, np.integer, np.floating)):
                        return '<c r="%s"%s><f>%s</f><v>%s</v></c>' % (coord, attrs, f, repr(float(v)) if not float(v).is_integer() else int(v))
                    return '<c r="%s"%s t="str"><f>%s</f><v>%s</v></c>' % (coord, attrs, f, escape(str(v)))
                text = re.sub(r'<c r="([A-Z]+\d+)"([^>]*)><f>(.*?)</f><v\s*/?>(?:</v>)?</c>', fix, data.decode("utf-8"))
                data = text.encode("utf-8")
            zout.writestr(item, data)
    return out.getvalue()


# =============================================================== learning the template
_GENERIC = {"label", "name", "value", "toggle", "toggles", "key", "parameter", "param", "setting", "switch", "flag"}
_PARAM_REF = re.compile(r"(?<![A-Za-z0-9_.!'$:])\$([A-Z]{1,3})\$(\d+)(?![\d:(])")
_SHEET_REFS = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!")


def toggles_from_tables(tables: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Every Yes/No (TRUE/FALSE, 1/0) row of a sheet called Toggle / Toggles... is a switch."""
    out: List[Dict[str, Any]] = []
    for name, spec in tables.items():
        if not is_toggle_sheet(name):
            continue
        for ri, row in enumerate(spec.get("grid") or []):
            if len(row) > 1 and row[0] is not None and str(row[0]).strip().lower() not in _GENERIC and yes_no(row[1]):
                label, n = str(row[0]).strip(), 2
                while any(t["name"] == label for t in out):
                    label = "%s (%d)" % (str(row[0]).strip(), n)
                    n += 1
                out.append({"name": label, "value": yes_no(row[1]), "sheet": name, "cell": "B%d" % (ri + 1)})
    return out


def _param_cells(e: Dict[str, Any], cols: List[Dict[str, Any]], hr: int) -> List[Dict[str, Any]]:
    """Single cells like $G$1 that a sample formula points at on its own sheet but that are not part of the table."""
    cells = []
    idx = [c["col"] for c in cols]
    for c in cols:
        if c["kind"] != "formula":
            continue
        for m in _PARAM_REF.finditer(re.sub(r'"(?:[^"]|"")*"', "", c["raw"])):
            col = 0
            for ch in m.group(1):
                col = col * 26 + ord(ch) - 64
            col -= 1
            row = int(m.group(2))
            inside = (min(idx) <= col <= max(idx)) and row > hr + 1
            addr = "%s%d" % (m.group(1), row)
            if not inside and addr not in [x["cell"] for x in cells]:
                cells.append({"cell": addr, "row": row - 1, "col": col})
    if cells and e.get("bytes"):
        try:
            ws = openpyxl.load_workbook(io.BytesIO(e["bytes"]), data_only=True)[e["sheet"]]
            for x in cells:
                x["value"] = jsonable(ws.cell(row=x["row"] + 1, column=x["col"] + 1).value)
        except Exception:  # noqa: BLE001
            pass
    return cells


def route_sheets(X: pd.DataFrame, masks: Dict[str, np.ndarray], parts: List[pd.DataFrame], labels: List[str]
                 ) -> Tuple[Dict[str, Any], Optional[str]]:
    """Which input rows does each sample sheet hold?  Expressed as filters on INPUT columns only."""
    off = {"enabled": False, "mode": "column", "column": "", "map": {}, "new_values": "create", "sheets": [], "rest_sheet": ""}
    if len(masks) < 2:
        return off, None
    vals = list(masks.values())
    disjoint = not any((a & b).any() for i, a in enumerate(vals) for b in vals[i + 1:])
    if disjoint:
        sp, _ = detect_split(pd.concat(parts, ignore_index=True), pd.Series(labels))
        if sp.get("enabled") and sp.get("mode") == "column":
            return sp, None
    sp = dict(off, enabled=True, mode="custom")
    miss = []
    for name, m in masks.items():
        conds: Optional[List[Dict[str, Any]]] = [] if m.all() else _route_conditions(X, m, name)
        if conds is None:
            miss.append(name)
            conds = []
        sp["sheets"].append({"name": name, "match": "all", "conditions": conds})
    return sp, (("Could not work out which rows belong on sheet(s) " + ", ".join(miss) +
                 " - set their filters under Filters & outputs.") if miss else None)


def build_system(entries: List[Dict[str, Any]], options: Dict[str, Any]) -> Dict[str, Any]:
    """``entries``: [{file, sheet, role, df, header_row, col_offset, col_idx, formula_cells, grid, bytes}]."""
    warnings: List[str] = []
    ins = [e for e in entries if e["role"] == "input"]
    maps = [e for e in entries if e["role"] == "mapping"]
    outs = [e for e in entries if e["role"] == "output"]
    if not ins:
        raise ValueError("Add at least one input sheet.")
    if not outs:
        raise ValueError("Add the expected output sheet(s).")

    # ---- the input: sheets with the same columns are stacked; the biggest stack is the data
    groups: List[List[Dict[str, Any]]] = []
    for e in ins:
        for g in groups:
            if [nk(c) for c in g[0]["df"].columns] == [nk(c) for c in e["df"].columns]:
                g.append(e)
                break
        else:
            groups.append([e])
    groups.sort(key=lambda g: -sum(len(x["df"]) for x in g))
    primary = groups[0]
    names = [e["sheet"] for e in primary]
    parts = []
    for e in primary:
        f = canonical_columns(e["df"], list(primary[0]["df"].columns)).copy()
        if len(primary) > 1:      # the input's own sheets are a split: remember which sheet a row came from
            f[SRC] = e["sheet"] if names.count(e["sheet"]) == 1 else "%s/%s" % (_stem(e["file"]), e["sheet"])
        parts.append(f)
    X = pd.concat(parts, ignore_index=True)
    if len(groups) > 1:
        warnings.append("Input sheets with a different layout were kept as reference tables: "
                        + ", ".join(x["sheet"] for g in groups[1:] for x in g))

    # ---- mapping / toggle sheets (uploaded ones win over copies inside the output workbook)
    out_files = {o["file"] for o in outs}
    tables: Dict[str, Dict[str, Any]] = {}
    table_src: Dict[str, str] = {}
    taken: set = set()
    for e in sorted(maps, key=lambda e: e["file"] in out_files) + [x for g in groups[1:] for x in g]:
        dup = next((t for t in tables if nk(t) == nk(e["sheet"])), None)
        if e["file"] in out_files and dup:
            if tables[dup]["headers"] != [str(c) for c in e["df"].columns] or len(tables[dup]["records"]) != len(e["df"]):
                warnings.append(f"Mapping '{e['sheet']}' in the output file differs from the uploaded mapping; "
                                "the uploaded mapping is used.")
            continue
        nm = _unique_name(e["sheet"] if e["sheet"] not in taken else f"{_stem(e['file'])}_{e['sheet']}", taken)
        taken.add(nm)
        tables[nm] = df_to_spec(e["df"], e.get("header_row", 0), e.get("col_offset", 0), e.get("grid"))
        table_src[nm] = f"{e['file']} / {e['sheet']}"

    # ---- one template per output sheet
    in_names = {nk(c): str(c) for c in X.columns if c != SRC}
    sheets: List[Dict[str, Any]] = []
    masks: Dict[str, np.ndarray] = {}
    xs_parts: List[pd.DataFrame] = []
    labels: List[str] = []
    for e in outs:
        df = e["df"]
        fc = e.get("formula_cells") or {}
        idx = e.get("col_idx") or list(range(e.get("col_offset", 0), e.get("col_offset", 0) + df.shape[1]))
        cols: List[Dict[str, Any]] = []
        for j, c in enumerate(df.columns):
            c = str(c)
            if c in fc:
                cols.append({"name": c, "col": int(idx[j]), "kind": "formula", "raw": fc[c]["raw"], "sample_row": int(fc[c]["row"])})
            elif nk(c) in in_names:
                cols.append({"name": c, "col": int(idx[j]), "kind": "input", "source": in_names[nk(c)]})
            else:
                cols.append({"name": c, "col": int(idx[j]), "kind": "empty"})
        empty = [c["name"] for c in cols if c["kind"] == "empty"]
        if empty:
            warnings.append(f"Sheet '{e['sheet']}': no formula in the sample for {', '.join(empty)} - these columns stay empty.")
        hr = int(e.get("header_row", 0))
        in_cols = [c["name"] for c in cols if c["kind"] == "input"]
        if not in_cols:
            raise ValueError(f"Sheet '{e['sheet']}': none of its columns match a column of the input.")
        try:
            Xs, m, _ = align(X, df[in_cols])
        except ValueError as exc:
            raise ValueError(f"Sheet '{e['sheet']}': {exc}")
        masks[e["sheet"]] = m
        xs_parts.append(Xs)
        labels += [e["sheet"]] * len(Xs)
        sheets.append({"name": e["sheet"], "origin": {"row": hr, "col": int(min(c["col"] for c in cols))},
                       "columns": cols, "params": _param_cells(e, cols, hr), "sample_rows": int(len(df))})

    # ---- filters: the input filter + which rows each sheet holds (the user can edit both)
    flt: Dict[str, Any] = {"pre": [], "split": {"enabled": False, "mode": "column", "column": "", "map": {},
                                                "new_values": "create", "sheets": [], "rest_sheet": ""}}
    included = np.zeros(len(X), bool)
    for m in masks.values():
        included |= m
    pre, pre_warn = detect_pre_filter(X, included)
    flt["pre"] = pre
    if pre_warn:
        warnings.append(pre_warn)
    sp_opt = options.get("split") or {}
    mode = sp_opt.get("mode", "auto")
    if mode == "none":
        pass
    elif mode == "column" and sp_opt.get("column"):
        col = sp_opt["column"]
        vals = list(dict.fromkeys(X[col].map(lambda v: v if isinstance(v, str) else jsonable(v))))
        flt["split"].update(enabled=True, mode="column", column=col, map={str(v): disp(v) for v in vals if not is_blank(v)})
    elif mode == "custom":
        flt["split"].update(enabled=True, mode="custom", sheets=sp_opt.get("sheets", []))
    else:
        sp, sw = route_sheets(X, masks, xs_parts, labels)
        flt["split"] = sp
        if sw:
            warnings.append(sw)

    system = {
        "engine": ENGINE_BUILD,
        "id": options.get("id") or uuid.uuid4().hex[:10],
        "name": options.get("name") or "Untitled system",
        "description": options.get("description", ""),
        "created": dt.datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "input": {"columns": [str(c) for c in X.columns], "sources": [{"file": e["file"], "sheet": e["sheet"], "rows": int(len(e["df"]))} for e in primary]},
        "tables": tables, "table_sources": table_src, "toggles": toggles_from_tables(tables),
        "output": {"sheet_name": outs[0]["sheet"],
                   "extra_tables": [e["sheet"] for e in maps if e["file"] in out_files]},
        "sheets": sheets, "filters": flt, "warnings": warnings,
    }
    finalize_system(system)
    return system


def finalize_system(system: Dict[str, Any]) -> None:
    """Input columns the filters need (a new input must have them); everything else is pasted if present."""
    incols = {nk(c): c for c in system["input"]["columns"]}
    required: List[str] = []

    def need(c):
        if c and nk(c) in incols and incols[nk(c)] not in required:
            required.append(incols[nk(c)])
    f = system.get("filters", {})
    for c in f.get("pre", []):
        need(c["col"])
    sp = f.get("split", {})
    if sp.get("enabled"):
        if sp.get("mode") == "column":
            need(sp.get("column"))
        for s in sp.get("sheets", []):
            for c in s.get("conditions", []):
                need(c["col"])
    system["input"]["required"] = required


# ======================================================================= applying it
def _sheet_frame(d: Dict[str, Any], rows: pd.DataFrame) -> pd.DataFrame:
    """The sheet as it will be written: pasted input values + the sample's formulas filled down (nothing else)."""
    first = d["origin"]["row"] + 2
    data: "OrderedDict[str, List[Any]]" = OrderedDict()
    n = len(rows)
    for c in d["columns"]:
        if c["kind"] == "input":
            data[c["name"]] = list(rows[c["source"]]) if c["source"] in rows.columns else [None] * n
        elif c["kind"] == "formula":
            data[c["name"]] = [shift_formula(c["raw"], c["sample_row"], first + i) for i in range(n)]
        else:
            data[c["name"]] = [None] * n
    return pd.DataFrame(data, columns=[c["name"] for c in d["columns"]], dtype=object)


def apply_system(system: Dict[str, Any], df_in: pd.DataFrame,
                 table_override: Optional[Dict[str, pd.DataFrame]] = None,
                 toggle_override: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
    df, notes = prepare_input(system, df_in)
    n_in = len(df)
    pre = system.get("filters", {}).get("pre", [])
    if pre:
        df = df[conditions_mask(df, pre).values].reset_index(drop=True)
        notes.append(f"Input filter kept {len(df)} of {n_in} rows.")
    frames = tables_to_frames(system, table_override)
    toggles_now = effective_toggles(system, frames, toggle_override)
    rows, snotes = split_frames(system, df, df)
    notes += snotes
    defs = {s["name"]: s for s in system["sheets"]}
    sheets: "OrderedDict[str, Dict[str, Any]]" = OrderedDict()
    for name, rdf in rows.items():
        d = defs.get(name)
        if d is None:
            d = system["sheets"][0]
            notes.append(f"Sheet '{name}' is new: laid out like '{d['name']}'.")
        sheets[name] = {"df": _sheet_frame(d, rdf.reset_index(drop=True)), "def": d}
    return {"sheets": sheets, "notes": notes, "toggles": toggles_now, "input_rows": n_in,
            "rows": int(sum(len(v["df"]) for v in sheets.values()))}


# ================================================================ the output workbook
def _write_tables(wb, system, tables, toggles, table_sheet):
    for t, sn in table_sheet.items():
        ws = wb.create_sheet(sn)
        T = tables[t]
        spec = system.get("tables", {}).get(t, {})
        if spec.get("grid"):                          # Toggle sheet: exact cell layout, current switch values
            grid = [list(r) for r in spec["grid"]]
            for tg in toggles:
                if tg.get("cell") and nk(tg.get("sheet", "")) == nk(t) and tg.get("value"):
                    r, c = _cell_addr(tg["cell"])
                    if r < len(grid) and c < len(grid[r]):
                        grid[r][c] = _like(grid[r][c], tg["value"])
            for row in grid:
                ws.append([xl_val(v) for v in row])
            continue
        pad = [None] * int(spec.get("col_offset", 0))
        for _ in range(int(spec.get("header_row", 0))):
            ws.append([])
        hr = int(spec.get("header_row", 0)) + 1
        ws.append(pad + [str(c) for c in T.columns])
        for row in T.itertuples(index=False, name=None):
            ws.append(pad + [xl_val(v) for v in row])
        for j in range(len(pad) + 1, len(pad) + T.shape[1] + 1):
            ws.cell(row=hr, column=j).font = Font(bold=True)


def write_workbook(system: Dict[str, Any], sheets: "OrderedDict[str, Dict[str, Any]]",
                   summary: Optional[List[Tuple[str, Any]]] = None,
                   tables: Optional[Dict[str, pd.DataFrame]] = None,
                   toggles: Optional[List[Dict[str, Any]]] = None) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    tables = tables if tables is not None else tables_to_frames(system)
    toggles = toggles if toggles is not None else system.get("toggles", [])
    for sname, item in sheets.items():
        df, d = item["df"], item["def"]
        ws = wb.create_sheet(sname)
        hr = int(d["origin"]["row"])
        for c, name in zip(d["columns"], df.columns):          # same cells as the sample
            cell = ws.cell(row=hr + 1, column=c["col"] + 1, value=name)
            cell.font = HDR_FONT
            cell.fill = HDR_FILL if c["kind"] != "formula" else PatternFill("solid", fgColor="2F6B4F")
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(c["col"] + 1)].width = max(9, min(40, len(str(name)) + 3))
        for ri, row in enumerate(df.itertuples(index=False, name=None), start=hr + 2):
            for c, v in zip(d["columns"], row):
                if v is not None and not (isinstance(v, float) and np.isnan(v)):
                    # a formula is written as a formula; pasted input text starting with "=" stays text
                    ws.cell(row=ri, column=c["col"] + 1, value=v if c["kind"] == "formula" else xl_val(v))
        for p in d.get("params", []):                          # parameter cell(s) the formulas point at ($G$1 ...)
            if p.get("value") is not None:
                ws.cell(row=p["row"] + 1, column=p["col"] + 1, value=xl_val(p["value"]))
        ws.freeze_panes = ws.cell(row=hr + 2, column=1).coordinate

    # mapping and toggle sheets are dependencies of the formulas: always part of the output workbook
    taken = set(sheets.keys()) | {"Run summary"}
    dependent: List[str] = list(tables)                 # every mapping + toggle sheet of the system comes along
    dependent += [tg["sheet"] for tg in toggles if tg.get("sheet")]
    for s in system["sheets"]:
        for c in s["columns"]:
            if c["kind"] == "formula":
                dependent += [(m.group(1) or m.group(2)).replace("''", "'") for m in _SHEET_REFS.finditer(c["raw"])]
    table_sheet: Dict[str, str] = {}
    for t in dict.fromkeys(dependent):
        t = next((n for n in tables if nk(n) == nk(t)), None)
        if t is not None and t not in table_sheet:
            table_sheet[t] = safe_sheet_name(t, taken)
    _write_tables(wb, system, tables, toggles, table_sheet)

    sm = wb.create_sheet("Run summary")
    sm.append(["System", system["name"]])
    sm.append(["Generated", dt.datetime.now().strftime("%Y-%m-%d %H:%M")])
    sm.append([])
    sm.append(["Sheet", "Rows"])
    for c in sm[4]:
        c.font = Font(bold=True)
    for sname, item in sheets.items():
        sm.append([sname, int(len(item["df"]))])
    if summary:
        sm.append([])
        for k, v in summary:
            sm.append([k, v if not isinstance(v, (list, dict)) else "; ".join(map(str, v))])
    sm.column_dimensions["A"].width = 30
    sm.column_dimensions["B"].width = 90
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
