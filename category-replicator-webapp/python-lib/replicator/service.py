"""Application services on top of a store: drafts, systems, runs.  No Flask, no Dataiku."""

from __future__ import annotations

import datetime as dt
import io
import threading
import re
import uuid
from collections import OrderedDict
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from . import engine, formats, perf
from .rules import disp, is_blank, nk
from .store import delete_prefix, get_json, put_json
from .tables import (file_ext, formula_sheet_refs, guess_role, holds_formulas, jsonable, read_many, read_workbook, summarize,
                     supported, canonical_columns)

_DRAFT_LOCK = threading.RLock()   # serialises the draft record update (uploads can overlap)
SUPPORTED_TEXT = ".xlsx, .xlsm, .xlsb, .xls, .xltx, .xltm, .ods, .csv, .tsv, .txt, .gz or .parquet"
ROLES = ("input", "mapping", "output", "ignore")


def _now() -> str:
    return dt.datetime.utcnow().isoformat(timespec="seconds") + "Z"


def _slug(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", s)[:60] or "file"


# ====================================================================== drafts
def new_draft(store) -> str:
    did = uuid.uuid4().hex[:10]
    put_json(store, f"drafts/{did}/draft.json", {"id": did, "created": _now(), "files": []})
    return did


def get_draft(store, did: str) -> Dict[str, Any]:
    d = get_json(store, f"drafts/{did}/draft.json")
    if d is None:
        raise KeyError("draft not found")
    return d


def add_draft_files(store, did: str, files: List[Tuple[str, bytes]], hint: str = "") -> Dict[str, Any]:
    get_draft(store, did)                      # fail early if the draft is unknown
    entries = []
    for name, data in files:                   # slow part (parsing) runs outside the lock
        if not supported(name):
            raise ValueError(f"{name}: unsupported file type (use {SUPPORTED_TEXT})")
        if hint == "output" and not holds_formulas(name):
            raise ValueError(f"{name}: the expected-output sample must be .xlsx / .xlsm (a {file_ext(name)} file keeps no readable "
                             "formulas). Open it in Excel and Save As .xlsx.")
        report: List[Dict[str, Any]] = []
        sheets = read_workbook(data, name, report)
        if not sheets:
            why = "; ".join("%s %s" % (r["sheet"], r["issue"]) for r in report)
            raise ValueError(f"{name}: no data found" + (f" ({why})" if why else ""))
        fid = uuid.uuid4().hex[:8]
        path = f"drafts/{did}/files/{fid}_{_slug(name)}"
        store.put(path, data)
        fr = formula_sheet_refs(data, name) if hint == "output" else {"referenced": set()}
        ent = {"fid": fid, "name": name, "path": path, "sheets": [], "issues": report}
        for i, s in enumerate(sheets):
            role = hint if hint in ROLES and hint != "ignore" else guess_role(name, s["sheet"], s["n_formulas"], i == 0)
            if guess_role(name, s["sheet"], s["n_formulas"], i == 0) == "ignore" and hint in ("input", "mapping"):
                role = "ignore"     # data dictionaries / notes sheets are not data
            if hint == "output":
                role = "output"
                # an output workbook often carries the mapping sheets its formulas use
                if fr["referenced"] and s["sheet"] in fr["referenced"]:
                    role = "mapping"
                elif any(x["n_formulas"] for x in sheets) and not s["n_formulas"]                         and guess_role("", s["sheet"], 0, False) == "mapping":
                    role = "mapping"      # value-only sheet called Rates / Lookup / Mapping next to formula sheets
            ent["sheets"].append({"sheet": s["sheet"], "role": role, "n_formulas": s["n_formulas"],
                                  "header_row": s["header_row"], **summarize(s["df"])})
        entries.append(ent)
    with _DRAFT_LOCK:                          # tiny read-modify-write: never lose a concurrent upload
        d = get_draft(store, did)
        d["files"].extend(entries)
        put_json(store, f"drafts/{did}/draft.json", d)
    return d


def set_draft_roles(store, did: str, roles: Dict[str, str]) -> Dict[str, Any]:
    with _DRAFT_LOCK:
        d = get_draft(store, did)
        for f in d["files"]:
            for s in f["sheets"]:
                k = f"{f['fid']}|{s['sheet']}"
                if k in roles and roles[k] in ROLES:
                    s["role"] = roles[k]
        put_json(store, f"drafts/{did}/draft.json", d)
        return d


def remove_draft_file(store, did: str, fid: str) -> Dict[str, Any]:
    with _DRAFT_LOCK:
        d = get_draft(store, did)
        for f in d["files"]:
            if f["fid"] == fid:
                store.delete(f["path"])
        d["files"] = [f for f in d["files"] if f["fid"] != fid]
        put_json(store, f"drafts/{did}/draft.json", d)
        return d


def _entries(store, files: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    out = []
    for f in files:
        data = store.get(f["path"])
        if data is None:
            raise ValueError(f"stored file missing: {f['name']}")
        roles = {s["sheet"]: s["role"] for s in f["sheets"]}
        for s in read_workbook(data, f["name"]):
            role = roles.get(s["sheet"], "ignore")
            if role == "ignore":
                continue
            out.append({"file": f["name"], "sheet": s["sheet"], "role": role, "df": s["df"],
                        "header_row": s["header_row"], "n_formulas": s["n_formulas"], "bytes": data,
                        "col_offset": s.get("col_offset", 0), "grid": s.get("grid"),
                        "col_idx": s.get("col_idx"), "formula_cells": s.get("formula_cells")})
    return out


# ===================================================================== systems
INDEX = "index.json"


def _read_index(store) -> Dict[str, Any]:
    idx = get_json(store, INDEX)
    if idx is None:  # first run, or index lost: rebuild from the system folders
        idx = {"systems": {}}
        for p in store.list("systems/"):
            if p.endswith("/system.json"):
                sid = p.split("/")[-2]
                try:
                    idx["systems"][sid] = system_card(get_system(store, sid))
                except Exception:  # noqa: BLE001
                    continue
        put_json(store, INDEX, idx)
    return idx


def list_systems(store) -> List[Dict[str, Any]]:
    res = list(_read_index(store)["systems"].values())
    res.sort(key=lambda s: s.get("updated") or s.get("created") or "", reverse=True)
    return res


def system_card(s: Dict[str, Any]) -> Dict[str, Any]:
    sheets = s.get("sheets", [])
    n_formula = sum(1 for d in sheets for c in d["columns"] if c["kind"] == "formula")
    n_empty = sum(1 for d in sheets for c in d["columns"] if c["kind"] == "empty")
    sp = s.get("filters", {}).get("split", {})
    return {"id": s["id"], "name": s["name"], "description": s.get("description", ""),
            "created": s.get("created"), "updated": s.get("updated"),
            "n_sheets": len(sheets), "n_formulas": n_formula, "n_review": n_empty,
            "n_tables": len(s.get("tables", {})), "split": bool(sp.get("enabled")),
            "columns": sum(len(d["columns"]) for d in sheets)}


def _tslug(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name)[:80] or "table"


def get_system(store, sid: str) -> Dict[str, Any]:
    """Assemble a system from its files: system.json + rules.json + tables/<name>.json."""
    s = get_json(store, f"systems/{sid}/system.json")
    if s is None:
        raise KeyError("system not found")
    if "sheets" not in s:
        s["sheets"] = (get_json(store, f"systems/{sid}/template.json") or {}).get("sheets", [])
    if "tables" not in s:  # (older systems kept everything inline - handled by the `in` checks)
        s["tables"] = {}
        for name, fn in (s.get("table_files") or {}).items():
            t = get_json(store, f"systems/{sid}/tables/{fn}.json")
            if t is not None:
                s["tables"][name] = t
    return s


def save_system(store, s: Dict[str, Any]) -> None:
    """Readable layout in the folder:
        index.json                       one card per system (fast home page)
        systems/<id>/system.json         name, input/output layout, filters, toggles, training info
        systems/<id>/template.json       every output sheet: its columns, formulas, parameter cells
        systems/<id>/tables/<name>.json  one file per mapping table (rewritten only when changed)
    """
    import hashlib
    import json as _json
    s["updated"] = _now()
    sid = s["id"]
    base = f"systems/{sid}"
    tables = s.get("tables", {})
    files = {n: _tslug(n) for n in tables}
    hashes = dict(s.get("table_hashes") or {})
    for n, spec in tables.items():
        h = hashlib.md5(_json.dumps(spec, sort_keys=True, default=str).encode("utf-8")).hexdigest()
        if hashes.get(n) != h or store.get(f"{base}/tables/{files[n]}.json") is None:
            put_json(store, f"{base}/tables/{files[n]}.json", spec, indent=1)
            hashes[n] = h
    for old in set(hashes) - set(tables):
        store.delete(f"{base}/tables/{_tslug(old)}.json")
        hashes.pop(old, None)
    lean = {k: v for k, v in s.items() if k not in ("sheets", "tables")}
    lean["table_files"], lean["table_hashes"] = files, hashes
    put_json(store, f"{base}/template.json", {"sheets": s.get("sheets", [])}, indent=1)
    put_json(store, f"{base}/system.json", lean, indent=1)
    idx = _read_index(store)
    idx["systems"][sid] = system_card(s)
    put_json(store, INDEX, idx)


def delete_system(store, sid: str) -> None:
    delete_prefix(store, f"systems/{sid}/")
    delete_prefix(store, f"runs/{sid}/")
    idx = _read_index(store)
    idx["systems"].pop(sid, None)
    put_json(store, INDEX, idx)


def create_system(store, did: str, options: Dict[str, Any]) -> Dict[str, Any]:
    d = get_draft(store, did)
    system = engine.build_system(_entries(store, d["files"]), options)
    sid = system["id"]
    files = []                       # keep the example files with the system so it can be re-learned later
    for f in d["files"]:
        newp = f"systems/{sid}/training/{f['fid']}_{_slug(f['name'])}"
        store.put(newp, store.get(f["path"]))
        files.append({**f, "path": newp})
    put_json(store, f"systems/{sid}/training.json", {"files": files})
    delete_prefix(store, f"drafts/{did}/")
    save_system(store, system)
    return get_system(store, sid)


def relearn(store, sid: str, options: Dict[str, Any]) -> Dict[str, Any]:
    s = get_system(store, sid)
    tr = get_json(store, f"systems/{sid}/training.json")
    opts = {"name": s["name"], "description": s.get("description", ""), "id": sid}
    opts.update(options or {})
    new = engine.build_system(_entries(store, tr["files"]), opts)
    new["created"] = s.get("created", new["created"])
    save_system(store, new)
    return get_system(store, sid)


def update_system(store, sid: str, patch: Dict[str, Any]) -> Dict[str, Any]:
    s = get_system(store, sid)
    for k in ("name", "description"):
        if k in patch:
            s[k] = str(patch[k])
    if "filters" in patch:
        s["filters"] = engine.clean_filters(patch["filters"])
    if "toggles" in patch:
        s["toggles"] = patch["toggles"]
    if "max_rows_per_sheet" in patch:                    # output limit: a sheet continues on a new tab after this many rows
        v = patch["max_rows_per_sheet"]
        if v in (None, "", 0):
            s["output"].pop("max_rows_per_sheet", None)
        else:
            s["output"]["max_rows_per_sheet"] = clean_run_options({"max_rows_per_sheet": v})["max_rows_per_sheet"]
    if patch.get("formulas"):
        engine.apply_formula_edits(s, patch["formulas"])
    if patch.get("dtypes"):
        engine.apply_dtype_edits(s, patch["dtypes"])
    if "tables" in patch:  # replace mapping tables (e.g. refreshed rates)
        for n, spec in patch["tables"].items():
            s["tables"][n] = spec
    engine.finalize_system(s)
    save_system(store, s)
    return s


def _read_inputs(system: Dict[str, Any], files: List[Tuple[str, bytes]], workers: Optional[int] = None
                 ) -> Tuple[pd.DataFrame, List[str], List[Dict[str, Any]]]:
    cols = system["input"]["columns"]
    want = {nk(c) for c in cols if c != engine.SRC}
    req = {nk(c) for c in system["input"].get("required", []) if c != engine.SRC}
    frames, used = [], []
    for (name, data), sheets in zip(files, read_many(files, workers, values_only=True)):      # big files are read at the same time
        scored = []
        for s in sheets:
            have = {nk(c) for c in s["df"].columns}
            scored.append((len(want & have) / max(1, len(want)), req <= have, s))
        good = [s for sc, ok, s in scored if sc >= 0.6 and ok and len(s["df"])]
        if not good:
            best = max(scored, key=lambda t: t[0]) if scored else None
            miss = sorted((want | req) - {nk(c) for c in best[2]["df"].columns}) if best else []
            raise ValueError(f"{name}: no sheet looks like the input this system was built from "
                             f"(missing e.g. {', '.join(miss[:6])}).")
        multi = engine.SRC in cols
        for s in (good if multi else sorted(good, key=lambda x: -len(x["df"]))[:1]):
            f = canonical_columns(s["df"], cols)
            if multi:   # the input's own sheets are the split; keep each row's sheet name
                f = f.copy()
                f[engine.SRC] = s["sheet"]
            frames.append(f)
            used.append({"file": name, "sheet": s["sheet"], "rows": int(len(s["df"]))})
    return pd.concat(frames, ignore_index=True), [], used


def clean_run_options(options: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """Per-run settings for big data: max_rows_per_sheet, workers, fast_writer (True / False / None = by size)."""
    o = options or {}
    out: Dict[str, Any] = {}

    def num(key: str, lo: int, hi: int) -> Optional[int]:
        v = o.get(key)
        if v in (None, ""):
            return None
        try:
            n = int(str(v).replace(",", "").replace("_", ""))
        except ValueError:
            raise ValueError("Option %s must be a whole number" % key)
        if not lo <= n <= hi:
            raise ValueError("Option %s must be between %s and %s" % (key, format(lo, ","), format(hi, ",")))
        return n
    for key, lo, hi in (("max_rows_per_sheet", 1, perf.EXCEL_MAX_ROWS - 1), ("workers", 1, 64)):
        v = num(key, lo, hi)
        if v is not None:
            out[key] = v
    fw = o.get("fast_writer")
    out["fast_writer"] = None if fw in (None, "", "auto") else (fw is True or str(fw).lower() in ("1", "true", "yes", "on"))
    return out


def run_system(store, sid: str, input_files: List[Tuple[str, bytes]],
               mapping_files: Optional[List[Tuple[str, bytes]]] = None,
               toggle_values: Optional[Dict[str, str]] = None,
               run_filter: Optional[Dict[str, Any]] = None,
               options: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    timer = perf.Timer()
    s = get_system(store, sid)
    opts = clean_run_options(options)
    nw = opts.get("workers")
    with timer.stage("read"):
        df, _, used = _read_inputs(s, input_files, nw)
        override = {}
        for name, data in mapping_files or []:
            for sh in read_workbook(data, name):
                for tname in s["tables"]:
                    if nk(tname) == nk(sh["sheet"]) or nk(tname) == nk(re.sub(r"\.[^.]+$", "", name)):
                        override[tname] = sh["df"]
    res = engine.apply_system(s, df, override, toggle_values, run_filter, timer=timer)
    notes = list(res["notes"])
    if override:
        notes.append("Mapping tables replaced for this run: " + ", ".join(override))
    empty = [c["name"] for d in s["sheets"] for c in d["columns"] if c["kind"] == "empty"]
    if empty:
        notes.append("No formula was found in your sample for: " + ", ".join(dict.fromkeys(empty)) + " - left empty.")
    parts, pnotes = engine.plan_parts(s, res["sheets"], opts.get("max_rows_per_sheet"))
    notes += pnotes
    tg = {t["name"]: t.get("value") for t in res["toggles"]}
    rf = engine.clean_run_filter(run_filter)
    summary = [("Run filter", engine.describe_conditions(rf["conditions"], rf["match"]) if rf else "none"),
               ("Toggles used", ", ".join("%s = %s" % kv for kv in tg.items()) or "none"),
               ("Rows", "%s read, %s written" % (format(res["input_rows"], ","), format(res["rows"], ","))),
               ("Input files", [f"{u['file']} / {u['sheet']} ({u['rows']} rows)" for u in used])]
    if notes:
        summary.append(("Notes", notes))
    tables = engine.tables_to_frames(s, override)
    winfo: Dict[str, Any] = {}
    with timer.stage("write"):
        data = engine.write_workbook(s, res["sheets"], summary, tables, res["toggles"], max_rows=opts.get("max_rows_per_sheet"),
                                     fast=opts.get("fast_writer"), workers=nw, info=winfo)
    rid = dt.datetime.utcnow().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:4]
    with timer.stage("original format"):                 # the input came as csv / xlsm / ...: the output in that format too
        extras, fnotes = formats.original_outputs(res["sheets"], parts, data, used, f"{_slug(s['name'])}_{rid}")
    notes += fnotes
    with timer.stage("save"):
        store.put(f"runs/{sid}/{rid}/output.xlsx", data)
        for e in extras:
            store.put(f"runs/{sid}/{rid}/original/{e['filename']}", e["data"])
    preview = {}
    for sn, item in res["sheets"].items():
        f = engine.full_frame(item, 0, 25)
        preview[sn] = {"rows": int(item["n"]), "headers": list(map(str, f.columns)),
                       "kinds": [c["kind"] for c in item["def"]["columns"]],
                       "data": [[jsonable(v) for v in r] for r in f.itertuples(index=False, name=None)]}
    split: Dict[str, List[str]] = {}
    for pn, (sn, _, _) in parts.items():
        split.setdefault(sn, []).append(pn)
    timing = timer.report(res["input_rows"], res["rows"], writer=winfo.get("writer"), workers=winfo.get("workers", 1),
                          tabs=len(parts), output_mb=round(len(data) / 1048576.0, 1))
    if winfo.get("pool_error"):
        notes.append("Worker processes were not available (%s); the file was written in one process." % winfo["pool_error"])
    meta = {"run_id": rid, "system": sid, "created": _now(), "rows": int(res["rows"]),
            "input_rows": int(res["input_rows"]), "sheets": {k: v["rows"] for k, v in preview.items()},
            "split": {k: v for k, v in split.items() if len(v) > 1},
            "extra_outputs": [{"format": e["format"], "filename": e["filename"], "note": e["note"], "mb": round(len(e["data"]) / 1048576.0, 2)}
                              for e in extras],
            "notes": notes, "inputs": used, "toggles": res["toggles"], "run_filter": rf, "timing": timing,
            "filename": f"{_slug(s['name'])}_{rid}.xlsx"}
    put_json(store, f"runs/{sid}/{rid}/meta.json", meta)
    return {**meta, "preview": preview}


_FRAMES: "OrderedDict[Tuple[str, str], pd.DataFrame]" = OrderedDict()      # the last loaded inputs (at most 2)


def _remember(sid: str, token: str, df: pd.DataFrame) -> None:
    _FRAMES[(sid, token)] = df
    while len(_FRAMES) > 2:
        _FRAMES.popitem(last=False)


def _stash(store, files: List[Tuple[str, bytes]]) -> str:
    """Keep a loaded input briefly (tmp/<token>/) so filter previews do not re-upload it every keystroke."""
    token = dt.datetime.utcnow().strftime("%Y%m%d%H%M%S") + "-" + uuid.uuid4().hex[:6]
    for i, (name, data) in enumerate(files):
        store.put(f"tmp/{token}/{i}__{_slug(name)}", data)
    cutoff = (dt.datetime.utcnow() - dt.timedelta(days=1)).strftime("%Y%m%d%H%M%S")
    for p in store.list("tmp/"):                                  # drop stashes older than a day
        t = p.strip("/").split("/")[1] if p.count("/") >= 2 else ""
        if t and t[:14] < cutoff:
            store.delete(p)
    return token


def _unstash(store, token: str) -> List[Tuple[str, bytes]]:
    if not re.fullmatch(r"[0-9]{14}-[0-9a-f]{6}", token or ""):
        raise ValueError("Drop the new input file again (its temporary copy is gone).")
    paths = sorted(store.list(f"tmp/{token}/"), key=lambda p: int(p.rsplit("/", 1)[1].split("__", 1)[0]))
    if not paths:
        raise ValueError("Drop the new input file again (its temporary copy is gone).")
    return [(p.rsplit("/", 1)[1].split("__", 1)[1], store.get(p)) for p in paths]


def preview_filter(store, sid: str, token: str, run_filter: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """Row counts before/after the filters for an already loaded input (see inspect_input)."""
    s = get_system(store, sid)
    df = _FRAMES.get((sid, token))                       # the parsed file stays in memory: a big file is not re-read per keystroke
    if df is None:
        df, _, _ = _read_inputs(s, _unstash(store, token))
        _remember(sid, token, df)
    return engine.preview_counts(s, df, run_filter)


def inspect_input(store, sid: str, files: List[Tuple[str, bytes]]) -> Dict[str, Any]:
    """What the new input looks like, so a run filter can offer its real columns and values."""
    s = get_system(store, sid)
    df, _, used = _read_inputs(s, files)
    token = _stash(store, files)
    _remember(sid, token, df)
    values: Dict[str, List[str]] = {}
    for c in df.columns:
        u = df[c].dropna().unique()                      # distinct values first (fast), only then the readable form
        if 0 < len(u) <= 300:
            vals = sorted({disp(v) for v in u if not is_blank(v)})
            if 0 < len(vals) <= 30:
                values[str(c)] = [v[:40] for v in vals]
    return {"columns": [str(c) for c in df.columns], "rows": int(len(df)), "sheets": used, "values": values, "token": token,
            "preview": engine.preview_counts(s, df, None)}


def list_runs(store, sid: str) -> List[Dict[str, Any]]:
    out = []
    for p in store.list(f"runs/{sid}/"):
        if p.endswith("/meta.json"):
            m = get_json(store, p.lstrip("/"))
            if m:
                m.pop("compare", None)
                out.append(m)
    out.sort(key=lambda m: m["created"], reverse=True)
    return out[:30]


def get_run_file(store, sid: str, rid: str, extra: Optional[str] = None) -> Tuple[bytes, str]:
    """The run's .xlsx, or - with ``extra`` = a file name from meta['extra_outputs'] - its output in the input's own format."""
    meta = get_json(store, f"runs/{sid}/{rid}/meta.json") or {}
    if extra:
        if extra not in [e["filename"] for e in meta.get("extra_outputs", [])]:
            raise KeyError("file not found")
        data = store.get(f"runs/{sid}/{rid}/original/{extra}")
        if data is None:
            raise KeyError("file not found")
        return data, extra
    data = store.get(f"runs/{sid}/{rid}/output.xlsx")
    if data is None:
        raise KeyError("run not found")
    return data, meta.get("filename", f"{rid}.xlsx")


def input_template(system: Dict[str, Any]) -> bytes:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Input"
    ws.append([c for c in system["input"]["columns"] if c != engine.SRC])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
