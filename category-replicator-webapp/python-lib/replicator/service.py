"""Application services on top of a store: drafts, systems, runs.  No Flask, no Dataiku."""

from __future__ import annotations

import datetime as dt
import io
import threading
import re
import uuid
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from . import engine
from .rules import nk
from .store import delete_prefix, get_json, put_json
from .tables import formula_sheet_refs, guess_role, jsonable, read_workbook, summarize, canonical_columns

_DRAFT_LOCK = threading.RLock()   # serialises the draft record update (uploads can overlap)
ALLOWED = (".xlsx", ".xlsm", ".xls", ".csv", ".txt")
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
        if not name.lower().endswith(ALLOWED):
            raise ValueError(f"{name}: unsupported file type (use .xlsx, .xls or .csv)")
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
    for k in ("filters", "toggles"):
        if k in patch:
            s[k] = patch[k]
    if "tables" in patch:  # replace mapping tables (e.g. refreshed rates)
        for n, spec in patch["tables"].items():
            s["tables"][n] = spec
    engine.finalize_system(s)
    save_system(store, s)
    return s


def _read_inputs(system: Dict[str, Any], files: List[Tuple[str, bytes]]
                 ) -> Tuple[pd.DataFrame, List[str], List[Dict[str, Any]]]:
    cols = system["input"]["columns"]
    want = {nk(c) for c in cols if c != engine.SRC}
    req = {nk(c) for c in system["input"].get("required", []) if c != engine.SRC}
    frames, used = [], []
    for name, data in files:
        scored = []
        for s in read_workbook(data, name):
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


def run_system(store, sid: str, input_files: List[Tuple[str, bytes]],
               mapping_files: Optional[List[Tuple[str, bytes]]] = None,
               toggle_values: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
    s = get_system(store, sid)
    df, _, used = _read_inputs(s, input_files)
    override = {}
    for name, data in mapping_files or []:
        for sh in read_workbook(data, name):
            for tname in s["tables"]:
                if nk(tname) == nk(sh["sheet"]) or nk(tname) == nk(re.sub(r"\.[^.]+$", "", name)):
                    override[tname] = sh["df"]
    res = engine.apply_system(s, df, override, toggle_values)
    notes = list(res["notes"])
    if override:
        notes.append("Mapping tables replaced for this run: " + ", ".join(override))
    empty = [c["name"] for d in s["sheets"] for c in d["columns"] if c["kind"] == "empty"]
    if empty:
        notes.append("No formula was found in your sample for: " + ", ".join(dict.fromkeys(empty)) + " - left empty.")
    tg = {t["name"]: t.get("value") for t in res["toggles"]}
    summary = [("Toggles used", ", ".join("%s = %s" % kv for kv in tg.items()) or "none"),
               ("Input files", [f"{u['file']} / {u['sheet']} ({u['rows']} rows)" for u in used])]
    if notes:
        summary.append(("Notes", notes))
    tables = engine.tables_to_frames(s, override)
    data = engine.write_workbook(s, res["sheets"], summary, tables, res["toggles"])
    rid = dt.datetime.utcnow().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:4]
    store.put(f"runs/{sid}/{rid}/output.xlsx", data)
    preview = {}
    for sn, item in res["sheets"].items():
        f = item["df"]
        preview[sn] = {"rows": int(len(f)), "headers": list(map(str, f.columns)),
                       "kinds": [c["kind"] for c in item["def"]["columns"]],
                       "data": [[jsonable(v) for v in r] for r in f.head(25).itertuples(index=False, name=None)]}
    meta = {"run_id": rid, "system": sid, "created": _now(), "rows": int(res["rows"]),
            "input_rows": int(res["input_rows"]), "sheets": {k: v["rows"] for k, v in preview.items()},
            "notes": notes, "inputs": used, "toggles": res["toggles"],
            "filename": f"{_slug(s['name'])}_{rid}.xlsx"}
    put_json(store, f"runs/{sid}/{rid}/meta.json", meta)
    return {**meta, "preview": preview}


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


def get_run_file(store, sid: str, rid: str) -> Tuple[bytes, str]:
    data = store.get(f"runs/{sid}/{rid}/output.xlsx")
    meta = get_json(store, f"runs/{sid}/{rid}/meta.json") or {}
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
