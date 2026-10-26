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
                        "col_offset": s.get("col_offset", 0), "grid": s.get("grid")})
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
    rules = s.get("rules", [])
    bad = [r for r in rules if r.get("status") in ("unresolved", "approx")]
    sp = s.get("filters", {}).get("split", {})
    return {"id": s["id"], "name": s["name"], "description": s.get("description", ""),
            "created": s.get("created"), "updated": s.get("updated"),
            "n_rules": len(rules), "n_review": len(bad), "n_tables": len(s.get("tables", {})),
            "split": bool(sp.get("enabled")), "match": (s.get("training") or {}).get("replay_match"),
            "columns": len(s.get("output", {}).get("columns", []))}


def _tslug(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name)[:80] or "table"


def get_system(store, sid: str) -> Dict[str, Any]:
    """Assemble a system from its files: system.json + rules.json + tables/<name>.json."""
    s = get_json(store, f"systems/{sid}/system.json")
    if s is None:
        raise KeyError("system not found")
    if "rules" not in s:
        s["rules"] = (get_json(store, f"systems/{sid}/rules.json") or {}).get("rules", [])
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
        systems/<id>/rules.json          every column rule (formulas, lookups, ...)
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
    lean = {k: v for k, v in s.items() if k not in ("rules", "tables")}
    lean["table_files"], lean["table_hashes"] = files, hashes
    put_json(store, f"{base}/rules.json", {"rules": s.get("rules", [])}, indent=1)
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


def _training_frames(store, s: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    tr = get_json(store, f"systems/{s['id']}/training.json")
    if not tr:
        raise ValueError("No training files stored for this system.")
    es = _entries(store, tr["files"])
    ins = [e for e in es if e["role"] == "input"]
    outs = [e for e in es if e["role"] == "output"]
    cols = s["input"]["columns"]
    same = [e for e in ins if {nk(c) for c in e["df"].columns} == {nk(c) for c in cols if c != engine.SRC}]
    frames = []
    for e in same or [ins[0]]:
        f = canonical_columns(e["df"], cols).copy()
        if engine.SRC in cols:
            f[engine.SRC] = e["sheet"]
        frames.append(f)
    X = pd.concat(frames, ignore_index=True)
    ocols = s["output"]["columns"]
    Y_sheets = {e["sheet"]: canonical_columns(e["df"], ocols) for e in outs}
    Y = pd.concat(list(Y_sheets.values()), ignore_index=True)
    return X, Y, Y_sheets


def replay(store, sid: str, persist: bool = True) -> Dict[str, Any]:
    s = get_system(store, sid)
    X, Y, Y_sheets = _training_frames(store, s)
    res = engine.replay_training(s, X, Y, Y_sheets)
    if persist:
        s.setdefault("training", {})["replay_match"] = res.get("match")
        s["training"]["replay_cols"] = {c["column"]: c["match"] for c in res.get("columns", [])}
        save_system(store, s)
    return res


def create_system(store, did: str, options: Dict[str, Any]) -> Dict[str, Any]:
    d = get_draft(store, did)
    es = _entries(store, d["files"])
    system = engine.build_system(es, options)
    sid = system["id"]
    # move the draft files under the system so it can be re-learned / replayed later
    files = []
    for f in d["files"]:
        data = store.get(f["path"])
        newp = f"systems/{sid}/training/{f['fid']}_{_slug(f['name'])}"
        store.put(newp, data)
        files.append({**f, "path": newp})
    put_json(store, f"systems/{sid}/training.json", {"files": files})
    delete_prefix(store, f"drafts/{did}/")
    save_system(store, system)
    try:
        replay(store, sid)
    except Exception as exc:  # noqa: BLE001
        system = get_system(store, sid)
        system["warnings"].append(f"Replay check failed: {exc}")
        save_system(store, system)
    return get_system(store, sid)


def relearn(store, sid: str, options: Dict[str, Any]) -> Dict[str, Any]:
    s = get_system(store, sid)
    tr = get_json(store, f"systems/{sid}/training.json")
    es = _entries(store, tr["files"])
    opts = {"name": s["name"], "description": s.get("description", ""), "id": sid}
    opts.update(options or {})
    new = engine.build_system(es, opts)
    new["created"] = s.get("created", new["created"])
    save_system(store, new)
    replay(store, sid)
    return get_system(store, sid)


def update_system(store, sid: str, patch: Dict[str, Any]) -> Dict[str, Any]:
    s = get_system(store, sid)
    for k in ("name", "description"):
        if k in patch:
            s[k] = str(patch[k])
    if "rules" in patch:
        old = {r["name"]: r for r in s["rules"]}
        order = []
        for r in patch["rules"]:
            if r["name"] in old:
                if r.get("edited"):
                    merged = {k: v for k, v in r.items() if k != "edited"}
                    merged.update(source="manual", status="manual", match=None)
                else:
                    merged = dict(old[r["name"]])
                order.append(merged)
        s["rules"] = order
        s["output"]["columns"] = [r["name"] for r in order]
    if "filters" in patch:
        s["filters"] = patch["filters"]
    if "toggles" in patch:
        s["toggles"] = patch["toggles"]
    if "tables" in patch:  # replace mapping tables (e.g. refreshed rates)
        for n, spec in patch["tables"].items():
            s["tables"][n] = spec
    engine.finalize_system(s)
    save_system(store, s)
    return s


# ========================================================================= runs
def _read_inputs(system: Dict[str, Any], files: List[Tuple[str, bytes]]
                 ) -> Tuple[pd.DataFrame, List[str], List[Dict[str, Any]]]:
    cols = system["input"]["columns"]
    want = {nk(c) for c in cols if c != engine.SRC}
    req = {nk(c) for c in system["input"].get("required", []) if c != engine.SRC} or want
    frames, used, notes = [], [], []
    for name, data in files:
        sheets = read_workbook(data, name)
        scored = []
        for s in sheets:
            have = {nk(c) for c in s["df"].columns}
            scored.append((len(req & have) / max(1, len(req)), s))
        scored.sort(key=lambda t: -t[0])
        good = [s for sc, s in scored if sc >= 0.999 and len(s["df"])]
        if not good:
            best = scored[0] if scored else None
            miss = sorted(req - {nk(c) for c in best[1]["df"].columns}) if best else []
            raise ValueError(f"{name}: no sheet has the expected input columns "
                             f"(missing e.g. {', '.join(miss[:6])}).")
        multi = engine.SRC in cols
        for s in (good if multi else good[:1]):
            f = canonical_columns(s["df"], cols)
            if multi:   # the input's own sheets are the split; keep each row's sheet name
                f = f.copy()
                f[engine.SRC] = s["sheet"]
            frames.append(f)
            used.append({"file": name, "sheet": s["sheet"], "rows": int(len(s["df"]))})
    return pd.concat(frames, ignore_index=True), notes, used


def run_system(store, sid: str, input_files: List[Tuple[str, bytes]],
               mapping_files: Optional[List[Tuple[str, bytes]]] = None,
               expected_file: Optional[Tuple[str, bytes]] = None,
               live_formulas: bool = False,
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
    tg = {t["name"]: t.get("value") for t in res["toggles"]}
    summary = [("Toggles used", ", ".join("%s = %s" % kv for kv in tg.items()) or "none"), ("Input files", [f"{u['file']} / {u['sheet']} ({u['rows']} rows)" for u in used])]
    if notes:
        summary.append(("Notes", notes))
    cmp_ = None
    if expected_file:
        ex = read_workbook(*(expected_file[1], expected_file[0]))
        ocols = s["output"]["columns"]
        ysh = {x["sheet"]: canonical_columns(x["df"], ocols) for x in ex
               if len({nk(c) for c in x["df"].columns} & {nk(c) for c in ocols}) >= max(1, len(ocols) // 2)}
        if ysh:
            if len(ysh) == 1 and len(res["sheets"]) == 1:
                cmp_ = engine.compare_frames(next(iter(res["sheets"].values())), next(iter(ysh.values())))
            else:
                cmp_ = engine.compare_sheets(res["sheets"], ysh)
            summary.append(("Match vs expected output", f"{cmp_.get('match', 0) * 100:.2f}%" if cmp_.get("aligned") else cmp_.get("note")))
    tables = engine.tables_to_frames(s, override)
    data = engine.write_workbook(s, res["sheets"], res["work"], summary, live_formulas, tables, res["toggles"])
    rid = dt.datetime.utcnow().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:4]
    store.put(f"runs/{sid}/{rid}/output.xlsx", data)
    preview = {}
    for sn, f in res["sheets"].items():
        preview[sn] = {"rows": int(len(f)), "headers": list(map(str, f.columns)),
                       "data": [[jsonable(v) for v in r] for r in f.head(25).itertuples(index=False, name=None)]}
    meta = {"run_id": rid, "system": sid, "created": _now(), "rows": int(res["rows"]),
            "input_rows": int(res["input_rows"]), "sheets": {k: v["rows"] for k, v in preview.items()},
            "notes": notes, "rule_errors": res["rule_errors"], "inputs": used,
            "live_formulas": bool(live_formulas), "compare": cmp_, "toggles": res["toggles"],
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
