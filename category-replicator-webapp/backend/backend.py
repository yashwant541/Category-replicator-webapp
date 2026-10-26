# -*- coding: utf-8 -*-
"""**Category Replicator** - Dataiku Standard webapp, Python pane.

Teach a *category* from examples (input file(s) + mapping file(s) + the expected output) and the
app learns a rule for every output column - copy, lookup into a mapping, arithmetic, concat,
IF/bands, date parts, or the live Excel formula itself.  Run the category on new data to get the
same output back, optionally split into several sheets by a filter.

Needs ``replicator_formulas`` and ``replicator`` on the project's Python libraries (both in ``python-lib/``).
Storage: one managed folder whose id is the project variable ``REPLICATOR_FOLDER`` (see store.py).
"""

from __future__ import annotations

import io
import traceback
from functools import wraps

from flask import Response, jsonify, request

from replicator import service
from replicator.engine import ENGINE_BUILD
from replicator.store import make_store

try:  # pragma: no cover - Dataiku provides `app`
    app  # type: ignore  # noqa: F821
except NameError:  # pragma: no cover
    from flask import Flask
    app = Flask(__name__)

# ---- GLOBAL SETTING: the managed folder where systems, mappings, rules and runs are saved.
# Put the folder's name or id here. If left empty, the project variable REPLICATOR_FOLDER is used.
REPLICATOR_FOLDER = ""

_STORE = None
_MAX_UPLOAD = 200 * 1024 * 1024


def store():
    global _STORE
    if _STORE is None:
        try:
            _STORE = make_store(REPLICATOR_FOLDER)
        except TypeError:  # stale replicator/store.py in the project libraries: use the env fallback
            import os
            if REPLICATOR_FOLDER:
                os.environ["REPLICATOR_FOLDER"] = REPLICATOR_FOLDER
            _STORE = make_store()
    return _STORE


def ok(**kw):
    kw["success"] = True
    return jsonify(kw)


def fail(msg, code=400, details=None):
    return jsonify({"success": False, "error": str(msg), "details": details}), code


def guarded(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        try:
            return fn(*a, **kw)
        except KeyError as exc:
            return fail(exc.args[0] if exc.args else "not found", 404)
        except ValueError as exc:
            return fail(exc)
        except Exception as exc:  # noqa: BLE001
            return fail(f"{type(exc).__name__}: {exc}", 500, traceback.format_exc()[-1500:])
    return wrapper


def _files(field="files"):
    out = []
    for f in request.files.getlist(field):
        if f and f.filename:
            data = f.read()
            if len(data) > _MAX_UPLOAD:
                raise ValueError(f"{f.filename} is too large")
            out.append((f.filename, data))
    return out


def _download(data, name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
    return Response(data, mimetype=mime, headers={
        "Content-Disposition": 'attachment; filename="%s"' % name, "Content-Length": str(len(data))})


# ---------------------------------------------------------------------------- routes
@app.route("/api/health", methods=["GET"])
@guarded
def api_health():
    return ok(engine_build=ENGINE_BUILD, storage=type(store()).__name__)


@app.route("/api/systems", methods=["GET"])
@guarded
def api_systems():
    return ok(systems=service.list_systems(store()))


@app.route("/api/drafts", methods=["POST"])
@guarded
def api_new_draft():
    return ok(draft_id=service.new_draft(store()))


@app.route("/api/drafts/<did>", methods=["GET"])
@guarded
def api_get_draft(did):
    return ok(draft=service.get_draft(store(), did))


@app.route("/api/drafts/<did>/files", methods=["POST"])
@guarded
def api_draft_files(did):
    d = service.add_draft_files(store(), did, _files(), request.form.get("hint", ""))
    return ok(draft=d)


@app.route("/api/drafts/<did>/files/<fid>", methods=["DELETE"])
@guarded
def api_draft_file_delete(did, fid):
    return ok(draft=service.remove_draft_file(store(), did, fid))


@app.route("/api/drafts/<did>/roles", methods=["PUT"])
@guarded
def api_draft_roles(did):
    return ok(draft=service.set_draft_roles(store(), did, (request.get_json(silent=True) or {}).get("roles", {})))


@app.route("/api/drafts/<did>/learn", methods=["POST"])
@guarded
def api_learn(did):
    body = request.get_json(silent=True) or {}
    s = service.create_system(store(), did, {
        "name": (body.get("name") or "").strip() or "Untitled system",
        "description": body.get("description", ""), "split": body.get("split") or {"mode": "auto"}})
    return ok(system=s)


@app.route("/api/systems/<sid>", methods=["GET"])
@guarded
def api_system(sid):
    return ok(system=service.get_system(store(), sid))


@app.route("/api/systems/<sid>", methods=["PUT"])
@guarded
def api_system_update(sid):
    return ok(system=service.update_system(store(), sid, request.get_json(silent=True) or {}))


@app.route("/api/systems/<sid>", methods=["DELETE"])
@guarded
def api_system_delete(sid):
    service.delete_system(store(), sid)
    return ok()


@app.route("/api/systems/<sid>/relearn", methods=["POST"])
@guarded
def api_relearn(sid):
    body = request.get_json(silent=True) or {}
    return ok(system=service.relearn(store(), sid, {"split": body.get("split") or {"mode": "auto"}}))


@app.route("/api/systems/<sid>/replay", methods=["POST"])
@guarded
def api_replay(sid):
    return ok(report=service.replay(store(), sid), system=service.get_system(store(), sid))


@app.route("/api/systems/<sid>/run", methods=["POST"])
@guarded
def api_run(sid):
    inputs = _files("files")
    if not inputs:
        raise ValueError("Upload at least one input file.")
    exp = _files("expected")
    import json as _json
    try:
        tv = _json.loads(request.form.get("toggles") or "{}")
    except ValueError:
        tv = {}
    res = service.run_system(store(), sid, inputs, _files("mappings"), exp[0] if exp else None,
                             request.form.get("live_formulas") in ("1", "true", "on"), tv or None)
    return ok(run=res)


@app.route("/api/systems/<sid>/runs", methods=["GET"])
@guarded
def api_runs(sid):
    return ok(runs=service.list_runs(store(), sid))


@app.route("/api/systems/<sid>/runs/<rid>/download", methods=["GET"])
@guarded
def api_run_download(sid, rid):
    data, name = service.get_run_file(store(), sid, rid)
    return _download(data, name)


@app.route("/api/systems/<sid>/template", methods=["GET"])
@guarded
def api_template(sid):
    s = service.get_system(store(), sid)
    return _download(service.input_template(s), "%s_input_template.xlsx" % s["name"])


@app.route("/api/systems/<sid>/tables/<tname>/download", methods=["GET"])
@guarded
def api_table_download(sid, tname):
    import openpyxl
    s = service.get_system(store(), sid)
    spec = s["tables"][tname]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = tname[:31]
    ws.append(spec["headers"])
    for r in spec["records"]:
        ws.append([r.get(h) for h in spec["headers"]])
    buf = io.BytesIO()
    wb.save(buf)
    return _download(buf.getvalue(), "%s.xlsx" % tname)
