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
import os
import shutil
import tempfile
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

# ---- HEAVY FILE MODE: two more managed folders (name or id). Leave empty to use sub-folders of the folder above.
UPLOAD_FOLDER = ""      # where big input files are uploaded to (and run from)
DOWNLOAD_FOLDER = ""    # where the outputs of heavy runs are written for download

_STORE = None
_MAX_UPLOAD = int(os.environ.get("REPLICATOR_MAX_UPLOAD_MB", "2000")) * 1024 * 1024     # per file; big inputs are expected


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
        _setup_heavy(_STORE)
    return _STORE


def _setup_heavy(main):
    """Point the service at the upload / download folders (separate managed folders, or sub-folders of the main one)."""
    def pick(folder, sub):
        if folder:
            st = make_store(folder)
            if type(st).__name__ == "DataikuStore":      # its own managed folder: files sit at its root
                return st, ""
            return st, sub                                  # local development: a sub-folder of the local store
        return main, sub
    up, upp = pick(UPLOAD_FOLDER, "heavy_uploads")
    down, downp = pick(DOWNLOAD_FOLDER, "heavy_downloads")
    service.configure_heavy(up, upp, down, downp)


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


@app.route("/api/systems/<sid>/run", methods=["POST"])
@guarded
def api_run(sid):
    inputs = _files("files")
    inbox, token = _inbox_and_token()
    if not inputs and not inbox and not token:
        raise ValueError("Upload at least one input file.")
    import json as _json
    try:
        tv = _json.loads(request.form.get("toggles") or "{}")
    except ValueError:
        tv = {}
    try:
        rf = _json.loads(request.form.get("filters") or "null")
    except ValueError:
        raise ValueError("The run filter is not valid JSON")
    try:
        opts = _json.loads(request.form.get("options") or "null")
    except ValueError:
        raise ValueError("The run options are not valid JSON")
    maps = _files("mappings")
    st = store()
    # a big run is a background job (the request would be cut by the proxy long before it ends); poll /api/jobs/<id>
    heavy = request.form.get("heavy") == "1"
    jid = service.start_job("run", lambda progress: service.run_system(
        st, sid, service.resolve_inputs(st, inputs, inbox, token), maps, tv or None, rf, opts, progress=progress, heavy=heavy))
    return ok(job_id=jid)


def _inbox_and_token():
    import json as _json
    try:
        inbox = [str(x) for x in (_json.loads(request.form.get("inbox") or "[]") or [])]
    except ValueError:
        raise ValueError("The inbox list is not valid JSON")
    return inbox, (request.form.get("token") or "").strip() or None


# ---- heavy-file mode: upload (in chunks), list, delete, clear; downloads likewise
_CHUNK_ROOT = os.path.join(tempfile.gettempdir(), "replicator_chunks")


@app.route("/api/heavy/<area>", methods=["GET"])
@guarded
def api_heavy_list(area):
    return ok(files=service.list_area(store(), area))


@app.route("/api/heavy/upload", methods=["POST"])
@guarded
def api_heavy_upload():
    """One chunk of a big file (raw request body). Chunks are put together on the server and stored in the upload folder;
    a browser never has to send 500 MB in a single request."""
    import re
    name = service._safe_name(request.args.get("name", ""))
    uid = request.args.get("id", "")
    if not re.fullmatch(r"[A-Za-z0-9_-]{6,64}", uid):
        raise ValueError("Bad upload id.")
    index, total = int(request.args.get("index", "0")), int(request.args.get("total", "1"))
    if not 0 <= index < total <= 20000:
        raise ValueError("Bad chunk number.")
    store()                                                   # make sure the folders are configured
    d = os.path.join(_CHUNK_ROOT, uid)
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, "%06d" % index), "wb") as fh:
        fh.write(request.get_data())
    if index < total - 1:
        return ok(done=False, index=index)
    parts = sorted(os.listdir(d))
    if len(parts) != total:
        raise ValueError("A chunk is missing (%d of %d arrived). Upload the file again." % (len(parts), total))
    whole = os.path.join(d, "whole")
    with open(whole, "wb") as out:
        for part in parts:
            with open(os.path.join(d, part), "rb") as fh:
                shutil.copyfileobj(fh, out, 8 * 1024 * 1024)
    try:
        service.area_put_file(store(), "uploads", name, whole)
    finally:
        shutil.rmtree(d, ignore_errors=True)
    return ok(done=True, name=name)


@app.route("/api/heavy/<area>/delete", methods=["POST"])
@guarded
def api_heavy_delete(area):
    names = (request.get_json(silent=True) or {}).get("names") or []
    return ok(deleted=service.area_delete(store(), area, [str(n) for n in names]))


@app.route("/api/heavy/<area>/clear", methods=["POST"])
@guarded
def api_heavy_clear(area):
    return ok(deleted=service.area_clear(store(), area))


@app.route("/api/heavy/downloads/file", methods=["GET"])
@guarded
def api_heavy_download():
    name = request.args.get("name", "")
    return _download(service.area_get(store(), "downloads", name), name, "application/octet-stream")


@app.route("/api/jobs/<jid>", methods=["GET"])
@guarded
def api_job(jid):
    return ok(job=service.get_job(jid))


@app.route("/api/systems/<sid>/inspect-input", methods=["POST"])
@guarded
def api_inspect_input(sid):
    files = _files("files")
    inbox, _ = _inbox_and_token()
    if not files and not inbox:
        raise ValueError("Upload at least one input file.")
    st = store()
    jid = service.start_job("inspect", lambda progress: service.inspect_input(
        st, sid, service.resolve_inputs(st, files, inbox), inbox))
    return ok(job_id=jid)


@app.route("/api/systems/<sid>/preview-filter", methods=["POST"])
@guarded
def api_preview_filter(sid):
    body = request.get_json(silent=True) or {}
    return ok(**service.preview_filter(store(), sid, body.get("token", ""), body.get("filters")))


@app.route("/api/systems/<sid>/runs", methods=["GET"])
@guarded
def api_runs(sid):
    return ok(runs=service.list_runs(store(), sid))


@app.route("/api/systems/<sid>/runs/<rid>/download", methods=["GET"])
@guarded
def api_run_download(sid, rid):
    data, name = service.get_run_file(store(), sid, rid, request.args.get("file"))
    return _download(data, name, "application/octet-stream" if request.args.get("file") else
                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


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
