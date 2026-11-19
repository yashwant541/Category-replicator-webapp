"""Big files: background jobs, running from the server's inbox/ folder, and re-using the already uploaded copy."""
import io
import json
import time

import pytest

from replicator import service
from replicator.store import LocalStore
from test_run_filter import HDR, ROWS, book, output_book  # noqa: F401  (output_book: shared sample builder)
from test_template import b
import openpyxl


def _system(c):
    did = c.post("/api/drafts").get_json()["draft_id"]
    out = openpyxl.Workbook()
    o = out.active
    o.title = "Out"
    o.append(list(HDR) + ["Fee"])
    for i, r in enumerate(ROWS, start=2):
        o.append(list(r) + [f"=D{i}*0.1"])
    for n, data, h in (("i.xlsx", book(), "input"), ("o.xlsx", b(out), "output")):
        c.post(f"/api/drafts/{did}/files", data={"files": (io.BytesIO(data), n), "hint": h}, content_type="multipart/form-data")
    return c.post(f"/api/drafts/{did}/learn", json={"name": "jobs"}).get_json()["system"]["id"]


def _wait(c, r, secs=60):
    jid = r.get_json()["job_id"]
    t = time.time()
    while time.time() - t < secs:
        j = c.get("/api/jobs/" + jid).get_json()["job"]
        if j["state"] != "running":
            return j
        time.sleep(0.05)
    raise AssertionError("slow job")


def raw_client():
    from test_api import client
    return client()


def test_run_returns_a_job_id_at_once_and_reports_stages(monkeypatch):
    from flask.testing import FlaskClient
    c = raw_client()
    sid = _system(c)
    import test_api  # noqa: F401
    # bypass the conftest shim: we want the raw response
    orig = FlaskClient.post.__closure__[0].cell_contents if FlaskClient.post.__closure__ else FlaskClient.post
    r = orig(c, f"/api/systems/{sid}/run", data={"files": (io.BytesIO(book()), "n.xlsx")}, content_type="multipart/form-data")
    body = r.get_json()
    assert r.status_code == 200 and body["success"] and body["job_id"] and "run" not in body
    j = _wait(c, r)
    assert j["state"] == "done" and j["result"]["rows"] == len(ROWS) and "timing" in j["result"]
    assert c.get("/api/jobs/nope").status_code == 404


def test_job_error_is_reported_not_raised():
    c = raw_client()
    sid = _system(c)
    bad = openpyxl.Workbook()
    bad.active.append(["x", "y"])
    r = c.post(f"/api/systems/{sid}/run", data={"files": (io.BytesIO(b(bad)), "bad.xlsx")}, content_type="multipart/form-data").get_json()
    assert not r["success"] and "no sheet looks like the input" in r["error"]


def test_inbox_file_runs_without_upload_and_token_reuses_it():
    c = raw_client()
    sid = _system(c)
    # find the store the test backend uses and drop a file into inbox/
    import test_api
    st = c.application.view_functions["api_heavy_list"].__globals__["store"]()
    st.put("heavy_uploads/big.xlsx", book())
    files = c.get("/api/heavy/uploads").get_json()["files"]
    assert [f["name"] for f in files] == ["big.xlsx"] and files[0]["mb"] is not None
    info = c.post(f"/api/systems/{sid}/inspect-input", data={"inbox": json.dumps(["big.xlsx"])}, content_type="multipart/form-data").get_json()
    assert info["success"] and info["rows"] == len(ROWS) and info["token"]
    # the stash holds a pointer, not a second copy of the file
    stashed = [p for p in st.list("tmp/") if info["token"] in p]
    assert stashed and all(len(st.get(p)) < 100 for p in stashed)
    # run from the token only: nothing is uploaded again
    r = c.post(f"/api/systems/{sid}/run", data={"token": info["token"]}, content_type="multipart/form-data").get_json()
    assert r["success"] and r["run"]["rows"] == len(ROWS)
    # straight from the inbox too
    r = c.post(f"/api/systems/{sid}/run", data={"inbox": json.dumps(["big.xlsx"])}, content_type="multipart/form-data").get_json()
    assert r["success"] and r["run"]["rows"] == len(ROWS)


def test_inbox_rejects_paths_and_missing_files():
    c = raw_client()
    sid = _system(c)
    for name in ("../secret.xlsx", "a/b.xlsx", ".hidden"):
        r = c.post(f"/api/systems/{sid}/run", data={"inbox": json.dumps([name])}, content_type="multipart/form-data").get_json()
        assert not r["success"] and "Bad file name" in r["error"]
    r = c.post(f"/api/systems/{sid}/run", data={"inbox": json.dumps(["nothere.xlsx"])}, content_type="multipart/form-data").get_json()
    assert not r["success"] and "not in the uploads" in r["error"]


def test_heavy_mode_upload_chunks_run_download_delete_clear():
    c = raw_client()
    sid = _system(c)
    data = book()
    n, size = 4, -(-len(data) // 4)
    for i in range(n):                                   # the browser sends the file in pieces
        r = c.post(f"/api/heavy/upload?name=big.xlsx&id=abc123xyz&index={i}&total={n}", data=data[i * size:(i + 1) * size],
                   content_type="application/octet-stream").get_json()
        assert r["success"] and r["done"] == (i == n - 1)
    assert [f["name"] for f in c.get("/api/heavy/uploads").get_json()["files"]] == ["big.xlsx"]
    r = c.post(f"/api/systems/{sid}/run", data={"inbox": json.dumps(["big.xlsx"]), "heavy": "1"}, content_type="multipart/form-data").get_json()
    run = r["run"]
    assert r["success"] and run["rows"] == len(ROWS) and run["heavy"] and run["download_files"] == [run["filename"]]
    downs = c.get("/api/heavy/downloads").get_json()["files"]
    assert [f["name"] for f in downs] == [run["filename"]]
    got = c.get("/api/heavy/downloads/file?name=" + run["filename"])
    assert got.status_code == 200 and got.data[:2] == b"PK"
    # the History download of that run reads from the download folder too
    assert c.get(f"/api/systems/{sid}/runs/{run['run_id']}/download").data[:2] == b"PK"
    # delete one / clear all
    assert c.post("/api/heavy/downloads/delete", json={"names": [run["filename"]]}).get_json()["deleted"] == 1
    assert c.get("/api/heavy/downloads").get_json()["files"] == []
    assert c.get(f"/api/systems/{sid}/runs/{run['run_id']}/download").status_code == 404
    assert c.post("/api/heavy/uploads/clear").get_json()["deleted"] == 1
    assert c.get("/api/heavy/uploads").get_json()["files"] == []
    # a missing chunk is reported
    r = c.post("/api/heavy/upload?name=x.xlsx&id=zzzzzz1&index=1&total=2", data=b"abc", content_type="application/octet-stream").get_json()
    assert not r["success"] and "chunk is missing" in r["error"]
    assert not c.post("/api/heavy/upload?name=../x.xlsx&id=zzzzzz2&index=0&total=1", data=b"a").get_json()["success"]
