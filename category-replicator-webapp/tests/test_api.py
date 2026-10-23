import io
import os
import pathlib
import tempfile

import openpyxl
from flask import Flask
from helpers import make_case


def client():
    os.environ["REPLICATOR_LOCAL_ROOT"] = tempfile.mkdtemp()
    app = Flask(__name__)
    g = {"app": app, "__name__": "b"}
    src = (pathlib.Path(__file__).resolve().parents[1] / "backend" / "backend.py").read_text(encoding="utf-8")
    exec(compile(src, "backend.py", "exec"), g)
    return app.test_client()


def upload(c, path, name, data, **form):
    return c.post(path, data={"files": (io.BytesIO(data), name), **form},
                  content_type="multipart/form-data").get_json()


def test_full_flow():
    c = client()
    inp, mp, out, *_ = make_case(n=70)
    did = c.post("/api/drafts").get_json()["draft_id"]
    for name, data, hint in (("in.xlsx", inp, "input"), ("rates_map.xlsx", mp, "mapping"), ("out.xlsx", out, "output")):
        r = upload(c, f"/api/drafts/{did}/files", name, data, hint=hint)
        assert r["success"], r
    d = r["draft"]
    roles = {s["sheet"]: s["role"] for f in d["files"] for s in f["sheets"] if f["name"] == "out.xlsx"}
    assert set(roles.values()) == {"output"}
    r = c.post(f"/api/drafts/{did}/learn", json={"name": "Demo cat"}).get_json()
    assert r["success"], r
    s = r["system"]
    assert s["training"]["replay_match"] == 1.0, s["training"]
    sid = s["id"]
    assert c.get("/api/systems").get_json()["systems"][0]["name"] == "Demo cat"
    inp2, *_ = make_case(n=30, seed=5)
    r = upload(c, f"/api/systems/{sid}/run", "new.xlsx", inp2)
    assert r["success"], r
    run = r["run"]
    assert set(run["sheets"]) == {"EU", "US", "ASIA"} and run["rows"] == 30
    dl = c.get(f"/api/systems/{sid}/runs/{run['run_id']}/download")
    wb = openpyxl.load_workbook(io.BytesIO(dl.data))
    assert {"EU", "US", "ASIA", "Run summary"} <= set(wb.sheetnames)
    # turn the split off, round trip
    s["filters"]["split"]["enabled"] = False
    assert c.put(f"/api/systems/{sid}", json={"filters": s["filters"]}).get_json()["success"]
    r = upload(c, f"/api/systems/{sid}/run", "new.xlsx", inp2)
    assert len(r["run"]["sheets"]) == 1
    # wrong input
    bad = openpyxl.Workbook()
    bad.active.append(["x", "y"])
    bad.active.append([1, 2])
    bio = io.BytesIO()
    bad.save(bio)
    r = upload(c, f"/api/systems/{sid}/run", "bad.xlsx", bio.getvalue())
    assert not r["success"] and "expected input columns" in r["error"]
