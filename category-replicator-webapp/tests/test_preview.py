"""Live preview of the filters on loaded data: row counts before / after, and they equal a real run."""
import io
import json

import pytest

from replicator import service
from test_run_filter import ROWS, book, cond, ids, system
from test_template import output_book


def load(st, s):
    return service.inspect_input(st, s["id"], [("n.xlsx", book())])


def test_loading_the_file_already_gives_the_counts():
    st, s = system()
    info = load(st, s)
    assert info["token"] and info["rows"] == 6
    p = info["preview"]
    assert p["loaded"] == 6 and p["after"] == 6 and p["stages"] == [{"key": "loaded", "label": "Loaded", "rows": 6}]
    assert p["sheets"] == {"Out": 6}


def test_counts_before_and_after_each_filter_and_per_condition():
    st, s = system()
    token = load(st, s)["token"]
    flt = {"match": "all", "conditions": [cond("Region", "==", "EU"), cond("Tier", "==", "A")]}
    p = service.preview_filter(st, s["id"], token, flt)
    assert [(x["key"], x["rows"]) for x in p["stages"]] == [("loaded", 6), ("run_filter", 2)]
    assert [e["matches"] for e in p["each"]] == [3, 4] and p["after"] == 2 and p["sheets"] == {"Out": 2}
    p = service.preview_filter(st, s["id"], token, {"match": "any", "conditions": flt["conditions"]})
    assert p["after"] == 5                                              # EU(3) or Tier A(4), overlap 2
    assert service.preview_filter(st, s["id"], token, None)["after"] == 6


def test_stored_input_filter_is_its_own_step():
    st, s = system()
    s = service.update_system(st, s["id"], {"filters": dict(s["filters"], pre=[cond("Amt", ">=", 600)])})
    token = load(st, s)["token"]
    p = service.preview_filter(st, s["id"], token, {"conditions": [cond("Region", "==", "EU")]})
    assert [(x["key"], x["rows"]) for x in p["stages"]] == [("loaded", 6), ("input_filter", 4), ("run_filter", 2)]


def test_the_preview_equals_what_the_real_run_produces():
    st, s = system()
    token = load(st, s)["token"]
    for flt in (None, {"match": "all", "conditions": [cond("Region", "in", "US, APAC"), cond("Amt", ">", 100)]},
                {"match": "any", "conditions": [cond("Tier", "==", "B"), cond("Amt", ">=", 900)]}):
        p = service.preview_filter(st, s["id"], token, flt)
        run, wb = output_book(st, s, book(), run_filter=flt)
        assert p["sheets"] == run["sheets"] and p["after"] == run["rows"] == len(ids(wb))


def test_preview_errors_are_clear_and_a_lost_file_is_reported():
    st, s = system()
    token = load(st, s)["token"]
    with pytest.raises(ValueError, match="uses column 'Nope'"):
        service.preview_filter(st, s["id"], token, {"conditions": [cond("Nope", "==", 1)]})
    with pytest.raises(ValueError, match="Drop the new input file again"):
        service.preview_filter(st, s["id"], "20200101000000-abcdef", None)
    with pytest.raises(ValueError, match="Drop the new input file again"):
        service.preview_filter(st, s["id"], "../../etc/passwd", None)               # never touches other paths


def test_old_stashes_are_cleaned_up():
    st, s = system()
    st.put("tmp/20200101000000-abcdef/0__old.xlsx", b"x")
    load(st, s)
    assert not [p for p in st.list("tmp/") if "20200101000000-abcdef" in p]


def test_preview_through_the_api():
    from test_api import client
    import openpyxl
    from test_template import b
    c = client()
    did = c.post("/api/drafts").get_json()["draft_id"]
    out = openpyxl.Workbook()
    o = out.active
    o.title = "Out"
    o.append(["ID", "Region", "Tier", "Amt", "Fee"])
    for i, r in enumerate(ROWS, start=2):
        o.append(list(r) + [f"=D{i}*0.1"])
    for n, data, h in (("i.xlsx", book(), "input"), ("o.xlsx", b(out), "output")):
        c.post(f"/api/drafts/{did}/files", data={"files": (io.BytesIO(data), n), "hint": h}, content_type="multipart/form-data")
    sid = c.post(f"/api/drafts/{did}/learn", json={"name": "api"}).get_json()["system"]["id"]
    info = c.post(f"/api/systems/{sid}/inspect-input", data={"files": (io.BytesIO(book()), "n.xlsx")}, content_type="multipart/form-data").get_json()
    r = c.post(f"/api/systems/{sid}/preview-filter", json={"token": info["token"], "filters": {"match": "all", "conditions": [cond("Region", "==", "US")]}}).get_json()
    assert r["success"] and r["loaded"] == 6 and r["after"] == 2
    bad = c.post(f"/api/systems/{sid}/preview-filter", json={"token": "nope", "filters": None})
    assert bad.status_code == 400 and "Drop the new input file again" in bad.get_json()["error"]
