"""Filtering the NEW data of one run with several conditions (on top of the system's stored input filter)."""
import io

import openpyxl
import pytest

from replicator import service
from test_template import b, make, output_book

ROWS = [(1, "EU", "A", 100), (2, "EU", "B", 600), (3, "US", "A", 700), (4, "US", "B", 50), (5, "APAC", "A", 900), (6, "EU", "A", 800)]
HDR = ("ID", "Region", "Tier", "Amt")


def book(rows=ROWS):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(list(HDR))
    for r in rows:
        ws.append(list(r))
    return b(wb)


def system():
    out = openpyxl.Workbook()
    o = out.active
    o.title = "Out"
    o.append(list(HDR) + ["Fee"])
    for i, r in enumerate(ROWS, start=2):                  # every input row is in the sample: no stored input filter is learned
        o.append(list(r) + [f"=D{i}*0.1"])
    return make([("i.xlsx", book(), "input"), ("o.xlsx", b(out), "output")])


def ids(wb):
    ws = wb["Out"]
    return [ws.cell(r, 1).value for r in range(2, ws.max_row + 1)]


def cond(col, op, value=None):
    return {"col": col, "op": op, "value": value}


def test_several_conditions_all_must_match():
    st, s = system()
    run, wb = output_book(st, s, book(), run_filter={"match": "all", "conditions": [
        cond("Region", "==", "EU"), cond("Tier", "==", "A"), cond("Amt", ">=", 700)]})
    assert ids(wb) == [6] and any("Run filter (ALL of 3 condition(s)) kept 1 of 6" in n for n in run["notes"])


def test_any_of_the_conditions_and_list_values():
    st, s = system()
    run, wb = output_book(st, s, book(), run_filter={"match": "any", "conditions": [
        cond("Region", "in", "US, APAC"), cond("Amt", "<", 100)]})
    assert ids(wb) == [3, 4, 5]
    run, wb = output_book(st, s, book(), run_filter={"match": "all", "conditions": [cond("Region", "not in", "EU,US"), cond("ID", "not blank")]})
    assert ids(wb) == [5]


def test_it_stacks_on_top_of_the_systems_stored_input_filter():
    st, s = system()
    s = service.update_system(st, s["id"], {"filters": dict(s["filters"], pre=[cond("Amt", ">=", 600)])})
    run, wb = output_book(st, s, book(), run_filter={"match": "all", "conditions": [cond("Region", "==", "EU")]})
    assert ids(wb) == [2, 6]
    assert [n for n in run["notes"] if "Input filter kept" in n] and [n for n in run["notes"] if "Run filter" in n]


def test_no_filter_means_all_rows_and_the_run_records_what_was_used():
    st, s = system()
    run, wb = output_book(st, s, book())
    assert len(ids(wb)) == 6 and run["run_filter"] is None
    run, wb = output_book(st, s, book(), run_filter={"match": "all", "conditions": [cond("Tier", "==", "B")]})
    assert run["run_filter"]["conditions"][0]["col"] == "Tier"
    summary = [r for r in wb["Run summary"].iter_rows(values_only=True) if r[0] == "Run filter"][0][1]
    assert summary == "Tier == B"


def test_bad_filters_are_rejected_with_a_clear_message():
    st, s = system()
    with pytest.raises(ValueError, match="uses column 'Nope'"):
        service.run_system(st, s["id"], [("n.xlsx", book())], run_filter={"conditions": [cond("Nope", "==", 1)]})
    with pytest.raises(ValueError, match="unknown operator"):
        service.run_system(st, s["id"], [("n.xlsx", book())], run_filter={"conditions": [cond("Tier", "DROP", 1)]})
    with pytest.raises(ValueError, match="choose a column"):
        service.run_system(st, s["id"], [("n.xlsx", book())], run_filter={"conditions": [cond("", "==", 1)]})
    with pytest.raises(ValueError, match="match must be"):
        service.run_system(st, s["id"], [("n.xlsx", book())], run_filter={"match": "xor", "conditions": [cond("Tier", "==", 1)]})


def test_inspect_gives_the_real_columns_and_value_suggestions():
    st, s = system()
    info = service.inspect_input(st, s["id"], [("n.xlsx", book())])
    assert info["columns"] == list(HDR) and info["rows"] == 6
    assert info["values"]["Region"] == ["APAC", "EU", "US"] and "ID" in info["values"]


def test_through_the_api():
    from test_api import client
    c = client()
    did = c.post("/api/drafts").get_json()["draft_id"]
    out = openpyxl.Workbook()
    o = out.active
    o.title = "Out"
    o.append(list(HDR) + ["Fee"])
    for i, r in enumerate(ROWS, start=2):
        o.append(list(r) + [f"=D{i}*0.1"])
    for n, data, h in (("i.xlsx", book(), "input"), ("o.xlsx", b(out), "output")):
        c.post(f"/api/drafts/{did}/files", data={"files": (io.BytesIO(data), n), "hint": h}, content_type="multipart/form-data")
    sid = c.post(f"/api/drafts/{did}/learn", json={"name": "api"}).get_json()["system"]["id"]
    import json
    r = c.post(f"/api/systems/{sid}/inspect-input", data={"files": (io.BytesIO(book()), "n.xlsx")}, content_type="multipart/form-data").get_json()
    assert r["success"] and r["columns"] == list(HDR)
    flt = json.dumps({"match": "all", "conditions": [cond("Region", "==", "US"), cond("Amt", ">", 100)]})
    r = c.post(f"/api/systems/{sid}/run", data={"files": (io.BytesIO(book()), "n.xlsx"), "filters": flt}, content_type="multipart/form-data").get_json()
    assert r["success"] and r["run"]["rows"] == 1
    bad = c.post(f"/api/systems/{sid}/run", data={"files": (io.BytesIO(book()), "n.xlsx"), "filters": "{oops"}, content_type="multipart/form-data")
    assert bad.status_code == 400 and "not valid JSON" in bad.get_json()["error"]
