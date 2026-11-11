"""AND / OR groups, 'does not contain', and filters typed by column name (the column need not be in the sample)."""
import io

import openpyxl
import pytest

from replicator import engine, service
from test_run_filter import cond, ids, system
from test_template import b, output_book

# the sample (see test_run_filter.system) has no Import column: the NEW data does
NEW = [(1, "EU", "A", 100, "NONTP-ABCD"), (2, "EU", "B", 600, "ABCD-NONTP"), (3, "US", "A", 700, "ABCD"),
       (4, "US", "B", 50, "xx-nontp-yy"), (5, "APAC", "A", 900, None), (6, "EU", "A", 800, "TP-1")]


def book(rows=NEW, header=("ID", "Region", "Tier", "Amt", "Import")):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(list(header))
    for r in rows:
        ws.append(list(r))
    return b(wb)


def grp(match, *conds):
    return {"match": match, "conditions": list(conds)}


def run(st, s, flt=None, data=None):
    r, wb = output_book(st, s, data or book(), run_filter=flt)
    return r, ids(wb)


def test_does_not_contain_finds_the_text_anywhere_in_the_value_and_in_any_case():
    st, s = system()
    _, got = run(st, s, {"conditions": [cond("Import", "not contains", "NONTP")]})
    assert got == [3, 5, 6]                                    # NONTP-ABCD, ABCD-NONTP and xx-nontp-yy are all gone; a blank Import stays
    _, got = run(st, s, {"conditions": [cond("Import", "contains", "nontp")]})
    assert got == [1, 2, 4]


def test_or_group_inside_an_and_list():
    st, s = system()
    flt = {"match": "all", "conditions": [cond("Import", "not contains", "NONTP"),
                                           grp("any", cond("Region", "==", "EU"), cond("Amt", ">=", 900))]}
    assert run(st, s, flt)[1] == [5, 6]
    # (EU AND Tier A) OR Amt >= 900
    flt = {"match": "any", "conditions": [grp("all", cond("Region", "==", "EU"), cond("Tier", "==", "A")), cond("Amt", ">=", 900)]}
    assert run(st, s, flt)[1] == [1, 5, 6]


def test_groups_nest_and_describe_themselves():
    st, s = system()
    inner = grp("all", cond("Tier", "==", "A"), cond("Amt", ">", 100))
    flt = {"match": "all", "conditions": [grp("any", cond("Region", "==", "US"), inner)]}
    assert run(st, s, flt)[1] == [3, 4, 5, 6]
    r, wb = output_book(st, s, book(), run_filter=flt)
    summary = [x for x in wb["Run summary"].iter_rows(values_only=True) if x[0] == "Run filter"][0][1]
    assert summary == "(Region == US OR (Tier == A AND Amt > 100))"


def test_the_column_is_found_by_name_in_any_case_even_if_the_sample_never_had_it():
    st, s = system()
    for typed in ("import", "IMPORT", " Import ", "I_mport"):
        assert run(st, s, {"conditions": [cond(typed, "not contains", "NONTP")]})[1] == [3, 5, 6], typed
    r, _ = run(st, s, {"conditions": [cond("import", "contains", "TP-1")]})
    assert any("matched to 'Import'" in n for n in r["notes"])


def test_a_missing_column_is_a_clear_error_and_an_ambiguous_one_is_not_guessed():
    st, s = system()
    with pytest.raises(ValueError, match="uses column 'Imports'"):
        service.run_system(st, s["id"], [("n.xlsx", book())], run_filter={"conditions": [cond("Imports", "contains", "x")]})
    two = book([(1, "EU", "A", 100, "a", "b")], ("ID", "Region", "Tier", "Amt", "Import", "import "))
    with pytest.raises(ValueError, match="matches several columns"):
        service.preview_filter(st, s["id"], service.inspect_input(st, s["id"], [("n.xlsx", two)])["token"],
                               {"conditions": [cond("IMPORT", "contains", "a")]})


def test_the_stored_input_filter_takes_groups_or_and_a_typed_column():
    st, s = system()
    flt = dict(s["filters"], pre=[cond("import", "not contains", "NONTP"), grp("any", cond("Region", "==", "EU"), cond("Region", "==", "APAC"))])
    s = service.update_system(st, s["id"], {"filters": flt})
    assert s["filters"]["pre_match"] == "all" and "import" in [c["col"] for c in engine.condition_leaves(s["filters"]["pre"])]
    r, got = run(st, s)
    assert got == [5, 6] and any("Input filter kept 2 of 6" in n for n in r["notes"])
    s = service.update_system(st, s["id"], {"filters": dict(s["filters"], pre_match="any")})
    assert run(st, s)[1] == [1, 2, 3, 5, 6]                 # NOT NONTP (3, 5, 6) OR EU/APAC (1, 2, 5, 6)


def test_the_stored_filter_rejects_a_bad_tree_and_drops_empty_groups():
    st, s = system()
    with pytest.raises(ValueError, match="unknown operator"):
        service.update_system(st, s["id"], {"filters": dict(s["filters"], pre=[grp("all", cond("Region", "LIKE", "x"))])})
    with pytest.raises(ValueError, match="ALL or ANY"):
        service.update_system(st, s["id"], {"filters": dict(s["filters"], pre=[grp("xor", cond("Region", "==", "x"))])})
    with pytest.raises(ValueError, match="choose a column"):
        service.update_system(st, s["id"], {"filters": dict(s["filters"], pre=[cond("", "==", "x")])})
    s2 = service.update_system(st, s["id"], {"filters": dict(s["filters"], pre=[grp("any")])})
    assert s2["filters"]["pre"] == []
    deep = cond("Region", "==", "EU")
    for _ in range(8):
        deep = grp("all", deep)
    with pytest.raises(ValueError, match="nested too deeply"):
        service.update_system(st, s["id"], {"filters": dict(s["filters"], pre=[deep])})


def test_preview_with_groups_equals_the_run_and_reports_each_top_level_item():
    st, s = system()
    flt = {"match": "all", "conditions": [cond("Import", "not contains", "NONTP"), grp("any", cond("Region", "==", "EU"), cond("Amt", ">=", 900))]}
    token = service.inspect_input(st, s["id"], [("n.xlsx", book())])["token"]
    p = service.preview_filter(st, s["id"], token, flt)
    r, got = run(st, s, flt)
    assert p["after"] == r["rows"] == len(got) == 2
    assert [e["matches"] for e in p["each"]] == [3, 4]      # not-NONTP alone: 3 rows, the OR group alone: EU(1,2,6) + 900(5) = 4 rows


def test_custom_sheet_filters_use_groups_too():
    st, s = system()
    sp = {"enabled": True, "mode": "custom", "rest_sheet": "",
          "sheets": [{"name": "Out", "match": "all", "conditions": [grp("any", cond("Region", "==", "EU"), cond("Region", "==", "US"))]}]}
    s = service.update_system(st, s["id"], {"filters": dict(s["filters"], split=sp)})
    assert s["input"]["required"] and run(st, s)[1] == [1, 2, 3, 4, 6]


def test_through_the_api():
    from test_api import client
    c = client()
    did = c.post("/api/drafts").get_json()["draft_id"]
    out = openpyxl.Workbook()
    o = out.active
    o.title = "Out"
    o.append(["ID", "Region", "Tier", "Amt", "Fee"])
    for i, r in enumerate(NEW, start=2):
        o.append(list(r[:4]) + [f"=D{i}*0.1"])
    for n, data, h in (("i.xlsx", book([r[:4] for r in NEW], ("ID", "Region", "Tier", "Amt")), "input"), ("o.xlsx", b(out), "output")):
        c.post(f"/api/drafts/{did}/files", data={"files": (io.BytesIO(data), n), "hint": h}, content_type="multipart/form-data")
    sid = c.post(f"/api/drafts/{did}/learn", json={"name": "api"}).get_json()["system"]["id"]
    flt = {"pre": [{"col": "import", "op": "not contains", "value": "NONTP"}], "pre_match": "all", "split": {"enabled": False}}
    r = c.put(f"/api/systems/{sid}", json={"filters": flt}).get_json()
    assert r["success"] and r["system"]["filters"]["pre"][0]["op"] == "not contains"
    bad = c.put(f"/api/systems/{sid}", json={"filters": {"pre": [{"col": "x", "op": "??", "value": 1}]}})
    assert bad.status_code == 400 and "unknown operator" in bad.get_json()["error"]
