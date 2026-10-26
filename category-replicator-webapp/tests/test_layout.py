"""Simple model: each output sheet = the input columns (in the sample INPUT's order) + that sheet's derived
columns (formulas from the sample).  Mapping and toggle sheets always come along."""
import io
import tempfile

import openpyxl

from replicator import engine, service
from replicator.store import LocalStore

TIERS = {"A": 0.1, "B": 0.2}
ROWS = [(1, "A", 1000.0, "EMEA"), (2, "B", 2000.0, "EMEA"), (3, "A", 300.0, "EMEA")]   # Region identical on every row


def b(wb):
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def files():
    inp = openpyxl.Workbook()
    ws = inp.active
    ws.title = "Data"
    ws.append(["ID", "Tier", "Amount", "Region"])
    [ws.append(list(r)) for r in ROWS]
    mp = openpyxl.Workbook()
    r = mp.active
    r.title = "Rates"
    r.append(["Tier", "Rate"])
    [r.append([k, v]) for k, v in TIERS.items()]
    t = mp.create_sheet("Toggle")
    t.append(["Label", "Value"])
    t.append(["Apply", True])
    # sample OUTPUT: pass-through columns deliberately in a different order than the input, derived after
    out = openpyxl.Workbook()
    o = out.active
    o.title = "Result"
    o.append(["Region", "Amount", "ID", "Tier", "Rate", "Fee"])
    cached = {}
    for i, (a, tier, amt, reg) in enumerate(ROWS, start=2):
        o.append([reg, amt, a, tier, f"=VLOOKUP(D{i},Rates!$A$2:$B$3,2,FALSE)", f"=IF(Toggle!$B$2,ROUND(B{i}*E{i},2),0)"])
        cached[f"E{i}"] = TIERS[tier]
        cached[f"F{i}"] = round(amt * TIERS[tier], 2)
    r2 = out.create_sheet("Rates")
    r2.append(["Tier", "Rate"])
    [r2.append([k, v]) for k, v in TIERS.items()]
    t2 = out.create_sheet("Toggle")
    t2.append(["Label", "Value"])
    t2.append(["Apply", True])
    return b(inp), b(mp), engine.inject_cached(b(out), ["Result"], {"Result": cached})


def system():
    inp, mp, out = files()
    st = LocalStore(tempfile.mkdtemp())
    did = service.new_draft(st)
    for n, d, h in (("in.xlsx", inp, "input"), ("map.xlsx", mp, "mapping"), ("out.xlsx", out, "output")):
        service.add_draft_files(st, did, [(n, d)], h)
    return st, service.create_system(st, did, {"name": "layout"})


NEW = [(7, "B", 500.0, "APAC"), (8, "A", 90.0, "AMER")]


def new_input():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["ID", "Tier", "Amount", "Region"])
    [ws.append(list(r)) for r in NEW]
    return b(wb)


def test_layout_is_exactly_the_samples():
    st, s = system()
    assert s["output"]["sheet_columns"]["Result"] == ["Region", "Amount", "ID", "Tier", "Rate", "Fee"]   # as in the sample
    by = {r["name"]: r for r in s["rules"]}
    assert by["Region"]["kind"] == "copy"                 # NOT the constant "EMEA"
    assert by["Rate"]["raw_formula"] == "=VLOOKUP(D2,Rates!$A$2:$B$3,2,FALSE)" and by["Rate"]["sample_row"] == 2
    assert by["Fee"]["raw_formula"] == "=IF(Toggle!$B$2,ROUND(B2*E2,2),0)"


def test_formulas_verbatim_and_dependent_sheets_in_both_modes():
    st, s = system()
    for live in (True, False):
        run = service.run_system(st, s["id"], [("new.xlsx", new_input())], live_formulas=live)
        p = run["preview"]["Result"]
        assert p["headers"] == ["Region", "Amount", "ID", "Tier", "Rate", "Fee"]
        rows = [dict(zip(p["headers"], r)) for r in p["data"]]
        assert [r["Region"] for r in rows] == ["APAC", "AMER"]            # carried from the new input
        assert [r["Fee"] for r in rows] == [100.0, 9.0]
        data, _ = service.get_run_file(st, s["id"], run["run_id"])
        wb = openpyxl.load_workbook(io.BytesIO(data))
        assert {"Rates", "Toggle"} <= set(wb.sheetnames)                  # mapping + toggle always included
        assert wb["Toggle"]["B2"].value is True
        if live:
            ws = wb["Result"]
            assert ws["E2"].value == "=VLOOKUP(D2,Rates!$A$2:$B$3,2,FALSE)"        # byte-for-byte the sample's
            assert ws["F2"].value == "=IF(Toggle!$B$2,ROUND(B2*E2,2),0)"
            assert ws["F3"].value == "=IF(Toggle!$B$2,ROUND(B3*E3,2),0)"           # next row: only the row moved
            assert ws["C2"].value == 7                                              # raw column stays a value


def test_shift_formula_is_excel_fill_down():
    sh = engine.shift_formula
    assert sh("=VLOOKUP(D2,Rates!$A$2:$B$3,2,FALSE)", 2, 5) == "=VLOOKUP(D5,Rates!$A$2:$B$3,2,FALSE)"
    assert sh("=C2-C1", 2, 4) == "=C4-C3"                                           # previous-row reference follows
    assert sh("=IF(A2=\"A2\",'My Sheet A1'!B2,LOG10(A2))", 2, 3) == "=IF(A3=\"A2\",'My Sheet A1'!B3,LOG10(A3))"
    assert sh("=SUM($C$2:$C$40)+Toggle!$B$2", 2, 9) == "=SUM($C$2:$C$40)+Toggle!$B$2"
    assert sh("=A2", 2, 2) == "=A2"


def test_sheet_origin_is_the_samples():
    """Sample with a title row and the table starting in column C: the output starts in the same place."""
    inp = openpyxl.Workbook()
    ws = inp.active
    ws.title = "Data"
    ws.append(["ID", "Amount"])
    ws.append([1, 100])
    ws.append([2, 250])
    out = openpyxl.Workbook()
    o = out.active
    o.title = "Out"
    o["A1"] = "Report"
    o.append([])
    o.append([])
    o["C4"], o["D4"], o["E4"] = "ID", "Amount", "Tax"
    for r, (i, a) in enumerate([(1, 100), (2, 250)], start=5):
        o.cell(r, 3, i); o.cell(r, 4, a); o.cell(r, 5, f"=D{r}*$G$1")
    o["G1"] = 0.1
    st = LocalStore(tempfile.mkdtemp())
    did = service.new_draft(st)
    service.add_draft_files(st, did, [("i.xlsx", b(inp))], "input")
    service.add_draft_files(st, did, [("o.xlsx", b(out))], "output")
    s = service.create_system(st, did, {"name": "origin"})
    assert s["output"]["sample_origin"]["Out"] == {"row": 3, "col": 2}
    new = openpyxl.Workbook()
    n = new.active
    n.title = "Data"
    n.append(["ID", "Amount"])
    n.append([7, 40])
    run = service.run_system(st, s["id"], [("n.xlsx", b(new))], live_formulas=True)
    data, _ = service.get_run_file(st, s["id"], run["run_id"])
    w = openpyxl.load_workbook(io.BytesIO(data))["Out"]
    assert [w["C4"].value, w["D4"].value, w["E4"].value] == ["ID", "Amount", "Tax"]
    assert [w["C5"].value, w["D5"].value] == [7, 40] and w["E5"].value == "=D5*$G$1"       # same cells as the sample
    assert w["G1"].value == 0.1                                                           # the parameter cell it points at
