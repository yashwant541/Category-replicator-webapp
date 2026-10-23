import io
import openpyxl
import pandas as pd
from replicator import engine
from replicator.tables import read_workbook
from test_learn import entries


def test_formula_output_without_cached_values():
    inp = openpyxl.Workbook()
    ws = inp.active
    ws.title = "Data"
    ws.append(["ID", "Tier", "Amount"])
    for i, (t, a) in enumerate([("A", 100), ("B", 200), ("A", 300), ("C", 50)], start=1):
        ws.append([i, t, a])
    out = openpyxl.Workbook()
    o = out.active
    o.title = "Out"
    o.append(["ID", "Tier", "Amount", "Weight", "RWA"])
    for r, (i, t, a) in enumerate([(1, "A", 100), (2, "B", 200), (3, "A", 300), (4, "C", 50)], start=2):
        o.append([i, t, a, f"=VLOOKUP(B{r},Weights!$A$2:$B$4,2,FALSE)", f"=ROUND(C{r}*D{r},1)"])
    w = out.create_sheet("Weights")
    w.append(["Tier", "W"])
    for t, v in [("A", 0.5), ("B", 1.5), ("C", 2)]:
        w.append([t, v])

    def b(x):
        bio = io.BytesIO(); x.save(bio); return bio.getvalue()
    ib, ob = b(inp), b(out)
    es = []
    for role, name, data in (("input", "in.xlsx", ib), ("output", "out.xlsx", ob)):
        for s in read_workbook(data, name):
            if role == "output" and s["sheet"] != "Out":
                continue
            es.append(dict(file=name, sheet=s["sheet"], role=role, df=s["df"], header_row=s["header_row"],
                           n_formulas=s["n_formulas"], bytes=data))
    system = engine.build_system(es, {"name": "f"})
    by = {r["name"]: r for r in system["rules"]}
    assert by["Weight"]["kind"] == "formula" and by["RWA"]["kind"] == "formula"
    assert "Weights" in system["tables"]
    X = pd.DataFrame({"ID": [10, 11], "Tier": ["B", "C"], "Amount": [1000, 20]})
    res = engine.apply_system(system, X)
    assert res["out"]["Weight"].tolist() == [1.5, 2]
    assert res["out"]["RWA"].tolist() == [1500.0, 40.0]


def test_prefilter_detected():
    from helpers import make_case
    inp, mp, out, rows, out_rows = make_case(n=60, split=False)
    # keep only Amount >= 1000 rows in the output
    wb = openpyxl.load_workbook(io.BytesIO(out))
    ws = wb.active
    hdr = [c.value for c in ws[1]]
    ai = hdr.index("Amount")
    keep = [[c.value for c in r] for r in ws.iter_rows(min_row=2) if r[ai].value >= 1000]
    ws.delete_rows(2, ws.max_row)
    for r in keep:
        ws.append(r)
    bio = io.BytesIO(); wb.save(bio)
    system = engine.build_system(entries(inp, mp, bio.getvalue()), {"name": "pf"})
    assert system["filters"]["pre"], system["warnings"]
    print(system["filters"]["pre"])
