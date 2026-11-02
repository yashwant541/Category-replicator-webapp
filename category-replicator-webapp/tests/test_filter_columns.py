"""The whole-data (input) filter can use every column of every input sheet; split filters are unchanged."""
import openpyxl
import pytest

from replicator import service
from test_template import b, make, output_book

A = [(1, "x", 10), (2, "y", 20), (3, "x", 30)]
BROWS = [(4, "y", 40, "EU"), (5, "x", 50, "US")]


def input_book():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    a = wb.create_sheet("PartA")
    a.append(["ID", "Grp", "Amt"])
    for r in A:
        a.append(list(r))
    bb = wb.create_sheet("PartB")                    # a different layout: one extra column, only here
    bb.append(["ID", "Grp", "Amt", "Region"])
    for r in BROWS:
        bb.append(list(r))
    return b(wb)


def system():
    out = openpyxl.Workbook()
    o = out.active
    o.title = "Out"
    o.append(["ID", "Grp", "Amt", "Tax"])
    for i, r in enumerate(A, start=2):
        o.append(list(r) + [f"=C{i}*0.1"])
    return make([("i.xlsx", input_book(), "input"), ("o.xlsx", b(out), "output")])


def new_input(rows, header):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(list(header))
    for r in rows:
        ws.append(list(r))
    return b(wb)


def test_filter_columns_cover_every_input_sheet_and_splits_keep_theirs():
    st, s = system()
    assert "Region" in s["input"]["filter_columns"]                       # only in the second input sheet
    assert "Region" not in s["input"]["columns"]                          # split conditions keep the main layout's columns
    assert s["input"]["columns"] == [c for c in s["input"]["filter_columns"] if c != "Region"]


def test_whole_data_filter_on_a_column_from_the_other_sheet_works_on_a_run():
    st, s = system()
    s = service.update_system(st, s["id"], {"filters": dict(s["filters"], pre=[{"col": "Region", "op": "==", "value": "EU"}])})
    run, wb = output_book(st, s, new_input([(7, "x", 5, "EU"), (8, "y", 6, "US"), (9, "x", 7, "EU")], ("ID", "Grp", "Amt", "Region")))
    assert run["rows"] == 2 and [wb["Out"].cell(r, 1).value for r in (2, 3)] == [7, 9]
    assert any("Input filter kept 2 of 3" in n for n in run["notes"])


def test_clear_message_when_the_run_input_lacks_the_filter_column():
    st, s = system()
    s = service.update_system(st, s["id"], {"filters": dict(s["filters"], pre=[{"col": "Region", "op": "==", "value": "EU"}])})
    with pytest.raises(ValueError, match="input filter uses column 'Region'"):
        service.run_system(st, s["id"], [("n.xlsx", new_input([(7, "x", 5)], ("ID", "Grp", "Amt")))])
