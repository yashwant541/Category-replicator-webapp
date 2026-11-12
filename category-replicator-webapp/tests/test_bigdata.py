"""Big data: more file formats, the fast multi-process writer, sheets that continue on new tabs, timing."""
import gzip
import io
import tempfile

import openpyxl
import pytest

from replicator import engine, fastxlsx, perf, service
from replicator.store import LocalStore
from replicator.tables import file_ext, holds_formulas, read_workbook, supported
from test_run_filter import HDR, ROWS, book, ids, system
from test_template import b

REGIONS = ["EU", "US", "APAC", "LATAM"]


def many(n):
    return [(i, REGIONS[i % 4], "AB"[i % 2], 100 + i % 700) for i in range(1, n + 1)]


def run_file(st, s, rows, **options):
    r = service.run_system(st, s["id"], [("n.xlsx", book(rows))], options=options)
    data, _ = service.get_run_file(st, s["id"], r["run_id"])
    return r, openpyxl.load_workbook(io.BytesIO(data))


# ------------------------------------------------------------------------------- formulas
@pytest.mark.parametrize("f", [
    "=D2*0.1", "=VLOOKUP(B2,Rates!$A$2:$B$9,2,FALSE)", "=IF(D2>$G$1,\"D2 stays\",'Rate Sheet'!C2)", "=SUM(D2:D9)+E$2+$F3",
    "=A1+B1", "=IFERROR(INDEX(Map!B:B,MATCH(C2,Map!A:A,0)),\"n/a\")", "=\"it's 2\"&B2", "=D2", "", "=NOW()"])
@pytest.mark.parametrize("row", [2, 3, 17, 1048000])
def test_compiled_formulas_equal_shift_formula(f, row):
    c = engine.compile_formula(f, 2)
    assert fastxlsx.fill_formula(c, row) == engine.shift_formula(f, 2, row)


def test_a_row_that_would_move_above_row_one_stays_as_in_shift_formula():
    c = engine.compile_formula("=A2+B5", 6)
    assert fastxlsx.fill_formula(c, 3) == engine.shift_formula("=A2+B5", 6, 3)


# ------------------------------------------------------------------------------- the fast writer
def cells(wb):
    out = {}
    for ws in wb.worksheets:
        out[ws.title] = {(c.row, c.column): (c.value, c.number_format, c.font.b, c.fill.fgColor.rgb[-6:] if c.fill.fill_type else None)       # openpyxl writes alpha 00, we FF: Excel ignores it
                         for row in ws.iter_rows() for c in row if c.value is not None}
    return out


def test_the_fast_writer_writes_the_same_workbook_as_openpyxl():
    st, s = system()
    s = service.update_system(st, s["id"], {"dtypes": [{"sheet": "Out", "column": "Region", "type": "text"},
                                                        {"sheet": "Out", "column": "Amt", "type": "decimal", "decimals": 1}]})
    rows = many(40)
    r1, slow = run_file(st, s, rows, fast_writer=False)
    r2, fast = run_file(st, s, rows, fast_writer=True)
    assert r1["timing"]["writer"] == "openpyxl" and r2["timing"]["writer"] == "fast"
    assert [w.title for w in slow.worksheets] == [w.title for w in fast.worksheets]
    a, c = cells(slow), cells(fast)
    for name in a:
        if name == "Run summary":
            continue                                           # the time stamp differs
        assert a[name] == c[name], name
    for ws in slow.worksheets:
        f = fast[ws.title]
        assert f.freeze_panes == ws.freeze_panes
        assert {k: round(v.width) for k, v in f.column_dimensions.items()} == {k: round(v.width) for k, v in ws.column_dimensions.items()}


def test_fast_writer_through_worker_processes_gives_the_same_cells():
    st, s = system()
    rows = many(300)
    _, one = run_file(st, s, rows, fast_writer=True, workers=1)
    r, par = run_file(st, s, rows, fast_writer=True, workers=2)
    assert r["timing"]["workers"] == 2
    assert cells(one)["Out"] == cells(par)["Out"] and len(ids(par)) == 300


def test_every_kind_of_value_survives_the_fast_writer():
    import datetime as dt
    st, s = system()
    rows = [(1, "a & b <c> \"q\" 'x'", "line1\nline2", 1.5), (2, "plain", "  padded  ", 10 ** 15 + 3),
            (3, "é ñ 漢字 😀", dt.datetime(2026, 3, 4, 5, 6, 7), float("nan")), (4, True, dt.date(2026, 1, 2), None)]
    _, slow = run_file(st, s, rows, fast_writer=False)
    _, fast = run_file(st, s, rows, fast_writer=True)
    a, c = slow["Out"], fast["Out"]
    for r in range(1, 7):
        for col in range(1, 6):
            x, y = a.cell(r, col).value, c.cell(r, col).value
            assert x == y or (isinstance(x, str) and isinstance(y, str) and x.strip() == y.strip()), (r, col, x, y)
    assert "'=x" in fastxlsx.cell_xml("A1", "=x", 0)                           # pasted text that starts with = stays text
    assert "<f>x+1</f>" in fastxlsx.cell_xml("A1", "=x+1", 0, formula=True)    # a formula column's text is a formula


def test_the_fast_file_is_a_well_formed_package():
    st, s = system()
    r, wb = run_file(st, s, many(10), fast_writer=True)
    data, _ = service.get_run_file(st, s["id"], r["run_id"])
    import zipfile
    from xml.dom import minidom
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        assert z.testzip() is None
        for n in z.namelist():
            minidom.parseString(z.read(n))                                    # every part is well-formed XML
        assert b"fullCalcOnLoad" in z.read("xl/workbook.xml")
    assert wb["Out"]["E2"].value.startswith("=")                              # derived columns are formulas


# ------------------------------------------------------------------------------ splitting sheets
def test_a_big_sheet_continues_on_new_tabs_and_formulas_restart_on_each():
    st, s = system()
    for fast in (False, True):
        r, wb = run_file(st, s, many(2500), max_rows_per_sheet=1000, fast_writer=fast)
        assert [w.title for w in wb.worksheets][:3] == ["Out", "Out (2)", "Out (3)"]
        assert [wb[n].max_row - 1 for n in ("Out", "Out (2)", "Out (3)")] == [1000, 1000, 500]
        assert wb["Out"]["A2"].value == 1 and wb["Out (2)"]["A2"].value == 1001 and wb["Out (3)"]["A501"].value == 2500
        assert wb["Out (2)"]["E2"].value == "=D2*0.1" and wb["Out (3)"]["E501"].value == "=D501*0.1"      # each tab is filled down from row 2
        assert wb["Out (2)"].freeze_panes == "A2"
        assert r["rows"] == 2500 and r["sheets"] == {"Out": 2500} and r["split"] == {"Out": ["Out", "Out (2)", "Out (3)"]}
        assert any("continues on 2 more tab" in n for n in r["notes"])
        summary = {row[0]: row[1] for row in wb["Run summary"].iter_rows(min_row=5, max_row=7, values_only=True)}
        assert summary == {"Out": 1000, "Out (2)": 1000, "Out (3)": 500}


def test_the_limit_can_be_saved_with_the_system_and_is_capped_by_excel():
    st, s = system()
    s = service.update_system(st, s["id"], {"max_rows_per_sheet": 400})
    assert s["output"]["max_rows_per_sheet"] == 400
    r, wb = run_file(st, s, many(900))
    assert len(wb["Out"]["A"]) - 1 == 400 and "Out (3)" in wb.sheetnames
    s = service.update_system(st, s["id"], {"max_rows_per_sheet": ""})
    assert "max_rows_per_sheet" not in s["output"]
    parts, _ = engine.plan_parts(s, {"Out": {"def": {"origin": {"row": 0}}, "n": 2_000_000}})
    assert [p[2] - p[1] for p in parts.values()] == [1000000, 1000000]       # default 1,000,000 rows a tab
    parts, _ = engine.plan_parts(s, {"Out": {"def": {"origin": {"row": 5}}, "n": 2_000_000}}, max_rows=perf.EXCEL_MAX_ROWS)
    assert max(p[2] - p[1] for p in parts.values()) <= perf.EXCEL_MAX_ROWS - 6        # header rows count against Excel's limit


def test_continuation_names_never_clash_or_exceed_31_characters():
    sheets = {"A" * 31: {"def": {"origin": {"row": 0}}, "n": 30}, "Out (2)": {"def": {"origin": {"row": 0}}, "n": 1}}
    parts, _ = engine.plan_parts({}, sheets, max_rows=10)
    assert all(len(n) <= 31 for n in parts) and len(set(parts)) == len(parts) == 4 and "Out (2)" in parts


def test_bad_run_options_are_rejected():
    st, s = system()
    for bad in ({"max_rows_per_sheet": 0}, {"max_rows_per_sheet": "lots"}, {"workers": 0}, {"workers": 1000}):
        with pytest.raises(ValueError, match="Option"):
            service.run_system(st, s["id"], [("n.xlsx", book())], options=bad)


# --------------------------------------------------------------------------------- timing
def test_every_run_reports_its_time_and_row_counts():
    st, s = system()
    r, wb = run_file(st, s, many(500))
    t = r["timing"]
    assert t["rows_in"] == 500 and t["rows_out"] == 500 and t["seconds"] > 0 and t["rows_per_sec"] > 0
    assert set(t["stages"]) >= {"read", "filter", "build", "write", "save"} and t["tabs"] == 1 and t["output_mb"] >= 0
    assert abs(sum(t["stages"].values()) - t["seconds"]) < 0.5 * max(t["seconds"], 1)
    assert [x for x in wb["Run summary"].iter_rows(values_only=True) if x[0] == "Rows"][0][1] == "500 read, 500 written"
    assert service.list_runs(st, s["id"])[0]["timing"]["rows_in"] == 500
    assert perf.human_seconds(5.04) == "5.0 s" and perf.human_seconds(125) == "2 min 05 s"


# --------------------------------------------------------------------------------- formats
def csv_bytes(sep=",", enc="utf-8", bom=False):
    text = "\n".join(sep.join(map(str, r)) for r in [HDR] + ROWS) + "\n"
    return (b"\xef\xbb\xbf" if bom else b"") + text.encode(enc)


@pytest.mark.parametrize("name,data", [
    ("n.csv", csv_bytes()), ("n.CSV", csv_bytes(bom=True)), ("n.txt", csv_bytes(";")), ("n.tsv", csv_bytes("\t")),
    ("n.csv", csv_bytes("|")), ("n.csv.gz", gzip.compress(csv_bytes())), ("n.tsv.gz", gzip.compress(csv_bytes("\t")))])
def test_text_files_in_any_delimiter_are_read(name, data):
    [sh] = read_workbook(data, name)
    assert list(sh["df"].columns) == list(HDR) and len(sh["df"]) == 6


def test_text_files_in_windows_encoding_are_read():
    raw = "ID,Region,Tier,Amt\n1,Zürich,A,10\n2,São Paulo,B,20\n".encode("cp1252")
    [sh] = read_workbook(raw, "w.csv")
    assert list(sh["df"]["Region"]) == ["Zürich", "São Paulo"]


@pytest.mark.parametrize("name", ["n.xlsx", "n.xlsm", "n.xltx", "n.xltm", "N.XLSM"])
def test_the_whole_xlsx_family_is_read_including_formulas(name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(list(HDR) + ["Fee"])
    for i, r in enumerate(ROWS, start=2):
        ws.append(list(r) + [f"=D{i}*0.1"])
    [sh] = read_workbook(b(wb), name)
    assert len(sh["df"]) == 6 and sh["formula_cells"]["Fee"]["raw"] == "=D2*0.1" and holds_formulas(name)


def test_an_xlsm_or_xlsx_sample_makes_a_system_and_a_csv_input_runs_on_it():
    st, s = system()
    r = service.run_system(st, s["id"], [("new.csv", csv_bytes())])
    assert r["rows"] == 6 and r["timing"]["rows_in"] == 6


def test_unreadable_formats_say_what_to_do():
    assert file_ext("a.csv.gz") == ".csv" and file_ext("A.XLSB") == ".xlsb" and supported("x.xlsb") and not supported("x.pdf")
    with pytest.raises(ValueError, match="pyxlsb|xlsb"):
        read_workbook(b"not really an xlsb", "x.xlsb")
    with pytest.raises(ValueError):
        read_workbook(b"junk", "x.ods")


def test_a_sample_output_must_keep_its_formulas():
    st = LocalStore(tempfile.mkdtemp())
    did = service.new_draft(st)
    for name in ("o.xlsb", "o.xls", "o.csv", "o.ods"):
        with pytest.raises(ValueError, match="expected-output sample must be .xlsx"):
            service.add_draft_files(st, did, [(name, b"x")], "output")
    with pytest.raises(ValueError, match="unsupported file type"):
        service.add_draft_files(st, did, [("o.pdf", b"x")], "input")


# ------------------------------------------------------------------- reading in parallel
def test_big_workbooks_are_read_sheet_by_sheet_in_worker_processes(monkeypatch):
    monkeypatch.setenv("REPLICATOR_PARALLEL_MIN_BYTES", "1")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for n in ("North", "South", "East"):
        ws = wb.create_sheet(n)
        ws.append(list(HDR))
        for r in many(50):
            ws.append(list(r))
    data = b(wb)
    seq = read_workbook(data, "m.xlsx", workers=1)
    par = read_workbook(data, "m.xlsx", workers=3)
    assert [x["sheet"] for x in par] == [x["sheet"] for x in seq] == ["North", "South", "East"]
    for x, y in zip(seq, par):
        assert x["df"].equals(y["df"]) and x["header_row"] == y["header_row"]


def test_several_big_files_are_read_at_the_same_time(monkeypatch):
    from replicator.tables import read_many
    monkeypatch.setenv("REPLICATOR_PARALLEL_MIN_BYTES", "1")
    files = [("a.xlsx", book(many(30))), ("b.csv", csv_bytes()), ("c.xlsx", book(many(10)))]
    par = read_many(files, workers=3)
    seq = read_many(files, workers=1)
    assert [len(f[0]["df"]) for f in par] == [30, 6, 10]
    for x, y in zip(par, seq):
        assert x[0]["df"].equals(y[0]["df"])


# ------------------------------------------------------------------------------------ API
def test_options_and_limits_through_the_api():
    import json
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
    r = c.put(f"/api/systems/{sid}", json={"max_rows_per_sheet": 2000}).get_json()
    assert r["success"] and r["system"]["output"]["max_rows_per_sheet"] == 2000
    r = c.post(f"/api/systems/{sid}/run", data={"files": (io.BytesIO(book(many(30))), "n.xlsx"),
                                                "options": json.dumps({"max_rows_per_sheet": 10, "fast_writer": True})},
               content_type="multipart/form-data").get_json()
    assert r["success"] and r["run"]["split"] == {"Out": ["Out", "Out (2)", "Out (3)"]} and r["run"]["timing"]["writer"] == "fast"
    bad = c.post(f"/api/systems/{sid}/run", data={"files": (io.BytesIO(book()), "n.xlsx"), "options": "{oops"}, content_type="multipart/form-data")
    assert bad.status_code == 400 and "options" in bad.get_json()["error"]
    did2 = c.post("/api/drafts").get_json()["draft_id"]
    d = c.post(f"/api/drafts/{did2}/files", data={"files": (io.BytesIO(b"x"), "s.xlsb"), "hint": "output"}, content_type="multipart/form-data")
    assert d.status_code == 400 and "xlsx" in d.get_json()["error"]


# ------------------------------------------------------------------- the fast values reader
def _same_sheets(a, b):
    assert [x["sheet"] for x in a] == [x["sheet"] for x in b]
    for x, y in zip(a, b):
        assert x["header_row"] == y["header_row"] and x.get("col_offset") == y.get("col_offset"), x["sheet"]
        assert list(x["df"].columns) == list(y["df"].columns), x["sheet"]
        assert x["df"].astype(object).where(x["df"].notna(), None).values.tolist() == \
            y["df"].astype(object).where(y["df"].notna(), None).values.tolist(), x["sheet"]


def test_fast_reader_equals_openpyxl_on_every_sample_workbook():
    import pathlib
    root = pathlib.Path(__file__).resolve().parents[1] / "dev" / "sample_data"
    files = sorted(root.rglob("*.xlsx"))
    assert files
    for f in files:
        data = f.read_bytes()
        _same_sheets(read_workbook(data, f.name, workers=1), read_workbook(data, f.name, workers=1, values_only=True))


def test_fast_reader_equals_openpyxl_on_awkward_values():
    import datetime as dt
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Odd & <Sheet>"
    ws.append([None, None, None])                                        # leading blank row and column
    ws.append([None, "Name", "When", "Amount", "Flag", "Note"])
    ws.append([None, "a & b <c>", dt.datetime(2026, 3, 4, 5, 6, 7), 1.5, True, "line1\nline2"])
    ws.append([None, "é ñ 漢字 😀", dt.date(2026, 1, 2), 10 ** 12, False, "  padded  "])
    ws.append([None, None, None, None, None, None])                      # a blank row in the middle
    ws.append([None, "  ", None, -3, None, "x"])
    ws["B9"] = "after a gap"
    ws["D9"] = 0.1 + 0.2
    ws2 = wb.create_sheet("Second")
    ws2.append(["Key", "Val"])
    for i in range(5):
        ws2.append(["k%d" % i, i * 1.25])
    data = b(wb)
    _same_sheets(read_workbook(data, "o.xlsx", workers=1), read_workbook(data, "o.xlsx", workers=1, values_only=True))
    # and a workbook written by the fast writer itself (inline strings, our date styles)
    st, s = system()
    r = service.run_system(st, s["id"], [("n.xlsx", book(many(30)))], options={"fast_writer": True})
    out, _ = service.get_run_file(st, s["id"], r["run_id"])
    _same_sheets(read_workbook(out, "o.xlsx", workers=1), read_workbook(out, "o.xlsx", workers=1, values_only=True))


def test_fast_reader_falls_back_instead_of_guessing():
    from replicator import fastread
    with pytest.raises(fastread.Unsupported):
        fastread.read_sheets(b"not a zip")
    [sh] = read_workbook(book(many(5)), "n.xlsx", workers=1, values_only=True)
    assert len(sh["df"]) == 5


# ------------------------------------------------------------ output in the input's own format
def run_in(st, s, name, data, **options):
    r = service.run_system(st, s["id"], [(name, data)], options=options)
    return r, {e["filename"]: service.get_run_file(st, s["id"], r["run_id"], e["filename"])[0] for e in r["extra_outputs"]}


def test_a_csv_input_gives_the_xlsx_only():
    st, s = system()
    r, extra = run_in(st, s, "new.csv", csv_bytes(";"))
    assert r["extra_outputs"] == [] and extra == {} and r["rows"] == 6
    assert not any("CSV" in n or "Also produced" in n for n in r["notes"])
    xlsx, _ = service.get_run_file(st, s["id"], r["run_id"])
    assert openpyxl.load_workbook(io.BytesIO(xlsx))["Out"]["E2"].value == "=D2*0.1"


@pytest.mark.parametrize("ext", [".xlsm", ".xltx", ".xltm"])
def test_xlsm_and_template_inputs_give_the_same_workbook_in_that_format(ext):
    import zipfile
    st, s = system()
    r, extra = run_in(st, s, "new" + ext, book(many(8)))
    [(fname, blob)] = extra.items()
    assert fname.endswith(ext) and r["extra_outputs"][0]["format"] == ext
    with zipfile.ZipFile(io.BytesIO(blob)) as z:
        assert z.testzip() is None and b"/xl/workbook.xml" in z.read("[Content_Types].xml")
        assert b"spreadsheetml.sheet.main+xml" not in z.read("[Content_Types].xml")
    xlsx, _ = service.get_run_file(st, s["id"], r["run_id"])
    a, c = openpyxl.load_workbook(io.BytesIO(xlsx)), openpyxl.load_workbook(io.BytesIO(blob))
    assert cells(a) == cells(c) or [w.title for w in a.worksheets] == [w.title for w in c.worksheets]


def test_xlsx_input_gives_no_extra_file_and_unwritable_formats_say_so():
    from replicator import formats
    st, s = system()
    r, extra = run_in(st, s, "new.xlsx", book(many(5)))
    assert r["extra_outputs"] == [] and extra == {}
    out, notes = formats.original_outputs({}, {}, b"", [{"file": "a.xlsb"}, {"file": "b.XLS"}, {"file": "c.ods"}, {"file": "d.parquet"}], "x")
    assert out == [] and len(notes) == 4 and all("only the .xlsx was produced" in n for n in notes)


def test_extra_files_download_only_by_their_listed_name():
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
    run = c.post(f"/api/systems/{sid}/run", data={"files": (io.BytesIO(book()), "n.xlsm")}, content_type="multipart/form-data").get_json()["run"]
    name = run["extra_outputs"][0]["filename"]
    got = c.get(f"/api/systems/{sid}/runs/{run['run_id']}/download?file={name}")
    assert got.status_code == 200 and got.data[:2] == b"PK"
    assert c.get(f"/api/systems/{sid}/runs/{run['run_id']}/download?file=../meta.json").status_code == 404
    assert c.get(f"/api/systems/{sid}/runs/{run['run_id']}/download").data[:2] == b"PK"
