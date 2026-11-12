# Category Replicator (Dataiku Standard webapp)

Show the app **one worked example** — input file(s), mapping/toggle sheets and the output workbook you expect — and it keeps a
**template**. Give it new input and it produces the same workbook again.

## What a template is (nothing is calculated or guessed)
For every output sheet of your sample it remembers

1. **Layout** — exactly like the sample: same columns in the same order, same header row, same first column, same gaps.
2. **Input columns** — pasted as values from the input rows that pass the sheet's **filter**.
3. **Derived columns** — the header + the sample's formula, **stored verbatim** and filled down (only the row number moves, like
   Excel's fill-down; `$`-anchored refs, ranges, sheet names and text are never touched).
4. **Parameter cells** a formula points at (e.g. `$G$1`), written back at the same address.
5. **Mapping and Toggle sheets** — copied into every output workbook (a Toggle sheet keeps its exact cells; its Yes/No value
   comes from the updated mapping you attach to a run, or from a per-run override on the Run tab).
6. **Filters** — an input filter (when the sample has fewer rows than the input) and which input rows each sheet holds. They are
   worked out from the sample **on input columns only**, shown on the *Filters & outputs* tab, and editable.

Derived columns are never computed by the app: the output cells contain the formulas and Excel calculates them when the file is
opened. A derived column with **no formula in the sample** (values only) stays empty and is flagged.

## Deploy in Dataiku
See **DEPLOY.md**. In short: paste `frontend/body.html`, `style.css`, `app.js` and **`backend_single.py`** (rebuilt with
`python build_single.py`) into the webapp panes, set `REPLICATOR_FOLDER` (managed folder name/id) at the top of the Python pane,
restart the backend.

## Using it
1. **New system** → name → upload input / mapping / expected output (each sheet gets a role you can change; a workbook may mix roles;
   mapping/toggle sheets inside the output file are recognised).
2. **Template** tab shows every sheet: cell, column, "input" (pasted) or "formula" (the stored text).
3. **Formulas** tab: every derived cell (sheet, address, formula, what the next row reads). Edit a formula or a reference, use
   find & replace across sheets, add a formula to an empty column, then save. The original stays one click away (↺).
4. **Data types** tab: lock the type of chosen output columns (text / whole number / decimal / date). Every run forces those columns into
   that type and sets the Excel cell format; unreadable values are left empty and reported in the run notes.
5. **Run on new data** → drop new input → (optionally) add several filter conditions for this run, matching ALL or ANY, on any column of
   the file (its real columns and values are offered). Row counts update live: rows loaded → after the system's input filter → after your
   filters, how many rows each single condition matches, and how many rows each output sheet will get. Then download the `.xlsx`.

## File formats and big data
- **Read:** `.xlsx .xlsm .xltx .xltm` (formulas kept), `.xls .xlsb .ods` (values only; need `xlrd` / `pyxlsb` / `odfpy` installed), `.csv .tsv .txt`
  (delimiter and encoding detected, `.gz` allowed), `.parquet` (needs `pyarrow`). The expected-output **sample** must be `.xlsx/.xlsm`
  because only those keep readable formulas.
- **Speed:** big workbooks are read sheet by sheet in worker processes, several big files at the same time; big outputs are written by a
  streaming writer (`fastxlsx.py`) whose row chunks are encoded by worker processes. Formulas are generated while writing (never stored
  per cell in memory). Small outputs still use openpyxl. Pyspark is not used: the work is cell-by-cell text building and zip writing,
  which Spark cannot spread any better than local processes.
- **Limits:** a sheet with more rows than fit on one tab continues on `Name (2)`, `Name (3)`… (each with its header and formulas filled down
  from its own first row). Default 1,000,000 rows per tab (Excel's maximum is 1,048,576); set it per system under *Filters & outputs → Large
  outputs* or per run under *Large data options*.
- **Output in the input's format:** the `.xlsx` is always produced. If the input was `.xlsm/.xltx/.xltm` you also get the same workbook in that
  format. `.xls/.xlsb/.ods/.parquet` cannot be written here - the run says so. Extra files have their own download buttons.
- **Time:** every run reports total time, rows/second and the time of each stage (read, filter, build, write, save) in the result and History.
- **Environment knobs:** `REPLICATOR_WORKERS` (default CPU count, max 8; 1 = no parallelism), `REPLICATOR_ROWS_PER_SHEET`,
  `REPLICATOR_FAST_WRITE_ROWS` (default 100000), `REPLICATOR_PARALLEL_MIN_ROWS`, `REPLICATOR_PARALLEL_MIN_BYTES`, `REPLICATOR_MAX_UPLOAD_MB`
  (default 2000). Benchmark: `python dev/benchmark.py 1000000`.

Local preview: `python dev/serve.py` → http://localhost:5181. Sample data: `dev/sample_data/` (`python dev/make_dummy_data.py`).
Tests: `cd tests && python -m pytest -q`.
