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
3. **Run on new data** → drop new input → download the `.xlsx`.

Local preview: `python dev/serve.py` → http://localhost:5181. Sample data: `dev/sample_data/` (`python dev/make_dummy_data.py`).
Tests: `cd tests && python -m pytest -q`.
