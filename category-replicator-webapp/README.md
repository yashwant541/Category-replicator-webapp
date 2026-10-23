# Category Replicator (Dataiku Standard webapp)

Show the app **one worked example** — input file(s), mapping file(s) and the output you expect — and it
learns how every output column is built. Run the same *system* (category) on new data and it rebuilds the
output exactly, optionally split into several sheets by a filter.

## What it learns (per output column)
| Kind | Example |
|---|---|
| Excel formula (output file has live formulas) | `=VLOOKUP(B2,Weights!$A$2:$B$4,2,FALSE)` → translated & replayed |
| Copy / constant | `[ID]`, `"EUR"` |
| Mapping lookup (exact, 2-column key, approximate/band) | `Rates.Rate where Tier = [Tier]` |
| Maths | `ROUND([Amount] * [Rate], 2)`, `[Amount]*(1-[Crm])`, `[Fee]/[Amount]`, sums, max/min, looked-up operands |
| Join text | `[Region] & "-" & [ID]` |
| IF / bands | `IF([Amount] >= 1000, "High", "Low")` (round thresholds are picked) |
| Text / date parts | `LOWER([Tier])`, `year of [Date]` |
| Learned map | value → value table, when nothing simpler explains it |

Every rule is validated by running the **same** engine that production runs use, so what is learned is
what is replayed. A *replay check* re-runs the system on the training input and diffs it with your
expected output (per column, with example mismatches).

Also detected automatically: an **input filter** (output has fewer rows than input) and a **multi-sheet
split** (one output sheet per value of a column). Both are editable under *Filters & outputs*; the split
can also be set up front in the wizard (single / auto-detect / by column / custom filters per sheet).

## Deploy in Dataiku (short version — full checklist in DEPLOY.md; one-file backend: `backend_single.py`, rebuilt with `python build_single.py`)
1. Project → **Libraries → python/**: copy the whole `python-lib/` contents (`webapp_core/` and `replicator/`).
2. Create a **managed folder** for storage; set the project variable `REPLICATOR_FOLDER` to its id
   (or edit `FOLDER_FALLBACK` in `replicator/store.py`).
3. New **Standard webapp** → paste `frontend/body.html`, `frontend/style.css`, `frontend/app.js`
   and `backend/backend.py` into the HTML / CSS / JS / Python panes. Restart the backend.
4. After any upgrade replace the **whole** `python-lib/` and restart the backend.

Local preview (no Dataiku): `python dev/serve.py` → http://localhost:5181 (sample files are written to `dev/samples/`).
Tests: `cd tests && python -m pytest -q`.

## Using it
1. **New system** → name, how many outputs → upload input / mapping / expected output. Each sheet gets a role
   (Input / Mapping / Expected output / Ignore); a workbook can mix roles; several input sheets with the same
   columns are stacked.
2. Review the **Rules** tab. Anything amber/red can be corrected by editing the rule JSON
   (kind `formula` accepts a pandas expression such as `df['Amount'] * 0.05`). Rules can be reordered.
3. **Run on new data**: drop new input (same layout). Optionally attach updated mapping files (matched by
   sheet name) and an expected output to compare. Download the `.xlsx` (one sheet per split + `Run summary`).
   "Live Excel formulas (beta)" writes formulas instead of values and embeds the mapping sheets.

## Limits / notes
- Learning needs enough rows for a pattern to be unambiguous; constants and thresholds are inferred from data.
- Output files saved by tools that don't store cached formula results (e.g. openpyxl) still work: the
  formulas are used as-is (status *From formula*, no match % since there are no values to compare).
- Rows are matched between input and output by order, or by an identifying column carried into the output.
- Multi-key lookups and approximate (band) lookups are values-only in live-formula mode.
