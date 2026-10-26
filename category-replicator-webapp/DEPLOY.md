# Deploy checklist — Category Replicator on Dataiku

No project-library upload needed: `backend_single.py` embeds everything (verified by a test that runs it with `python-lib/` off the path).

## One-time
- [ ] **Code env** (the webapp's, or project default) has: `pandas`, `numpy`, `openpyxl`, `flask` (already in Dataiku). Add `xlrd` only if you will upload legacy `.xls`. Python 3.9–3.11.
- [ ] **Managed folder** for storage (any connection; holds systems, training files, run outputs). Copy its id (shown in the folder URL).
- [ ] **Folder setting**: at the top of the Python pane set `REPLICATOR_FOLDER = "<folder name or id>"` (global variable). If left empty, the project variable `REPLICATOR_FOLDER` is used instead.
- [ ] The user running the webapp has **read + write** on that folder.

## Create the webapp
- [ ] Project → Webapps → **New webapp → Code webapp → Standard**.
- [ ] **HTML** pane ← `frontend/body.html`
- [ ] **CSS** pane ← `frontend/style.css`
- [ ] **JS** pane ← `frontend/app.js`
- [ ] **Python** pane ← `backend_single.py` (whole file; it already ends with the routes)
- [ ] Settings → pick the code env → **Save & restart backend**.

## Smoke test (5 min)
- [ ] Open the webapp: top bar shows `engine 2026.10a`. If it shows nothing, the backend didn't start — check the backend log.
- [ ] `<webapp-url>/api/health` (via the backend URL) returns `"storage": "DataikuStore"`. `LocalStore` means the variable isn't set.
- [ ] **New system** → upload your input, mapping and expected-output files → *Build the template*. Open the **Template** tab: every output sheet should list its input columns and its formulas.
- [ ] **Run on new data** with a second input. Download the `.xlsx`: formulas in the derived columns, mapping/Toggle sheets included.
- [ ] Reopen the webapp in a new tab: the system is still listed (persisted in the folder).

## Upgrading
Rebuild with `python build_single.py`, paste the new `backend_single.py` over the Python pane, restart the backend. Systems saved by older builds must be **Rebuilt** (button on the system page) or re-created: the template format changed.

## Troubleshooting
| Symptom | Cause / fix |
|---|---|
| `storage: LocalStore` | `REPLICATOR_FOLDER` not set or not visible to the webapp; restart backend after setting. |
| `ModuleNotFoundError: replicator_formulas` / `replicator` | You pasted `backend/backend.py` instead of `backend_single.py` (or didn't upload `python-lib/`). |
| Upload fails / very large files | Backend limit is 200 MB per file; Dataiku/nginx may impose a lower one. |
| "no sheet has the expected input columns" | Run input is missing columns the rules read; use **Input template** on the system page. |
| A derived column is empty | Your sample had values, not a formula, in that column - the app only replays formulas. Put the formula in the sample and rebuild. |
| Derived cells look blank in a viewer | The workbook holds formulas without results; open it in Excel (it calculates on load). |

Not yet exercised: running inside a real Dataiku instance (only the local Flask harness and the single-file build were tested).
