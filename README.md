# ERS Consolidation

Streamlit app to consolidate multiple tabs of an Excel workbook into one editable view, then write edits back to the exact source cell.

## What it does

- **Consolidate** — reads every tab of an `.xlsx` / `.xlsm` workbook, auto-detects each tab's header row, matches columns by full header text (numbering prefixes like `3.2` are not treated as equivalent — the entire label is the key), and produces a single consolidated view. Every consolidated cell keeps a coordinate pointer back to its exact source (`SheetName!A5`).
- **Merge (dry run)** — compares your edits against the source workbook and produces a plan (`WILL WRITE` / `ORPHAN` / `SKIP (formula)` / `SKIP (match)` / `SKIP (protected)` / `SKIP (blank = keep)`). Nothing is written.
- **Write Back** — pushes edits into the raw source cells via the stored coordinates, highlights each written cell **yellow**, and emits an audit log (CSV + embedded sheet).

## Guarantees

1. Edits land in the exact source cell via stored coordinates — never by value match.
2. Blank cells in the consolidated view mean "leave unchanged" — the source is never accidentally cleared.
3. Values typed in a column that doesn't exist on the source tab for that row are surfaced as `ORPHAN` — never silently dropped.
4. Formulas and match-already-equal cells are skipped.
5. Sheet protection is honored (opt-in override in the sidebar for source owners).
6. Row/column shape changes to the consolidated view are BLOCKED at Merge and Write Back time.
7. Every run produces a downloadable audit log and an embedded `AuditLog_<timestamp>` sheet.
8. Backup download for the original workbook is always available before any write.

## Matching by meaning

Header labels are canonicalized (lowercased, whitespace/punctuation collapsed) but the **full text is preserved**. Different labels never merge automatically.

To treat differently-labeled fields as the same column, add an alias group in the sidebar:

```
Port of Loading | POL | Loading Port
4.19 E-Mail address | 4.20 E-Mail address
```

The first entry is the canonical label; every alias maps to it.

## Running

```bash
pip install streamlit openpyxl pandas
streamlit run app.py
```

## Flow

1. Upload the raw workbook.
2. Click **① Consolidate** to build the consolidated view.
3. Either edit inline in the Streamlit table, or download the consolidated file, edit in Excel, and re-upload.
4. Click **② Merge (dry run)** to preview all intended writes and catch ORPHANs.
5. Click **③ Write Back** to commit; download the updated workbook and audit log.

## Files

- `app.py` — the Streamlit application.
- `code.py` — headless CLI version of the consolidation logic.
