"""
Streamlit UI — consolidate multi-tab Excel workbooks into one editable view,
then push edits back to the exact source cell.

Run:  streamlit run app.py

Enforces the following operational conditions:
  1. Consolidate multiple tabs into one editable view; push edits back.
  2. Match fields by meaning (normalized header + optional aliases), not position.
  3. Every consolidated value traces back to (sheet, row, column).
  4. Edits land in the source cell via stored coordinate — not by value match.
  5. Dry-run preview before any write.
  6. Unmapped edits are surfaced, never silently dropped.
  7. Blank means "leave unchanged" OR "clear" — chosen explicitly.
  8. Respects worksheet protection; reports skipped cells.
  9. Every run emits an audit log (CSV + sheet inside the workbook).
 10. Original workbook is backed up before any destructive write.
"""

from __future__ import annotations

import re
from collections import Counter
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Optional

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


# ==================== BACKEND CONFIG ====================
# These were previously sidebar controls. They are now fixed defaults —
# edit the values here if the behavior ever needs to change.

# Merge headers that differ only by a leading numeric outline (4.19 / 4.20).
STRIP_NUMBER_PREFIX = True

# Drop satellite columns (Quantity, Currency, …) that are blank/zero across
# every row. Kept OFF so every satellite column is preserved in the output.
DROP_EMPTY_QUANTITY = False

# Sheet-name prefixes to skip (hidden helper sheets like 'Formulas_…').
SKIP_SHEET_PREFIXES = ["formulas_"]

# Alias groups: each inner list is one field; the first entry is canonical.
# Every alias in a group is matched to the first. Example:
#   ["Port of Loading", "POL", "Loading Port"]
ALIAS_GROUPS = [
    # ["4.20 E-Mail address", "4.19 E-Mail address", "4.18 E-Mail address"],
]

# Write to cells on protected sheets. The tender template protects the
# region tabs, so this must be True for edits to actually land.
IGNORE_PROTECTION = True
# ========================================================


def _build_alias_text() -> str:
    """Render ALIAS_GROUPS into the '|'-separated form parse_aliases expects."""
    return "\n".join(" | ".join(group) for group in ALIAS_GROUPS if len(group) >= 2)


# ---------------- Header normalization / matching by meaning ----------------

# Leading outline number followed by whitespace: '4.19 ', '2.3.1 ', or a
# plain integer like '17 '. Requiring trailing whitespace avoids eating the
# number in tokens like '20ft'.
_NUM_PREFIX = re.compile(r"^\s*\d+(?:\.\d+)*\s+")


def strip_number_prefix(s: str) -> str:
    """Remove leading numeric outline like '4.19 ', '2.3.1 ' or '17 '."""
    return _NUM_PREFIX.sub("", str(s or "")).strip()


def norm(s: str, strip_prefix: bool = False) -> str:
    """Canonical form of a header — case/whitespace/punctuation-agnostic.
    If strip_prefix=True, leading numeric outline is dropped first, so
    '4.19 E-Mail address' and '4.20 E-Mail address' collapse to the same key."""
    s = str(s or "")
    if strip_prefix:
        s = strip_number_prefix(s)
    s = s.lower()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_aliases(text: str, strip_prefix: bool = False) -> dict:
    """Parse alias groups: one group per line, '|'-separated.
    Returns {norm(any_alias): norm(first_alias_in_group)}."""
    out = {}
    for line in text.splitlines():
        parts = [p.strip() for p in line.split("|") if p.strip()]
        if len(parts) < 2:
            continue
        canonical_key = norm(parts[0], strip_prefix)
        for p in parts:
            out[norm(p, strip_prefix)] = canonical_key
    return out


def canonical(header: str, aliases: dict, strip_prefix: bool = False) -> str:
    n = norm(header, strip_prefix)
    return aliases.get(n, n)


# Separator used to attach a satellite/qualifier column (Quantity, Currency,
# UOM, …) to the primary column it belongs to: "Main Carriage — Quantity".
SAT_SEP = " — "


def _satellite_key(value) -> str:
    """Normalized key for satellite detection: drop the leading number
    outline and casefold, so '17 Currency', '19 Currency' and bare
    'Currency' all reduce to the same key."""
    return strip_number_prefix(value).strip().casefold()


# ---------------- Reading sheets with cell coordinates ----------------

def detect_header_row(ws, scan_rows: int = 40) -> int:
    best_row, best_score = 1, -1
    for r in range(1, min(ws.max_row, scan_rows) + 1):
        score = sum(1 for c in ws[r] if isinstance(c.value, str) and c.value.strip())
        if score > best_score:
            best_score, best_row = score, r
    return best_row


def detect_satellite_labels(wb, skip_prefixes, min_repeat: int = 2) -> set:
    """Return the set of satellite labels for a workbook.

    A satellite is any header label that repeats within a sheet's header row
    (e.g. 'Quantity' after every cost line, 'Currency' after every rate).
    Such a column has no standalone meaning — it qualifies the column to its
    left. Detection is global: a label that repeats on ANY sheet is treated
    as a satellite on EVERY sheet, so a lone occurrence on one tab is still
    qualified consistently instead of colliding with a real column elsewhere.
    """
    satellites = set()
    for name in wb.sheetnames:
        if any(name.lower().startswith(p) for p in skip_prefixes):
            continue
        ws = wb[name]
        hdr = detect_header_row(ws)
        keys = [_satellite_key(c.value) for c in ws[hdr]
                if c.value is not None and str(c.value).strip()]
        for key, count in Counter(keys).items():
            if key and count >= min_repeat:
                satellites.add(key)
    return satellites


def read_sheet_with_coords(ws, aliases: dict, strip_prefix: bool = False,
                           satellite_labels: set = None):
    """Return (display_by_canonical, rows).
    - display_by_canonical: {canonical_key: nice display header seen first}
    - rows: list of {canonical_key: (value, coord_str)} for each data row.
      coord_str format: 'SheetName!A5'

    Any header whose label is in `satellite_labels` (e.g. 'Quantity',
    'Currency') is treated as a satellite: it is renamed to
    '<preceding primary> — <satellite>' so it stays distinct from other
    satellites of the same type and lines up with the matching satellite on
    other sheets. If strip_prefix is on, the primary's leading number outline
    is dropped so renumbered siblings across tabs collapse together.
    """
    satellite_labels = satellite_labels or set()
    hdr_row = detect_header_row(ws)
    raw = [(str(c.value).strip() if c.value is not None else "") for c in ws[hdr_row]]

    display_headers, last_desc = [], ""
    for h in raw:
        if not h:
            display_headers.append("")
        elif _satellite_key(h) in satellite_labels:
            # Satellite column — qualify with the preceding primary header.
            sat_label = strip_number_prefix(h).strip() or h  # e.g. "Currency"
            if last_desc:
                qualifier = strip_number_prefix(last_desc) if strip_prefix else last_desc
                display_headers.append(f"{qualifier}{SAT_SEP}{sat_label}")
            else:
                display_headers.append(sat_label)
        else:
            last_desc = h
            display_headers.append(h)

    display_by_canonical: dict = {}
    canon_by_col: list = []
    for h in display_headers:
        if not h:
            canon_by_col.append(None)
            continue
        c = canonical(h, aliases, strip_prefix)
        canon_by_col.append(c)
        display_by_canonical.setdefault(c, h)

    rows = []
    for r_idx, row_cells in enumerate(
        ws.iter_rows(min_row=hdr_row + 1, values_only=False), start=hdr_row + 1
    ):
        vals = [c.value for c in row_cells]
        if all(v is None or v == "" for v in vals):
            continue
        record: dict = {}
        for i, cell in enumerate(row_cells):
            c = canon_by_col[i] if i < len(canon_by_col) else None
            if not c:
                continue
            coord = f"{ws.title}!{get_column_letter(cell.column)}{cell.row}"
            # If the same canonical appears twice in one sheet, keep first coord.
            record.setdefault(c, (cell.value, coord))
        rows.append(record)

    return display_by_canonical, rows


# ---------------- Consolidation ----------------

def _is_empty_satellite(v) -> bool:
    """A satellite cell (Quantity, Currency, …) counts as empty when blank
    or zero."""
    if v is None:
        return True
    s = str(v).strip()
    if s == "" or s.lower() == "none":
        return True
    try:
        return float(s) == 0.0
    except (TypeError, ValueError):
        return False


def consolidate(file_bytes: bytes, skip_prefixes: list, aliases: dict,
                strip_prefix: bool = False, drop_empty_quantity: bool = True):
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)

    # Pass 1 — find satellite labels (repeating headers) across the workbook.
    satellite_labels = detect_satellite_labels(wb, skip_prefixes)

    # Pass 2 — read each sheet, qualifying its satellite columns.
    per_sheet: dict = {}
    for name in wb.sheetnames:
        if any(name.lower().startswith(p) for p in skip_prefixes):
            continue
        display_map, rows = read_sheet_with_coords(
            wb[name], aliases, strip_prefix, satellite_labels)
        if display_map and rows:
            per_sheet[name] = (display_map, rows)

    if not per_sheet:
        return None, "No data sheets found in the workbook."

    n = len(per_sheet)
    canon_freq = Counter()
    for display_map, _ in per_sheet.values():
        canon_freq.update(display_map.keys())

    # Keep ALL headers (union across sheets). Column order = first sheet's
    # order, then any leftover canonicals from later sheets.
    keep = set(canon_freq.keys())
    ordered_canon, seen = [], set()
    for display_map, _ in per_sheet.values():
        for c in display_map:
            if c in keep and c not in seen:
                seen.add(c)
                ordered_canon.append(c)

    # Pick one display label per canonical (first sheet that had it wins).
    canon_to_display = {}
    for display_map, _ in per_sheet.values():
        for c, disp in display_map.items():
            canon_to_display.setdefault(c, disp)

    # Ensure display column names are unique.
    display_cols, used = [], set()
    for c in ordered_canon:
        base = canon_to_display[c]
        label = base
        i = 2
        while label in used:
            label = f"{base} ({i})"
            i += 1
        used.add(label)
        display_cols.append(label)
    display_to_canon = dict(zip(display_cols, ordered_canon))

    # Build the value DataFrame and the parallel coordinate DataFrame.
    data_records, coord_records = [], []
    for sheet_name, (_, rows) in per_sheet.items():
        for row in rows:
            data_rec = {"Source Sheet": sheet_name}
            coord_rec = {"Source Sheet": ""}  # not a source-mapped column
            for disp, canon in zip(display_cols, ordered_canon):
                if canon in row:
                    val, coord = row[canon]
                    data_rec[disp] = "" if val is None else val
                    coord_rec[disp] = coord
                else:
                    data_rec[disp] = ""
                    coord_rec[disp] = ""  # this cell has no source coordinate
            data_records.append(data_rec)
            coord_records.append(coord_rec)

    cols = ["Source Sheet"] + display_cols
    df = pd.DataFrame(data_records, columns=cols)
    coords = pd.DataFrame(coord_records, columns=cols)

    # Drop satellite columns (Quantity, Currency, …) that have no real data
    # anywhere. The source templates pair a satellite with every descriptive
    # field, even ones where it is meaningless (e.g. a Quantity after an
    # E-Mail address). Keeping empty satellites just clutters the view;
    # dropping them is safe because there is nothing to write back for a
    # column that was never populated.
    dropped_empty_qty = []
    if drop_empty_quantity:
        for col in display_cols:
            if SAT_SEP not in col:  # only satellites carry the separator
                continue
            if all(_is_empty_satellite(v) for v in df[col]):
                dropped_empty_qty.append(col)
        if dropped_empty_qty:
            df = df.drop(columns=dropped_empty_qty)
            coords = coords.drop(columns=dropped_empty_qty)
            for c in dropped_empty_qty:
                canon = display_to_canon.pop(c, None)
                if canon and canon in ordered_canon:
                    ordered_canon.remove(canon)
            display_cols = [c for c in display_cols if c not in dropped_empty_qty]

    meta = {
        "sheets": n,
        "sheet_names": list(per_sheet.keys()),
        "canon_freq": canon_freq,
        "kept_canonicals": ordered_canon,
        "dropped_canonicals": [],
        "dropped_empty_qty": dropped_empty_qty,
        "satellite_labels": sorted(satellite_labels),
        "display_to_canon": display_to_canon,
        "canon_to_display_all": {c: [] for c in canon_freq},
    }
    for display_map, _ in per_sheet.values():
        for c, disp in display_map.items():
            if disp not in meta["canon_to_display_all"][c]:
                meta["canon_to_display_all"][c].append(disp)
    return (df, coords, meta), None


# ---------------- Diff + write-back ----------------

YELLOW = PatternFill("solid", fgColor="FFFF00")


def _values_equal(a, b) -> bool:
    """True when a and b represent the same value, ignoring blank normalization
    and int/float coercion caused by round-tripping through Excel."""
    a_blank = a is None or (isinstance(a, str) and a.strip() == "") or (isinstance(a, float) and pd.isna(a))
    b_blank = b is None or (isinstance(b, str) and b.strip() == "") or (isinstance(b, float) and pd.isna(b))
    if a_blank and b_blank:
        return True
    if a_blank or b_blank:
        return False
    try:
        return float(a) == float(b)
    except (TypeError, ValueError):
        return str(a) == str(b)


def _shape_signature(df: pd.DataFrame) -> tuple:
    """Snapshot used to detect row/column tampering between Consolidate and Write Back."""
    return (
        tuple(df.columns),
        len(df),
        tuple(df["Source Sheet"].astype(str).tolist()) if "Source Sheet" in df.columns else (),
    )


def compute_planned(orig_df: pd.DataFrame, edited_df: pd.DataFrame,
                    coords_df: pd.DataFrame, source_wb) -> pd.DataFrame:
    """Build the plan of intended writes. Applies the write-back rules that
    can be evaluated without mutating anything:
      - Blank in edit  -> SKIP (blank = keep)
      - No coordinate  -> ORPHAN (edit typed in a column that doesn't map back)
      - Cell holds a formula -> SKIP (formula)
      - Source already equals new value -> SKIP (match)
      - Sheet protected & cell locked -> SKIP (protected)
      - Otherwise -> WILL WRITE
    """
    rows = []
    common_cols = [c for c in orig_df.columns
                   if c in edited_df.columns and c != "Source Sheet"]
    for idx in orig_df.index:
        for col in common_cols:
            old_view = orig_df.at[idx, col]
            new_val = edited_df.at[idx, col] if idx in edited_df.index else old_view
            if _values_equal(old_view, new_val):
                continue  # user didn't touch this cell
            src = orig_df.at[idx, "Source Sheet"]
            coord = coords_df.at[idx, col] if idx in coords_df.index else ""

            is_blank = new_val is None or str(new_val).strip() == ""
            if is_blank:
                status, source_val = "SKIP (blank = keep)", ""
            elif not coord:
                status, source_val = "ORPHAN", ""
            else:
                try:
                    sheet_name, cell_addr = str(coord).split("!", 1)
                    cell = source_wb[sheet_name][cell_addr]
                except Exception:
                    status, source_val = "ERROR (bad coord)", ""
                else:
                    source_val = cell.value
                    if cell.data_type == "f" or (
                        isinstance(source_val, str) and source_val.startswith("=")
                    ):
                        status = "SKIP (formula)"
                    elif _values_equal(source_val, new_val):
                        status = "SKIP (match — already equal)"
                    else:
                        ws = source_wb[sheet_name]
                        sheet_protected = bool(ws.protection.sheet)
                        cell_locked = bool(getattr(cell.protection, "locked", True))
                        if sheet_protected and cell_locked:
                            status = "SKIP (protected)"
                        else:
                            status = "WILL WRITE"

            rows.append({
                "Source Sheet": src,
                "Coordinate": coord,
                "Header": col,
                "Source value": "" if source_val is None else source_val,
                "New value": new_val,
                "Status": status,
            })
    return pd.DataFrame(rows, columns=[
        "Source Sheet", "Coordinate", "Header", "Source value", "New value", "Status"
    ])


def apply_planned(wb, planned: pd.DataFrame, ignore_protection: bool) -> pd.DataFrame:
    """Execute the plan. Only rows with status 'WILL WRITE' are written; all
    other statuses are carried through so the audit log shows exactly what
    happened. Written cells are highlighted yellow."""
    out = planned.copy()
    for i, row in out.iterrows():
        if row["Status"] != "WILL WRITE" and not (
            ignore_protection and row["Status"] == "SKIP (protected)"
        ):
            # non-writable status stays as-is (SKIP / ORPHAN / ERROR)
            if row["Status"].startswith("SKIP") or row["Status"] == "ORPHAN":
                out.at[i, "Status"] = row["Status"].replace("SKIP", "SKIPPED")
            continue
        try:
            sheet_name, cell_addr = str(row["Coordinate"]).split("!", 1)
        except ValueError:
            out.at[i, "Status"] = "ERROR (bad coord)"
            continue
        if sheet_name not in wb.sheetnames:
            out.at[i, "Status"] = "ERROR (missing sheet)"
            continue
        cell = wb[sheet_name][cell_addr]
        cell.value = row["New value"]
        cell.fill = YELLOW
        out.at[i, "Status"] = "WRITTEN"
    return out


def embed_audit_log(wb, audit: pd.DataFrame, run_id: str):
    name = f"AuditLog_{run_id}"[:31]  # Excel sheet name limit
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws.append(list(audit.columns))
    bold = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="305496")
    for cell in ws[1]:
        cell.font = bold
        cell.fill = fill
    for row in audit.itertuples(index=False, name=None):
        ws.append(["" if pd.isna(v) else v for v in row])
    ws.freeze_panes = "A2"


def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Consolidated") -> bytes:
    """Serialize a DataFrame to a standalone .xlsx workbook for external editing."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(list(df.columns))
    bold = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="305496")
    for cell in ws[1]:
        cell.font = bold
        cell.fill = fill
    for row in df.itertuples(index=False, name=None):
        ws.append(["" if v is None else v for v in row])
    ws.freeze_panes = "A2"
    for i, h in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = min(max(len(str(h)), 15) + 2, 40)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def read_edited_consolidated(file_bytes: bytes, expected_cols: list) -> tuple:
    """Load a user-edited consolidated workbook. Returns (df, error_message).
    The header row + Source Sheet column + row count must match what was
    originally exported — that's how each edited cell still maps to a
    coordinate in the parallel coords_df.
    """
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return None, "The uploaded file is empty."
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    if headers != list(expected_cols):
        return None, (
            "Column headers in the uploaded file do not match the consolidated view.\n"
            f"Expected {len(expected_cols)} columns, got {len(headers)}. "
            "Do not add, remove, rename, or reorder columns — only edit cell values."
        )
    data = [dict(zip(headers, r)) for r in rows[1:]]
    df = pd.DataFrame(data, columns=headers)
    # Normalize blanks: None → "" so shape signature matches.
    for col in df.columns:
        df[col] = df[col].apply(lambda v: "" if v is None else v)
    return df, None


def write_consolidated_sheet(wb, df: pd.DataFrame):
    if "Consolidated" in wb.sheetnames:
        del wb["Consolidated"]
    ws = wb.create_sheet("Consolidated", 0)
    ws.append(list(df.columns))
    bold = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="305496")
    for cell in ws[1]:
        cell.font = bold
        cell.fill = fill
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))
    ws.freeze_panes = "A2"
    for i, h in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = min(max(len(str(h)), 15) + 2, 40)


# ---------------- Streamlit UI ----------------

st.set_page_config(
    page_title="GEODIS · RFQ Consolidator",
    page_icon="Geodis Logo.webp" if Path("Geodis Logo.webp").exists() else None,
    layout="wide",
)

# ---------------- Brand styling ----------------
BRAND = "#2E1CB8"
BRAND_DARK = "#1D138A"
BRAND_TINT = "#F4F3FB"
INK = "#12123A"

st.markdown(f"""
<style>
    /* Buttons */
    .stButton > button {{
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.15s ease;
    }}
    .stButton > button[kind="primary"] {{
        background-color: {BRAND};
        border-color: {BRAND};
    }}
    .stButton > button[kind="primary"]:hover:not(:disabled) {{
        background-color: {BRAND_DARK};
        border-color: {BRAND_DARK};
    }}
    .stDownloadButton > button {{
        border-radius: 8px;
        font-weight: 600;
    }}
    /* Metric cards */
    div[data-testid="stMetric"] {{
        background-color: {BRAND_TINT};
        border-left: 4px solid {BRAND};
        padding: 12px 16px;
        border-radius: 6px;
    }}
    /* Section header line */
    h2, h3 {{
        border-bottom: 1px solid #E5E4F0;
        padding-bottom: 6px;
        margin-top: 1.5rem !important;
    }}
    /* Step badge */
    .step-badge {{
        display: inline-block;
        background: {BRAND};
        color: white;
        border-radius: 999px;
        width: 26px;
        height: 26px;
        text-align: center;
        line-height: 26px;
        font-weight: 700;
        margin-right: 8px;
        font-size: 14px;
    }}
</style>
""", unsafe_allow_html=True)

# ---------------- Header ----------------
_logo_candidates = ["Geodis Logo.webp", "geodis_logo.webp", "geodis_logo.png",
                    "geodis_logo.jpg", "geodis_logo.svg", "logo.png"]
_logo_path = next((p for p in _logo_candidates if Path(p).exists()), None)

hdr_cols = st.columns([1, 5], vertical_alignment="center")
with hdr_cols[0]:
    if _logo_path:
        st.image(_logo_path, width=190)
with hdr_cols[1]:
    st.markdown(
        f"<h1 style='color:{INK};margin:0;'>RFQ Consolidator</h1>",
        unsafe_allow_html=True,
    )

# ---------------- How it works ----------------
with st.expander("**How it works — read this before your first run**", expanded=False):
    st.markdown(f"""
### The 3-step workflow

<span class='step-badge'>1</span> **Consolidate** — Reads every region tab of your uploaded workbook and produces one big table where the same field from different tabs sits in one column. Every consolidated cell remembers its exact source (sheet + row + column).

<span class='step-badge'>2</span> **Merge (dry run)** — Compares your edits against the source. Writes **nothing**. Produces a plan showing exactly which cells will change, plus any `ORPHAN` rows (values typed in a column that doesn't exist on the source tab for that row). *Optional but recommended — it catches "value in the wrong column" mistakes before they're saved.*

<span class='step-badge'>3</span> **Write Back** — Pushes your edits into the source tabs. Changed cells are highlighted **yellow** so you can see what was updated. Downloads the updated workbook + a full audit log.

### Two ways to edit

- **Inline (Option A)**: edit cells directly in the table shown in this page. Simplest for small changes.
- **Excel round-trip (Option B)**: download the consolidated file, edit it in Excel, then re-upload. Best for many changes or when you want to use Excel's find/replace, filters, etc.

### Rules you should know

- **Empty cell = leave source unchanged.** Deleting a value in the consolidated view will *never* clear the source cell.
- **Do not add or delete rows/columns** in the consolidated view — Merge and Write Back will BLOCK if the shape changes.
- **Protected source sheets** are skipped by default. Tick *"Ignore sheet protection"* in the sidebar Advanced section if you own the workbook.
- **Formulas in source cells are never overwritten.**
- Always **download the backup** before Write Back — that action cannot be undone by the app.
""", unsafe_allow_html=True)

# ---- Workflow summary (centered, under "How it works") ----
st.markdown(
    f"""
    <div style='text-align:center; margin: 18px 0 8px;'>
        <h3 style='color:{INK}; margin:0; border:none;'>Consolidate → Edit → Merge → Write Back</h3>
        <p style='color:{INK}; opacity:0.65; font-size:14px; margin:6px 0 0;'>
            Click <strong>Consolidate</strong> first. Then edit inline or via Excel.
            <strong>Merge</strong> previews changes safely; <strong>Write Back</strong> commits them.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

ss = st.session_state
for k, default in [
    ("file_bytes", None), ("file_name", None), ("orig_df", None),
    ("edited_df", None), ("coords_df", None), ("meta", None),
    ("shape_sig", None), ("planned", None), ("committed", False),
    ("output_bytes", None), ("audit", None), ("run_id", None),
]:
    ss.setdefault(k, default)

# ---- Backend settings (no UI — configured at top of file) ----
strip_prefix = STRIP_NUMBER_PREFIX
drop_empty_qty = DROP_EMPTY_QUANTITY
skip_prefixes = [p.strip().lower() for p in SKIP_SHEET_PREFIXES if p.strip()]
aliases = parse_aliases(_build_alias_text(), strip_prefix)
ignore_protection = IGNORE_PROTECTION
# Blank cells always mean "leave unchanged" — never clear the source.
blank_policy = "Leave unchanged"


# ---- Upload ----
st.markdown(
    "<h3>Upload RFQ file</h3>",
    unsafe_allow_html=True,
)
uploaded = st.file_uploader(
    "Drag and drop your .xlsx / .xlsm file here",
    type=["xlsx", "xlsm"],
    label_visibility="collapsed",
)
if uploaded:
    if ss.file_name != uploaded.name or ss.file_bytes != uploaded.getvalue():
        ss.file_bytes = uploaded.getvalue()
        ss.file_name = uploaded.name
        ss.orig_df = ss.edited_df = ss.coords_df = ss.meta = None
        ss.planned = ss.shape_sig = ss.output_bytes = None
        ss.committed = False
    st.success(f"Loaded **{uploaded.name}** ({len(ss.file_bytes)/1024:.1f} KB)")

# ---- Three buttons ----
if ss.file_bytes:
    b1, b2, b3 = st.columns(3)
    do_consolidate = b1.button("① Consolidate", type="primary", use_container_width=True,
                               help="Read every tab and build one editable consolidated view.")
    do_merge = b2.button("② Merge (dry run)",
                         disabled=ss.orig_df is None, use_container_width=True,
                         help="Compare edits against source. Writes NOTHING. Flags ORPHANs.")
    do_write = b3.button("③ Write Back",
                         disabled=ss.orig_df is None, use_container_width=True,
                         help="Push edits into source cells. Highlights changes yellow. Irreversible.")

    # ---- Button 1: Consolidate ----
    if do_consolidate:
        with st.spinner("Reading tabs…"):
            result, err = consolidate(ss.file_bytes, skip_prefixes, aliases,
                                       strip_prefix, drop_empty_qty)
        if err:
            st.error(err)
        else:
            df, coords, meta = result
            ss.orig_df = df
            ss.edited_df = df.copy()
            ss.coords_df = coords
            ss.meta = meta
            ss.shape_sig = _shape_signature(df)
            ss.planned = None
            ss.committed = False
            ss.output_bytes = None

    # ---- Editable view + metrics + external round-trip ----
    if ss.orig_df is not None:
        meta = ss.meta
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Sheets consolidated", meta["sheets"])
        c2.metric("Headers kept (union)", len(meta["kept_canonicals"]))
        c3.metric("Consolidated rows", len(ss.orig_df))
        # Sparsity: fraction of mapped cells (i.e. cells with a real source coord).
        data_cells = ss.coords_df.drop(columns=["Source Sheet"])
        mapped = int((data_cells != "").sum().sum())
        total = data_cells.size
        c4.metric("Cells with source coord", f"{mapped}/{total}",
                  help=("Blanks in the consolidated view are expected: each row only "
                        "fills columns that existed on its source tab. Use Alias groups "
                        "in the sidebar to merge equivalent labels."))

        sats = meta.get("satellite_labels", [])
        if sats:
            st.caption(
                "Detected satellite columns (repeating labels attached to the column "
                f"on their left): **{', '.join(s.title() for s in sats)}**."
            )

        head_l, head_r = st.columns([3, 1], vertical_alignment="bottom")
        with head_l:
            st.subheader("Multi-Tab Data Consolidation")
        with head_r:
            st.download_button(
                "Download Consolidated file",
                data=df_to_xlsx_bytes(ss.orig_df),
                file_name=ss.file_name.rsplit(".", 1)[0] + "_Consolidated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="Download to edit in Excel. Do NOT add/rename/reorder columns or add/remove rows.",
            )
        st.caption(
            "Click any cell to change its value. **Empty cell = leave source unchanged.** "
            "Do not add or delete rows/columns — the Merge and Write Back steps will BLOCK if the shape changes."
        )
        ss.edited_df = st.data_editor(
            ss.edited_df,
            num_rows="fixed",
            use_container_width=True,
            height=420,
            disabled=["Source Sheet"],
            key="editor",
        )

        st.subheader("Post Consolidation Update Upload")
        st.caption(
            "Download the file using the button above, edit it in Excel, then upload it back here. "
            "**Do not rename, add, remove or reorder columns or rows** in Excel — only change cell values."
        )
        edited_upload = st.file_uploader(
            "Upload edited Consolidated file",
            type=["xlsx"],
            key="edited_upload",
            help="After editing in Excel, upload the file here. Merge will use these values.",
        )
        if edited_upload is not None:
            new_df, err = read_edited_consolidated(edited_upload.getvalue(),
                                                    list(ss.orig_df.columns))
            if err:
                st.error(err)
            else:
                if _shape_signature(new_df) != ss.shape_sig:
                    st.error(
                        "BLOCKED: the uploaded file's row count or Source Sheet column "
                        "does not match the original consolidated view. Do not add or "
                        "delete rows, and do not alter the Source Sheet column."
                    )
                else:
                    ss.edited_df = new_df
                    st.success(
                        f"Loaded edited consolidated file — {len(new_df)} rows will be "
                        "compared against the source. Click **② Merge (dry run)** next."
                    )

    # ---- Button 2: Merge (dry run) ----
    if do_merge:
        current_sig = _shape_signature(ss.edited_df)
        if current_sig != ss.shape_sig:
            st.error("BLOCKED: rows or columns of the Consolidated view have changed. "
                     "Press Consolidate again to rebuild, then redo your edits.")
            ss.planned = None
        else:
            wb_src = openpyxl.load_workbook(BytesIO(ss.file_bytes), data_only=False)
            ss.planned = compute_planned(ss.orig_df, ss.edited_df, ss.coords_df, wb_src)
            ss.committed = False

    if ss.planned is not None:
        st.subheader("Dry-run result — nothing has been written yet")
        with st.expander("What do these statuses mean?"):
            st.markdown("""
| Status | Meaning |
|---|---|
| `WILL WRITE` | Value will be written to the source cell when you click Write Back. |
| `ORPHAN` | You typed a value in a column that doesn't exist on this row's source tab. **Move it to the correct column on the consolidated view.** Nothing is written. |
| `SKIP (blank = keep)` | The consolidated cell is empty — source is left alone. |
| `SKIP (match — already equal)` | The source already contains this value — no work to do. |
| `SKIP (formula)` | Source cell contains a formula — never overwritten. |
| `SKIP (protected)` | Source sheet is protected — tick *Ignore sheet protection* in the sidebar Advanced section to write anyway. |
""")
        if ss.planned.empty:
            st.info("No edits detected in the consolidated view.")
        else:
            counts = ss.planned["Status"].value_counts().to_dict()
            will = int((ss.planned["Status"] == "WILL WRITE").sum())
            orph = int((ss.planned["Status"] == "ORPHAN").sum())
            m1, m2, m3 = st.columns(3)
            m1.metric("Total edits", len(ss.planned))
            m2.metric("Will write", will)
            m3.metric("ORPHAN (blocked)", orph)
            if orph:
                st.warning(
                    f"{orph} edit(s) landed in a column that does not exist on the source tab for that row. "
                    "These are the classic \"value in the wrong column\" cases — fix them on the Consolidated view "
                    "before Write Back. Nothing is silently dropped: every ORPHAN is listed below."
                )
            st.write("**Status breakdown:**", counts)
            st.dataframe(ss.planned, use_container_width=True, height=320)

    # ---- Button 3: Write Back ----
    if do_write:
        current_sig = _shape_signature(ss.edited_df)
        if current_sig != ss.shape_sig:
            st.error("BLOCKED: rows or columns of the Consolidated view have changed. "
                     "Press Consolidate again to rebuild, then redo your edits.")
        else:
            # Build the plan just-in-time if Merge wasn't run.
            plan = ss.planned
            if plan is None:
                wb_src = openpyxl.load_workbook(BytesIO(ss.file_bytes), data_only=False)
                plan = compute_planned(ss.orig_df, ss.edited_df, ss.coords_df, wb_src)
                ss.planned = plan
            if plan.empty:
                st.info("No edits detected — nothing to write.")
            else:
                wb = openpyxl.load_workbook(BytesIO(ss.file_bytes), data_only=False)
                audit = apply_planned(wb, plan, ignore_protection)
                write_consolidated_sheet(wb, ss.edited_df)
                ss.run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
                embed_audit_log(wb, audit, ss.run_id)
                out_buf = BytesIO()
                wb.save(out_buf)
                ss.output_bytes = out_buf.getvalue()
                ss.audit = audit
                ss.committed = True

    if ss.committed:
        st.subheader("Write Back complete")
        st.success(
            "Your changes have been applied to a copy of the workbook. Changed cells are highlighted **yellow** in every source tab."
        )
        st.write("**Write summary:**", ss.audit["Status"].value_counts().to_dict())
        ext = ".xlsm" if ss.file_name.lower().endswith(".xlsm") else ".xlsx"
        st.download_button(
            "Download UPDATED workbook",
            data=ss.output_bytes,
            file_name=ss.file_name.rsplit(".", 1)[0] + f"_Updated_{ss.run_id}" + ext,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.download_button(
            "Download audit log (CSV)",
            data=ss.audit.to_csv(index=False).encode("utf-8"),
            file_name=f"AuditLog_{ss.run_id}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # ---- Diagnostics ----
    if ss.meta:
        with st.expander("Kept headers (with per-sheet coverage & aliases)"):
            rows = []
            for c in ss.meta["kept_canonicals"]:
                picked = next((d for d, k in ss.meta["display_to_canon"].items() if k == c), c)
                all_seen = " | ".join(ss.meta["canon_to_display_all"].get(c, []))
                rows.append((picked, ss.meta["canon_freq"][c],
                             f"{ss.meta['canon_freq'][c]}/{ss.meta['sheets']}", all_seen))
            st.dataframe(
                pd.DataFrame(rows, columns=["Display header", "Count", "Coverage", "Source labels grouped"]),
                use_container_width=True,
            )
else:
    st.info("Upload an .xlsx file to get started.")
