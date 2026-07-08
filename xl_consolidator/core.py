"""Pandas-based reimplementation of a working VBA "two-way consolidation"
macro, mirroring its exact matching/filtering rules rather than a generic
outer-join (see xl_consolidator/streamlit_app.py for the Consolidate ->
Merge -> Write Back UI built on top of this):

  1. Consolidate: read every "source" tab (see `_tab_is_source`), matching
     columns across tabs by NORMALIZED header text so the same field merges
     into one column even when its section number differs (e.g.
     "4.18 E-Mail address" / "4.19 E-Mail address" -> one column), while
     text that genuinely differs stays separate even if the number
     coincides. Every included cell's exact (row, column) origin is
     recorded as "provenance" - this is what write-back uses later instead
     of re-deriving a column position by name.
  2. Merge (dry run): diff an edited copy of the consolidated table against
     the provenance-linked source cells. Blank edits never register as a
     change (fill-only - clearing a cell keeps the old value). Edits into a
     cell with no provenance link are reported as orphans rather than
     silently dropped.
  3. Write Back: apply just the real, non-blank, non-formula changes to a
     fresh copy of the original workbook, highlighting written cells
     yellow. Formula cells are always skipped, matching the VBA's
     `cell.HasFormula` guard.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

SOURCE_TAB_COL = "Source Tab"
ROW_TYPE_COL = "Row Type"
META_COLS = [SOURCE_TAB_COL, ROW_TYPE_COL]
CONSOLIDATED_SHEET_NAME = "Consolidated"

# These four constants encode the specific ratecard layout the VBA macro
# this module mirrors was built for. Adjust them if reusing this against a
# differently-labeled workbook.
HEADER_ANCHOR_TEXT = "Service package"          # literal header cell that marks the header row
FIELD_DESCRIPTION_HEADER_SUBSTRING = "field description"  # column whose text gates row inclusion
ROW_INCLUDE_SUBSTRING = "offer"                 # only rows whose Field description contains this are kept
EXCLUDED_TAB_PREFIXES = ("Formulas", "_bk_")    # helper/backup tabs, never treated as data sources
EXCLUDED_TAB_NAMES = {CONSOLIDATED_SHEET_NAME}  # a leftover output sheet from a prior run

_HIGHLIGHT = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
_NUMBER_PREFIX_RE = re.compile(r"^\d+\.\d+\s*")
DRY_RUN_COLUMNS = ["Action", "SrcTab", "SrcAddr", "OldValue", "NewValue", "ConsRow", "ConsCol"]


def _to_bytes(source) -> bytes:
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if isinstance(source, (str, Path)):
        return Path(source).read_bytes()
    return source.read()  # file-like / BytesIO


def _cell_text(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def norm_text(text: str) -> str:
    """Matching key for a header: strip a leading "N.N " section-number
    prefix, collapse whitespace, lowercase. Two headers normalize to the
    same key - and therefore merge into one consolidated column - only
    when their text matches after the number is stripped."""
    stripped = _NUMBER_PREFIX_RE.sub("", text.strip())
    return re.sub(r"\s+", " ", stripped).strip().lower()


def _col_key(header: str, left_header: str) -> str:
    """"Quantity" columns are ambiguous by text alone (every fee category
    has one) - key them by their preceding column's NORMALIZED text
    instead, so "Quantity" after "1.2 DG Surcharge" and "Quantity" after
    "2.2 DG Surcharge" merge into the same column despite the different
    section number, exactly like every other header does."""
    if header.strip().lower() == "quantity":
        return f"quantity::{norm_text(left_header)}"
    return norm_text(header)


def _tab_is_source(sheet_name: str) -> bool:
    if sheet_name.startswith(EXCLUDED_TAB_PREFIXES):
        return False
    return sheet_name not in EXCLUDED_TAB_NAMES


def _find_header_row(ws, scan_rows: int = 30, scan_cols: int = 6) -> int:
    """Locate the literal HEADER_ANCHOR_TEXT cell rather than using a
    generic structural heuristic - this workbook's header row always
    contains it, and anchoring on it avoids misdetecting one of the many
    metadata rows above it."""
    anchor = HEADER_ANCHOR_TEXT.lower()
    for r in range(1, scan_rows + 1):
        for c in range(1, scan_cols + 1):
            if _cell_text(ws.cell(row=r, column=c).value).lower() == anchor:
                return r
    return 0


def _last_nonblank_column(ws, header_row: int) -> int:
    last = 0
    for c in range(1, ws.max_column + 1):
        if _cell_text(ws.cell(row=header_row, column=c).value):
            last = c
    return last


def _last_nonblank_row(ws, col: int = 1) -> int:
    last = 0
    for r in range(1, ws.max_row + 1):
        if _cell_text(ws.cell(row=r, column=col).value):
            last = r
    return last


# --------------------------------------------------------------------------- #
# Stage 1: Consolidate
# --------------------------------------------------------------------------- #

@dataclass
class ProvenanceEntry:
    source_tab: str
    source_row: int
    source_col: int
    is_formula: bool

    @property
    def address(self) -> str:
        return f"{get_column_letter(self.source_col)}{self.source_row}"


@dataclass
class ConsolidationResult:
    consolidated: pd.DataFrame  # [Source Tab, Row Type, <every matched column>] - baseline, never mutated after this
    provenance: dict[tuple[int, str], ProvenanceEntry] = field(default_factory=dict)  # (row position, col) -> source cell
    signature: tuple = ()  # (row count, col count, column names) snapshot used to detect a reshaped table


def consolidate_workbook(source) -> ConsolidationResult:
    original_bytes = _to_bytes(source)
    # data_only=True to read computed values; a second, formula-preserving
    # load is needed alongside it because data_only strips formula text
    # entirely, and write-back must never overwrite a formula cell.
    wb_values = openpyxl.load_workbook(io.BytesIO(original_bytes), data_only=True)
    wb_formulas = openpyxl.load_workbook(io.BytesIO(original_bytes), data_only=False)

    source_sheets = [name for name in wb_values.sheetnames if _tab_is_source(name)]

    # ---- Pass 1: discover the ordered, deduplicated set of columns ---------
    key_order: list[str] = []
    display_name: dict[str, str] = {}
    tab_headers: dict[str, tuple[int, list[str]]] = {}

    for name in source_sheets:
        ws = wb_values[name]
        header_row = _find_header_row(ws)
        if header_row == 0:
            continue  # no recognizable header on this tab - skip it entirely
        last_col = _last_nonblank_column(ws, header_row)
        raw_headers = [_cell_text(ws.cell(row=header_row, column=c).value) for c in range(1, last_col + 1)]
        tab_headers[name] = (header_row, raw_headers)

        for idx, header in enumerate(raw_headers):
            if not header:
                continue
            left = raw_headers[idx - 1] if idx > 0 else ""
            key = _col_key(header, left)
            if key not in display_name:
                key_order.append(key)
                display_name[key] = f"Quantity [{left}]" if key.startswith("quantity::") else header

    if not key_order:
        raise ValueError(f'No header rows found (looked for "{HEADER_ANCHOR_TEXT}").')

    canonical_columns = [display_name[k] for k in key_order]
    all_columns = META_COLS + canonical_columns

    # ---- Pass 2: build rows + full-coverage provenance (incl. blanks) ------
    rows: list[dict] = []
    provenance: dict[tuple[int, str], ProvenanceEntry] = {}

    for name in source_sheets:
        if name not in tab_headers:
            continue
        header_row, raw_headers = tab_headers[name]
        ws_val = wb_values[name]
        ws_formula = wb_formulas[name]
        last_col = len(raw_headers)
        last_row = _last_nonblank_row(ws_val, col=1)

        field_desc_col = next(
            (i for i, h in enumerate(raw_headers, start=1) if FIELD_DESCRIPTION_HEADER_SUBSTRING in h.lower()),
            0,
        )

        for r in range(header_row + 1, last_row + 1):
            if field_desc_col:
                fd_text = _cell_text(ws_val.cell(row=r, column=field_desc_col).value).lower()
                if ROW_INCLUDE_SUBSTRING not in fd_text:
                    continue

            row_position = len(rows)
            row_dict = {SOURCE_TAB_COL: name, ROW_TYPE_COL: "Offer"}

            for c in range(1, last_col + 1):
                header = raw_headers[c - 1]
                if not header:
                    continue
                left = raw_headers[c - 2] if c >= 2 else ""
                key = _col_key(header, left)
                col_name = display_name[key]

                row_dict[col_name] = ws_val.cell(row=r, column=c).value
                provenance[(row_position, col_name)] = ProvenanceEntry(
                    source_tab=name,
                    source_row=r,
                    source_col=c,
                    is_formula=ws_formula.cell(row=r, column=c).data_type == "f",
                )

            rows.append(row_dict)

    consolidated = pd.DataFrame(rows, columns=all_columns)
    signature = (len(consolidated), len(consolidated.columns), tuple(consolidated.columns))

    return ConsolidationResult(consolidated=consolidated, provenance=provenance, signature=signature)


def export_consolidated_sheet(original_bytes: bytes, consolidated: pd.DataFrame) -> bytes:
    """Write `consolidated` into a brand-new "Consolidated" sheet in a
    fresh copy of the original workbook."""
    wb = openpyxl.load_workbook(io.BytesIO(original_bytes))
    if CONSOLIDATED_SHEET_NAME in wb.sheetnames:
        del wb[CONSOLIDATED_SHEET_NAME]
    ws = wb.create_sheet(CONSOLIDATED_SHEET_NAME)

    ws.append(list(consolidated.columns))
    for row in consolidated.itertuples(index=False):
        ws.append([None if pd.isna(v) else v for v in row])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# Stage 2: Merge (dry run)
# --------------------------------------------------------------------------- #

def _is_blank(value: object) -> bool:
    if value is None:
        return True
    # Covers float NaN, numpy NaN, pandas.NA and NaT - not just float('nan').
    # A grid widget backed by Arrow (e.g. Streamlit's data_editor) can hand
    # back missing values as pandas.NA rather than a bare float NaN, which
    # `isinstance(value, float)` would have missed entirely.
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass  # some objects (e.g. arrays) don't support a scalar isna check
    if isinstance(value, str) and not value.strip():
        return True
    return False


def _as_float(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _values_differ(old: object, new: object) -> bool:
    old_num, new_num = _as_float(old), _as_float(new)
    if old_num is not None and new_num is not None:
        return old_num != new_num
    return str("" if _is_blank(old) else old).strip() != str("" if _is_blank(new) else new).strip()


@dataclass
class MergePreview:
    entries: pd.DataFrame  # DRY_RUN_COLUMNS - one row per populated/mapped cell examined
    blocked: bool
    blocked_reason: str | None

    def _rows(self, action: str) -> pd.DataFrame:
        if self.entries.empty:
            return self.entries
        return self.entries[self.entries["Action"] == action].reset_index(drop=True)

    @property
    def writes(self) -> pd.DataFrame:
        return self._rows("WRITE")

    @property
    def orphans(self) -> pd.DataFrame:
        return self._rows("ORPHAN_NO_PROV")

    @property
    def skipped_formulas(self) -> pd.DataFrame:
        return self._rows("SKIP_FORMULA")


def merge_preview(edited: pd.DataFrame, result: ConsolidationResult) -> MergePreview:
    """Stage "Merge": diff the edited table against the provenance-linked
    source cells, without writing anything. Mirrors the VBA's
    DryRunWriteBack action categories: WRITE, UNCHANGED, BLANK_MAPPED,
    SKIP_FORMULA, ORPHAN_NO_PROV.
    """
    current_signature = (len(edited), len(edited.columns), tuple(edited.columns))
    if current_signature != result.signature:
        return MergePreview(
            entries=pd.DataFrame(columns=DRY_RUN_COLUMNS),
            blocked=True,
            blocked_reason="Consolidated table changed shape since consolidation. Click Consolidate to reset.",
        )

    data_columns = [c for c in edited.columns if c not in META_COLS]
    baseline = result.consolidated
    out_rows: list[dict] = []

    for row_position in range(len(edited)):
        edited_row = edited.iloc[row_position]
        for col in data_columns:
            new_value = edited_row[col]
            prov = result.provenance.get((row_position, col))

            if prov is None:
                if not _is_blank(new_value):
                    out_rows.append(
                        {
                            "Action": "ORPHAN_NO_PROV",
                            "SrcTab": None,
                            "SrcAddr": None,
                            "OldValue": None,
                            "NewValue": new_value,
                            "ConsRow": row_position,
                            "ConsCol": col,
                        }
                    )
                continue

            old_value = baseline.iloc[row_position][col]
            if prov.is_formula:
                action = "SKIP_FORMULA"
            elif _is_blank(new_value):
                action = "BLANK_MAPPED"
            elif not _values_differ(old_value, new_value):
                action = "UNCHANGED"
            else:
                action = "WRITE"

            out_rows.append(
                {
                    "Action": action,
                    "SrcTab": prov.source_tab,
                    "SrcAddr": prov.address,
                    "OldValue": old_value,
                    "NewValue": new_value,
                    "ConsRow": row_position,
                    "ConsCol": col,
                }
            )

    return MergePreview(entries=pd.DataFrame(out_rows, columns=DRY_RUN_COLUMNS), blocked=False, blocked_reason=None)


# --------------------------------------------------------------------------- #
# Stage 3: Write Back
# --------------------------------------------------------------------------- #

def write_back(
    original_bytes: bytes,
    provenance: dict[tuple[int, str], ProvenanceEntry],
    writes: pd.DataFrame,
) -> tuple[bytes, pd.DataFrame]:
    """Apply only the WRITE-action rows (from `MergePreview.writes`) to a
    fresh copy of the original workbook. Changed cells are filled yellow.
    Blank edits and formula cells never reach here - merge_preview already
    filtered them out into BLANK_MAPPED / SKIP_FORMULA.
    """
    wb = openpyxl.load_workbook(io.BytesIO(original_bytes))

    written_by_sheet: dict[str, int] = {}
    for _, entry in writes.iterrows():
        prov = provenance[(entry["ConsRow"], entry["ConsCol"])]
        ws = wb[prov.source_tab]
        cell = ws.cell(row=prov.source_row, column=prov.source_col)
        cell.value = entry["NewValue"]
        cell.fill = _HIGHLIGHT
        written_by_sheet[prov.source_tab] = written_by_sheet.get(prov.source_tab, 0) + 1

    audit = pd.DataFrame(
        [{"sheet": sheet, "cells_written": count} for sheet, count in written_by_sheet.items()],
        columns=["sheet", "cells_written"],
    )

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), audit
