"""Apply a SyncPlan to a fresh copy of the originally uploaded workbook.

The uploaded bytes are only ever read here, never mutated in place — this
always produces a new in-memory workbook. Within one tab, operations run
in a fixed order so column indices stay valid throughout:

  1. write modified cell values (indices refer to the tab's *original*
     layout, so this must happen before anything that could shift columns)
  2. append new columns at the end (their position doesn't depend on the
     original layout, so order relative to deletions doesn't matter, but
     doing it before deletions keeps every index computed from the same
     still-intact original layout)
  3. delete removed columns, highest column index first, so deleting one
     doesn't shift the index of another still-pending deletion
"""
from __future__ import annotations

import io

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from sync_studio.diff_sync import SyncPlan
from sync_studio.master import effective_columns
from sync_studio.parser import TabData

_HIGHLIGHT = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


def apply_sync_plan(original_bytes: bytes, baseline_tabs: list[TabData], plan: SyncPlan) -> bytes:
    wb = load_workbook(io.BytesIO(original_bytes))
    baseline_by_name = {tab.sheet_name: tab for tab in baseline_tabs}

    audit_rows = []

    for tab_plan in plan.tabs:
        baseline = baseline_by_name[tab_plan.sheet_name]
        ws = wb[tab_plan.sheet_name]
        eff_cols = effective_columns(baseline.columns)

        for change in tab_plan.modified_cells:
            cell = ws.cell(row=change.row_number, column=change.col_index, value=change.value)
            cell.fill = _HIGHLIGHT

        next_col = ws.max_column + 1
        for new_col in tab_plan.new_columns:
            header_cell = ws.cell(row=baseline.header_row, column=next_col, value=new_col.column_name)
            header_cell.fill = _HIGHLIGHT
            for row_number, value in new_col.values.items():
                cell = ws.cell(row=row_number, column=next_col, value=value)
                cell.fill = _HIGHLIGHT
            next_col += 1

        delete_indices = sorted(
            (eff_cols.index(name) + 1 for name in tab_plan.deleted_columns if name in eff_cols),
            reverse=True,
        )
        for col_idx in delete_indices:
            ws.delete_cols(col_idx, 1)

        audit_rows.append(
            {
                "sheet": tab_plan.sheet_name,
                "cells_modified": len(tab_plan.modified_cells),
                "columns_added": ", ".join(c.column_name for c in tab_plan.new_columns),
                "columns_removed": ", ".join(tab_plan.deleted_columns),
            }
        )

    _write_audit_sheet(wb, audit_rows)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _write_audit_sheet(wb, audit_rows: list[dict]) -> None:
    ws = wb.create_sheet("Audit")
    ws.append(["Sheet", "Cells Modified", "Columns Added", "Columns Removed"])
    for row in audit_rows:
        ws.append([row["sheet"], row["cells_modified"], row["columns_added"], row["columns_removed"]])
