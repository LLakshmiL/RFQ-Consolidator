"""Diff the browser's current edited master state against the original
parse baseline (sync_studio.parser.TabData, as captured at upload time —
never a previous push's result) and produce a per-tab write-back plan.

Every push re-diffs against the same original baseline, so pushing twice
in a row without further edits produces the same plan both times (no
"diff of a diff" drift), and the uploaded file itself is never mutated.
"""
from __future__ import annotations

import datetime
from dataclasses import dataclass, field

from sync_studio.header_detect import normalize_text
from sync_studio.master import effective_columns
from sync_studio.parser import TabData


def _match_key(text: str) -> str:
    return normalize_text(text).lower()


def to_jsonable(value: object) -> object:
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.isoformat()
    return value


def _as_float(value: object) -> float | None:
    if isinstance(value, bool):
        return None  # keep True/False from colliding with 1.0/0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _values_differ(original: object, edited: object) -> bool:
    """True cell edits should be flagged; round-trip type noise should not.
    A dataframe-backed grid (e.g. Streamlit's data_editor, which infers an
    Arrow dtype per column) can turn an original int cell like 1 into the
    string "1" for a row that was never touched, simply because some other
    row in the same column holds a non-numeric value. Comparing "1" != 1
    would misreport that as a user edit, so numeric-looking values compare
    by number, and everything else compares by trimmed string form."""
    a = to_jsonable(original)
    a = "" if a is None else a
    b = "" if edited is None else edited
    if a == b:
        return False
    a_num, b_num = _as_float(a), _as_float(b)
    if a_num is not None and b_num is not None:
        return a_num != b_num
    return str(a).strip() != str(b).strip()


@dataclass
class CellChange:
    row_number: int
    col_index: int  # 1-based, in the tab's ORIGINAL column layout
    value: object


@dataclass
class NewColumnPlan:
    column_name: str
    values: dict[int, object]  # row_number -> value, only rows with a non-blank value


@dataclass
class TabSyncPlan:
    sheet_name: str
    modified_cells: list[CellChange] = field(default_factory=list)
    new_columns: list[NewColumnPlan] = field(default_factory=list)
    deleted_columns: list[str] = field(default_factory=list)


@dataclass
class SyncPlan:
    tabs: list[TabSyncPlan]


def build_sync_plan(
    baseline_tabs: list[TabData], edited_columns: list[str], edited_rows: list[dict]
) -> SyncPlan:
    edited_by_key: dict[tuple[str, int], dict] = {
        (row.get("__source_tab"), row.get("__source_row")): row for row in edited_rows
    }
    edited_columns_lower = {_match_key(c) for c in edited_columns}

    tab_plans: list[TabSyncPlan] = []
    for tab in baseline_tabs:
        eff_cols = effective_columns(tab.columns)
        col_index = {_match_key(c): idx for idx, c in enumerate(eff_cols) if c}

        modified_cells: list[CellChange] = []
        new_col_values: dict[str, dict[int, object]] = {}

        for values, row_number in zip(tab.rows, tab.row_numbers):
            edited_row = edited_by_key.get((tab.sheet_name, row_number))
            if edited_row is None:
                # Row-level delete isn't a feature this tool exposes; if a row
                # is missing from the edited payload, leave it untouched
                # rather than guessing at deletion.
                continue

            for col in edited_columns:
                key = _match_key(col)
                if key in col_index:
                    idx = col_index[key]
                    if _values_differ(values[idx], edited_row.get(col)):
                        modified_cells.append(
                            CellChange(row_number=row_number, col_index=idx + 1, value=edited_row.get(col))
                        )
                else:
                    edited_value = edited_row.get(col)
                    if edited_value is not None and edited_value != "":
                        new_col_values.setdefault(col, {})[row_number] = edited_value

        new_columns = [NewColumnPlan(column_name=col, values=vals) for col, vals in new_col_values.items()]
        deleted_columns = [c for c in eff_cols if c and _match_key(c) not in edited_columns_lower]

        tab_plans.append(
            TabSyncPlan(
                sheet_name=tab.sheet_name,
                modified_cells=modified_cells,
                new_columns=new_columns,
                deleted_columns=deleted_columns,
            )
        )

    return SyncPlan(tabs=tab_plans)
