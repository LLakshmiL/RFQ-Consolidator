"""Outer-union consolidation across tabs by literal (normalized) header
text — no crosswalk config. A column that only exists in some tabs is
still included in the master; rows from tabs lacking it just get a blank
for that column.

Columns with no header text (a blank header cell) have no name to match
on, so they are not represented in the master and are left untouched by
the rest of this tool — this is a deliberate limitation of "identify by
header text", not a bug.
"""
from __future__ import annotations

from dataclasses import dataclass

from sync_studio.header_detect import normalize_text
from sync_studio.parser import TabData


def _match_key(text: str) -> str:
    """Whitespace-and-case-insensitive join key. Parser.py already runs
    header text through normalize_text, but this stays defensive in case
    build_master is ever called on non-parser-sourced TabData."""
    return normalize_text(text).lower()


@dataclass
class MasterResult:
    columns: list[str]  # canonical display text, first-seen order across tabs
    rows: list[dict]  # each row: {"__source_tab": ..., "__source_row": ..., <col>: value, ...}


def effective_columns(columns: list[str]) -> list[str]:
    """A single tab can legitimately repeat a header verbatim (e.g. Sample1's
    ratecard tabs each have 11 columns literally named "Currency" for
    different fees). Since this tool has no crosswalk to disambiguate them
    by position, same-tab duplicates are suffixed " (2)", " (3)", ... in
    left-to-right order instead of silently collapsing into one column and
    losing the later ones' data."""
    seen: dict[str, int] = {}
    effective: list[str] = []
    for col in columns:
        if not col:
            effective.append("")
            continue
        key = _match_key(col)
        seen[key] = seen.get(key, 0) + 1
        n = seen[key]
        effective.append(col if n == 1 else f"{col} ({n})")
    return effective


def _column_index_by_key(effective_cols: list[str]) -> dict[str, int]:
    index: dict[str, int] = {}
    for idx, col in enumerate(effective_cols):
        if not col:
            continue
        index.setdefault(_match_key(col), idx)
    return index


def build_master(tabs: list[TabData]) -> MasterResult:
    canonical_columns: list[str] = []
    key_to_canonical: dict[str, str] = {}

    for tab in tabs:
        for col in effective_columns(tab.columns):
            if not col:
                continue
            key = _match_key(col)
            if key not in key_to_canonical:
                key_to_canonical[key] = col
                canonical_columns.append(col)

    rows: list[dict] = []
    for tab in tabs:
        col_index = _column_index_by_key(effective_columns(tab.columns))
        for values, row_number in zip(tab.rows, tab.row_numbers):
            row_dict: dict = {"__source_tab": tab.sheet_name, "__source_row": row_number}
            for col in canonical_columns:
                idx = col_index.get(_match_key(col))
                row_dict[col] = values[idx] if idx is not None else None
            rows.append(row_dict)

    return MasterResult(columns=canonical_columns, rows=rows)
