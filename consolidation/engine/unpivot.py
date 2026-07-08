"""Generic melt for the repeating (label, value) column-pair pattern.

Some source sheets ("bid matrix" style) list a fixed set of identifier
columns followed by a variable number of repeating column pairs: a label
column (e.g. "1.1 Pre Carriage ex Postcode 74xxx (incl. OTHC)") immediately
followed by a value column whose header is a generic word like "Quantity".
The number of pairs differs per sheet, so this cannot be modeled as fixed
output columns — it is unpivoted ("melted") into one output row per
non-blank (label, value) pair instead.

This module is generic over the pair *count*: it works whether a sheet has
6 pairs or 40, as long as the fixed identifier columns are resolved by the
crosswalk first (see engine/mapping.py) and the value column of each pair
matches ``value_header_match``.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from consolidation.engine.header import normalize_text
from consolidation.engine.mapping import ColumnMapping

_CODE_PREFIX_RE = re.compile(r"^(\d+(?:\.\d+)*)\s+(.*)$")


@dataclass
class ColumnPair:
    label_col: int
    value_col: int
    label_text: str


@dataclass
class UnpivotPlan:
    identifier_end_col: int  # last column index that belongs to the fixed identifier prefix
    pairs: list[ColumnPair]
    leftover_cols: list[int]  # columns in the repeating region that didn't form a valid pair


def find_repeating_region_start(column_mappings: list[ColumnMapping]) -> int:
    """Return the 1-based index of the first column after the contiguous
    run of crosswalk-mapped identifier columns starting at column 1."""
    start = 1
    for cm in column_mappings:
        if cm.status == "mapped":
            start = cm.col_index + 1
        else:
            break
    return start


def build_unpivot_plan(
    raw_headers: list[object],
    column_mappings: list[ColumnMapping],
    value_header_match: str = "quantity",
) -> UnpivotPlan:
    region_start = find_repeating_region_start(column_mappings)
    texts = [normalize_text(h) for h in raw_headers]

    pairs: list[ColumnPair] = []
    leftover: list[int] = []
    i = region_start
    n = len(texts)
    while i <= n:
        label_text = texts[i - 1]
        if not label_text:
            i += 1
            continue
        if i + 1 <= n and texts[i].lower() == value_header_match:
            pairs.append(ColumnPair(label_col=i, value_col=i + 1, label_text=label_text))
            i += 2
        else:
            leftover.append(i)
            i += 1

    return UnpivotPlan(identifier_end_col=region_start - 1, pairs=pairs, leftover_cols=leftover)


def split_attribute_code(label_text: str) -> tuple[str, str]:
    """'1.1 Pre Carriage ex Postcode 74xxx' -> ('1.1', 'Pre Carriage ex Postcode 74xxx')."""
    m = _CODE_PREFIX_RE.match(label_text)
    if m:
        return m.group(1), m.group(2)
    return "", label_text


def melt_row(row_values: list[object], plan: UnpivotPlan) -> list[dict]:
    """Emit one dict per non-blank (label, value) pair for a single data row."""
    out = []
    for pair in plan.pairs:
        value = row_values[pair.value_col - 1]
        if value is None or (isinstance(value, str) and not value.strip()):
            continue
        code, name = split_attribute_code(pair.label_text)
        out.append({"attribute_code": code, "attribute_name": name, "attribute_value": value})
    return out
