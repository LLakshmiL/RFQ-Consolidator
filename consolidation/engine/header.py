"""Anchor-keyword based header row location.

Header position varies per sheet (title/metadata blocks of differing
length precede the real header row). Instead of assuming a fixed row
number, we scan a window of rows and pick the one that contains the most
of a profile's declared anchor keywords.
"""
from __future__ import annotations

import re
from dataclasses import dataclass


def normalize_text(value: object) -> str:
    """Collapse whitespace and trim. Does not alter case or strip numbering."""
    if value is None:
        return ""
    text = str(value)
    return re.sub(r"\s+", " ", text).strip()


@dataclass
class HeaderMatch:
    row_index: int
    hits: int
    matched_keywords: list[str]


def locate_header_row(
    ws,
    anchor_keywords: list[str],
    scan_rows: int = 40,
    scan_cols: int = 40,
    min_hits: int = 2,
) -> HeaderMatch | None:
    """Scan the top-left window of a sheet for the row most likely to be
    the header row, i.e. the row containing the most anchor keywords as
    case-insensitive substrings of its cell text.

    Returns None if no row in the window reaches ``min_hits``.
    """
    max_row = min(ws.max_row, scan_rows)
    max_col = min(ws.max_column, scan_cols)
    anchors_lower = [a.lower() for a in anchor_keywords]

    best: HeaderMatch | None = None
    for r in range(1, max_row + 1):
        row_texts = [
            normalize_text(ws.cell(row=r, column=c).value) for c in range(1, max_col + 1)
        ]
        row_texts_lower = [t.lower() for t in row_texts if t]
        matched = []
        for anchor, anchor_l in zip(anchor_keywords, anchors_lower):
            if any(anchor_l in cell for cell in row_texts_lower):
                matched.append(anchor)
        if best is None or len(matched) > best.hits:
            best = HeaderMatch(row_index=r, hits=len(matched), matched_keywords=matched)

    if best is None or best.hits < min_hits:
        return None
    return best


def score_sheet(ws, anchor_keywords: list[str], scan_rows: int = 40, scan_cols: int = 40) -> int:
    """Detection score for profile matching: the best anchor-hit count
    found for any row in the scan window (same search as locate_header_row,
    exposed separately so profile detection doesn't require a min_hits cutoff)."""
    match = locate_header_row(
        ws, anchor_keywords, scan_rows=scan_rows, scan_cols=scan_cols, min_hits=0
    )
    return match.hits if match else 0
