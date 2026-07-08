"""Generic, keyword-free header-row detection.

Unlike consolidation/engine/header.py (which scores rows against a known
list of anchor keywords per format profile), this scores rows on purely
structural signals so it works on *any* uploaded workbook with no
per-format configuration:

- a header row has several non-blank cells, each usually distinct from its
  neighbors (metadata/title rows tend to be one long label plus blanks)
- header cells are short; free-text description paragraphs are not headers
- the row immediately below a header is usually about as "wide" (as many
  non-blank cells) as the header itself

This is a heuristic, not a certainty — the caller should treat the result
as a starting point and let the user override it per sheet.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

_LONG_TEXT_CHARS = 60
_LONG_TEXT_PENALTY = 3.0
_DENSITY_BONUS = 2.0
_DENSITY_RATIO_THRESHOLD = 0.5
_MIN_NON_BLANK = 2
_HIGH_CONFIDENCE_SCORE = 4.0


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


@dataclass
class HeaderRowMatch:
    row_index: int
    score: float
    non_blank_count: int
    confidence: str  # "high" | "low"


def _row_texts(ws, row_idx: int, max_col: int) -> list[str]:
    return [normalize_text(ws.cell(row=row_idx, column=c).value) for c in range(1, max_col + 1)]


def _score_row(texts: list[str], next_texts: list[str]) -> tuple[float, int]:
    non_blank = [t for t in texts if t]
    non_blank_count = len(non_blank)
    if non_blank_count < _MIN_NON_BLANK:
        return float("-inf"), non_blank_count

    unique_ratio = len({t.lower() for t in non_blank}) / non_blank_count
    long_text_count = sum(1 for t in non_blank if len(t) > _LONG_TEXT_CHARS)
    score = non_blank_count * unique_ratio - long_text_count * _LONG_TEXT_PENALTY

    next_non_blank = sum(1 for t in next_texts if t)
    density_ratio = next_non_blank / non_blank_count if non_blank_count else 0
    if density_ratio >= _DENSITY_RATIO_THRESHOLD:
        score += _DENSITY_BONUS

    return score, non_blank_count


def locate_header_row(ws, scan_rows: int = 40, scan_cols: int = 80) -> HeaderRowMatch:
    """Best-guess header row for an arbitrary sheet. Falls back to row 1
    if nothing in the scan window looks header-like."""
    max_row = min(ws.max_row, scan_rows)
    max_col = min(ws.max_column, scan_cols) or 1

    best: HeaderRowMatch | None = None
    for r in range(1, max_row + 1):
        texts = _row_texts(ws, r, max_col)
        next_texts = _row_texts(ws, r + 1, max_col) if r < ws.max_row else []
        score, non_blank_count = _score_row(texts, next_texts)
        if best is None or score > best.score:
            confidence = "high" if score >= _HIGH_CONFIDENCE_SCORE else "low"
            best = HeaderRowMatch(
                row_index=r, score=score, non_blank_count=non_blank_count, confidence=confidence
            )

    if best is None or best.score == float("-inf"):
        return HeaderRowMatch(row_index=1, score=0.0, non_blank_count=0, confidence="low")
    return best
