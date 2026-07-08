"""Per-tab extraction: locate the header row (header_detect.py), then read
the header text and data rows below it. This is the baseline every later
edit gets diffed against — see diff_sync.py.
"""
from __future__ import annotations

from dataclasses import dataclass

from sync_studio.header_detect import HeaderRowMatch, locate_header_row, normalize_text


@dataclass
class TabData:
    sheet_name: str
    header_row: int
    detection: HeaderRowMatch
    columns: list[str]  # raw header text per column index, "" if blank
    rows: list[list[object]]  # data rows, each same length as columns
    row_numbers: list[int]  # actual workbook row number for each entry in `rows` (blank rows are skipped,
                             # so this is NOT simply header_row + index)


def _is_blank_row(values: list[object]) -> bool:
    return all(v is None or (isinstance(v, str) and not v.strip()) for v in values)


def parse_tab(ws, header_row: int | None = None) -> TabData:
    """Extract header + data for one sheet. If header_row is None, it is
    auto-detected; callers can pass an explicit row to support a manual
    per-tab override from the UI."""
    detection = locate_header_row(ws)
    resolved_header_row = header_row if header_row is not None else detection.row_index

    max_col = ws.max_column
    columns = [normalize_text(ws.cell(row=resolved_header_row, column=c).value) for c in range(1, max_col + 1)]

    rows: list[list[object]] = []
    row_numbers: list[int] = []
    for r in range(resolved_header_row + 1, ws.max_row + 1):
        values = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
        if _is_blank_row(values):
            continue
        rows.append(values)
        row_numbers.append(r)

    return TabData(
        sheet_name=ws.title,
        header_row=resolved_header_row,
        detection=detection,
        columns=columns,
        rows=rows,
        row_numbers=row_numbers,
    )


def parse_workbook(wb) -> list[TabData]:
    """Parse every visible sheet in the workbook."""
    return [parse_tab(wb[name]) for name in wb.sheetnames if wb[name].sheet_state == "visible"]
