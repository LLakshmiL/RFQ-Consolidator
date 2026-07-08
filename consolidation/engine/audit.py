"""Audit trail: records every decision the engine makes so nothing about
the consolidation is a black box.

One AuditEntry per (source_file, source_sheet) processed or skipped. This
is what makes the engine "auditable" rather than just "runs and produces
output": every sheet's fate (which profile matched, which header row was
used, which columns mapped/didn't, how many rows went in vs. came out) is
recorded and can be inspected without re-reading the source file.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd


@dataclass
class AuditEntry:
    source_file: str
    source_sheet: str
    profile: str | None = None
    confidence: int = 0
    status: str = "processed"  # processed | skipped_hidden | skipped_no_profile_match | skipped_no_header_row
    header_row: int | None = None
    mapped_columns: int = 0
    unmapped_headers: list[str] = field(default_factory=list)
    ambiguous_headers: list[str] = field(default_factory=list)
    rows_in: int = 0
    rows_out: int = 0
    notes: list[str] = field(default_factory=list)


class AuditLog:
    def __init__(self) -> None:
        self.entries: list[AuditEntry] = []

    def add(self, entry: AuditEntry) -> None:
        self.entries.append(entry)

    def to_dataframe(self) -> pd.DataFrame:
        rows = []
        for e in self.entries:
            rows.append(
                {
                    "source_file": e.source_file,
                    "source_sheet": e.source_sheet,
                    "profile": e.profile,
                    "confidence": e.confidence,
                    "status": e.status,
                    "header_row": e.header_row,
                    "mapped_columns": e.mapped_columns,
                    "unmapped_headers": "; ".join(e.unmapped_headers),
                    "ambiguous_headers": "; ".join(e.ambiguous_headers),
                    "rows_in": e.rows_in,
                    "rows_out": e.rows_out,
                    "notes": "; ".join(e.notes),
                }
            )
        return pd.DataFrame(rows)

    def write(self, path: str) -> None:
        self.to_dataframe().to_excel(path, sheet_name="Audit", index=False)
