"""Profile for the 'bid_matrix' family (Sample2.xlsx): a metadata block
(matrix name, free-text description) in the first ~12 rows, then a header
row with 4 fixed identifier columns followed by a variable number of
repeating (charge/attribute label, "Quantity") column pairs."""
from __future__ import annotations

from pathlib import Path

from consolidation.engine.header import normalize_text
from consolidation.profiles.base import Profile, register

_CROSSWALK_PATH = Path(__file__).resolve().parent.parent / "crosswalks" / "bid_matrix.json"

_IDENTIFIER_FIELDS = ["record_id", "service_package", "row_description", "field_description"]

_UNPIVOT_FIELDS = ["attribute_code", "attribute_name", "attribute_value"]

PROVENANCE_FIELDS = ["source_file", "source_sheet", "source_header_row", "source_data_row"]


def extract_sheet_metadata(ws, scan_rows: int = 15) -> dict:
    """Find the 'Matrix: <name>' row and return its value as matrix_name."""
    max_row = min(ws.max_row, scan_rows)
    max_col = min(ws.max_column, 10)
    for r in range(1, max_row + 1):
        row = [normalize_text(ws.cell(row=r, column=c).value) for c in range(1, max_col + 1)]
        for i, cell in enumerate(row):
            if cell.rstrip(":").lower() == "matrix":
                rest = [v for v in row[i + 1 :] if v]
                return {"matrix_name": rest[0] if rest else ""}
    return {"matrix_name": ""}


PROFILE = register(
    Profile(
        name="bid_matrix",
        anchor_keywords=[
            "Service package",
            "Field description",
            "Zeilenbeschreibung",
            "Quantity",
        ],
        crosswalk_path=_CROSSWALK_PATH,
        template_columns=_IDENTIFIER_FIELDS
        + ["matrix_name"]
        + _UNPIVOT_FIELDS
        + PROVENANCE_FIELDS
        + ["extra_fields"],
        kind="unpivot",
        min_hits=3,
        value_header_match="quantity",
        extract_sheet_metadata=extract_sheet_metadata,
        text_columns=("attribute_code",),
    )
)
