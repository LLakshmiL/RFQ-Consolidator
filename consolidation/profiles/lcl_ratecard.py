"""Profile for the 'lcl_ratecard' family (Sample1.xlsx): a title/metadata
block in the first few rows, then a header row with numbered field names
(e.g. '1 Origin Country [ISO country code]'), then one data row per lane
rate."""
from __future__ import annotations

from pathlib import Path

from consolidation.profiles.base import Profile, register

_CROSSWALK_PATH = Path(__file__).resolve().parent.parent / "crosswalks" / "lcl_ratecard.json"

_IDENTIFIER_FIELDS = [
    "origin_country",
    "origin_place",
    "port_of_loading",
    "port_of_destination",
    "destination_country",
    "destination_place",
    "service_type",
    "bulkiness_cbm_to_ton",
    "service_mode",
    "volume_cbm_annual",
    "pre_carriage_rate",
    "pre_carriage_currency",
    "pre_carriage_mode",
    "export_customs_clearance_fee",
    "export_customs_clearance_currency",
    "consolidation_fee",
    "consolidation_fee_currency",
    "port_origin_surcharge",
    "port_origin_surcharge_currency",
    "main_carriage_rate",
    "main_carriage_special_surcharges",
    "main_carriage_currency",
    "port_destination_surcharge",
    "port_destination_surcharge_currency",
    "demurrage_free_days",
    "demurrage_cost",
    "demurrage_currency",
    "import_customs_clearance_fee",
    "import_customs_clearance_currency",
    "deconsolidation_fee",
    "deconsolidation_fee_currency",
    "on_carriage_rate",
    "on_carriage_currency",
    "on_carriage_mode",
    "detention_free_days",
    "detention_cost",
    "detention_currency",
    "total_transit_time_days",
]

PROVENANCE_FIELDS = ["source_file", "source_sheet", "source_header_row", "source_data_row"]

PROFILE = register(
    Profile(
        name="lcl_ratecard",
        anchor_keywords=[
            "Origin Country",
            "Port of Loading",
            "Destination Country",
            "Port of Destination",
            "Type of service",
        ],
        crosswalk_path=_CROSSWALK_PATH,
        template_columns=_IDENTIFIER_FIELDS + PROVENANCE_FIELDS + ["extra_fields"],
        kind="flat",
        min_hits=3,
    )
)
