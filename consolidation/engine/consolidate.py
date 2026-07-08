"""Top-level orchestration: workbooks -> sheets -> profile detection ->
extraction -> mapping -> (optional) unpivot -> per-profile master rows.

Consolidation happens *within* each detected profile only. Two sheets that
match the same profile (whether from the same workbook, a different
workbook, or a different customer entirely) land in the same master
sheet; sheets that match different profiles never mix.
"""
from __future__ import annotations

import json
from pathlib import Path

import openpyxl
import pandas as pd

from consolidation.engine.audit import AuditEntry, AuditLog
from consolidation.engine.header import locate_header_row, normalize_text, score_sheet
from consolidation.engine.mapping import map_headers
from consolidation.engine.unpivot import build_unpivot_plan, melt_row
from consolidation.profiles.base import Profile, all_profiles

# Importing these registers them with profiles.base's registry.
from consolidation.profiles import bid_matrix as _bid_matrix  # noqa: F401
from consolidation.profiles import lcl_ratecard as _lcl_ratecard  # noqa: F401


def _read_row(ws, row_idx: int, max_col: int) -> list[object]:
    return [ws.cell(row=row_idx, column=c).value for c in range(1, max_col + 1)]


def _is_blank_row(values: list[object]) -> bool:
    return all(v is None or (isinstance(v, str) and not v.strip()) for v in values)


def detect_profile(ws, profiles: list[Profile]) -> tuple[Profile | None, int]:
    best_profile = None
    best_score = 0
    for profile in profiles:
        score = score_sheet(ws, profile.anchor_keywords)
        if score > best_score:
            best_profile, best_score = profile, score
    if best_profile is not None and best_score >= best_profile.min_hits:
        return best_profile, best_score
    return None, best_score


def process_sheet(
    ws,
    profile: Profile,
    header_row: int,
    source_file: str,
    source_sheet: str,
    audit_entry: AuditEntry,
) -> list[dict]:
    crosswalk = profile.load_crosswalk()
    max_col = ws.max_column
    raw_headers = _read_row(ws, header_row, max_col)
    column_mappings = map_headers(raw_headers, crosswalk)

    audit_entry.mapped_columns = sum(1 for cm in column_mappings if cm.status == "mapped")
    audit_entry.ambiguous_headers = sorted(
        {cm.raw_header for cm in column_mappings if cm.status == "ambiguous_unresolved"}
    )

    sheet_metadata = profile.extract_sheet_metadata(ws) if profile.extract_sheet_metadata else {}

    output_rows: list[dict] = []
    rows_in = 0

    if profile.kind == "flat":
        audit_entry.unmapped_headers = sorted(
            {cm.raw_header for cm in column_mappings if cm.status == "unmapped"}
        )
        for r in range(header_row + 1, ws.max_row + 1):
            values = _read_row(ws, r, max_col)
            if _is_blank_row(values):
                continue
            rows_in += 1
            row_dict: dict = dict(sheet_metadata)
            extras: dict = {}
            for cm, value in zip(column_mappings, values):
                if cm.canonical_field:
                    row_dict[cm.canonical_field] = value
                elif cm.status != "blank" and value is not None and value != "":
                    extras[cm.raw_header] = value
            row_dict["source_file"] = source_file
            row_dict["source_sheet"] = source_sheet
            row_dict["source_header_row"] = header_row
            row_dict["source_data_row"] = r
            row_dict["extra_fields"] = json.dumps(extras) if extras else ""
            output_rows.append(row_dict)

    elif profile.kind == "unpivot":
        plan = build_unpivot_plan(raw_headers, column_mappings, profile.value_header_match)
        # Columns folded into (label, value) pairs are handled data, not lost
        # data — only genuinely unpaired leftover columns count as "unmapped".
        audit_entry.unmapped_headers = sorted(
            {normalize_text(raw_headers[c - 1]) for c in plan.leftover_cols}
        )
        if plan.leftover_cols:
            audit_entry.notes.append(
                f"{len(plan.leftover_cols)} unpaired column(s) in repeating region: {audit_entry.unmapped_headers}"
            )
        identifier_cols = [
            cm for cm in column_mappings if cm.canonical_field and cm.col_index <= plan.identifier_end_col
        ]

        for r in range(header_row + 1, ws.max_row + 1):
            values = _read_row(ws, r, max_col)
            if _is_blank_row(values):
                continue
            rows_in += 1
            identifier_values = {cm.canonical_field: values[cm.col_index - 1] for cm in identifier_cols}
            extras = {}
            for c in plan.leftover_cols:
                v = values[c - 1]
                if v is not None and v != "":
                    extras[normalize_text(raw_headers[c - 1])] = v
            extras_json = json.dumps(extras) if extras else ""

            melted = melt_row(values, plan)
            for attr in melted:
                row_dict = dict(sheet_metadata)
                row_dict.update(identifier_values)
                row_dict.update(attr)
                row_dict["source_file"] = source_file
                row_dict["source_sheet"] = source_sheet
                row_dict["source_header_row"] = header_row
                row_dict["source_data_row"] = r
                row_dict["extra_fields"] = extras_json
                output_rows.append(row_dict)

    audit_entry.rows_in = rows_in
    audit_entry.rows_out = len(output_rows)
    return output_rows


def consolidate(input_paths: list[str], output_dir: str) -> AuditLog:
    profiles = all_profiles()
    audit_log = AuditLog()
    master_rows: dict[str, list[dict]] = {p.name: [] for p in profiles}

    for path in input_paths:
        wb = openpyxl.load_workbook(path, data_only=True)
        source_file = Path(path).name
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            if ws.sheet_state != "visible":
                audit_log.add(
                    AuditEntry(
                        source_file=source_file,
                        source_sheet=sheet_name,
                        status="skipped_hidden",
                        notes=[f"sheet_state={ws.sheet_state}"],
                    )
                )
                continue

            profile, confidence = detect_profile(ws, profiles)
            if profile is None:
                audit_log.add(
                    AuditEntry(
                        source_file=source_file,
                        source_sheet=sheet_name,
                        status="skipped_no_profile_match",
                        confidence=confidence,
                    )
                )
                continue

            header_match = locate_header_row(ws, profile.anchor_keywords, min_hits=profile.min_hits)
            if header_match is None:
                audit_log.add(
                    AuditEntry(
                        source_file=source_file,
                        source_sheet=sheet_name,
                        profile=profile.name,
                        confidence=confidence,
                        status="skipped_no_header_row",
                    )
                )
                continue

            entry = AuditEntry(
                source_file=source_file,
                source_sheet=sheet_name,
                profile=profile.name,
                confidence=confidence,
                header_row=header_match.row_index,
                status="processed",
            )
            rows = process_sheet(ws, profile, header_match.row_index, source_file, sheet_name, entry)
            master_rows[profile.name].extend(rows)
            audit_log.add(entry)

    Path(output_dir).mkdir(parents=True, exist_ok=True)
    for profile in profiles:
        df = pd.DataFrame(master_rows[profile.name])
        df = df.reindex(columns=profile.template_columns)
        out_path = Path(output_dir) / f"consolidated_{profile.name}.xlsx"
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Master", index=False)
            if profile.text_columns:
                ws_out = writer.sheets["Master"]
                for col_name in profile.text_columns:
                    if col_name not in df.columns:
                        continue
                    # Force Excel Text format so numeric-looking codes (e.g. "3.10")
                    # don't get silently reinterpreted as numbers (3.1) on re-open.
                    col_idx = df.columns.get_loc(col_name) + 1
                    for row_idx in range(2, len(df) + 2):
                        ws_out.cell(row=row_idx, column=col_idx).number_format = "@"

    audit_log.write(str(Path(output_dir) / "audit_report.xlsx"))
    return audit_log
