"""Streamlit UI for Sync Studio. Reuses the exact same backend as app.py
(header_detect / parser / master / diff_sync / writer) — only the web
layer differs, so consolidation and two-way-sync behavior is identical.

Run with:  streamlit run sync_studio/streamlit_app.py
This starts Streamlit's own local dev server and opens it in your
browser — nothing further to deploy or configure. Bound to localhost via
the repo's .streamlit/config.toml.
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import openpyxl
import pandas as pd
import streamlit as st

from sync_studio.diff_sync import build_sync_plan, to_jsonable
from sync_studio.master import MasterResult, build_master
from sync_studio.parser import parse_tab, parse_workbook
from sync_studio.writer import apply_sync_plan

SOURCE_COL = "Source Tab"
ROW_KEY_COL = "__source_row"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="Sync Studio", page_icon="\U0001f504", layout="wide")
st.title("Sync Studio")
st.caption(
    "Upload a workbook, consolidate every tab by matching header text, edit the result, "
    "then push changes back into the original tabs (changed cells highlighted yellow)."
)


def _master_to_df(master: MasterResult) -> pd.DataFrame:
    records = []
    for row in master.rows:
        record = {SOURCE_COL: row["__source_tab"], ROW_KEY_COL: row["__source_row"]}
        for col in master.columns:
            record[col] = to_jsonable(row.get(col))
        records.append(record)
    return pd.DataFrame(records, columns=[SOURCE_COL, ROW_KEY_COL] + master.columns)


def _df_to_rows(df: pd.DataFrame, columns: list[str]) -> list[dict]:
    rows = []
    for _, r in df.iterrows():
        row = {"__source_tab": r[SOURCE_COL], "__source_row": int(r[ROW_KEY_COL])}
        for col in columns:
            v = r[col]
            row[col] = None if pd.isna(v) else v
        rows.append(row)
    return rows


def _reset_from_master(master: MasterResult, baseline_tabs) -> None:
    st.session_state.baseline_tabs = baseline_tabs
    st.session_state.columns = list(master.columns)
    st.session_state.master_df = _master_to_df(master)
    st.session_state.detections = [
        {"sheet": t.sheet_name, "header_row": t.header_row, "confidence": t.detection.confidence}
        for t in baseline_tabs
    ]
    st.session_state.pop("push_result", None)


if "baseline_tabs" not in st.session_state:
    st.session_state.baseline_tabs = None
    st.session_state.original_bytes = None
    st.session_state.master_df = None
    st.session_state.columns = []
    st.session_state.detections = []

uploaded = st.file_uploader("Workbook", type=["xlsx"])

if uploaded is not None and st.button("Upload & Consolidate", type="primary"):
    original_bytes = uploaded.getvalue()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(original_bytes), data_only=True)
        baseline_tabs = parse_workbook(wb)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read workbook: {exc}")
        st.stop()

    st.session_state.original_bytes = original_bytes
    _reset_from_master(build_master(baseline_tabs), baseline_tabs)

if st.session_state.baseline_tabs is not None:
    low_conf = [d for d in st.session_state.detections if d["confidence"] == "low"]
    if low_conf:
        with st.expander(f"{len(low_conf)} sheet(s) had low-confidence header detection", expanded=True):
            for d in low_conf:
                c1, c2, c3 = st.columns([3, 2, 2])
                c1.write(d["sheet"])
                new_row = c2.number_input(
                    "Header row", min_value=1, value=d["header_row"], key=f"override_row_{d['sheet']}"
                )
                if c3.button("Use this row", key=f"apply_override_{d['sheet']}"):
                    wb = openpyxl.load_workbook(io.BytesIO(st.session_state.original_bytes), data_only=True)
                    new_tab = parse_tab(wb[d["sheet"]], header_row=int(new_row))
                    updated_tabs = [
                        new_tab if t.sheet_name == d["sheet"] else t for t in st.session_state.baseline_tabs
                    ]
                    _reset_from_master(build_master(updated_tabs), updated_tabs)
                    st.rerun()

    st.subheader("Master")

    add_col, del_col, _spacer = st.columns([2, 2, 3])
    with add_col:
        new_col_name = st.text_input("New column name", key="new_col_name", label_visibility="collapsed", placeholder="New column name")
        if st.button("+ Add Column"):
            name = (new_col_name or "").strip()
            if not name:
                st.warning("Enter a column name first.")
            elif name.lower() in [c.lower() for c in st.session_state.columns]:
                st.warning("A column with that name already exists.")
            else:
                st.session_state.columns.append(name)
                st.session_state.master_df[name] = ""
                st.rerun()
    with del_col:
        target_col = st.selectbox("Delete column", options=["-"] + st.session_state.columns, label_visibility="collapsed")
        if st.button("Delete Column", disabled=(target_col == "-")):
            st.session_state.columns.remove(target_col)
            st.session_state.master_df = st.session_state.master_df.drop(columns=[target_col])
            st.rerun()

    edited_df = st.data_editor(
        st.session_state.master_df,
        column_order=[SOURCE_COL] + st.session_state.columns,
        disabled=[SOURCE_COL],
        hide_index=True,
        width="stretch",
        num_rows="fixed",
        key="master_editor",
    )
    st.session_state.master_df = edited_df

    dl_col, push_col = st.columns(2)
    with dl_col:
        if st.button("Download Master"):
            wb_out = openpyxl.Workbook()
            ws = wb_out.active
            ws.title = "Master"
            ws.append([SOURCE_COL] + st.session_state.columns)
            for _, r in edited_df.iterrows():
                ws.append([r[SOURCE_COL]] + [None if pd.isna(r[c]) else r[c] for c in st.session_state.columns])
            buf = io.BytesIO()
            wb_out.save(buf)
            st.session_state.master_download = buf.getvalue()

        if st.session_state.get("master_download"):
            st.download_button(
                "Save consolidated_master.xlsx", data=st.session_state.master_download,
                file_name="consolidated_master.xlsx", mime=XLSX_MIME,
            )

    with push_col:
        if st.button("Push Changes back to Tabs", type="primary"):
            rows = _df_to_rows(edited_df, st.session_state.columns)
            plan = build_sync_plan(st.session_state.baseline_tabs, st.session_state.columns, rows)
            result_bytes = apply_sync_plan(st.session_state.original_bytes, st.session_state.baseline_tabs, plan)
            st.session_state.push_result = result_bytes
            st.session_state.push_summary = pd.DataFrame(
                [
                    {
                        "Sheet": tp.sheet_name,
                        "Cells Modified": len(tp.modified_cells),
                        "Columns Added": ", ".join(c.column_name for c in tp.new_columns),
                        "Columns Removed": ", ".join(tp.deleted_columns),
                    }
                    for tp in plan.tabs
                ]
            )

        if st.session_state.get("push_result"):
            st.success("Push complete - changed cells are highlighted yellow in the download below.")
            st.download_button(
                "Save synced_result.xlsx", data=st.session_state.push_result,
                file_name="synced_result.xlsx", mime=XLSX_MIME, type="primary",
            )
            st.dataframe(st.session_state.push_summary, width="stretch", hide_index=True)
