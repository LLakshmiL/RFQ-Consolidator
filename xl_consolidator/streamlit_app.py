"""Streamlit UI for xl_consolidator: a pandas/openpyxl reimplementation of a
3-button "Consolidate -> Merge -> Write Back" Excel macro workflow.

Run with:  streamlit run xl_consolidator/streamlit_app.py

The uploaded file is never modified in place - every stage reads from the
originally uploaded bytes and produces new in-memory output, so re-running
Consolidate always starts from a clean slate.
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import streamlit as st

from xl_consolidator.core import (
    CONSOLIDATED_SHEET_NAME,
    META_COLS,
    consolidate_workbook,
    export_consolidated_sheet,
    merge_preview,
    write_back,
)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="Consolidator", page_icon="\U0001f5c3️", layout="wide")

# --------------------------------------------------------------------------- #
# Button styling: big rounded teal buttons stacked in the sidebar, matching
# the "Consolidate / Merge / Write Back" macro buttons this tool mirrors.
# --------------------------------------------------------------------------- #
st.markdown(
    """
    <style>
    section[data-testid="stSidebar"] div.stButton > button {
        background-color: #1b6e7c;
        color: white;
        font-size: 1.35rem;
        font-weight: 600;
        border: none;
        border-radius: 16px;
        padding: 1.5rem 0.5rem;
        width: 100%;
        margin-bottom: 1.1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.15);
    }
    section[data-testid="stSidebar"] div.stButton > button:hover {
        background-color: #14545f;
        color: white;
    }
    section[data-testid="stSidebar"] div.stButton > button:disabled {
        background-color: #9fb9bd;
        color: #eef2f2;
        box-shadow: none;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if "stage" not in st.session_state:
    st.session_state.stage = "idle"  # idle -> consolidated -> merged -> written_back
    st.session_state.result = None
    st.session_state.original_bytes = None
    st.session_state.edited_df = None
    st.session_state.preview = None

st.title("Consolidator")
st.caption(
    "Upload a multi-tab workbook, consolidate every source tab into one master table, "
    "edit values, then push the changes back into the original tabs."
)

with st.sidebar:
    st.header("Actions")
    uploaded = st.file_uploader("Workbook", type=["xlsx"], label_visibility="collapsed")

    consolidate_clicked = st.button("Consolidate", disabled=uploaded is None)
    merge_clicked = st.button("Merge", disabled=st.session_state.stage == "idle")
    write_back_clicked = st.button("Write Back", disabled=st.session_state.stage != "merged")

    st.divider()
    st.markdown(
        "**The routine, every time**\n"
        "1. Save a backup copy — this tool never edits your original file, "
        "but downloads below overwrite each other.\n"
        "2. **Consolidate** — combines every source tab into one master table "
        f'(rows are kept only where "Field description" contains "Offer").\n'
        "3. Type your changes directly in the table. Leave a cell blank to "
        "keep its old value — blanks never erase data.\n"
        "4. **Merge** — previews your changes. Check the **Orphan edits** "
        "table for values typed somewhere with no matching source cell.\n"
        "5. **Write Back** — pushes changes into the source tabs. Updated "
        "cells turn yellow. Formula cells are always skipped.\n\n"
        "**Limits:** rows/columns can't be added or removed in the master "
        "table — doing so blocks Merge until you click Consolidate again."
    )

# --------------------------------------------------------------------------- #
# Stage: Consolidate
# --------------------------------------------------------------------------- #
if consolidate_clicked and uploaded is not None:
    st.session_state.original_bytes = uploaded.getvalue()
    try:
        result = consolidate_workbook(st.session_state.original_bytes)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read workbook: {exc}")
        st.stop()

    st.session_state.result = result
    st.session_state.edited_df = result.consolidated.copy()
    st.session_state.stage = "consolidated"
    st.session_state.preview = None
    st.rerun()

if st.session_state.stage == "idle":
    st.info("Upload a workbook, then click **Consolidate** to build the master table.")
    st.stop()

result = st.session_state.result

st.caption(f"{len(result.consolidated)} rows x {len(result.consolidated.columns) - len(META_COLS)} data columns")

st.subheader(CONSOLIDATED_SHEET_NAME)
edited = st.data_editor(
    st.session_state.edited_df,
    disabled=META_COLS,
    hide_index=True,
    width="stretch",
    num_rows="fixed",  # mirrors the source tool's "no new rows" limitation
    key="master_editor",
)
st.session_state.edited_df = edited

dl_col, _ = st.columns([1, 3])
with dl_col:
    out_bytes = export_consolidated_sheet(st.session_state.original_bytes, result.consolidated)
    st.download_button(
        f'Download workbook with "{CONSOLIDATED_SHEET_NAME}" sheet',
        data=out_bytes,
        file_name="consolidated.xlsx",
        mime=XLSX_MIME,
    )

# --------------------------------------------------------------------------- #
# Stage: Merge
# --------------------------------------------------------------------------- #
if merge_clicked:
    preview = merge_preview(st.session_state.edited_df, result)
    st.session_state.preview = preview
    st.session_state.stage = "merged" if not preview.blocked else "consolidated"
    # Force a fresh top-to-bottom run so the sidebar buttons re-render with
    # the updated `stage` — without this, "Write Back" stays visually
    # disabled for one click because its `disabled=` was already evaluated
    # earlier in this same script pass, before `stage` changed above.
    st.rerun()

preview = st.session_state.preview
if preview is not None:
    st.subheader("Merge preview")
    if preview.blocked:
        st.error(f"BLOCKED — {preview.blocked_reason}")
    else:
        writes, orphans, skipped_formulas = preview.writes, preview.orphans, preview.skipped_formulas
        st.success(f"{len(writes)} cell(s) will be written across {writes['SrcTab'].nunique()} tab(s).")
        if not writes.empty:
            st.dataframe(writes, width="stretch", hide_index=True)

        if not orphans.empty:
            st.warning(
                f"{len(orphans)} orphan edit(s) — data typed somewhere with no matching source cell "
                "(e.g. outside the original table shape). These will NOT be written."
            )
            st.dataframe(orphans, width="stretch", hide_index=True)

        if not skipped_formulas.empty:
            st.info(f"{len(skipped_formulas)} edit(s) target a formula cell and will be skipped.")
            st.dataframe(skipped_formulas, width="stretch", hide_index=True)

# --------------------------------------------------------------------------- #
# Stage: Write Back
# --------------------------------------------------------------------------- #
if write_back_clicked and st.session_state.stage == "merged":
    # Recompute the diff from the CURRENT table state rather than reusing the
    # cached Merge-time preview - if the user kept editing after clicking
    # Merge, this ensures Write Back reflects what's actually on screen
    # instead of a stale snapshot.
    fresh_preview = merge_preview(st.session_state.edited_df, result)
    out_bytes, audit = write_back(st.session_state.original_bytes, result.provenance, fresh_preview.writes)
    st.session_state.push_result = out_bytes
    st.session_state.push_audit = audit
    st.session_state.stage = "written_back"
    st.rerun()

if st.session_state.get("push_result"):
    st.subheader("Write Back result")
    st.success("Push complete — changed cells are highlighted yellow in the download below.")
    st.download_button(
        "Save synced_workbook.xlsx",
        data=st.session_state.push_result,
        file_name="synced_workbook.xlsx",
        mime=XLSX_MIME,
        type="primary",
    )
    st.dataframe(st.session_state.push_audit, width="stretch", hide_index=True)
