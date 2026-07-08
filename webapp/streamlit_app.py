"""Streamlit UI for the consolidation engine: upload workbook(s), click
Consolidate, get back a zip containing the consolidated master sheet(s)
and the audit report.

Run with:  streamlit run webapp/streamlit_app.py
This starts Streamlit's own local dev server (default: http://localhost:8501)
and opens it in your browser — nothing further to deploy or configure.
"""
from __future__ import annotations

import io
import sys
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from consolidation.engine.consolidate import consolidate

st.set_page_config(page_title="Ratecard Consolidator", page_icon="📊")

st.title("Ratecard Consolidator")
st.caption(
    "Upload one or more workbooks. Each sheet is matched to a known format "
    "and folded into that format's consolidated master sheet."
)

uploaded_files = st.file_uploader("Workbooks", type=["xlsx"], accept_multiple_files=True)

if "result" not in st.session_state:
    st.session_state.result = None

if st.button("Consolidate", type="primary", disabled=not uploaded_files):
    with tempfile.TemporaryDirectory() as in_dir, tempfile.TemporaryDirectory() as out_dir:
        input_paths = []
        for f in uploaded_files:
            path = Path(in_dir) / f.name
            path.write_bytes(f.getvalue())
            input_paths.append(str(path))

        try:
            with st.spinner("Consolidating..."):
                audit_log = consolidate(input_paths, out_dir)
        except Exception as exc:  # noqa: BLE001 - surface any parse failure to the user
            st.error(f"Could not process upload: {exc}")
            st.stop()

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in sorted(Path(out_dir).glob("*.xlsx")):
                zf.write(path, arcname=path.name)

    summary_df = pd.DataFrame(
        [
            {
                "source_file": e.source_file,
                "source_sheet": e.source_sheet,
                "profile": e.profile or "-",
                "status": e.status,
                "rows_in": e.rows_in,
                "rows_out": e.rows_out,
            }
            for e in audit_log.entries
        ]
    )
    processed = sum(1 for e in audit_log.entries if e.status == "processed")

    # Stashed in session_state because Streamlit reruns this whole script on
    # every widget interaction — including the click on the download button
    # below — so the result must survive that rerun instead of only existing
    # inside this `if st.button(...)` block.
    st.session_state.result = {
        "zip_bytes": buf.getvalue(),
        "summary_df": summary_df,
        "processed": processed,
        "total": len(audit_log.entries),
    }

if st.session_state.result:
    r = st.session_state.result
    st.success(f"Done — {r['processed']}/{r['total']} sheet(s) processed")
    st.download_button(
        "Download consolidation_result.zip",
        data=r["zip_bytes"],
        file_name="consolidation_result.zip",
        mime="application/zip",
        type="primary",
    )
    st.subheader("Audit")
    st.dataframe(r["summary_df"], width="stretch", hide_index=True)
