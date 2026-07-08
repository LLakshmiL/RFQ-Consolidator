"""Flask app for the generic interactive consolidation & two-way sync
tool. Local, single-user use only (binds to 127.0.0.1, in-memory session
store) — see session_store.py.

Run with:  python sync_studio/app.py
Then open: http://127.0.0.1:5050
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import openpyxl
from flask import Flask, abort, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from sync_studio.diff_sync import build_sync_plan, to_jsonable
from sync_studio.master import build_master
from sync_studio.parser import parse_tab, parse_workbook
from sync_studio.session_store import create_session, get_session
from sync_studio.writer import apply_sync_plan

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024


def _master_json(master, detections=None):
    payload = {
        "columns": master.columns,
        "rows": [{k: to_jsonable(v) for k, v in row.items()} for row in master.rows],
    }
    if detections is not None:
        payload["detections"] = detections
    return payload


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("workbook")
    if not f or not f.filename:
        return jsonify({"error": "Choose a .xlsx file."}), 400
    if not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": f"'{f.filename}' is not a .xlsx file."}), 400

    original_bytes = f.read()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(original_bytes), data_only=True)
        baseline_tabs = parse_workbook(wb)
    except Exception as exc:  # noqa: BLE001
        return jsonify({"error": f"Could not read workbook: {exc}"}), 400

    master = build_master(baseline_tabs)
    token = create_session(original_bytes, secure_filename(f.filename), baseline_tabs, master)

    detections = [
        {
            "sheet": tab.sheet_name,
            "header_row": tab.header_row,
            "confidence": tab.detection.confidence,
        }
        for tab in baseline_tabs
    ]
    return jsonify({"token": token, **_master_json(master, detections)})


@app.route("/header-override", methods=["POST"])
def header_override():
    body = request.get_json(force=True)
    token = body.get("token")
    sheet_name = body.get("sheet_name")
    header_row = body.get("header_row")

    session = get_session(token)
    if session is None:
        abort(404)

    wb = openpyxl.load_workbook(io.BytesIO(session.original_bytes), data_only=True)
    if sheet_name not in wb.sheetnames:
        return jsonify({"error": f"Unknown sheet '{sheet_name}'."}), 400

    new_tab = parse_tab(wb[sheet_name], header_row=int(header_row))
    session.baseline_tabs = [new_tab if t.sheet_name == sheet_name else t for t in session.baseline_tabs]
    session.master = build_master(session.baseline_tabs)

    detections = [
        {"sheet": t.sheet_name, "header_row": t.header_row, "confidence": t.detection.confidence}
        for t in session.baseline_tabs
    ]
    return jsonify(_master_json(session.master, detections))


@app.route("/download-master", methods=["POST"])
def download_master():
    body = request.get_json(force=True)
    columns = body.get("columns", [])
    rows = body.get("rows", [])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master"
    header = ["Source Tab"] + columns
    ws.append(header)
    for row in rows:
        ws.append([row.get("__source_tab", "")] + [row.get(col) for col in columns])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True, download_name="consolidated_master.xlsx",
    )


@app.route("/push", methods=["POST"])
def push():
    body = request.get_json(force=True)
    token = body.get("token")
    columns = body.get("columns", [])
    rows = body.get("rows", [])

    session = get_session(token)
    if session is None:
        abort(404)

    plan = build_sync_plan(session.baseline_tabs, columns, rows)
    result_bytes = apply_sync_plan(session.original_bytes, session.baseline_tabs, plan)

    buf = io.BytesIO(result_bytes)
    return send_file(
        buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True, download_name="synced_result.xlsx",
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5050, debug=False)
