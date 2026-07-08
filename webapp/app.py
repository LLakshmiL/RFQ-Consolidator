"""Local web UI for the consolidation engine: upload workbook(s), click
Consolidate, get back a zip containing the consolidated master sheet(s)
and the audit report.

Run with:  python webapp/app.py
Then open: http://127.0.0.1:5000

Binds to 127.0.0.1 only — this is a local single-user tool with no
authentication, not meant to be exposed on a network as-is.
"""
from __future__ import annotations

import io
import sys
import tempfile
import uuid
import zipfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from flask import Flask, abort, render_template, request, send_file
from werkzeug.utils import secure_filename

from consolidation.engine.consolidate import consolidate

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB per request

# token -> zip bytes. In-memory is fine for a local, single-user tool.
_RESULTS: dict[str, bytes] = {}


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/consolidate", methods=["POST"])
def consolidate_route():
    uploads = [f for f in request.files.getlist("workbooks") if f and f.filename]
    if not uploads:
        return render_template("index.html", error="Choose at least one .xlsx file.")

    for f in uploads:
        if not f.filename.lower().endswith(".xlsx"):
            return render_template("index.html", error=f"'{f.filename}' is not a .xlsx file.")

    with tempfile.TemporaryDirectory() as in_dir, tempfile.TemporaryDirectory() as out_dir:
        input_paths = []
        for f in uploads:
            path = Path(in_dir) / secure_filename(f.filename)
            f.save(path)
            input_paths.append(str(path))

        try:
            audit_log = consolidate(input_paths, out_dir)
        except Exception as exc:  # noqa: BLE001 - surface any parse failure to the user, not a 500 page
            return render_template("index.html", error=f"Could not process upload: {exc}")

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in sorted(Path(out_dir).glob("*.xlsx")):
                zf.write(path, arcname=path.name)
        buf.seek(0)

        token = uuid.uuid4().hex
        _RESULTS[token] = buf.getvalue()

    summary = [
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
    processed = sum(1 for e in audit_log.entries if e.status == "processed")

    return render_template(
        "result.html",
        token=token,
        summary=summary,
        processed=processed,
        total=len(audit_log.entries),
    )


@app.route("/download/<token>", methods=["GET"])
def download(token: str):
    data = _RESULTS.get(token)
    if data is None:
        abort(404)
    return send_file(
        io.BytesIO(data),
        mimetype="application/zip",
        as_attachment=True,
        download_name="consolidation_result.zip",
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
