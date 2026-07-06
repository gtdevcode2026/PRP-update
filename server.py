"""
PRP Dashboard — local Flask server.
Serves index.html as the frontend and runs the Python automation scripts
server-side, returning the generated XLSX + embedded chart images.

Run:  python server.py
Then open:  http://localhost:5000
"""

import base64
import shutil
import subprocess
import sys
import tempfile
import traceback
import zipfile
from pathlib import Path

from flask import Flask, jsonify, request, send_file

BASE = Path(__file__).parent

app = Flask(__name__)

# Map report ID → (subdirectory, script filename, expected output filename)
REPORTS = {
    "d1":  ("daigram 1 automation",             "automation.py",  "PRP_Output.xlsx"),
    "d2":  ("daigram 2 automation",             "automation.py",  "output file D2.xlsx"),
    "d3":  ("daigram 3 automation",             "automation.py",  "Risk_Output.xlsx"),
    "s1c": ("Slide 12 1st daigram Automation",  "automation3.py", "PRP_Final_Output3.xlsx"),
    "s2a": ("Slide 12 2nd daigram Automation",  "automation.py",  "OneTrust_Report.xlsx"),
    "s3d": ("Slide 12 3rd daigram Automation",  "automation4.py", "Risk_Output.xlsx"),
    "s4":  ("diagram4",                         "automation.py",  "Risk_Output.xlsx"),
}

# Reports that read/write a History.xlsx file (needs to persist between runs)
HISTORY_REPORTS = {"d3", "s4"}


def extract_xlsx_images(xlsx_path: Path) -> list:
    """Return list of base64-encoded PNG strings from the XLSX xl/media/ folder."""
    images = []
    try:
        with zipfile.ZipFile(xlsx_path, "r") as zf:
            media = sorted(
                n for n in zf.namelist()
                if n.startswith("xl/media/") and n.lower().endswith(".png")
            )
            for name in media:
                images.append(base64.b64encode(zf.read(name)).decode())
    except Exception:
        pass
    return images


@app.route("/api/health")
def health():
    return jsonify(ok=True)


@app.route("/api/run", methods=["POST"])
def run_report():
    report_id = request.form.get("reportId", "").strip()
    uploaded = request.files.get("file")

    if report_id not in REPORTS:
        return jsonify(error=f"Unknown report: {report_id!r}"), 400
    if not uploaded:
        return jsonify(error="No file uploaded"), 400

    script_dir, script_file, output_name = REPORTS[report_id]
    tmpdir = Path(tempfile.mkdtemp(prefix="prp_"))

    try:
        # Save the uploaded workbook with the expected filename
        input_path = tmpdir / "PRP Sample Jun (2).xlsx"
        uploaded.save(str(input_path))

        # For reports that use History.xlsx, seed the temp dir with the project copy
        if report_id in HISTORY_REPORTS:
            hist_src = BASE / "History.xlsx"
            if hist_src.exists():
                shutil.copy(str(hist_src), str(tmpdir / "History.xlsx"))

        script_path = BASE / script_dir / script_file
        proc = subprocess.run(
            [sys.executable, str(script_path)],
            cwd=str(tmpdir),
            capture_output=True,
            text=True,
            timeout=120,
        )

        stdout = (proc.stdout or "").strip()
        stderr = (proc.stderr or "").strip()

        if proc.returncode != 0:
            return jsonify(
                error=stderr or stdout or "Script exited with non-zero status",
                stdout=stdout,
                stderr=stderr,
            ), 500

        # Locate the output XLSX
        output_path = tmpdir / output_name
        if not output_path.exists():
            xlsx_files = sorted(tmpdir.glob("*.xlsx"))
            if not xlsx_files:
                return jsonify(
                    error="Output file not found after script ran",
                    stdout=stdout,
                    stderr=stderr,
                ), 500
            output_path = xlsx_files[0]

        xlsx_b64 = base64.b64encode(output_path.read_bytes()).decode()

        # Collect chart images: first from XLSX embedded media, then standalone PNGs
        charts = extract_xlsx_images(output_path)
        for png in sorted(tmpdir.glob("*.png")):
            charts.append(base64.b64encode(png.read_bytes()).decode())

        # Persist updated History.xlsx back to the project root
        if report_id in HISTORY_REPORTS:
            hist_out = tmpdir / "History.xlsx"
            if hist_out.exists():
                shutil.copy(str(hist_out), str(BASE / "History.xlsx"))

        return jsonify(
            ok=True,
            xlsx_b64=xlsx_b64,
            filename=output_path.name,
            charts=charts,
            stdout=stdout,
            stderr=stderr,
        )

    except subprocess.TimeoutExpired:
        return jsonify(error="Script timed out after 120 seconds"), 500
    except Exception as exc:
        return jsonify(error=str(exc), trace=traceback.format_exc()), 500
    finally:
        shutil.rmtree(str(tmpdir), ignore_errors=True)


# ---------------------------------------------------------------------------
# Static file serving
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return send_file(str(BASE / "index.html"))


@app.route("/<path:path>")
def static_files(path):
    fp = BASE / path
    if fp.exists() and fp.is_file():
        return send_file(str(fp))
    return "Not found", 404


if __name__ == "__main__":
    print("PRP Dashboard server starting at http://localhost:5000")
    print("Open that URL in Chrome or Edge.")
    app.run(host="localhost", port=5000, debug=False, threaded=True)
