"""
PRP Automation Dashboard — Flask server
Run: python server.py
"""

from __future__ import annotations

import base64
import io
import math
import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd
from flask import Flask, Response, jsonify, redirect, request, send_file, send_from_directory

APP_DIR = Path(__file__).resolve().parent
INPUT_NAME = "PRP Sample Jun (2).xlsx"
ABS_PREFIX = (
    r"C:\Users\C915662\OneDrive - Anheuser-Busch InBev"
    r"\AB-InBev Automations\Assessment_Automation"
)


@dataclass
class ScriptEntry:
    id: str
    group: str
    label: str
    rel_path: str
    sheet: str
    notes: str = ""
    patch_abs: bool = False

    @property
    def path(self) -> Path:
        return APP_DIR / self.rel_path


REGISTRY: list[ScriptEntry] = [
    ScriptEntry(
        "d1", "Diagram 1 — Suppliers",
        "Zone-wise Tier-1 Suppliers (2026 Technology filter) + chart",
        "daigram 1 automation/automation.py",
        "TPRM Web-Portal Export",
        "Filters 2026 + TECHNOLOGY, merges static Tier-1 data, embeds a matplotlib chart.",
    ),
    ScriptEntry(
        "d2", "Diagram 2 — Assessments",
        "Cyber assessments 2026: Open vs Closed by zone + KPI charts",
        "daigram 2 automation/automation.py",
        "OneTrust Assessment",
        "Tags=Cyber & year 2026, Open/Closed pivot, Q2 KPI, two native Excel charts.",
    ),
    ScriptEntry(
        "d3", "Diagram 3 — Assessments",
        "Beyond-1-year-overdue assessments: Completed vs Open pivot + chart",
        "daigram 3 automation/automation.py",
        "OneTrust Assessment",
        "Has an absolute path baked in — the app rewrites it to run locally.",
        patch_abs=True,
    ),
    ScriptEntry(
        "s4", "Diagram 4 — Risk Dashboard",
        "Cumulative Risk Treatment Progress + History",
        "diagram4/automation.py",
        "OneTrust - Risk Export",
        "Bar chart of open risks vs baseline/target; outputs Risk_Output.xlsx + History.xlsx.",
    ),
    ScriptEntry(
        "s1c", "Slide 12 · 1st — Suppliers",
        "Status Overview + Active-by-Zone (combines ACTIVE/Active)",
        "Slide 12 1st daigram Automation/automation3.py",
        "TPRM Web-Portal Export",
        "Most robust 1st-diagram variant: merges case-variant Active columns.",
    ),
    ScriptEntry(
        "s2a", "Slide 12 · 2nd — Assessments",
        "Open vs Closed by Zone (pivot + stacked chart)",
        "Slide 12 2nd daigram Automation/automation.py",
        "OneTrust Assessment",
        "Appends two summary sheets to a copy of the workbook.",
    ),
    ScriptEntry(
        "s3d", "Slide 12 · 3rd — Risks",
        "Open/Overdue pivot (Aging > 90 days) — most complete risk report",
        "Slide 12 3rd daigram Automation/automation4.py",
        "OneTrust - Risk Export",
        "Most robust 3rd-diagram variant: numeric Aging>90 rule, column auto-detect.",
    ),
]

REGISTRY_BY_ID = {e.id: e for e in REGISTRY}


# ---------------------------------------------------------------------------
# Runner (unchanged business logic)
# ---------------------------------------------------------------------------

@dataclass
class RunResult:
    ok: bool
    returncode: int
    stdout: str
    stderr: str
    outputs: list[tuple[str, bytes]] = field(default_factory=list)
    tables: dict[str, dict[str, pd.DataFrame]] = field(default_factory=dict)


def _snapshot(folder: Path) -> dict[str, float]:
    return {p.name: p.stat().st_mtime for p in folder.iterdir() if p.is_file()}


def run_script(entry: ScriptEntry, uploaded_bytes: bytes) -> RunResult:
    tmpdir = Path(tempfile.mkdtemp(prefix="excel_auto_"))
    try:
        (tmpdir / INPUT_NAME).write_bytes(uploaded_bytes)
        source = entry.path.read_text(encoding="utf-8")
        if entry.patch_abs:
            source = source.replace(ABS_PREFIX, str(tmpdir))
        script_copy = tmpdir / "script_to_run.py"
        script_copy.write_text(source, encoding="utf-8")
        before = _snapshot(tmpdir)
        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"
        env["MPLBACKEND"] = "Agg"
        proc = subprocess.run(
            [sys.executable, script_copy.name],
            cwd=str(tmpdir), env=env,
            capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=180,
        )
        after = _snapshot(tmpdir)
        ignore = {INPUT_NAME, script_copy.name}
        produced = [
            n for n, mt in after.items()
            if n not in ignore and (n not in before or before[n] != mt)
        ]
        result = RunResult(
            ok=proc.returncode == 0,
            returncode=proc.returncode,
            stdout=proc.stdout or "",
            stderr=proc.stderr or "",
        )
        for name in sorted(produced):
            data = (tmpdir / name).read_bytes()
            result.outputs.append((name, data))
            if name.lower().endswith(".xlsx"):
                try:
                    result.tables[name] = pd.read_excel(
                        tmpdir / name, sheet_name=None, engine="openpyxl"
                    )
                except Exception as exc:
                    result.tables[name] = {"(could not read)": pd.DataFrame({"error": [str(exc)]})}
        return result
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Data helpers
# ---------------------------------------------------------------------------

_NAN_STRINGS = {"nan", "none", "nat", ""}


def _trim_sparse_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    threshold = max(2, len(df.columns) // 2)
    def _filled(row):
        return sum(1 for v in row if not pd.isna(v) and str(v).strip().lower() not in _NAN_STRINGS)
    return df[df.apply(_filled, axis=1) >= threshold].reset_index(drop=True)


def _trim_sparse_cols(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    threshold = max(2, len(df) // 2)
    def _filled(col):
        return sum(1 for v in col if not pd.isna(v) and str(v).strip().lower() not in _NAN_STRINGS)
    return df.loc[:, df.apply(_filled, axis=0) >= threshold]


def _clean_for_tsv(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join(str(c) for c in col if str(c) not in ("", "nan")).strip() for col in df.columns]
    else:
        df.columns = [str(c) for c in df.columns]
    df.columns = ["" if c.startswith("Unnamed:") else c for c in df.columns]
    col_names = list(df.columns)
    out_cols = []
    for i in range(len(col_names)):
        series = df.iloc[:, i]
        if pd.api.types.is_numeric_dtype(series):
            def _fmt(v):
                if pd.isna(v): return ""
                if isinstance(v, float) and v.is_integer(): return str(int(v))
                return str(v)
            out_cols.append(series.apply(_fmt).rename(col_names[i]))
        else:
            out_cols.append(
                series.astype(str)
                .str.replace(r"[\t\r\n]+", " ", regex=True).str.strip()
                .apply(lambda x: "" if str(x).lower() in _NAN_STRINGS else x)
                .rename(col_names[i])
            )
    return pd.concat(out_cols, axis=1)


def _chartable(df: pd.DataFrame):
    if df is None or df.empty:
        return None
    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    label_cols   = [c for c in df.columns if c not in numeric_cols]
    if not numeric_cols or not label_cols:
        return None
    label = label_cols[0]
    chart_df = df[[label] + numeric_cols].copy().dropna(subset=[label])
    mask = ~chart_df[label].astype(str).str.strip().str.lower().isin({"grand total", "total"})
    chart_df = chart_df[mask]
    if chart_df.empty:
        return None
    chart_df = chart_df.set_index(label)
    chart_df.columns = [
        (str(c).replace(":", "_").replace(".", "_") or f"col_{i}")
        for i, c in enumerate(chart_df.columns)
    ]
    chart_df.index.name = (
        str(chart_df.index.name).replace(":", "_").replace(".", "_")
        if chart_df.index.name else "label"
    )
    return chart_df


def _safe_val(v: Any) -> Any:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    try:
        import numpy as np  # noqa: PLC0415
        if isinstance(v, np.generic):
            v = v.item()
    except ImportError:
        pass
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return v
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return ""
        if v == int(v) and abs(v) < 1e15:
            return int(v)
        return round(v, 6)
    s = str(v)
    return "" if s.lower() in _NAN_STRINGS else s


def _df_headers(df: pd.DataFrame) -> list[str]:
    if isinstance(df.columns, pd.MultiIndex):
        return [" ".join(str(c) for c in col if str(c) not in ("", "nan")).strip() for col in df.columns]
    return ["" if str(c).startswith("Unnamed:") else str(c) for c in df.columns]


def _df_to_rows(df: pd.DataFrame) -> list[list]:
    return [[_safe_val(v) for v in row] for row in df.values]


def _chart_json(chart_df: pd.DataFrame) -> dict:
    return {
        "type": "bar",
        "labels": chart_df.index.astype(str).tolist(),
        "series": [
            {"name": str(col), "values": pd.to_numeric(chart_df[col], errors="coerce").fillna(0).tolist()}
            for col in chart_df.columns
        ],
    }


def _d2_chart_json(xlsx_bytes: bytes) -> dict | None:
    try:
        raw = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name="Dashboard", header=None, engine="openpyxl")
    except Exception:
        return None
    kpi_row = next((i for i, r in raw.iterrows() if "Metric" in r.values), None)
    if kpi_row is None:
        return None
    kv = raw.iloc[kpi_row:].copy()
    kv.columns = kv.iloc[0]
    kv = kv.iloc[1:].reset_index(drop=True)
    kv = kv[kv["Metric"].notna() & (kv["Metric"].astype(str).str.strip() != "")].copy()
    kv["Value"] = pd.to_numeric(kv["Value"], errors="coerce")
    kv = kv.dropna(subset=["Value"])
    kv = kv[~kv["Metric"].astype(str).str.contains("formula", case=False, na=False)]
    if kv.empty:
        return None
    return {
        "type": "hbar",
        "labels": kv["Metric"].astype(str).tolist(),
        "series": [{"name": "Value", "values": kv["Value"].tolist()}],
    }


def _mime(name: str) -> str:
    if name.lower().endswith(".xlsx"):
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if name.lower().endswith(".png"):
        return "image/png"
    return "application/octet-stream"


def _process_results(entry: ScriptEntry, result: RunResult) -> dict:
    processed_sheets: list[tuple[str, str, pd.DataFrame]] = []
    for fname, sheets in result.tables.items():
        for sheet_name, df in sheets.items():
            if entry.id == "s4":
                if fname.lower() == "history.xlsx":
                    continue
                if sheet_name == "History Used":
                    continue
            if entry.id == "s3d" and sheet_name == "Open Overdue Pivot":
                continue
            if entry.id == "s1c" and sheet_name == "Pivot":
                continue
            if entry.id == "s2a" and sheet_name == "Auto Pivot Summary":
                continue
            processed_sheets.append((fname, sheet_name, _trim_sparse_cols(_trim_sparse_rows(df))))

    filtered_outputs = [
        (n, d) for n, d in result.outputs
        if not (entry.id in ("d1", "d2") and n.lower().endswith(".png"))
    ]

    d2_chart = None
    if entry.id == "d2":
        xlsx_bytes = next((d for n, d in result.outputs if n.lower().endswith(".xlsx")), None)
        if xlsx_bytes:
            d2_chart = _d2_chart_json(xlsx_bytes)

    s4_chart = None
    if entry.id == "s4":
        for _fn, _sheets in result.tables.items():
            if _fn.lower() == "history.xlsx":
                for _sn, _df in _sheets.items():
                    _cdf = _chartable(_trim_sparse_cols(_trim_sparse_rows(_df)))
                    if _cdf is not None:
                        s4_chart = _chart_json(_cdf)
                        break

    charts_done = 0
    sheets_out = []
    for fname, sheet_name, df in processed_sheets:
        show_chart = True
        if entry.id == "d2":
            show_chart = charts_done == 0
        elif entry.id == "s4":
            show_chart = charts_done == 0

        chart_data = None
        if show_chart:
            if entry.id == "d2":
                if d2_chart:
                    chart_data = d2_chart
                    charts_done += 1
            elif entry.id == "s4":
                if s4_chart:
                    chart_data = s4_chart
                    charts_done += 1
            else:
                chart_df = _chartable(df)
                if chart_df is not None:
                    chart_data = _chart_json(chart_df)
                    charts_done += 1

        sheets_out.append({
            "file": fname,
            "name": sheet_name,
            "rows": len(df),
            "cols": len(df.columns),
            "headers": _df_headers(df),
            "preview": _df_to_rows(df.head(50)),
            "chart": chart_data,
        })

    copy_tsv = ""
    if processed_sheets:
        _, _, df0 = processed_sheets[0]
        copy_tsv = _clean_for_tsv(df0).head(5).to_csv(sep="\t", index=False)

    return {
        "ok": result.ok,
        "returncode": result.returncode,
        "stdout": result.stdout,
        "stderr": result.stderr,
        "files": [
            {"name": n, "data_b64": base64.b64encode(d).decode(), "mime": _mime(n)}
            for n, d in filtered_outputs
        ],
        "sheets": sheets_out,
        "copy_tsv": copy_tsv,
    }


# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------

_PLOTLY_URL = "https://cdn.plot.ly/plotly-2.35.2.min.js"
_PLOTLY_LOCAL = APP_DIR / "static" / "plotly.min.js"


def _ensure_plotly() -> None:
    """Download Plotly.js once to static/ so the browser never hits an external CDN."""
    if _PLOTLY_LOCAL.exists():
        return
    import urllib.request
    _PLOTLY_LOCAL.parent.mkdir(exist_ok=True)
    try:
        print("Downloading Plotly.js (one-time) ...", end=" ", flush=True)
        with urllib.request.urlopen(_PLOTLY_URL, timeout=15) as r:
            _PLOTLY_LOCAL.write_bytes(r.read())
        print("done.")
    except Exception as exc:
        print(f"warning: {exc}")
        print("  -> Charts will load from CDN at runtime.")


app = Flask(__name__)


@app.get("/")
def index():
    html = (APP_DIR / "index.html").read_text(encoding="utf-8")
    logo_path = APP_DIR / "logo.png"
    if logo_path.exists():
        logo_b64 = base64.b64encode(logo_path.read_bytes()).decode()
        html = html.replace(
            'src="/logo.png"',
            f'src="data:image/png;base64,{logo_b64}"',
            1,
        )
    return Response(html, mimetype="text/html")


@app.get("/static/plotly.min.js")
def serve_plotly():
    if _PLOTLY_LOCAL.exists():
        return send_file(_PLOTLY_LOCAL, mimetype="application/javascript")
    return redirect(_PLOTLY_URL, 302)


@app.post("/api/run")
def api_run():
    script_id = request.form.get("script_id")
    f = request.files.get("file")
    if not script_id or not f:
        return jsonify({"error": "Missing script_id or file"}), 400
    entry = REGISTRY_BY_ID.get(script_id)
    if not entry:
        return jsonify({"error": f"Unknown script: {script_id}"}), 400
    try:
        result = run_script(entry, f.read())
    except subprocess.TimeoutExpired:
        return jsonify({"error": "Script exceeded 3-minute timeout."}), 504
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    return jsonify(_process_results(entry, result))


if __name__ == "__main__":
    import threading
    import webbrowser
    _ensure_plotly()
    port = 5050
    url = f"http://localhost:{port}"
    print(f"\n  PRP Dashboard  ->  {url}\n  Ctrl+C to stop\n")
    threading.Timer(0.9, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=port, debug=False)
