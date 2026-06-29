"""
Excel Automation Studio — AB InBev
===================================
Professional Streamlit dashboard for the 12 automation scripts in this folder.
Black-and-gold design system with custom CSS injection.
"""

from __future__ import annotations

import os
import sys
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

APP_DIR = Path(__file__).resolve().parent

# The filename every relative-path script expects to read.
INPUT_NAME = "PRP Sample Jun (2).xlsx"

# The one script that hardcodes an absolute path; we rewrite this prefix to the
# temp dir so it reads/writes locally and never touches the real location.
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


# All 11 scripts, grouped by folder, with friendly labels.
REGISTRY: list[ScriptEntry] = [
    ScriptEntry(
        "d1", "Daigram 1 — Suppliers",
        "Zone-wise Tier-1 Suppliers (2026 Technology filter) + chart",
        "daigram 1 automation/automation.py",
        "TPRM Web-Portal Export",
        "Filters 2026 + TECHNOLOGY, merges static Tier-1 data, embeds a matplotlib chart.",
    ),
    ScriptEntry(
        "d2", "Daigram 2 — Assessments",
        "Cyber assessments 2026: Open vs Closed by zone + KPI charts",
        "daigram 2 automation/automation.py",
        "OneTrust Assessment",
        "Tags=Cyber & year 2026, Open/Closed pivot, Q2 KPI, two native Excel charts.",
    ),
    ScriptEntry(
        "d3", "Daigram 3 — Assessments",
        "Beyond-1-year-overdue assessments: Completed vs Open pivot + chart",
        "daigram 3 automation/automation.py",
        "OneTrust Assessment",
        "Has an absolute path baked in — the app rewrites it to run locally.",
        patch_abs=True,
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
        "s2b", "Slide 12 · 2nd — Assessments",
        "Open/Closed + Overdue (90d) — standalone report, 3 sheets",
        "Slide 12 2nd daigram Automation/automation2.py",
        "OneTrust Assessment",
        "Most complete 2nd-diagram variant: adds an overdue-by-zone report.",
    ),
    ScriptEntry(
        "s3d", "Slide 12 · 3rd — Risks",
        "Open/Overdue pivot (Aging > 90 days) — most complete risk report",
        "Slide 12 3rd daigram Automation/automation4.py",
        "OneTrust - Risk Export",
        "Most robust 3rd-diagram variant: numeric Aging>90 rule, column auto-detect.",
    ),
    ScriptEntry(
        "s4", "Diagram 4 — Risk Dashboard",
        "Cumulative Risk Treatment Progress + History",
        "diagram4/automation.py",
        "OneTrust - Risk Export",
        "Bar chart of open risks vs baseline/target; outputs Risk_Output.xlsx + History.xlsx.",
    ),
]

REGISTRY_BY_ID = {e.id: e for e in REGISTRY}


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------

@dataclass
class RunResult:
    ok: bool
    returncode: int
    stdout: str
    stderr: str
    outputs: list[tuple[str, bytes]] = field(default_factory=list)   # (filename, bytes)
    tables: dict[str, dict[str, pd.DataFrame]] = field(default_factory=dict)  # file -> {sheet: df}


def _snapshot(folder: Path) -> dict[str, float]:
    snap = {}
    for p in folder.iterdir():
        if p.is_file():
            snap[p.name] = p.stat().st_mtime
    return snap


def run_script(entry: ScriptEntry, uploaded_bytes: bytes) -> RunResult:
    """Run one automation script in an isolated temp dir and collect its outputs."""
    tmpdir = Path(tempfile.mkdtemp(prefix="excel_auto_"))
    try:
        # 1. Place the uploaded workbook under the name the scripts expect.
        (tmpdir / INPUT_NAME).write_bytes(uploaded_bytes)

        # 2. Copy (and possibly patch) the script into the temp dir.
        source = entry.path.read_text(encoding="utf-8")
        if entry.patch_abs:
            source = source.replace(ABS_PREFIX, str(tmpdir))
        script_copy = tmpdir / "script_to_run.py"
        script_copy.write_text(source, encoding="utf-8")

        # 3. Snapshot existing files so we can detect what the script creates.
        before = _snapshot(tmpdir)

        # 4. Run it. Force UTF-8 (scripts print emoji that crash cp1252 on Windows)
        #    and a headless matplotlib backend.
        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"
        env["MPLBACKEND"] = "Agg"
        proc = subprocess.run(
            [sys.executable, script_copy.name],
            cwd=str(tmpdir),
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
        )

        # 5. Diff to find new/modified output files (ignore the input + the copy).
        after = _snapshot(tmpdir)
        ignore = {INPUT_NAME, script_copy.name}
        produced = [
            name for name, mtime in after.items()
            if name not in ignore and (name not in before or before[name] != mtime)
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
                except Exception as exc:  # preview is best-effort
                    result.tables[name] = {
                        "(could not read)": pd.DataFrame({"error": [str(exc)]})
                    }
        return result
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Preview helpers
# ---------------------------------------------------------------------------

_NAN_STRINGS = {"nan", "none", "nat", ""}


def _trim_sparse_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    threshold = max(2, len(df.columns) // 2)

    def _filled(row):
        return sum(
            1 for v in row
            if not pd.isna(v) and str(v).strip().lower() not in _NAN_STRINGS
        )

    mask = df.apply(_filled, axis=1) >= threshold
    return df[mask].reset_index(drop=True)


def _trim_sparse_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Drop columns that are mostly empty — removes side-by-side tables that pandas
    reads as extra columns when a sheet has multiple logical tables placed next to
    each other (e.g. a pivot at col 0 and a summary table starting at col 8)."""
    if df.empty:
        return df
    threshold = max(2, len(df) // 2)

    def _filled(col):
        return sum(
            1 for v in col
            if not pd.isna(v) and str(v).strip().lower() not in _NAN_STRINGS
        )

    mask = df.apply(_filled, axis=0) >= threshold
    return df.loc[:, mask]


def _clean_for_tsv(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # 1. Flatten MultiIndex column headers
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            " ".join(str(c) for c in col if str(c) not in ("", "nan")).strip()
            for col in df.columns
        ]
    else:
        df.columns = [str(c) for c in df.columns]

    # 2. Clear pandas "Unnamed: N" artefacts from merged header cells
    df.columns = [
        "" if c.startswith("Unnamed:") else c
        for c in df.columns
    ]

    # 3. Clean every column
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            # Whole-number floats (e.g. 536.0) → "536"; NaN → ""
            def _fmt(v):
                if pd.isna(v):
                    return ""
                if isinstance(v, float) and v.is_integer():
                    return str(int(v))
                return str(v)
            df[col] = df[col].apply(_fmt)
        else:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(r"[\t\r\n]+", " ", regex=True)
                .str.strip()
                .apply(lambda x: "" if x.lower() in _NAN_STRINGS else x)
            )

    return df


def _chartable(df: pd.DataFrame):
    if df is None or df.empty:
        return None
    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    label_cols = [c for c in df.columns if c not in numeric_cols]
    if not numeric_cols or not label_cols:
        return None
    label = label_cols[0]
    chart_df = df[[label] + numeric_cols].copy()
    chart_df = chart_df.dropna(subset=[label])
    mask = ~chart_df[label].astype(str).str.strip().str.lower().isin({"grand total", "total"})
    chart_df = chart_df[mask]
    if chart_df.empty:
        return None
    chart_df = chart_df.set_index(label)
    # Sanitize column names — altair v6 parse_shorthand crashes on empty strings
    # or names containing ':' which it interprets as type qualifiers.
    chart_df.columns = [
        (str(c).replace(":", "_").replace(".", "_") or f"col_{i}")
        for i, c in enumerate(chart_df.columns)
    ]
    chart_df.index.name = (
        str(chart_df.index.name).replace(":", "_").replace(".", "_")
        if chart_df.index.name
        else "label"
    )
    return chart_df


# ---------------------------------------------------------------------------
# Design system — Training Status(4) exact design
# ---------------------------------------------------------------------------

_CSS = """<style>
:root {
  --bg:       #000000;
  --bg2:      #0d0d0d;
  --bg3:      #1a1a1a;
  --bg4:      #242424;
  --border:   #3a3000;
  --border2:  #5a4a00;
  --text:     #d4a800;
  --text2:    #b08800;
  --text3:    #7a6200;
  --accent:   #d4a800;
  --accent2:  #f0c000;
  --green:    #4ab840;
  --red:      #d44040;
  --orange:   #d47800;
  --card-glow: 0 0 40px rgba(212,168,0,0.10);
}
* { box-sizing: border-box; }
body, .stApp { font-family: Arial, sans-serif !important; background: var(--bg) !important; color: var(--text) !important; }
.stApp > header, [data-testid="stHeader"] { background: var(--bg) !important; display: none !important; }
#MainMenu, footer, [data-testid="stDecoration"], [data-testid="stToolbar"] { display: none !important; }
.main .block-container { padding-top: 0 !important; max-width: 1600px !important; padding-left: 0 !important; padding-right: 0 !important; }

/* ── Headings ── */
h1,h2,h3,h4 { color: var(--text) !important; font-family: Arial, sans-serif !important; }

/* ── KPI cards ── */
.kpi-grid { display:grid; grid-template-columns:repeat(6,1fr); gap:14px; margin-bottom:24px; }
@media(max-width:1200px){.kpi-grid{grid-template-columns:repeat(3,1fr);}}
@media(max-width:700px){.kpi-grid{grid-template-columns:repeat(2,1fr);}}
.kpi-card {
  background:var(--bg2); border:1px solid var(--border); border-radius:14px;
  padding:18px 20px; position:relative; overflow:hidden;
  transition:border-color .2s,box-shadow .2s; box-shadow:var(--card-glow);
}
.kpi-card:hover { border-color:var(--border2); box-shadow:0 0 30px rgba(212,168,0,0.18); }
.kpi-card::before {
  content:''; position:absolute; top:0; left:0; right:0; height:2px;
  background:linear-gradient(90deg,var(--accent),var(--accent2)); opacity:0; transition:opacity .2s;
}
.kpi-card:hover::before { opacity:1; }
.kpi-label { font-size:11px; color:#fff; text-transform:uppercase; letter-spacing:0.8px; font-weight:500; margin-bottom:8px; }
.kpi-val { font-family:Arial,sans-serif; font-size:28px; font-weight:800; line-height:1; letter-spacing:-1px; color:var(--accent); }
.kpi-sub { font-size:11px; color:#fff; margin-top:4px; }

/* ── Tabs (Streamlit native override) ── */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
  background: transparent !important; gap: 6px !important;
  border-bottom: 1px solid var(--border) !important; padding-bottom: 0 !important;
}
[data-testid="stTabs"] [data-baseweb="tab"] {
  background: var(--bg2) !important; border: 1px solid var(--border) !important;
  border-radius: 8px 8px 0 0 !important; color: var(--text2) !important;
  font-size: 12px !important; font-weight: 500 !important; font-family: Arial,sans-serif !important;
  padding: 7px 16px !important; white-space: nowrap !important;
  transition: all .15s !important;
}
[data-testid="stTabs"] [data-baseweb="tab"]:hover {
  border-color: var(--accent) !important; color: var(--accent) !important;
}
[data-testid="stTabs"] [aria-selected="true"] {
  background: var(--accent) !important; border-color: var(--accent) !important;
  color: #000 !important; font-weight: 600 !important;
}
[data-testid="stTabs"] [data-baseweb="tab-highlight"] { display: none !important; }
[data-testid="stTabs"] [data-baseweb="tab-border"] { display: none !important; }

/* ── Section card ── */
.section-card {
  background: var(--bg2); border: 1px solid var(--border);
  border-radius: 16px; overflow: hidden; box-shadow: var(--card-glow);
  margin-bottom: 20px;
}
.section-header {
  display: flex; align-items: center; justify-content: space-between;
  padding: 16px 22px; border-bottom: 1px solid var(--border); flex-wrap: wrap; gap: 10px;
}
.section-title-block { display: flex; align-items: center; gap: 10px; }
.section-dot { width: 10px; height: 10px; border-radius: 50%; background: var(--accent); flex-shrink: 0; }
.section-title { font-family: Arial,sans-serif; font-weight: 700; font-size: 14px; color: var(--text); }
.section-desc { font-size: 11px; color: var(--text3); margin-top: 2px; }
.section-body { padding: 18px 22px; }

/* ── Metrics row ── */
.metrics-row {
  display: grid; grid-template-columns: repeat(4,1fr);
  gap: 1px; background: var(--border); border-bottom: 1px solid var(--border);
}
.metric-box { background: var(--bg2); padding: 13px 20px; display: flex; flex-direction: column; gap: 4px; }
.metric-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.7px; color: #fff; font-weight: 500; }
.metric-val { font-family: Arial,sans-serif; font-size: 20px; font-weight: 700; line-height: 1; color: var(--accent); }

/* ── Upload zone ── */
[data-testid="stFileUploader"] {
  background: var(--bg3) !important;
  border: 2px dashed var(--border2) !important;
  border-radius: 12px !important;
  transition: all .2s !important;
}
[data-testid="stFileUploader"]:hover {
  border-color: var(--accent) !important;
  background: rgba(212,168,0,0.06) !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] div,
[data-testid="stFileUploaderDropzoneInstructions"] span { color: var(--text3) !important; font-size: 12px !important; }
[data-testid="stFileUploader"] label { color: var(--text3) !important; font-size: 11px !important; }
[data-testid="stFileUploader"] small { color: var(--text3) !important; }

/* ── Primary button (Run) ── */
.stButton > button[kind="primary"] {
  background: var(--accent) !important; color: #000 !important;
  border: none !important; border-radius: 8px !important;
  font-family: Arial,sans-serif !important; font-weight: 600 !important;
  font-size: 12px !important; padding: 8px 20px !important;
  transition: background .15s, transform .15s !important;
  letter-spacing: 0.3px !important;
}
.stButton > button[kind="primary"]:hover {
  background: var(--accent2) !important; transform: translateY(-1px) !important;
}
.stButton > button[kind="primary"]:disabled {
  background: var(--bg3) !important; color: var(--text3) !important;
  transform: none !important; opacity: 0.5 !important;
}

/* ── Ghost / secondary button ── */
.stButton > button:not([kind="primary"]) {
  background: var(--bg3) !important; color: var(--text2) !important;
  border: 1px solid var(--border) !important; border-radius: 8px !important;
  font-family: Arial,sans-serif !important; font-weight: 600 !important;
  font-size: 12px !important; transition: all .15s !important;
}
.stButton > button:not([kind="primary"]):hover {
  border-color: var(--accent) !important; color: var(--accent) !important;
}

/* ── Download button ── */
[data-testid="stDownloadButton"] > button {
  background: var(--bg3) !important; color: var(--text2) !important;
  border: 1px solid var(--border) !important; border-radius: 8px !important;
  font-family: Arial,sans-serif !important; font-weight: 600 !important; font-size: 12px !important;
  transition: all .15s !important;
}
[data-testid="stDownloadButton"] > button:hover {
  border-color: var(--accent) !important; color: var(--accent) !important;
}

/* ── Radio ── */
[data-testid="stRadio"] label { color: var(--text3) !important; font-size: 11px !important; text-transform: uppercase !important; letter-spacing: 0.7px !important; font-family: Arial,sans-serif !important; }
[data-testid="stRadio"] [data-testid="stMarkdownContainer"] p { color: var(--text2) !important; font-size: 12px !important; font-family: Arial,sans-serif !important; }

/* ── Expander ── */
[data-testid="stExpander"] { background: var(--bg2) !important; border: 1px solid var(--border) !important; border-radius: 8px !important; }
[data-testid="stExpander"] summary { color: var(--text3) !important; font-size: 12px !important; font-family: Arial,sans-serif !important; }
[data-testid="stExpander"] summary:hover { color: var(--accent) !important; }

/* ── Alerts ── */
[data-testid="stAlert"] { background: var(--bg2) !important; border-radius: 8px !important; border: 1px solid var(--border) !important; }
[data-testid="stAlert"] p { color: var(--text2) !important; font-family: Arial,sans-serif !important; }

/* ── Dataframe ── */
[data-testid="stDataFrame"] { border: 1px solid var(--border) !important; border-radius: 8px !important; overflow: hidden !important; }

/* ── Code ── */
[data-testid="stCode"], .stCodeBlock pre { background: var(--bg3) !important; border: 1px solid var(--border) !important; border-radius: 8px !important; color: var(--text2) !important; }

/* ── Image ── */
[data-testid="stImage"] img { border-radius: 8px !important; border: 1px solid var(--border) !important; }

/* ── Spinner ── */
[data-testid="stSpinner"] > div { border-top-color: var(--accent) !important; }

/* ── Text ── */
.stMarkdown p, .stMarkdown li { color: var(--text2) !important; font-family: Arial,sans-serif !important; font-size: 13px !important; }
.stCaption p { color: var(--text3) !important; font-size: 11px !important; }

/* ── Divider ── */
hr { border-color: var(--border) !important; margin: 1rem 0 !important; }

/* ── Badge pills ── */
.badge { display:inline-flex; align-items:center; border-radius:5px; padding:2px 8px; font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:0.5px; }
.badge-green { background:rgba(74,184,64,.15); color:var(--green); }
.badge-red   { background:rgba(212,64,64,.15);  color:var(--red); }
.badge-yellow{ background:rgba(212,168,0,.15);  color:var(--accent); }
.badge-gray  { background:rgba(122,98,0,.25);   color:var(--text3); }

/* ── Vega chart ── */
[data-testid="stArrowVegaLiteChart"] {
  background: var(--bg2) !important; border: 1px solid var(--border) !important;
  border-radius: 8px !important; padding: 12px !important;
}

/* ── Scrollbar ── */
::-webkit-scrollbar { width:6px; height:6px; }
::-webkit-scrollbar-track { background: var(--bg); }
::-webkit-scrollbar-thumb { background: var(--border2); border-radius:3px; }
::-webkit-scrollbar-thumb:hover { background: var(--text3); }
</style>"""


# ---------------------------------------------------------------------------
# HTML helpers
# ---------------------------------------------------------------------------

def _header_html() -> str:
    return """
<div style="
  position:sticky;top:0;z-index:100;
  background:rgba(0,0,0,0.96);
  backdrop-filter:blur(20px);
  border-bottom:1px solid #3a3000;
  padding:0 32px;
  display:flex;align-items:center;justify-content:space-between;
  height:58px;margin-bottom:24px;
">
  <div style="display:flex;align-items:center;gap:10px;">
    <div style="width:30px;height:30px;background:linear-gradient(135deg,#d4a800,#f0c000);
      border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:14px;">&#128202;</div>
    <span style="font-family:Arial,sans-serif;font-weight:800;font-size:18px;color:#d4a800;letter-spacing:-0.3px;">
      Excel Automation Studio
    </span>
  </div>
  <div style="display:flex;align-items:center;gap:12px;">
    <span style="font-family:Arial,sans-serif;font-size:11px;color:#d4a800;background:#1a1a1a;
      border:1px solid #3a3000;border-radius:6px;padding:4px 10px;font-weight:700;">
      AB InBev &middot; TPRM / Risk Automation
    </span>
    <span style="font-family:Arial,sans-serif;font-size:11px;color:#d4a800;background:#1a1a1a;
      border:1px solid #3a3000;border-radius:6px;padding:4px 10px;font-weight:700;">
      8 Reports
    </span>
  </div>
</div>"""


def _kpi_row_html(result: "RunResult | None" = None) -> str:
    n_outputs = len(result.outputs) if result else 0
    n_tables  = sum(len(s) for s in result.tables.values()) if result else 0
    status_html = (
        '<span class="badge badge-green">&#10003; Completed</span>' if result and result.ok
        else '<span class="badge badge-red">&#10007; Error</span>' if result
        else '<span class="badge badge-gray">&#8212;</span>'
    )
    cards = [
        ("Total Reports",  "8",           "8 automation scripts"),
        ("Script Groups",  "7",           "Daigram 1/2/3 &middot; Slide 12"),
        ("Input Sheets",   "3",           "TPRM &middot; OneTrust &middot; Risk"),
        ("Output Files",   str(n_outputs) if result else "&mdash;", "From last run"),
        ("Tables",         str(n_tables)  if result else "&mdash;", "Sheets in output"),
        ("Run Status",     status_html,   "&nbsp;"),
    ]
    inner = "".join(
        f'<div class="kpi-card">'
        f'<div class="kpi-label">{lbl}</div>'
        f'<div class="kpi-val">{val}</div>'
        f'<div class="kpi-sub">{sub}</div>'
        f'</div>'
        for lbl, val, sub in cards
    )
    return f'<div class="kpi-grid">{inner}</div>'


def _section_label(text: str) -> str:
    return (
        f'<div style="font-size:10px;text-transform:uppercase;letter-spacing:0.7px;'
        f'color:#7a6200;font-weight:600;margin:16px 0 10px;">{text}</div>'
    )


def _info_card_html(sheet: str, notes: str, group: str) -> str:
    return f"""
<div class="section-card" style="margin-top:10px;">
  <div class="section-header">
    <div class="section-title-block">
      <div class="section-dot"></div>
      <div>
        <div class="section-title">{group}</div>
        <div class="section-desc">{notes}</div>
      </div>
    </div>
  </div>
  <div class="metrics-row" style="grid-template-columns:1fr 1fr;">
    <div class="metric-box">
      <div class="metric-label">Required Sheet</div>
      <div style="font-family:Arial,sans-serif;font-size:12px;color:#d4a800;margin-top:4px;">{sheet}</div>
    </div>
    <div class="metric-box">
      <div class="metric-label">Format</div>
      <div style="font-family:Arial,sans-serif;font-size:12px;color:#b08800;margin-top:4px;">.xlsx + chart</div>
    </div>
  </div>
</div>"""


# ---------------------------------------------------------------------------
# Results renderer
# ---------------------------------------------------------------------------

def render_results(entry: ScriptEntry, result: RunResult, fmt: str) -> None:
    # ── Metrics row ────────────────────────────────────────────────────────
    n_tables = sum(len(s) for s in result.tables.values())
    status_badge = (
        '<span class="badge badge-green">&#10003; Completed</span>'
        if result.ok else
        f'<span class="badge badge-red">&#10007; Error &mdash; code {result.returncode}</span>'
    )
    st.markdown(f"""
<div class="section-card">
  <div class="section-header">
    <div class="section-title-block">
      <div class="section-dot" style="background:{'#4ab840' if result.ok else '#d44040'};"></div>
      <div>
        <div class="section-title">{entry.label}</div>
        <div class="section-desc">{entry.group}</div>
      </div>
    </div>
    <div>{status_badge}</div>
  </div>
  <div class="metrics-row">
    <div class="metric-box">
      <div class="metric-label">Script</div>
      <div class="metric-val">{entry.id.upper()}</div>
    </div>
    <div class="metric-box">
      <div class="metric-label">Output Files</div>
      <div class="metric-val">{len(result.outputs)}</div>
    </div>
    <div class="metric-box">
      <div class="metric-label">Tables</div>
      <div class="metric-val">{n_tables}</div>
    </div>
    <div class="metric-box">
      <div class="metric-label">Return Code</div>
      <div class="metric-val" style="color:{'#4ab840' if result.ok else '#d44040'};">{result.returncode}</div>
    </div>
  </div>
</div>""", unsafe_allow_html=True)

    with st.expander("Script log (stdout / stderr)", expanded=not result.ok):
        if result.stdout.strip():
            st.code(result.stdout, language="text")
        if result.stderr.strip():
            st.markdown(_section_label("stderr"), unsafe_allow_html=True)
            st.code(result.stderr, language="text")
        if not result.stdout.strip() and not result.stderr.strip():
            st.markdown('<span style="color:#7a6200;font-size:12px;">No console output</span>', unsafe_allow_html=True)

    if not result.outputs:
        st.warning("The script produced no output file.")
        return

    # ── Downloads ──────────────────────────────────────────────────────────
    st.markdown(_section_label("Downloads"), unsafe_allow_html=True)
    dl_cols = st.columns(min(4, len(result.outputs)))
    for i, (name, data) in enumerate(result.outputs):
        mime = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if name.lower().endswith(".xlsx") else
            "image/png" if name.lower().endswith(".png") else
            "application/octet-stream"
        )
        icon = "&#128202;" if name.endswith(".xlsx") else "&#128444;"
        dl_cols[i % len(dl_cols)].download_button(
            f"{icon}  {name}", data=data, file_name=name, mime=mime,
            key=f"dl_{entry.id}_{i}", use_container_width=True,
        )

    # ── Chart preview ──────────────────────────────────────────────────────
    st.markdown(_section_label("Chart Preview"), unsafe_allow_html=True)
    png_outputs = [(n, d) for n, d in result.outputs if n.lower().endswith(".png")]
    rendered_any = False

    if png_outputs:
        img_cols = st.columns(len(png_outputs))
        for col, (name, data) in zip(img_cols, png_outputs):
            with col:
                st.image(data, caption=name, use_container_width=True)
        rendered_any = True

    for fname, sheets in result.tables.items():
        for sheet_name, df in sheets.items():
            chart_df = _chartable(df)
            if chart_df is not None and len(chart_df.columns) <= 6:
                st.markdown(
                    f'<div style="font-size:11px;color:#7a6200;margin-bottom:6px;">'
                    f'{fname} &rarr; {sheet_name}</div>', unsafe_allow_html=True,
                )
                try:
                    st.bar_chart(chart_df, stack=True, use_container_width=True)
                    rendered_any = True
                except Exception:
                    pass

    if not rendered_any:
        st.markdown(
            '<div style="color:#7a6200;font-size:12px;padding:8px 0;">No chartable data detected.</div>',
            unsafe_allow_html=True,
        )
    elif not png_outputs:
        st.markdown(
            '<div style="color:#7a6200;font-size:11px;margin-top:4px;">'
            'Re-rendered from output data &mdash; downloaded .xlsx has the original styled chart.</div>',
            unsafe_allow_html=True,
        )

    # ── Output tables ──────────────────────────────────────────────────────
    st.markdown(_section_label("Output Tables"), unsafe_allow_html=True)
    for fname, sheets in result.tables.items():
        for sheet_name, df in sheets.items():
            df = _trim_sparse_rows(df)
            df = _trim_sparse_cols(df)
            st.markdown(f"""
<div class="section-card" style="margin-bottom:14px;">
  <div class="section-header">
    <div class="section-title-block">
      <div class="section-dot"></div>
      <div>
        <div class="section-title">{sheet_name}</div>
        <div class="section-desc">{fname} &nbsp;&middot;&nbsp; {len(df)} rows &times; {len(df.columns)} cols</div>
      </div>
    </div>
  </div>
  <div style="padding:14px 18px 4px;">""", unsafe_allow_html=True)
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.markdown("</div>", unsafe_allow_html=True)
            with st.expander(f"Copy '{sheet_name}' as text"):
                safe_key = f"cp_{entry.id}_{fname}_{sheet_name}".replace(" ", "_")
                if fmt == "Markdown":
                    try:
                        text = df.to_markdown(index=False)
                    except Exception:
                        text = df.to_csv(sep="\t", index=False)
                    st.text_area("", text, height=220, key=safe_key + "_md", label_visibility="collapsed")
                else:
                    tsv = _clean_for_tsv(df).to_csv(sep="\t", index=False, lineterminator="\r\n")
                    st.caption("Click inside the box, Ctrl+A to select all, Ctrl+C to copy, then paste into Excel — columns split automatically.")
                    st.text_area("", tsv, height=220, key=safe_key, label_visibility="collapsed")
            st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Main UI
# ---------------------------------------------------------------------------

def main() -> None:
    st.set_page_config(
        page_title="Excel Automation Studio",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    st.markdown(_CSS, unsafe_allow_html=True)
    st.markdown(_header_html(), unsafe_allow_html=True)

    # Resolve last result for dynamic KPI
    last_result: "RunResult | None" = st.session_state.get("last_result")

    # ── KPI row ────────────────────────────────────────────────────────────
    padding = '<div style="padding:0 32px;">'
    st.markdown(padding + _kpi_row_html(last_result) + "</div>", unsafe_allow_html=True)

    # ── Upload + Report selector ───────────────────────────────────────────
    st.markdown(padding, unsafe_allow_html=True)
    col_left, col_right = st.columns([1, 1], gap="large")

    with col_left:
        st.markdown(_section_label("1 &middot; Input Workbook"), unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "Upload workbook",
            type=["xlsx"],
            label_visibility="collapsed",
        )
        if uploaded:
            st.markdown(
                f'<div style="font-size:12px;color:#4ab840;margin-top:6px;font-family:Arial,sans-serif;">'
                f'&#10003; &nbsp;<strong style="color:#d4a800;">{uploaded.name}</strong>'
                f'&nbsp;<span style="color:#7a6200;">({uploaded.size/1024:.1f} KB)</span></div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<div style="font-size:11px;color:#7a6200;margin-top:6px;font-family:Arial,sans-serif;">'
                'Accepts .xlsx &middot; same workbook the scripts expect</div>',
                unsafe_allow_html=True,
            )

    with col_right:
        st.markdown(_section_label("2 &middot; Select Report"), unsafe_allow_html=True)

        groups: dict[str, list[ScriptEntry]] = {}
        for e in REGISTRY:
            groups.setdefault(e.group, []).append(e)
        group_names = list(groups.keys())

        _short = ["Daigram 1", "Daigram 2", "Daigram 3", "Slide 12·1", "Slide 12·2", "Slide 12·3", "Diagram 4"]
        short_of = dict(zip(group_names, _short))

        if "active_group" not in st.session_state:
            st.session_state["active_group"] = group_names[0]

        # Tab row — explicit buttons so active group is always known
        tab_cols = st.columns(len(group_names))
        for col, gname in zip(tab_cols, group_names):
            with col:
                is_active = st.session_state["active_group"] == gname
                if st.button(
                    short_of.get(gname, gname),
                    key=f"grp_{gname}",
                    type="primary" if is_active else "secondary",
                    use_container_width=True,
                ):
                    st.session_state["active_group"] = gname
                    st.rerun()

        # Show only the active group's variants
        active_group = st.session_state["active_group"]
        g_entries = groups[active_group]
        labels = [e.label for e in g_entries]

        selected_label = st.radio(
            "Variant", labels,
            key=f"variant_{active_group}",
            label_visibility="collapsed",
        )
        entry = g_entries[labels.index(selected_label)]
        st.session_state["selected_entry"] = entry
        st.markdown(
            _info_card_html(entry.sheet, entry.notes, entry.group),
            unsafe_allow_html=True,
        )

    st.markdown("</div>", unsafe_allow_html=True)

    # ── Options & Run ──────────────────────────────────────────────────────
    st.markdown(padding, unsafe_allow_html=True)
    st.markdown(_section_label("3 &middot; Options &amp; Run"), unsafe_allow_html=True)
    opt_col, run_col = st.columns([2, 1], gap="large")

    with opt_col:
        fmt_raw = st.radio("Copy format", ["TSV (Excel / Sheets)", "Markdown"], horizontal=True)
        fmt = "Markdown" if "Markdown" in fmt_raw else "TSV"

    with run_col:
        run = st.button(
            "&#9654;  Run Report",
            type="primary",
            disabled=uploaded is None,
            use_container_width=True,
        )
        if uploaded is None:
            st.markdown(
                '<div style="font-size:11px;color:#7a6200;text-align:center;margin-top:4px;'
                'font-family:Arial,sans-serif;">Upload a workbook to enable</div>',
                unsafe_allow_html=True,
            )

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown('<hr style="border-color:#3a3000;margin:16px 32px;">', unsafe_allow_html=True)

    with st.expander("How it works"):
        st.markdown(
            "1. **Upload** your `.xlsx` workbook.\n"
            "2. **Select** a report group tab, then pick a variant.\n"
            "3. **Run** — the script executes in an isolated temp folder; you get the log, "
            "chart preview, output tables with copy, and the exact `.xlsx` the script generates.\n\n"
            "> Each report requires a specific sheet — shown in the info card."
        )

    # ── Execute ────────────────────────────────────────────────────────────
    if run and uploaded is not None:
        with st.spinner(f"Running  ·  {entry.label}"):
            try:
                result = run_script(entry, uploaded.getvalue())
            except subprocess.TimeoutExpired:
                st.error("Script exceeded 3-minute timeout.")
                return
            except Exception as exc:
                st.error(f"Could not launch script: {exc}")
                return
        st.session_state["last_result"]    = result
        st.session_state["last_entry_id"]  = entry.id
        st.session_state["last_fmt"]       = fmt
        st.rerun()

    # ── Render persisted results ───────────────────────────────────────────
    if "last_result" in st.session_state:
        _entry = REGISTRY_BY_ID[st.session_state["last_entry_id"]]
        st.markdown(padding, unsafe_allow_html=True)
        render_results(_entry, st.session_state["last_result"], st.session_state["last_fmt"])
        st.markdown("</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
