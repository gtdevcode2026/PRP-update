"""
PRP Automation Dashboard — AB InBev
====================================
Professional Streamlit dashboard for the PRP automation scripts.
Premium black-and-gold design system.
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
import streamlit.components.v1 as components

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

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
    outputs: list[tuple[str, bytes]] = field(default_factory=list)
    tables: dict[str, dict[str, pd.DataFrame]] = field(default_factory=dict)


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
            cwd=str(tmpdir),
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
        )

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
                except Exception as exc:
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

    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            " ".join(str(c) for c in col if str(c) not in ("", "nan")).strip()
            for col in df.columns
        ]
    else:
        df.columns = [str(c) for c in df.columns]

    df.columns = [
        "" if c.startswith("Unnamed:") else c
        for c in df.columns
    ]

    col_names = list(df.columns)
    out_cols = []
    for i in range(len(col_names)):
        series = df.iloc[:, i]
        if pd.api.types.is_numeric_dtype(series):
            def _fmt(v):
                if pd.isna(v):
                    return ""
                if isinstance(v, float) and v.is_integer():
                    return str(int(v))
                return str(v)
            out_cols.append(series.apply(_fmt).rename(col_names[i]))
        else:
            out_cols.append(
                series
                .astype(str)
                .str.replace(r"[\t\r\n]+", " ", regex=True)
                .str.strip()
                .apply(lambda x: "" if str(x).lower() in _NAN_STRINGS else x)
                .rename(col_names[i])
            )

    return pd.concat(out_cols, axis=1)


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
# Design system
# ---------------------------------------------------------------------------

_CSS = """<style>
:root {
  --bg:       #070707;
  --s1:       #0d0d0d;
  --s2:       #131313;
  --s3:       #191919;
  --s4:       #212121;
  --gold:     #C8A84B;
  --gold-hi:  #E0BC5A;
  --gold-lo:  #7A6535;
  --gold-xlo: #221B0B;
  --gold-xxlo:#160F05;
  --glow:     rgba(200,168,75,0.08);
  --glow2:    rgba(200,168,75,0.18);
  --tx1:      #F0E4C0;
  --tx2:      #B89A68;
  --tx3:      #6A5A38;
  --green:    #4DBB80;
  --red:      #D45050;
  --font:         Arial, sans-serif;
  --font-display: Arial, sans-serif;
}

* { box-sizing: border-box; }

body, .stApp {
  font-family: var(--font) !important;
  background: var(--bg) !important;
  color: var(--tx1) !important;
}

.stApp::after {
  content: '';
  position: fixed;
  top: -300px; right: -200px;
  width: 700px; height: 700px;
  background: radial-gradient(ellipse, rgba(200,168,75,0.04) 0%, transparent 65%);
  pointer-events: none;
  z-index: 0;
}

/* Hide Streamlit chrome */
.stApp > header,
[data-testid="stHeader"],
[data-testid="stDecoration"],
[data-testid="stToolbar"],
#MainMenu, footer {
  display: none !important;
  height: 0 !important;
  min-height: 0 !important;
  overflow: hidden !important;
}

.main .block-container {
  padding-top: 0 !important;
  margin-top: 0 !important;
  max-width: 1440px !important;
  padding-left: 0 !important;
  padding-right: 0 !important;
}

[data-testid="stAppViewContainer"] {
  padding-top: 0 !important;
  margin-top: 0 !important;
}
[data-testid="stMain"] {
  padding-top: 0 !important;
  margin-top: 0 !important;
}
.stMainBlockContainer {
  padding-top: 0 !important;
}

/* ── Typography ── */
h1,h2,h3,h4 {
  font-family: var(--font-display) !important;
  color: var(--tx1) !important;
  letter-spacing: -0.02em !important;
}
.stMarkdown p, .stMarkdown li {
  font-family: var(--font) !important;
  color: var(--tx2) !important;
  font-size: 13px !important;
  line-height: 1.6 !important;
}
.stCaption p { color: var(--tx3) !important; font-size: 11px !important; font-family: var(--font) !important; }

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: var(--bg); }
::-webkit-scrollbar-thumb { background: var(--gold-xlo); border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: var(--gold-lo); }

/* ═══════════════════════════════════════════════════════
   HEADER
═══════════════════════════════════════════════════════ */
.prp-header {
  position: sticky; top: 0; z-index: 100;
  background: rgba(7,7,7,0.95);
  backdrop-filter: blur(28px);
  -webkit-backdrop-filter: blur(28px);
  border-bottom: 1px solid var(--gold-xlo);
  padding: 0 44px;
  display: flex; align-items: center; justify-content: space-between;
  height: 64px;
  margin-bottom: 0;
}
.prp-logo { display: flex; align-items: center; gap: 14px; }
.prp-abinbev-logo {
  height: 36px;
  width: auto;
  display: flex;
  align-items: center;
  flex-shrink: 0;
}
.prp-abinbev-logo img {
  height: 36px;
  width: auto;
  object-fit: contain;
}
.prp-abinbev-logo svg {
  height: 36px;
  width: auto;
}
.prp-brand-name {
  font-size: 15px; font-weight: 800;
  color: var(--tx1); letter-spacing: -0.01em; line-height: 1.1;
  font-family: var(--font-display);
}
.prp-brand-sub {
  font-size: 10px; color: var(--tx3);
  letter-spacing: 0.1em; text-transform: uppercase;
  margin-top: 3px; font-weight: 600;
  font-family: var(--font);
}
.prp-header-right { display: flex; align-items: center; gap: 8px; }
.prp-badge {
  font-size: 9.5px; font-weight: 700;
  color: var(--gold);
  background: rgba(200,168,75,0.08);
  border: 1px solid rgba(200,168,75,0.16);
  border-radius: 100px;
  padding: 5px 13px;
  letter-spacing: 0.06em;
  text-transform: uppercase;
  font-family: var(--font);
}
.prp-badge-live {
  color: var(--green);
  background: rgba(77,187,128,0.08);
  border-color: rgba(77,187,128,0.16);
}
.prp-badge-live::before {
  content: '';
  display: inline-block;
  width: 5px; height: 5px;
  background: var(--green);
  border-radius: 50%;
  margin-right: 6px;
  box-shadow: 0 0 6px var(--green);
  vertical-align: middle;
  margin-top: -1px;
}

/* ═══════════════════════════════════════════════════════
   BADGES
═══════════════════════════════════════════════════════ */
.badge {
  display: inline-flex; align-items: center;
  border-radius: 6px; padding: 3px 10px;
  font-size: 9.5px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.08em;
  font-family: var(--font);
}
.badge-green { background: rgba(77,187,128,0.12); color: var(--green); }
.badge-red   { background: rgba(212,80,80,0.12);  color: var(--red); }
.badge-gold  { background: rgba(200,168,75,0.12); color: var(--gold); }
.badge-gray  { background: rgba(94,80,48,0.18);   color: var(--tx3); }

/* ═══════════════════════════════════════════════════════
   BUTTONS
═══════════════════════════════════════════════════════ */
.stButton > button {
  font-family: var(--font) !important;
  border-radius: 8px !important;
  font-size: 11px !important;
  font-weight: 700 !important;
  letter-spacing: 0.04em !important;
  padding: 6px 10px !important;
  transition: all 0.2s cubic-bezier(0.32,0.72,0,1) !important;
}
.stButton > button[kind="primary"] {
  background: var(--gold) !important;
  color: #000 !important;
  border: none !important;
  box-shadow: 0 0 24px rgba(200,168,75,0.28), inset 0 1px 1px rgba(255,255,255,0.12) !important;
}
.stButton > button[kind="primary"]:hover {
  background: var(--gold-hi) !important;
  transform: translateY(-1px) !important;
  box-shadow: 0 0 36px rgba(200,168,75,0.38) !important;
}
.stButton > button[kind="primary"]:active { transform: scale(0.98) !important; }
.stButton > button:not([kind="primary"]) {
  background: var(--s2) !important;
  color: var(--tx3) !important;
  border: 1px solid var(--gold-xlo) !important;
}
.stButton > button:not([kind="primary"]):hover {
  border-color: rgba(200,168,75,0.28) !important;
  color: var(--tx2) !important;
  background: rgba(200,168,75,0.04) !important;
}

/* ═══════════════════════════════════════════════════════
   UPLOAD ZONE
═══════════════════════════════════════════════════════ */
[data-testid="stFileUploader"] {
  background: var(--s2) !important;
  border: 1.5px dashed var(--gold-lo) !important;
  border-radius: 14px !important;
  transition: all 0.2s !important;
}
[data-testid="stFileUploader"]:hover {
  border-color: var(--gold) !important;
  background: rgba(200,168,75,0.03) !important;
  box-shadow: 0 0 0 1px rgba(200,168,75,0.08),
              0 0 28px rgba(200,168,75,0.06) !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] div,
[data-testid="stFileUploaderDropzoneInstructions"] span {
  color: var(--tx3) !important;
  font-size: 12px !important;
  font-family: var(--font) !important;
}
[data-testid="stFileUploader"] label,
[data-testid="stFileUploader"] small {
  color: var(--tx3) !important;
  font-family: var(--font) !important;
}

/* ═══════════════════════════════════════════════════════
   RADIO
═══════════════════════════════════════════════════════ */
[data-testid="stRadio"] label {
  color: var(--tx3) !important;
  font-size: 10px !important;
  text-transform: uppercase !important;
  letter-spacing: 0.1em !important;
  font-family: var(--font) !important;
  font-weight: 700 !important;
}
[data-baseweb="radio-group"] {
  flex-direction: column !important;
  gap: 5px !important;
}
[data-baseweb="radio-group"] [data-baseweb="radio"] {
  background: var(--s2) !important;
  border: 1px solid var(--gold-xlo) !important;
  border-radius: 8px !important;
  padding: 7px 10px !important;
  gap: 8px !important;
  align-items: center !important;
  transition: border-color 0.18s, background 0.18s !important;
  cursor: pointer !important;
  width: 100% !important;
  box-sizing: border-box !important;
  margin: 0 !important;
}
[data-baseweb="radio-group"] [data-baseweb="radio"]:hover {
  border-color: rgba(200,168,75,0.22) !important;
  background: rgba(200,168,75,0.025) !important;
}
[data-baseweb="radio-group"] [data-baseweb="radio"]:has([aria-checked="true"]) {
  background: linear-gradient(90deg, rgba(200,168,75,0.10), rgba(200,168,75,0.04)) !important;
  border-color: rgba(200,168,75,0.28) !important;
  border-left-width: 2px !important;
  border-left-color: var(--gold) !important;
}
[data-baseweb="radio-group"] [role="radio"] {
  border-color: var(--gold-lo) !important;
  background: transparent !important;
  flex-shrink: 0 !important;
  width: 13px !important;
  height: 13px !important;
}
[data-baseweb="radio-group"] [role="radio"][aria-checked="true"] {
  background: var(--gold) !important;
  border-color: var(--gold) !important;
  box-shadow: 0 0 6px rgba(200,168,75,0.4) !important;
}
[data-baseweb="radio-group"] [data-testid="stMarkdownContainer"] p {
  font-size: 11px !important;
  font-weight: 600 !important;
  color: var(--tx2) !important;
  font-family: var(--font) !important;
  line-height: 1.3 !important;
  margin: 0 !important;
}
[data-baseweb="radio-group"] [data-baseweb="radio"]:has([aria-checked="true"])
[data-testid="stMarkdownContainer"] p {
  color: var(--tx1) !important;
  font-weight: 700 !important;
}
[data-testid="stRadio"] > label > div > p {
  font-size: 9px !important;
  font-weight: 800 !important;
  text-transform: uppercase !important;
  letter-spacing: 0.14em !important;
  color: var(--tx3) !important;
  font-family: var(--font) !important;
  margin-bottom: 6px !important;
}

/* ═══════════════════════════════════════════════════════
   EXPANDER
═══════════════════════════════════════════════════════ */
[data-testid="stExpander"] {
  background: var(--s1) !important;
  border: 1px solid var(--gold-xlo) !important;
  border-radius: 12px !important;
}
[data-testid="stExpander"] summary {
  color: var(--tx3) !important;
  font-size: 12px !important;
  font-family: var(--font) !important;
  font-weight: 600 !important;
  letter-spacing: 0.01em !important;
}
[data-testid="stExpander"] summary:hover { color: var(--gold) !important; }

/* ═══════════════════════════════════════════════════════
   ALERTS
═══════════════════════════════════════════════════════ */
[data-testid="stAlert"] {
  background: var(--s2) !important;
  border-radius: 12px !important;
  border: 1px solid var(--gold-xlo) !important;
}
[data-testid="stAlert"] p {
  color: var(--tx2) !important;
  font-family: var(--font) !important;
}

/* ═══════════════════════════════════════════════════════
   DOWNLOAD BUTTON
═══════════════════════════════════════════════════════ */
[data-testid="stDownloadButton"] > button {
  background: var(--s2) !important;
  color: var(--tx2) !important;
  border: 1px solid var(--gold-xlo) !important;
  border-radius: 10px !important;
  font-family: var(--font) !important;
  font-weight: 600 !important;
  font-size: 11.5px !important;
  transition: all 0.18s cubic-bezier(0.32,0.72,0,1) !important;
  letter-spacing: 0.01em !important;
}
[data-testid="stDownloadButton"] > button:hover {
  border-color: rgba(200,168,75,0.35) !important;
  color: var(--gold) !important;
  background: rgba(200,168,75,0.05) !important;
  transform: translateY(-1px) !important;
}

/* ═══════════════════════════════════════════════════════
   DATAFRAME
═══════════════════════════════════════════════════════ */
[data-testid="stDataFrame"] {
  border: 1px solid var(--gold-xlo) !important;
  border-radius: 10px !important;
  overflow: hidden !important;
}

/* ═══════════════════════════════════════════════════════
   CODE BLOCK
═══════════════════════════════════════════════════════ */
[data-testid="stCode"], .stCodeBlock pre {
  background: var(--s3) !important;
  border: 1px solid var(--gold-xlo) !important;
  border-radius: 10px !important;
  color: var(--tx1) !important;
  font-size: 12px !important;
  line-height: 1.7 !important;
  font-family: 'JetBrains Mono', 'Fira Code', 'Cascadia Code', monospace !important;
}

/* ═══════════════════════════════════════════════════════
   IMAGE
═══════════════════════════════════════════════════════ */
[data-testid="stImage"] img {
  border-radius: 12px !important;
  border: 1px solid var(--gold-xlo) !important;
  box-shadow: 0 0 40px rgba(200,168,75,0.06) !important;
}

/* ═══════════════════════════════════════════════════════
   SPINNER
═══════════════════════════════════════════════════════ */
[data-testid="stSpinner"] > div { border-top-color: var(--gold) !important; }

/* ═══════════════════════════════════════════════════════
   VEGA / ALTAIR
═══════════════════════════════════════════════════════ */
[data-testid="stArrowVegaLiteChart"] {
  background: var(--s1) !important;
  border: 1px solid var(--gold-xlo) !important;
  border-radius: 14px !important;
  padding: 18px !important;
}

/* ═══════════════════════════════════════════════════════
   TEXT AREA
═══════════════════════════════════════════════════════ */
[data-testid="stTextArea"] textarea {
  background: var(--s3) !important;
  border: 1px solid var(--gold-xlo) !important;
  border-radius: 10px !important;
  color: var(--tx2) !important;
  font-family: 'JetBrains Mono', 'Fira Code', monospace !important;
  font-size: 11px !important;
}
[data-testid="stTextArea"] textarea:focus {
  border-color: var(--gold-lo) !important;
  box-shadow: 0 0 0 2px rgba(200,168,75,0.12) !important;
}

/* ═══════════════════════════════════════════════════════
   ANIMATIONS
═══════════════════════════════════════════════════════ */
@keyframes slideUp {
  from { opacity: 0; transform: translateY(20px); }
  to   { opacity: 1; transform: translateY(0); }
}
@keyframes slideInLeft {
  from { opacity: 0; transform: translateX(-14px); }
  to   { opacity: 1; transform: translateX(0); }
}
@keyframes fadeUp {
  from { opacity: 0; transform: translateY(32px); }
  to   { opacity: 1; transform: translateY(0); }
}
@keyframes scanLine {
  0%   { left: -60%; }
  100% { left: 110%; }
}

.anim-slide-left { animation: slideInLeft 0.38s cubic-bezier(0.32,0.72,0,1) both; }
.anim-slide-up   { animation: slideUp     0.42s cubic-bezier(0.32,0.72,0,1) both; }
.anim-fade-up    { animation: fadeUp      0.5s  cubic-bezier(0.32,0.72,0,1) both; }
.results-reveal  { animation: fadeUp 0.52s cubic-bezier(0.32,0.72,0,1) both; }

/* ═══════════════════════════════════════════════════════
   MAIN LAYOUT
═══════════════════════════════════════════════════════ */
.main-layout { padding: 6px 24px 0; }

/* ═══════════════════════════════════════════════════════
   NAV PANEL (left sidebar)
═══════════════════════════════════════════════════════ */
.sidebar-label {
  font-size: 9px; font-weight: 800;
  text-transform: uppercase; letter-spacing: 0.18em;
  color: var(--tx3); padding: 8px 0 5px;
  font-family: Arial, sans-serif;
}
/* ── Flat nav panel ── */
.nav-panel {
  padding: 4px 0;
  position: sticky;
  top: 80px;
}
/* Section labels with trailing rule */
.nav-section-head {
  display: flex; align-items: center; gap: 10px;
  font-size: 9px; font-weight: 800;
  text-transform: uppercase; letter-spacing: 0.22em;
  color: rgba(200,168,75,0.4);
  padding: 10px 4px 5px;
  font-family: Arial, sans-serif;
}
.nav-section-head::after {
  content: '';
  flex: 1; height: 1px;
  background: rgba(200,168,75,0.12);
}
.nav-section-sep { margin-top: 8px; }
/* Nav items */
.nav-item {
  display: flex; align-items: center; gap: 11px;
  padding: 7px 8px 7px 12px;
  border-radius: 0 6px 6px 0;
  cursor: pointer;
  min-width: 0;
  margin-bottom: 1px;
  box-shadow: inset 0 0 0 transparent;
  transition: background 0.18s, box-shadow 0.18s;
}
.nav-item:hover:not(.active) {
  background: rgba(200,168,75,0.04);
}
/* Active: left accent via inset shadow (avoids side-stripe ban) */
.nav-item.active {
  box-shadow: inset 2px 0 0 var(--gold);
  background: rgba(200,168,75,0.06);
}
/* Number */
.nav-item-num {
  font-size: 10.5px; font-weight: 500;
  color: var(--tx3); flex-shrink: 0;
  width: 14px; text-align: right;
  font-family: Arial, sans-serif;
  font-variant-numeric: tabular-nums;
  transition: color 0.18s;
}
.nav-item.active .nav-item-num { color: var(--gold); font-weight: 700; }
.nav-item:hover:not(.active) .nav-item-num { color: var(--tx2); }
/* Label text */
.nav-item-text {
  font-size: 12px; font-weight: 600;
  color: var(--tx3); font-family: Arial, sans-serif;
  line-height: 1.3;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  min-width: 0;
  transition: color 0.18s;
}
.nav-item.active .nav-item-text { color: var(--tx1); font-weight: 700; }
.nav-item:hover:not(.active) .nav-item-text { color: var(--tx2); }

/* ═══════════════════════════════════════════════════════
   HIDE NAV BUTTONS (wired via JS)
═══════════════════════════════════════════════════════ */
[data-testid="stHorizontalBlock"]:first-of-type
[data-testid="stColumn"]:first-child
.stButton,
[data-testid="stHorizontalBlock"]:first-of-type
[data-testid="stColumn"]:first-child
iframe {
  display: none !important;
}

/* ═══════════════════════════════════════════════════════
   CONTROL STRIP SECTIONS
═══════════════════════════════════════════════════════ */
.ctrl-section-label {
  font-size: 8.5px; font-weight: 800;
  text-transform: uppercase; letter-spacing: 0.16em;
  color: var(--tx3); margin-bottom: 8px;
  font-family: Arial, sans-serif;
}
.desc-title {
  font-size: 12px; font-weight: 700;
  color: var(--gold); line-height: 1.35;
  margin-bottom: 5px;
  font-family: Arial, sans-serif;
}
.desc-notes {
  font-size: 10.5px; color: var(--tx3);
  line-height: 1.5; margin-bottom: 6px;
  font-family: Arial, sans-serif;
}
.desc-sheet {
  display: inline-block;
  font-size: 9px; color: var(--gold-lo);
  font-weight: 700; text-transform: uppercase;
  letter-spacing: 0.08em; background: var(--gold-xlo);
  border-radius: 4px; padding: 2px 7px;
  font-family: Arial, sans-serif;
}
.run-hint {
  font-size: 9.5px; color: var(--tx3);
  text-align: center; margin-top: 5px;
  font-family: Arial, sans-serif;
}

/* ═══════════════════════════════════════════════════════
   CONTROL STRIP — flat, no cards
═══════════════════════════════════════════════════════ */
.file-ready-compact {
  font-size: 10.5px; color: var(--tx2);
  padding: 4px 0; font-family: Arial, sans-serif;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  display: flex; align-items: center; gap: 5px; flex-wrap: wrap;
}
.file-size-sm { color: var(--tx3); font-size: 9.5px; }

/* ═══════════════════════════════════════════════════════
   SECTION LABELS (output area)
═══════════════════════════════════════════════════════ */
.section-bar {
  display: flex; align-items: center; gap: 8px;
  margin-bottom: 10px;
}
.section-bar-label {
  font-size: 9.5px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.14em;
  color: var(--tx3); font-family: Arial, sans-serif;
}
.output-divider {
  border: none; border-top: 1px solid var(--gold-xlo);
  margin: 14px 0 16px;
}

/* ═══════════════════════════════════════════════════════
   TABLE SECTION
═══════════════════════════════════════════════════════ */
.table-sheet-name {
  font-size: 11.5px; font-weight: 600;
  color: var(--tx2); margin-bottom: 5px;
  font-family: Arial, sans-serif;
}
.table-meta {
  font-size: 9.5px; color: var(--tx3);
  margin-left: 5px; font-weight: 400;
}
.table-hint {
  font-size: 10px; color: var(--tx3);
  margin-top: 4px; font-family: Arial, sans-serif;
}

/* ═══════════════════════════════════════════════════════
   EMPTY STATES
═══════════════════════════════════════════════════════ */
.empty-state {
  display: flex; flex-direction: column;
  align-items: center; justify-content: center;
  min-height: 380px;
  border: 1.5px dashed var(--gold-xlo);
  border-radius: 20px; margin-top: 16px;
  padding: 48px 32px;
  background: radial-gradient(ellipse at 50% 80%, rgba(200,168,75,0.03) 0%, transparent 65%);
}
.empty-state-icon {
  font-size: 40px; margin-bottom: 18px;
  opacity: 0.2; color: var(--gold); line-height: 1;
}
.empty-state-title {
  font-size: 15px; font-weight: 700;
  color: var(--tx3); font-family: Arial, sans-serif;
  letter-spacing: -0.01em; margin-bottom: 8px;
}
.empty-state-desc {
  font-size: 12px; color: var(--tx3);
  font-family: Arial, sans-serif;
  text-align: center; max-width: 280px; line-height: 1.7;
}
.empty-chart {
  color: var(--tx3); font-size: 12px;
  padding: 48px 16px; text-align: center;
  border: 1px dashed var(--gold-xlo);
  border-radius: 12px; font-family: Arial, sans-serif;
}

/* file-ready scan shimmer */
.file-ready {
  position: relative; overflow: hidden;
  margin-top: 6px; padding: 8px 12px;
  background: rgba(77,187,128,0.07);
  border: 1px solid rgba(77,187,128,0.22);
  border-radius: 8px;
  display: flex; align-items: center; gap: 8px;
  animation: slideInLeft 0.38s cubic-bezier(0.32,0.72,0,1) both;
  font-family: Arial, sans-serif;
}
.file-ready::after {
  content: '';
  position: absolute; top: 0; bottom: 0; width: 60%;
  background: linear-gradient(90deg, transparent, rgba(255,255,255,0.04), transparent);
  animation: scanLine 1.1s cubic-bezier(0.4,0,0.6,1) 0.38s 1 both;
  pointer-events: none;
}

hr { border-color: var(--gold-xlo) !important; margin: 20px 0 !important; }
</style>"""


# ---------------------------------------------------------------------------
# Nav JS wiring
# ---------------------------------------------------------------------------

_NAV_JS = """<script>
(function() {
  var doc = window.parent.document;
  function setup() {
    var items = doc.querySelectorAll('.nav-item');
    if (!items.length) return;
    var hb = doc.querySelector('[data-testid="stHorizontalBlock"]');
    if (!hb) return;
    var col = hb.querySelector('[data-testid="stColumn"]');
    if (!col) return;
    var btns = Array.from(col.querySelectorAll('button'));
    if (!btns.length) return;
    items.forEach(function(el, i) {
      if (el._navW) return;
      el._navW = 1;
      el.addEventListener('click', function() {
        if (btns[i]) btns[i].click();
      });
    });
  }
  var t;
  new MutationObserver(function() { clearTimeout(t); t = setTimeout(setup, 160); })
    .observe(doc.body, { childList: true, subtree: true });
  [350, 800, 1600].forEach(function(d) { setTimeout(setup, d); });
})();
</script>"""


# ---------------------------------------------------------------------------
# HTML helpers
# ---------------------------------------------------------------------------

_ABINBEV_LOGO_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 260 60" height="36" aria-label="AB InBev">
  <g transform="translate(30,30)">
    <path d="M1,-1 C3,-22 20,-24 20,-8 C14,2 4,4 1,-1Z" fill="#F5C518"/>
    <path d="M1,-1 C3,-22 20,-24 20,-8 C14,2 4,4 1,-1Z" fill="#D4A800" transform="rotate(120,0,0)"/>
    <path d="M1,-1 C3,-22 20,-24 20,-8 C14,2 4,4 1,-1Z" fill="#F5C518" transform="rotate(240,0,0)"/>
  </g>
  <text x="68" y="42" font-family="Arial Black,Arial,sans-serif" font-size="32" font-weight="900" fill="#ffffff" letter-spacing="-0.5">ABInBev</text>
</svg>"""


def _header_html() -> str:
    import base64 as _b64
    logo_path = APP_DIR / "logo.png"
    if logo_path.exists():
        logo_b64 = _b64.b64encode(logo_path.read_bytes()).decode()
        logo_el = f'<img src="data:image/png;base64,{logo_b64}" class="prp-abinbev-logo" alt="AB InBev">'
    else:
        logo_el = f'<div class="prp-abinbev-logo">{_ABINBEV_LOGO_SVG}</div>'
    return f"""
<div class="prp-header">
  <div class="prp-logo">
    {logo_el}
    <div>
      <div class="prp-brand-name">PRP Automation Dashboard</div>
      <div class="prp-brand-sub">TPRM &amp; Risk Intelligence</div>
    </div>
  </div>
  <div class="prp-header-right">
    <span class="prp-badge prp-badge-live">Live</span>
    <span class="prp-badge">8 Reports</span>
    <span class="prp-badge">AB InBev</span>
  </div>
</div>"""


# ---------------------------------------------------------------------------
# Live chart builder
# ---------------------------------------------------------------------------

def _build_plotly_bar(chart_df: pd.DataFrame, title: str = ""):
    """Return a styled Plotly Figure matching the dark-gold design system."""
    import plotly.graph_objects as go

    gold_palette = [
        "#C8A84B", "#4DBB80", "#D45050",
        "#6496D2", "#BE6EBE", "#DCA03C",
    ]
    fill_palette = [
        "rgba(200,168,75,0.82)", "rgba(77,187,128,0.75)",
        "rgba(212,80,80,0.72)", "rgba(100,150,210,0.72)",
        "rgba(190,110,190,0.72)", "rgba(220,160,60,0.75)",
    ]

    x_labels = chart_df.index.astype(str).tolist()
    fig = go.Figure()
    for i, col in enumerate(chart_df.columns):
        y_vals = pd.to_numeric(chart_df[col], errors="coerce").fillna(0).tolist()
        fig.add_trace(go.Bar(
            name=str(col),
            x=x_labels,
            y=y_vals,
            marker=dict(
                color=fill_palette[i % len(fill_palette)],
                line=dict(color=gold_palette[i % len(gold_palette)], width=0),
                cornerradius=4,
            ),
            hovertemplate=(
                "<b>%{x}</b><br>"
                + str(col) + ": <b>%{y:,.0f}</b><extra></extra>"
            ),
        ))

    grid_color = "rgba(34,27,11,0.9)"
    tick_color = "#A89060"
    font_cfg   = dict(family="Arial, sans-serif", color=tick_color, size=11)

    fig.update_layout(
        title=dict(
            text=title,
            font=dict(family="Arial, sans-serif", color="#A89060", size=11),
            x=0, xanchor="left", pad=dict(l=4, b=12),
        ) if title else None,
        paper_bgcolor="#0d0d0d",
        plot_bgcolor="#0d0d0d",
        font=font_cfg,
        barmode="group",
        bargap=0.22,
        bargroupgap=0.06,
        margin=dict(l=48, r=20, t=40 if title else 20, b=60),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
            font=dict(family="Arial, sans-serif", color="#A89060", size=10),
            bgcolor="rgba(0,0,0,0)",
            bordercolor="rgba(0,0,0,0)",
        ),
        xaxis=dict(
            gridcolor=grid_color, linecolor="#2A1F08", tickcolor="#2A1F08",
            tickfont=dict(family="Arial, sans-serif", color=tick_color, size=10),
            tickangle=-25,
        ),
        yaxis=dict(
            gridcolor=grid_color, linecolor="#2A1F08", tickcolor="#2A1F08",
            tickfont=dict(family="Arial, sans-serif", color=tick_color, size=10),
            zeroline=False,
        ),
        hoverlabel=dict(
            bgcolor="#191919",
            bordercolor="#221B0B",
            font=dict(family="Arial, sans-serif", color="#F0E4C0", size=12),
        ),
        transition=dict(duration=700, easing="cubic-in-out"),
    )
    return fig


# ---------------------------------------------------------------------------
# Diagram 2 — custom chart renderer
# ---------------------------------------------------------------------------

def _d2_charts(xlsx_bytes: bytes) -> int:
    """Parse Diagram 2 Dashboard sheet and render two Plotly charts.
    Returns the number of charts successfully rendered."""
    import io as _io
    import plotly.graph_objects as go

    try:
        raw = pd.read_excel(
            _io.BytesIO(xlsx_bytes), sheet_name="Dashboard",
            header=None, engine="openpyxl"
        )
    except Exception:
        return 0

    # Locate KPI table header row (contains "Metric")
    kpi_row = next(
        (i for i, r in raw.iterrows() if "Metric" in r.values), None
    )

    _grid  = "rgba(34,27,11,0.9)"
    _tick  = "#A89060"
    _font  = dict(family="Arial, sans-serif", color=_tick, size=10)
    _hover = dict(bgcolor="#191919", bordercolor="#221B0B",
                  font=dict(family="Arial, sans-serif", color="#F0E4C0", size=12))
    charts = 0

    # ── Improve in Supplier Response Time (KPI horizontal bar) ──────────────
    if kpi_row is not None:
        kv = raw.iloc[kpi_row:].copy()
        kv.columns = kv.iloc[0]
        kv = kv.iloc[1:].reset_index(drop=True)
        kv = kv[kv["Metric"].notna() & (kv["Metric"].astype(str).str.strip() != "")].copy()
        kv["Value"] = pd.to_numeric(kv["Value"], errors="coerce")
        kv = kv.dropna(subset=["Value"])
        kv = kv[~kv["Metric"].astype(str).str.contains("formula", case=False, na=False)]

        if not kv.empty:
            metrics = kv["Metric"].astype(str).tolist()
            values  = kv["Value"].tolist()
            max_val = max(values) if values else 1
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(
                name="", x=values, y=metrics, orientation="h",
                marker=dict(color="rgba(200,168,75,0.78)", cornerradius=4),
                text=[f"{v:.0%}" for v in values],
                textposition="outside",
                textfont=dict(color=_tick, size=10, family="Arial, sans-serif"),
                hovertemplate="<b>%{y}</b>: <b>%{x:.0%}</b><extra></extra>",
            ))
            fig2.update_layout(
                title=dict(text="Improve in Supplier Response Time",
                           font=dict(family="Arial, sans-serif", color=_tick, size=11),
                           x=0, xanchor="left", pad=dict(l=4, b=12)),
                paper_bgcolor="#0d0d0d", plot_bgcolor="#0d0d0d", font=_font,
                margin=dict(l=130, r=70, t=40, b=40),
                xaxis=dict(gridcolor=_grid, linecolor="#2A1F08", tickformat=".0%",
                           tickfont=_font, zeroline=False,
                           range=[0, max_val * 1.3]),
                yaxis=dict(gridcolor=_grid, linecolor="#2A1F08", tickfont=_font),
                hoverlabel=_hover,
                showlegend=False,
            )
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})
            charts += 1

    return charts


# ---------------------------------------------------------------------------
# Results renderer
# ---------------------------------------------------------------------------

def render_results(entry: ScriptEntry, result: RunResult, fmt: str) -> None:
    st.markdown('<div class="results-reveal">', unsafe_allow_html=True)

    if not result.outputs:
        st.warning("The script produced no output file.")
        st.markdown('</div>', unsafe_allow_html=True)
        return

    # Pre-process sheets
    processed_sheets: list[tuple[str, str, pd.DataFrame]] = []
    for fname, sheets in result.tables.items():
        for sheet_name, df in sheets.items():
            # Diagram 4: skip "History Used" sheet; rename Sheet1 from History.xlsx
            if entry.id == "s4":
                if sheet_name == "History Used":
                    continue
                if sheet_name == "Sheet1" and fname.lower() == "history.xlsx":
                    sheet_name = "History.xlsx"
            # Risks Identified: skip Open Overdue Pivot table entirely
            if entry.id == "s3d" and sheet_name == "Open Overdue Pivot":
                continue
            # Suppliers in Scope: skip Pivot table entirely
            if entry.id == "s1c" and sheet_name == "Pivot":
                continue
            # Risk Assessment Progress: skip Auto Pivot Summary, show the other sheets with charts
            if entry.id in ("s2a", "s2b") and sheet_name == "Auto Pivot Summary":
                continue
            processed_sheets.append((fname, sheet_name, _trim_sparse_cols(_trim_sparse_rows(df))))

    st.markdown('<hr class="output-divider">', unsafe_allow_html=True)

    # ── Row 1: Download Options | Copy Text ──────────────────────────────────
    c_dl, c_copy = st.columns([6, 4], gap="small")

    filtered_outputs = [
        (name, data) for name, data in result.outputs
        if not (entry.id == "d1" and name.lower().endswith(".png"))
    ]

    with c_dl:
        st.markdown(
            f'<div class="section-bar">'
            f'<span class="section-bar-label">Download Options</span>'
            f'<span class="badge badge-gold">{len(filtered_outputs)} file{"s" if len(filtered_outputs) != 1 else ""}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )
        dl_cols = st.columns(min(3, len(filtered_outputs)))
        for i, (name, data) in enumerate(filtered_outputs):
            mime = (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                if name.lower().endswith(".xlsx") else
                "image/png" if name.lower().endswith(".png") else
                "application/octet-stream"
            )
            dl_cols[i % len(dl_cols)].download_button(
                f"⬇  {name}", data=data, file_name=name, mime=mime,
                key=f"dl_{entry.id}_{i}", use_container_width=True,
            )

    with c_copy:
        st.markdown(
            '<div class="section-bar"><span class="section-bar-label">Copy Text</span></div>',
            unsafe_allow_html=True,
        )
        if processed_sheets:
            import json as _json
            fname, sheet_name, df = processed_sheets[0]
            clean_df = _clean_for_tsv(df)
            preview_df = clean_df.head(5)
            if fmt == "Markdown":
                try:
                    copy_text = preview_df.to_markdown(index=False)
                except Exception:
                    copy_text = preview_df.to_string(index=False)
            else:
                copy_text = preview_df.to_csv(sep="\t", index=False)
            js_text = (
                _json.dumps(copy_text)
                .replace("<", r"<")
                .replace(">", r">")
                .replace("&", r"&")
            )
            components.html(f"""
<style>
  body {{ margin: 0; background: transparent; }}
  .cb {{
    background: #0d0d0d;
    color: #B89A68;
    border: 1px solid #221B0B;
    border-radius: 8px;
    padding: 9px 14px;
    font-family: Arial, sans-serif;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.06em;
    cursor: pointer;
    width: 100%;
    text-transform: uppercase;
    transition: border-color 0.18s, color 0.18s, background 0.18s;
  }}
  .cb:hover {{ border-color: rgba(200,168,75,0.35); color: #C8A84B; background: rgba(200,168,75,0.05); }}
  .cb.ok {{ border-color: rgba(77,187,128,0.35); color: #4DBB80; background: rgba(77,187,128,0.07); }}
  .hint {{ font-size: 9px; color: #6A5A38; font-family: Arial, sans-serif; margin-top: 7px; }}
</style>
<button class="cb" id="cb">⧉  Copy Text</button>
<div class="hint">First 5 rows &middot; download .xlsx for full data</div>
<script>
  var _t = {js_text};
  document.getElementById('cb').addEventListener('click', function() {{
    var b = this;
    function done() {{
      b.textContent = '✓ Copied!';
      b.classList.add('ok');
      setTimeout(function() {{ b.textContent = '⧉  Copy Text'; b.classList.remove('ok'); }}, 2000);
    }}
    if (navigator.clipboard && window.isSecureContext) {{
      navigator.clipboard.writeText(_t).then(done);
    }} else {{
      var a = document.createElement('textarea');
      a.value = _t;
      a.style.cssText = 'position:fixed;opacity:0';
      document.body.appendChild(a);
      a.focus(); a.select();
      document.execCommand('copy');
      document.body.removeChild(a);
      done();
    }}
  }});
</script>
""", height=68)

    st.markdown('<div style="height:18px;"></div>', unsafe_allow_html=True)

    # ── Chart (left) | Table (right) per sheet ───────────────────────────────
    st.markdown(
        '<div class="section-bar"><span class="section-bar-label">Process &middot; Preview</span></div>',
        unsafe_allow_html=True,
    )

    charts_rendered = 0
    d2_xlsx_bytes = (
        next((d for n, d in result.outputs if n.lower().endswith(".xlsx")), None)
        if entry.id == "d2" else None
    )

    for fname, sheet_name, df in processed_sheets:
        row_count = len(df)
        col_count = len(df.columns)

        # Determine if a chart should render for this sheet
        show_chart = True
        if entry.id == "d2":
            show_chart = charts_rendered == 0
        elif entry.id == "s4" and charts_rendered > 0:
            show_chart = False
        elif entry.id == "s3d" and sheet_name == "Open Overdue Pivot":
            show_chart = False

        if show_chart:
            c_chart, c_table = st.columns([7, 5], gap="small")
            with c_chart:
                if entry.id == "d2":
                    if d2_xlsx_bytes:
                        charts_rendered = _d2_charts(d2_xlsx_bytes)
                else:
                    chart_df = _chartable(df)
                    if chart_df is not None:
                        st.plotly_chart(
                            _build_plotly_bar(chart_df, sheet_name),
                            use_container_width=True,
                            config={"displayModeBar": False},
                        )
                        charts_rendered += 1
            with c_table:
                st.markdown(
                    f'<div class="table-sheet-name">{sheet_name}'
                    f'<span class="table-meta">{fname} &nbsp;&middot;&nbsp; {row_count}r &times; {col_count}c</span></div>',
                    unsafe_allow_html=True,
                )
                st.dataframe(
                    df.head(50),
                    use_container_width=True,
                    hide_index=True,
                    height=min(320, 38 + len(df.head(50)) * 35),
                )
                if row_count > 50:
                    st.markdown(
                        f'<div class="table-hint">Showing 50 of {row_count} rows — download .xlsx for full dataset.</div>',
                        unsafe_allow_html=True,
                    )
        else:
            st.markdown(
                f'<div class="table-sheet-name">{sheet_name}'
                f'<span class="table-meta">{fname} &nbsp;&middot;&nbsp; {row_count}r &times; {col_count}c</span></div>',
                unsafe_allow_html=True,
            )
            st.dataframe(
                df.head(50),
                use_container_width=True,
                hide_index=True,
                height=min(320, 38 + len(df.head(50)) * 35),
            )
            if row_count > 50:
                st.markdown(
                    f'<div class="table-hint">Showing 50 of {row_count} rows — download .xlsx for full dataset.</div>',
                    unsafe_allow_html=True,
                )

    if charts_rendered == 0 and not processed_sheets:
        st.markdown('<div class="empty-chart">No chartable data in output.</div>', unsafe_allow_html=True)

    # ── Script log ────────────────────────────────────────────────────────────
    st.markdown('<div style="margin-top:20px;"></div>', unsafe_allow_html=True)
    with st.expander("Script log  ·  stdout / stderr", expanded=not result.ok):
        if result.stdout.strip():
            st.code(result.stdout, language="text")
        if result.stderr.strip():
            st.markdown(
                '<div style="font-size:10px;text-transform:uppercase;letter-spacing:0.1em;'
                'color:var(--tx3);margin:10px 0 6px;font-family:Arial,sans-serif;">stderr</div>',
                unsafe_allow_html=True,
            )
            st.code(result.stderr, language="text")
        if not result.stdout.strip() and not result.stderr.strip():
            st.markdown(
                '<div style="color:var(--tx3);font-size:12px;padding:4px 0;'
                'font-family:Arial,sans-serif;">No console output.</div>',
                unsafe_allow_html=True,
            )

    st.markdown('</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Main UI
# ---------------------------------------------------------------------------

def main() -> None:
    st.set_page_config(
        page_title="PRP Automation Dashboard — AB InBev",
        page_icon="⬡",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    st.markdown(_CSS, unsafe_allow_html=True)
    st.markdown(_header_html(), unsafe_allow_html=True)

    groups: dict[str, list[ScriptEntry]] = {}
    for e in REGISTRY:
        groups.setdefault(e.group, []).append(e)
    group_names = list(groups.keys())

    _short = {
        group_names[0]: "Vendor Onboarding (Critical Tech)",
        group_names[1]: "Response Time Reduction",
        group_names[2]: "Reduce Long Over-Due Assessments",
        group_names[3]: "Risk Treatment",
        group_names[4]: "Suppliers in Scope",
        group_names[5]: "Risk Assessment Progress",
        group_names[6]: "Risks Identified",
    }

    if "active_group" not in st.session_state:
        st.session_state["active_group"] = group_names[0]

    active_group = st.session_state["active_group"]
    g_entries = groups[active_group]
    labels = [e.label for e in g_entries]

    st.markdown('<div class="main-layout">', unsafe_allow_html=True)

    # ── 2-column outer layout: sidebar | main ─────────────────────────────────
    col_sidebar, col_main = st.columns([2, 10], gap="small")

    # ── SIDEBAR ───────────────────────────────────────────────────────────────
    with col_sidebar:
        nav_html = '<div class="nav-panel">'
        for idx, gname in enumerate(group_names, 1):
            is_active = st.session_state["active_group"] == gname
            active_cls = " active" if is_active else ""
            short_label = _short.get(gname, gname)
            if idx == 1:
                nav_html += '<div class="nav-section-head">Diagrams</div>'
            elif idx == 5:
                nav_html += '<div class="nav-section-head nav-section-sep">Slide 12</div>'
            nav_html += (
                f'<div class="nav-item{active_cls}" title="{gname}">'
                f'<div class="nav-item-num">{idx}</div>'
                f'<div class="nav-item-text">{short_label}</div>'
                f'</div>'
            )
        nav_html += '</div>'
        st.markdown(nav_html, unsafe_allow_html=True)

        for gname in group_names:
            if st.button(
                _short.get(gname, gname),
                key=f"grp_{gname}",
                type="secondary",
                use_container_width=True,
            ):
                st.session_state["active_group"] = gname
                st.rerun()
        components.html(_NAV_JS, height=0, scrolling=False)

    # ── MAIN ──────────────────────────────────────────────────────────────────
    with col_main:

        # ── CONTROL STRIP: input | variant | process | description | ● | run ──
        fmt = "TSV"

        c_input, c_variant, c_desc, c_run = st.columns(
            [3, 3, 6, 3], gap="small"
        )

        # INPUT
        with c_input:
            st.markdown('<div class="ctrl-section-label">Input</div>', unsafe_allow_html=True)
            uploaded = st.file_uploader(
                "Workbook", type=["xlsx"], label_visibility="collapsed",
            )
            if uploaded is not None:
                st.session_state["uploaded_bytes"] = uploaded.getvalue()
                st.session_state["uploaded_name"] = uploaded.name
                st.session_state["uploaded_size"] = uploaded.size

            file_bytes = st.session_state.get("uploaded_bytes")
            file_name  = st.session_state.get("uploaded_name", "")
            file_size  = st.session_state.get("uploaded_size", 0)
            has_file   = file_bytes is not None

            if has_file:
                st.markdown(
                    f'<div class="file-ready">'
                    f'<span style="color:var(--green);font-weight:700;">&#10003;</span>'
                    f'<span style="color:var(--tx1);font-size:11px;">{file_name}</span>'
                    f'<span class="file-size-sm" style="margin-left:auto;">{file_size/1024:.1f} KB</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

        # VARIANT
        with c_variant:
            st.markdown('<div class="ctrl-section-label">Variant</div>', unsafe_allow_html=True)
            selected_label = st.radio(
                "Variant", labels,
                key=f"variant_{active_group}",
                label_visibility="collapsed",
            )
            entry = g_entries[labels.index(selected_label)]
            st.session_state["selected_entry"] = entry

        # DESCRIPTION
        with c_desc:
            st.markdown('<div class="ctrl-section-label">Description</div>', unsafe_allow_html=True)
            st.markdown(
                f'<div class="desc-title">{entry.label}</div>'
                f'<div class="desc-notes">{entry.notes}</div>'
                f'<span class="desc-sheet">Sheet: {entry.sheet}</span>',
                unsafe_allow_html=True,
            )

        # RUN
        with c_run:
            st.markdown('<div class="ctrl-section-label">Run</div>', unsafe_allow_html=True)
            run = st.button(
                "▶  Run",
                type="primary",
                disabled=not has_file,
                use_container_width=True,
            )
            if not has_file:
                st.markdown(
                    '<div class="run-hint">Upload first</div>',
                    unsafe_allow_html=True,
                )

        # ── RUN HANDLER ───────────────────────────────────────────────────────
        if run and has_file:
            with st.spinner(f"Running  ·  {entry.label}"):
                try:
                    result = run_script(entry, file_bytes)
                except subprocess.TimeoutExpired:
                    st.error("Script exceeded 3-minute timeout.")
                    st.markdown('</div>', unsafe_allow_html=True)
                    return
                except Exception as exc:
                    st.error(f"Could not launch script: {exc}")
                    st.markdown('</div>', unsafe_allow_html=True)
                    return
            st.session_state["last_result"]   = result
            st.session_state["last_entry_id"] = entry.id
            st.session_state["last_fmt"]      = fmt
            st.rerun()

        # ── OUTPUT ────────────────────────────────────────────────────────────
        if "last_result" in st.session_state:
            render_results(
                REGISTRY_BY_ID[st.session_state["last_entry_id"]],
                st.session_state["last_result"],
                st.session_state["last_fmt"],
            )
        else:
            st.markdown("""
<div class="empty-state anim-fade-up">
  <div class="empty-state-icon">&#9678;</div>
  <div class="empty-state-title">Output will appear here</div>
  <div class="empty-state-desc">
    Select a diagram type, upload your Excel workbook,
    and click <strong style="color:var(--gold);">&#9654;&nbsp;Run</strong> to generate results.
  </div>
</div>""", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)  # /main-layout
    st.markdown('<div style="height:40px;"></div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
