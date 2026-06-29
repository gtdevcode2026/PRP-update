"""
PRP Automation Studio — AB InBev
=================================
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

    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
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
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800&display=swap');

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
  --tx1:      #EDE0B8;
  --tx2:      #A89060;
  --tx3:      #5E5030;
  --green:    #4DBB80;
  --red:      #D45050;
  --font:     'Plus Jakarta Sans', system-ui, -apple-system, sans-serif;
}

* { box-sizing: border-box; }

body, .stApp {
  font-family: var(--font) !important;
  background: var(--bg) !important;
  color: var(--tx1) !important;
}

/* Ambient background glow */
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
.stApp > header, [data-testid="stHeader"],
#MainMenu, footer,
[data-testid="stDecoration"],
[data-testid="stToolbar"] { display: none !important; }

.main .block-container {
  padding-top: 0 !important;
  max-width: 1440px !important;
  padding-left: 0 !important;
  padding-right: 0 !important;
}

/* ── Typography ── */
h1,h2,h3,h4 {
  font-family: var(--font) !important;
  color: var(--tx1) !important;
  letter-spacing: -0.03em !important;
}
.stMarkdown p, .stMarkdown li {
  font-family: var(--font) !important;
  color: var(--tx2) !important;
  font-size: 13px !important;
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
.prp-logo-mark {
  width: 38px; height: 38px;
  background: linear-gradient(135deg, #C8A84B 0%, #E8C84E 50%, #9A7A30 100%);
  border-radius: 11px;
  display: flex; align-items: center; justify-content: center;
  font-size: 19px; line-height: 1;
  box-shadow: 0 0 28px rgba(200,168,75,0.35), inset 0 1px 1px rgba(255,255,255,0.2);
  flex-shrink: 0;
}
.prp-brand-name {
  font-size: 16px; font-weight: 800;
  color: var(--tx1); letter-spacing: -0.025em; line-height: 1.1;
}
.prp-brand-sub {
  font-size: 10.5px; color: var(--tx3);
  letter-spacing: 0.06em; text-transform: uppercase;
  margin-top: 2px; font-weight: 500;
}
.prp-header-right { display: flex; align-items: center; gap: 8px; }
.prp-badge {
  font-size: 10px; font-weight: 700;
  color: var(--gold);
  background: rgba(200,168,75,0.08);
  border: 1px solid rgba(200,168,75,0.16);
  border-radius: 100px;
  padding: 5px 13px;
  letter-spacing: 0.04em;
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
   LAYOUT WRAPPER
═══════════════════════════════════════════════════════ */
.prp-wrap { padding: 0 44px; }

/* ═══════════════════════════════════════════════════════
   KPI STRIP
═══════════════════════════════════════════════════════ */
.kpi-strip {
  display: grid;
  grid-template-columns: repeat(6, 1fr);
  background: var(--gold-xxlo);
  border: 1px solid var(--gold-xlo);
  border-radius: 16px;
  overflow: hidden;
  margin: 28px 0 32px;
  gap: 1px;
}
@media(max-width:1100px){ .kpi-strip { grid-template-columns: repeat(3,1fr); } }
@media(max-width:700px)  { .kpi-strip { grid-template-columns: repeat(2,1fr); } }

.kpi-cell {
  background: var(--s1);
  padding: 22px 26px;
  transition: background 0.2s;
  position: relative;
}
.kpi-cell:hover { background: var(--s2); }
.kpi-cell:first-child::after {
  content: '';
  position: absolute; top: 0; left: 0; right: 0; height: 2px;
  background: linear-gradient(90deg, var(--gold), var(--gold-hi), transparent);
  opacity: 0.7;
}
.kpi-eyebrow {
  font-size: 9.5px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.12em;
  color: var(--tx3); margin-bottom: 10px;
  font-family: var(--font);
}
.kpi-value {
  font-size: 32px; font-weight: 800;
  line-height: 1; letter-spacing: -0.04em;
  color: var(--gold);
  font-family: var(--font);
}
.kpi-value-sm { font-size: 14px; letter-spacing: -0.01em; font-weight: 600; }
.kpi-detail { font-size: 10.5px; color: var(--tx3); margin-top: 7px; font-family: var(--font); }

/* ═══════════════════════════════════════════════════════
   SECTION LABEL
═══════════════════════════════════════════════════════ */
.prp-label {
  display: flex; align-items: center; gap: 10px;
  font-size: 9.5px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.14em;
  color: var(--tx3); margin: 28px 0 14px;
  font-family: var(--font);
}
.prp-label .prp-label-num {
  display: inline-flex; align-items: center; justify-content: center;
  width: 18px; height: 18px;
  background: var(--gold-xlo); border: 1px solid var(--gold-lo);
  border-radius: 5px; font-size: 9px; font-weight: 800;
  color: var(--gold-lo);
}
.prp-label::after {
  content: ''; flex: 1; height: 1px;
  background: linear-gradient(90deg, var(--gold-xlo), transparent);
}

/* ═══════════════════════════════════════════════════════
   CARDS
═══════════════════════════════════════════════════════ */
.prp-card {
  background: var(--s1);
  border: 1px solid var(--gold-xlo);
  border-radius: 16px;
  overflow: hidden;
  transition: border-color 0.25s cubic-bezier(0.32,0.72,0,1),
              box-shadow   0.25s cubic-bezier(0.32,0.72,0,1);
}
.prp-card:hover {
  border-color: rgba(200,168,75,0.2);
  box-shadow: 0 0 48px rgba(200,168,75,0.06);
}
.prp-card-accent-bar {
  height: 2px;
  background: linear-gradient(90deg, var(--gold) 0%, var(--gold-hi) 35%, transparent 100%);
  opacity: 0.65;
}
.prp-card-header {
  display: flex; align-items: center; justify-content: space-between;
  padding: 18px 24px;
  border-bottom: 1px solid var(--gold-xlo);
  flex-wrap: wrap; gap: 10px;
}
.prp-card-title-group { display: flex; align-items: center; gap: 10px; }
.prp-card-dot {
  width: 8px; height: 8px;
  border-radius: 50%; background: var(--gold);
  box-shadow: 0 0 10px rgba(200,168,75,0.5);
  flex-shrink: 0;
}
.prp-card-title {
  font-size: 13px; font-weight: 700;
  color: var(--tx1); letter-spacing: -0.015em;
  font-family: var(--font);
}
.prp-card-sub { font-size: 11px; color: var(--tx3); margin-top: 2px; font-family: var(--font); }
.prp-card-body { padding: 22px 24px; }

/* ── Result metrics band ── */
.result-band {
  display: grid; grid-template-columns: repeat(4,1fr);
  background: var(--gold-xxlo);
  border-top: 1px solid var(--gold-xlo);
  gap: 1px;
}
.result-cell { background: var(--s1); padding: 16px 22px; }
.result-cell-label {
  font-size: 9.5px; text-transform: uppercase; letter-spacing: 0.12em;
  color: var(--tx3); font-weight: 700; font-family: var(--font);
}
.result-cell-value {
  font-size: 24px; font-weight: 800;
  letter-spacing: -0.04em; color: var(--gold);
  margin-top: 5px; line-height: 1;
  font-family: var(--font);
}

/* ═══════════════════════════════════════════════════════
   INFO CARD  (script details)
═══════════════════════════════════════════════════════ */
.info-card {
  background: var(--s2);
  border: 1px solid var(--gold-xlo);
  border-radius: 14px;
  overflow: hidden;
  margin-top: 16px;
}
.info-card-head {
  padding: 15px 20px;
  border-bottom: 1px solid var(--gold-xlo);
  display: flex; align-items: flex-start; gap: 12px;
}
.info-card-dot {
  width: 8px; height: 8px;
  border-radius: 50%; background: var(--gold);
  box-shadow: 0 0 8px rgba(200,168,75,0.45);
  flex-shrink: 0; margin-top: 3px;
}
.info-card-title { font-size: 12px; font-weight: 700; color: var(--tx1); font-family: var(--font); }
.info-card-desc  { font-size: 11px; color: var(--tx3); margin-top: 3px; line-height: 1.55; font-family: var(--font); }
.info-card-meta  { display: grid; grid-template-columns: 1fr 1fr; background: var(--gold-xxlo); gap: 1px; }
.info-meta-cell  { background: var(--s2); padding: 13px 20px; }
.info-meta-lbl   { font-size: 9.5px; text-transform: uppercase; letter-spacing: 0.12em; color: var(--tx3); font-weight: 700; font-family: var(--font); }
.info-meta-val   { font-size: 12px; color: var(--gold); font-weight: 600; margin-top: 5px; font-family: var(--font); }

/* ═══════════════════════════════════════════════════════
   CHART AREA LABEL
═══════════════════════════════════════════════════════ */
.chart-label {
  font-size: 10px; font-weight: 600; text-transform: uppercase;
  letter-spacing: 0.1em; color: var(--tx3);
  margin-bottom: 8px; font-family: var(--font);
}

/* ═══════════════════════════════════════════════════════
   BADGES
═══════════════════════════════════════════════════════ */
.badge {
  display: inline-flex; align-items: center;
  border-radius: 6px; padding: 3px 10px;
  font-size: 10px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.06em;
  font-family: var(--font);
}
.badge-green { background: rgba(77,187,128,0.12); color: var(--green); }
.badge-red   { background: rgba(212,80,80,0.12);  color: var(--red); }
.badge-gold  { background: rgba(200,168,75,0.12); color: var(--gold); }
.badge-gray  { background: rgba(94,80,48,0.18);   color: var(--tx3); }

/* ═══════════════════════════════════════════════════════
   GROUP SELECTOR BUTTONS
═══════════════════════════════════════════════════════ */
.stButton > button {
  font-family: var(--font) !important;
  border-radius: 9px !important;
  font-size: 11px !important;
  font-weight: 600 !important;
  letter-spacing: 0.01em !important;
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

/* ── Big run CTA ── */
.run-wrap .stButton > button[kind="primary"] {
  height: 54px !important;
  font-size: 13px !important;
  font-weight: 800 !important;
  border-radius: 13px !important;
  letter-spacing: 0.04em !important;
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
[data-testid="stRadio"] [data-testid="stMarkdownContainer"] p {
  color: var(--tx2) !important;
  font-size: 12px !important;
  font-family: var(--font) !important;
  font-weight: 500 !important;
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
  color: var(--tx2) !important;
  font-size: 11.5px !important;
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
   VEGA / ALTAIR CHART
═══════════════════════════════════════════════════════ */
[data-testid="stArrowVegaLiteChart"] {
  background: var(--s1) !important;
  border: 1px solid var(--gold-xlo) !important;
  border-radius: 14px !important;
  padding: 18px !important;
}

/* ═══════════════════════════════════════════════════════
   DIVIDER
═══════════════════════════════════════════════════════ */
hr { border-color: var(--gold-xlo) !important; margin: 24px 0 !important; }

/* ═══════════════════════════════════════════════════════
   TEXT AREA (copy box)
═══════════════════════════════════════════════════════ */
[data-testid="stTextArea"] textarea {
  background: var(--s3) !important;
  border: 1px solid var(--gold-xlo) !important;
  border-radius: 10px !important;
  color: var(--tx2) !important;
  font-family: 'JetBrains Mono', 'Fira Code', monospace !important;
  font-size: 11.5px !important;
}
[data-testid="stTextArea"] textarea:focus {
  border-color: var(--gold-lo) !important;
  box-shadow: 0 0 0 2px rgba(200,168,75,0.12) !important;
}
</style>"""


# ---------------------------------------------------------------------------
# HTML helpers
# ---------------------------------------------------------------------------

def _header_html() -> str:
    return """
<div class="prp-header">
  <div class="prp-logo">
    <div class="prp-logo-mark">&#9678;</div>
    <div>
      <div class="prp-brand-name">PRP Automation Studio</div>
      <div class="prp-brand-sub">AB InBev &middot; TPRM &amp; Risk Intelligence</div>
    </div>
  </div>
  <div class="prp-header-right">
    <span class="prp-badge prp-badge-live">Live</span>
    <span class="prp-badge">8 Reports</span>
    <span class="prp-badge">AB InBev</span>
  </div>
</div>"""


def _kpi_strip_html(result: "RunResult | None" = None) -> str:
    n_outputs = len(result.outputs) if result else 0
    n_tables  = sum(len(s) for s in result.tables.values()) if result else 0

    if result and result.ok:
        status = '<span class="badge badge-green">&#9679;&nbsp; Success</span>'
    elif result:
        status = '<span class="badge badge-red">&#9679;&nbsp; Error</span>'
    else:
        status = '<span class="badge badge-gray">&mdash;&nbsp; Idle</span>'

    cells = [
        ("Total Reports",  "8",                              "Automation scripts"),
        ("Script Groups",  "7",                              "Diagram &middot; Slide 12"),
        ("Input Sheets",   "3",                              "TPRM &middot; OneTrust &middot; Risk"),
        ("Output Files",   str(n_outputs) if result else "—", "From last run"),
        ("Data Tables",    str(n_tables)  if result else "—", "Sheets in output"),
        ("Run Status",     status,                           "&nbsp;"),
    ]

    html = ""
    for lbl, val, detail in cells:
        is_badge = "<span" in str(val)
        val_class = "kpi-value kpi-value-sm" if is_badge else "kpi-value"
        html += f"""
<div class="kpi-cell">
  <div class="kpi-eyebrow">{lbl}</div>
  <div class="{val_class}">{val}</div>
  <div class="kpi-detail">{detail}</div>
</div>"""
    return f'<div class="kpi-strip">{html}</div>'


def _section_label(num: str, text: str) -> str:
    return (
        f'<div class="prp-label">'
        f'<span class="prp-label-num">{num}</span>'
        f'<span>{text}</span>'
        f'</div>'
    )


def _info_card_html(sheet: str, notes: str, group: str) -> str:
    return f"""
<div class="info-card">
  <div class="info-card-head">
    <div class="info-card-dot"></div>
    <div>
      <div class="info-card-title">{group}</div>
      <div class="info-card-desc">{notes}</div>
    </div>
  </div>
  <div class="info-card-meta">
    <div class="info-meta-cell">
      <div class="info-meta-lbl">Required Sheet</div>
      <div class="info-meta-val">{sheet}</div>
    </div>
    <div class="info-meta-cell">
      <div class="info-meta-lbl">Output Format</div>
      <div class="info-meta-val">.xlsx + Charts</div>
    </div>
  </div>
</div>"""


# ---------------------------------------------------------------------------
# Results renderer
# ---------------------------------------------------------------------------

def render_results(entry: ScriptEntry, result: RunResult, fmt: str) -> None:
    n_tables = sum(len(s) for s in result.tables.values())

    if result.ok:
        status_badge = '<span class="badge badge-green">&#10003;&nbsp; Completed</span>'
        dot_color    = "var(--green)"
    else:
        status_badge = f'<span class="badge badge-red">&#10007;&nbsp; Error &mdash; code {result.returncode}</span>'
        dot_color    = "var(--red)"

    # ── Run summary card ────────────────────────────────────────────────────
    st.markdown(f"""
<div class="prp-card" style="margin-bottom:20px;">
  <div class="prp-card-accent-bar"></div>
  <div class="prp-card-header">
    <div class="prp-card-title-group">
      <div class="prp-card-dot" style="background:{dot_color};box-shadow:0 0 10px {dot_color};"></div>
      <div>
        <div class="prp-card-title">{entry.label}</div>
        <div class="prp-card-sub">{entry.group}</div>
      </div>
    </div>
    <div>{status_badge}</div>
  </div>
  <div class="result-band">
    <div class="result-cell">
      <div class="result-cell-label">Script ID</div>
      <div class="result-cell-value">{entry.id.upper()}</div>
    </div>
    <div class="result-cell">
      <div class="result-cell-label">Output Files</div>
      <div class="result-cell-value">{len(result.outputs)}</div>
    </div>
    <div class="result-cell">
      <div class="result-cell-label">Data Tables</div>
      <div class="result-cell-value">{n_tables}</div>
    </div>
    <div class="result-cell">
      <div class="result-cell-label">Exit Code</div>
      <div class="result-cell-value" style="color:{'var(--green)' if result.ok else 'var(--red)'};">{result.returncode}</div>
    </div>
  </div>
</div>""", unsafe_allow_html=True)

    # ── Script log ──────────────────────────────────────────────────────────
    with st.expander("Script log  ·  stdout / stderr", expanded=not result.ok):
        if result.stdout.strip():
            st.code(result.stdout, language="text")
        if result.stderr.strip():
            st.markdown(
                '<div style="font-size:10px;text-transform:uppercase;letter-spacing:0.1em;'
                'color:var(--tx3);margin:10px 0 6px;font-family:var(--font);">stderr</div>',
                unsafe_allow_html=True,
            )
            st.code(result.stderr, language="text")
        if not result.stdout.strip() and not result.stderr.strip():
            st.markdown(
                '<div style="color:var(--tx3);font-size:12px;padding:4px 0;'
                'font-family:var(--font);">No console output.</div>',
                unsafe_allow_html=True,
            )

    if not result.outputs:
        st.warning("The script produced no output file.")
        return

    # ── Downloads ───────────────────────────────────────────────────────────
    st.markdown(
        '<div style="font-size:9.5px;text-transform:uppercase;letter-spacing:0.12em;'
        'color:var(--tx3);font-weight:700;margin:28px 0 12px;font-family:var(--font);">'
        'Downloads</div>',
        unsafe_allow_html=True,
    )
    dl_cols = st.columns(min(4, len(result.outputs)))
    for i, (name, data) in enumerate(result.outputs):
        mime = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if name.lower().endswith(".xlsx") else
            "image/png" if name.lower().endswith(".png") else
            "application/octet-stream"
        )
        icon = "⬇  " + name
        dl_cols[i % len(dl_cols)].download_button(
            icon, data=data, file_name=name, mime=mime,
            key=f"dl_{entry.id}_{i}", use_container_width=True,
        )

    # ── Chart preview ───────────────────────────────────────────────────────
    png_outputs = [(n, d) for n, d in result.outputs if n.lower().endswith(".png")]
    rendered_any = False

    if png_outputs:
        st.markdown(
            '<div style="font-size:9.5px;text-transform:uppercase;letter-spacing:0.12em;'
            'color:var(--tx3);font-weight:700;margin:28px 0 12px;font-family:var(--font);">'
            'Chart Preview</div>',
            unsafe_allow_html=True,
        )
        img_cols = st.columns(len(png_outputs))
        for col, (name, data) in zip(img_cols, png_outputs):
            with col:
                st.image(data, caption=name, use_container_width=True)
        rendered_any = True

    chart_items = []
    for fname, sheets in result.tables.items():
        for sheet_name, df in sheets.items():
            cdf = _chartable(df)
            if cdf is not None and len(cdf.columns) <= 6:
                chart_items.append((fname, sheet_name, cdf))

    if chart_items:
        st.markdown(
            '<div style="font-size:9.5px;text-transform:uppercase;letter-spacing:0.12em;'
            'color:var(--tx3);font-weight:700;margin:28px 0 12px;font-family:var(--font);">'
            'Data Charts</div>',
            unsafe_allow_html=True,
        )
        n_cols = min(2, len(chart_items))
        chart_cols = st.columns(n_cols) if n_cols > 1 else [st]
        for idx, (fname, sheet_name, cdf) in enumerate(chart_items):
            with chart_cols[idx % n_cols]:
                st.markdown(
                    f'<div class="chart-label">{sheet_name}</div>',
                    unsafe_allow_html=True,
                )
                try:
                    st.bar_chart(cdf, stack=True, use_container_width=True)
                    rendered_any = True
                except Exception:
                    pass

    if not rendered_any and not png_outputs:
        st.markdown(
            '<div style="color:var(--tx3);font-size:12px;padding:10px 0;'
            'font-family:var(--font);">No chartable data detected in output.</div>',
            unsafe_allow_html=True,
        )
    elif chart_items and not png_outputs:
        st.markdown(
            '<div style="color:var(--tx3);font-size:10.5px;margin-top:6px;'
            'font-family:var(--font);">Re-rendered from output data — '
            'downloaded .xlsx contains the original styled charts.</div>',
            unsafe_allow_html=True,
        )

    # ── Output tables ────────────────────────────────────────────────────────
    st.markdown(
        '<div style="font-size:9.5px;text-transform:uppercase;letter-spacing:0.12em;'
        'color:var(--tx3);font-weight:700;margin:28px 0 12px;font-family:var(--font);">'
        'Output Tables</div>',
        unsafe_allow_html=True,
    )
    for fname, sheets in result.tables.items():
        for sheet_name, df in sheets.items():
            df = _trim_sparse_rows(df)
            df = _trim_sparse_cols(df)
            st.markdown(f"""
<div class="prp-card" style="margin-bottom:16px;">
  <div class="prp-card-header">
    <div class="prp-card-title-group">
      <div class="prp-card-dot"></div>
      <div>
        <div class="prp-card-title">{sheet_name}</div>
        <div class="prp-card-sub">{fname} &nbsp;&middot;&nbsp; {len(df)} rows &times; {len(df.columns)} cols</div>
      </div>
    </div>
  </div>
  <div class="prp-card-body">""", unsafe_allow_html=True)

            st.dataframe(df, use_container_width=True, hide_index=True)

            safe_key = f"cp_{entry.id}_{fname}_{sheet_name}".replace(" ", "_")
            with st.expander(f"Copy as text  ·  {sheet_name}"):
                if fmt == "Markdown":
                    try:
                        text = df.to_markdown(index=False)
                    except Exception:
                        text = df.to_csv(sep="\t", index=False)
                    st.text_area("", text, height=200, key=safe_key + "_md",
                                 label_visibility="collapsed")
                else:
                    tsv = _clean_for_tsv(df).to_csv(sep="\t", index=False, lineterminator="\r\n")
                    st.caption(
                        "Click inside · Ctrl+A to select all · Ctrl+C to copy · "
                        "paste into Excel — columns split automatically."
                    )
                    st.text_area("", tsv, height=200, key=safe_key,
                                 label_visibility="collapsed")

            st.markdown("</div></div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Main UI
# ---------------------------------------------------------------------------

def main() -> None:
    st.set_page_config(
        page_title="PRP Automation Studio — AB InBev",
        page_icon="⬡",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    st.markdown(_CSS, unsafe_allow_html=True)
    st.markdown(_header_html(), unsafe_allow_html=True)

    last_result: "RunResult | None" = st.session_state.get("last_result")

    # ── KPI strip ─────────────────────────────────────────────────────────
    st.markdown('<div class="prp-wrap">', unsafe_allow_html=True)
    st.markdown(_kpi_strip_html(last_result), unsafe_allow_html=True)

    # ── Input + Selector (2 columns) ──────────────────────────────────────
    col_up, col_sel = st.columns([5, 7], gap="large")

    with col_up:
        st.markdown(_section_label("1", "Upload Workbook"), unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "Workbook",
            type=["xlsx"],
            label_visibility="collapsed",
        )
        if uploaded:
            st.markdown(
                f'<div style="margin-top:10px;padding:12px 16px;background:rgba(77,187,128,0.07);'
                f'border:1px solid rgba(77,187,128,0.2);border-radius:10px;">'
                f'<span style="color:var(--green);font-size:11px;font-weight:700;font-family:var(--font);">'
                f'&#10003;&nbsp; Ready</span>'
                f'<span style="color:var(--tx1);font-size:12px;font-weight:600;font-family:var(--font);margin-left:10px;">'
                f'{uploaded.name}</span>'
                f'<span style="color:var(--tx3);font-size:11px;font-family:var(--font);margin-left:8px;">'
                f'{uploaded.size/1024:.1f} KB</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<div style="margin-top:10px;font-size:11.5px;color:var(--tx3);'
                'font-family:var(--font);">Accepts .xlsx &middot; '
                'the same workbook your scripts expect</div>',
                unsafe_allow_html=True,
            )

    with col_sel:
        st.markdown(_section_label("2", "Select Report"), unsafe_allow_html=True)

        groups: dict[str, list[ScriptEntry]] = {}
        for e in REGISTRY:
            groups.setdefault(e.group, []).append(e)
        group_names = list(groups.keys())

        _short = {
            group_names[0]: "Diagram 1",
            group_names[1]: "Diagram 2",
            group_names[2]: "Diagram 3",
            group_names[3]: "Slide 12 · 1st",
            group_names[4]: "Slide 12 · 2nd",
            group_names[5]: "Slide 12 · 3rd",
            group_names[6]: "Diagram 4",
        }

        if "active_group" not in st.session_state:
            st.session_state["active_group"] = group_names[0]

        tab_cols = st.columns(len(group_names))
        for col, gname in zip(tab_cols, group_names):
            with col:
                is_active = st.session_state["active_group"] == gname
                if st.button(
                    _short.get(gname, gname),
                    key=f"grp_{gname}",
                    type="primary" if is_active else "secondary",
                    use_container_width=True,
                ):
                    st.session_state["active_group"] = gname
                    st.rerun()

        st.markdown('<div style="margin-top:14px;"></div>', unsafe_allow_html=True)

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

    # ── Divider ───────────────────────────────────────────────────────────
    st.markdown(
        '<hr style="border-color:var(--gold-xlo);margin:28px 0;">',
        unsafe_allow_html=True,
    )

    # ── Options & Run ─────────────────────────────────────────────────────
    col_opt, col_run = st.columns([6, 3], gap="large")

    with col_opt:
        st.markdown(_section_label("3", "Options"), unsafe_allow_html=True)
        fmt_raw = st.radio(
            "Copy format",
            ["TSV (Excel / Sheets)", "Markdown"],
            horizontal=True,
        )
        fmt = "Markdown" if "Markdown" in fmt_raw else "TSV"
        st.markdown(
            '<div style="margin-top:8px;font-size:11px;color:var(--tx3);font-family:var(--font);">'
            'TSV — paste directly into Excel or Google Sheets with columns split automatically.'
            '</div>',
            unsafe_allow_html=True,
        )

    with col_run:
        st.markdown(_section_label("4", "Execute"), unsafe_allow_html=True)
        st.markdown('<div class="run-wrap">', unsafe_allow_html=True)
        run = st.button(
            "▶  Run Report",
            type="primary",
            disabled=uploaded is None,
            use_container_width=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)
        if uploaded is None:
            st.markdown(
                '<div style="font-size:10.5px;color:var(--tx3);text-align:center;'
                'margin-top:8px;font-family:var(--font);">Upload a workbook first</div>',
                unsafe_allow_html=True,
            )

    # ── How it works ──────────────────────────────────────────────────────
    st.markdown('<div style="margin-top:20px;"></div>', unsafe_allow_html=True)
    with st.expander("How it works"):
        st.markdown(
            "1. **Upload** your `.xlsx` workbook.\n"
            "2. **Select** a report group, then pick a variant.\n"
            "3. **Run** — the script executes in an isolated temp folder. "
            "You get the log, chart preview, output tables with copy, "
            "and the exact `.xlsx` the script generates.\n\n"
            "> Each report requires a specific sheet — shown in the script info card."
        )

    # ── Execute ───────────────────────────────────────────────────────────
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
        st.session_state["last_result"]   = result
        st.session_state["last_entry_id"] = entry.id
        st.session_state["last_fmt"]      = fmt
        st.rerun()

    # ── Render persisted results ───────────────────────────────────────────
    if "last_result" in st.session_state:
        st.markdown(
            '<hr style="border-color:var(--gold-xlo);margin:32px 0 28px;">',
            unsafe_allow_html=True,
        )
        st.markdown(_section_label("↓", "Results"), unsafe_allow_html=True)
        _entry = REGISTRY_BY_ID[st.session_state["last_entry_id"]]
        render_results(
            _entry,
            st.session_state["last_result"],
            st.session_state["last_fmt"],
        )

    st.markdown('</div>', unsafe_allow_html=True)  # /prp-wrap

    # Bottom spacer
    st.markdown('<div style="height:60px;"></div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
