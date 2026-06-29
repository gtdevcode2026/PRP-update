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
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=Instrument+Sans:wght@400;500;600;700&display=swap');

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
  --font:         'Instrument Sans', system-ui, -apple-system, sans-serif;
  --font-display: 'Syne', system-ui, sans-serif;
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

/* Hide Streamlit chrome — collapse to zero so it takes no space */
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

/* Belt-and-suspenders top-gap kill */
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
   LAYOUT WRAPPER
═══════════════════════════════════════════════════════ */
.prp-wrap { padding: 0 44px; }

/* ═══════════════════════════════════════════════════════
   KPI STRIP
═══════════════════════════════════════════════════════ */
.kpi-strip {
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  background: var(--gold-xxlo);
  border: 1px solid var(--gold-xlo);
  border-radius: 16px;
  overflow: hidden;
  margin: 16px 0 18px;
  gap: 1px;
}
@media(max-width:1100px){ .kpi-strip { grid-template-columns: repeat(3,1fr); } }
@media(max-width:700px)  { .kpi-strip { grid-template-columns: repeat(2,1fr); } }

.kpi-cell {
  background: var(--s1);
  padding: 16px 22px;
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
  font-size: 9px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.14em;
  color: var(--tx3); margin-bottom: 11px;
  font-family: var(--font);
}
.kpi-value {
  font-size: 28px; font-weight: 800;
  line-height: 1; letter-spacing: -0.03em;
  color: var(--gold);
  font-family: var(--font-display);
}
.kpi-value-sm { font-size: 15px; letter-spacing: -0.01em; font-weight: 700; font-family: var(--font) !important; }
.kpi-detail { font-size: 10px; color: var(--tx3); margin-top: 8px; font-family: var(--font); letter-spacing: 0.01em; }

/* ═══════════════════════════════════════════════════════
   SECTION LABEL
═══════════════════════════════════════════════════════ */
.prp-label {
  display: flex; align-items: center; gap: 10px;
  font-size: 9px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.16em;
  color: var(--tx3); margin: 4px 0 6px;
  font-family: var(--font);
}
.info-card { margin-top: 8px; }
.prp-label .prp-label-num {
  display: inline-flex; align-items: center; justify-content: center;
  width: 19px; height: 19px;
  background: var(--gold-xlo); border: 1px solid var(--gold-lo);
  border-radius: 5px; font-size: 9px; font-weight: 800;
  color: var(--gold-lo);
  font-family: var(--font-display);
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
  font-size: 13.5px; font-weight: 700;
  color: var(--tx1); letter-spacing: -0.01em;
  font-family: var(--font-display);
}
.prp-card-sub { font-size: 11px; color: var(--tx3); margin-top: 3px; font-family: var(--font); letter-spacing: 0.01em; }
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
  font-size: 9px; text-transform: uppercase; letter-spacing: 0.14em;
  color: var(--tx3); font-weight: 700; font-family: var(--font);
}
.result-cell-value {
  font-size: 26px; font-weight: 800;
  letter-spacing: -0.02em; color: var(--gold);
  margin-top: 6px; line-height: 1;
  font-family: var(--font-display);
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
.info-card-title { font-size: 13px; font-weight: 700; color: var(--tx1); font-family: var(--font-display); letter-spacing: -0.01em; }
.info-card-desc  { font-size: 11.5px; color: var(--tx3); margin-top: 4px; line-height: 1.6; font-family: var(--font); }
.info-card-meta  { display: grid; grid-template-columns: 1fr 1fr; background: var(--gold-xxlo); gap: 1px; }
.info-meta-cell  { background: var(--s2); padding: 13px 20px; }
.info-meta-lbl   { font-size: 9px; text-transform: uppercase; letter-spacing: 0.14em; color: var(--tx3); font-weight: 700; font-family: var(--font); }
.info-meta-val   { font-size: 12.5px; color: var(--gold); font-weight: 700; margin-top: 5px; font-family: var(--font); }

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
  font-size: 9.5px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.08em;
  font-family: var(--font);
}
.badge-green { background: rgba(77,187,128,0.12); color: var(--green); }
.badge-red   { background: rgba(212,80,80,0.12);  color: var(--red); }
.badge-gold  { background: rgba(200,168,75,0.12); color: var(--gold); }
.badge-gray  { background: rgba(94,80,48,0.18);   color: var(--tx3); }

/* ═══════════════════════════════════════════════════════
   LEFT NAV PANEL
═══════════════════════════════════════════════════════ */
.nav-panel {
  background: var(--s1);
  border: 1px solid var(--gold-xlo);
  border-radius: 16px;
  overflow: hidden;
  position: sticky;
  top: 80px;
}
.nav-panel-head {
  padding: 14px 18px;
  border-bottom: 1px solid var(--gold-xlo);
  font-size: 8.5px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.18em;
  color: var(--tx3); font-family: var(--font);
}
.nav-item {
  display: flex; align-items: center; gap: 11px;
  padding: 12px 18px;
  border-bottom: 1px solid var(--gold-xlo);
  cursor: pointer;
  transition: background 0.15s cubic-bezier(0.32,0.72,0,1);
  text-decoration: none;
}
.nav-item:last-child { border-bottom: none; }
.nav-item:hover { background: var(--s2); }
.nav-item.active {
  background: var(--gold-xlo);
  border-left: 2px solid var(--gold);
}
.nav-item-num {
  width: 22px; height: 22px; flex-shrink: 0;
  border-radius: 6px;
  background: var(--gold-xlo);
  border: 1px solid var(--gold-lo);
  display: flex; align-items: center; justify-content: center;
  font-size: 10px; font-weight: 800; color: var(--gold-lo);
  font-family: var(--font-display);
}
.nav-item.active .nav-item-num {
  background: var(--gold); color: #000; border-color: var(--gold);
}
.nav-item-text {
  font-size: 12px; font-weight: 600;
  color: var(--tx2); font-family: var(--font);
  line-height: 1.3; letter-spacing: -0.005em;
}
.nav-item.active .nav-item-text { color: var(--tx1); font-weight: 700; }

/* ═══════════════════════════════════════════════════════
   BUTTONS (run + generic)
═══════════════════════════════════════════════════════ */
.stButton > button {
  font-family: var(--font) !important;
  border-radius: 9px !important;
  font-size: 11.5px !important;
  font-weight: 600 !important;
  letter-spacing: 0.02em !important;
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
  font-weight: 700 !important;
  border-radius: 13px !important;
  letter-spacing: 0.06em !important;
  font-family: var(--font) !important;
  text-transform: uppercase !important;
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
   CODE BLOCK  (used for aligned table display)
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
/* Column headers in the code block get gold colour */
[data-testid="stCode"] pre > code > span:first-child {
  color: var(--gold) !important;
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
@keyframes pulseGold {
  0%, 100% { box-shadow: 0 0 0 0 rgba(200,168,75,0); }
  50%       { box-shadow: 0 0 0 6px rgba(200,168,75,0.18); }
}
@keyframes scanLine {
  0%   { left: -60%; }
  100% { left: 110%; }
}
@keyframes fadeUp {
  from { opacity: 0; transform: translateY(32px); }
  to   { opacity: 1; transform: translateY(0); }
}

.anim-slide-left { animation: slideInLeft 0.38s cubic-bezier(0.32,0.72,0,1) both; }
.anim-slide-up   { animation: slideUp     0.42s cubic-bezier(0.32,0.72,0,1) both; }
.anim-fade-up    { animation: fadeUp      0.5s  cubic-bezier(0.32,0.72,0,1) both; }

/* stagger children */
.anim-stagger > *:nth-child(1) { animation-delay: 0ms; }
.anim-stagger > *:nth-child(2) { animation-delay: 60ms; }
.anim-stagger > *:nth-child(3) { animation-delay: 120ms; }
.anim-stagger > *:nth-child(4) { animation-delay: 180ms; }

/* file-ready scan shimmer */
.file-ready {
  position: relative; overflow: hidden;
  margin-top: 8px; padding: 10px 16px;
  background: rgba(77,187,128,0.07);
  border: 1px solid rgba(77,187,128,0.22);
  border-radius: 10px;
  display: flex; align-items: center; gap: 10px;
  animation: slideInLeft 0.38s cubic-bezier(0.32,0.72,0,1) both;
}
.file-ready::after {
  content: '';
  position: absolute; top: 0; bottom: 0; width: 60%;
  background: linear-gradient(90deg, transparent, rgba(255,255,255,0.04), transparent);
  animation: scanLine 1.1s cubic-bezier(0.4,0,0.6,1) 0.38s 1 both;
  pointer-events: none;
}
.file-ready-tick  { color: var(--green); font-size: 11px; font-weight: 700; font-family: var(--font); }
.file-ready-name  { color: var(--tx1);   font-size: 12px; font-weight: 600; font-family: var(--font); }
.file-ready-size  { color: var(--tx3);   font-size: 11px; font-family: var(--font); margin-left: auto; }

/* results reveal wrapper */
.results-reveal { animation: fadeUp 0.52s cubic-bezier(0.32,0.72,0,1) both; }

/* thumbnail chart grid */
.chart-thumb-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
  gap: 12px;
  margin: 10px 0 0;
}
.chart-thumb {
  background: var(--s2);
  border: 1px solid var(--gold-xlo);
  border-radius: 12px;
  overflow: hidden;
  animation: slideUp 0.42s cubic-bezier(0.32,0.72,0,1) both;
}
.chart-thumb-label {
  font-size: 9.5px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.1em;
  color: var(--tx3); padding: 10px 12px 6px;
  font-family: var(--font);
  border-bottom: 1px solid var(--gold-xlo);
}
.chart-thumb img { display: block; width: 100%; height: auto; max-height: 160px; object-fit: contain; }

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

/* ═══════════════════════════════════════════════════════
   HIDE NAV DUPLICATE BUTTONS
   Nav buttons are isolated in col_nav (first column).
   Run button lives in col_input (second column).
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
   LEFT CONFIG PANEL — sticky scroll
═══════════════════════════════════════════════════════ */
.left-config-panel {
  position: sticky;
  top: 80px;
  max-height: calc(100vh - 96px);
  overflow-y: auto;
  overflow-x: hidden;
  padding-right: 4px;
  scrollbar-width: thin;
  scrollbar-color: var(--gold-xlo) transparent;
}
.left-config-panel::-webkit-scrollbar { width: 3px; }
.left-config-panel::-webkit-scrollbar-track { background: transparent; }
.left-config-panel::-webkit-scrollbar-thumb { background: var(--gold-xlo); border-radius: 3px; }

/* ═══════════════════════════════════════════════════════
   RIGHT OUTPUT PANEL
═══════════════════════════════════════════════════════ */
.right-output-panel {
  min-height: 500px;
}

/* ═══════════════════════════════════════════════════════
   EMPTY STATE
═══════════════════════════════════════════════════════ */
.empty-state {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  min-height: 420px;
  border: 1.5px dashed var(--gold-xlo);
  border-radius: 20px;
  margin-top: 16px;
  padding: 48px 32px;
  background: radial-gradient(ellipse at 50% 80%, rgba(200,168,75,0.03) 0%, transparent 65%);
}
.empty-state-icon {
  font-size: 42px;
  margin-bottom: 20px;
  opacity: 0.2;
  color: var(--gold);
  line-height: 1;
}
.empty-state-title {
  font-size: 16px;
  font-weight: 700;
  color: var(--tx3);
  font-family: var(--font-display);
  letter-spacing: -0.01em;
  margin-bottom: 10px;
}
.empty-state-desc {
  font-size: 12px;
  color: var(--tx3);
  font-family: var(--font);
  text-align: center;
  max-width: 300px;
  line-height: 1.7;
}

/* ═══════════════════════════════════════════════════════
   DOWNLOAD STRIP (top of right panel)
═══════════════════════════════════════════════════════ */
.dl-strip-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 12px;
}
.dl-strip-title {
  font-size: 9.5px;
  text-transform: uppercase;
  letter-spacing: 0.12em;
  color: var(--tx3);
  font-weight: 700;
  font-family: var(--font);
}

/* ═══════════════════════════════════════════════════════
   INPUT FIELD LABEL
═══════════════════════════════════════════════════════ */
.input-section-box {
  background: var(--s2);
  border: 1px solid var(--gold-xlo);
  border-radius: 14px;
  padding: 16px 18px 18px;
  margin-bottom: 4px;
}
.input-section-label {
  font-size: 9px;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: 0.14em;
  color: var(--tx3);
  font-family: var(--font);
  margin-bottom: 12px;
  display: flex;
  align-items: center;
  gap: 8px;
}
.input-section-label::before {
  content: '';
  display: inline-block;
  width: 4px;
  height: 4px;
  border-radius: 50%;
  background: var(--gold);
  box-shadow: 0 0 6px var(--gold);
}

/* ═══════════════════════════════════════════════════════
   NAV PANEL — active glow + depth
═══════════════════════════════════════════════════════ */
.nav-item.active {
  background: linear-gradient(90deg, rgba(200,168,75,0.12), rgba(200,168,75,0.04)) !important;
  border-left: 2px solid var(--gold) !important;
  box-shadow: inset 0 0 16px rgba(200,168,75,0.05) !important;
}
.nav-item.active .nav-item-num {
  background: var(--gold) !important;
  color: #000 !important;
  border-color: var(--gold) !important;
  box-shadow: 0 0 8px rgba(200,168,75,0.4) !important;
}
.nav-item.active .nav-item-text {
  color: var(--tx1) !important;
  font-weight: 700 !important;
}
.nav-item:hover:not(.active) {
  background: rgba(200,168,75,0.03) !important;
  border-left: 1px solid rgba(200,168,75,0.12) !important;
}

/* ═══════════════════════════════════════════════════════
   RADIO — premium selectable cards
═══════════════════════════════════════════════════════ */
[data-baseweb="radio-group"] {
  flex-direction: column !important;
  gap: 6px !important;
}
[data-baseweb="radio-group"] [data-baseweb="radio"] {
  background: var(--s2) !important;
  border: 1px solid var(--gold-xlo) !important;
  border-radius: 10px !important;
  padding: 10px 14px !important;
  gap: 10px !important;
  align-items: center !important;
  transition: border-color 0.18s cubic-bezier(0.32,0.72,0,1),
              background   0.18s cubic-bezier(0.32,0.72,0,1) !important;
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
  width: 14px !important;
  height: 14px !important;
}
[data-baseweb="radio-group"] [role="radio"][aria-checked="true"] {
  background: var(--gold) !important;
  border-color: var(--gold) !important;
  box-shadow: 0 0 6px rgba(200,168,75,0.4) !important;
}
[data-baseweb="radio-group"] [data-testid="stMarkdownContainer"] p {
  font-size: 12px !important;
  font-weight: 600 !important;
  color: var(--tx2) !important;
  font-family: var(--font) !important;
  line-height: 1.4 !important;
  margin: 0 !important;
  letter-spacing: -0.005em !important;
}
[data-baseweb="radio-group"] [data-baseweb="radio"]:has([aria-checked="true"])
[data-testid="stMarkdownContainer"] p {
  color: var(--tx1) !important;
  font-weight: 700 !important;
}

/* ═══════════════════════════════════════════════════════
   RADIO LABEL (the "Copy format" header text)
═══════════════════════════════════════════════════════ */
[data-testid="stRadio"] > label > div > p {
  font-size: 9px !important;
  font-weight: 800 !important;
  text-transform: uppercase !important;
  letter-spacing: 0.14em !important;
  color: var(--tx3) !important;
  font-family: var(--font) !important;
  margin-bottom: 8px !important;
}

/* ═══════════════════════════════════════════════════════
   FILE UPLOADER — more dramatic hover
═══════════════════════════════════════════════════════ */
[data-testid="stFileUploader"]:hover {
  border-color: var(--gold) !important;
  box-shadow: 0 0 0 1px rgba(200,168,75,0.08),
              0 0 28px rgba(200,168,75,0.06) !important;
}

/* ═══════════════════════════════════════════════════════
   SECTION LABEL number badge — more prominent
═══════════════════════════════════════════════════════ */
.prp-label .prp-label-num {
  background: var(--gold-xlo) !important;
  border: 1px solid var(--gold-lo) !important;
  color: var(--gold) !important;
  font-weight: 900 !important;
}

/* top-gap handled in chrome-hide block above */

/* ═══════════════════════════════════════════════════════
   STATUS BAR  (replaces KPI strip)
═══════════════════════════════════════════════════════ */
.status-bar {
  display: flex; align-items: center; gap: 10px;
  padding: 7px 44px;
  background: var(--gold-xxlo);
  border-bottom: 1px solid var(--gold-xlo);
  font-family: var(--font); font-size: 11px;
  letter-spacing: 0.02em;
}
.status-sep { color: var(--gold-xlo); }
.status-item { color: var(--tx3); }

/* ═══════════════════════════════════════════════════════
   MAIN LAYOUT  (content area padding)
═══════════════════════════════════════════════════════ */
.main-layout { padding: 6px 32px 0; }
</style>"""


# ---------------------------------------------------------------------------
# Nav JS wiring (runs in components.html iframe — same origin → parent access)
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

def _header_html() -> str:
    return """
<div class="prp-header">
  <div class="prp-logo">
    <div class="prp-logo-mark">&#9678;</div>
    <div>
      <div class="prp-brand-name">PRP Automation Dashboard</div>
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
        status_html = '<span class="badge badge-green" style="font-size:11px;">&#9679;&nbsp; Success</span>'
        status_detail = "Last run completed"
    elif result:
        status_html = '<span class="badge badge-red" style="font-size:11px;">&#9679;&nbsp; Error</span>'
        status_detail = f"Exit code {result.returncode}"
    else:
        status_html = '<span class="badge badge-gray" style="font-size:11px;">&mdash;&nbsp; Ready</span>'
        status_detail = "Upload &amp; run a report"

    cells = [
        ("Reports",      "8",                                       "Automation scripts"),
        ("Input Sheets", "3",                                       "TPRM &middot; OneTrust &middot; Risk"),
        ("Output Files", str(n_outputs) if result else "&mdash;",  "From last run"),
        ("Data Tables",  str(n_tables)  if result else "&mdash;",  "Sheets in output"),
        ("Run Status",   status_html,                               status_detail),
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
<div class="info-card anim-slide-up">
  <div class="info-card-head">
    <div class="info-card-dot"></div>
    <div style="flex:1;min-width:0;">
      <div class="info-card-title">{group}</div>
      <div class="info-card-desc">{notes}</div>
    </div>
    <span class="badge badge-gold" style="flex-shrink:0;align-self:flex-start;">AB InBev</span>
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
# Live chart builder  (Plotly — no external CDN, already installed with Streamlit)
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
    font_cfg   = dict(family="'Instrument Sans', system-ui, sans-serif", color=tick_color, size=11)

    fig.update_layout(
        title=dict(
            text=title,
            font=dict(family="'Syne', system-ui", color="#A89060", size=11),
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
            font=dict(family="'Instrument Sans', system-ui", color="#A89060", size=10),
            bgcolor="rgba(0,0,0,0)",
            bordercolor="rgba(0,0,0,0)",
        ),
        xaxis=dict(
            gridcolor=grid_color, linecolor="#2A1F08", tickcolor="#2A1F08",
            tickfont=dict(family="'Instrument Sans', system-ui", color=tick_color, size=10),
            tickangle=-25,
        ),
        yaxis=dict(
            gridcolor=grid_color, linecolor="#2A1F08", tickcolor="#2A1F08",
            tickfont=dict(family="'Instrument Sans', system-ui", color=tick_color, size=10),
            zeroline=False,
        ),
        hoverlabel=dict(
            bgcolor="#191919",
            bordercolor="#221B0B",
            font=dict(family="'Instrument Sans', system-ui", color="#F0E4C0", size=12),
        ),
        # Animate bars growing in from zero on first render
        transition=dict(duration=700, easing="cubic-in-out"),
    )
    return fig


# ---------------------------------------------------------------------------
# Results renderer
# ---------------------------------------------------------------------------

def render_results(entry: ScriptEntry, result: RunResult, fmt: str) -> None:
    st.markdown('<div class="results-reveal">', unsafe_allow_html=True)

    if not result.outputs:
        st.warning("The script produced no output file.")
        st.markdown('</div>', unsafe_allow_html=True)
        return

    # ── 1. Download Options ──────────────────────────────────────────────────
    st.markdown(f"""
<div style="display:flex;align-items:center;justify-content:space-between;
  padding:10px 0 10px;margin-bottom:10px;
  border-bottom:1px solid var(--gold-xlo);">
  <div style="display:flex;align-items:center;gap:8px;">
    <div style="width:7px;height:7px;border-radius:50%;background:var(--gold);
      box-shadow:0 0 8px rgba(200,168,75,0.5);"></div>
    <span style="font-size:15px;font-weight:700;color:var(--tx1);
      font-family:var(--font-display);letter-spacing:0.06em;text-transform:uppercase;">
      Download Options</span>
  </div>
  <span class="badge badge-gold">{len(result.outputs)} file{'s' if len(result.outputs) != 1 else ''}</span>
</div>""", unsafe_allow_html=True)

    dl_cols = st.columns(min(3, len(result.outputs)))
    for i, (name, data) in enumerate(result.outputs):
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

    st.markdown('<div style="margin-bottom:28px;"></div>', unsafe_allow_html=True)

    # ── Pre-process sheets ───────────────────────────────────────────────────
    processed_sheets: list[tuple[str, str, pd.DataFrame]] = []
    for fname, sheets in result.tables.items():
        for sheet_name, df in sheets.items():
            processed_sheets.append((fname, sheet_name, _trim_sparse_cols(_trim_sparse_rows(df))))

    # ── 2. Charts ────────────────────────────────────────────────────────────
    charts_rendered = 0
    for fname, sheet_name, df in processed_sheets:
        chart_df = _chartable(df)
        if chart_df is not None:
            if charts_rendered == 0:
                st.markdown(
                    '<div style="font-size:9.5px;text-transform:uppercase;letter-spacing:0.12em;'
                    'color:var(--tx3);font-weight:700;margin:0 0 14px;font-family:var(--font);">'
                    'Charts</div>',
                    unsafe_allow_html=True,
                )
            st.plotly_chart(
                _build_plotly_bar(chart_df, sheet_name),
                use_container_width=True,
                config={"displayModeBar": False},
            )
            charts_rendered += 1

    # ── 3. Output Preview ────────────────────────────────────────────────────
    if processed_sheets:
        st.markdown(
            '<div style="font-size:9.5px;text-transform:uppercase;letter-spacing:0.12em;'
            'color:var(--tx3);font-weight:700;margin:18px 0 14px;font-family:var(--font);">'
            'Output Preview</div>',
            unsafe_allow_html=True,
        )
        for fname, sheet_name, df in processed_sheets:
            row_count = len(df)
            col_count = len(df.columns)
            st.markdown(f"""
<div class="prp-card" style="margin-bottom:14px;">
  <div class="prp-card-header">
    <div class="prp-card-title-group">
      <div class="prp-card-dot"></div>
      <div>
        <div class="prp-card-title">{sheet_name}</div>
        <div class="prp-card-sub">{fname} &nbsp;&middot;&nbsp; {row_count} rows &times; {col_count} cols</div>
      </div>
    </div>
  </div>
  <div class="prp-card-body">""", unsafe_allow_html=True)

            st.dataframe(
                df.head(50),
                use_container_width=True,
                hide_index=True,
                height=min(300, 38 + len(df.head(50)) * 35),
            )
            if row_count > 50:
                st.markdown(
                    f'<div style="font-size:10.5px;color:var(--tx3);margin-top:6px;'
                    f'font-family:var(--font);">Showing 50 of {row_count} rows — '
                    f'download the .xlsx for the full dataset.</div>',
                    unsafe_allow_html=True,
                )

            safe_key = f"cp_{entry.id}_{fname}_{sheet_name}".replace(" ", "_")
            with st.expander(f"Copy as text  ·  {sheet_name}"):
                clean_df = _clean_for_tsv(df)
                if fmt == "Markdown":
                    try:
                        display_text = clean_df.to_markdown(index=False)
                    except Exception:
                        display_text = clean_df.to_string(index=False)
                    st.code(display_text, language=None)
                else:
                    try:
                        display_text = clean_df.to_string(index=False)
                    except Exception:
                        display_text = clean_df.to_csv(sep="\t", index=False)
                    st.code(display_text, language=None)
                    st.markdown(
                        '<div style="font-size:9px;text-transform:uppercase;letter-spacing:0.12em;'
                        'color:var(--tx3);font-family:var(--font);margin:10px 0 4px;">'
                        'Raw TSV &mdash; Ctrl+A &rarr; Ctrl+C &rarr; paste into Excel</div>',
                        unsafe_allow_html=True,
                    )
                    tsv = clean_df.to_csv(sep="\t", index=False, lineterminator="\r\n")
                    st.text_area("", tsv, height=90, key=safe_key,
                                 label_visibility="collapsed")

            st.markdown("</div></div>", unsafe_allow_html=True)

    # ── 4. Script log ────────────────────────────────────────────────────────
    st.markdown('<div style="margin-top:24px;"></div>', unsafe_allow_html=True)
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

    st.markdown('</div>', unsafe_allow_html=True)  # /results-reveal


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

    last_result: "RunResult | None" = st.session_state.get("last_result")

    groups: dict[str, list[ScriptEntry]] = {}
    for e in REGISTRY:
        groups.setdefault(e.group, []).append(e)
    group_names = list(groups.keys())

    _short = {
        group_names[0]: "Diagram 1",
        group_names[1]: "Diagram 2",
        group_names[2]: "Diagram 3",
        group_names[3]: "Diagram 4",
        group_names[4]: "Slide 12 · 1st",
        group_names[5]: "Slide 12 · 2nd",
        group_names[6]: "Slide 12 · 3rd",
    }

    if "active_group" not in st.session_state:
        st.session_state["active_group"] = group_names[0]

    # ── 3-column layout: nav | inputs | output ─────────────────────────────
    st.markdown('<div class="main-layout">', unsafe_allow_html=True)
    col_nav, col_input, col_output = st.columns([2, 3, 7], gap="small")

    # ── NAV column ─────────────────────────────────────────────────────────
    with col_nav:
        st.markdown(_section_label("1", "Report"), unsafe_allow_html=True)
        nav_html = '<div class="nav-panel"><div class="nav-panel-head">Type</div>'
        for idx, gname in enumerate(group_names, 1):
            is_active = st.session_state["active_group"] == gname
            active_cls = " active" if is_active else ""
            nav_html += (
                f'<div class="nav-item{active_cls}">'
                f'<div class="nav-item-num">{idx}</div>'
                f'<div class="nav-item-text">{_short.get(gname, gname)}</div>'
                f'</div>'
            )
        nav_html += '</div>'
        st.markdown(nav_html, unsafe_allow_html=True)

        # Hidden Streamlit buttons — CSS hides, JS wires .nav-item clicks to them
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

    # ── INPUT column ────────────────────────────────────────────────────────
    with col_input:
        st.markdown(_section_label("2", "Input File"), unsafe_allow_html=True)
        st.markdown('<div class="input-section-box">', unsafe_allow_html=True)
        st.markdown(
            '<div class="input-section-label">Upload Excel Workbook (.xlsx)</div>',
            unsafe_allow_html=True,
        )
        uploaded = st.file_uploader(
            "Workbook",
            type=["xlsx"],
            label_visibility="collapsed",
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
                f'<span class="file-ready-tick">&#10003;&nbsp; Ready</span>'
                f'<span class="file-ready-name">{file_name}</span>'
                f'<span class="file-ready-size">{file_size/1024:.1f} KB</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
        st.markdown('</div>', unsafe_allow_html=True)

        active_group = st.session_state["active_group"]
        g_entries = groups[active_group]
        labels = [e.label for e in g_entries]

        st.markdown(_section_label("3", "Variant"), unsafe_allow_html=True)
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

        st.markdown('<hr style="border-color:var(--gold-xlo);margin:10px 0;">', unsafe_allow_html=True)
        fmt_raw = st.radio(
            "Copy format",
            ["TSV (Excel / Sheets)", "Markdown"],
            horizontal=True,
        )
        fmt = "Markdown" if "Markdown" in fmt_raw else "TSV"

        st.markdown('<div class="run-wrap" style="margin-top:10px;">', unsafe_allow_html=True)
        run = st.button(
            "▶  Run Report",
            type="primary",
            disabled=not has_file,
            use_container_width=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)
        if not has_file:
            st.markdown(
                '<div style="font-size:10.5px;color:var(--tx3);text-align:center;'
                'margin-top:6px;font-family:var(--font);">Upload a workbook first</div>',
                unsafe_allow_html=True,
            )

    # ── OUTPUT column ────────────────────────────────────────────────────────
    with col_output:
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

        if "last_result" in st.session_state:
            st.markdown(_section_label("↓", "Results"), unsafe_allow_html=True)
            _entry = REGISTRY_BY_ID[st.session_state["last_entry_id"]]
            render_results(
                _entry,
                st.session_state["last_result"],
                st.session_state["last_fmt"],
            )
        else:
            st.markdown("""
<div class="empty-state anim-fade-up">
  <div class="empty-state-icon">&#9678;</div>
  <div class="empty-state-title">Output will appear here</div>
  <div class="empty-state-desc">
    Select a report type, upload your Excel workbook,
    and click <strong style="color:var(--gold);">Run Report</strong> to generate results.
  </div>
</div>""", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)  # /main-layout
    st.markdown('<div style="height:40px;"></div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
