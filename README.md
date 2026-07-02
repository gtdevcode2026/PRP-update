# PRP Automation Dashboard

AB InBev TPRM & Risk Intelligence dashboard. It runs entirely in the browser —
no Python, no install, no server required to use it.

## Quick start — just open it (recommended)

1. Download this repo: **Code → Download ZIP**, then unzip it.
2. Keep `index.html` next to the `pyodide/` folder.
3. **Double-click `index.html`.** Upload your Excel workbook, pick a report, click **Run**.

Everything (the Python runtime, pandas, matplotlib, all automation scripts)
is bundled in the `pyodide/` folder and runs offline in your browser. Recommended
browsers: Chrome or Edge.

> Sharing it with others via GitHub? See **[DEPLOY.md](DEPLOY.md)** — the `pyodide/`
> runtime is ~95 MB, so publish with `git push` (GitHub's web upload caps at 25 MB).

## Developer setup (optional — Flask/Streamlit backends)

The scripts can also be run through a small server (`server.py` for Flask, or the
legacy `app.py` for Streamlit) instead of the offline bundle.

### Prerequisites

- Python 3.9 or higher
- pip

## Setup

1. **Clone the repository**

   ```bash
   git clone https://github.com/gtdevcode2026/PRP-update.git
   cd PRP-update
   ```

2. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

3. **Place the input file**

   Copy `PRP Sample Jun (2).xlsx` into the project root (same folder as `app.py`).

## Running the App

```bash
python -m streamlit run app.py
```

The dashboard opens automatically at `http://localhost:8501`.

## Usage

1. Upload `PRP Sample Jun (2).xlsx` using the **Input** file uploader (or it auto-loads if the file is in the project root).
2. Select a **report** from the left sidebar (Vendor Onboarding, Response Time Reduction, etc.).
3. Choose a **variant** if multiple options are available.
4. Click **Run** to execute the automation script and preview charts and tables.
5. Download the generated `.xlsx` report from the **Download Options** section.

## Reports

| # | Name | Source Sheet |
|---|------|-------------|
| 1 | Vendor Onboarding (Critical Tech) | TPRM Web-Portal Export |
| 2 | Response Time Reduction | OneTrust Assessment |
| 3 | Reduce Long Over-Due Assessments | OneTrust Assessment |
| 4 | Risk Treatment | OneTrust - Risk Export |
| 5 | Suppliers in Scope | TPRM Web-Portal Export |
| 6 | Risk Assessment Progress | OneTrust Assessment |
| 7 | Risks Identified | OneTrust Assessment |
