# PRP Automation Dashboard

AB InBev TPRM & Risk Intelligence dashboard built with Streamlit.

## Prerequisites

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
