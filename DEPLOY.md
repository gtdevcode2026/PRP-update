# Sharing the dashboard (download & double-click)

The dashboard is a **self-contained offline app**. Every report runs in plain
JavaScript in the browser (SheetJS/ExcelJS/Plotly, all vendored locally). A
person who receives it needs **no Python, no install, no server** — they just
open `index.html`.

## What must be shared together

Keep these next to each other (this is what makes double-click work):

```
index.html        ← the dashboard UI
logo.png          ← branding
vendor/           ← SheetJS, ExcelJS, Plotly (≈6 MB total)
js/                ← report engine + the 7 report generators
```

Everything else (the `*.py` sources, the sample `.xlsx`) is optional — the
`.py` files are kept only as a historical reference for the report logic.

## Publish it to GitHub

The whole working tree (excluding git history) is only a few MB now, so a
normal **web upload** or `git push` both work fine — no special handling needed.

```bash
git add -A
git commit -m "Publish offline dashboard (double-click index.html to run)"
git push origin main
```

## How the recipient runs it

1. On the GitHub repo page: **Code → Download ZIP**.
2. **Unzip** it (keep `index.html`, `vendor/`, and `js/` together).
3. **Double-click `index.html`.**

It opens instantly — no boot delay of any kind. They upload their Excel
workbook, pick a report, and click **Run** — all offline.

> Alternatively they can `git clone` the repo and open `index.html` — same result.

## Before you share: test it yourself

You already have the folder locally — **double-click `index.html` right now** to
confirm it launches in your browser. Recommended browsers: **Chrome or Edge**
(latest). If it opens and you can run a report, recipients will get the same.

## Notes

- **Cross-platform:** works on Windows/Mac/Linux — it's just a web page + files.
- **No secrets are published.** Uploaded workbooks are processed only in the
  visitor's own browser; nothing is sent anywhere.
- Want a **click-a-link** version instead of download-and-open (hosted on the
  web, no download)? That's also possible via GitHub Pages — ask and it can be
  added.
