# Sharing the dashboard (download & double-click)

The dashboard is a **self-contained offline app**. It runs Python entirely in
the browser (via Pyodide, bundled in the `pyodide/` folder). A person who
receives it needs **no Python, no install, no server** — they just open
`index.html`.

## What must be shared together

Keep these next to each other (this is what makes double-click work):

```
index.html        ← the dashboard (UI + all automation scripts embedded)
logo.png          ← branding
pyodide/          ← the offline Python runtime + packages  (≈95 MB)
```

Everything else (the `*.py` sources, the sample `.xlsx`) is optional.

## Publish it to GitHub

> ⚠️ The runtime file `pyodide/assets-b64.js` is ~95 MB. GitHub's **web upload
> caps at 25 MB**, so you must publish with **git** (not drag-and-drop). git
> accepts it (limit is 100 MB); the first push may take a minute.

```bash
git add -A
git commit -m "Publish offline dashboard (double-click index.html to run)"
git push origin main
```

The remote is already set to `gtdevcode2026/PRP-update`.

## How the recipient runs it

1. On the GitHub repo page: **Code → Download ZIP**.
2. **Unzip** it (keep `index.html` and the `pyodide/` folder together).
3. **Double-click `index.html`.**

The first launch takes a moment while the browser loads the Python runtime from
the local folder; after that it's cached. They upload their Excel workbook,
pick a report, and click **Run** — all offline.

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
