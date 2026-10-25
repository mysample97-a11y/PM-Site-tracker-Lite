# PM-Site-tracker-Lite

A lighter version of a site progress tracker, helping project managers and coordinators track works across sites.

Load the tracker Excel workbook and the app produces a **one-page status brief**: which sites are tracked, task groups finished, tasks done / WIP / blocked / not started, per-task-group progress, pending prerequisites, blocked tasks, overall progress, and a conclusion on whether the sites will finish on time and where strategy or resources need to change.

## How it works

1. **The app calculates every figure itself** (`js/core/brief.js`, on the parser/engine/report modules reused unchanged from the full tracker). No AI is involved, so the numbers are repeatable and tested.
2. **The AI writes only the conclusion**, from the finished figures, never from the spreadsheet. It is labelled *AI-written* with the model name, its reply is cut back to the selected sites before it reaches the page, and it makes **one request for all selected sites**.
3. The brief exports as **PDF, Word and Excel**.

The workbook is read in the browser and never uploaded. The only network call is the one you ask for: the AI conclusion, straight from your browser to Google Gemini.

## Using it

1. Open the page and drop in your tracker workbook (an optional second slot takes last week's workbook, to show what changed).
2. **Check what was read**: sites found, which sheet each matched, warnings, and a flag if the workbook seems to contain personal or place names. Tick the sites to include.
3. **Generate brief.** Optionally **Get AI conclusion** (needs a Gemini API key from Google AI Studio; you see exactly what will be sent first).
4. Export as PDF / Word / Excel, or **Save session** to continue later.

**Reloading the page or closing the tab clears everything.** The only way to keep your work is *Save session*, which writes a file holding the selected sites only and never the API key.

## Privacy and security

- Use **codes only** in the workbook (site `A-01`, person `R1`). The app flags likely names, but the check is a guess and only warns.
- This is a public repository. `.gitignore` blocks every `.xlsx` except the template and any session file. **Never commit a workbook or a session file.**
- **API key:** held in memory for the tab; optionally kept through a refresh (cleared when the tab closes). Never saved to a session file, an export or `localStorage`. Anything that can run script on the page could read it, so use a key made for this tool with a spend cap.
- **Free Gemini tier:** Google may use what you send to improve its products, and people may read it. The review step says so before anything is sent.
- Settings (model names, limits) are kept in the browser's `localStorage`; they are not project data.

## Publishing (GitHub Pages)

In the repository: **Settings → Pages → Source: GitHub Actions**. A push to `main` runs the tests and, if they pass, publishes the page (only `index.html`, `assets/`, `js/` and `templates/`).

## Development

```bash
npm ci
npm test                                   # everything below, in order
node tools/print-brief.mjs workbook.xlsx   # print the brief for a workbook, no browser
python3 -m http.server                     # serve the page locally
```

| Suite | Covers |
|---|---|
| `tools/check-refs.mjs` | every function call resolves |
| `tools/verify-core.mjs` | parser, engine, report (ported from the full app) |
| `tools/verify-brief.mjs` | every figure on the page, against hand-worked numbers |
| `tools/verify-ai.mjs` | retry, rate-limit, bad key, fallback model, cancel, reply checking (stubbed Gemini) |
| `tools/verify-view.mjs` | page escaping, print document, Word/Excel content, session file |
| `tools/smoke-dom.mjs` | the whole workflow in a headless browser (jsdom) |

jsdom has no layout engine: these tests prove behaviour, not appearance. Check the PDF in a real browser's print preview.

## Layout

```
index.html  assets/css/lite.css  assets/vendor/   page, styles, SheetJS (Apache-2.0, vendored)
js/views/app.js                                   the page
js/core/parser.js engine.js report.js util.js     reused unchanged from the full tracker
js/core/brief.js                                  the one-page figures
js/core/briefview.js                              draws the brief (screen and print)
js/core/ai.js tokens.js                           Gemini client, retry, usage meter
js/core/session.js                                state, settings, key, session file
js/core/exports.js                                PDF / Word / Excel
templates/                                        the tracker workbook template
tools/                                            tests and print-brief
```
