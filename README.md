# PM-Site-tracker-Lite

A lighter version of a site progress tracker, helping project managers and coordinators track work across sites.

The app reads an Excel tracker workbook, calculates the progress figures itself, and sends the finished figures to an AI to write a short interpretation and conclusion.

## How it works

1. **The engine calculates every figure** (task counts, task-group progress, pending prerequisites, blocked tasks, forecast, overall pace). No AI is involved, so the numbers are repeatable and tested.
2. **The AI writes only the conclusion**: whether the sites will finish on time, and where a strategy change or resource reallocation is needed. It is always labelled as AI-written and it never recalculates a figure.
3. The result is a one-page brief, exportable as PDF, Word and Excel.

The workbook is read in the browser and is never uploaded anywhere except, if you ask for a conclusion, the summary figures sent to the AI provider (Gemini).

## Status

| Stage | What | State |
|---|---|---|
| 0 | Repo skeleton, parser/engine/report reused unchanged from the full app, tests | done |
| 1 | `js/core/brief.js`: the one-page figures, with tests | done |
| 2 | Browser page: load, check what was read, pick sites, show the brief | next |
| 3 | Gemini conclusion | planned |
| 4 | PDF / Word / Excel export | planned |
| 5 | Session save / load | planned |
| 6 | GitHub Pages deploy | planned |

There is no browser page yet. You can already check the figures against your own workbook.

## Try it on a workbook

```bash
node tools/print-brief.mjs path/to/your-workbook.xlsx
node tools/print-brief.mjs this-week.xlsx --previous last-week.xlsx --sites A-01,B-02
node tools/print-brief.mjs this-week.xlsx --json      # the exact data the AI would be given
```

Needs Node 20 or later. There is nothing to install.

## Tests

```bash
npm test
```

Runs the undefined-function check, the ported parser/engine/report suite and the brief suite. The same three run in CI on every push (`.github/workflows/test.yml`).

## Privacy

- Workbooks use **codes only** (site `A-01`, person `R1`); no real names or places.
- `.gitignore` blocks every `.xlsx` except the template, and any session file. Never commit a workbook or session file: this repository is public.
- The app flags possible personal or place names in the workbook before anything is sent to an AI. The check is a heuristic and only warns.

## Layout

```
js/core/parser.js engine.js report.js util.js   reused unchanged from the full app
js/core/brief.js                                the one-page figures (new)
tools/verify-core.mjs verify-brief.mjs          tests
tools/print-brief.mjs                           prints the brief for a workbook
templates/                                      the tracker workbook template
assets/vendor/                                  SheetJS (Apache-2.0), vendored
```
