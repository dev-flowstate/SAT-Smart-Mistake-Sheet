# SAT Smart Mistake Sheet

An Excel mistake log for the **Digital SAT (Bluebook)** — log every question you get wrong,
find the topics you keep missing, and review them before the next test.

**Current version: [`SAT_Mistake_Log_v9.xlsm`](versions/SAT_Mistake_Log_v9.xlsm)**

---

## What it does

- **300 log rows**, everything picked from lists instead of typed.
- **In-cell buttons.** Click a cell and a button appears *inside* it — `Choose ▼` on any
  list cell, `Add Photo` / `Zoom In` on the screenshot cell.
- **Dependent topics.** Pick the Section and the Topic list narrows to that section's
  official Digital SAT skills (Reading & Writing → 11 topics, Math → 19).
- **Bluebook aware.** Test 4–11, Module 1/2, and a `Q #` field where typing `12` displays `Q12`.
- **Question screenshots** that sit inside the cell and open full-size when clicked.
- **Dashboard** with a *Top Repeat Offenders* ranking — the topics you miss most, ranked
  automatically. That's the list you actually study from.
- Priority and Status colour themselves red / yellow / green.

## Requirements

| Setup | What works |
|---|---|
| Windows + Microsoft 365 | Everything |
| Excel 2016 / 2019 / 2021, Excel for Mac | Everything, but photos use the older picture mode (shrunk to fit the cell, click to zoom) |
| Excel on the web, Google Sheets | Log, dropdowns, colours, dashboard and charts. **No buttons** — macros can't run |

### First run — unblock the file

Windows marks downloaded files as untrusted and Excel silently switches the buttons off.

1. Close Excel.
2. Right-click the file → **Properties**.
3. Tick **Unblock** at the bottom → **OK**.
4. Reopen. If a yellow bar appears, click **Enable Content**.

Skip this and the log still works — you just lose the buttons, photos and zoom.

---

## Version history

Files are in [`versions/`](versions/). Each one is the whole workbook at that point.

| Version | What changed |
|---|---|
| `SAT_Mistake_Log_ORIGINAL_backup.xlsx` | The starting point. 64 rows, four dropdowns, a dashboard whose *Mistakes by Type* chart was wired to a free-text column and always read zero. |
| `v2.xlsx` | 1000 rows. Added dependent Section → Topic dropdowns (hidden helper + `INDIRECT` + named ranges), a real **Mistake Type** column so the broken chart actually counts something, and a hidden `Lists` sheet driving every dropdown. |
| `v3.xlsx` | Dropdown hints on every list column and `▼` markers in the headings. Fixed the total, which counted only rows with a **Date** and so under-reported. Added axis titles and plain-English chart titles. |
| `v4.xlsx` | Switched to **Digital SAT** format (Reading & Writing / Math) with the official domain topics. Added the **Top Repeat Offenders** ranking. Filter buttons on every heading. Merged the duplicate *Concept to Review* column away. |
| `v5.xlsx` | **Bluebook Test** (4–11) and **Module** dropdowns, plus a `Q #` field where you type `12` and it shows `Q12`. Added a per-test breakdown chart. |
| `v6.xlsx` | Performance. 500 → 300 rows, row height 60 → 38, removed ~4,000 banded cell fills. Scrolling went from **3.22s → 1.63s** and the file from 55 KB → 43 KB. |
| `v7.xlsm` | First macro version. A photo button that centres itself in the screenshot cell, click-to-zoom, and Ctrl+V pastes auto-fitting into their row. |
| `v8.xlsm` | The **`Choose ▼` button now centres itself in every dropdown cell** — measured at `offset (0,0)` on all seven list columns. Zoom became a toggle: a second click puts the picture back. |
| `v9.xlsm` | Made it safe to hand to other people. Falls back to a classic sized-to-fit picture where in-cell pictures aren't supported, Mac-safe paths, and every button catches its own errors instead of throwing a VBA dialog. |

## How it works under the hood

A few things that were not obvious and are worth knowing if you edit this:

- **Dependent dropdowns** use a hidden column that maps the Section to a named range
  (`secRW` / `secMath`), and the validation source is `=INDIRECT($O5)`. Excel adjusts that
  reference per row on its own.
- **Conditional formatting fills must use `bgColor`**, not `fgColor`. With `fgColor` the rule
  applies correctly and simply paints nothing — which is exactly how the colours went missing.
- **Excel cannot paste a picture into a cell.** `Ctrl+V` always creates a floating picture.
  The macro copies it through a throwaway chart, exports a PNG, and re-inserts it with
  `InsertPictureInCell`.
- **The repeat-offender ranking** can't use `SORT`, so it ranks with `LARGE` over a
  `COUNTIF + ROW()/100000` key, which makes every value unique and breaks ties cleanly.

## Layout

```
Mistake Log   the 300-row log (hidden helper columns O and P)
Dashboard     counts, Top Repeat Offenders, 6 charts
How To Use    student-facing instructions
Lists         hidden - every dropdown's options; edit here to change them
```

To change what any dropdown offers: right-click a sheet tab → **Unhide** → **Lists**.
