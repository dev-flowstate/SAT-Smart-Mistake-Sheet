# SAT Smart Mistake Sheet

**Made by Muhammad Salar Khan** · Instagram [@salars_catalogue](https://www.instagram.com/salars_catalogue/)

An Excel mistake log for the **Digital SAT (Bluebook)** — log every question you get wrong,
find the topics you keep missing, and review them before the next test.

## Download

| File | Use it when |
|---|---|
| **[`SAT_Mistake_Log_v11.xlsm`](SAT_Mistake_Log_v11.xlsm)** | Windows Excel. Buttons inside the cells, paste question screenshots with Ctrl+V, click to zoom. Needs a one-time unblock. |
| **[`SAT_Mistake_Log_v11_NoMacros.xlsx`](SAT_Mistake_Log_v11_NoMacros.xlsx)** | Mac, older Excel, Excel on the web, or anyone who can't run macros. No warning, nothing to enable. |
| **[`SAT_Mistake_Log_Guide.pdf`](SAT_Mistake_Log_Guide.pdf)** | Step-by-step student guide for Windows, with real screenshots and arrows: unblocking, enabling macros, logging, pictures, and the dashboard. |

Check back here for new versions.

---

## What it does

- **300 log rows**, almost everything picked from lists instead of typed.
- **In-cell buttons** (macro edition). Click a cell and a button appears *inside* it —
  `Choose ▼` on any list cell, `Add Photo` / `Photo ▼` on the screenshot cell.
- **Dependent topics.** Pick the Section and the Topic list narrows to that section's
  official Digital SAT skills (Reading & Writing → 11, Math → 19).
- **Bluebook aware.** Test 4–11, Module 1/2, and a `Q #` field where typing `12` shows `Q12`.
- **Question screenshots** — copy a screenshot, click the cell, press **Ctrl+V**. It shrinks
  into the cell, is saved inside the workbook, and moves with its row when you sort.
  Click it to zoom; **Photo ▼** replaces or removes it; **Delete** on the cell removes it too.
- **Dashboard** with a *Top Repeat Offenders* ranking — the topics you miss most. That's the
  list you actually study from.
- Priority and Status colour themselves red / yellow / green.

## First run — unblock the macro edition

Windows marks downloaded files as untrusted, and Excel switches the buttons off with a red
**SECURITY RISK** bar that has no Enable button.

1. Close Excel.
2. Right-click the file → **Properties** → tick **Unblock** → **OK**.
3. Open it and click **Enable Content** on the yellow bar.

The [guide](SAT_Mistake_Log_Guide.pdf) shows each screen. The macros only move the buttons,
show the menus and place pictures inside the file — no internet, no other programs, no file
access. Every line is in [`vba/`](vba/).

## Compatibility

| Platform | Macro edition | NoMacros edition |
|---|---|---|
| Windows Excel 365 / 2024 / 2016–2021 | Everything, once unblocked | Everything except the buttons |
| Excel for Mac | Dropdowns, colours, dashboard | Full — recommended |
| Excel on the web | Dropdowns, colours, dashboard (macros never run) | Same |
| Google Sheets | **Not supported** — the Section → Topic link doesn't survive import | Same |

## What is locked

Students can type in rows 5–304 and use filters and sorting. The title, credits, headings,
dashboard, instructions and dropdown lists are protected, and sheets can't be deleted, renamed
or unhidden, so nothing important is lost by accident.

This is protection against accidental and casual edits, not security — Excel sheet protection
can be removed by someone determined. The guide PDF is edit-restricted as well.

---

## Version history

Files are in [`versions/`](versions/). Each one is the whole workbook at that point.

| Version | What changed |
|---|---|
| `ORIGINAL_backup.xlsx` | The starting point. 64 rows, four dropdowns, and a *Mistakes by Type* chart wired to a free-text column so it always read zero. |
| `v2.xlsx` | 1000 rows. Dependent Section → Topic dropdowns, a real **Mistake Type** column, and a hidden `Lists` sheet driving every dropdown. |
| `v3.xlsx` | Dropdown hints and `▼` headings. Fixed the total, which only counted rows with a Date. Axis titles and plain-English chart titles. |
| `v4.xlsx` | **Digital SAT** format with the official topics. **Top Repeat Offenders** ranking. Filter buttons. |
| `v5.xlsx` | **Bluebook Test**, **Module** and a `Q #` field that shows `12` as `Q12`. Per-test chart. |
| `v6.xlsx` | Performance: 300 rows, shorter rows, no banded fills. Scrolling **3.22s → 1.63s**. |
| `v7.xlsm` | First macro version: photo button, click-to-zoom, Ctrl+V fitting. |
| `v8.xlsm` | `Choose ▼` button centred in every dropdown cell. Zoom became a toggle. |
| `v9.xlsm` | Fallback picture mode, Mac-safe paths, per-button error handling. |
| `v10` | Fixed the dropdown popup that couldn't write a value. Pictures switched to portable anchored mode. Added the **NoMacros** edition. |
| **`v11`** | Photos can be **replaced and removed**. **Ctrl+V works** the moment you press it. Photos live inside the file, so zoom works after the log is emailed. Dashboard title and long chart labels fixed. **Author credits added and locked.** Student guide PDF. |

### What was wrong before v11

- **No way to remove or replace a photo.** Once a row had a picture the button only zoomed,
  and clicking the picture ran zoom instead of selecting it. v11 adds a **Photo ▼** menu
  (Zoom In / Replace Photo / Remove Photo), Delete on the cell, and Ctrl+V-to-replace.
- **Ctrl+V didn't work reliably.** Fitting a pasted picture exported it through a chart, and a
  single failed export silently disabled pasting for the whole session. v11 fits the pasted
  picture itself — no export — and a listener inside the workbook does it immediately.
- **Why not take over Ctrl+V with `Application.OnKey`?** It was tried and rejected. `OnKey`
  sets a switch on Excel itself, not on the file; if it is ever left on, pressing Ctrl+V in any
  other workbook makes Excel **reopen the log**. A real-keypress test caught exactly that.

## How it works under the hood

- **Dependent dropdowns**: a hidden column maps the Section to a named range (`secRW` /
  `secMath`); the validation source is `=INDIRECT($O5)`, which Excel adjusts per row.
- **Menus**: `CommandBar` popups. Each entry carries its index in `Tag` and calls one
  argument-free macro that reads `CommandBars.ActionControl.Tag` — no fragile `OnAction`
  argument quoting. A literal `&` in a caption must be written `&&`.
- **Pictures**: ordinary pictures set to *move and size with cells*, found by position rather
  than name so they stay matched after sorting. Not Excel's newer in-cell pictures, which
  show `#UNKNOWN!` in Excel 2016/2019/2021.
- **Paste listener**: a class module handling `CommandBars.OnUpdate` notices a new picture and
  fits it. It lives and dies with the workbook, so nothing is left behind when it closes.
- **Conditional formatting fills use `bgColor`**, not `fgColor` — with `fgColor` the rule
  applies and paints nothing.
- **Repeat-offender ranking** uses `LARGE` over a `COUNTIF + ROW()/100000` key, since `SORT`
  isn't available everywhere.

## Layout

```
Mistake Log   300 entry rows (hidden helper columns O and P)
Dashboard     counts, Top Repeat Offenders, 6 charts
How To Use    student instructions
Lists         hidden, locked - every dropdown's options
vba/          the macros as readable source
```

---

**SAT Smart Mistake Sheet** · Made by **Muhammad Salar Khan** ·
Instagram [@salars_catalogue](https://www.instagram.com/salars_catalogue/) ·
Updates: [github.com/dev-flowstate/SAT-Smart-Mistake-Sheet](https://github.com/dev-flowstate/SAT-Smart-Mistake-Sheet)
