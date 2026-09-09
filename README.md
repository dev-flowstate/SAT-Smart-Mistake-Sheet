# SAT Smart Mistake Sheet

An Excel mistake log for the **Digital SAT (Bluebook)** — log every question you get wrong,
find the topics you keep missing, and review them before the next test.

## Download

Two editions, same 300 rows, same dropdowns, same dashboard. Pick by machine:

| File | Use it when |
|---|---|
| **[`SAT_Mistake_Log_v10.xlsm`](SAT_Mistake_Log_v10.xlsm)** | Windows Excel. Adds buttons **inside** the cells, one-click photo, click-to-zoom. Must be unblocked first (see below). |
| **[`SAT_Mistake_Log_v10_NoMacros.xlsx`](SAT_Mistake_Log_v10_NoMacros.xlsx)** | Anything else — Mac, older Excel, Excel on the web, or anyone who can't run macros. No warning, nothing to enable. |

---

## What it does

- **300 log rows**, everything picked from lists instead of typed.
- **In-cell buttons** (macro edition). Click a cell and a button appears *inside* it —
  `Choose ▼` on any list cell, `Add Photo` / `Zoom In` on the screenshot cell.
- **Dependent topics.** Pick the Section and the Topic list narrows to that section's
  official Digital SAT skills (Reading & Writing → 11, Math → 19).
- **Bluebook aware.** Test 4–11, Module 1/2, and a `Q #` field where typing `12` shows `Q12`.
- **Question screenshots** anchored inside the cell, travelling with the row when you sort.
- **Dashboard** with a *Top Repeat Offenders* ranking — the topics you miss most, ranked
  automatically. That's the list you actually study from.
- Priority and Status colour themselves red / yellow / green.

## First run — unblock the macro edition

Windows marks downloaded files as untrusted and Excel switches the buttons off. Since 2022
this is a red **SECURITY RISK** bar with **no "Enable Content" button**, so it has to be
cleared on the file itself:

1. Close Excel.
2. Right-click the file → **Properties**.
3. Tick **Unblock** at the bottom → **OK**.
4. Reopen. If a yellow bar appears now, click **Enable Content**.

There is no way to avoid this from the sender's side short of a code-signing certificate
that each recipient already trusts. If that's a hassle, hand out the **NoMacros** edition —
it needs none of this.

## Compatibility

| Platform | Macro edition | NoMacros edition |
|---|---|---|
| Windows Excel 365 / 2024 | Everything, once unblocked | Everything except buttons |
| Windows Excel 2016 / 2019 / 2021 | Everything, once unblocked | Everything except buttons |
| Excel for Mac | Dropdowns, colours, dashboard | Full — recommended here |
| Excel on the web | Dropdowns, colours, dashboard. Macros never run | Same |
| Google Sheets | **Not recommended** — the Section→Topic link doesn't survive import | Same |

**Pictures use ordinary anchored pictures, not Excel's newer "Place in Cell".** In-cell
pictures are stored as a rich value that Excel 2016/2019/2021 cannot read — those versions
display `#UNKNOWN!` instead of the image. Anchored pictures render everywhere, so the log
survives being emailed around a mixed classroom.

---

## Version history

Files are in [`versions/`](versions/). Each is the whole workbook at that point.

| Version | What changed |
|---|---|
| `ORIGINAL_backup.xlsx` | The starting point. 64 rows, four dropdowns, and a *Mistakes by Type* chart wired to a free-text column so it always read zero. |
| `v2.xlsx` | 1000 rows. Dependent Section → Topic dropdowns (hidden helper + `INDIRECT` + named ranges), a real **Mistake Type** column so the broken chart counts something, and a hidden `Lists` sheet driving every dropdown. |
| `v3.xlsx` | Dropdown hints and `▼` markers on every list heading. Fixed the total, which counted only rows with a **Date** and under-reported. Axis titles and plain-English chart titles. |
| `v4.xlsx` | Switched to **Digital SAT** format (Reading & Writing / Math) with the official domain topics. Added the **Top Repeat Offenders** ranking. Filter buttons. Merged away the duplicate *Concept to Review* column. |
| `v5.xlsx` | **Bluebook Test** (4–11) and **Module** dropdowns, plus a `Q #` field that displays `12` as `Q12`. Per-test breakdown chart. |
| `v6.xlsx` | Performance. 500 → 300 rows, row height 60 → 38, removed ~4,000 banded cell fills. Scrolling **3.22s → 1.63s**, file 55 KB → 43 KB. |
| `v7.xlsm` | First macro version. Photo button centring itself in the screenshot cell, click-to-zoom, and Ctrl+V pastes auto-fitting into their row. |
| `v8.xlsm` | The **`Choose ▼` button now centres itself in every dropdown cell** — measured `offset (0,0)` on all seven list columns. Zoom became a toggle. |
| `v9.xlsm` | Hardening for sharing: fallback picture mode, Mac-safe paths, per-button error handling. |
| **`v10`** | **Fixed the popup that couldn't write a value** (see below), flipped pictures to the portable anchored mode so they don't break on older Excel, added per-column validation messages, and split out a **NoMacros** edition. |

### The v10 bug, for the record

A `CommandBar` control's `OnAction` **must have the whole call wrapped in single quotes**
when it passes an argument. Without them the macro silently doesn't run:

```vba
ci.OnAction = "PickByIndex " & i      ' runs nothing, no error
ci.OnAction = "'PickByIndex " & i & "'"  ' works
```

It shipped because `ShowPopup` is modal and couldn't be driven from a test. The fix was to
split `BuildPopupBar()` out from `ShowListPopup()` so the menu can be built and its entries
fired with `.Execute()` in a test, plus a `Tag`-based fallback in the handler for builds and
locales that drop the argument anyway.

## How it works under the hood

- **Dependent dropdowns** use a hidden column mapping the Section to a named range
  (`secRW` / `secMath`), with the validation source `=INDIRECT($O5)`. Excel adjusts that
  reference per row by itself.
- **Conditional formatting fills must use `bgColor`**, not `fgColor`. With `fgColor` the rule
  applies correctly and paints nothing — which is how the colours went missing for a while.
- **Excel cannot paste a picture into a cell.** `Ctrl+V` always creates a floating picture.
  The macro copies it through a throwaway chart, exports a PNG, and re-places it anchored.
- **The repeat-offender ranking** can't use `SORT`, so it ranks with `LARGE` over a
  `COUNTIF + ROW()/100000` key, which makes every value unique and breaks ties cleanly.

## Layout

```
Mistake Log   the 300-row log (hidden helper columns O and P)
Dashboard     counts, Top Repeat Offenders, 6 charts
How To Use    student-facing instructions
Lists         hidden - every dropdown's options; edit here to change them
vba/          the macros as readable source, so they can be reviewed without opening Excel
```

To change what any dropdown offers: right-click a sheet tab → **Unhide** → **Lists**.
