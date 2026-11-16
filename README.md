# Kimai Timesheet Updater

Automatically generates the next month's **Excel timesheet** from a **Kimai CSV export**
— filling every person's sheet (dates, activity, hours), preserving the original
template's formatting exactly, and updating the **Team Budget Report** summary.

**Status:** standalone script OrbisDS uses today. Being integrated into the OrbisDS
platform (Kimai ↔ Timesheets); see [Roadmap](#roadmap).

---

## Table of contents

- [What it does](#what-it-does)
- [How it works — pipeline](#how-it-works--pipeline)
- [Quick start](#quick-start)
- [Folder layout](#folder-layout)
- [CLI reference](#cli-reference)
- [Kimai CSV format](#kimai-csv-format)
- [Template "contracts"](#template-contracts)
- [Development / how to edit correctly](#development--how-to-edit-correctly)
- [Building the .exe](#building-the-exe)
- [Testing a change](#testing-a-change)
- [Roadmap](#roadmap)
- [Author](#author)

---

## What it does

Every month someone manually copies Kimai hours into styled Excel timesheets. This
script removes the manual work and produces a file that matches the template exactly:

- **the CSV is the only source of people** — exactly the people in the CSV export get
  a sheet in the result (nothing stale from the template carries over),
- fills each person's sheet with the month's days (`dd.mm.yyyy`, activity, decimal hours),
  cloning the template's one placeholder person sheet per person,
- keeps the template's borders, colours, column widths, fonts and number formats,
- updates every `SUBTOTAL` formula (dynamic, so it still works if rows are added by hand),
- updates the **Team Budget Report** with **live `INDEX/MATCH` formulas** into each
  person's own sheet,
- writes the result as a **new file** — the template is never overwritten.

It reads **`.xlsx`** and legacy **`.xls`** templates (`.xls` is converted in-memory).

---

## How it works — pipeline

```
Kimai CSV ──parse──▶ { name → [ {date, hours, description}, … ] }
                            │
Excel template ──load──▶ workbook (.xlsx / .xls→.xlsx)
   summary sheet + one "placeholder timesheet" sheet
                            │
    1. Detect months (template month → target month)
    2. Find the placeholder person sheet
    3. Per CSV person: clone the placeholder → "Timesheet <Name>", fill it
    4. Remove the placeholder sheet
    5. Rebuild Team Budget Report rows from the CSV people (live formulas)
                            │
                            ▼
                     output/<new_month>.xlsx
```

Concrete steps, matching the function names in `update_timesheet.py` so edits are
easy to trace:

1. **Load** — `load_workbook_any()`: `.xlsx` via `openpyxl`; `.xls` via `xlrd`, converting
   values + styles (font, borders, alignment, fill, number format, column widths,
   merged cells) into an `openpyxl` workbook.
2. **Parse CSV** — `parse_kimai_csv()` → `{Name: [entries]}`. Rows without a name or with
   an unparseable date/duration are skipped silently.
3. **Detect months** — reads the template's current month/year (from the sheet title via
   `extract_month_year` regex, or the budget sheet's title). The **target month** is the
   month with the most CSV entries (`csv_target_month`), falling back to the month
   *after* the template (`next_month`).
4. **The CSV is the only source of people** — the template holds one empty placeholder
   person sheet (its name starts with `placeholder`). For every person in the CSV the
   script `copy_worksheet`s that placeholder into `Timesheet <Name>` (`_unique_sheet_title`
   keeps titles unique/pre-31 chars) and fills it.
5. **Fill each person's sheet** — `update_person_sheet_v2()`: write the worker's name in
   the first data row, insert/delete rows as needed, write values only, set the
   `SUBTOTAL(9, …:INDEX(…,ROW()-1))` formula, then stamp template styles onto every data
   row (`_apply_template_style`).
6. **Remove the placeholder** — the scaffolding never belongs in the result.
7. **Team Budget Report** — `update_budget_report()`: drops the template's placeholder
   people rows, then writes one row **per CSV person** with live `=INDEX('Timesheet X'!D:D,
   MATCH("Total hours", …))` and name references, plus a team `SUBTOTAL`. The Role column
   is left blank for the PM to fill in.
8. **Save** — a new file named like the template but with the new month, into `output/`
   (or `--output` / `--outdir`).

---

## Quick start

### A. Windows — pre-built `.exe` (no Python)

1. Put your previous month's template in `input\` (e.g. `Timesheet Report June 2026.xlsx`).
2. Put the Kimai CSV export in `input\` — **all people in one file**.
3. Double-click **`run.bat`**.
4. Result appears in `output\` with the new month's name.

> `input\` must contain **exactly one** template and **exactly one** CSV, or `run.bat`
> stops with "Multiple files found". See [Building the .exe](#building-the-exe).

### B. From source (Python)

```bash
pip install -r requirements.txt

python update_timesheet.py \
  "examples/timesheets/Acme - Timesheet Report October 2026.xlsx" \
  examples/kimai_export.csv          # → output/Acme - Timesheet Report November 2026.xlsx
```

---

## Folder layout

```
.
├── update_timesheet.py    # the whole script (single file, no framework)
├── run.bat                # Windows launcher (.exe preferred, source fallback)
├── build_exe.bat          # Windows: rebuild dist\timesheet-updater.exe
├── requirements.txt       # openpyxl, xlrd
├── timesheet-updater.spec # PyInstaller spec
├── examples/              # sample templates + sample Kimai CSV (demo/validation inputs)
│   ├── timesheets/        #   anonymized fake-company template (Acme Studio)
│   └── kimai_export.csv   #   demo export (people from the Acme template)
├── scripts/               # maintenance tools
│   └── build_fake_template.py  #  regenerate the anonymized example from a real template
├── input/                 # drop template + CSV here, then run run.bat
├── output/                # results
└── dist/                  # built .exe (created by build_exe.bat)
```

---

## CLI reference

```
python update_timesheet.py <excel_template> <kimai_csv> [options]

options:
  -o, --output <file>   exact output path
      --outdir <dir>    output directory for the auto-named file (default: ./output)
```

Output naming: the template's filename with the old month/year replaced by the new one.

---

## Kimai CSV format

| Column | Example | Notes |
|--------|---------|-------|
| `Name` | `John A. Doe` | names the person's sheet (the CSV drives the people) |
| `Date` | `2026-11-02` | `%Y-%m-%d` |
| `Duration` | `8:00` | `h:mm`; summed per day, rounded to 2 decimals |
| `Description` | `Frontend work` | multiple entries per day joined with `; ` |

- **Header row required**; columns are read **by name** (`csv.DictReader`), so the order
  and any extra columns (the full Kimai export has ~27) do not matter — `Name`, `Date`,
  `Duration`, `Description` are picked out regardless of position. Other columns
  (`Project`, `Activity`, `Customer`, …) are currently not used.
- **Only the people in the CSV get a sheet** — the result contains nothing stale from the
  template. A person with no entries or 0 hours still appears (cloned + listed) so the PM
  can spot and fix it.

---

## Template "contracts"

These are the assumptions the script relies on. **Changes to templates must respect
them** (or the logic below must change with them):

- **The template is minimal and per project**: one **summary sheet** (`Budget` in the name,
  or an `Actual Hours` header) plus **one empty placeholder person sheet** whose name
  starts with `placeholder` (e.g. `placeholder timesheet 01`) — the formatting source.
- **The placeholder is scaffolding only.** The script `copy_worksheet`s it into a
  `Timesheet <Name>` per CSV person and **removes it** from the result — it is what makes
  every generated person sheet match the template exactly.
- **Summary sheet** has `Budget` in its name, or contains an `Actual Hours` header.
- **Month/year** appears in a title as `<Full Month> <year>` (e.g. `Timesheet Report June 2026`);
  `extract_month_year()` regex uses the `MONTHS` list in `update_timesheet.py`.
- **Two person-sheet layouts** are detected by the `Timesheet date` header (in row 6 or 7):
  - **Format A** (date is a real date value) → columns `B C D E` = date, name, activity, hours.
  - **Format B** (date is a text string) → columns `A B C D` = date, name, activity, hours.
  - Detection lives in `detect_sheet_format()`; the returned dict drives all writes.
- **Total** is a row whose activity cell contains `Total hours` (found below the data).
- **Budget rows** come from the CSV people: the script drops any placeholder rows, then
  writes one row per CSV person after the `Worker name` header (column B), with the
  summary's `Total hours` / `SUBTOTAL` row below. Machine-assigned **Role is left blank**
  for the PM to fill in.

---

## Development / how to edit correctly

Keep these invariants in mind — they are why the output stays "as the template".

1. **Values don't carry style.** The script writes values *and* relies on openpyxl keeping
   the template's existing cell style. When you insert rows (for a person with many days),
   `insert_rows()` creates **unstyled** cells, so `_apply_template_style()` stamps the
   first-data-row style onto every written row. **If you change the data-writing path,
   re-apply styles the same way** — otherwise new rows lose borders/fill.
2. **Date number format.** Written dates are forced to `dd.mm.yyyy`
   (`DATE_NUM_FMT`), matching the Bulgarian-style day display the template uses.
3. **Formulas are dynamic.** `SUBTOTAL(9, D{first}:INDEX(D:D, ROW()-1))` sums only
   numeric cells and covers rows added manually. Don't replace with a fixed-range `SUM`.
4. **People come from the CSV only.** Person sheets are *cloned* from the placeholder per
   CSV person and the placeholder is removed — the script never matches CSV names to
   existing sheets. If you add a project/role mapping later, feed it into the summary from
   the CSV (or a per-project source) rather than from the template's own sheets.
5. **Month advance is data-driven.** Prefer the CSV's dominant month over
   "next month". Only when the CSV is empty does it fall back to `next_month`.
6. **Don't add a Python convention the project doesn't use.** It's a single, dependency-free
   stdlib script: no classes, no dataclasses, no type-only scaffolding beyond the `dict`
   shapes already present. Match the surrounding style (plain functions, `dict`-typed
   helpers with `Optional` where needed).
7. **The CSV drives the people.** Sheet titles come from the CSV `Name`s
   (`_person_sheet_title`/`_unique_sheet_title` sanitize + de-dupe). No name→sheet fuzzy
   matching exists anymore.

### Where to change what

| I want to… | Look at |
|---|---|
| change the CSV columns read | `parse_kimai_csv()` |
| add/remove a printed field per row | `update_person_sheet_v2()` + `apply_table_style` |
| support a new sheet layout | `detect_sheet_format()` (and the `fmt` dict consumers) |
| change month/title handling | `extract_month_year`, `next_month`, `replace_month_in_text` |
| change the budget summary | `update_budget_report()` |
| change output naming/folder | `main()` output block |
| rebuild the .exe | `build_exe.bat` / `timesheet-updater.spec` |

---

## Building the .exe

```bat
:: on Windows, from this folder
build_exe.bat
```

(PyInstaller + deps are installed into `.venv` if present, else system python; output is
`dist\timesheet-updater.exe`. `run.bat` uses the exe if it exists, otherwise falls back to
`python update_timesheet.py`.)

---

## Testing a change

The repo ships ready-to-run fixtures under `examples/` so a change can be validated
without touching real data:

```bash
# 1. Source run against the demo (fake Acme template + demo CSV)
python update_timesheet.py \
  "examples/timesheets/Acme - Timesheet Report October 2026.xlsx" \
  examples/kimai_export.csv --outdir /tmp/kts-out

# 2. Regenerate the anonymized example from a real template (kept offline, never committed)
python scripts/build_fake_template.py "/path/to/a/real/timesheet.xlsx"
```

Check: exit code `0`, output file exists with the new month in the name, a person sheet
shows the written days + hours, and the Team Budget Report totals are non-zero. The
critical regression to watch is **formatting** — visually compare an output sheet against
its template (borders/fill/widths intact), especially for people with *more* entry days
than the template had (that exercises `_apply_template_style`/insert/delete paths).

---

## Roadmap

- [x] Standalone Kimai CSV → Excel timesheet updater (this repo)
- [ ] **Integrate into the OrbisDS platform** — Kimai read/write becomes a first-class
      part of the platform's Timesheet module, replacing this standalone script
- [ ] Push completed timesheets back to Kimai
- [ ] Per-project / per-client reporting on top of the team budget summary

---

## Privacy / sample data

`examples/timesheets/` contains **only an anonymized fake-company template**
(`Acme Studio`, workers `John A. Doe` etc.). No real OrbisDS names, employees,
clients or hours are committed. The real templates stay offline; regenerate the
example from one with `scripts/build_fake_template.py <real.xlsx>`.

---

## Author

Built and maintained by **OrbisDS** (`OrbisDS-Corp`).
