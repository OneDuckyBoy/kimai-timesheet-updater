"""Build an EMPTY fake-company timesheet template from a real one's structure.

Keeps the real template's formatting/layout (borders, fills, widths, sheet
setup) but strips all real data and real employee names, replacing them with
a fake company + obviously-fake worker names. Result ships in the repo as the
example, so no real personnel/company data is published.
"""
import argparse
import os
import re

import openpyxl

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
parser = argparse.ArgumentParser(
    description="Anonymize a real timesheet template into the fake example.")
parser.add_argument("source", help="Path to a REAL timesheet template (.xlsx) to strip data from")
parser.add_argument("--out", default=os.path.join(REPO_ROOT, "examples/timesheets/Acme - Timesheet Report October 2026.xlsx"),
                    help="Output path for the fake template (default: the repo example)")
args = parser.parse_args()

SRC = args.source
DST = args.out

FAKE_COMPANY = "Acme Studio"
# Obviously fake worker names (same count as the source's person sheets).
FAKE_WORKERS = ["John A. Doe", "Jane B. Smith", "Sam C. Johnson",
                "Alex D. Brown", "Maria E. White", "Chris F. Wilson"]

wb = openpyxl.load_workbook(SRC)

person_sheets = [n for n in wb.sheetnames if n.startswith("Timesheet")]
assert len(person_sheets) <= len(FAKE_WORKERS), "more person sheets than fake names"

MONTHS = "(January|February|March|April|May|June|July|August|September|October|November|December)"
TITLE_RE = re.compile(r"^(.*?Timesheet Report )(%s \d{4})(.*)$" % MONTHS)

def blank_cell(ws, coord):
    c = ws[coord]
    if c.value is not None:
        try:
            c.value = None
        except Exception:
            pass

for idx, sheet_name in enumerate(person_sheets):
    ws = wb[sheet_name]
    fake = FAKE_WORKERS[idx]
    print(f"   processing {sheet_name!r} -> fake {fake!r}")

    # 1. Blank the company from the title, keep the month/year.
    for r in range(1, 5):
        for col in ("A", "B"):
            cell = ws[f"{col}{r}"]
            if cell.value and "Timesheet Report" in str(cell.value):
                m = TITLE_RE.match(str(cell.value))
                base = f"{m.group(1)}{m.group(2)}" if m else "Timesheet Report October 2026"
                cell.value = f"{FAKE_COMPANY} - {base}"
    # 1b. Blank client / project metadata (leaks the real client & project).
    for coord in ("B3", "C3", "B4", "C4", "B5", "C5"):
        blank_cell(ws, coord)

    # 2. Blank all data cells (dates/activity/hours) but keep the worker's
    #    name in the first data row so name-matching + budget still resolve.
    #    Layout detection mirrors detect_sheet_format() in update_timesheet.py.
    header_row = None
    for hr in (6, 7):
        for col in ("A", "B"):
            if ws[f"{col}{hr}"].value and "Timesheet date" in str(ws[f"{col}{hr}"].value):
                header_row = hr
                if col == "A":
                    col_date, col_name, col_activity, col_hours = "A", "B", "C", "D"
                else:
                    col_date, col_name, col_activity, col_hours = "B", "C", "D", "E"
                break
        if header_row:
            break
    if header_row is None:
        header_row = 7  # fallback: Format B, header row 7
        col_date, col_name, col_activity, col_hours = "A", "B", "C", "D"

    # Find the "Total hours" row.
    total_row = None
    for row in range(header_row + 1, 200):
        if ws[f"{col_activity}{row}"].value and "Total hours" in str(ws[f"{col_activity}{row}"].value):
            total_row = row
            break
    if total_row is None:
        continue

    for row in range(header_row + 1, total_row):
        blank_cell(ws, f"{col_date}{row}")
        blank_cell(ws, f"{col_activity}{row}")
        blank_cell(ws, f"{col_hours}{row}")
        if row > header_row + 1:
            blank_cell(ws, f"{col_name}{row}")
    # First data row keeps the new (fake) worker name.
    ws[f"{col_name}{header_row + 1}"] = fake

    # 3. Rename the sheet to the fake worker.
    ws.title = f"Timesheet {fake}"
    print(f"   renamed {sheet_name!r} -> {ws.title!r}")

# Blank the Team Budget Report: data rows keep FAKE worker names, hours blanked,
# and client/project metadata + cross-sheet formulas removed.
budget = [n for n in wb.sheetnames if "Budget" in n]
if budget:
    ws = wb[budget[0]]

    # 1. Blank client / project metadata header area.
    for coord in ("A3", "B3", "A4", "B4", "A5", "B5"):
        blank_cell(ws, coord)

    # 2. Title -> fake company.
    for r in range(1, 5):
        for col in ("A", "B"):
            cell = ws[f"{col}{r}"]
            if cell.value and "Timesheet Report" in str(cell.value):
                cell.value = f"{FAKE_COMPANY} - Timesheet Report October 2026"

    first_data_row = None
    for row in range(1, 50):
        if ws[f"B{row}"].value and str(ws[f"B{row}"].value).strip() == "Worker name":
            first_data_row = row + 1
            break
    if first_data_row:
        wi = 0
        for row in range(first_data_row, 60):
            if not ws[f"B{row}"].value:
                break
            blank_cell(ws, f"A{row}")                      # date range -> empty
            if wi < len(FAKE_WORKERS):
                ws[f"B{row}"] = FAKE_WORKERS[wi]           # name -> fake, plain text
                wi += 1
            blank_cell(ws, f"D{row}")                      # hours -> empty
        # blank the totals row too
        for row in range(first_data_row, 60):
            if ws[f"C{row}"].value and "Total" in str(ws[f"C{row}"].value):
                blank_cell(ws, f"A{row}")
                blank_cell(ws, f"D{row}")
                break

wb.save(DST)
print("wrote", DST)
print("sheets:", wb.sheetnames)