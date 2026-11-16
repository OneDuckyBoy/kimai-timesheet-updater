#!/usr/bin/env python3
"""
Kimai → Excel Timesheet Updater
================================
Reads a Kimai CSV export and updates an Excel timesheet template for the next calendar month.

Usage:
    python update_timesheet.py <excel_template> <kimai_csv> [--output OUTPUT]

Example:
    python update_timesheet.py "Timesheet Report Stella Care June 2026.xlsx" kimai_export.csv

What it does:
    1. Detects the current month from the Excel title (e.g. "June 2026")
    2. Advances to the next calendar month (e.g. "July 2026")
    3. For each person's sheet:
       - Clears old daily entries (keeps the "Total hours" row)
       - Fills new entries from the Kimai CSV (grouped by day)
       - Updates the SUBTOTAL formula
       - Updates the title to the new month
    4. Updates the Team Budget Report with new totals
    5. Saves as a new Excel file
"""

import argparse
import csv
import calendar
import os
import re
import sys
from collections import defaultdict
from datetime import datetime, timedelta
from typing import Optional

import openpyxl
from openpyxl.utils import get_column_letter


# ── Excel Loading (supports .xls and .xlsx) ──────────────────────────────────

def load_workbook_any(path: str):
    """
    Load an Excel file, supporting both .xls and .xlsx.
    Returns an openpyxl Workbook.
    """
    ext = os.path.splitext(path)[1].lower()

    if ext == ".xlsx":
        return openpyxl.load_workbook(path)

    if ext == ".xls":
        try:
            import xlrd
        except ImportError:
            print("ERROR: xlrd not installed. Run: pip install xlrd==1.2.0")
            sys.exit(1)

        print("   Converting .xls → .xlsx (old format detected)...")
        xls = xlrd.open_workbook(path, formatting_info=True)
        wb = openpyxl.Workbook()
        # Remove default sheet
        wb.remove(wb.active)

        from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
        from openpyxl.utils import get_column_letter

        # Excel border style names → openpyxl side styles
        _BORDER_STYLES = {
            0: None,  # none
            1: "thin", 2: "medium", 3: "dashed", 4: "dotted",
            5: "thick", 6: "double", 7: "hair", 8: "mediumDashed",
            9: "dashDot", 10: "mediumDashDot", 11: "dashDotDot",
            12: "mediumDashDotDot", 13: "slantDashDot",
        }

        def _map_colour(colour_map, cidx):
            """Map an xlrd palette index to an RGB hex string (or None)."""
            if cidx is None or cidx < 0:
                return None
            rgb = colour_map.get(cidx)
            if rgb is None:
                return None
            # xlrd colour_map values are (r, g, b) or (r, g, b, a) tuples
            r, g, b = rgb[0], rgb[1], rgb[2]
            return f"{r:02X}{g:02X}{b:02X}"

        for sheet_name in xls.sheet_names():
            xs = xls.sheet_by_name(sheet_name)
            ws = wb.create_sheet(title=sheet_name)

            # ── Copy column widths ─────────────────────────────────────
            for colx, cinfo in xs.colinfo_map.items():
                if cinfo.width:
                    width_chars = cinfo.width / 256.0
                    ws.column_dimensions[get_column_letter(colx + 1)].width = width_chars

            # ── Copy cell values AND full styles ───────────────────────
            for row_idx in range(xs.nrows):
                for col_idx in range(xs.ncols):
                    xcell = xs.cell(row_idx, col_idx)
                    val = xcell.value
                    # Convert XLDate to datetime
                    if xcell.ctype == 3:  # XL_CELL_DATE
                        try:
                            date_tuple = xlrd.xldate_as_tuple(val, xls.datemode)
                        except Exception:
                            date_tuple = None
                        if date_tuple:
                            if date_tuple[:3] == (0, 0, 0):  # time only
                                val = timedelta(
                                    hours=date_tuple[3],
                                    minutes=date_tuple[4],
                                    seconds=date_tuple[5]
                                )
                            else:
                                val = datetime(*date_tuple)
                    elif xcell.ctype == 2 and isinstance(val, float) and val == int(val):
                        # keep true integers as ints (matches original .xlsx)
                        val = int(val)

                    cell = ws.cell(row=row_idx + 1, column=col_idx + 1, value=val)

                    xf_idx = xs.cell_xf_index(row_idx, col_idx)
                    xf = xls.xf_list[xf_idx]
                    font = xls.font_list[xf.font_index]

                    # Font
                    cell.font = Font(
                        name=font.name,
                        bold=bool(font.bold),
                        italic=bool(font.italic),
                        underline=("single" if font.underline_type == 1 else None),
                        size=(font.height / 20.0) if font.height else None,
                        color=(_map_colour(xls.colour_map, font.colour_index)),
                    )

                    # Number format
                    if xf.format_key in xls.format_map:
                        _nf = xls.format_map[xf.format_key].format_str or "General"
                    else:
                        _nf = "General"
                    cell.number_format = _nf

                    # Border
                    cell.border = Border(
                        left=Side(style=_BORDER_STYLES.get(xf.border.left_line_style)),
                        right=Side(style=_BORDER_STYLES.get(xf.border.right_line_style)),
                        top=Side(style=_BORDER_STYLES.get(xf.border.top_line_style)),
                        bottom=Side(style=_BORDER_STYLES.get(xf.border.bottom_line_style)),
                    )

                    # Alignment
                    from openpyxl.styles import Alignment as _Align
                    _HOR = {0: "general", 1: "left", 2: "center", 3: "right",
                            4: "fill", 5: "justify", 6: "centerContinuous", 7: "distributed"}
                    _VER = {0: "top", 1: "center", 2: "bottom", 3: "justify", 4: "distributed"}
                    cell.alignment = _Align(
                        horizontal=_HOR.get(xf.alignment.hor_align, "general"),
                        vertical=_VER.get(xf.alignment.vert_align, "bottom"),
                        wrap_text=bool(xf.alignment.text_wrapped),
                    )

                    # Fill (background colour / pattern)
                    bg = xf.background
                    fg_colour = _map_colour(xls.colour_map, bg.pattern_colour_index)
                    if bg.fill_pattern and fg_colour:
                        cell.fill = PatternFill(
                            patternType="solid",
                            fgColor=fg_colour,
                        )

            # Copy merged cells
            for merge in xs.merged_cells:
                rlo, rhi, clo, chi = merge
                ws.merge_cells(
                    start_row=rlo + 1, start_column=clo + 1,
                    end_row=rhi, end_column=chi
                )

        return wb

    print(f"ERROR: Unsupported file format: {ext}. Use .xls or .xlsx")
    sys.exit(1)


# ── Time Formatting ─────────────────────────────────────────────────────────

def format_hours_cell(cell, hours: float, keep_num_fmt: bool = True):
    """
    Write hours as a plain decimal number (e.g. 8 -> displayed '8.00'),
    matching the template's number format. Preserves an existing real
    number format; falls back to '0.00' only if the cell has none yet.
    """
    cell.value = round(float(hours), 2)
    if keep_num_fmt:
        if not cell.number_format or cell.number_format == "General":
            cell.number_format = "0.00"
    else:
        cell.number_format = "0.00"


# ── Table Styling ────────────────────────────────────────────────────────────

def apply_table_style(ws, fmt: dict, first_data_row: int, last_data_row: int,
                      total_row: int, num_entries: int):
    """
    Apply borders to a person's timesheet.
    Column widths are left as-is (from the template) — nothing is resized.
    """
    from openpyxl.styles import Border, Side, Alignment, Font

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin'),
    )
    header_font = Font(bold=True)
    wrap_align = Alignment(wrap_text=True, vertical='center')

    # Columns in order: date, name, activity, hours
    cols = [fmt["col_date"], fmt["col_name"], fmt["col_activity"], fmt["col_hours"]]

    # ── Apply borders ───────────────────────────────────────────────────
    # Header row
    for col in cols:
        cell = ws[f"{col}{fmt['header_row']}"]
        cell.border = thin_border
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Data rows
    for r in range(first_data_row, first_data_row + num_entries):
        for col in cols:
            cell = ws[f"{col}{r}"]
            cell.border = thin_border
            if col == fmt["col_activity"]:
                cell.alignment = wrap_align
            else:
                cell.alignment = Alignment(vertical='bottom')

    # Total row
    total_label_col = fmt["col_activity"]  # "Total hours" is in activity column
    ws[f"{total_label_col}{total_row}"].border = thin_border
    ws[f"{total_label_col}{total_row}"].font = Font(bold=True)
    ws[f"{fmt['col_hours']}{total_row}"].border = thin_border
    ws[f"{fmt['col_hours']}{total_row}"].font = Font(bold=True)


# ── Constants ────────────────────────────────────────────────────────────────

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

BG_MONTHS = {
    "January": "януари", "February": "февруари", "March": "март",
    "April": "април", "May": "май", "June": "юни",
    "July": "юли", "August": "август", "September": "септември",
    "October": "октомври", "November": "ноември", "December": "декември"
}

# ── CSV Parsing ──────────────────────────────────────────────────────────────

def parse_duration(duration_str: str) -> float:
    """Convert '1:30' or '0:45' to decimal hours."""
    parts = duration_str.strip().split(":")
    return int(parts[0]) + int(parts[1]) / 60


def parse_kimai_csv(csv_path: str) -> dict[str, list[dict]]:
    """
    Parse Kimai CSV export.
    Returns: { "Full Name": [ {date, hours, description}, ... ] }
    Entries are NOT grouped by date yet — that happens per-sheet.
    """
    data: dict[str, list[dict]] = defaultdict(list)

    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = row.get("Name", "").strip()
            if not name:
                continue

            date_str = row.get("Date", "").strip()
            dur_str = row.get("Duration", "0:00").strip()
            desc = row.get("Description", "").strip()

            try:
                date = datetime.strptime(date_str, "%Y-%m-%d")
                hours = parse_duration(dur_str)
            except (ValueError, IndexError):
                continue

            data[name].append({
                "date": date,
                "hours": hours,
                "description": desc
            })

    return dict(data)


def csv_client_project(csv_path: str) -> tuple[str, str]:
    """
    The export's customer and project (the `Customer` and `Project` columns), which the
    sheets show as Client / Project. Taken from the first row that carries any; a Kimai
    export's rows share one project, so the values are the same no matter the row.
    """
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            client = (row.get("Customer") or "").strip()
            project = (row.get("Project") or "").strip()
            if client or project:
                return client, project
    return "", ""


def group_by_date(entries: list[dict]) -> list[dict]:
    """
    Group entries by date: sum hours, concatenate descriptions with '; '.
    Returns sorted list of {date, hours, description}.
    """
    grouped: dict[datetime, dict] = defaultdict(lambda: {"hours": 0.0, "descriptions": []})

    for entry in entries:
        d = entry["date"]
        grouped[d]["hours"] += entry["hours"]
        grouped[d]["descriptions"].append(entry["description"])

    result = []
    for date in sorted(grouped.keys()):
        result.append({
            "date": date,
            "hours": round(grouped[date]["hours"], 2),
            "description": "; ".join(filter(None, grouped[date]["descriptions"]))
        })
    return result


# ── Month Helpers ────────────────────────────────────────────────────────────

def extract_month_year(text: str) -> Optional[tuple[str, int]]:
    """
    Extract (month_name, year) from a title like 'Timesheet Report June 2026'.
    """
    pattern = r"(" + "|".join(MONTHS) + r")\s+(\d{4})"
    match = re.search(pattern, text)
    if match:
        return match.group(1), int(match.group(2))
    return None


def next_month(month_name: str, year: int) -> tuple[str, int]:
    """Return the next calendar month."""
    idx = MONTHS.index(month_name)
    next_idx = (idx + 1) % 12
    next_year = year + 1 if next_idx == 0 else year
    return MONTHS[next_idx], next_year


def month_date_range(month_name: str, year: int) -> tuple[datetime, datetime, str]:
    """
    Return (first_day, last_day, display_string).
    Display string example: '1.07.2026 - 31.07.2026'
    """
    idx = MONTHS.index(month_name)
    first = datetime(year, idx + 1, 1)
    last_day = calendar.monthrange(year, idx + 1)[1]
    last = datetime(year, idx + 1, last_day)

    display = f"{first.day}.{first.month:02d}.{first.year} - {last.day}.{last.month:02d}.{last.year}"
    return first, last, display


def replace_month_in_text(text: str, old_month: str, old_year: int,
                          new_month: str, new_year: int) -> str:
    """Replace all occurrences of the old month/year with the new one."""
    text = text.replace(f"{old_month} {old_year}", f"{new_month} {new_year}")
    return text


def csv_target_month(kimai_data: dict[str, list[dict]]) -> Optional[tuple[str, int]]:
    """
    Determine the target (month, year) from the CSV data itself — the month
    with the most entries. Returns None if the CSV has no entries.
    """
    counts: dict[tuple[int, int], int] = {}
    for entries in kimai_data.values():
        for entry in entries:
            key = (entry["date"].month, entry["date"].year)
            counts[key] = counts.get(key, 0) + 1
    if not counts:
        return None
    (month_num, year), _ = max(counts.items(), key=lambda kv: kv[1])
    return MONTHS[month_num - 1], year


# ── Sheet Analysis ───────────────────────────────────────────────────────────

def detect_sheet_format(ws) -> Optional[dict]:
    """
    Detect whether a person's sheet uses Format A (B-E columns) or Format B (A-D columns).

    Returns a dict with keys describing the layout, or None if not a person sheet.
    """
    # Check rows 6-7 for "Timesheet date" header
    for check_row in [6, 7]:
        for col_idx, col_letter in enumerate(["A", "B"], start=1):
            cell = ws[f"{col_letter}{check_row}"]
            if cell.value and "Timesheet date" in str(cell.value):
                if col_letter == "A":
                    # Format B: A=Date, B=Name, C=Activity, D=Hours
                    return {
                        "format": "B",
                        "header_row": check_row,
                        "col_date": "A",
                        "col_name": "B",
                        "col_activity": "C",
                        "col_hours": "D",
                        "date_type": "string",  # Numbers export gives strings
                    }
                else:
                    # Format A: B=Date, C=Name, D=Activity, E=Hours
                    return {
                        "format": "A",
                        "header_row": check_row,
                        "col_date": "B",
                        "col_name": "C",
                        "col_activity": "D",
                        "col_hours": "E",
                        "date_type": "datetime",
                    }
    return None


def find_total_row(ws, fmt: dict) -> Optional[int]:
    """Find the row number containing 'Total hours'."""
    col = fmt["col_activity"]  # "Total hours" is in the activity column
    for row in range(fmt["header_row"] + 1, 200):
        cell = ws[f"{col}{row}"]
        if cell.value and "Total hours" in str(cell.value):
            return row
    return None


def find_title_info(ws) -> Optional[dict]:
    """
    Find the title cell and extract month/year.
    Returns {"cell": "B2", "text": "Timesheet Report June 2026", "month": "June", "year": 2026}
    or None.
    """
    for row in range(1, 5):
        for col in ["A", "B"]:
            cell = ws[f"{col}{row}"]
            if cell.value and ("Timesheet Report" in str(cell.value) or "Team Budget Report" in str(cell.value)):
                month_info = extract_month_year(str(cell.value))
                if month_info:
                    return {
                        "cell": cell.coordinate,
                        "text": str(cell.value),
                        "month": month_info[0],
                        "year": month_info[1],
                    }
    return None


# ── Updating a Person's Sheet ────────────────────────────────────────────────

def _write_label_value(ws, label: str, value: str) -> None:
    """Write a client/project value into the cells a sheet labels, whichever layout it has.
    A person sheet keeps `Client`/`Project` in column A and the value in B; a summary joins
    them in one cell (`Client All Digital`). Both are handled, so the shared template's own
    placeholder never leaks into a different project's file."""
    if not value:
        return
    for row in range(1, 6):
        a = ws[f"A{row}"].value
        if not a:
            continue
        text = str(a).strip()
        if text.lower() == label.lower():
            ws[f"B{row}"] = value
            return
        if text.lower().startswith(label.lower() + " "):
            ws[f"A{row}"] = f"{label} {value}"
            return


def update_person_sheet_v2(ws, fmt: dict, worker_name: str,
                           entries_by_date: list[dict],
                           new_month: str, new_year: int,
                           old_month: str, old_year: int,
                           client: str = "", project: str = "") -> None:
    """
    Update a single person's timesheet.
    v2: worker_name passed explicitly.
    """
    header_row = fmt["header_row"]
    first_data_row = header_row + 1
    total_row = find_total_row(ws, fmt)
    if not total_row:
        print(f"  ⚠ Could not find 'Total hours' row")
        return

    # ── Update title ─────────────────────────────────────────────────────
    title_info = find_title_info(ws)
    if title_info:
        new_title = replace_month_in_text(
            title_info["text"], old_month, old_year, new_month, new_year
        )
        ws[title_info["cell"]] = new_title

    # ── Client / Project from the export, never the template's placeholder ──
    _write_label_value(ws, "Client", client)
    _write_label_value(ws, "Project", project)

    # ── Capture the template's cell styles BEFORE modifying anything ────
    # We copy style from the first data row of the template so that rows
    # created by insert_rows() (which start with NO border/fill) get the
    # exact same look as the template's existing rows.
    from copy import copy as _copy_style
    _style_cols = [fmt["col_date"], fmt["col_name"], fmt["col_activity"], fmt["col_hours"]]
    _ref_cells = {c: ws[f"{c}{first_data_row}"] for c in _style_cols}
    # Date column is displayed as  ден.месец.година (dd.mm.yyyy), not US mm-dd-yy
    DATE_NUM_FMT = "dd.mm.yyyy"

    def _apply_template_style(r):
        """Copy border/fill/font/alignment/number_format from the template's
        first data row onto row r (only for cells that have a template style)."""
        for c in _style_cols:
            src = _ref_cells[c]
            if not src.has_style:
                continue
            dst = ws[f"{c}{r}"]
            dst.font = _copy_style(src.font)
            dst.border = _copy_style(src.border)
            dst.fill = _copy_style(src.fill)
            dst.alignment = _copy_style(src.alignment)
            # dates -> dd.mm.yyyy; everything else keeps the template format
            dst.number_format = DATE_NUM_FMT if c == fmt["col_date"] else src.number_format

    # ── Clear old data rows ──────────────────────────────────────────────
    for r in range(first_data_row, total_row):
        for col in [fmt["col_date"], fmt["col_name"], fmt["col_activity"], fmt["col_hours"]]:
            ws[f"{col}{r}"] = None

    # Always write the worker's name in the first data row so the summary
    # can pull it live, even when the person has no entries this month.
    ws[f"{fmt['col_name']}{first_data_row}"] = worker_name

    # ── Calculate how many data rows we need ─────────────────────────────
    num_needed = max(len(entries_by_date), 0)
    num_existing = (total_row - 1) - first_data_row  # rows available (excluding blank row)

    # If we need more rows, insert them before the total row
    if num_needed > num_existing:
        rows_to_insert = num_needed - num_existing
        # Insert above the blank row (total_row - 1)
        ws.insert_rows(total_row - 1, rows_to_insert)
        # total_row has shifted — re-find it
        total_row = find_total_row(ws, fmt)
        if not total_row:
            print(f"  ⚠ Lost 'Total hours' row after insert")
            return

    # ── Write data ───────────────────────────────────────────────────────
    # IMPORTANT: we only write VALUES. The template cells already carry the
    # correct style (borders, fills, number formats, widths). openpyxl keeps
    # a cell's existing style when you set .value, so writing here preserves
    # the template's formatting exactly — we do NOT re-apply any styling.
    for i, entry in enumerate(entries_by_date):
        r = first_data_row + i
        ws[f"{fmt['col_date']}{r}"] = entry["date"]
        ws[f"{fmt['col_name']}{r}"] = worker_name
        ws[f"{fmt['col_activity']}{r}"] = entry["description"]
        format_hours_cell(ws[f"{fmt['col_hours']}{r}"], entry["hours"])

    # ── Update formula ───────────────────────────────────────────────────
    # SUBTOTAL(9, ...) sums only numeric cells (ignores text). The range is
    # DYNAMIC: from the first data row down to the row just above this total
    # row (ROW()-1), so rows the user adds manually are included too.
    hours_col = fmt["col_hours"]
    formula = f"=SUBTOTAL(9,{hours_col}{first_data_row}:INDEX({hours_col}:{hours_col},ROW()-1))"
    if num_needed > 0:
        last_data_row = first_data_row + num_needed - 1
    else:
        last_data_row = first_data_row - 1  # signal "no data"
    total_cell = ws[f"{hours_col}{total_row}"]
    total_orig_nf = total_cell.number_format
    total_cell.value = formula
    # keep the template's own number format (decimal hours), not '[h]:mm'
    if not total_orig_nf or total_orig_nf == "General":
        total_cell.number_format = "0.00"

    # ── Delete extra empty rows (keep one blank row between data and total) ──
    blank_row = last_data_row + 1 if num_needed > 0 else first_data_row
    # Delete rows from blank_row+1 up to total_row-1
    if blank_row + 1 <= total_row - 1:
        num_to_delete = (total_row - 1) - blank_row
        ws.delete_rows(blank_row + 1, num_to_delete)
        # Re-find total row (it shifted up)
        total_row = find_total_row(ws, fmt)

    # ── Apply the template's style to every data row ────────────────────
    # This is the critical fix: insert_rows() creates rows with NO border or
    # fill. We stamp the template's border/fill/font/alignment/number_format
    # onto every data row so new rows match the originals exactly.
    if num_needed > 0:
        for r in range(first_data_row, first_data_row + num_needed):
            _apply_template_style(r)

    return last_data_row if num_needed > 0 else None


# ── Team Budget Report ───────────────────────────────────────────────────────

def find_budget_sheet(wb) -> Optional[str]:
    """
    Find the budget/summary sheet. Looks for:
    1. Sheet name containing 'Budget'
    2. Sheet with 'Actual Hours' header (the summary table)
    """
    for name in wb.sheetnames:
        if "budget" in name.lower():
            return name

    # Fallback: look for sheet with "Actual Hours" header
    for name in wb.sheetnames:
        if name.startswith("Timesheet") or name == "Export Summary":
            continue
        ws = wb[name]
        for row in ws.iter_rows(min_row=1, max_row=10, max_col=5, values_only=True):
            for val in row:
                if val and "Actual Hours" in str(val):
                    return name

    return None


def update_budget_report(wb, kimai_data: dict[str, list[dict]],
                         name_to_sheet: dict[str, str],
                         sheet_totals: dict[str, float],
                         new_month: str, new_year: int,
                         old_month: str, old_year: int,
                         client: str = "", project: str = "") -> None:
    """
    Rebuild the Team Budget Report sheet's data rows from the CSV people:
    one row per person with live INDEX/MATCH formulas into their sheet, a
    dynamic SUBTOTAL and the new date range. The template's placeholder
    person rows are replaced, never carried over — the summary lists exactly
    who is in the CSV.
    """
    budget_name = find_budget_sheet(wb)
    if not budget_name:
        print("  ⚠ Could not find budget/summary sheet")
        return
    ws = wb[budget_name]

    # ── Update title ─────────────────────────────────────────────────────
    for row in range(1, 6):
        for col in ["A", "B"]:
            cell = ws[f"{col}{row}"]
            if cell.value:
                month_info = extract_month_year(str(cell.value))
                if month_info:
                    cell.value = replace_month_in_text(
                        str(cell.value), old_month, old_year, new_month, new_year
                    )
                    break

    _, _, date_display = month_date_range(new_month, new_year)

    # ── Client / Project from the export, never the template's placeholder ──
    _write_label_value(ws, "Client", client)
    _write_label_value(ws, "Project", project)

    # ── Locate the header, the data area and the total row ───────────────
    first_data_row = None
    for row in range(1, 30):
        b_val = ws[f"B{row}"].value
        if b_val and str(b_val).strip() == "Worker name":
            first_data_row = row + 1
            break
    if not first_data_row:
        print("  ⚠ Could not find data in Team Budget Report")
        return

    total_row = None
    for row in range(first_data_row + 1, 60):
        c_val = ws[f"C{row}"].value
        if c_val and "Total" in str(c_val):
            total_row = row
            break
    if not total_row:
        print("  ⚠ Could not find 'Total hours' row in Team Budget Report")
        return

    # ── Drop the template's placeholder people rows ──────────────────────
    if total_row - 1 > first_data_row:
        ws.delete_rows(first_data_row, total_row - 1 - first_data_row)
    # Re-locate the total row after the delete (it slid up by the deleted count).
    total_row = None
    for row in range(first_data_row, 60):
        c_val = ws[f"C{row}"].value
        if c_val and "Total" in str(c_val):
            total_row = row
            break
    if not total_row:
        print("  ⚠ Could not find 'Total hours' row in Team Budget Report (after delete)")
        return

    from openpyxl.styles import Border, Side, Font
    thin = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    bold = Font(bold=True)

    people = [n for n in kimai_data if name_to_sheet.get(n) in wb.sheetnames]
    if people:
        ws.insert_rows(total_row, len(people))
        total_row += len(people)

    # ── One row per CSV person; hours read live from their own sheet ─────
    for i, csv_name in enumerate(people):
        row = first_data_row + i
        sheet_title = name_to_sheet[csv_name]
        # Role is left blank — the PM fills it in when checking.
        ws[f"A{row}"] = date_display

        person_sheet = wb[sheet_title]
        pfmt = detect_sheet_format(person_sheet)
        if pfmt and find_total_row(person_sheet, pfmt):
            quoted = f"'{sheet_title}'"
            # Name comes live from the person's sheet.
            pname_cell = f"{pfmt['col_name']}{pfmt['header_row'] + 1}"
            ws[f"B{row}"] = f"={quoted}!{pname_cell}"
            # Hours come live from the person's "Total hours" row.
            act_col = pfmt["col_activity"]
            hrs_col = pfmt["col_hours"]
            d_cell = ws[f"D{row}"]
            d_cell.value = (
                f"=INDEX({quoted}!{hrs_col}:{hrs_col}, "
                f"MATCH(\"Total hours\", {quoted}!{act_col}:{act_col}, 0))"
            )
            if not d_cell.number_format or d_cell.number_format == "General":
                d_cell.number_format = "0.00"
        else:
            # No computable total — keep the person visible as a plain zero
            # so a PM spots and fixes the sheet when checking.
            ws[f"B{row}"] = csv_name
            format_hours_cell(ws[f"D{row}"], 0)

        for col in ["A", "B", "C", "D"]:
            ws[f"{col}{row}"].border = thin

    last_data_row = first_data_row + len(people) - 1 if people else first_data_row

    # ── SUBTOTAL formula (dynamic, sums the rows above it) ───────────────
    total_cell = ws[f"D{total_row}"]
    total_orig_nf = total_cell.number_format
    total_cell.value = f"=SUBTOTAL(9,D{first_data_row}:INDEX(D:D,ROW()-1))"
    if not total_orig_nf or total_orig_nf == "General":
        total_cell.number_format = "0.00"

    # ── Borders: header, data, total ─────────────────────────────────────
    for col in ["A", "B", "C", "D"]:
        hdr = ws[f"{col}{first_data_row - 1}"]
        hdr.border = thin
        hdr.font = bold
    for r in range(first_data_row, last_data_row + 1):
        for col in ["A", "B", "C", "D"]:
            ws[f"{col}{r}"].border = thin
    ws[f"C{total_row}"].border = thin
    ws[f"C{total_row}"].font = bold
    ws[f"D{total_row}"].border = thin
    ws[f"D{total_row}"].font = bold


# ── Person-sheet naming (CSV drives the people) ─────────────────────────────

def _person_sheet_title(name: str) -> str:
    """A safe Excel sheet title for one person: 'Timesheet <name>', <=31 chars."""
    title = f"Timesheet {name}"
    for ch in "[]:*?/\\":
        title = title.replace(ch, " ")
    title = re.sub(r"\s+", " ", title).strip()
    return title[:31]


def _unique_sheet_title(wb, name: str) -> str:
    """A person sheet title that no sheet in `wb` already has."""
    base = _person_sheet_title(name)
    title = base
    n = 2
    while title in wb.sheetnames:
        suffix = f"-{n}"
        title = f"{base[:31 - len(suffix)]}{suffix}"
        n += 1
    return title


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    # Make stdout/stderr encoding-safe for the frozen .exe regardless of the
    # console code page (cp1252 would crash on the emoji in print()).
    for _stream_name in ("stdout", "stderr"):
        _s = getattr(sys, _stream_name, None)
        if _s is not None:
            try:
                _s.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

    parser = argparse.ArgumentParser(
        description="Update Excel timesheet from Kimai CSV export"
    )
    parser.add_argument("excel", help="Path to the Excel template (.xlsx)")
    parser.add_argument("csv", help="Path to the Kimai CSV export")
    parser.add_argument("--output", "-o", help="Output Excel file path (auto-generated if omitted)")
    parser.add_argument("--outdir", help="Directory for the auto-generated output file (default: 'output' next to the input)")
    args = parser.parse_args()

    # ── Load files ───────────────────────────────────────────────────────
    if not os.path.exists(args.excel):
        print(f"ERROR: Excel file not found: {args.excel}")
        sys.exit(1)
    if not os.path.exists(args.csv):
        print(f"ERROR: CSV file not found: {args.csv}")
        sys.exit(1)

    print(f"📂 Loading Excel: {args.excel}")
    wb = load_workbook_any(args.excel)

    print(f"📂 Loading CSV:   {args.csv}")
    kimai_data = parse_kimai_csv(args.csv)
    print(f"   Found {len(kimai_data)} people in CSV")
    client, project = csv_client_project(args.csv)

    # ── Detect month ─────────────────────────────────────────────────────
    old_month, old_year = None, None
    # Try to extract from any sheet
    for sheet_name in wb.sheetnames:
        if sheet_name.startswith("Timesheet"):
            ws = wb[sheet_name]
            month_info = extract_month_year(sheet_name)
            if not month_info:
                title_info = find_title_info(ws)
                if title_info:
                    old_month, old_year = title_info["month"], title_info["year"]
                    break
            else:
                old_month, old_year = month_info
                break

    if not old_month:
        # Try budget/summary sheet
        budget_name = find_budget_sheet(wb)
        if budget_name:
            ws = wb[budget_name]
            for row in range(1, 5):
                for col in ["A", "B"]:
                    cell = ws[f"{col}{row}"]
                    if cell.value:
                        month_info = extract_month_year(str(cell.value))
                        if month_info:
                            old_month, old_year = month_info
                            break

    if not old_month:
        print("ERROR: Could not detect current month in Excel file")
        sys.exit(1)

    # ── Determine the target month ──────────────────────────────────────
    # Prefer the month/year present in the CSV data (the month with the
    # most entries) — each CSV export covers a single month, so the target
    # month should be whichever month the data belongs to, not blindly "the
    # month after the template". Only falls back to next-month-after-the-
    # template when the CSV has no entries.
    csv_target = csv_target_month(kimai_data)
    if csv_target:
        new_month, new_year = csv_target
    else:
        new_month, new_year = next_month(old_month, old_year)
    print(f"\n📅 Current month (template): {old_month} {old_year}")
    print(f"📅 Target month (from CSV): {new_month} {new_year}")

    # ── The CSV is the only source of people ─────────────────────────────
    # The template is minimal: one summary sheet + one placeholder person
    # sheet. Every person in the CSV gets a clone of that placeholder sheet,
    # filled from their own entries; the placeholder itself is removed from
    # the result so only real people remain.
    placeholder = next(
        (sn for sn in wb.sheetnames if sn.lower().startswith("placeholder")),
        None,
    )
    if placeholder is None:
        print(
            "ERROR: template has no 'placeholder timesheet' sheet. Expected a "
            "minimal template (summary + one empty placeholder person sheet)."
        )
        sys.exit(1)
    ph_ws = wb[placeholder]
    ph_fmt = detect_sheet_format(ph_ws)
    if not ph_fmt:
        print(f"ERROR: placeholder sheet '{placeholder}' does not look like a person sheet")
        sys.exit(1)

    print(f"\n👥 CSV people → cloned person sheets:")
    created = {}
    for csv_name, entries in kimai_data.items():
        sheet_title = _unique_sheet_title(wb, csv_name)
        new_ws = wb.copy_worksheet(ph_ws)
        new_ws.title = sheet_title
        grouped = group_by_date(entries)
        update_person_sheet_v2(
            new_ws, ph_fmt, csv_name, grouped,
            new_month, new_year, old_month, old_year,
            client, project,
        )
        total = sum(e["hours"] for e in grouped)
        has = "✓" if grouped else "✗ (0 hours)"
        print(f"   {has} {csv_name}: {len(grouped)} days, {total:.1f}h → {sheet_title}")
        created[csv_name] = sheet_title

    # The placeholder is scaffolding only — it never belongs in the result.
    wb.remove(wb[placeholder])
    print(f"   Removed placeholder sheet: {placeholder}")

    # ── Update Team Budget Report for the CSV people ─────────────────────
    print(f"\n📊 Updating Team Budget Report...")
    update_budget_report(
        wb, kimai_data, created, {},
        new_month, new_year, old_month, old_year,
        client, project,
    )
    print(f"   Done")

    # ── Save ─────────────────────────────────────────────────────────────
    if args.output:
        output_path = args.output
    else:
        # Auto-generate: name the output like the template but with the new
        # month, and place it in --outdir (or an "output" folder next to the
        # input when --outdir is not given).
        base = os.path.splitext(os.path.basename(args.excel))[0]
        new_base = replace_month_in_text(base, old_month, old_year, new_month, new_year)
        if args.outdir:
            out_dir = args.outdir
        else:
            basedir = os.path.dirname(args.excel) or "."
            out_dir = os.path.join(basedir, "output")
        os.makedirs(out_dir, exist_ok=True)
        output_path = os.path.join(out_dir, f"{new_base}.xlsx")

    print(f"\n💾 Saving: {output_path}")
    wb.save(output_path)
    print(f"✅ Done! Open '{output_path}' to verify.")


if __name__ == "__main__":
    main()