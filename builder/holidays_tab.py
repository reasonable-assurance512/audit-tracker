"""
Builder for the Holidays & Skeleton tab.

Produces a reference tab listing state holidays (full closures) and
skeleton crew days across FY2026 and FY2027. Returns the row ranges
used by downstream tabs for COUNTIFS formulas.
"""

from .constants import (
    BROWN,
    DARK_BLUE,
    GRAY_LT,
    GRAY_MD,
    WHITE,
    make_align,
    make_border,
    make_fill,
    make_font,
)
from .holidays_data import CLOSURE_DAYS, SKELETON_DAYS


def _write_section(ws, start_row, title, header_bg, data):
    """Write one section (header + table + rows) to the tab."""
    ws.merge_cells(f"A{start_row}:D{start_row}")
    title_cell = ws[f"A{start_row}"]
    title_cell.value = title
    title_cell.font = make_font(bold=True, color=WHITE, size=10)
    title_cell.fill = make_fill(header_bg)
    title_cell.alignment = make_align()
    ws.row_dimensions[start_row].height = 18

    header_row = start_row + 1
    for col_idx, label in enumerate(["Holiday", "Date", "Day of Week", "FY"], 1):
        cell = ws.cell(row=header_row, column=col_idx, value=label)
        cell.font = make_font(bold=True, size=9)
        cell.fill = make_fill(GRAY_MD)
        cell.alignment = make_align("center")
        cell.border = make_border()
    ws.row_dimensions[header_row].height = 16

    data_start = header_row + 1
    current_row = data_start
    for name, holiday_date, fiscal_year in data:
        ws.cell(row=current_row, column=1, value=name).font = make_font(size=9)
        date_cell = ws.cell(row=current_row, column=2, value=holiday_date)
        date_cell.number_format = "MM/DD/YYYY"
        date_cell.font = make_font(size=9)
        ws.cell(row=current_row, column=3, value=holiday_date.strftime("%A")).font = make_font(size=9)
        ws.cell(row=current_row, column=4, value=fiscal_year).font = make_font(size=9)
        for col_idx in range(1, 5):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.border = make_border()
            cell.alignment = make_align("center" if col_idx > 1 else "left")
        ws.row_dimensions[current_row].height = 15
        current_row += 1

    return data_start, current_row - 1


def build_holidays_tab(wb):
    """
    Build the Holidays & Skeleton tab in the given workbook.

    Returns a dict with range references:
        - closed_range: cell range string for full closure dates
        - skeleton_range: cell range string for skeleton crew dates
    These are used by downstream tabs in COUNTIFS formulas.
    """
    ws = wb.active
    ws.title = "Holidays & Skeleton"
    ws.sheet_properties.tabColor = DARK_BLUE

    for col_letter, width in [("A", 30), ("B", 14), ("C", 14), ("D", 10)]:
        ws.column_dimensions[col_letter].width = width

    # Title
    ws.merge_cells("A1:D1")
    title = ws["A1"]
    title.value = "State Holiday Reference — FY2026 & FY2027 (Weekday Holidays Only)"
    title.font = make_font(bold=True, color=WHITE, size=12)
    title.fill = make_fill(DARK_BLUE)
    title.alignment = make_align("center")
    ws.row_dimensions[1].height = 22

    closed_start, closed_end = _write_section(
        ws, 3, "ALL AGENCIES CLOSED", DARK_BLUE, CLOSURE_DAYS
    )
    skeleton_start, skeleton_end = _write_section(
        ws, closed_end + 2, "SKELETON CREW DAYS", BROWN, SKELETON_DAYS
    )

    # Footer note
    note_row = skeleton_end + 2
    ws.merge_cells(f"A{note_row}:D{note_row}")
    note = ws[f"A{note_row}"]
    note.value = "Weekend holidays excluded from this table."
    note.font = make_font(italic=True, size=8, color="595959")
    note.fill = make_fill(GRAY_LT)
    note.alignment = make_align("left", wrap=True)
    ws.row_dimensions[note_row].height = 28

    closed_range = f"'Holidays & Skeleton'!$B${closed_start}:$B${closed_end}"
    skeleton_range = f"'Holidays & Skeleton'!$B${skeleton_start}:$B${skeleton_end}"

    return {
        "closed_range": closed_range,
        "skeleton_range": skeleton_range,
        "closed_start": closed_start,
        "closed_end": closed_end,
        "skeleton_start": skeleton_start,
        "skeleton_end": skeleton_end,
    }
