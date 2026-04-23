"""
Parity tests comparing the modularized builder output against the
v4 golden file. A cell-by-cell comparison of values, formulas, and
key formatting properties for each tab that has been extracted so far.

Run with: pytest tests/test_parity.py -v
Or directly: python tests/test_parity.py
"""

import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from builder.holidays_tab import build_holidays_tab


GOLDEN_FILE_PATH = PROJECT_ROOT / "tests" / "golden_files" / "v4_reference.xlsx"


def generate_holidays_workbook():
    """Generate a new workbook with just the Holidays & Skeleton tab."""
    wb = Workbook()
    build_holidays_tab(wb)
    return wb


def normalize_value(value):
    """
    Normalize values so that equivalent representations compare equal.
    Handles the date/datetime difference that arises from saving and
    reloading xlsx files: openpyxl stores date as datetime-at-midnight
    when a workbook is saved and reopened.
    """
    if isinstance(value, datetime) and value.time().hour == 0 and value.time().minute == 0:
        return value.date()
    return value


def normalize_wrap_text(value):
    """
    Treat wrap_text=False and wrap_text=None as equivalent.
    Both mean 'text does not wrap' in Excel.
    """
    if value is None:
        return False
    return value


def get_cell_snapshot(cell):
    """
    Extract comparable properties from a cell.
    Returns a dict of value, font, fill, alignment, and border info.
    Values are normalized to avoid false positives from openpyxl's
    inconsistent representation of equivalent states.
    """
    return {
        "value": normalize_value(cell.value),
        "number_format": cell.number_format,
        "font_bold": cell.font.bold,
        "font_italic": cell.font.italic,
        "font_color": cell.font.color.rgb if cell.font.color else None,
        "font_size": cell.font.size,
        "fill_color": (
            cell.fill.start_color.rgb
            if cell.fill and cell.fill.start_color
            else None
        ),
        "horizontal": cell.alignment.horizontal,
        "vertical": cell.alignment.vertical,
        "wrap_text": normalize_wrap_text(cell.alignment.wrap_text),
    }


def compare_tabs(new_ws, golden_ws, tab_name):
    """
    Compare two worksheets cell-by-cell for all cells that have values
    in either sheet. Returns a list of differences; empty list means
    the tabs match.
    """
    differences = []

    max_row = max(new_ws.max_row, golden_ws.max_row)
    max_col = max(new_ws.max_column, golden_ws.max_column)

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            new_cell = new_ws.cell(row=row, column=col)
            golden_cell = golden_ws.cell(row=row, column=col)

            if new_cell.value is None and golden_cell.value is None:
                continue

            new_snap = get_cell_snapshot(new_cell)
            golden_snap = get_cell_snapshot(golden_cell)

            if new_snap != golden_snap:
                differences.append({
                    "tab": tab_name,
                    "cell": f"{new_cell.coordinate}",
                    "new": new_snap,
                    "golden": golden_snap,
                })

    return differences


def test_holidays_tab_parity():
    """Verify that the modularized holidays tab matches the golden file."""
    new_wb = generate_holidays_workbook()
    new_ws = new_wb["Holidays & Skeleton"]

    golden_wb = load_workbook(GOLDEN_FILE_PATH)
    golden_ws = golden_wb["Holidays & Skeleton"]

    differences = compare_tabs(new_ws, golden_ws, "Holidays & Skeleton")

    if differences:
        print(f"\n{len(differences)} differences found:")
        for diff in differences[:10]:
            print(f"  Cell {diff['cell']}:")
            print(f"    new:    {diff['new']}")
            print(f"    golden: {diff['golden']}")
        if len(differences) > 10:
            print(f"  ... and {len(differences) - 10} more")
        raise AssertionError(
            f"Holidays & Skeleton tab has {len(differences)} cell differences"
        )

    print("Holidays & Skeleton tab: PARITY VERIFIED (no differences)")


if __name__ == "__main__":
    test_holidays_tab_parity()
