"""
Holiday and skeleton-crew day data for FY2026 and FY2027.

This module is the single source of truth for holiday dates used by both
the Holidays & Skeleton tab builder and the date-math module that
powers the live preview in the Streamlit app.

Data shape: list of (name, date, fiscal_year_label) tuples. Weekend
holidays are intentionally excluded — only weekday holidays are tracked,
matching the behavior of the v4 Excel workbook.

Source: Texas state holiday calendar, FY2026 and FY2027.

Future replacement: backlog item F-03 will allow user-uploaded holiday
calendars. When implemented, this module's contents will be replaced
by data loaded from a user-provided source. The module's public
interface (CLOSURE_DAYS, SKELETON_DAYS as lists of (name, date, fy)
tuples) will be preserved.
"""

from datetime import date


CLOSURE_DAYS = [
    ("Labor Day", date(2025, 9, 1), "FY2026"),
    ("Veterans Day", date(2025, 11, 11), "FY2026"),
    ("Thanksgiving Day", date(2025, 11, 27), "FY2026"),
    ("Day after Thanksgiving", date(2025, 11, 28), "FY2026"),
    ("Christmas Eve Day", date(2025, 12, 24), "FY2026"),
    ("Christmas Day", date(2025, 12, 25), "FY2026"),
    ("Day after Christmas", date(2025, 12, 26), "FY2026"),
    ("New Year's Day", date(2026, 1, 1), "FY2026"),
    ("Martin Luther King, Jr. Day", date(2026, 1, 19), "FY2026"),
    ("Presidents' Day", date(2026, 2, 16), "FY2026"),
    ("Memorial Day", date(2026, 5, 25), "FY2026"),
    ("Labor Day", date(2026, 9, 7), "FY2027"),
    ("Veterans Day", date(2026, 11, 11), "FY2027"),
    ("Thanksgiving Day", date(2026, 11, 26), "FY2027"),
    ("Day after Thanksgiving", date(2026, 11, 27), "FY2027"),
    ("Christmas Eve Day", date(2026, 12, 24), "FY2027"),
    ("Christmas Day", date(2026, 12, 25), "FY2027"),
    ("New Year's Day", date(2027, 1, 1), "FY2027"),
    ("Martin Luther King, Jr. Day", date(2027, 1, 18), "FY2027"),
    ("Presidents' Day", date(2027, 2, 15), "FY2027"),
    ("Memorial Day", date(2027, 5, 31), "FY2027"),
]

SKELETON_DAYS = [
    ("Texas Independence Day", date(2026, 3, 2), "FY2026"),
    ("San Jacinto Day", date(2026, 4, 21), "FY2026"),
    ("Emancipation Day", date(2026, 6, 19), "FY2026"),
    ("LBJ Day", date(2026, 8, 27), "FY2026"),
    ("Confederate Heroes Day", date(2027, 1, 19), "FY2027"),
    ("Texas Independence Day", date(2027, 3, 2), "FY2027"),
    ("San Jacinto Day", date(2027, 4, 21), "FY2027"),
    ("LBJ Day", date(2027, 8, 27), "FY2027"),
]
