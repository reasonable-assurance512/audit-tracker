"""
Workbook orchestrator for the Audit Resource Tracker.

Coordinates the modular tab builders to produce a complete workbook
ready for download. Returns a BytesIO object suitable for
Streamlit's st.download_button.

Per Sprint 2 scope: inputs are accepted but not yet threaded through
to the tab builders. The output workbook uses reference dates that
match the v4 golden file. Input-driven generation is Sprint 3 work.
"""

from io import BytesIO

from openpyxl import Workbook

from .bbt_tab import build_bbt_tab
from .holidays_tab import build_holidays_tab
from .mbdd_tab import build_mbdd_tab
from .resource_tab import build_all_resource_tabs
from .setup_tab import build_setup_tab


def build_workbook(kickoff_date, planning_weeks, fieldwork_weeks, reporting_weeks):
    """
    Build the complete Audit Resource Tracker workbook.

    Args:
        kickoff_date: date object (currently unused; see module docstring)
        planning_weeks: int (currently unused)
        fieldwork_weeks: int (currently unused)
        reporting_weeks: int (currently unused)

    Returns:
        BytesIO containing the generated .xlsx file, positioned at start.
    """
    wb = Workbook()

    holidays_info = build_holidays_tab(wb)
    closed_range = holidays_info["closed_range"]
    skeleton_range = holidays_info["skeleton_range"]

    build_setup_tab(
        wb,
        closed_range=closed_range,
        skeleton_range=skeleton_range,
    )

    build_all_resource_tabs(
        wb,
        closed_range=closed_range,
        skeleton_range=skeleton_range,
    )

    build_mbdd_tab(
        wb,
        closed_range=closed_range,
        skeleton_range=skeleton_range,
    )

    build_bbt_tab(
        wb,
        closed_range=closed_range,
        skeleton_range=skeleton_range,
    )

    # Minimal metadata strip (full strip deferred to Sprint 5)
    wb.properties.creator = "Audit Resource Tracker"
    wb.properties.lastModifiedBy = ""
    wb.properties.title = ""
    wb.properties.subject = ""
    wb.properties.description = ""

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
