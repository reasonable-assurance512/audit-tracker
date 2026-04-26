"""
Workbook orchestrator for the Audit Resource Tracker.

Coordinates the modular tab builders to produce a complete workbook
ready for download. Returns a BytesIO object suitable for
Streamlit's st.download_button.

Per Sprint 3 (conservative scope): inputs are accepted as an AuditConfig
dataclass and forwarded to tab builders. The tab builders themselves
do not yet consume config values; that is Phase 3 of Sprint 3 work.
For now, the workbook output is identical regardless of config because
the tab builders use their existing hardcoded defaults.
"""

from io import BytesIO

from openpyxl import Workbook

from .bbt_tab import build_bbt_tab
from .config import AuditConfig
from .holidays_tab import build_holidays_tab
from .mbdd_tab import build_mbdd_tab
from .resource_tab import build_all_resource_tabs
from .setup_tab import build_setup_tab


def build_workbook(config: AuditConfig) -> BytesIO:
    """
    Build the complete Audit Resource Tracker workbook.

    Args:
        config: AuditConfig dataclass with kickoff date, phase weeks,
            on-target buffer, and hours-per-holiday values.

    Returns:
        BytesIO containing the generated .xlsx file, positioned at start.

    Note:
        Per Sprint 3 conservative scope, config is currently accepted but
        not yet threaded through to tab builders. The output workbook uses
        hardcoded defaults that match the v4 reference. Sprint 3 Phase 3
        will make config values flow into tab content.
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
