"""
Workbook orchestrator for the Audit Resource Tracker.

Coordinates the modular tab builders to produce a complete workbook
ready for download. Returns a BytesIO object suitable for
Streamlit's st.download_button.

Per Sprint 3 Phase 3: AuditConfig values flow through to tab builders
that need them (setup_tab, resource_tab, mbdd_tab). The Holidays &
Skeleton and Budget by Task builders do not currently consume config
because their content is independent of audit-specific inputs.
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
    """
    wb = Workbook()

    holidays_info = build_holidays_tab(wb)
    closed_range = holidays_info["closed_range"]
    skeleton_range = holidays_info["skeleton_range"]

    build_setup_tab(
        wb,
        config=config,
        closed_range=closed_range,
        skeleton_range=skeleton_range,
    )

    build_all_resource_tabs(
        wb,
        config=config,
        closed_range=closed_range,
        skeleton_range=skeleton_range,
    )

    build_mbdd_tab(
        wb,
        config=config,
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
