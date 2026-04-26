"""
Audit Resource Tracker — Streamlit entry point.

Accepts kickoff date and phase weeks, generates a complete 13-tab
Excel workbook via the modular builder package, and offers it as
a browser download.

Per Sprint 3 conservative scope: user inputs are now wrapped in an
AuditConfig dataclass and passed to the orchestrator. Tab builders
do not yet consume config values; that is Phase 3 of Sprint 3 work.
"""

from datetime import date

import streamlit as st

from builder import build_workbook
from builder.config import AuditConfig


st.set_page_config(
    page_title="Audit Resource Tracker",
    page_icon="📋",
    layout="centered",
)

st.title("Audit Resource Tracker")
st.caption("Sprint 3 in progress — config dataclass wired, tab builders updating")

st.markdown(
    """
    > **Notice.** This tool is a personal project. It contains no client,
    > employer, or agency data, and does not connect to any workplace system.
    > It does not use any AI or LLM service.
    """
)

st.markdown(
    """
    > **Scope note.** Output uses reference dates regardless of inputs.
    > Input-driven generation is the next phase of Sprint 3.
    """
)

st.divider()

st.subheader("Audit parameters")

col1, col2 = st.columns(2)
with col1:
    kickoff_date = st.date_input(
        "Kickoff date",
        value=date(2026, 5, 4),
        help="The audit's kickoff / project launch date",
    )
with col2:
    st.write("")

planning_weeks = st.number_input(
    "Planning weeks",
    min_value=1,
    max_value=50,
    value=4,
    step=1,
    help="Number of weeks for the Planning phase (minimum 1)",
)

fieldwork_weeks = st.number_input(
    "Fieldwork weeks",
    min_value=1,
    max_value=50,
    value=16,
    step=1,
    help="Number of weeks for the Fieldwork phase (minimum 1)",
)

reporting_weeks = st.number_input(
    "Reporting weeks",
    min_value=1,
    max_value=50,
    value=4,
    step=1,
    help="Number of weeks for the Reporting phase (minimum 1)",
)

total_weeks = planning_weeks + fieldwork_weeks + reporting_weeks

if total_weeks > 52:
    st.warning(
        f"Total audit duration is {total_weeks} weeks. "
        "The v1 tool enforces a 52-week maximum (see living document Section 13.7). "
        "Reduce one or more phases before generating."
    )
    can_generate = False
else:
    st.info(f"Total audit duration: {total_weeks} weeks")
    can_generate = True

st.divider()

if st.button("Generate workbook", type="primary", disabled=not can_generate):
    config = AuditConfig(
        kickoff_date=kickoff_date,
        planning_weeks=int(planning_weeks),
        fieldwork_weeks=int(fieldwork_weeks),
        reporting_weeks=int(reporting_weeks),
    )

    errors = config.validate()
    if errors:
        for error in errors:
            st.error(error)
    else:
        output = build_workbook(config)
        filename = f"Audit_Tracker_{kickoff_date.strftime('%Y-%m-%d')}.xlsx"
        st.success("Workbook generated. Click below to download.")
        st.download_button(
            label=f"Download {filename}",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
