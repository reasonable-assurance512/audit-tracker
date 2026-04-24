"""
Audit Resource Tracker — modular builder package.

Exports build_workbook for use by the Streamlit application.
Tab-specific builders (build_holidays_tab, build_setup_tab,
build_all_resource_tabs) remain available via direct submodule
imports for use by the parity test framework.
"""

from .workbook import build_workbook

__all__ = ["build_workbook"]
