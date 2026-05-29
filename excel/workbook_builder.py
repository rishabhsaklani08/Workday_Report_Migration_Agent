"""
workbook_builder.py — Orchestrate multi-tab Excel workbook creation.
"""

import logging
import openpyxl

from config import UNMATCHED_TAB, INDUSTRY_CONFIG
from excel.summary_tab import create_summary_tab
from excel.industry_tab import create_industry_tab

logger = logging.getLogger(__name__)

TAB_ORDER = INDUSTRY_CONFIG["tab_order"]


def create_excel_workbook(processed_data: dict, output_path: str) -> str:
    """
    Create the full multi-tab Excel workbook from processed data.

    Args:
        processed_data: {industry_name: [report_dicts, ...]}.
        output_path: File path for the .xlsx output.

    Returns:
        The output_path on success.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    # Sort industries: Summary first, Other last, rest alphabetical
    industries = sorted(
        processed_data.keys(),
        key=lambda x: TAB_ORDER.get(x, 500),
    )

    # Summary tab (always first)
    create_summary_tab(wb, processed_data)

    # Industry tabs
    for industry in industries:
        reports = processed_data[industry]
        create_industry_tab(wb, industry, reports)

    wb.save(output_path)
    logger.info(f"✅ Excel file saved: {output_path}")
    return output_path
