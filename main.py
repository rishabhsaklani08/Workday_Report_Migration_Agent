"""
main.py — Workday Report Catalog Agent entry point.

Orchestrates:
    1. Fetch from Workday RaaS API (or use sample data)
    2. Parse JSON response (filters to AI_* tagged reports only)
    3. Categorize reports by industry via Report_Tag
    4. Flag calculated-field dependencies
    5. Generate multi-tab Excel workbook
"""

import os
import sys
import logging
from datetime import datetime
from pathlib import Path

# Force UTF-8 output on Windows consoles (avoids cp1252 emoji crashes)
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from api.raas_client import fetch_workday_report
from processors.json_parser import parse_json_response
from processors.industry_router import handle_multiple_tags
from processors.dependency_checker import (
    get_dependency_label,
    get_dependency_details,
    has_dependencies,
)
from excel.workbook_builder import create_excel_workbook
from config import EXCEL_CONFIG

# ─────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("agent.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────
# Core Processing
# ─────────────────────────────────────────────────────────────
def process_reports(raw_reports: list[dict]) -> dict:
    """
    Process all AI-tagged reports:
        1. Extract all real Workday fields (with graceful defaults)
        2. Detect industry from Report_Tag
        3. Flag dependencies from Calculated_Fields_for_Report
        4. Structure into {industry: [report_rows]} for Excel output

    Only reports that resolve to at least one industry tab are included.
    """
    processed: dict[str, list[dict]] = {}

    for report in raw_reports:
        # ── Extract all real Workday JSON fields ──────────────
        report_data = {
            "Report_Name": report.get("Report_Name", "N/A"),
            "Report_Type": report.get("Report_Type", "N/A"),
            "Output_Type": report.get("Output_Type", "N/A"),
            "Category": report.get("Category", "N/A"),
            "Data_Source": report.get("Data_Source", "N/A"),
            "Report_Owner": report.get("Report_Owner", "N/A"),
            "Report_Tag": report.get("Report_Tag", ""),
            "Fields_Referenced_in_Report": report.get(
                "Fields_Referenced_in_Report", ""
            ),
            "Calculated_Fields_for_Report": report.get(
                "Calculated_Fields_for_Report", ""
            ),
            "Created_On": report.get("Created_On", "N/A"),
            "Last_Updated": report.get("Last_Updated", "N/A"),
            "Shared": report.get("Shared", "N/A"),
            "Worklet": report.get("Worklet", "N/A"),
        }

        # ── Dependency flagging ───────────────────────────────
        calc_fields = report_data["Calculated_Fields_for_Report"]
        report_data["Has_Dependencies"] = get_dependency_label(calc_fields)
        report_data["Dependency_Details"] = get_dependency_details(calc_fields)

        # ── Industry detection & routing ──────────────────────
        industries = handle_multiple_tags(report_data["Report_Tag"])

        # Skip if no valid AI_ industry resolved
        if not industries:
            logger.debug(
                f"Skipping report '{report_data['Report_Name']}' — "
                f"no AI_ industry resolved from tag '{report_data['Report_Tag']}'"
            )
            continue

        for industry in industries:
            processed.setdefault(industry, []).append(report_data)

    return processed


# ─────────────────────────────────────────────────────────────
# Sample Data (mirrors actual Workday RaaS JSON structure)
# ─────────────────────────────────────────────────────────────
def get_sample_data() -> dict:
    """
    Realistic sample dataset matching the actual Workday RaaS JSON
    structure — includes tagged, untagged, and multi-tagged reports.
    """
    return {
        "Report_Entry": [
            # ── Reports WITH AI_* tags (these should be included) ──
            {
                "Report_Name": "<tenant> Patient Demographics Report",
                "Report_Type": "Advanced",
                "Output_Type": "Table",
                "Category": "Patient Data",
                "Data_Source": "Workers for HCM Reporting",
                "Report_Owner": "admin-user / Admin User",
                "Report_Tag": "AI_Healthcare",
                "Fields_Referenced_in_Report": "Worker; Age; Worker Types; Worker Terminated (Indexed Workers Filter)",
                "Calculated_Fields_for_Report": "PTN_<tenant>_AGE_CALC; PTN_<tenant>_LOS_CALC",
                "Created_On": "2026-05-01T10:30:00.000-07:00",
                "Last_Updated": "2026-05-10T14:22:00.000-07:00",
                "Shared": "1",
                "Worklet": "0",
            },
            {
                "Report_Name": "<tenant> Healthcare Billing Analysis",
                "Report_Type": "Advanced",
                "Output_Type": "Table",
                "Category": "Financial Data",
                "Data_Source": "All Active Workers",
                "Report_Owner": "finance-user / Finance Admin",
                "Report_Tag": "AI_Healthcare",
                "Fields_Referenced_in_Report": "Invoice_ID; Patient_ID; Amount; Insurance_Code",
                "Calculated_Fields_for_Report": "PTN_<tenant>_NET_REVENUE; PTN_<tenant>_MARGIN",
                "Created_On": "2026-04-15T08:00:00.000-07:00",
                "Last_Updated": "2026-05-12T09:45:00.000-07:00",
                "Shared": "0",
                "Worklet": "0",
            },
            {
                "Report_Name": "<tenant> Claims Summary Q1",
                "Report_Type": "Advanced",
                "Output_Type": "Table",
                "Category": "Claims Data",
                "Data_Source": "All Active and Terminated Workers",
                "Report_Owner": "claims-user / Claims Analyst",
                "Report_Tag": "AI_Insurance",
                "Fields_Referenced_in_Report": "Claim_ID; Policy_Number; Amount; Status",
                "Calculated_Fields_for_Report": "",
                "Created_On": "2026-03-01T12:00:00.000-07:00",
                "Last_Updated": "2026-05-08T16:30:00.000-07:00",
                "Shared": "1",
                "Worklet": "0",
            },
            {
                "Report_Name": "<tenant> Premium Collection Report",
                "Report_Type": "Simple",
                "Output_Type": "Table",
                "Category": "Premium Data",
                "Data_Source": "Policy Admin System",
                "Report_Owner": "admin-user / Admin User",
                "Report_Tag": "AI_Insurance",
                "Fields_Referenced_in_Report": "Policy_ID; Premium; Region; Start_Date",
                "Calculated_Fields_for_Report": "PTN_<tenant>_ANNUAL_TOTAL",
                "Created_On": "2026-02-20T09:15:00.000-07:00",
                "Last_Updated": "2026-05-11T11:00:00.000-07:00",
                "Shared": "0",
                "Worklet": "0",
            },
            {
                "Report_Name": "<tenant> Daily Transactions",
                "Report_Type": "Composite",
                "Output_Type": "Table",
                "Category": "Transaction Data",
                "Data_Source": "Transaction Ledger",
                "Report_Owner": "banking-user / Banking Ops",
                "Report_Tag": "AI_Banking",
                "Fields_Referenced_in_Report": "TXN_ID; Amount; Date; Account",
                "Calculated_Fields_for_Report": "PTN_<tenant>_NET_BALANCE",
                "Created_On": "2026-01-10T07:00:00.000-07:00",
                "Last_Updated": "2026-05-13T08:20:00.000-07:00",
                "Shared": "1",
                "Worklet": "0",
            },
            {
                "Report_Name": "<tenant> Risk Assessment Dashboard",
                "Report_Type": "Advanced",
                "Output_Type": "Table",
                "Category": "Risk Data",
                "Data_Source": "Risk Engine",
                "Report_Owner": "risk-user / Risk Manager",
                "Report_Tag": "AI_Banking",
                "Fields_Referenced_in_Report": "Account; Risk_Score; Exposure; Collateral",
                "Calculated_Fields_for_Report": "PTN_<tenant>_LTV_RATIO",
                "Created_On": "2026-02-05T14:30:00.000-07:00",
                "Last_Updated": "2026-05-09T10:15:00.000-07:00",
                "Shared": "0",
                "Worklet": "0",
            },
            {
                "Report_Name": "<tenant> Monthly Sales",
                "Report_Type": "Simple",
                "Output_Type": "Table",
                "Category": "Sales Data",
                "Data_Source": "POS System",
                "Report_Owner": "retail-user / Retail Analyst",
                "Report_Tag": "AI_Retail",
                "Fields_Referenced_in_Report": "Store_ID; Category; Units_Sold; Revenue",
                "Calculated_Fields_for_Report": "PTN_<tenant>_AVG_PRICE",
                "Created_On": "2026-03-15T11:00:00.000-07:00",
                "Last_Updated": "2026-05-12T13:00:00.000-07:00",
                "Shared": "1",
                "Worklet": "0",
            },
            {
                "Report_Name": "<tenant> Production Output Weekly",
                "Report_Type": "Advanced",
                "Output_Type": "Table",
                "Category": "Production Data",
                "Data_Source": "Production DB",
                "Report_Owner": "mfg-user / Manufacturing Lead",
                "Report_Tag": "AI_Manufacturing",
                "Fields_Referenced_in_Report": "Line_ID; Units_Produced; Defects; Shift",
                "Calculated_Fields_for_Report": "PTN_<tenant>_DEFECT_RATE",
                "Created_On": "2026-04-01T06:00:00.000-07:00",
                "Last_Updated": "2026-05-13T07:00:00.000-07:00",
                "Shared": "0",
                "Worklet": "0",
            },
            {
                "Report_Name": "<tenant> Sprint Velocity",
                "Report_Type": "Simple",
                "Output_Type": "Table",
                "Category": "Project Data",
                "Data_Source": "Jira Export",
                "Report_Owner": "tech-user / Tech Lead",
                "Report_Tag": "AI_Technology",
                "Fields_Referenced_in_Report": "Sprint_ID; Story_Points; Completed; Remaining",
                "Calculated_Fields_for_Report": "",
                "Created_On": "2026-04-20T09:00:00.000-07:00",
                "Last_Updated": "2026-05-11T17:00:00.000-07:00",
                "Shared": "1",
                "Worklet": "0",
            },
            {
                "Report_Name": "<tenant> Cross-Industry Health-Insurance",
                "Report_Type": "Composite",
                "Output_Type": "Table",
                "Category": "Cross-Industry",
                "Data_Source": "Data Warehouse",
                "Report_Owner": "admin-user / Admin User",
                "Report_Tag": "AI_Healthcare; AI_Insurance",
                "Fields_Referenced_in_Report": "Patient_ID; Claim_ID; Amount",
                "Calculated_Fields_for_Report": "PTN_<tenant>_COMBINED_COST",
                "Created_On": "2026-05-05T10:00:00.000-07:00",
                "Last_Updated": "2026-05-13T12:00:00.000-07:00",
                "Shared": "0",
                "Worklet": "0",
            },
            # ── Reports WITHOUT AI_* tags (should be EXCLUDED) ─────
            {
                "Report_Name": "<snguvvala-trn> Worker Details and Dependents",
                "Report_Type": "Advanced",
                "Output_Type": "Table",
                "Category": "Worker Data",
                "Data_Source": "Workers for HCM Reporting",
                "Report_Owner": "snguvvala-trn / siva nagi guvvala",
                "Fields_Referenced_in_Report": "Worker; Age; Worker Types; Worker Terminated (Indexed Workers Filter)",
                "Calculated_Fields_for_Report": "PTN_<snguvvala-trn>_LRV_EXTRACT_OLDEST; PTN_<snguvvala-trn>_DD_Oldest",
                "Created_On": "2026-05-11T04:43:40.048-07:00",
                "Last_Updated": "2026-05-13T02:05:39.208-07:00",
                "Shared": "0",
                "Worklet": "0",
            },
            {
                "Report_Name": "<snguvvala-trn>_Employee Comp Levels",
                "Report_Type": "Advanced",
                "Output_Type": "Table",
                "Category": "Worker Data",
                "Data_Source": "All Active and Terminated Workers",
                "Report_Owner": "snguvvala-trn / siva nagi guvvala",
                "Fields_Referenced_in_Report": "-; Total Base Pay Annualized in Reporting Currency",
                "Calculated_Fields_for_Report": "PTN_<snguvvala-trn>_CF_EEB_COMP",
                "Created_On": "2026-05-12T02:18:52.002-07:00",
                "Last_Updated": "2026-05-12T02:51:48.893-07:00",
                "Shared": "0",
                "Worklet": "0",
            },
            {
                "Report_Name": "<tenant> General Worker List",
                "Report_Type": "Simple",
                "Output_Type": "Table",
                "Category": "Worker Data",
                "Data_Source": "All Workers",
                "Report_Owner": "admin / system admin",
                "Fields_Referenced_in_Report": "Worker; Name; Employee_ID",
                "Calculated_Fields_for_Report": "",
                "Created_On": "2026-01-05T08:00:00.000-07:00",
                "Last_Updated": "2026-05-10T10:00:00.000-07:00",
                "Shared": "1",
                "Worklet": "0",
            },
        ]
    }


# ─────────────────────────────────────────────────────────────
# Agent Orchestration
# ─────────────────────────────────────────────────────────────
def run_agent(output_path=None, use_sample_data=False):
    """
    Main agent pipeline:
        1. Fetch data from Workday RaaS (or sample)
        2. Parse JSON (filter to AI_* tagged reports)
        3. Categorize reports by industry
        4. Generate multi-tab Excel workbook
    """
    print("🤖 Workday Report Catalog Agent Starting...")
    print("=" * 55)

    # Step 1: Fetch
    print("\n📡 Step 1: Fetching data from Workday RaaS API...")
    if use_sample_data:
        logger.info("Using sample data (no live API call)")
        raw_data = get_sample_data()
    else:
        raw_data = fetch_workday_report()

    # Step 2: Parse & filter to AI_* tagged reports
    print("\n🔍 Step 2: Parsing JSON & filtering AI_* tagged reports...")
    reports = parse_json_response(raw_data)
    if not reports:
        logger.warning(
            "⚠️ No AI-tagged reports found in response. "
            "Ensure reports have a Report_Tag field with AI_* values."
        )
        return None

    # Step 3: Categorize
    print("\n🏷️  Step 3: Categorizing reports by industry...")
    processed = process_reports(reports)

    if not processed:
        logger.warning("⚠️ No reports matched any industry after processing.")
        return None

    print("\n📊 Categorization Summary:")
    print("-" * 50)
    for industry, rpts in sorted(processed.items()):
        deps = sum(1 for r in rpts if r["Has_Dependencies"] == "✅ Yes")
        print(
            f"  {industry:<30} {len(rpts):>3} reports "
            f"({deps} with dependencies)"
        )
    print("-" * 50)
    print(
        f"  {'TOTAL':<30} "
        f"{sum(len(r) for r in processed.values()):>3} reports"
    )

    # Step 4: Generate Excel
    print("\n📊 Step 4: Generating Excel workbook...")
    output_dir = EXCEL_CONFIG["output_directory"]
    os.makedirs(output_dir, exist_ok=True)

    if not output_path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        prefix = EXCEL_CONFIG["filename_prefix"]
        output_path = os.path.join(output_dir, f"{prefix}_{timestamp}.xlsx")

    create_excel_workbook(processed, output_path)

    # Summary
    print("\n" + "=" * 55)
    print("✅ Agent completed successfully!")
    print(f"📁 Output file: {os.path.abspath(output_path)}")
    print(f"📋 Total AI-tagged reports processed: {len(reports)}")
    print(f"🗂️  Industry tabs created: {len(processed)}")
    print("=" * 55)

    return output_path


# ─────────────────────────────────────────────────────────────
# Entry Point
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # CLI flag: pass --sample to use sample data without live API
    use_sample = "--sample" in sys.argv
    run_agent(use_sample_data=use_sample)
