"""
json_parser.py — Parse Workday RaaS JSON responses.

Handles the various JSON envelope formats Workday may return,
normalizes them into a flat list of report dictionaries, and
filters to only reports carrying an AI_* Report_Tag.
"""

import re
import logging

logger = logging.getLogger(__name__)

# Regex: tag must start with "AI_" followed by at least one word character
_AI_TAG_PATTERN = re.compile(r"AI_\w+", re.IGNORECASE)


def parse_json_response(json_data: dict) -> list[dict]:
    """
    Parse a Workday RaaS JSON response and return report entries
    that have a Report_Tag matching the AI_(IndustryName) pattern.

    Reports without a Report_Tag field, or whose tag does not match
    AI_*, are silently excluded.

    Args:
        json_data: Raw JSON dict from the RaaS API.

    Returns:
        list[dict]: Filtered list of report-entry dicts.
    """
    all_reports: list[dict] = []

    # ── Detect envelope key ──────────────────────────────────
    if "Report_Entry" in json_data:
        all_reports = json_data["Report_Entry"]

    elif "report" in json_data:
        all_reports = json_data["report"]

    elif isinstance(json_data, list):
        all_reports = json_data

    else:
        # Heuristic: first list value in the top-level dict
        for key, value in json_data.items():
            if isinstance(value, list) and len(value) > 0:
                logger.info(
                    f"🔍 Auto-detected report array under key: '{key}'"
                )
                all_reports = value
                break

    # Guard: keep only dict entries
    all_reports = [r for r in all_reports if isinstance(r, dict)]

    logger.info(f"📊 Total reports in JSON response: {len(all_reports)}")

    # ── Filter: keep only reports with an AI_* Report_Tag ────
    tagged_reports: list[dict] = []
    skipped = 0

    for report in all_reports:
        tag = report.get("Report_Tag", "")
        if tag and _AI_TAG_PATTERN.search(str(tag)):
            tagged_reports.append(report)
        else:
            skipped += 1

    logger.info(
        f"🏷️  Reports with AI_* tag: {len(tagged_reports)}  |  "
        f"Skipped (no AI tag): {skipped}"
    )

    return tagged_reports
