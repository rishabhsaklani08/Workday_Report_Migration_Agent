"""
industry_router.py — Map Report_Tag values to industry names.

Supports direct mapping, partial matching, multi-tag delimiters,
and smart AI_ prefix extraction for tags not in the mapping table.

Since the parser already filters to AI_* tagged reports only,
detect_industry will never return UNMATCHED_TAB under normal
conditions.  The fallback is retained purely as a safety net.
"""

import logging
from config import INDUSTRY_CONFIG, UNMATCHED_TAB

logger = logging.getLogger(__name__)

# Flat tag → industry mapping from config
INDUSTRY_TAG_MAPPING: dict[str, str] = INDUSTRY_CONFIG["custom_mappings"]


def detect_industry(report_tag: str) -> str:
    """
    Detect the industry name from a single Report_Tag string.

    Resolution order:
        1. Exact match (case-insensitive)
        2. Partial / substring match
        3. Smart prefix extraction  (AI_Xyz → Xyz)
        4. Fallback → UNMATCHED_TAB  (safety net only)

    Args:
        report_tag: A single tag string (e.g. "AI_Healthcare").

    Returns:
        str: Resolved industry name.
    """
    if not report_tag or str(report_tag).strip() == "":
        return UNMATCHED_TAB

    tag = str(report_tag).strip()

    # 1. Exact match (case-insensitive)
    for mapped_tag, industry in INDUSTRY_TAG_MAPPING.items():
        if mapped_tag.lower() == tag.lower():
            return industry

    # 2. Partial / substring match
    for mapped_tag, industry in INDUSTRY_TAG_MAPPING.items():
        if mapped_tag.lower() in tag.lower():
            return industry

    # 3. Smart prefix extraction: AI_SomeName → Some Name
    if tag.upper().startswith("AI_"):
        industry_name = tag[3:]  # Strip "AI_"
        industry_name = industry_name.replace("_", " ").title()
        logger.info(
            f"🏷️  Auto-detected industry '{industry_name}' from tag '{tag}'"
        )
        return industry_name

    logger.warning(f"⚠️  Unmatched tag: '{tag}' → routed to '{UNMATCHED_TAB}'")
    return UNMATCHED_TAB


def handle_multiple_tags(report_tag: str) -> list[str]:
    """
    Handle reports whose Report_Tag contains multiple tags separated
    by common delimiters (; , | /).

    A report with multiple tags is duplicated into each matching
    industry tab.

    Args:
        report_tag: Raw tag value, possibly multi-valued.

    Returns:
        list[str]: De-duplicated list of industry names.
                   Empty list if no valid AI_ tags found.
    """
    if not report_tag or str(report_tag).strip() == "":
        return []

    report_tag = str(report_tag).strip()

    # Split by the first delimiter found
    delimiters = [";", ",", "|", "/"]
    tags = [report_tag]

    for delimiter in delimiters:
        if delimiter in report_tag:
            tags = [t.strip() for t in report_tag.split(delimiter) if t.strip()]
            break

    # Resolve each tag independently and de-duplicate
    industries: list[str] = []
    for tag in tags:
        industry = detect_industry(tag)
        if industry != UNMATCHED_TAB and industry not in industries:
            industries.append(industry)

    return industries
