"""
dependency_checker.py — Detect calculated-field dependencies.

Inspects the Calculated_Fields_for_Report value and returns
a boolean flag indicating whether the report has dependencies.
"""

import logging
from config import DEPENDENCY_CONFIG

logger = logging.getLogger(__name__)

# Values treated as "no dependencies"
_EMPTY_VALUES: set[str] = {
    v.strip().lower() for v in DEPENDENCY_CONFIG["empty_values"]
}


def has_dependencies(calculated_fields) -> bool:
    """
    Determine whether a report has calculated-field dependencies.

    Args:
        calculated_fields: Raw value from the Calculated_Fields_for_Report
                           JSON field (may be str, None, or other).

    Returns:
        bool: True if the report has meaningful calculated fields.
    """
    if calculated_fields is None:
        return False

    value = str(calculated_fields).strip()

    if value == "":
        return False

    if value.lower() in _EMPTY_VALUES:
        return False

    return True


def get_dependency_label(calculated_fields) -> str:
    """Return the user-facing dependency flag text."""
    if has_dependencies(calculated_fields):
        return DEPENDENCY_CONFIG["has_dependencies_text"]
    return DEPENDENCY_CONFIG["no_dependencies_text"]


def get_dependency_details(calculated_fields) -> str:
    """Return the detail string (the calc expression, or 'None')."""
    if has_dependencies(calculated_fields):
        return str(calculated_fields).strip()
    return "None"
