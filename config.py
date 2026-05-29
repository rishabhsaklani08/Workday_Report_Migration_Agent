"""
config.py — Central configuration for Workday Report Catalog Agent.

All tunable settings for API connectivity, Excel output formatting,
industry tag mappings, and dependency detection live here.
"""

import os
from pathlib import Path

# Load .env file into os.environ (no external dependency needed)
_env_path = Path(__file__).resolve().parent / ".env"
if _env_path.exists():
    for _line in _env_path.read_text(encoding="utf-8").splitlines():
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _key, _, _val = _line.partition("=")
            os.environ.setdefault(_key.strip(), _val.strip())

# ─────────────────────────────────────────────────────────────
# Workday RaaS API Configuration
# ─────────────────────────────────────────────────────────────
WORKDAY_CONFIG = {
    # Direct RaaS report URL (without ?format=json — that's added automatically)
    "raas_url": os.getenv("WORKDAY_RAAS_URL", "https://[TENANT].workday.com/ccx/service/customreport2/[TENANT]/[USER]/[REPORT_NAME]"),
    "isu_username": os.getenv("WORKDAY_ISU_USERNAME", "[ISU_USERNAME]"),
    "isu_password": os.getenv("WORKDAY_ISU_PASSWORD", "[ISU_PASSWORD]"),

    # Timeout & retry settings
    "timeout_seconds": int(os.getenv("WORKDAY_TIMEOUT", "60")),
    "retry_attempts": int(os.getenv("WORKDAY_RETRIES", "3")),
    "retry_delay_seconds": int(os.getenv("WORKDAY_RETRY_DELAY", "5")),
}

# ─────────────────────────────────────────────────────────────
# Excel Output Configuration
# ─────────────────────────────────────────────────────────────
EXCEL_CONFIG = {
    # Output settings
    "output_directory": os.getenv("OUTPUT_DIR", "./output"),
    "filename_prefix": "Workday_Report_Catalog",
    "include_timestamp": True,

    # Sheet settings
    "freeze_panes": True,
    "auto_filter": True,
    "alternating_rows": True,
    "max_column_width": 50,
    "min_row_height": 20,
    "header_row_height": 35,
}

# ─────────────────────────────────────────────────────────────
# Color Scheme (hex without #)
# ─────────────────────────────────────────────────────────────
COLORS = {
    "header_bg": "1E3A5F",       # Dark navy blue
    "header_font": "FFFFFF",     # White
    "alt_row_1": "F0F4F8",       # Light blue-gray
    "alt_row_2": "FFFFFF",       # White
    "dependency_yes": "C6EFCE",  # Light green
    "dependency_no": "FFCCCC",   # Light red
    "summary_header": "2E4057",  # Darker navy
    "tab_accent": "4A90D9",      # Blue accent
    "title_bg": "0D47A1",        # Deep blue title
    "border_color": "B0BEC5",    # Gray border
}

# ─────────────────────────────────────────────────────────────
# Industry Tag Mapping Configuration
# ─────────────────────────────────────────────────────────────
INDUSTRY_CONFIG = {
    # Custom tag → industry name mappings
    "custom_mappings": {
        # Healthcare
        "AI_Healthcare": "Healthcare",
        "AI_HEALTHCARE": "Healthcare",
        "healthcare": "Healthcare",

        # Insurance
        "AI_Insurance": "Insurance",
        "AI_INSURANCE": "Insurance",
        "insurance": "Insurance",

        # Banking & Financial Services
        "AI_Banking": "Banking",
        "AI_BANKING": "Banking",
        "AI_FinancialServices": "Financial Services",
        "AI_Finance": "Financial Services",

        # Retail
        "AI_Retail": "Retail",
        "AI_RETAIL": "Retail",

        # Manufacturing
        "AI_Manufacturing": "Manufacturing",
        "AI_MANUFACTURING": "Manufacturing",

        # Technology
        "AI_Technology": "Technology",
        "AI_Tech": "Technology",

        # Government & Public Sector
        "AI_Government": "Government",
        "AI_PublicSector": "Government",

        # Education
        "AI_Education": "Education",
        "AI_EDUCATION": "Education",

        # Real Estate
        "AI_RealEstate": "Real Estate",

        # Add more mappings as needed
    },

    # Tab ordering (0 = first, higher = later)
    "tab_order": {
        "Summary": 0,
        "Healthcare": 1,
        "Insurance": 2,
        "Banking": 3,
        "Financial Services": 4,
        "Retail": 5,
        "Manufacturing": 6,
        "Technology": 7,
        "Government": 8,
        "Education": 9,
        "Real Estate": 10,
        "Other / Untagged": 999,
    },

    # Tab colors (hex without #)
    "tab_colors": {
        "Healthcare": "E91E63",
        "Insurance": "2196F3",
        "Banking": "4CAF50",
        "Financial Services": "FF9800",
        "Retail": "9C27B0",
        "Manufacturing": "795548",
        "Technology": "00BCD4",
        "Government": "607D8B",
        "Education": "FF5722",
        "Real Estate": "8BC34A",
        "Other / Untagged": "9E9E9E",
    },
}

# Fallback tab for unmatched tags
UNMATCHED_TAB = "Other / Untagged"
SUMMARY_TAB = "Summary"

# ─────────────────────────────────────────────────────────────
# Dependency Detection Configuration
# ─────────────────────────────────────────────────────────────
DEPENDENCY_CONFIG = {
    # Values that indicate NO dependencies
    "empty_values": ["", "null", "none", "n/a", "-", "no", "N/A"],

    # Display text
    "has_dependencies_text": "✅ Yes",
    "no_dependencies_text": "❌ No",

    # Row highlighting
    "highlight_dependency_rows": True,
    "dependency_row_color": "FFF9C4",  # Soft yellow
}
