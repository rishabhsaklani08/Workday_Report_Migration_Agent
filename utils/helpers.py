"""
helpers.py — Shared utility functions.
"""

from openpyxl.styles import Border, Side
from config import COLORS


def get_border():
    """Return a standard thin cell border using the configured border color."""
    thin = Side(style="thin", color=COLORS["border_color"])
    return Border(left=thin, right=thin, top=thin, bottom=thin)
