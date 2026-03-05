"""Unit conversion utilities for HWPX generation.

HWPX internal unit: 1/7200 inch (≈ 0.00353mm), referred to as "HWP unit".
"""

import math


def mm_to_hwp(mm: float) -> int:
    """Convert millimeters to HWP units (1/7200 inch)."""
    return round(mm * 7200 / 25.4)


def hwp_to_mm(hwp: int) -> float:
    """Convert HWP units to millimeters."""
    return hwp * 25.4 / 7200


def pt_to_hwp(pt: float) -> int:
    """Convert points to HWP font size units (1/100 pt in HWPX)."""
    return round(pt * 100)


def color_to_hex(color: str) -> str:
    """Normalize color string to 6-digit hex (no #)."""
    color = color.strip().lstrip("#")
    if len(color) == 3:
        color = "".join(c * 2 for c in color)
    return color.upper()


def color_to_rgb_int(color: str) -> int:
    """Convert hex color string to integer (for HWPX attributes)."""
    h = color_to_hex(color)
    return int(h, 16)
