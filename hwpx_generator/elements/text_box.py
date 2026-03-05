"""TextBox element -- highlighted text box using hp:rect with embedded text."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass
class TextBox:
    """A highlighted text box (hp:rect with embedded text).

    Attributes:
        text: The text content displayed inside the box.
        border_color: Border color as hex string (e.g. "C00000").
        bg_color: Background fill color as hex string (e.g. "FFF2CC").
        font_name: Font name; None uses document default.
        font_size_pt: Font size in points; None uses document default.
        font_bold: Whether text is bold.
        font_italic: Whether text is italic.
        font_color: Text color as hex string; None uses "000000".
        padding_mm: Internal padding in millimeters.
        width_mm: Box width in mm; None = auto (full content width).
        height_mm: Box height in mm; None = auto-calculated.
        align: Horizontal alignment -- "left", "center", or "right".
    """

    text: str = ""
    border_color: str = "000000"
    bg_color: Optional[str] = None
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    font_bold: bool = False
    font_italic: bool = False
    font_color: Optional[str] = None
    padding_mm: float = 4.0
    width_mm: Optional[float] = None
    height_mm: Optional[float] = None
    align: str = "left"
