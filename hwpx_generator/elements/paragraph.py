"""Paragraph and TextRun elements."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class TextRun:
    """A run of text with uniform character properties."""

    text: str = ""
    font_name: Optional[str] = None   # override document default
    font_size_pt: Optional[float] = None
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: Optional[str] = None        # hex e.g. "000000"
    strike: bool = False


@dataclass
class Paragraph:
    """A paragraph containing one or more TextRuns."""

    runs: List[TextRun] = field(default_factory=list)

    # Paragraph properties
    align: str = "JUSTIFY"             # LEFT, CENTER, RIGHT, JUSTIFY
    line_spacing_type: str = "PERCENT" # PERCENT or FIXED
    line_spacing_value: float = 160.0  # % or HWP units
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    indent_left_mm: float = 0.0
    indent_right_mm: float = 0.0
    indent_first_mm: float = 0.0       # positive = indent, negative = hanging
    keep_with_next: bool = False
    widow_orphan: bool = False
    page_break_before: bool = False

    # Style reference (optional — overrides above if set)
    style_name: Optional[str] = None

    # Character defaults for this paragraph (used if TextRun doesn't override)
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: bool = False
    italic: bool = False
    color: Optional[str] = None

    def add_run(self, text: str, **kwargs) -> TextRun:
        """Add a text run to this paragraph."""
        run = TextRun(text=text, **kwargs)
        self.runs.append(run)
        return run

    @property
    def plain_text(self) -> str:
        return "".join(r.text for r in self.runs)
