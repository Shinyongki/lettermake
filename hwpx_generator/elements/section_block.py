"""SectionBlock element — section divider with auto-incrementing number."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass
class SectionBlock:
    """Section divider block: 2-row 3-col table with number + divider + title.

    Structure (from 도형.hwpx):
    - Col 0: Number cell (2-row merged, blue bg, white bold text)
    - Col 1: Thin divider cell (2-row merged, gray side borders)
    - Col 2 Row 0: Title text cell (gradation fill)
    - Col 2 Row 1: Thin bottom bar (gradation fill, height=100)

    Number auto-increments across all SectionBlock instances in the document.
    """

    text: str = ""
    num: Optional[int] = None  # Manual override; None = auto-increment
