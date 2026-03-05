"""DocPurpose element — document purpose block (1x1 table with text)."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class DocPurpose:
    """Document purpose block: 1-row 1-col table with centered text.

    Matches 도형.hwpx reference structure:
    - Table: borderFillIDRef for outer frame (SOLID 0.12mm)
    - Cell: borderFillIDRef for inner (DOUBLE_SLIM top/bottom 0.7mm, white fill)
    - Text centered vertically in the cell
    """

    text: str = ""
    font_name: str = "맑은 고딕"  # preset.font_table에서 설정됨
