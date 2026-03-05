"""Table, TableRow, TableCell elements."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional, Union

from .paragraph import Paragraph, TextRun


@dataclass
class TableCell:
    """A single table cell."""

    paragraphs: List[Paragraph] = field(default_factory=list)
    colspan: int = 1
    rowspan: int = 1
    bg_color: Optional[str] = None     # hex, None = inherit from row/table
    valign: str = "CENTER"             # TOP, CENTER, BOTTOM
    align: Optional[str] = None        # override paragraph alignment

    def set_text(self, text: str, **char_kwargs) -> None:
        """Set cell content as a single paragraph with one run."""
        p = Paragraph(align=self.align or "LEFT")
        p.add_run(text, **char_kwargs)
        self.paragraphs = [p]


@dataclass
class TableRow:
    """A table row."""

    cells: List[TableCell] = field(default_factory=list)
    is_header: bool = False
    height_mm: Optional[float] = None  # None = auto


@dataclass
class Table:
    """A table element.

    Usage::

        table = Table(col_widths_mm=[25, 145])
        table.add_header_row(["항목", "내용"])
        table.add_row(["기간", "2026년 3월"])
    """

    col_widths_mm: List[float] = field(default_factory=list)
    rows: List[TableRow] = field(default_factory=list)

    # Table-level settings
    keep_together: bool = True
    border_color: str = "000000"
    border_width: str = "0.12mm"

    # Cell margins (mm)
    cell_margin_lr_mm: float = 3.0
    cell_margin_tb_mm: float = 2.0

    # Header row styling
    header_bg_color: str = "D9D9D9"
    header_font_color: str = "000000"
    header_font_bold: bool = True
    header_font_size_pt: Optional[float] = None  # None = use table font
    header_align: str = "CENTER"

    # Data row styling
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None

    def add_header_row(self, texts: List[str]) -> TableRow:
        """Add a header row with text values."""
        row = TableRow(is_header=True)
        for text in texts:
            cell = TableCell(
                bg_color=self.header_bg_color,
                align=self.header_align,
            )
            p = Paragraph(align=self.header_align)
            p.add_run(
                text,
                font_name=self.font_name,
                font_size_pt=self.header_font_size_pt or self.font_size_pt,
                bold=self.header_font_bold,
                color=self.header_font_color,
            )
            cell.paragraphs.append(p)
            row.cells.append(cell)
        self.rows.append(row)
        return row

    def add_row(
        self,
        texts: List[str],
        bg_color: Optional[str] = None,
    ) -> TableRow:
        """Add a data row with text values."""
        row = TableRow()
        for text in texts:
            cell = TableCell(bg_color=bg_color)
            p = Paragraph(align="LEFT")
            p.add_run(
                text,
                font_name=self.font_name,
                font_size_pt=self.font_size_pt,
            )
            cell.paragraphs.append(p)
            row.cells.append(cell)
        self.rows.append(row)
        return row

    def add_merged_row(
        self,
        cell_defs: List[dict],
    ) -> TableRow:
        """Add a row with merge support.

        Args:
            cell_defs: List of dicts with keys:
                - text (str): Cell text
                - colspan (int): Column span (default 1)
                - rowspan (int): Row span (default 1)
                - bg_color (str, optional): Background color
                - bold (bool): Bold text
                - align (str): Text alignment
        """
        row = TableRow()
        for cd in cell_defs:
            cell = TableCell(
                colspan=cd.get("colspan", 1),
                rowspan=cd.get("rowspan", 1),
                bg_color=cd.get("bg_color"),
                align=cd.get("align", "LEFT"),
            )
            p = Paragraph(align=cd.get("align", "LEFT"))
            p.add_run(
                cd.get("text", ""),
                font_name=self.font_name,
                font_size_pt=self.font_size_pt,
                bold=cd.get("bold", False),
                color=cd.get("color"),
            )
            cell.paragraphs.append(p)
            row.cells.append(cell)
        self.rows.append(row)
        return row

    @property
    def col_count(self) -> int:
        return len(self.col_widths_mm)

    @property
    def row_count(self) -> int:
        return len(self.rows)
