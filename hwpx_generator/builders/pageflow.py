"""PageFlowController — automatic page break insertion based on height estimation.

Runs during save() to ensure:
1. Tables don't split across pages (keep-together)
2. Section headings stay with their following content (keep-with-next)
3. Widow/orphan lines are prevented (min 2 lines together)
4. Manual page breaks are respected
"""

from __future__ import annotations

import math
from typing import List, Any, TYPE_CHECKING

if TYPE_CHECKING:
    from ..document import PageSettings


class PageFlowController:
    """Estimates block heights and inserts page breaks to prevent layout issues."""

    # Average character width factor (mm per character at 1pt font).
    # This is a rough heuristic for Korean text mixed with Latin.
    AVG_CHAR_WIDTH_FACTOR = 0.45  # mm per pt per character (approximate)

    def __init__(self, page_settings: "PageSettings") -> None:
        self.page_settings = page_settings
        self.content_width_mm = page_settings.content_width_mm
        self.content_height_mm = page_settings.content_height_mm

    def process(self, blocks: List[Any]) -> List[Any]:
        """Main entry point: analyze blocks and insert page breaks where needed.

        Returns a new list of blocks with PageBreak objects inserted.
        """
        from ..elements.page_break import PageBreak
        from ..elements.paragraph import Paragraph
        from ..elements.table import Table

        result: List[Any] = []
        remaining_height = self.content_height_mm

        i = 0
        while i < len(blocks):
            block = blocks[i]

            # Manual page break — reset remaining height
            if isinstance(block, PageBreak):
                result.append(block)
                remaining_height = self.content_height_mm
                i += 1
                continue

            block_height = self.estimate_height(block)

            # --- Rule 1: Table keep-together ---
            if isinstance(block, Table) and block.keep_together:
                if block_height > remaining_height and remaining_height < self.content_height_mm:
                    # Table won't fit, insert page break before it
                    result.append(PageBreak())
                    remaining_height = self.content_height_mm

            # --- Rule 2: Keep-with-next (section headings) ---
            if isinstance(block, Paragraph) and block.keep_with_next:
                group_height = self._compute_keep_with_next_group_height(blocks, i)
                if group_height > remaining_height and remaining_height < self.content_height_mm:
                    # Heading group won't fit, insert page break before heading
                    result.append(PageBreak())
                    remaining_height = self.content_height_mm

            # --- Rule 3: Widow/orphan control ---
            if isinstance(block, Paragraph) and not isinstance(block, type(PageBreak())):
                line_count = self._estimate_line_count(block)
                if line_count >= 2:
                    line_height = self._line_height_mm(block)
                    # Check for orphan: only first line fits on current page
                    first_line_height = line_height + self._space_before_mm(block)
                    if (first_line_height <= remaining_height
                            < first_line_height + line_height
                            and remaining_height < self.content_height_mm):
                        # Only 1 line would fit — move entire paragraph to next page
                        result.append(PageBreak())
                        remaining_height = self.content_height_mm

                    # Check for widow: only last line would go to next page
                    all_but_last = block_height - line_height
                    if (all_but_last <= remaining_height < block_height
                            and remaining_height < self.content_height_mm
                            and line_count > 2):
                        # All but last line fits — push paragraph so at least 2 lines
                        # go to the next page. We insert a break before this block.
                        result.append(PageBreak())
                        remaining_height = self.content_height_mm

            # Add the block
            result.append(block)

            # Update remaining height
            if block_height >= remaining_height:
                # Block fills the page (or more) — calculate how many pages it spans
                overflow = block_height - remaining_height
                if overflow > 0:
                    full_pages = math.floor(overflow / self.content_height_mm)
                    remaining_height = self.content_height_mm - (
                        overflow - full_pages * self.content_height_mm
                    )
                else:
                    remaining_height = self.content_height_mm
            else:
                remaining_height -= block_height

            i += 1

        return result

    def estimate_height(self, block: Any) -> float:
        """Estimate the height of a block in mm."""
        from ..elements.paragraph import Paragraph
        from ..elements.table import Table
        from ..elements.page_break import PageBreak

        if isinstance(block, PageBreak):
            return 0.0

        if isinstance(block, Paragraph):
            return self._estimate_paragraph_height(block)

        if isinstance(block, Table):
            return self._estimate_table_height(block)

        # Try Image — use height_mm if available
        if hasattr(block, 'height_mm') and block.height_mm:
            h = block.height_mm
            # Add caption height estimate if present
            if hasattr(block, 'caption') and block.caption:
                h += 6.0  # ~6mm for a caption line
            return h

        # Unknown block type — give a minimal default
        return 5.0

    def _estimate_paragraph_height(self, para) -> float:
        """Estimate paragraph height in mm.

        Formula: line_height * line_count + space_before + space_after
        where line_height = font_size_pt * line_spacing_pct / 100 * 0.3528
        """
        line_count = self._estimate_line_count(para)
        line_h = self._line_height_mm(para)
        space_before = self._space_before_mm(para)
        space_after = self._space_after_mm(para)

        return line_h * line_count + space_before + space_after

    def _estimate_table_height(self, table) -> float:
        """Estimate table height in mm.

        Formula: row_height * row_count + 2mm margin
        where row_height = font_size_pt * 1.6 * 0.3528 + cell_margin_tb * 2
        """
        font_size = table.font_size_pt or 10.0
        row_height = font_size * 1.6 * 0.3528 + table.cell_margin_tb_mm * 2
        return row_height * table.row_count + 2.0

    def _line_height_mm(self, para) -> float:
        """Calculate single line height in mm for a paragraph."""
        font_size = para.font_size_pt or 13.0
        line_spacing_pct = para.line_spacing_value if para.line_spacing_type == "PERCENT" else 160.0
        return font_size * line_spacing_pct / 100.0 * 0.3528

    def _space_before_mm(self, para) -> float:
        """Convert space_before_pt to mm."""
        return para.space_before_pt * 0.3528

    def _space_after_mm(self, para) -> float:
        """Convert space_after_pt to mm."""
        return para.space_after_pt * 0.3528

    def _estimate_line_count(self, para) -> int:
        """Estimate the number of lines a paragraph will occupy.

        Uses heuristic: ceil(text_length * avg_char_width / content_width)
        """
        text = para.plain_text if hasattr(para, 'plain_text') else ""
        if not text:
            return 1  # Empty paragraph still takes one line

        font_size = para.font_size_pt or 13.0
        avg_char_width = font_size * self.AVG_CHAR_WIDTH_FACTOR
        text_width = len(text) * avg_char_width

        # Account for indentation reducing available width
        available_width = self.content_width_mm - para.indent_left_mm - para.indent_right_mm
        if available_width <= 0:
            available_width = self.content_width_mm

        line_count = math.ceil(text_width / available_width)
        return max(1, line_count)

    def _compute_keep_with_next_group_height(
        self, blocks: list, start_idx: int
    ) -> float:
        """Compute the combined height of a keep-with-next group.

        For a heading at start_idx, find how many following blocks it must
        stay with, then sum their heights.
        """
        from ..elements.paragraph import Paragraph

        block = blocks[start_idx]
        total_height = self.estimate_height(block)

        # Determine how many blocks to keep together
        keep_count = self._get_keep_count(block)

        # Sum the heights of the following blocks
        for j in range(1, keep_count + 1):
            idx = start_idx + j
            if idx >= len(blocks):
                break
            total_height += self.estimate_height(blocks[idx])

        return total_height

    def _get_keep_count(self, block: Any) -> int:
        """Determine how many following blocks a heading should keep with.

        Based on spec:
        - section heading (contains '\\u25a0'): keep with next 3 blocks
        - sub heading (contains '\\u25cb'): keep with next 1 block
        - sub item (contains Korean letters like '\\uac00'~): keep with next 1 block
        Default for any keep_with_next paragraph: 1 block.
        """
        from ..elements.paragraph import Paragraph

        if not isinstance(block, Paragraph):
            return 0

        text = block.plain_text if hasattr(block, 'plain_text') else ""

        # Section heading: starts with black square
        if text.startswith("\u25a0"):  # ■
            return 3

        # Sub heading: starts with white circle
        if text.startswith("\u25cb"):  # ○
            return 1

        # Sub item: starts with Korean syllable followed by period (가. 나. 다.)
        korean_sub_prefixes = list("가나다라마바사아자차카타파하")
        for prefix in korean_sub_prefixes:
            if text.startswith(prefix + ".") or text.startswith(prefix + " "):
                return 1

        # Default for any keep_with_next paragraph
        return 1
