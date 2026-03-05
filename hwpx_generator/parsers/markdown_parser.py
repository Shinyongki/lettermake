"""Markdown → HwpxDocument parser.

Parses Markdown text and drives HwpxDocument methods directly,
producing the same block list that the JSON loader would create.
No intermediate JSON file is written.

Mapping rules:
    # text              → doc_title
    목적: text          → doc_purpose
    ## text             → section (Roman-numeral table block)
    ### text            → sub_heading (○ auto)
    #### text           → bullet sub1 (가.나.다. auto)
    - text              → bullet li1
      - text            → bullet li2
    ※ text              → note
    | col | col |       → table (standard markdown)
    ---                 → page_break
    > text              → text_box
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import TYPE_CHECKING, List, Optional, Tuple

if TYPE_CHECKING:
    from ..document import HwpxDocument

# Korean sub-numbering sequence: 가, 나, 다, ...
_KOREAN_SUBS = list("가나다라마바사아자차카타파하")

# Regex patterns
_RE_DOC_TITLE = re.compile(r"^#\s+(.+)$")
_RE_DOC_PURPOSE = re.compile(r"^목적:\s*(.+)$")
_RE_SECTION = re.compile(r"^##\s+(.+)$")
_RE_SUB_HEADING = re.compile(r"^###\s+(.+)$")
_RE_SUB1 = re.compile(r"^####\s+(.+)$")
_RE_BULLET2 = re.compile(r"^[ \t]{2,}-\s+(.+)$")
_RE_BULLET1 = re.compile(r"^-\s+(.+)$")
_RE_NOTE = re.compile(r"^※\s*(.+)$")
_RE_TABLE_ROW = re.compile(r"^\|(.+)\|$")
_RE_TABLE_SEP = re.compile(r"^\|[\s:]*-[-\s:|]*\|$")
_RE_PAGE_BREAK = re.compile(r"^-{3,}\s*$")
_RE_BLOCKQUOTE = re.compile(r"^>\s*(.*)$")


def _parse_table_cells(line: str) -> List[str]:
    """Extract cell texts from a markdown table row like '| a | b |'."""
    inner = line.strip().strip("|")
    return [cell.strip() for cell in inner.split("|")]


def _estimate_col_widths(
    headers: List[str],
    rows: List[List[str]],
    total_width_mm: float = 170.0,
) -> List[float]:
    """Estimate column widths proportionally based on content length.

    Uses a simple heuristic: wider columns for longer content,
    with a minimum width to keep things readable.
    """
    ncols = len(headers)
    if ncols == 0:
        return []

    # Compute max content length per column
    max_lens = [len(h) for h in headers]
    for row in rows:
        for i, cell in enumerate(row):
            if i < len(max_lens):
                max_lens[i] = max(max_lens[i], len(cell))

    # Each Korean/CJK char ≈ 2 units width
    total_len = sum(max(ml, 2) for ml in max_lens)
    if total_len == 0:
        return [total_width_mm / ncols] * ncols

    widths = []
    for ml in max_lens:
        ratio = max(ml, 2) / total_len
        widths.append(round(total_width_mm * ratio, 1))

    # Adjust rounding errors
    diff = total_width_mm - sum(widths)
    if abs(diff) > 0.01:
        widths[-1] = round(widths[-1] + diff, 1)

    return widths


class MarkdownParser:
    """Parse Markdown text and populate an HwpxDocument."""

    def __init__(self) -> None:
        self._section_counter: int = 0
        self._sub1_counter: int = 0  # resets per section/sub_heading

    def parse(self, text: str, doc: "HwpxDocument") -> None:
        """Parse markdown *text* and add blocks to *doc*.

        Args:
            text: Raw markdown string (UTF-8).
            doc: HwpxDocument to populate.
        """
        lines = text.splitlines()
        idx = 0

        while idx < len(lines):
            line = lines[idx]
            stripped = line.rstrip()

            # --- blank line: skip ---
            if not stripped:
                idx += 1
                continue

            # --- page break: --- ---
            if _RE_PAGE_BREAK.match(stripped):
                doc.add_page_break()
                idx += 1
                continue

            # --- doc_title: # ---
            m = _RE_DOC_TITLE.match(stripped)
            if m:
                title_text = m.group(1).strip()
                # Collect continuation lines (non-blank, non-heading, non-special)
                idx += 1
                while idx < len(lines):
                    cont = lines[idx].rstrip()
                    if not cont:
                        break
                    if cont.startswith("#") or cont.startswith("-") or cont.startswith("|") or cont.startswith(">") or cont.startswith("※") or cont.startswith("목적:"):
                        break
                    title_text += "\n" + cont
                    idx += 1
                doc.add_doc_title(title_text)
                continue

            # --- doc_purpose: 목적: ---
            m = _RE_DOC_PURPOSE.match(stripped)
            if m:
                doc.add_doc_purpose(m.group(1).strip())
                idx += 1
                continue

            # --- section block: ## ---
            m = _RE_SECTION.match(stripped)
            if m:
                heading_text = m.group(1).strip()
                # Strip leading arabic number: "1. 텍스트" → "텍스트"
                heading_text = re.sub(r'^\d+\.\s*', '', heading_text)
                self._sub1_counter = 0  # reset sub-numbering
                doc.add_section_block(heading_text)
                idx += 1
                continue

            # --- sub heading: ### ---
            m = _RE_SUB_HEADING.match(stripped)
            if m:
                self._sub1_counter = 0  # reset sub-numbering
                doc.add_sub_heading(m.group(1).strip())
                idx += 1
                continue

            # --- sub1: #### (가.나.다. auto) ---
            m = _RE_SUB1.match(stripped)
            if m:
                sub_text = m.group(1).strip()
                if self._sub1_counter < len(_KOREAN_SUBS):
                    korean_num = _KOREAN_SUBS[self._sub1_counter]
                else:
                    korean_num = str(self._sub1_counter + 1)
                self._sub1_counter += 1
                doc.add_sub_item(korean_num, sub_text)
                idx += 1
                continue

            # --- note: ※ ---
            m = _RE_NOTE.match(stripped)
            if m:
                doc.add_note(m.group(1).strip())
                idx += 1
                continue

            # --- blockquote → doc_purpose: > ---
            m = _RE_BLOCKQUOTE.match(stripped)
            if m:
                box_lines = [m.group(1)]
                idx += 1
                while idx < len(lines):
                    bm = _RE_BLOCKQUOTE.match(lines[idx].rstrip())
                    if bm:
                        box_lines.append(bm.group(1))
                        idx += 1
                    else:
                        break
                box_text = "\n".join(box_lines).strip()
                doc.add_doc_purpose(box_text)
                continue

            # --- table: | ... | ---
            m = _RE_TABLE_ROW.match(stripped)
            if m:
                idx = self._parse_table(lines, idx, doc)
                continue

            # --- bullet2: indented - ---
            m = _RE_BULLET2.match(stripped)
            if m:
                doc.add_bullet2(m.group(1).strip())
                idx += 1
                continue

            # --- bullet1: - ---
            m = _RE_BULLET1.match(stripped)
            if m:
                doc.add_bullet1(m.group(1).strip())
                idx += 1
                continue

            # --- fallback: plain paragraph ---
            doc.add_paragraph(stripped)
            idx += 1

    def _parse_table(
        self,
        lines: List[str],
        start_idx: int,
        doc: "HwpxDocument",
    ) -> int:
        """Parse a contiguous markdown table starting at *start_idx*.

        Returns the index of the first line after the table.
        """
        idx = start_idx
        header_cells: Optional[List[str]] = None
        data_rows: List[List[str]] = []
        found_separator = False

        while idx < len(lines):
            stripped = lines[idx].rstrip()
            if not stripped:
                break

            # Table separator line (|---|---|)
            if _RE_TABLE_SEP.match(stripped):
                found_separator = True
                idx += 1
                continue

            m = _RE_TABLE_ROW.match(stripped)
            if not m:
                break

            cells = _parse_table_cells(stripped)

            if not found_separator and header_cells is None:
                header_cells = cells
            else:
                data_rows.append(cells)
            idx += 1

        # Build the table via HwpxDocument
        ncols = 0
        if header_cells:
            ncols = len(header_cells)
        elif data_rows:
            ncols = len(data_rows[0])

        if ncols == 0:
            return idx

        col_widths = _estimate_col_widths(
            header_cells or [],
            data_rows,
            total_width_mm=doc.page_settings.content_width_mm - 2.0,
        )

        table = doc.add_table(col_widths)
        if header_cells:
            table.add_header_row(header_cells)
        for row in data_rows:
            # Pad or trim to match column count
            padded = row[:ncols]
            while len(padded) < ncols:
                padded.append("")
            table.add_row(padded)

        return idx

    def parse_file(self, path: str, doc: "HwpxDocument") -> None:
        """Read a markdown file and add its blocks to *doc*.

        Args:
            path: Path to the .md file (UTF-8).
            doc: HwpxDocument to populate.
        """
        p = Path(path)
        if not p.exists():
            raise FileNotFoundError(f"Markdown file not found: {path}")
        text = p.read_text(encoding="utf-8")
        self.parse(text, doc)
