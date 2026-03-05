"""Markdown loader — parse a Markdown file and build an HwpxDocument.

Delegates to parsers.markdown_parser.MarkdownParser for the actual parsing.
Provides the same public API that cli.py expects.

Supported Markdown elements:
    # doc_title
    목적: doc_purpose
    ## section (Roman-numeral table block)
    ### sub_heading (○ auto)
    #### sub1 (가.나.다. auto)
    - bullet li1
      - bullet li2
    ※ note
    | table |
    --- page_break
    > text_box
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from .document import HwpxDocument
from .parsers.markdown_parser import MarkdownParser

_PRESET_MAP: dict[str, type] = {}


def _get_preset_map() -> dict[str, type]:
    """Lazy-load preset map to avoid circular imports."""
    if not _PRESET_MAP:
        from .presets.gov_document import GovDocumentPreset
        _PRESET_MAP["gov_document"] = GovDocumentPreset
    return _PRESET_MAP


def load_from_markdown(
    text: str, preset_name: Optional[str] = None
) -> HwpxDocument:
    """Parse Markdown text and build an HwpxDocument.

    Args:
        text: Markdown source text (UTF-8).
        preset_name: Optional preset to apply (e.g. "gov_document").

    Returns:
        A fully-populated HwpxDocument ready for .save().
    """
    doc = HwpxDocument()

    if preset_name:
        presets = _get_preset_map()
        cls = presets.get(preset_name)
        if cls is None:
            raise ValueError(
                f"Unknown preset '{preset_name}'. "
                f"Available: {', '.join(presets)}"
            )
        doc.apply_preset(cls())

    parser = MarkdownParser()
    parser.parse(text, doc)
    return doc


def load_from_md_file(
    path: str, preset_name: Optional[str] = None
) -> HwpxDocument:
    """Read a Markdown file and build an HwpxDocument.

    Args:
        path: Path to the .md file (UTF-8 encoded).
        preset_name: Optional preset to apply. Defaults to "gov_document".

    Returns:
        A fully-populated HwpxDocument ready for .save().
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Markdown file not found: {path}")

    if preset_name is None:
        preset_name = "gov_document"

    text = p.read_text(encoding="utf-8")
    return load_from_markdown(text, preset_name=preset_name)
