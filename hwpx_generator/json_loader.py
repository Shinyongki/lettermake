"""JSON loader — parse JSON content blocks and build an HwpxDocument."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Optional

from .document import HwpxDocument
from .presets.gov_document import GovDocumentPreset


# ---------------------------------------------------------------------------
# Preset factory
# ---------------------------------------------------------------------------

_PRESET_MAP = {
    "gov_document": GovDocumentPreset,
}


def _build_preset(preset_name: str, options: Dict[str, Any]) -> GovDocumentPreset:
    """Instantiate a preset from its name and option overrides."""
    cls = _PRESET_MAP.get(preset_name)
    if cls is None:
        raise ValueError(
            f"Unknown preset '{preset_name}'. "
            f"Available presets: {', '.join(_PRESET_MAP)}"
        )

    kwargs: Dict[str, Any] = {}

    # Direct scalar mappings
    _SIMPLE_KEYS = {
        "font_body", "font_table", "font_title",
        "size_title", "size_body", "size_table",
        "line_spacing",
        "color_main", "color_sub", "color_accent",
    }
    for key in _SIMPLE_KEYS:
        if key in options:
            kwargs[key] = options[key]

    # Margin dict → individual fields
    margin = options.get("margin")
    if isinstance(margin, dict):
        for side in ("top", "bottom", "left", "right"):
            if side in margin:
                kwargs[f"margin_{side}"] = margin[side]

    # Cell-margin dict → individual fields
    cell_margin = options.get("cell_margin")
    if isinstance(cell_margin, dict):
        if "lr" in cell_margin:
            kwargs["cell_margin_lr"] = cell_margin["lr"]
        if "tb" in cell_margin:
            kwargs["cell_margin_tb"] = cell_margin["tb"]

    return cls(**kwargs)


# ---------------------------------------------------------------------------
# Content-block dispatchers
# ---------------------------------------------------------------------------

def _dispatch_title(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    doc.add_title(block["text"])


def _dispatch_doc_title(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    doc.add_doc_title(block["text"])


def _dispatch_doc_purpose(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    doc.add_doc_purpose(block["text"])


def _dispatch_section(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    kwargs = {}
    if "num" in block:
        kwargs["num"] = block["num"]
    doc.add_section_block(block["text"], **kwargs)


def _dispatch_section_heading(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    doc.add_section_heading(block.get("num", ""), block["text"])


def _dispatch_sub_heading(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    doc.add_sub_heading(block["text"])


def _dispatch_sub_item(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    doc.add_sub_item(block.get("num", ""), block["text"])


def _dispatch_bullet(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    level = block.get("level", 1)
    text = block["text"]
    # Support both numeric (1~5) and string ("li1"~"li5", "sub1") levels
    if level in (5, "li5"):
        doc.add_bullet5(text)
    elif level in (4, "li4"):
        doc.add_bullet4(text)
    elif level in (3, "li3"):
        doc.add_bullet3(text)
    elif level in (2, "li2"):
        doc.add_bullet2(text)
    elif level in ("sub1",):
        doc.add_sub_item("", text)
    else:
        doc.add_bullet1(text)


def _dispatch_note(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    doc.add_note(block["text"])


def _dispatch_paragraph(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    text = block.get("text", "")
    kwargs: Dict[str, Any] = {}
    if "align" in block:
        kwargs["align"] = block["align"]
    if "bold" in block:
        kwargs["bold"] = block["bold"]
    if "italic" in block:
        kwargs["italic"] = block["italic"]
    if "underline" in block:
        kwargs["underline"] = block["underline"]
    if "color" in block:
        kwargs["color"] = block["color"]
    if "font_name" in block:
        kwargs["font_name"] = block["font_name"]
    if "font_size_pt" in block:
        kwargs["font_size_pt"] = block["font_size_pt"]
    doc.add_paragraph(text, **kwargs)


def _dispatch_table(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    col_widths = block.get("col_widths_mm", [])
    table_kwargs: Dict[str, Any] = {}
    if "border_color" in block:
        table_kwargs["border_color"] = block["border_color"]
    if "header_bg_color" in block:
        table_kwargs["header_bg_color"] = block["header_bg_color"]
    if "header_font_color" in block:
        table_kwargs["header_font_color"] = block["header_font_color"]

    table = doc.add_table(col_widths, **table_kwargs)

    if "header" in block:
        table.add_header_row(block["header"])
    for row_data in block.get("rows", []):
        table.add_row(row_data)


def _dispatch_page_break(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    doc.add_page_break()


def _dispatch_image(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    kwargs: Dict[str, Any] = {}
    if "width_mm" in block:
        kwargs["width_mm"] = block["width_mm"]
    if "height_mm" in block:
        kwargs["height_mm"] = block["height_mm"]
    if "align" in block:
        kwargs["align"] = block["align"]
    if "caption" in block:
        kwargs["caption"] = block["caption"]
    doc.add_image(block["src"], **kwargs)


def _dispatch_diagram(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    kwargs: Dict[str, Any] = {}
    for key in ("layout", "direction", "width_mm", "height_mm", "theme_color"):
        if key in block:
            kwargs[key] = block[key]

    diagram = doc.add_diagram(**kwargs)

    for node_def in block.get("nodes", []):
        node_kwargs: Dict[str, Any] = {}
        for key in ("fill_color", "line_color", "font_color", "font_size_pt"):
            if key in node_def:
                node_kwargs[key] = node_def[key]
        diagram.add_node(
            id=node_def["id"],
            label=node_def.get("label", ""),
            shape=node_def.get("shape", "rect"),
            **node_kwargs,
        )

    for edge_def in block.get("edges", []):
        edge_kwargs: Dict[str, Any] = {}
        for key in ("head_style", "tail_style", "tail_sz", "line_color"):
            if key in edge_def:
                edge_kwargs[key] = edge_def[key]
        # Accept both "from"/"to" and "from_id"/"to_id" keys
        from_id = edge_def.get("from_id") or edge_def.get("from", "")
        to_id = edge_def.get("to_id") or edge_def.get("to", "")
        diagram.add_edge(
            from_id=from_id,
            to_id=to_id,
            **edge_kwargs,
        )

    # Caption as a centered paragraph after the diagram
    if "caption" in block:
        doc.add_paragraph(block["caption"], align="CENTER", font_size_pt=9, italic=True)


def _dispatch_chart(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    data = block.get("data", {})
    datasets = data.get("datasets", [])
    doc.add_chart(
        chart_type=block.get("chart_type", "bar"),
        width_mm=block.get("width_mm", 120.0),
        height_mm=block.get("height_mm", 70.0),
        title=block.get("title", ""),
        labels=data.get("labels", []),
        datasets=datasets,
    )

    # Caption as a centered paragraph after the chart
    if "caption" in block:
        doc.add_paragraph(block["caption"], align="CENTER", font_size_pt=9, italic=True)


def _dispatch_svg(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    kwargs: Dict[str, Any] = {}
    if "src" in block:
        kwargs["src"] = block["src"]
    if "svg_string" in block:
        kwargs["svg_string"] = block["svg_string"]
    if "width_mm" in block:
        kwargs["width_mm"] = block["width_mm"]
    if "height_mm" in block:
        kwargs["height_mm"] = block["height_mm"]
    if "align" in block:
        kwargs["align"] = block["align"]
    if "caption" in block:
        kwargs["caption"] = block["caption"]
    doc.add_svg(**kwargs)


def _dispatch_text_box(doc: HwpxDocument, block: Dict[str, Any]) -> None:
    text = block.get("text", "")
    kwargs: Dict[str, Any] = {}
    for key in (
        "border_color", "bg_color",
        "font_name", "font_size_pt", "font_bold", "font_italic", "font_color",
        "padding_mm", "width_mm", "height_mm", "align",
    ):
        if key in block:
            kwargs[key] = block[key]
    doc.add_text_box(text, **kwargs)


# Dispatcher registry
_DISPATCHERS = {
    "title": _dispatch_title,
    "doc_title": _dispatch_doc_title,
    "doc_purpose": _dispatch_doc_purpose,
    "section": _dispatch_section,
    "section_heading": _dispatch_section_heading,
    "sub_heading": _dispatch_sub_heading,
    "sub_item": _dispatch_sub_item,
    "bullet": _dispatch_bullet,
    "note": _dispatch_note,
    "paragraph": _dispatch_paragraph,
    "table": _dispatch_table,
    "page_break": _dispatch_page_break,
    "image": _dispatch_image,
    "diagram": _dispatch_diagram,
    "chart": _dispatch_chart,
    "svg": _dispatch_svg,
    "text_box": _dispatch_text_box,
}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def load_from_json(data: Dict[str, Any]) -> HwpxDocument:
    """Build an HwpxDocument from a parsed JSON dictionary.

    Args:
        data: Dictionary with optional "preset", "preset_options", and
              required "content" list.

    Returns:
        A fully-populated HwpxDocument ready for .save().
    """
    doc = HwpxDocument()

    # Apply preset if specified
    preset_name = data.get("preset")
    if preset_name:
        preset_options = data.get("preset_options", {})
        preset = _build_preset(preset_name, preset_options)
        doc.apply_preset(preset)

    # Process content blocks
    content = data.get("content", [])
    for block in content:
        block_type = block.get("type")
        if block_type is None:
            raise ValueError(f"Content block missing 'type' key: {block}")

        dispatcher = _DISPATCHERS.get(block_type)
        if dispatcher is None:
            raise ValueError(
                f"Unknown content block type '{block_type}'. "
                f"Supported types: {', '.join(sorted(_DISPATCHERS))}"
            )

        dispatcher(doc, block)

    return doc


def load_from_file(path: str) -> HwpxDocument:
    """Read a JSON file and build an HwpxDocument.

    Args:
        path: Path to the JSON file (UTF-8 encoded).

    Returns:
        A fully-populated HwpxDocument ready for .save().
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"JSON file not found: {path}")

    text = p.read_text(encoding="utf-8")
    data = json.loads(text)
    return load_from_json(data)
