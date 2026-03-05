"""HwpxDocument — top-level document object."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Any, Optional, Dict

from .builders.package_builder import build_hwpx_package
from .builders.pageflow import PageFlowController
from .styles import StyleManager, Style, CharProps, ParaProps
from .elements.paragraph import Paragraph, TextRun
from .elements.table import Table, TableRow, TableCell
from .elements.page_break import PageBreak
from .elements.image import Image
from .elements.diagram import Diagram, Node, Edge
from .elements.chart import Chart, ChartData, ChartDataset
from .elements.svg_element import SvgElement
from .elements.text_box import TextBox
from .elements.doc_title import DocTitle
from .elements.doc_purpose import DocPurpose
from .elements.section_block import SectionBlock
from .presets.gov_document import GovDocumentPreset, BulletStyle

# Korean sub-numbering: 가, 나, 다, ...
KOREAN_SUB_NUMS = list("가나다라마바사아자차카타파하")


@dataclass
class PageSettings:
    """Page layout settings (all dimensions in mm)."""

    width_mm: float = 210.0
    height_mm: float = 297.0
    orientation: str = "portrait"
    margin_top: float = 20.0
    margin_bottom: float = 20.0
    margin_left: float = 20.0
    margin_right: float = 20.0
    header_mm: float = 0.0
    footer_mm: float = 0.0

    @property
    def content_width_mm(self) -> float:
        if self.orientation == "landscape":
            return self.height_mm - self.margin_left - self.margin_right
        return self.width_mm - self.margin_left - self.margin_right

    @property
    def content_height_mm(self) -> float:
        if self.orientation == "landscape":
            return self.width_mm - self.margin_top - self.margin_bottom
        return self.height_mm - self.margin_top - self.margin_bottom


class HwpxDocument:
    """Top-level HWPX document object."""

    def __init__(self) -> None:
        self.page_settings = PageSettings()
        self.style_manager = StyleManager()
        self.blocks: List[Any] = []
        self._bullet_styles: Dict[str, BulletStyle] = {}
        self._preset: Optional[GovDocumentPreset] = None
        # Auto-numbering counters for li4/li5
        self._li4_counter: int = 0
        self._li5_counter: int = 0

    # --- Preset ---

    def apply_preset(self, preset: GovDocumentPreset) -> None:
        """Apply a style preset to the document."""
        self._preset = preset
        self._bullet_styles = dict(preset.bullet_styles)

        # Apply page settings
        ps = self.page_settings
        ps.margin_top = preset.margin_top
        ps.margin_bottom = preset.margin_bottom
        ps.margin_left = preset.margin_left
        ps.margin_right = preset.margin_right

        # Register default style with preset fonts
        default_style = Style(
            name="바탕글",
            char_props=CharProps(
                font_name=preset.font_body,
                font_size_pt=preset.size_body,
            ),
            para_props=ParaProps(
                line_spacing_value=preset.line_spacing,
            ),
        )
        self.style_manager.register_style(default_style)

    def _get_bullet_style(self, key: str) -> BulletStyle:
        """Get bullet style, falling back to defaults."""
        if key in self._bullet_styles:
            return self._bullet_styles[key]
        # Return a basic default
        return BulletStyle(size=13.0)

    def _body_font(self) -> Optional[str]:
        return self._preset.font_body if self._preset else None

    def _body_size(self) -> float:
        return self._preset.size_body if self._preset else 13.0

    # --- Title ---

    def add_title(self, text: str) -> Paragraph:
        """Add a document title (centered, bold, large)."""
        bs = self._get_bullet_style("title")
        return self._add_bullet_paragraph(bs, text)

    # --- DocTitle (decorated title with top/bottom images) ---

    def add_doc_title(self, text: str) -> DocTitle:
        """Add a decorated document title with top/bottom decoration lines.

        Args:
            text: Title text displayed between decoration lines.

        Returns:
            The DocTitle element added to the document.
        """
        dt = DocTitle(text=text)
        self.blocks.append(dt)
        return dt

    # --- DocPurpose (1x1 table with purpose text) ---

    def add_doc_purpose(self, text: str) -> DocPurpose:
        """Add a document purpose block (text inside a framed table).

        Args:
            text: Purpose text displayed inside the bordered frame.

        Returns:
            The DocPurpose element added to the document.
        """
        font = self._preset.font_table if self._preset else "맑은 고딕"
        dp = DocPurpose(text=text, font_name=font)
        self.blocks.append(dp)
        return dp

    # --- SectionBlock (section divider with number + title) ---

    def add_section_block(self, text: str, *, num: int = None) -> SectionBlock:
        """Add a section divider block with auto-incrementing number.

        Args:
            text: Section title text.
            num: Manual section number (None = auto-increment).

        Returns:
            The SectionBlock element added to the document.
        """
        sb = SectionBlock(text=text, num=num)
        self.blocks.append(sb)
        # Reset bullet counters at section boundary
        self._li4_counter = 0
        self._li5_counter = 0
        return sb

    # --- Section heading: ■ 1. 제목 ---

    def add_section_heading(self, num, text: str) -> HwpxDocument:
        """Add a section heading: ■ {num}. {text}

        Args:
            num: Section number (e.g. 1, "1", "Ⅰ")
            text: Heading text
        """
        bs = self._get_bullet_style("h1")
        prefix = bs.prefix.replace("{num}", str(num))
        self._add_bullet_paragraph(bs, text, prefix_text=prefix)
        return self

    # --- Sub heading: ○ 소제목 ---

    def add_sub_heading(self, text: str) -> HwpxDocument:
        """Add a sub-heading: ○ {text}"""
        bs = self._get_bullet_style("h2")
        self._add_bullet_paragraph(bs, text, prefix_text=bs.prefix)
        return self

    # --- Sub item: 가. / 나. / 다. ---

    def add_sub_item(self, num: str, text: str) -> HwpxDocument:
        """Add a sub-item: {num}. {text}

        Args:
            num: Korean letter or number (e.g. "가", "나", "1")
            text: Item text
        """
        bs = self._get_bullet_style("sub1")
        prefix = bs.prefix.replace("{num}", str(num))
        self._add_bullet_paragraph(bs, text, prefix_text=prefix)
        return self

    # --- Bullet level 1: \uf06d 항목 (동그라미 특수문자) ---

    def add_bullet1(self, text: str) -> HwpxDocument:
        """Add bullet level 1: \\uf06d prefix (휴먼명조 15pt) + text (body font)."""
        bs = self._get_bullet_style("li1")
        self._add_bullet_paragraph(bs, text, prefix_text=" \uf06d ")
        return self

    # --- Bullet level 2: - 항목 ---

    def add_bullet2(self, text: str) -> HwpxDocument:
        """Add bullet level 2: - text (body font, single run)."""
        font = self._body_font()
        size = self._body_size()
        ls = self._preset.line_spacing if self._preset else 160.0
        para = Paragraph(
            align="JUSTIFY", line_spacing_value=ls,
            font_name=font, font_size_pt=size,
        )
        para.add_run("  - " + text, font_name=font, font_size_pt=size)
        self.blocks.append(para)
        return self

    # --- Bullet level 3: ∙ 항목 ---

    def add_bullet3(self, text: str) -> HwpxDocument:
        """Add bullet level 3: space(body) + ∙ text(휴먼명조 13pt), 2 runs."""
        font = self._body_font()
        size = self._body_size()
        ls = self._preset.line_spacing if self._preset else 160.0
        para = Paragraph(
            align="JUSTIFY", line_spacing_value=ls,
            font_name="휴먼명조", font_size_pt=13.0,
        )
        # Run 1: single space in body font (indent)
        para.add_run(" ", font_name=font, font_size_pt=size)
        # Run 2: ∙ + text in 휴먼명조 13pt
        para.add_run("\u2219 " + text, font_name="휴먼명조", font_size_pt=13.0)
        self.blocks.append(para)
        return self

    # --- Bullet level 4: N. 항목 (auto-number) ---

    def add_bullet4(self, text: str) -> HwpxDocument:
        """Add bullet level 4: auto-numbered (1. 2. 3...), 휴먼명조 13pt."""
        self._li4_counter += 1
        self._li5_counter = 0  # reset li5 when li4 increments
        num = self._li4_counter
        ls = self._preset.line_spacing if self._preset else 160.0
        para = Paragraph(
            align="JUSTIFY", line_spacing_value=ls,
            font_name="휴먼명조", font_size_pt=13.0,
        )
        para.add_run(f"  {num}. " + text, font_name="휴먼명조", font_size_pt=13.0)
        self.blocks.append(para)
        return self

    # --- Bullet level 5: 가. 항목 (auto-letter Korean) ---

    def add_bullet5(self, text: str) -> HwpxDocument:
        """Add bullet level 5: auto-lettered (가. 나. 다...), 휴먼명조 13pt."""
        self._li5_counter += 1
        idx = self._li5_counter - 1
        letter = KOREAN_SUB_NUMS[idx] if idx < len(KOREAN_SUB_NUMS) else str(self._li5_counter)
        ls = self._preset.line_spacing if self._preset else 160.0
        para = Paragraph(
            align="JUSTIFY", line_spacing_value=ls,
            font_name="휴먼명조", font_size_pt=13.0,
        )
        para.add_run(f"   {letter}. " + text, font_name="휴먼명조", font_size_pt=13.0)
        self.blocks.append(para)
        return self

    # --- Note: ※ 주석 ---

    def add_note(self, text: str) -> HwpxDocument:
        """Add a note: ※ {text}"""
        bs = self._get_bullet_style("note")
        self._add_bullet_paragraph(bs, text, prefix_text=bs.prefix)
        return self

    # --- Low-level bullet paragraph builder ---

    def _add_bullet_paragraph(
        self, bs: BulletStyle, text: str, *, prefix_text: str = ""
    ) -> Paragraph:
        """Create a paragraph from a BulletStyle and add it to blocks.

        If prefix_text is provided and BulletStyle has prefix_font_name/prefix_size,
        the prefix is rendered as a separate run with its own font/size.
        """
        font = bs.font_name or self._body_font()
        indent_first = -bs.hanging_mm if bs.hanging_mm else 0.0

        para = Paragraph(
            align=bs.align,
            line_spacing_value=self._preset.line_spacing if self._preset else 160.0,
            space_before_pt=bs.space_before_pt,
            space_after_pt=bs.space_after_pt,
            indent_left_mm=bs.indent_mm,
            indent_first_mm=indent_first,
            keep_with_next=bs.keep_with_next,
            font_name=font,
            font_size_pt=bs.size,
            bold=bs.bold,
            italic=bs.italic,
            color=bs.color,
        )

        if prefix_text and (bs.prefix_font_name or bs.prefix_size):
            # Prefix in different font/size → separate run
            pfx_font = bs.prefix_font_name or font
            pfx_size = bs.prefix_size or bs.size
            para.add_run(
                prefix_text,
                font_name=pfx_font,
                font_size_pt=pfx_size,
                bold=bs.bold,
                italic=bs.italic,
                color=bs.color,
            )
            para.add_run(
                text,
                font_name=font,
                font_size_pt=bs.size,
                bold=bs.bold,
                italic=bs.italic,
                color=bs.color,
            )
        else:
            # Single run (prefix + text combined)
            combined = prefix_text + text if prefix_text else text
            para.add_run(
                combined,
                font_name=font,
                font_size_pt=bs.size,
                bold=bs.bold,
                italic=bs.italic,
                color=bs.color,
            )

        self.blocks.append(para)
        return para

    # --- Raw paragraph methods (from Phase 2) ---

    def add_paragraph(
        self,
        text: str = "",
        *,
        font_name: Optional[str] = None,
        font_size_pt: Optional[float] = None,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        color: Optional[str] = None,
        align: str = "JUSTIFY",
        line_spacing_type: str = "PERCENT",
        line_spacing_value: float = 160.0,
        space_before_pt: float = 0.0,
        space_after_pt: float = 0.0,
        indent_left_mm: float = 0.0,
        indent_right_mm: float = 0.0,
        indent_first_mm: float = 0.0,
        keep_with_next: bool = False,
        widow_orphan: bool = False,
        style_name: Optional[str] = None,
    ) -> Paragraph:
        """Add a paragraph with a single text run."""
        para = Paragraph(
            align=align,
            line_spacing_type=line_spacing_type,
            line_spacing_value=line_spacing_value,
            space_before_pt=space_before_pt,
            space_after_pt=space_after_pt,
            indent_left_mm=indent_left_mm,
            indent_right_mm=indent_right_mm,
            indent_first_mm=indent_first_mm,
            keep_with_next=keep_with_next,
            widow_orphan=widow_orphan,
            font_name=font_name,
            font_size_pt=font_size_pt,
            bold=bold,
            italic=italic,
            color=color,
            style_name=style_name,
        )
        if text:
            para.add_run(
                text,
                font_name=font_name,
                font_size_pt=font_size_pt,
                bold=bold,
                italic=italic,
                underline=underline,
                color=color,
            )
        self.blocks.append(para)
        return para

    def add_empty_paragraph(self) -> Paragraph:
        """Add an empty paragraph (blank line)."""
        para = Paragraph()
        self.blocks.append(para)
        return para

    # --- Page break ---

    def add_page_break(self) -> PageBreak:
        """Insert an explicit page break."""
        pb = PageBreak()
        self.blocks.append(pb)
        return pb

    # --- Image ---

    def add_image(
        self,
        src: str,
        *,
        width_mm: float = 80.0,
        height_mm: Optional[float] = None,
        align: str = "center",
        caption: Optional[str] = None,
    ) -> Image:
        """Add an image to the document.

        Args:
            src: Path to the image file (PNG or JPG).
            width_mm: Display width in millimeters.
            height_mm: Display height in mm (auto-calculated from aspect ratio if omitted).
            align: Horizontal alignment — "left", "center", or "right".
            caption: Optional caption text below the image.

        Returns:
            The Image element added to the document.
        """
        img = Image(
            src=src,
            width_mm=width_mm,
            height_mm=height_mm,
            align=align,
            caption=caption,
        )
        # Resolve pixel dimensions and auto-calculate height
        img.resolve_dimensions()
        self.blocks.append(img)
        return img

    def add_image_from_bytes(
        self,
        data: bytes,
        ext: str,
        *,
        width_mm: float = 80.0,
        height_mm: Optional[float] = None,
        align: str = "center",
        caption: Optional[str] = None,
    ) -> Image:
        """Add an image from raw bytes (useful for testing or generated images).

        Args:
            data: Raw image bytes.
            ext: File extension ("png" or "jpg").
            width_mm: Display width in millimeters.
            height_mm: Display height in mm (auto-calculated if omitted).
            align: Horizontal alignment — "left", "center", or "right".
            caption: Optional caption text below the image.

        Returns:
            The Image element added to the document.
        """
        ext_clean = ext.lower().lstrip(".")
        img = Image(
            src=f"__memory__.{ext_clean}",
            width_mm=width_mm,
            height_mm=height_mm,
            align=align,
            caption=caption,
        )
        img.resolve_dimensions_from_bytes(data, ext_clean)
        img._image_data = data  # type: ignore[attr-defined]
        self.blocks.append(img)
        return img

    # --- Diagram ---

    def add_diagram(
        self,
        *,
        layout: str = "step_flow",
        direction: str = "horizontal",
        width_mm: float = 160.0,
        height_mm: float = 30.0,
        theme_color: Optional[str] = None,
    ) -> Diagram:
        """Add a diagram to the document.

        Args:
            layout: Layout algorithm — "step_flow" or "free".
            direction: Flow direction — "horizontal" or "vertical".
            width_mm: Total diagram width in mm.
            height_mm: Total diagram height in mm.
            theme_color: Hex color for lines/fills (default from preset).

        Returns:
            The Diagram element (call .add_node() and .add_edge() on it).
        """
        color = theme_color or (
            self._preset.color_main if self._preset else "1F4E79"
        )
        diagram = Diagram(
            layout=layout,
            direction=direction,
            width_mm=width_mm,
            height_mm=height_mm,
            theme_color=color,
        )
        self.blocks.append(diagram)
        return diagram

    # --- Chart ---

    def add_chart(
        self,
        *,
        chart_type: str = "bar",
        width_mm: float = 120.0,
        height_mm: float = 70.0,
        title: str = "",
        labels: Optional[List[str]] = None,
        datasets: Optional[List[dict]] = None,
    ) -> Chart:
        """Add a chart to the document.

        Args:
            chart_type: One of "bar", "barh", "line", "pie", "stacked_bar".
            width_mm: Chart display width in mm.
            height_mm: Chart display height in mm.
            title: Chart title text.
            labels: Category labels (x-axis).
            datasets: List of dicts with keys: label, values, color.

        Returns:
            The Chart element.
        """
        ds_list = []
        for ds_dict in (datasets or []):
            ds_list.append(ChartDataset(
                label=ds_dict.get("label", ""),
                values=ds_dict.get("values", []),
                color=ds_dict.get("color"),
            ))
        chart = Chart(
            chart_type=chart_type,
            width_mm=width_mm,
            height_mm=height_mm,
            title=title,
            data=ChartData(
                labels=labels or [],
                datasets=ds_list,
            ),
        )
        self.blocks.append(chart)
        return chart

    # --- Table ---

    def add_table(
        self,
        col_widths: List[float],
        *,
        border_color: str = "000000",
        border_width: str = "0.12mm",
        cell_margin_lr: Optional[float] = None,
        cell_margin_tb: Optional[float] = None,
        header_bg_color: str = "D9D9D9",
        header_font_color: str = "000000",
        header_font_bold: bool = True,
        header_align: str = "CENTER",
        font_name: Optional[str] = None,
        font_size_pt: Optional[float] = None,
        keep_together: bool = True,
    ) -> Table:
        """Add a table to the document.

        Args:
            col_widths: Column widths in mm.
        """
        cm_lr = cell_margin_lr if cell_margin_lr is not None else (
            self._preset.cell_margin_lr if self._preset else 3.0)
        cm_tb = cell_margin_tb if cell_margin_tb is not None else (
            self._preset.cell_margin_tb if self._preset else 2.0)
        fn = font_name or (self._preset.font_table if self._preset else None)
        fs = font_size_pt or (self._preset.size_table if self._preset else None)

        table = Table(
            col_widths_mm=col_widths,
            border_color=border_color,
            border_width=border_width,
            cell_margin_lr_mm=cm_lr,
            cell_margin_tb_mm=cm_tb,
            header_bg_color=header_bg_color,
            header_font_color=header_font_color,
            header_font_bold=header_font_bold,
            header_align=header_align,
            font_name=fn,
            font_size_pt=fs,
            keep_together=keep_together,
        )
        self.blocks.append(table)
        return table

    # --- SVG ---

    def add_svg(
        self,
        *,
        src: str = "",
        svg_string: str = "",
        width_mm: float = 160.0,
        height_mm: Optional[float] = None,
        align: str = "center",
        caption: Optional[str] = None,
    ) -> SvgElement:
        """Add an SVG element to the document (converted to native shapes).

        Args:
            src: Path to the SVG file.
            svg_string: Raw SVG XML string (alternative to src).
            width_mm: Display width in millimeters.
            height_mm: Display height in mm (auto-calculated if omitted).
            align: Horizontal alignment -- "left", "center", or "right".
            caption: Optional caption text below the SVG.

        Returns:
            The SvgElement added to the document.
        """
        svg = SvgElement(
            src=src,
            svg_string=svg_string,
            width_mm=width_mm,
            height_mm=height_mm,
            align=align,
            caption=caption,
        )
        self.blocks.append(svg)
        return svg

    # --- TextBox ---

    def add_text_box(
        self,
        text: str,
        *,
        border_color: str = "000000",
        bg_color: Optional[str] = None,
        font_name: Optional[str] = None,
        font_size_pt: Optional[float] = None,
        font_bold: bool = False,
        font_italic: bool = False,
        font_color: Optional[str] = None,
        padding_mm: float = 4.0,
        width_mm: Optional[float] = None,
        height_mm: Optional[float] = None,
        align: str = "left",
    ) -> TextBox:
        """Add a highlighted text box to the document.

        Args:
            text: Text content displayed inside the box.
            border_color: Border hex color (e.g. "C00000").
            bg_color: Background fill hex color (e.g. "FFF2CC").
            font_name: Font name; None uses document default.
            font_size_pt: Font size in points.
            font_bold: Whether text is bold.
            font_italic: Whether text is italic.
            font_color: Text hex color; None uses "000000".
            padding_mm: Internal padding in mm.
            width_mm: Box width in mm; None = auto.
            height_mm: Box height in mm; None = auto.
            align: Horizontal alignment.

        Returns:
            The TextBox element added to the document.
        """
        fn = font_name or (self._preset.font_body if self._preset else None)
        fs = font_size_pt or (self._preset.size_body if self._preset else None)

        tb = TextBox(
            text=text,
            border_color=border_color,
            bg_color=bg_color,
            font_name=fn,
            font_size_pt=fs,
            font_bold=font_bold,
            font_italic=font_italic,
            font_color=font_color,
            padding_mm=padding_mm,
            width_mm=width_mm,
            height_mm=height_mm,
            align=align,
        )
        self.blocks.append(tb)
        return tb

    # --- Style registration ---

    def register_style(self, style: Style) -> None:
        self.style_manager.register_style(style)

    # --- Markdown loading ---

    def load_markdown(self, path: str) -> None:
        """Parse a Markdown file and add its blocks to this document.

        Args:
            path: Path to the .md file (UTF-8).
        """
        from .parsers.markdown_parser import MarkdownParser

        parser = MarkdownParser()
        parser.parse_file(path, self)

    def load_markdown_string(self, text: str) -> None:
        """Parse a Markdown string and add its blocks to this document.

        Args:
            text: Raw Markdown text.
        """
        from .parsers.markdown_parser import MarkdownParser

        parser = MarkdownParser()
        parser.parse(text, self)

    # --- Save ---

    def save(self, path: str, *, auto_page_flow: bool = True) -> None:
        """Build and write the .hwpx file.

        Args:
            path: Output file path.
            auto_page_flow: If True (default), run PageFlowController to
                automatically insert page breaks where needed.
        """
        if auto_page_flow:
            controller = PageFlowController(self.page_settings)
            self.blocks = controller.process(self.blocks)
        self.style_manager.finalize(self.blocks)
        build_hwpx_package(self, Path(path))
