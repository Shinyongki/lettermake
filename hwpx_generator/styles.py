"""StyleManager — document-level style definitions and ID tracking."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .utils import mm_to_hwp, pt_to_hwp

# Reserved paraPr ID for DocTitle paragraph (CENTER, 120% line spacing).
# This ID is always hardcoded in header.xml and referenced by _render_doc_title_v7().
RESERVED_DOC_TITLE_PP_ID = 10


@dataclass
class CharProps:
    """Character-level properties (maps to hh:charPr)."""

    font_name: str = "함초롬돋움"
    font_size_pt: float = 10.0
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    color: str = "000000"
    char_spacing: int = 0  # character spacing (hh:spacing value per lang)

    def key(self) -> tuple:
        return (self.font_name, self.font_size_pt, self.bold, self.italic,
                self.underline, self.strike, self.color, self.char_spacing)


@dataclass
class ParaProps:
    """Paragraph-level properties (maps to hh:paraPr)."""

    align: str = "JUSTIFY"
    line_spacing_type: str = "PERCENT"
    line_spacing_value: float = 160.0
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    indent_left_mm: float = 0.0
    indent_right_mm: float = 0.0
    indent_first_mm: float = 0.0
    keep_with_next: bool = False
    widow_orphan: bool = False
    page_break_before: bool = False

    def key(self) -> tuple:
        return (self.align, self.line_spacing_type, self.line_spacing_value,
                self.space_before_pt, self.space_after_pt,
                self.indent_left_mm, self.indent_right_mm, self.indent_first_mm,
                self.keep_with_next, self.widow_orphan, self.page_break_before)


@dataclass
class BorderFillProps:
    """Border and fill properties for table cells."""

    border_type: str = "SOLID"
    border_width: str = "0.4mm"
    border_color: str = "000000"
    bg_color: Optional[str] = None  # None = transparent

    # Per-side overrides: (type, width, color) tuples; None = use defaults above
    left_border: Optional[Tuple[str, str, str]] = None
    right_border: Optional[Tuple[str, str, str]] = None
    top_border: Optional[Tuple[str, str, str]] = None
    bottom_border: Optional[Tuple[str, str, str]] = None

    # Raw fill XML for complex fills (e.g., gradation)
    fill_xml: Optional[str] = None

    def key(self) -> tuple:
        return (self.border_type, self.border_width, self.border_color, self.bg_color,
                self.left_border, self.right_border, self.top_border, self.bottom_border,
                self.fill_xml)


@dataclass
class Style:
    """A named style combining paragraph and character properties."""

    name: str
    char_props: CharProps = field(default_factory=CharProps)
    para_props: ParaProps = field(default_factory=ParaProps)
    next_style_name: Optional[str] = None


class StyleManager:
    """Manages named styles and assigns unique IDs for charPr/paraPr/faceName/borderFill."""

    def __init__(self) -> None:
        self._styles: Dict[str, Style] = {}
        self._face_name_map: Dict[str, int] = {}
        self._char_pr_map: Dict[tuple, int] = {}
        self._para_pr_map: Dict[tuple, int] = {}
        self._style_id_map: Dict[str, int] = {}
        self._border_fill_map: Dict[tuple, int] = {}

        self._face_names: List[str] = []
        self._char_prs: List[CharProps] = []
        self._para_prs: List[ParaProps] = []
        self._style_list: List[Style] = []
        self._border_fills: List[BorderFillProps] = []

        # Image binData tracking (populated during finalize)
        self._bin_data_images: list = []  # List[Image] with _bin_id assigned

        self.register_style(Style(name="바탕글"))

    def register_style(self, style: Style) -> None:
        self._styles[style.name] = style

    def get_style(self, name: str) -> Optional[Style]:
        return self._styles.get(name)

    def finalize(self, blocks: list) -> None:
        """Scan all blocks and assign sequential IDs to unique properties."""
        from .elements.paragraph import Paragraph, TextRun
        from .elements.table import Table
        from .elements.page_break import PageBreak
        from .elements.image import Image
        from .elements.diagram import Diagram
        from .elements.chart import Chart
        from .elements.svg_element import SvgElement
        from .elements.text_box import TextBox
        from .elements.doc_title import DocTitle
        from .elements.doc_purpose import DocPurpose
        from .elements.section_block import SectionBlock

        face_set: Dict[str, None] = {}
        char_set: Dict[tuple, CharProps] = {}
        para_set: Dict[tuple, ParaProps] = {}
        bf_set: Dict[tuple, BorderFillProps] = {}

        # Default borderFill (no border) — always ID 1
        default_bf = BorderFillProps(border_type="NONE", border_width="0.12mm",
                                     border_color="000000", bg_color=None)
        bf_set[default_bf.key()] = default_bf

        # From registered styles
        for style in self._styles.values():
            cp = style.char_props
            pp = style.para_props
            face_set[cp.font_name] = None
            char_set[cp.key()] = cp
            para_set[pp.key()] = pp

        # Scan Image and DocTitle blocks and assign bin IDs
        self._bin_data_images = []
        bin_id_counter = 0
        for block in blocks:
            if isinstance(block, Image):
                bin_id_counter += 1
                block._bin_id = bin_id_counter
                self._bin_data_images.append(block)
            elif isinstance(block, DocTitle):
                bin_id_counter += 1
                block._top_bin_id = bin_id_counter
                bin_id_counter += 1
                block._bottom_bin_id = bin_id_counter

        # Auto-number SectionBlock instances
        section_counter = 0
        for block in blocks:
            if isinstance(block, SectionBlock):
                section_counter += 1
                if block.num is None:
                    block.num = section_counter

        # From blocks
        for block in blocks:
            if isinstance(block, PageBreak):
                pb_para = Paragraph(page_break_before=True)
                pp = self._para_props_from_paragraph(pb_para)
                para_set[pp.key()] = pp
            elif isinstance(block, Paragraph):
                self._collect_paragraph_props(block, face_set, char_set, para_set)
            elif isinstance(block, Table):
                self._collect_table_props(block, face_set, char_set, para_set, bf_set)
            elif isinstance(block, Image):
                self._collect_image_props(block, face_set, char_set, para_set)
            elif isinstance(block, TextBox):
                self._collect_text_box_props(block, face_set, char_set, para_set)
            elif isinstance(block, DocTitle):
                self._collect_doc_title_props(block, face_set, char_set, para_set)
            elif isinstance(block, DocPurpose):
                self._collect_doc_purpose_props(block, face_set, char_set, para_set, bf_set)
            elif isinstance(block, SectionBlock):
                self._collect_section_block_props(block, face_set, char_set, para_set, bf_set)

        # Assign IDs
        self._face_names = list(face_set.keys())
        self._face_name_map = {name: i for i, name in enumerate(self._face_names)}

        self._char_prs = list(char_set.values())
        self._char_pr_map = {cp.key(): i for i, cp in enumerate(self._char_prs)}

        self._para_prs = list(para_set.values())
        # Assign IDs, skipping RESERVED_DOC_TITLE_PP_ID (always 10)
        self._para_pr_map = {}
        dynamic_id = 0
        for pp in self._para_prs:
            if dynamic_id == RESERVED_DOC_TITLE_PP_ID:
                dynamic_id += 1
            self._para_pr_map[pp.key()] = dynamic_id
            dynamic_id += 1

        self._style_list = list(self._styles.values())
        self._style_id_map = {s.name: i for i, s in enumerate(self._style_list)}

        # BorderFills: user-defined IDs start at 3 (header.xml hardcodes id=1 and id=2)
        self._border_fills = list(bf_set.values())
        self._border_fill_map = {bf.key(): i + 3 for i, bf in enumerate(self._border_fills)}

    def _collect_paragraph_props(self, para, face_set, char_set, para_set):
        pp = self._para_props_from_paragraph(para)
        para_set[pp.key()] = pp
        for run in para.runs:
            cp = self._char_props_from_run(run, para)
            face_set[cp.font_name] = None
            char_set[cp.key()] = cp

    def _collect_table_props(self, table, face_set, char_set, para_set, bf_set):
        from .elements.table import Table
        # Collect borderFills for table cells
        for row in table.rows:
            for cell in row.cells:
                bf = BorderFillProps(
                    border_type="SOLID",
                    border_width=table.border_width,
                    border_color=table.border_color,
                    bg_color=cell.bg_color,
                )
                bf_set[bf.key()] = bf
                # Collect from cell paragraphs
                for para in cell.paragraphs:
                    self._collect_paragraph_props(para, face_set, char_set, para_set)

    def _collect_text_box_props(self, text_box, face_set, char_set, para_set):
        """Collect font/paragraph props needed for TextBox text."""
        fn = text_box.font_name or "함초롬돋움"
        fs = text_box.font_size_pt or 10.0
        fc = text_box.font_color or "000000"
        cp = CharProps(
            font_name=fn,
            font_size_pt=fs,
            bold=text_box.font_bold,
            italic=text_box.font_italic,
            color=fc,
        )
        face_set[cp.font_name] = None
        char_set[cp.key()] = cp

    def _collect_doc_title_props(self, block, face_set, char_set, para_set):
        """Collect props needed for DocTitle (decoration images + title text).

        NOTE: ParaProps (CENTER, 120%) is NOT registered here.
        It is hardcoded as paraPr id=RESERVED_DOC_TITLE_PP_ID in styles_builder.
        """
        face_set["HY헤드라인M"] = None
        cp = CharProps(font_name="HY헤드라인M", font_size_pt=27.0, bold=True, color="000000")
        char_set[cp.key()] = cp

    def _collect_doc_purpose_props(self, block, face_set, char_set, para_set, bf_set):
        """Collect borderFills and font props for DocPurpose."""
        fn = block.font_name
        face_set[fn] = None
        cp = CharProps(font_name=fn, font_size_pt=13.0, bold=False, color="000000")
        char_set[cp.key()] = cp
        # Table outer border (DOUBLE 0.12mm all sides)
        bf_tbl = BorderFillProps(
            border_type="DOUBLE", border_width="0.12 mm", border_color="000000",
        )
        bf_set[bf_tbl.key()] = bf_tbl
        # Cell inner border (DOUBLE_SLIM top/bottom, NONE left/right, white fill)
        bf_cell = BorderFillProps(
            border_type="NONE", border_width="0.12 mm", border_color="000000",
            left_border=("NONE", "0.12 mm", "#000000"),
            right_border=("NONE", "0.12 mm", "#000000"),
            top_border=("DOUBLE_SLIM", "0.7 mm", "#000000"),
            bottom_border=("DOUBLE_SLIM", "0.7 mm", "#000000"),
            fill_xml='<hc:fillBrush><hc:winBrush faceColor="#FFFFFF" hatchColor="#000000" alpha="0"/></hc:fillBrush>',
        )
        bf_set[bf_cell.key()] = bf_cell

    def _collect_section_block_props(self, block, face_set, char_set, para_set, bf_set):
        """Collect borderFills, charPr, paraPr for SectionBlock (섹션.hwpx reference)."""
        face_set["휴먼명조"] = None
        face_set["HY헤드라인M"] = None
        face_set["고딕"] = None
        face_set["맑은 고딕"] = None

        # charPr id=8: Number cell (휴먼명조 18pt, white, bold)
        cp_num = CharProps(font_name="휴먼명조", font_size_pt=18.0, bold=True, color="FFFFFF")
        char_set[cp_num.key()] = cp_num
        # charPr id=9: Divider cell (고딕 17pt, white, bold)
        cp_div = CharProps(font_name="고딕", font_size_pt=17.0, bold=True, color="FFFFFF")
        char_set[cp_div.key()] = cp_div
        # charPr id=10: Title cell (HY헤드라인M 17pt, black, char_spacing=10)
        cp_title = CharProps(font_name="HY헤드라인M", font_size_pt=17.0, bold=False, color="000000", char_spacing=10)
        char_set[cp_title.key()] = cp_title
        # charPr id=11: Bottom bar (맑은 고딕 1.5pt, bold, char_spacing=15)
        cp_bar = CharProps(font_name="맑은 고딕", font_size_pt=1.5, bold=True, color="000000", char_spacing=15)
        char_set[cp_bar.key()] = cp_bar

        # paraPr: number cell — CENTER, lineSpacing=180
        pp_num = ParaProps(align="CENTER", line_spacing_value=180.0)
        para_set[pp_num.key()] = pp_num
        # paraPr: divider cell — JUSTIFY, lineSpacing=180
        pp_div = ParaProps(align="JUSTIFY", line_spacing_value=180.0)
        para_set[pp_div.key()] = pp_div
        # paraPr: title cell — LEFT, lineSpacing=150
        pp_title = ParaProps(align="LEFT", line_spacing_value=150.0)
        para_set[pp_title.key()] = pp_title
        # paraPr: bottom bar — LEFT, lineSpacing=180
        pp_bar = ParaProps(align="LEFT", line_spacing_value=180.0)
        para_set[pp_bar.key()] = pp_bar

        # borderFill: table outer — NONE all sides
        bf_tbl = BorderFillProps(
            border_type="NONE", border_width="0.1 mm", border_color="000000",
            left_border=("NONE", "0.1 mm", "none"),
            right_border=("NONE", "0.1 mm", "none"),
            top_border=("NONE", "0.1 mm", "none"),
            bottom_border=("NONE", "0.1 mm", "none"),
        )
        bf_set[bf_tbl.key()] = bf_tbl
        # borderFill: number cell — SOLID 0.4mm #999999 all sides, bg #1F5B9B
        bf_num = BorderFillProps(
            border_type="SOLID", border_width="0.4 mm", border_color="999999",
            bg_color="1F5B9B",
        )
        bf_set[bf_num.key()] = bf_num
        # borderFill: divider cell — SOLID left+right #999999, NONE top+bottom
        bf_div = BorderFillProps(
            border_type="NONE", border_width="0.1 mm", border_color="000000",
            left_border=("SOLID", "0.4 mm", "#999999"),
            right_border=("SOLID", "0.4 mm", "#999999"),
            top_border=("NONE", "0.1 mm", "none"),
            bottom_border=("NONE", "0.1 mm", "none"),
        )
        bf_set[bf_div.key()] = bf_div
        # borderFill: title cell — SOLID left #999999, NONE others, gradation #DFEAF5→#FFFFFF
        bf_title = BorderFillProps(
            border_type="NONE", border_width="0.1 mm", border_color="000000",
            left_border=("SOLID", "0.4 mm", "#999999"),
            right_border=("NONE", "0.1 mm", "none"),
            top_border=("NONE", "0.1 mm", "none"),
            bottom_border=("NONE", "0.1 mm", "none"),
            fill_xml=(
                '<hc:fillBrush>'
                '<hc:gradation type="LINEAR" angle="90" centerX="0" centerY="0"'
                ' step="250" colorNum="2" stepCenter="50" alpha="0">'
                '<hc:color value="#DFEAF5"/><hc:color value="#FFFFFF"/>'
                '</hc:gradation></hc:fillBrush>'
            ),
        )
        bf_set[bf_title.key()] = bf_title
        # borderFill: bottom bar — SOLID left #999999, NONE others, gradation #999999→#FFFFFF
        bf_bar = BorderFillProps(
            border_type="NONE", border_width="0.1 mm", border_color="000000",
            left_border=("SOLID", "0.4 mm", "#999999"),
            right_border=("NONE", "0.1 mm", "none"),
            top_border=("NONE", "0.1 mm", "none"),
            bottom_border=("NONE", "0.1 mm", "none"),
            fill_xml=(
                '<hc:fillBrush>'
                '<hc:gradation type="LINEAR" angle="90" centerX="0" centerY="0"'
                ' step="255" colorNum="2" stepCenter="50" alpha="0">'
                '<hc:color value="#999999"/><hc:color value="#FFFFFF"/>'
                '</hc:gradation></hc:fillBrush>'
            ),
        )
        bf_set[bf_bar.key()] = bf_bar

    def _collect_image_props(self, image, face_set, char_set, para_set):
        """Collect props needed for image block.

        In real HWPX structure (reverse-engineered from reference files),
        the image paragraph uses default paraPr (id=0). Caption is inside
        hp:pic > hp:caption, not a separate paragraph, so no extra
        keep_with_next or caption paragraph styles are needed.
        """
        pass

    # --- Lookup methods ---

    def get_face_id(self, font_name: str) -> int:
        return self._face_name_map.get(font_name, 0)

    def get_char_pr_id(self, cp: CharProps) -> int:
        return self._char_pr_map.get(cp.key(), 0)

    def get_para_pr_id(self, pp: ParaProps) -> int:
        # DocTitle paraPr (CENTER, 120%) → always reserved ID
        _reserved_key = ParaProps(align="CENTER", line_spacing_value=120.0).key()
        if pp.key() == _reserved_key:
            return RESERVED_DOC_TITLE_PP_ID
        return self._para_pr_map.get(pp.key(), 0)

    def get_style_id(self, name: str) -> int:
        return self._style_id_map.get(name, 0)

    def get_border_fill_id(self, bf: BorderFillProps) -> int:
        return self._border_fill_map.get(bf.key(), 1)

    def get_cell_border_fill_id(self, table, cell) -> int:
        """Get borderFill ID for a specific table cell."""
        bf = BorderFillProps(
            border_type="SOLID",
            border_width=table.border_width,
            border_color=table.border_color,
            bg_color=cell.bg_color,
        )
        return self.get_border_fill_id(bf)

    # --- Property resolution ---

    def _para_props_from_paragraph(self, p) -> ParaProps:
        return ParaProps(
            align=p.align,
            line_spacing_type=p.line_spacing_type,
            line_spacing_value=p.line_spacing_value,
            space_before_pt=p.space_before_pt,
            space_after_pt=p.space_after_pt,
            indent_left_mm=p.indent_left_mm,
            indent_right_mm=p.indent_right_mm,
            indent_first_mm=p.indent_first_mm,
            keep_with_next=p.keep_with_next,
            widow_orphan=p.widow_orphan,
            page_break_before=p.page_break_before,
        )

    def _char_props_from_run(self, run, para) -> CharProps:
        default_style = self._styles.get("바탕글", Style(name="바탕글"))
        dc = default_style.char_props
        return CharProps(
            font_name=run.font_name or para.font_name or dc.font_name,
            font_size_pt=run.font_size_pt or para.font_size_pt or dc.font_size_pt,
            bold=run.bold or para.bold,
            italic=run.italic or para.italic,
            underline=run.underline,
            strike=run.strike,
            color=run.color or para.color or dc.color,
        )

    def resolve_char_props(self, run, para) -> CharProps:
        return self._char_props_from_run(run, para)

    def resolve_para_props(self, para) -> ParaProps:
        return self._para_props_from_paragraph(para)
