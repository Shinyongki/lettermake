"""Generate Contents/section0.xml — the main body content.

Rewritten to match real HWPX file structure (reverse-engineered from samples):
- hp:secPr inside first hp:p (not hs:pageDef at section level)
- hp:linesegarray required on every hp:p
- hp:t instead of hp:char for text content
- No &#x000D; terminator runs
- hp:p attributes: id, pageBreak="0", columnBreak="0", merged="0"
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Any
from xml.sax.saxutils import escape

from ..utils import mm_to_hwp, pt_to_hwp

if TYPE_CHECKING:
    from ..document import HwpxDocument

# Namespace URIs
NS_HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
NS_HC = "http://www.hancom.co.kr/hwpml/2011/core"
NS_HS = "http://www.hancom.co.kr/hwpml/2011/section"
NS_HH = "http://www.hancom.co.kr/hwpml/2011/head"

# Global paragraph ID counter
_para_id = 0


def _next_para_id() -> int:
    global _para_id
    _id = _para_id
    _para_id += 1
    return _id


def _lineseg_xml(
    font_size_pt: float = 10.0,
    line_spacing_pct: float = 160.0,
    content_width_hwp: int = 42520,
) -> str:
    """Build hp:linesegarray with a single hp:lineseg."""
    height = pt_to_hwp(font_size_pt)         # e.g. 1000 for 10pt
    baseline = round(height * 0.85)           # e.g. 850
    spacing = round(height * (line_spacing_pct - 100) / 100)  # e.g. 600
    return (
        f'<hp:linesegarray>'
        f'<hp:lineseg textpos="0" vertpos="0" vertsize="{height}" textheight="{height}"'
        f' baseline="{baseline}" spacing="{spacing}" horzpos="0"'
        f' horzsize="{content_width_hwp}" flags="393216"/>'
        f'</hp:linesegarray>'
    )


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def build_section_xml(doc: HwpxDocument) -> bytes:
    """Build Contents/section0.xml."""
    global _para_id
    _para_id = 0

    ps = doc.page_settings
    sm = doc.style_manager

    if ps.orientation == "landscape":
        w, h = ps.height_mm, ps.width_mm
    else:
        w, h = ps.width_mm, ps.height_mm

    content_w = mm_to_hwp(ps.content_width_mm)

    lines: List[str] = []
    lines.append('<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>')
    lines.append(
        '<hs:sec'
        ' xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"'
        f' xmlns:hp="{NS_HP}"'
        ' xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph"'
        f' xmlns:hs="{NS_HS}"'
        f' xmlns:hc="{NS_HC}"'
        f' xmlns:hh="{NS_HH}"'
        ' xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history"'
        ' xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page"'
        ' xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"'
        ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
        ' xmlns:opf="http://www.idpf.org/2007/opf/"'
        ' xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart"'
        ' xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"'
        ' xmlns:epub="http://www.idpf.org/2007/ops"'
        ' xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">'
    )

    # Detect DocTitle as first block (secPr will be merged into it)
    from ..elements.doc_title import DocTitle
    from ..elements.doc_purpose import DocPurpose
    from ..elements.section_block import SectionBlock

    first_is_doc_title = bool(doc.blocks) and isinstance(doc.blocks[0], DocTitle)

    # Section header paragraph — only if first block is NOT DocTitle
    if not first_is_doc_title:
        lines.append(_section_header_para(doc))

    # Content blocks
    if not doc.blocks:
        lines.append(_empty_paragraph(content_w))
    else:
        skip_next = False
        for i, block in enumerate(doc.blocks):
            if skip_next:
                skip_next = False
                continue

            if isinstance(block, DocTitle):
                # Look for DocPurpose immediately after
                purpose = None
                if i + 1 < len(doc.blocks) and isinstance(doc.blocks[i + 1], DocPurpose):
                    purpose = doc.blocks[i + 1]
                    skip_next = True
                include_secpr = (i == 0 and first_is_doc_title)
                lines.append(_render_doc_title_v7(
                    block, purpose, doc if include_secpr else None, sm, content_w))
            elif isinstance(block, SectionBlock) and i > 0:
                lines.append(_empty_paragraph(content_w))
                lines.append(_render_block(block, sm, content_w))
            else:
                lines.append(_render_block(block, sm, content_w))

    lines.append('</hs:sec>')
    return "\n".join(lines).encode("utf-8")


# ---------------------------------------------------------------------------
# Section header paragraph (secPr + colPr)
# ---------------------------------------------------------------------------

def _secpr_run_xml(doc: "HwpxDocument") -> str:
    """Generate the hp:run containing secPr + colPr (reusable)."""
    ps = doc.page_settings

    if ps.orientation == "landscape":
        w_hwp = mm_to_hwp(ps.height_mm)
        h_hwp = mm_to_hwp(ps.width_mm)
        landscape = "NARROWLY"
    else:
        w_hwp = mm_to_hwp(ps.width_mm)
        h_hwp = mm_to_hwp(ps.height_mm)
        landscape = "WIDELY"

    ml = mm_to_hwp(ps.margin_left)
    mr = mm_to_hwp(ps.margin_right)
    mt = mm_to_hwp(ps.margin_top)
    mb = mm_to_hwp(ps.margin_bottom)
    mh = mm_to_hwp(ps.header_mm) if ps.header_mm else mm_to_hwp(15.0)
    mf = mm_to_hwp(ps.footer_mm) if ps.footer_mm else mm_to_hwp(15.0)

    return (
        f'<hp:run charPrIDRef="0">'
        f'<hp:secPr textDirection="HORIZONTAL" spaceColumns="1134"'
        f' tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT"'
        f' outlineShapeIDRef="1" memoShapeIDRef="0"'
        f' textVerticalWidthHead="0" masterPageCnt="0">'
        f'<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>'
        f'<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
        f'<hp:visibility hideFirstHeader="0" hideFirstFooter="0"'
        f' hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL"'
        f' hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>'
        f'<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>'
        f'<hp:pagePr landscape="{landscape}" width="{w_hwp}" height="{h_hwp}"'
        f' gutterType="LEFT_ONLY">'
        f'<hp:margin header="{mh}" footer="{mf}" gutter="0"'
        f' left="{ml}" right="{mr}" top="{mt}" bottom="{mb}"/>'
        f'</hp:pagePr>'
        f'<hp:footNotePr>'
        f'<hp:autoNumFormat type="DIGIT" suffixChar=")" superscript="0"/>'
        f'<hp:noteLine length="-1" type="SOLID" width="0.12mm" color="#000000"/>'
        f'<hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/>'
        f'</hp:footNotePr>'
        f'<hp:endNotePr>'
        f'<hp:autoNumFormat type="DIGIT" suffixChar=")" superscript="0"/>'
        f'<hp:noteLine length="-1" type="SOLID" width="0.12mm" color="#000000"/>'
        f'<hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/>'
        f'</hp:endNotePr>'
        f'<hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER"'
        f' headerInside="0" footerInside="0" fillArea="PAPER"/>'
        f'<hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER"'
        f' headerInside="0" footerInside="0" fillArea="PAPER"/>'
        f'<hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER"'
        f' headerInside="0" footerInside="0" fillArea="PAPER"/>'
        f'</hp:secPr>'
        f'<hp:ctrl>'
        f'<hp:colPr id="" type="NEWSPAPER" layout="LEFT"'
        f' colCount="1" sameSz="1" sameGap="0"/>'
        f'</hp:ctrl>'
        f'</hp:run>'
    )


def _section_header_para(doc: HwpxDocument) -> str:
    """First hp:p containing hp:secPr for page settings."""
    pid = _next_para_id()
    return (
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'{_secpr_run_xml(doc)}'
        f'<hp:linesegarray>'
        f'<hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000"'
        f' baseline="850" spacing="600" horzpos="0" horzsize="42520" flags="393216"/>'
        f'</hp:linesegarray>'
        f'</hp:p>'
    )


# ---------------------------------------------------------------------------
# Empty / page-break paragraphs
# ---------------------------------------------------------------------------

def _empty_paragraph(content_w: int) -> str:
    pid = _next_para_id()
    lseg = _lineseg_xml(10.0, 160.0, content_w)
    return (
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="0"/>'
        f'{lseg}'
        f'</hp:p>'
    )


def _render_page_break(sm, content_w: int) -> str:
    from ..elements.paragraph import Paragraph
    para = Paragraph(page_break_before=True)
    pp = sm.resolve_para_props(para)
    pp_id = sm.get_para_pr_id(pp)
    pid = _next_para_id()
    lseg = _lineseg_xml(10.0, 160.0, content_w)
    return (
        f'<hp:p id="{pid}" paraPrIDRef="{pp_id}" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="0"/>'
        f'{lseg}'
        f'</hp:p>'
    )


# ---------------------------------------------------------------------------
# Block dispatcher
# ---------------------------------------------------------------------------

def _render_block(block: Any, sm, content_w: int) -> str:
    from ..elements.paragraph import Paragraph
    from ..elements.table import Table
    from ..elements.page_break import PageBreak
    from ..elements.image import Image
    from ..elements.diagram import Diagram
    from ..elements.chart import Chart
    from ..elements.svg_element import SvgElement
    from ..elements.text_box import TextBox
    from ..elements.doc_title import DocTitle
    from ..elements.doc_purpose import DocPurpose
    from ..elements.section_block import SectionBlock

    if isinstance(block, PageBreak):
        return _render_page_break(sm, content_w)
    elif isinstance(block, DocTitle):
        return _render_doc_title_v7(block, None, None, sm, content_w)
    elif isinstance(block, DocPurpose):
        return _render_doc_purpose(block, sm, content_w)
    elif isinstance(block, SectionBlock):
        return _render_section_block(block, sm, content_w)
    elif isinstance(block, Paragraph):
        return _render_paragraph(block, sm, content_w)
    elif isinstance(block, Table):
        return _render_table(block, sm, content_w)
    elif isinstance(block, Image):
        return _render_image(block, sm, content_w)
    elif isinstance(block, Diagram):
        return _render_diagram(block, content_w)
    elif isinstance(block, Chart):
        return _render_chart(block, content_w)
    elif isinstance(block, SvgElement):
        return _render_svg(block, content_w)
    elif isinstance(block, TextBox):
        return _render_text_box(block, sm, content_w)

    return _empty_paragraph(content_w)


# ---------------------------------------------------------------------------
# Paragraph rendering
# ---------------------------------------------------------------------------

def _render_paragraph(para, sm, content_w: int) -> str:
    pp = sm.resolve_para_props(para)
    pp_id = sm.get_para_pr_id(pp)
    style_id = 0
    if para.style_name:
        style_id = sm.get_style_id(para.style_name)

    pid = _next_para_id()
    font_size = para.font_size_pt or 10.0
    line_sp = para.line_spacing_value or 160.0

    parts: List[str] = []
    parts.append(
        f'<hp:p id="{pid}" paraPrIDRef="{pp_id}" styleIDRef="{style_id}"'
        f' pageBreak="0" columnBreak="0" merged="0">'
    )

    if not para.runs:
        parts.append('<hp:run charPrIDRef="0"/>')
    else:
        for run in para.runs:
            cp = sm.resolve_char_props(run, para)
            cp_id = sm.get_char_pr_id(cp)
            text = escape(run.text)
            parts.append(f'<hp:run charPrIDRef="{cp_id}"><hp:t>{text}</hp:t></hp:run>')

    parts.append(_lineseg_xml(font_size, line_sp, content_w))
    parts.append('</hp:p>')
    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Table rendering
# ---------------------------------------------------------------------------

def _render_table(table, sm, content_w: int) -> str:
    """Render a table matching real Hancom HWPX structure (reverse-engineered from sample)."""
    from ..elements.table import Table, TableRow, TableCell
    from ..styles import BorderFillProps

    # --- Column widths ---
    num_cols = _detect_col_count(table)
    out_margin = 283  # matches sample
    if table.col_widths_mm and len(table.col_widths_mm) == num_cols:
        col_widths_hwp = [mm_to_hwp(w) for w in table.col_widths_mm]
    else:
        # Auto: equal widths filling content area
        table_w = content_w - out_margin * 2
        col_w = table_w // num_cols
        col_widths_hwp = [col_w] * num_cols
        col_widths_hwp[-1] = table_w - col_w * (num_cols - 1)

    total_width = sum(col_widths_hwp)
    cm_lr = mm_to_hwp(table.cell_margin_lr_mm)
    cm_tb = mm_to_hwp(table.cell_margin_tb_mm)

    merge_map = _build_merge_map(table)

    # Row heights: text height + cell margins
    font_height = pt_to_hwp(table.font_size_pt or 10.0)
    row_heights = []
    for row in table.rows:
        if row.height_mm:
            rh = mm_to_hwp(row.height_mm)
        else:
            rh = font_height + cm_tb * 2
        row_heights.append(rh)
    total_height = sum(row_heights)

    # Table border fill
    tbl_bf = BorderFillProps(
        border_type="SOLID", border_width=table.border_width,
        border_color=table.border_color, bg_color=None,
    )
    tbl_bf_id = sm.get_border_fill_id(tbl_bf)

    pid = _next_para_id()
    tbl_id = pid + 1000

    parts: List[str] = []
    # Wrapper paragraph
    parts.append(
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
    )
    parts.append('<hp:run charPrIDRef="0">')

    # hp:tbl — all attributes matching sample
    parts.append(
        f'<hp:tbl id="{tbl_id}" zOrder="0" numberingType="TABLE"'
        f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
        f' dropcapstyle="None" pageBreak="{"NONE" if table.keep_together else "CELL"}" repeatHeader="1"'
        f' rowCnt="{len(table.rows)}" colCnt="{num_cols}"'
        f' cellSpacing="0" borderFillIDRef="{tbl_bf_id}" noAdjust="0">'
    )
    parts.append(
        f'<hp:sz width="{total_width}" widthRelTo="ABSOLUTE"'
        f' height="{total_height}" heightRelTo="ABSOLUTE" protect="0"/>'
    )
    parts.append(
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1"'
        f' allowOverlap="0" holdAnchorAndSO="0"'
        f' vertRelTo="PARA" horzRelTo="COLUMN"'
        f' vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
    )
    parts.append(
        f'<hp:outMargin left="{out_margin}" right="{out_margin}"'
        f' top="{out_margin}" bottom="{out_margin}"/>'
    )
    parts.append(
        f'<hp:inMargin left="{cm_lr}" right="{cm_lr}"'
        f' top="{cm_tb}" bottom="{cm_tb}"/>'
    )

    # Rows (NO cellZone/gridCol — sample doesn't have it)
    for row_idx, row in enumerate(table.rows):
        parts.append('<hp:tr>')
        col_offset = 0
        for cell_idx, cell in enumerate(row.cells):
            while (row_idx, col_offset) in merge_map:
                col_offset += 1

            bf_id = sm.get_cell_border_fill_id(table, cell)
            cell_width = sum(
                col_widths_hwp[col_offset + j]
                for j in range(cell.colspan)
                if col_offset + j < len(col_widths_hwp)
            )
            cell_height = cm_tb * 2  # minimum = top+bottom margin

            cell_font_size = table.font_size_pt or 10.0
            cell_content_w = cell_width - cm_lr * 2
            if cell_content_w < 100:
                cell_content_w = cell_width

            # hp:tc with all attributes matching sample
            parts.append(
                f'<hp:tc name="" header="0" hasMargin="1" protect="0"'
                f' editable="0" dirty="0" borderFillIDRef="{bf_id}">'
            )

            # hp:subList wrapping cell content (required by Hancom)
            valign = cell.valign or "CENTER"
            parts.append(
                f'<hp:subList id="" textDirection="HORIZONTAL"'
                f' lineWrap="BREAK" vertAlign="{valign}"'
                f' linkListIDRef="0" linkListNextIDRef="0"'
                f' textWidth="0" textHeight="0"'
                f' hasTextRef="0" hasNumRef="0">'
            )

            if cell.paragraphs:
                for para in cell.paragraphs:
                    parts.append(_render_cell_paragraph(para, sm, cell_content_w, cell_font_size))
            else:
                parts.append(_cell_empty_paragraph(cell_content_w, cell_font_size))

            parts.append('</hp:subList>')

            # Cell metadata AFTER subList (sample order)
            parts.append(f'<hp:cellAddr colAddr="{col_offset}" rowAddr="{row_idx}"/>')
            parts.append(f'<hp:cellSpan colSpan="{cell.colspan}" rowSpan="{cell.rowspan}"/>')
            parts.append(f'<hp:cellSz width="{cell_width}" height="{cell_height}"/>')
            parts.append(
                f'<hp:cellMargin left="{cm_lr}" right="{cm_lr}"'
                f' top="{cm_tb}" bottom="{cm_tb}"/>'
            )
            parts.append('</hp:tc>')
            col_offset += cell.colspan

        parts.append('</hp:tr>')

    parts.append('</hp:tbl>')
    parts.append('<hp:t/>')  # empty text after table (matches sample)
    parts.append('</hp:run>')

    # lineseg for wrapper paragraph (horzsize=0 for table paras in sample)
    parts.append(
        '<hp:linesegarray>'
        '<hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000"'
        ' baseline="850" spacing="600" horzpos="0" horzsize="0" flags="393216"/>'
        '</hp:linesegarray>'
    )
    parts.append('</hp:p>')
    return "\n".join(parts)


def _detect_col_count(table) -> int:
    """Detect column count from widths or first row."""
    if table.col_widths_mm:
        return len(table.col_widths_mm)
    if table.rows:
        return sum(c.colspan for c in table.rows[0].cells)
    return 1


def _render_cell_paragraph(para, sm, cell_w: int, font_size: float) -> str:
    pp = sm.resolve_para_props(para)
    pp_id = sm.get_para_pr_id(pp)
    pid = _next_para_id()

    parts: List[str] = []
    parts.append(
        f'<hp:p id="{pid}" paraPrIDRef="{pp_id}" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
    )

    if not para.runs:
        parts.append('<hp:run charPrIDRef="0"/>')
    else:
        for run in para.runs:
            cp = sm.resolve_char_props(run, para)
            cp_id = sm.get_char_pr_id(cp)
            text = escape(run.text)
            parts.append(f'<hp:run charPrIDRef="{cp_id}"><hp:t>{text}</hp:t></hp:run>')

    parts.append(_lineseg_xml(font_size, 160.0, cell_w))
    parts.append('</hp:p>')
    return "\n".join(parts)


def _cell_empty_paragraph(cell_w: int, font_size: float) -> str:
    pid = _next_para_id()
    lseg = _lineseg_xml(font_size, 160.0, cell_w)
    return (
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="0"/>'
        f'{lseg}'
        f'</hp:p>'
    )


def _build_merge_map(table) -> set:
    covered = set()
    for row_idx, row in enumerate(table.rows):
        col_offset = 0
        for cell in row.cells:
            while (row_idx, col_offset) in covered:
                col_offset += 1
            if cell.rowspan > 1 or cell.colspan > 1:
                for dr in range(cell.rowspan):
                    for dc in range(cell.colspan):
                        if dr == 0 and dc == 0:
                            continue
                        covered.add((row_idx + dr, col_offset + dc))
            col_offset += cell.colspan
    return covered


# ---------------------------------------------------------------------------
# DocTitle rendering (decoration images + title text)
# ---------------------------------------------------------------------------

def _render_decoration_pic(bin_id: int, pixel_w: int, pixel_h: int,
                           content_w: int, vert_offset: int = 0) -> str:
    """Render an hp:pic for a decoration line image (matching 도형.hwpx)."""
    # Display size matches reference: width=47907 (content width), height=708
    w_hwp = content_w  # fill content area width
    h_hwp = 708  # thin decoration line
    dim_w = pixel_w * 75
    dim_h = pixel_h * 75
    cx = w_hwp // 2
    cy = h_hwp // 2
    pic_id = _next_para_id() + 2000000
    instid = pic_id + 1000

    return (
        f'<hp:pic id="{pic_id}" zOrder="0" numberingType="PICTURE"'
        f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
        f' dropcapstyle="None" href=";0;0;0;" groupLevel="0" instid="{instid}" reverse="0">'
        f'<hp:offset x="0" y="0"/>'
        f'<hp:orgSz width="{w_hwp}" height="{h_hwp}"/>'
        f'<hp:curSz width="0" height="0"/>'
        f'<hp:flip horizontal="0" vertical="0"/>'
        f'<hp:rotationInfo angle="0" centerX="{cx}" centerY="{cy}" rotateimage="1"/>'
        f'<hp:renderingInfo>'
        f'<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        f'<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        f'<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        f'</hp:renderingInfo>'
        f'<hp:imgRect>'
        f'<hc:pt0 x="0" y="0"/><hc:pt1 x="{w_hwp}" y="0"/>'
        f'<hc:pt2 x="{w_hwp}" y="{h_hwp}"/><hc:pt3 x="0" y="{h_hwp}"/>'
        f'</hp:imgRect>'
        f'<hp:imgClip left="0" right="{dim_w}" top="0" bottom="{dim_h}"/>'
        f'<hp:inMargin left="0" right="0" top="0" bottom="0"/>'
        f'<hp:imgDim dimwidth="{dim_w}" dimheight="{dim_h}"/>'
        f'<hc:img binaryItemIDRef="image{bin_id}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>'
        f'<hp:effects/>'
        f'<hp:sz width="{w_hwp}" widthRelTo="ABSOLUTE" height="{h_hwp}" heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0"'
        f' holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN"'
        f' vertAlign="TOP" horzAlign="LEFT" vertOffset="{vert_offset}" horzOffset="0"/>'
        f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
        f'</hp:pic>'
    )


def _render_doc_title_v7(block, purpose_block, doc, sm, content_w: int) -> str:
    """Render DocTitle v7: single hp:p with secPr + floating elements + title text.

    Structure (from test_notice_v7.hwpx reference):
        run[0]: secPr + colPr (if doc is provided)
        run[1]: hp:pic(image1, vertOffset=6968) + hp:tbl(purpose, vertOffset=6976)
                + hp:pic(image2, vertOffset=0)
        run[2]: title text (HY헤드라인M 27pt)
        lineseg: vertsize=1300, spacing=780
        paraPrIDRef → CENTER, 120% line spacing
    """
    from ..styles import CharProps, RESERVED_DOC_TITLE_PP_ID
    parts: List[str] = []

    pid = _next_para_id()

    # charPr: HY헤드라인M 27pt bold
    cp_title = CharProps(font_name="HY헤드라인M", font_size_pt=27.0, bold=True, color="000000")
    cp_title_id = sm.get_char_pr_id(cp_title)

    # paraPrIDRef → 고정 ID (CENTER, 120% line spacing)
    parts.append(
        f'<hp:p id="{pid}" paraPrIDRef="{RESERVED_DOC_TITLE_PP_ID}" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
    )

    # --- run[0]: secPr + colPr ---
    if doc is not None:
        parts.append(_secpr_run_xml(doc))

    # --- run[1]: floating images + purpose table ---
    parts.append('<hp:run charPrIDRef="0">')

    # image1 (위쪽 선도형 — floats below title at vertOffset=6968)
    parts.append(_render_decoration_pic(
        block._top_bin_id, block._top_pixel_w, block._top_pixel_h,
        content_w, vert_offset=6968))

    # purpose table (vertOffset=6976, right below image1)
    if purpose_block is not None:
        parts.append(_render_purpose_table_float(purpose_block, sm, content_w, 6976))

    # image2 (아래쪽 선도형 — 제목 위 배치, vertOffset=-708: 선도형 높이만큼 위로)
    parts.append(_render_decoration_pic(
        block._bottom_bin_id, block._bottom_pixel_w, block._bottom_pixel_h,
        content_w, vert_offset=-708))

    parts.append('<hp:t/>')
    parts.append('</hp:run>')

    # --- run[2]: title text ---
    title_text = escape(block.text)
    parts.append(
        f'<hp:run charPrIDRef="{cp_title_id}"><hp:t>{title_text}</hp:t></hp:run>'
    )

    # lineseg: vertsize=1300, spacing=260 (120% line spacing: 1300 * 0.20)
    parts.append(
        '<hp:linesegarray>'
        '<hp:lineseg textpos="0" vertpos="0" vertsize="1300" textheight="1300"'
        f' baseline="1105" spacing="260" horzpos="0" horzsize="{content_w}" flags="393216"/>'
        '</hp:linesegarray>'
    )

    parts.append('</hp:p>')
    return "\n".join(parts)


def _render_purpose_table_float(block, sm, content_w: int, vert_offset: int) -> str:
    """Render DocPurpose as a floating hp:tbl (inside a run, no wrapper paragraph).

    Used by _render_doc_title_v7 to embed purpose table in the title paragraph.
    """
    from ..styles import BorderFillProps, CharProps

    tbl_width = content_w - 283 * 2  # 47623
    tbl_height = 4228
    out_margin = 283
    in_margin_lr = 510
    in_margin_tb = 141

    # BorderFills — 목적 표 외곽: DOUBLE 이중선
    bf_tbl = BorderFillProps(
        border_type="DOUBLE", border_width="0.12 mm", border_color="000000",
    )
    bf_tbl_id = sm.get_border_fill_id(bf_tbl)

    bf_cell = BorderFillProps(
        border_type="NONE", border_width="0.12 mm", border_color="000000",
        left_border=("NONE", "0.12 mm", "#000000"),
        right_border=("NONE", "0.12 mm", "#000000"),
        top_border=("DOUBLE_SLIM", "0.7 mm", "#000000"),
        bottom_border=("DOUBLE_SLIM", "0.7 mm", "#000000"),
        fill_xml='<hc:fillBrush><hc:winBrush faceColor="#FFFFFF" hatchColor="#000000" alpha="0"/></hc:fillBrush>',
    )
    bf_cell_id = sm.get_border_fill_id(bf_cell)

    # Font from block (preset.font_table)
    fn = block.font_name
    cp = CharProps(font_name=fn, font_size_pt=13.0, bold=False, color="000000")
    cp_id = sm.get_char_pr_id(cp)

    tbl_id = _next_para_id() + 1000
    cell_content_w = tbl_width - in_margin_lr * 2
    cell_text = escape(block.text)
    inner_pid = _next_para_id()
    inner_lseg = _lineseg_xml(13.0, 160.0, cell_content_w)

    parts: List[str] = []
    parts.append(
        f'<hp:tbl id="{tbl_id}" zOrder="0" numberingType="TABLE"'
        f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
        f' dropcapstyle="None" pageBreak="CELL" repeatHeader="1"'
        f' rowCnt="1" colCnt="1" cellSpacing="0"'
        f' borderFillIDRef="{bf_tbl_id}" noAdjust="0">'
    )
    parts.append(
        f'<hp:sz width="{tbl_width}" widthRelTo="ABSOLUTE"'
        f' height="{tbl_height}" heightRelTo="ABSOLUTE" protect="0"/>'
    )
    parts.append(
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1"'
        f' allowOverlap="0" holdAnchorAndSO="0"'
        f' vertRelTo="PARA" horzRelTo="COLUMN"'
        f' vertAlign="TOP" horzAlign="LEFT"'
        f' vertOffset="{vert_offset}" horzOffset="0"/>'
    )
    parts.append(
        f'<hp:outMargin left="{out_margin}" right="{out_margin}"'
        f' top="{out_margin}" bottom="{out_margin}"/>'
    )
    parts.append(
        f'<hp:inMargin left="{in_margin_lr}" right="{in_margin_lr}"'
        f' top="{in_margin_tb}" bottom="{in_margin_tb}"/>'
    )
    parts.append('<hp:tr>')
    parts.append(
        f'<hp:tc name="" header="0" hasMargin="1" protect="0"'
        f' editable="0" dirty="0" borderFillIDRef="{bf_cell_id}">'
    )
    parts.append(
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK"'
        f' vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0"'
        f' textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
    )
    parts.append(
        f'<hp:p id="{inner_pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{cp_id}"><hp:t>{cell_text}</hp:t></hp:run>'
        f'{inner_lseg}</hp:p>'
    )
    parts.append('</hp:subList>')
    parts.append('<hp:cellAddr colAddr="0" rowAddr="0"/>')
    parts.append('<hp:cellSpan colSpan="1" rowSpan="1"/>')
    parts.append(f'<hp:cellSz width="{tbl_width}" height="{tbl_height}"/>')
    parts.append(
        f'<hp:cellMargin left="{in_margin_lr}" right="{in_margin_lr}"'
        f' top="{in_margin_tb}" bottom="{in_margin_tb}"/>'
    )
    parts.append('</hp:tc>')
    parts.append('</hp:tr>')
    parts.append('</hp:tbl>')
    return "\n".join(parts)


# ---------------------------------------------------------------------------
# DocPurpose rendering (1x1 table with text)
# ---------------------------------------------------------------------------

def _render_doc_purpose(block, sm, content_w: int) -> str:
    """Render DocPurpose: 1x1 table matching 도형.hwpx structure.

    Table: borderFillIDRef for outer (SOLID 0.12mm)
    Cell:  borderFillIDRef for inner (DOUBLE_SLIM top/bottom, white fill)
    """
    from ..styles import BorderFillProps, CharProps

    # Table dimensions (from reference)
    tbl_width = content_w - 283 * 2  # 47624 in reference (content_w minus outMargin)
    tbl_height = 4228
    out_margin = 283
    in_margin_lr = 510
    in_margin_tb = 141

    # BorderFill IDs
    bf_tbl = BorderFillProps(
        border_type="SOLID", border_width="0.12 mm", border_color="000000",
    )
    bf_tbl_id = sm.get_border_fill_id(bf_tbl)

    bf_cell = BorderFillProps(
        border_type="NONE", border_width="0.12 mm", border_color="000000",
        left_border=("NONE", "0.12 mm", "#000000"),
        right_border=("NONE", "0.12 mm", "#000000"),
        top_border=("DOUBLE_SLIM", "0.7 mm", "#000000"),
        bottom_border=("DOUBLE_SLIM", "0.7 mm", "#000000"),
        fill_xml='<hc:fillBrush><hc:winBrush faceColor="#FFFFFF" hatchColor="#000000" alpha="0"/></hc:fillBrush>',
    )
    bf_cell_id = sm.get_border_fill_id(bf_cell)

    # Font from block (preset.font_table)
    fn = block.font_name
    cp = CharProps(font_name=fn, font_size_pt=13.0, bold=False, color="000000")
    cp_id = sm.get_char_pr_id(cp)

    pid = _next_para_id()
    tbl_id = pid + 1000
    cell_content_w = tbl_width - in_margin_lr * 2
    cell_text = escape(block.text)

    inner_pid = _next_para_id()
    inner_lseg = _lineseg_xml(13.0, 160.0, cell_content_w)

    parts: List[str] = []
    # Wrapper paragraph
    parts.append(
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
    )
    parts.append('<hp:run charPrIDRef="0">')

    # hp:tbl (1x1)
    parts.append(
        f'<hp:tbl id="{tbl_id}" zOrder="0" numberingType="TABLE"'
        f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
        f' dropcapstyle="None" pageBreak="CELL" repeatHeader="1"'
        f' rowCnt="1" colCnt="1" cellSpacing="0"'
        f' borderFillIDRef="{bf_tbl_id}" noAdjust="0">'
    )
    parts.append(
        f'<hp:sz width="{tbl_width}" widthRelTo="ABSOLUTE"'
        f' height="{tbl_height}" heightRelTo="ABSOLUTE" protect="0"/>'
    )
    parts.append(
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1"'
        f' allowOverlap="0" holdAnchorAndSO="0"'
        f' vertRelTo="PARA" horzRelTo="COLUMN"'
        f' vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
    )
    parts.append(
        f'<hp:outMargin left="{out_margin}" right="{out_margin}"'
        f' top="{out_margin}" bottom="{out_margin}"/>'
    )
    parts.append(
        f'<hp:inMargin left="{in_margin_lr}" right="{in_margin_lr}"'
        f' top="{in_margin_tb}" bottom="{in_margin_tb}"/>'
    )

    # Single row, single cell
    parts.append('<hp:tr>')
    parts.append(
        f'<hp:tc name="" header="0" hasMargin="1" protect="0"'
        f' editable="0" dirty="0" borderFillIDRef="{bf_cell_id}">'
    )
    parts.append(
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK"'
        f' vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0"'
        f' textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
    )
    parts.append(
        f'<hp:p id="{inner_pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{cp_id}"><hp:t>{cell_text}</hp:t></hp:run>'
        f'{inner_lseg}</hp:p>'
    )
    parts.append('</hp:subList>')
    parts.append(f'<hp:cellAddr colAddr="0" rowAddr="0"/>')
    parts.append(f'<hp:cellSpan colSpan="1" rowSpan="1"/>')
    parts.append(f'<hp:cellSz width="{tbl_width}" height="{tbl_height}"/>')
    parts.append(
        f'<hp:cellMargin left="{in_margin_lr}" right="{in_margin_lr}"'
        f' top="{in_margin_tb}" bottom="{in_margin_tb}"/>'
    )
    parts.append('</hp:tc>')
    parts.append('</hp:tr>')
    parts.append('</hp:tbl>')
    parts.append('<hp:t/>')
    parts.append('</hp:run>')

    # lineseg for wrapper
    parts.append(
        '<hp:linesegarray>'
        '<hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000"'
        ' baseline="850" spacing="600" horzpos="0" horzsize="0" flags="393216"/>'
        '</hp:linesegarray>'
    )
    parts.append('</hp:p>')
    return "\n".join(parts)


# ---------------------------------------------------------------------------
# SectionBlock rendering (2x3 table with number + divider + title)
# ---------------------------------------------------------------------------

def _render_section_block(block, sm, content_w: int) -> str:
    """Render SectionBlock: 2-row 3-col table matching 섹션.hwpx reference exactly.

    Col 0: Number cell (rowSpan=2, blue bg #1F5B9B, white bold 휴먼명조 18pt)
    Col 1: Divider cell (rowSpan=2, gray side borders only)
    Col 2 Row 0: Title text (gradation #DFEAF5→#FFFFFF, HY헤드라인M 17pt)
    Col 2 Row 1: Bottom bar (height=150, gradation #999999→#FFFFFF)
    """
    from ..styles import BorderFillProps, CharProps, ParaProps

    # Dimensions (dynamic width)
    total_width = content_w
    total_height = 2715
    col0_w = 2850   # number
    col1_w = 565    # divider
    col2_w = total_width - col0_w - col1_w  # title (fills remaining)
    row0_h = 2565   # title row
    row1_h = 150    # bottom bar

    # --- BorderFill IDs (must match _collect_section_block_props exactly) ---
    bf_tbl = BorderFillProps(
        border_type="NONE", border_width="0.1 mm", border_color="000000",
        left_border=("NONE", "0.1 mm", "none"),
        right_border=("NONE", "0.1 mm", "none"),
        top_border=("NONE", "0.1 mm", "none"),
        bottom_border=("NONE", "0.1 mm", "none"),
    )
    bf_tbl_id = sm.get_border_fill_id(bf_tbl)

    bf_num = BorderFillProps(
        border_type="SOLID", border_width="0.4 mm", border_color="999999",
        bg_color="1F5B9B",
    )
    bf_num_id = sm.get_border_fill_id(bf_num)

    bf_div = BorderFillProps(
        border_type="NONE", border_width="0.1 mm", border_color="000000",
        left_border=("SOLID", "0.4 mm", "#999999"),
        right_border=("SOLID", "0.4 mm", "#999999"),
        top_border=("NONE", "0.1 mm", "none"),
        bottom_border=("NONE", "0.1 mm", "none"),
    )
    bf_div_id = sm.get_border_fill_id(bf_div)

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
    bf_title_id = sm.get_border_fill_id(bf_title)

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
    bf_bar_id = sm.get_border_fill_id(bf_bar)

    # --- CharPr IDs ---
    cp_num = CharProps(font_name="휴먼명조", font_size_pt=18.0, bold=True, color="FFFFFF")
    cp_num_id = sm.get_char_pr_id(cp_num)
    cp_div = CharProps(font_name="고딕", font_size_pt=17.0, bold=True, color="FFFFFF")
    cp_div_id = sm.get_char_pr_id(cp_div)
    cp_title = CharProps(font_name="HY헤드라인M", font_size_pt=17.0, bold=False, color="000000", char_spacing=10)
    cp_title_id = sm.get_char_pr_id(cp_title)
    cp_bar = CharProps(font_name="맑은 고딕", font_size_pt=1.5, bold=True, color="000000", char_spacing=15)
    cp_bar_id = sm.get_char_pr_id(cp_bar)

    # --- ParaPr IDs ---
    pp_num = ParaProps(align="CENTER", line_spacing_value=180.0)
    pp_num_id = sm.get_para_pr_id(pp_num)
    pp_div = ParaProps(align="JUSTIFY", line_spacing_value=180.0)
    pp_div_id = sm.get_para_pr_id(pp_div)
    pp_title = ParaProps(align="LEFT", line_spacing_value=150.0)
    pp_title_id = sm.get_para_pr_id(pp_title)
    pp_bar = ParaProps(align="LEFT", line_spacing_value=180.0)
    pp_bar_id = sm.get_para_pr_id(pp_bar)

    pid = _next_para_id()
    tbl_id = pid + 1000

    # Roman numeral mapping for section numbers
    _ROMAN = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ"]
    idx = (block.num or 1) - 1
    num_text = _ROMAN[idx] if 0 <= idx < len(_ROMAN) else str(block.num)
    title_text = escape(block.text)

    parts: List[str] = []
    # Wrapper paragraph
    parts.append(
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
    )
    parts.append('<hp:run charPrIDRef="0">')

    # hp:tbl (2x3, treatAsChar=1) — matching 섹션.hwpx exactly
    parts.append(
        f'<hp:tbl id="{tbl_id}" zOrder="0" numberingType="TABLE"'
        f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
        f' dropcapstyle="None" pageBreak="CELL" repeatHeader="0"'
        f' rowCnt="2" colCnt="3" cellSpacing="0"'
        f' borderFillIDRef="{bf_tbl_id}" noAdjust="0">'
    )
    parts.append(
        f'<hp:sz width="{total_width}" widthRelTo="ABSOLUTE"'
        f' height="{total_height}" heightRelTo="ABSOLUTE" protect="0"/>'
    )
    parts.append(
        f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="0"'
        f' allowOverlap="0" holdAnchorAndSO="0"'
        f' vertRelTo="PARA" horzRelTo="COLUMN"'
        f' vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
    )
    parts.append('<hp:outMargin left="0" right="0" top="0" bottom="0"/>')
    parts.append('<hp:inMargin left="0" right="0" top="0" bottom="0"/>')

    # --- Row 0 ---
    parts.append('<hp:tr>')

    # Col 0: Number cell (rowSpan=2, blue bg, 휴먼명조 18pt white bold CENTER)
    num_pid = _next_para_id()
    parts.append(
        f'<hp:tc name="" header="0" hasMargin="1" protect="0"'
        f' editable="0" dirty="0" borderFillIDRef="{bf_num_id}">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK"'
        f' vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0"'
        f' textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
        f'<hp:p id="{num_pid}" paraPrIDRef="{pp_num_id}" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{cp_num_id}"><hp:t>{num_text}</hp:t></hp:run>'
        f'<hp:linesegarray>'
        f'<hp:lineseg textpos="0" vertpos="0" vertsize="1800" textheight="1800"'
        f' baseline="1530" spacing="1440" horzpos="0" horzsize="{col0_w}" flags="393216"/>'
        f'</hp:linesegarray>'
        f'</hp:p>'
        f'</hp:subList>'
        f'<hp:cellAddr colAddr="0" rowAddr="0"/>'
        f'<hp:cellSpan colSpan="1" rowSpan="2"/>'
        f'<hp:cellSz width="{col0_w}" height="{total_height}"/>'
        f'<hp:cellMargin left="0" right="0" top="0" bottom="0"/>'
        f'</hp:tc>'
    )

    # Col 1: Divider cell (rowSpan=2, hasMargin=0, 고딕 17pt white bold empty)
    div_pid = _next_para_id()
    parts.append(
        f'<hp:tc name="" header="0" hasMargin="0" protect="0"'
        f' editable="0" dirty="0" borderFillIDRef="{bf_div_id}">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK"'
        f' vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0"'
        f' textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
        f'<hp:p id="{div_pid}" paraPrIDRef="{pp_div_id}" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{cp_div_id}"/>'
        f'<hp:linesegarray>'
        f'<hp:lineseg textpos="0" vertpos="0" vertsize="1700" textheight="1700"'
        f' baseline="1445" spacing="1360" horzpos="0" horzsize="{col1_w}" flags="393216"/>'
        f'</hp:linesegarray>'
        f'</hp:p>'
        f'</hp:subList>'
        f'<hp:cellAddr colAddr="1" rowAddr="0"/>'
        f'<hp:cellSpan colSpan="1" rowSpan="2"/>'
        f'<hp:cellSz width="{col1_w}" height="{total_height}"/>'
        f'<hp:cellMargin left="0" right="0" top="0" bottom="0"/>'
        f'</hp:tc>'
    )

    # Col 2 Row 0: Title text cell (gradation, HY헤드라인M 17pt LEFT ls=150)
    title_pid = _next_para_id()
    title_hz = col2_w - 565  # subtract left cellMargin for horzsize
    parts.append(
        f'<hp:tc name="" header="0" hasMargin="1" protect="0"'
        f' editable="0" dirty="0" borderFillIDRef="{bf_title_id}">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK"'
        f' vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0"'
        f' textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
        f'<hp:p id="{title_pid}" paraPrIDRef="{pp_title_id}" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{cp_title_id}"><hp:t>{title_text}</hp:t></hp:run>'
        f'<hp:linesegarray>'
        f'<hp:lineseg textpos="0" vertpos="0" vertsize="1700" textheight="1700"'
        f' baseline="1445" spacing="852" horzpos="0" horzsize="{title_hz}" flags="393216"/>'
        f'</hp:linesegarray>'
        f'</hp:p>'
        f'</hp:subList>'
        f'<hp:cellAddr colAddr="2" rowAddr="0"/>'
        f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
        f'<hp:cellSz width="{col2_w}" height="{row0_h}"/>'
        f'<hp:cellMargin left="565" right="0" top="140" bottom="0"/>'
        f'</hp:tc>'
    )
    parts.append('</hp:tr>')

    # --- Row 1 ---
    parts.append('<hp:tr>')

    # Col 2 Row 1: Bottom bar (height=150, gradation #999999→#FFFFFF)
    bar_pid = _next_para_id()
    bar_hz = col2_w  # full cell width for horzsize
    parts.append(
        f'<hp:tc name="" header="0" hasMargin="1" protect="0"'
        f' editable="0" dirty="0" borderFillIDRef="{bf_bar_id}">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK"'
        f' vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0"'
        f' textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
        f'<hp:p id="{bar_pid}" paraPrIDRef="{pp_bar_id}" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{cp_bar_id}"/>'
        f'<hp:linesegarray>'
        f'<hp:lineseg textpos="0" vertpos="0" vertsize="150" textheight="150"'
        f' baseline="128" spacing="120" horzpos="600" horzsize="{bar_hz}" flags="393216"/>'
        f'</hp:linesegarray>'
        f'</hp:p>'
        f'</hp:subList>'
        f'<hp:cellAddr colAddr="2" rowAddr="1"/>'
        f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
        f'<hp:cellSz width="{col2_w}" height="{row1_h}"/>'
        f'<hp:cellMargin left="0" right="0" top="0" bottom="0"/>'
        f'</hp:tc>'
    )
    parts.append('</hp:tr>')

    parts.append('</hp:tbl>')
    parts.append('<hp:t/>')
    parts.append('</hp:run>')

    # lineseg for wrapper paragraph
    parts.append(
        '<hp:linesegarray>'
        '<hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000"'
        ' baseline="850" spacing="600" horzpos="0" horzsize="0" flags="393216"/>'
        '</hp:linesegarray>'
    )
    parts.append('</hp:p>')
    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Diagram rendering
# ---------------------------------------------------------------------------

def _render_diagram(diagram, content_w: int) -> str:
    from .shape_builder import build_diagram_xml
    return build_diagram_xml(diagram, content_w=content_w, para_id_fn=_next_para_id)


# ---------------------------------------------------------------------------
# Chart rendering
# ---------------------------------------------------------------------------

def _render_chart(chart, content_w: int) -> str:
    """Render chart matching real HWPX structure.

    Chart is wrapped in hp:switch > hp:case for ooxmlchart namespace,
    matching reference file pattern.
    """
    w_hwp = mm_to_hwp(chart.width_mm)
    h_hwp = mm_to_hwp(chart.height_mm)
    chart_ref = f"Chart/chart{chart.chart_id}.xml"

    pid = _next_para_id()
    chart_id = pid + 5000
    lseg = _lineseg_xml(10.0, 160.0, content_w)

    parts: List[str] = []
    parts.append(
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
    )
    parts.append('<hp:run charPrIDRef="0">')

    # hp:switch wrapping chart — matching reference structure
    parts.append('<hp:switch>')
    parts.append(
        '<hp:case hp:required-namespace='
        '"http://www.hancom.co.kr/hwpml/2016/ooxmlchart">'
    )
    parts.append(
        f'<hp:chart id="{chart_id}" zOrder="0"'
        f' numberingType="PICTURE" textWrap="SQUARE"'
        f' textFlow="BOTH_SIDES" lock="0" dropcapstyle="None"'
        f' chartIDRef="{chart_ref}">'
    )
    parts.append(
        f'<hp:sz width="{w_hwp}" widthRelTo="ABSOLUTE"'
        f' height="{h_hwp}" heightRelTo="ABSOLUTE" protect="0"/>'
    )
    parts.append(
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1"'
        f' allowOverlap="0" holdAnchorAndSO="0"'
        f' vertRelTo="PARA" horzRelTo="COLUMN"'
        f' vertAlign="TOP" horzAlign="LEFT"'
        f' vertOffset="0" horzOffset="0"/>'
    )
    parts.append('<hp:outMargin left="0" right="0" top="0" bottom="0"/>')
    parts.append('</hp:chart>')
    parts.append('</hp:case>')
    parts.append('<hp:default/>')
    parts.append('</hp:switch>')

    parts.append('<hp:t/>')
    parts.append('</hp:run>')
    parts.append(lseg)
    parts.append('</hp:p>')

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Image rendering
# ---------------------------------------------------------------------------

def _render_image(image, sm, content_w: int) -> str:
    """Render an image using hp:pic structure (reverse-engineered from real HWPX files).

    Reference structure from image1.hwpx / image2.hwpx:
    - hp:pic inside hp:run (with hp:t/ at end)
    - hc:img binaryItemIDRef="imageN" references content.hpf opf:item id
    - Caption inside hp:pic as hp:caption > hp:subList > hp:p
    - Image manifest in content.hpf, NOT in header.xml or manifest.xml
    """
    w_hwp = mm_to_hwp(image.width_mm)
    h_mm = image.height_mm if image.height_mm else image.width_mm
    h_hwp = mm_to_hwp(h_mm)

    # imgDim = pixel dimensions × 75 (7200/96 dpi)
    px_to_dim = 75
    dim_w = image._pixel_width * px_to_dim if image._pixel_width else w_hwp * 2
    dim_h = image._pixel_height * px_to_dim if image._pixel_height else h_hwp * 2

    # Alignment: use horzRelTo="COLUMN" with horzAlign
    align_map = {"left": "LEFT", "center": "CENTER", "right": "RIGHT"}
    hwp_align = align_map.get(image.align.lower(), "CENTER")

    # binaryItemIDRef matches the id in content.hpf
    bin_item_ref = f"image{image._bin_id}"

    pic_id = _next_para_id() + 2000000
    instid = pic_id + 1000

    pid = _next_para_id()
    lseg = _lineseg_xml(10.0, 160.0, content_w)

    parts: List[str] = []
    parts.append(
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
    )
    parts.append('<hp:run charPrIDRef="0">')

    # hp:pic element — matching reference structure
    parts.append(
        f'<hp:pic id="{pic_id}" zOrder="0" numberingType="PICTURE"'
        f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
        f' dropcapstyle="None" href="" groupLevel="0" instid="{instid}" reverse="0">'
    )
    parts.append('<hp:offset x="0" y="0"/>')
    parts.append(f'<hp:orgSz width="{w_hwp}" height="{h_hwp}"/>')
    parts.append('<hp:curSz width="0" height="0"/>')
    parts.append('<hp:flip horizontal="0" vertical="0"/>')
    parts.append(
        f'<hp:rotationInfo angle="0" centerX="{w_hwp // 2}"'
        f' centerY="{h_hwp // 2}" rotateimage="1"/>'
    )
    parts.append(
        '<hp:renderingInfo>'
        '<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        '<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        '<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        '</hp:renderingInfo>'
    )
    parts.append(
        f'<hc:img binaryItemIDRef="{bin_item_ref}"'
        f' bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>'
    )
    parts.append(
        f'<hp:imgRect>'
        f'<hc:pt0 x="0" y="0"/>'
        f'<hc:pt1 x="{w_hwp}" y="0"/>'
        f'<hc:pt2 x="{w_hwp}" y="{h_hwp}"/>'
        f'<hc:pt3 x="0" y="{h_hwp}"/>'
        f'</hp:imgRect>'
    )
    parts.append(f'<hp:imgClip left="0" right="{dim_w}" top="0" bottom="{dim_h}"/>')
    parts.append('<hp:inMargin left="0" right="0" top="0" bottom="0"/>')
    parts.append(f'<hp:imgDim dimwidth="{dim_w}" dimheight="{dim_h}"/>')
    parts.append('<hp:effects/>')
    parts.append(
        f'<hp:sz width="{w_hwp}" widthRelTo="ABSOLUTE"'
        f' height="{h_hwp}" heightRelTo="ABSOLUTE" protect="0"/>'
    )
    parts.append(
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1"'
        f' allowOverlap="1" holdAnchorAndSO="0"'
        f' vertRelTo="PARA" horzRelTo="COLUMN"'
        f' vertAlign="TOP" horzAlign="{hwp_align}"'
        f' vertOffset="0" horzOffset="0"/>'
    )
    parts.append('<hp:outMargin left="0" right="0" top="0" bottom="0"/>')

    # Caption inside hp:pic (reference: hp:caption > hp:subList > hp:p)
    if image.caption:
        cap_text = escape(image.caption)
        cap_pid = _next_para_id()
        cap_lseg = _lineseg_xml(10.0, 160.0, w_hwp)
        parts.append(
            f'<hp:caption side="BOTTOM" fullSz="0" width="8504" gap="850"'
            f' lastWidth="{w_hwp}">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK"'
            f' vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0"'
            f' textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
            f'<hp:p id="{cap_pid}" paraPrIDRef="0" styleIDRef="0"'
            f' pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="0"><hp:t>{cap_text}</hp:t></hp:run>'
            f'{cap_lseg}'
            f'</hp:p>'
            f'</hp:subList>'
            f'</hp:caption>'
        )

    parts.append('</hp:pic>')
    parts.append('<hp:t/>')  # empty text after pic (matches reference)
    parts.append('</hp:run>')
    parts.append(lseg)
    parts.append('</hp:p>')

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# SVG rendering
# ---------------------------------------------------------------------------

def _render_svg(svg_elem, content_w: int) -> str:
    from .svg_converter import build_svg_xml
    return build_svg_xml(svg_elem, content_w=content_w, para_id_fn=_next_para_id)


# ---------------------------------------------------------------------------
# TextBox rendering
# ---------------------------------------------------------------------------

def _render_text_box(text_box, sm, content_w: int) -> str:
    """Render TextBox matching real HWPX structure (ref_textbox.hwpx).

    Key structure: hp:rect > ... > hp:drawText > hp:subList > hp:p
    Uses hp:drawText (NOT hp:textbox) with hp:textMargin.
    """
    from ..styles import CharProps

    width_mm = text_box.width_mm if text_box.width_mm else 170.0
    padding = text_box.padding_mm

    if text_box.height_mm:
        height_mm = text_box.height_mm
    else:
        font_size = text_box.font_size_pt or 10.0
        line_count = text_box.text.count("\n") + 1
        line_height_mm = font_size * 0.45
        height_mm = line_height_mm * line_count + padding * 2 + 2.0

    w_hwp = mm_to_hwp(width_mm)
    h_hwp = mm_to_hwp(height_mm)
    pad_hwp = mm_to_hwp(padding)

    fn = text_box.font_name or "함초롬돋움"
    fs = text_box.font_size_pt or 10.0
    fc = text_box.font_color or "000000"

    cp = CharProps(
        font_name=fn, font_size_pt=fs,
        bold=text_box.font_bold, italic=text_box.font_italic, color=fc,
    )
    cp_id = sm.get_char_pr_id(cp)

    pid = _next_para_id()
    rect_id = pid + 3000
    instid = rect_id + 7000
    border_color = text_box.border_color
    bg_color = text_box.bg_color
    line_style = "SOLID"

    lseg = _lineseg_xml(10.0, 160.0, content_w)
    cx = w_hwp // 2
    cy = h_hwp // 2

    # Inner text area width for lineseg
    inner_w = w_hwp - pad_hwp * 2

    parts: List[str] = []
    parts.append(
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
    )
    parts.append('<hp:run charPrIDRef="0">')

    # hp:rect with full attributes — matching ref_textbox.hwpx
    parts.append(
        f'<hp:rect id="{rect_id}" zOrder="0"'
        f' numberingType="PICTURE" textWrap="IN_FRONT_OF_TEXT"'
        f' textFlow="BOTH_SIDES" lock="0" dropcapstyle="None"'
        f' href="" groupLevel="0" instid="{instid}" ratio="0">'
    )

    # Common pre-elements: offset, orgSz, curSz, flip, rotationInfo, renderingInfo
    parts.append('<hp:offset x="0" y="0"/>')
    parts.append(f'<hp:orgSz width="{w_hwp}" height="{h_hwp}"/>')
    parts.append('<hp:curSz width="0" height="0"/>')
    parts.append('<hp:flip horizontal="0" vertical="0"/>')
    parts.append(
        f'<hp:rotationInfo angle="0" centerX="{cx}" centerY="{cy}" rotateimage="1"/>'
    )
    parts.append('<hp:renderingInfo>')
    parts.append('<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>')
    parts.append('<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>')
    parts.append('<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>')
    parts.append('</hp:renderingInfo>')

    # lineShape — matching reference
    parts.append(
        f'<hp:lineShape color="#{border_color}" width="33" style="{line_style}"'
        f' endCap="FLAT" headStyle="NORMAL" tailStyle="NORMAL"'
        f' headfill="1" tailfill="1"'
        f' headSz="MEDIUM_MEDIUM" tailSz="MEDIUM_MEDIUM"'
        f' outlineStyle="NORMAL" alpha="0"/>'
    )

    # fillBrush — matching reference (always present, default white)
    fill_color = bg_color or "FFFFFF"
    parts.append(
        f'<hc:fillBrush>'
        f'<hc:winBrush faceColor="#{fill_color}" hatchColor="#000000" alpha="0"/>'
        f'</hc:fillBrush>'
    )

    # shadow
    parts.append('<hp:shadow type="NONE" color="#B2B2B2" offsetX="0" offsetY="0" alpha="0"/>')

    # hp:drawText — THE KEY DIFFERENCE from wrong hp:textbox
    parts.append(f'<hp:drawText lastWidth="{w_hwp}" name="" editable="0">')
    parts.append(
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK"'
        f' vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0"'
        f' textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
    )

    for line_text in text_box.text.split("\n"):
        inner_pid = _next_para_id()
        inner_lseg = _lineseg_xml(fs, 160.0, inner_w)
        escaped = escape(line_text)
        parts.append(
            f'<hp:p id="{inner_pid}" paraPrIDRef="0" styleIDRef="0"'
            f' pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="{cp_id}"><hp:t>{escaped}</hp:t></hp:run>'
            f'{inner_lseg}'
            f'</hp:p>'
        )

    parts.append('</hp:subList>')
    parts.append(
        f'<hp:textMargin left="{pad_hwp}" right="{pad_hwp}"'
        f' top="{pad_hwp}" bottom="{pad_hwp}"/>'
    )
    parts.append('</hp:drawText>')

    # Corner points (hc:pt0 ~ hc:pt3) — matching reference
    parts.append(f'<hc:pt0 x="0" y="0"/>')
    parts.append(f'<hc:pt1 x="{w_hwp}" y="0"/>')
    parts.append(f'<hc:pt2 x="{w_hwp}" y="{h_hwp}"/>')
    parts.append(f'<hc:pt3 x="0" y="{h_hwp}"/>')

    # sz, pos, outMargin — matching reference
    parts.append(
        f'<hp:sz width="{w_hwp}" widthRelTo="ABSOLUTE"'
        f' height="{h_hwp}" heightRelTo="ABSOLUTE" protect="0"/>'
    )
    parts.append(
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="0"'
        f' allowOverlap="1" holdAnchorAndSO="0"'
        f' vertRelTo="PARA" horzRelTo="PARA"'
        f' vertAlign="TOP" horzAlign="LEFT"'
        f' vertOffset="0" horzOffset="0"/>'
    )
    parts.append('<hp:outMargin left="0" right="0" top="0" bottom="0"/>')

    parts.append('</hp:rect>')
    parts.append('<hp:t/>')
    parts.append('</hp:run>')
    parts.append(lseg)
    parts.append('</hp:p>')

    return "\n".join(parts)
