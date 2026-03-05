"""Build native HWPX shape XML for diagrams.

Generates hp:rect, hp:ellipse, hp:polygon, and hp:connectLine elements
matching real Hancom HWPX structure (reverse-engineered from shape.hwpx).

Key structural rules from reference:
- Shapes are direct children of <hp:run>, NOT wrapped in <hp:drawingObject>
- Each shape needs full attribute set: numberingType, textWrap, textFlow,
  offset, orgSz, curSz, flip, rotationInfo, renderingInfo, lineShape,
  fillBrush, shadow, sz, pos, outMargin
- hp:rect uses hc:pt0~pt3 for corner points
- hp:ellipse uses hc:center, hc:ax1, hc:ax2, hc:start1/end1/start2/end2
- hp:connectLine uses hp:startPt/hp:endPt with control points
- hp:polygon closes path by repeating first point
"""

from __future__ import annotations

from typing import List, Dict, Tuple, TYPE_CHECKING
from xml.sax.saxutils import escape

from ..utils import mm_to_hwp

if TYPE_CHECKING:
    from ..elements.diagram import Diagram, Node, Edge


# ── Layout engine ────────────────────────────────────────────────

def _compute_step_flow_layout(
    diagram: Diagram,
) -> Dict[str, Tuple[int, int, int, int]]:
    """Compute positions for step_flow layout.

    Returns a dict mapping node_id -> (x, y, w, h) in HWP units.
    """
    n = len(diagram.nodes)
    if n == 0:
        return {}

    total_w = mm_to_hwp(diagram.width_mm)
    total_h = mm_to_hwp(diagram.height_mm)

    positions: Dict[str, Tuple[int, int, int, int]] = {}

    if diagram.direction == "horizontal":
        if n == 1:
            node_w = total_w
            gap = 0
        else:
            gap_ratio = 0.20
            node_w = int(total_w / (n + gap_ratio * (n - 1)))
            gap = int(node_w * gap_ratio)

        node_h = total_h
        x_cursor = 0
        for node in diagram.nodes:
            positions[node.id] = (x_cursor, 0, node_w, node_h)
            x_cursor += node_w + gap
    else:
        # Vertical layout
        if n == 1:
            node_h = total_h
            gap = 0
        else:
            gap_ratio = 0.20
            node_h = int(total_h / (n + gap_ratio * (n - 1)))
            gap = int(node_h * gap_ratio)

        node_w = total_w
        y_cursor = 0
        for node in diagram.nodes:
            positions[node.id] = (0, y_cursor, node_w, node_h)
            y_cursor += node_h + gap

    return positions


# ── Shared building blocks matching reference structure ──────────

_INSTID_BASE = 1024860000


def _shape_common_pre(
    shape_id: int,
    z_order: int,
    instid: int,
    w: int, h: int,
    extra_attrs: str = "",
) -> List[str]:
    """Common attributes and child elements BEFORE shape-specific content.

    Matches reference file element order:
      hp:offset, hp:orgSz, hp:curSz, hp:flip, hp:rotationInfo, hp:renderingInfo
    """
    cx = w // 2
    cy = h // 2

    parts: List[str] = []
    # hp:offset
    parts.append('<hp:offset x="0" y="0"/>')
    # hp:orgSz — original size = current size for newly created shapes
    parts.append(f'<hp:orgSz width="{w}" height="{h}"/>')
    # hp:curSz — 0,0 means same as orgSz (reference pattern)
    parts.append('<hp:curSz width="0" height="0"/>')
    # hp:flip
    parts.append('<hp:flip horizontal="0" vertical="0"/>')
    # hp:rotationInfo
    parts.append(
        f'<hp:rotationInfo angle="0" centerX="{cx}" centerY="{cy}" rotateimage="1"/>'
    )
    # hp:renderingInfo with identity matrices
    parts.append('<hp:renderingInfo>')
    parts.append('<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>')
    parts.append('<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>')
    parts.append('<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>')
    parts.append('</hp:renderingInfo>')
    return parts


def _line_shape_xml(
    color: str,
    width: str = "33",
    style: str = "SOLID",
    head_style: str = "NORMAL",
    tail_style: str = "NORMAL",
    tail_sz: str = "MEDIUM_MEDIUM",
) -> str:
    """Build hp:lineShape matching reference structure."""
    return (
        f'<hp:lineShape color="#{color}" width="{width}" style="{style}"'
        f' endCap="FLAT" headStyle="{head_style}" tailStyle="{tail_style}"'
        f' headfill="1" tailfill="1"'
        f' headSz="MEDIUM_MEDIUM" tailSz="{tail_sz}"'
        f' outlineStyle="NORMAL" alpha="0"/>'
    )


def _fill_brush_xml(face_color: str) -> str:
    """Build hc:fillBrush matching reference structure."""
    return (
        f'<hc:fillBrush>'
        f'<hc:winBrush faceColor="#{face_color}" hatchColor="#000000" alpha="0"/>'
        f'</hc:fillBrush>'
    )


def _shadow_xml() -> str:
    """Build hp:shadow (NONE type) matching reference."""
    return '<hp:shadow type="NONE" color="#B2B2B2" offsetX="0" offsetY="0" alpha="0"/>'


def _shape_common_post(
    w: int, h: int,
    vert_offset: int, horz_offset: int,
    comment: str = "",
) -> List[str]:
    """Common elements AFTER shape-specific content.

    Matches reference file element order:
      hp:sz, hp:pos, hp:outMargin, hp:shapeComment
    """
    parts: List[str] = []
    parts.append(
        f'<hp:sz width="{w}" widthRelTo="ABSOLUTE"'
        f' height="{h}" heightRelTo="ABSOLUTE" protect="0"/>'
    )
    parts.append(
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1"'
        f' allowOverlap="1" holdAnchorAndSO="0"'
        f' vertRelTo="PARA" horzRelTo="PARA"'
        f' vertAlign="TOP" horzAlign="LEFT"'
        f' vertOffset="{vert_offset}" horzOffset="{horz_offset}"/>'
    )
    parts.append('<hp:outMargin left="0" right="0" top="0" bottom="0"/>')
    if comment:
        parts.append(f'<hp:shapeComment>{escape(comment)}</hp:shapeComment>')
    return parts


# ── XML builders for individual shapes ───────────────────────────

def _build_rect_xml(
    node: Node,
    x: int, y: int, w: int, h: int,
    theme_color: str,
    id_counter: List[int],
    instid_counter: List[int],
    rounded: bool = False,
) -> str:
    """Build hp:rect XML matching real HWPX structure."""
    shape_id = id_counter[0]
    id_counter[0] += 1
    instid = instid_counter[0]
    instid_counter[0] += 1
    z_order = shape_id - id_counter[0] + 100  # sequential

    fill_color = node.fill_color or _lighter_color(theme_color)
    line_color = node.line_color or theme_color

    ratio_attr = ' ratio="0"'
    if rounded:
        ratio_attr = ' ratio="20"'

    parts: List[str] = []
    parts.append(
        f'<hp:rect id="{shape_id}" zOrder="{z_order}"'
        f' numberingType="PICTURE" textWrap="IN_FRONT_OF_TEXT"'
        f' textFlow="BOTH_SIDES" lock="0" dropcapstyle="None"'
        f' href="" groupLevel="0" instid="{instid}"{ratio_attr}>'
    )

    # Common pre-elements
    parts.extend(_shape_common_pre(shape_id, z_order, instid, w, h))

    # lineShape
    parts.append(_line_shape_xml(line_color))

    # fillBrush
    parts.append(_fill_brush_xml(fill_color))

    # shadow
    parts.append(_shadow_xml())

    # Corner points (hc:pt0 ~ hc:pt3) — matches reference
    parts.append(f'<hc:pt0 x="0" y="0"/>')
    parts.append(f'<hc:pt1 x="{w}" y="0"/>')
    parts.append(f'<hc:pt2 x="{w}" y="{h}"/>')
    parts.append(f'<hc:pt3 x="0" y="{h}"/>')

    # Common post-elements (sz, pos, outMargin, shapeComment)
    parts.extend(_shape_common_post(w, h, y, x))

    # Internal text (textbox with hp:subList)
    if node.label:
        parts.extend(_build_textbox_xml(node, w, h))

    parts.append('</hp:rect>')
    return "\n".join(parts)


def _build_textbox_xml(node: Node, w: int, h: int) -> List[str]:
    """Build hp:drawText inside a shape for node label text.

    Uses hp:drawText (NOT hp:textbox) matching ref_textbox.hwpx structure.
    """
    pad = 283  # ~1mm padding
    parts: List[str] = []
    parts.append(f'<hp:drawText lastWidth="{w}" name="" editable="0">')
    parts.append(
        '<hp:subList id="" textDirection="HORIZONTAL"'
        ' lineWrap="BREAK" vertAlign="CENTER"'
        ' linkListIDRef="0" linkListNextIDRef="0"'
        ' textWidth="0" textHeight="0"'
        ' hasTextRef="0" hasNumRef="0">'
    )

    # Each line of the label as a paragraph
    lines = node.label.split("\n")
    for line_text in lines:
        escaped = escape(line_text)
        parts.append(
            '<hp:p id="0" paraPrIDRef="0" styleIDRef="0"'
            ' pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="0"><hp:t>{escaped}</hp:t></hp:run>'
            '<hp:linesegarray>'
            '<hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000"'
            f' baseline="850" spacing="600" horzpos="0" horzsize="{w}" flags="393216"/>'
            '</hp:linesegarray>'
            '</hp:p>'
        )

    parts.append('</hp:subList>')
    parts.append(
        f'<hp:textMargin left="{pad}" right="{pad}"'
        f' top="{pad}" bottom="{pad}"/>'
    )
    parts.append('</hp:drawText>')
    return parts


def _build_ellipse_xml(
    node: Node,
    x: int, y: int, w: int, h: int,
    theme_color: str,
    id_counter: List[int],
    instid_counter: List[int],
) -> str:
    """Build hp:ellipse XML matching real HWPX structure."""
    shape_id = id_counter[0]
    id_counter[0] += 1
    instid = instid_counter[0]
    instid_counter[0] += 1
    z_order = shape_id - id_counter[0] + 100

    fill_color = node.fill_color or _lighter_color(theme_color)
    line_color = node.line_color or theme_color

    cx = w // 2
    cy = h // 2

    parts: List[str] = []
    parts.append(
        f'<hp:ellipse id="{shape_id}" zOrder="{z_order}"'
        f' numberingType="PICTURE" textWrap="IN_FRONT_OF_TEXT"'
        f' textFlow="BOTH_SIDES" lock="0" dropcapstyle="None"'
        f' href="" groupLevel="0" instid="{instid}"'
        f' intervalDirty="0" hasArcPr="0" arcType="NORMAL">'
    )

    # Common pre-elements
    parts.extend(_shape_common_pre(shape_id, z_order, instid, w, h))

    # lineShape
    parts.append(_line_shape_xml(line_color))

    # fillBrush
    parts.append(_fill_brush_xml(fill_color))

    # shadow
    parts.append(_shadow_xml())

    # Ellipse geometry — matching reference
    parts.append(f'<hc:center x="{cx}" y="{cy}"/>')
    parts.append(f'<hc:ax1 x="{w}" y="{cy}"/>')
    parts.append(f'<hc:ax2 x="{cx}" y="0"/>')
    parts.append(f'<hc:start1 x="0" y="0"/>')
    parts.append(f'<hc:end1 x="0" y="0"/>')
    parts.append(f'<hc:start2 x="0" y="0"/>')
    parts.append(f'<hc:end2 x="0" y="0"/>')

    # Common post-elements
    parts.extend(_shape_common_post(w, h, y, x))

    # Internal text
    if node.label:
        parts.extend(_build_textbox_xml(node, w, h))

    parts.append('</hp:ellipse>')
    return "\n".join(parts)


def _build_diamond_xml(
    node: Node,
    x: int, y: int, w: int, h: int,
    theme_color: str,
    id_counter: List[int],
    instid_counter: List[int],
) -> str:
    """Build hp:polygon (4-point diamond) XML matching real HWPX structure."""
    shape_id = id_counter[0]
    id_counter[0] += 1
    instid = instid_counter[0]
    instid_counter[0] += 1
    z_order = shape_id - id_counter[0] + 100

    fill_color = node.fill_color or _lighter_color(theme_color)
    line_color = node.line_color or theme_color

    parts: List[str] = []
    parts.append(
        f'<hp:polygon id="{shape_id}" zOrder="{z_order}"'
        f' numberingType="PICTURE" textWrap="IN_FRONT_OF_TEXT"'
        f' textFlow="BOTH_SIDES" lock="0" dropcapstyle="None"'
        f' href="" groupLevel="0" instid="{instid}">'
    )

    # Common pre-elements
    parts.extend(_shape_common_pre(shape_id, z_order, instid, w, h))

    # lineShape
    parts.append(_line_shape_xml(line_color))

    # fillBrush
    parts.append(_fill_brush_xml(fill_color))

    # shadow
    parts.append(_shadow_xml())

    # Diamond points — top, right, bottom, left, close with repeat of first
    pts = [
        (w // 2, 0),
        (w, h // 2),
        (w // 2, h),
        (0, h // 2),
        (w // 2, 0),  # close polygon (reference pattern)
    ]
    for px, py in pts:
        parts.append(f'<hc:pt x="{px}" y="{py}"/>')

    # Common post-elements
    parts.extend(_shape_common_post(w, h, y, x))

    # Internal text
    if node.label:
        parts.extend(_build_textbox_xml(node, w, h))

    parts.append('</hp:polygon>')
    return "\n".join(parts)


def _build_connectline_xml(
    edge: Edge,
    positions: Dict[str, Tuple[int, int, int, int]],
    direction: str,
    theme_color: str,
    id_counter: List[int],
    instid_counter: List[int],
) -> str:
    """Build hp:connectLine XML matching real HWPX structure."""
    shape_id = id_counter[0]
    id_counter[0] += 1
    instid = instid_counter[0]
    instid_counter[0] += 1
    z_order = shape_id - id_counter[0] + 100

    line_color = edge.line_color or theme_color

    from_pos = positions.get(edge.from_id)
    to_pos = positions.get(edge.to_id)
    if from_pos is None or to_pos is None:
        return ""

    fx, fy, fw, fh = from_pos
    tx, ty, tw, th = to_pos

    if direction == "horizontal":
        # Arrow from right edge of 'from' to left edge of 'to'
        x1 = fx + fw
        y1 = fy + fh // 2
        x2 = tx
        y2 = ty + th // 2
    else:
        # Arrow from bottom edge of 'from' to top edge of 'to'
        x1 = fx + fw // 2
        y1 = fy + fh
        x2 = tx + tw // 2
        y2 = ty

    # Line bounding box
    line_w = abs(x2 - x1)
    line_h = abs(y2 - y1) or 1  # avoid zero height
    min_x = min(x1, x2)
    min_y = min(y1, y2)
    cx = line_w // 2
    cy = line_h // 2

    # orgSz uses the original coordinate range
    org_w = abs(x2 - x1) or 1
    org_h = abs(y2 - y1) or 1

    parts: List[str] = []
    parts.append(
        f'<hp:connectLine id="{shape_id}" zOrder="{z_order}"'
        f' numberingType="PICTURE" textWrap="IN_FRONT_OF_TEXT"'
        f' textFlow="BOTH_SIDES" lock="0" dropcapstyle="None"'
        f' href="" groupLevel="0" instid="{instid}"'
        f' type="STRAIGHT_ONEWAY">'
    )

    # hp:offset, orgSz, curSz, flip, rotationInfo, renderingInfo
    parts.append('<hp:offset x="0" y="0"/>')
    parts.append(f'<hp:orgSz width="{org_w}" height="{org_h}"/>')
    parts.append(f'<hp:curSz width="{line_w}" height="{line_h}"/>')
    parts.append('<hp:flip horizontal="0" vertical="0"/>')
    parts.append(
        f'<hp:rotationInfo angle="0" centerX="{cx}" centerY="{cy}" rotateimage="1"/>'
    )
    parts.append('<hp:renderingInfo>')
    parts.append('<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>')
    # Scale matrix
    sx = line_w / org_w if org_w else 1
    sy = line_h / org_h if org_h else 1
    parts.append(f'<hc:scaMatrix e1="{sx:.6f}" e2="0" e3="0" e4="0" e5="{sy:.6f}" e6="0"/>')
    parts.append('<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>')
    parts.append('</hp:renderingInfo>')

    # lineShape with arrow
    parts.append(_line_shape_xml(
        line_color,
        head_style=edge.head_style,
        tail_style=edge.tail_style,
        tail_sz=edge.tail_sz,
    ))

    # shadow
    parts.append(_shadow_xml())

    # Start/end points — matching reference structure
    parts.append(f'<hp:startPt x="0" y="0"/>')
    parts.append(f'<hp:endPt x="{org_w}" y="{org_h}"/>')
    parts.append('<hp:controlPoints>')
    parts.append(f'<hp:point x="0" y="0" type="3"/>')
    parts.append(f'<hp:point x="{org_w}" y="{org_h}" type="26"/>')
    parts.append('</hp:controlPoints>')

    # sz, pos, outMargin
    parts.append(
        f'<hp:sz width="{line_w}" widthRelTo="ABSOLUTE"'
        f' height="{line_h}" heightRelTo="ABSOLUTE" protect="0"/>'
    )
    parts.append(
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1"'
        f' allowOverlap="1" holdAnchorAndSO="0"'
        f' vertRelTo="PARA" horzRelTo="PARA"'
        f' vertAlign="TOP" horzAlign="LEFT"'
        f' vertOffset="{min_y}" horzOffset="{min_x}"/>'
    )
    parts.append('<hp:outMargin left="0" right="0" top="0" bottom="0"/>')

    parts.append('</hp:connectLine>')
    return "\n".join(parts)


# ── Main entry point ─────────────────────────────────────────────

def build_diagram_xml(
    diagram: Diagram,
    start_id: int = 100,
    content_w: int = 42520,
    para_id_fn=None,
) -> str:
    """Build the full XML fragment for a Diagram block.

    All shapes are placed as direct children of <hp:run> inside a single
    <hp:p>, matching real Hancom HWPX structure (no hp:drawingObject wrapper).
    """
    positions = _compute_step_flow_layout(diagram)

    id_counter = [start_id]
    instid_counter = [_INSTID_BASE]
    theme = diagram.theme_color

    total_w = mm_to_hwp(diagram.width_mm)
    total_h = mm_to_hwp(diagram.height_mm)

    shape_fragments: List[str] = []

    for node in diagram.nodes:
        pos = positions.get(node.id)
        if pos is None:
            continue
        x, y, w, h = pos

        if node.shape == "rect":
            shape_fragments.append(
                _build_rect_xml(node, x, y, w, h, theme, id_counter, instid_counter)
            )
        elif node.shape == "rounded_rect":
            shape_fragments.append(
                _build_rect_xml(node, x, y, w, h, theme, id_counter, instid_counter, rounded=True)
            )
        elif node.shape == "ellipse":
            shape_fragments.append(
                _build_ellipse_xml(node, x, y, w, h, theme, id_counter, instid_counter)
            )
        elif node.shape == "diamond":
            shape_fragments.append(
                _build_diamond_xml(node, x, y, w, h, theme, id_counter, instid_counter)
            )

    for edge in diagram.edges:
        frag = _build_connectline_xml(
            edge, positions, diagram.direction, theme, id_counter, instid_counter
        )
        if frag:
            shape_fragments.append(frag)

    pid = para_id_fn() if para_id_fn else 0

    # lineseg for this paragraph
    lseg = (
        '<hp:linesegarray>'
        '<hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000"'
        f' baseline="850" spacing="600" horzpos="0"'
        f' horzsize="{content_w}" flags="393216"/>'
        '</hp:linesegarray>'
    )

    # Build paragraph — shapes are direct children of hp:run
    parts: List[str] = []
    parts.append(
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
    )
    parts.append('<hp:run charPrIDRef="0">')

    for frag in shape_fragments:
        parts.append(frag)

    parts.append('<hp:t/>')
    parts.append('</hp:run>')
    parts.append(lseg)
    parts.append('</hp:p>')

    return "\n".join(parts)


# ── Helpers ──────────────────────────────────────────────────────

def _lighter_color(hex_color: str) -> str:
    """Create a lighter tint of a hex color for fill backgrounds.

    Blends the color towards white by ~60%.
    """
    h = hex_color.lstrip("#")
    if len(h) != 6:
        return "D6E4F0"  # fallback light blue

    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)

    factor = 0.6
    r2 = int(r + (255 - r) * factor)
    g2 = int(g + (255 - g) * factor)
    b2 = int(b + (255 - b) * factor)

    return f"{r2:02X}{g2:02X}{b2:02X}"
