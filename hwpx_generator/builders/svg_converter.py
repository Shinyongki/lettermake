"""Convert SVG XML to native HWPX shape XML.

Parses SVG using Python's built-in xml.etree.ElementTree, then converts
each SVG element to the corresponding HWPX shape:
    <rect>        -> hp:rect
    <circle>      -> hp:ellipse
    <ellipse>     -> hp:ellipse
    <line>        -> hp:connectLine
    <polygon>     -> hp:polygon
    <polyline>    -> hp:polygon
    <path>        -> hp:polygon (M/L/Z commands; curves approximated)

Structure matches real Hancom HWPX files (reverse-engineered from ref_svg_shapes.hwpx):
- Shapes are direct children of <hp:run>, NOT wrapped in <hp:drawingObject>
- Each shape needs full attribute set: numberingType, textWrap, textFlow,
  offset, orgSz, curSz, flip, rotationInfo, renderingInfo, lineShape,
  fillBrush, shadow, geometry, sz, pos, outMargin
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from typing import List, Tuple, Optional, TYPE_CHECKING

from ..utils import mm_to_hwp

if TYPE_CHECKING:
    from ..elements.svg_element import SvgElement


# SVG namespace
SVG_NS = "http://www.w3.org/2000/svg"

# Fallback default viewBox
_DEFAULT_VIEWBOX = (0, 0, 100, 100)

# instid base (distinct from shape_builder to avoid collisions)
_INSTID_BASE = 1024880000


# ── Public API ───────────────────────────────────────────────────

def build_svg_xml(
    svg_elem: SvgElement,
    start_id: int = 200,
    content_w: int = 42520,
    para_id_fn=None,
) -> str:
    """Build HWPX XML for an SVG element.

    Shapes are direct children of hp:run (matching real Hancom structure).
    """
    svg_text = svg_elem.load_svg()
    root = ET.fromstring(svg_text)

    vb = _parse_viewbox(root)
    vb_x, vb_y, vb_w, vb_h = vb

    total_w_mm = svg_elem.width_mm
    if svg_elem.height_mm is None:
        if vb_w > 0:
            svg_elem.height_mm = total_w_mm * vb_h / vb_w
        else:
            svg_elem.height_mm = total_w_mm * 0.75
    total_h_mm = svg_elem.height_mm

    total_w = mm_to_hwp(total_w_mm)
    total_h = mm_to_hwp(total_h_mm)

    sx = total_w / vb_w if vb_w > 0 else 1.0
    sy = total_h / vb_h if vb_h > 0 else 1.0

    id_counter = [start_id]
    instid_counter = [_INSTID_BASE]
    shape_fragments: List[str] = []

    _process_element(root, vb_x, vb_y, sx, sy, id_counter, instid_counter, shape_fragments)

    pid = para_id_fn() if para_id_fn else 0

    # lineseg
    lseg = (
        f'<hp:linesegarray>'
        f'<hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000"'
        f' baseline="850" spacing="600" horzpos="0"'
        f' horzsize="{content_w}" flags="393216"/>'
        f'</hp:linesegarray>'
    )

    parts: List[str] = []
    parts.append(
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
    )
    parts.append('<hp:run charPrIDRef="0">')

    # Shapes as direct children of hp:run (no hp:drawingObject wrapper)
    for frag in shape_fragments:
        parts.append(frag)

    parts.append('<hp:t/>')
    parts.append('</hp:run>')
    parts.append(lseg)
    parts.append('</hp:p>')

    if svg_elem.caption:
        from xml.sax.saxutils import escape
        caption_text = escape(svg_elem.caption)
        cap_pid = para_id_fn() if para_id_fn else 0
        parts.append(
            f'<hp:p id="{cap_pid}" paraPrIDRef="0" styleIDRef="0"'
            f' pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="0"><hp:t>{caption_text}</hp:t></hp:run>'
            f'{lseg}'
            f'</hp:p>'
        )

    return "\n".join(parts)


# ── ViewBox parsing ──────────────────────────────────────────────

def _parse_viewbox(root: ET.Element) -> Tuple[float, float, float, float]:
    """Extract viewBox from SVG root element."""
    vb_attr = root.get("viewBox")
    if vb_attr:
        parts = re.split(r"[\s,]+", vb_attr.strip())
        if len(parts) >= 4:
            return (float(parts[0]), float(parts[1]),
                    float(parts[2]), float(parts[3]))

    # Fall back to width/height attributes
    w = _parse_length(root.get("width", "100"))
    h = _parse_length(root.get("height", "100"))
    return (0, 0, w, h)


def _parse_length(s: str) -> float:
    """Parse a length value, stripping units like 'px', 'mm', etc."""
    s = s.strip()
    for unit in ("px", "mm", "cm", "in", "pt", "em", "ex", "%"):
        if s.endswith(unit):
            s = s[:-len(unit)].strip()
            break
    try:
        return float(s)
    except (ValueError, TypeError):
        return 100.0


# ── Recursive element processing ─────────────────────────────────

def _process_element(
    elem: ET.Element,
    vb_x: float, vb_y: float,
    sx: float, sy: float,
    id_counter: List[int],
    instid_counter: List[int],
    fragments: List[str],
) -> None:
    """Recursively process SVG elements into HWPX shape fragments."""
    tag = _local_tag(elem)

    converters = {
        "rect": _convert_rect,
        "circle": _convert_circle,
        "ellipse": _convert_ellipse,
        "line": _convert_line,
        "polygon": _convert_polygon,
        "polyline": _convert_polyline,
        "path": _convert_path,
    }

    converter = converters.get(tag)
    if converter:
        frag = converter(elem, vb_x, vb_y, sx, sy, id_counter, instid_counter)
        if frag:
            fragments.append(frag)

    # Recurse into child elements (for <g>, <svg>, etc.)
    for child in elem:
        _process_element(child, vb_x, vb_y, sx, sy, id_counter, instid_counter, fragments)


def _local_tag(elem: ET.Element) -> str:
    """Get the local tag name, stripping namespace."""
    tag = elem.tag
    if "}" in tag:
        tag = tag.split("}", 1)[1]
    return tag


# ── Common HWPX shape structure (matching reference) ─────────────

def _shape_common_pre(w: int, h: int) -> List[str]:
    """Common elements BEFORE shape-specific content.

    Matches reference: hp:offset, hp:orgSz, hp:curSz, hp:flip,
    hp:rotationInfo, hp:renderingInfo.
    """
    cx = w // 2
    cy = h // 2
    parts: List[str] = []
    parts.append('<hp:offset x="0" y="0"/>')
    parts.append(f'<hp:orgSz width="{w}" height="{h}"/>')
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
    return parts


def _line_shape_xml(
    color: str,
    width: str = "33",
    style: str = "SOLID",
    head_style: str = "NORMAL",
    tail_style: str = "NORMAL",
) -> str:
    """Build hp:lineShape matching reference structure."""
    return (
        f'<hp:lineShape color="#{color}" width="{width}" style="{style}"'
        f' endCap="FLAT" headStyle="{head_style}" tailStyle="{tail_style}"'
        f' headfill="1" tailfill="1"'
        f' headSz="MEDIUM_MEDIUM" tailSz="MEDIUM_MEDIUM"'
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
) -> List[str]:
    """Common elements AFTER shape-specific content.

    Matches reference: hp:sz, hp:pos, hp:outMargin.
    """
    parts: List[str] = []
    parts.append(
        f'<hp:sz width="{w}" widthRelTo="ABSOLUTE"'
        f' height="{h}" heightRelTo="ABSOLUTE" protect="0"/>'
    )
    parts.append(
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="0"'
        f' allowOverlap="1" holdAnchorAndSO="0"'
        f' vertRelTo="PARA" horzRelTo="PARA"'
        f' vertAlign="TOP" horzAlign="LEFT"'
        f' vertOffset="{vert_offset}" horzOffset="{horz_offset}"/>'
    )
    parts.append('<hp:outMargin left="0" right="0" top="0" bottom="0"/>')
    return parts


# ── Style extraction helpers ─────────────────────────────────────

def _get_fill_color(elem: ET.Element) -> Optional[str]:
    """Extract fill color from SVG element (attribute or style)."""
    fill = elem.get("fill")
    if not fill:
        style = elem.get("style", "")
        m = re.search(r"fill\s*:\s*([^;]+)", style)
        if m:
            fill = m.group(1).strip()
    if fill and fill.lower() != "none":
        return _svg_color_to_hex(fill)
    return None


def _get_stroke_color(elem: ET.Element) -> Optional[str]:
    """Extract stroke color from SVG element."""
    stroke = elem.get("stroke")
    if not stroke:
        style = elem.get("style", "")
        m = re.search(r"stroke\s*:\s*([^;]+)", style)
        if m:
            stroke = m.group(1).strip()
    if stroke and stroke.lower() != "none":
        return _svg_color_to_hex(stroke)
    return None


def _get_stroke_width(elem: ET.Element) -> str:
    """Extract stroke width, returning HWPX lineShape width value."""
    sw = elem.get("stroke-width")
    if not sw:
        style = elem.get("style", "")
        m = re.search(r"stroke-width\s*:\s*([^;]+)", style)
        if m:
            sw = m.group(1).strip()
    if sw:
        try:
            val = float(sw.replace("px", "").strip())
            return str(max(1, int(val * 33)))
        except (ValueError, TypeError):
            pass
    return "33"


def _svg_color_to_hex(color: str) -> str:
    """Convert an SVG color value to a 6-digit hex string (no #)."""
    color = color.strip()

    if color.startswith("#"):
        h = color[1:]
        if len(h) == 3:
            h = h[0]*2 + h[1]*2 + h[2]*2
        return h.upper()

    m = re.match(r"rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", color)
    if m:
        r, g, b = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"{r:02X}{g:02X}{b:02X}"

    named = {
        "black": "000000", "white": "FFFFFF", "red": "FF0000",
        "green": "008000", "blue": "0000FF", "yellow": "FFFF00",
        "orange": "FFA500", "purple": "800080", "gray": "808080",
        "grey": "808080", "navy": "000080", "teal": "008080",
        "silver": "C0C0C0", "maroon": "800000", "lime": "00FF00",
        "aqua": "00FFFF", "fuchsia": "FF00FF",
    }
    return named.get(color.lower(), "000000")


# ── SVG element converters ───────────────────────────────────────

def _convert_rect(
    elem: ET.Element,
    vb_x: float, vb_y: float,
    sx: float, sy: float,
    id_counter: List[int],
    instid_counter: List[int],
) -> str:
    """Convert SVG <rect> to hp:rect (matching ref_svg_shapes.hwpx)."""
    shape_id = id_counter[0]
    id_counter[0] += 1
    instid = instid_counter[0]
    instid_counter[0] += 1

    x = int((float(elem.get("x", "0")) - vb_x) * sx)
    y = int((float(elem.get("y", "0")) - vb_y) * sy)
    w = int(float(elem.get("width", "0")) * sx)
    h = int(float(elem.get("height", "0")) * sy)

    if w <= 0 or h <= 0:
        return ""

    rx = elem.get("rx")
    ry = elem.get("ry")

    fill = _get_fill_color(elem)
    stroke = _get_stroke_color(elem)
    sw = _get_stroke_width(elem)
    line_style = "SOLID" if stroke else "NONE"
    line_color = stroke or "000000"

    ratio_attr = ' ratio="0"'
    if rx or ry:
        rx_val = float(rx or ry or "0") * sx
        ratio_pct = min(100, max(0, int(rx_val * 100 / (w / 2)))) if w > 0 else 0
        ratio_attr = f' ratio="{ratio_pct}"'

    parts: List[str] = []
    parts.append(
        f'<hp:rect id="{shape_id}" zOrder="0"'
        f' numberingType="PICTURE" textWrap="IN_FRONT_OF_TEXT"'
        f' textFlow="BOTH_SIDES" lock="0" dropcapstyle="None"'
        f' href="" groupLevel="0" instid="{instid}"{ratio_attr}>'
    )

    parts.extend(_shape_common_pre(w, h))
    parts.append(_line_shape_xml(line_color, sw, line_style))
    if fill:
        parts.append(_fill_brush_xml(fill))
    parts.append(_shadow_xml())

    # Corner points (hc:pt0 ~ hc:pt3) — matching reference
    parts.append(f'<hc:pt0 x="0" y="0"/>')
    parts.append(f'<hc:pt1 x="{w}" y="0"/>')
    parts.append(f'<hc:pt2 x="{w}" y="{h}"/>')
    parts.append(f'<hc:pt3 x="0" y="{h}"/>')

    parts.extend(_shape_common_post(w, h, y, x))
    parts.append('</hp:rect>')
    return "\n".join(parts)


def _convert_circle(
    elem: ET.Element,
    vb_x: float, vb_y: float,
    sx: float, sy: float,
    id_counter: List[int],
    instid_counter: List[int],
) -> str:
    """Convert SVG <circle> to hp:ellipse (matching ref_svg_shapes.hwpx)."""
    shape_id = id_counter[0]
    id_counter[0] += 1
    instid = instid_counter[0]
    instid_counter[0] += 1

    cx_svg = float(elem.get("cx", "0"))
    cy_svg = float(elem.get("cy", "0"))
    r = float(elem.get("r", "0"))

    # Bounding box in HWP units
    left = int((cx_svg - r - vb_x) * sx)
    top = int((cy_svg - r - vb_y) * sy)
    w = int(2 * r * sx)
    h = int(2 * r * sy)

    if w <= 0 or h <= 0:
        return ""

    ecx = w // 2
    ecy = h // 2

    fill = _get_fill_color(elem)
    stroke = _get_stroke_color(elem)
    sw = _get_stroke_width(elem)
    line_style = "SOLID" if stroke else "NONE"
    line_color = stroke or "000000"

    parts: List[str] = []
    parts.append(
        f'<hp:ellipse id="{shape_id}" zOrder="0"'
        f' numberingType="PICTURE" textWrap="IN_FRONT_OF_TEXT"'
        f' textFlow="BOTH_SIDES" lock="0" dropcapstyle="None"'
        f' href="" groupLevel="0" instid="{instid}"'
        f' intervalDirty="0" hasArcPr="0" arcType="NORMAL">'
    )

    parts.extend(_shape_common_pre(w, h))
    parts.append(_line_shape_xml(line_color, sw, line_style))
    if fill:
        parts.append(_fill_brush_xml(fill))
    parts.append(_shadow_xml())

    # Ellipse geometry — matching reference
    parts.append(f'<hc:center x="{ecx}" y="{ecy}"/>')
    parts.append(f'<hc:ax1 x="{w}" y="{ecy}"/>')
    parts.append(f'<hc:ax2 x="{ecx}" y="0"/>')
    parts.append(f'<hc:start1 x="0" y="0"/>')
    parts.append(f'<hc:end1 x="0" y="0"/>')
    parts.append(f'<hc:start2 x="0" y="0"/>')
    parts.append(f'<hc:end2 x="0" y="0"/>')

    parts.extend(_shape_common_post(w, h, top, left))
    parts.append('</hp:ellipse>')
    return "\n".join(parts)


def _convert_ellipse(
    elem: ET.Element,
    vb_x: float, vb_y: float,
    sx: float, sy: float,
    id_counter: List[int],
    instid_counter: List[int],
) -> str:
    """Convert SVG <ellipse> to hp:ellipse."""
    shape_id = id_counter[0]
    id_counter[0] += 1
    instid = instid_counter[0]
    instid_counter[0] += 1

    cx_svg = float(elem.get("cx", "0"))
    cy_svg = float(elem.get("cy", "0"))
    erx = float(elem.get("rx", "0"))
    ery = float(elem.get("ry", "0"))

    left = int((cx_svg - erx - vb_x) * sx)
    top = int((cy_svg - ery - vb_y) * sy)
    w = int(2 * erx * sx)
    h = int(2 * ery * sy)

    if w <= 0 or h <= 0:
        return ""

    ecx = w // 2
    ecy = h // 2

    fill = _get_fill_color(elem)
    stroke = _get_stroke_color(elem)
    sw = _get_stroke_width(elem)
    line_style = "SOLID" if stroke else "NONE"
    line_color = stroke or "000000"

    parts: List[str] = []
    parts.append(
        f'<hp:ellipse id="{shape_id}" zOrder="0"'
        f' numberingType="PICTURE" textWrap="IN_FRONT_OF_TEXT"'
        f' textFlow="BOTH_SIDES" lock="0" dropcapstyle="None"'
        f' href="" groupLevel="0" instid="{instid}"'
        f' intervalDirty="0" hasArcPr="0" arcType="NORMAL">'
    )

    parts.extend(_shape_common_pre(w, h))
    parts.append(_line_shape_xml(line_color, sw, line_style))
    if fill:
        parts.append(_fill_brush_xml(fill))
    parts.append(_shadow_xml())

    parts.append(f'<hc:center x="{ecx}" y="{ecy}"/>')
    parts.append(f'<hc:ax1 x="{w}" y="{ecy}"/>')
    parts.append(f'<hc:ax2 x="{ecx}" y="0"/>')
    parts.append(f'<hc:start1 x="0" y="0"/>')
    parts.append(f'<hc:end1 x="0" y="0"/>')
    parts.append(f'<hc:start2 x="0" y="0"/>')
    parts.append(f'<hc:end2 x="0" y="0"/>')

    parts.extend(_shape_common_post(w, h, top, left))
    parts.append('</hp:ellipse>')
    return "\n".join(parts)


def _convert_line(
    elem: ET.Element,
    vb_x: float, vb_y: float,
    sx: float, sy: float,
    id_counter: List[int],
    instid_counter: List[int],
) -> str:
    """Convert SVG <line> to hp:connectLine (matching ref_svg_shapes.hwpx)."""
    shape_id = id_counter[0]
    id_counter[0] += 1
    instid = instid_counter[0]
    instid_counter[0] += 1

    x1 = (float(elem.get("x1", "0")) - vb_x) * sx
    y1 = (float(elem.get("y1", "0")) - vb_y) * sy
    x2 = (float(elem.get("x2", "0")) - vb_x) * sx
    y2 = (float(elem.get("y2", "0")) - vb_y) * sy

    stroke = _get_stroke_color(elem) or "000000"
    sw = _get_stroke_width(elem)

    # Bounding box
    min_x = int(min(x1, x2))
    min_y = int(min(y1, y2))
    line_w = max(1, int(abs(x2 - x1)))
    line_h = max(1, int(abs(y2 - y1)))

    # orgSz uses relative coordinates
    org_w = max(1, int(abs(x2 - x1)))
    org_h = max(1, int(abs(y2 - y1)))

    # Check for arrow marker
    marker_end = elem.get("marker-end", "")
    tail_style = "ARROW" if "arrow" in marker_end.lower() else "NORMAL"

    parts: List[str] = []
    parts.append(
        f'<hp:connectLine id="{shape_id}" zOrder="0"'
        f' numberingType="PICTURE" textWrap="IN_FRONT_OF_TEXT"'
        f' textFlow="BOTH_SIDES" lock="0" dropcapstyle="None"'
        f' href="" groupLevel="0" instid="{instid}"'
        f' type="STRAIGHT_ONEWAY">'
    )

    # Common pre-elements
    cx = line_w // 2
    cy = line_h // 2
    parts.append('<hp:offset x="0" y="0"/>')
    parts.append(f'<hp:orgSz width="{org_w}" height="{org_h}"/>')
    parts.append(f'<hp:curSz width="0" height="0"/>')
    parts.append('<hp:flip horizontal="0" vertical="0"/>')
    parts.append(
        f'<hp:rotationInfo angle="0" centerX="{cx}" centerY="{cy}" rotateimage="1"/>'
    )
    parts.append('<hp:renderingInfo>')
    parts.append('<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>')
    parts.append('<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>')
    parts.append('<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>')
    parts.append('</hp:renderingInfo>')

    parts.append(_line_shape_xml(stroke, sw, "SOLID", tail_style=tail_style))
    parts.append(_shadow_xml())

    # Start/end points — matching reference
    parts.append(f'<hp:startPt x="0" y="0" subjectIDRef="0" subjectIdx="0"/>')
    parts.append(
        f'<hp:endPt x="{org_w}" y="{org_h}" subjectIDRef="0" subjectIdx="0"/>'
    )
    parts.append('<hp:controlPoints>')
    parts.append(f'<hp:point x="0" y="0" type="3"/>')
    parts.append(f'<hp:point x="{org_w}" y="0" type="26"/>')
    parts.append('</hp:controlPoints>')

    parts.extend(_shape_common_post(line_w, line_h, min_y, min_x))
    parts.append('</hp:connectLine>')
    return "\n".join(parts)


def _convert_polygon(
    elem: ET.Element,
    vb_x: float, vb_y: float,
    sx: float, sy: float,
    id_counter: List[int],
    instid_counter: List[int],
) -> str:
    """Convert SVG <polygon> to hp:polygon."""
    points = _parse_points(elem.get("points", ""))
    if not points:
        return ""
    return _build_polygon_from_points(
        points, elem, vb_x, vb_y, sx, sy, id_counter, instid_counter
    )


def _convert_polyline(
    elem: ET.Element,
    vb_x: float, vb_y: float,
    sx: float, sy: float,
    id_counter: List[int],
    instid_counter: List[int],
) -> str:
    """Convert SVG <polyline> to hp:polygon."""
    points = _parse_points(elem.get("points", ""))
    if not points:
        return ""
    return _build_polygon_from_points(
        points, elem, vb_x, vb_y, sx, sy, id_counter, instid_counter
    )


def _convert_path(
    elem: ET.Element,
    vb_x: float, vb_y: float,
    sx: float, sy: float,
    id_counter: List[int],
    instid_counter: List[int],
) -> str:
    """Convert SVG <path> to hp:polygon.

    Supports M, L, H, V, Z commands (absolute and relative).
    C, Q, S, T, A commands are approximated by using endpoint positions.
    """
    d = elem.get("d", "")
    if not d:
        return ""
    points = _parse_path_d(d)
    if not points:
        return ""
    return _build_polygon_from_points(
        points, elem, vb_x, vb_y, sx, sy, id_counter, instid_counter
    )


# ── Polygon builder helper ───────────────────────────────────────

def _build_polygon_from_points(
    points: List[Tuple[float, float]],
    elem: ET.Element,
    vb_x: float, vb_y: float,
    sx: float, sy: float,
    id_counter: List[int],
    instid_counter: List[int],
) -> str:
    """Build hp:polygon from SVG coordinate points (matching reference structure)."""
    shape_id = id_counter[0]
    id_counter[0] += 1
    instid = instid_counter[0]
    instid_counter[0] += 1

    # Transform to HWP units
    hwp_points = []
    for px, py in points:
        hx = int((px - vb_x) * sx)
        hy = int((py - vb_y) * sy)
        hwp_points.append((hx, hy))

    if not hwp_points:
        return ""

    # Bounding box
    xs = [p[0] for p in hwp_points]
    ys = [p[1] for p in hwp_points]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    w = max(1, max_x - min_x)
    h = max(1, max_y - min_y)

    # Make points relative to bounding box origin
    rel_points = [(px - min_x, py - min_y) for px, py in hwp_points]
    # Close polygon by repeating first point (reference pattern)
    if rel_points and rel_points[0] != rel_points[-1]:
        rel_points.append(rel_points[0])

    fill = _get_fill_color(elem)
    stroke = _get_stroke_color(elem)
    sw = _get_stroke_width(elem)
    line_style = "SOLID" if stroke else ("NONE" if fill else "SOLID")
    line_color = stroke or "000000"
    line_w = sw if stroke else "0"

    parts: List[str] = []
    parts.append(
        f'<hp:polygon id="{shape_id}" zOrder="0"'
        f' numberingType="PICTURE" textWrap="IN_FRONT_OF_TEXT"'
        f' textFlow="BOTH_SIDES" lock="0" dropcapstyle="None"'
        f' href="" groupLevel="0" instid="{instid}">'
    )

    parts.extend(_shape_common_pre(w, h))
    parts.append(_line_shape_xml(line_color, line_w, line_style))
    if fill:
        parts.append(_fill_brush_xml(fill))
    parts.append(_shadow_xml())

    # Polygon points
    for px, py in rel_points:
        parts.append(f'<hc:pt x="{px}" y="{py}"/>')

    parts.extend(_shape_common_post(w, h, min_y, min_x))
    parts.append('</hp:polygon>')
    return "\n".join(parts)


# ── Points parsing ───────────────────────────────────────────────

def _parse_points(points_str: str) -> List[Tuple[float, float]]:
    """Parse SVG points attribute (space or comma separated pairs)."""
    if not points_str.strip():
        return []
    nums = re.findall(r"[-+]?[0-9]*\.?[0-9]+(?:[eE][-+]?[0-9]+)?", points_str)
    result = []
    for i in range(0, len(nums) - 1, 2):
        result.append((float(nums[i]), float(nums[i + 1])))
    return result


# ── SVG path `d` attribute parser ────────────────────────────────

def _parse_path_d(d: str) -> List[Tuple[float, float]]:
    """Parse SVG path data string into a list of points.

    Handles M/m, L/l, H/h, V/v, Z/z, C/c, S/s, Q/q, T/t, A/a.
    For curves, we approximate by taking endpoint positions.
    """
    tokens = re.findall(
        r"[MmLlHhVvZzCcSsQqTtAa]|[-+]?[0-9]*\.?[0-9]+(?:[eE][-+]?[0-9]+)?",
        d,
    )

    points: List[Tuple[float, float]] = []
    cx, cy = 0.0, 0.0
    start_x, start_y = 0.0, 0.0

    i = 0
    while i < len(tokens):
        cmd = tokens[i]
        if not cmd:
            i += 1
            continue

        if cmd in ("M", "m"):
            i += 1
            is_rel = cmd == "m"
            first = True
            while i < len(tokens) and _is_number(tokens[i]):
                x_val = float(tokens[i]); i += 1
                y_val = 0.0
                if i < len(tokens) and _is_number(tokens[i]):
                    y_val = float(tokens[i]); i += 1
                if is_rel:
                    cx += x_val; cy += y_val
                else:
                    cx = x_val; cy = y_val
                if first:
                    start_x, start_y = cx, cy
                    first = False
                points.append((cx, cy))

        elif cmd in ("L", "l"):
            i += 1
            is_rel = cmd == "l"
            while i < len(tokens) and _is_number(tokens[i]):
                x_val = float(tokens[i]); i += 1
                y_val = 0.0
                if i < len(tokens) and _is_number(tokens[i]):
                    y_val = float(tokens[i]); i += 1
                if is_rel:
                    cx += x_val; cy += y_val
                else:
                    cx = x_val; cy = y_val
                points.append((cx, cy))

        elif cmd in ("H", "h"):
            i += 1
            is_rel = cmd == "h"
            while i < len(tokens) and _is_number(tokens[i]):
                x_val = float(tokens[i]); i += 1
                if is_rel:
                    cx += x_val
                else:
                    cx = x_val
                points.append((cx, cy))

        elif cmd in ("V", "v"):
            i += 1
            is_rel = cmd == "v"
            while i < len(tokens) and _is_number(tokens[i]):
                y_val = float(tokens[i]); i += 1
                if is_rel:
                    cy += y_val
                else:
                    cy = y_val
                points.append((cx, cy))

        elif cmd in ("C", "c"):
            i += 1
            is_rel = cmd == "c"
            while i < len(tokens) and _is_number(tokens[i]):
                nums = []
                for _ in range(6):
                    if i < len(tokens) and _is_number(tokens[i]):
                        nums.append(float(tokens[i])); i += 1
                    else:
                        nums.append(0.0)
                if is_rel:
                    cx += nums[4]; cy += nums[5]
                else:
                    cx = nums[4]; cy = nums[5]
                points.append((cx, cy))

        elif cmd in ("S", "s"):
            i += 1
            is_rel = cmd == "s"
            while i < len(tokens) and _is_number(tokens[i]):
                nums = []
                for _ in range(4):
                    if i < len(tokens) and _is_number(tokens[i]):
                        nums.append(float(tokens[i])); i += 1
                    else:
                        nums.append(0.0)
                if is_rel:
                    cx += nums[2]; cy += nums[3]
                else:
                    cx = nums[2]; cy = nums[3]
                points.append((cx, cy))

        elif cmd in ("Q", "q"):
            i += 1
            is_rel = cmd == "q"
            while i < len(tokens) and _is_number(tokens[i]):
                nums = []
                for _ in range(4):
                    if i < len(tokens) and _is_number(tokens[i]):
                        nums.append(float(tokens[i])); i += 1
                    else:
                        nums.append(0.0)
                if is_rel:
                    cx += nums[2]; cy += nums[3]
                else:
                    cx = nums[2]; cy = nums[3]
                points.append((cx, cy))

        elif cmd in ("T", "t"):
            i += 1
            is_rel = cmd == "t"
            while i < len(tokens) and _is_number(tokens[i]):
                x_val = float(tokens[i]); i += 1
                y_val = 0.0
                if i < len(tokens) and _is_number(tokens[i]):
                    y_val = float(tokens[i]); i += 1
                if is_rel:
                    cx += x_val; cy += y_val
                else:
                    cx = x_val; cy = y_val
                points.append((cx, cy))

        elif cmd in ("A", "a"):
            i += 1
            is_rel = cmd == "a"
            while i < len(tokens) and _is_number(tokens[i]):
                nums = []
                for _ in range(7):
                    if i < len(tokens) and _is_number(tokens[i]):
                        nums.append(float(tokens[i])); i += 1
                    else:
                        nums.append(0.0)
                if is_rel:
                    cx += nums[5]; cy += nums[6]
                else:
                    cx = nums[5]; cy = nums[6]
                points.append((cx, cy))

        elif cmd in ("Z", "z"):
            cx, cy = start_x, start_y
            i += 1

        else:
            i += 1

    return points


def _is_number(token: str) -> bool:
    """Check if a token is a number (not a command letter)."""
    try:
        float(token)
        return True
    except (ValueError, TypeError):
        return False
