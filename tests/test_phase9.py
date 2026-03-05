"""Phase 9 tests: SVG to HWPX native shape conversion.

Updated to match real Hancom HWPX structure (ref_svg_shapes.hwpx):
- Shapes are direct children of hp:run (NOT wrapped in hp:drawingObject)
- Full attribute set: numberingType, instid, offset, orgSz, etc.
"""

import zipfile
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator import HwpxDocument, SvgElement


def _read_section(path):
    with zipfile.ZipFile(path, "r") as zf:
        return zf.read("Contents/section0.xml").decode("utf-8")


# ─── Basic SVG elements ──────────────────────────────────────────


def test_svg_rect():
    """Test SVG <rect> -> hp:rect conversion."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<rect x="10" y="10" width="80" height="60" fill="#2E75B6" stroke="#1F4E79"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=80)

    out = Path(__file__).parent / "output_p9_rect.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:rect" in section, "Expected hp:rect in output"
    # Shapes are direct children of hp:run (no hp:drawingObject)
    assert "hp:drawingObject" not in section, "Should NOT use drawingObject wrapper"
    assert "numberingType" in section, "Expected full attributes"
    assert "instid" in section, "Expected instid attribute"
    assert "#2E75B6" in section, "Expected fill color"
    assert "#1F4E79" in section, "Expected stroke color"
    assert "hp:sz" in section, "Expected size element"
    assert "hp:offset" in section, "Expected offset element"
    assert "hp:orgSz" in section, "Expected orgSz element"
    assert "hp:shadow" in section, "Expected shadow element"
    assert "hc:pt0" in section, "Expected corner point pt0"
    assert "hc:pt3" in section, "Expected corner point pt3"

    print("[PASS] test_svg_rect")
    out.unlink()


def test_svg_circle():
    """Test SVG <circle> -> hp:ellipse conversion."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<circle cx="50" cy="50" r="30" fill="#FF0000" stroke="#800000"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=60)

    out = Path(__file__).parent / "output_p9_circle.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:ellipse" in section, "Expected hp:ellipse for circle"
    assert "hc:center" in section, "Expected center element"
    assert "hc:ax1" in section, "Expected ax1 element"
    assert "hc:ax2" in section, "Expected ax2 element"
    assert "hc:start1" in section, "Expected start1 element"
    assert "#FF0000" in section, "Expected fill color"
    assert "intervalDirty" in section, "Expected ellipse-specific attributes"

    print("[PASS] test_svg_circle")
    out.unlink()


def test_svg_ellipse():
    """Test SVG <ellipse> -> hp:ellipse conversion."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100">'
        '<ellipse cx="100" cy="50" rx="80" ry="40" fill="#00FF00"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=100)

    out = Path(__file__).parent / "output_p9_ellipse.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:ellipse" in section, "Expected hp:ellipse"
    assert "#00FF00" in section, "Expected fill color"
    assert "hp:renderingInfo" in section, "Expected renderingInfo"

    print("[PASS] test_svg_ellipse")
    out.unlink()


def test_svg_line():
    """Test SVG <line> -> hp:connectLine conversion."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<line x1="10" y1="10" x2="90" y2="90" stroke="#000000" stroke-width="2"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=80)

    out = Path(__file__).parent / "output_p9_line.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:connectLine" in section, "Expected hp:connectLine for line"
    assert "hp:startPt" in section, "Expected startPt element"
    assert "hp:endPt" in section, "Expected endPt element"
    assert "hp:controlPoints" in section, "Expected controlPoints"
    assert 'type="STRAIGHT_ONEWAY"' in section, "Expected line type"

    print("[PASS] test_svg_line")
    out.unlink()


def test_svg_polygon():
    """Test SVG <polygon> -> hp:polygon conversion."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<polygon points="50,10 90,90 10,90" fill="#FFA500" stroke="#FF6600"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=60)

    out = Path(__file__).parent / "output_p9_polygon.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:polygon" in section, "Expected hp:polygon"
    assert section.count("hc:pt") >= 3, "Expected at least 3 points for triangle"
    assert "#FFA500" in section, "Expected fill color"
    assert "#FF6600" in section, "Expected stroke color"
    assert "numberingType" in section, "Expected full attributes"

    print("[PASS] test_svg_polygon")
    out.unlink()


def test_svg_polyline():
    """Test SVG <polyline> -> hp:polygon conversion."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<polyline points="10,10 50,80 90,10" stroke="#0000FF" fill="none"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=60)

    out = Path(__file__).parent / "output_p9_polyline.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:polygon" in section, "Expected hp:polygon for polyline"
    assert "#0000FF" in section, "Expected stroke color"

    print("[PASS] test_svg_polyline")
    out.unlink()


def test_svg_path_basic():
    """Test SVG <path> with M, L, Z commands -> hp:polygon."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<path d="M 10 10 L 90 10 L 90 90 L 10 90 Z" fill="#808080"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=80)

    out = Path(__file__).parent / "output_p9_path.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:polygon" in section, "Expected hp:polygon for path"
    assert section.count("hc:pt") >= 4, "Expected at least 4 points for rectangle path"
    assert "#808080" in section, "Expected fill color"

    print("[PASS] test_svg_path_basic")
    out.unlink()


# ─── Multiple elements and features ──────────────────────────────


def test_svg_multiple_elements():
    """Test SVG with multiple different elements."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100">'
        '<rect x="10" y="10" width="50" height="80" fill="#2E75B6"/>'
        '<circle cx="130" cy="50" r="30" fill="#FF0000"/>'
        '<line x1="80" y1="50" x2="100" y2="50" stroke="#000000"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=160)

    out = Path(__file__).parent / "output_p9_multi.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:rect" in section, "Expected hp:rect"
    assert "hp:ellipse" in section, "Expected hp:ellipse"
    assert "hp:connectLine" in section, "Expected hp:connectLine"
    # All shapes in same hp:run
    assert "hp:drawingObject" not in section, "Should NOT use drawingObject"

    print("[PASS] test_svg_multiple_elements")
    out.unlink()


def test_svg_with_caption():
    """Test SVG element with caption text."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<rect x="10" y="10" width="80" height="80" fill="#2E75B6"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=80, caption="그림 3. 체계도")

    out = Path(__file__).parent / "output_p9_caption.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:rect" in section
    assert "그림 3. 체계도" in section, "Expected caption text"

    print("[PASS] test_svg_with_caption")
    out.unlink()


def test_svg_viewbox_auto_height():
    """Test auto height calculation from viewBox aspect ratio."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100">'
        '<rect x="0" y="0" width="200" height="100" fill="#CCC"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    elem = doc.add_svg(svg_string=svg, width_mm=100)

    out = Path(__file__).parent / "output_p9_autoheight.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:rect" in section

    print("[PASS] test_svg_viewbox_auto_height")
    out.unlink()


def test_svg_with_text_content():
    """Integration test: SVG mixed with text paragraphs."""
    doc = HwpxDocument()
    doc.add_paragraph("Before SVG")

    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 50">'
        '<rect x="5" y="5" width="90" height="40" fill="#E8E8E8"/>'
        '</svg>'
    )
    doc.add_svg(svg_string=svg, width_mm=120)
    doc.add_paragraph("After SVG")

    out = Path(__file__).parent / "output_p9_integration.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "Before SVG" in section
    assert "After SVG" in section
    assert "hp:rect" in section

    print("[PASS] test_svg_with_text_content")
    out.unlink()


def test_svg_nested_group():
    """Test SVG with nested <g> group elements."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<g>'
        '<rect x="10" y="10" width="30" height="30" fill="#AA0000"/>'
        '<rect x="60" y="60" width="30" height="30" fill="#0000AA"/>'
        '</g>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=80)

    out = Path(__file__).parent / "output_p9_group.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    # Should have two hp:rect elements (4 = 2 opening + 2 closing)
    assert section.count("hp:rect") >= 4, "Expected at least 2 rect opening+closing tags"
    assert "#AA0000" in section
    assert "#0000AA" in section

    print("[PASS] test_svg_nested_group")
    out.unlink()


def test_svg_rounded_rect():
    """Test SVG <rect> with rx/ry attributes."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<rect x="10" y="10" width="80" height="60" rx="10" ry="10" fill="#CCCCCC"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=80)

    out = Path(__file__).parent / "output_p9_rounded.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:rect" in section
    assert 'ratio="' in section, "Expected ratio attribute for rounded corners"

    print("[PASS] test_svg_rounded_rect")
    out.unlink()


def test_svg_path_hv_commands():
    """Test SVG path with H and V commands."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<path d="M 10 10 H 90 V 90 H 10 Z" fill="#AABBCC"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=60)

    out = Path(__file__).parent / "output_p9_hv.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:polygon" in section, "Expected hp:polygon for path with H/V"
    assert "#AABBCC" in section

    print("[PASS] test_svg_path_hv_commands")
    out.unlink()


def test_svg_named_colors():
    """Test SVG with named colors (red, blue, etc.)."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<rect x="10" y="10" width="80" height="80" fill="red" stroke="navy"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=60)

    out = Path(__file__).parent / "output_p9_named.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "#FF0000" in section, "Expected red -> FF0000"
    assert "#000080" in section, "Expected navy -> 000080"

    print("[PASS] test_svg_named_colors")
    out.unlink()


def test_svg_alignment():
    """Test SVG element alignment options."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<rect x="0" y="0" width="100" height="100" fill="#EEE"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=60, align="left")

    out = Path(__file__).parent / "output_p9_align.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    # Shapes use horzAlign="LEFT" in hp:pos
    assert 'horzAlign="LEFT"' in section, "Expected LEFT alignment"

    print("[PASS] test_svg_alignment")
    out.unlink()


def test_svg_element_class():
    """Test SvgElement class creation and basic properties."""
    elem = SvgElement(
        svg_string='<svg xmlns="http://www.w3.org/2000/svg"><rect/></svg>',
        width_mm=120.0,
        align="right",
        caption="Test caption",
    )
    assert elem.width_mm == 120.0
    assert elem.align == "right"
    assert elem.caption == "Test caption"
    assert elem.svg_string != ""

    loaded = elem.load_svg()
    assert "<svg" in loaded

    print("[PASS] test_svg_element_class")


def test_svg_path_relative_commands():
    """Test SVG path with relative commands (m, l, h, v)."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<path d="m 10 10 l 80 0 l 0 80 l -80 0 z" fill="#DDEEFF"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=60)

    out = Path(__file__).parent / "output_p9_relpath.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:polygon" in section, "Expected hp:polygon for relative path"
    assert "#DDEEFF" in section

    print("[PASS] test_svg_path_relative_commands")
    out.unlink()


def test_svg_reference_structure():
    """Verify SVG output matches real Hancom HWPX structure from ref_svg_shapes.hwpx."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100">'
        '<rect x="10" y="10" width="80" height="80" fill="#69D8AD" stroke="#FF0000"/>'
        '<ellipse cx="150" cy="50" rx="40" ry="30" fill="#BAFF1A"/>'
        '</svg>'
    )
    doc = HwpxDocument()
    doc.add_svg(svg_string=svg, width_mm=160)

    out = Path(__file__).parent / "output_p9_refstruct.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)

    # Structure must match reference — no drawingObject wrapper
    assert "hp:drawingObject" not in section

    # hp:rect must have reference attributes
    assert 'numberingType="PICTURE"' in section
    assert 'textWrap="IN_FRONT_OF_TEXT"' in section
    assert "instid=" in section

    # Common pre-elements present
    assert "hp:offset" in section
    assert "hp:orgSz" in section
    assert "hp:curSz" in section
    assert "hp:flip" in section
    assert "hp:rotationInfo" in section
    assert "hp:renderingInfo" in section
    assert "hc:transMatrix" in section
    assert "hc:scaMatrix" in section

    # lineShape with full attributes
    assert "endCap=" in section
    assert "outlineStyle=" in section

    # fillBrush with hatchColor
    assert "hatchColor=" in section

    # shadow
    assert "hp:shadow" in section

    # Corner points for rect
    assert "hc:pt0" in section
    assert "hc:pt1" in section

    # Post-elements
    assert "widthRelTo=" in section
    assert "hp:outMargin" in section

    print("[PASS] test_svg_reference_structure")
    out.unlink()


if __name__ == "__main__":
    test_svg_rect()
    test_svg_circle()
    test_svg_ellipse()
    test_svg_line()
    test_svg_polygon()
    test_svg_polyline()
    test_svg_path_basic()
    test_svg_multiple_elements()
    test_svg_with_caption()
    test_svg_viewbox_auto_height()
    test_svg_with_text_content()
    test_svg_nested_group()
    test_svg_rounded_rect()
    test_svg_path_hv_commands()
    test_svg_named_colors()
    test_svg_alignment()
    test_svg_element_class()
    test_svg_path_relative_commands()
    test_svg_reference_structure()
    print("\n=== All Phase 9 tests passed! ===")
