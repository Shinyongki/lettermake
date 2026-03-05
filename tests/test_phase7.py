"""Phase 7 tests: native shape diagram generation."""

import zipfile
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator import HwpxDocument, GovDocumentPreset, Diagram, Node, Edge


def _read_section(path):
    with zipfile.ZipFile(path, "r") as zf:
        return zf.read("Contents/section0.xml").decode("utf-8")


def test_basic_diagram():
    """Test a basic horizontal step flow diagram."""
    doc = HwpxDocument()
    diag = doc.add_diagram(
        layout="step_flow",
        direction="horizontal",
        width_mm=160,
        height_mm=30,
        theme_color="1F4E79",
    )
    diag.add_node("s1", label="Step 1.\n준비 및 안내", shape="rect")
    diag.add_node("s2", label="Step 2.\n자체점검 회수", shape="rect")
    diag.add_edge("s1", "s2")

    out = Path(__file__).parent / "output_p7_basic.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)

    # Check for shape elements
    assert "hp:rect" in section, "Expected hp:rect in section XML"
    assert "hp:connectLine" in section, "Expected hp:connectLine in section XML"
    assert "Step 1." in section, "Expected node label 'Step 1.' in XML"
    assert "Step 2." in section, "Expected node label 'Step 2.' in XML"
    # Check line/fill attributes
    assert "#1F4E79" in section, "Expected theme color in line"
    assert "hp:lineShape" in section, "Expected lineShape element"
    assert "hc:fillBrush" in section, "Expected fillBrush element"
    # Check arrow style
    assert 'tailStyle="ARROW"' in section, "Expected arrow tail style"

    print("[PASS] test_basic_diagram")
    out.unlink()


def test_vertical_diagram():
    """Test vertical step flow diagram."""
    doc = HwpxDocument()
    diag = doc.add_diagram(
        direction="vertical",
        width_mm=60,
        height_mm=120,
    )
    diag.add_node("a", label="Start", shape="ellipse")
    diag.add_node("b", label="Process", shape="rect")
    diag.add_node("c", label="End", shape="ellipse")
    diag.add_edge("a", "b")
    diag.add_edge("b", "c")

    out = Path(__file__).parent / "output_p7_vertical.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:ellipse" in section, "Expected hp:ellipse"
    assert "hp:rect" in section, "Expected hp:rect"
    assert section.count("hp:connectLine") >= 2, "Expected 2 connectlines"
    assert "hc:center" in section, "Expected center element for ellipse"
    assert "hc:ax1" in section, "Expected ax1 for ellipse"

    print("[PASS] test_vertical_diagram")
    out.unlink()


def test_diamond_shape():
    """Test diamond (polygon) shape."""
    doc = HwpxDocument()
    diag = doc.add_diagram(width_mm=80, height_mm=40)
    diag.add_node("d1", label="Decision?", shape="diamond")

    out = Path(__file__).parent / "output_p7_diamond.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:polygon" in section, "Expected hp:polygon for diamond"
    assert section.count("hc:pt ") >= 5, "Expected 5 points for diamond (polygon closes path)"
    assert "Decision?" in section

    print("[PASS] test_diamond_shape")
    out.unlink()


def test_rounded_rect():
    """Test rounded rectangle shape."""
    doc = HwpxDocument()
    diag = doc.add_diagram(width_mm=80, height_mm=30)
    diag.add_node("r1", label="Rounded", shape="rounded_rect")

    out = Path(__file__).parent / "output_p7_rounded.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:rect" in section, "Expected hp:rect for rounded_rect"
    assert 'ratio="20"' in section, "Expected ratio attribute for rounded corners"

    print("[PASS] test_rounded_rect")
    out.unlink()


def test_custom_colors():
    """Test custom node/edge colors."""
    doc = HwpxDocument()
    diag = doc.add_diagram(theme_color="2E75B6")
    diag.add_node("n1", label="A", shape="rect", fill_color="FFE0B2", line_color="FF6F00")
    diag.add_node("n2", label="B", shape="rect")
    diag.add_edge("n1", "n2", line_color="C00000")

    out = Path(__file__).parent / "output_p7_colors.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "#FFE0B2" in section, "Expected custom fill color"
    assert "#FF6F00" in section, "Expected custom line color"
    assert "#C00000" in section, "Expected custom edge color"

    print("[PASS] test_custom_colors")
    out.unlink()


def test_multiple_shapes():
    """Test a diagram with mixed shape types."""
    doc = HwpxDocument()
    diag = doc.add_diagram(width_mm=160, height_mm=40)
    diag.add_node("start", label="Start", shape="ellipse")
    diag.add_node("proc", label="Process", shape="rect")
    diag.add_node("dec", label="OK?", shape="diamond")
    diag.add_node("end", label="End", shape="rounded_rect")
    diag.add_edge("start", "proc")
    diag.add_edge("proc", "dec")
    diag.add_edge("dec", "end")

    out = Path(__file__).parent / "output_p7_mixed.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:ellipse" in section
    assert "hp:rect" in section
    assert "hp:polygon" in section
    assert section.count("hp:connectLine") >= 3

    print("[PASS] test_multiple_shapes")
    out.unlink()


def test_diagram_with_preset():
    """Test diagram uses preset theme color."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(color_main="C00000"))
    diag = doc.add_diagram()
    diag.add_node("s1", label="A", shape="rect")

    out = Path(__file__).parent / "output_p7_preset.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "#C00000" in section, "Expected preset color_main in diagram"

    print("[PASS] test_diagram_with_preset")
    out.unlink()


def test_diagram_with_text():
    """Integration test: diagram mixed with text content."""
    doc = HwpxDocument()
    doc.add_paragraph("Before diagram", bold=True)

    diag = doc.add_diagram(width_mm=120, height_mm=25)
    diag.add_node("a", label="Input", shape="rect")
    diag.add_node("b", label="Output", shape="rect")
    diag.add_edge("a", "b")

    doc.add_paragraph("After diagram")

    out = Path(__file__).parent / "output_p7_integration.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "Before diagram" in section
    assert "After diagram" in section
    assert "Input" in section
    assert "Output" in section

    print("[PASS] test_diagram_with_text")
    out.unlink()


def test_single_node():
    """Test diagram with a single node (no edges)."""
    doc = HwpxDocument()
    diag = doc.add_diagram(width_mm=60, height_mm=30)
    diag.add_node("only", label="Solo", shape="ellipse")

    out = Path(__file__).parent / "output_p7_single.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:ellipse" in section
    assert "Solo" in section
    # No connectlines expected
    assert "hp:connectLine" not in section

    print("[PASS] test_single_node")
    out.unlink()


if __name__ == "__main__":
    test_basic_diagram()
    test_vertical_diagram()
    test_diamond_shape()
    test_rounded_rect()
    test_custom_colors()
    test_multiple_shapes()
    test_diagram_with_preset()
    test_diagram_with_text()
    test_single_node()
    print("\n=== All Phase 7 tests passed! ===")
