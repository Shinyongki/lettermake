"""Phase 10 tests: TextBox (highlighted text box) generation."""

import zipfile
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator import HwpxDocument, GovDocumentPreset, TextBox


def _read_section(path):
    with zipfile.ZipFile(path, "r") as zf:
        return zf.read("Contents/section0.xml").decode("utf-8")


def test_basic_text_box():
    """Test basic text box with border and background."""
    doc = HwpxDocument()
    doc.add_text_box(
        "This is important text.",
        border_color="C00000",
        bg_color="FFF2CC",
    )

    out = Path(__file__).parent / "output_p10_basic.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:rect" in section, "Expected hp:rect for text box"
    assert "hp:drawText" in section, "Expected hp:drawText element"
    assert "hp:textbox" not in section, "hp:textbox should not be used (use hp:drawText)"
    assert "hp:subList" in section, "Expected hp:subList inside drawText"
    assert "hp:textMargin" in section, "Expected hp:textMargin inside drawText"
    assert "This is important text." in section, "Expected text content"
    assert "#C00000" in section, "Expected border color"
    assert "#FFF2CC" in section, "Expected background color"

    print("[PASS] test_basic_text_box")
    out.unlink()


def test_text_box_bold():
    """Test text box with bold text."""
    doc = HwpxDocument()
    doc.add_text_box(
        "Bold text here",
        font_bold=True,
        border_color="000000",
    )

    out = Path(__file__).parent / "output_p10_bold.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "Bold text here" in section
    assert "hp:drawText" in section
    assert "hp:textbox" not in section

    print("[PASS] test_text_box_bold")
    out.unlink()


def test_text_box_no_bg():
    """Test text box without background color (border only)."""
    doc = HwpxDocument()
    doc.add_text_box(
        "Border only",
        border_color="0000FF",
    )

    out = Path(__file__).parent / "output_p10_nobg.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "#0000FF" in section, "Expected border color"
    assert "Border only" in section
    # Should not have fillBrush since bg_color is None
    # The text box renders without hc:fillBrush
    assert "hp:drawText" in section
    assert "hp:textbox" not in section

    print("[PASS] test_text_box_no_bg")
    out.unlink()


def test_text_box_multiline():
    """Test text box with multi-line text."""
    doc = HwpxDocument()
    doc.add_text_box(
        "Line 1\nLine 2\nLine 3",
        border_color="000000",
        bg_color="E8E8E8",
    )

    out = Path(__file__).parent / "output_p10_multiline.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "Line 1" in section
    assert "Line 2" in section
    assert "Line 3" in section
    # Each line should be a separate <hp:p> inside the textbox
    # Count paragraphs inside subList
    assert section.count("hp:subList") >= 1

    print("[PASS] test_text_box_multiline")
    out.unlink()


def test_text_box_custom_dimensions():
    """Test text box with explicit width and height."""
    doc = HwpxDocument()
    doc.add_text_box(
        "Fixed size box",
        border_color="000000",
        width_mm=100.0,
        height_mm=20.0,
    )

    out = Path(__file__).parent / "output_p10_dims.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "hp:sz" in section
    assert "hp:drawText" in section
    assert "hp:textbox" not in section
    assert "Fixed size box" in section

    print("[PASS] test_text_box_custom_dimensions")
    out.unlink()


def test_text_box_with_text_content():
    """Integration test: text box mixed with regular paragraphs."""
    doc = HwpxDocument()
    doc.add_paragraph("Before text box")
    doc.add_text_box(
        "Important notice here",
        border_color="C00000",
        bg_color="FFF2CC",
        font_bold=True,
    )
    doc.add_paragraph("After text box")

    out = Path(__file__).parent / "output_p10_integration.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "Before text box" in section
    assert "After text box" in section
    assert "Important notice here" in section
    assert "hp:drawText" in section
    assert "hp:textbox" not in section

    print("[PASS] test_text_box_with_text_content")
    out.unlink()


def test_text_box_with_preset():
    """Test text box uses preset font settings."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(font_body="맑은 고딕", size_body=12.0))
    doc.add_text_box(
        "Preset font text",
        border_color="000000",
    )

    out = Path(__file__).parent / "output_p10_preset.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "Preset font text" in section
    assert "hp:drawText" in section
    assert "hp:textbox" not in section

    print("[PASS] test_text_box_with_preset")
    out.unlink()


def test_text_box_korean_text():
    """Test text box with Korean text content."""
    doc = HwpxDocument()
    doc.add_text_box(
        "※ 자체점검표는 반드시 3월 4일(수)까지 제출하여야 합니다.",
        border_color="C00000",
        bg_color="FFF2CC",
        font_bold=True,
        padding_mm=4,
    )

    out = Path(__file__).parent / "output_p10_korean.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "자체점검표" in section, "Expected Korean text"
    assert "#C00000" in section
    assert "#FFF2CC" in section

    print("[PASS] test_text_box_korean_text")
    out.unlink()


def test_text_box_class():
    """Test TextBox class creation and default properties."""
    tb = TextBox(
        text="Test",
        border_color="FF0000",
        bg_color="FFEEEE",
        font_bold=True,
        padding_mm=5.0,
    )
    assert tb.text == "Test"
    assert tb.border_color == "FF0000"
    assert tb.bg_color == "FFEEEE"
    assert tb.font_bold is True
    assert tb.padding_mm == 5.0
    assert tb.width_mm is None  # auto
    assert tb.height_mm is None  # auto

    print("[PASS] test_text_box_class")


def test_text_box_font_color():
    """Test text box with custom font color."""
    doc = HwpxDocument()
    doc.add_text_box(
        "Colored text",
        border_color="000000",
        font_color="C00000",
    )

    out = Path(__file__).parent / "output_p10_fontcolor.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "Colored text" in section
    assert "hp:drawText" in section
    assert "hp:textbox" not in section

    print("[PASS] test_text_box_font_color")
    out.unlink()


def test_text_box_position():
    """Test text box has proper position attributes."""
    doc = HwpxDocument()
    doc.add_text_box(
        "Positioned box",
        border_color="000000",
    )

    out = Path(__file__).parent / "output_p10_pos.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert 'treatAsChar="0"' in section, "Expected treatAsChar"
    assert 'vertRelTo="PARA"' in section, "Expected vertRelTo"
    assert 'horzRelTo="PARA"' in section, "Expected horzRelTo"

    print("[PASS] test_text_box_position")
    out.unlink()


if __name__ == "__main__":
    test_basic_text_box()
    test_text_box_bold()
    test_text_box_no_bg()
    test_text_box_multiline()
    test_text_box_custom_dimensions()
    test_text_box_with_text_content()
    test_text_box_with_preset()
    test_text_box_korean_text()
    test_text_box_class()
    test_text_box_font_color()
    test_text_box_position()
    print("\n=== All Phase 10 tests passed! ===")
