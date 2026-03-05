"""Phase 2 tests: text/paragraph style implementation."""

import zipfile
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator import HwpxDocument, Paragraph, TextRun, Style, CharProps, ParaProps


def test_single_paragraph():
    """Test a document with a single styled paragraph."""
    doc = HwpxDocument()
    doc.add_paragraph("안녕하세요, HWPX 문서입니다.", font_size_pt=13)

    out = Path(__file__).parent / "output_p2_single.hwpx"
    doc.save(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        section = zf.read("Contents/section0.xml").decode("utf-8")
        assert "안녕하세요, HWPX 문서입니다." in section
        assert "hp:p" in section
        assert "hp:run" in section
        assert "hp:t" in section

    print("[PASS] test_single_paragraph")
    out.unlink()


def test_bold_italic_underline():
    """Test character formatting: bold, italic, underline."""
    doc = HwpxDocument()

    p = doc.add_paragraph()
    p.add_run("굵은 텍스트 ", bold=True, font_size_pt=13)
    p.add_run("기울임 텍스트 ", italic=True, font_size_pt=13)
    p.add_run("밑줄 텍스트", underline=True, font_size_pt=13)

    out = Path(__file__).parent / "output_p2_format.hwpx"
    doc.save(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        header = zf.read("Contents/header.xml").decode("utf-8")
        section = zf.read("Contents/section0.xml").decode("utf-8")

        # Check bold/italic/underline charPr exists
        assert '<hh:bold/>' in header
        assert '<hh:italic/>' in header
        assert 'type="BOTTOM"' in header  # underline type

        # Check text content
        assert "굵은 텍스트" in section
        assert "기울임 텍스트" in section
        assert "밑줄 텍스트" in section

    print("[PASS] test_bold_italic_underline")
    out.unlink()


def test_font_and_color():
    """Test custom font name and color."""
    doc = HwpxDocument()
    doc.add_paragraph("제목 텍스트", font_name="맑은 고딕", font_size_pt=15, color="1F4E79")

    out = Path(__file__).parent / "output_p2_font.hwpx"
    doc.save(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        header = zf.read("Contents/header.xml").decode("utf-8")
        assert "맑은 고딕" in header
        assert "1500" in header   # 15pt = 1500
        assert "#1F4E79" in header

    print("[PASS] test_font_and_color")
    out.unlink()


def test_alignment():
    """Test paragraph alignment options."""
    doc = HwpxDocument()
    doc.add_paragraph("왼쪽 정렬", align="LEFT")
    doc.add_paragraph("가운데 정렬", align="CENTER")
    doc.add_paragraph("오른쪽 정렬", align="RIGHT")
    doc.add_paragraph("양쪽 정렬", align="JUSTIFY")

    out = Path(__file__).parent / "output_p2_align.hwpx"
    doc.save(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        header = zf.read("Contents/header.xml").decode("utf-8")
        assert 'horizontal="LEFT"' in header
        assert 'horizontal="CENTER"' in header
        assert 'horizontal="RIGHT"' in header
        assert 'horizontal="JUSTIFY"' in header

    print("[PASS] test_alignment")
    out.unlink()


def test_line_spacing():
    """Test line spacing percent and fixed."""
    doc = HwpxDocument()
    doc.add_paragraph("줄간격 160%", line_spacing_value=160)
    doc.add_paragraph("줄간격 200%", line_spacing_value=200)

    out = Path(__file__).parent / "output_p2_spacing.hwpx"
    doc.save(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        header = zf.read("Contents/header.xml").decode("utf-8")
        assert 'type="PERCENT" value="160"' in header
        assert 'type="PERCENT" value="200"' in header

    print("[PASS] test_line_spacing")
    out.unlink()


def test_indent():
    """Test left indent and hanging indent."""
    doc = HwpxDocument()
    # 7mm left indent, 4mm hanging (first line outdented)
    doc.add_paragraph(
        "들여쓰기 테스트",
        indent_left_mm=7.0,
        indent_first_mm=-4.0,  # negative = hanging
    )

    out = Path(__file__).parent / "output_p2_indent.hwpx"
    doc.save(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        header = zf.read("Contents/header.xml").decode("utf-8")
        # left=7mm -> 1984 HWP units (hc:left value="1984")
        assert 'value="1984"' in header
        # indent=-4mm -> -1134 HWP units (hc:intent value="-1134")
        assert 'value="-1134"' in header

    print("[PASS] test_indent")
    out.unlink()


def test_paragraph_spacing():
    """Test before/after paragraph spacing."""
    doc = HwpxDocument()
    doc.add_paragraph("간격 테스트", space_before_pt=6.0, space_after_pt=3.0)

    out = Path(__file__).parent / "output_p2_paraspacing.hwpx"
    doc.save(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        header = zf.read("Contents/header.xml").decode("utf-8")
        # 6pt * 100 = 600 (hc:prev value="600"), 3pt * 100 = 300 (hc:next value="300")
        assert 'value="600"' in header
        assert 'value="300"' in header

    print("[PASS] test_paragraph_spacing")
    out.unlink()


def test_multiple_runs():
    """Test paragraph with multiple differently-styled runs."""
    doc = HwpxDocument()
    p = doc.add_paragraph()
    p.add_run("일반 텍스트 ", font_size_pt=13)
    p.add_run("굵은 빨간 텍스트 ", font_size_pt=13, bold=True, color="C00000")
    p.add_run("작은 파란 텍스트", font_size_pt=10, color="2E75B6")

    out = Path(__file__).parent / "output_p2_multi.hwpx"
    doc.save(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        section = zf.read("Contents/section0.xml").decode("utf-8")
        header = zf.read("Contents/header.xml").decode("utf-8")

        # Each run should have a different charPrIDRef
        assert "일반 텍스트" in section
        assert "굵은 빨간 텍스트" in section
        assert "작은 파란 텍스트" in section

        # Different font sizes
        assert "1300" in header  # 13pt
        assert "1000" in header  # 10pt

        # Different colors
        assert "#C00000" in header
        assert "#2E75B6" in header

    print("[PASS] test_multiple_runs")
    out.unlink()


def test_named_style():
    """Test named style registration and reference."""
    doc = HwpxDocument()
    doc.register_style(Style(
        name="제목",
        char_props=CharProps(font_name="맑은 고딕", font_size_pt=15, bold=True, color="1F4E79"),
        para_props=ParaProps(align="CENTER", line_spacing_value=130),
    ))
    doc.add_paragraph("문서 제목", style_name="제목", font_name="맑은 고딕",
                      font_size_pt=15, bold=True, color="1F4E79", align="CENTER",
                      line_spacing_value=130)

    out = Path(__file__).parent / "output_p2_style.hwpx"
    doc.save(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        header = zf.read("Contents/header.xml").decode("utf-8")
        section = zf.read("Contents/section0.xml").decode("utf-8")

        assert '제목' in header  # style name in styleList
        assert 'styleIDRef=' in section

    print("[PASS] test_named_style")
    out.unlink()


def test_keep_with_next():
    """Test keep-with-next paragraph property."""
    doc = HwpxDocument()
    doc.add_paragraph("제목 (다음과 묶음)", keep_with_next=True)
    doc.add_paragraph("본문 내용")

    out = Path(__file__).parent / "output_p2_kwn.hwpx"
    doc.save(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        header = zf.read("Contents/header.xml").decode("utf-8")
        assert 'keepWithNext="1"' in header

    print("[PASS] test_keep_with_next")
    out.unlink()


def test_full_document():
    """Integration test: a realistic document with mixed styles."""
    doc = HwpxDocument()
    doc.page_settings.margin_top = 10
    doc.page_settings.margin_bottom = 10
    doc.page_settings.margin_left = 20
    doc.page_settings.margin_right = 20

    # Title
    doc.add_paragraph(
        "2026년 모니터링 실시 안내서",
        font_name="맑은 고딕", font_size_pt=15, bold=True,
        align="CENTER", line_spacing_value=130,
        space_after_pt=12,
    )

    # Section heading
    doc.add_paragraph(
        "■ 1. 모니터링 개요",
        font_size_pt=13, bold=True, color="1F4E79",
        keep_with_next=True,
    )

    # Sub heading
    doc.add_paragraph(
        "○ 목적",
        font_size_pt=13, bold=True,
        indent_left_mm=4,
    )

    # Body text
    doc.add_paragraph(
        "신규 수행기관의 사업 운영 적정성을 확보하고 서비스 품질을 향상시키기 위함",
        font_size_pt=13,
        indent_left_mm=7,
    )

    # Bullet
    doc.add_paragraph(
        "- 대상: 2026년 신규 지정 지역수행기관 5개소",
        font_size_pt=13,
        indent_left_mm=7, indent_first_mm=-4,
    )

    out = Path(__file__).parent / "output_p2_full.hwpx"
    doc.save(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        section = zf.read("Contents/section0.xml").decode("utf-8")
        header = zf.read("Contents/header.xml").decode("utf-8")

        assert "2026년 모니터링 실시 안내서" in section
        assert "■ 1. 모니터링 개요" in section
        assert "○ 목적" in section
        assert "맑은 고딕" in header
        assert "#1F4E79" in header

    print("[PASS] test_full_document")
    out.unlink()


if __name__ == "__main__":
    test_single_paragraph()
    test_bold_italic_underline()
    test_font_and_color()
    test_alignment()
    test_line_spacing()
    test_indent()
    test_paragraph_spacing()
    test_multiple_runs()
    test_named_style()
    test_keep_with_next()
    test_full_document()
    print("\n=== All Phase 2 tests passed! ===")
