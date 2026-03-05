"""Phase 4 tests: table generation."""

import zipfile
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator import HwpxDocument, GovDocumentPreset, Table, TableCell


def _read_hwpx(path):
    with zipfile.ZipFile(path, "r") as zf:
        section = zf.read("Contents/section0.xml").decode("utf-8")
        header = zf.read("Contents/header.xml").decode("utf-8")
    return section, header


def test_basic_table():
    """Test a simple 2-column table."""
    doc = HwpxDocument()
    table = doc.add_table(col_widths=[25, 145])
    table.add_header_row(["항목", "내용"])
    table.add_row(["기간", "2026년 3월 9일(월) ~ 3월 13일(금)"])
    table.add_row(["대상", "신규 지정 지역수행기관 5개소"])

    out = Path(__file__).parent / "output_p4_basic.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)

    assert "hp:tbl" in section
    assert "hp:tr" in section
    assert "hp:tc" in section
    assert "항목" in section
    assert "내용" in section
    assert "2026년 3월" in section
    assert 'rowCnt="3"' in section
    assert 'colCnt="2"' in section
    # Col widths: 25mm = 7087, 145mm = 41102
    assert "hp:colPr" in section

    print("[PASS] test_basic_table")
    out.unlink()


def test_table_with_preset():
    """Test table with GovDocumentPreset applied."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(
        font_table="맑은 고딕",
        size_table=10,
        cell_margin_lr=3,
        cell_margin_tb=2,
    ))

    table = doc.add_table(col_widths=[25, 145])
    table.add_header_row(["항목", "내용"])
    table.add_row(["기간", "2026년 3월"])

    out = Path(__file__).parent / "output_p4_preset.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    assert "맑은 고딕" in header
    # Cell margin LR 3mm = 850
    assert '850' in section
    # Cell margin TB 2mm = 567
    assert '567' in section

    print("[PASS] test_table_with_preset")
    out.unlink()


def test_header_styling():
    """Test header row background color and font color."""
    doc = HwpxDocument()
    table = doc.add_table(
        col_widths=[50, 50],
        header_bg_color="1F4E79",
        header_font_color="FFFFFF",
    )
    table.add_header_row(["Col A", "Col B"])
    table.add_row(["val1", "val2"])

    out = Path(__file__).parent / "output_p4_header.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    # Header bg color should be in borderFill
    assert "#1F4E79" in header
    # Header font color should be in charPr
    assert "#FFFFFF" in header

    print("[PASS] test_header_styling")
    out.unlink()


def test_cell_merge_colspan():
    """Test colspan merge."""
    doc = HwpxDocument()
    table = doc.add_table(col_widths=[30, 60, 60])
    table.add_header_row(["A", "B", "C"])
    table.add_merged_row([
        {"text": "병합 셀", "colspan": 2, "bold": True},
        {"text": "일반"},
    ])

    out = Path(__file__).parent / "output_p4_colspan.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    assert 'colSpan="2"' in section
    assert "병합 셀" in section

    print("[PASS] test_cell_merge_colspan")
    out.unlink()


def test_cell_merge_rowspan():
    """Test rowspan merge."""
    doc = HwpxDocument()
    table = doc.add_table(col_widths=[30, 60, 60])
    table.add_header_row(["A", "B", "C"])
    table.add_merged_row([
        {"text": "병합", "rowspan": 2},
        {"text": "행1-B"},
        {"text": "행1-C"},
    ])
    table.add_row(["행2-B", "행2-C"])  # first col covered by rowspan

    out = Path(__file__).parent / "output_p4_rowspan.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    assert 'rowSpan="2"' in section
    assert "병합" in section

    print("[PASS] test_cell_merge_rowspan")
    out.unlink()


def test_border_fills():
    """Test that borderFills are generated for header vs data cells."""
    doc = HwpxDocument()
    table = doc.add_table(col_widths=[50, 50], header_bg_color="2E75B6")
    table.add_header_row(["H1", "H2"])
    table.add_row(["D1", "D2"])

    out = Path(__file__).parent / "output_p4_bf.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    # Should have multiple borderFill entries (default + header bg + data)
    bf_count = header.count("<hh:borderFill ")
    assert bf_count >= 2, f"Expected at least 2 borderFills, got {bf_count}"
    # Header cell bg
    assert "#2E75B6" in header

    print("[PASS] test_border_fills")
    out.unlink()


def test_keep_together():
    """Test keep-together table property."""
    doc = HwpxDocument()
    table = doc.add_table(col_widths=[50, 50], keep_together=True)
    table.add_row(["a", "b"])

    out = Path(__file__).parent / "output_p4_keep.hwpx"
    doc.save(str(out))

    section, _ = _read_hwpx(out)
    assert 'pageBreak="NONE"' in section

    # Test keep_together=False
    doc2 = HwpxDocument()
    table2 = doc2.add_table(col_widths=[50, 50], keep_together=False)
    table2.add_row(["a", "b"])

    out2 = Path(__file__).parent / "output_p4_nokeep.hwpx"
    doc2.save(str(out2))

    section2, _ = _read_hwpx(out2)
    assert 'pageBreak="CELL"' in section2

    print("[PASS] test_keep_together")
    out.unlink()
    out2.unlink()


def test_full_document_with_table():
    """Integration test: document with text and tables."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(
        font_body="휴먼명조",
        font_table="맑은 고딕",
        size_title=15,
        size_body=13,
        size_table=10,
        margin_top=10, margin_bottom=10,
        margin_left=20, margin_right=20,
        cell_margin_lr=3, cell_margin_tb=2,
    ))

    doc.add_title("2026년 모니터링 실시 안내서")
    doc.add_section_heading(1, "모니터링 개요")
    doc.add_sub_heading("기간 및 대상")

    table = doc.add_table(col_widths=[25, 145])
    table.add_header_row(["항목", "내용"])
    table.add_row(["기간", "2026년 3월 9일(월) ~ 3월 13일(금)"])
    table.add_row(["대상", "2026년 신규 지정 지역수행기관 5개소"])
    table.add_row(["수행인력", "광역지원기관 전담사회복지사"])

    doc.add_note("자체점검표는 반드시 기한 내 제출 요망")

    out = Path(__file__).parent / "output_p4_full.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)

    # Text content
    assert "2026년 모니터링 실시 안내서" in section
    assert "■ 1. 모니터링 개요" in section
    assert "○ 기간 및 대상" in section
    assert "※ 자체점검표" in section

    # Table content
    assert "hp:tbl" in section
    assert "항목" in section
    assert "광역지원기관 전담사회복지사" in section
    assert 'rowCnt="4"' in section  # header + 3 data rows

    # Fonts
    assert "휴먼명조" in header
    assert "맑은 고딕" in header

    print("[PASS] test_full_document_with_table")
    out.unlink()


def test_custom_bg_colors():
    """Test rows with custom background colors."""
    doc = HwpxDocument()
    table = doc.add_table(col_widths=[80, 80])
    table.add_header_row(["Name", "Value"])
    table.add_row(["Row 1", "val"], bg_color="F2F2F2")
    table.add_row(["Row 2", "val"])
    table.add_row(["Row 3", "val"], bg_color="F2F2F2")

    out = Path(__file__).parent / "output_p4_altrow.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    assert "#F2F2F2" in header

    print("[PASS] test_custom_bg_colors")
    out.unlink()


if __name__ == "__main__":
    test_basic_table()
    test_table_with_preset()
    test_header_styling()
    test_cell_merge_colspan()
    test_cell_merge_rowspan()
    test_border_fills()
    test_keep_together()
    test_full_document_with_table()
    test_custom_bg_colors()
    print("\n=== All Phase 4 tests passed! ===")
