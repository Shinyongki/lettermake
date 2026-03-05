"""Phase 3 tests: government document bullet system."""

import zipfile
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator import HwpxDocument, GovDocumentPreset, BulletStyle


def _read_hwpx(path):
    """Helper: read section0.xml and header.xml from hwpx."""
    with zipfile.ZipFile(path, "r") as zf:
        section = zf.read("Contents/section0.xml").decode("utf-8")
        header = zf.read("Contents/header.xml").decode("utf-8")
    return section, header


def test_apply_preset():
    """Test applying GovDocumentPreset."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(
        font_body="휴먼명조",
        font_table="맑은 고딕",
        size_title=15,
        size_body=13,
        line_spacing=160,
        margin_top=10, margin_bottom=10,
        margin_left=20, margin_right=20,
    ))

    assert doc.page_settings.margin_top == 10
    assert doc.page_settings.margin_left == 20
    assert doc._preset.font_body == "휴먼명조"

    doc.add_paragraph("테스트", font_name="휴먼명조", font_size_pt=13)
    out = Path(__file__).parent / "output_p3_preset.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    assert "휴먼명조" in header
    print("[PASS] test_apply_preset")
    out.unlink()


def test_title():
    """Test add_title method."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(font_body="맑은 고딕", size_title=15))
    doc.add_title("2026년 모니터링 실시 안내서")

    out = Path(__file__).parent / "output_p3_title.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    assert "2026년 모니터링 실시 안내서" in section
    assert 'horizontal="CENTER"' in header  # title is centered
    assert "1500" in header  # 15pt = 1500
    print("[PASS] test_title")
    out.unlink()


def test_section_heading():
    """Test ■ section heading."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(color_main="1F4E79"))
    doc.add_section_heading(1, "모니터링 개요")

    out = Path(__file__).parent / "output_p3_h1.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    assert "■ 1. 모니터링 개요" in section
    assert "#1F4E79" in header
    assert 'keepWithNext="1"' in header
    print("[PASS] test_section_heading")
    out.unlink()


def test_sub_heading():
    """Test ○ sub heading."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(color_sub="2E75B6"))
    doc.add_sub_heading("목적")

    out = Path(__file__).parent / "output_p3_h2.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    assert "○ 목적" in section
    assert "#2E75B6" in header
    assert 'keepWithNext="1"' in header
    print("[PASS] test_sub_heading")
    out.unlink()


def test_sub_item():
    """Test 가./나./다. sub items."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset())
    doc.add_sub_item("가", "세부항목 내용")
    doc.add_sub_item("나", "두 번째 항목")

    out = Path(__file__).parent / "output_p3_sub.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    assert "가. 세부항목 내용" in section
    assert "나. 두 번째 항목" in section
    # sub1 has indent_mm=7, hanging_mm=4 (hc:left value="1984", hc:intent value="-1134")
    assert 'value="1984"' in header   # 7mm
    assert 'value="-1134"' in header  # -4mm hanging
    print("[PASS] test_sub_item")
    out.unlink()


def test_bullet1():
    """Test - bullet level 1."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset())
    doc.add_bullet1("대상: 2026년 신규 지정 지역수행기관 5개소")

    out = Path(__file__).parent / "output_p3_li1.hwpx"
    doc.save(str(out))

    section, _ = _read_hwpx(out)
    assert "- 대상: 2026년 신규 지정 지역수행기관 5개소" in section
    print("[PASS] test_bullet1")
    out.unlink()


def test_bullet2():
    """Test · bullet level 2."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset())
    doc.add_bullet2("보충 설명 텍스트")

    out = Path(__file__).parent / "output_p3_li2.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    assert "· 보충 설명 텍스트" in section
    # li2: indent_mm=13 -> 3685 (hc:left value="3685")
    assert 'value="3685"' in header
    print("[PASS] test_bullet2")
    out.unlink()


def test_note():
    """Test ※ note."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(color_accent="C00000"))
    doc.add_note("자체점검표는 반드시 3월 4일(수)까지 제출해야 합니다.")

    out = Path(__file__).parent / "output_p3_note.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    assert "※ 자체점검표는 반드시 3월 4일(수)까지 제출해야 합니다." in section
    assert "#C00000" in header
    print("[PASS] test_note")
    out.unlink()


def test_full_hierarchy():
    """Integration test: complete bullet hierarchy."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(
        font_body="휴먼명조",
        size_title=15,
        size_body=13,
        margin_top=10, margin_bottom=10,
        margin_left=20, margin_right=20,
    ))

    doc.add_title("2026년 경상남도 노인맞춤돌봄서비스\n신규 수행기관 모니터링 실시 안내서")

    doc.add_section_heading(1, "모니터링 개요")
    doc.add_sub_heading("목적")
    doc.add_bullet1("신규 수행기관의 사업 운영 적정성을 확보")
    doc.add_bullet2("서비스 품질 향상을 위한 점검")

    doc.add_sub_heading("기간 및 대상")
    doc.add_sub_item("가", "기간: 2026년 3월 9일(월) ~ 3월 13일(금)")
    doc.add_sub_item("나", "대상: 신규 지정 5개소")

    doc.add_note("자체점검표는 반드시 기한 내 제출 요망")

    doc.add_section_heading(2, "세부 추진계획")
    doc.add_sub_heading("추진 절차")
    doc.add_bullet1("1단계: 사전 안내")
    doc.add_bullet1("2단계: 현장 점검")
    doc.add_bullet1("3단계: 결과 보고")

    out = Path(__file__).parent / "output_p3_full.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)

    # Verify all bullet types present
    assert "■ 1. 모니터링 개요" in section
    assert "■ 2. 세부 추진계획" in section
    assert "○ 목적" in section
    assert "○ 기간 및 대상" in section
    assert "가. 기간:" in section
    assert "나. 대상:" in section
    assert "- 신규 수행기관" in section
    assert "· 서비스 품질" in section
    assert "※ 자체점검표" in section

    # Verify styles in header
    assert "휴먼명조" in header
    assert "#1F4E79" in header  # main color
    assert "#2E75B6" in header  # sub color
    assert "#C00000" in header  # accent color
    assert 'keepWithNext="1"' in header  # headings keep with next

    # Count paragraphs (hp:p tags)
    p_count = section.count("<hp:p ")
    assert p_count == 15, f"Expected 15 paragraphs, got {p_count}"

    print("[PASS] test_full_hierarchy")
    out.unlink()


def test_custom_bullet_styles():
    """Test customizing bullet styles after preset."""
    doc = HwpxDocument()
    preset = GovDocumentPreset()
    preset.bullet_styles["h1"] = BulletStyle(
        prefix="▶ {num}. ",
        bold=True,
        color="FF0000",
        size=14,
        keep_with_next=True,
    )
    doc.apply_preset(preset)
    doc.add_section_heading(1, "커스텀 제목")

    out = Path(__file__).parent / "output_p3_custom.hwpx"
    doc.save(str(out))

    section, header = _read_hwpx(out)
    assert "▶ 1. 커스텀 제목" in section
    assert "#FF0000" in header
    assert "1400" in header  # 14pt
    print("[PASS] test_custom_bullet_styles")
    out.unlink()


def test_chaining():
    """Test fluent API chaining."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset())

    # Methods should return self for chaining
    result = doc.add_section_heading(1, "제목")
    assert result is doc

    result = doc.add_sub_heading("소제목")
    assert result is doc

    result = doc.add_bullet1("항목")
    assert result is doc

    print("[PASS] test_chaining")


if __name__ == "__main__":
    test_apply_preset()
    test_title()
    test_section_heading()
    test_sub_heading()
    test_sub_item()
    test_bullet1()
    test_bullet2()
    test_note()
    test_full_hierarchy()
    test_custom_bullet_styles()
    test_chaining()
    print("\n=== All Phase 3 tests passed! ===")
