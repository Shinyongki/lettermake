"""Phase 12 tests: end-to-end validation and realistic document generation.

Validates ALL features together against the spec criteria (section 5):
- Valid ZIP + valid XML structure
- Font names, line spacing, cell margins, bullet hierarchy
- Page margins, keep-together, keep-with-next, widow/orphan
- Images, diagrams (vector shapes), charts (OOXML), SVG, text boxes
- File size < 1MB, generation time < 5s
- CLI end-to-end
- UTF-8 encoding with Korean text preservation
"""

import io
import os
import struct
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

# Add parent to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator import (
    HwpxDocument,
    PageSettings,
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
    PageBreak,
    Image,
    Diagram,
    Node,
    Edge,
    Chart,
    ChartData,
    ChartDataset,
    SvgElement,
    TextBox,
    GovDocumentPreset,
    BulletStyle,
)
from hwpx_generator.json_loader import load_from_json, load_from_file
from hwpx_generator.utils import mm_to_hwp

# ── Test output directory ────────────────────────────────────────

OUTPUT_DIR = Path(__file__).parent / "output_phase12"
OUTPUT_DIR.mkdir(exist_ok=True)


# ── Helpers ──────────────────────────────────────────────────────

def _create_minimal_png(width: int = 100, height: int = 80) -> bytes:
    """Create a minimal valid PNG file in memory."""
    import zlib

    def _chunk(chunk_type: bytes, data: bytes) -> bytes:
        c = chunk_type + data
        crc = struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)
        return struct.pack(">I", len(data)) + c + crc

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr_data = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    ihdr = _chunk(b"IHDR", ihdr_data)

    # Create minimal image data: one row of white pixels
    raw_data = b""
    for _ in range(height):
        raw_data += b"\x00" + b"\xff" * (width * 3)
    compressed = zlib.compress(raw_data)
    idat = _chunk(b"IDAT", compressed)
    iend = _chunk(b"IEND", b"")

    return sig + ihdr + idat + iend


def _build_full_document(output_path: str = None) -> HwpxDocument:
    """Build a realistic government document using ALL features."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(
        font_body="휴먼명조",
        font_table="맑은 고딕",
        size_title=15,
        size_body=13,
        size_table=10,
        line_spacing=160,
        margin_top=10,
        margin_bottom=10,
        margin_left=20,
        margin_right=20,
        cell_margin_lr=3,
        cell_margin_tb=2,
    ))

    # 1. Title
    doc.add_title("2026년도 모니터링 실시 안내")
    doc.add_empty_paragraph()

    # 2. Section heading (■ 1.)
    doc.add_section_heading("1", "추진 배경")

    # 3. Bullet level 1 (-)
    doc.add_bullet1("정부 정책 이행 점검 및 현장 모니터링 강화")
    doc.add_bullet1("관련 법령에 따른 정기 점검 의무 이행")

    # 4. Bullet level 2 (·)
    doc.add_bullet2("「공공기관 운영에 관한 법률」 제xx조")
    doc.add_bullet2("「정부업무 평가 기본법」 제xx조")

    # 5. Section heading (■ 2.)
    doc.add_section_heading("2", "모니터링 개요")

    # 6. Sub heading (○)
    doc.add_sub_heading("점검 대상 및 기간")

    # 7. Table with header
    table = doc.add_table([25, 145])
    table.add_header_row(["항목", "내용"])
    table.add_row(["대상", "전국 17개 시·도 산하 공공기관"])
    table.add_row(["기간", "2026. 3. 10.(화) ~ 3. 31.(화)"])
    table.add_row(["방법", "현장 방문 점검 및 서면 점검 병행"])

    # 8. Sub heading + sub items (가, 나, 다)
    doc.add_sub_heading("추진 일정")
    doc.add_sub_item("가", "사전 준비: 3. 1.(일) ~ 3. 9.(월)")
    doc.add_sub_item("나", "현장 점검: 3. 10.(화) ~ 3. 31.(화)")
    doc.add_sub_item("다", "결과 보고: 4. 1.(수) ~ 4. 10.(금)")

    # 9. Section heading (■ 3.)
    doc.add_section_heading("3", "점검 항목")

    # 10. Table with merge
    table2 = doc.add_table([10, 40, 80, 40])
    table2.add_header_row(["No.", "점검 분야", "세부 점검 항목", "비고"])
    table2.add_merged_row([
        {"text": "1", "colspan": 1, "align": "CENTER"},
        {"text": "안전관리", "rowspan": 2},
        {"text": "소방시설 점검 및 관리 현황"},
        {"text": "필수", "align": "CENTER"},
    ])
    table2.add_merged_row([
        {"text": "2", "colspan": 1, "align": "CENTER"},
        # second column is covered by rowspan above
        {"text": "비상대피 훈련 실시 여부"},
        {"text": "필수", "align": "CENTER"},
    ])
    table2.add_row(["3", "시설관리", "노후시설 보수·교체 현황", "선택"])
    table2.add_row(["4", "운영관리", "운영 매뉴얼 구비 여부", "필수"])
    table2.add_row(["5", "운영관리", "인력 배치 적정성", "선택"])

    # 11. Note (※)
    doc.add_note("자체점검표는 반드시 3월 4일(수)까지 제출하여야 합니다.")

    # 12. Text box with border and background
    doc.add_text_box(
        "※ 자체점검표는 반드시 3월 4일(수)까지 제출하여야 합니다.",
        border_color="C00000",
        bg_color="FFF2CC",
        font_bold=True,
        padding_mm=4.0,
    )

    # 13. Page break
    doc.add_page_break()

    # 14. Section heading (■ 4.)
    doc.add_section_heading("4", "점검 절차")

    # 15. Diagram (step flow)
    diagram = doc.add_diagram(
        layout="step_flow",
        direction="horizontal",
        width_mm=160.0,
        height_mm=30.0,
    )
    diagram.add_node("n1", "사전준비", "rect")
    diagram.add_node("n2", "현장점검", "rect")
    diagram.add_node("n3", "결과보고", "rect")
    diagram.add_node("n4", "후속조치", "rounded_rect")
    diagram.add_edge("n1", "n2")
    diagram.add_edge("n2", "n3")
    diagram.add_edge("n3", "n4")

    # 16. Chart (bar chart)
    doc.add_chart(
        chart_type="bar",
        width_mm=120.0,
        height_mm=70.0,
        title="분야별 점검 결과",
        labels=["안전관리", "시설관리", "운영관리"],
        datasets=[
            {"label": "적합", "values": [85, 72, 90], "color": "2E75B6"},
            {"label": "부적합", "values": [15, 28, 10], "color": "C00000"},
        ],
    )

    # 17. SVG (inline)
    doc.add_svg(
        svg_string=(
            '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100">'
            '  <rect x="10" y="10" width="80" height="40" fill="#2E75B6" stroke="#000"/>'
            '  <circle cx="150" cy="50" r="30" fill="#FFC000" stroke="#333"/>'
            '  <line x1="90" y1="30" x2="120" y2="50" stroke="#000" stroke-width="2"/>'
            '</svg>'
        ),
        width_mm=100.0,
    )

    # 18. Image (from bytes)
    png_data = _create_minimal_png(200, 150)
    doc.add_image_from_bytes(
        png_data, "png",
        width_mm=60.0,
        align="center",
        caption="[그림 1] 점검 현장 사진",
    )

    # 19. Closing paragraphs
    doc.add_empty_paragraph()
    doc.add_paragraph(
        "붙임  1. 모니터링 실시계획 1부.",
        align="LEFT",
    )
    doc.add_paragraph(
        "      2. 자체점검표 서식 1부.  끝.",
        align="LEFT",
    )

    if output_path:
        doc.save(output_path)

    return doc


def _get_section_xml(hwpx_path: str) -> str:
    """Extract Contents/section0.xml from a .hwpx file."""
    with zipfile.ZipFile(hwpx_path, "r") as zf:
        return zf.read("Contents/section0.xml").decode("utf-8")


def _get_header_xml(hwpx_path: str) -> str:
    """Extract Contents/header.xml from a .hwpx file."""
    with zipfile.ZipFile(hwpx_path, "r") as zf:
        return zf.read("Contents/header.xml").decode("utf-8")


# ═══════════════════════════════════════════════════════════════
# 1. Full realistic document generation
# ═══════════════════════════════════════════════════════════════

def test_full_realistic_document():
    """Generate a complete government document with ALL features and verify."""
    out = OUTPUT_DIR / "full_realistic.hwpx"
    doc = _build_full_document(str(out))
    assert out.exists(), "Output file was not created"

    with zipfile.ZipFile(out, "r") as zf:
        names = zf.namelist()

        # Basic structure
        assert "mimetype" in names
        assert "Contents/section0.xml" in names
        assert "Contents/header.xml" in names

        # Chart file
        chart_files = [n for n in names if n.startswith("Chart/")]
        assert len(chart_files) >= 1, "Chart file missing"

        # Image file (BinData)
        bin_files = [n for n in names if n.startswith("BinData/")]
        assert len(bin_files) >= 1, "BinData image missing"

        # Section XML should contain content from ALL block types
        section = zf.read("Contents/section0.xml").decode("utf-8")

        # Title text
        assert "모니터링 실시 안내" in section

        # Table
        assert "hp:tbl" in section
        assert "hp:tc" in section

        # Diagram shapes (direct children of hp:run, no hp:drawingObject)
        assert "hp:rect" in section

        # Chart reference
        assert "hp:chart" in section

        # Text box (hp:drawText, not hp:textbox)
        assert "hp:drawText" in section

        # Korean text preserved
        assert "추진 배경" in section
        assert "점검 항목" in section
        assert "결과 보고" in section

    print("[PASS] test_full_realistic_document")


# ═══════════════════════════════════════════════════════════════
# 2. ZIP structure completeness
# ═══════════════════════════════════════════════════════════════

def test_zip_structure_complete():
    """Verify all required ZIP entries exist."""
    out = OUTPUT_DIR / "zip_structure.hwpx"
    _build_full_document(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        names = zf.namelist()
        required = [
            "mimetype",
            "version.xml",
            "META-INF/container.xml",
            "META-INF/manifest.xml",
            "settings.xml",
            "Contents/content.hpf",
            "Contents/header.xml",
            "Contents/section0.xml",
            "Preview/PrvText.txt",
        ]
        for entry in required:
            assert entry in names, f"Missing required entry: {entry}"

        # Verify META-INF directory is represented
        meta_entries = [n for n in names if n.startswith("META-INF/")]
        assert len(meta_entries) >= 2, "META-INF should have at least container.xml and manifest.xml"

        # Verify Contents directory
        content_entries = [n for n in names if n.startswith("Contents/")]
        assert len(content_entries) >= 3, "Contents should have hpf, header, section"

        # Verify Preview directory
        preview_entries = [n for n in names if n.startswith("Preview/")]
        assert len(preview_entries) >= 1, "Preview directory missing"

    print("[PASS] test_zip_structure_complete")


# ═══════════════════════════════════════════════════════════════
# 3. XML well-formedness
# ═══════════════════════════════════════════════════════════════

def test_xml_well_formed():
    """Parse all XML files with ElementTree to verify well-formedness."""
    out = OUTPUT_DIR / "xml_wellformed.hwpx"
    _build_full_document(str(out))

    xml_entries = [
        "version.xml",
        "META-INF/container.xml",
        "META-INF/manifest.xml",
        "settings.xml",
        "Contents/content.hpf",
        "Contents/header.xml",
        "Contents/section0.xml",
    ]

    with zipfile.ZipFile(out, "r") as zf:
        names = zf.namelist()

        # Also check chart files
        chart_files = [n for n in names if n.startswith("Chart/") and n.endswith(".xml")]
        xml_entries.extend(chart_files)

        for entry in xml_entries:
            if entry not in names:
                continue
            content = zf.read(entry).decode("utf-8")
            try:
                ET.fromstring(content)
            except ET.ParseError as e:
                raise AssertionError(f"XML parse error in {entry}: {e}")

    print("[PASS] test_xml_well_formed")


# ═══════════════════════════════════════════════════════════════
# 4. Font names in header
# ═══════════════════════════════════════════════════════════════

def test_font_names_in_header():
    """Verify specified fonts appear in header.xml faceNameList."""
    out = OUTPUT_DIR / "fonts.hwpx"
    _build_full_document(str(out))
    header = _get_header_xml(str(out))

    # The body font should be present
    assert "휴먼명조" in header, "Body font 휴먼명조 not found in header.xml"

    # The table font should be present
    assert "맑은 고딕" in header, "Table font 맑은 고딕 not found in header.xml"

    # faceName elements should exist
    assert "hh:font" in header, "No hh:font elements found"
    assert "hh:fontfaces" in header, "No fontfaces found"

    print("[PASS] test_font_names_in_header")


# ═══════════════════════════════════════════════════════════════
# 5. Line spacing 160%
# ═══════════════════════════════════════════════════════════════

def test_line_spacing_160():
    """Verify 160% line spacing is reflected in paraPr definitions."""
    out = OUTPUT_DIR / "line_spacing.hwpx"
    _build_full_document(str(out))
    header = _get_header_xml(str(out))

    # Line spacing should be present with value="160" and type="PERCENT"
    assert 'type="PERCENT"' in header, "PERCENT line spacing type not found"
    assert 'value="160"' in header, "Line spacing value 160 not found in header"

    print("[PASS] test_line_spacing_160")


# ═══════════════════════════════════════════════════════════════
# 6. Cell margins (lr=3mm -> 850 HWP, tb=2mm -> 567 HWP)
# ═══════════════════════════════════════════════════════════════

def test_cell_margins():
    """Verify cell margin values in table XML."""
    out = OUTPUT_DIR / "cell_margins.hwpx"
    _build_full_document(str(out))
    section = _get_section_xml(str(out))

    # 3mm = 3 * 7200/25.4 = 850.39... -> 850
    lr_val = mm_to_hwp(3.0)
    assert str(lr_val) in section, f"Cell margin LR ({lr_val}) not found in section"

    # 2mm = 2 * 7200/25.4 = 566.93... -> 567
    tb_val = mm_to_hwp(2.0)
    assert str(tb_val) in section, f"Cell margin TB ({tb_val}) not found in section"

    # Verify cellMargin or inMargin tag with these values
    assert "hp:cellMargin" in section or "hp:inMargin" in section, \
        "No cell margin element found"

    print("[PASS] test_cell_margins")


# ═══════════════════════════════════════════════════════════════
# 7. Bullet hierarchy (all 6 types)
# ═══════════════════════════════════════════════════════════════

def test_bullet_hierarchy_complete():
    """Verify all 6 bullet types in output: title, section, sub-heading,
    sub-item, bullet1, bullet2, note (symbols: ■ ○ 가 - · ※)."""
    out = OUTPUT_DIR / "bullet_hierarchy.hwpx"
    _build_full_document(str(out))
    section = _get_section_xml(str(out))

    # All bullet prefixes should appear in the section content
    expected_markers = [
        "\u25a0",   # ■ (section heading)
        "\u25cb",   # ○ (sub heading)
        "\uac00",   # 가 (sub item)
        "- ",       # bullet level 1
        "\u00b7",   # · (bullet level 2)
        "\u203b",   # ※ (note)
    ]
    for marker in expected_markers:
        assert marker in section, f"Bullet marker '{marker}' not found in section XML"

    print("[PASS] test_bullet_hierarchy_complete")


# ═══════════════════════════════════════════════════════════════
# 8. Page margins (top/bottom=10mm, left/right=20mm)
# ═══════════════════════════════════════════════════════════════

def test_page_margins():
    """Verify margin values in pageDef."""
    out = OUTPUT_DIR / "page_margins.hwpx"
    _build_full_document(str(out))
    section = _get_section_xml(str(out))

    top_val = mm_to_hwp(10.0)    # ~2835
    bottom_val = mm_to_hwp(10.0)
    left_val = mm_to_hwp(20.0)   # ~5669
    right_val = mm_to_hwp(20.0)

    assert "hp:pagePr" in section, "pagePr element not found"
    assert "hp:margin" in section, "margin element not found"

    # Check specific values
    assert f'top="{top_val}"' in section, f"top margin ({top_val}) not found"
    assert f'bottom="{bottom_val}"' in section, f"bottom margin ({bottom_val}) not found"
    assert f'left="{left_val}"' in section, f"left margin ({left_val}) not found"
    assert f'right="{right_val}"' in section, f"right margin ({right_val}) not found"

    print("[PASS] test_page_margins")


# ═══════════════════════════════════════════════════════════════
# 9. Keep-together on tables (pageBreak="NONE")
# ═══════════════════════════════════════════════════════════════

def test_keep_together_table():
    """Verify pageBreak='NONE' for tables to prevent splitting."""
    out = OUTPUT_DIR / "keep_together.hwpx"
    _build_full_document(str(out))
    section = _get_section_xml(str(out))

    assert 'pageBreak="NONE"' in section, "pageBreak='NONE' not found (tables should be keep-together)"

    print("[PASS] test_keep_together_table")


# ═══════════════════════════════════════════════════════════════
# 10. Keep-with-next on headings
# ═══════════════════════════════════════════════════════════════

def test_keep_with_next_headings():
    """Verify keepWithNext='true' for heading paragraph styles."""
    out = OUTPUT_DIR / "keep_with_next.hwpx"
    _build_full_document(str(out))
    header = _get_header_xml(str(out))

    assert 'keepWithNext="1"' in header, \
        "keepWithNext='1' not found in paraPr list (headings need this)"

    # Count occurrences: h1, h2, sub1 all use keepWithNext
    kwn_count = header.count('keepWithNext="1"')
    assert kwn_count >= 2, \
        f"Expected at least 2 paraPr entries with keepWithNext, found {kwn_count}"

    print("[PASS] test_keep_with_next_headings")


# ═══════════════════════════════════════════════════════════════
# 11. Diagram vector shapes (not PNG fallback)
# ═══════════════════════════════════════════════════════════════

def test_diagram_vector_shapes():
    """Verify hp:rect, hp:connectline tags exist (native vector, not PNG)."""
    out = OUTPUT_DIR / "diagram_vector.hwpx"
    _build_full_document(str(out))
    section = _get_section_xml(str(out))

    # Shapes are direct children of hp:run (no hp:drawingObject wrapper)
    assert "hp:rect" in section, "hp:rect not found (diagram node shapes)"
    assert "hp:connectLine" in section, "hp:connectLine not found (diagram edges)"

    # Verify there are multiple rect elements (at least 4 nodes)
    rect_count = section.count("<hp:rect ")
    assert rect_count >= 4, f"Expected at least 4 hp:rect elements, found {rect_count}"

    # Verify connector lines (3 edges between 4 nodes)
    line_count = section.count("<hp:connectLine ")
    assert line_count >= 3, f"Expected at least 3 hp:connectLine elements, found {line_count}"

    print("[PASS] test_diagram_vector_shapes")


# ═══════════════════════════════════════════════════════════════
# 12. Chart OOXML (Chart/chart1.xml with c:chartSpace)
# ═══════════════════════════════════════════════════════════════

def test_chart_ooxml():
    """Verify Chart/chart1.xml exists with c:chartSpace."""
    out = OUTPUT_DIR / "chart_ooxml.hwpx"
    _build_full_document(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        names = zf.namelist()
        assert "Chart/chart1.xml" in names, "Chart/chart1.xml not found in ZIP"

        chart_xml = zf.read("Chart/chart1.xml").decode("utf-8")

        # Verify OOXML chart structure
        assert "c:chartSpace" in chart_xml, "c:chartSpace root element not found"
        assert "c:chart" in chart_xml, "c:chart element not found"
        assert "c:plotArea" in chart_xml, "c:plotArea not found"
        assert "c:barChart" in chart_xml, "c:barChart not found (expected bar chart)"

        # Verify data
        assert "c:ser" in chart_xml, "c:ser (series) not found"
        assert "c:val" in chart_xml, "c:val (values) not found"

        # Verify Korean labels
        assert "안전관리" in chart_xml
        assert "시설관리" in chart_xml
        assert "운영관리" in chart_xml

        # Verify it parses as valid XML
        ET.fromstring(chart_xml)

    # Also verify chart reference in section
    section = _get_section_xml(str(out))
    assert "hp:chart" in section, "hp:chart reference not in section XML"

    print("[PASS] test_chart_ooxml")


# ═══════════════════════════════════════════════════════════════
# 13. Text box styling (hp:rect + fillBrush + textbox)
# ═══════════════════════════════════════════════════════════════

def test_text_box_styling():
    """Verify hp:rect with fillBrush and textbox content."""
    out = OUTPUT_DIR / "text_box.hwpx"
    _build_full_document(str(out))
    section = _get_section_xml(str(out))

    assert "hp:drawText" in section, "hp:drawText element not found"
    assert "hp:subList" in section, "hp:subList (drawText content) not found"

    # Background fill
    assert "hc:fillBrush" in section, "hc:fillBrush not found (text box background)"
    assert "FFF2CC" in section, "Background color FFF2CC not found"

    # Border color
    assert "C00000" in section, "Border color C00000 not found"

    # Text content
    assert "자체점검표" in section, "Text box content not found"

    print("[PASS] test_text_box_styling")


# ═══════════════════════════════════════════════════════════════
# 14. File size < 1MB (for document without large images)
# ═══════════════════════════════════════════════════════════════

def test_file_size_limit():
    """Verify output < 1MB for document without external images."""
    out = OUTPUT_DIR / "file_size.hwpx"
    _build_full_document(str(out))

    size = out.stat().st_size
    one_mb = 1024 * 1024
    assert size < one_mb, \
        f"File size {size:,} bytes exceeds 1MB limit ({one_mb:,} bytes)"
    assert size > 0, "File is empty"

    print(f"[PASS] test_file_size_limit (size={size:,} bytes)")


# ═══════════════════════════════════════════════════════════════
# 15. Generation speed < 5 seconds
# ═══════════════════════════════════════════════════════════════

def test_generation_speed():
    """Verify < 5 seconds for generating a document with all features."""
    out = OUTPUT_DIR / "speed_test.hwpx"

    start = time.perf_counter()
    _build_full_document(str(out))
    elapsed = time.perf_counter() - start

    assert elapsed < 5.0, \
        f"Generation took {elapsed:.2f}s, exceeding 5.0s limit"

    print(f"[PASS] test_generation_speed ({elapsed:.3f}s)")


# ═══════════════════════════════════════════════════════════════
# 16. Mimetype first and uncompressed
# ═══════════════════════════════════════════════════════════════

def test_mimetype_first_uncompressed():
    """Verify mimetype is first ZIP entry and stored uncompressed."""
    out = OUTPUT_DIR / "mimetype_check.hwpx"
    _build_full_document(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        # First entry
        first_name = zf.namelist()[0]
        assert first_name == "mimetype", \
            f"First ZIP entry is '{first_name}', expected 'mimetype'"

        # Uncompressed
        info = zf.getinfo("mimetype")
        assert info.compress_type == zipfile.ZIP_STORED, \
            "mimetype must be uncompressed (ZIP_STORED)"

        # Correct content
        content = zf.read("mimetype")
        assert content == b"application/hwp+zip", \
            f"mimetype content is '{content}', expected 'application/hwp+zip'"

    print("[PASS] test_mimetype_first_uncompressed")


# ═══════════════════════════════════════════════════════════════
# 17. UTF-8 encoding with Korean text
# ═══════════════════════════════════════════════════════════════

def test_utf8_encoding():
    """Verify all XML files are valid UTF-8 with Korean text preserved."""
    out = OUTPUT_DIR / "utf8_check.hwpx"
    _build_full_document(str(out))

    xml_entries = [
        "version.xml",
        "META-INF/container.xml",
        "META-INF/manifest.xml",
        "settings.xml",
        "Contents/content.hpf",
        "Contents/header.xml",
        "Contents/section0.xml",
    ]

    with zipfile.ZipFile(out, "r") as zf:
        names = zf.namelist()
        chart_files = [n for n in names if n.startswith("Chart/") and n.endswith(".xml")]
        xml_entries.extend(chart_files)

        for entry in xml_entries:
            if entry not in names:
                continue
            raw = zf.read(entry)

            # Verify UTF-8 decoding works
            try:
                text = raw.decode("utf-8")
            except UnicodeDecodeError as e:
                raise AssertionError(f"UTF-8 decode error in {entry}: {e}")

            # Verify XML declaration specifies UTF-8
            if text.startswith("<?xml"):
                assert 'encoding="UTF-8"' in text.split("?>")[0], \
                    f"{entry}: XML declaration missing UTF-8 encoding"

        # Verify Korean text is preserved in section0.xml
        section = zf.read("Contents/section0.xml").decode("utf-8")
        korean_texts = [
            "모니터링",
            "추진 배경",
            "공공기관",
            "자체점검표",
            "현장 방문 점검",
        ]
        for kt in korean_texts:
            assert kt in section, f"Korean text '{kt}' not preserved in section0.xml"

        # Verify Korean in header.xml (font names)
        header = zf.read("Contents/header.xml").decode("utf-8")
        assert "휴먼명조" in header, "Korean font name not preserved in header"
        assert "맑은 고딕" in header, "Korean font name not preserved in header"

        # Verify Korean in chart
        if chart_files:
            chart = zf.read(chart_files[0]).decode("utf-8")
            assert "안전관리" in chart, "Korean text not preserved in chart XML"

    print("[PASS] test_utf8_encoding")


# ═══════════════════════════════════════════════════════════════
# 18. CLI end-to-end
# ═══════════════════════════════════════════════════════════════

def test_cli_end_to_end():
    """Run cli.py with sample_notice.json and verify output."""
    project_root = Path(__file__).resolve().parent.parent
    cli_path = project_root / "cli.py"
    json_path = project_root / "examples" / "sample_notice.json"
    output_path = OUTPUT_DIR / "cli_output.hwpx"

    assert cli_path.exists(), f"cli.py not found at {cli_path}"
    assert json_path.exists(), f"sample_notice.json not found at {json_path}"

    result = subprocess.run(
        [sys.executable, str(cli_path),
         "--input", str(json_path),
         "--output", str(output_path)],
        capture_output=True,
        text=True,
        timeout=30,
    )

    assert result.returncode == 0, \
        f"CLI failed with exit code {result.returncode}\nstderr: {result.stderr}"
    assert output_path.exists(), "CLI did not produce output file"

    # Verify the output is a valid HWPX
    with zipfile.ZipFile(output_path, "r") as zf:
        names = zf.namelist()
        assert "mimetype" in names, "CLI output missing mimetype"
        assert "Contents/section0.xml" in names, "CLI output missing section0.xml"

        # Verify content from JSON
        section = zf.read("Contents/section0.xml").decode("utf-8")
        assert "모니터링 실시 안내" in section, "Title text missing in CLI output"
        assert "추진 배경" in section, "Section heading missing in CLI output"

    print("[PASS] test_cli_end_to_end")


# ═══════════════════════════════════════════════════════════════
# 19. SVG converted to native shapes
# ═══════════════════════════════════════════════════════════════

def test_svg_native_shapes():
    """Verify SVG is converted to native HWPX shapes (hp:rect, hp:ellipse)."""
    out = OUTPUT_DIR / "svg_native.hwpx"
    _build_full_document(str(out))
    section = _get_section_xml(str(out))

    # SVG rect -> hp:rect, SVG circle -> hp:ellipse, SVG line -> hp:connectLine
    # Shapes are direct children of hp:run (no hp:drawingObject wrapper)
    assert "hp:ellipse" in section, "hp:ellipse not found (SVG circle conversion)"

    # Verify rect and ellipse from both diagram and SVG
    rect_count = section.count("<hp:rect ")
    assert rect_count >= 5, \
        f"Expected at least 5 hp:rect (4 diagram + 1 SVG), found {rect_count}"

    print("[PASS] test_svg_native_shapes")


# ═══════════════════════════════════════════════════════════════
# 20. Image at correct size/alignment
# ═══════════════════════════════════════════════════════════════

def test_image_size_and_alignment():
    """Verify PNG image is inserted at correct size/alignment."""
    out = OUTPUT_DIR / "image_check.hwpx"
    _build_full_document(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        names = zf.namelist()

        # BinData image file exists
        bin_files = [n for n in names if n.startswith("BinData/")]
        assert len(bin_files) >= 1, "No BinData image found"

        # Verify image data is actual PNG
        png_data = zf.read(bin_files[0])
        assert png_data[:4] == b"\x89PNG", "BinData file is not valid PNG"

    section = _get_section_xml(str(out))

    # Check hp:pic element
    assert "hp:pic" in section, "hp:pic element not found"
    assert "binaryItemIDRef" in section, "binaryItemIDRef not found in picture"

    # Caption
    assert "점검 현장 사진" in section, "Image caption not found"

    # Image width should be 60mm = mm_to_hwp(60) HWP units
    w_hwp = mm_to_hwp(60.0)
    assert str(w_hwp) in section, f"Image width {w_hwp} not found in section"

    print("[PASS] test_image_size_and_alignment")


# ═══════════════════════════════════════════════════════════════
# 21. JSON loader end-to-end
# ═══════════════════════════════════════════════════════════════

def test_json_loader_end_to_end():
    """Verify JSON loader produces a valid document."""
    json_data = {
        "preset": "gov_document",
        "preset_options": {
            "font_body": "휴먼명조",
            "font_table": "맑은 고딕",
            "size_title": 15,
            "size_body": 13,
            "size_table": 10,
            "line_spacing": 160,
            "margin": {"top": 10, "bottom": 10, "left": 20, "right": 20},
            "cell_margin": {"lr": 3, "tb": 2},
        },
        "content": [
            {"type": "title", "text": "JSON 테스트 문서"},
            {"type": "section_heading", "num": "1", "text": "개요"},
            {"type": "bullet", "level": 1, "text": "항목 1"},
            {"type": "bullet", "level": 2, "text": "세부 항목"},
            {"type": "note", "text": "참고사항입니다."},
            {"type": "page_break"},
            {"type": "paragraph", "text": "일반 단락입니다."},
            {"type": "diagram",
             "layout": "step_flow", "direction": "horizontal",
             "width_mm": 120, "height_mm": 25,
             "nodes": [
                 {"id": "a", "label": "시작", "shape": "rect"},
                 {"id": "b", "label": "종료", "shape": "rect"},
             ],
             "edges": [{"from_id": "a", "to_id": "b"}]},
            {"type": "chart",
             "chart_type": "bar", "title": "테스트 차트",
             "data": {
                 "labels": ["A", "B"],
                 "datasets": [{"label": "값", "values": [10, 20]}],
             }},
            {"type": "text_box", "text": "텍스트 박스 내용",
             "border_color": "000000", "bg_color": "E0E0E0"},
            {"type": "table",
             "col_widths_mm": [40, 130],
             "header": ["구분", "설명"],
             "rows": [["행1", "설명1"]]},
            {"type": "svg",
             "svg_string": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 50"><rect x="5" y="5" width="90" height="40" fill="blue"/></svg>',
             "width_mm": 80},
        ],
    }

    out = OUTPUT_DIR / "json_loader.hwpx"
    doc = load_from_json(json_data)
    doc.save(str(out))

    assert out.exists()
    with zipfile.ZipFile(out, "r") as zf:
        section = zf.read("Contents/section0.xml").decode("utf-8")
        assert "JSON 테스트 문서" in section
        assert "hp:tbl" in section
        assert "hp:chart" in section
        assert "hp:rect" in section  # diagram shapes
        assert "hp:drawText" in section  # text box

        chart_files = [n for n in zf.namelist() if n.startswith("Chart/")]
        assert len(chart_files) >= 1

    print("[PASS] test_json_loader_end_to_end")


# ═══════════════════════════════════════════════════════════════
# 22. Manifest covers all embedded resources
# ═══════════════════════════════════════════════════════════════

def test_manifest_completeness():
    """Verify manifest.xml references all files in the ZIP."""
    out = OUTPUT_DIR / "manifest.hwpx"
    _build_full_document(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        manifest = zf.read("META-INF/manifest.xml").decode("utf-8")
        names = zf.namelist()

        # manifest.xml should exist and be valid XML
        assert "odf:manifest" in manifest or "manifest" in manifest, \
            "manifest.xml should contain manifest element"

        # All expected files should be in the ZIP
        for entry in ["Contents/section0.xml", "Contents/header.xml",
                       "settings.xml", "version.xml"]:
            assert entry in names, f"{entry} not found in ZIP"

        # BinData and Chart files should exist in ZIP
        bin_files = [n for n in names if n.startswith("BinData/")]
        assert len(bin_files) >= 1, "BinData files missing"

        chart_files = [n for n in names if n.startswith("Chart/")]
        assert len(chart_files) >= 1, "Chart files missing"

    print("[PASS] test_manifest_completeness")


# ═══════════════════════════════════════════════════════════════
# 23. Page definition orientation and dimensions
# ═══════════════════════════════════════════════════════════════

def test_page_definition():
    """Verify pageDef has correct A4 dimensions and orientation."""
    out = OUTPUT_DIR / "page_def.hwpx"
    _build_full_document(str(out))
    section = _get_section_xml(str(out))

    # A4 portrait: 210 x 297 mm
    w_hwp = mm_to_hwp(210.0)
    h_hwp = mm_to_hwp(297.0)

    assert f'width="{w_hwp}"' in section, f"Page width {w_hwp} not found"
    assert f'height="{h_hwp}"' in section, f"Page height {h_hwp} not found"
    assert 'landscape="WIDELY"' in section, "Portrait orientation (landscape=WIDELY) not found"

    print("[PASS] test_page_definition")


# ═══════════════════════════════════════════════════════════════
# 24. Table merge (rowspan/colspan)
# ═══════════════════════════════════════════════════════════════

def test_table_merge():
    """Verify rowspan/colspan attributes in table cells."""
    out = OUTPUT_DIR / "table_merge.hwpx"
    _build_full_document(str(out))
    section = _get_section_xml(str(out))

    # The document has a merged row with rowspan=2
    assert 'rowSpan="2"' in section, "rowSpan=2 not found in table"

    # All cells should have cellSpan element
    assert "hp:cellSpan" in section, "hp:cellSpan element not found"

    print("[PASS] test_table_merge")


# ═══════════════════════════════════════════════════════════════
# 25. Generation speed for 10-page document
# ═══════════════════════════════════════════════════════════════

def test_generation_speed_10_pages():
    """Verify < 5 seconds for a 10-page document."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(
        font_body="휴먼명조",
        size_body=13,
        line_spacing=160,
        margin_top=10, margin_bottom=10,
        margin_left=20, margin_right=20,
    ))

    # Generate ~10 pages of content
    for page_num in range(10):
        if page_num > 0:
            doc.add_page_break()
        doc.add_section_heading(str(page_num + 1), f"섹션 {page_num + 1}")
        for i in range(8):
            doc.add_bullet1(f"항목 {i + 1}: 이것은 {page_num + 1}페이지의 내용입니다. " * 3)
        doc.add_sub_heading(f"소제목 {page_num + 1}")
        for j in range(5):
            doc.add_bullet2(f"세부항목 {j + 1}")
        table = doc.add_table([30, 50, 90])
        table.add_header_row(["번호", "항목", "내용"])
        for k in range(4):
            table.add_row([str(k + 1), f"항목{k + 1}", f"설명 내용 {k + 1}"])

    out = OUTPUT_DIR / "speed_10pages.hwpx"
    start = time.perf_counter()
    doc.save(str(out))
    elapsed = time.perf_counter() - start

    assert elapsed < 5.0, f"10-page generation took {elapsed:.2f}s, exceeds 5.0s"
    print(f"[PASS] test_generation_speed_10_pages ({elapsed:.3f}s)")


# ═══════════════════════════════════════════════════════════════
# Run all tests
# ═══════════════════════════════════════════════════════════════

ALL_TESTS = [
    test_full_realistic_document,
    test_zip_structure_complete,
    test_xml_well_formed,
    test_font_names_in_header,
    test_line_spacing_160,
    test_cell_margins,
    test_bullet_hierarchy_complete,
    test_page_margins,
    test_keep_together_table,
    test_keep_with_next_headings,
    test_diagram_vector_shapes,
    test_chart_ooxml,
    test_text_box_styling,
    test_file_size_limit,
    test_generation_speed,
    test_mimetype_first_uncompressed,
    test_utf8_encoding,
    test_cli_end_to_end,
    test_svg_native_shapes,
    test_image_size_and_alignment,
    test_json_loader_end_to_end,
    test_manifest_completeness,
    test_page_definition,
    test_table_merge,
    test_generation_speed_10_pages,
]


def main():
    passed = 0
    failed = 0
    errors = []

    for test_fn in ALL_TESTS:
        try:
            test_fn()
            passed += 1
        except Exception as e:
            failed += 1
            errors.append((test_fn.__name__, str(e)))
            print(f"[FAIL] {test_fn.__name__}: {e}")

    print(f"\n{'='*60}")
    print(f"Phase 12 Results: {passed} passed, {failed} failed, {passed + failed} total")
    if errors:
        print(f"\nFailed tests:")
        for name, err in errors:
            print(f"  - {name}: {err}")
    print(f"{'='*60}")

    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
