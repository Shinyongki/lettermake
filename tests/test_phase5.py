"""Phase 5 tests: page flow control (PageFlowController, PageBreak)."""

import zipfile
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator import HwpxDocument, GovDocumentPreset, PageBreak
from hwpx_generator.builders.pageflow import PageFlowController
from hwpx_generator.document import PageSettings
from hwpx_generator.elements.paragraph import Paragraph
from hwpx_generator.elements.table import Table


def _read_hwpx(path):
    with zipfile.ZipFile(path, "r") as zf:
        section = zf.read("Contents/section0.xml").decode("utf-8")
        header = zf.read("Contents/header.xml").decode("utf-8")
    return section, header


# ─── Manual page break tests ────────────────────────────────────

def test_manual_page_break():
    """Test doc.add_page_break() adds a PageBreak to blocks."""
    doc = HwpxDocument()
    doc.add_paragraph("First paragraph")
    pb = doc.add_page_break()
    doc.add_paragraph("After page break")

    assert isinstance(pb, PageBreak)
    assert isinstance(doc.blocks[1], PageBreak)
    assert len(doc.blocks) == 3

    print("[PASS] test_manual_page_break")


def test_page_break_renders_xml():
    """Test that PageBreak produces pageBreakBefore='true' in section XML."""
    doc = HwpxDocument()
    doc.add_paragraph("Page one content")
    doc.add_page_break()
    doc.add_paragraph("Page two content")

    out = Path(__file__).parent / "output_p5_pb.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section, header = _read_hwpx(out)

    assert "Page one content" in section
    assert "Page two content" in section
    # The header should contain a paraPr with pageBreakBefore="true"
    assert 'pageBreakBefore="1"' in header

    print("[PASS] test_page_break_renders_xml")
    out.unlink()


def test_page_break_import():
    """Test that PageBreak can be imported from hwpx_generator."""
    from hwpx_generator import PageBreak as PB
    pb = PB()
    assert pb is not None

    print("[PASS] test_page_break_import")


# ─── PageFlowController unit tests ──────────────────────────────

def test_height_estimation_paragraph():
    """Test paragraph height estimation."""
    ps = PageSettings()
    controller = PageFlowController(ps)

    # Simple paragraph: 13pt font, 160% line spacing, ~20 chars
    para = Paragraph(font_size_pt=13.0, line_spacing_value=160.0)
    para.add_run("Short text here.")

    height = controller.estimate_height(para)
    # line_height = 13 * 160/100 * 0.3528 = 7.338mm
    # 1 line * 7.338 + 0 + 0 = ~7.34mm
    expected_line_height = 13.0 * 160.0 / 100.0 * 0.3528
    assert abs(height - expected_line_height) < 1.0, f"Expected ~{expected_line_height}mm, got {height}mm"

    print("[PASS] test_height_estimation_paragraph")


def test_height_estimation_table():
    """Test table height estimation."""
    ps = PageSettings()
    controller = PageFlowController(ps)

    table = Table(
        col_widths_mm=[50, 50],
        font_size_pt=10.0,
        cell_margin_tb_mm=2.0,
    )
    table.add_header_row(["A", "B"])
    table.add_row(["1", "2"])
    table.add_row(["3", "4"])

    height = controller.estimate_height(table)
    # row_height = 10 * 1.6 * 0.3528 + 2*2 = 5.645 + 4 = 9.645mm
    # total = 9.645 * 3 + 2 = 30.935mm
    expected_row_h = 10.0 * 1.6 * 0.3528 + 2.0 * 2
    expected = expected_row_h * 3 + 2.0
    assert abs(height - expected) < 0.1, f"Expected ~{expected}mm, got {height}mm"

    print("[PASS] test_height_estimation_table")


def test_height_estimation_page_break():
    """Test that PageBreak has zero height."""
    ps = PageSettings()
    controller = PageFlowController(ps)

    pb = PageBreak()
    assert controller.estimate_height(pb) == 0.0

    print("[PASS] test_height_estimation_page_break")


# ─── Table keep-together tests ───────────────────────────────────

def test_table_keep_together_auto_break():
    """Test that auto page flow inserts break before a table that won't fit."""
    doc = HwpxDocument()
    ps = doc.page_settings
    controller = PageFlowController(ps)

    # Fill up most of the page with paragraphs
    blocks = []
    content_h = ps.content_height_mm

    # Create paragraphs that fill ~90% of the page
    para_font = 13.0
    line_h = para_font * 160.0 / 100.0 * 0.3528
    lines_needed = int(content_h * 0.9 / line_h)
    for _ in range(lines_needed):
        p = Paragraph(font_size_pt=para_font, line_spacing_value=160.0)
        p.add_run("X")
        blocks.append(p)

    # Add a table that needs ~30mm (won't fit in remaining ~10%)
    table = Table(col_widths_mm=[80, 80], font_size_pt=10.0, cell_margin_tb_mm=2.0)
    for r in range(5):
        table.add_row([f"r{r}c0", f"r{r}c1"])
    blocks.append(table)

    result = controller.process(blocks)

    # There should be a PageBreak inserted before the table
    page_breaks = [b for b in result if isinstance(b, PageBreak)]
    assert len(page_breaks) >= 1, "Expected at least one auto page break before table"

    # The table should come right after a PageBreak
    table_idx = None
    for idx, b in enumerate(result):
        if isinstance(b, Table):
            table_idx = idx
            break
    assert table_idx is not None
    assert isinstance(result[table_idx - 1], PageBreak), \
        "Expected PageBreak immediately before table"

    print("[PASS] test_table_keep_together_auto_break")


def test_table_no_break_when_fits():
    """Test that no break is inserted if the table fits on the current page."""
    doc = HwpxDocument()
    ps = doc.page_settings
    controller = PageFlowController(ps)

    blocks = []
    # One short paragraph
    p = Paragraph(font_size_pt=13.0, line_spacing_value=160.0)
    p.add_run("Short intro.")
    blocks.append(p)

    # Small table
    table = Table(col_widths_mm=[80, 80], font_size_pt=10.0, cell_margin_tb_mm=2.0)
    table.add_row(["a", "b"])
    blocks.append(table)

    result = controller.process(blocks)

    page_breaks = [b for b in result if isinstance(b, PageBreak)]
    assert len(page_breaks) == 0, "No page break should be inserted when table fits"

    print("[PASS] test_table_no_break_when_fits")


# ─── Keep-with-next (heading) tests ─────────────────────────────

def test_section_heading_keep_with_next():
    """Test that section heading (starts with \\u25a0) keeps with next 3 blocks."""
    ps = PageSettings()
    controller = PageFlowController(ps)

    # Fill page to near-full
    blocks = []
    content_h = ps.content_height_mm
    para_font = 13.0
    line_h = para_font * 160.0 / 100.0 * 0.3528
    lines_needed = int(content_h * 0.93 / line_h)
    for _ in range(lines_needed):
        p = Paragraph(font_size_pt=para_font, line_spacing_value=160.0)
        p.add_run("Filler")
        blocks.append(p)

    # Section heading with keep_with_next
    heading = Paragraph(
        font_size_pt=13.0, line_spacing_value=160.0,
        keep_with_next=True, space_before_pt=6.0, space_after_pt=3.0,
    )
    heading.add_run("\u25a0 1. Section Title")
    blocks.append(heading)

    # 3 following blocks
    for text in ["Sub content 1", "Sub content 2", "Sub content 3"]:
        p = Paragraph(font_size_pt=13.0, line_spacing_value=160.0)
        p.add_run(text)
        blocks.append(p)

    result = controller.process(blocks)

    # Find the heading in result
    heading_idx = None
    for idx, b in enumerate(result):
        if isinstance(b, Paragraph) and hasattr(b, 'plain_text'):
            if b.plain_text.startswith("\u25a0"):
                heading_idx = idx
                break

    assert heading_idx is not None, "Heading should be in result"
    # There should be a PageBreak before the heading
    assert heading_idx > 0 and isinstance(result[heading_idx - 1], PageBreak), \
        "Expected PageBreak before section heading"

    print("[PASS] test_section_heading_keep_with_next")


def test_sub_heading_keep_with_next():
    """Test that sub heading (starts with \\u25cb) keeps with next 1 block."""
    ps = PageSettings()
    controller = PageFlowController(ps)

    blocks = []
    content_h = ps.content_height_mm
    para_font = 13.0
    line_h = para_font * 160.0 / 100.0 * 0.3528
    lines_needed = int(content_h * 0.95 / line_h)
    for _ in range(lines_needed):
        p = Paragraph(font_size_pt=para_font, line_spacing_value=160.0)
        p.add_run("Filler")
        blocks.append(p)

    # Sub heading
    subh = Paragraph(
        font_size_pt=13.0, line_spacing_value=160.0,
        keep_with_next=True, space_before_pt=3.0,
    )
    subh.add_run("\u25cb Subheading")
    blocks.append(subh)

    # Following block
    p = Paragraph(font_size_pt=13.0, line_spacing_value=160.0)
    p.add_run("Content after subheading")
    blocks.append(p)

    result = controller.process(blocks)

    # Find the subheading
    subh_idx = None
    for idx, b in enumerate(result):
        if isinstance(b, Paragraph) and hasattr(b, 'plain_text'):
            if b.plain_text.startswith("\u25cb"):
                subh_idx = idx
                break

    assert subh_idx is not None
    assert subh_idx > 0 and isinstance(result[subh_idx - 1], PageBreak), \
        "Expected PageBreak before sub heading"

    print("[PASS] test_sub_heading_keep_with_next")


# ─── auto_page_flow=False test ───────────────────────────────────

def test_auto_page_flow_disabled():
    """Test that auto_page_flow=False skips PageFlowController."""
    doc = HwpxDocument()

    # Add lots of paragraphs (would normally trigger page breaks)
    for i in range(100):
        doc.add_paragraph(f"Paragraph {i} with enough text to fill some lines. " * 3)

    # Add a table
    table = doc.add_table(col_widths=[80, 80])
    for r in range(10):
        table.add_row([f"r{r}", f"v{r}"])

    # Save with auto_page_flow disabled
    out = Path(__file__).parent / "output_p5_noflow.hwpx"
    doc.save(str(out), auto_page_flow=False)

    # No PageBreak should have been inserted
    page_breaks = [b for b in doc.blocks if isinstance(b, PageBreak)]
    assert len(page_breaks) == 0, "No auto page breaks when auto_page_flow=False"

    print("[PASS] test_auto_page_flow_disabled")
    out.unlink()


# ─── Integration: full document with auto page flow ──────────────

def test_full_document_auto_page_flow():
    """Integration: document with mixed content and auto page flow."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(
        font_body="함초롬돋움",
        font_table="함초롬돋움",
        size_body=13,
        size_table=10,
        margin_top=20, margin_bottom=20,
        margin_left=20, margin_right=20,
    ))

    doc.add_title("Phase 5 Integration Test")

    # Fill a page with content
    for i in range(30):
        doc.add_paragraph(
            f"Paragraph {i}: " + "This is filler text for testing. " * 5,
            font_size_pt=13.0,
        )

    # Section heading near the bottom should trigger keep-with-next
    doc.add_section_heading(1, "Important Section")
    doc.add_sub_heading("Details")
    doc.add_sub_item("가", "First sub-item")
    doc.add_bullet1("Explanation point one")
    doc.add_bullet2("Minor detail")

    # Table
    table = doc.add_table(col_widths=[40, 130])
    table.add_header_row(["Key", "Value"])
    for i in range(8):
        table.add_row([f"Item {i}", f"Description for item {i}"])

    # More content
    for i in range(10):
        doc.add_paragraph(f"After table paragraph {i}. " * 4)

    # Manual page break
    doc.add_page_break()
    doc.add_paragraph("Content after manual page break.")

    doc.add_note("Important note at the end")

    out = Path(__file__).parent / "output_p5_full.hwpx"
    doc.save(str(out), auto_page_flow=True)

    section, header = _read_hwpx(out)

    # Basic content checks
    assert "Phase 5 Integration Test" in section
    assert "Important Section" in section
    assert "Content after manual page break" in section
    assert "hp:tbl" in section

    # Check that pageBreakBefore="true" appears in the header styles
    assert 'pageBreakBefore="1"' in header

    print("[PASS] test_full_document_auto_page_flow")
    out.unlink()


def test_page_break_before_paragraph_property():
    """Test that Paragraph.page_break_before renders correctly."""
    doc = HwpxDocument()
    doc.add_paragraph("Before")

    # Manually create a paragraph with page_break_before
    p = doc.add_paragraph("After break", font_size_pt=13.0)
    p.page_break_before = True

    out = Path(__file__).parent / "output_p5_pbb.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section, header = _read_hwpx(out)
    assert 'pageBreakBefore="1"' in header

    print("[PASS] test_page_break_before_paragraph_property")
    out.unlink()


# ─── Widow/orphan control test ───────────────────────────────────

def test_widow_orphan_control():
    """Test that widow/orphan control prevents single lines alone."""
    ps = PageSettings()
    controller = PageFlowController(ps)

    blocks = []
    content_h = ps.content_height_mm
    para_font = 13.0
    line_h = para_font * 160.0 / 100.0 * 0.3528

    # Fill page to leave room for only ~1 line
    lines_to_fill = int(content_h / line_h) - 1
    for _ in range(lines_to_fill):
        p = Paragraph(font_size_pt=para_font, line_spacing_value=160.0)
        p.add_run("X")
        blocks.append(p)

    # Multi-line paragraph that would orphan its first line
    long_para = Paragraph(font_size_pt=para_font, line_spacing_value=160.0)
    long_text = "A" * 500  # Very long text that wraps to many lines
    long_para.add_run(long_text)
    blocks.append(long_para)

    result = controller.process(blocks)

    # There should be a PageBreak somewhere before the long paragraph
    page_breaks = [b for b in result if isinstance(b, PageBreak)]
    assert len(page_breaks) >= 1, "Expected page break for widow/orphan control"

    print("[PASS] test_widow_orphan_control")


# ─── Line count estimation test ──────────────────────────────────

def test_line_count_estimation():
    """Test _estimate_line_count heuristic."""
    ps = PageSettings()  # content_width = 170mm
    controller = PageFlowController(ps)

    # Short text — should be 1 line
    p1 = Paragraph(font_size_pt=13.0, line_spacing_value=160.0)
    p1.add_run("Hello")
    assert controller._estimate_line_count(p1) == 1

    # Very long text — should be multiple lines
    p2 = Paragraph(font_size_pt=13.0, line_spacing_value=160.0)
    p2.add_run("A" * 1000)
    assert controller._estimate_line_count(p2) > 1

    # Empty paragraph — should be 1 line
    p3 = Paragraph()
    assert controller._estimate_line_count(p3) == 1

    print("[PASS] test_line_count_estimation")


# ─── Keep count detection tests ──────────────────────────────────

def test_keep_count_detection():
    """Test _get_keep_count for different heading types."""
    ps = PageSettings()
    controller = PageFlowController(ps)

    # Section heading: keeps 3
    h1 = Paragraph(keep_with_next=True)
    h1.add_run("\u25a0 1. Title")
    assert controller._get_keep_count(h1) == 3

    # Sub heading: keeps 1
    h2 = Paragraph(keep_with_next=True)
    h2.add_run("\u25cb Subtitle")
    assert controller._get_keep_count(h2) == 1

    # Korean sub item: keeps 1
    sub = Paragraph(keep_with_next=True)
    sub.add_run("가. First item")
    assert controller._get_keep_count(sub) == 1

    # Regular keep_with_next paragraph: keeps 1
    reg = Paragraph(keep_with_next=True)
    reg.add_run("Some regular text")
    assert controller._get_keep_count(reg) == 1

    print("[PASS] test_keep_count_detection")


# ─── Manual page break preserves in auto flow ────────────────────

def test_manual_break_preserved_in_auto_flow():
    """Manual PageBreak should be preserved when auto_page_flow is on."""
    doc = HwpxDocument()
    doc.add_paragraph("Page 1 content")
    doc.add_page_break()
    doc.add_paragraph("Page 2 content")

    out = Path(__file__).parent / "output_p5_manual_auto.hwpx"
    doc.save(str(out), auto_page_flow=True)

    # The manual PageBreak should still be in blocks
    page_breaks = [b for b in doc.blocks if isinstance(b, PageBreak)]
    assert len(page_breaks) >= 1, "Manual page break should be preserved"

    section, header = _read_hwpx(out)
    assert "Page 1 content" in section
    assert "Page 2 content" in section
    assert 'pageBreakBefore="1"' in header

    print("[PASS] test_manual_break_preserved_in_auto_flow")
    out.unlink()


# ─── Table keep_together=False should not trigger auto break ─────

def test_table_keep_together_false_no_auto_break():
    """Table with keep_together=False should not get auto page break."""
    ps = PageSettings()
    controller = PageFlowController(ps)

    blocks = []
    content_h = ps.content_height_mm
    para_font = 13.0
    line_h = para_font * 160.0 / 100.0 * 0.3528
    lines_needed = int(content_h * 0.9 / line_h)
    for _ in range(lines_needed):
        p = Paragraph(font_size_pt=para_font, line_spacing_value=160.0)
        p.add_run("X")
        blocks.append(p)

    # Table with keep_together=False
    table = Table(
        col_widths_mm=[80, 80], font_size_pt=10.0,
        cell_margin_tb_mm=2.0, keep_together=False,
    )
    for r in range(5):
        table.add_row([f"r{r}c0", f"r{r}c1"])
    blocks.append(table)

    result = controller.process(blocks)

    # Check: no page break should be inserted directly before the table
    table_idx = None
    for idx, b in enumerate(result):
        if isinstance(b, Table):
            table_idx = idx
            break
    assert table_idx is not None
    if table_idx > 0:
        assert not isinstance(result[table_idx - 1], PageBreak), \
            "No auto page break before table with keep_together=False"

    print("[PASS] test_table_keep_together_false_no_auto_break")


if __name__ == "__main__":
    test_manual_page_break()
    test_page_break_renders_xml()
    test_page_break_import()
    test_height_estimation_paragraph()
    test_height_estimation_table()
    test_height_estimation_page_break()
    test_table_keep_together_auto_break()
    test_table_no_break_when_fits()
    test_section_heading_keep_with_next()
    test_sub_heading_keep_with_next()
    test_auto_page_flow_disabled()
    test_full_document_auto_page_flow()
    test_page_break_before_paragraph_property()
    test_widow_orphan_control()
    test_line_count_estimation()
    test_keep_count_detection()
    test_manual_break_preserved_in_auto_flow()
    test_table_keep_together_false_no_auto_break()
    print("\n=== All Phase 5 tests passed! ===")
