"""Phase 11 tests: JSON loader and CLI."""

import json
import subprocess
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator import HwpxDocument, GovDocumentPreset
from hwpx_generator.json_loader import load_from_json, load_from_file


def _read_section(path):
    with zipfile.ZipFile(path, "r") as zf:
        return zf.read("Contents/section0.xml").decode("utf-8")


# ---------------------------------------------------------------------------
# json_loader unit tests
# ---------------------------------------------------------------------------

def test_load_empty_content():
    """Minimal valid JSON with no content blocks."""
    data = {"content": []}
    doc = load_from_json(data)
    assert isinstance(doc, HwpxDocument)
    assert len(doc.blocks) == 0
    print("[PASS] test_load_empty_content")


def test_load_title():
    """Title block dispatches to add_title."""
    data = {"content": [{"type": "title", "text": "My Title"}]}
    doc = load_from_json(data)
    assert len(doc.blocks) == 1
    assert "My Title" in doc.blocks[0].runs[0].text
    print("[PASS] test_load_title")


def test_load_section_heading():
    """Section heading with num and text (with preset for prefix)."""
    data = {
        "preset": "gov_document",
        "content": [
            {"type": "section_heading", "num": "1", "text": "Overview"}
        ],
    }
    doc = load_from_json(data)
    assert len(doc.blocks) == 1
    run_text = doc.blocks[0].runs[0].text
    assert "1" in run_text
    assert "Overview" in run_text
    print("[PASS] test_load_section_heading")


def test_load_sub_heading():
    """Sub-heading block."""
    data = {"content": [{"type": "sub_heading", "text": "Details"}]}
    doc = load_from_json(data)
    assert len(doc.blocks) == 1
    print("[PASS] test_load_sub_heading")


def test_load_sub_item():
    """Sub-item with Korean numbering (with preset for prefix)."""
    data = {
        "preset": "gov_document",
        "content": [
            {"type": "sub_item", "num": "가", "text": "First item"}
        ],
    }
    doc = load_from_json(data)
    assert len(doc.blocks) == 1
    run_text = doc.blocks[0].runs[0].text
    assert "가" in run_text
    assert "First item" in run_text
    print("[PASS] test_load_sub_item")


def test_load_bullets():
    """Bullet level 1 and level 2."""
    data = {"content": [
        {"type": "bullet", "level": 1, "text": "Item A"},
        {"type": "bullet", "level": 2, "text": "Sub-item B"},
    ]}
    doc = load_from_json(data)
    assert len(doc.blocks) == 2
    print("[PASS] test_load_bullets")


def test_load_note():
    """Note block."""
    data = {"content": [{"type": "note", "text": "Important note"}]}
    doc = load_from_json(data)
    assert len(doc.blocks) == 1
    print("[PASS] test_load_note")


def test_load_paragraph():
    """Paragraph with optional kwargs."""
    data = {"content": [
        {"type": "paragraph", "text": "Hello", "align": "CENTER", "bold": True}
    ]}
    doc = load_from_json(data)
    assert len(doc.blocks) == 1
    p = doc.blocks[0]
    assert p.align == "CENTER"
    assert p.bold is True
    print("[PASS] test_load_paragraph")


def test_load_paragraph_defaults():
    """Paragraph with only text uses defaults."""
    data = {"content": [{"type": "paragraph", "text": "Simple"}]}
    doc = load_from_json(data)
    assert len(doc.blocks) == 1
    print("[PASS] test_load_paragraph_defaults")


def test_load_table():
    """Table with header and rows."""
    data = {"content": [
        {
            "type": "table",
            "col_widths_mm": [30, 140],
            "header": ["Key", "Value"],
            "rows": [["A", "1"], ["B", "2"]],
        }
    ]}
    doc = load_from_json(data)
    assert len(doc.blocks) == 1
    table = doc.blocks[0]
    assert table.col_count == 2
    assert table.row_count == 3  # 1 header + 2 data
    print("[PASS] test_load_table")


def test_load_table_no_header():
    """Table without header row."""
    data = {"content": [
        {
            "type": "table",
            "col_widths_mm": [50, 120],
            "rows": [["X", "Y"]],
        }
    ]}
    doc = load_from_json(data)
    table = doc.blocks[0]
    assert table.row_count == 1
    print("[PASS] test_load_table_no_header")


def test_load_page_break():
    """Page break block."""
    data = {"content": [{"type": "page_break"}]}
    doc = load_from_json(data)
    assert len(doc.blocks) == 1
    from hwpx_generator.elements.page_break import PageBreak
    assert isinstance(doc.blocks[0], PageBreak)
    print("[PASS] test_load_page_break")


def test_load_diagram():
    """Diagram with nodes and edges."""
    data = {"content": [
        {
            "type": "diagram",
            "layout": "step_flow",
            "direction": "horizontal",
            "width_mm": 160,
            "height_mm": 30,
            "theme_color": "1F4E79",
            "nodes": [
                {"id": "s1", "label": "Step 1", "shape": "rect"},
                {"id": "s2", "label": "Step 2", "shape": "rect"},
            ],
            "edges": [
                {"from_id": "s1", "to_id": "s2"},
            ],
        }
    ]}
    doc = load_from_json(data)
    assert len(doc.blocks) == 1
    diagram = doc.blocks[0]
    assert len(diagram.nodes) == 2
    assert len(diagram.edges) == 1
    assert diagram.nodes[0].label == "Step 1"
    assert diagram.edges[0].from_id == "s1"
    print("[PASS] test_load_diagram")


def test_load_chart():
    """Chart with data."""
    data = {"content": [
        {
            "type": "chart",
            "chart_type": "bar",
            "width_mm": 120,
            "height_mm": 70,
            "title": "Sales",
            "data": {
                "labels": ["Q1", "Q2"],
                "datasets": [
                    {"label": "Revenue", "values": [100, 200], "color": "2E75B6"}
                ],
            },
        }
    ]}
    doc = load_from_json(data)
    assert len(doc.blocks) == 1
    chart = doc.blocks[0]
    assert chart.chart_type == "bar"
    assert chart.title == "Sales"
    assert chart.data.labels == ["Q1", "Q2"]
    assert len(chart.data.datasets) == 1
    assert chart.data.datasets[0].values == [100, 200]
    print("[PASS] test_load_chart")


def test_load_svg():
    """SVG block with src."""
    data = {"content": [
        {
            "type": "svg",
            "src": "diagram.svg",
            "width_mm": 160,
            "align": "center",
        }
    ]}
    doc = load_from_json(data)
    assert len(doc.blocks) == 1
    svg = doc.blocks[0]
    assert svg.src == "diagram.svg"
    assert svg.width_mm == 160
    print("[PASS] test_load_svg")


def test_load_text_box():
    """Text box with styling."""
    data = {"content": [
        {
            "type": "text_box",
            "text": "Warning!",
            "border_color": "C00000",
            "bg_color": "FFF2CC",
            "font_bold": True,
            "padding_mm": 5,
        }
    ]}
    doc = load_from_json(data)
    assert len(doc.blocks) == 1
    tb = doc.blocks[0]
    assert tb.text == "Warning!"
    assert tb.border_color == "C00000"
    assert tb.bg_color == "FFF2CC"
    assert tb.font_bold is True
    assert tb.padding_mm == 5
    print("[PASS] test_load_text_box")


# ---------------------------------------------------------------------------
# Preset tests
# ---------------------------------------------------------------------------

def test_preset_applied():
    """Preset 'gov_document' sets margins and styles."""
    data = {
        "preset": "gov_document",
        "preset_options": {
            "font_body": "맑은 고딕",
            "size_body": 12,
            "margin": {"top": 15, "left": 25},
            "cell_margin": {"lr": 4},
        },
        "content": [],
    }
    doc = load_from_json(data)
    assert doc._preset is not None
    assert doc._preset.font_body == "맑은 고딕"
    assert doc._preset.size_body == 12
    assert doc.page_settings.margin_top == 15
    assert doc.page_settings.margin_left == 25
    assert doc._preset.cell_margin_lr == 4
    print("[PASS] test_preset_applied")


def test_preset_defaults():
    """Preset with no options uses defaults."""
    data = {
        "preset": "gov_document",
        "content": [],
    }
    doc = load_from_json(data)
    assert doc._preset is not None
    assert doc._preset.font_body == "휴먼명조"
    assert doc._preset.size_body == 13.0
    print("[PASS] test_preset_defaults")


def test_unknown_preset_raises():
    """Unknown preset name raises ValueError."""
    data = {"preset": "unknown_preset", "content": []}
    try:
        load_from_json(data)
        assert False, "Expected ValueError"
    except ValueError as e:
        assert "unknown_preset" in str(e)
    print("[PASS] test_unknown_preset_raises")


def test_unknown_block_type_raises():
    """Unknown content type raises ValueError."""
    data = {"content": [{"type": "nonexistent"}]}
    try:
        load_from_json(data)
        assert False, "Expected ValueError"
    except ValueError as e:
        assert "nonexistent" in str(e)
    print("[PASS] test_unknown_block_type_raises")


def test_missing_type_key_raises():
    """Content block without 'type' raises ValueError."""
    data = {"content": [{"text": "no type"}]}
    try:
        load_from_json(data)
        assert False, "Expected ValueError"
    except ValueError as e:
        assert "type" in str(e).lower()
    print("[PASS] test_missing_type_key_raises")


# ---------------------------------------------------------------------------
# load_from_file tests
# ---------------------------------------------------------------------------

def test_load_from_file():
    """load_from_file reads a JSON file and builds a document."""
    tmp = Path(__file__).parent / "_test_p11_input.json"
    data = {
        "content": [
            {"type": "title", "text": "File Test"},
            {"type": "paragraph", "text": "From file."},
        ]
    }
    tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

    try:
        doc = load_from_file(str(tmp))
        assert len(doc.blocks) == 2
    finally:
        tmp.unlink(missing_ok=True)

    print("[PASS] test_load_from_file")


def test_load_from_file_not_found():
    """load_from_file raises FileNotFoundError for missing file."""
    try:
        load_from_file("/nonexistent/path.json")
        assert False, "Expected FileNotFoundError"
    except FileNotFoundError:
        pass
    print("[PASS] test_load_from_file_not_found")


# ---------------------------------------------------------------------------
# End-to-end: JSON -> save -> verify HWPX contents
# ---------------------------------------------------------------------------

def test_e2e_json_to_hwpx():
    """Full round-trip: JSON dict -> HwpxDocument -> .hwpx file."""
    data = {
        "preset": "gov_document",
        "preset_options": {
            "font_body": "휴먼명조",
            "size_title": 15,
            "size_body": 13,
        },
        "content": [
            {"type": "title", "text": "테스트 문서"},
            {"type": "section_heading", "num": "1", "text": "개요"},
            {"type": "bullet", "level": 1, "text": "첫 번째 항목"},
            {"type": "bullet", "level": 2, "text": "하위 항목"},
            {"type": "note", "text": "주의 사항"},
            {"type": "paragraph", "text": "일반 문단입니다."},
            {"type": "table", "col_widths_mm": [30, 140],
             "header": ["항목", "내용"],
             "rows": [["기간", "3월"]]},
            {"type": "text_box", "text": "강조 문구",
             "border_color": "C00000", "bg_color": "FFF2CC"},
        ],
    }

    out = Path(__file__).parent / "output_p11_e2e.hwpx"
    doc = load_from_json(data)
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "테스트 문서" in section, "Expected title text"
    assert "개요" in section, "Expected heading text"
    assert "첫 번째 항목" in section, "Expected bullet text"
    assert "하위 항목" in section, "Expected sub-bullet text"
    assert "주의 사항" in section, "Expected note text"
    assert "일반 문단입니다" in section, "Expected paragraph text"
    assert "기간" in section, "Expected table cell text"
    assert "강조 문구" in section, "Expected text box text"

    print("[PASS] test_e2e_json_to_hwpx")
    out.unlink()


def test_e2e_with_page_flow():
    """End-to-end test with auto page flow enabled."""
    data = {
        "preset": "gov_document",
        "content": [
            {"type": "title", "text": "Page Flow Test"},
            {"type": "paragraph", "text": "Content A"},
            {"type": "page_break"},
            {"type": "paragraph", "text": "Content B"},
        ],
    }

    out = Path(__file__).parent / "output_p11_pageflow.hwpx"
    doc = load_from_json(data)
    doc.save(str(out), auto_page_flow=True)

    section = _read_section(out)
    assert "Page Flow Test" in section
    assert "Content A" in section
    assert "Content B" in section

    print("[PASS] test_e2e_with_page_flow")
    out.unlink()


def test_e2e_diagram_and_chart():
    """Diagram and chart blocks produce valid HWPX."""
    data = {
        "content": [
            {
                "type": "diagram",
                "layout": "step_flow",
                "direction": "horizontal",
                "width_mm": 160,
                "height_mm": 30,
                "nodes": [
                    {"id": "a", "label": "Alpha", "shape": "rect"},
                    {"id": "b", "label": "Beta", "shape": "rect"},
                ],
                "edges": [{"from_id": "a", "to_id": "b"}],
            },
            {
                "type": "chart",
                "chart_type": "bar",
                "width_mm": 120,
                "height_mm": 70,
                "title": "Test Chart",
                "data": {
                    "labels": ["X", "Y"],
                    "datasets": [
                        {"label": "Series", "values": [10, 20], "color": "2E75B6"}
                    ],
                },
            },
        ],
    }

    out = Path(__file__).parent / "output_p11_diag_chart.hwpx"
    doc = load_from_json(data)
    doc.save(str(out), auto_page_flow=False)

    assert out.exists(), "HWPX file should be created"
    # Verify it's a valid zip
    with zipfile.ZipFile(out, "r") as zf:
        names = zf.namelist()
        assert "Contents/section0.xml" in names

    print("[PASS] test_e2e_diagram_and_chart")
    out.unlink()


# ---------------------------------------------------------------------------
# CLI tests
# ---------------------------------------------------------------------------

def test_cli_basic():
    """CLI produces an .hwpx file from a JSON input."""
    cli_py = str(Path(__file__).resolve().parent.parent / "cli.py")

    # Create a temporary JSON input
    tmp_json = Path(__file__).parent / "_test_cli_input.json"
    tmp_hwpx = Path(__file__).parent / "_test_cli_output.hwpx"
    data = {
        "preset": "gov_document",
        "content": [
            {"type": "title", "text": "CLI Test"},
            {"type": "paragraph", "text": "Created by CLI"},
        ],
    }
    tmp_json.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

    try:
        result = subprocess.run(
            [sys.executable, cli_py, "--input", str(tmp_json), "--output", str(tmp_hwpx)],
            capture_output=True, text=True, timeout=30,
        )
        assert result.returncode == 0, f"CLI failed: {result.stderr}"
        assert "Success" in result.stdout
        assert tmp_hwpx.exists(), "Output .hwpx should exist"

        # Verify contents
        section = _read_section(tmp_hwpx)
        assert "CLI Test" in section
        assert "Created by CLI" in section
    finally:
        tmp_json.unlink(missing_ok=True)
        tmp_hwpx.unlink(missing_ok=True)

    print("[PASS] test_cli_basic")


def test_cli_no_page_flow():
    """CLI with --no-page-flow flag."""
    cli_py = str(Path(__file__).resolve().parent.parent / "cli.py")

    tmp_json = Path(__file__).parent / "_test_cli_npf_input.json"
    tmp_hwpx = Path(__file__).parent / "_test_cli_npf_output.hwpx"
    data = {
        "content": [
            {"type": "paragraph", "text": "No page flow"},
        ],
    }
    tmp_json.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

    try:
        result = subprocess.run(
            [sys.executable, cli_py, "-i", str(tmp_json), "-o", str(tmp_hwpx),
             "--no-page-flow"],
            capture_output=True, text=True, timeout=30,
        )
        assert result.returncode == 0, f"CLI failed: {result.stderr}"
        assert tmp_hwpx.exists()
    finally:
        tmp_json.unlink(missing_ok=True)
        tmp_hwpx.unlink(missing_ok=True)

    print("[PASS] test_cli_no_page_flow")


def test_cli_missing_input():
    """CLI exits with error when input file is missing."""
    cli_py = str(Path(__file__).resolve().parent.parent / "cli.py")

    result = subprocess.run(
        [sys.executable, cli_py, "--input", "/no/such/file.json",
         "--output", "out.hwpx"],
        capture_output=True, text=True, timeout=30,
    )
    assert result.returncode != 0, "CLI should fail for missing input"
    assert "Error" in result.stderr or "error" in result.stderr.lower()
    print("[PASS] test_cli_missing_input")


def test_cli_missing_args():
    """CLI exits with error when required args are missing."""
    cli_py = str(Path(__file__).resolve().parent.parent / "cli.py")

    result = subprocess.run(
        [sys.executable, cli_py],
        capture_output=True, text=True, timeout=30,
    )
    assert result.returncode != 0, "CLI should fail without args"
    print("[PASS] test_cli_missing_args")


# ---------------------------------------------------------------------------
# Multiple block types in sequence
# ---------------------------------------------------------------------------

def test_mixed_content_sequence():
    """All content types together in a single document."""
    data = {
        "preset": "gov_document",
        "preset_options": {"font_body": "맑은 고딕", "size_body": 12},
        "content": [
            {"type": "title", "text": "종합 테스트"},
            {"type": "section_heading", "num": "1", "text": "섹션 A"},
            {"type": "sub_heading", "text": "소제목 A"},
            {"type": "sub_item", "num": "가", "text": "세부항목"},
            {"type": "bullet", "level": 1, "text": "항목 1"},
            {"type": "bullet", "level": 2, "text": "항목 1-1"},
            {"type": "note", "text": "참고 사항"},
            {"type": "paragraph", "text": "일반 문단"},
            {"type": "table", "col_widths_mm": [40, 130],
             "header": ["Col1", "Col2"],
             "rows": [["R1C1", "R1C2"]]},
            {"type": "page_break"},
            {"type": "diagram", "layout": "step_flow", "width_mm": 100, "height_mm": 25,
             "nodes": [{"id": "n1", "label": "Node"}], "edges": []},
            {"type": "chart", "chart_type": "line", "title": "Chart",
             "data": {"labels": ["A"], "datasets": [{"label": "S", "values": [5]}]}},
            {"type": "text_box", "text": "Box", "border_color": "000000"},
        ],
    }

    out = Path(__file__).parent / "output_p11_mixed.hwpx"
    doc = load_from_json(data)
    # 13 content blocks but page_break is separate element
    assert len(doc.blocks) == 13

    doc.save(str(out), auto_page_flow=False)
    assert out.exists()

    section = _read_section(out)
    assert "종합 테스트" in section
    assert "섹션 A" in section
    assert "소제목 A" in section
    assert "일반 문단" in section

    print("[PASS] test_mixed_content_sequence")
    out.unlink()


# ---------------------------------------------------------------------------
# sample_notice.json integration test
# ---------------------------------------------------------------------------

def test_sample_notice_json():
    """The provided sample_notice.json produces a valid .hwpx."""
    sample = Path(__file__).resolve().parent.parent / "examples" / "sample_notice.json"
    if not sample.exists():
        print("[SKIP] test_sample_notice_json — sample file not found")
        return

    doc = load_from_file(str(sample))
    assert len(doc.blocks) > 0, "Expected content blocks from sample"

    out = Path(__file__).parent / "output_p11_sample.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section = _read_section(out)
    assert "모니터링" in section, "Expected sample content in output"

    print("[PASS] test_sample_notice_json")
    out.unlink()


if __name__ == "__main__":
    # json_loader unit tests
    test_load_empty_content()
    test_load_title()
    test_load_section_heading()
    test_load_sub_heading()
    test_load_sub_item()
    test_load_bullets()
    test_load_note()
    test_load_paragraph()
    test_load_paragraph_defaults()
    test_load_table()
    test_load_table_no_header()
    test_load_page_break()
    test_load_diagram()
    test_load_chart()
    test_load_svg()
    test_load_text_box()

    # Preset tests
    test_preset_applied()
    test_preset_defaults()
    test_unknown_preset_raises()
    test_unknown_block_type_raises()
    test_missing_type_key_raises()

    # File loading tests
    test_load_from_file()
    test_load_from_file_not_found()

    # End-to-end tests
    test_e2e_json_to_hwpx()
    test_e2e_with_page_flow()
    test_e2e_diagram_and_chart()

    # CLI tests
    test_cli_basic()
    test_cli_no_page_flow()
    test_cli_missing_input()
    test_cli_missing_args()

    # Mixed content
    test_mixed_content_sequence()

    # Sample notice
    test_sample_notice_json()

    print("\n=== All Phase 11 tests passed! ===")
