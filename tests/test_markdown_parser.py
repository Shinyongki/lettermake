#!/usr/bin/env python3
"""End-to-end test for MarkdownParser → HWPX generation.

Tests:
  1. Unit: parse each markdown element into correct block types
  2. E2E:  parse sample_notice.md → .hwpx file via CLI and Python API
"""

from __future__ import annotations

import io
import os
import sys
import zipfile
from pathlib import Path

# Force UTF-8 output on Windows
if sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# Allow importing from project root
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator.document import HwpxDocument
from hwpx_generator.presets.gov_document import GovDocumentPreset
from hwpx_generator.parsers.markdown_parser import MarkdownParser
from hwpx_generator.elements.paragraph import Paragraph
from hwpx_generator.elements.table import Table
from hwpx_generator.elements.page_break import PageBreak
from hwpx_generator.elements.text_box import TextBox

TESTS_DIR = Path(__file__).resolve().parent
SAMPLE_MD = TESTS_DIR / "sample_notice.md"
OUTPUT_HWPX = TESTS_DIR / "test_markdown_output.hwpx"

passed = 0
failed = 0


def check(name: str, condition: bool, detail: str = "") -> None:
    global passed, failed
    if condition:
        passed += 1
        print(f"  [PASS] {name}")
    else:
        failed += 1
        msg = f"  [FAIL] {name}"
        if detail:
            msg += f" -{detail}"
        print(msg)


# ---------------------------------------------
# Test 1: Unit -parse individual elements
# ---------------------------------------------
def test_unit_parsing():
    print("\n=== Test 1: Unit - individual element parsing ===")

    md = """\
# 문서 제목
부제목 라인

## 1. 첫 번째 섹션

### 소제목

#### 세부항목 A

#### 세부항목 B

- 불릿 항목 1
  - 하위 불릿

※ 주의사항입니다

| 이름 | 값 |
|------|-----|
| A    | 100 |
| B    | 200 |

---

> 참고 텍스트 박스입니다.
"""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset())
    parser = MarkdownParser()
    parser.parse(md, doc)

    blocks = doc.blocks
    check("block count >= 10", len(blocks) >= 10, f"got {len(blocks)}")

    # Title
    b0 = blocks[0]
    check("title is Paragraph", isinstance(b0, Paragraph))
    if b0.runs:
        check("title text contains '문서 제목'", "문서 제목" in b0.runs[0].text)
        check("title text contains '부제목'", "부제목" in b0.runs[0].text)

    # Section heading
    b1 = blocks[1]
    check("section_heading is Paragraph", isinstance(b1, Paragraph))
    if b1.runs:
        check("section contains ■", "■" in b1.runs[0].text)
        check("section contains '1.'", "1." in b1.runs[0].text)

    # Sub heading
    b2 = blocks[2]
    check("sub_heading is Paragraph", isinstance(b2, Paragraph))
    if b2.runs:
        check("sub_heading contains ○", "○" in b2.runs[0].text)

    # Sub1 items (가. 나.)
    b3 = blocks[3]
    check("sub1_a is Paragraph", isinstance(b3, Paragraph))
    if b3.runs:
        check("sub1_a contains '가.'", "가." in b3.runs[0].text)

    b4 = blocks[4]
    check("sub1_b is Paragraph", isinstance(b4, Paragraph))
    if b4.runs:
        check("sub1_b contains '나.'", "나." in b4.runs[0].text)

    # Bullet1
    b5 = blocks[5]
    check("bullet1 is Paragraph", isinstance(b5, Paragraph))
    if b5.runs:
        check("bullet1 contains '-'", "- " in b5.runs[0].text or b5.runs[0].text.startswith("-"))

    # Bullet2
    b6 = blocks[6]
    check("bullet2 is Paragraph", isinstance(b6, Paragraph))
    if b6.runs:
        check("bullet2 contains '·'", "·" in b6.runs[0].text)

    # Note
    b7 = blocks[7]
    check("note is Paragraph", isinstance(b7, Paragraph))
    if b7.runs:
        check("note contains ※", "※" in b7.runs[0].text)

    # Table
    b8 = blocks[8]
    check("table is Table", isinstance(b8, Table))
    if isinstance(b8, Table):
        check("table has 2 columns", len(b8.col_widths_mm) == 2)
        # header row + 2 data rows = 3 rows total
        check("table has 3 rows", len(b8.rows) == 3, f"got {len(b8.rows)}")

    # Page break
    b9 = blocks[9]
    check("page_break is PageBreak", isinstance(b9, PageBreak))

    # Text box
    b10 = blocks[10]
    check("text_box is TextBox", isinstance(b10, TextBox))
    if isinstance(b10, TextBox):
        check("text_box contains '참고'", "참고" in b10.text)


# ---------------------------------------------
# Test 2: Korean sub-numbering reset
# ---------------------------------------------
def test_sub_numbering_reset():
    print("\n=== Test 2: Korean sub-numbering resets per section ===")

    md = """\
## 1. 섹션 A

#### 항목 1
#### 항목 2

### 소제목

#### 항목 3
#### 항목 4
"""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset())
    parser = MarkdownParser()
    parser.parse(md, doc)

    blocks = doc.blocks
    # blocks[0] = section A
    # blocks[1] = 가. 항목 1
    # blocks[2] = 나. 항목 2
    # blocks[3] = 소제목
    # blocks[4] = 가. 항목 3  (reset!)
    # blocks[5] = 나. 항목 4

    check("total blocks = 6", len(blocks) == 6, f"got {len(blocks)}")

    if len(blocks) >= 3 and blocks[1].runs and blocks[2].runs:
        check("first sub1 = 가.", "가." in blocks[1].runs[0].text)
        check("second sub1 = 나.", "나." in blocks[2].runs[0].text)

    if len(blocks) >= 6 and blocks[4].runs and blocks[5].runs:
        check("after reset sub1 = 가.", "가." in blocks[4].runs[0].text)
        check("after reset sub1 = 나.", "나." in blocks[5].runs[0].text)


# ---------------------------------------------
# Test 3: Multi-line blockquote
# ---------------------------------------------
def test_multiline_blockquote():
    print("\n=== Test 3: Multi-line blockquote → single TextBox ===")

    md = """\
> 첫 번째 줄
> 두 번째 줄
> 세 번째 줄
"""
    doc = HwpxDocument()
    parser = MarkdownParser()
    parser.parse(md, doc)

    check("one block created", len(doc.blocks) == 1)
    b = doc.blocks[0]
    check("is TextBox", isinstance(b, TextBox))
    if isinstance(b, TextBox):
        check("contains all lines", "첫 번째" in b.text and "세 번째" in b.text)


# ---------------------------------------------
# Test 4: Section auto-numbering
# ---------------------------------------------
def test_section_auto_numbering():
    print("\n=== Test 4: Section auto-numbering without explicit num ===")

    md = """\
## 섹션 하나

## 섹션 둘

## 3. 명시적 번호
"""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset())
    parser = MarkdownParser()
    parser.parse(md, doc)

    check("3 blocks", len(doc.blocks) == 3)

    if doc.blocks[0].runs:
        check("auto num 1", "1." in doc.blocks[0].runs[0].text)
    if doc.blocks[1].runs:
        check("auto num 2", "2." in doc.blocks[1].runs[0].text)
    if doc.blocks[2].runs:
        check("explicit num 3", "3." in doc.blocks[2].runs[0].text)


# ---------------------------------------------
# Test 5: End-to-end -sample_notice.md → .hwpx
# ---------------------------------------------
def test_e2e_sample_notice():
    print("\n=== Test 5: E2E -sample_notice.md → .hwpx ===")

    check("sample .md exists", SAMPLE_MD.exists())

    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset())
    doc.load_markdown(str(SAMPLE_MD))

    blocks = doc.blocks
    check("blocks > 0", len(blocks) > 0, f"got {len(blocks)}")

    # Count block types
    para_count = sum(1 for b in blocks if isinstance(b, Paragraph))
    table_count = sum(1 for b in blocks if isinstance(b, Table))
    pb_count = sum(1 for b in blocks if isinstance(b, PageBreak))
    tb_count = sum(1 for b in blocks if isinstance(b, TextBox))

    check("paragraphs present", para_count > 0, f"got {para_count}")
    check("2 tables present", table_count == 2, f"got {table_count}")
    check("1 page_break", pb_count == 1, f"got {pb_count}")
    check("1 text_box", tb_count == 1, f"got {tb_count}")

    # Save to .hwpx
    doc.save(str(OUTPUT_HWPX))
    check(".hwpx file created", OUTPUT_HWPX.exists())

    # Verify ZIP structure
    if OUTPUT_HWPX.exists():
        with zipfile.ZipFile(str(OUTPUT_HWPX), "r") as zf:
            names = zf.namelist()
            check("contains mimetype", "mimetype" in names)
            check("contains section0.xml", "Contents/section0.xml" in names)
            check("contains header.xml", "Contents/header.xml" in names)

            # Read section0.xml and verify content
            sec_xml = zf.read("Contents/section0.xml").decode("utf-8")
            check("section0 has ■", "■" in sec_xml)
            check("section0 has ○", "○" in sec_xml)
            check("section0 has 가.", "가." in sec_xml)
            check("section0 has ※", "※" in sec_xml)
            check("section0 has table tag", "<hp:tbl" in sec_xml)

    print(f"\n  Output: {OUTPUT_HWPX}")


# ---------------------------------------------
# Test 6: CLI integration
# ---------------------------------------------
def test_cli():
    print("\n=== Test 6: CLI -python cli.py --input .md --output .hwpx ===")

    cli_output = TESTS_DIR / "test_cli_md_output.hwpx"
    if cli_output.exists():
        cli_output.unlink()

    from cli import main as cli_main
    try:
        cli_main([
            "--input", str(SAMPLE_MD),
            "--output", str(cli_output),
            "--preset", "gov_document",
        ])
        check("CLI completed without error", True)
    except SystemExit as e:
        check("CLI completed without error", e.code == 0 or e.code is None, f"exit code {e.code}")
    except Exception as e:
        check("CLI completed without error", False, str(e))

    check("CLI output file created", cli_output.exists())

    # Cleanup
    if cli_output.exists():
        cli_output.unlink()


# ---------------------------------------------
# Run all tests
# ---------------------------------------------
if __name__ == "__main__":
    print("=" * 60)
    print("  MarkdownParser End-to-End Tests")
    print("=" * 60)

    test_unit_parsing()
    test_sub_numbering_reset()
    test_multiline_blockquote()
    test_section_auto_numbering()
    test_e2e_sample_notice()
    test_cli()

    print("\n" + "=" * 60)
    print(f"  Results: {passed} passed, {failed} failed")
    print("=" * 60)

    if failed > 0:
        sys.exit(1)
