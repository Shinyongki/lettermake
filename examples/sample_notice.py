#!/usr/bin/env python3
"""Generate a complete government monitoring notice document using the Python API.

This script demonstrates UC-01: generating a government notice document
(모니터링 실시 안내서) using all features of hwpx_generator:
- Document preset (fonts, margins, spacing)
- Title, section headings, sub-headings, sub-items
- Bullet lists at multiple levels
- Tables with headers, merged cells
- Step flow diagrams (native vector)
- Bar/line charts (OOXML native)
- SVG converted to native shapes
- Text boxes with borders and backgrounds
- Notes and annotations
- Page breaks
"""

from __future__ import annotations

import sys
from pathlib import Path

# Allow running from the repository root
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator import HwpxDocument
from hwpx_generator.presets import GovDocumentPreset


def main() -> None:
    # ── 1. Create document and apply government preset ──────────
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

    # ── 2. Title ────────────────────────────────────────────────
    doc.add_title("2026년도 모니터링 실시 안내")
    doc.add_empty_paragraph()

    # ── 3. Section 1: 추진 배경 ────────────────────────────────
    doc.add_section_heading("1", "추진 배경")
    doc.add_bullet1("정부 정책 이행 점검 및 현장 모니터링 강화")
    doc.add_bullet1("관련 법령에 따른 정기 점검 의무 이행")
    doc.add_bullet2("「공공기관 운영에 관한 법률」 제xx조")
    doc.add_bullet2("「정부업무 평가 기본법」 제xx조")

    # ── 4. Section 2: 모니터링 개요 ─────────────────────────────
    doc.add_section_heading("2", "모니터링 개요")

    # Sub-heading: 점검 대상 및 기간
    doc.add_sub_heading("점검 대상 및 기간")

    # Summary table
    table1 = doc.add_table([25, 145])
    table1.add_header_row(["항목", "내용"])
    table1.add_row(["대상", "전국 17개 시·도 산하 공공기관"])
    table1.add_row(["기간", "2026. 3. 10.(화) ~ 3. 31.(화)"])
    table1.add_row(["방법", "현장 방문 점검 및 서면 점검 병행"])
    table1.add_row(["점검인력", "본부 및 지역본부 합동 점검단"])

    # Sub-heading: 추진 일정
    doc.add_sub_heading("추진 일정")
    doc.add_sub_item("가", "사전 준비: 3. 1.(일) ~ 3. 9.(월)")
    doc.add_sub_item("나", "현장 점검: 3. 10.(화) ~ 3. 31.(화)")
    doc.add_sub_item("다", "결과 보고: 4. 1.(수) ~ 4. 10.(금)")

    # ── 5. Section 3: 점검 항목 ─────────────────────────────────
    doc.add_section_heading("3", "점검 항목")

    # Detailed table with merged cells
    table2 = doc.add_table([10, 40, 80, 40])
    table2.add_header_row(["No.", "점검 분야", "세부 점검 항목", "비고"])
    table2.add_merged_row([
        {"text": "1", "align": "CENTER"},
        {"text": "안전관리", "rowspan": 2},
        {"text": "소방시설 점검 및 관리 현황"},
        {"text": "필수", "align": "CENTER"},
    ])
    table2.add_merged_row([
        {"text": "2", "align": "CENTER"},
        {"text": "비상대피 훈련 실시 여부"},
        {"text": "필수", "align": "CENTER"},
    ])
    table2.add_row(["3", "시설관리", "노후시설 보수·교체 현황", "선택"])
    table2.add_row(["4", "운영관리", "운영 매뉴얼 구비 여부", "필수"])
    table2.add_row(["5", "운영관리", "인력 배치 적정성", "선택"])

    # ── 6. Section 4: 점검 절차 (Step Flow Diagram) ─────────────
    doc.add_section_heading("4", "점검 절차")
    doc.add_paragraph(
        "점검은 아래 4단계로 진행되며, 각 단계별 세부 사항은 첨부 참조.",
        align="JUSTIFY",
    )

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

    doc.add_empty_paragraph()

    # ── 7. Section 5: 점검 결과 (Chart) ─────────────────────────
    doc.add_section_heading("5", "점검 결과 (예시)")
    doc.add_paragraph(
        "아래 차트는 전년도 점검 결과 요약입니다.",
        align="JUSTIFY",
    )

    doc.add_chart(
        chart_type="bar",
        width_mm=130.0,
        height_mm=75.0,
        title="2025년 분야별 점검 결과",
        labels=["안전관리", "시설관리", "운영관리", "인력관리"],
        datasets=[
            {"label": "적합", "values": [85, 72, 90, 78], "color": "2E75B6"},
            {"label": "부적합", "values": [15, 28, 10, 22], "color": "C00000"},
        ],
    )

    # ── 8. Section 6: 협조 사항 ─────────────────────────────────
    doc.add_page_break()
    doc.add_section_heading("6", "협조 사항")
    doc.add_bullet1("점검 대상기관은 자체점검표를 사전에 작성하여 제출")
    doc.add_bullet1("현장 점검 시 담당자 입회 필수")
    doc.add_bullet1("관련 서류 및 증빙자료 사전 준비")

    # Note
    doc.add_note("자체점검표는 반드시 3월 4일(수)까지 제출하여야 합니다.")

    # Highlighted text box
    doc.add_text_box(
        "※ 자체점검표는 반드시 3월 4일(수)까지 제출하여야 합니다.\n"
        "※ 미제출 시 점검 시 불이익이 있을 수 있습니다.",
        border_color="C00000",
        bg_color="FFF2CC",
        font_bold=True,
        padding_mm=5.0,
    )

    doc.add_empty_paragraph()

    # ── 9. SVG diagram (simple organization chart) ──────────────
    doc.add_sub_heading("점검단 구성도")
    doc.add_svg(
        svg_string=(
            '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 300 150">'
            '  <rect x="100" y="5" width="100" height="35" rx="5" '
            '        fill="#1F4E79" stroke="#000"/>'
            '  <rect x="10" y="80" width="80" height="35" rx="5" '
            '        fill="#2E75B6" stroke="#000"/>'
            '  <rect x="110" y="80" width="80" height="35" rx="5" '
            '        fill="#2E75B6" stroke="#000"/>'
            '  <rect x="210" y="80" width="80" height="35" rx="5" '
            '        fill="#2E75B6" stroke="#000"/>'
            '  <line x1="150" y1="40" x2="50" y2="80" stroke="#333" stroke-width="2"/>'
            '  <line x1="150" y1="40" x2="150" y2="80" stroke="#333" stroke-width="2"/>'
            '  <line x1="150" y1="40" x2="250" y2="80" stroke="#333" stroke-width="2"/>'
            '</svg>'
        ),
        width_mm=120.0,
        caption="[그림 1] 합동 점검단 구성도",
    )

    # ── 10. Closing ─────────────────────────────────────────────
    doc.add_empty_paragraph()
    doc.add_paragraph(
        "붙임  1. 모니터링 실시계획 1부.",
        align="LEFT",
    )
    doc.add_paragraph(
        "      2. 자체점검표 서식 1부.",
        align="LEFT",
    )
    doc.add_paragraph(
        "      3. 점검 체크리스트 1부.  끝.",
        align="LEFT",
    )

    # ── 11. Save ────────────────────────────────────────────────
    output_path = Path(__file__).parent / "모니터링_실시안내서.hwpx"
    doc.save(str(output_path))
    print(f"Generated: {output_path.resolve()}")


if __name__ == "__main__":
    main()
