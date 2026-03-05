"""외부 스타일 동적 본문 생성기.

템플릿(report-template.hwpx)에서 원형을 추출하고,
마크다운 파싱 결과를 기반으로 본문(파트2)을 동적으로 조립한다.

모든 XML 조작은 xml.etree.ElementTree 노드 단위로 처리한다.
"""

from __future__ import annotations

import copy
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import List, Optional


# ── 경로 ──────────────────────────────────────────────
_SCRIPT_DIR = Path(__file__).resolve().parent
_ASSETS_DIR = _SCRIPT_DIR.parent / "assets"
TEMPLATE = _ASSETS_DIR / "report-template.hwpx"
FIX_NS = _SCRIPT_DIR / "fix_namespaces.py"

# 장식선 PNG (hwpx_generator에서 가져옴)
_HG_ASSETS = Path(__file__).resolve().parent.parent.parent.parent / "hwpx_generator" / "assets"
_DECO_TOP_PNG = _HG_ASSETS / "doc_title_top.png"
_DECO_BOTTOM_PNG = _HG_ASSETS / "doc_title_bottom.png"


# ── 네임스페이스 ─────────────────────────────────────
_NS = {
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hp10": "http://www.hancom.co.kr/hwpml/2016/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hhs": "http://www.hancom.co.kr/hwpml/2011/history",
    "hm": "http://www.hancom.co.kr/hwpml/2011/master-page",
    "hpf": "http://www.hancom.co.kr/schema/2011/hpf",
    "dc": "http://purl.org/dc/elements/1.1/",
    "opf": "http://www.idpf.org/2007/opf/",
    "ooxmlchart": "http://www.hancom.co.kr/hwpml/2016/ooxmlchart",
    "hwpunitchar": "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar",
    "epub": "http://www.idpf.org/2007/ops",
    "config": "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
}

_HP = _NS["hp"]
_HH = _NS["hh"]


def _register_namespaces() -> None:
    for prefix, uri in _NS.items():
        ET.register_namespace(prefix, uri)


def _tag(ns_prefix: str, local: str) -> str:
    return f"{{{_NS[ns_prefix]}}}{local}"


def _get_text_elements(elem: ET.Element) -> list[ET.Element]:
    """elem 하위의 모든 hp:t 요소를 반환."""
    return list(elem.iter(_tag("hp", "t")))


def _has_tbl(elem: ET.Element) -> bool:
    return elem.find(f".//{_tag('hp', 'tbl')}") is not None


# ── 데이터 구조 ──────────────────────────────────────
# ── 기호 접두사 (frontmatter로 오버라이드 가능) ────────
DEFAULT_PREFIX = {
    "square": "□",   # 2-run: run[0] 기호, run[1] 텍스트
    "circle": "○",   # 1-run: "  ○ " + 텍스트
    "dash":   "-",   # 1-run: "   - " + 텍스트
    "dot":    "•",   # 1-run: "    • " + 텍스트
    "note":   "※",   # 1-run: "     ※ " + 텍스트
}

# 들여쓰기 (기호 앞 공백) — 코드 고정, frontmatter 변경 불가
_INDENT = {
    "square": "",
    "circle": "  ",
    "dash":   "   ",
    "dot":    "    ",
    "note":   "     ",
}


@dataclass
class BulletItem:
    type: str          # "square", "circle", "dash", "dot", "note"
    text: str
    children: List[BulletItem] = field(default_factory=list)


@dataclass
class TableItem:
    type: str = "table"
    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)


@dataclass
class Section:
    title: str
    items: list = field(default_factory=list)  # BulletItem | TableItem


@dataclass
class ParsedDocument:
    title: str
    subtitle: str
    purpose: str = ""
    sections: List[Section] = field(default_factory=list)


@dataclass
class Prototypes:
    section_bar_p: ET.Element    # 섹션바 포함 <hp:p> 노드
    title_block_p: ET.Element    # 본문 제목 포함 <hp:p> 노드
    slot_square: ET.Element      # □ 단락 노드
    slot_circle: ET.Element      # ○ 단락 노드
    slot_dash: ET.Element        # ― 단락 노드
    slot_dot: ET.Element         # • 단락 노드 (dash 복제)
    slot_note: ET.Element        # ※ 단락 노드
    empty_para: Optional[ET.Element]  # 섹션바 앞 빈 단락


# ── 1. split_sections (ET 기반) ──────────────────────
def split_sections(root: ET.Element) -> tuple[list[ET.Element], list[ET.Element], list[ET.Element]]:
    """루트의 직계 자식을 pageBreak='1' 기준으로 3개 파트로 분리.

    Returns: (part0_nodes, part1_nodes, part2_nodes)
    """
    children = list(root)
    pb_indices = [i for i, ch in enumerate(children) if ch.get("pageBreak") == "1"]

    if len(pb_indices) < 2:
        raise ValueError(f"pageBreak='1' 이 2개 미만입니다 (발견: {len(pb_indices)})")

    cut1 = pb_indices[0]
    cut2 = pb_indices[1]

    part0 = children[:cut1]
    part1 = children[cut1:cut2]
    part2 = children[cut2:]

    return part0, part1, part2


# ── 2. extract_prototypes (ET 기반) ──────────────────
def extract_prototypes(part2: list[ET.Element]) -> Prototypes:
    """파트2 노드 리스트에서 원형을 추출한다."""
    title_block_p = None
    section_bar_p = None
    slot_square = None
    slot_circle = None
    slot_dash = None
    slot_note = None
    empty_para = None

    section_bar_found = False

    for i, p in enumerate(part2):
        has_table = _has_tbl(p)

        # 첫 번째 테이블 포함 단락 = 본문 제목
        if has_table and title_block_p is None:
            title_block_p = p
            continue

        # 두 번째 이후 테이블 포함 단락 = 섹션바
        if has_table and title_block_p is not None and section_bar_p is None:
            section_bar_p = p
            section_bar_found = True
            # 섹션바 바로 앞 단락이 빈 단락인지 확인
            if i > 0:
                prev = part2[i - 1]
                prev_texts = _get_text_elements(prev)
                if not any((t.text or "").strip() for t in prev_texts):
                    if not _has_tbl(prev) and prev != title_block_p:
                        empty_para = prev
            continue

        # 슬롯 단락 추출 (테이블 없는 단락)
        if not has_table and section_bar_found:
            pr_id = p.get("paraPrIDRef", "")
            # charPrIDRef 찾기
            runs = list(p.iter(_tag("hp", "run")))
            char_ids = [r.get("charPrIDRef", "") for r in runs]

            if pr_id == "28" and "5" in char_ids and slot_square is None:
                slot_square = p
            elif pr_id == "29" and "18" in char_ids and slot_circle is None:
                slot_circle = p
            elif pr_id == "30" and "18" in char_ids and slot_dash is None:
                slot_dash = p
            elif pr_id == "30" and "20" in char_ids and slot_note is None:
                slot_note = p

    if title_block_p is None:
        raise ValueError("본문 제목 블록을 찾을 수 없습니다")
    if section_bar_p is None:
        raise ValueError("섹션바를 찾을 수 없습니다")

    # slot_dot은 slot_dash 원형 복제 (paraPrIDRef=30, charPrIDRef=18 동일)
    slot_dot = copy.deepcopy(slot_dash) if slot_dash is not None else None

    return Prototypes(
        section_bar_p=section_bar_p,
        title_block_p=title_block_p,
        slot_square=slot_square,
        slot_circle=slot_circle,
        slot_dash=slot_dash,
        slot_dot=slot_dot,
        slot_note=slot_note,
        empty_para=empty_para,
    )


# ── 3. parse_markdown_sections ───────────────────────
_ROMAN = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ"]


def _parse_frontmatter(lines: list[str]) -> tuple[dict[str, str], list[str]]:
    """frontmatter(---블록)를 파싱하여 (설정dict, 나머지줄) 반환.

    외부 라이브러리 없이 `key: "value"` 패턴을 직접 파싱한다.
    """
    if not lines or lines[0].strip() != "---":
        return {}, lines

    config: dict[str, str] = {}
    end_idx = -1
    for i, line in enumerate(lines[1:], start=1):
        if line.strip() == "---":
            end_idx = i
            break
        m = re.match(r'^(\w+)\s*:\s*["\']?([^"\']*)["\']?\s*$', line.strip())
        if m:
            config[m.group(1)] = m.group(2)

    if end_idx < 0:
        return {}, lines

    return config, lines[end_idx + 1:]


def parse_markdown_sections(md_path: str) -> tuple[ParsedDocument, dict[str, str]]:
    """마크다운을 섹션별 계층 구조로 파싱한다.

    Returns: (ParsedDocument, prefix_overrides)
    """
    with open(md_path, encoding="utf-8") as f:
        lines = f.readlines()

    # frontmatter 파싱
    fm_config, lines = _parse_frontmatter(lines)
    prefix_overrides: dict[str, str] = {}
    for key, val in fm_config.items():
        if key.startswith("prefix_"):
            slot_type = key[len("prefix_"):]
            if slot_type in DEFAULT_PREFIX:
                prefix_overrides[slot_type] = val

    doc = ParsedDocument(title="", subtitle="")
    current_section: Optional[Section] = None
    stack_square: Optional[BulletItem] = None
    stack_circle: Optional[BulletItem] = None
    stack_dash: Optional[BulletItem] = None
    # ###/#### 헤딩 아래 `-` 항목의 들여쓰기 보정값
    indent_offset = 0

    for line in lines:
        stripped = line.rstrip("\n")
        # 탭 → 공백 2칸 정규화
        stripped = stripped.replace("\t", "  ")

        if re.match(r"^# [^#]", stripped):
            doc.title = stripped.lstrip("# ").strip()
            continue

        if doc.title and not doc.subtitle and not doc.sections:
            if stripped.strip() and not stripped.startswith("#") and not stripped.startswith("-"):
                text = stripped.strip()
                # "목적:" 또는 "목적 :" 패턴 감지
                m_purpose = re.match(r"^목적\s*[:：]\s*(.+)$", text)
                if m_purpose:
                    doc.purpose = m_purpose.group(1).strip()
                doc.subtitle = text
                continue

        if re.match(r"^## ", stripped):
            section_title = stripped.lstrip("# ").strip()
            current_section = Section(title=section_title)
            doc.sections.append(current_section)
            stack_square = None
            stack_circle = None
            stack_dash = None
            indent_offset = 0
            continue

        if re.match(r"^###+ ", stripped):
            # ### 또는 #### 헤딩 → square 슬롯
            if current_section is None:
                current_section = Section(title="")
                doc.sections.append(current_section)
            text = stripped.lstrip("# ").strip()
            item = BulletItem(type="square", text=text)
            current_section.items.append(item)
            stack_square = item
            stack_circle = None
            stack_dash = None
            # 헤딩 아래 `-` 항목은 1단계 깊은 것이므로 +2칸 보정
            indent_offset = 2
            continue

        # 표 파싱: | ... | 패턴
        if re.match(r"^\|.+\|$", stripped.strip()):
            if current_section is None:
                current_section = Section(title="")
                doc.sections.append(current_section)
            # 구분선(|---|---|) 건너뛰기
            if re.match(r"^\|[\s\-:|]+\|$", stripped.strip()):
                continue
            cells = [c.strip() for c in stripped.strip().strip("|").split("|")]
            # 현재 섹션의 마지막 항목이 TableItem이면 행 추가, 아니면 새 TableItem
            if current_section.items and isinstance(current_section.items[-1], TableItem):
                current_section.items[-1].rows.append(cells)
            else:
                tbl = TableItem(headers=cells, rows=[])
                current_section.items.append(tbl)
            continue

        if stripped.lstrip().startswith("※"):
            if current_section is None:
                current_section = Section(title="")
                doc.sections.append(current_section)
            text = stripped.strip()
            text = text.lstrip("※").strip()
            item = BulletItem(type="note", text=text)
            current_section.items.append(item)
            continue

        m = re.match(r"^( *)-\s+(.+)$", stripped)
        if m:
            if current_section is None:
                current_section = Section(title="")
                doc.sections.append(current_section)

            indent = len(m.group(1)) + indent_offset
            text = m.group(2).strip()

            if indent >= 6:
                # 4단계: dot
                item = BulletItem(type="dot", text=text)
                if stack_dash:
                    stack_dash.children.append(item)
                elif stack_circle:
                    stack_circle.children.append(item)
                elif stack_square:
                    stack_square.children.append(item)
                else:
                    current_section.items.append(item)
            elif indent >= 4:
                # 3단계: dash
                item = BulletItem(type="dash", text=text)
                stack_dash = item
                if stack_circle:
                    stack_circle.children.append(item)
                elif stack_square:
                    stack_square.children.append(item)
                else:
                    current_section.items.append(item)
            elif indent >= 2:
                # 2단계: circle
                item = BulletItem(type="circle", text=text)
                stack_circle = item
                stack_dash = None
                if stack_square:
                    stack_square.children.append(item)
                else:
                    current_section.items.append(item)
            else:
                # 1단계: square
                item = BulletItem(type="square", text=text)
                stack_square = item
                stack_circle = None
                stack_dash = None
                indent_offset = 0
                current_section.items.append(item)
            continue

    return doc, prefix_overrides


# ── 4. build_part2 (ET 기반) ─────────────────────────
def _clone_and_replace_slot(
    proto: ET.Element,
    new_text: str,
    slot_type: str = "",
    prefix_map: dict[str, str] | None = None,
) -> ET.Element:
    """슬롯 원형을 복제하고 기호+텍스트를 올바르게 삽입한다.

    - square: 2-run 구조. run[0] 기호(" □ "), run[1] 텍스트만 교체.
    - circle/dash/dot/note: 1-run 구조. "들여쓰기+기호 "+텍스트 합쳐서 삽입.
    """
    pmap = prefix_map or DEFAULT_PREFIX
    node = copy.deepcopy(proto)
    t_elems = _get_text_elements(node)

    if slot_type == "square":
        # 2-run: run[0]=" □ "(기호), run[1]=텍스트
        sym = pmap.get("square", "□")
        if len(t_elems) >= 2:
            t_elems[0].text = f" {sym} "
            t_elems[-1].text = new_text
        elif t_elems:
            t_elems[-1].text = f" {sym} {new_text}"
    elif slot_type in ("circle", "dash", "dot", "note"):
        # 1-run: 들여쓰기 + 기호 + 텍스트
        indent = _INDENT.get(slot_type, "")
        sym = pmap.get(slot_type, "")
        combined = f"{indent}{sym} {new_text}"
        if t_elems:
            t_elems[-1].text = combined
    else:
        # fallback
        if t_elems:
            t_elems[-1].text = new_text

    # linesegarray 제거 (한글이 자동 재계산)
    for ls in list(node.iter(_tag("hp", "linesegarray"))):
        parent = _find_parent(node, ls)
        if parent is not None:
            parent.remove(ls)
    return node


def _clone_and_replace_section_bar(proto: ET.Element, roman: str, title: str) -> ET.Element:
    """섹션바 원형을 복제하고 로마숫자+제목을 교체한다.

    섹션바 구조: <hp:p> → <hp:run> → <hp:tbl> → 내부에 hp:t 2개 (로마숫자, 제목)
    tbl 바깥의 <hp:t/>는 건드리지 않는다.
    """
    node = copy.deepcopy(proto)
    # tbl 내부의 텍스트만 찾기
    tbl_elem = node.find(f".//{_tag('hp', 'tbl')}")
    if tbl_elem is not None:
        tbl_texts = list(tbl_elem.iter(_tag("hp", "t")))
        if len(tbl_texts) >= 2:
            tbl_texts[0].text = roman
            tbl_texts[-1].text = " " + title
    # linesegarray 제거
    for ls in list(node.iter(_tag("hp", "linesegarray"))):
        parent = _find_parent(node, ls)
        if parent is not None:
            parent.remove(ls)
    return node


def _find_parent(root: ET.Element, target: ET.Element) -> Optional[ET.Element]:
    """target의 부모 요소를 찾는다."""
    for parent in root.iter():
        for child in parent:
            if child is target:
                return parent
    return None


def _flatten_items(items: list[BulletItem]) -> list[BulletItem]:
    """계층 구조를 순서 유지하며 평탄화 (4단계까지 지원)."""
    result = []
    for item in items:
        result.append(BulletItem(type=item.type, text=item.text))
        for child in item.children:
            result.append(BulletItem(type=child.type, text=child.text))
            for grandchild in child.children:
                result.append(BulletItem(type=grandchild.type, text=grandchild.text))
                for great in grandchild.children:
                    result.append(BulletItem(type=great.type, text=great.text))
    return result


# ── 테이블 XML 생성 ─────────────────────────────────
_TBL_BORDER_FILL_ID = 15       # 데이터 셀 (테두리만, 배경 없음)
_TBL_HEADER_BORDER_FILL_ID = 16  # 헤더 셀 (테두리 + 배경 D9D9D9)

# borderFill XML 삽입용 (header.xml에 추가)
_TBL_BORDER_FILLS = (
    '<hh:borderFill id="15" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
    '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
    '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
    '<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:diagonal type="SOLID" width="0.12 mm" color="#000000"/>'
    '</hh:borderFill>'
    '<hh:borderFill id="16" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
    '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
    '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
    '<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:diagonal type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hc:fillBrush><hc:winBrush faceColor="#D9D9D9" hatchColor="#000000" alpha="0"/></hc:fillBrush>'
    '</hh:borderFill>'
    '<hh:borderFill id="17" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
    '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
    '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
    '<hh:leftBorder type="NONE" width="0.12 mm" color="#000000"/>'
    '<hh:rightBorder type="NONE" width="0.12 mm" color="#000000"/>'
    '<hh:topBorder type="SOLID" width="0.4 mm" color="#000000"/>'
    '<hh:bottomBorder type="SOLID" width="0.4 mm" color="#000000"/>'
    '<hh:diagonal type="NONE" width="0.12 mm" color="#000000"/>'
    '<hc:fillBrush><hc:winBrush faceColor="#FFFFFF" hatchColor="#000000" alpha="0"/></hc:fillBrush>'
    '</hh:borderFill>'
)


def _estimate_col_widths(
    headers: list[str],
    rows: list[list[str]],
    total_width_mm: float = 170.0,
) -> list[float]:
    """내용 길이 기준 열 너비 자동 배분 (mm)."""
    ncols = len(headers)
    if ncols == 0:
        return []
    max_lens = [len(h) for h in headers]
    for row in rows:
        for i, cell in enumerate(row):
            if i < len(max_lens):
                max_lens[i] = max(max_lens[i], len(cell))
    total_len = sum(max(ml, 2) for ml in max_lens)
    if total_len == 0:
        return [total_width_mm / ncols] * ncols
    widths = []
    for ml in max_lens:
        ratio = max(ml, 2) / total_len
        widths.append(round(total_width_mm * ratio, 1))
    return widths


def _mm_to_hwp(mm: float) -> int:
    """mm → hwp unit (1mm = 283.46 hwp unit)."""
    return round(mm * 283.46)


def _pt_to_hwp(pt: float) -> int:
    """pt → hwp unit (1pt ≈ 100 hwp unit)."""
    return round(pt * 100)


_tbl_counter = 0


def _next_tbl_id() -> int:
    global _tbl_counter
    _tbl_counter += 1
    return 100000 + _tbl_counter


def _build_table_element(tbl_item: TableItem) -> ET.Element:
    """TableItem → <hp:p> 노드 (테이블 포함)."""
    headers = tbl_item.headers
    rows = tbl_item.rows
    ncols = len(headers)
    nrows = len(rows) + 1  # +1 for header row

    col_widths_mm = _estimate_col_widths(headers, rows)
    col_widths_hwp = [_mm_to_hwp(w) for w in col_widths_mm]
    total_width = sum(col_widths_hwp)

    cm_lr = _mm_to_hwp(3.0)
    cm_tb = _mm_to_hwp(2.0)
    font_height = _pt_to_hwp(10.0)
    row_height = font_height + cm_tb * 2
    total_height = row_height * nrows

    tbl_id = _next_tbl_id()
    pid = tbl_id + 500

    hp = _NS["hp"]
    hc = _NS["hc"]

    def _cell_xml(text: str, bf_id: int, align: str, bold: bool, char_pr: str) -> str:
        """셀 1개의 XML."""
        bold_attr = ' bold="1"' if bold else ""
        return (
            f'<hp:tc name="" header="0" hasMargin="1" protect="0"'
            f' editable="0" dirty="0" borderFillIDRef="{bf_id}">'
            f'<hp:subList id="" textDirection="HORIZONTAL"'
            f' lineWrap="BREAK" vertAlign="CENTER"'
            f' linkListIDRef="0" linkListNextIDRef="0"'
            f' textWidth="0" textHeight="0"'
            f' hasTextRef="0" hasNumRef="0">'
            f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
            f' pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="{char_pr}">'
            f'<hp:t>{_xml_escape(text)}</hp:t>'
            f'</hp:run>'
            f'</hp:p>'
            f'</hp:subList>'
        )

    def _cell_meta(col: int, row_idx: int, width: int) -> str:
        return (
            f'<hp:cellAddr colAddr="{col}" rowAddr="{row_idx}"/>'
            f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
            f'<hp:cellSz width="{width}" height="{row_height}"/>'
            f'<hp:cellMargin left="{cm_lr}" right="{cm_lr}"'
            f' top="{cm_tb}" bottom="{cm_tb}"/>'
            f'</hp:tc>'
        )

    parts = []
    # Wrapper <hp:p>
    parts.append(
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0"'
        f' xmlns:hp="{hp}" xmlns:hc="{hc}">'
    )
    parts.append(f'<hp:run charPrIDRef="18">')

    # <hp:tbl>
    parts.append(
        f'<hp:tbl id="{tbl_id}" zOrder="0" numberingType="TABLE"'
        f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
        f' dropcapstyle="None" pageBreak="NONE" repeatHeader="1"'
        f' rowCnt="{nrows}" colCnt="{ncols}"'
        f' cellSpacing="0" borderFillIDRef="{_TBL_BORDER_FILL_ID}" noAdjust="0">'
    )
    parts.append(
        f'<hp:sz width="{total_width}" widthRelTo="ABSOLUTE"'
        f' height="{total_height}" heightRelTo="ABSOLUTE" protect="0"/>'
    )
    parts.append(
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1"'
        f' allowOverlap="0" holdAnchorAndSO="0"'
        f' vertRelTo="PARA" horzRelTo="COLUMN"'
        f' vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
    )
    out_m = 283
    parts.append(
        f'<hp:outMargin left="{out_m}" right="{out_m}"'
        f' top="{out_m}" bottom="{out_m}"/>'
    )
    parts.append(
        f'<hp:inMargin left="{cm_lr}" right="{cm_lr}"'
        f' top="{cm_tb}" bottom="{cm_tb}"/>'
    )

    # Header row
    parts.append('<hp:tr>')
    for ci, h in enumerate(headers):
        w = col_widths_hwp[ci] if ci < len(col_widths_hwp) else col_widths_hwp[-1]
        parts.append(_cell_xml(h, _TBL_HEADER_BORDER_FILL_ID, "CENTER", True, "18"))
        parts.append(_cell_meta(ci, 0, w))
    parts.append('</hp:tr>')

    # Data rows
    for ri, row in enumerate(rows):
        parts.append('<hp:tr>')
        for ci in range(ncols):
            cell_text = row[ci] if ci < len(row) else ""
            w = col_widths_hwp[ci] if ci < len(col_widths_hwp) else col_widths_hwp[-1]
            parts.append(_cell_xml(cell_text, _TBL_BORDER_FILL_ID, "LEFT", False, "18"))
            parts.append(_cell_meta(ci, ri + 1, w))
        parts.append('</hp:tr>')

    parts.append('</hp:tbl>')
    parts.append('<hp:t/>')
    parts.append('</hp:run>')
    parts.append('</hp:p>')

    xml_str = "".join(parts)
    return ET.fromstring(xml_str)


def _xml_escape(text: str) -> str:
    """XML 특수문자 이스케이프."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def build_part2(
    doc: ParsedDocument,
    proto: Prototypes,
    prefix_map: dict[str, str] | None = None,
    purpose: str = "",
) -> list[ET.Element]:
    """파싱된 문서 + 원형으로 새 파트2 노드 리스트를 조립한다.

    파트2의 첫 번째 노드(pageBreak=1, secPr 포함)는 호출자가 별도 보존.
    여기서는 제목 블록 + 섹션바 + 슬롯만 반환.
    """
    pmap = {**DEFAULT_PREFIX, **(prefix_map or {})}
    nodes: list[ET.Element] = []

    # 본문 제목 블록 (빈 단락 없이 바로 시작)
    title_node = copy.deepcopy(proto.title_block_p)
    title_node.set("pageBreak", "0")  # pageBreak는 first_p가 담당
    for t in _get_text_elements(title_node):
        if t.text and "제 목" in t.text:
            t.text = t.text.replace("제 목", doc.title)
    for ls in list(title_node.iter(_tag("hp", "linesegarray"))):
        parent = _find_parent(title_node, ls)
        if parent is not None:
            parent.remove(ls)
    nodes.append(title_node)

    # 목적 블록 (첫 섹션바 앞에 삽입)
    if purpose:
        nodes.append(_build_purpose_element(purpose))

    # 슬롯 타입 → 원형 매핑
    _slot_map = {
        "square": proto.slot_square,
        "circle": proto.slot_circle,
        "dash": proto.slot_dash,
        "dot": proto.slot_dot,
        "note": proto.slot_note,
    }

    # 각 섹션
    for i, section in enumerate(doc.sections):
        roman = _ROMAN[i] if i < len(_ROMAN) else str(i + 1)

        # 빈 단락
        if proto.empty_para is not None:
            nodes.append(copy.deepcopy(proto.empty_para))

        # 섹션바
        bar_node = _clone_and_replace_section_bar(proto.section_bar_p, roman, section.title)
        nodes.append(bar_node)

        # 슬롯/테이블 단락
        # items에는 BulletItem과 TableItem이 혼재할 수 있음
        bullet_items = []
        for item in section.items:
            if isinstance(item, TableItem):
                # 쌓인 bullet items 먼저 처리
                if bullet_items:
                    for bi in _flatten_items(bullet_items):
                        slot_proto = _slot_map.get(bi.type)
                        if slot_proto is not None:
                            nodes.append(_clone_and_replace_slot(
                                slot_proto, bi.text,
                                slot_type=bi.type, prefix_map=pmap,
                            ))
                    bullet_items = []
                # 테이블 삽입
                nodes.append(_build_table_element(item))
            else:
                bullet_items.append(item)
        # 남은 bullet items
        if bullet_items:
            for bi in _flatten_items(bullet_items):
                slot_proto = _slot_map.get(bi.type)
                if slot_proto is not None:
                    nodes.append(_clone_and_replace_slot(
                        slot_proto, bi.text,
                        slot_type=bi.type, prefix_map=pmap,
                    ))

    return nodes


# ── 5. apply_colors ──────────────────────────────────
def _lighten_color(hex_color: str, factor: float = 0.4) -> str:
    """hex 색상을 밝게 만든다. factor=0이면 원색, 1이면 흰색."""
    c = hex_color.lstrip("#")
    r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return f"{r:02X}{g:02X}{b:02X}"


def apply_colors(
    header_xml: str,
    color_main: str = "1F4E79",
    color_sub: str = "2E75B6",
    color_accent: str = "C00000",
) -> str:
    """header.xml의 borderFill/charPr 색상을 오버라이드 + 섹션바 그라데이션 적용."""
    result = header_xml

    # 기존 색상 치환 (표지 borderFill 7,8 등)
    result = result.replace('faceColor="#193AAA"', f'faceColor="#{color_main}"')
    result = result.replace('color="#006699"', f'color="#{color_main}"')

    # charPr ID=20 textColor → color_accent
    pattern = r'(<hh:charPr [^>]*id="20"[^>]*textColor=")#[0-9A-Fa-f]+(")'
    result = re.sub(pattern, rf'\g<1>#{color_accent}\2', result)

    # 섹션바 그라데이션 적용
    # borderFill 9 (좌측 셀): 단색 → 좌→우 그라데이션 (color_main → 밝은 계열)
    color_main_light = _lighten_color(color_main, 0.35)
    grad_left = (
        f'<hc:gradation type="LINEAR" angle="90" centerX="0" centerY="0"'
        f' step="250" colorNum="2" stepCenter="50" alpha="0">'
        f'<hc:color value="#{color_main}"/><hc:color value="#{color_main_light}"/>'
        f'</hc:gradation>'
    )
    # borderFill 10 (우측 셀): 단색 → 연한 회색→흰색 그라데이션
    grad_right = (
        '<hc:gradation type="LINEAR" angle="90" centerX="0" centerY="0"'
        ' step="250" colorNum="2" stepCenter="50" alpha="0">'
        '<hc:color value="#E8E8E8"/><hc:color value="#FFFFFF"/>'
        '</hc:gradation>'
    )

    # borderFill 9: winBrush → gradation
    result = re.sub(
        r'(<hh:borderFill id="9"[^>]*>.*?<hc:fillBrush>)'
        r'<hc:winBrush[^/]*/>'
        r'(</hc:fillBrush>)',
        rf'\g<1>{grad_left}\2',
        result,
        flags=re.DOTALL,
    )

    # borderFill 10: winBrush → gradation
    result = re.sub(
        r'(<hh:borderFill id="10"[^>]*>.*?<hc:fillBrush>)'
        r'<hc:winBrush[^/]*/>'
        r'(</hc:fillBrush>)',
        rf'\g<1>{grad_right}\2',
        result,
        flags=re.DOTALL,
    )

    return result


def _inject_table_border_fills(header_xml: str) -> str:
    """header.xml에 테이블/목적 블록용 borderFill 엔트리를 추가한다 (ID 15, 16, 17)."""
    if 'borderFill id="15"' in header_xml:
        return header_xml  # 이미 존재
    # </hh:borderFills> 앞에 삽입
    result = header_xml.replace(
        "</hh:borderFills>",
        _TBL_BORDER_FILLS + "</hh:borderFills>",
    )
    # itemCnt 업데이트 (+3)
    result = re.sub(
        r'(<hh:borderFills\s+itemCnt=")(\d+)(")',
        lambda m: f'{m.group(1)}{int(m.group(2)) + 3}{m.group(3)}',
        result,
    )
    return result


# ── 6. replace_cover_text (ET 기반) ──────────────────
def replace_cover_text(
    part0: list[ET.Element],
    title: str,
    org: str,
    date_str: str,
) -> None:
    """표지(파트0) 노드의 텍스트를 치환한다. (in-place)"""
    for p in part0:
        for t in _get_text_elements(p):
            if t.text:
                t.text = t.text.replace("브라더 공기관", org)
                t.text = t.text.replace("기본 보고서 양식", title)
                t.text = t.text.replace("2024. 5. 23.", date_str)


_PURPOSE_BORDER_FILL_ID = 17  # 상하 SOLID 0.5mm, 좌우 NONE


def _build_purpose_element(purpose: str) -> ET.Element:
    """목적 블록(1x1 테두리 테이블) ET 노드를 생성한다."""
    hp = _NS["hp"]
    hc = _NS["hc"]

    tbl_w = 48180
    cm_lr = _mm_to_hwp(5.0)
    cm_tb = _mm_to_hwp(3.0)
    row_h = _mm_to_hwp(12.0)
    pid = 9999000

    purpose_xml = (
        f'<hp:p xmlns:hp="{hp}" xmlns:hc="{hc}"'
        f' id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="5">'
        f'<hp:tbl id="{pid+1}" zOrder="0" numberingType="TABLE"'
        f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
        f' dropcapstyle="None" pageBreak="NONE" repeatHeader="0"'
        f' rowCnt="1" colCnt="1"'
        f' cellSpacing="0" borderFillIDRef="{_PURPOSE_BORDER_FILL_ID}" noAdjust="0">'
        f'<hp:sz width="{tbl_w}" widthRelTo="ABSOLUTE"'
        f' height="{row_h}" heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1"'
        f' allowOverlap="0" holdAnchorAndSO="0"'
        f' vertRelTo="PARA" horzRelTo="COLUMN"'
        f' vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
        f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
        f'<hp:inMargin left="{cm_lr}" right="{cm_lr}"'
        f' top="{cm_tb}" bottom="{cm_tb}"/>'
        f'<hp:tr>'
        f'<hp:tc name="" header="0" hasMargin="1" protect="0"'
        f' editable="0" dirty="0" borderFillIDRef="{_PURPOSE_BORDER_FILL_ID}">'
        f'<hp:subList id="" textDirection="HORIZONTAL"'
        f' lineWrap="BREAK" vertAlign="CENTER"'
        f' linkListIDRef="0" linkListNextIDRef="0"'
        f' textWidth="0" textHeight="0"'
        f' hasTextRef="0" hasNumRef="0">'
        f'<hp:p id="{pid+2}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="5">'
        f'<hp:t>{_xml_escape(purpose)}</hp:t>'
        f'</hp:run>'
        f'</hp:p>'
        f'</hp:subList>'
        f'<hp:cellAddr colAddr="0" rowAddr="0"/>'
        f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
        f'<hp:cellSz width="{tbl_w}" height="{row_h}"/>'
        f'<hp:cellMargin left="{cm_lr}" right="{cm_lr}"'
        f' top="{cm_tb}" bottom="{cm_tb}"/>'
        f'</hp:tc>'
        f'</hp:tr>'
        f'</hp:tbl>'
        f'<hp:t/>'
        f'</hp:run>'
        f'</hp:p>'
    )

    return ET.fromstring(purpose_xml)


def replace_cover_decoration(part0: list[ET.Element]) -> None:
    """표지 테이블 Row0/Row2 색상 바를 제거하고, 테이블 앞/뒤에 장식선 <hp:p>를 삽입한다.

    Row0/Row2 배경만 제거하고, 장식선은 테이블 밖 독립 단락으로 삽입하여
    셀 폭 제약 없이 본문 영역 전체 폭(48190)을 사용한다.
    """
    hp = _NS["hp"]
    hc = _NS["hc"]

    # 표지 테이블을 포함한 단락(p) 찾기
    tbl_p_idx = None
    tbl = None
    for i, p in enumerate(part0):
        t = p.find(f".//{_tag('hp', 'tbl')}")
        if t is not None:
            tbl_p_idx = i
            tbl = t
            break
    if tbl is None:
        return

    rows = list(tbl.iter(_tag("hp", "tr")))
    if len(rows) < 3:
        return

    # Row0/Row2 배경 제거
    for ri in (0, 2):
        for tc in rows[ri].iter(f"{{{hp}}}tc"):
            tc.set("borderFillIDRef", "1")
            for fb in list(tc.iter(f"{{{hc}}}fillBrush")):
                parent = _find_parent(tc, fb)
                if parent is not None:
                    parent.remove(fb)

    # 장식선 단락 생성 (테이블 밖, 본문 영역 폭)
    content_w = 48190  # pagePr width(59528) - margin left(5669) - margin right(5669)
    top_p = _build_deco_paragraph("image2", content_w, hp, hc)
    bot_p = _build_deco_paragraph("image3", content_w, hp, hc)

    # 테이블 단락 앞에 상단 장식선, 뒤에 하단 장식선 삽입
    part0.insert(tbl_p_idx, top_p)
    # tbl_p_idx가 1 밀렸으므로 +2
    part0.insert(tbl_p_idx + 2, bot_p)


def _build_deco_paragraph(bin_ref: str, content_w: int, hp: str, hc: str) -> ET.Element:
    """장식선 이미지를 포함한 독립 <hp:p> 단락을 생성한다."""
    img_w = content_w
    img_h = 708  # 약 2.5mm
    dim_w = 59000  # 원본 PNG 픽셀폭 * 75 근사
    dim_h = 750    # 원본 PNG 픽셀높이 * 75 근사
    cx = img_w // 2
    cy = img_h // 2

    p_xml = (
        f'<hp:p xmlns:hp="{hp}" xmlns:hc="{hc}"'
        f' id="0" paraPrIDRef="31" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="0">'
        f'<hp:pic id="0" zOrder="0" numberingType="PICTURE"'
        f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
        f' dropcapstyle="None" href=";0;0;0;" groupLevel="0" instid="0" reverse="0">'
        f'<hp:offset x="0" y="0"/>'
        f'<hp:orgSz width="{img_w}" height="{img_h}"/>'
        f'<hp:curSz width="0" height="0"/>'
        f'<hp:flip horizontal="0" vertical="0"/>'
        f'<hp:rotationInfo angle="0" centerX="{cx}" centerY="{cy}" rotateimage="1"/>'
        f'<hp:renderingInfo>'
        f'<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        f'<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        f'<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        f'</hp:renderingInfo>'
        f'<hp:imgRect>'
        f'<hc:pt0 x="0" y="0"/><hc:pt1 x="{img_w}" y="0"/>'
        f'<hc:pt2 x="{img_w}" y="{img_h}"/><hc:pt3 x="0" y="{img_h}"/>'
        f'</hp:imgRect>'
        f'<hp:imgClip left="0" right="{dim_w}" top="0" bottom="{dim_h}"/>'
        f'<hp:inMargin left="0" right="0" top="0" bottom="0"/>'
        f'<hp:imgDim dimwidth="{dim_w}" dimheight="{dim_h}"/>'
        f'<hc:img binaryItemIDRef="{bin_ref}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>'
        f'<hp:effects/>'
        f'<hp:sz width="{img_w}" widthRelTo="ABSOLUTE" height="{img_h}" heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0"'
        f' holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN"'
        f' vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
        f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
        f'</hp:pic>'
        f'<hp:t/>'
        f'</hp:run>'
        f'</hp:p>'
    )
    return ET.fromstring(p_xml)


# ── 7. replace_toc_text ─────────────────────────────
def replace_toc_text(part1: list[ET.Element], title: str) -> None:
    """목차(파트1) 노드의 제목 텍스트를 치환한다. (in-place)"""
    for p in part1:
        for t in _get_text_elements(p):
            if t.text and "제 목" in t.text:
                t.text = t.text.replace("제 목", title)


# ── 통합 실행 ────────────────────────────────────────
def run(
    input_path: str,
    output_path: str,
    org: str = "",
    color_main: str = "1F4E79",
    color_sub: str = "2E75B6",
    color_accent: str = "C00000",
) -> None:
    """외부 스타일 동적 변환 실행."""
    if not TEMPLATE.exists():
        print(f"Error: 템플릿 파일을 찾을 수 없습니다: {TEMPLATE}", file=sys.stderr)
        sys.exit(1)

    _register_namespaces()

    # 1. 마크다운 파싱
    doc, prefix_overrides = parse_markdown_sections(input_path)
    prefix_map = {**DEFAULT_PREFIX, **prefix_overrides} if prefix_overrides else None
    today = date.today()
    date_str = f"{today.year}. {today.month}. {today.day}."

    # 2. 템플릿 읽기
    shutil.copy2(str(TEMPLATE), output_path)
    with zipfile.ZipFile(output_path, "r") as zin:
        names = zin.namelist()
        contents: dict[str, bytes] = {}
        for name in names:
            contents[name] = zin.read(name)

    section0_xml = contents["Contents/section0.xml"].decode("utf-8")
    header_xml = contents["Contents/header.xml"].decode("utf-8")

    # 3. ET 파싱
    root = ET.fromstring(section0_xml)

    # 4. 3분할
    part0, part1, part2 = split_sections(root)

    # 5. 원형 추출
    proto = extract_prototypes(part2)

    # 6. 표지/목차 텍스트 치환 + 장식선 교체
    replace_cover_text(part0, doc.title, org, date_str)
    replace_cover_decoration(part0)
    replace_toc_text(part1, doc.title)

    # 7. 새 파트2 조립
    # 파트2[0]은 pageBreak="1" + secPr 포함 단락 → 보존
    first_p = part2[0]
    # pageBreak="1" 보존 (part1→part2 경계에 필요)
    # 제목 테이블이 이 안에 있으면 제거 (별도로 재생성하므로)
    for run_elem in list(first_p.iter(_tag("hp", "run"))):
        tbls_in_run = list(run_elem.iter(_tag("hp", "tbl")))
        if tbls_in_run:
            for tbl in tbls_in_run:
                run_elem.remove(tbl)
    # linesegarray 제거
    for ls in list(first_p.iter(_tag("hp", "linesegarray"))):
        parent = _find_parent(first_p, ls)
        if parent is not None:
            parent.remove(ls)

    new_body_nodes = build_part2(doc, proto, prefix_map=prefix_map, purpose=doc.purpose)

    # 8. 루트 재구성
    for child in list(root):
        root.remove(child)

    # 파트0 + 파트1 + 파트2[0](secPr) + 새 본문 추가
    for node in part0:
        root.append(node)
    for node in part1:
        root.append(node)
    root.append(first_p)
    for node in new_body_nodes:
        root.append(node)

    # 9. XML 직렬화
    new_section0 = ET.tostring(root, encoding="unicode", xml_declaration=False)
    new_section0 = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>' + new_section0

    # 10. 색상 적용 (header.xml은 문자열 치환이 안전)
    header_xml = apply_colors(header_xml, color_main, color_sub, color_accent)
    header_xml = _inject_table_border_fills(header_xml)

    # 11. 장식선 PNG 추가 및 content.hpf 업데이트
    if _DECO_TOP_PNG.exists() and _DECO_BOTTOM_PNG.exists():
        contents["BinData/image2.png"] = _DECO_TOP_PNG.read_bytes()
        contents["BinData/image3.png"] = _DECO_BOTTOM_PNG.read_bytes()
        if "BinData/image2.png" not in names:
            names.append("BinData/image2.png")
        if "BinData/image3.png" not in names:
            names.append("BinData/image3.png")
        # content.hpf에 BinData 항목 추가
        if "Contents/content.hpf" in contents:
            hpf = contents["Contents/content.hpf"].decode("utf-8")
            if 'id="image2"' not in hpf:
                hpf = hpf.replace(
                    "</opf:manifest>",
                    '<opf:item id="image2" href="BinData/image2.png" media-type="image/png" isEmbeded="1"/>'
                    '<opf:item id="image3" href="BinData/image3.png" media-type="image/png" isEmbeded="1"/>'
                    "</opf:manifest>",
                )
                contents["Contents/content.hpf"] = hpf.encode("utf-8")

    # 12. ZIP 재작성
    contents["Contents/section0.xml"] = new_section0.encode("utf-8")
    contents["Contents/header.xml"] = header_xml.encode("utf-8")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, contents[name])

    # 12. 네임스페이스 후처리
    if FIX_NS.exists():
        subprocess.run([sys.executable, str(FIX_NS), output_path], check=True)

    abs_out = str(Path(output_path).resolve())
    print(f"Success: {abs_out}")
