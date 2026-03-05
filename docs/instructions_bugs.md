# hwpx_generator 버그 수정 지시

sample_notice.hwpx 한글 확인 결과 발견된 문제 6가지 수정.

---

## 문제 1. DocTitle — tbl(목적 표)이 run[1]에서 누락됨

### 현재 상태
run[1] 자식 순서: pic(6968) → pic(0) → t
tbl이 없음 → 목적 텍스트가 표 안에 안 들어가 있음

### 원인
md_loader에서 `>` (blockquote) → doc_purpose 변환이 아니라
text_box로 처리되고 있음.

### 수정 1-A: markdown_parser.py
`>` blockquote를 text_box 대신 doc_purpose로 매핑 변경:

```python
# 변경 전
m = _RE_BLOCKQUOTE.match(stripped)
if m:
    ...
    doc.add_text_box(box_text)

# 변경 후
m = _RE_BLOCKQUOTE.match(stripped)
if m:
    ...
    doc.add_doc_purpose(box_text)
```

### 수정 1-B: _render_doc_title_v7() (xml_builder.py)
DocTitle 렌더링 시 purpose_block이 None이 아닐 때만 tbl 삽입.
현재 purpose_block을 별도 인자로 받는지 확인 후,
build_section_xml()에서 DocTitle 다음에 오는 DocPurpose를
DocTitle 렌더링 시 같이 전달하도록 수정.

---

## 문제 2. 제목 텍스트가 두 줄로 나뉘며 선도형이 중간에 끼어듦

### 원인
image2(아래 선도형) vertOffset=0 → 제목 첫 줄 위에 붙어야 하는데
제목이 2줄일 때 줄 사이에 끼어드는 현상.

### 수정
image2 vertOffset을 0 → -708 (선도형 높이만큼 위로 올림) 으로 변경.
또는 image2를 run[2](제목 텍스트 run) 앞에 배치하는 방식 검토.

v7 실측값 재확인: test_notice_v7.hwpx에서
image2 vertOffset=0 이 맞는지 확인 후 적용.

---

## 문제 3. 본문 폰트가 휴먼명조여야 함 (현재 맑은 고딕으로 출력)

md_loader에서 preset을 적용하지 않고 생성하면
기본 폰트가 맑은 고딕으로 설정됨.

### 수정: load_from_md_file() 기본 preset 적용

```python
def load_from_md_file(path, preset_name=None):
    ...
    # preset_name이 없으면 gov_document 기본 적용
    if preset_name is None:
        preset_name = "gov_document"
    return load_from_markdown(text, preset_name=preset_name)
```

---

## 문제 4. 표 테두리선이 굵게 표시됨

현재 border_width 기본값이 너무 굵음.
표 일반 테두리: "0.4mm" → "0.12mm" 로 변경.
목적 텍스트 표(DocPurpose): 이중선(DOUBLE) 적용 필요.

### 수정: _render_purpose_table_float() 테두리를 이중선으로

v7 실측: 목적 표 borderFill → type="DOUBLE", width="0.12mm"

```python
# 목적 표 border 설정
BorderFillProps(
    border_type="DOUBLE",
    border_width="0.12 mm",
    border_color="000000",
)
```

---

## 문제 5. 기관별 방문 일정 표가 겹쳐 보임

표 col_widths 자동 계산 시 총합이 content_width_mm을 초과함.

### 수정: _estimate_col_widths() total_width_mm 조정

```python
# 현재
total_width_mm=doc.page_settings.content_width_mm

# 수정: out_margin(2mm) 고려해서 약간 줄임
total_width_mm=doc.page_settings.content_width_mm - 2.0
```

---

## 문제 6. SectionBlock 제목에 아라비아 숫자가 붙음

### 현재 상태
md에서 `## 1. 모니터링 개요` → SectionBlock 텍스트가 "1. 모니터링 개요" 로 됨
로마숫자 셀(Ⅰ)과 별도로 아라비아 숫자 "1."이 제목 텍스트에 포함됨.

### 수정: markdown_parser.py _RE_SECTION 처리

```python
m = _RE_SECTION.match(stripped)
if m:
    heading_text = m.group(1).strip()
    # 앞의 숫자 제거: "1. 텍스트" → "텍스트"
    heading_text = re.sub(r'^\d+\.\s*', '', heading_text)
    doc.add_section_block(heading_text)
```

---

## 완료 후 검증

```bash
python3 -c "
from hwpx_generator.md_loader import load_from_md_file
doc = load_from_md_file('sample_notice.md')
doc.save('sample_notice_fixed.hwpx')
print('done')
"
```

한글에서 확인:
- [ ] 선도형 2개 사이에 제목 텍스트 정상 배치
- [ ] 목적 텍스트가 이중선 표 안에 맑은 고딕으로 표시
- [ ] 본문 텍스트 휴먼명조
- [ ] 표 테두리 얇게 (0.12mm)
- [ ] 기관별 방문 일정 표 겹침 없음
- [ ] 로마숫자 섹션 제목에 아라비아 숫자 없음 (Ⅰ 모니터링 개요)
