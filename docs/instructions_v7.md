# hwpx_generator v7 수정 지시

레퍼런스: test_notice_v7.hwpx 직접 XML 분석 결과 기반

---

## 수정 1. DocTitle 구조 변경 (xml_builder.py — _render_doc_title)

### 현재 문제
- 제목 위에 빈 공간이 생김
- 선도형(image1) 위치가 잘못됨

### v7 실측 구조

단일 hp:p 안에 아래 순서로 모두 포함:

```
run[0]: secPr + colPr (섹션 헤더)
run[1]:
  ├── hp:pic (image1 — 위쪽 선도형)
  │     width=48189, height=708
  │     treatAsChar=0, vertRelTo=PARA, horzRelTo=COLUMN
  │     vertOffset=6968, horzOffset=0   ← 제목 텍스트 아래쪽에 float
  ├── hp:tbl (목적 텍스트 1x1 표)
  │     width=47623, height=4228
  │     treatAsChar=0
  │     vertOffset=6976, horzOffset=0   ← 위 선도형 바로 아래 붙음
  │     inMargin left=510 right=510 top=141 bottom=141
  │     셀 내부: charPrIDRef=13, 줄간격 780(spacing)
  └── hp:pic (image2 — 아래쪽 선도형)
        width=48189, height=708
        treatAsChar=0
        vertOffset=0, horzOffset=0      ← 제목 텍스트 맨 위에 붙음
run[2]: 제목 텍스트 (charPrIDRef=1, HY헤드라인M 27pt)
  <hp:t>2026년 경상남도 노인맞춤돌봄서비스</hp:t>

lineseg: vertsize=1300, spacing=780
paraPrIDRef=10 (줄간격 120% 전용)
```

### 핵심 포인트
- image2(아래 선도형) vertOffset=0 → 제목 텍스트 위에 바로 붙음
- image1(위 선도형) vertOffset=6968 → 제목 텍스트 아래쪽에 float
- 목적 표 vertOffset=6976 → 선도형1 바로 아래 붙음
- 세 요소 모두 같은 run 안, treatAsChar=0, flowWithText=1

### paraPr id=10 정의 (header.xml styles_builder에 추가)

```xml
<hh:paraPr id="10" tabPrIDRef="0" condense="0" fontLineHeight="0"
           snapToGrid="1" suppressLineNumbers="0" checked="0">
  <hh:align horizontal="CENTER" vertical="BASELINE"/>
  <hh:heading type="NONE" idRef="0" level="0"/>
  <hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD"
                   widowOrphan="0" keepWithNext="0" keepLines="0"
                   pageBreakBefore="0" lineWrap="BREAK"/>
  <hh:autoSpacing eAsianEng="0" eAsianNum="0"/>
  <hp:switch>
    <hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">
      <hh:margin>
        <hc:intent value="0" unit="HWPUNIT"/>
        <hc:left value="0" unit="HWPUNIT"/>
        <hc:right value="0" unit="HWPUNIT"/>
        <hc:prev value="0" unit="HWPUNIT"/>
        <hc:next value="0" unit="HWPUNIT"/>
      </hh:margin>
      <hh:lineSpacing type="PERCENT" value="120" unit="HWPUNIT"/>
    </hp:case>
    <hp:default>
      <hh:margin>
        <hc:intent value="0" unit="HWPUNIT"/>
        <hc:left value="0" unit="HWPUNIT"/>
        <hc:right value="0" unit="HWPUNIT"/>
        <hc:prev value="0" unit="HWPUNIT"/>
        <hc:next value="0" unit="HWPUNIT"/>
      </hh:margin>
      <hh:lineSpacing type="PERCENT" value="120" unit="HWPUNIT"/>
    </hp:default>
  </hp:switch>
  <hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0"
             offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>
</hh:paraPr>
```

---

## 수정 2. 본문 줄간격 160 유지

본문(휴먼명조 13pt) 줄간격은 160% 그대로.
GovDocumentPreset line_spacing 기본값 변경 없음.
줄간격 120은 제목 단락(paraPr id=10)에만 적용.

---

## 수정 3. add_bullet1() 동그라미 앞 공백 1칸 (document.py)

```python
# 변경 전
self._add_bullet_paragraph(bs, text, prefix_text=bs.prefix)
# bs.prefix = "\uf06d "

# 변경 후
self._add_bullet_paragraph(bs, text, prefix_text=" \uf06d ")
```

---

## 수정 4. add_bullet2() 대시 앞 공백 2칸 (document.py)

```python
# 변경 전
para.add_run("- " + text, ...)

# 변경 후
para.add_run("  - " + text, ...)
```

---

## 수정 5. SectionBlock 앞 빈 단락 삽입 (xml_builder.py)

build_section_xml() 블록 루프 수정:

```python
# 변경 전
for block in doc.blocks:
    lines.append(_render_block(block, sm, content_w))

# 변경 후
for i, block in enumerate(doc.blocks):
    if isinstance(block, SectionBlock) and i > 0:
        lines.append(_empty_paragraph(content_w))
    lines.append(_render_block(block, sm, content_w))
```

---

## 검증 체크리스트

- [ ] 문서 맨 위가 선도형으로 시작 (빈 줄 없음)
- [ ] 제목 줄간격 120% (촘촘하게)
- [ ] 아래 선도형이 제목에 바로 붙음
- [ ] 목적 텍스트 표가 아래 선도형 바로 아래 (공백 없음)
- [ ] 본문 줄간격 160% 유지
- [ ] 동그라미 불릿 왼쪽에서 1칸
- [ ] 대시 불릿 왼쪽에서 2칸
- [ ] 각 로마숫자 섹션 앞 빈 줄 1개
