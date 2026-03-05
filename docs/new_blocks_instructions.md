# hwpx_generator 신규 블록 타입 구현 지시문
> 실제 한글 파일(도형.hwpx) 분석 기반

---

## 구현할 블록 4종

### 1. doc_title (전체 제목)
장식선(이미지) + 제목 텍스트 + 장식선(이미지) 구조

**실측 구조 (도형.hwpx 기준):**
```
hp:p
  └── hp:run → hp:pic (image1 — 상단 장식선, width=47907, height=708)
  └── hp:linesegarray

hp:p  
  └── hp:run → hp:pic (image2 — 하단 장식선, width=47907, height=708)
  └── 제목 텍스트는 이미지 위에 겹쳐서 배치 또는 별도 단락
  └── hp:linesegarray
```

**장식 이미지 파일:**
- 상단: `assets/doc_title_top.png` (786×10px)
- 하단: `assets/doc_title_bottom.png` (787×10px)
- BinData에 패키징해서 binaryItemIDRef로 참조

**JSON 입력:**
```json
{ "type": "doc_title", "text": "모니터링 개요" }
```

**구현 위치:** `elements/doc_title.py`, `builders/xml_builder.py`

---

### 2. doc_purpose (전체 목적)
1행1열 표 안에 텍스트 배치

**실측 구조 (도형.hwpx 기준):**
```xml
<hp:tbl rowCnt="1" colCnt="1" borderFillIDRef="6" pageBreak="CELL">
  <hp:sz width="47624" widthRelTo="ABSOLUTE" height="4228"/>
  <hp:pos treatAsChar="0" flowWithText="1" ...
         vertRelTo="PARA" horzRelTo="COLUMN"
         vertAlign="TOP" horzAlign="LEFT"/>
  <hp:outMargin left="283" right="283" top="283" bottom="283"/>
  <hp:inMargin left="510" right="510" top="141" bottom="141"/>
  <hp:tr>
    <hp:tc borderFillIDRef="7">
      <hp:subList vertAlign="CENTER">
        <hp:p>
          <hp:run><hp:t>텍스트 내용</hp:t></hp:run>
          <hp:linesegarray>...</hp:linesegarray>
        </hp:p>
      </hp:subList>
      <hp:cellAddr colAddr="0" rowAddr="0"/>
      <hp:cellSpan colSpan="1" rowSpan="1"/>
      <hp:cellSz width="47624" height="4228"/>
      <hp:cellMargin left="510" right="510" top="141" bottom="141"/>
    </hp:tc>
  </hp:tr>
</hp:tbl>
```

**JSON 입력:**
```json
{ "type": "doc_purpose", "text": "신규 수행기관의 사업 운영 적정성을 확보하고..." }
```

**구현 위치:** `elements/doc_purpose.py`, `builders/xml_builder.py`

---

### 3. section (단락나눔 — 아라비아 숫자 자동증가)
2행3열 표 구조. 1열=번호(Ⅰ→아라비아숫자로 변경), 2열=구분선, 3열=제목

**실측 구조 (도형.hwpx 기준):**
```xml
<hp:tbl rowCnt="2" colCnt="3" borderFillIDRef="4" pageBreak="NONE">
  <hp:sz width="30670" widthRelTo="ABSOLUTE" height="2715"/>
  <hp:pos treatAsChar="1" flowWithText="1" .../>  ← treatAsChar="1" 중요
  <hp:outMargin left="0" right="0" top="0" bottom="0"/>
  <hp:inMargin left="0" right="0" top="0" bottom="0"/>
  <hp:tr>
    <!-- 1열: 번호 (2행 병합) -->
    <hp:tc borderFillIDRef="8">
      <hp:subList vertAlign="CENTER">
        <hp:p><hp:run><hp:t>1</hp:t></hp:run>...</hp:p>
      </hp:subList>
      <hp:cellAddr colAddr="0" rowAddr="0"/>
      <hp:cellSpan colSpan="1" rowSpan="2"/>  ← 2행 병합
      <hp:cellSz width="2850" height="2715"/>
      <hp:cellMargin left="0" right="0" top="0" bottom="0"/>
    </hp:tc>
    <!-- 2열: 구분선 (2행 병합) -->
    <hp:tc borderFillIDRef="9">
      <hp:subList vertAlign="CENTER">
        <hp:p><hp:run/><hp:linesegarray>...</hp:linesegarray></hp:p>
      </hp:subList>
      <hp:cellAddr colAddr="1" rowAddr="0"/>
      <hp:cellSpan colSpan="1" rowSpan="2"/>  ← 2행 병합
      <hp:cellSz width="565" height="2715"/>
      <hp:cellMargin left="0" right="0" top="0" bottom="0"/>
    </hp:tc>
    <!-- 3열 1행: 제목 텍스트 -->
    <hp:tc borderFillIDRef="10">
      <hp:subList vertAlign="CENTER">
        <hp:p><hp:run><hp:t>모니터링 개요</hp:t></hp:run>...</hp:p>
      </hp:subList>
      <hp:cellAddr colAddr="2" rowAddr="0"/>
      <hp:cellSpan colSpan="1" rowSpan="1"/>
      <hp:cellSz width="27255" height="2565"/>
      <hp:cellMargin left="565" right="0" top="140" bottom="0"/>
    </hp:tc>
  </hp:tr>
  <hp:tr>
    <!-- 3열 2행: 얇은 구분선 역할 (height=100) -->
    <hp:tc borderFillIDRef="11">
      <hp:subList vertAlign="CENTER">
        <hp:p><hp:run/><hp:linesegarray>...</hp:linesegarray></hp:p>
      </hp:subList>
      <hp:cellAddr colAddr="2" rowAddr="1"/>
      <hp:cellSpan colSpan="1" rowSpan="1"/>
      <hp:cellSz width="27255" height="100"/>  ← 얇은 구분선
      <hp:cellMargin left="0" right="0" top="0" bottom="0"/>
    </hp:tc>
  </hp:tr>
</hp:tbl>
```

**JSON 입력 (번호 자동증가):**
```json
{ "type": "section", "text": "모니터링 개요" }
{ "type": "section", "text": "기관별 방문 일정" }
```

**번호 자동증가 로직:**
- 문서 내 section 블록 등장 순서대로 1, 2, 3... 자동 부여
- 수동 지정도 가능: `{ "type": "section", "num": 3, "text": "..." }`

**구현 위치:** `elements/section.py`, `builders/xml_builder.py`

---

### 4. bullet (불릿) — 기존 그대로
변경 없음.

---

## borderFill 참조 구조

도형.hwpx header.xml에서 실측한 borderFill ID 매핑:
- `borderFillIDRef="4"` — section 표 외곽선
- `borderFillIDRef="6"` — doc_purpose 표 외곽선  
- `borderFillIDRef="7"` — doc_purpose 셀 내부
- `borderFillIDRef="8~11"` — section 각 셀

> ⚠️ borderFill 정의는 header.xml의 hh:borderFills에 추가 필요
> 도형.hwpx의 header.xml에서 해당 borderFill 항목 그대로 복사

---

## 구현 순서

1. `assets/` 폴더 생성, 장식 이미지 2개 포함
2. `elements/doc_title.py` 구현
3. `elements/doc_purpose.py` 구현  
4. `elements/section.py` 구현 (번호 자동증가 포함)
5. `builders/xml_builder.py` 에 3개 블록 렌더러 추가
6. `builders/package_builder.py` 에 assets 이미지 BinData 패키징 추가
7. `presets/gov_document.py` 에 관련 스타일 추가
8. 최소 테스트 파일 생성 → 한글에서 열기 확인
9. test_notice.json에 새 블록 추가 → 전체 테스트

---

## 테스트용 JSON

```json
{
  "preset": "gov_document",
  "content": [
    { "type": "doc_title", "text": "모니터링 개요" },
    { "type": "doc_purpose", "text": "신규 수행기관의 사업 운영 적정성을 확보하고, 2026년 개정 지침의 이행 여부 확인 및 현장 모니터링을 통한 운영 안정화 지원함" },
    { "type": "section", "text": "모니터링 개요" },
    { "type": "bullet", "level": "li1", "text": "일시: 2026. 3. 9.(월) ~ 3. 13.(금)" },
    { "type": "bullet", "level": "li1", "text": "대상: 2026년 신규 지정 지역수행기관 5개소" },
    { "type": "section", "text": "기관별 방문 일정" },
    { "type": "bullet", "level": "li1", "text": "방문 일정은 별도 공문 참조" }
  ]
}
```

---

## 레퍼런스 파일

- `도형.hwpx` — 한글에서 직접 생성한 기준 파일
- `assets/doc_title_top.png` — 상단 장식선 (786×10px)
- `assets/doc_title_bottom.png` — 하단 장식선 (787×10px)

추측으로 XML 생성하지 말고, 도형.hwpx의 실제 XML 구조를 그대로 참고해서 구현해줘.
구현할 때마다 최소 테스트 파일 생성 → 한글에서 열어서 확인하는 방식으로 진행해줘.
