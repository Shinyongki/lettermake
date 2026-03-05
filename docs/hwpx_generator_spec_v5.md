# HWPX 공문서 자동 생성기 — PRD / IA / Use Case

> 작성일: 2026-03-01  
> 작성자: 경상남도사회서비스원 광역지원기관  
> 대상 AI: Claude Code (또는 동급 코드 에이전트)

---

## 1. PRD (Product Requirements Document)

### 1-1. 배경 및 목적

경상남도사회서비스원 노인맞춤돌봄서비스 광역지원기관에서는 매년 수십 건의 공문서(모니터링 안내서, 교육 안내문, 결과보고서 등)를 한글(HWP/HWPX) 형식으로 작성·배포한다.

현재 문제:
- Claude 등 AI로 내용을 작성해도 최종적으로 한글 프로그램에서 서식을 수동으로 다시 잡아야 함
- 줄간격, 자간, 폰트, 표 여백 등이 매번 틀어져 반복 수작업 발생
- 한글 프로그램을 자동으로 실행할 수 없는 환경(Linux 서버)에서는 HWPX 직접 생성이 유일한 해결책

목표:
- **Python 라이브러리(CLI + API)** 형태로, 구조화된 문서 데이터(JSON/dict)를 입력받아 서식이 완벽하게 적용된 `.hwpx` 파일을 생성한다.
- 한글에서 열었을 때 추가 서식 작업 없이 바로 사용 가능한 수준의 품질을 목표로 한다.

---

### 1-2. 핵심 요구사항 (Must Have)

#### M1. HWPX 파일 생성
- 표준 HWPX 포맷(ZIP + XML) 으로 파일 출력
- 한글 2018 이상에서 정상적으로 열릴 것
- 파일 손상 없이 재저장 가능할 것

#### M2. 문단 스타일 지원
- 폰트명, 크기(pt), 굵게, 기울임, 밑줄 설정 가능
- 줄간격: %, 고정값(pt) 두 가지 모드
- 문단 위/아래 간격(pt) 설정
- 들여쓰기(left indent), 내어쓰기(hanging indent) 설정
- 정렬: 왼쪽 / 가운데 / 오른쪽 / 양쪽

#### M3. 공문서 불릿 체계 지원
아래 계층 구조를 자동으로 처리:

```
■ 1. 대섹션 제목       (굵게, 색상 강조)
  ○ 소제목             (굵게)
    가. 세부항목        (일반)
      - 설명 항목       (들여쓰기 1단계)
        · 보충 설명     (들여쓰기 2단계, 소폰트)
  ※ 주석/참고사항      (색상 강조)
```

각 레벨별 들여쓰기, 폰트 크기, 색상을 스타일로 미리 정의하고 참조 방식으로 적용.

#### M4. 표(Table) 생성
- 열 너비 지정 (mm 또는 twip 단위)
- 셀 내부 여백 설정 (상하/좌우 각각 mm 단위)
- 셀 배경색, 테두리 스타일/두께/색상
- 셀 병합 (colspan, rowspan)
- 셀 내 텍스트 정렬 (가로/세로)
- 표 헤더 행 스타일 별도 지정

#### M5. 페이지 설정
- 용지 크기: A4 기본, 커스텀 가능
- 여백: 상/하/좌/우 각각 mm 단위 지정
- 용지 방향: 세로/가로

#### M6. 스타일 시스템
- 문서 수준에서 스타일을 사전 정의 (이름, 폰트, 크기, 줄간격 등)
- 각 문단/텍스트에서 스타일 이름으로 참조
- 기본 스타일 프리셋 제공 (공문서 표준, 보고서 등)

#### M7. 페이지 흐름 제어 (Page Flow Control) ★ 핵심

**배경:**
공문서에서 가장 빈번한 문제는 내용이 페이지 경계에서 어중간하게 잘리는 것이다.
섹션 제목만 한 페이지에 남고 내용은 다음 페이지로 넘어가거나, 표가 중간에 잘려
두 페이지에 걸쳐 출력되는 경우가 대표적이다. 이를 구조적으로 방지해야 한다.

**요구사항:**

1. **표 분리 금지 (keep-together)**
   - 모든 표는 기본적으로 페이지 중간에서 잘리지 않음
   - 표 전체가 현재 페이지에 들어가지 않으면 → 표 전체를 다음 페이지로 이동
   - HWPX XML: `<hp:tblPr><hp:protect>` + 페이지 분리 방지 속성 적용

2. **제목-내용 분리 금지 (keep-with-next)**
   - ■ 대섹션 제목: 최소 다음 3개 블록(소제목 또는 첫 번째 항목)과 항상 같은 페이지
   - ○ 소제목: 바로 다음 블록(첫 번째 li1 또는 표)과 항상 같은 페이지
   - 가./나./다. 항목: 바로 다음 li1과 항상 같은 페이지
   - HWPX XML: 해당 문단에 `<hp:breakLatinWord>` + `keepWithNext` 속성 적용

3. **고아/과부 방지 (widow/orphan control)**
   - 문단의 마지막 1줄만 다음 페이지에 남는 상황(고아, orphan) 방지
   - 문단의 첫 1줄만 이전 페이지에 남는 상황(과부, widow) 방지
   - 최소 2줄 이상이 같은 페이지에 함께 있도록 강제

4. **수동 페이지 브레이크**
   - `doc.add_page_break()` API 제공
   - JSON에서 `{"type": "page_break"}` 로 명시적 삽입 가능

5. **섹션 단위 자동 배치 전략**
   - 빌더가 save() 시점에 각 블록의 예상 높이를 계산
   - 현재 페이지 잔여 공간보다 다음 섹션(제목+첫 표/내용)이 크면 자동으로 페이지 브레이크 삽입
   - 높이 계산 기준:
     - 문단: `줄 수 × 줄간격(mm) + 문단 위아래 간격`
     - 표: `행 수 × 행 높이(mm) + 표 위아래 여백`
     - 폰트 크기(pt) → 줄 높이 환산: `줄높이(mm) = 폰트크기(pt) × 줄간격(%) × 0.3528`

**페이지 흐름 제어 규칙 요약표:**

| 요소 | 기본 동작 | keep-with-next | keep-together |
|------|-----------|---------------|---------------|
| ■ 대섹션 제목 | - | ✅ 다음 3블록과 묶음 | - |
| ○ 소제목 | - | ✅ 다음 1블록과 묶음 | - |
| 가./나./다. 항목 | - | ✅ 다음 1블록과 묶음 | - |
| - 항목 (li1) | 자유 흐름 | - | - |
| · 항목 (li2) | 자유 흐름 | - | - |
| 표 (4행 이하) | - | - | ✅ 전체 묶음 |
| 표 (5행 이상) | - | - | ✅ 전체 묶음 (불가시 다음 페이지) |
| ※ 주석 | 자유 흐름 | - | - |

---

### 1-3. 추가 요구사항 (Nice to Have)

#### N1. 머리말/꼬리말
- 문서 제목, 페이지 번호 삽입

#### N2. 이미지 삽입 (PNG/JPG)
- PNG/JPG 파일 경로를 지정하여 본문에 삽입
- 크기(mm 단위), 정렬(좌/가운데/우) 지정
- 캡션 텍스트 추가 가능
- 이미지는 HWPX BinData/ 폴더에 패키징
- 페이지 흐름 제어 대상: 이미지+캡션은 분리 금지

**JSON 입력 예시:**
```json
{
  "type": "image",
  "src": "images/photo.jpg",
  "width_mm": 80,
  "align": "center",
  "caption": "사진 1. 현장 모니터링 장면"
}
```

---

#### N3. 도형 흐름도 (Shape Diagram) ★ HWPX 네이티브 벡터

> **[샘플 파일 분석 결과 반영]** 실제 .hwpx 파일을 분석한 결과, 도형은 PNG가 아닌
> HWPX 네이티브 XML 태그로 직접 생성 가능함이 확인됨. 화질 손실 없는 완전한 벡터 출력.

**구현 전략: HWPX 네이티브 도형 XML 직접 생성**
- PNG 래스터화 방식 사용 안 함 → 확대/인쇄해도 선명
- 한글에서 열면 도형을 직접 클릭·편집 가능
- 단위: HWP unit (1mm ≈ 283 HWP unit)

**실제 확인된 HWPX 도형 XML 태그 (샘플 파일 리버스 엔지니어링):**

```xml
<!-- 사각형 -->
<hp:rect id="..." zOrder="0" ...>
  <hp:lineShape color="#1F4E79" width="33" style="SOLID"/>
  <hc:fillBrush>
    <hc:winBrush faceColor="#D6E4F0"/>
  </hc:fillBrush>
  <hp:sz width="14175" height="5669"/>          <!-- mm → HWP unit 변환 -->
  <hp:pos horzOffset="0" vertOffset="0"/>
  <hp:caption>Step 1. 준비 및 안내</hp:caption>
</hp:rect>

<!-- 타원 (원형) -->
<hp:ellipse id="..." ...>
  <hc:center x="4007" y="3631"/>
  <hc:ax1 x="8014" y="3631"/>
  <hc:ax2 x="4007" y="0"/>
  ...
</hp:ellipse>

<!-- 연결선 + 화살표 -->
<hp:connectline id="..." ...>
  <hp:lineShape headStyle="NORMAL" tailStyle="ARROW" tailSz="MEDIUM_MEDIUM"/>
  <hc:pt x="14175" y="2835"/>   <!-- 시작점 -->
  <hc:pt x="17010" y="2835"/>   <!-- 끝점 -->
</hp:connectline>

<!-- 다각형 (마름모 등 커스텀) -->
<hp:polygon id="..." ...>
  <hc:pt x="0"    y="2835"/>
  <hc:pt x="5669" y="0"/>
  <hc:pt x="11339" y="2835"/>
  <hc:pt x="5669" y="5669"/>
</hp:polygon>
```

**지원 도형 타입:**

| JSON 타입 | HWPX 태그 | 설명 |
|-----------|-----------|------|
| `rect` | `hp:rect` | 사각형 (Step 박스 등) |
| `rounded_rect` | `hp:rect` + `rx` 속성 | 모서리 둥근 사각형 |
| `ellipse` | `hp:ellipse` | 타원/원 |
| `diamond` | `hp:polygon` (4점) | 마름모 (조건 분기) |
| `arrow_line` | `hp:connectline` | 화살표 연결선 |
| `polygon` | `hp:polygon` | 임의 다각형 |

**지원 레이아웃:**

| 레이아웃 | 설명 |
|----------|------|
| `step_flow` | 가로/세로 일렬 흐름도 (Step1→Step2→…) |
| `timeline` | 날짜 기반 타임라인 |
| `org_chart` | 조직도 (트리 구조) |
| `free` | 노드 좌표 직접 지정 |

**JSON 입력 예시 — Step 흐름도:**
```json
{
  "type": "diagram",
  "layout": "step_flow",
  "direction": "horizontal",
  "width_mm": 160,
  "height_mm": 30,
  "theme_color": "1F4E79",
  "nodes": [
    { "id": "s1", "label": "Step 1.\n준비 및 안내",  "shape": "rect" },
    { "id": "s2", "label": "Step 2.\n자체점검 회수", "shape": "rect" },
    { "id": "s3", "label": "Step 3.\n현장 모니터링", "shape": "rect" },
    { "id": "s4", "label": "Step 4.\n결과 보고",     "shape": "rect" }
  ],
  "edges": [
    { "from": "s1", "to": "s2" },
    { "from": "s2", "to": "s3" },
    { "from": "s3", "to": "s4" }
  ]
}
```

---

#### N4. 차트 (Chart) ★ HWPX 네이티브 OOXML 차트

> **[샘플 파일 분석 결과 반영]** 실제 .hwpx 파일을 분석한 결과, 차트는
> `Chart/chart1.xml` 파일에 **OOXML DrawingML 표준 포맷**으로 저장됨이 확인됨.
> Excel/PowerPoint와 동일한 스펙 → 이미 잘 문서화되어 있어 구현 수월.
> PNG 래스터화 없이 한글에서 완전한 벡터 차트로 렌더링됨.

**실제 확인된 HWPX 차트 구조 (샘플 파일 리버스 엔지니어링):**

```
output.hwpx
└── Chart/
    └── chart1.xml        ← OOXML DrawingML 차트 스펙 그대로
        xmlns:c = "http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a = "http://schemas.openxmlformats.org/drawingml/2006/main"
```

```xml
<!-- section0.xml에서 차트 참조 방식 -->
<hp:chart id="..." chartIDRef="Chart/chart1.xml">
  <hp:sz width="32250" height="18750"/>          <!-- 크기 (HWP unit) -->
  <hp:pos vertOffset="0" horzOffset="0"/>
</hp:chart>

<!-- Chart/chart1.xml 내부 구조 (OOXML 표준) -->
<c:chartSpace>
  <c:chart>
    <c:plotArea>
      <c:barChart>                               <!-- bar/line/pie 등 교체 -->
        <c:barDir val="col"/>                    <!-- col=세로, bar=가로 -->
        <c:grouping val="stacked"/>              <!-- clustered/stacked 등 -->
        <c:ser>                                  <!-- 데이터 계열 -->
          <c:cat>                                <!-- X축 레이블 -->
            <c:strCache>
              <c:pt idx="0"><c:v>항목 1</c:v></c:pt>
            </c:strCache>
          </c:cat>
          <c:val>                                <!-- Y축 값 -->
            <c:numCache>
              <c:pt idx="0"><c:v>4.3</c:v></c:pt>
            </c:numCache>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
```

**데이터 입력 방식 두 가지 모두 지원:**
- 방식 A: JSON으로 데이터 직접 입력 → chart1.xml 자동 생성
- 방식 B: 외부 이미지(PNG) 경로 지정 → BinData에 래스터 이미지로 삽입 (fallback)

**지원 차트 타입:**

| JSON 타입 | OOXML 태그 | 설명 |
|-----------|-----------|------|
| `bar` | `c:barChart barDir=col` | 세로 막대 |
| `barh` | `c:barChart barDir=bar` | 가로 막대 |
| `line` | `c:lineChart` | 꺾은선 |
| `pie` | `c:pieChart` | 원형 |
| `stacked_bar` | `c:barChart grouping=stacked` | 누적 막대 |

**JSON 입력 예시 — 방식 A:**
```json
{
  "type": "chart",
  "chart_type": "bar",
  "width_mm": 120,
  "height_mm": 70,
  "title": "기관별 자체점검 제출 현황",
  "data": {
    "labels": ["하청교회", "은빛노인", "남양양로원", "노인세상", "기관5"],
    "datasets": [
      { "label": "제출 완료", "values": [1, 1, 1, 0, 1], "color": "2E75B6" }
    ]
  },
  "caption": "그림 2. 자체점검표 제출 현황"
}
```

---

#### N5. SVG → HWPX 네이티브 도형 변환

> **[변경]** SVG → PNG 래스터화 방식 폐기.
> SVG path 데이터를 파싱하여 HWPX `hp:polygon` 좌표로 변환하는 방식으로 변경.
> 벡터 품질 완전 보존.

**구현 전략:**
- `svgpathtools` 라이브러리로 SVG path 파싱
- path 좌표 → HWP unit 변환 후 `hp:polygon` 태그 생성
- 단순 도형(rect, circle, line)은 각각 `hp:rect`, `hp:ellipse`, `hp:connectline`으로 매핑
- 복잡한 곡선 path는 `hp:polygon` 근사 좌표열로 변환

**JSON 입력 예시:**
```json
{
  "type": "svg",
  "src": "icons/process_diagram.svg",
  "width_mm": 160,
  "align": "center",
  "caption": "그림 3. 상시관리 연계 체계"
}
```

---

#### N6. 텍스트 박스 (강조 박스)
- 테두리+배경색이 있는 강조용 박스
- 공문서의 "※ 참고", "중요" 강조 표시에 활용
- HWPX `hp:rect` 도형 안에 텍스트 삽입 방식으로 구현 (네이티브 벡터)

**JSON 입력 예시:**
```json
{
  "type": "text_box",
  "text": "※ 자체점검표는 반드시 3월 4일(수)까지 제출하여야 합니다.",
  "border_color": "C00000",
  "bg_color": "FFF2CC",
  "font_bold": true,
  "padding_mm": 4
}
```

---

#### N7. 색상 테마
- 주색상(main), 보조색상(sub), 강조색상(accent) 3가지를 테마로 정의
- 섹션 제목, 표 헤더, 도형 fill/선색, 차트 계열색에 일관되게 자동 적용

#### N8. JSON 스키마 입력
- 문서 구조를 JSON으로 정의하여 CLI에서 바로 실행 가능

---

### 1-4. 비기능 요구사항

| 항목 | 기준 |
|------|------|
| 실행 환경 | Python 3.9 이상, Linux/macOS/Windows |
| 외부 의존성 | zipfile, lxml, svgpathtools (SVG 변환용) |
| 한글 미설치 환경 | 한글 프로그램 없이 생성 가능 |
| 파일 크기 | 이미지 없는 일반 문서 기준 1MB 이하 |
| 생성 시간 | 도형·차트 포함 10페이지 이내 문서 기준 5초 이내 |
| 벡터 품질 | 도형·차트 확대/인쇄 시 화질 손실 없음 (네이티브 벡터) |
| 인코딩 | UTF-8 전용 |

---

### 1-5. 제약사항 및 가정

- HWPX 공식 SDK 사용 불가 (한글과컴퓨터 비공개)
- HWPX 스펙은 실제 샘플 파일(.hwpx) 리버스 엔지니어링으로 확인된 구조 기반으로 구현
- 네임스페이스는 2011 버전 기준 (`hwpml/2011/...`) — 샘플 파일에서 실측 확인
- 도형과 차트는 PNG 래스터화 없이 HWPX 네이티브 벡터로 생성 (화질 손실 없음)
- 차트는 OOXML DrawingML 표준 (`Chart/chart1.xml`) — Excel/PowerPoint와 동일 스펙
- SVG 삽입 시 단순 도형은 네이티브 태그로, 복잡한 곡선은 polygon 근사 좌표로 변환
- 한글 2018 이상 기준으로 호환성 검증
- 복잡한 수식, 매크로는 범위 외

---

## 2. IA (Information Architecture)

### 2-1. 모듈 구조

```
hwpx_generator/
│
├── hwpx_generator/
│   ├── __init__.py
│   ├── document.py          # HwpxDocument — 최상위 문서 객체
│   ├── styles.py            # StyleManager — 스타일 정의/참조
│   ├── elements/
│   │   ├── __init__.py
│   │   ├── paragraph.py     # Paragraph, TextRun
│   │   ├── table.py         # Table, TableRow, TableCell
│   │   ├── image.py         # Image (PNG/JPG 파일 삽입)
│   │   ├── diagram.py       # ★ ShapeDiagram (hp:rect/ellipse/connectline 네이티브 벡터)
│   │   ├── chart.py         # ★ Chart (OOXML chart1.xml 네이티브 벡터)
│   │   ├── svg_element.py   # ★ SvgElement (SVG path → hp:polygon 좌표 변환)
│   │   └── text_box.py      # TextBox (hp:rect 안에 텍스트)
│   ├── builders/
│   │   ├── __init__.py
│   │   ├── xml_builder.py        # section0.xml 조립
│   │   ├── styles_builder.py     # header.xml 스타일 정의
│   │   ├── settings_builder.py   # settings.xml 생성
│   │   ├── chart_builder.py      # ★ Chart/chart1.xml OOXML 생성
│   │   ├── shape_builder.py      # ★ hp:rect/ellipse/polygon/connectline XML 생성
│   │   ├── svg_converter.py      # ★ SVG path → hp:polygon 좌표 변환
│   │   ├── pageflow.py           # PageFlowController — 높이 계산 및 페이지 브레이크
│   │   └── package_builder.py    # ZIP 패키징 (Chart/, BinData/ 포함)
│   ├── presets/
│   │   ├── __init__.py
│   │   └── gov_document.py       # 공문서 표준 스타일 프리셋
│   └── utils.py                  # 단위 변환 (mm→HWP unit), 색상 헬퍼
│
├── cli.py
├── tests/
│   ├── test_paragraph.py
│   ├── test_table.py
│   ├── test_diagram.py
│   ├── test_chart.py
│   ├── test_svg.py
│   └── test_document.py
├── examples/
│   ├── sample_notice.py
│   └── sample_notice.json
├── requirements.txt
└── README.md
```

---

### 2-2. 핵심 클래스 관계

```
HwpxDocument
  ├── PageSettings
  ├── StyleManager
  │     └── Style[]
  ├── PageFlowController
  │     ├── estimate_height(block) → mm
  │     ├── check_page_overflow(blocks)
  │     └── insert_page_breaks(blocks)
  └── Section[]
        └── Block[]  # Paragraph | Table | Image | Diagram | Chart | SvgElement | TextBox | PageBreak
              ├── Paragraph
              │     ├── keep_with_next: bool
              │     ├── widow_orphan: int
              │     └── TextRun[]
              ├── Table
              │     ├── keep_together: bool
              │     └── TableRow[] → TableCell[] → Paragraph[]
              ├── Image                          # PNG/JPG → BinData 패키징
              │     ├── src: str
              │     ├── width_mm: float
              │     ├── align: str
              │     └── caption: str
              ├── Diagram                        # ★ 네이티브 벡터 도형
              │     ├── nodes: list[Node]
              │     │     └── Node(id, label, shape, x_mm, y_mm, w_mm, h_mm, fill, border)
              │     ├── edges: list[Edge]
              │     │     └── Edge(from_id, to_id, style)
              │     ├── layout: str              #   step_flow / timeline / org_chart / free
              │     └── → shape_builder.py 에서
              │           hp:rect / hp:ellipse / hp:polygon / hp:connectline XML 생성
              ├── Chart                          # ★ OOXML 네이티브 차트
              │     ├── chart_type: str          #   bar/line/pie/stacked_bar
              │     ├── data: dict               #   방식 A: JSON 데이터
              │     ├── src: str                 #   방식 B: 외부 이미지 경로 (fallback)
              │     ├── width_mm, height_mm
              │     └── → chart_builder.py 에서 Chart/chartN.xml OOXML 생성
              ├── SvgElement                     # ★ SVG → 네이티브 도형 변환
              │     ├── src: str
              │     ├── width_mm: float
              │     └── → svg_converter.py 에서
              │           SVG path → hp:polygon 좌표 변환
              ├── TextBox                        # hp:rect + 내부 텍스트
              │     ├── text: str
              │     ├── border_color: str
              │     ├── bg_color: str
              │     └── padding_mm: float
              └── PageBreak
```

---

### 2-3. HWPX 파일 내부 구조 (ZIP)

```
output.hwpx  (ZIP)
├── mimetype                          # "application/hwp+zip" (압축 없음)
├── version.xml                       # 한글 버전 정보
├── META-INF/
│   ├── container.xml
│   ├── container.rdf
│   └── manifest.xml
├── settings.xml                      # 문서 설정 (용지, 여백)
├── Contents/
│   ├── content.hpf                   # 패키지 매니페스트
│   ├── header.xml                    # 스타일, 도형 속성 정의
│   └── section0.xml                  # 본문 (도형/차트 참조 포함)
├── Chart/                            # ★ 차트 데이터 (OOXML 표준)
│   ├── chart1.xml                    #   첫 번째 차트
│   └── chart2.xml                    #   두 번째 차트 (복수 가능)
├── BinData/                          # 이미지 바이너리
│   ├── image1.png
│   └── image2.jpg
└── Preview/
    ├── PrvText.txt
    └── PrvImage.png
```

> ※ 실제 샘플 파일 분석 결과 기준. `Styles/styles.xml` 별도 파일이 아닌
> `Contents/header.xml` 안에 스타일이 포함되는 구조임을 확인.

---

### 2-4. 핵심 XML 네임스페이스

> ※ 실제 샘플 파일 분석으로 확인된 네임스페이스 (2011 버전 기준)

```xml
<!-- section0.xml, header.xml 공통 -->
xmlns:hp  = "http://www.hancom.co.kr/hwpml/2011/paragraph"
xmlns:hc  = "http://www.hancom.co.kr/hwpml/2011/core"
xmlns:hs  = "http://www.hancom.co.kr/hwpml/2011/section"
xmlns:hh  = "http://www.hancom.co.kr/hwpml/2011/head"
xmlns:hhs = "http://www.hancom.co.kr/hwpml/2011/history"
xmlns:hm  = "http://www.hancom.co.kr/hwpml/2011/master-page"

<!-- 차트 (OOXML 표준 — Excel/PPT와 동일) -->
xmlns:c   = "http://schemas.openxmlformats.org/drawingml/2006/chart"
xmlns:a   = "http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:r   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

<!-- 신형 한글(2016+) 확장 -->
xmlns:hp10      = "http://www.hancom.co.kr/hwpml/2016/paragraph"
xmlns:ooxmlchart = "http://www.hancom.co.kr/hwpml/2016/ooxmlchart"
```

---

### 2-5. 단위 변환 기준

| 단위 | 변환 |
|------|------|
| mm → HWP unit (twip) | 1mm = 2834.6 / 25.4 ≈ 72 HWP unit (1/7200 inch 기준) |
| pt → HWP unit | 1pt = 100 (HWPX 폰트 크기는 1/100 pt) |
| 줄간격 % | 값 그대로 (예: 160) |
| 줄간격 고정 | mm → HWP unit 변환 |

> ※ HWPX의 내부 길이 단위는 **1/7200 인치 (≈ 0.00353mm)** 이며, 편의상 "HWP unit" 으로 칭함.
> 실제 변환: `hwp_unit = round(mm_value * 7200 / 25.4)`

---

## 3. Use Case

### UC-01. 공문서 생성 (핵심)

**행위자:** 사용자 (광역지원기관 담당자 또는 AI 에이전트)  
**전제조건:** hwpx_generator 설치 완료  
**기본 흐름:**

```
1. 사용자가 HwpxDocument 객체 생성
2. PageSettings로 용지/여백 설정
3. StyleManager에 문서 스타일 등록 (또는 프리셋 로드)
4. 섹션에 Paragraph/Table 요소 추가
5. document.save("output.hwpx") 호출
6. ZIP 패키징 → .hwpx 파일 출력
7. 한글에서 파일을 열어 확인
```

**Python 코드 예시:**
```python
from hwpx_generator import HwpxDocument
from hwpx_generator.presets import GovDocumentPreset

doc = HwpxDocument()
doc.apply_preset(GovDocumentPreset(
    font_body="HY헤드라인M",
    font_table="맑은 고딕",
    size_title=15,
    size_body=13,
    size_table=10,
    line_spacing=160,
    margin_top=10, margin_bottom=10,
    margin_left=20, margin_right=20,
    cell_margin_lr=3, cell_margin_tb=2,
))

# 제목
doc.add_title("2026년 경상남도 노인맞춤돌봄서비스\n신규 수행기관 모니터링 실시 안내서")

# 섹션
sec1 = doc.add_section_heading(1, "모니터링 개요")
sec1.add_sub_heading("목적")
sec1.add_bullet1("신규 수행기관의 사업 운영 적정성을 확보하고...")

# 표
table = doc.add_table(col_widths=[25, 145])  # mm 단위
table.add_header_row(["항목", "내용"])
table.add_row(["기간", "2026년 3월 9일(월) ~ 3월 13일(금)"])
table.add_row(["대상", "2026년 신규 지정 지역수행기관 5개소"])

doc.save("모니터링_실시안내서.hwpx")
```

---

### UC-02. JSON 입력으로 CLI 실행

**행위자:** AI 에이전트 (Claude Code 등)  
**전제조건:** cli.py 실행 가능  
**기본 흐름:**

```bash
python cli.py --input sample_notice.json --output 모니터링_실시안내서.hwpx
```

**JSON 구조 예시:**
```json
{
  "preset": "gov_document",
  "preset_options": {
    "font_body": "HY헤드라인M",
    "font_table": "맑은 고딕",
    "size_title": 15,
    "size_body": 13,
    "size_table": 10,
    "line_spacing": 160,
    "margin": { "top": 10, "bottom": 10, "left": 20, "right": 20 },
    "cell_margin": { "lr": 3, "tb": 2 }
  },
  "content": [
    {
      "type": "title",
      "text": "2026년 경상남도 노인맞춤돌봄서비스\n신규 수행기관 모니터링 실시 안내서"
    },
    {
      "type": "section_heading",
      "level": 1,
      "num": "1",
      "text": "모니터링 개요"
    },
    {
      "type": "sub_heading",
      "text": "목적"
    },
    {
      "type": "bullet",
      "level": 1,
      "text": "신규 수행기관의 사업 운영 적정성을 확보하고..."
    },
    {
      "type": "table",
      "col_widths_mm": [25, 145],
      "header": ["항목", "내용"],
      "rows": [
        ["기간", "2026년 3월 9일(월) ~ 3월 13일(금)"],
        ["대상", "2026년 신규 지정 지역수행기관 5개소"],
        ["수행인력", "광역지원기관 전담사회복지사"]
      ]
    }
  ]
}
```

---

### UC-03. 기존 문서 내용 재사용

**행위자:** 사용자  
**전제조건:** 기존 .hwp/.hwpx 또는 PDF 파일 존재  
**기본 흐름:**

```
1. hwp5txt 또는 PDF 텍스트 추출로 내용 파싱
2. AI(Claude)가 구조화된 JSON으로 변환
3. UC-02 방식으로 새 HWPX 생성
```

---

### UC-05. 페이지 흐름 자동 제어 ★

**행위자:** 사용자 (자동 처리, 별도 호출 불필요)  
**전제조건:** 문서 내용 구성 완료, `doc.save()` 호출 직전  
**기본 흐름:**

```
1. save() 호출 시 PageFlowController.process(blocks) 자동 실행
2. 각 블록의 예상 높이(mm) 계산
3. 현재 페이지 잔여 공간 추적
4. 규칙 적용:
   a. 표: keep_together=True → 잔여 공간 부족 시 앞에 페이지 브레이크 삽입
   b. 섹션 제목(■): 다음 3블록과 묶음 → 합산 높이가 잔여 공간 초과 시 페이지 브레이크
   c. 소제목(○): 다음 1블록과 묶음 → 동일 처리
   d. 고아/과부: 문단 마지막/첫 줄 분리 감지 → 페이지 브레이크 조정
5. 최종 블록 리스트(페이지 브레이크 포함)로 XML 생성
```

**높이 계산 로직:**
```python
def estimate_height_mm(block, style) -> float:
    if isinstance(block, Paragraph):
        line_h = style.font_size_pt * style.line_spacing_pct / 100 * 0.3528
        lines  = estimate_line_count(block.text, content_width_mm, style)
        return line_h * lines + style.space_before_mm + style.space_after_mm

    if isinstance(block, Table):
        row_h = style.table_font_pt * 1.6 * 0.3528 + style.cell_margin_tb * 2
        return row_h * len(block.rows) + 2  # +2mm 여유
```

**사용자가 자동 제어를 끄고 싶을 때:**
```python
doc.save("output.hwpx", auto_page_flow=False)  # 수동 제어 모드
```

**JSON에서 수동 페이지 브레이크 삽입:**
```json
{ "type": "page_break" }
```

---

### UC-04. 스타일 프리셋 커스터마이징

**행위자:** 사용자  
**기본 흐름:**

```python
from hwpx_generator.presets import GovDocumentPreset

# 기본 프리셋에서 일부만 변경
preset = GovDocumentPreset()
preset.color_main = "1F4E79"
preset.color_accent = "C00000"
preset.bullet_styles = {
    "h1":    {"prefix": "■ {num}. ", "bold": True,  "color": "1F4E79", "size": 13},
    "h2":    {"prefix": "○ ",        "bold": True,  "color": "2E75B6", "size": 13},
    "sub1":  {"prefix": "가. ",       "bold": False, "color": "000000", "size": 13},
    "li1":   {"prefix": "- ",         "indent_mm": 7,  "hanging_mm": 4},
    "li2":   {"prefix": "· ",         "indent_mm": 13, "hanging_mm": 4, "size": 12},
    "note":  {"prefix": "※ ",         "bold": True,  "color": "C00000", "size": 11},
}
doc.apply_preset(preset)
```

---

### UC-06. 마크다운 → HWPX 직접 변환

**행위자:** 사용자  
**전제조건:** hwpx_generator 설치 완료, .md 파일 준비  
**기본 흐름:**

```
1. CLI에 .md 파일 경로 입력
2. 확장자 감지 → MarkdownParser 자동 호출
3. 마크다운 파싱 → 메모리상 블록 구조로 변환 (JSON 중간 파일 없음)
4. 기존 hwpx_generator 파이프라인으로 .hwpx 생성
5. 파일 출력
```

**CLI 사용법:**
```bash
python cli.py --input 안내서.md --output 안내서.hwpx
python cli.py --input 안내서.md --output 안내서.hwpx --preset gov_document
```

**Python API 사용법:**
```python
from hwpx_generator import HwpxDocument
from hwpx_generator.parsers import MarkdownParser

doc = HwpxDocument()
doc.apply_preset(GovDocumentPreset())
doc.load_markdown("안내서.md")   # 내부적으로 MarkdownParser 호출
doc.save("안내서.hwpx")
```

**마크다운 → 공문서 블록 매핑 규칙:**

| 마크다운 문법 | 공문서 블록 | 비고 |
|---|---|---|
| `# 제목` | `title` | 문서 제목 |
| `## 1. 텍스트` | `section_heading` | ■ 자동 부여 |
| `### 텍스트` | `sub_heading` | ○ 자동 부여 |
| `#### 텍스트` | `bullet sub1` | 가. 나. 다. 자동 부여 |
| `- 텍스트` | `bullet li1` | - 유지 |
| `  - 텍스트` | `bullet li2` | · 변환 |
| `※ 텍스트` | `note` | ※ 그대로 |
| `\| 표 \|` | `table` | 표준 마크다운 테이블 |
| `---` | `page_break` | 수평선 → 페이지 나누기 |
| `> 텍스트` | `text_box` | 인용구 → 강조 박스 |

**마크다운 예시 (입력):**
```markdown
# 2026년 경상남도 노인맞춤돌봄서비스
신규 수행기관 모니터링 실시 안내서

## 1. 모니터링 개요

### 목적

- 신규 수행기관의 사업 운영 적정성을 확보하고, 2026년 개정 지침의
  이행 여부 확인 및 현장 모니터링을 통한 운영 안정화 지원함

### 세부 운영 현황

| 항목 | 내용 |
|------|------|
| 일시 | 2026. 3. 9.(월) ~ 3. 13.(금) |
| 대상 | 2026년 신규 지정 지역수행기관 5개소 |

## 2. 수행기관 협조 사항

#### 사전 자체점검 철저

- 지표별 운영 현황을 객관적으로 진단하고 요청사항 기재
- 제출기한: 3. 4.(수)까지 gnscc@naver.com으로 제출

> ※ 자체점검표는 반드시 기한 내 제출하여야 합니다.
```

**구현 위치:**
```
hwpx_generator/
└── parsers/
    ├── __init__.py
    ├── markdown_parser.py   # ★ 신규 — 마크다운 → 블록 구조 변환
    └── json_parser.py       # 기존 JSON 파서 (리팩토링)
```

**핵심 설계 원칙:**
- JSON 중간 파일 저장 없이 메모리에서 직접 블록 리스트로 변환
- 변환된 블록 리스트는 기존 `xml_builder.py`가 그대로 처리
- 기존 JSON → HWPX 파이프라인 코드 변경 없음

---

## 4. 개발 우선순위 및 마일스톤

| Phase | 내용 | 산출물 |
|-------|------|--------|
| Phase 1 | HWPX 기본 구조 생성 (빈 문서, 페이지 설정) | `HwpxDocument.save()` 작동 |
| Phase 2 | 텍스트/문단 스타일 구현 | 폰트·줄간격 적용된 본문 출력 |
| Phase 3 | 공문서 불릿 체계 구현 | ■/○/가./- /· 계층 정상 출력 |
| Phase 4 | 표 생성 구현 | 헤더·셀여백·병합 포함 표 출력 |
| Phase 5 | 페이지 흐름 제어 구현 | 표 분리 없음, 제목-내용 묶음, 자동 페이지 브레이크 |
| Phase 6 | 이미지(PNG/JPG) 삽입 구현 | BinData 패키징, 크기·정렬·캡션 |
| Phase 7 | 네이티브 도형 흐름도 구현 ★ | hp:rect/ellipse/connectline XML 생성, Step 흐름도 벡터 출력 |
| Phase 8 | OOXML 네이티브 차트 구현 ★ | Chart/chart1.xml 생성, 막대/꺾은선/원형 벡터 차트 |
| Phase 9 | SVG → 네이티브 도형 변환 구현 ★ | SVG path → hp:polygon 변환, 벡터 품질 보존 |
| Phase 10 | 텍스트 박스 구현 | hp:rect + 내부 텍스트, 테두리·배경색 |
| Phase 11 | 프리셋 시스템 + CLI | JSON → HWPX 원스텝 변환 |
| Phase 12 | 테스트 & 한글 호환성 검증 | 한글 2018~2024 오픈, 도형·차트 클릭 편집 가능 확인 |

---

## 5. 검증 기준 (Definition of Done)

- [ ] 한글 2020 이상에서 파일 오류 없이 열림
- [ ] 폰트명이 그대로 적용됨 (HY헤드라인M, 맑은 고딕)
- [ ] 줄간격 160% 정확히 반영
- [ ] 표 셀 여백 좌우 3mm / 상하 2mm 정확히 반영
- [ ] 공문서 불릿 5단계 계층 정상 출력
- [ ] 페이지 여백 상하 10mm / 좌우 20mm 적용
- [ ] 표가 페이지 중간에서 잘리지 않음
- [ ] ■ 섹션 제목이 페이지 맨 아래에 혼자 남지 않음
- [ ] ○ 소제목이 다음 내용과 항상 같은 페이지에 위치
- [ ] 고아/과부 줄 발생 없음 (최소 2줄 묶음)
- [ ] PNG/JPG 이미지가 지정 크기·정렬로 문서에 삽입됨
- [ ] Step 흐름도가 HWPX 네이티브 벡터 도형으로 출력됨 ★ (확대·인쇄 시 선명)
- [ ] 한글에서 생성된 도형을 클릭하여 직접 편집 가능 ★
- [ ] 막대/꺾은선/원형 차트가 OOXML 네이티브 벡터로 출력됨 ★
- [ ] 한글에서 생성된 차트를 더블클릭하여 데이터 편집 가능 ★
- [ ] SVG 파일이 네이티브 도형으로 변환되어 벡터 품질 보존됨 ★
- [ ] 강조 텍스트 박스 테두리·배경색 정상 출력
- [ ] 도형·차트 포함 문서 5초 이내 생성
- [ ] 생성 후 한글에서 재저장 시 파일 손상 없음

---

*본 문서는 경상남도사회서비스원 광역지원기관 내부 도구 개발용이며, hwpx_generator 라이브러리의 설계 기준으로 활용됩니다.*
