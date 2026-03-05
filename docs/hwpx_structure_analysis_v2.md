# 실제 HWPX 파일 구조 분석 결과
> 샘플 .hwpx 파일 직접 분석 (리버스 엔지니어링)
> 파일 손상 원인 파악 및 수정을 위한 실측 데이터

---

## 1. hp:p (단락) 구조 — 실측

### 1-1. 빈 단락 (가장 단순한 형태)

```xml
<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="0"/>
  <hp:linesegarray>
    <hp:lineseg textpos="0" vertpos="1600" vertsize="1000" textheight="1000"
                baseline="850" spacing="600" horzpos="0" horzsize="42520" flags="393216"/>
  </hp:linesegarray>
</hp:p>
```

**핵심 확인사항:**
- `hp:linesegarray` + `hp:lineseg` 는 **필수 요소** (없으면 파일 손상)
- 빈 단락도 반드시 `hp:lineseg` 1개 포함
- `hp:run` 이 비어있어도 (`<hp:run charPrIDRef="0"/>`) lineseg는 있어야 함

### 1-2. lineseg 속성 의미

| 속성 | 실측값 | 의미 |
|------|--------|------|
| `textpos` | 0 | 텍스트 시작 위치 |
| `vertpos` | 1600, 3200... | 수직 위치 (1600씩 증가, 줄간격) |
| `vertsize` | 1000 | 줄 높이 |
| `textheight` | 1000 | 텍스트 높이 |
| `baseline` | 850 | 베이스라인 위치 |
| `spacing` | 600 | 줄 간격 |
| `horzpos` | 0 | 수평 시작 위치 |
| `horzsize` | 42520 | 수평 크기 (텍스트 영역 너비) |
| `flags` | 393216 | 단락 플래그 |

### 1-3. 텍스트가 있는 단락 (추정 구조)

샘플 파일에 순수 텍스트 단락이 없어서 구조 추정:

```xml
<hp:p id="[고유ID]" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="0">
    <hp:t>실제 텍스트 내용</hp:t>
  </hp:run>
  <hp:linesegarray>
    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000"
                baseline="850" spacing="600" horzpos="0" horzsize="42520" flags="393216"/>
  </hp:linesegarray>
</hp:p>
```

**단락 종료 문자:** `&#x000D;` 사용 여부 불확실 — 샘플에서 미확인.
`<hp:t>` 안의 텍스트로 처리하는 것이 안전.

---

## 2. header.xml 필수 구조 — 실측

### 2-1. 최상위 구조

```xml
<hh:head xmlns:hp="..." xmlns:hh="..." xmlns:hc="...">
  <hh:beginNum page="1" footnote="1" endnote="1"/>    ← 필수
  <hh:refList>                                         ← 필수
    <hh:fontfaces itemCnt="7">...</hh:fontfaces>
    <hh:borderFills itemCnt="2">...</hh:borderFills>
    <hh:charProperties itemCnt="7">...</hh:charProperties>
    <hh:tabProperties itemCnt="3">...</hh:tabProperties>
    <hh:numberings itemCnt="1">...</hh:numberings>
    <hh:paraProperties itemCnt="20">...</hh:paraProperties>
    <hh:styles itemCnt="22">...</hh:styles>
  </hh:refList>
  <hh:compatibleDocument targetProgram="HWP201X">    ← 필수
    <hh:layoutCompatibility/>
  </hh:compatibleDocument>
  <hh:docOption>...</hh:docOption>
  <hh:trackchageConfig flags="56"/>
</hh:head>
```

> ⚠️ `hh:mappingTable`은 실측에서 **존재하지 않음** — 불필요한 요소

### 2-2. hh:charPr (글자 속성) — 실측

```xml
<hh:charPr id="0" height="1000" textColor="#000000" shadeColor="none"
           useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="2">
  <hh:fontRef hangul="1" latin="1" hanja="1" japanese="1" other="1" symbol="1" user="1"/>
  <hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
  <hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
  <hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
  <hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
  <hh:underline type="NONE" shape="SOLID" color="#000000"/>
  <hh:strikeout shape="NONE" color="#000000"/>
  <hh:outline type="NONE"/>
  <hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/>
</hh:charPr>
```

**height="1000"** = 폰트 크기 10pt (100 단위 → pt 환산)

### 2-3. hh:paraPr (단락 속성) — 실측

```xml
<hh:paraPr id="0" tabPrIDRef="0" condense="0" fontLineHeight="0"
           snapToGrid="1" suppressLineNumbers="0" checked="0">
  <hh:align horizontal="JUSTIFY" vertical="BASELINE"/>
  <hh:heading type="NONE" idRef="0" level="0"/>
  <hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="KEEP_WORD"
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
      <hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>
    </hp:case>
    <hp:default>
      <hh:margin>...</hh:margin>
      <hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>
    </hp:default>
  </hp:switch>
  <hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0"
             offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>
</hh:paraPr>
```

**핵심 확인사항:**
- `hh:margin`은 `hp:switch` / `hp:case` / `hp:default` 구조로 감싸야 함 ← 중요!
- `keepWithNext`, `widowOrphan`, `keepLines` 속성이 `hh:breakSetting`에 있음
- 줄간격: `hh:lineSpacing type="PERCENT" value="160"`

### 2-4. hh:styles (스타일 목록) — 실측

```xml
<hh:styles itemCnt="22">
  <hh:style id="0" type="PARA" name="바탕글" engName="Normal"
            paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0"
            langID="1042" lockForm="0"/>
  <hh:style id="1" type="PARA" name="본문" engName="Body"
            paraPrIDRef="1" charPrIDRef="0" nextStyleIDRef="1"
            langID="1042" lockForm="0"/>
  <!-- ... 22개 스타일 -->
</hh:styles>
```

**필수 기본 스타일 (id 0~11 필수 포함):**
- id=0: 바탕글 (Normal)
- id=1: 본문 (Body)
- id=2~11: 개요 1~10 (Outline)
- id=12: 쪽 번호 (CHAR 타입)
- id=13: 머리말, 14: 각주, 15: 미주, 16: 메모

### 2-5. hh:fontfaces — 실측

```xml
<hh:fontfaces itemCnt="7">
  <hh:fontface lang="HANGUL" fontCnt="2">
    <hh:font id="0" face="함초롬돋움" type="TTF" isEmbedded="0">
      <hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="4"
                   contrast="0" strokeVariation="1" armStyle="1"
                   letterform="1" midline="1" xHeight="1"/>
    </hh:font>
    <hh:font id="1" face="함초롬바탕" type="TTF" isEmbedded="0">...</hh:font>
  </hh:fontface>
  <!-- lang: HANGUL, LATIN, HANJA, JAPANESE, OTHER, SYMBOL, USER — 7개 -->
</hh:fontfaces>
```

> lang 7개(HANGUL/LATIN/HANJA/JAPANESE/OTHER/SYMBOL/USER) 모두 있어야 함

---

## 3. 페이지 설정 — 실측 (section0.xml 첫 번째 hp:p 안 hp:secPr)

```xml
<hp:secPr textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000"
          tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1"
          memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0">
  <hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>
  <hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>
  <hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0"
                 border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0"
                 hideFirstEmptyLine="0" showLineNumber="0"/>
  <hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>
  <hp:pagePr landscape="WIDELY" width="59528" height="84186" gutterType="LEFT_ONLY">
    <hp:margin header="4252" footer="4252" gutter="0"
               left="8504" right="8504" top="5668" bottom="4252"/>
  </hp:pagePr>
  <!-- footNotePr, endNotePr, pageBorderFill × 3 (BOTH/EVEN/ODD) 필수 -->
</hp:secPr>
```

**실측 단위 환산:**
- width=59528, height=84186 → A4 (210×297mm)
- left=right=8504 → 약 30mm (8504 / 283.46 ≈ 30mm)
- top=5668 → 약 20mm

---

## 4. 섹션 구조 시작부 패턴 — 실측

첫 번째 `hp:p`는 항상 `hp:secPr` + `hp:colPr`를 포함:

```xml
<hp:p id="[고유ID]" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="0">
    <hp:secPr ...>
      <!-- 페이지 설정 전체 -->
    </hp:secPr>
    <hp:ctrl>
      <hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/>
    </hp:ctrl>
  </hp:run>
  <hp:linesegarray>    ← ⚠️ secPr 단락에도 linesegarray 필수! (이전 분석 오류 수정)
    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000"
                baseline="850" spacing="600" horzpos="0" horzsize="42520" flags="393216"/>
  </hp:linesegarray>
</hp:p>
```

---

## 5. ZIP 패키징 규칙 — 실측 확정 ★ 가장 중요

> ⚠️ **가장 흔한 실수:** manifest.xml과 container.rdf 내용을 뒤바꿔 넣는 것.
> manifest.xml은 거의 빈 파일, 실제 매니페스트 내용은 container.rdf에 있음.

| 항목 | 올바른 값 | 잘못된 값 |
|------|-----------|-----------|
| `mimetype` 압축 방식 | `ZIP_STORED` (압축 없음) | `ZIP_DEFLATED` |
| `mimetype` 순서 | ZIP 첫 번째 항목 필수 | 다른 순서 |
| `flag_bits` | `4` (data descriptor) | `0` |
| `create_system` | `11` (NTFS/Windows) | `0` |

### manifest.xml 올바른 내용 (빈 파일)
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>
```

### container.rdf 올바른 내용 (실제 매니페스트)
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
  <rdf:Description rdf:about="">
    <ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#"
                 rdf:resource="Contents/header.xml"/>
  </rdf:Description>
  <rdf:Description rdf:about="Contents/header.xml">
    <rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile"/>
  </rdf:Description>
  <rdf:Description rdf:about="">
    <ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#"
                 rdf:resource="Contents/section0.xml"/>
  </rdf:Description>
  <rdf:Description rdf:about="Contents/section0.xml">
    <rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/>
  </rdf:Description>
  <rdf:Description rdf:about="">
    <rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document"/>
  </rdf:Description>
</rdf:RDF>
```

---

## 6. 파일 손상 원인 체크리스트 (실제 테스트로 확정)

| 체크 항목 | 손상 영향 | 확인 방법 |
|-----------|-----------|-----------|
| `manifest.xml` ↔ `container.rdf` 내용 뒤바뀜 | 🔴 치명적 — 실제 테스트 확정 | manifest.xml은 빈 self-closing 태그 |
| `hp:linesegarray` 누락 (secPr 단락 포함 전체) | 🔴 치명적 — 실제 테스트 확정 | 빈 단락·secPr 단락도 예외 없음 |
| `hs:sec` xmlns 15개 누락 | 🔴 치명적 — 실제 테스트 확정 | 샘플 루트 태그 xmlns 그대로 복사 |
| `mimetype` ZIP_STORED 아닌 경우 | 🔴 치명적 | compress_type=ZIP_STORED 필수 |
| `hh:mappingTable` 포함 | 🟠 제거 필요 | 실측 샘플에 없음 |
| `hh:margin` hp:switch/case/default 미적용 | 🟠 구버전 오류 | 샘플 구조 그대로 |
| fontfaces lang 7개 미완성 | 🟡 폰트 깨짐 | HANGUL/LATIN/HANJA/JAPANESE/OTHER/SYMBOL/USER |
| styles 기본 22개 미만 | 🟡 스타일 오류 | id=0~21 필수 |
| `hh:compatibleDocument` 누락 | 🟡 구버전 미호환 | targetProgram="HWP201X" |
| `"0.12mm"` 공백 없음 | 🟢 렌더링 이슈 | "0.12 mm" 공백 포함 |
