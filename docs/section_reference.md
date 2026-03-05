# section 블록 실측 레퍼런스
> 파일: 섹션.hwpx (한글에서 직접 생성한 정상 파일)

---

## section 블록 완전한 XML 구조

```xml
<hp:tbl id="[ID]" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM"
        textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL"
        repeatHeader="0" rowCnt="2" colCnt="3" cellSpacing="0"
        borderFillIDRef="3" noAdjust="0">
  <hp:sz width="30670" widthRelTo="ABSOLUTE" height="2715" heightRelTo="ABSOLUTE" protect="0"/>
  <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="0" allowOverlap="0"
          holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN"
          vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>
  <hp:outMargin left="0" right="0" top="0" bottom="0"/>
  <hp:inMargin left="0" right="0" top="0" bottom="0"/>
  <hp:tr>
    <!-- 1열: 번호 (rowSpan=2 병합) — 파란 배경 -->
    <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="4">
      <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                  linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0"
                  hasTextRef="0" hasNumRef="0">
        <hp:p id="2147483648" paraPrIDRef="21" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
          <hp:run charPrIDRef="8">
            <hp:t>1</hp:t>   <!-- 번호: 자동증가 -->
          </hp:run>
          <hp:linesegarray>
            <hp:lineseg textpos="0" vertpos="0" vertsize="1800" textheight="1800"
                        baseline="1530" spacing="1440" horzpos="0" horzsize="2848" flags="393216"/>
          </hp:linesegarray>
        </hp:p>
      </hp:subList>
      <hp:cellAddr colAddr="0" rowAddr="0"/>
      <hp:cellSpan colSpan="1" rowSpan="2"/>
      <hp:cellSz width="2850" height="2715"/>
      <hp:cellMargin left="0" right="0" top="0" bottom="0"/>
    </hp:tc>

    <!-- 2열: 구분선 (rowSpan=2 병합) — 테두리만 -->
    <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="5">
      <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                  linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0"
                  hasTextRef="0" hasNumRef="0">
        <hp:p id="2147483648" paraPrIDRef="22" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
          <hp:run charPrIDRef="9"/>
          <hp:linesegarray>
            <hp:lineseg textpos="0" vertpos="0" vertsize="1700" textheight="1700"
                        baseline="1445" spacing="1360" horzpos="0" horzsize="1440" flags="393216"/>
          </hp:linesegarray>
        </hp:p>
      </hp:subList>
      <hp:cellAddr colAddr="1" rowAddr="0"/>
      <hp:cellSpan colSpan="1" rowSpan="2"/>
      <hp:cellSz width="565" height="2715"/>
      <hp:cellMargin left="0" right="0" top="0" bottom="0"/>
    </hp:tc>

    <!-- 3열 1행: 제목 텍스트 — 그라데이션 배경 -->
    <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="6">
      <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                  linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0"
                  hasTextRef="0" hasNumRef="0">
        <hp:p id="2147483648" paraPrIDRef="23" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
          <hp:run charPrIDRef="10">
            <hp:t>추진 배경</hp:t>   <!-- 제목 텍스트 -->
          </hp:run>
          <hp:linesegarray>
            <hp:lineseg textpos="0" vertpos="0" vertsize="1700" textheight="1700"
                        baseline="1445" spacing="852" horzpos="0" horzsize="26688" flags="393216"/>
          </hp:linesegarray>
        </hp:p>
      </hp:subList>
      <hp:cellAddr colAddr="2" rowAddr="0"/>
      <hp:cellSpan colSpan="1" rowSpan="1"/>
      <hp:cellSz width="27255" height="2565"/>
      <hp:cellMargin left="565" right="0" top="140" bottom="0"/>
    </hp:tc>
  </hp:tr>

  <hp:tr>
    <!-- 3열 2행: 얇은 하단 구분선 (height=150) — 그라데이션 다른 색 -->
    <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="7">
      <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                  linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0"
                  hasTextRef="0" hasNumRef="0">
        <hp:p id="2147483648" paraPrIDRef="24" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
          <hp:run charPrIDRef="11"/>
          <hp:linesegarray>
            <hp:lineseg textpos="0" vertpos="0" vertsize="150" textheight="150"
                        baseline="128" spacing="120" horzpos="600" horzsize="26652" flags="393216"/>
          </hp:linesegarray>
        </hp:p>
      </hp:subList>
      <hp:cellAddr colAddr="2" rowAddr="1"/>
      <hp:cellSpan colSpan="1" rowSpan="1"/>
      <hp:cellSz width="27255" height="150"/>
      <hp:cellMargin left="0" right="0" top="0" bottom="0"/>
    </hp:tc>
  </hp:tr>
</hp:tbl>
```

---

## borderFill 정의 (header.xml에 추가)

```xml
<!-- id=3: section 표 외곽 — 테두리 없음 -->
<hh:borderFill id="3" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">
  <hh:slash type="NONE" Crooked="0" isCounter="0"/>
  <hh:backSlash type="NONE" Crooked="0" isCounter="0"/>
  <hh:leftBorder type="NONE" width="0.1 mm" color="none"/>
  <hh:rightBorder type="NONE" width="0.1 mm" color="none"/>
  <hh:topBorder type="NONE" width="0.1 mm" color="none"/>
  <hh:bottomBorder type="NONE" width="0.1 mm" color="none"/>
</hh:borderFill>

<!-- id=4: 1열 번호 셀 — 파란 배경 (#1F5B9B), 회색 테두리 -->
<hh:borderFill id="4" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">
  <hh:slash type="NONE" Crooked="0" isCounter="0"/>
  <hh:backSlash type="NONE" Crooked="0" isCounter="0"/>
  <hh:leftBorder type="SOLID" width="0.4 mm" color="#999999"/>
  <hh:rightBorder type="SOLID" width="0.4 mm" color="#999999"/>
  <hh:topBorder type="SOLID" width="0.4 mm" color="#999999"/>
  <hh:bottomBorder type="SOLID" width="0.4 mm" color="#999999"/>
  <hc:fillBrush>
    <hc:winBrush faceColor="#1F5B9B" hatchColor="none" alpha="0"/>
  </hc:fillBrush>
</hh:borderFill>

<!-- id=5: 2열 구분선 셀 — 좌우 테두리만, 배경 없음 -->
<hh:borderFill id="5" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">
  <hh:slash type="NONE" Crooked="0" isCounter="0"/>
  <hh:backSlash type="NONE" Crooked="0" isCounter="0"/>
  <hh:leftBorder type="SOLID" width="0.4 mm" color="#999999"/>
  <hh:rightBorder type="SOLID" width="0.4 mm" color="#999999"/>
  <hh:topBorder type="NONE" width="0.1 mm" color="none"/>
  <hh:bottomBorder type="NONE" width="0.1 mm" color="none"/>
</hh:borderFill>

<!-- id=6: 3열 1행 제목 셀 — 왼쪽 테두리 + 그라데이션 (#DFEAF5 → #FFFFFF) -->
<hh:borderFill id="6" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">
  <hh:slash type="NONE" Crooked="0" isCounter="0"/>
  <hh:backSlash type="NONE" Crooked="0" isCounter="0"/>
  <hh:leftBorder type="SOLID" width="0.4 mm" color="#999999"/>
  <hh:rightBorder type="NONE" width="0.1 mm" color="none"/>
  <hh:topBorder type="NONE" width="0.1 mm" color="none"/>
  <hh:bottomBorder type="NONE" width="0.1 mm" color="none"/>
  <hc:fillBrush>
    <hc:gradation type="LINEAR" angle="90" centerX="0" centerY="0"
                  step="250" colorNum="2" stepCenter="50" alpha="0">
      <hc:color value="#DFEAF5"/>
      <hc:color value="#FFFFFF"/>
    </hc:gradation>
  </hc:fillBrush>
</hh:borderFill>

<!-- id=7: 3열 2행 구분선 셀 — 왼쪽 테두리 + 그라데이션 (#999999 → #FFFFFF) -->
<hh:borderFill id="7" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">
  <hh:slash type="NONE" Crooked="0" isCounter="0"/>
  <hh:backSlash type="NONE" Crooked="0" isCounter="0"/>
  <hh:leftBorder type="SOLID" width="0.4 mm" color="#999999"/>
  <hh:rightBorder type="NONE" width="0.1 mm" color="none"/>
  <hh:topBorder type="NONE" width="0.1 mm" color="none"/>
  <hh:bottomBorder type="NONE" width="0.1 mm" color="none"/>
  <hc:fillBrush>
    <hc:gradation type="LINEAR" angle="90" centerX="0" centerY="0"
                  step="255" colorNum="2" stepCenter="50" alpha="0">
      <hc:color value="#999999"/>
      <hc:color value="#FFFFFF"/>
    </hc:gradation>
  </hc:fillBrush>
</hh:borderFill>
```

---

## charPr 정의 (header.xml에 추가)

```
charPr id=8:  height=1800, fontRef=3(휴먼명조) — 번호 셀 (흰색 텍스트)
charPr id=9:  height=1700, fontRef=5(고딕) — 구분선 셀
charPr id=10: height=1700, fontRef=4(HY헤드라인M) — 제목 텍스트 17pt
charPr id=11: height=150,  fontRef=0(맑은 고딕) — 하단 구분선
```

**charPr id=8 (번호 — 흰색 텍스트):**
```xml
<hh:charPr id="8" height="1800" textColor="#FFFFFF" shadeColor="none"
           useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="2">
  <hh:fontRef hangul="3" latin="3" hanja="3" japanese="3" other="3" symbol="3" user="3"/>
  <!-- ratio, spacing, relSz, offset 모두 기본값 100/0 -->
  <hh:underline type="NONE" shape="SOLID" color="#000000"/>
  <hh:strikeout shape="NONE" color="#000000"/>
  <hh:outline type="NONE"/>
  <hh:shadow type="NONE" color="#B2B2B2" offsetX="10" offsetY="10"/>
</hh:charPr>
```

**charPr id=10 (제목 텍스트 — HY헤드라인M 17pt):**
```xml
<hh:charPr id="10" height="1700" textColor="#000000" shadeColor="none"
           useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="2">
  <hh:fontRef hangul="4" latin="4" hanja="4" japanese="4" other="4" symbol="4" user="4"/>
  <hh:underline type="NONE" shape="SOLID" color="#000000"/>
  <hh:strikeout shape="NONE" color="#000000"/>
  <hh:outline type="NONE"/>
  <hh:shadow type="NONE" color="#B2B2B2" offsetX="10" offsetY="10"/>
</hh:charPr>
```

---

## 구현 지시

기존 section 블록을 위 실측 구조로 전면 교체해줘.

핵심 변경사항:
1. borderFill id=4~7 추가 (그라데이션 포함)
2. charPr id=8~11 추가 (번호 흰색, 제목 HY헤드라인M 17pt)
3. 번호 셀 텍스트 색상 흰색 (#FFFFFF)
4. 3열 너비 27255 유지 → 텍스트 1행 처리
5. 번호 자동증가 로직 유지 (1, 2, 3...)

수정 후 test_section_v2.hwpx 생성해줘. 내가 한글에서 열어서 확인할게.
섹션.hwpx 레퍼런스 파일도 첨부할게.
