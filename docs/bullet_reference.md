# 불릿 순서 실측 레퍼런스
> 파일: 불릿순서.hwpx (한글에서 직접 생성한 정상 파일)

---

## 불릿 계층 순서 (확정)

| 레벨 | 기호 | 예시 | charPr | 폰트 | 크기 |
|------|------|------|--------|------|------|
| li1 | `-` | `- 항목` | id=0 | 함초롬바탕 | 10pt |
| li2 | `∙` | ` ∙ 항목` | id=7 | 휴먼명조 | 13pt |
| li3 | `1.` | `  1. 항목` | id=7 | 휴먼명조 | 13pt |
| li4 | `가.` | `   가. 항목` | id=7 | 휴먼명조 | 13pt |

---

## 실측 단락 구조

### li1 (`-`)
```xml
<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="0">
    <hp:t>- 항목 텍스트</hp:t>
  </hp:run>
  <hp:linesegarray>
    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000"
                baseline="850" spacing="600" horzpos="0" horzsize="42520" flags="393216"/>
  </hp:linesegarray>
</hp:p>
```

### li2 (`∙`) — 공백 + 중점 별도 run
```xml
<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="0">
    <hp:t> </hp:t>   <!-- 공백 (들여쓰기) -->
  </hp:run>
  <hp:run charPrIDRef="7">
    <hp:t>∙ 항목 텍스트</hp:t>   <!-- ∙ U+2219 -->
  </hp:run>
  <hp:linesegarray>
    <hp:lineseg textpos="0" vertpos="0" vertsize="1300" textheight="1300"
                baseline="1105" spacing="780" horzpos="0" horzsize="42520" flags="393216"/>
  </hp:linesegarray>
</hp:p>
```

### li3 (`1.`) — 공백 2개 + 번호
```xml
<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="7">
    <hp:t>  1. 항목 텍스트</hp:t>   <!-- 공백 2개 + 번호 -->
  </hp:run>
  <hp:linesegarray>
    <hp:lineseg textpos="0" vertpos="0" vertsize="1300" textheight="1300"
                baseline="1105" spacing="780" horzpos="0" horzsize="42520" flags="393216"/>
  </hp:linesegarray>
</hp:p>
```

### li4 (`가.`) — 공백 3개 + 한글 번호
```xml
<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="7">
    <hp:t>   가. 항목 텍스트</hp:t>   <!-- 공백 3개 + 가나다 -->
  </hp:run>
  <hp:linesegarray>
    <hp:lineseg textpos="0" vertpos="0" vertsize="1300" textheight="1300"
                baseline="1105" spacing="780" horzpos="0" horzsize="42520" flags="393216"/>
  </hp:linesegarray>
</hp:p>
```

---

## charPr 정의

```
charPr id=0: height=1000, fontRef=1(함초롬바탕), color=#000000  → li1 용
charPr id=7: height=1300, fontRef=2(휴먼명조), color=#000000    → li2/li3/li4 용
```

---

## 자동 번호 로직

- `li3`: 문서 내 li3 등장 순서대로 1, 2, 3... 자동증가 (섹션 바뀌면 리셋)
- `li4`: 문서 내 li4 등장 순서대로 가, 나, 다... 자동증가 (li3 바뀌면 리셋)
- `li2`: `∙` 고정 (U+2219, 불릿 중점)

---

## 구현 지시

기존 bullet 블록 li1~li4 매핑을 위 실측 구조로 교체해줘.

변경사항:
1. li1: `-` (기존 그대로, charPr id=0)
2. li2: ` ∙` (공백+중점, charPr id=7, 휴먼명조 13pt) — 기존 `○` 제거
3. li3: `  1.` (공백2+번호 자동증가, charPr id=7)
4. li4: `   가.` (공백3+한글번호 자동증가, charPr id=7)
5. lineseg vertsize/textheight: li1=1000, li2~li4=1300

수정 후 test_bullet_v2.hwpx 생성해줘. 내가 한글에서 열어서 확인할게.
불릿순서.hwpx 레퍼런스 파일도 첨부할게.
