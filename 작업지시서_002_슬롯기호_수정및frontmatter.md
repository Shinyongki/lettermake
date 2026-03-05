# 작업지시서 #002
**제목**: 슬롯 기호 수정 + frontmatter 설정 지원  
**작성**: Claude.ai  
**대상**: Claude Code  
**상태**: 신규

---

## 배경

외부 스타일 변환 시 ○/-/• 슬롯의 기호가 누락되고 공란 페이지/첫 줄 문제가 확인됨.
추가로 문서마다 기호를 다르게 쓸 수 있도록 마크다운 frontmatter 설정을 지원.

---

## 수정 대상 파일

```
gonggong_hwpxskills-main/scripts/dynamic_builder.py
```

---

## 작업 1 — 버그 수정 (먼저 처리)

### 1-1. 기호 접두사 누락 수정

템플릿 원본 분석 결과:
- □: 2-run 구조 — run[0](" □ ", charPr=0) + run[1](텍스트, charPr=5)
- ○/-/•/※: 1-run 구조 — "기호 + 텍스트" 합쳐서 단일 run

`_clone_and_replace_slot()`에서 타입별로 아래와 같이 처리합니다.

```python
DEFAULT_PREFIX = {
    "square": "□",   # run[0] 유지, run[1] 텍스트만 교체
    "circle": "○",   # "  ○ " + 텍스트 합쳐서 run[0]에 삽입
    "dash":   "-",   # "   - " + 텍스트
    "dot":    "•",   # "    • " + 텍스트
    "note":   "※",   # "     ※ " + 텍스트
}
```

공백(들여쓰기 깊이)은 코드에서 고정. 기호만 교체 대상.

### 1-2. 4단계(•) 슬롯 추가

- `extract_prototypes()`에서 `slot_dot`을 `slot_dash` 원형으로부터 복사 등록
- paraPrIDRef=30, charPrIDRef=18 동일 사용, 기호만 `•`로 적용
- `parse_markdown_sections()`에서 들여쓰기 6칸(`      -`) → `type: "dot"` 파싱

### 1-3. 공란 페이지 제거

`assemble()`에서 파트1 끝과 파트2 시작에 pageBreak 단락이 중복되는 문제 수정.
파트2 시작에 pageBreak 단락을 추가하지 않도록 수정.

### 1-4. 본문 첫 줄 공란 제거

`build_part2()` 시작 시 삽입되는 `paraPrIDRef=0` 빈 단락 제거.

---

## 작업 2 — frontmatter 기호 설정 지원 (작업 1 완료 후 처리)

### 형식

마크다운 파일 맨 앞에 아래 블록을 추가하면 기호를 덮어씁니다.
미지정 항목은 기본값 사용. 공백(들여쓰기)은 변경 불가.

```markdown
---
prefix_square: "■"
prefix_circle: "●"
prefix_dash: "–"
prefix_dot: "◦"
prefix_note: "▶"
---

# 문서 제목
## Ⅰ. 섹션...
```

### 구현

`parse_markdown_sections()` 맨 앞에서 `---` 블록 감지 및 파싱.
외부 라이브러리 없이 `key: "value"` 패턴을 직접 파싱 (표준 라이브러리만 사용).
파싱된 기호값을 `DEFAULT_PREFIX`에 오버라이드 후 기존 로직 진행.

---

## 검증

### 작업 1 검증
`test_3depth.md`로 변환 후 아래 확인:
- □/○/-/• 기호 모두 정상 출력
- 공란 페이지 없음
- 본문 첫 줄 공란 없음

### 작업 2 검증
아래 frontmatter 포함 마크다운으로 변환 후 기호 변경 확인:
```markdown
---
prefix_square: "■"
prefix_circle: "●"
---
```
기호만 바뀌고 들여쓰기는 동일한지 확인.

---

## 주의사항

- 작업 1 완료 후 검증 통과 시 작업 2 진행
- 작업 중 판단이 필요한 상황 발생 시 멈추고 보고할 것
