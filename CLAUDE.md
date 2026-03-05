# 문서작성자 프로젝트

## 프로젝트 개요
마크다운/JSON 문서를 HWPX(한글) 파일로 변환하는 시스템.
두 가지 변환 방식을 선택적으로 사용한다.

## 프로젝트 구조
D:\Project2\문서작성자\
├── CLAUDE.md                        ← 프로젝트 지침
├── cli.py                           ← 변환 실행 진입점
├── hwpx_generator\                  ← 기존 변환 라이브러리
├── gonggong_hwpxskills-main\
│   ├── assets\report-template.hwpx  ← 외부 보고서 템플릿
│   └── scripts\
│       ├── fix_namespaces.py        ← 필수 후처리 스크립트
│       └── dynamic_builder.py       ← 외부 스타일 동적 본문 생성 엔진
├── docs\                            ← 문서
├── tests\                           ← 테스트
└── examples\                        ← 예제

## 환경
- Python 3.8 이상
- 외부 의존성 없음 (표준 라이브러리만 사용)

---

## 변환 방식

### 기존
hwpx_generator 라이브러리 기반 변환.
내용 중심, 공문서/안내문/일반 문서에 적합.

### 외부
gonggong_hwpxskills-main 템플릿 기반 변환.
표지 + 목차 + 섹션바 디자인이 있는 내부 보고서에 적합.
표지/목차는 템플릿 보존, 본문은 동적 생성 (슬롯 제한 없음).

---

## 문서 변환 요청 규칙

사용자가 문서 변환을 요청할 때 아래 기준으로 판단한다.

"기존으로 변환해줘"
→ python cli.py -i [입력파일] -o [출력파일] --style 기존

"외부로 변환해줘"
→ python cli.py -i [입력파일] -o [출력파일] --style 외부

"외부로 변환해줘, 기관명 OOO"
→ python cli.py -i [입력파일] -o [출력파일] --style 외부 --org OOO

스타일 언급 없으면 --style 기존 적용.

---

## 주요 명령어

### 기존 방식
python cli.py -i doc.md -o result.hwpx --style 기존
python cli.py -i doc.json -o result.hwpx --style 기존
python cli.py -i doc.md -o result.hwpx --style 기존 --preset gov_document
python cli.py -i doc.md -o result.hwpx --style 기존 --no-page-flow

### 외부 방식
python cli.py -i doc.md -o result.hwpx --style 외부
python cli.py -i doc.md -o result.hwpx --style 외부 --org 경기도교육청
python cli.py -i doc.md -o result.hwpx --style 외부 --color-main 1F4E79 --color-sub 2E75B6 --color-accent C00000

---

## 입력 파일 포맷

### Markdown (.md)
- 기존/외부 모두 지원
- 외부 스타일은 아래 마크다운 작성 규칙을 따를 것

### JSON (.json)
- 기존 스타일만 지원
- preset 키를 JSON 안에 포함 가능

---

## 외부 스타일 마크다운 작성 규칙

외부 스타일 사용 시 들여쓰기에 따라 아래와 같이 매핑된다.

# 제목           → 표지 제목, 본문 제목
목적: ...         → 표지 하단 목적 박스 (선택)
## Ⅰ. 섹션      → 섹션 바 제목
- 항목           → □ 슬롯 (들여쓰기 0칸)
  - 항목         → ○ 슬롯 (들여쓰기 2칸)
    - 항목       → ― 슬롯 (들여쓰기 4칸)
      - 항목     → • 슬롯 (들여쓰기 6칸)
### 소제목        → 이하 `-` 항목의 들여쓰기 +2 (○부터 시작)
#### 소소제목     → 이하 `-` 항목의 들여쓰기 +2
※ 주석           → ※ 슬롯
| 표 |           → hwpx 테이블로 변환

슬롯 수, 섹션 수 제한 없음 (원형 복제 방식).

### Frontmatter (선택)
기호를 커스터마이즈할 때 마크다운 최상단에 작성:
```
---
prefix_square: "■"
prefix_circle: "●"
---
```
지원 키: prefix_square, prefix_circle, prefix_dash, prefix_dot, prefix_note

### 작성 예시
# 2026년 상반기 업무 추진 계획

목적: 부서별 핵심 과제를 정리하고 효율적인 업무 추진 방향을 공유하기 위함

## Ⅰ. 추진 배경
### 전년도 성과 요약
- 주요 사업 목표 달성률 94% 기록
  - 예산 집행률 98.2%로 전년 대비 3.1%p 향상
※ 2025년 행안부 조사 기준

## Ⅱ. 추진 일정
| 구분 | 1분기 | 2분기 | 담당 |
|------|-------|-------|------|
| 전략 수립 | 완료 | 점검 | 기획조정팀 |

---

## 기존 스타일 추가 옵션

--preset gov_document : 공문서 스타일 프리셋 적용
--no-page-flow        : 자동 페이지 나누기 비활성화

---

## 주의사항
- fix_namespaces.py 는 subprocess.run() 으로 실행 (exec() 사용 금지)
- 외부 스타일은 .md 파일만 지원 (.json 미지원)
- 공문서 날짜 형식: 2026. 3. 5. (월·일 앞 0 생략)
- dynamic_builder.py: 외부 스타일 동적 본문 생성 엔진 (섹션바/슬롯 원형 복제 방식)
- 색상 옵션(--color-main 등) 미지정 시 기본값 자동 적용 (1F4E79 / 2E75B6 / C00000)

## 외부 스타일 기술 참고

### 템플릿 구조 (section0.xml)
- pageBreak="1" 기준으로 3파트 분리: part0(표지), part1(목차), part2(본문)
- part2에서 원형(prototype) 추출 후 deepcopy로 동적 생성

### 슬롯 원형 (paraPrIDRef / charPrIDRef)
- □ square: pr=28, char=5 (2-run 구조: run[0]=기호, run[1]=텍스트)
- ○ circle: pr=29, char=18 (1-run 구조: 인라인 기호+텍스트)
- ― dash: pr=30, char=18
- • dot: pr=30, char=18 (dash 복제)
- ※ note: pr=30, char=20

### 페이지/마진
- pagePr: width=59528, height=84188
- margin: left=5669, right=5669 → content_w=48190

### 표지 장식선
- 테이블 밖 독립 `<hp:p>`로 삽입 (treatAsChar="0", textWrap="TOP_AND_BOTTOM")
- Row0/Row2는 배경 제거만 수행

### 섹션바
- 3셀 구조: Cell0(로마숫자, charPr=24 흰색), Cell1(공백), Cell2(제목, charPr=2 검정)
- 좌측 셀: LINEAR 그라데이션 (color_main → 밝은 계열)
- 우측 셀: 회색→흰색 그라데이션

### borderFill 할당
- 1~14: 템플릿 기본
- 15: 데이터 테이블 셀 (테두리만)
- 16: 헤더 테이블 셀 (테두리 + 배경 D9D9D9)
- 17: 목적 박스 (상하 SOLID 0.4mm, 좌우 NONE, 배경 #FFFFFF)

---

## 클로드 코드 행동 규칙

### CLAUDE.md 자동 업데이트 규칙
아래 상황이 발생하면 작업 완료 후 반드시 사용자에게 먼저 물어볼 것:

"이번 작업에서 아래 내용이 추가/변경됐습니다.
CLAUDE.md에 업데이트할까요?
- [변경 항목 1]
- [변경 항목 2]"

### 업데이트 물어볼 타이밍
- 새 기능이 추가됐을 때
- 새 규칙이나 약속이 정해졌을 때
- 폴더/파일 구조가 바뀔 때
- 오류 해결 후 주의사항이 생겼을 때
- 작업 세션을 끝낼 때

### 사용자가 "응" 또는 "추가해줘" 하면
CLAUDE.md를 직접 수정하고 완료 보고할 것.

---

## 협업 방식

### 도구 역할 분담
- Claude.ai: 전략/방향 결정, 복잡한 판단, 이미지 분석, 지시문 작성
- Claude Code: 직접 코딩 실행, 파일 작업, 변환 실행

### 정보 전달 방식
- Claude.ai → Claude Code: 파일로 저장 후 전달 (붙여넣기 대신)
- Claude Code → Claude.ai: 텍스트 결과는 복사, 이미지는 캡처

### Claude Code 행동 규칙
1. 방향이 불명확하거나 복잡한 판단이 필요하면 작업 중단 후 사용자에게 알릴 것
   "이 부분은 Claude.ai에서 방향을 잡고 오시면 좋을 것 같습니다"
2. 작업 완료 후 CLAUDE.md 업데이트 필요 여부를 반드시 먼저 물어볼 것
3. 오류 발생 시 원인과 함께 Claude.ai 공유가 필요한지 판단해서 알려줄 것
