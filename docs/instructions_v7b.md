# hwpx_generator 추가 수정 지시

아직 적용 안 된 수정사항 4가지.
Python API / md 변환 / JSON 변환 3가지 경로 모두 동일하게 적용할 것.

---

## 수정 1. 동그라미 불릿 앞 공백 1칸

### document.py
```python
# add_bullet1() 변경 전
self._add_bullet_paragraph(bs, text, prefix_text=bs.prefix)
# bs.prefix = "\uf06d "

# 변경 후
self._add_bullet_paragraph(bs, text, prefix_text=" \uf06d ")
```

### md_loader.py / markdown_parser.py
- 항목 → bullet1 변환 시 동일하게 prefix " \uf06d " 적용

### json_loader.py
{"type": "bullet", "level": 1} 변환 시 동일 적용

---

## 수정 2. 대시 불릿 앞 공백 2칸

### document.py
```python
# add_bullet2() 변경 전
para.add_run("- " + text, ...)

# 변경 후
para.add_run("  - " + text, ...)
```

### md_loader.py / markdown_parser.py / json_loader.py
동일하게 "  - " + text 적용

---

## 수정 3. SectionBlock 앞 빈 단락 자동 삽입

### xml_builder.py — build_section_xml()

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

v7 실측: Ⅰ~Ⅴ 섹션 앞에 각각 빈 단락 1개 존재.
첫 번째 블록이 SectionBlock인 경우만 빈 단락 생략 (i > 0 조건).

---

## 수정 4. 본문 줄간격 160% 확인

- GovDocumentPreset line_spacing 기본값 = 160 유지
- 제목 단락(paraPr id=10)만 120%
- 본문 / 불릿 / 표 내부 등 나머지는 모두 160%
- 현재 값이 맞는지 확인 후 틀리면 수정

---

## 완료 후 검증

```bash
python generate_notice.py
```

test_notice_v7b.hwpx 생성 후 XML 구조 검증:
- [ ] 동그라미 불릿 run 텍스트가 " \uf06d " 로 시작하는지
- [ ] 대시 불릿 run 텍스트가 "  - " 로 시작하는지
- [ ] SectionBlock(Ⅱ~Ⅴ) 앞에 빈 단락 존재하는지
- [ ] 본문 paraPr lineSpacing value=160 인지
