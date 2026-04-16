---
name: hwpx
description: HWPX 문서(.hwpx 파일)를 생성, 읽기, 편집, 템플릿 관리하는 스킬. '한글 문서', 'hwpx', 'HWPX', '한글파일', '.hwpx 파일 만들어줘', 'HWP 문서 생성', '보고서 작성', '공문', '기안문', '한글로 작성', '양식', '서식', '템플릿 등록' 등의 키워드가 나오면 반드시 이 스킬을 사용할 것. 한글과컴퓨터(한컴)의 HWPX 포맷(OWPML 기반, ZIP+XML 구조)을 직접 zipfile+lxml로 다루는 hwpx_engine을 사용한다.
---

# HWPX 문서 플러그인

HWPX 문서를 생성, 편집, 검증하는 플러그인. 양식(template)을 등록하면 빌드 스크립트로 문서를 생성하고, HwpxEditor로 기존 문서를 편집한다.

## 핵심 워크플로우

```
1. 초안 생성: 빌드 스크립트 작성 → build() 실행 → .hwpx 생성
2. 수정(빌드 스크립트 있음): 빌드 스크립트를 Edit 도구로 부분 수정 → 재빌드
3. 수정(기존 .hwpx 직접): HwpxEditor로 편집 스크립트 작성 → 실행
4. PDF 검증 (Windows+한글): pyhwpx로 PDF 변환 → Read 도구로 시각 확인
5. 버전관리: 스크립트 파일이 소스 → git으로 이력 관리
```

### 대규모 편집 워크플로우 (3-pass)

기존 문서를 대규모로 편집할 때 반드시 아래 순서를 따른다:

```
PASS 1 (ZIP): batch_replace — 텍스트 일괄 교체
PASS 2 (DOM): delete_row/delete_table/remove_paragraph — 구조 삭제
PASS 3 (DOM): set_cell/batch_set_cell/add_row/set_paragraph_style — 셀·행·스타일 수정
각 PASS 사이에 save → reopen 필수
```

상세: `references/document-editing-guide.md` 참조

### PDF 검증 규칙

- **Windows + 한글 설치 환경에서만 가능** (pyhwpx는 COM 자동화 기반, macOS/Linux 미지원)
- **새 양식 분석/등록 테스트**: PDF 변환 후 시각 확인 **필수**
- **문서 생성/편집**: 완료 후 사용자에게 "PDF로 변환해서 확인해볼까요?" 질문
- Windows가 아니거나 한글 미설치 시: "한글에서 직접 열어서 확인해주세요" 안내
- 상세: `references/pdf-verification-guide.md` 참조

## 의사결정 트리

| 요청 | 방법 |
|------|------|
| "보고서 만들어줘" | 등록 양식의 usage-guide.md 읽기 → 빌드 스크립트 작성 → 실행 |
| "2장에 표 추가해줘" | 기존 빌드 스크립트를 Edit 도구로 수정 → 재실행 |
| "제목 바꿔줘" | 스크립트의 cover_replacements 수정 → 재실행 |
| "이 양식으로 문서 만들어줘" | 양식 분석 → 등록 → 빌드 스크립트 작성 |
| "이미 만든 .hwpx 텍스트 수정" | `HwpxEditor.str_replace()` |
| "이미 만든 .hwpx에 새 문단 추가" | `HwpxEditor.open(path, template_id=...)` → `insert_paragraph()` |
| "이미 만든 .hwpx 표 셀 수정" | `editor.set_cell(table_index, row, col, text)` |
| "여러 셀 한꺼번에 수정" | `editor.batch_set_cell(table_index, [(r,c,text), ...])` |
| "표 안에서만 특정 텍스트 교체" | `editor.replace_in_table(table_index, old, new)` |
| "이미 만든 .hwpx에 새 표 추가" | `editor.insert_table(rows_data, ..., after=앵커)` |
| "이미 만든 .hwpx에 그림 추가" | `editor.insert_figure_box(width, height, ..., after=앵커)` |
| 기존 표에서 행 삭제 + 내용 교체 | `find_table`→ 인덱스 확보 → `delete_row`(뒤→앞) → `batch_set_cell`(새 주소) |
| "표/그림 캡션 스타일 변경" | `editor.get_nearby_paragraphs(idx)` → 확인 → `editor.set_paragraph_style(idx, offset, ...)` |
| "새 양식 분석해줘" | `references/template-analysis-guide.md` 따라 분석 |
| "문서에서 텍스트 검색" | `editor.find_text(pattern)` — 패턴과 일치하는 모든 위치 반환 |
| "표 셀 값 읽기" | `editor.get_cell(table_index, row, col)` — 셀 텍스트 반환 |
| "표 전체 데이터 추출" | `editor.get_table_data(table_index)` — 2D 리스트 반환 |
| "표 내용으로 테이블 검색" | `editor.find_table(text_pattern)` — 패턴을 포함하는 표 인덱스 목록 반환 |

## 등록된 양식 조회 및 선택

서식을 고르기 전에 **반드시** `list_templates()`로 목록을 확인한다. ID 축약명만 보고 추측하지 말 것.

```python
from hwpx_engine import list_templates
templates = list_templates()
for t in templates:
    print(t["id"], "|", t["display_name"], "|", t["summary"], "|", t["status"])
```

리턴 각 엔트리 필드:
- `id`: 서식 식별자 (build/editor 호출 시 사용)
- `display_name`: 사람이 읽는 이름
- `summary`: ⭐ 1줄 목적 설명 (매칭의 핵심)
- `description`: 상세 설명 (선택)
- `path`: 디스크 경로 (usage-guide.md 등 읽을 때)
- `status`: `"ok"` / `"incomplete"` / `"invalid"`
- `missing_fields`: incomplete 시 비어있는 필드 목록
- `error`: invalid 시 예외 메시지

### 서식 선택 규칙

1. `list_templates()` 호출 → 전체 목록 확보
2. 사용자 요청과 각 엔트리의 `display_name` + `summary` 매칭
3. **사용자의 요청 문구가 어떤 서식 하나의 ID/display_name에 명백히 1:1 매칭되지 않으면, summary 전체를 사용자에게 제시하고 선택받는다.** 유사도 판단 금지 — 확신이 없으면 무조건 사용자에게 묻는다.
4. ID 이름만 보고 고르지 말 것

### 불완전/손상 서식 처리

목록에 `status != "ok"` 엔트리가 있으면:

- **incomplete**: `repair_template_metadata(id)` 호출 → `filled`에 채워진 필드 사용자에게 제시
  - `warnings_summary_missing`에 `summary`가 남아 있으면 자동으로 채우지 않는다 (repair가 일부러 비워둠). 에이전트는 사용자에게 "이 서식의 한 줄 요약을 알려주세요" 라고 묻거나, `template.hwpx`와 `builder.py`를 직접 읽고 요약을 초안으로 제시한 뒤 사용자 확인을 받는다
  - `missing_fields`에 스타일 매핑 관련 항목이 있으면 `references/template-analysis-guide.md` 절차로 심층 분석 후 재등록
- **invalid**: `error` 메시지 사용자에게 보고 + 재등록 또는 수동 수정 요청

**양식이 없는 경우**: 사용자에게 .hwpx 양식 파일을 요청 → 양식 등록 워크플로우 진행.

## 문서 생성 절차

### 1단계: usage-guide.md 읽기

**반드시** `list_templates()`의 결과에서 해당 `id`의 `path`를 구한 뒤 `{path}/usage-guide.md`를 읽어 content dict 구성법을 파악한다.

### 2단계: 빌드 스크립트 작성

```python
# scripts/build_문서명.py
from hwpx_engine.build import build

content = {
    'cover_replacements': { ... },   # 표지 텍스트 교체
    'global_replacements': { ... },  # 전 섹션 텍스트 교체 (머리말 등)
    'chapters': [ ... ],             # 본문 (절/소절/표/그림/이미지)
    'references': [ ... ],
    # ... (양식별 usage-guide.md 참조)
}

result = build('template_id', content, 'output/document.hwpx')
```

### 3단계: 실행

```bash
python3 scripts/build_문서명.py
```

### 수정 시: Edit 도구로 스크립트 부분 수정 → 재실행

## 기존 문서 편집 (HwpxEditor)

### 텍스트/수치 수정 (template_id 불필요)

```python
from hwpx_engine.editor import HwpxEditor

editor = HwpxEditor.open('output.hwpx')
editor.str_replace('기존 텍스트', '새 텍스트')
editor.save('output_modified.hwpx')
```

### 새 문단 삽입 (template_id로 올바른 스타일 자동 적용)

```python
editor = HwpxEditor.open('output.hwpx', template_id='your_template_id')

# 단순 삽입
editor.insert_paragraph('새 내용', style='bullet_2', after='기존 텍스트')

# 인라인 서식 (볼드/이탤릭/색상/크기)
editor.insert_paragraph(parts=[
    {'text': '일반 텍스트 '},
    {'text': '볼드', 'bold': True},
    {'text': ' 빨간색', 'color': '#FF0000'},
    {'text': ' 큰 글씨', 'size': 1400},  # 14pt
], style='bullet_1', after='기존 텍스트')
```

**주의**: template_id는 해당 등록 양식으로 생성된 문서에서만 유효.

## content dict 핵심 구조

content dict의 정확한 구조는 양식별 usage-guide.md에 정의된다. 공통 패턴:

```python
content = {
    'cover_replacements': {원본텍스트: 새텍스트},  # 빈 문자열 금지!
    'global_replacements': {원본: 새것},            # 전 섹션 적용
    'chapters': [ ... ],                           # 본문
    'references': ['...'],
}
```

### items 타입

| type | 설명 | 주요 필드 |
|------|------|----------|
| (기본) | 불릿 문단 | `style`, `text` 또는 `parts` |
| `table` | 표 | `caption`, `headers`, `rows`, `unit`, `source`, `merges` |
| `figure` | 빈 그림 박스 | `caption`, `source` |
| `image` | 실제 이미지 삽입 | `path`, `caption`, `source` |

### 인라인 서식 (parts)

`text` 대신 `parts`를 사용하면 한 문단 안에서 서로 다른 서식을 적용할 수 있다.

```python
{'style': 'bullet_2', 'parts': [
    {'text': '일반 텍스트 '},
    {'text': '볼드', 'bold': True},
    {'text': ' 빨간색', 'color': '#FF0000'},
    {'text': ' 볼드+이탤릭+파랑', 'bold': True, 'italic': True, 'color': '#0000FF'},
    {'text': ' 큰 글씨', 'size': 1400},  # 14pt (0.1pt 단위)
]}
```

**지원 속성**: `bold`, `italic`, `color` (#RRGGBB), `size` (int, 0.1pt 단위)

build()와 HwpxEditor 모두 동일한 parts 포맷을 지원한다.

### 셀병합 (merges)

```python
'merges': [
    {'row': 0, 'col': 0, 'rowspan': 2, 'colspan': 1},
    {'row': 0, 'col': 1, 'rowspan': 1, 'colspan': 2},
]
```

HWPX에서 병합된 셀은 XML에 **존재하지 않음**. headers 길이 = 전체 컬럼 수.

## 절대 금지 규칙

1. **빈 문자열(`''`) 교체 금지** — cover_replacements에서 `''`로 교체하면 문서 손상
2. **기존 문단의 `<hp:t>` 노드 삭제 금지** — 문단 전체를 삭제하고 새 문단으로 삽입
3. **스타일 ID 하드코딩 금지** — 양식마다 전부 다름. metadata.json에서 로드
4. **charPrIDRef 추정 금지** — 반드시 원본 양식에서 추출한 정확한 값 사용
5. **새 문단에 linesegarray 누락 금지** — xml_primitives가 자동 처리
6. **`fillBrush`/`winBrush` 네임스페이스는 `hc:`** — `hh:`로 쓰면 배경색 무시됨
7. **charProperties/borderFills의 `itemCnt` 업데이트 필수** — CharPrManager가 자동 처리
8. **표 안 텍스트를 anchor로 사용 시 주의** — 표 밖 텍스트를 anchor로 사용할 것
9. **DOM 직접 조작 금지** — `p.addnext()`, `sec.mark_dirty()` 등 lxml DOM 직접 조작은 save() 시 소실됨. 반드시 editor의 메서드(`set_cell`, `insert_table`, `insert_figure_box` 등)를 사용할 것
10. **TableEditor 직접 생성 금지** — `TableEditor(editor.doc)`로 직접 만들면 `str_replace()` 후 참조가 끊어져 변경이 소실됨. 반드시 `editor.set_cell()`, `editor.add_row()` 등 editor 메서드 사용
11. **`make_table_xml()`/`make_figure_box()` 이중 래핑 금지** — 이 함수들은 `<hp:p>` 전체를 반환하므로 `make_para()`로 다시 감싸면 이중 중첩되어 렌더링 실패
12. **set_cell로 표 헤더 행(R0) 덮어쓰기 금지** — 표 캡션/제목은 batch_replace로 변경하고, 셀 내용은 set_cell로 변경할 것. 헤더 행의 컬럼명("소속", "구분" 등)을 다른 텍스트로 덮어쓰면 표 구조가 깨짐
13. **delete_row 후 set_cell 시 원래 행 주소 사용 금지** — delete_row는 _reindex_addresses를 호출하여 모든 cellAddr.rowAddr를 갱신함. 별도 pass에서 set_cell 할 때는 삭제 후 새 주소를 사용해야 함. 같은 pass에서 set_cell→delete_row 순으로 하면 원래 주소를 쓸 수 있음
14. **`get_cell()`/`get_table_data()`는 읽기 전용이다** — 반환값을 수정해도 원본에 반영되지 않는다. 셀을 수정하려면 반드시 `set_cell()`을 사용할 것
15. **테이블 인덱스 하드코딩 금지** — `editor.set_cell(42, ...)` 같이 고정 인덱스를 사용하면 표 순서 변경 시 깨짐. 반드시 `editor.find_table('패턴')` → 반환된 인덱스 사용
16. **수동 rowAddr/rowSpan/height 조작 금지** — `cellAddr.set('rowAddr', ...)` 등 직접 XML 속성 조작은 `editor.delete_row()`/`editor.add_row()`가 자동 처리함. 수동 조작 시 크래시 위험
17. **수동 XML 파싱으로 셀 텍스트 교체 금지** — `tc.find(HP+'t').text = '...'` 대신 `editor.set_cell()` 또는 `editor.batch_set_cell()` 사용. 수동 교체는 formulaScript LastResult 동기화가 누락됨
18. **표 내부만 바꿀 때 str_replace/batch_replace 사용 금지** — 문서 전체가 아닌 특정 표 안에서만 교체할 때는 `editor.replace_in_table()` 사용. 전체 치환은 의도치 않은 곳이 바뀔 위험

## 테스트 실행

```bash
pytest                    # 유닛 + 통합
pytest -m pdf             # PDF 변환 테스트 (한글 설치 필요)
pytest --cov=hwpx_engine  # 커버리지
```

## 참조 문서

| 상황 | 파일 |
|------|------|
| **양식 분석 방법** (필독) | `references/template-analysis-guide.md` |
| **PDF 변환 검증** | `references/pdf-verification-guide.md` |
| **기존 표 편집** | `references/table-editing-guide.md` |
| **대규모 문서 편집** | `references/document-editing-guide.md` |
| 등록된 양식 사용법 | `list_templates()` 로 `path` 확인 후 `{path}/usage-guide.md` |

## 양식 저장 위치

양식은 **글로벌 경로 한 곳**에서 관리한다:

### 서식 저장 경로

```
~/.claude/hwpx-engine/registered/{id}/          ← 글로벌 영구 저장소 (유일한 공식 경로)
~/.claude/hwpx-engine/.trash/{id}_{stamp}/      ← unregister 시 이동 (복구 가능)
```

글로벌 경로는 플러그인 업데이트에도 유지된다. 에이전트가 양식을 수정할 때도 이 경로의 파일을 직접 수정한다.

### 등록/수정/삭제 API

```python
from hwpx_engine import register_template, unregister_template, repair_template_metadata

# 신규 등록 (기존 ID 충돌 시 에러)
register_template('my_template', './작업폴더/')

# 강제 덮어쓰기 (명시적 의도 필요)
register_template('my_template', './new_source/', force=True)

# 삭제 (기본: .trash/로 백업)
unregister_template('my_template')

# 완전 삭제
unregister_template('my_template', backup=False)

# 불완전 메타 자동 보강 (display_name만 — summary/description은 사용자 판단)
result = repair_template_metadata('my_template')
# → {"filled": ["display_name"], "skipped": [...], "warnings_summary_missing": [...]}
```

## 양식 등록 워크플로우

새 양식 파일(.hwpx)을 받으면 아래 싸이클을 **모든 테스트가 통과할 때까지** 자동 반복한다.

### 전체 싸이클

```
1. 분석: references/template-analysis-guide.md 따라 스타일 매핑 추출
2. 에셋 초안 작성: metadata.json, usage-guide.md, builder.py, modules/
3. 테스트 문서 생성: build()로 전체 모듈 테스트 문서 빌드
4. PDF 변환: pyhwpx로 변환 (references/pdf-verification-guide.md)
5. PDF 시각 확인: Read 도구로 주요 페이지 검토
6. 실패 시 → 에셋 수정 후 3번으로 돌아가 반복
7. 빌드 테스트 통과 → 편집 테스트 실행 (아래 항목)
8. PDF 변환 → 시각 확인
9. 실패 시 → 에셋 또는 엔진 수정 후 7번으로 반복
10. 전체 통과 → 글로벌 등록: register_template(id, source_dir)
```

**에이전트는 사용자 개입 없이 6→3, 9→7 루프를 자동 반복한다. 모든 항목이 통과해야 싸이클 종료.**

### 빌드 테스트 범위

build()로 생성한 문서에서 **해당 양식의 모든 모듈과 스타일**이 올바르게 동작하는지 확인한다. 테스트 대상은 양식의 `builder.py`에 선언된 섹션 파이프라인과 `metadata.json`의 스타일 목록에서 도출한다.

**확인 원칙:**
- `builder.py`의 **모든 모듈** → 각 모듈이 생성하는 섹션이 정상 렌더링되는지
- `metadata.json`의 **모든 스타일** → 각 스타일이 적용된 문단/표/그림이 올바른지
- `cover_replacements` / `global_replacements` → 텍스트 교체가 빠짐없이 반영되는지
- 표, 그림, 이미지 등 **content dict에서 지원하는 모든 item type** → 각각 최소 1회 포함
- 인라인 서식 (parts) → 볼드, 이탤릭, 색상, 크기 혼합이 정상 렌더링되는지

**예시** (양식에 따라 대상이 달라짐):

| 대상 | 확인 사항 |
|------|----------|
| 표지 (cover) | cover_replacements 반영, 레이아웃 유지 |
| 판권/정책건의 | global_replacements 반영 |
| 목차 (toc) | 자동 생성 페이지 번호 정상 |
| 본문 각 장 (body) | 간지, 절/소절 제목, 모든 불릿 스타일 (bullet_1~3) |
| 표 (table) | 캡션, 헤더/바디 셀 스타일, 단위, 출처, 셀병합 |
| 그림 (figure) | 빈 그림 박스, 캡션, 출처 |
| 이미지 (image) | 실제 이미지 삽입, 크기, 캡션 |
| 참고문헌 (references) | 번호 매기기, 들여쓰기 |
| Abstract/키워드 | 영문 텍스트 렌더링 |
| 인라인 서식 (parts) | 볼드, 이탤릭, 색상, 크기 혼합 |

### 편집 테스트 범위

HwpxEditor로 기존 문서를 편집하여 **모든 편집 연산**이 정상 동작하는지 확인:

| 연산 | 메서드 | 확인 사항 |
|------|--------|----------|
| 텍스트 교체 | `editor.str_replace()` | 본문 텍스트 정확히 교체됨 |
| 표 셀 교체 | `editor.set_cell()` | 특정 셀 텍스트 변경됨 |
| 표 로우 추가 | `editor.add_row()` | 행 추가 후 rowCnt 증가 |
| 표 로우 삭제 | `editor.delete_row()` | 행 삭제 후 rowCnt 감소 |
| 표 칼럼 추가 | `editor.add_column()` | 열 추가 후 colCnt 증가 |
| 표 칼럼 삭제 | `editor.delete_column()` | 열 삭제 후 colCnt 감소 |
| 문단 삽입 (서식) | `editor.insert_paragraph(parts=...)` | 볼드 등 인라인 서식 유지 |
| 문단 삽입 (단순) | `editor.insert_paragraph(text=...)` | 올바른 스타일 적용 |
| 표 삽입 | `editor.insert_table()` | 캡션, 헤더/바디 스타일, 데이터 정확 |
| 그림 박스 삽입 | `editor.insert_figure_box()` | 빈 박스 + 캡션 렌더링 |
| 문단 삭제 | `editor.remove_paragraph()` | 대상 문단만 제거됨 |

추가로, 양식 고유 특성에 따라 테스트 항목을 확장한다:
- 셀병합이 있는 표 → 병합 유지 확인
- 다단 레이아웃 → 단 구분 유지 확인
- 머리말/꼬리말 → global_replacements 반영 확인
- 페이지 번호 양식 → 연속 번호 유지 확인

### 테스트 문서 설계 원칙

1. **모든 모듈을 1개 문서에서 테스트** — cover부터 abstract까지 전부 포함
2. **각 연산을 최소 1회 이상 실행** — 누락 없이 전체 커버
3. **앵커를 여러 절에 분산** — 특정 위치에 몰리지 않게
4. **XML 검증 + PDF 시각 검증 모두 수행** — XML에 있어도 렌더링 안 되면 실패
5. **편집 테스트는 빌드 산출물에 대해 실행** — 빌드 테스트 통과 후 진행

**중요**: 분석 완료 후 반드시 글로벌 등록까지 수행할 것. 플러그인 캐시에만 두면 업데이트 시 유실된다.

## 빌드 후 에셋 개선 규칙

등록된 서식으로 문서를 빌드할 때, 아래 상황이 발생하면 **빌드 완료 후 반드시 해당 서식 에셋을 개선**한다. 사용자에게 묻지 않고 능동적으로 수행한다.

### 트리거 조건

| 상황 | 예시 |
|------|------|
| usage-guide.md에 없는 기능이 필요했음 | col_styles 파라미터가 필요한데 가이드에 설명 없음 |
| modules의 process 함수를 우회하여 해결함 | body.py가 지원하지 않는 표 스타일을 HwpxEditor로 후처리 |
| metadata.json에 없는 스타일을 사용함 | "표점본문" 스타일이 필요한데 metadata에 미등록 |
| content dict 구조가 가이드와 불일치 | 가이드에는 없는 필드를 임의로 추가하여 동작시킴 |

### 개선 대상 (3개 파일)

```
1. modules/*.py의 process 함수
   → 우회한 기능을 정식 파라미터/로직으로 통합
   → 새 파라미터 추가 시 기존 호출과 하위 호환 유지

2. usage-guide.md
   → 새 기능/파라미터의 사용법, 예시 코드, 선택 가이드 추가
   → 기존 Iron Laws / Red Flags에 해당 사항 반영

3. metadata.json (필요시)
   → 새 스타일 항목 추가 (pPr/sRef/cPr/name)
   → 새 설정값 추가
```

### 수행 절차

1. 빌드 중 우회/누락 발생 → 메모 (어떤 기능이 부족했는지)
2. 빌드 완료 및 사용자 요청 충족 확인
3. 에셋 3개 파일 개선 (modules → metadata → usage-guide 순서)
4. 개선 사항을 사용자에게 간략히 보고

**이 규칙은 에이전트가 능동적으로 수행한다. 사용자에게 "개선할까요?" 묻지 않는다.**

## 아키텍처

```
src/hwpx_engine/
  ├── build.py              ← 빌드 오케스트레이터
  ├── xml_primitives.py     ← 순수 XML 빌딩 블록 (make_para, make_table_xml, ...)
  ├── charpr_manager.py     ← charPr/borderFill 복제/재활용
  ├── editor.py             ← HwpxEditor (str_replace, insert_paragraph)
  ├── formatter.py          ← StyleMapper (metadata.json → 스타일 해석)
  ├── validator.py          ← 3단계 검증 + 자동수정
  ├── utils.py              ← 네임스페이스 수정 등 유틸
  └── elements.py           ← XML 요소 역분석 API (v1.2.0에서 구현 완료)

~/.claude/hwpx-engine/registered/{template_id}/   (글로벌 영구 저장소)
  ├── template.hwpx         ← 원본 양식
  ├── metadata.json         ← 스타일 매핑
  ├── usage-guide.md        ← content dict 사용법
  ├── builder.py            ← 섹션 파이프라인 선언
  └── modules/              ← 섹션별 처리 모듈
      ├── cover.py, toc.py, body.py, references.py, ...
```
