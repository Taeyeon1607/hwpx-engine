# hwpx-engine

HWPX 문서(.hwpx) 생성, 편집, 검증을 위한 Claude Code 플러그인.

한글과컴퓨터(한컴)의 HWPX 포맷(OWPML 기반, ZIP+XML 구조)을 다룹니다.

## 설치

```bash
claude plugin add Taeyeon1607/hwpx-engine
```

## 요구사항

- Python 3.10+
- `lxml >= 4.9`
- `Pillow >= 9.0` (이미지 삽입 시)

```bash
pip install lxml Pillow
```

## 주요 기능

### 1. 양식 기반 문서 생성 (build)

등록된 양식 템플릿에 데이터를 넣어 새 문서를 자동 생성합니다.

```python
from hwpx_engine.build import build

content = {
    'cover_replacements': {'원본 제목': '새 제목'},
    'chapters': [
        {
            'title': '서론',
            'items': [
                {'style': 'heading_1', 'text': '1. 연구의 배경'},
                {'style': 'bullet_1', 'text': '본문 내용'},
                {'type': 'table', 'caption': '연구 개요',
                 'headers': ['구분', '내용'],
                 'rows': [['기간', '2026.1~12']],
                 'source': '자료 : 연구진.'},
                {'type': 'image', 'caption': '현황 분석',
                 'path': '/path/to/image.jpg',
                 'source': '자료 : 통계청.'},
            ]
        }
    ],
    'references': ['참고문헌 1'],
}

result = build('template_id', content, 'output/document.hwpx')
```

### 2. 기존 문서 편집 (HwpxEditor)

```python
from hwpx_engine.editor import HwpxEditor

editor = HwpxEditor.open('document.hwpx')
editor.str_replace('기존 텍스트', '새 텍스트')
editor.insert_paragraph(
    parts=[{'text': '일반 '}, {'text': '볼드', 'bold': True}],
    style='bullet_2', after='기존 텍스트'
)
editor.save('output.hwpx')
```

### 3. 표 조작

```python
# 표 찾기 / 셀 읽기 / 셀 쓰기
indices = editor.find_table("구분")
data = editor.get_table_data(0)
value = editor.get_cell(0, 1, 2)

# 일괄 셀 교체
editor.batch_set_cell(0, {(1,0): '새값1', (2,1): '새값2'})
```

### 4. 주변 문단 / 스타일 설정

```python
# 표 주변 문단 찾기
nearby = editor.get_nearby_paragraphs(0, before=2, after=2)

# 문단 스타일 변경
editor.set_paragraph_style(section=0, para_index=5,
                           pPr='46', sRef='6', cPr='17')
```

## 양식 에셋 등록

새 양식을 등록하려면 `~/.claude/hwpx-engine/registered/{id}/`에 다음 파일을 준비합니다:

| 파일 | 설명 |
|------|------|
| `template.hwpx` | 원본 양식 파일 |
| `metadata.json` | 스타일 매핑 (pPr, sRef, cPr, borderFill 등) |
| `usage-guide.md` | content dict 구성법, 주의사항 |
| `builder.py` | 섹션 파이프라인 + prepare() 훅 |
| `modules/` | 섹션별 처리 모듈 (cover.py, body.py 등) |

양식 분석 방법은 `skills/hwpx/references/template-analysis-guide.md`를 참조하세요.

### 에셋 배포

다른 사용자에게 에셋을 배포하려면:
1. 에셋 폴더를 zip으로 압축
2. 수신자가 `~/.claude/hwpx-engine/registered/` 아래에 압축 해제
3. `build('template_id', content, output_path)`로 사용

## 아키텍처

```
src/hwpx_engine/
  ├── build.py              ← 빌드 오케스트레이터 (prepare hook, 섹션 핸들러)
  ├── xml_primitives.py     ← XML 빌딩 블록 (make_para, make_table_xml 등)
  ├── charpr_manager.py     ← charPr/borderFill 동적 관리
  ├── editor.py             ← HwpxEditor (편집 API)
  ├── tables.py             ← 표 조작 (batch_set_cell, formulaScript 등)
  ├── elements.py           ← 각주, 머리말/꼬리말
  ├── formatter.py          ← StyleMapper
  ├── validator.py          ← 3단계 검증
  └── utils.py              ← 유틸리티

~/.claude/hwpx-engine/registered/{template_id}/
  ├── template.hwpx         ← 원본 양식
  ├── metadata.json          ← 스타일 매핑
  ├── usage-guide.md         ← 사용 가이드
  ├── builder.py             ← 섹션 파이프라인
  └── modules/               ← 섹션별 처리 모듈
```

## 라이선스

MIT
