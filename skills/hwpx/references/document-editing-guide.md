# 대규모 문서 편집 가이드

기존 HWPX 문서를 대규모로 수정할 때의 범용 워크플로우와 패턴.

## 3-pass 아키텍처

HWPX 편집은 반드시 아래 순서로 진행한다. ZIP 레벨과 DOM 레벨 조작을 같은 세션에 섞으면 데이터 변조가 발생한다.

```
PASS 1 (ZIP): batch_replace — 텍스트 일괄 교체
PASS 2 (DOM): delete_row/delete_table/remove_paragraph — 구조 삭제
PASS 3 (DOM): set_cell/batch_set_cell/add_row/set_paragraph_style — 셀·행·스타일 수정
각 PASS 사이에 editor.save() → del editor → HwpxEditor.open() 필요
```

### 왜 순서가 중요한가

- batch_replace (ZIP)는 원본 바이트를 직접 치환. DOM 변경사항이 있으면 `_flush_dom()`이 먼저 실행되어 ZIP과 DOM이 충돌.
- delete_row/set_cell은 둘 다 DOM 레벨이지만, delete_row가 행 인덱스를 변경하므로 set_cell보다 먼저 실행해야 인덱스가 맞음.

### 메서드 선택 플로우차트

```
텍스트를 바꿀 곳이 어디인가?
├─ 문서 전체 (서술 + 표 + 캡션) → batch_replace (PASS 1)
├─ 특정 표 안에서만 → replace_in_table (PASS 3)
├─ 특정 셀 하나 → set_cell (PASS 3)
├─ 특정 표의 여러 셀 → batch_set_cell (PASS 3)
└─ 캡션/주변 paragraph 스타일 → set_paragraph_style (PASS 3)

구조를 바꿀 곳은?
├─ 행 삭제 → delete_row (PASS 2, 뒤→앞)
├─ 행 추가 → add_row (PASS 3)
├─ 표 전체 삭제 → delete_table (PASS 2, 뒤→앞)
└─ 문단 삭제 → remove_paragraph (PASS 2)

표를 어떻게 찾는가?
├─ find_table('패턴') → 동적 인덱스 (항상 이것을 사용)
└─ 고정 인덱스 (42 등) → 절대 사용하지 말 것!
```

## 전체 예제: 대규모 편집

```python
from hwpx_engine.editor import HwpxEditor

# === PASS 1 (ZIP): 텍스트 일괄 교체 ===
editor = HwpxEditor.open('report.hwpx')
editor.batch_replace([
    ('120명', '53명'),
    ('22개 사업', '17개 사업'),
    ('10팀+1센터', '6팀'),
    ('A팀', 'B팀'),
])
editor.save('report.hwpx')
del editor

# === PASS 2 (DOM): 구조 삭제 ===
editor = HwpxEditor.open('report.hwpx')

# 표 삭제 (find_table으로 동적 탐색, 뒤에서부터)
for pattern in ['ESG 경영 지원', '직원 교육 프로그램']:
    indices = editor.find_table(pattern)
    for idx in reversed(indices):
        editor.delete_table(idx)

# 행 삭제 (find_table로 표 찾기, 뒤에서 앞으로)
idx = editor.find_table('사업비 산정')[0]
for row in [12, 8, 5]:
    editor.delete_row(idx, row)

# 문단 삭제
editor.remove_paragraph('R&D 혁신 허브 사업은')

editor.save('report.hwpx')
del editor

# === PASS 3 (DOM): 셀·행·스타일 수정 ===
editor = HwpxEditor.open('report.hwpx')

# 여러 셀 일괄 수정
idx = editor.find_table('사업비 산정')[0]
editor.batch_set_cell(idx, [
    (1, 2, '300'),
    (2, 2, '500'),
    (3, 2, '100'),
    (5, 3, '1,500'),
])

# 행 추가
editor.add_row(idx, ['국내복귀기업 지원', '500', '-', '-', '-'], position=8)

# 캡션 스타일 변경
paras = editor.get_nearby_paragraphs(idx)
# offset=-1이 캡션인지 확인 후:
editor.set_paragraph_style(idx, offset=-1, para_pr='56', style_ref='28')

# 교차 검증
data = editor.get_table_data(idx)
print(f'Row count: {len(data)}')
assert editor.get_cell(idx, 1, 2) == '300'

editor.save('report_final.hwpx')
```

## 패턴 1: 반복 섹션 일괄 삭제

문서에서 특정 항목(사업, 제품, 부서 등)을 삭제할 때:

1. `find_text()`로 삭제 대상 섹션의 시작/끝 경계 확인
2. `find_table()` + `delete_table()`로 관련 표 삭제 (clean_surrounding=True로 캡션/출처 동시 제거)
3. `remove_paragraph()`로 관련 서술 텍스트 삭제
4. `find_text()`로 고아 텍스트(삭제 대상의 잔존 참조) 탐지

## 패턴 2: 수치 연쇄 수정

한 값이 변경되면 관련된 모든 표와 서술을 함께 수정:

1. `find_text()`로 변경 대상 수치가 나오는 모든 위치 파악
2. `batch_replace()`로 서술 텍스트의 수치 일괄 변경 (PASS 1)
3. `find_table()` + `batch_set_cell()`로 표 셀 값 변경 (PASS 3)
4. `get_cell()`/`get_table_data()`로 교차 검증 — 표 값과 서술 텍스트의 일치 확인

## 패턴 3: 넘버링 재정렬

항목 삭제 후 남은 항목의 번호 수정:

1. 삭제된 항목 이후의 번호 목록 작성 (예: (6)→(5), (7)→(6))
2. `batch_replace()`로 일괄 변경
3. `find_text()`로 변경 결과 확인

## 패턴 4: 고아 텍스트 탐지/제거

삭제된 항목을 참조하는 잔존 텍스트 제거:

1. 삭제된 항목의 고유 키워드 목록 작성
2. `find_text()`로 각 키워드 검색
3. 발견된 참조를 `remove_paragraph()` 또는 `batch_replace()`로 제거

## 패턴 5: 교차 검증

표 데이터와 서술 텍스트의 일치 확인:

1. `find_table()` + `get_table_data()`로 표 값 추출
2. `find_text()`로 서술 텍스트의 수치 추출
3. 두 값 비교 — 불일치 시 수정

## 패턴 6: 캡션 스타일 일괄 변경

여러 표의 캡션 스타일을 한꺼번에 변경:

```python
# 3장의 모든 표 캡션 스타일을 변경
for pattern in ['사업별 예산', '조직 현황', '인력 계획']:
    indices = editor.find_table(pattern)
    for idx in indices:
        paras = editor.get_nearby_paragraphs(idx)
        # 캡션이 있는지 확인
        caption = [p for p in paras if p['offset'] < 0 and p['has_auto_num']]
        if caption:
            editor.set_paragraph_style(idx, offset=caption[0]['offset'],
                                       para_pr='56', style_ref='28')
```
