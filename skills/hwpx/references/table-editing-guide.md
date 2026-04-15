# 기존 표 편집 가이드

## 1. 표 탐색과 구조 확인

기존 표를 편집하기 전에 반드시 대상 표를 찾고 구조를 파악해야 한다.

### 표 찾기 (find_table)

하드코딩된 인덱스(`set_cell(42, ...)`)는 표 순서 변경 시 깨진다. 반드시 `find_table()`로 동적 탐색한다.

```python
from hwpx_engine.editor import HwpxEditor

editor = HwpxEditor.open('document.hwpx')

# 헤더 행(R0)에 특정 텍스트가 있는 표 찾기
indices = editor.find_table('사업비 산정')
print(f'Found {len(indices)} tables: {indices}')

# 여러 표 중 하나 선택
idx = indices[0]
```

### 표 데이터 확인 (get_table_data)

```python
data = editor.get_table_data(idx)
for r, row in enumerate(data):
    print(f'R{r}: {" | ".join(row)}')
```

### 개별 셀 확인 (get_cell)

```python
value = editor.get_cell(idx, 1, 2)  # row 1, col 2
print(f'Cell (1,2): {value}')
```

Key concepts:
- `cellAddr`: (rowAddr, colAddr) — 셀의 논리적 위치
- `cellSpan`: (rowSpan, colSpan) — merge 범위 (anchor 셀에만 표시)
- deactivated cell: merge에 의해 가려진 셀 — `get_cell()`은 빈 문자열 반환
- `set_cell(table_index, row, col, text)`은 cellAddr로 셀을 찾음

## 2. 편집 패턴

### 패턴 A: 단일 셀 텍스트 교체

```python
editor.set_cell(idx, row, col, '새 텍스트')
```

- 첫 번째 hp:p의 첫 번째 hp:run만 남기고 나머지를 제거
- 원래 paragraph의 스타일(paraPrIDRef, styleIDRef, charPrIDRef)은 보존됨
- 수식 셀인 경우 formulaScript의 LastResult도 자동 동기화됨

### 패턴 B: 여러 셀 일괄 수정 (batch_set_cell)

한 표에서 여러 셀을 바꿀 때 set_cell을 반복하는 대신 batch_set_cell을 사용한다. DOM 순회가 1회로 줄어 효율적이다.

```python
idx = editor.find_table('사업비 산정')[0]
editor.batch_set_cell(idx, [
    (1, 2, '300'),
    (2, 2, '500'),
    (3, 2, '100'),
    (4, 3, '1,500'),
])
```

- set_cell과 동일한 동작 (스타일 보존, LastResult 자동 동기화)
- 한 번의 DOM 탐색으로 모든 셀을 처리

### 패턴 C: 표 내부 텍스트 교체 (replace_in_table)

특정 표 안에서만 텍스트를 교체할 때 사용한다. 문서 전체를 대상으로 하는 str_replace/batch_replace와 달리 범위가 제한되어 안전하다.

```python
idx = editor.find_table('부서 현황')[0]
editor.replace_in_table(idx, '영업1팀', '영업2팀')
```

- 기본적으로 고유성 검증: 같은 표 안에서 한 번만 나타나야 함
- `replace_all=True`로 모든 매칭 교체 가능

### 패턴 D: 행 삭제 + 내용 교체 (3-pass)

대규모 표 편집 시 권장하는 구조:

```
PASS 1 (ZIP): batch_replace — 서술 텍스트, 캡션 등 문자열 치환
PASS 2 (DOM): delete_row — 구조 변경 (행 삭제)
PASS 3 (DOM): batch_set_cell — 삭제 후 새 주소 기준으로 셀 내용 설정
```

**핵심 원칙:**
- batch_replace와 DOM 연산(set_cell, delete_row)은 별도 open/save 세션으로 분리
- delete_row는 항상 뒤에서 앞으로 (큰 인덱스부터)
- delete_row 후 `_reindex_addresses`가 자동 호출되어 모든 rowAddr 갱신
- PASS 3의 set_cell/batch_set_cell은 삭제 후 새 행 주소를 사용해야 함

```python
# PASS 2
editor = HwpxEditor.open('doc.hwpx')
idx = editor.find_table('사업비 산정')[0]
for row in [12, 8, 5]:  # 뒤에서 앞으로
    editor.delete_row(idx, row)
editor.save('doc.hwpx')

# PASS 3
editor = HwpxEditor.open('doc.hwpx')
idx = editor.find_table('사업비 산정')[0]
editor.batch_set_cell(idx, [
    (1, 2, '300'), (2, 2, '500'),
])
editor.save('doc.hwpx')
```

### 패턴 E: merge 구조 조정 (전략적 행 삭제)

직급별 merge 그룹에서 특정 행을 삭제하면 merge span이 자동 축소된다.

**예시**: 5급(merge=3x1)을 5급(merge=2x1)로 줄이려면:
- 5급 merge 그룹 내에서 마지막 행을 delete_row
- 엔진이 자동으로 rowSpan 3→2로 감소, rowAddr 연속 재번호, height 비례 조정

```python
# 원본: 5급(R2, merge=3x1), 6급(R5, merge=4x1), 7급(R9, merge=4x1)
# 목표: 5급(2x1), 6급(2x1), 7급(3x1)
# 삭제: R12(7급 마지막), R8(6급 마지막), R7(6급 3rd), R4(5급 마지막)
idx = editor.find_table('소요인력')[0]
for r in [12, 8, 7, 4]:  # 뒤에서 앞으로
    editor.delete_row(idx, r)
```

**주의사항:**
- merge anchor 행을 삭제하면 anchor가 다음 행으로 자동 이전됨
- 삭제 순서가 중요: 항상 큰 인덱스부터 삭제해야 앞쪽 인덱스가 안정적
- 삭제 후 실제 데이터를 get_table_data()로 확인하는 것이 안전

### 패턴 F: 행 추가 (add_row)

```python
idx = editor.find_table('사업비 산정')[0]
editor.add_row(idx, ['국내복귀기업 지원', '500', '-', '-', '-'], position=8)
```

- merge-aware: 삽입 위치를 가로지르는 rowSpan 셀이 있으면 자동으로 rowSpan 증가
- rowAddr 자동 재번호
- rowCnt 자동 업데이트

### 패턴 G: 캡션 스타일 변경

표/그림 캡션의 스타일(paraPrIDRef/styleIDRef)을 변경할 때 사용한다. 엔진은 캡션을 자동 판정하지 않으므로, 먼저 주변 paragraph를 확인하고 호출자가 판단한다.

```python
# 1. 주변 paragraph 확인
idx = editor.find_table('사업비 산정')[0]
paras = editor.get_nearby_paragraphs(idx)
for p in paras:
    print(f"offset={p['offset']}: {p['text'][:50]}  "
          f"style={p['style_ref']}  auto_num={p['has_auto_num']}")

# 2. 캡션 식별 후 스타일 변경
# offset=-1이 캡션이라고 확인된 경우:
editor.set_paragraph_style(idx, offset=-1, para_pr='56', style_ref='28')
```

- `get_nearby_paragraphs(idx, before=3, after=2)`: 표 전후 paragraph의 텍스트, 스타일, autoNum 여부를 반환
- `set_paragraph_style(idx, offset, ...)`: offset 위치의 paragraph 스타일을 변경
- `char_pr` 파라미터로 모든 run의 charPrIDRef도 일괄 변경 가능

## 3. 메서드 선택 가이드

| 대상 | 방법 | 이유 |
|------|------|------|
| 서술 텍스트 ("120명"→"53명") | batch_replace | 문서 전체에서 고유 문자열 치환 |
| 표 캡션/제목 | batch_replace | 캡션은 표 외부에 위치 |
| 표 안의 개별 셀 | set_cell | 행/열 주소로 정확히 접근 |
| 한 표의 여러 셀 | batch_set_cell | 단일 DOM 순회로 효율적 |
| 표 안에서만 특정 텍스트 교체 | replace_in_table | 범위 제한으로 안전 |
| 산출기초 (x 120명 → x 53명) | batch_replace | 고유 문자열이면 batch_replace가 효율적 |
| 합계 행 숫자 | set_cell | 특정 셀 하나만 변경 |
| 캡션 스타일 변경 | get_nearby_paragraphs → set_paragraph_style | 범용적 |

**batch_replace 한계:**
- 고유하지 않은 숫자(예: "100", "300")는 batch_replace로 치환하면 의도치 않은 곳도 변경
- 이런 경우 `replace_in_table`이나 `set_cell`로 개별 접근

## 4. 구조 삭제 패턴

### delete_table

`delete_table(table_index, clean_surrounding=True)`는 표와 함께 주변 관련 문단을 자동 제거한다.

**자동 제거 대상 (clean_surrounding=True):**
- 표 앞: 캡션, 단위 행
- 표 뒤: 출처, 주석, 빈 줄/대시

```python
# 표와 주변 텍스트 한꺼번에 삭제 (뒤에서부터, find_table로 찾기)
for pattern in ['삭제 대상 A', '삭제 대상 B']:
    indices = editor.find_table(pattern)
    for idx in reversed(indices):
        editor.delete_table(idx)

# 표만 삭제, 주변 유지
editor.delete_table(idx, clean_surrounding=False)
```

**주의사항:**
- 표 삭제 후 인덱스가 변경되므로 항상 뒤에서부터 삭제
- 본문 서술은 자동 제거 대상이 아님 — `remove_paragraph`로 별도 제거

### remove_paragraph

`remove_paragraph(containing, remove_all=False)`는 고유성 검증을 수행한다.

```python
# 고유 텍스트로 1건만 제거 (기본)
editor.remove_paragraph('기업 간 공급망 협력 사업')

# 모든 매칭 제거
editor.remove_paragraph('(삭제됨)', remove_all=True)
```

**주의사항:**
- 짧은 텍스트는 여러 문단에 매칭될 수 있으므로, 충분히 고유한 문자열 사용
- XML entity는 lxml이 자동 해석하므로, `R&D`로 검색 (NOT `R&amp;D`)

## 5. 읽기 및 검증 패턴

### 수정 후 교차 검증

```python
editor.set_cell(idx, 1, 1, '999')
assert editor.get_cell(idx, 1, 1) == '999'
```

### 전체 표 추출 후 비교

```python
data = editor.get_table_data(idx)
for row in data:
    print(' | '.join(row))
```

### find_table → get_table_data 연계

```python
# 여러 표 데이터 한꺼번에 확인
for pattern in ['사업비', '인건비', '경상비']:
    indices = editor.find_table(pattern)
    for idx in indices:
        data = editor.get_table_data(idx)
        print(f'Table {idx} ({pattern}): {len(data)}R x {len(data[0])}C')
```
