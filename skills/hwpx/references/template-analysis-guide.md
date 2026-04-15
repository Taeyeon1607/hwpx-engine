# 양식 분석 완전 가이드

> 새 양식을 등록할 때 이 가이드를 따라 분석한다.
> 이 가이드대로 분석하면 해당 양식의 서식을 완벽히 재현하여 새 문서를 생성할 수 있다.

## 왜 이 가이드가 필요한가

HWPX 양식은 겉보기에 단순해 보여도 내부적으로 복잡한 구조를 갖고 있다:
- 같은 □ 기호라도 직접 타이핑한 텍스트와 `styleIDRef`로 자동 생성되는 글머리표는 완전히 다르다
- 같은 "제목"이라도 표지 제목, 간지 제목, 장 제목, 절 제목이 모두 다른 스타일 조합을 사용한다
- 빈 줄 하나도 `paraPrIDRef`가 다르면 간격이 달라진다
- **같은 양식이라도 메모 포함/미포함 버전에서 모든 ID가 달라진다**

**코드로 자동 분류하지 마라. 추출된 데이터를 에이전트(AI)가 직접 읽고 이해해야 한다.**

---

## 분석 절차

### Phase 1: 원시 데이터 추출

스크립트로 다음을 추출한다 (해석하지 않고 있는 그대로):

```python
# 직접 스크립트로 추출 (양식의 ZIP 내 XML을 파싱)
# 각 섹션 파일(section0~N)의 모든 문단을 순서대로 출력:
#   인덱스, paraPrIDRef, styleIDRef, heading 속성, charPrIDRef(각 런), 텍스트, 표/이미지 여부
```

추출할 항목:

| 항목 | 위치 | 설명 |
|------|------|------|
| 섹션 파일 목록 | ZIP 내 Contents/section*.xml | 몇 개의 섹션으로 구성되어 있는지 |
| 문단 시퀀스 | 각 section*.xml | 모든 문단을 순서대로 |
| paraPrIDRef | `<hp:p>` 속성 | 문단 모양 참조 ID |
| styleIDRef | `<hp:p>` 속성 | **스타일 참조 ID — 글머리표/개요번호의 핵심** |
| charPrIDRef | `<hp:run>` 속성 | 글자 모양 참조 ID |
| heading 속성 | header.xml의 paraPr 내 | `type="BULLET\|NUMBER"`, `idRef`, `level` |
| pageBreak | `<hp:p>` 속성 | 페이지 나눔 여부 |
| 표 존재 여부 | `<hp:tbl>` | rowCnt, colCnt, borderFillIDRef |
| 이미지 존재 여부 | `<hp:pic>` | binaryItemIDRef |
| secPr 존재 여부 | `<hp:secPr>` | 섹션 속성 (첫 문단에 포함) |
| charPr 정의 | header.xml | 글꼴, 크기, 볼드, 색상, 장평, 자간 |
| paraPr 정의 | header.xml | 정렬, 행간, 들여쓰기, 여백, 문단 간격 |
| heading 정의 | header.xml paraPr 내 | BULLET/NUMBER 타입, idRef, level |

### Phase 2: 에이전트가 읽고 이해

추출 결과를 에이전트가 직접 읽으면서 다음을 판단한다:

#### 2-1. 섹션별 역할 파악

각 section*.xml이 문서의 어떤 부분인지 판단한다. **양식마다 구조가 완전히 다르므로 미리 가정하지 말 것.**

가능한 구조 유형들 (예시일 뿐, 이 외에도 다양함):

```
# 유형 A: 장별로 별도 섹션 파일
section0.xml → 표지
section1.xml → 목차
section2.xml → 제1장 (간지 + 본문)
section3.xml → 제2장 (간지 + 본문)
section4.xml → 참고문헌

# 유형 B: 본문이 하나의 섹션
section0.xml → 표지
section1.xml → 본문 전체 (장 구분은 내부 요소로)

# 유형 C: 더 복잡한 구조
section0.xml → 표지 + 판권
section1.xml → 요약
section2.xml → 목차
section3~N.xml → 본문 장들
sectionN+1.xml → 부록
```

**판단 근거**: 첫 몇 개 문단의 텍스트 내용, pageBreak 위치, 이미지/표 유무

#### 2-2. 구역(Zone) 경계 식별

각 섹션 내에서 구역을 나눈다. 예시:

```
section0.xml:
  p[0]~p[?]   : 표지     (이미지, 제목, 저자)
  p[?]~p[?]   : 판권지   (발행정보 표, 연구진 표)
  p[?]~p[?]   : 요약     (불릿 본문)
  p[?]~p[?]   : 키워드   (키워드 표)
```

**판단 근거**: pageBreak 위치, 스타일 변화 패턴, 내용

#### 2-3. 문단 역할 분류

각 문단이 다음 중 어느 역할인지 판단:

| 역할 | 설명 | 처리 방법 |
|------|------|----------|
| **구조** | 이미지, secPr, 장식용 표, 페이지나눔, 간지 요소 | **유지** (건드리지 않음) |
| **교체 가능 텍스트** | 제목, 저자, 날짜, 기관명 등 | **텍스트만 교체** (런 구조 유지) |
| **내용** | 불릿 본문, 절 제목, 참고문헌 항목 등 | **제거 후 새로 삽입** |
| **메타** | "스타일 목록 →", "KoPub돋움체 Bold 9pt..." | **제거** (양식 안내용 텍스트) |
| **빈줄** | 텍스트 없는 문단 | **구역에 따라 유지 또는 제거** |

**중요**: 코드로 자동 분류하면 틀린다. 에이전트가 문맥을 보고 판단해야 한다.

#### 2-4. 스타일 매핑 테이블 작성

**이것이 가장 중요한 산출물이다.** 양식에서 실제 사용된 스타일 조합을 의미론적 이름으로 매핑한다. 모든 값은 원본에서 직접 추출해야 하며 추정하지 않는다.

매핑해야 하는 항목들 (양식에 따라 존재 여부가 다름):

```
[본문 계층 — 글머리표/개요번호]
  1수준:  paraPrIDRef=?  styleIDRef=?  charPrIDRef=?  (heading 속성 확인)
  2수준:  paraPrIDRef=?  styleIDRef=?  charPrIDRef=?
  3수준:  paraPrIDRef=?  styleIDRef=?  charPrIDRef=?

[절/소절 제목]
  소절:   paraPrIDRef=?  styleIDRef=?  charPrIDRef=?  (런 구조 확인: 1런? 2런?)
  소소절: paraPrIDRef=?  styleIDRef=?  charPrIDRef=?

[장 구조 — 양식에 따라 다름]
  장 제목바:    paraPrIDRef=?  (표로 구현? 단독 문단?)
  장 제목:      paraPrIDRef=?
  간지 제목:    paraPrIDRef=?  (있는 경우. 위치: 일반 문단? RECT drawText?)
  간지 절 목차: paraPrIDRef=?  (있는 경우)

[목차 — 있는 경우]
  TOC 장 제목: paraPrIDRef=?  (탭+리더 있는지? 텍스트에 _ 포함?)
  TOC 절 제목: paraPrIDRef=?
  TOC 표/그림: paraPrIDRef=?

[표]
  표 캡션:     paraPrIDRef=?  (자동번호 styleIDRef 확인: NUMBER? 없음?)
  표 머리 셀:  paraPrIDRef=?  charPrIDRef=?  borderFillIDRef=?
  표 본문 셀:  paraPrIDRef=?  charPrIDRef=?  borderFillIDRef=?
  표 단위:     paraPrIDRef=?  (우측정렬, 있는 경우)
  출처:        paraPrIDRef=?

[참고문헌/Abstract — 있는 경우]
  참고문헌:    paraPrIDRef=?  styleIDRef=?  charPrIDRef=?
  Abstract:    paraPrIDRef=?  styleIDRef=?  charPrIDRef=?
  키워드:      paraPrIDRef=?  (위치: 테이블 셀? 일반 문단?)
```

**charPrIDRef를 절대 빠뜨리지 마라.** 이 값이 틀리면 글꼴, 크기, 볼드가 전부 달라진다. 추정하지 말고 원본에서 직접 추출한다.

추출된 매핑은 `metadata.json`에 저장하여 빌드 엔진이 참조하도록 한다.

#### 2-5. 표/그림 구조 정밀 분석

**이 단계를 건너뛰면 표/그림이 원본과 달라진다.** 양식마다 표/그림 구조가 완전히 다르다.

##### 표 캡션 구조

많은 양식에서 표 캡션은 단순 텍스트가 아니라 **별도 1행 TBL**로 구현되어 있다:

```
캡션 TBL (1행 4열):
  C0: 캡션 텍스트 (bf=77, 회색 배경)
  C1: 대각선 삼각형 (bf=50)
  C2: 단위 텍스트 (bf=28)
  C3: 빈 셀 (bf=28)
```

분석 시 확인할 것:
- **캡션이 별도 TBL인지 단순 paragraph인지**
- **캡션 TBL의 열 너비**: ratio가 아닌 원본의 정확한 픽셀 값 추출
- **캡션 셀의 cellMargin**: hasMargin, left/right/top/bottom 값
- **캡션 sRef**: 자동번호를 사용하는 sRef 확인 (캡션 TBL의 C0만 적용, C1~C3은 sRef=0)

##### 표 데이터 셀의 위치별 borderFill

**모든 셀이 같은 borderFill을 사용하지 않는다.** 좌측 첫 열, 가운데 열, 우측 마지막 열이 다른 borderFill을 사용하여 테두리가 겹치지 않게 한다:

```
헤더행: [bf_left=64, bf_center=73, bf_right=74]
첫번째 데이터행: [bf_left=46, bf_center=68, bf_right=47]
나머지 데이터행: [bf_left=48, bf_center=69, bf_right=49]
```

분석 방법: 원본 표의 첫 데이터 TBL에서 모든 tr의 tc borderFillIDRef를 추출하여 패턴을 파악한다.

##### 그림 구조 (복합 TBL)

빈 그림 박스가 단순 RECT가 아닌 **복합 TBL**인 양식이 있다. 반드시 원본의 그림 TBL 구조를 정확히 분석해야 한다:

```
예: 7열 4행 TBL
  R0: 캡션 (C0 colspan=4 + C4 + C5 + C6)  ← 각 셀 다른 bf
  R1: 상단 border (C0 + C1 colspan=5 + C6) ← 얇은 행, cPr=46(5pt)
  R2: 이미지 영역 (C0 + C1 colspan=5 + C6) ← 큰 행
  R3: 하단 border (C0 + C1 colspan=5 + C6) ← 얇은 행, cPr=46(5pt)
```

분석 시 **반드시 확인할 것**:
1. **정확한 열 수 (colCnt)**: 보이는 셀 수와 다를 수 있음 (colspan 때문)
2. **각 셀의 borderFillIDRef**: 모든 셀이 다른 bf를 사용
3. **각 셀의 cellSz width/height**: 원본 값 그대로 사용
4. **border 행의 charPrIDRef**: 글자크기가 작아야 행 높이가 원본과 동일 (예: 5pt)
5. **cellMargin과 hasMargin**: border 행은 hasMargin=0 (기본 여백 141)
6. **subList 자식 순서**: HWPX에서 tc 자식 순서가 중요함 (subList → cellAddr → cellSpan → cellSz → cellMargin)

##### 이미지 삽입 위치

이미지(PIC)를 그림 TBL 셀에 넣을 때:
- `make_image_pic()`은 `<hp:p><hp:run><hp:tbl>...<hp:pic>...</hp:tbl></hp:run></hp:p>` 구조를 생성
- 그림 TBL의 셀에 넣을 때는 **내부 PIC paragraph만 추출**하여 삽입해야 중첩 TBL 문제를 방지

```python
tmp_p = make_image_pic(bin_id, org_w, org_h, cur_w, cur_h, ...)
inner_p = tmp_p.find(f'.//{{{HP}}}tc/{{{HP}}}subList/{{{HP}}}p')
sl.append(inner_p)  # 셀의 subList에 직접 추가
```

##### wrapper paragraph 스타일

표/그림 TBL을 감싸는 paragraph에도 스타일이 있다. 원본에서 확인:

```
표 캡션 wrapper: paraPrIDRef=24 ("표그림제목 위치")
데이터 표 wrapper: paraPrIDRef=2 ("표그림 위치")
그림 TBL wrapper: paraPrIDRef=2 ("표그림 위치")
```

이 스타일이 틀리면 표/그림 위치가 어긋난다.

##### 머리말(masterpage) 분석

머리말/꼬리말에 연구 제목이 포함된 양식이 있다. `Contents/masterpage*.xml` 파일을 확인하고, cover_replacements로 교체할 텍스트가 있는지 분석한다.

### Phase 3: 처리 규칙 결정

각 구역에 대해 처리 전략을 결정한다:

```
1. 표지: 어떤 텍스트가 교체 가능한지 (RECT vs 일반 문단)
2. TOC: 위치와 구조 (별도 섹션? 표지 내 테이블? RECT 안 테이블?)
3. 장 경계: 어떻게 구분하는지 (별도 섹션 파일? 1x1 테이블? pageBreak?)
4. 간지: 있는지, 있으면 RECT drawText인지 일반 문단인지
5. 머리말: 위치와 교체 가능 여부
6. 참고문헌/부록: 위치와 삭제 범위
```

### Phase 4: 테스트

```
1. 빈 content로 build → 원본과 동일하게 열리는지 확인
2. 기능 하나씩 추가하면서 확인 (요약 → 1장 → 전장 → 참고문헌)
3. 손상 발생 시: 단계별 비활성화로 원인 분리
4. 표 스타일 확인: 헤더 색상, 본문 스타일이 원본과 일치하는지
```

### Phase 5: 등록

```
assets/registered/{id}/
  template.hwpx    ← 양식 파일 복사
  metadata.json    ← 스타일 매핑 + 메타데이터
  usage-guide.md   ← content dict 구성법, 특이사항
  builder.py       ← 섹션 파이프라인 선언
  modules/         ← 섹션별 처리 모듈
    __init__.py
    cover.py       ← 표지 처리
    toc.py         ← 목차 처리 (있는 경우)
    body.py        ← 본문 처리
    references.py  ← 참고문헌 등 (있는 경우)
    interleaf.py   ← 간지 처리 (있는 경우)
    ...            ← 양식 구조에 따라 추가
```

---

## HWPX 핵심 개념

### styleIDRef가 글머리표를 결정한다

#### ❌ 잘못된 방법 (텍스트로 마커 입력)
```xml
<hp:p paraPrIDRef="0" styleIDRef="0">
  <hp:run charPrIDRef="0"><hp:t>□ 내용</hp:t></hp:run>
</hp:p>
```
→ □가 텍스트 문자로 표시됨. 한글의 Ctrl+2,3,4 글머리표와 완전히 다름.

#### ✅ 올바른 방법 (styleIDRef로 자동 생성)
```xml
<!-- pPr, sRef, cPr 값은 양식마다 다름 — metadata.json에서 조회 -->
<hp:p paraPrIDRef="?" styleIDRef="?" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="?"><hp:t>내용만 작성</hp:t></hp:run>
</hp:p>
```
→ 글머리표가 paraPr의 heading BULLET 설정에 의해 자동으로 생성됨.

#### heading 속성 확인 방법

header.xml에서 해당 paraPrIDRef의 정의를 확인:

```xml
<hh:paraPr id="?">
  <hh:heading type="BULLET" idRef="?" level="0"/>
  <!-- type="BULLET"이면 글머리표, "NUMBER"이면 개요번호 -->
  <!-- type="NONE"이면 일반 문단 -->
</hh:paraPr>
```

### 런(Run) 구조 이해

#### 단일 런 (대부분의 경우)
```xml
<hp:p paraPrIDRef="?" styleIDRef="?">
  <hp:run charPrIDRef="?"><hp:t>전체 내용</hp:t></hp:run>
</hp:p>
```

#### 복수 런 (서식이 중간에 바뀌는 경우)
```xml
<!-- 번호와 제목이 분리된 2런 구조 -->
<hp:p paraPrIDRef="?" styleIDRef="?">
  <hp:run charPrIDRef="?"><hp:t>(1)</hp:t></hp:run>
  <hp:run charPrIDRef="?"><hp:t> 배경</hp:t></hp:run>
</hp:p>
```

**왜 분리되는가**: 한글에서 개요번호 입력 시 번호 부분과 텍스트 부분이 별도 런으로 생성됨. 새 문단 생성 시에도 이 구조를 재현해야 한글이 동일하게 인식한다.

#### 런 경계를 넘는 텍스트 교체

표지 제목처럼 편집 시 서식이 중간에 바뀌면 하나의 문장이 여러 런으로 쪼개짐. raw XML 문자열 레벨에서 교체하면 런 경계를 넘어도 안전하게 교체 가능.

### 간지(Interleaf) 구조

일부 양식에서 장 앞에 간지 페이지가 있다. 간지의 번호, 제목, 절 목차는 일반 문단이 아니라 **RECT drawText 서브리스트** 안에 있을 수 있다.

#### 간지 구조 분석 방법

```python
# 간지 문단(보통 첫 2-3개 문단)의 RECT drawText 구조 확인
for para in paras[:5]:
    for drawText in para.iter(f'{{{HP}}}drawText'):
        subList = drawText.find(f'{{{HP}}}subList')
        sub_paras = subList.findall(f'{{{HP}}}p')
        first_ppr = sub_paras[0].get('paraPrIDRef', '?')
        print(f'  drawText: {len(sub_paras)} sub-paras, first pPr={first_ppr}')
```

#### 간지 요소별 위치 (예시)

```
Para 0 (secPr 포함):
  RECT drawText → sub-paras with 숫자 ("0", "1" → "01")

Para 1:
  RECT 1 drawText → 절 목차 sub-paras ("1. 배경", "2. 목적")
  RECT 2 drawText → 푸터 sub-para (기관명)
  RECT 3 drawText → 큰 제목 sub-paras ("서론")
```

**간지 숫자 구조 주의**: 양식에 따라 숫자가 2개 sub-para로 분리되거나("0" + "1"), 1개 sub-para에 2개 text node로 들어있을 수 있다. 분석 시 반드시 실제 구조를 확인할 것.

**간지 run-level charPr 보존**: "제N장"처럼 각 글자가 다른 크기/스타일인 경우가 있다(예: 제=20pt, N=32pt, 장=20pt). `_set_para_text()`로 전체를 덮어쓰면 크기 차이가 사라진다. 반드시 숫자 run만 찾아서 텍스트만 변경해야 한다.

**간지 sRef별 역할 판별**: 인덱스가 아닌 sRef로 역할을 판단한다:
- sRef=42: 장 제목 (크게)
- sRef=43: 절 목차 항목 (heading_1과 매핑)
- sRef=46/47: 꼬리말 ("제N장 제목" 형식)
- sRef=5: 본문 큰 제목 (첫 번째만 교체, 나머지는 비움)

#### 간지 텍스트 교체

RECT 내부의 서브 문단은 `<hp:t>` 노드 수정이 안전하다 (최상위 문단이 아니므로).

#### 간지 절 목차 추가/삭제

새 콘텐츠의 절 수가 원본과 다를 때:
- **추가**: 같은 pPr, sRef, cPr로 새 `<hp:p>` 생성하여 subList에 append
- **삭제**: subList에서 초과 `<hp:p>` 직접 remove (서브리스트 내부라 안전)

---

## 산출물 상세

분석 완료 후 다음 파일들을 생성한다.

### 1. metadata.json

양식의 스타일 매핑과 메타데이터를 담는 핵심 파일. 빌드 엔진과 모듈이 이 파일에서 모든 스타일 ID를 조회한다.

**필수 필드:**

```json
{
  "id": "template_id",
  "display_name": "양식 표시 이름",
  "description": "양식 설명",
  "tags": ["태그1", "태그2"],
  "sections": [
    {"file": "section0.xml", "role": "역할 설명"},
    {"file": "section1.xml", "role": "역할 설명"}
  ],
  "styles": {
    "bullet_1": {"pPr": "?", "sRef": "?", "cPr": "?", "name": "설명"},
    "table_caption": {"pPr": "?", "sRef": "?", "cPr": "?"},
    "table_header_cell": {"pPr": "?", "sRef": "?", "cPr": "?", "borderFill": "?"},
    "reference": {"pPr": "?", "sRef": "?", "cPr": "?"}
  },
  "table_page_width": "?",
  "figure_box_width": "?",
  "figure_box_height": "?",
  "figure_box_borderFill": "?",
  "safety_notes": ["양식 특이사항"]
}
```

**styles 키 네이밍 규칙**: `bullet_1`, `bullet_2`, `bullet_3`, `subsection`, `subsubsection`, `chapter_title_bar`, `table_caption`, `table_header_cell`, `table_body_cell`, `table_unit`, `source_note`, `toc_chapter`, `toc_section`, `reference`, `abstract`, `keyword` 등. 양식에 없는 스타일은 생략.

**치수 필드**: `table_page_width`, `figure_box_width/height/borderFill`은 양식의 기존 표/그림에서 `<hp:sz width="?">`, `<hp:tc borderFillIDRef="?">`를 직접 추출한다.

> **예시**: `examples/metadata-annual-report.json`, `examples/metadata-issue-paper.json` 참조

### 2. usage-guide.md

에이전트가 content dict를 올바르게 구성하기 위한 안내서. 포함할 내용:

- content dict의 전체 구조와 각 필드 설명
- 각 필드의 예시 값
- items 타입별 사용법 (paragraph, table, figure, image)
- parts 포맷 (인라인 서식)
- 이 양식의 특이사항/주의사항
- 자주 틀리는 것들

### 3. builder.py

섹션 파이프라인을 선언하는 파일. 어떤 section 파일에 어떤 모듈을 적용할지 정의한다.

```python
# handler 값은 modules/ 안의 파일명 (문자열)
# build.py가 로드 시 자동으로 modules/{handler}.py → process() 함수로 연결
sections = [
    {'file': 'section0.xml', 'handler': 'cover'},
    {'file': 'section1.xml', 'handler': 'body', 'sec_num': 1},
    # ... 양식 구조에 맞게 선언
]
```

추가 kwargs(예: `sec_num`)는 handler의 `**kwargs`로 전달된다.

> **예시**: `examples/builder-annual-report.py`, `examples/builder-issue-paper.py` 참조

### 4. modules/

각 모듈은 다음 시그니처를 따른다:

```python
def process(xml_data: bytes, content: dict, ctx, **kwargs) -> bytes:
    """섹션 XML을 받아 처리하고 반환."""
```

- `xml_data`: 해당 section 파일의 원본 XML bytes
- `content`: 사용자가 전달한 content dict
- `ctx`: BuildContext 인스턴스
  - `ctx.styles`: metadata.json의 styles dict
  - `ctx.metadata`: metadata.json 전체
  - `ctx.charpr_mgr`: CharPrManager (charPr 복제/재활용)
  - `ctx.register_image(path)`: 이미지 등록
- `**kwargs`: builder.py에서 전달한 추가 인자 (sec_num 등)

xml_primitives 모듈의 함수들을 사용하여 XML 요소를 생성한다:

```python
from hwpx_engine.xml_primitives import make_para, make_table_xml, make_figure_box
```

> **예시**: `examples/module-cover-example.py`, `examples/module-body-example.py` 참조

---

## 핵심 설계 원칙: 예시 텍스트 자동 교체 (필수)

> **이 원칙은 모든 양식 모듈에 반드시 적용해야 한다.**

### 문제

HWPX 양식에는 헤더 문단(TBL/secPr을 포함하는 구조 문단)의 **직접 run**에 예시 텍스트가 잔류하는 경우가 많다. 이 헤더 문단은 구조 요소이므로 삭제할 수 없다. 예시 텍스트를 그대로 두면 사용자가 `cover_replacements`나 `global_replacements`에 원본 텍스트를 정확히 입력해야 하는데, 이는:
- 사용자에게 불필요한 부담
- 원본 텍스트가 런 경계에 걸쳐있으면 교체 자체가 어려움
- 빌드 스크립트가 비대해지고 토큰 낭비

### 해결: 모듈에서 자동 교체

**모듈이 content dict의 해당 필드를 읽어 헤더 문단의 직접 run 텍스트를 자동으로 교체한다.** 사용자는 원본 텍스트를 알 필요 없이 content dict만 작성하면 된다.

### 적용 패턴

양식 분석 시 **삭제할 수 없는 구조 문단(TBL/secPr 포함)에 예시 텍스트가 잔류하는 구역**을 모두 찾아내고, 각각에 대해 content dict의 어떤 필드로 자동 교체할지 결정한다. 양식마다 해당 구역의 종류와 위치가 다르므로 분석 시 직접 확인해야 한다.

**판별 기준**: 문단에 TBL이나 secPr이 있어서 문단 단위 삭제가 불가능한데, 그 문단의 직접 run에 예시 텍스트가 들어있는 경우.

### 구현 방법

```python
# 헤더 문단의 "직접 run"만 교체 (TBL 안의 run은 건드리지 않음)
for run in header_para.findall(f'{{{HP}}}run'):
    if run.getparent() is header_para:  # ← 핵심: 직접 자식 run만
        for t_el in run.findall(f'{{{HP}}}t'):
            t_el.text = content_first_item  # content dict에서 가져온 값
```

**`run.getparent() is header_para`** 체크가 핵심이다. 이 조건 없이 `iter()`하면 TBL 내부의 라벨 run까지 교체되어 구조가 깨진다.

### 분석 시 확인 사항

- [ ] 각 구역의 헤더 문단(TBL/secPr 포함)에 **직접 run 텍스트가 잔류**하는지 확인
- [ ] 잔류하는 텍스트가 content dict의 어떤 필드로 교체되어야 하는지 매핑
- [ ] 모듈에서 해당 자동 교체 로직을 구현
- [ ] 첫 번째 항목은 헤더에서 자동 교체하고, **나머지 항목만 새 문단으로 삽입** (`items[1:]`)

---

## 문서 구역별 처리 전략

### 표지 (Cover)
- **전략**: 텍스트 교체 (구조 유지)
- raw XML 문자열 레벨에서 `<hp:t>` 텍스트를 매칭/교체
- 이미지, 구분선 등 구조 요소는 절대 건드리지 않음
- 긴 키부터 매칭 (부분문자열 충돌 방지)

### 판권지 (Colophon)
- **전략**: 표지와 동일 (텍스트 교체)
- 표 내부 텍스트도 raw 문자열로 교체

### 요약/정책건의
- **전략**: 내용 제거 + 새로 삽입
- 구조 요소 유지: 라벨, 제목, 장식용 구분 표, 빈줄
- 불릿 문단(해당 pPr, sRef) 제거 → 새 내용을 올바른 styleIDRef로 삽입
- 메타 텍스트(스타일 안내) 제거

### 목차 (TOC)
- **전략**: 내용 제거 + 새로 삽입
- 구조 요소 유지: 라벨, 제목, 장식 표, 페이지나눔
- 항목(해당 pPr, sRef) 제거 후 새로 삽입

### 간지 (Interleaf) — 있는 경우
- **전략**: 텍스트 교체 (구조 유지)
- RECT drawText 내부의 서브 문단 텍스트만 교체
- 절 목차: 원본보다 절이 많으면 같은 스타일로 새 문단 추가

### 본문 (Body)
- **전략**: 내용 제거 + 새로 삽입
- 구조 요소(secPr, 간지, 장식용 요소) 유지
- 불릿 문단, 절 제목 등 내용 영역 전체 제거
- 새 내용을 올바른 스타일 조합으로 삽입

### 참고문헌 (References)
- **전략**: 내용 제거 + 새로 삽입
- 헤더 문단(secPr + 라벨 테이블 포함)은 반드시 유지
- 기존 항목만 제거 후 새 항목 삽입

### Abstract / 키워드
- **전략**: 내용 제거 + 새로 삽입 (Abstract), 텍스트 교체 (키워드)
- Abstract 헤더 문단(테이블 포함)은 반드시 유지
- 키워드는 테이블 셀 내부의 `<hp:t>` 텍스트만 변경

---

## XML 규칙과 주의사항

### 기존 문단의 `<hp:t>` 노드 삭제 금지

기존 문단에서 `<hp:t>` 요소를 제거하면 **한글에서 "문서가 손상되었습니다" 오류 발생**.

```python
# ❌ 문서 손상
extra_t.getparent().remove(extra_t)

# ✅ 안전: 문단 전체 삭제 후 새 문단으로 삽입
sec_root.remove(old_para)
anchor.addnext(new_para)
```

| 작업 | 안전? | 설명 |
|------|-------|------|
| `t_el.text = "새 텍스트"` | ✅ | 텍스트 값 변경은 OK |
| `sec_root.remove(para)` | ✅ | 문단 전체 삭제는 OK |
| `anchor.addnext(new_para)` | ✅ | 새 문단 삽입은 OK |
| `run.remove(t_el)` | ❌ | **`<hp:t>` 노드 삭제는 손상** |
| drawText subList 내 `<hp:t>` 제거 | ✅ | RECT 내부의 서브문단은 OK |

**핵심 원리**: 최상위 문단(`<hs:sec>` 직속 `<hp:p>`)의 내부 구조는 건드리지 말고, 문단 단위로 삭제/삽입하라.

### secPr 네임스페이스 주의 — HP (paragraph), HS (section) 아님

`<hp:secPr>`은 **HP 네임스페이스**(`{http://www.hancom.co.kr/hwpml/2011/paragraph}`)에 속한다. 직관적으로 HS(section)일 것 같지만 아니다. secPr은 `<hp:p>` > `<hp:run>` 내부에 위치하므로, 찾을 때는 반드시 깊은 검색을 사용해야 한다:

```python
# 올바른 검색
sec_pr = para.find(f'.//{{{HP}}}secPr')

# 잘못된 검색 (항상 None 반환)
HS = '{http://www.hancom.co.kr/hwpml/2011/section}'
sec_pr = para.find(f'{{{HS}}}secPr')  # ← 이러면 찾지 못함
```

secPr에는 페이지 여백(`<hp:pageMargin>`), masterpage 참조(`<hp:pageDef>`) 등 섹션 레이아웃 정보가 포함되어 있다. 본문 모듈에서 모든 문단을 삭제하고 재생성할 때, secPr을 보존하지 않으면:
- 페이지 여백이 기본값으로 리셋됨
- masterpage(머리말/꼬리말) 참조가 사라짐
- 용지 방향/크기 설정이 유실될 수 있음

**보존 패턴**: 삭제 전 deep copy → 모든 문단 삭제 → 첫 번째 재생성 문단의 `<hp:run>` 안에 복원.

### tc 자식 요소 순서 규격

HWPX는 `<hp:tc>` 내부 자식 요소의 순서가 엄격하다. 순서가 틀리면 한글에서 문서 손상으로 인식한다:

```
올바른 순서: subList → cellAddr → cellSpan → cellSz → cellMargin
```

`make_table_xml()`은 cellAddr를 먼저 넣을 수 있으므로, 모든 make_table_xml 호출 후 반드시 `_fix_tc_child_order()` 후처리를 적용해야 한다:

```python
tbl_p = make_table_xml(...)
_fix_tc_child_order(tbl_p)  # ← 반드시 호출
```

### linesegarray 필수

모든 새 문단에는 반드시 `<hp:linesegarray>` 포함. 없으면 렌더링이 깨진다. `xml_primitives.make_para()`가 자동으로 추가한다.

### charProperties/borderFills의 itemCnt 갱신

header.xml에서 `<charProperties itemCnt="?">` 속성이 있다. 새 charPr이나 borderFill을 추가할 때 이 값을 갱신하지 않으면, 한글 프로그램이 `itemCnt`까지만 읽고 나머지를 **완전히 무시**한다. `CharPrManager`가 자동 처리하지만, 직접 수정할 경우 반드시 갱신할 것.

### fillBrush/winBrush 네임스페이스는 hc: (core)

borderFill 자체는 `hh:` (head) 네임스페이스이지만, 하위 채우기 요소인 `<fillBrush>`와 `<winBrush>`는 **`hc:` (core)** 네임스페이스를 사용해야 한다. `hh:`로 생성하면 한글이 배경색을 무시한다.

```python
HC = '{http://www.hancom.co.kr/hwpml/2011/core}'
fill = etree.SubElement(new_bf, f'{HC}fillBrush')
wb = etree.SubElement(fill, f'{HC}winBrush')
wb.set('faceColor', '#FFFF00')
```

### 셀 음영 변경 시 원래 borderFill 복제

셀 배경색을 바꿀 때 아무 borderFill이나 복제하면 원래 셀의 테두리가 사라진다. 반드시 section XML에서 해당 셀의 현재 `borderFillIDRef`를 찾아서 그것을 복제 + fillBrush 추가.

### 표 안 텍스트를 anchor로 사용하면 표 안에 삽입됨

`insert_paragraph(after='표_안_텍스트')`하면 표 셀 내부에 삽입된다. **표 밖 텍스트(출처 등)를 anchor로 사용할 것.**

### RECT drawText 안 테이블 — 바이트 길이 손상

일부 양식에서 표지 TOC나 머리말이 RECT drawText 안의 테이블 셀에 위치한다. 이런 깊은 중첩 구조의 텍스트를 더 짧은 텍스트로 교체하면 문서가 손상된다.

**해결: 바이트 패딩** — 교체 후 텍스트의 UTF-8 바이트 길이를 원본과 동일하게 맞춘다.

```python
def safe_replace(xml_text, old, new):
    old_bytes = len(old.encode('utf-8'))
    new_bytes = len(new.encode('utf-8'))
    if new_bytes < old_bytes:
        new = new + ' ' * (old_bytes - new_bytes)
    return xml_text.replace(old, new)
```

**판별 방법**: 분석 시 RECT drawText 안에 테이블이 있는지 확인.

### header.xml 지연 쓰기

빌드 엔진에서 section 처리 중 charPr/borderFill을 추가하면 header.xml에 반영해야 하는데, ZIP 순회 시 header.xml이 section*.xml보다 먼저 나온다. header.xml 쓰기를 지연(deferred write)하여 모든 section 처리 완료 후 마지막에 쓰는 패턴 사용. (build.py가 자동 처리)

### save() ZIP-level 수정 유실 방지

hwpx_engine v1.1.0부터 python-hwpx 의존성을 제거하고 직접 ZIP/XML을 다룬다. `HwpxDoc.save_to_path()`는 수정된 섹션 XML만 재직렬화하고 나머지는 원본 바이트를 그대로 보존한다. HwpxEditor는 `_zip_modified` 플래그로 ZIP-level 수정과 DOM-level 수정을 구분한다.

### 빈 문자열 교체 금지

`cover_replacements`에서 텍스트를 빈 문자열(`''`)로 교체하면 `<hp:t></hp:t>`가 되어 문서 손상. 반드시 의미 있는 텍스트로 교체할 것.

### cover_replacements 부분문자열 충돌

짧은 키가 긴 키 안에도 매칭되면 교체 순서에 따라 결과가 달라진다. **긴 키부터 매칭**하면 해결된다. 빌드 엔진이 자동으로 긴 키부터 처리한다.

---

## 표 생성 주의사항

### 캡션 자동번호

양식에 따라 캡션의 자동번호 동작이 다르다:

- **자동번호 있는 양식**: `styleIDRef`의 NUMBER 속성이 `[표 N-M]`을 자동 생성. 캡션 텍스트에 번호를 직접 입력하면 중복된다. 제목만 입력할 것.
- **자동번호 없는 양식**: 캡션에 대괄호를 포함하여 직접 입력 (`[표제목]`).

분석 시 header.xml에서 해당 캡션 스타일의 heading 속성(NUMBER 타입 여부)을 반드시 확인한다.

### 캡션의 장별 sRef 조정

일부 양식에서 표 캡션의 `styleIDRef`가 장 번호에 따라 달라진다 (예: base_sRef에서 장번호만큼 차감). 이 패턴은 양식마다 다르므로 분석 시 확인 필요.

### 셀 스타일 / borderFill

표 셀 스타일과 borderFillIDRef는 양식마다 전부 다르다. 반드시 metadata.json에서 조회:
- 머리(header) 셀: pPr, sRef, cPr, borderFill
- 본문(data) 셀: pPr, sRef, cPr, borderFill
- 좌측/우측 셀의 borderFill이 다를 수 있음

### 빈 셀의 charPr이 셀 높이를 결정

표의 빈 셀(장식용 대각선 셀, border 구분 행 등)도 내부 `<hp:run>`의 `charPrIDRef`에 지정된 폰트 크기에 따라 최소 셀 높이가 달라진다. 텍스트가 없어도 한글은 해당 charPr의 line height를 적용하기 때문이다.

```
charPr height=2400 (12pt, 기본값) → 셀 높이가 불필요하게 커짐
charPr height=100  (1pt)          → 셀 높이가 cellSz만으로 결정됨
```

분석 시 장식 셀/border 행의 charPrIDRef를 반드시 확인하고, metadata에 기록해야 한다. 예: heading TBL의 대각선 셀은 charPr=4(1pt)를 사용.

### 출처/주의 위치: 테이블 내부 vs 외부

양식에 따라 출처/주(`자료:`, `주:`)의 위치가 다르다:

- **테이블 내부 마지막 행**: 이슈&진단처럼 출처를 TBL의 마지막 행으로 넣는 양식. 해당 행의 borderFill은 상단 경계선만 있다(예: bf=14). 이 경우 source를 별도 paragraph로 넣으면 원본과 다르게 보인다.
- **테이블 외부 별도 문단**: 시군전략처럼 출처를 표 바로 아래 독립 paragraph로 넣는 양식. 이 경우 source를 테이블 안에 넣으면 안 된다.

분석 시 반드시 원본 표의 마지막 행이 출처인지, 아니면 표 아래 별도 문단인지 확인할 것.

### 셀병합

HWPX에서 병합된 셀은 XML에 **존재하지 않음**. `headers` 길이 = 전체 컬럼 수. 병합은 `colSpan`/`rowSpan` 속성으로 표현하고, 가려진 위치의 `<hp:tc>`는 생성하지 않는다.

---

## 메타 텍스트 식별 규칙

양식에는 스타일 안내용 텍스트가 포함되어 있다. 이것들은 실제 문서에서 제거해야 한다:

| 패턴 | 예시 |
|------|------|
| `스타일 목록 →` | "스타일 목록 → 표제지 연구번호" |
| 글꼴명으로 시작 | "KoPub돋움체 Bold", "KoPub바탕체 Light" |
| `Npt,` 패턴 | "9pt, 장평 100, 자간 –3," |
| 문단 속성 설명 | "행간 150, 왼쪽 정렬," |
| 작성 안내 | "키워드 입력 필수", "길이에 따라 조절 가능" |

---

## 장 경계(Chapter Delimiter) 판별 — 가짜 경계 주의

일부 양식에서 장 구분이 1x1 테이블 + `pageBreak=1`로 구현된다. 이때 **데이터 테이블도 pageBreak=1을 가질 수 있어** 가짜 장 경계로 오인될 수 있다.

**안전한 판별법**: `pageBreak=1` + `rowCnt=1, colCnt=1` + 제목 텍스트 패턴을 **모두** 확인.

```python
def is_chapter_header_1x1(para):
    tbl = para.find(f'.//{{{HP}}}tbl')
    if tbl is None: return False, ''
    if tbl.get('rowCnt') != '1' or tbl.get('colCnt') != '1':
        return False, ''  # 다중 셀 테이블 = 데이터 → 가짜 경계
    text = ...  # 텍스트 패턴으로 추가 확인
    return True, text
```

---

## charPrIDRef — 가장 자주 틀리는 값

**모든 문단 유형에 대해 원본의 `<hp:run charPrIDRef="?">`를 직접 추출하라.** 추정 금지.

```python
target_styles = {}
for p in paras:
    ppr = p.get('paraPrIDRef', '0')
    sref = p.get('styleIDRef', '0')
    key = (ppr, sref)
    if key not in target_styles:
        for run in p.findall(f'{{{HP}}}run'):
            if run.find(f'.//{{{HP}}}fieldBegin') is not None:
                continue
            cpr = run.get('charPrIDRef', '?')
            target_styles[key] = cpr
            break
```

### 스타일 ID는 양식 버전마다 다르다

같은 양식이라도 메모 포함/미포함, 판형, 스타일 추가/제거에 따라 모든 ID가 달라진다. **코드에 스타일 ID를 절대 하드코딩하지 말 것.** 반드시 metadata.json에서 동적으로 로드.

---

## TOC(목차) 구조

### TOC 감지 방법
pPr 하드코딩이 아닌 **텍스트 기반**으로 감지 (`'TABLES' in text`, `'CONTENTS' in text` 등).

### 절 제목의 탭/리더
원본에 `<hp:tab>` 요소가 인라인으로 포함될 수 있음:
```xml
<hp:t>1. 연구의 배경<hp:tab width="?" leader="3" type="2"/>5</hp:t>
```
- `leader="3"`: 점선 리더
- `type="2"`: 오른쪽 탭 정지

---

## 이미지 삽입

### 네임스페이스 주의
- `<hp:renderingInfo>` 안의 `<transMatrix>`, `<scaMatrix>`, `<rotMatrix>`는 **`hc:`** 네임스페이스
- `<img>`, `<pt0>`~`<pt3>`도 **`hc:`** 네임스페이스
- `hp:`로 생성하면 문서 손상. raw XML 문자열로 직접 작성하여 제어.

### HWPX 내장 구조
- ZIP 내 `BinData/imageN.ext`에 저장
- `Contents/content.hpf`의 `<opf:manifest>`에 항목 추가
- `<hp:pic>` > `<hc:img binaryItemIDRef="imageN">`으로 참조

---

## 체크리스트: 분석 시 확인할 것

### 기본 구조
- [ ] 섹션 파일이 몇 개인지, 각각의 역할
- [ ] 장 구분 방식: 별도 섹션 파일? 1x1 테이블+pageBreak? 기타?
- [ ] 글머리표/개요번호가 텍스트인지 styleIDRef인지 (heading 속성 확인)
- [ ] 각 계층 수준의 정확한 **(paraPrIDRef, styleIDRef, charPrIDRef)** 3개 값 조합
- [ ] 복수 런 구조가 필요한 문단 유형

### 간지 — 있는 경우
- [ ] 간지 구조: RECT drawText 안인지, 일반 문단인지
- [ ] 간지 숫자, 제목, 절 목차의 pPr 확인
- [ ] 간지 숫자 구조: 2개 sub-para? 1개 sub-para에 2개 text node?
- [ ] 간지 처리 시 **장 제목바 문단은 반드시 skip**

### 본문
- [ ] 장 제목바가 표로 구현되어 있는지, 단독 문단인지
- [ ] 본문 시작점 판별 기준
- [ ] 장식용 요소가 있는지
- [ ] secPr 위치 확인: 어떤 문단의 어떤 run 안에 있는지 (HP 네임스페이스로 검색)
- [ ] 본문 재구성 시 secPr 보존/복원 로직이 있는지

### 표
- [ ] 표 캡션이 `<hp:caption>` 안에 있는지, 별도 문단인지
- [ ] 캡션 스타일에 NUMBER 자동번호가 있는지
- [ ] 캡션 sRef가 장 번호에 따라 달라지는지
- [ ] 표 머리/본문 셀의 정확한 (pPr, sRef, cPr, borderFillIDRef)
- [ ] 좌측/우측 셀의 borderFill이 다른지
- [ ] 출처/주가 테이블 내부 마지막 행인지, 외부 별도 문단인지
- [ ] 장식 셀(대각선, border 행 등)의 charPrIDRef — 폰트 크기가 셀 높이에 영향
- [ ] make_table_xml 호출 후 `_fix_tc_child_order()` 적용 여부

### 목차 — 있는 경우
- [ ] `<hp:tab>` 요소 사용 여부
- [ ] 점선 리더 구조

### 참고문헌/Abstract
- [ ] 참고문헌 항목의 정확한 (pPr, sRef, **cPr**)
- [ ] 첫 항목이 헤더 문단에 포함되어 있는지
- [ ] 키워드 값의 정확한 위치

### 치수/크기
- [ ] `table_page_width`: 기존 표의 `<hp:sz width=?>`에서 추출
- [ ] `figure_box_width/height/borderFill`: 기존 그림 박스에서 추출

### 손상 방지
- [ ] RECT drawText 안에 테이블이 있는지 → 바이트 패딩 필수
- [ ] 장 경계 판별 시 데이터 테이블 오인식 방지
- [ ] cover_replacements 부분문자열 충돌 없는지
- [ ] 새 문단에 linesegarray 포함되는지
- [ ] 기존 문단의 `<hp:t>` 삭제 코드가 없는지
- [ ] 빈 문자열(`''`) 교체가 없는지
- [ ] secPr 검색 시 HP 네임스페이스 사용 (HS 아님)
- [ ] tc 자식 순서가 규격대로인지 (subList → cellAddr → cellSpan → cellSz → cellMargin)
- [ ] 본문 전체 삭제/재구성 시 secPr deep copy 보존/복원 있는지

---

## 발견된 교훈 모음

- **스타일 ID 하드코딩 금지**: 같은 양식이라도 버전에 따라 ID가 전부 달라진다. metadata.json에서 동적 로드 필수.
- **빈 문자열 교체는 손상**: `<hp:t></hp:t>`가 되어 한글에서 오류. 의미 있는 텍스트로 교체.
- **TOC 감지는 텍스트 기반**: pPr 하드코딩 대신 `'TABLES' in text` 등으로 범용 감지.
- **source_note 뒤 빈줄은 같은 스타일 유지**: 기본 empty가 아닌 source_note 스타일로.
- **참고문헌 첫 항목이 헤더에 포함된 경우**: 문단 삭제 불가 → `global_replacements`로 텍스트만 교체.
- **절 제목바 복제 시 장식 제거**: clone 후 `<tbl>`, `<ctrl>` 제거, 텍스트 요소만 유지.
- **본문 이미지는 보존하지 않기**: 원본 이미지와 새 콘텐츠 충돌. 이미지 문단도 삭제 대상.
- **캡션 자동번호 양식 vs 수동 양식**: NUMBER 속성 확인. 자동이면 번호 직접 입력 금지.
- **부록은 구분자가 있어야 가능**: 참고문헌 뒤 부록 구분자가 없으면 부록 삽입 불가.
- **그림 박스 borderFill은 양식별**: metadata.json의 `figure_box_borderFill`에서 조회.
- **치수 필드 필수**: `table_page_width`, `figure_box_*`는 양식의 기존 요소에서 직접 추출.
- **외부 문서 편집 시 template_id 주의**: 등록 양식으로 생성한 문서에서만 유효. 외부 문서는 전용 등록 필요.
- **secPr 네임스페이스는 HP (paragraph)**: secPr은 `{http://www.hancom.co.kr/hwpml/2011/paragraph}secPr`이다. HS(section)가 아님. secPr은 `<hp:run>` 안에 위치하므로 `.//{HP}secPr`로 깊은 검색 필수. HS로 검색하면 secPr을 찾지 못해 본문 재구성 시 페이지 여백과 masterpage 참조가 유실된다.
- **tc 자식 순서 규격**: HWPX는 tc 자식 순서가 엄격하다: subList -> cellAddr -> cellSpan -> cellSz -> cellMargin. `make_table_xml`이 cellAddr를 먼저 넣을 수 있으므로, `_fix_tc_child_order()` 후처리를 make_table_xml 호출 후 반드시 적용할 것.
- **빈 셀의 charPr이 셀 높이를 결정**: 표의 장식용 빈 셀(대각선 등)도 charPr의 height(폰트 크기)에 따라 셀 높이가 달라진다. 기본 charPr(12pt)를 쓰면 의도보다 셀이 훨씬 커진다. 예: 대각선 셀은 charPr=4(1pt)로 설정해야 원본과 동일한 높이.
- **출처/주가 테이블 내부인 양식 있음**: 이슈&진단처럼 출처/주를 TBL 마지막 행(bf with 상단 경계선만)으로 넣는 양식이 있다. 시군전략처럼 별도 paragraph로 넣는 양식도 있다. 분석 시 원본 표 구조에서 출처의 위치를 반드시 확인.
- **secPr 복원 필수**: body.py가 모든 문단을 삭제하고 재구성할 때, secPr(페이지 여백, masterpage 참조 포함)이 사라진다. 삭제 전 deep copy하고, 첫 번째 재생성 문단의 run 안에 복원해야 한다.
- **불릿 마커 heading BULLET vs plain 구분 필수**: 양식마다 불릿 스타일이 heading BULLET(마커 자동생성)인지, plain style(마커를 텍스트에 직접 넣어야 함)인지 확인해야 한다. header.xml에서 해당 paraPr의 `<hh:heading type="BULLET"/>` 존재 여부로 판별. heading BULLET이면 usage-guide에서 "텍스트에 마커 넣지 말 것" 명시, plain이면 "마커를 텍스트에 포함" 명시.
- **표 본문 칼럼별 셀 스타일 분석**: 원본 표에서 첫 칼럼(라벨/카테고리)과 나머지 칼럼(본문)이 다른 스타일을 사용하는지 확인. 예: "표가"(볼드)는 첫 칼럼 전용, "표점본문"/"표-왼"/"표본문 오른쪽 정렬"은 본문 칼럼. metadata.json에 `table_first_col`, `table_body_center`, `table_body_right`, `table_body_left` 등으로 등록하고, col_styles 파라미터로 칼럼별 스타일 지정 지원.
