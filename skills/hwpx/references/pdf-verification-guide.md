# PDF 변환 검증 가이드

## 개요

HWPX 문서를 생성/편집한 후, 한글 프로그램의 COM 자동화를 통해 PDF로 변환하여 렌더링 결과를 확인한다.

## 환경 요구사항

| 환경 | PDF 변환 | 비고 |
|------|---------|------|
| **Windows + 한글 설치** | **가능** | pyhwpx (COM 자동화) 사용 |
| macOS | 불가 | pyhwpx는 pywin32 기반 — macOS 미지원 |
| Linux | 불가 | 동일 |

pyhwpx는 **Windows 전용**이다. 내부적으로 `pywin32`의 COM 자동화로 한글 프로그램을 백그라운드 실행하여 변환한다. macOS/Linux에서는 사용할 수 없으며, 대안도 현재 없다.

```bash
pip install pyhwpx
```

## 변환 코드

반드시 아래 함수를 사용할 것. headless 모드(`visible=False`)로 실행하고, 파일 잠금 시 다른 이름으로 저장하는 fallback이 포함되어 있다.

```python
import os
from pyhwpx import Hwp

def hwpx_to_pdf(hwpx_path, pdf_path=None):
    """HWPX를 PDF로 변환한다. Windows + 한글 설치 환경 전용.

    - headless(visible=False)로 실행
    - pdf_path에 파일 잠금(사용자가 열어놓은 경우) 시 _1, _2 ... 붙여서 저장
    """
    hwpx_path = os.path.abspath(hwpx_path)
    if pdf_path is None:
        pdf_path = hwpx_path.rsplit('.', 1)[0] + '.pdf'
    pdf_path = os.path.abspath(pdf_path)

    hwp = Hwp(visible=False)
    try:
        hwp.open(hwpx_path)
        # PDF 저장 시도 — 잠금 시 fallback
        saved_path = _save_with_fallback(hwp, pdf_path)
    finally:
        hwp.quit()
    return saved_path

def _save_with_fallback(hwp, pdf_path, max_retries=10):
    """PDF 저장을 시도하고, 파일 잠금 시 번호를 붙여 다른 이름으로 저장."""
    base, ext = os.path.splitext(pdf_path)
    for i in range(max_retries):
        try_path = pdf_path if i == 0 else f'{base}_{i}{ext}'
        try:
            hwp.save_as(try_path, format='pdf')
            return try_path
        except Exception:
            continue
    raise RuntimeError(f'PDF 저장 실패: {pdf_path} (잠금 또는 권한 문제)')
```

HWPX 저장(`editor.save()`)에서도 동일한 파일 잠금이 발생할 수 있다. 사용자가 hwpx 파일을 한글에서 열어놓은 경우, 에이전트는 다른 이름으로 저장한 뒤 사용자에게 안내한다:
- 예: `output/문서.hwpx` 잠금 시 → `output/문서_1.hwpx`로 저장 → "파일이 잠겨있어서 문서_1.hwpx로 저장했습니다"

## 검증 시점

| 상황 | PDF 검증 | 비고 |
|------|---------|------|
| 새 양식 분석/등록 테스트 | **필수** (자동 반복) | 스타일 매핑이 올바른지 반드시 시각 확인 |
| build()로 새 문서 생성 | 사용자에게 질문 | "PDF로 변환해서 확인해볼까요?" |
| HwpxEditor로 기존 문서 편집 | 사용자에게 질문 | "PDF로 변환해서 확인해볼까요?" |

Windows가 아니거나 한글이 설치되지 않은 환경에서는 PDF 변환을 건너뛰고, 사용자에게 직접 한글로 열어서 확인하도록 안내한다.

## 검증 절차

1. HWPX → PDF 변환 (`hwpx_to_pdf()` 사용)
2. Read 도구로 PDF 열기 (페이지 범위 지정)
3. 변경 사항이 포함된 페이지 확인
4. 문제 발견 시 수정 후 재변환

## 주의사항

- `visible=False`로 headless 실행할 것 — 화면에 한글 창이 뜨지 않음
- 한글이 이미 열려 있으면 COM 충돌 가능 — 한글을 먼저 닫거나 `on_quit=True` 사용
- 대용량 문서(300+ 페이지)는 변환에 1-2분 소요될 수 있음
- PDF 변환 후 Read 도구로 확인할 때 `pages` 파라미터로 관련 페이지만 지정할 것
- `os.path.abspath()`로 절대 경로 변환 필수 — 상대 경로는 한글 COM에서 인식 못할 수 있음

## 삽입 순서 주의

여러 요소(문단, 표, 그림)를 연속 삽입할 때, **그림 박스를 마지막에 삽입**해야 캡션과 박스 사이에 다른 요소가 끼지 않는다. `_zip_insert_paragraph`는 앵커 텍스트의 첫 번째 출현 위치 바로 뒤에 삽입하므로, 같은 앵커를 여러 삽입에 사용하면 나중 삽입이 먼저 삽입된 내용 앞에 끼어든다.

```
올바른 순서: bullet → table → 다른 요소 → figure_box (마지막)
잘못된 순서: bullet → figure_box → table (표가 캡션-박스 사이에 끼어듦)
```

## COM 안정화 패턴

pyhwpx COM 자동화에서 빈번한 오류를 방지하는 패턴:

```python
import os
import subprocess
import time

def safe_hwpx_to_pdf(hwpx_path, pdf_path, max_retries=3):
    """안정적인 PDF 변환. 실패 시 HWP 프로세스 종료 후 재시도."""
    for attempt in range(max_retries):
        subprocess.run(["taskkill", "/F", "/IM", "Hwp.exe"],
                       capture_output=True)
        time.sleep(5)

        try:
            from pyhwpx import Hwp
            hwp = Hwp(visible=False)
            time.sleep(3)
            hwp.open(os.path.abspath(hwpx_path))
            time.sleep(3)
            hwp.save_as(os.path.abspath(pdf_path))
            hwp.quit()

            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
                return True
        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            try:
                hwp.quit()
            except:
                pass

    return False
```

### 주요 에러 및 대처

| 에러 | 원인 | 대처 |
|------|------|------|
| "서버에서 예외 오류가 발생했습니다" | HWP 프로세스 충돌 또는 이전 인스턴스 잔존 | `taskkill /F /IM Hwp.exe` → sleep 5초 → 재시도 |
| PermissionError: 다른 프로세스가 사용 중 | Dropbox 동기화 또는 HWP 파일 잠금 | sleep 후 재시도, 또는 /tmp에 복사 후 작업 |
| PDF 파일 크기 < 1KB | 변환 실패 (빈 PDF) | HWP 프로세스 완전 종료 후 재시도 |
