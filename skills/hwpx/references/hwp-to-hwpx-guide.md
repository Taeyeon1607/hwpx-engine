# HWP → HWPX/PDF 배치 변환 가이드

## 개요

레거시 `.hwp` 파일을 `.hwpx`(OWPML) 또는 `.pdf`로 일괄 변환한다. pyhwpx가 한글 프로그램을 백그라운드(`visible=False`)로 구동해 "다른 이름으로 저장".

## 환경 요구사항

- Windows 10/11
- 한컴 한글 2020 이상
- Python 3.10+
- `pip install pyhwpx`
- **첫 실행 1회**: `hwpx-apply-appid` 콘솔 명령으로 관리자 UAC 승인

## 첫 실행 2줄

```powershell
pip install pyhwpx
hwpx-apply-appid
```

UAC 승인 후 `AppID 상태: {'wow64_runas': 'Interactive User', 'native_runas': 'Interactive User'}` 출력되면 준비 완료. 이후 재부팅해도 유지.

## 사용

```python
from hwpx_engine import hwp_to_hwpx_pdf

result = hwp_to_hwpx_pdf("C:/work/reports")
print(result)
# {"success": [Path, ...], "skipped": [Path, ...], "fail": [(Path, err), ...]}
```

- 디렉토리 → `rglob("*.hwp")`
- 같은 이름 산출물 이미 있으면 skip
- Dropbox/OneDrive 경로 자동 감지 → `%TEMP%`로 lazy 복사 후 변환
- 단일 `Hwp()` 인스턴스 재사용 + 실패 시 자동 재생성

## 파라미터

| 이름 | 기본 | 설명 |
|------|------|------|
| `sources` | 필수 | 단일 `.hwp` / 디렉토리 / 경로 iterable (1-level) |
| `hwpx` | True | HWPX 산출 |
| `pdf` | True | PDF 산출 |
| `output_dir` | None | 지정 시 그 폴더에 저장 |
| `skip_existing` | True | 이미 있으면 skip |
| `ensure_appid` | True | 첫 호출 시 자동 패치 시도 |
| `progress` | True | True=print, callable=사용자 정의, False=무음 |
| `copy_to_temp` | `"auto"` | True/False/"auto". 클라우드 경로 자동 감지 |
| `preserve_tree` | False | output_dir 지정 시 하위 폴더 구조 유지 |

## 반환

```python
{
  "success": [Path, ...],
  "skipped": [Path, ...],
  "fail":    [(Path, err_str), ...],
}
```

파일별 실패는 격리 — 93개 중 1개 깨져도 92개는 정상 완료.

## output_dir 덮어쓰기 방지

`output_dir` 지정 + `preserve_tree=False`(기본)에서 하위 폴더의 동명 파일이 겹치면 `ValueError: 출력 경로 충돌`. 이 경우 `preserve_tree=True`로 호출해 폴더 트리를 유지할 것.

```python
hwp_to_hwpx_pdf("C:/work", output_dir="C:/out", preserve_tree=True)
```

## tqdm 진행바 연동 (선택)

```python
from tqdm import tqdm
bar = tqdm()
def cb(i, total, path):
    if bar.total != total:
        bar.reset(total=total)
    bar.update(1); bar.set_description(path.name)

hwp_to_hwpx_pdf("C:/work", progress=cb)
bar.close()
```

## 주의사항

### 이미 한글이 떠 있는 경우
배치 시작 전 한글 완전 종료 권장. 실패 다발 시 `taskkill /F /IM Hwp.exe` 후 재시도.

### 구버전 HWP (한글 2014 이하)
`visible=False`에서 호환성 대화상자로 교착 가능. 사전에 한글로 한 번 열어 "한글 2020 형식으로 저장" 후 배치.

### 비밀번호 걸린 HWP
`hwp.open`이 프롬프트로 교착. 배치 전 제외.

### 한글 업데이트 후 AppID 소실
`hwpx-apply-appid` 재실행으로 복구. 관리자 UAC 한 번 더.

## 수동 AppID 적용 (자동 실패 시)

```powershell
hwpx-apply-appid
```

한 줄. 관리자 UAC 프롬프트 → 승인. 실패 시 sha8 캐시 파일이 `%TEMP%\hwpx_engine_set_hwp_appid_<sha8>.ps1` 로 남으므로 PowerShell 관리자로 `& "<경로>"` 직접 실행도 가능.

## 성능 참고

- 로컬 SSD + 한글 2020: 파일당 수초(단순 문서). 대용량/복잡 문서는 증가
- 단일 인스턴스 재사용이 핵심. 인스턴스 재생성은 2~3배 느림
- 클라우드 경로는 `copy_to_temp="auto"` 기본 동작으로 로컬 복사 후 변환

## 에러 대처

| 증상 | 원인 | 대처 |
|------|------|------|
| `RuntimeError: pyhwpx 패키지가 설치돼 있지 않습니다` | pyhwpx 없음 | `pip install pyhwpx` |
| `RuntimeError: Windows + 한글 ...` | macOS/Linux 호출 | Windows 환경에서 실행 |
| `RuntimeError: UAC 승인 대기 120초 초과` | UAC 승인 지연 | `hwpx-apply-appid` 재실행 후 즉시 승인 |
| `pywintypes.com_error 서버 실행 실패` | AppID 미패치 | `hwpx-apply-appid` 실행 |
| `ValueError: 출력 경로 충돌` | output_dir 평탄화 충돌 | `preserve_tree=True` |
| `fail` 다수 | 한글 이미 실행 중 | `taskkill /F /IM Hwp.exe` 후 재시도 |
