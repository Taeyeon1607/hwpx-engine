"""HWP → HWPX/PDF 배치 변환기 (Windows + 한글 전용).

공개 API:
    hwp_to_hwpx_pdf(sources, hwpx=True, pdf=True, output_dir=None,
                    skip_existing=True, ensure_appid=True, progress=True,
                    copy_to_temp="auto", preserve_tree=False)

콘솔 진입점 (pyproject.toml [project.scripts]):
    hwpx-apply-appid  →  _cli_apply_appid()

계약:
    - sources: str/Path (파일/디렉토리) 또는 iterable of str/Path (1-level만)
    - 첫 실행 시 DCOM AppID 자동 패치 (ShellExecuteEx + WaitForSingleObject 동기)
    - Dropbox/OneDrive 자동 감지 → %TEMP%로 lazy 복사·변환·publish
    - 파일별 실패 격리 + 실패 시 hwp 인스턴스 재생성 fallback

반환:
    {"success": [Path, ...], "skipped": [Path, ...], "fail": [(Path, str), ...]}
"""
from __future__ import annotations

import sys
import time
from pathlib import Path
from typing import Callable, Iterable, Optional, Union, List, Dict


PathLike = Union[str, Path]
SourcesArg = Union[PathLike, Iterable[PathLike]]


# NOTE: 아래 _is_windows 등 모듈 레벨 유틸은 반드시 `_is_windows()` 형태로 호출
# 한다(모듈 __dict__ lookup → LOAD_GLOBAL). 이렇게 해야 테스트에서
# monkeypatch.setattr(cv, "_is_windows", ...) 가 내부 호출에도 적용된다.
# 내부 함수에서 `from .converter import _is_windows` 같은 from-import를 하면
# 로컬 바인딩이 되어 monkeypatch가 먹히지 않으니 금지.
def _is_windows() -> bool:
    return sys.platform.startswith("win") or sys.platform == "cygwin"


def _is_admin() -> bool:
    if not _is_windows():
        return False
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _ensure_pyhwpx():
    """pyhwpx import → Hwp 클래스 반환. 실패 시 한글 안내 RuntimeError."""
    try:
        from pyhwpx import Hwp
        return Hwp
    except ImportError as e:
        raise RuntimeError(
            "pyhwpx 패키지가 설치돼 있지 않습니다.\n\n"
            "설치:\n  pip install pyhwpx\n\n"
            "이 기능은 Windows + 한글(HWP 2020 이상) 설치 환경 전용입니다."
        ) from e


def hwp_to_hwpx_pdf(
    sources: SourcesArg,
    hwpx: bool = True,
    pdf: bool = True,
    output_dir: Optional[PathLike] = None,
    skip_existing: bool = True,
    ensure_appid: bool = True,
    progress: Union[bool, Callable[[int, int, Path], None]] = True,
    copy_to_temp: Union[bool, str] = "auto",
    preserve_tree: bool = False,
) -> dict:
    raise NotImplementedError


def _cli_apply_appid() -> None:
    """Console entry point: apply HWP DCOM AppID patch (UAC prompt)."""
    raise NotImplementedError
