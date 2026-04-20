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
