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


_HWP_CLSID = "{2291CF00-64A1-4877-A9B4-68CFE89612D6}"
_TARGET_RUNAS = "Interactive User"
_PS1_TIMEOUT_SEC = 120


def _resolve_ps1_path() -> Optional[Path]:
    """패키지 리소스에서 set_hwp_appid.ps1 경로를 얻는다.

    importlib.resources는 editable 설치에서도 Traversable 인터페이스를 반환할 수
    있어 `isinstance(x, Path)` 검사는 믿을 수 없다. 항상 내용을 TEMP에 sha8 기반
    파일명으로 캐시하되 내용 비교로 불필요한 재쓰기를 회피한다.
    """
    if not _is_windows():
        return None
    try:
        from importlib.resources import files
        import hashlib
        import tempfile
        ref = files("hwpx_engine._ps1") / "set_hwp_appid.ps1"
        if not ref.is_file():
            return None
        content = ref.read_bytes()
        digest = hashlib.sha1(content).hexdigest()[:8]
        tmp = Path(tempfile.gettempdir()) / "hwpx_engine_set_hwp_appid_{}.ps1".format(digest)
        if not tmp.exists() or tmp.read_bytes() != content:
            tmp.write_bytes(content)
        return tmp
    except Exception:
        return None


def _read_appid_status() -> dict:
    """reg query로 AppID 두 뷰의 RunAs 값을 읽어 dict로 반환."""
    status = {"wow64_runas": None, "native_runas": None}
    if not _is_windows():
        return status
    import subprocess
    for key, view in [("wow64_runas", "32"), ("native_runas", "64")]:
        cmd = ["reg", "query",
               r"HKLM\SOFTWARE\Classes\AppID\{}".format(_HWP_CLSID),
               "/v", "RunAs", "/reg:" + view]
        try:
            out = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
            if out.returncode == 0:
                for line in out.stdout.splitlines():
                    line = line.strip()
                    if line.startswith("RunAs") and "REG_SZ" in line:
                        parts = line.split("REG_SZ", 1)
                        if len(parts) == 2:
                            status[key] = parts[1].strip()
                            break
        except Exception:
            pass
    return status


def _appid_already_applied(status: Optional[dict] = None) -> bool:
    s = status if status is not None else _read_appid_status()
    return s.get("wow64_runas") == _TARGET_RUNAS and s.get("native_runas") == _TARGET_RUNAS


def _run_ps1_sync(ps1_path: Optional[Path], elevate: bool,
                   timeout_sec: int = _PS1_TIMEOUT_SEC) -> bool:
    """PS1을 동기 실행. 성공 시 True, 실패 시 False, 타임아웃 시 RuntimeError.

    elevate=False: subprocess.run으로 직접 실행 (이미 관리자)
    elevate=True: ShellExecuteExW + WaitForSingleObject로 UAC 동기 대기
    """
    if not _is_windows() or ps1_path is None or not ps1_path.exists():
        return False

    if not elevate:
        import subprocess
        try:
            r = subprocess.run(
                ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
                 "-File", str(ps1_path)],
                capture_output=True, text=True, timeout=timeout_sec,
            )
            return r.returncode == 0
        except subprocess.TimeoutExpired as e:
            raise RuntimeError(
                "PS1 실행 타임아웃({}s): {}".format(timeout_sec, ps1_path)
            ) from e
        except Exception:
            return False

    # elevate=True: ShellExecuteExW(runas) + WaitForSingleObject
    import ctypes
    from ctypes import wintypes

    SEE_MASK_NOCLOSEPROCESS = 0x00000040
    SW_HIDE = 0
    WAIT_OBJECT_0 = 0
    WAIT_TIMEOUT = 0x00000102

    class SHELLEXECUTEINFOW(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("fMask", ctypes.c_ulong),
            ("hwnd", wintypes.HWND),
            ("lpVerb", wintypes.LPCWSTR),
            ("lpFile", wintypes.LPCWSTR),
            ("lpParameters", wintypes.LPCWSTR),
            ("lpDirectory", wintypes.LPCWSTR),
            ("nShow", ctypes.c_int),
            ("hInstApp", wintypes.HINSTANCE),
            ("lpIDList", ctypes.c_void_p),
            ("lpClass", wintypes.LPCWSTR),
            ("hkeyClass", wintypes.HKEY),
            ("dwHotKey", wintypes.DWORD),
            ("hIconOrMonitor", wintypes.HANDLE),
            ("hProcess", wintypes.HANDLE),
        ]

    info = SHELLEXECUTEINFOW()
    info.cbSize = ctypes.sizeof(SHELLEXECUTEINFOW)
    info.fMask = SEE_MASK_NOCLOSEPROCESS
    info.hwnd = 0
    info.lpVerb = "runas"
    info.lpFile = "powershell.exe"
    info.lpParameters = '-NoProfile -ExecutionPolicy Bypass -File "{}"'.format(str(ps1_path))
    info.lpDirectory = None
    info.nShow = SW_HIDE
    info.hInstApp = 0

    shell32 = ctypes.windll.shell32
    kernel32 = ctypes.windll.kernel32
    ok = shell32.ShellExecuteExW(ctypes.byref(info))
    if not ok or not info.hProcess:
        return False

    wait_result = kernel32.WaitForSingleObject(info.hProcess, int(timeout_sec * 1000))

    if wait_result == WAIT_OBJECT_0:
        exit_code = wintypes.DWORD()
        kernel32.GetExitCodeProcess(info.hProcess, ctypes.byref(exit_code))
        kernel32.CloseHandle(info.hProcess)
        return exit_code.value == 0
    elif wait_result == WAIT_TIMEOUT:
        kernel32.CloseHandle(info.hProcess)
        raise RuntimeError(
            "UAC 승인 대기 {}초 초과. 관리자 PowerShell에서 `hwpx-apply-appid` 수동 실행하세요."
            .format(timeout_sec)
        )
    else:
        kernel32.CloseHandle(info.hProcess)
        return False


def _ensure_hwp_appid_patch(auto_elevate: bool = True) -> bool:
    """DCOM AppID 패치 보장. 이미/성공이면 True, 실패면 False. UAC 타임아웃이면 RuntimeError."""
    if not _is_windows():
        return False
    if _appid_already_applied():
        return True

    ps1 = _resolve_ps1_path()
    if ps1 is None:
        return False

    if _is_admin():
        ok = _run_ps1_sync(ps1, elevate=False)
    elif auto_elevate:
        ok = _run_ps1_sync(ps1, elevate=True)  # 타임아웃 시 RuntimeError 전파
    else:
        return False

    if not ok:
        return False
    # 반영 대기
    for _ in range(6):
        if _appid_already_applied():
            return True
        time.sleep(0.5)
    return _appid_already_applied()


def _iter_hwp_sources(sources: SourcesArg) -> List[Path]:
    """입력을 정규화해 HWP 파일 경로 리스트를 반환.

    - str/Path: 디렉토리면 rglob("*.hwp"), 파일이면 그 파일
    - iterable of str/Path: 1-level만 (중첩은 TypeError)
    """
    result: List[Path] = []

    def _handle(item) -> None:
        if not isinstance(item, (str, Path)):
            raise TypeError(
                "Nested iterables not supported. Use a flat sequence of paths."
            )
        p = Path(item)
        if p.is_dir():
            for f in sorted(p.rglob("*.hwp")):
                if f.is_file():
                    result.append(f)
            return
        if p.is_file():
            if p.suffix.lower() != ".hwp":
                raise ValueError("Not a .hwp file: {}".format(p))
            result.append(p)
            return
        raise FileNotFoundError(str(p))

    if isinstance(sources, (str, Path)):
        _handle(sources)
    else:
        try:
            items = list(sources)
        except TypeError:
            raise TypeError("sources must be str/Path or iterable of str/Path")
        for it in items:
            _handle(it)

    return result


def _detect_cloud_sync(path: Path) -> bool:
    """경로의 어떤 부분에 'Dropbox' 또는 'OneDrive'가 포함되는지 판정."""
    markers = ("dropbox", "onedrive")
    parts = [p.lower() for p in Path(path).resolve().parts]
    return any(m in parts for m in markers)


class _TempWorkspace:
    """소스 파일을 %TEMP%로 요청 시(lazy) 복사. 파일별 release 지원.

    사용:
        with _TempWorkspace() as ws:
            local = ws.local_path(original)  # 요청 시 복사
            ... 변환 ...
            ws.publish(local_output, dest)
            ws.release(original)             # 로컬 원본 + 남은 산출물 cleanup

    release는 주로 **실패 경로 cleanup**이 목적이다. 성공 경로에서는 publish가
    local 산출물을 대상 위치로 이미 이동시켰으므로 side 산출물은 존재하지
    않는다 — FileNotFoundError는 try/except로 묵살해 no-op이 된다.
    """

    def __init__(self) -> None:
        import tempfile
        self._root = Path(tempfile.mkdtemp(prefix="hwpx_batch_"))
        self._map: Dict[str, Path] = {}

    def _copy_in(self, src: Path) -> Path:
        import shutil
        local = self._root / src.name
        i = 1
        while local.exists():
            local = self._root / "{}_{}{}".format(src.stem, i, src.suffix)
            i += 1
        shutil.copy2(str(src), str(local))
        return local

    def local_path(self, original: PathLike) -> Path:
        key = str(Path(original).resolve())
        if key not in self._map:
            self._map[key] = self._copy_in(Path(original))
        return self._map[key]

    def release(self, original: PathLike) -> None:
        key = str(Path(original).resolve())
        local = self._map.pop(key, None)
        if local is None:
            return
        try:
            local.unlink()
        except Exception:
            pass
        # 변환 실패 시 남은 산출물 정리 (성공 경로에서는 이미 publish로 이동됨)
        for ext in (".hwpx", ".pdf"):
            side = local.with_suffix(ext)
            if side.exists():
                try:
                    side.unlink()
                except Exception:
                    pass

    def publish(self, local_output: Path, dest: Path) -> Path:
        import shutil
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(local_output), str(dest))
        return dest

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        import shutil
        shutil.rmtree(str(self._root), ignore_errors=True)
        return False


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
    """Console entry point: hwpx-apply-appid.

    Applies the HWP DCOM AppID patch. UAC will prompt if not admin.
    """
    import sys as _sys
    if not _is_windows():
        print("이 명령은 Windows 전용입니다.", file=_sys.stderr)
        _sys.exit(2)
    try:
        ok = _ensure_hwp_appid_patch(auto_elevate=True)
    except RuntimeError as e:
        print("실패: {}".format(e), file=_sys.stderr)
        _sys.exit(1)
    print("AppID 상태:", _read_appid_status())
    _sys.exit(0 if ok else 1)
