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


def _default_progress(i: int, total: int, path: Path) -> None:
    """기본 진행률 출력. tqdm이 필요하면 가이드의 callback 예시 참조."""
    print("[{}/{}] {}".format(i, total, path.name), flush=True)


def _compute_target(src: Path, out_dir: Optional[Path], ext: str,
                     preserve_tree: bool, src_root: Optional[Path]) -> Path:
    """산출물 저장 경로 계산.

    - out_dir 없음: src와 같은 폴더에 확장자 교체
    - out_dir 있음 + preserve_tree=False: out_dir/<filename>
    - out_dir 있음 + preserve_tree=True: src_root 기준 상대경로 미러링
      (src_root가 None이면 ValueError — iterable sources는 공통 루트 모호)
    """
    if out_dir is None:
        return src.with_suffix(ext)
    if not preserve_tree:
        return out_dir / src.with_suffix(ext).name
    if src_root is None:
        raise ValueError(
            "preserve_tree=True를 iterable sources와 함께 쓰려면 공통 루트가 필요합니다. "
            "디렉토리 단일 인자로 호출하거나 preserve_tree=False를 사용하세요."
        )
    try:
        rel = src.relative_to(src_root)
    except ValueError:
        rel = Path(src.name)
    return out_dir / rel.with_suffix(ext)


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
    if not hwpx and not pdf:
        raise ValueError("hwpx와 pdf 중 적어도 하나는 True여야 합니다.")
    if not _is_windows():
        raise RuntimeError(
            "이 기능은 Windows + 한글 설치 환경에서만 동작합니다. 현재: {}".format(sys.platform)
        )

    if ensure_appid:
        # MJ-A: 타임아웃 RuntimeError가 전파되면 배치 전체 중단 → 이미 적용돼 있을
        # 수도 있으므로 진행은 계속하고 실패는 파일별 fail로 드러나게 한다.
        try:
            _ensure_hwp_appid_patch(auto_elevate=True)
        except RuntimeError as _appid_err:
            print(
                "[warn] AppID 자동 패치 타임아웃/거부: {}. "
                "변환은 계속 시도합니다. 전부 실패하면 관리자 PowerShell에서 "
                "`hwpx-apply-appid`를 먼저 실행하세요.".format(_appid_err),
                file=sys.stderr,
            )

    Hwp = _ensure_pyhwpx()
    out_dir = Path(output_dir).resolve() if output_dir else None
    if out_dir:
        out_dir.mkdir(parents=True, exist_ok=True)

    paths = _iter_hwp_sources(sources)
    total = len(paths)
    result: dict = {"success": [], "skipped": [], "fail": []}
    if total == 0:
        return result

    # src_root 계산 (preserve_tree용)
    src_root: Optional[Path] = None
    if isinstance(sources, (str, Path)) and Path(sources).is_dir():
        src_root = Path(sources).resolve()

    # MJ-2: target 충돌 사전 검사 (output_dir 평탄화 덮어쓰기 방지)
    if out_dir is not None:
        targets_seen: Dict[str, Path] = {}
        for p in paths:
            p_res = p.resolve()
            for ext_flag, ext in ((hwpx, ".hwpx"), (pdf, ".pdf")):
                if not ext_flag:
                    continue
                t = _compute_target(p_res, out_dir, ext, preserve_tree, src_root)
                key = str(t.resolve())
                if key in targets_seen:
                    raise ValueError(
                        "출력 경로 충돌: {} 와 {} 가 동일한 산출물({})을 생성합니다. "
                        "preserve_tree=True를 쓰거나 sources/output_dir 구조를 조정하세요."
                        .format(p, targets_seen[key], t)
                    )
                targets_seen[key] = p

    # copy_to_temp 결정
    use_temp = copy_to_temp
    if use_temp == "auto":
        use_temp = any(_detect_cloud_sync(p) for p in paths)

    def _emit(i: int, src: Path) -> None:
        if progress is False or progress is None:
            return
        if callable(progress):
            progress(i, total, src)
            return
        _default_progress(i, total, src)

    # MN-7: 초기 안내
    if progress is True and total > 0:
        print("HWP 변환 시작: {}개 파일".format(total), flush=True)

    hwp = Hwp(visible=False)
    ws: Optional[_TempWorkspace] = None
    try:
        if use_temp:
            ws = _TempWorkspace()

        for i, src in enumerate(paths, 1):
            src = src.resolve()
            target_hwpx = _compute_target(src, out_dir, ".hwpx", preserve_tree, src_root) if hwpx else None
            target_pdf = _compute_target(src, out_dir, ".pdf", preserve_tree, src_root) if pdf else None

            need_hwpx = bool(target_hwpx) and not (skip_existing and target_hwpx.exists())
            need_pdf = bool(target_pdf) and not (skip_existing and target_pdf.exists())
            if not need_hwpx and not need_pdf:
                # ws.local_path 호출 전이므로 release 불필요
                result["skipped"].append(src)
                _emit(i, src)
                continue

            _emit(i, src)

            work_src = ws.local_path(src) if ws else src
            local_hwpx = work_src.with_suffix(".hwpx") if need_hwpx else None
            local_pdf = work_src.with_suffix(".pdf") if need_pdf else None

            try:
                hwp.open(str(work_src))
                # Task 7 결과: save_as는 format 인자가 필수. 대문자로 명시.
                if need_hwpx:
                    hwp.save_as(str(local_hwpx), format="HWP")
                if need_pdf:
                    hwp.save_as(str(local_pdf), format="PDF")
                try:
                    hwp.clear(1)
                except Exception:
                    # clear 실패 → 인스턴스 오염 가능성 → 재생성 (C4 fallback)
                    try: hwp.quit()
                    except Exception: pass
                    hwp = Hwp(visible=False)
                # 산출물 publish
                if ws:
                    if need_hwpx: ws.publish(local_hwpx, target_hwpx)
                    if need_pdf: ws.publish(local_pdf, target_pdf)
                else:
                    if need_hwpx and local_hwpx != target_hwpx:
                        import shutil
                        target_hwpx.parent.mkdir(parents=True, exist_ok=True)
                        shutil.move(str(local_hwpx), str(target_hwpx))
                    if need_pdf and local_pdf != target_pdf:
                        import shutil
                        target_pdf.parent.mkdir(parents=True, exist_ok=True)
                        shutil.move(str(local_pdf), str(target_pdf))
                result["success"].append(src)
            except Exception as e:
                result["fail"].append((src, repr(e)))
                # 실패 시 인스턴스 재생성으로 다음 파일 오염 방지
                try: hwp.quit()
                except Exception: pass
                hwp = Hwp(visible=False)
            finally:
                if ws:
                    ws.release(src)
    finally:
        try: hwp.quit()
        except Exception: pass
        if ws is not None:
            try: ws.__exit__(None, None, None)
            except Exception: pass

    return result


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
