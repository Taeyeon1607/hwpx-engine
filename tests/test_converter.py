"""HWP→HWPX/PDF 배치 변환기 테스트.

파일 레벨 pytestmark 없음 — 한글·pyhwpx 불필요한 테스트는 기본 CI에서 실행.
실 한글 프로세스가 필요한 테스트에만 @pytest.mark.pdf 개별 부착.
"""
import os
import sys
from pathlib import Path

import pytest


def test_public_api_importable():
    from hwpx_engine import hwp_to_hwpx_pdf
    assert callable(hwp_to_hwpx_pdf)


def test_is_windows_detection():
    from hwpx_engine.converter import _is_windows
    assert _is_windows() == (sys.platform.startswith("win") or sys.platform == "cygwin")


def test_is_admin_returns_bool():
    from hwpx_engine.converter import _is_admin
    result = _is_admin()
    assert result is True or result is False


def test_ensure_pyhwpx_raises_helpful_message(monkeypatch):
    """pyhwpx import 실패 시 한글 안내 RuntimeError."""
    import builtins
    monkeypatch.delitem(sys.modules, "pyhwpx", raising=False)
    real_import = builtins.__import__

    def fake_import(name, *args, **kwargs):
        if name == "pyhwpx":
            raise ImportError("No module named 'pyhwpx'")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fake_import)
    from hwpx_engine.converter import _ensure_pyhwpx

    with pytest.raises(RuntimeError) as exc:
        _ensure_pyhwpx()
    msg = str(exc.value)
    assert "pyhwpx" in msg and "pip install pyhwpx" in msg
    assert "Windows" in msg and "한글" in msg


def test_read_appid_status_returns_dict():
    from hwpx_engine.converter import _read_appid_status
    status = _read_appid_status()
    assert isinstance(status, dict)
    assert set(status.keys()) >= {"wow64_runas", "native_runas"}


def test_ensure_appid_returns_true_when_already_applied(monkeypatch):
    import hwpx_engine.converter as cv
    calls = {"run": 0}
    monkeypatch.setattr(cv, "_appid_already_applied", lambda status=None: True)
    monkeypatch.setattr(cv, "_run_ps1_sync",
                        lambda *a, **k: calls.__setitem__("run", calls["run"] + 1) or True)
    assert cv._ensure_hwp_appid_patch(auto_elevate=True) is True
    assert calls["run"] == 0


def test_ensure_appid_non_windows_returns_false(monkeypatch):
    import hwpx_engine.converter as cv
    monkeypatch.setattr(cv, "_is_windows", lambda: False)
    assert cv._ensure_hwp_appid_patch(auto_elevate=True) is False


def test_resolve_ps1_path_returns_cached_file():
    from hwpx_engine.converter import _resolve_ps1_path
    if not (sys.platform.startswith("win") or sys.platform == "cygwin"):
        pytest.skip("Windows only")
    p = _resolve_ps1_path()
    assert p is not None and p.exists()
    # sha8 기반 이름이라 두 번 호출해도 같은 경로
    p2 = _resolve_ps1_path()
    assert p == p2


def test_run_ps1_sync_returns_false_when_no_path():
    """ps1_path=None이면 False."""
    from hwpx_engine.converter import _run_ps1_sync
    assert _run_ps1_sync(None, elevate=False) is False


@pytest.mark.skipif(not sys.platform.startswith("win"),
                     reason="Windows only")
def test_run_ps1_sync_timeout_raises(tmp_path):
    """실 타임아웃 유도 — 30초 sleep ps1에 1초 타임아웃 (MJ-B)."""
    from hwpx_engine.converter import _run_ps1_sync
    ps1 = tmp_path / "slow.ps1"
    ps1.write_text("Start-Sleep -Seconds 30", encoding="utf-8")
    with pytest.raises(RuntimeError, match="타임아웃"):
        _run_ps1_sync(ps1, elevate=False, timeout_sec=1)


@pytest.fixture
def tmp_hwp_dir(tmp_path):
    """임시 디렉토리에 빈 HWP/텍스트 파일 혼재 생성."""
    for rel in ["a.hwp", "sub/b.hwp", "c.txt"]:
        p = tmp_path / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(b"")
    return tmp_path


def test_iter_sources_single_file(tmp_hwp_dir):
    from hwpx_engine.converter import _iter_hwp_sources
    out = list(_iter_hwp_sources(tmp_hwp_dir / "a.hwp"))
    assert len(out) == 1 and out[0].name == "a.hwp"


def test_iter_sources_directory_rglob(tmp_hwp_dir):
    from hwpx_engine.converter import _iter_hwp_sources
    out = sorted(p.name for p in _iter_hwp_sources(tmp_hwp_dir))
    assert out == ["a.hwp", "b.hwp"]  # c.txt 제외


def test_iter_sources_list_tuple_set(tmp_hwp_dir):
    from hwpx_engine.converter import _iter_hwp_sources
    paths = [tmp_hwp_dir / "a.hwp", tmp_hwp_dir / "sub" / "b.hwp"]
    assert len(list(_iter_hwp_sources(paths))) == 2
    assert len(list(_iter_hwp_sources(tuple(paths)))) == 2
    assert len(list(_iter_hwp_sources({paths[0]}))) == 1


def test_iter_sources_rejects_nested(tmp_hwp_dir):
    from hwpx_engine.converter import _iter_hwp_sources
    with pytest.raises(TypeError):
        list(_iter_hwp_sources([[tmp_hwp_dir / "a.hwp"]]))


def test_iter_sources_rejects_nonexistent(tmp_hwp_dir):
    from hwpx_engine.converter import _iter_hwp_sources
    with pytest.raises(FileNotFoundError):
        list(_iter_hwp_sources(tmp_hwp_dir / "nope.hwp"))
