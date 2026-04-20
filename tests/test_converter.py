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
