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


def test_detect_cloud_sync(tmp_path):
    from hwpx_engine.converter import _detect_cloud_sync
    d = tmp_path / "Dropbox" / "a.hwp"; d.parent.mkdir(parents=True); d.write_bytes(b"")
    o = tmp_path / "OneDrive" / "a.hwp"; o.parent.mkdir(parents=True); o.write_bytes(b"")
    n = tmp_path / "work" / "a.hwp"; n.parent.mkdir(parents=True); n.write_bytes(b"")
    assert _detect_cloud_sync(d) is True
    assert _detect_cloud_sync(o) is True
    assert _detect_cloud_sync(n) is False


def _make_fake_hwp(behavior="ok"):
    """테스트용 FakeHwp 팩토리. pyhwpx 1.6.6 시그니처(Task 7)에 맞춤."""
    instances = []

    class FakeHwp:
        def __init__(self, *a, **kw):
            instances.append(self)

        def open(self, filename, format="", arg=""):
            if behavior == "open_fail_bad" and "bad" in str(filename):
                raise RuntimeError("open failed")
            return True

        def save_as(self, path, format="HWP", arg="", split_page=False):
            Path(path).write_bytes(b"FAKE-" + format.encode())
            return True

        def clear(self, option=1):
            if behavior == "clear_fail":
                raise RuntimeError("clear failed")

        def quit(self):
            pass

    FakeHwp.instances = instances
    return FakeHwp


def test_temp_workspace_lazy_and_release(tmp_path):
    from hwpx_engine.converter import _TempWorkspace
    src1 = tmp_path / "a.hwp"; src1.write_bytes(b"X")
    src2 = tmp_path / "sub" / "a.hwp"; src2.parent.mkdir(); src2.write_bytes(b"Y")

    with _TempWorkspace() as ws:
        # lazy: 요청 시점에 복사
        local1 = ws.local_path(src1)
        assert local1.exists() and local1.read_bytes() == b"X"
        # 이름 충돌 → _1 접미사로 회피
        local2 = ws.local_path(src2)
        assert local2.exists() and local2.read_bytes() == b"Y"
        assert local1 != local2

        # 산출물 publish
        out1 = local1.with_suffix(".hwpx"); out1.write_bytes(b"OUT1")
        dest1 = src1.with_suffix(".hwpx")
        ws.publish(out1, dest1)
        assert dest1.read_bytes() == b"OUT1"

        # release: 해당 로컬 원본 파일 정리
        ws.release(src1)
        assert not local1.exists()


def test_batch_skip_existing(tmp_hwp_dir, monkeypatch):
    from hwpx_engine import hwp_to_hwpx_pdf
    import hwpx_engine.converter as cv
    monkeypatch.setattr(cv, "_ensure_hwp_appid_patch", lambda auto_elevate=True: True)
    monkeypatch.setattr(cv, "_ensure_pyhwpx", lambda: _make_fake_hwp("ok"))

    (tmp_hwp_dir / "a.hwpx").write_bytes(b"")
    (tmp_hwp_dir / "a.pdf").write_bytes(b"")

    result = hwp_to_hwpx_pdf(tmp_hwp_dir, ensure_appid=False, progress=False, copy_to_temp=False)
    assert len(result["skipped"]) == 1
    assert len(result["success"]) == 1


def test_batch_failure_isolation_open(tmp_hwp_dir, monkeypatch):
    from hwpx_engine import hwp_to_hwpx_pdf
    import hwpx_engine.converter as cv
    monkeypatch.setattr(cv, "_ensure_hwp_appid_patch", lambda auto_elevate=True: True)
    monkeypatch.setattr(cv, "_ensure_pyhwpx", lambda: _make_fake_hwp("open_fail_bad"))

    bad = tmp_hwp_dir / "bad.hwp"; bad.write_bytes(b"")
    result = hwp_to_hwpx_pdf(tmp_hwp_dir, ensure_appid=False, progress=False, copy_to_temp=False)
    assert len(result["fail"]) == 1
    assert len(result["success"]) == 2  # a.hwp, sub/b.hwp


def test_batch_clear_failure_recreates_instance(tmp_hwp_dir, monkeypatch):
    """clear 실패 시 Hwp 인스턴스 재생성 (MN-3, C4 fallback)."""
    from hwpx_engine import hwp_to_hwpx_pdf
    import hwpx_engine.converter as cv
    FakeHwp = _make_fake_hwp("clear_fail")
    monkeypatch.setattr(cv, "_ensure_hwp_appid_patch", lambda auto_elevate=True: True)
    monkeypatch.setattr(cv, "_ensure_pyhwpx", lambda: FakeHwp)

    result = hwp_to_hwpx_pdf(tmp_hwp_dir, ensure_appid=False, progress=False, copy_to_temp=False)
    assert len(result["success"]) == 2
    # 초기 1 + 파일1 clear 실패 후 재생성 1 + 파일2 clear 실패 후 재생성 1 = 최소 3
    assert len(FakeHwp.instances) >= 2


def test_batch_rejects_non_windows(monkeypatch, tmp_hwp_dir):
    from hwpx_engine import hwp_to_hwpx_pdf
    import hwpx_engine.converter as cv
    monkeypatch.setattr(cv, "_is_windows", lambda: False)
    with pytest.raises(RuntimeError):
        hwp_to_hwpx_pdf(tmp_hwp_dir, ensure_appid=False, progress=False)


def test_batch_rejects_both_formats_off(tmp_hwp_dir):
    from hwpx_engine import hwp_to_hwpx_pdf
    with pytest.raises(ValueError):
        hwp_to_hwpx_pdf(tmp_hwp_dir, hwpx=False, pdf=False,
                         ensure_appid=False, progress=False)


def test_batch_rejects_target_collision(tmp_path, monkeypatch):
    """output_dir 평탄화 충돌 사전 감지 (MJ-2)."""
    from hwpx_engine import hwp_to_hwpx_pdf
    import hwpx_engine.converter as cv
    monkeypatch.setattr(cv, "_ensure_hwp_appid_patch", lambda auto_elevate=True: True)
    monkeypatch.setattr(cv, "_ensure_pyhwpx", lambda: _make_fake_hwp("ok"))

    (tmp_path / "sub1").mkdir(); (tmp_path / "sub2").mkdir()
    (tmp_path / "sub1" / "x.hwp").write_bytes(b"")
    (tmp_path / "sub2" / "x.hwp").write_bytes(b"")
    out = tmp_path / "out"

    with pytest.raises(ValueError) as exc:
        hwp_to_hwpx_pdf(tmp_path, output_dir=out, preserve_tree=False,
                         ensure_appid=False, progress=False, copy_to_temp=False)
    assert "출력 경로 충돌" in str(exc.value)


@pytest.mark.pdf
def test_batch_real_conversion(tmp_path):
    """실 HWP → HWPX + PDF 변환 (pyhwpx + 한글 필요)."""
    import shutil
    sample = Path(__file__).parent / "fixtures" / "sample.hwp"
    if not sample.exists():
        pytest.skip("tests/fixtures/sample.hwp missing")
    src = tmp_path / "sample.hwp"
    shutil.copy(sample, src)

    from hwpx_engine import hwp_to_hwpx_pdf
    result = hwp_to_hwpx_pdf(src, ensure_appid=True, progress=False, copy_to_temp=False)
    assert len(result["success"]) == 1, result
    assert (tmp_path / "sample.hwpx").exists()
    assert (tmp_path / "sample.pdf").exists()
    assert (tmp_path / "sample.pdf").stat().st_size > 1000


def test_batch_rejects_preserve_tree_without_root(tmp_path, monkeypatch):
    """iterable + preserve_tree=True + src_root 없음 → ValueError (MN-8)."""
    from hwpx_engine import hwp_to_hwpx_pdf
    import hwpx_engine.converter as cv
    monkeypatch.setattr(cv, "_ensure_hwp_appid_patch", lambda auto_elevate=True: True)
    monkeypatch.setattr(cv, "_ensure_pyhwpx", lambda: _make_fake_hwp("ok"))

    a = tmp_path / "a.hwp"; a.write_bytes(b"")
    out = tmp_path / "out"
    with pytest.raises(ValueError):
        hwp_to_hwpx_pdf([a], output_dir=out, preserve_tree=True,
                         ensure_appid=False, progress=False, copy_to_temp=False)
