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
