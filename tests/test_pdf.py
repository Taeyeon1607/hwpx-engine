"""PDF conversion tests — requires pyhwpx and Hangul WP installed.

Run with: pytest -m pdf
Skip automatically if pyhwpx is not installed.
"""

import os
import time

import pytest

try:
    from pyhwpx import Hwp
    HAS_PYHWPX = True
except ImportError:
    HAS_PYHWPX = False


pytestmark = pytest.mark.pdf


@pytest.mark.skipif(not HAS_PYHWPX, reason="pyhwpx not installed")
class TestPdfConversion:
    def test_minimal_to_pdf(self, minimal_hwpx, tmp_dir):
        """Convert minimal hwpx to PDF and verify file is created."""
        pdf_path = os.path.join(tmp_dir, "output.pdf")
        hwp = Hwp(visible=False)
        try:
            hwp.open(os.path.abspath(minimal_hwpx))
            time.sleep(2)
            hwp.save_as(os.path.abspath(pdf_path))
        finally:
            try:
                hwp.quit()
            except Exception:
                pass

        assert os.path.exists(pdf_path)
        assert os.path.getsize(pdf_path) > 1000

    def test_edited_file_to_pdf(self, minimal_hwpx, tmp_dir):
        """Edit a file then convert to PDF — verifies editability."""
        from hwpx_engine.editor import HwpxEditor

        editor = HwpxEditor.open(minimal_hwpx)
        editor.set_cell(0, 1, 1, "PDF_TEST_VALUE")
        editor.save(minimal_hwpx)
        del editor

        pdf_path = os.path.join(tmp_dir, "edited.pdf")
        hwp = Hwp(visible=False)
        try:
            hwp.open(os.path.abspath(minimal_hwpx))
            time.sleep(2)
            hwp.save_as(os.path.abspath(pdf_path))
        finally:
            try:
                hwp.quit()
            except Exception:
                pass

        assert os.path.exists(pdf_path)
        assert os.path.getsize(pdf_path) > 1000
