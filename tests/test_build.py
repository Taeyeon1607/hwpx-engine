"""Tests for build pipeline improvements."""
import json
import os
import pytest


class TestMetadataValidation:
    def test_missing_styles_raises(self, tmp_dir):
        """build() should raise clear error when metadata.json lacks 'styles'."""
        from hwpx_engine.build import _validate_metadata
        with pytest.raises(ValueError, match="styles"):
            _validate_metadata({"id": "test", "sections": []})

    def test_missing_multiple_fields(self, tmp_dir):
        from hwpx_engine.build import _validate_metadata
        with pytest.raises(ValueError, match="styles.*sections|sections.*styles"):
            _validate_metadata({"id": "test"})

    def test_valid_metadata_passes(self, tmp_dir):
        from hwpx_engine.build import _validate_metadata
        # Should not raise
        _validate_metadata({"id": "test", "styles": {}, "sections": []})

    def test_missing_id_raises(self, tmp_dir):
        """build() should raise clear error when metadata.json lacks 'id'."""
        from hwpx_engine.build import _validate_metadata
        with pytest.raises(ValueError, match="id"):
            _validate_metadata({"styles": {}, "sections": []})

    def test_empty_metadata_raises(self, tmp_dir):
        """Empty metadata dict should report all missing fields."""
        from hwpx_engine.build import _validate_metadata
        with pytest.raises(ValueError, match="id"):
            _validate_metadata({})

    def test_error_message_lists_all_missing(self, tmp_dir):
        """Error message should name all missing fields, not just the first."""
        from hwpx_engine.build import _validate_metadata
        with pytest.raises(ValueError) as exc_info:
            _validate_metadata({"id": "test"})
        msg = str(exc_info.value)
        assert "styles" in msg
        assert "sections" in msg

    def test_context_appears_in_error_message(self, tmp_dir):
        """Error message should include the context string passed in."""
        from hwpx_engine.build import _validate_metadata
        with pytest.raises(ValueError, match="register"):
            _validate_metadata({"id": "test", "sections": []}, context="register")
