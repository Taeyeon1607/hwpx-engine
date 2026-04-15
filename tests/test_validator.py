"""Tests for HWPX validator."""
import os
import pytest
from hwpx_engine.validator import HwpxValidator


class TestLevel1:
    def test_valid_minimal(self, minimal_hwpx):
        result = HwpxValidator.validate(minimal_hwpx)
        assert result.level1_passed

    def test_valid_merge(self, merge_hwpx):
        result = HwpxValidator.validate(merge_hwpx)
        assert result.level1_passed

    def test_invalid_file(self, tmp_dir):
        bad_path = os.path.join(tmp_dir, "bad.hwpx")
        with open(bad_path, "w") as f:
            f.write("not a zip file")
        result = HwpxValidator.validate(bad_path)
        assert not result.level1_passed

    def test_missing_file(self, tmp_dir):
        missing_path = os.path.join(tmp_dir, "does_not_exist.hwpx")
        result = HwpxValidator.validate(missing_path)
        assert not result.level1_passed
        assert not result.valid
        assert len(result.errors) >= 1


class TestLevel2:
    def test_valid_minimal_passes_level2(self, minimal_hwpx):
        result = HwpxValidator.validate(minimal_hwpx)
        assert result.level2_passed

    def test_result_valid_when_no_errors(self, minimal_hwpx):
        result = HwpxValidator.validate(minimal_hwpx)
        assert result.valid
        assert result.errors == []


class TestValidationResult:
    def test_stats_populated(self, minimal_hwpx):
        result = HwpxValidator.validate(minimal_hwpx)
        assert "file_size_kb" in result.stats
        assert result.stats["file_size_kb"] > 0

    def test_stats_paragraph_count(self, minimal_hwpx):
        result = HwpxValidator.validate(minimal_hwpx)
        # minimal_hwpx has 3 paragraphs + 1 table wrapper paragraph
        assert result.stats.get("paragraphs", 0) >= 3

    def test_stats_table_count(self, minimal_hwpx):
        result = HwpxValidator.validate(minimal_hwpx)
        assert result.stats.get("tables", 0) >= 1
