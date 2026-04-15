"""Tests for HwpxEditor methods."""
import re
import pytest
from hwpx_engine.editor import HwpxEditor


class TestFindText:
    def test_find_exact_string(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        results = editor.find_text("Second paragraph")
        assert len(results) >= 1
        assert any("Second paragraph" in r["text"] for r in results)
        del editor

    def test_find_returns_context(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        results = editor.find_text("항목A")
        assert len(results) >= 1
        assert "항목A" in results[0]["context"]
        del editor

    def test_find_regex_pattern(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        results = editor.find_text(re.compile(r"\d{3}"))
        # Should find "100", "200", "300", "400" in table cells
        assert len(results) >= 4
        del editor

    def test_find_no_match(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        results = editor.find_text("nonexistent text xyz")
        assert results == []
        del editor

    def test_find_in_specific_section(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        results = editor.find_text("First", section="Contents/section0.xml")
        assert len(results) >= 1
        del editor


class TestGetCell:
    def test_get_header_cell(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        assert editor.get_cell(0, 0, 0) == "구분"
        assert editor.get_cell(0, 0, 1) == "2027"
        assert editor.get_cell(0, 0, 2) == "2028"
        del editor

    def test_get_data_cell(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        assert editor.get_cell(0, 1, 0) == "항목A"
        assert editor.get_cell(0, 1, 1) == "100"
        assert editor.get_cell(0, 2, 1) == "300"
        del editor

    def test_get_cell_merged_returns_empty(self, merge_hwpx):
        editor = HwpxEditor.open(merge_hwpx)
        assert editor.get_cell(0, 0, 0) == "Header1"
        assert editor.get_cell(0, 1, 0) == "A"
        # Deactivated cell should return empty
        assert editor.get_cell(0, 2, 0) == ""
        del editor

    def test_get_cell_index_error_table(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        with pytest.raises(IndexError):
            editor.get_cell(99, 0, 0)
        del editor

    def test_get_cell_index_error_row(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        with pytest.raises(IndexError):
            editor.get_cell(0, 99, 0)
        del editor

    def test_get_cell_index_error_col(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        with pytest.raises(IndexError):
            editor.get_cell(0, 0, 99)
        del editor


class TestGetTableData:
    def test_full_table(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        data = editor.get_table_data(0)
        assert data == [
            ["구분", "2027", "2028"],
            ["항목A", "100", "200"],
            ["항목B", "300", "400"],
        ]
        del editor

    def test_merged_table(self, merge_hwpx):
        editor = HwpxEditor.open(merge_hwpx)
        data = editor.get_table_data(0)
        assert len(data) == 4
        assert data[0][0] == "Header1"
        assert data[1][0] == "A"
        assert data[2][0] == ""  # deactivated
        assert data[3] == ["F", "G", "H"]
        del editor

    def test_get_table_data_index_error(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        with pytest.raises(IndexError):
            editor.get_table_data(99)
        del editor


class TestInsertParagraph:
    def test_insert_after_top_level_paragraph(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        editor.insert_paragraph(text="Inserted text", after="First paragraph", style="body")
        results = editor.find_text("Inserted text")
        assert len(results) >= 1
        editor.save(minimal_hwpx)
        del editor

        # Verify persistence
        editor2 = HwpxEditor.open(minimal_hwpx)
        results2 = editor2.find_text("Inserted text")
        assert len(results2) >= 1
        del editor2

    def test_anchor_skips_table_cell_text(self, minimal_hwpx):
        """When anchor text exists both in a table cell and in a top-level paragraph,
        the insertion should happen after the top-level paragraph, not inside the table."""
        editor = HwpxEditor.open(minimal_hwpx)
        # "Third paragraph" is a top-level paragraph
        editor.insert_paragraph(text="After third", after="Third paragraph", style="body")
        results = editor.find_text("After third")
        assert len(results) >= 1
        del editor

    def test_anchor_not_found_raises(self, minimal_hwpx):
        """Inserting with a non-existent anchor raises TextNotFoundError."""
        from hwpx_engine.editor import TextNotFoundError
        editor = HwpxEditor.open(minimal_hwpx)
        with pytest.raises(TextNotFoundError):
            editor.insert_paragraph(text="X", after="nonexistent anchor text xyz", style="body")
        del editor

    def test_anchor_inside_table_cell_skipped(self, ambiguous_hwpx):
        """When anchor text exists only inside a table cell (not as top-level paragraph),
        the function should raise TextNotFoundError (not silently insert inside cell)."""
        from hwpx_engine.editor import TextNotFoundError
        editor = HwpxEditor.open(ambiguous_hwpx)
        # "cell_only_text" appears only inside a table cell, not as a top-level paragraph
        with pytest.raises(TextNotFoundError):
            editor.insert_paragraph(text="After cell", after="cell_only_text", style="body")
        del editor


class TestFindTable:
    def test_find_by_header_text(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        indices = editor.find_table("구분")
        assert 0 in indices
        del editor

    def test_find_by_data_row(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        indices = editor.find_table("항목A", match_row=1)
        assert 0 in indices
        del editor

    def test_find_no_match(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        indices = editor.find_table("nonexistent")
        assert indices == []
        del editor


class TestBatchReplace:
    def test_single_replacement(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        result = editor.batch_replace([("First paragraph", "Modified paragraph")])
        assert result["applied"] == 1
        editor.save(minimal_hwpx)
        del editor

        editor2 = HwpxEditor.open(minimal_hwpx)
        results = editor2.find_text("Modified paragraph")
        assert len(results) >= 1
        assert editor2.find_text("First paragraph") == []
        del editor2

    def test_multiple_replacements(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        result = editor.batch_replace([
            ("First", "1st"),
            ("Second", "2nd"),
        ])
        assert result["applied"] == 2
        editor.save(minimal_hwpx)
        del editor

    def test_skipped_when_not_found(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        result = editor.batch_replace([
            ("First paragraph", "Modified"),
            ("nonexistent xyz 999", "whatever"),
        ])
        assert result["applied"] == 1
        assert result["skipped"] == 1
        del editor


class TestSetCell:
    def test_set_and_verify(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        editor.set_cell(0, 1, 1, "999")
        assert editor.get_cell(0, 1, 1) == "999"
        del editor

    def test_set_cell_persists(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        editor.set_cell(0, 1, 1, "CHANGED")
        editor.save(minimal_hwpx)
        del editor

        editor2 = HwpxEditor.open(minimal_hwpx)
        assert editor2.get_cell(0, 1, 1) == "CHANGED"
        del editor2

    def test_set_cell_empty_string(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        editor.set_cell(0, 0, 0, "")
        assert editor.get_cell(0, 0, 0) == ""
        del editor


class TestRemoveParagraph:
    def test_remove_existing(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        editor.remove_paragraph("Second paragraph")
        results = editor.find_text("Second paragraph")
        assert results == []
        del editor

    def test_remove_nonexistent_raises(self, minimal_hwpx):
        from hwpx_engine.editor import TextNotFoundError
        editor = HwpxEditor.open(minimal_hwpx)
        with pytest.raises(TextNotFoundError):
            editor.remove_paragraph("nonexistent xyz")
        del editor

    def test_remove_returns_count(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        count = editor.remove_paragraph("Third paragraph")
        assert count == 1
        del editor
