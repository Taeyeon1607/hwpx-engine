"""Tests for table structural operations (delete_row, add_row)."""
import pytest
from hwpx_engine.editor import HwpxEditor


class TestDeleteRow:
    def test_delete_last_row(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        data_before = editor.get_table_data(0)
        assert len(data_before) == 3
        editor.delete_row(0, 2)  # Delete last data row (항목B)
        data_after = editor.get_table_data(0)
        assert len(data_after) == 2
        assert data_after[1][0] == "항목A"
        del editor

    def test_delete_preserves_other_rows(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        editor.delete_row(0, 1)  # Delete first data row (항목A)
        data = editor.get_table_data(0)
        assert len(data) == 2
        assert data[0] == ["구분", "2027", "2028"]
        assert data[1][0] == "항목B"
        del editor

    def test_delete_row_out_of_range_raises(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        with pytest.raises(IndexError):
            editor.delete_row(0, 99)
        del editor

    def test_delete_row_invalid_table_raises(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        with pytest.raises(IndexError):
            editor.delete_row(99, 0)
        del editor

    def test_delete_row_persists_after_save(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        editor.delete_row(0, 2)
        editor.save(minimal_hwpx)
        del editor

        editor2 = HwpxEditor.open(minimal_hwpx)
        data = editor2.get_table_data(0)
        assert len(data) == 2
        del editor2


class TestAddRow:
    def test_add_row_at_end(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        editor.add_row(0, ["항목C", "500", "600"])
        data = editor.get_table_data(0)
        assert len(data) == 4
        assert data[3][0] == "항목C"
        assert data[3][1] == "500"
        assert data[3][2] == "600"
        del editor

    def test_add_row_increases_count(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        before = len(editor.get_table_data(0))
        editor.add_row(0, ["new", "1", "2"])
        after = len(editor.get_table_data(0))
        assert after == before + 1
        del editor

    def test_add_row_persists_after_save(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        editor.add_row(0, ["항목D", "700", "800"])
        editor.save(minimal_hwpx)
        del editor

        editor2 = HwpxEditor.open(minimal_hwpx)
        data = editor2.get_table_data(0)
        assert len(data) == 4
        assert data[3][0] == "항목D"
        del editor2

    def test_add_row_invalid_table_raises(self, minimal_hwpx):
        editor = HwpxEditor.open(minimal_hwpx)
        with pytest.raises(IndexError):
            editor.add_row(99, ["a", "b", "c"])
        del editor
