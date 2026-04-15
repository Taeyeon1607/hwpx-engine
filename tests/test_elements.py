"""Tests for elements.py — footnotes, endnotes, headers, footers, page numbers, memos."""
from __future__ import annotations

import os
import zipfile

import pytest
from lxml import etree

from hwpx_engine.hwpx_doc import HwpxDoc
from hwpx_engine.elements import (
    add_footnote,
    add_footnote_to_doc,
    add_endnote,
    set_header,
    set_footer,
    set_page_number,
    add_memo_to_doc,
)

_HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _section_xml(root: etree._Element) -> str:
    """Serialize section root to string for inspection."""
    return etree.tostring(root, encoding="unicode")


def _open_and_get_root(path: str) -> tuple:
    """Open a doc and return (doc, filename, root) for first section."""
    doc = HwpxDoc.open(path)
    sections = doc.sections
    assert len(sections) >= 1
    filename, root = sections[0]
    return doc, filename, root


# ---------------------------------------------------------------------------
# Footnote tests
# ---------------------------------------------------------------------------

class TestAddFootnote:
    def test_basic_footnote(self, minimal_hwpx):
        """Add a footnote anchored to existing text."""
        doc = HwpxDoc.open(minimal_hwpx)
        add_footnote(doc, "First paragraph", "This is footnote 1")

        _, root = doc.sections[0]
        fn_elements = list(root.iter(f"{_HP}footNote"))
        assert len(fn_elements) == 1

        fn = fn_elements[0]
        assert fn.get("number") == "1"
        assert fn.get("suffixChar") == "41"

        # Check footnote body text
        texts = [t.text for t in fn.iter(f"{_HP}t") if t.text]
        assert any("This is footnote 1" in t for t in texts)

        # Check autoNum
        auto_nums = list(fn.iter(f"{_HP}autoNum"))
        assert len(auto_nums) == 1
        assert auto_nums[0].get("numType") == "FOOTNOTE"
        assert auto_nums[0].get("num") == "1"

        doc.close()

    def test_multiple_footnotes(self, minimal_hwpx):
        """Add multiple footnotes; numbering should increment."""
        doc = HwpxDoc.open(minimal_hwpx)
        add_footnote(doc, "First paragraph", "Footnote A")
        add_footnote(doc, "Second paragraph", "Footnote B")

        _, root = doc.sections[0]
        fn_elements = list(root.iter(f"{_HP}footNote"))
        assert len(fn_elements) == 2
        assert fn_elements[0].get("number") == "1"
        assert fn_elements[1].get("number") == "2"

        doc.close()

    def test_footnote_persists_after_save(self, minimal_hwpx, tmp_dir):
        """Footnote survives save/reload cycle."""
        doc = HwpxDoc.open(minimal_hwpx)
        add_footnote(doc, "First paragraph", "Persistent footnote")

        out_path = os.path.join(tmp_dir, "fn_save.hwpx")
        doc.save_to_path(out_path)
        doc.close()

        # Reload and verify
        doc2 = HwpxDoc.open(out_path)
        _, root2 = doc2.sections[0]
        fn_elements = list(root2.iter(f"{_HP}footNote"))
        assert len(fn_elements) == 1
        texts = [t.text for t in fn_elements[0].iter(f"{_HP}t") if t.text]
        assert any("Persistent footnote" in t for t in texts)
        doc2.close()

    def test_footnote_anchor_not_found(self, minimal_hwpx):
        """ValueError when anchor text doesn't exist."""
        doc = HwpxDoc.open(minimal_hwpx)
        with pytest.raises(ValueError, match="not found"):
            add_footnote(doc, "nonexistent text", "Note")
        doc.close()

    def test_footnote_invalid_section(self, minimal_hwpx):
        """ValueError when section index is out of range."""
        doc = HwpxDoc.open(minimal_hwpx)
        with pytest.raises(ValueError, match="out of range"):
            add_footnote(doc, "First", "Note", section_index=99)
        doc.close()

    def test_backward_compat_alias(self, minimal_hwpx):
        """add_footnote_to_doc is the same function as add_footnote."""
        assert add_footnote_to_doc is add_footnote

    def test_footnote_marks_dirty(self, minimal_hwpx):
        """Adding a footnote marks the section as dirty."""
        doc = HwpxDoc.open(minimal_hwpx)
        assert len(doc._dirty_sections) == 0
        add_footnote(doc, "First paragraph", "Note")
        assert "Contents/section0.xml" in doc._dirty_sections
        doc.close()


# ---------------------------------------------------------------------------
# Endnote tests
# ---------------------------------------------------------------------------

class TestAddEndnote:
    def test_basic_endnote(self, minimal_hwpx):
        """Add an endnote anchored to existing text."""
        doc = HwpxDoc.open(minimal_hwpx)
        add_endnote(doc, "Second paragraph", "This is endnote 1")

        _, root = doc.sections[0]
        en_elements = list(root.iter(f"{_HP}endNote"))
        assert len(en_elements) == 1

        en = en_elements[0]
        assert en.get("number") == "1"

        # Check autoNum type
        auto_nums = list(en.iter(f"{_HP}autoNum"))
        assert len(auto_nums) == 1
        assert auto_nums[0].get("numType") == "ENDNOTE"

        doc.close()

    def test_endnote_persists(self, minimal_hwpx, tmp_dir):
        """Endnote survives save/reload."""
        doc = HwpxDoc.open(minimal_hwpx)
        add_endnote(doc, "Third paragraph", "Endnote text")

        out_path = os.path.join(tmp_dir, "en_save.hwpx")
        doc.save_to_path(out_path)
        doc.close()

        doc2 = HwpxDoc.open(out_path)
        _, root2 = doc2.sections[0]
        en_elements = list(root2.iter(f"{_HP}endNote"))
        assert len(en_elements) == 1
        doc2.close()

    def test_endnote_anchor_not_found(self, minimal_hwpx):
        doc = HwpxDoc.open(minimal_hwpx)
        with pytest.raises(ValueError, match="not found"):
            add_endnote(doc, "missing text", "Note")
        doc.close()


# ---------------------------------------------------------------------------
# Header tests
# ---------------------------------------------------------------------------

class TestSetHeader:
    def test_add_header(self, minimal_hwpx):
        """Add a new header to a section."""
        doc = HwpxDoc.open(minimal_hwpx)
        set_header(doc, "My Header Text")

        _, root = doc.sections[0]
        headers = list(root.iter(f"{_HP}header"))
        assert len(headers) >= 1

        # Verify text
        found_text = False
        for h in headers:
            for t in h.iter(f"{_HP}t"):
                if t.text and "My Header Text" in t.text:
                    found_text = True
        assert found_text, "Header text not found in XML"

        doc.close()

    def test_header_apply_page_type(self, minimal_hwpx):
        """Header has correct applyPageType attribute."""
        doc = HwpxDoc.open(minimal_hwpx)
        set_header(doc, "Both header", apply_page_type="BOTH")
        set_header(doc, "Odd header", apply_page_type="ODD")

        _, root = doc.sections[0]
        headers = list(root.iter(f"{_HP}header"))
        types = {h.get("applyPageType") for h in headers}
        assert "BOTH" in types
        assert "ODD" in types

        doc.close()

    def test_replace_existing_header(self, minimal_hwpx):
        """Setting header twice with same page type replaces text."""
        doc = HwpxDoc.open(minimal_hwpx)
        set_header(doc, "First version")
        set_header(doc, "Second version")

        _, root = doc.sections[0]
        headers = list(root.iter(f"{_HP}header"))
        both_headers = [h for h in headers if h.get("applyPageType") == "BOTH"]
        assert len(both_headers) == 1

        texts = [t.text for t in both_headers[0].iter(f"{_HP}t") if t.text]
        assert "Second version" in texts
        assert "First version" not in texts

        doc.close()

    def test_header_persists(self, minimal_hwpx, tmp_dir):
        """Header survives save/reload cycle."""
        doc = HwpxDoc.open(minimal_hwpx)
        set_header(doc, "Persistent Header")

        out_path = os.path.join(tmp_dir, "hdr_save.hwpx")
        doc.save_to_path(out_path)
        doc.close()

        doc2 = HwpxDoc.open(out_path)
        _, root2 = doc2.sections[0]
        headers = list(root2.iter(f"{_HP}header"))
        assert len(headers) >= 1
        found = any(
            t.text and "Persistent Header" in t.text
            for h in headers
            for t in h.iter(f"{_HP}t")
        )
        assert found
        doc2.close()

    def test_header_invalid_section(self, minimal_hwpx):
        doc = HwpxDoc.open(minimal_hwpx)
        with pytest.raises(ValueError, match="out of range"):
            set_header(doc, "Header", section_index=5)
        doc.close()

    def test_header_marks_dirty(self, minimal_hwpx):
        doc = HwpxDoc.open(minimal_hwpx)
        set_header(doc, "Header")
        assert "Contents/section0.xml" in doc._dirty_sections
        doc.close()


# ---------------------------------------------------------------------------
# Footer tests
# ---------------------------------------------------------------------------

class TestSetFooter:
    def test_add_footer(self, minimal_hwpx):
        """Add a new footer to a section."""
        doc = HwpxDoc.open(minimal_hwpx)
        set_footer(doc, "My Footer Text")

        _, root = doc.sections[0]
        footers = list(root.iter(f"{_HP}footer"))
        assert len(footers) >= 1

        found = any(
            t.text and "My Footer Text" in t.text
            for f in footers
            for t in f.iter(f"{_HP}t")
        )
        assert found

        doc.close()

    def test_footer_vert_align(self, minimal_hwpx):
        """Footer subList should have vertAlign=BOTTOM."""
        doc = HwpxDoc.open(minimal_hwpx)
        set_footer(doc, "Bottom footer")

        _, root = doc.sections[0]
        for footer in root.iter(f"{_HP}footer"):
            sub = footer.find(f"{_HP}subList")
            if sub is not None:
                assert sub.get("vertAlign") == "BOTTOM"

        doc.close()

    def test_replace_existing_footer(self, minimal_hwpx):
        """Setting footer twice replaces the text."""
        doc = HwpxDoc.open(minimal_hwpx)
        set_footer(doc, "Old footer")
        set_footer(doc, "New footer")

        _, root = doc.sections[0]
        both_footers = [
            f for f in root.iter(f"{_HP}footer")
            if f.get("applyPageType") == "BOTH"
        ]
        assert len(both_footers) == 1
        texts = [t.text for t in both_footers[0].iter(f"{_HP}t") if t.text]
        assert "New footer" in texts

        doc.close()

    def test_footer_persists(self, minimal_hwpx, tmp_dir):
        doc = HwpxDoc.open(minimal_hwpx)
        set_footer(doc, "Persistent Footer")
        out_path = os.path.join(tmp_dir, "ftr_save.hwpx")
        doc.save_to_path(out_path)
        doc.close()

        doc2 = HwpxDoc.open(out_path)
        _, root2 = doc2.sections[0]
        footers = list(root2.iter(f"{_HP}footer"))
        found = any(
            t.text and "Persistent Footer" in t.text
            for f in footers
            for t in f.iter(f"{_HP}t")
        )
        assert found
        doc2.close()


# ---------------------------------------------------------------------------
# Page number tests
# ---------------------------------------------------------------------------

class TestSetPageNumber:
    def test_basic_page_number(self, minimal_hwpx):
        """Add a page number in footer."""
        doc = HwpxDoc.open(minimal_hwpx)
        set_page_number(doc)

        _, root = doc.sections[0]
        # Should create a footer with autoNum PAGE
        auto_nums = [
            an for an in root.iter(f"{_HP}autoNum")
            if an.get("numType") == "PAGE"
        ]
        assert len(auto_nums) >= 1
        assert auto_nums[0].get("numType") == "PAGE"

        # Check format type
        fmt = auto_nums[0].find(f"{_HP}autoNumFormat")
        assert fmt is not None
        assert fmt.get("type") == "DIGIT"

        doc.close()

    def test_page_number_format(self, minimal_hwpx):
        """Page number with custom format string."""
        doc = HwpxDoc.open(minimal_hwpx)
        set_page_number(doc, fmt="Page {page}")

        _, root = doc.sections[0]
        # Find the footer containing the page number
        for footer in root.iter(f"{_HP}footer"):
            texts = [t.text for t in footer.iter(f"{_HP}t") if t.text]
            if any("Page" in t for t in texts):
                break
        else:
            # Also check if autoNum is in a header (for header position)
            pass

        # The prefix "Page " should exist as text
        all_t = list(root.iter(f"{_HP}t"))
        found_prefix = any(t.text and "Page " in t.text for t in all_t)
        assert found_prefix, "Format prefix 'Page ' not found in XML"

        doc.close()

    def test_page_number_roman(self, minimal_hwpx):
        """Page number with Roman numeral format."""
        doc = HwpxDoc.open(minimal_hwpx)
        set_page_number(doc, number_type="ROMAN_SMALL")

        _, root = doc.sections[0]
        auto_nums = [
            an for an in root.iter(f"{_HP}autoNum")
            if an.get("numType") == "PAGE"
        ]
        assert len(auto_nums) >= 1
        fmt = auto_nums[0].find(f"{_HP}autoNumFormat")
        assert fmt.get("type") == "ROMAN_SMALL"

        doc.close()

    def test_page_number_in_header(self, minimal_hwpx):
        """Page number placed in header position."""
        doc = HwpxDoc.open(minimal_hwpx)
        set_page_number(doc, position="header_center")

        _, root = doc.sections[0]
        # Should have a header (not footer) with autoNum PAGE
        for header in root.iter(f"{_HP}header"):
            auto_nums = list(header.iter(f"{_HP}autoNum"))
            if auto_nums:
                assert auto_nums[0].get("numType") == "PAGE"
                break
        else:
            pytest.fail("No header with page number found")

        doc.close()

    def test_page_number_invalid_position(self, minimal_hwpx):
        doc = HwpxDoc.open(minimal_hwpx)
        with pytest.raises(ValueError, match="position must be"):
            set_page_number(doc, position="invalid_pos")
        doc.close()

    def test_page_number_persists(self, minimal_hwpx, tmp_dir):
        doc = HwpxDoc.open(minimal_hwpx)
        set_page_number(doc, fmt="[{page}]")
        out_path = os.path.join(tmp_dir, "pn_save.hwpx")
        doc.save_to_path(out_path)
        doc.close()

        doc2 = HwpxDoc.open(out_path)
        _, root2 = doc2.sections[0]
        auto_nums = [
            an for an in root2.iter(f"{_HP}autoNum")
            if an.get("numType") == "PAGE"
        ]
        assert len(auto_nums) >= 1
        doc2.close()


# ---------------------------------------------------------------------------
# Memo tests
# ---------------------------------------------------------------------------

class TestAddMemo:
    def test_basic_memo(self, minimal_hwpx):
        """Add a memo anchored to existing text."""
        doc = HwpxDoc.open(minimal_hwpx)
        add_memo_to_doc(doc, "First paragraph", "Review this section")

        _, root = doc.sections[0]
        memos = list(root.iter(f"{_HP}memo"))
        assert len(memos) == 1

        texts = [t.text for t in memos[0].iter(f"{_HP}t") if t.text]
        assert any("Review this section" in t for t in texts)

        doc.close()

    def test_memo_with_author(self, minimal_hwpx):
        """Memo includes author attribute."""
        doc = HwpxDoc.open(minimal_hwpx)
        add_memo_to_doc(doc, "Second paragraph", "Comment", author="tester")

        _, root = doc.sections[0]
        memos = list(root.iter(f"{_HP}memo"))
        assert len(memos) == 1
        assert memos[0].get("author") == "tester"

        doc.close()

    def test_memo_no_author(self, minimal_hwpx):
        """Memo without author should not have author attribute."""
        doc = HwpxDoc.open(minimal_hwpx)
        add_memo_to_doc(doc, "First paragraph", "Note")

        _, root = doc.sections[0]
        memos = list(root.iter(f"{_HP}memo"))
        assert len(memos) == 1
        # author attribute should not be set when empty string
        assert memos[0].get("author") is None

        doc.close()

    def test_memo_anchor_not_found(self, minimal_hwpx):
        doc = HwpxDoc.open(minimal_hwpx)
        with pytest.raises(ValueError, match="not found"):
            add_memo_to_doc(doc, "nonexistent", "Memo")
        doc.close()

    def test_memo_persists(self, minimal_hwpx, tmp_dir):
        doc = HwpxDoc.open(minimal_hwpx)
        add_memo_to_doc(doc, "Third paragraph", "Saved memo")
        out_path = os.path.join(tmp_dir, "memo_save.hwpx")
        doc.save_to_path(out_path)
        doc.close()

        doc2 = HwpxDoc.open(out_path)
        _, root2 = doc2.sections[0]
        memos = list(root2.iter(f"{_HP}memo"))
        assert len(memos) == 1
        doc2.close()

    def test_memo_marks_dirty(self, minimal_hwpx):
        doc = HwpxDoc.open(minimal_hwpx)
        add_memo_to_doc(doc, "First paragraph", "Note")
        assert "Contents/section0.xml" in doc._dirty_sections
        doc.close()


# ---------------------------------------------------------------------------
# Integration: combining multiple elements
# ---------------------------------------------------------------------------

class TestCombinedElements:
    def test_footnote_and_header_together(self, minimal_hwpx, tmp_dir):
        """Add both a footnote and header, save, and verify both survive."""
        doc = HwpxDoc.open(minimal_hwpx)
        add_footnote(doc, "First paragraph", "A footnote")
        set_header(doc, "Document Title")

        out_path = os.path.join(tmp_dir, "combined.hwpx")
        doc.save_to_path(out_path)
        doc.close()

        doc2 = HwpxDoc.open(out_path)
        _, root2 = doc2.sections[0]
        assert len(list(root2.iter(f"{_HP}footNote"))) == 1
        assert len(list(root2.iter(f"{_HP}header"))) >= 1
        doc2.close()

    def test_header_footer_page_number(self, minimal_hwpx, tmp_dir):
        """Add header, footer, and page number together."""
        doc = HwpxDoc.open(minimal_hwpx)
        set_header(doc, "Chapter 1")
        set_footer(doc, "Confidential")
        set_page_number(doc, fmt="- {page} -")

        out_path = os.path.join(tmp_dir, "hdr_ftr_pn.hwpx")
        doc.save_to_path(out_path)
        doc.close()

        doc2 = HwpxDoc.open(out_path)
        _, root2 = doc2.sections[0]
        assert len(list(root2.iter(f"{_HP}header"))) >= 1
        # Footer count: one explicit + one from page number
        footers = list(root2.iter(f"{_HP}footer"))
        assert len(footers) >= 1
        # Page number autoNum
        page_nums = [
            an for an in root2.iter(f"{_HP}autoNum")
            if an.get("numType") == "PAGE"
        ]
        assert len(page_nums) >= 1
        doc2.close()

    def test_all_note_types(self, minimal_hwpx, tmp_dir):
        """Add footnote, endnote, and memo together."""
        doc = HwpxDoc.open(minimal_hwpx)
        add_footnote(doc, "First paragraph", "Footnote text")
        add_endnote(doc, "Second paragraph", "Endnote text")
        add_memo_to_doc(doc, "Third paragraph", "Memo text")

        out_path = os.path.join(tmp_dir, "all_notes.hwpx")
        doc.save_to_path(out_path)
        doc.close()

        doc2 = HwpxDoc.open(out_path)
        _, root2 = doc2.sections[0]
        assert len(list(root2.iter(f"{_HP}footNote"))) == 1
        assert len(list(root2.iter(f"{_HP}endNote"))) == 1
        assert len(list(root2.iter(f"{_HP}memo"))) == 1
        doc2.close()
