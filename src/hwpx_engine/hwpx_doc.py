"""Direct ZIP/XML HWPX document handler.

Replaces python-hwpx dependency with direct zipfile + lxml implementation.
Key improvement: save_to_path preserves all non-XML entries byte-for-byte,
preventing data loss that occurred with python-hwpx's re-serialization.
"""
from __future__ import annotations

import os
import zipfile
from pathlib import Path
from typing import List, Optional, Tuple

from lxml import etree

_HP_URI = 'http://www.hancom.co.kr/hwpml/2011/paragraph'
_HP = f'{{{_HP_URI}}}'


class HwpxTable:
    """Lightweight table wrapper over an lxml <hp:tbl> element."""

    def __init__(self, element: etree._Element, section_filename: str):
        self.element = element
        self._section = section_filename

    @property
    def row_count(self) -> int:
        """Count hp:tr children."""
        value = self.element.get('rowCnt')
        if value is not None and value.isdigit():
            return int(value)
        return len(self.element.findall(f'{_HP}tr'))

    @property
    def column_count(self) -> int:
        value = self.element.get('colCnt')
        if value is not None and value.isdigit():
            return int(value)
        first_row = self.element.find(f'{_HP}tr')
        if first_row is not None:
            return len(first_row.findall(f'{_HP}tc'))
        return 0

    def mark_dirty(self):
        """Mark the section containing this table as DOM-modified."""
        if hasattr(self, '_doc') and self._doc is not None:
            self._doc._dirty_sections.add(self._section)

    def set_cell_text(self, row: int, col: int, text: str) -> None:
        """Set text of cell at (row, col). Basic implementation for compatibility."""
        for tr in self.element.findall(f'{_HP}tr'):
            for tc in tr.findall(f'{_HP}tc'):
                addr = tc.find(f'{_HP}cellAddr')
                if addr is not None:
                    r = int(addr.get('rowAddr', '-1'))
                    c = int(addr.get('colAddr', '-1'))
                    if r == row and c == col:
                        sub = tc.find(f'{_HP}subList')
                        if sub is None:
                            continue
                        # Find first hp:t and set text
                        for t_el in sub.iter(f'{_HP}t'):
                            t_el.text = str(text) if text is not None else ''
                            return
        raise IndexError(f'Cell ({row}, {col}) not found in table')

    def get_cell_text(self, row: int, col: int) -> str:
        """Get text from cell at (row, col)."""
        for tr in self.element.findall(f'{_HP}tr'):
            for tc in tr.findall(f'{_HP}tc'):
                addr = tc.find(f'{_HP}cellAddr')
                if addr is not None:
                    r = int(addr.get('rowAddr', '-1'))
                    c = int(addr.get('colAddr', '-1'))
                    if r == row and c == col:
                        texts = []
                        for t_el in tc.iter(f'{_HP}t'):
                            if t_el.text:
                                texts.append(t_el.text)
                        return ''.join(texts)
        raise IndexError(f'Cell ({row}, {col}) not found in table')

    def merge_cells(self, r1: int, c1: int, r2: int, c2: int) -> None:
        """Merge cells from (r1,c1) to (r2,c2).

        Sets the anchor cell's span, and deactivates covered cells
        (size 0x0, span 1x1).
        """
        anchor_tc = None
        covered_tcs = []

        for tr in self.element.findall(f'{_HP}tr'):
            for tc in tr.findall(f'{_HP}tc'):
                addr = tc.find(f'{_HP}cellAddr')
                if addr is None:
                    continue
                r = int(addr.get('rowAddr', '-1'))
                c = int(addr.get('colAddr', '-1'))
                if r1 <= r <= r2 and c1 <= c <= c2:
                    if r == r1 and c == c1:
                        anchor_tc = tc
                    else:
                        covered_tcs.append(tc)

        if anchor_tc is None:
            return

        # Set anchor span
        span = anchor_tc.find(f'{_HP}cellSpan')
        if span is not None:
            span.set('rowSpan', str(r2 - r1 + 1))
            span.set('colSpan', str(c2 - c1 + 1))

        # Deactivate covered cells
        for tc in covered_tcs:
            span = tc.find(f'{_HP}cellSpan')
            if span is not None:
                span.set('rowSpan', '1')
                span.set('colSpan', '1')
            csz = tc.find(f'{_HP}cellSz')
            if csz is not None:
                csz.set('width', '0')
                csz.set('height', '0')


class _Paragraph:
    """Minimal paragraph wrapper for iteration compatibility."""

    def __init__(self, element: etree._Element, section_filename: str):
        self.element = element
        self._section = section_filename

    @property
    def text(self) -> str:
        """Concatenated text from all hp:t elements."""
        texts = []
        for t_el in self.element.iter(f'{_HP}t'):
            if t_el.text:
                texts.append(t_el.text)
        return ''.join(texts)

    @property
    def tables(self) -> List[HwpxTable]:
        """Return tables embedded within this paragraph's runs."""
        result = []
        for run in self.element.findall(f'{_HP}run'):
            for child in run:
                if child.tag == f'{_HP}tbl':
                    result.append(HwpxTable(child, self._section))
        return result


class _Run:
    """Minimal run wrapper for replace_text compatibility."""

    def __init__(self, element: etree._Element):
        self._element = element
        self._t_elements = element.findall(f'{_HP}t')

    @property
    def text(self) -> Optional[str]:
        """Concatenated text from all hp:t elements in this run."""
        texts = []
        for t_el in self._t_elements:
            if t_el.text:
                texts.append(t_el.text)
        return ''.join(texts) if texts else None

    def replace_text(self, old: str, new: str) -> None:
        """Replace text within this run's hp:t elements."""
        for t_el in self._t_elements:
            if t_el.text and old in t_el.text:
                t_el.text = t_el.text.replace(old, new)


class HwpxDoc:
    """Direct ZIP/XML hwpx document handler. No python-hwpx dependency.

    Reads an hwpx file as a ZIP archive, parses section XMLs with lxml,
    and provides a compatible API surface for HwpxEditor/TableEditor.
    """

    def __init__(self, source_path: str, zip_entries: dict, section_trees: dict,
                 dirty_sections: set = None):
        """Internal constructor. Use HwpxDoc.open() instead.

        Args:
            source_path: Original file path.
            zip_entries: {filename: bytes} for ALL entries in the ZIP.
            section_trees: {filename: lxml.etree._Element} for parsed section XMLs.
        """
        self._source_path = source_path
        self._zip_entries = zip_entries  # Raw bytes for every ZIP entry
        self._section_trees = section_trees  # Parsed lxml trees for section XMLs
        self._entry_order = list(zip_entries.keys())  # Preserve original ZIP order
        self._dirty_sections: set = dirty_sections or set()  # Sections modified via DOM
        self._closed = False

    @classmethod
    def open(cls, path) -> 'HwpxDoc':
        """Open hwpx file. Read ZIP, parse all section XMLs with lxml."""
        path = str(path)
        zip_entries = {}
        section_trees = {}
        entry_order = []

        with zipfile.ZipFile(path, 'r') as zf:
            for item in zf.infolist():
                data = zf.read(item.filename)
                zip_entries[item.filename] = data
                entry_order.append(item.filename)

                # Parse section XMLs
                if (item.filename.startswith('Contents/section')
                        and item.filename.endswith('.xml')):
                    root = etree.fromstring(data)
                    section_trees[item.filename] = root

        doc = cls(path, zip_entries, section_trees)
        doc._entry_order = entry_order
        return doc

    def close(self) -> None:
        """Release resources."""
        self._closed = True
        self._zip_entries.clear()
        self._section_trees.clear()

    def save_to_path(self, path) -> str:
        """Save document to path.

        Re-serializes modified section XMLs with lxml.etree.tostring().
        All other entries (images, bindata, mimetype, header, manifest, etc.)
        are copied byte-for-byte from the in-memory data, preserving
        everything exactly as loaded.
        """
        path = str(path)

        # Build updated entries: for sections that have lxml trees,
        # serialize them; for everything else, use original bytes.
        with zipfile.ZipFile(path, 'w') as zout:
            for filename in self._entry_order:
                if filename in self._dirty_sections:
                    # Section was modified via DOM — serialize from lxml tree
                    root = self._section_trees[filename]
                    data = etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding='UTF-8',
                    )
                else:
                    # Use original bytes (preserves exact formatting)
                    data = self._zip_entries[filename]

                if filename == 'mimetype':
                    zout.writestr(filename, data, compress_type=zipfile.ZIP_STORED)
                else:
                    zout.writestr(filename, data, compress_type=zipfile.ZIP_DEFLATED)

        return path

    def reload_from_path(self, path: str) -> None:
        """Reload document state from a (potentially modified) ZIP file.

        Used after ZIP-level operations (str_replace, etc.) modify the file
        directly, requiring the in-memory DOM to be refreshed.
        """
        path = str(path)
        zip_entries = {}
        section_trees = {}
        entry_order = []

        with zipfile.ZipFile(path, 'r') as zf:
            for item in zf.infolist():
                data = zf.read(item.filename)
                zip_entries[item.filename] = data
                entry_order.append(item.filename)

                if (item.filename.startswith('Contents/section')
                        and item.filename.endswith('.xml')):
                    root = etree.fromstring(data)
                    section_trees[item.filename] = root

        self._source_path = path
        self._zip_entries = zip_entries
        self._section_trees = section_trees
        self._entry_order = entry_order
        self._dirty_sections = set()  # Reset after reload
        self._closed = False

    @property
    def sections(self) -> List[Tuple[str, etree._Element]]:
        """List of (filename, lxml root element) for each section XML."""
        result = []
        for filename in self._entry_order:
            if filename in self._section_trees:
                result.append((filename, self._section_trees[filename]))
        return result

    @property
    def paragraphs(self) -> List[_Paragraph]:
        """All paragraphs across all sections."""
        result = []
        for filename, root in self.sections:
            for p_el in root.findall(f'{_HP}p'):
                result.append(_Paragraph(p_el, filename))
        return result

    def iter_runs(self):
        """Yield every run element across all sections (for replace_text)."""
        for filename, root in self.sections:
            for run_el in root.iter(f'{_HP}run'):
                yield _Run(run_el)

    def get_tables(self) -> List[HwpxTable]:
        """Find all hp:tbl elements across all sections."""
        result = []
        for para in self.paragraphs:
            for tbl in para.tables:
                tbl._doc = self  # Back-reference for dirty tracking
                result.append(tbl)
        return result

    def add_paragraph(self, text: str, char_pr_id_ref: str = '0',
                      para_pr_id_ref: str = '0') -> None:
        """Append a paragraph to the last section."""
        sections = self.sections
        if not sections:
            return
        filename, root = sections[-1]
        self._dirty_sections.add(filename)

        p_el = etree.SubElement(root, f'{_HP}p', {
            'id': '0',
            'paraPrIDRef': str(para_pr_id_ref),
            'styleIDRef': '0',
            'pageBreak': '0',
            'columnBreak': '0',
            'merged': '0',
        })
        run_el = etree.SubElement(p_el, f'{_HP}run', {
            'charPrIDRef': str(char_pr_id_ref),
        })
        t_el = etree.SubElement(run_el, f'{_HP}t')
        t_el.text = text

    def remove_paragraph(self, para: _Paragraph) -> None:
        """Remove a paragraph from its section."""
        for filename, root in self.sections:
            if filename == para._section:
                parent = para.element.getparent()
                if parent is not None:
                    parent.remove(para.element)
                return

    def export_text(self) -> str:
        """Extract all text from the document."""
        texts = []
        for para in self.paragraphs:
            t = para.text
            if t:
                texts.append(t)
        return '\n'.join(texts)

    def export_markdown(self) -> str:
        """Extract text in simple markdown format (one paragraph per line)."""
        return self.export_text()
