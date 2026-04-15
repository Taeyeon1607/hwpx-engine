"""Complex table creation and editing for HWPX documents.

Provides:
  - add_table_to_doc(): compatibility wrapper for quick table creation
  - TableEditor: read/write existing tables via direct lxml element access
    (add_row, delete_row, add_column, delete_column use direct XML manipulation)
"""
from __future__ import annotations

from copy import deepcopy
from typing import Any, List, Optional, TYPE_CHECKING

from hwpx_engine.hwpx_doc import HwpxDoc

if TYPE_CHECKING:
    from hwpx_engine.formatter import StyleMapper

_HP = '{http://www.hancom.co.kr/hwpml/2011/paragraph}'


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _get_tbl_element(table) -> Any:
    """Return the underlying <hp:tbl> lxml Element from a table wrapper."""
    return table.element


def _cell_info(tc):
    """Return (rowAddr, colAddr, rowSpan, colSpan) for a <hp:tc>."""
    addr = tc.find(f'{_HP}cellAddr')
    span = tc.find(f'{_HP}cellSpan')
    r = int(addr.get('rowAddr', '0')) if addr is not None else 0
    c = int(addr.get('colAddr', '0')) if addr is not None else 0
    rs = int(span.get('rowSpan', '1')) if span is not None else 1
    cs = int(span.get('colSpan', '1')) if span is not None else 1
    return r, c, rs, cs


def _is_deactivated(tc) -> bool:
    """Check if a cell is deactivated (hidden by a merge): size 0x0 and span 1x1."""
    span = tc.find(f'{_HP}cellSpan')
    csz = tc.find(f'{_HP}cellSz')
    if span is None or csz is None:
        return False
    rs = int(span.get('rowSpan', '1'))
    cs = int(span.get('colSpan', '1'))
    w = int(csz.get('width', '0'))
    h = int(csz.get('height', '0'))
    return rs == 1 and cs == 1 and w == 0 and h == 0


def _make_deactivated_cell(src_tc, row_idx: int, col_idx: int):
    """Create a deactivated (merge-hidden) cell cloned from src_tc."""
    new_tc = deepcopy(src_tc)
    addr = new_tc.find(f'{_HP}cellAddr')
    if addr is not None:
        addr.set('rowAddr', str(row_idx))
        addr.set('colAddr', str(col_idx))
    span = new_tc.find(f'{_HP}cellSpan')
    if span is not None:
        span.set('rowSpan', '1')
        span.set('colSpan', '1')
    csz = new_tc.find(f'{_HP}cellSz')
    if csz is not None:
        csz.set('width', '0')
        csz.set('height', '0')
    for t_el in new_tc.iter(f'{_HP}t'):
        t_el.text = ''
    return new_tc


def _clone_cell(src_tc, row_idx: int, col_idx: int):
    """Deep-copy a <hp:tc> element, clear its text, update its address."""
    new_tc = deepcopy(src_tc)
    addr = new_tc.find(f'{_HP}cellAddr')
    if addr is not None:
        addr.set('rowAddr', str(row_idx))
        addr.set('colAddr', str(col_idx))
    span = new_tc.find(f'{_HP}cellSpan')
    if span is not None:
        span.set('rowSpan', '1')
        span.set('colSpan', '1')
    # Restore normal size if source was deactivated
    csz = new_tc.find(f'{_HP}cellSz')
    if csz is not None:
        if int(csz.get('width', '0')) == 0 and int(csz.get('height', '0')) == 0:
            csz.set('height', '1200')
    for t_el in new_tc.iter(f'{_HP}t'):
        t_el.text = ''
    return new_tc


def _reindex_addresses(tbl_el) -> None:
    """Recompute rowAddr for every cell based on physical <hp:tr> position.

    colAddr is left unchanged — it is managed explicitly during column ops.
    """
    for row_idx, tr in enumerate(tbl_el.findall(f'{_HP}tr')):
        for tc in tr.findall(f'{_HP}tc'):
            addr = tc.find(f'{_HP}cellAddr')
            if addr is not None:
                addr.set('rowAddr', str(row_idx))


def _reindex_all(tbl_el) -> None:
    """Recompute both rowAddr and colAddr for every cell.

    Merge-aware: skips column positions covered by rowSpan from above,
    so cells in rows with partial coverage keep correct column addresses.
    """
    tr_list = tbl_el.findall(f'{_HP}tr')
    # coverage[row_idx] = set of col indices covered by rowSpan from above
    coverage = {}

    for row_idx, tr in enumerate(tr_list):
        covered = coverage.get(row_idx, set())
        col_cursor = 0

        for tc in tr.findall(f'{_HP}tc'):
            # Skip columns covered by merges from above
            while col_cursor in covered:
                col_cursor += 1

            addr = tc.find(f'{_HP}cellAddr')
            if addr is not None:
                addr.set('rowAddr', str(row_idx))
                addr.set('colAddr', str(col_cursor))

            span = tc.find(f'{_HP}cellSpan')
            rs = int(span.get('rowSpan', '1')) if span is not None else 1
            cs = int(span.get('colSpan', '1')) if span is not None else 1

            # Mark covered columns for future rows
            if rs > 1:
                for dr in range(1, rs):
                    future_row = row_idx + dr
                    if future_row not in coverage:
                        coverage[future_row] = set()
                    for dc in range(cs):
                        coverage[future_row].add(col_cursor + dc)

            col_cursor += cs


def _update_row_count(tbl_el) -> None:
    """Sync tbl@rowCnt with the actual number of <hp:tr>.

    Height is adjusted proportionally rather than recalculated from cellSz,
    because cellSz heights in skip-approach rows (where the first cell is
    absent due to merge) can return inflated values from the wrong cell.
    Bug-fix (v1.1.5): the old sum-of-cellSz approach caused table height to
    INCREASE after row deletion (e.g., 29944 → 69990), breaking rendering.
    """
    rows = tbl_el.findall(f'{_HP}tr')
    old_cnt = int(tbl_el.get('rowCnt', str(len(rows))))
    new_cnt = len(rows)
    tbl_el.set('rowCnt', str(new_cnt))
    # Proportionally adjust height (never recalculate from cellSz)
    sz = tbl_el.find(f'{_HP}sz')
    if sz is not None and old_cnt > 0 and new_cnt != old_cnt:
        old_h = int(sz.get('height', '0'))
        if old_h > 0:
            new_h = max(1, int(old_h * new_cnt / old_cnt))
            sz.set('height', str(new_h))


def _update_col_count(tbl_el, new_count: int) -> None:
    """Set tbl@colCnt."""
    tbl_el.set('colCnt', str(new_count))


def _build_merge_map(tbl_el) -> dict:
    """Build a map of (row, col) → (rowSpan, colSpan, anchor_tc) for merge anchors.

    Only cells with rowSpan > 1 or colSpan > 1 are included.
    """
    result = {}
    for tr in tbl_el.findall(f'{_HP}tr'):
        for tc in tr.findall(f'{_HP}tc'):
            r, c, rs, cs = _cell_info(tc)
            if rs > 1 or cs > 1:
                result[(r, c)] = (rs, cs, tc)
    return result


def _make_minimal_cell(tbl_el, row_idx: int, col_idx: int, deactivated: bool = False):
    """Create a minimal <hp:tc> from scratch."""
    new_tc = tbl_el.makeelement(f'{_HP}tc', {
        'name': '', 'header': '0', 'hasMargin': '0',
        'protect': '0', 'editable': '0', 'dirty': '0',
        'borderFillIDRef': tbl_el.get('borderFillIDRef', '2'),
    })
    addr = new_tc.makeelement(f'{_HP}cellAddr', {
        'colAddr': str(col_idx), 'rowAddr': str(row_idx),
    })
    new_tc.append(addr)
    span = new_tc.makeelement(f'{_HP}cellSpan', {'colSpan': '1', 'rowSpan': '1'})
    new_tc.append(span)
    w, h = ('0', '0') if deactivated else ('0', '1200')
    csz = new_tc.makeelement(f'{_HP}cellSz', {'width': w, 'height': h})
    new_tc.append(csz)
    sub = new_tc.makeelement(f'{_HP}subList', {
        'id': '', 'textDirection': 'HORIZONTAL', 'lineWrap': 'BREAK',
        'vertAlign': 'TOP', 'linkListIDRef': '0',
        'linkListNextIDRef': '0', 'textWidth': '0',
        'textHeight': '0', 'hasTextRef': '0', 'hasNumRef': '0',
    })
    new_tc.append(sub)
    p_el = sub.makeelement(f'{_HP}p', {
        'id': '0', 'paraPrIDRef': '0', 'styleIDRef': '0',
        'pageBreak': '0', 'columnBreak': '0', 'merged': '0',
    })
    sub.append(p_el)
    run = p_el.makeelement(f'{_HP}run', {'charPrIDRef': '0'})
    p_el.append(run)
    t_el = run.makeelement(f'{_HP}t', {})
    run.append(t_el)
    return new_tc


def _reactivate_cell(tr, col: int, mcs: int, anchor_tc):
    """Reactivate a deactivated cell in a row after a row-merge anchor is deleted.

    When a 2-row merge anchor is deleted, the deactivated cell(s) in the next
    row become normal cells, inheriting content from the deleted anchor.
    """
    for tc in tr.findall(f'{_HP}tc'):
        addr = tc.find(f'{_HP}cellAddr')
        if addr is not None and int(addr.get('colAddr', '0')) == col:
            if _is_deactivated(tc):
                # Copy the anchor's style and restore size
                tc.set('borderFillIDRef', anchor_tc.get('borderFillIDRef', '2'))
                csz = tc.find(f'{_HP}cellSz')
                anchor_csz = anchor_tc.find(f'{_HP}cellSz')
                if csz is not None and anchor_csz is not None:
                    csz.set('width', anchor_csz.get('width', '0'))
                    csz.set('height', anchor_csz.get('height', '1200'))
                span = tc.find(f'{_HP}cellSpan')
                if span is not None:
                    span.set('colSpan', str(mcs))
                    span.set('rowSpan', '1')
            break


def _transfer_anchor_down(tr, col: int, mcs: int, new_rowspan: int, anchor_tc):
    """Transfer merge anchor to the next row when the anchor row is deleted.

    The deactivated cell in the next row becomes the new merge anchor with
    the reduced rowSpan.
    """
    for tc in tr.findall(f'{_HP}tc'):
        addr = tc.find(f'{_HP}cellAddr')
        if addr is not None and int(addr.get('colAddr', '0')) == col:
            # Make this the new anchor
            tc.set('borderFillIDRef', anchor_tc.get('borderFillIDRef', '2'))
            csz = tc.find(f'{_HP}cellSz')
            anchor_csz = anchor_tc.find(f'{_HP}cellSz')
            if csz is not None and anchor_csz is not None:
                csz.set('width', anchor_csz.get('width', '0'))
                csz.set('height', anchor_csz.get('height', '1200'))
            span = tc.find(f'{_HP}cellSpan')
            if span is not None:
                span.set('rowSpan', str(new_rowspan))
                span.set('colSpan', str(mcs))
            # Copy content from old anchor
            old_sub = anchor_tc.find(f'{_HP}subList')
            new_sub = tc.find(f'{_HP}subList')
            if old_sub is not None and new_sub is not None:
                tc.remove(new_sub)
                tc.append(deepcopy(old_sub))
            break


def _reactivate_cell_col(tr, col: int, anchor_tc):
    """Reactivate a deactivated cell in a row after a col-merge anchor is deleted.

    anchor_tc is provided only for the first row (the actual anchor);
    for other rows it's None (those cells are plain deactivated cells).
    """
    for tc in tr.findall(f'{_HP}tc'):
        addr = tc.find(f'{_HP}cellAddr')
        if addr is not None and int(addr.get('colAddr', '0')) == col:
            if _is_deactivated(tc) or (anchor_tc is not None and tc is not anchor_tc):
                if anchor_tc is not None:
                    tc.set('borderFillIDRef', anchor_tc.get('borderFillIDRef', '2'))
                    anchor_csz = anchor_tc.find(f'{_HP}cellSz')
                    if anchor_csz is not None:
                        csz = tc.find(f'{_HP}cellSz')
                        if csz is not None:
                            csz.set('width', anchor_csz.get('width', '0'))
                            csz.set('height', anchor_csz.get('height', '1200'))
                span = tc.find(f'{_HP}cellSpan')
                if span is not None:
                    span.set('colSpan', '1')
                    span.set('rowSpan', '1')
            break


def _transfer_anchor_right(tr, col: int, new_colspan: int, anchor_tc):
    """Transfer merge anchor to the next column when the anchor column is deleted.

    anchor_tc is provided only for the first row; for other rows it's None.
    """
    for tc in tr.findall(f'{_HP}tc'):
        addr = tc.find(f'{_HP}cellAddr')
        if addr is not None and int(addr.get('colAddr', '0')) == col:
            if anchor_tc is not None:
                tc.set('borderFillIDRef', anchor_tc.get('borderFillIDRef', '2'))
                anchor_csz = anchor_tc.find(f'{_HP}cellSz')
                if anchor_csz is not None:
                    csz = tc.find(f'{_HP}cellSz')
                    if csz is not None:
                        csz.set('width', anchor_csz.get('width', '0'))
                        csz.set('height', anchor_csz.get('height', '1200'))
                old_sub = anchor_tc.find(f'{_HP}subList')
                new_sub = tc.find(f'{_HP}subList')
                if old_sub is not None and new_sub is not None:
                    tc.remove(new_sub)
                    tc.append(deepcopy(old_sub))
            span = tc.find(f'{_HP}cellSpan')
            if span is not None:
                span.set('colSpan', str(new_colspan))
            break


# ---------------------------------------------------------------------------
# Top-level convenience
# ---------------------------------------------------------------------------

def add_table_to_doc(
    doc: HwpxDoc,
    headers: List[str],
    rows: List[List[str]],
    merges: Optional[List[Any]] = None,
    style_mapper: Optional["StyleMapper"] = None,
    width: Optional[int] = None,
) -> Any:
    """Add a table with headers, data rows, and optional cell merges.

    NOTE: This function is a legacy convenience wrapper. For new table creation,
    prefer insert_table() in HwpxEditor which uses ZIP-level injection with
    full style control via xml_primitives.

    Parameters
    ----------
    doc : HwpxDoc
        The target document.
    headers : list[str]
        Column header texts.
    rows : list[list[str]]
        Data rows (each a list of cell texts).
    merges : list[tuple], optional
        Each tuple is ``(start_row, start_col, end_row, end_col)``.
    style_mapper : StyleMapper, optional
        Style mapper for header/cell formatting.
    width : int, optional
        Table width override.
    """
    if not headers:
        return None

    from lxml import etree

    col_count = len(headers)
    row_count = 1 + len(rows)

    # Build table element directly with lxml
    table_width = width if width is not None else col_count * 7200
    table_height = row_count * 1200
    col_w = table_width // col_count

    tbl = etree.Element(f'{_HP}tbl', {
        'id': '0', 'rowCnt': str(row_count), 'colCnt': str(col_count),
        'cellSpacing': '0', 'borderFillIDRef': '2', 'noAdjust': '0',
    })
    etree.SubElement(tbl, f'{_HP}sz', {
        'width': str(table_width), 'widthRelTo': 'ABSOLUTE',
        'height': str(table_height), 'heightRelTo': 'ABSOLUTE', 'protect': '0',
    })

    for row_idx in range(row_count):
        tr = etree.SubElement(tbl, f'{_HP}tr')
        for col_idx in range(col_count):
            tc = etree.SubElement(tr, f'{_HP}tc', {
                'name': '', 'header': '1' if row_idx == 0 else '0',
                'hasMargin': '0', 'protect': '0', 'editable': '0',
                'dirty': '0', 'borderFillIDRef': '2',
            })
            etree.SubElement(tc, f'{_HP}cellAddr', {
                'colAddr': str(col_idx), 'rowAddr': str(row_idx),
            })
            etree.SubElement(tc, f'{_HP}cellSpan', {'colSpan': '1', 'rowSpan': '1'})
            etree.SubElement(tc, f'{_HP}cellSz', {
                'width': str(col_w), 'height': '1200',
            })
            sub = etree.SubElement(tc, f'{_HP}subList', {
                'id': '', 'textDirection': 'HORIZONTAL', 'lineWrap': 'BREAK',
                'vertAlign': 'TOP', 'linkListIDRef': '0',
                'linkListNextIDRef': '0', 'textWidth': '0',
                'textHeight': '0', 'hasTextRef': '0', 'hasNumRef': '0',
            })
            p_el = etree.SubElement(sub, f'{_HP}p', {
                'id': '0', 'paraPrIDRef': '0', 'styleIDRef': '0',
            })
            run = etree.SubElement(p_el, f'{_HP}run', {'charPrIDRef': '0'})
            t_el = etree.SubElement(run, f'{_HP}t')

            if row_idx == 0:
                t_el.text = headers[col_idx]
            elif col_idx < len(rows[row_idx - 1]):
                t_el.text = str(rows[row_idx - 1][col_idx])

    # Append to last section's last paragraph as a run > tbl child
    sections = doc.sections
    if sections:
        filename, root = sections[-1]
        # Create a wrapper paragraph with run containing the table
        p_el = etree.SubElement(root, f'{_HP}p', {
            'id': '0', 'paraPrIDRef': '0', 'styleIDRef': '0',
        })
        run = etree.SubElement(p_el, f'{_HP}run', {'charPrIDRef': '0'})
        run.append(tbl)

    from hwpx_engine.hwpx_doc import HwpxTable
    table = HwpxTable(tbl, sections[-1][0] if sections else '')

    # Apply merges
    if merges:
        for r1, c1, r2, c2 in merges:
            try:
                table.merge_cells(r1, c1, r2, c2)
            except Exception:
                pass

    return table


# ---------------------------------------------------------------------------
# TableEditor
# ---------------------------------------------------------------------------

class TableEditor:
    """Read and edit existing tables in a document.

    All operations manipulate the underlying lxml Element directly.
    """

    def __init__(self, doc: HwpxDoc) -> None:
        self._doc = doc

    @property
    def tables(self) -> list:
        """Collect all tables from all paragraphs."""
        return self._doc.get_tables()

    # --- Simple delegations ---

    def set_cell(self, table_index: int, row: int, col: int, text: str) -> None:
        """Set the text of a specific cell.

        Fully replaces all cell content:
          1. Gets the raw lxml <hp:tc> element
          2. Finds the <hp:subList>
          3. Removes ALL existing <hp:p> children
          4. Creates one fresh <hp:p> with <hp:run> and <hp:t> containing the text
          5. Uses lxml.etree exclusively
        """
        from lxml import etree

        tables = self.tables
        if table_index >= len(tables):
            raise IndexError(f'table_index {table_index} out of range (have {len(tables)} tables)')

        table = tables[table_index]
        tbl_el = _get_tbl_element(table)

        # Find the target cell by walking rows/cells and matching addresses
        target_tc = None
        for tr in tbl_el.findall(f'{_HP}tr'):
            for tc in tr.findall(f'{_HP}tc'):
                addr = tc.find(f'{_HP}cellAddr')
                if addr is not None:
                    r = int(addr.get('rowAddr', '-1'))
                    c = int(addr.get('colAddr', '-1'))
                    if r == row and c == col:
                        target_tc = tc
                        break
            if target_tc is not None:
                break

        if target_tc is None:
            raise IndexError(f'Cell ({row}, {col}) not found in table {table_index}')

        # Find or create the subList
        sub = target_tc.find(f'{_HP}subList')
        if sub is None:
            sub = etree.SubElement(target_tc, f'{_HP}subList', {
                'id': '', 'textDirection': 'HORIZONTAL', 'lineWrap': 'BREAK',
                'vertAlign': 'TOP', 'linkListIDRef': '0',
                'linkListNextIDRef': '0', 'textWidth': '0',
                'textHeight': '0', 'hasTextRef': '0', 'hasNumRef': '0',
            })

        # Strategy: keep FIRST paragraph structure intact, just replace text.
        # Remove EXTRA paragraphs (2nd, 3rd, etc.) but keep the first one.
        # This preserves linesegarray, style refs, and other formatting.
        old_paras = sub.findall(f'{_HP}p')
        if old_paras:
            # Keep first paragraph, remove extras
            for old_p in old_paras[1:]:
                sub.remove(old_p)
            first_p = old_paras[0]
            # Find the first hp:t in the first hp:run and set text
            first_t = None
            for run in first_p.findall(f'{_HP}run'):
                t_el = run.find(f'{_HP}t')
                if t_el is not None:
                    if first_t is None:
                        first_t = t_el
                    else:
                        # Remove extra runs (keeps only first run)
                        first_p.remove(run)
            if first_t is not None:
                first_t.text = str(text) if text is not None else ''
            else:
                # No existing hp:t, create one
                char_pr_ref = '0'
                first_run = first_p.find(f'{_HP}run')
                if first_run is not None:
                    char_pr_ref = first_run.get('charPrIDRef', '0')
                else:
                    first_run = etree.SubElement(first_p, f'{_HP}run', {'charPrIDRef': '0'})
                new_t = etree.SubElement(first_run, f'{_HP}t')
                new_t.text = str(text) if text is not None else ''
        else:
            # No existing paragraphs — create minimal structure
            new_p = etree.SubElement(sub, f'{_HP}p', {
                'id': '0', 'paraPrIDRef': '0', 'styleIDRef': '0',
                'pageBreak': '0', 'columnBreak': '0', 'merged': '0',
            })
            new_run = etree.SubElement(new_p, f'{_HP}run', {'charPrIDRef': '0'})
            new_t = etree.SubElement(new_run, f'{_HP}t')
            new_t.text = str(text) if text is not None else ''

        # Sync formulaScript LastResult if present
        formula = target_tc.find(f'{_HP}formulaScript')
        if formula is not None:
            lr = formula.find(f'{_HP}stringParam[@name="LastResult"]')
            if lr is not None:
                lr.text = str(text) if text is not None else ''

        # Mark dirty
        target_tc.set('dirty', '1')
        table.mark_dirty()

    def merge(self, table_index: int, r1: int, c1: int, r2: int, c2: int) -> None:
        """Merge cells in a table."""
        tables = self.tables
        if table_index < len(tables):
            tables[table_index].merge_cells(r1, c1, r2, c2)

    # --- Structural: rows ---

    def add_row(
        self,
        table_index: int,
        cells: List[str],
        position: Optional[int] = None,
    ) -> None:
        """Add a row to a table. Merge-aware.

        If a merged cell spans across the insertion point, its rowSpan is
        increased by 1 and a deactivated (hidden) cell is placed in the new
        row for that column.

        Parameters
        ----------
        table_index : int
            Index of the target table (0-based).
        cells : list[str]
            Text for each cell in the new row. Length should match column count.
        position : int, optional
            Row index at which to insert. None or -1 = append at end.
        """
        tables = self.tables
        if table_index >= len(tables):
            raise IndexError(f'table_index {table_index} out of range (have {len(tables)} tables)')

        table = tables[table_index]
        tbl_el = _get_tbl_element(table)
        tr_list = tbl_el.findall(f'{_HP}tr')
        col_count = int(tbl_el.get('colCnt', '0'))

        if position is None or position < 0 or position >= len(tr_list):
            insert_idx = len(tr_list)  # append
        else:
            insert_idx = position

        # Build merge map BEFORE insertion to find spans that straddle insert_idx
        merge_map = _build_merge_map(tbl_el)

        # Columns covered by a merge that straddles the insertion point
        # (anchor row < insert_idx AND anchor row + rowSpan > insert_idx)
        covered_cols = set()
        for (mr, mc), (mrs, mcs, anchor_tc) in merge_map.items():
            if mrs > 1 and mr < insert_idx and mr + mrs > insert_idx:
                # This merge straddles — expand rowSpan
                span_el = anchor_tc.find(f'{_HP}cellSpan')
                if span_el is not None:
                    span_el.set('rowSpan', str(mrs + 1))
                for dc in range(mcs):
                    covered_cols.add(mc + dc)

        # Reference row for cloning style — prefer the row at insert_idx or the last row
        if insert_idx < len(tr_list):
            ref_tr = tr_list[insert_idx]
        else:
            ref_tr = tr_list[-1] if tr_list else None
        ref_cells_by_col = {}
        if ref_tr is not None:
            for tc in ref_tr.findall(f'{_HP}tc'):
                addr = tc.find(f'{_HP}cellAddr')
                if addr is not None:
                    ref_cells_by_col[int(addr.get('colAddr', '0'))] = tc

        # Build new <hp:tr>
        new_tr = tbl_el.makeelement(f'{_HP}tr', {})
        for col_idx in range(col_count):
            ref_tc = ref_cells_by_col.get(col_idx)
            if ref_tc is None and ref_cells_by_col:
                ref_tc = list(ref_cells_by_col.values())[-1]

            if col_idx in covered_cols:
                # This column is covered by a straddling merge → deactivated cell
                if ref_tc is not None:
                    new_tc = _make_deactivated_cell(ref_tc, insert_idx, col_idx)
                else:
                    new_tc = _make_minimal_cell(tbl_el, insert_idx, col_idx, deactivated=True)
            else:
                cell_text = cells[col_idx] if col_idx < len(cells) else ''
                if ref_tc is not None:
                    new_tc = _clone_cell(ref_tc, insert_idx, col_idx)
                else:
                    new_tc = _make_minimal_cell(tbl_el, insert_idx, col_idx)
                t_nodes = list(new_tc.iter(f'{_HP}t'))
                if t_nodes:
                    t_nodes[0].text = cell_text

            new_tr.append(new_tc)

        # Insert into tbl element
        if insert_idx >= len(tr_list):
            tbl_el.append(new_tr)
        else:
            tr_list[insert_idx].addprevious(new_tr)

        _reindex_addresses(tbl_el)
        _update_row_count(tbl_el)
        table.mark_dirty()

    def delete_row(self, table_index: int, row_index: int) -> None:
        """Delete a row from a table. Merge-aware.

        If a merged cell spans across the deleted row, its rowSpan is decreased
        by 1. If the merge anchor is on the deleted row, the anchor moves to
        the next row.

        Parameters
        ----------
        table_index : int
            Index of the target table (0-based).
        row_index : int
            Row index to delete (0-based).
        """
        tables = self.tables
        if table_index >= len(tables):
            raise IndexError(f'table_index {table_index} out of range')

        table = tables[table_index]
        tbl_el = _get_tbl_element(table)
        tr_list = tbl_el.findall(f'{_HP}tr')

        if row_index < 0 or row_index >= len(tr_list):
            raise IndexError(f'row_index {row_index} out of range (have {len(tr_list)} rows)')

        if len(tr_list) <= 1:
            raise ValueError('Cannot delete the only remaining row')

        # Handle merges that straddle the deleted row
        merge_map = _build_merge_map(tbl_el)
        for (mr, mc), (mrs, mcs, anchor_tc) in merge_map.items():
            if mrs <= 1:
                continue
            if mr <= row_index < mr + mrs:
                if mrs == 2 and mr == row_index:
                    # Anchor is on deleted row, span=2 → next row cell becomes
                    # normal (span=1). The anchor cell itself will be removed
                    # with the row. Find the deactivated cell in the next row
                    # and reactivate it.
                    next_tr = tr_list[row_index + 1] if row_index + 1 < len(tr_list) else None
                    if next_tr is not None:
                        _reactivate_cell(next_tr, mc, mcs, anchor_tc)
                elif mr == row_index:
                    # Anchor is on deleted row, span > 2 → move anchor to next row
                    # and reduce rowSpan by 1
                    next_tr = tr_list[row_index + 1] if row_index + 1 < len(tr_list) else None
                    if next_tr is not None:
                        _transfer_anchor_down(next_tr, mc, mcs, mrs - 1, anchor_tc)
                else:
                    # Anchor is above deleted row → just reduce rowSpan
                    span_el = anchor_tc.find(f'{_HP}cellSpan')
                    if span_el is not None:
                        span_el.set('rowSpan', str(mrs - 1))

        tbl_el.remove(tr_list[row_index])
        _reindex_addresses(tbl_el)
        _update_row_count(tbl_el)
        table.mark_dirty()

    # --- Structural: columns ---

    def add_column(
        self,
        table_index: int,
        cells: List[str],
        position: Optional[int] = None,
        width: Optional[int] = None,
    ) -> None:
        """Add a column to a table. Merge-aware.

        If a merged cell spans across the insertion point, its colSpan is
        increased by 1 and a deactivated cell is placed in that row.

        Parameters
        ----------
        table_index : int
            Index of the target table (0-based).
        cells : list[str]
            Text for each cell in the new column. Length should match row count.
        position : int, optional
            Column index at which to insert. None or -1 = append at end.
        width : int, optional
            Width for the new column in HWPML units. If omitted, the table's
            total width is divided equally among all columns.
        """
        tables = self.tables
        if table_index >= len(tables):
            raise IndexError(f'table_index {table_index} out of range')

        table = tables[table_index]
        tbl_el = _get_tbl_element(table)
        tr_list = tbl_el.findall(f'{_HP}tr')
        old_col_count = int(tbl_el.get('colCnt', '0'))
        new_col_count = old_col_count + 1

        if position is None or position < 0 or position >= old_col_count:
            insert_col = old_col_count  # append
        else:
            insert_col = position

        # Determine column width
        sz = tbl_el.find(f'{_HP}sz')
        table_width = int(sz.get('width', '0')) if sz is not None else 0
        if width is not None:
            col_width = width
        elif table_width > 0:
            col_width = table_width // new_col_count
        else:
            col_width = 7200

        # Build merge map to find colSpans that straddle insert_col
        merge_map = _build_merge_map(tbl_el)
        # covered_rows: rows where the new column cell should be deactivated
        covered_rows = set()
        for (mr, mc), (mrs, mcs, anchor_tc) in merge_map.items():
            if mcs > 1 and mc < insert_col and mc + mcs > insert_col:
                # This merge straddles — expand colSpan
                span_el = anchor_tc.find(f'{_HP}cellSpan')
                if span_el is not None:
                    span_el.set('colSpan', str(mcs + 1))
                for dr in range(mrs):
                    covered_rows.add(mr + dr)

        for row_idx, tr in enumerate(tr_list):
            existing_tcs = tr.findall(f'{_HP}tc')
            cell_text = cells[row_idx] if row_idx < len(cells) else ''

            # Find a reference cell for style cloning
            ref_tc = None
            if existing_tcs:
                ref_idx = min(insert_col, len(existing_tcs) - 1)
                ref_tc = existing_tcs[ref_idx]

            if row_idx in covered_rows:
                # Deactivated cell for merge continuation
                if ref_tc is not None:
                    new_tc = _make_deactivated_cell(ref_tc, row_idx, insert_col)
                else:
                    new_tc = _make_minimal_cell(tbl_el, row_idx, insert_col, deactivated=True)
            else:
                if ref_tc is not None:
                    new_tc = _clone_cell(ref_tc, row_idx, insert_col)
                else:
                    new_tc = _make_minimal_cell(tbl_el, row_idx, insert_col)
                csz = new_tc.find(f'{_HP}cellSz')
                if csz is not None:
                    csz.set('width', str(col_width))
                t_nodes = list(new_tc.iter(f'{_HP}t'))
                if t_nodes:
                    t_nodes[0].text = cell_text

            # Insert at correct position
            if insert_col < len(existing_tcs):
                existing_tcs[insert_col].addprevious(new_tc)
            else:
                tr.append(new_tc)

        # Reindex all addresses
        _reindex_all(tbl_el)

        # Scale existing columns proportionally to make room for new column
        if width is None and table_width > 0:
            scale = (table_width - col_width) / table_width
            for tr in tr_list:
                for tc in tr.findall(f'{_HP}tc'):
                    if not _is_deactivated(tc):
                        addr = tc.find(f'{_HP}cellAddr')
                        csz = tc.find(f'{_HP}cellSz')
                        if addr is not None and csz is not None:
                            ca = int(addr.get('colAddr', '-1'))
                            if ca != insert_col:
                                old_w = int(csz.get('width', '0'))
                                csz.set('width', str(max(1, int(old_w * scale))))

        # Update last-column borderFill pattern
        if insert_col == old_col_count and old_col_count > 0:
            for tr in tr_list:
                tcs = tr.findall(f'{_HP}tc')
                if len(tcs) >= 2:
                    old_last = tcs[-2]
                    new_last = tcs[-1]
                    if not _is_deactivated(new_last):
                        old_bf = old_last.get('borderFillIDRef', '')
                        new_bf = new_last.get('borderFillIDRef', '')
                        if old_bf != new_bf and len(tcs) >= 3:
                            inner_bf = tcs[-3].get('borderFillIDRef', old_bf)
                            old_last.set('borderFillIDRef', inner_bf)

        _update_col_count(tbl_el, new_col_count)
        table.mark_dirty()

    def delete_column(self, table_index: int, col_index: int) -> None:
        """Delete a column from a table. Merge-aware.

        If a merged cell spans across the deleted column, its colSpan is
        decreased by 1. If the merge anchor is on the deleted column, the
        anchor moves to the next column.

        Parameters
        ----------
        table_index : int
            Index of the target table (0-based).
        col_index : int
            Column index to delete (0-based).
        """
        tables = self.tables
        if table_index >= len(tables):
            raise IndexError(f'table_index {table_index} out of range')

        table = tables[table_index]
        tbl_el = _get_tbl_element(table)
        tr_list = tbl_el.findall(f'{_HP}tr')
        old_col_count = int(tbl_el.get('colCnt', '0'))

        if col_index < 0 or col_index >= old_col_count:
            raise IndexError(f'col_index {col_index} out of range (have {old_col_count} columns)')

        if old_col_count <= 1:
            raise ValueError('Cannot delete the only remaining column')

        # Handle merges that straddle the deleted column
        merge_map = _build_merge_map(tbl_el)
        # Track rows where the cell at col_index should NOT be removed (part of merge)
        merge_handled_rows = set()
        for (mr, mc), (mrs, mcs, anchor_tc) in merge_map.items():
            if mcs <= 1:
                continue
            if mc <= col_index < mc + mcs:
                if mcs == 2 and mc == col_index:
                    # Anchor is on deleted col, span=2 → next col cell becomes normal
                    for dr in range(mrs):
                        row_i = mr + dr
                        if row_i < len(tr_list):
                            _reactivate_cell_col(tr_list[row_i], col_index + 1, anchor_tc if dr == 0 else None)
                        merge_handled_rows.add(row_i)
                elif mc == col_index:
                    # Anchor is on deleted col, span > 2 → move anchor right, reduce colSpan
                    for dr in range(mrs):
                        row_i = mr + dr
                        if row_i < len(tr_list):
                            _transfer_anchor_right(tr_list[row_i], col_index + 1, mcs - 1, anchor_tc if dr == 0 else None)
                        merge_handled_rows.add(row_i)
                else:
                    # Anchor is to the left → just reduce colSpan
                    span_el = anchor_tc.find(f'{_HP}cellSpan')
                    if span_el is not None:
                        span_el.set('colSpan', str(mcs - 1))
                    # The deactivated cell at col_index will be removed below
                    # but we don't add the row to merge_handled_rows so it gets removed

        # Save deleted column's width before removal
        sz = tbl_el.find(f'{_HP}sz')
        table_width = int(sz.get('width', '0')) if sz is not None else 0
        deleted_width = 0
        for tr in tr_list:
            for tc in tr.findall(f'{_HP}tc'):
                addr = tc.find(f'{_HP}cellAddr')
                if addr is not None and int(addr.get('colAddr', '-1')) == col_index:
                    if not _is_deactivated(tc):
                        csz = tc.find(f'{_HP}cellSz')
                        if csz is not None:
                            deleted_width = max(deleted_width, int(csz.get('width', '0')))
                    break

        # If deleting last column, transfer borderFill
        is_last_col = (col_index == old_col_count - 1)

        for row_idx, tr in enumerate(tr_list):
            tcs = tr.findall(f'{_HP}tc')
            target_tc = None
            for tc in tcs:
                addr = tc.find(f'{_HP}cellAddr')
                if addr is not None and int(addr.get('colAddr', '-1')) == col_index:
                    target_tc = tc
                    break

            if target_tc is None:
                continue

            if is_last_col and len(tcs) >= 2:
                bf = target_tc.get('borderFillIDRef', '')
                prev_tc = None
                for tc in tcs:
                    addr = tc.find(f'{_HP}cellAddr')
                    if addr is not None and int(addr.get('colAddr', '-1')) == col_index - 1:
                        prev_tc = tc
                        break
                if prev_tc is not None and bf:
                    prev_tc.set('borderFillIDRef', bf)

            tr.remove(target_tc)

        _reindex_all(tbl_el)

        # Redistribute freed width proportionally to remaining columns
        if deleted_width > 0 and table_width > deleted_width:
            scale = table_width / (table_width - deleted_width)
            for tr in tr_list:
                for tc in tr.findall(f'{_HP}tc'):
                    if not _is_deactivated(tc):
                        csz = tc.find(f'{_HP}cellSz')
                        if csz is not None:
                            old_w = int(csz.get('width', '0'))
                            csz.set('width', str(max(1, int(old_w * scale))))

        _update_col_count(tbl_el, old_col_count - 1)
        table.mark_dirty()

    def get_cell(self, table_index: int, row: int, col: int) -> str:
        """Return the text content of a single cell.

        Deactivated (merged-away) cells return an empty string.
        Raises IndexError if table_index, row, or col is out of range.
        """
        tables = self.tables
        if table_index >= len(tables):
            raise IndexError(
                f"Table index {table_index} out of range (have {len(tables)} tables)"
            )
        tbl_el = _get_tbl_element(tables[table_index])
        trs = tbl_el.findall(f'{_HP}tr')
        if row >= len(trs):
            raise IndexError(f"Row {row} out of range (have {len(trs)} rows)")
        tcs = trs[row].findall(f'{_HP}tc')
        if col >= len(tcs):
            raise IndexError(f"Col {col} out of range (have {len(tcs)} cols)")
        tc = tcs[col]
        if _is_deactivated(tc):
            return ""
        texts = [t_el.text for t_el in tc.iter(f'{_HP}t') if t_el.text]
        return ''.join(texts)

    def get_table_data(self, table_index: int) -> list:
        """Extract a full table as a 2-D list of strings.

        Deactivated (merged-away) cells appear as empty strings.
        Raises IndexError if table_index is out of range.
        """
        tables = self.tables
        if table_index >= len(tables):
            raise IndexError(
                f"Table index {table_index} out of range (have {len(tables)} tables)"
            )
        tbl_el = _get_tbl_element(tables[table_index])
        result = []
        for tr in tbl_el.findall(f'{_HP}tr'):
            row_data = []
            for tc in tr.findall(f'{_HP}tc'):
                if _is_deactivated(tc):
                    row_data.append("")
                else:
                    texts = [t_el.text for t_el in tc.iter(f'{_HP}t') if t_el.text]
                    row_data.append(''.join(texts))
            result.append(row_data)
        return result

    def batch_set_cell(
        self,
        table_index: int,
        updates: list,
    ) -> None:
        """Set multiple cells in one table with a single DOM traversal.

        Parameters
        ----------
        table_index : int
            Index of the target table (0-based).
        updates : list[tuple[int, int, str]]
            List of ``(row, col, text)`` tuples.

        Raises
        ------
        IndexError
            If *table_index* is out of range or any *(row, col)* not found.
        """
        from lxml import etree

        tables = self.tables
        if table_index >= len(tables):
            raise IndexError(
                f'table_index {table_index} out of range (have {len(tables)} tables)'
            )

        table = tables[table_index]
        tbl_el = _get_tbl_element(table)

        # Build (rowAddr, colAddr) → tc map in a single traversal
        addr_map: dict = {}
        for tr in tbl_el.findall(f'{_HP}tr'):
            for tc in tr.findall(f'{_HP}tc'):
                addr = tc.find(f'{_HP}cellAddr')
                if addr is not None:
                    key = (int(addr.get('rowAddr', '-1')),
                           int(addr.get('colAddr', '-1')))
                    addr_map[key] = tc

        for row, col, text in updates:
            target_tc = addr_map.get((row, col))
            if target_tc is None:
                raise IndexError(
                    f'Cell ({row}, {col}) not found in table {table_index}'
                )

            # --- replicate set_cell() text-replacement logic ---
            sub = target_tc.find(f'{_HP}subList')
            if sub is None:
                sub = etree.SubElement(target_tc, f'{_HP}subList', {
                    'id': '', 'textDirection': 'HORIZONTAL',
                    'lineWrap': 'BREAK', 'vertAlign': 'TOP',
                    'linkListIDRef': '0', 'linkListNextIDRef': '0',
                    'textWidth': '0', 'textHeight': '0',
                    'hasTextRef': '0', 'hasNumRef': '0',
                })

            old_paras = sub.findall(f'{_HP}p')
            if old_paras:
                for old_p in old_paras[1:]:
                    sub.remove(old_p)
                first_p = old_paras[0]
                first_t = None
                for run in first_p.findall(f'{_HP}run'):
                    t_el = run.find(f'{_HP}t')
                    if t_el is not None:
                        if first_t is None:
                            first_t = t_el
                        else:
                            first_p.remove(run)
                if first_t is not None:
                    first_t.text = str(text) if text is not None else ''
                else:
                    first_run = first_p.find(f'{_HP}run')
                    if first_run is None:
                        first_run = etree.SubElement(
                            first_p, f'{_HP}run', {'charPrIDRef': '0'}
                        )
                    new_t = etree.SubElement(first_run, f'{_HP}t')
                    new_t.text = str(text) if text is not None else ''
            else:
                new_p = etree.SubElement(sub, f'{_HP}p', {
                    'id': '0', 'paraPrIDRef': '0', 'styleIDRef': '0',
                    'pageBreak': '0', 'columnBreak': '0', 'merged': '0',
                })
                new_run = etree.SubElement(new_p, f'{_HP}run', {'charPrIDRef': '0'})
                new_t = etree.SubElement(new_run, f'{_HP}t')
                new_t.text = str(text) if text is not None else ''

            # Sync formulaScript LastResult if present
            formula = target_tc.find(f'{_HP}formulaScript')
            if formula is not None:
                lr = formula.find(f'{_HP}stringParam[@name="LastResult"]')
                if lr is not None:
                    lr.text = str(text) if text is not None else ''

            target_tc.set('dirty', '1')

        table.mark_dirty()
