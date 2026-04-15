"""Pytest fixtures for hwpx-engine tests.

Generates minimal valid HWPX files (ZIP + XML) for use as test inputs.
Files are created in a temporary directory that is cleaned up after each test.
"""
from __future__ import annotations

import os
import shutil
import tempfile
import zipfile

import pytest
from lxml import etree

# ---------------------------------------------------------------------------
# XML namespace constants
# ---------------------------------------------------------------------------
_NS_HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
_NS_HS = "http://www.hancom.co.kr/hwpml/2011/section"
_NS_HC = "http://www.hancom.co.kr/hwpml/2011/core"
_NS_HH = "http://www.hancom.co.kr/hwpml/2011/head"
_NS_HPF = "http://www.hancom.co.kr/schema/2011/hpf"
_NS_OPF = "http://www.idpf.org/2007/opf/"

_NSMAP = {
    "hp": _NS_HP,
    "hs": _NS_HS,
    "hc": _NS_HC,
    "hh": _NS_HH,
    "hpf": _NS_HPF,
    "opf": _NS_OPF,
}

HP = f"{{{_NS_HP}}}"
HS = f"{{{_NS_HS}}}"
HC = f"{{{_NS_HC}}}"
HH = f"{{{_NS_HH}}}"
OPF = f"{{{_NS_OPF}}}"


# ---------------------------------------------------------------------------
# XML builder helpers
# ---------------------------------------------------------------------------

def _make_paragraph(text: str, para_id: str = "1", para_pr_id: str = "0",
                    style_id: str = "0", char_pr_id: str = "0") -> etree._Element:
    """Build a minimal <hp:p> element containing a single run with text."""
    p = etree.Element(f"{HP}p", {
        "id": para_id,
        "paraPrIDRef": para_pr_id,
        "styleIDRef": style_id,
        "pageBreak": "0",
        "columnBreak": "0",
        "merged": "0",
    })
    run = etree.SubElement(p, f"{HP}run", {"charPrIDRef": char_pr_id})
    t = etree.SubElement(run, f"{HP}t")
    t.text = text
    return p


def _make_cell(row_addr: int, col_addr: int, text: str,
               col_span: int = 1, row_span: int = 1,
               width: int = 3000, height: int = 1000,
               border_fill_id: str = "1") -> etree._Element:
    """Build a <hp:tc> element for a table cell."""
    tc = etree.Element(f"{HP}tc", {
        "name": "",
        "header": "0",
        "hasMargin": "0",
        "protect": "0",
        "editable": "0",
        "dirty": "0",
        "borderFillIDRef": border_fill_id,
    })
    etree.SubElement(tc, f"{HP}cellAddr", {
        "colAddr": str(col_addr),
        "rowAddr": str(row_addr),
    })
    etree.SubElement(tc, f"{HP}cellSpan", {
        "colSpan": str(col_span),
        "rowSpan": str(row_span),
    })
    etree.SubElement(tc, f"{HP}cellSz", {
        "width": str(width),
        "height": str(height),
    })
    etree.SubElement(tc, f"{HP}cellMargin", {
        "left": "141",
        "right": "141",
        "top": "141",
        "bottom": "141",
    })

    # Cell content lives inside hp:subList
    sub = etree.SubElement(tc, f"{HP}subList", {
        "id": "",
        "textDirection": "HORIZONTAL",
        "lineWrap": "BREAK",
        "vertAlign": "TOP",
        "linkListIDRef": "0",
        "linkListNextIDRef": "0",
        "textWidth": "0",
        "textHeight": "0",
        "hasTextRef": "0",
        "hasNumRef": "0",
    })
    cell_p = etree.SubElement(sub, f"{HP}p", {
        "id": "1",
        "paraPrIDRef": "0",
        "styleIDRef": "0",
        "pageBreak": "0",
        "columnBreak": "0",
        "merged": "0",
    })
    run = etree.SubElement(cell_p, f"{HP}run", {"charPrIDRef": "0"})
    t = etree.SubElement(run, f"{HP}t")
    t.text = text
    return tc


def _make_deactivated_cell(row_addr: int, col_addr: int,
                            border_fill_id: str = "1") -> etree._Element:
    """Build a deactivated (merged-away) <hp:tc> with width=0, height=0."""
    tc = etree.Element(f"{HP}tc", {
        "name": "",
        "header": "0",
        "hasMargin": "0",
        "protect": "0",
        "editable": "0",
        "dirty": "0",
        "borderFillIDRef": border_fill_id,
    })
    etree.SubElement(tc, f"{HP}cellAddr", {
        "colAddr": str(col_addr),
        "rowAddr": str(row_addr),
    })
    etree.SubElement(tc, f"{HP}cellSpan", {
        "colSpan": "1",
        "rowSpan": "1",
    })
    etree.SubElement(tc, f"{HP}cellSz", {
        "width": "0",
        "height": "0",
    })
    etree.SubElement(tc, f"{HP}cellMargin", {
        "left": "141",
        "right": "141",
        "top": "141",
        "bottom": "141",
    })
    sub = etree.SubElement(tc, f"{HP}subList", {
        "id": "",
        "textDirection": "HORIZONTAL",
        "lineWrap": "BREAK",
        "vertAlign": "TOP",
        "linkListIDRef": "0",
        "linkListNextIDRef": "0",
        "textWidth": "0",
        "textHeight": "0",
        "hasTextRef": "0",
        "hasNumRef": "0",
    })
    etree.SubElement(sub, f"{HP}p", {
        "id": "1",
        "paraPrIDRef": "0",
        "styleIDRef": "0",
        "pageBreak": "0",
        "columnBreak": "0",
        "merged": "0",
    })
    return tc


def _make_table(rows: list[list[etree._Element]],
                row_cnt: int, col_cnt: int) -> etree._Element:
    """Wrap a list-of-lists of <hp:tc> elements into a <hp:tbl>."""
    tbl = etree.Element(f"{HP}tbl", {
        "id": "1",
        "zOrder": "0",
        "numberingType": "TABLE",
        "textWrap": "TOP_AND_BOTTOM",
        "textFlow": "BOTH_SIDES",
        "lock": "0",
        "dropcapstyle": "None",
        "pageBreak": "CELL",
        "repeatHeader": "0",
        "rowCnt": str(row_cnt),
        "colCnt": str(col_cnt),
        "cellSpacing": "0",
        "borderFillIDRef": "1",
        "noAdjust": "0",
    })
    etree.SubElement(tbl, f"{HP}sz", {
        "width": str(col_cnt * 3000),
        "widthRelTo": "ABSOLUTE",
        "height": str(row_cnt * 1000),
        "heightRelTo": "ABSOLUTE",
        "protect": "0",
    })
    etree.SubElement(tbl, f"{HP}pos", {
        "treatAsChar": "1",
        "affectLSpacing": "0",
        "flowWithText": "1",
        "allowOverlap": "0",
        "holdAnchorAndSO": "0",
        "vertRelTo": "PARA",
        "horzRelTo": "PARA",
        "vertAlign": "TOP",
        "horzAlign": "LEFT",
        "vertOffset": "0",
        "horzOffset": "0",
    })
    for row_cells in rows:
        tr = etree.SubElement(tbl, f"{HP}tr")
        for tc in row_cells:
            tr.append(tc)
    return tbl


def _build_section_xml(paragraphs_text: list[str],
                        tables: list[etree._Element]) -> bytes:
    """Build section0.xml bytes.

    Paragraphs are emitted first, then each table is wrapped in a
    <hp:p><hp:run><hp:tbl>...</hp:tbl></hp:run></hp:p> structure.
    """
    root = etree.Element(f"{HS}sec", {
        "id": "0",
        "visibility": "SHOW",
        "numPageStartsAt": "0",
        "numLineStartsAt": "0",
        "outlineLevel": "0",
    }, nsmap={
        "hp": _NS_HP,
        "hs": _NS_HS,
        "hc": _NS_HC,
        "hh": _NS_HH,
    })

    for i, text in enumerate(paragraphs_text):
        p = _make_paragraph(text, para_id=str(i + 1))
        root.append(p)

    for tbl_el in tables:
        wrapper_p = etree.SubElement(root, f"{HP}p", {
            "id": "100",
            "paraPrIDRef": "0",
            "styleIDRef": "0",
            "pageBreak": "0",
            "columnBreak": "0",
            "merged": "0",
        })
        run = etree.SubElement(wrapper_p, f"{HP}run", {"charPrIDRef": "0"})
        run.append(tbl_el)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8")


def _build_header_xml() -> bytes:
    """Build a minimal Contents/header.xml."""
    root = etree.Element(f"{HH}head", {
        "version": "1.0",
        "secCnt": "1",
    }, nsmap={
        "hp": _NS_HP,
        "hs": _NS_HS,
        "hc": _NS_HC,
        "hh": _NS_HH,
    })

    etree.SubElement(root, f"{HH}beginNum", {
        "page": "1",
        "footnote": "1",
        "endnote": "1",
        "pic": "1",
        "tbl": "1",
        "equation": "1",
    })

    ref_list = etree.SubElement(root, f"{HH}refList")

    # borderFills (id=1 is referenced by table cells)
    border_fills = etree.SubElement(ref_list, f"{HH}borderFills", {"itemCnt": "1"})
    bf = etree.SubElement(border_fills, f"{HH}borderFill", {
        "id": "1",
        "threeD": "0",
        "shadow": "0",
        "centerLine": "0",
        "breakCellSeparateLine": "0",
    })
    for side in ("left", "right", "top", "bottom", "diagonal"):
        etree.SubElement(bf, f"{HH}border", {
            "type": "SOLID" if side != "diagonal" else "NONE",
            "width": "0.1mm",
            "color": "#000000",
        })

    # charProperties (id=0)
    char_props = etree.SubElement(ref_list, f"{HH}charProperties", {"itemCnt": "1"})
    etree.SubElement(char_props, f"{HH}charPr", {
        "id": "0",
        "height": "1000",
        "textColor": "#000000",
        "shadeColor": "#FFFFFF",
        "useFontSpace": "0",
        "useKerning": "0",
        "symMark": "NONE",
        "borderFillIDRef": "1",
    })

    # paraProperties (id=0)
    para_props = etree.SubElement(ref_list, f"{HH}paraProperties", {"itemCnt": "1"})
    etree.SubElement(para_props, f"{HH}paraPr", {
        "id": "0",
        "tabPrIDRef": "0",
        "condense": "0",
        "fontLineHeight": "0",
        "snapToGrid": "1",
        "suppressLineNumbers": "0",
        "checked": "0",
    })

    # styles (id=0, the default style)
    styles = etree.SubElement(ref_list, f"{HH}styles", {"itemCnt": "1"})
    etree.SubElement(styles, f"{HH}style", {
        "id": "0",
        "type": "PARA",
        "name": "바탕글",
        "engName": "Normal",
        "paraPrIDRef": "0",
        "charPrIDRef": "0",
        "nextStyleIDRef": "0",
        "langID": "1042",
        "lockForm": "0",
    })

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8")


def _build_content_hpf() -> bytes:
    """Build a minimal Contents/content.hpf manifest."""
    root = etree.Element(f"{OPF}package", {
        "version": "1.0",
        "unique-identifier": "",
        "id": "",
    }, nsmap={
        "opf": _NS_OPF,
        "hpf": _NS_HPF,
    })
    metadata = etree.SubElement(root, f"{OPF}metadata")
    etree.SubElement(metadata, f"{OPF}title").text = "Test Document"
    etree.SubElement(metadata, f"{OPF}language").text = "ko"

    manifest = etree.SubElement(root, f"{OPF}manifest")
    etree.SubElement(manifest, f"{OPF}item", {
        "id": "header",
        "href": "Contents/header.xml",
        "media-type": "application/xml",
    })
    etree.SubElement(manifest, f"{OPF}item", {
        "id": "section0",
        "href": "Contents/section0.xml",
        "media-type": "application/xml",
    })

    spine = etree.SubElement(root, f"{OPF}spine")
    etree.SubElement(spine, f"{OPF}itemref", {"idref": "section0"})

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8")


def _write_hwpx(path: str, section_xml: bytes) -> None:
    """Write a minimal HWPX ZIP to *path* using the provided section XML bytes."""
    header_xml = _build_header_xml()
    content_hpf = _build_content_hpf()

    with zipfile.ZipFile(path, "w") as zf:
        # mimetype must be first, stored uncompressed
        zf.writestr("mimetype", "application/hwp+zip",
                    compress_type=zipfile.ZIP_STORED)
        zf.writestr("Contents/header.xml", header_xml,
                    compress_type=zipfile.ZIP_DEFLATED)
        zf.writestr("Contents/content.hpf", content_hpf,
                    compress_type=zipfile.ZIP_DEFLATED)
        zf.writestr("Contents/section0.xml", section_xml,
                    compress_type=zipfile.ZIP_DEFLATED)


# ---------------------------------------------------------------------------
# Public helper functions (importable from conftest directly)
# ---------------------------------------------------------------------------

def _create_minimal_hwpx(path: str) -> str:
    """Create a minimal HWPX with 3 paragraphs and a 3×3 table at *path*."""
    paragraphs = ["First paragraph", "Second paragraph", "Third paragraph"]

    # 3x3 table: header row + 2 data rows
    header_row = [
        _make_cell(0, 0, "구분"),
        _make_cell(0, 1, "2027"),
        _make_cell(0, 2, "2028"),
    ]
    data_row1 = [
        _make_cell(1, 0, "항목A"),
        _make_cell(1, 1, "100"),
        _make_cell(1, 2, "200"),
    ]
    data_row2 = [
        _make_cell(2, 0, "항목B"),
        _make_cell(2, 1, "300"),
        _make_cell(2, 2, "400"),
    ]
    tbl = _make_table([header_row, data_row1, data_row2], row_cnt=3, col_cnt=3)

    section_xml = _build_section_xml(paragraphs, [tbl])
    _write_hwpx(path, section_xml)
    return path


def _create_merge_hwpx(path: str) -> str:
    """Create an HWPX with a merged-cell table at *path*.

    Table layout (4 rows × 3 cols):
      Row 0: "Header1"(colspan=2) | "Header2"
      Row 1: "A"(rowspan=2)       | "B" | "C"
      Row 2: (merged from A)      | "D" | "E"
      Row 3: "F"                  | "G" | "H"
    """
    # Row 0
    r0 = [
        _make_cell(0, 0, "Header1", col_span=2, row_span=1),  # anchor colspan=2
        _make_deactivated_cell(0, 1),                          # covered by Header1
        _make_cell(0, 2, "Header2"),
    ]
    # Row 1
    r1 = [
        _make_cell(1, 0, "A", col_span=1, row_span=2),         # anchor rowspan=2
        _make_cell(1, 1, "B"),
        _make_cell(1, 2, "C"),
    ]
    # Row 2: (0,0) is covered by "A"
    r2 = [
        _make_deactivated_cell(2, 0),                          # covered by A
        _make_cell(2, 1, "D"),
        _make_cell(2, 2, "E"),
    ]
    # Row 3
    r3 = [
        _make_cell(3, 0, "F"),
        _make_cell(3, 1, "G"),
        _make_cell(3, 2, "H"),
    ]

    tbl = _make_table([r0, r1, r2, r3], row_cnt=4, col_cnt=3)
    section_xml = _build_section_xml([], [tbl])
    _write_hwpx(path, section_xml)
    return path


def _create_ambiguous_hwpx(path: str) -> str:
    """Create an HWPX where 'cell_only_text' appears only inside a table cell.

    Used to verify that _find_toplevel_anchor skips cell-only matches and raises
    TextNotFoundError instead of inserting inside a table cell.
    """
    paragraphs = ["Top level paragraph"]

    # 1x1 table with cell text that is not a top-level paragraph
    data_row = [_make_cell(0, 0, "cell_only_text")]
    tbl = _make_table([data_row], row_cnt=1, col_cnt=1)

    section_xml = _build_section_xml(paragraphs, [tbl])
    _write_hwpx(path, section_xml)
    return path


# ---------------------------------------------------------------------------
# Pytest fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def tmp_dir():
    """Temporary directory, cleaned up after the test."""
    d = tempfile.mkdtemp(prefix="hwpx_test_")
    yield d
    shutil.rmtree(d, ignore_errors=True)


@pytest.fixture
def minimal_hwpx(tmp_dir):
    """Path to a minimal HWPX file with 3 paragraphs and a 3×3 table."""
    path = os.path.join(tmp_dir, "minimal.hwpx")
    _create_minimal_hwpx(path)
    return path


@pytest.fixture
def ambiguous_hwpx(tmp_dir):
    """Path to an HWPX where 'cell_only_text' appears only inside a table cell."""
    path = os.path.join(tmp_dir, "ambiguous.hwpx")
    _create_ambiguous_hwpx(path)
    return path


@pytest.fixture
def merge_hwpx(tmp_dir):
    """Path to an HWPX file containing a table with merged cells."""
    path = os.path.join(tmp_dir, "merge.hwpx")
    _create_merge_hwpx(path)
    return path
