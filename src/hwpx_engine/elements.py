"""Document elements: footnotes, endnotes, page numbers, memos, headers/footers.

Implemented by reverse-engineering actual HWPX files created in Hangul.
All functions manipulate section XML via lxml and mark dirty sections
so changes are persisted on save.

See docs/hwpx-elements-analysis.md for the XML structure reference.
"""
from __future__ import annotations

import random
from typing import Optional

from lxml import etree

from hwpx_engine.hwpx_doc import HwpxDoc

# ---------------------------------------------------------------------------
# Namespace constants
# ---------------------------------------------------------------------------
_HP_URI = "http://www.hancom.co.kr/hwpml/2011/paragraph"
_HP = f"{{{_HP_URI}}}"

# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

_inst_counter = 2_000_000_000


def _next_inst_id() -> str:
    """Generate a unique instId for note elements."""
    global _inst_counter
    _inst_counter += 1
    return str(_inst_counter)


def _make_sublist(text_width: str = "0", text_height: str = "0",
                  vert_align: str = "TOP") -> etree._Element:
    """Create a standard hp:subList container."""
    sub = etree.Element(f"{_HP}subList")
    sub.set("id", "")
    sub.set("textDirection", "HORIZONTAL")
    sub.set("lineWrap", "BREAK")
    sub.set("vertAlign", vert_align)
    sub.set("linkListIDRef", "0")
    sub.set("linkListNextIDRef", "0")
    sub.set("textWidth", text_width)
    sub.set("textHeight", text_height)
    sub.set("hasTextRef", "0")
    sub.set("hasNumRef", "0")
    return sub


def _make_linesegarray(horzsize: str = "42520") -> etree._Element:
    """Create a default linesegarray for a paragraph."""
    lsa = etree.SubElement(etree.Element("dummy"), f"{_HP}linesegarray")
    # Detach from dummy parent
    lsa = etree.Element(f"{_HP}linesegarray")
    ls = etree.SubElement(lsa, f"{_HP}lineseg")
    ls.set("textpos", "0")
    ls.set("vertpos", "0")
    ls.set("vertsize", "1000")
    ls.set("textheight", "1000")
    ls.set("baseline", "850")
    ls.set("spacing", "600")
    ls.set("horzpos", "0")
    ls.set("horzsize", horzsize)
    ls.set("flags", "393216")
    return lsa


def _make_note_paragraph(note_text: str, num: int, num_type: str,
                         para_pr: str = "0", style_ref: str = "0",
                         char_pr: str = "0") -> etree._Element:
    """Create a paragraph for inside a footnote/endnote subList.

    Contains autoNum control followed by the note text.
    """
    p = etree.Element(f"{_HP}p")
    p.set("id", "0")
    p.set("paraPrIDRef", para_pr)
    p.set("styleIDRef", style_ref)
    p.set("pageBreak", "0")
    p.set("columnBreak", "0")
    p.set("merged", "0")

    run = etree.SubElement(p, f"{_HP}run")
    run.set("charPrIDRef", char_pr)

    # autoNum control
    ctrl = etree.SubElement(run, f"{_HP}ctrl")
    auto_num = etree.SubElement(ctrl, f"{_HP}autoNum")
    auto_num.set("num", str(num))
    auto_num.set("numType", num_type)
    fmt = etree.SubElement(auto_num, f"{_HP}autoNumFormat")
    fmt.set("type", "DIGIT")
    fmt.set("userChar", "")
    fmt.set("prefixChar", "")
    fmt.set("suffixChar", ")")
    fmt.set("supscript", "0")

    # Note text (space before text matches Hangul convention)
    t = etree.SubElement(run, f"{_HP}t")
    t.text = f" {note_text}"

    p.append(_make_linesegarray())
    return p


def _find_run_containing_text(root: etree._Element, anchor_text: str):
    """Find the <hp:run> element whose <hp:t> children contain anchor_text.

    Searches only top-level paragraphs (not inside table cells, notes, etc.).
    Returns (run_element, hp_t_element) or (None, None).
    """
    for p in root.findall(f"{_HP}p"):
        for run in p.findall(f"{_HP}run"):
            for t_el in run.findall(f"{_HP}t"):
                if t_el.text and anchor_text in t_el.text:
                    return run, t_el
    return None, None


def _count_existing_notes(root: etree._Element, tag_local: str) -> int:
    """Count existing footnotes or endnotes in a section root."""
    count = 0
    for el in root.iter(f"{_HP}{tag_local}"):
        count += 1
    return count


def _get_text_width_from_section(root: etree._Element) -> str:
    """Extract text width from secPr > pagePr > margin for header/footer sizing."""
    sec_pr = root.find(f".//{_HP}secPr")
    if sec_pr is not None:
        page_pr = sec_pr.find(f"{_HP}pagePr")
        if page_pr is not None:
            width = int(page_pr.get("width", "59528"))
            margin = page_pr.find(f"{_HP}margin")
            if margin is not None:
                left = int(margin.get("left", "8504"))
                right = int(margin.get("right", "8504"))
                return str(width - left - right)
    return "42520"  # Default A4 text width


def _get_header_height_from_section(root: etree._Element) -> str:
    """Extract header margin height from secPr."""
    sec_pr = root.find(f".//{_HP}secPr")
    if sec_pr is not None:
        page_pr = sec_pr.find(f"{_HP}pagePr")
        if page_pr is not None:
            margin = page_pr.find(f"{_HP}margin")
            if margin is not None:
                return margin.get("header", "4252")
    return "4252"


def _get_footer_height_from_section(root: etree._Element) -> str:
    """Extract footer margin height from secPr."""
    sec_pr = root.find(f".//{_HP}secPr")
    if sec_pr is not None:
        page_pr = sec_pr.find(f"{_HP}pagePr")
        if page_pr is not None:
            margin = page_pr.find(f"{_HP}margin")
            if margin is not None:
                return margin.get("footer", "4252")
    return "4252"


def _find_secpr_run(root: etree._Element):
    """Find the <hp:run> that contains <hp:secPr> (first paragraph, first run).

    Headers and footers must be inserted in a run within the first paragraph
    that contains secPr. Returns (paragraph, run) or (None, None).
    """
    for p in root.findall(f"{_HP}p"):
        for run in p.findall(f"{_HP}run"):
            if run.find(f"{_HP}secPr") is not None:
                return p, run
    # Fallback: first paragraph, first run
    first_p = root.find(f"{_HP}p")
    if first_p is not None:
        first_run = first_p.find(f"{_HP}run")
        return first_p, first_run
    return None, None


def _max_header_footer_id(root: etree._Element) -> int:
    """Find the maximum id attribute across all headers and footers."""
    max_id = 0
    for tag in ("header", "footer"):
        for el in root.iter(f"{_HP}{tag}"):
            try:
                val = int(el.get("id", "0"))
                if val > max_id:
                    max_id = val
            except ValueError:
                pass
    return max_id


def _get_para_alignment(align: str) -> str:
    """Map alignment string to paraPrIDRef hint.

    Note: In a real document, alignment is controlled by paraPr definitions
    in header.xml. We use '0' as default since the actual alignment style
    must pre-exist in the document's header.xml charProperties.
    For simple use cases, the default style usually renders centered.
    """
    # We don't create new paraPr entries; use existing ones.
    # The caller can pass specific paraPrIDRef if needed.
    return "0"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def add_footnote(doc: HwpxDoc, anchor_text: str, note: str,
                 section_index: int = 0) -> None:
    """Add a footnote anchored to text containing *anchor_text*.

    The footnote marker is inserted after the anchor text within the same run.
    The footnote body appears at the bottom of the page.

    Args:
        doc: HwpxDoc instance (must be open).
        anchor_text: Text string to search for in the section.
        note: Footnote body text.
        section_index: Which section to search (default: 0 = first section).

    Raises:
        ValueError: If anchor_text is not found or section_index is invalid.
    """
    sections = doc.sections
    if section_index < 0 or section_index >= len(sections):
        raise ValueError(
            f"section_index {section_index} out of range "
            f"(document has {len(sections)} section(s))"
        )

    filename, root = sections[section_index]
    run, t_el = _find_run_containing_text(root, anchor_text)
    if run is None:
        raise ValueError(
            f"anchor_text {anchor_text!r} not found in section {section_index}"
        )

    # Determine footnote number
    num = _count_existing_notes(root, "footNote") + 1

    # Build footnote element
    ctrl = etree.Element(f"{_HP}ctrl")
    fn = etree.SubElement(ctrl, f"{_HP}footNote")
    fn.set("number", str(num))
    fn.set("suffixChar", "41")  # ')'
    fn.set("instId", _next_inst_id())

    sub = _make_sublist()
    fn.append(sub)
    note_para = _make_note_paragraph(note, num, "FOOTNOTE")
    sub.append(note_para)

    # Insert ctrl after the anchor text's <hp:t> element
    # Also add an empty <hp:t/> after ctrl (matches Hangul pattern)
    parent = t_el.getparent()
    t_index = list(parent).index(t_el)
    parent.insert(t_index + 1, ctrl)
    empty_t = etree.Element(f"{_HP}t")
    parent.insert(t_index + 2, empty_t)

    doc._dirty_sections.add(filename)


# Alias for backward compatibility
add_footnote_to_doc = add_footnote


def add_endnote(doc: HwpxDoc, anchor_text: str, note: str,
                section_index: int = 0) -> None:
    """Add an endnote anchored to text containing *anchor_text*.

    Endnotes appear at the end of the document (or section, depending on
    the document's endNotePr/placement setting).

    Args:
        doc: HwpxDoc instance (must be open).
        anchor_text: Text string to search for in the section.
        note: Endnote body text.
        section_index: Which section to search (default: 0 = first section).

    Raises:
        ValueError: If anchor_text is not found or section_index is invalid.
    """
    sections = doc.sections
    if section_index < 0 or section_index >= len(sections):
        raise ValueError(
            f"section_index {section_index} out of range "
            f"(document has {len(sections)} section(s))"
        )

    filename, root = sections[section_index]
    run, t_el = _find_run_containing_text(root, anchor_text)
    if run is None:
        raise ValueError(
            f"anchor_text {anchor_text!r} not found in section {section_index}"
        )

    # Determine endnote number
    num = _count_existing_notes(root, "endNote") + 1

    # Build endnote element
    ctrl = etree.Element(f"{_HP}ctrl")
    en = etree.SubElement(ctrl, f"{_HP}endNote")
    en.set("number", str(num))
    en.set("suffixChar", "41")  # ')'
    en.set("instId", _next_inst_id())

    sub = _make_sublist()
    en.append(sub)
    note_para = _make_note_paragraph(note, num, "ENDNOTE")
    sub.append(note_para)

    # Insert ctrl after the anchor text's <hp:t> element
    parent = t_el.getparent()
    t_index = list(parent).index(t_el)
    parent.insert(t_index + 1, ctrl)
    empty_t = etree.Element(f"{_HP}t")
    parent.insert(t_index + 2, empty_t)

    doc._dirty_sections.add(filename)


def set_header(doc: HwpxDoc, text: str, align: str = "center",
               apply_page_type: str = "BOTH",
               section_index: int = 0,
               para_pr_id_ref: str = "0",
               char_pr_id_ref: str = "0",
               style_id_ref: str = "0") -> None:
    """Set or replace a header in the specified section.

    If a header with the same applyPageType already exists, its text is replaced.
    Otherwise, a new header element is created.

    Args:
        doc: HwpxDoc instance (must be open).
        text: Header text content.
        align: Alignment hint (currently informational; actual alignment
               depends on the paraPr referenced by para_pr_id_ref).
        apply_page_type: "BOTH", "ODD", or "EVEN".
        section_index: Which section to modify (default: 0).
        para_pr_id_ref: paraPrIDRef for the header paragraph.
        char_pr_id_ref: charPrIDRef for the header run.
        style_id_ref: styleIDRef for the header paragraph.

    Raises:
        ValueError: If section_index is invalid.
    """
    sections = doc.sections
    if section_index < 0 or section_index >= len(sections):
        raise ValueError(
            f"section_index {section_index} out of range "
            f"(document has {len(sections)} section(s))"
        )

    filename, root = sections[section_index]
    text_width = _get_text_width_from_section(root)
    header_height = _get_header_height_from_section(root)

    # Check for existing header with same applyPageType
    for existing in root.iter(f"{_HP}header"):
        if existing.get("applyPageType") == apply_page_type:
            # Replace text in existing header
            sub = existing.find(f"{_HP}subList")
            if sub is not None:
                p = sub.find(f"{_HP}p")
                if p is not None:
                    run = p.find(f"{_HP}run")
                    if run is not None:
                        # Clear existing text elements
                        for old_t in run.findall(f"{_HP}t"):
                            run.remove(old_t)
                        t_el = etree.SubElement(run, f"{_HP}t")
                        t_el.text = text
                        doc._dirty_sections.add(filename)
                        return
            # subList exists but structure is odd -- rebuild paragraph
            break

    # Find the run containing secPr (or first paragraph's run)
    sec_p, sec_run = _find_secpr_run(root)
    if sec_run is None:
        raise ValueError(
            f"Cannot find a suitable run in section {section_index} "
            "to insert header"
        )

    # Generate unique ID
    new_id = _max_header_footer_id(root) + 1

    # Build header element
    ctrl = etree.Element(f"{_HP}ctrl")
    header = etree.SubElement(ctrl, f"{_HP}header")
    header.set("id", str(new_id))
    header.set("applyPageType", apply_page_type)

    sub = _make_sublist(text_width=text_width, text_height=header_height,
                        vert_align="TOP")
    header.append(sub)

    p = etree.SubElement(sub, f"{_HP}p")
    p.set("id", "0")
    p.set("paraPrIDRef", para_pr_id_ref)
    p.set("styleIDRef", style_id_ref)
    p.set("pageBreak", "0")
    p.set("columnBreak", "0")
    p.set("merged", "0")

    run = etree.SubElement(p, f"{_HP}run")
    run.set("charPrIDRef", char_pr_id_ref)
    t_el = etree.SubElement(run, f"{_HP}t")
    t_el.text = text

    p.append(_make_linesegarray(horzsize=text_width))

    # Insert after secPr run's existing ctrl elements (or at the end)
    # Find the right insertion point: after the last ctrl child of sec_run
    # But actually, headers go in a new <hp:run> inserted after sec_run
    # Looking at real files: headers are in the SECOND run of the first paragraph
    # (the first run has secPr + colPr, the second run has headers/footers)

    # Find or create the run for header/footer controls
    runs = sec_p.findall(f"{_HP}run")
    if len(runs) >= 2:
        # Use the second run (where headers/footers live)
        target_run = runs[1]
    else:
        # Create a new run after the secPr run
        target_run = etree.Element(f"{_HP}run")
        target_run.set("charPrIDRef", char_pr_id_ref)
        run_index = list(sec_p).index(sec_run)
        sec_p.insert(run_index + 1, target_run)

    # Insert the ctrl at the beginning of the target run (before any text)
    target_run.insert(0, ctrl)

    doc._dirty_sections.add(filename)


def set_footer(doc: HwpxDoc, text: str, align: str = "center",
               apply_page_type: str = "BOTH",
               section_index: int = 0,
               para_pr_id_ref: str = "0",
               char_pr_id_ref: str = "0",
               style_id_ref: str = "0") -> None:
    """Set or replace a footer in the specified section.

    If a footer with the same applyPageType already exists, its text is replaced.
    Otherwise, a new footer element is created.

    Args:
        doc: HwpxDoc instance (must be open).
        text: Footer text content.
        align: Alignment hint (informational).
        apply_page_type: "BOTH", "ODD", or "EVEN".
        section_index: Which section to modify (default: 0).
        para_pr_id_ref: paraPrIDRef for the footer paragraph.
        char_pr_id_ref: charPrIDRef for the footer run.
        style_id_ref: styleIDRef for the footer paragraph.

    Raises:
        ValueError: If section_index is invalid.
    """
    sections = doc.sections
    if section_index < 0 or section_index >= len(sections):
        raise ValueError(
            f"section_index {section_index} out of range "
            f"(document has {len(sections)} section(s))"
        )

    filename, root = sections[section_index]
    text_width = _get_text_width_from_section(root)
    footer_height = _get_footer_height_from_section(root)

    # Check for existing footer with same applyPageType
    for existing in root.iter(f"{_HP}footer"):
        if existing.get("applyPageType") == apply_page_type:
            # Replace text in existing footer
            sub = existing.find(f"{_HP}subList")
            if sub is not None:
                p = sub.find(f"{_HP}p")
                if p is not None:
                    run = p.find(f"{_HP}run")
                    if run is not None:
                        for old_t in run.findall(f"{_HP}t"):
                            run.remove(old_t)
                        t_el = etree.SubElement(run, f"{_HP}t")
                        t_el.text = text
                        doc._dirty_sections.add(filename)
                        return
            break

    # Find or create insertion point
    sec_p, sec_run = _find_secpr_run(root)
    if sec_run is None:
        raise ValueError(
            f"Cannot find a suitable run in section {section_index} "
            "to insert footer"
        )

    new_id = _max_header_footer_id(root) + 1

    # Build footer element
    ctrl = etree.Element(f"{_HP}ctrl")
    footer = etree.SubElement(ctrl, f"{_HP}footer")
    footer.set("id", str(new_id))
    footer.set("applyPageType", apply_page_type)

    sub = _make_sublist(text_width=text_width, text_height=footer_height,
                        vert_align="BOTTOM")
    footer.append(sub)

    p = etree.SubElement(sub, f"{_HP}p")
    p.set("id", "0")
    p.set("paraPrIDRef", para_pr_id_ref)
    p.set("styleIDRef", style_id_ref)
    p.set("pageBreak", "0")
    p.set("columnBreak", "0")
    p.set("merged", "0")

    run = etree.SubElement(p, f"{_HP}run")
    run.set("charPrIDRef", char_pr_id_ref)
    t_el = etree.SubElement(run, f"{_HP}t")
    t_el.text = text

    p.append(_make_linesegarray(horzsize=text_width))

    # Insert in appropriate run -- footers can go after headers
    # In real files, footers appear in later paragraphs' runs
    # But for simplicity, we insert in the same pattern as headers
    runs = sec_p.findall(f"{_HP}run")
    if len(runs) >= 2:
        target_run = runs[1]
    else:
        target_run = etree.Element(f"{_HP}run")
        target_run.set("charPrIDRef", char_pr_id_ref)
        run_index = list(sec_p).index(sec_run)
        sec_p.insert(run_index + 1, target_run)

    # Insert at end of target run (footers after headers)
    target_run.append(ctrl)

    doc._dirty_sections.add(filename)


def set_page_number(
    doc: HwpxDoc,
    position: str = "footer_center",
    fmt: str = "- {page} -",
    number_type: str = "DIGIT",
    section_index: int = 0,
    para_pr_id_ref: str = "0",
    char_pr_id_ref: str = "0",
    style_id_ref: str = "0",
) -> None:
    """Add page numbering to the document.

    Page numbers in HWPX are autoNum controls inside headers or footers.
    This function creates or updates a footer (or header) with a page number.

    Args:
        doc: HwpxDoc instance (must be open).
        position: Where to place the page number.
            "footer_center", "footer_left", "footer_right",
            "header_center", "header_left", "header_right".
        fmt: Format string. Use {page} as placeholder for the page number.
            Examples: "- {page} -", "{page}", "Page {page}".
        number_type: "DIGIT" (1,2,3), "ROMAN_SMALL" (i,ii,iii),
                     "ROMAN_CAPITAL" (I,II,III).
        section_index: Which section to modify.
        para_pr_id_ref: paraPrIDRef for the page number paragraph.
        char_pr_id_ref: charPrIDRef for the page number run.
        style_id_ref: styleIDRef for the page number paragraph.

    Raises:
        ValueError: If section_index is invalid or position is unrecognized.
    """
    valid_positions = {
        "footer_center", "footer_left", "footer_right",
        "header_center", "header_left", "header_right",
    }
    if position not in valid_positions:
        raise ValueError(
            f"position must be one of {valid_positions}, got {position!r}"
        )

    sections = doc.sections
    if section_index < 0 or section_index >= len(sections):
        raise ValueError(
            f"section_index {section_index} out of range "
            f"(document has {len(sections)} section(s))"
        )

    filename, root = sections[section_index]
    text_width = _get_text_width_from_section(root)

    is_header = position.startswith("header_")
    container_tag = "header" if is_header else "footer"
    vert_align = "TOP" if is_header else "BOTTOM"
    height = (_get_header_height_from_section(root) if is_header
              else _get_footer_height_from_section(root))

    # Split format string around {page}
    parts = fmt.split("{page}")
    prefix_text = parts[0] if len(parts) > 0 else ""
    suffix_text = parts[1] if len(parts) > 1 else ""

    # Build the page number paragraph
    page_p = etree.Element(f"{_HP}p")
    page_p.set("id", "0")
    page_p.set("paraPrIDRef", para_pr_id_ref)
    page_p.set("styleIDRef", style_id_ref)
    page_p.set("pageBreak", "0")
    page_p.set("columnBreak", "0")
    page_p.set("merged", "0")

    page_run = etree.SubElement(page_p, f"{_HP}run")
    page_run.set("charPrIDRef", char_pr_id_ref)

    # Prefix text (e.g., "- ")
    if prefix_text:
        t_prefix = etree.SubElement(page_run, f"{_HP}t")
        t_prefix.text = prefix_text

    # autoNum PAGE control
    ctrl = etree.SubElement(page_run, f"{_HP}ctrl")
    auto_num = etree.SubElement(ctrl, f"{_HP}autoNum")
    auto_num.set("num", "1")
    auto_num.set("numType", "PAGE")
    auto_fmt = etree.SubElement(auto_num, f"{_HP}autoNumFormat")
    auto_fmt.set("type", number_type)
    auto_fmt.set("userChar", "")
    auto_fmt.set("prefixChar", "")
    auto_fmt.set("suffixChar", "")
    auto_fmt.set("supscript", "0")

    # Suffix text (e.g., " -")
    t_suffix = etree.SubElement(page_run, f"{_HP}t")
    t_suffix.text = suffix_text if suffix_text else None

    page_p.append(_make_linesegarray(horzsize=text_width))

    # Find the secPr run
    sec_p, sec_run = _find_secpr_run(root)
    if sec_run is None:
        raise ValueError(
            f"Cannot find a suitable run in section {section_index} "
            "to insert page number"
        )

    new_id = _max_header_footer_id(root) + 1

    # Build the container (header or footer)
    outer_ctrl = etree.Element(f"{_HP}ctrl")
    container = etree.SubElement(outer_ctrl, f"{_HP}{container_tag}")
    container.set("id", str(new_id))
    container.set("applyPageType", "BOTH")

    sub = _make_sublist(text_width=text_width, text_height=height,
                        vert_align=vert_align)
    container.append(sub)
    sub.append(page_p)

    # Insert
    runs = sec_p.findall(f"{_HP}run")
    if len(runs) >= 2:
        target_run = runs[1]
    else:
        target_run = etree.Element(f"{_HP}run")
        target_run.set("charPrIDRef", char_pr_id_ref)
        run_index = list(sec_p).index(sec_run)
        sec_p.insert(run_index + 1, target_run)

    target_run.append(outer_ctrl)
    doc._dirty_sections.add(filename)


def add_memo_to_doc(
    doc: HwpxDoc,
    anchor_text: str,
    memo: str,
    author: str = "",
    section_index: int = 0,
) -> None:
    """Add a memo (comment) anchored to text containing *anchor_text*.

    Note: Memo/comment support is based on OWPML specification knowledge.
    The exact XML structure has not been verified against real Hangul memo
    output since no example files with memos were available. The structure
    follows the same pattern as footnotes/endnotes.

    Args:
        doc: HwpxDoc instance (must be open).
        anchor_text: Text string to search for.
        memo: Memo/comment text.
        author: Author name (optional).
        section_index: Which section to search (default: 0).

    Raises:
        ValueError: If anchor_text is not found or section_index is invalid.
    """
    sections = doc.sections
    if section_index < 0 or section_index >= len(sections):
        raise ValueError(
            f"section_index {section_index} out of range "
            f"(document has {len(sections)} section(s))"
        )

    filename, root = sections[section_index]
    run, t_el = _find_run_containing_text(root, anchor_text)
    if run is None:
        raise ValueError(
            f"anchor_text {anchor_text!r} not found in section {section_index}"
        )

    # Build memo element (best-effort based on OWPML spec pattern)
    ctrl = etree.Element(f"{_HP}ctrl")
    memo_el = etree.SubElement(ctrl, f"{_HP}memo")
    memo_el.set("width", "12000")  # Default memo width
    memo_el.set("lineType", "SOLID")
    memo_el.set("lineColor", "#FFFF00")  # Yellow highlight
    memo_el.set("fillColor", "#FFFFCC")  # Light yellow background
    memo_el.set("fillAlpha", "255")
    if author:
        memo_el.set("author", author)

    sub = _make_sublist()
    memo_el.append(sub)

    p = etree.SubElement(sub, f"{_HP}p")
    p.set("id", "0")
    p.set("paraPrIDRef", "0")
    p.set("styleIDRef", "0")
    p.set("pageBreak", "0")
    p.set("columnBreak", "0")
    p.set("merged", "0")

    memo_run = etree.SubElement(p, f"{_HP}run")
    memo_run.set("charPrIDRef", "0")
    t_el_memo = etree.SubElement(memo_run, f"{_HP}t")
    t_el_memo.text = memo

    p.append(_make_linesegarray())

    # Insert ctrl after the anchor text
    parent = t_el.getparent()
    t_index = list(parent).index(t_el)
    parent.insert(t_index + 1, ctrl)

    doc._dirty_sections.add(filename)
