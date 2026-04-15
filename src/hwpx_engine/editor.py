"""HWPX document editor — str_replace based precision editing with template-aware insertion."""

import os
import time
import zipfile
from pathlib import Path
from lxml import etree
from hwpx_engine.hwpx_doc import HwpxDoc
from hwpx_engine.charpr_manager import CharPrManager
from hwpx_engine.formatter import DEFAULT_STYLE_MAP, StyleMapper
from hwpx_engine.validator import HwpxValidator
from hwpx_engine.utils import fix_namespaces
from hwpx_engine.xml_primitives import has_part_overrides, xml_escape

HP = '{http://www.hancom.co.kr/hwpml/2011/paragraph}'
HH = '{http://www.hancom.co.kr/hwpml/2011/head}'
_HP_URI = 'http://www.hancom.co.kr/hwpml/2011/paragraph'


class TextNotFoundError(Exception):
    pass


class HwpxEditor:
    def __init__(self, doc, source_path, style_mapper=None):
        self._doc = doc
        self._source_path = source_path
        self._style = style_mapper or StyleMapper(DEFAULT_STYLE_MAP)
        self._zip_modified = False  # True when ZIP was modified directly (parts/charPr)
        self._dom_dirty = False     # True when in-memory DOM has unsaved changes

    @classmethod
    def open(cls, hwpx_path, template_id=None):
        """Open an HWPX file for editing.

        Args:
            hwpx_path: Path to the .hwpx file.
            template_id: Optional registered template ID (directory name under assets/registered/).
                         Enables style-name based insertion via metadata.json.
        """
        doc = HwpxDoc.open(hwpx_path)
        if template_id:
            mapper = StyleMapper.from_template_id(template_id)
        else:
            mapper = StyleMapper(DEFAULT_STYLE_MAP)
        return cls(doc, str(hwpx_path), mapper)

    # --- Internal: DOM/ZIP synchronization ---

    def _release_doc(self):
        """Release document file handle (needed on Windows before file replacement)."""
        if self._doc is not None:
            try:
                self._doc.close()
            except Exception:
                pass
            self._doc = None

    @staticmethod
    def _safe_replace(src, dst, retries=5, delay=0.1):
        """os.replace with retry for environments like Dropbox that briefly lock files."""
        for i in range(retries):
            try:
                os.replace(src, dst)
                return
            except PermissionError:
                if i == retries - 1:
                    raise
                time.sleep(delay)

    def _flush_dom(self):
        """Persist in-memory DOM changes to the source ZIP file.

        Called automatically before any ZIP-level operation to prevent
        DOM changes from being lost when the ZIP is rewritten.
        """
        if not self._dom_dirty:
            return
        # Save to tmp first, then replace, then reload in-memory state.
        tmp = self._source_path + ".dom_flush_tmp"
        self._doc.save_to_path(tmp)
        self._safe_replace(tmp, self._source_path)
        self._doc.reload_from_path(self._source_path)
        self._dom_dirty = False

    @staticmethod
    def _element_to_xml(element):
        """Serialize lxml element to XML string for ZIP-level injection.

        Normalizes namespace prefix to hp: and strips declarations
        since they're already present in the target section XML.
        """
        import re
        etree.register_namespace('hp', _HP_URI)
        xml = etree.tostring(element, encoding='unicode')
        # lxml may use ns0:/ns1: etc — find actual prefix and normalize to hp:
        ns_match = re.search(r'xmlns:(\w+)="' + re.escape(_HP_URI) + '"', xml)
        if ns_match:
            prefix = ns_match.group(1)
            if prefix != 'hp':
                xml = xml.replace(f'{prefix}:', 'hp:')
            xml = re.sub(r'\s*xmlns:\w+="' + re.escape(_HP_URI) + '"', '', xml)
        return xml

    # --- Text operations (no style needed) ---

    def find_text(self, pattern, section=None, context_chars=50):
        """Search for text across all paragraphs in the document.

        Args:
            pattern: String for substring match, or compiled regex for regex match.
            section: Optional section filename to limit search (e.g. "Contents/section0.xml").
            context_chars: Number of characters of surrounding context to include.

        Returns:
            List of dicts: [{"section": str, "para_index": int, "text": str, "context": str}]
        """
        import re as _re

        is_regex = isinstance(pattern, _re.Pattern)
        results = []

        for sec_name, sec_el in self._doc.sections:
            if section is not None and sec_name != section:
                continue

            for para_index, p_el in enumerate(sec_el.findall(f'.//{HP}p')):
                # Collect all text from <hp:t> elements within this paragraph
                full_text = ''.join(
                    (t.text or '') for t in p_el.iter(f'{HP}t')
                )
                if not full_text:
                    continue

                if is_regex:
                    for m in pattern.finditer(full_text):
                        idx = m.start()
                        ctx_start = max(0, idx - context_chars)
                        ctx_end = min(len(full_text), m.end() + context_chars)
                        results.append({
                            "section": sec_name,
                            "para_index": para_index,
                            "text": full_text,
                            "context": full_text[ctx_start:ctx_end],
                        })
                else:
                    idx = full_text.find(pattern)
                    if idx != -1:
                        ctx_start = max(0, idx - context_chars)
                        ctx_end = min(len(full_text), idx + len(pattern) + context_chars)
                        results.append({
                            "section": sec_name,
                            "para_index": para_index,
                            "text": full_text,
                            "context": full_text[ctx_start:ctx_end],
                        })

        return results

    def replace_text(self, find, replace, match="first", match_index=None):
        count = 0
        occurrence = 0
        for run in self._doc.iter_runs():
            if run.text and find in run.text:
                if match_index is not None:
                    if occurrence == match_index:
                        run.replace_text(find, replace)
                        self._dom_dirty = True
                        return 1
                    occurrence += 1
                elif match == "first":
                    run.replace_text(find, replace)
                    self._dom_dirty = True
                    return 1
                else:
                    run.replace_text(find, replace)
                    count += 1
        if count == 0 and match_index is None:
            raise TextNotFoundError(f"Text not found: '{find}'")
        if match_index is not None and occurrence <= match_index:
            raise TextNotFoundError(f"Text '{find}' occurrence {match_index} not found (only {occurrence} found)")
        if count > 0:
            self._dom_dirty = True
        return count

    def str_replace(self, old_string, new_string):
        """Edit-tool style precise replacement.

        Works like Claude Code's Edit tool: old_string must be a unique,
        context-rich string that appears exactly ONCE across all section XMLs.
        """
        self._flush_dom()

        matches = []
        with zipfile.ZipFile(self._source_path, "r") as zin:
            for item in zin.infolist():
                if item.filename.startswith("Contents/") and item.filename.endswith(".xml"):
                    text = zin.read(item.filename).decode("utf-8")
                    n = text.count(old_string)
                    if n > 0:
                        matches.append((item.filename, n))

        total = sum(n for _, n in matches)
        if total == 0:
            raise TextNotFoundError(f"Text not found: '{old_string[:50]}...'")
        if total > 1:
            locs = ", ".join(f"{f}({n})" for f, n in matches)
            raise TextNotFoundError(
                f"Text found {total} times ({locs}). "
                f"Use a longer, more unique string.")

        import re
        _lsa_re = re.compile(r'<hp:linesegarray>.*?</hp:linesegarray>', re.DOTALL)

        target_file = matches[0][0]
        tmp = self._source_path + ".str_replace_tmp"
        with zipfile.ZipFile(self._source_path, "r") as zin:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == target_file:
                        text = data.decode("utf-8")
                        text = text.replace(old_string, new_string, 1)
                        text = _lsa_re.sub('', text)
                        data = text.encode("utf-8")
                    if item.filename == "mimetype":
                        zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                    else:
                        zout.writestr(item, data)
        self._safe_replace(tmp, self._source_path)
        self._doc.reload_from_path(self._source_path)
        self._dom_dirty = False
        return 1

    def zip_str_replace(self, find, replace, match="all"):
        """Raw XML string replacement (use str_replace for precision editing)."""
        self._flush_dom()

        import re
        _lsa_re = re.compile(r'<hp:linesegarray>.*?</hp:linesegarray>', re.DOTALL)

        tmp = self._source_path + ".str_replace_tmp"
        count = 0
        with zipfile.ZipFile(self._source_path, "r") as zin:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename.startswith("Contents/") and item.filename.endswith(".xml"):
                        text = data.decode("utf-8")
                        if find in text:
                            if match == "first":
                                text = text.replace(find, replace, 1)
                                count += 1
                            else:
                                count += text.count(find)
                                text = text.replace(find, replace)
                            text = _lsa_re.sub('', text)
                        data = text.encode("utf-8")
                    if item.filename == "mimetype":
                        zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                    else:
                        zout.writestr(item, data)
        if count == 0:
            os.remove(tmp)
            raise TextNotFoundError(f"Text not found in XML: '{find}'")
        self._safe_replace(tmp, self._source_path)
        self._doc.reload_from_path(self._source_path)
        self._dom_dirty = False
        return count

    def batch_replace(self, pairs):
        """Apply multiple text replacements in a single ZIP read/write pass.

        Each call to str_replace or zip_str_replace rewrites the entire ZIP file.
        With 100+ replacements, repeated ZIP rewrites can cause progressive data
        loss (images, bindata shrinkage). This method applies ALL replacements
        in one pass, avoiding the degradation.

        Args:
            pairs: list of (old_text, new_text) tuples.
                   Each replacement is applied to all section XMLs.

        Returns:
            dict with 'applied' (count of pairs that matched) and 'skipped' (not found).
        """
        self._flush_dom()

        tmp = self._source_path + ".batch_replace_tmp"
        applied = 0
        skipped = 0
        # Count matches first (before replacing)
        with zipfile.ZipFile(self._source_path, "r") as zin:
            for old, new in pairs:
                found = False
                for item in zin.infolist():
                    fn = item.filename.replace("\\", "/")
                    if fn.startswith("Contents/") and fn.endswith(".xml"):
                        if old in zin.read(item.filename).decode("utf-8"):
                            found = True
                            break
                if found:
                    applied += 1
                else:
                    skipped += 1

        # Apply all replacements in a single ZIP pass
        import re
        _lsa_pattern = re.compile(
            r'<hp:linesegarray>.*?</hp:linesegarray>', re.DOTALL)

        with zipfile.ZipFile(self._source_path, "r") as zin:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    fn = item.filename.replace("\\", "/")
                    if fn.startswith("Contents/") and fn.endswith(".xml"):
                        text = data.decode("utf-8")
                        changed = False
                        for old, new in pairs:
                            if old in text:
                                text = text.replace(old, new)
                                changed = True
                        if changed:
                            # Strip linesegarray from modified sections.
                            # Text replacements change character counts, making
                            # existing lineseg textpos values invalid. 한글
                            # recalculates line segments when they're absent,
                            # but crashes when textpos exceeds the new text length.
                            text = _lsa_pattern.sub('', text)
                        data = text.encode("utf-8")
                    if item.filename == "mimetype":
                        zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                    else:
                        zout.writestr(item, data)

        self._safe_replace(tmp, self._source_path)
        self._doc.reload_from_path(self._source_path)
        self._dom_dirty = False
        return {"applied": applied, "skipped": skipped}

    # --- Paragraph insertion (template-aware) ---

    def insert_paragraph(self, text=None, after=None, style="body", parts=None):
        """Insert a new paragraph with template-aware styling.

        Simple usage:
            editor.insert_paragraph('텍스트', style='bullet_2', after='기존 텍스트')

        Rich text with inline formatting:
            editor.insert_paragraph(parts=[
                {'text': '일반 텍스트 '},
                {'text': '볼드', 'bold': True},
                {'text': ' 빨간색', 'color': '#FF0000'},
            ], style='bullet_2', after='기존 텍스트')

        Supported part overrides: bold, italic, color, size (in 0.1pt, e.g. 1100=11pt)
        """
        if parts is None and text is not None:
            parts = [{'text': text}]
        if not parts:
            raise ValueError("Either text or parts must be provided")

        cpr, ppr, sref = self._style.resolve(style)

        if len(parts) == 1 and not has_part_overrides(parts[0]) and after is None:
            # Simple case: append at end via direct lxml
            self._doc.add_paragraph(parts[0]['text'],
                                    char_pr_id_ref=str(cpr),
                                    para_pr_id_ref=str(ppr))
            self._dom_dirty = True
        else:
            # Anchored insertion or rich text — always use ZIP-level
            # (DOM add_paragraph ignores anchor position and appends at end)
            self._insert_rich_paragraph(parts, ppr, sref, cpr, after)

    def _insert_rich_paragraph(self, parts, ppr, sref, base_cpr, after):
        """Insert paragraph with multiple runs (rich text) via ZIP manipulation.

        Uses CharPrManager for charPr clone/reuse (loads header.xml once,
        writes back once if modified).
        """
        self._flush_dom()

        # Load header.xml and create CharPrManager
        with zipfile.ZipFile(self._source_path, "r") as z:
            header_root = etree.fromstring(z.read("Contents/header.xml"))
        mgr = CharPrManager(header_root)

        # Build run XML fragments
        run_xmls = []
        for part in parts:
            text = part.get('text', '')
            if has_part_overrides(part):
                cpr_id = mgr.find_or_create_charpr_from_part(str(base_cpr), part)
            else:
                cpr_id = str(base_cpr)
            escaped_text = xml_escape(text)
            run_xmls.append(
                f'<hp:run charPrIDRef="{cpr_id}"><hp:t>{escaped_text}</hp:t></hp:run>'
            )

        # Write back header.xml if charPr was added
        if mgr.modified:
            self._rewrite_zip_file("Contents/header.xml", mgr.serialize())

        # Build paragraph XML
        sref_attr = f' styleIDRef="{sref}"' if sref is not None else ''
        para_xml = (
            f'<hp:p id="0" paraPrIDRef="{ppr}"{sref_attr}>'
            + ''.join(run_xmls)
            + '</hp:p>'
        )

        # Find target section XML and insert after anchor
        self._zip_insert_paragraph(para_xml, after)

    def _zip_insert_paragraph(self, para_xml, after):
        """Insert paragraph XML into the document's section XML via ZIP manipulation."""
        self._flush_dom()

        tmp = self._source_path + ".insert_tmp"

        with zipfile.ZipFile(self._source_path, "r") as zin:
            # Find which section contains the anchor text
            target_section = None
            if after:
                for item in zin.infolist():
                    if item.filename.startswith("Contents/section") and item.filename.endswith(".xml"):
                        content = zin.read(item.filename).decode("utf-8")
                        if after in content:
                            target_section = item.filename
                            break
                if target_section is None:
                    raise TextNotFoundError(f"Anchor text not found: '{after}'")
            else:
                # Find last section
                sections = [i.filename for i in zin.infolist()
                            if i.filename.startswith("Contents/section") and i.filename.endswith(".xml")]
                target_section = sorted(sections)[-1] if sections else None

            # Pre-validate: anchor must exist at top-level (not only inside table cells)
            if after and target_section:
                section_text = zin.read(target_section).decode("utf-8")
                if _find_toplevel_anchor(section_text, after) < 0:
                    raise TextNotFoundError(
                        f"Anchor text '{after}' found only inside table cells, "
                        "not at top-level paragraph"
                    )

            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == target_section:
                        text = data.decode("utf-8")
                        if after:
                            # Find anchor text, then find the top-level </hp:p>
                            # Must escape table cells: if anchor is inside <hp:tc>,
                            # walk up to find the enclosing top-level </hp:p>
                            anchor_pos = _find_toplevel_anchor(text, after)
                            if anchor_pos >= 0:
                                insert_point = _find_toplevel_p_end(text, anchor_pos)
                                if insert_point >= 0:
                                    text = text[:insert_point] + para_xml + text[insert_point:]
                        else:
                            # Append before closing tag of root element
                            last_close = text.rfind('</')
                            if last_close >= 0:
                                text = text[:last_close] + para_xml + text[last_close:]
                        data = text.encode("utf-8")
                    if item.filename == "mimetype":
                        zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                    else:
                        zout.writestr(item, data)

        self._safe_replace(tmp, self._source_path)
        self._zip_modified = True
        self._doc.reload_from_path(self._source_path)
        self._dom_dirty = False

    def _rewrite_zip_file(self, target_filename, new_data):
        """Replace a single file inside the HWPX ZIP."""
        tmp = self._source_path + ".rewrite_tmp"
        with zipfile.ZipFile(self._source_path, "r") as zin:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == target_filename:
                        data = new_data if isinstance(new_data, bytes) else new_data.encode("utf-8")
                    if item.filename == "mimetype":
                        zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                    else:
                        zout.writestr(item, data)
        self._safe_replace(tmp, self._source_path)
        self._doc.reload_from_path(self._source_path)

    # --- Table/figure insertion (ZIP-level) ---

    def insert_table(self, rows_data, col_count, col_widths,
                     header_style, body_style, page_width,
                     after=None, caption_text=None, caption_style=None,
                     merges=None):
        """Insert a new table via ZIP-level injection.

        Args:
            rows_data: list of (row_type, cells) — row_type is 'header' or 'data'
            col_count: number of columns
            col_widths: list of column widths in HWPML units
            header_style: dict with 'pPr', 'sRef', 'cPr', 'borderFill', 'borderFill_right'
            body_style: dict with same keys
            page_width: total table width
            after: anchor text to insert after
            caption_text: optional table caption string
            caption_style: dict with 'pPr', 'sRef', 'cPr' for caption
            merges: optional list of merge dicts
        """
        from hwpx_engine.xml_primitives import make_table_xml, make_para

        caption_paras = []
        if caption_text and caption_style:
            cp = make_para(
                text=caption_text,
                char_pr=str(caption_style.get('cPr', '0')),
                para_pr=str(caption_style.get('pPr', '0')),
                style_ref=str(caption_style.get('sRef', '0')),
            )
            caption_paras.append(cp)

        tbl_para = make_table_xml(
            col_count=col_count,
            rows_data=rows_data,
            col_widths=col_widths,
            header_style=header_style,
            body_style=body_style,
            page_width=page_width,
            caption_paras=caption_paras,
            merges=merges,
        )

        para_xml = self._element_to_xml(tbl_para)
        self._zip_insert_paragraph(para_xml, after)

    def insert_figure_box(self, width, height, border_fill,
                          after=None, caption_text=None, caption_style=None,
                          figure_box_style=None):
        """Insert an empty figure box via ZIP-level injection.

        Caption is inserted ABOVE the figure box (Korean academic convention).

        Args:
            width: box width in HWPML units
            height: box height in HWPML units
            border_fill: borderFillIDRef for the box cell
            after: anchor text to insert after
            caption_text: optional figure caption string
            caption_style: dict with 'pPr', 'sRef', 'cPr' for caption
            figure_box_style: dict with 'pPr', 'sRef', 'cPr' for the figure box paragraph
                              (use CENTER-aligned pPr for centering)
        """
        from hwpx_engine.xml_primitives import make_figure_box as _make_fig, make_para

        fig_kwargs = {}
        if figure_box_style:
            fig_kwargs['para_pr'] = str(figure_box_style.get('pPr', '11'))
            fig_kwargs['style_ref'] = str(figure_box_style.get('sRef', '0'))
            fig_kwargs['char_pr'] = str(figure_box_style.get('cPr', '4'))

        fig_para = _make_fig(width, height, border_fill, **fig_kwargs)
        fig_xml = self._element_to_xml(fig_para)

        if caption_text and caption_style:
            cap_para = make_para(
                text=caption_text,
                char_pr=str(caption_style.get('cPr', '0')),
                para_pr=str(caption_style.get('pPr', '0')),
                style_ref=str(caption_style.get('sRef', '0')),
            )
            cap_xml = self._element_to_xml(cap_para)
            combined = cap_xml + fig_xml  # caption ABOVE figure
        else:
            combined = fig_xml

        self._zip_insert_paragraph(combined, after)

    # --- Table editing (DOM-level, always uses current doc) ---

    def set_cell(self, table_index, row, col, text):
        """Set cell text in an existing table.

        Fully replaces all cell content (removes all existing paragraphs,
        creates a fresh one). Uses lxml.etree exclusively to avoid
        xml.etree.ElementTree/lxml incompatibility.
        """
        from hwpx_engine.tables import TableEditor
        te = TableEditor(self._doc)
        te.set_cell(table_index, row, col, text)
        self._dom_dirty = True

    def replace_in_table(self, table_index, old_text, new_text, replace_all=False):
        """Replace text only within a specific table's XML subtree.

        Like str_replace, enforces uniqueness by default: old_text must appear
        exactly once within the table. If it appears multiple times, raises
        TextNotFoundError with the match count so the caller can provide a
        more specific string. Pass replace_all=True to replace every occurrence.

        Args:
            table_index: 0-based index of the target table.
            old_text: Text to find.
            new_text: Replacement text.
            replace_all: If True, replace all occurrences without uniqueness check.

        Returns:
            True if at least one replacement was made.

        Raises:
            TextNotFoundError: If old_text is not found, or found multiple times
                               (when replace_all=False).
        """
        from hwpx_engine.tables import TableEditor
        te = TableEditor(self._doc)
        tables = te.tables
        if table_index >= len(tables):
            raise IndexError(f'table_index {table_index} out of range (have {len(tables)} tables)')

        table = tables[table_index]
        tbl_el = table.element

        # Count occurrences first
        matches = []
        for t_node in tbl_el.iter(f'{HP}t'):
            if t_node.text and old_text in t_node.text:
                count = t_node.text.count(old_text)
                matches.append((t_node, count))

        total = sum(c for _, c in matches)
        if total == 0:
            raise TextNotFoundError(
                f"Text not found in table {table_index}: '{old_text[:50]}'"
            )
        if total > 1 and not replace_all:
            raise TextNotFoundError(
                f"Text found {total} times in table {table_index}. "
                f"Use a longer, more unique string, or pass replace_all=True."
            )

        # Apply replacement
        for t_node, _ in matches:
            t_node.text = t_node.text.replace(old_text, new_text)

        table.mark_dirty()
        self._dom_dirty = True
        return True

    def add_row(self, table_index, cells, position=None):
        """Add a row to an existing table."""
        from hwpx_engine.tables import TableEditor
        te = TableEditor(self._doc)
        te.add_row(table_index, cells, position=position)
        self._dom_dirty = True

    def delete_row(self, table_index, row_index):
        """Delete a row from an existing table."""
        from hwpx_engine.tables import TableEditor
        te = TableEditor(self._doc)
        te.delete_row(table_index, row_index)
        self._dom_dirty = True

    def add_column(self, table_index, cells, position=None):
        """Add a column to an existing table."""
        from hwpx_engine.tables import TableEditor
        te = TableEditor(self._doc)
        te.add_column(table_index, cells, position=position)
        self._dom_dirty = True

    def delete_column(self, table_index, col_index):
        """Delete a column from an existing table."""
        from hwpx_engine.tables import TableEditor
        te = TableEditor(self._doc)
        te.delete_column(table_index, col_index)
        self._dom_dirty = True

    def get_cell(self, table_index, row, col):
        """Return the text content of a table cell.

        Deactivated (merged-away) cells return an empty string.
        Raises IndexError if table_index, row, or col is out of range.
        """
        from hwpx_engine.tables import TableEditor
        return TableEditor(self._doc).get_cell(table_index, row, col)

    def get_table_data(self, table_index):
        """Extract a full table as a 2-D list of strings.

        Deactivated (merged-away) cells appear as empty strings.
        Raises IndexError if table_index is out of range.
        """
        from hwpx_engine.tables import TableEditor
        return TableEditor(self._doc).get_table_data(table_index)

    def batch_set_cell(self, table_index, updates):
        """Set multiple cells in one table with a single DOM traversal.

        Args:
            table_index: Index of the target table (0-based).
            updates: List of (row, col, text) tuples.
        """
        from hwpx_engine.tables import TableEditor
        TableEditor(self._doc).batch_set_cell(table_index, updates)
        self._dom_dirty = True

    def find_table(self, text_pattern, match_row=0):
        """Find tables containing text_pattern in a specific row.

        Args:
            text_pattern: String to search for in the row text.
            match_row: Which row to search (0 = header row).

        Returns:
            list[int]: Indices of matching tables.
        """
        tables = self._doc.get_tables()
        matches = []
        for i, tbl in enumerate(tables):
            trs = tbl.element.findall(f'{HP}tr')
            if match_row >= len(trs):
                continue
            row_texts = []
            for tc in trs[match_row].findall(f'{HP}tc'):
                for t_el in tc.iter(f'{HP}t'):
                    if t_el.text:
                        row_texts.append(t_el.text)
            row_full = ' '.join(row_texts)
            if text_pattern in row_full:
                matches.append(i)
        return matches

    # --- Table removal ---

    @staticmethod
    def _para_text(p_el):
        """Extract concatenated text from an <hp:p> element."""
        texts = []
        for t_el in p_el.iter(f'{HP}t'):
            if t_el.text:
                texts.append(t_el.text)
        return ''.join(texts).strip()

    @staticmethod
    def _is_table_related(text):
        """Check if paragraph text is table-related (caption, unit, source, note)."""
        if not text:
            return True  # empty paragraph
        t = text.strip()
        if not t or t == '-' or t == '–':
            return True
        # Source line: "자료:", "자료 :", "출처:"
        if t.startswith('자료') or t.startswith('출처'):
            return True
        # Note line: "주:", "주 :", "※", "* "
        if t.startswith('주:') or t.startswith('주 :') or t.startswith('※'):
            return True
        if t.startswith('* ') and len(t) < 200:
            return True
        # Unit line: "단위" with common units
        if '단위' in t and len(t) < 50:
            return True
        # Caption: "[표 " pattern
        if t.startswith('[표 ') or t.startswith('[표\xa0'):
            return True
        return False

    def _get_table_paragraph(self, table_index):
        """Return (tbl_p, container, tbl_idx_in_container) for a table.

        tbl_p is the <hp:p> element wrapping the table.
        container is the parent element holding all sibling paragraphs.
        """
        from hwpx_engine.tables import TableEditor
        te = TableEditor(self._doc)
        tables = te.tables
        if table_index >= len(tables):
            raise IndexError(
                f'table_index {table_index} out of range '
                f'(have {len(tables)} tables)')

        tbl_el = tables[table_index].element
        section_file = tables[table_index]._section

        # Walk up: tbl -> run -> p
        tbl_p = tbl_el.getparent()
        while tbl_p is not None:
            if tbl_p.tag == f'{HP}p':
                break
            tbl_p = tbl_p.getparent()
        if tbl_p is None:
            raise RuntimeError(
                f'Could not find wrapping <hp:p> for table {table_index}')

        container = tbl_p.getparent()
        if container is None:
            raise RuntimeError(
                f'Table paragraph has no parent container')

        all_children = list(container)
        idx = None
        for i, child in enumerate(all_children):
            if child is tbl_p:
                idx = i
                break

        return tbl_p, container, all_children, idx, section_file

    def get_nearby_paragraphs(self, table_index, before=3, after=2):
        """Return paragraphs near a table with metadata.

        Does NOT judge whether a paragraph is a caption — that decision
        belongs to the caller (agent or script).

        Parameters
        ----------
        table_index : int
            Target table index (0-based).
        before : int
            How many paragraphs before the table to include.
        after : int
            How many paragraphs after the table to include.

        Returns
        -------
        list[dict]
            Each dict:
            ``offset`` (int): negative=before, positive=after.
            ``text`` (str): concatenated text.
            ``para_pr`` (str): paraPrIDRef.
            ``style_ref`` (str): styleIDRef.
            ``char_pr`` (str): first run's charPrIDRef.
            ``has_auto_num`` (bool): ``<hp:autoNum>`` present.
        """
        tbl_p, container, all_children, tbl_ci, _ = \
            self._get_table_paragraph(table_index)

        result = []

        # Scan before
        for dist in range(1, before + 1):
            ci = tbl_ci - dist
            if ci < 0:
                break
            sib = all_children[ci]
            if sib.tag != f'{HP}p':
                break
            result.append(self._para_metadata(sib, -dist))

        # Scan after
        for dist in range(1, after + 1):
            ci = tbl_ci + dist
            if ci >= len(all_children):
                break
            sib = all_children[ci]
            if sib.tag != f'{HP}p':
                break
            result.append(self._para_metadata(sib, dist))

        result.sort(key=lambda d: d['offset'])
        return result

    def set_paragraph_style(self, table_index, offset,
                            para_pr=None, style_ref=None, char_pr=None):
        """Change style of a paragraph near a table.

        Parameters
        ----------
        table_index : int
        offset : int
            -1 = paragraph immediately before the table,
            -2 = two before, 1 = immediately after, etc.
        para_pr : str, optional
            New paraPrIDRef. None = no change.
        style_ref : str, optional
            New styleIDRef. None = no change.
        char_pr : str, optional
            New charPrIDRef for ALL runs. None = no change.
        """
        tbl_p, container, all_children, tbl_ci, section_file = \
            self._get_table_paragraph(table_index)

        target_ci = tbl_ci + offset
        if target_ci < 0 or target_ci >= len(all_children):
            raise IndexError(
                f'offset {offset} from table {table_index} is out of range')

        target_p = all_children[target_ci]
        if target_p.tag != f'{HP}p':
            raise ValueError(
                f'Element at offset {offset} is not a paragraph')

        if para_pr is not None:
            target_p.set('paraPrIDRef', str(para_pr))
        if style_ref is not None:
            target_p.set('styleIDRef', str(style_ref))
        if char_pr is not None:
            for run in target_p.findall(f'{HP}run'):
                run.set('charPrIDRef', str(char_pr))

        self._doc._dirty_sections.add(section_file)
        self._dom_dirty = True

    @staticmethod
    def _para_metadata(p_el, offset):
        """Build metadata dict for a paragraph element."""
        texts = []
        for t_el in p_el.iter(f'{HP}t'):
            if t_el.text:
                texts.append(t_el.text)
        text = ''.join(texts).strip()

        char_pr = ''
        first_run = p_el.find(f'{HP}run')
        if first_run is not None:
            char_pr = first_run.get('charPrIDRef', '')

        has_auto_num = p_el.find(f'.//{HP}autoNum') is not None

        return {
            'offset': offset,
            'text': text,
            'para_pr': p_el.get('paraPrIDRef', ''),
            'style_ref': p_el.get('styleIDRef', ''),
            'char_pr': char_pr,
            'has_auto_num': has_auto_num,
        }

    def delete_table(self, table_index, clean_surrounding=True):
        """Delete an entire table and its surrounding context from the document.

        Removes:
          1. The <hp:p> element that wraps the <hp:tbl>
          2. If clean_surrounding=True (default), also removes adjacent
             paragraphs that are table-related: caption ([표 ...]),
             unit line (단위: ...), source (자료: ...), notes (주: ..., ※),
             and empty/dash-only lines.

        Scanning stops at the first non-table-related paragraph in each
        direction, so body text is never accidentally removed.
        """
        tbl_p, container, all_children, tbl_idx, section_file = \
            self._get_table_paragraph(table_index)

        to_remove = [tbl_p]

        if clean_surrounding and tbl_idx is not None:
            # Scan BEFORE table (up to 3 paragraphs)
            for i in range(tbl_idx - 1, max(tbl_idx - 4, -1), -1):
                sib = all_children[i]
                if sib.tag == f'{HP}p' and self._is_table_related(self._para_text(sib)):
                    to_remove.append(sib)
                else:
                    break

            # Scan AFTER table (up to 5 paragraphs — source, notes, dashes)
            for i in range(tbl_idx + 1, min(tbl_idx + 6, len(all_children))):
                sib = all_children[i]
                if sib.tag == f'{HP}p' and self._is_table_related(self._para_text(sib)):
                    to_remove.append(sib)
                else:
                    break

        for p in to_remove:
            container.remove(p)

        # Mark section dirty
        self._doc._dirty_sections.add(section_file)
        self._dom_dirty = True

    # --- Paragraph removal ---

    def remove_paragraph(self, containing, remove_all=False):
        """Remove paragraph(s) containing the given text.

        Like str_replace, enforces uniqueness by default: the text must
        appear in exactly ONE paragraph. If it appears in multiple
        paragraphs, raises TextNotFoundError with the count so the caller
        can provide a more specific string.

        Pass remove_all=True to remove every matching paragraph without
        uniqueness checks.

        Args:
            containing: Text to search for in paragraph content.
            remove_all: If True, remove all matching paragraphs.

        Returns:
            int: Number of paragraphs removed.

        Raises:
            TextNotFoundError: If not found, or found multiple times
                               (when remove_all=False).
        """
        matches = [p for p in self._doc.paragraphs
                   if p.text and containing in p.text]

        if len(matches) == 0:
            raise TextNotFoundError(
                f"No paragraph containing: '{containing[:50]}'")

        if len(matches) > 1 and not remove_all:
            raise TextNotFoundError(
                f"Text found in {len(matches)} paragraphs. "
                f"Use a longer, more unique string, or pass remove_all=True.")

        for para in matches:
            self._doc._dirty_sections.add(para._section)
            self._doc.remove_paragraph(para)
        self._dom_dirty = True
        return len(matches)

    # --- Text extraction ---

    def extract_text(self, format="text"):
        if format == "markdown":
            return self._doc.export_markdown()
        return self._doc.export_text()

    # --- Batch operations ---

    def apply_operations(self, operations):
        applied = []
        for op in operations:
            op_type = op.get("op", "")
            try:
                if op_type == "replace":
                    self.replace_text(op["find"], op["replace"], match=op.get("match", "first"))
                    applied.append(f"replace: '{op['find']}' -> '{op['replace']}'")
                elif op_type == "insert_after":
                    self.insert_paragraph(op.get("text"), after=op.get("anchor"),
                                          style=op.get("style", "body"),
                                          parts=op.get("parts"))
                    applied.append(f"insert after '{op.get('anchor', 'end')}'")
                elif op_type == "remove":
                    self.remove_paragraph(op["containing"])
                    applied.append(f"remove: '{op['containing']}'")
                elif op_type == "table_set_cell":
                    self.set_cell(op["table_index"], op["row"], op["col"], op["text"])
                    applied.append(f"table[{op['table_index']}] cell({op['row']},{op['col']})")
                elif op_type == "table_add_row":
                    self.add_row(op["table_index"], op["cells"])
                    applied.append(f"table[{op['table_index']}] add row")
                elif op_type == "table_replace":
                    result = self.replace_in_table(
                        op["table_index"], op["find"], op["replace"],
                        replace_all=op.get("replace_all", False),
                    )
                    applied.append(f"table[{op['table_index']}] replace '{op['find']}' -> '{op['replace']}' ({'ok' if result else 'not found'})")
                elif op_type == "add_footnote":
                    from hwpx_engine.elements import add_footnote_to_doc
                    add_footnote_to_doc(self._doc, op["anchor"], op["note"])
                    self._dom_dirty = True
                    applied.append(f"footnote on '{op['anchor']}'")
                elif op_type == "set_page_number":
                    from hwpx_engine.elements import set_page_number
                    set_page_number(self._doc, op.get("position", "footer_center"))
                    self._dom_dirty = True
                    applied.append(f"page number: {op.get('position')}")
                else:
                    applied.append(f"UNKNOWN op: {op_type}")
            except TextNotFoundError as e:
                applied.append(f"FAILED {op_type}: {e}")
            except Exception as e:
                applied.append(f"ERROR {op_type}: {e}")
        return applied

    # --- Save ---

    def save(self, output_path, auto_fix=True):
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        # Flush any pending DOM changes first
        if self._dom_dirty:
            self._flush_dom()
        if self._zip_modified:
            # ZIP was modified directly (parts/charPr).
            import shutil
            if os.path.abspath(str(path)) != os.path.abspath(self._source_path):
                shutil.copy2(self._source_path, str(path))
        else:
            self._doc.save_to_path(str(path))
        if auto_fix:
            fix_namespaces(str(path))
        result = HwpxValidator.validate(str(path), auto_fix=auto_fix, reference_path=self._source_path)
        return {
            "success": result.valid,
            "output_path": str(path),
            "validation": {
                "valid": result.valid,
                "level1_passed": result.level1_passed,
                "level2_passed": result.level2_passed,
                "level3_passed": result.level3_passed,
                "auto_fixed": result.auto_fixed,
                "warnings": result.warnings,
                "errors": result.errors,
            },
            "stats": result.stats,
        }

    @property
    def doc(self):
        return self._doc


def _find_toplevel_anchor(xml_text, anchor):
    """Find anchor text that is NOT inside a <hp:tc> element.

    Returns the position of the anchor, or -1 if not found at top-level.
    """
    start = 0
    while True:
        pos = xml_text.find(anchor, start)
        if pos < 0:
            return -1
        # Check if this position is inside a <hp:tc>...</hp:tc>
        preceding = xml_text[:pos]
        tc_opens = preceding.count('<hp:tc')
        tc_closes = preceding.count('</hp:tc>')
        if tc_opens <= tc_closes:
            # Not inside a tc — this is a valid top-level match
            return pos
        # Inside a tc — try next occurrence
        start = pos + 1


def _find_toplevel_p_end(xml_text, anchor_pos):
    """Find the end of the top-level <hp:p> that contains anchor_pos.

    If anchor is inside a table (<hp:tbl>), walk past the closing </hp:tbl>,
    then find the enclosing </hp:p> so insertion happens OUTSIDE the table.

    HWPX table structure:
      <hp:p>          <- top-level paragraph
        <hp:run>
          <hp:tbl>    <- table
            <hp:tr><hp:tc><hp:subList>
              <hp:p>anchor text here</hp:p>  <- inner paragraph
            </hp:subList></hp:tc></hp:tr>
          </hp:tbl>
        </hp:run>
      </hp:p>         <- insert AFTER this
    """
    import re

    # Check if anchor is inside a table
    tbl_opens = len(re.findall(r'<hp:tbl[ >]', xml_text[:anchor_pos]))
    tbl_closes = len(re.findall(r'</hp:tbl>', xml_text[:anchor_pos]))
    inside_table = tbl_opens > tbl_closes

    if inside_table:
        # Walk forward from anchor, tracking <hp:tbl> nesting until depth=0
        depth = tbl_opens - tbl_closes
        pos = anchor_pos
        while depth > 0 and pos < len(xml_text):
            next_open = xml_text.find('<hp:tbl', pos)
            next_close = xml_text.find('</hp:tbl>', pos)
            if next_close < 0:
                break
            if next_open >= 0 and next_open < next_close:
                depth += 1
                pos = next_open + 8
            else:
                depth -= 1
                pos = next_close + 10  # len('</hp:tbl>')
        # Now past all </hp:tbl> — find the closing </hp:p> of the top-level paragraph
        close_p = xml_text.find('</hp:p>', pos)
        if close_p >= 0:
            return close_p + len('</hp:p>')
    else:
        # Not inside table — simple case
        close_p = xml_text.find('</hp:p>', anchor_pos)
        if close_p >= 0:
            return close_p + len('</hp:p>')

    return -1
