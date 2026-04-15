"""CharPr/BorderFill management for HWPX header.xml.

Provides a single class that handles:
- charPr clone/reuse with bold/italic/color/size overrides
- borderFill clone with background color (cell shading)
- Correct itemCnt maintenance
- Correct namespace usage (hc: for fillBrush/winBrush)

Used by both build.py (build orchestrator) and editor.py (HwpxEditor).
"""
import copy
from lxml import etree

HH = '{http://www.hancom.co.kr/hwpml/2011/head}'
HC = '{http://www.hancom.co.kr/hwpml/2011/core}'

# Standard child element order inside <charPr>
_CHARPR_CHILD_ORDER = [
    'fontRef', 'ratio', 'spacing', 'relSz', 'offset',
    'bold', 'italic', 'underline', 'strikeout', 'outline', 'shadow',
]


class CharPrManager:
    """Manage charPr and borderFill entries in header.xml.

    Keeps the header.xml root in memory during a build/edit session.
    Tracks whether any modifications were made so the caller knows
    whether to write back the header.
    """

    def __init__(self, header_root: etree._Element):
        self._root = header_root
        self._modified = False

    @property
    def modified(self) -> bool:
        return self._modified

    @property
    def root(self) -> etree._Element:
        return self._root

    def serialize(self) -> bytes:
        """Return modified header.xml as bytes."""
        return etree.tostring(self._root, xml_declaration=True, encoding='UTF-8')

    # ─── charPr ───────────────────────────────────────────────

    def find_or_create_charpr(self, base_id: str,
                              bold: bool = False, italic: bool = False,
                              color: str = None, size: int = None) -> str:
        """Find existing charPr matching overrides, or clone and create new one.

        Args:
            base_id: ID of the base charPr to clone from
            bold: apply bold
            italic: apply italic
            color: text color as '#RRGGBB'
            size: font size in 0.1pt units (e.g., 1100 = 11pt)

        Returns:
            charPr ID (str) — either existing match or newly created
        """
        if not bold and not italic and color is None and size is None:
            return str(base_id)

        char_props = self._root.findall(f'.//{HH}charPr')

        # Find base charPr
        base_cp = None
        for cp in char_props:
            if cp.get('id') == str(base_id):
                base_cp = cp
                break
        if base_cp is None:
            return str(base_id)

        # Target properties
        target_color = color or base_cp.get('textColor', '#000000')
        target_height = str(size) if size else base_cp.get('height', '1100')

        # Check for existing match
        for cp in char_props:
            if self._charpr_matches(cp, base_cp, target_color, target_height,
                                    bold, italic):
                return cp.get('id')

        # Clone and create
        new_cp = copy.deepcopy(base_cp)
        new_id = max(int(cp.get('id', 0)) for cp in char_props) + 1
        new_cp.set('id', str(new_id))

        if color:
            new_cp.set('textColor', color)
        if size:
            new_cp.set('height', str(size))
        if bold and new_cp.find(f'{HH}bold') is None:
            _insert_charpr_child(new_cp, 'bold')
        if italic and new_cp.find(f'{HH}italic') is None:
            _insert_charpr_child(new_cp, 'italic')

        charpr_list = base_cp.getparent()
        charpr_list.append(new_cp)
        charpr_list.set('itemCnt', str(len(charpr_list.findall(f'{HH}charPr'))))
        self._modified = True

        return str(new_id)

    def find_or_create_charpr_from_part(self, base_id: str, part: dict) -> str:
        """Convenience wrapper: extract overrides from a parts dict."""
        return self.find_or_create_charpr(
            base_id,
            bold=part.get('bold', False),
            italic=part.get('italic', False),
            color=part.get('color'),
            size=part.get('size'),
        )

    def _charpr_matches(self, candidate, base_cp, target_color: str,
                        target_height: str, want_bold: bool,
                        want_italic: bool) -> bool:
        """Check if candidate charPr matches desired properties."""
        if candidate.get('textColor') != target_color:
            return False
        if candidate.get('height') != target_height:
            return False

        # Must match bold/italic presence
        has_bold = candidate.find(f'{HH}bold') is not None
        has_italic = candidate.find(f'{HH}italic') is not None
        if has_bold != want_bold or has_italic != want_italic:
            return False

        # Must match fontRef
        base_font = base_cp.find(f'{HH}fontRef')
        cand_font = candidate.find(f'{HH}fontRef')
        if base_font is not None and cand_font is not None:
            if dict(base_font.attrib) != dict(cand_font.attrib):
                return False

        # Must match spacing
        base_spacing = base_cp.find(f'{HH}spacing')
        cand_spacing = candidate.find(f'{HH}spacing')
        if base_spacing is not None and cand_spacing is not None:
            if dict(base_spacing.attrib) != dict(cand_spacing.attrib):
                return False

        return True

    # ─── borderFill ───────────────────────────────────────────

    def create_shaded_border_fill(self, base_id: str, face_color: str) -> str:
        """Clone a borderFill and add background shading.

        IMPORTANT: fillBrush and winBrush use hc: namespace, not hh:.

        Args:
            base_id: ID of the existing borderFill to clone
            face_color: background color as '#RRGGBB' (e.g., '#FFE699')

        Returns:
            New borderFill ID (str)
        """
        border_fills = self._root.findall(f'.//{HH}borderFill')

        base_bf = None
        for bf in border_fills:
            if bf.get('id') == str(base_id):
                base_bf = bf
                break
        if base_bf is None:
            return str(base_id)

        new_bf = copy.deepcopy(base_bf)
        new_id = max(int(bf.get('id', 0)) for bf in border_fills) + 1
        new_bf.set('id', str(new_id))

        # Remove existing fill elements (if any)
        for tag in ['fillBrush', 'winBrush']:
            for old in new_bf.findall(f'{HC}{tag}'):
                new_bf.remove(old)
            # Also check wrong namespace (cleanup)
            for old in new_bf.findall(f'{HH}{tag}'):
                new_bf.remove(old)

        # Add fillBrush with hc: namespace (CRITICAL: not hh:)
        fill_brush = etree.SubElement(new_bf, f'{HC}fillBrush')
        win_brush = etree.SubElement(fill_brush, f'{HC}winBrush')
        win_brush.set('faceColor', face_color)
        win_brush.set('hatchColor', '#000000')
        win_brush.set('alpha', '0')

        bf_list = base_bf.getparent()
        bf_list.append(new_bf)
        bf_list.set('itemCnt', str(len(bf_list.findall(f'{HH}borderFill'))))
        self._modified = True

        return str(new_id)


# ─── Module-level helpers ─────────────────────────────────────────


def _insert_charpr_child(charpr: etree._Element, element_name: str) -> None:
    """Insert <bold/> or <italic/> at the correct position in charPr.

    Standard order: fontRef, ratio, spacing, relSz, offset, bold, italic, ...
    """
    target_idx = (_CHARPR_CHILD_ORDER.index(element_name)
                  if element_name in _CHARPR_CHILD_ORDER else -1)

    new_elem = etree.SubElement(charpr, f'{HH}{element_name}')

    if target_idx >= 0:
        insert_pos = 0
        for i, child in enumerate(charpr):
            child_name = child.tag.split('}')[-1]
            if child_name in _CHARPR_CHILD_ORDER:
                child_idx = _CHARPR_CHILD_ORDER.index(child_name)
                if child_idx < target_idx:
                    insert_pos = i + 1
        charpr.remove(new_elem)
        charpr.insert(insert_pos, new_elem)
