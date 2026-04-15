"""Pure XML building blocks for HWPX documents.

All functions are stateless and template-agnostic. They create XML elements
with the correct HWPML namespace structure but know nothing about specific
templates or their style configurations.

Template-specific logic (caption sRef formulas, bracket rules, borderFill
lookups) belongs in per-template modules, not here.
"""
from lxml import etree

HP = 'http://www.hancom.co.kr/hwpml/2011/paragraph'
HC = 'http://www.hancom.co.kr/hwpml/2011/core'
HH = 'http://www.hancom.co.kr/hwpml/2011/head'
HS = 'http://www.hancom.co.kr/hwpml/2011/section'

_id_counter = 900000


def next_id() -> str:
    """Generate a unique ID for HWPX elements."""
    global _id_counter
    _id_counter += 1
    return str(_id_counter)


def reset_id(start: int = 900000):
    """Reset the ID counter (call at start of each build session)."""
    global _id_counter
    _id_counter = start


def has_part_overrides(part: dict) -> bool:
    """Check if a parts dict has any formatting overrides (bold/italic/color/size)."""
    return any(k in part for k in ('bold', 'italic', 'color', 'size'))


# ─── Run / Paragraph ─────────────────────────────────────────────


def make_run(text: str, char_pr: str) -> etree._Element:
    """Create a single <hp:run> element with text."""
    run = etree.Element(f'{{{HP}}}run')
    run.set('charPrIDRef', str(char_pr))
    t = etree.SubElement(run, f'{{{HP}}}t')
    t.text = text
    return run


def make_para(text: str = None, char_pr: str = '0', para_pr: str = '0',
              style_ref: str = '0', parts: list = None,
              charpr_resolver=None) -> etree._Element:
    """Create <hp:p> with correct styleIDRef + paraPrIDRef + charPrIDRef.

    Simple: make_para('text', cPr, pPr, sRef)
    Rich:   make_para(parts=[{'text':'normal'},{'text':'bold','bold':True}],
                      char_pr=base_cPr, para_pr=pPr, style_ref=sRef,
                      charpr_resolver=fn)

    charpr_resolver: callable(base_cpr_id, part_dict) -> new_cpr_id
        Used to create/find charPr with bold/italic/color/size overrides.
        If None, parts overrides are ignored and base char_pr is used.
    """
    p = etree.Element(f'{{{HP}}}p')
    p.set('id', next_id())
    p.set('paraPrIDRef', str(para_pr))
    p.set('styleIDRef', str(style_ref))
    p.set('pageBreak', '0')
    p.set('columnBreak', '0')
    p.set('merged', '0')

    if parts is not None:
        for part in parts:
            part_text = part.get('text', '')
            if has_part_overrides(part) and charpr_resolver is not None:
                cpr_id = str(charpr_resolver(char_pr, part))
            else:
                cpr_id = str(char_pr)
            run = etree.SubElement(p, f'{{{HP}}}run')
            run.set('charPrIDRef', cpr_id)
            t = etree.SubElement(run, f'{{{HP}}}t')
            t.text = part_text
    else:
        run = etree.SubElement(p, f'{{{HP}}}run')
        run.set('charPrIDRef', str(char_pr))
        t = etree.SubElement(run, f'{{{HP}}}t')
        t.text = text or ''

    add_linesegarray(p)
    return p


def make_two_run_para(text1: str, text2: str, cpr1: str, cpr2: str,
                      para_pr: str = '0', style_ref: str = '0') -> etree._Element:
    """Create <hp:p> with two runs (e.g., '(1)' + ' title')."""
    p = etree.Element(f'{{{HP}}}p')
    p.set('id', next_id())
    p.set('paraPrIDRef', str(para_pr))
    p.set('styleIDRef', str(style_ref))
    p.set('pageBreak', '0')
    p.set('columnBreak', '0')
    p.set('merged', '0')

    run1 = etree.SubElement(p, f'{{{HP}}}run')
    run1.set('charPrIDRef', str(cpr1))
    t1 = etree.SubElement(run1, f'{{{HP}}}t')
    t1.text = text1

    run2 = etree.SubElement(p, f'{{{HP}}}run')
    run2.set('charPrIDRef', str(cpr2))
    t2 = etree.SubElement(run2, f'{{{HP}}}t')
    t2.text = text2

    add_linesegarray(p)
    return p


def add_linesegarray(p: etree._Element) -> None:
    """Add a default linesegarray to a paragraph (required by Hangul for rendering)."""
    lsa = etree.SubElement(p, f'{{{HP}}}linesegarray')
    ls = etree.SubElement(lsa, f'{{{HP}}}lineseg')
    ls.set('textpos', '0')
    ls.set('vertpos', '0')
    ls.set('vertsize', '1200')
    ls.set('textheight', '1200')
    ls.set('baseline', '1020')
    ls.set('spacing', '840')
    ls.set('horzpos', '0')
    ls.set('horzsize', '46488')
    ls.set('flags', '393216')


# ─── Table ────────────────────────────────────────────────────────


def make_table_xml(col_count: int, rows_data: list, col_widths: list,
                   header_style: dict, body_style: dict,
                   page_width: int,
                   caption_paras: list = None,
                   merges: list = None) -> etree._Element:
    """Create <hp:p> containing a <hp:tbl>.

    Args:
        col_count: number of columns
        rows_data: list of (row_type, cells) where row_type is 'header' or 'data'
        col_widths: list of column widths in HWPML units
        header_style: dict with keys 'pPr', 'sRef', 'cPr', 'borderFill', 'borderFill_right'
        body_style: dict with keys 'pPr', 'sRef', 'cPr', 'borderFill', 'borderFill_right'
        page_width: total table width in HWPML units
        caption_paras: list of pre-built <hp:p> elements for the caption subList (or None)
        merges: list of {'row':, 'col':, 'rowspan':, 'colspan':} dicts
    """
    row_count = len(rows_data)

    p = etree.Element(f'{{{HP}}}p')
    p.set('id', next_id())
    p.set('paraPrIDRef', '1')
    p.set('styleIDRef', '0')
    p.set('pageBreak', '0')
    p.set('columnBreak', '0')
    p.set('merged', '0')

    run = etree.SubElement(p, f'{{{HP}}}run')
    run.set('charPrIDRef', '0')

    tbl = etree.SubElement(run, f'{{{HP}}}tbl')
    tbl.set('id', next_id())
    tbl.set('zOrder', '0')
    tbl.set('numberingType', 'TABLE')
    tbl.set('textWrap', 'TOP_AND_BOTTOM')
    tbl.set('textFlow', 'BOTH_SIDES')
    tbl.set('lock', '0')
    tbl.set('dropcapstyle', 'None')
    tbl.set('pageBreak', 'CELL')
    tbl.set('repeatHeader', '1')
    tbl.set('rowCnt', str(row_count))
    tbl.set('colCnt', str(col_count))
    tbl.set('cellSpacing', '0')
    tbl.set('borderFillIDRef', '2')
    tbl.set('noAdjust', '0')

    total_height = row_count * 1200
    sz = etree.SubElement(tbl, f'{{{HP}}}sz')
    sz.set('width', str(page_width))
    sz.set('widthRelTo', 'ABSOLUTE')
    sz.set('height', str(total_height))
    sz.set('heightRelTo', 'ABSOLUTE')
    sz.set('protect', '0')

    pos = etree.SubElement(tbl, f'{{{HP}}}pos')
    pos.set('treatAsChar', '1')
    pos.set('affectLSpacing', '0')
    pos.set('flowWithText', '1')
    pos.set('allowOverlap', '0')
    pos.set('holdAnchorAndSO', '0')
    pos.set('vertRelTo', 'PARA')
    pos.set('horzRelTo', 'PARA')
    pos.set('vertAlign', 'TOP')
    pos.set('horzAlign', 'LEFT')
    pos.set('vertOffset', '0')
    pos.set('horzOffset', '0')

    # Caption (pre-built paragraphs from template module)
    if caption_paras:
        cap_el = etree.SubElement(tbl, f'{{{HP}}}caption')
        cap_el.set('side', 'TOP')
        cap_el.set('fullSz', '0')
        cap_el.set('width', '8504')
        cap_el.set('gap', '850')
        cap_el.set('lastWidth', str(page_width))
        cap_sub = etree.SubElement(cap_el, f'{{{HP}}}subList')
        cap_sub.set('id', '')
        cap_sub.set('textDirection', 'HORIZONTAL')
        cap_sub.set('lineWrap', 'BREAK')
        cap_sub.set('vertAlign', 'TOP')
        cap_sub.set('linkListIDRef', '0')
        cap_sub.set('linkListNextIDRef', '0')
        cap_sub.set('textWidth', '0')
        cap_sub.set('textHeight', '0')
        cap_sub.set('hasTextRef', '0')
        cap_sub.set('hasNumRef', '0')
        for cp in caption_paras:
            cap_sub.append(cp)

    # Margins
    for margin_name in ['outMargin', 'inMargin']:
        m = etree.SubElement(tbl, f'{{{HP}}}{margin_name}')
        val = '141' if margin_name == 'outMargin' else '566'
        m.set('left', val)
        m.set('right', val)
        m.set('top', val)
        m.set('bottom', val)

    # Merge map
    hidden = set()
    merge_master = {}
    if merges:
        for mg in merges:
            mr, mc = mg.get('row', 0), mg.get('col', 0)
            mrs, mcs = mg.get('rowspan', 1), mg.get('colspan', 1)
            merge_master[(mr, mc)] = (mrs, mcs)
            for dr in range(mrs):
                for dc in range(mcs):
                    if dr == 0 and dc == 0:
                        continue
                    hidden.add((mr + dr, mc + dc))

    # Build rows
    bf_hdr_l = header_style.get('borderFill', '17')
    bf_hdr_r = header_style.get('borderFill_right', header_style.get('borderFill', '18'))
    bf_body_l = body_style.get('borderFill', '16')
    bf_body_r = body_style.get('borderFill_right', body_style.get('borderFill', '15'))

    for row_idx, (row_type, cells_data) in enumerate(rows_data):
        tr = etree.SubElement(tbl, f'{{{HP}}}tr')

        for col_idx in range(col_count):
            if (row_idx, col_idx) in hidden:
                continue

            cell_text = cells_data[col_idx] if col_idx < len(cells_data) else ''
            rs, cs = merge_master.get((row_idx, col_idx), (1, 1))

            tc = etree.SubElement(tr, f'{{{HP}}}tc')
            tc.set('name', '')
            tc.set('header', '1' if row_type == 'header' else '0')
            tc.set('hasMargin', '0')
            tc.set('protect', '0')
            tc.set('editable', '0')
            tc.set('dirty', '0')

            last_visible_col = col_idx + cs - 1
            is_last_col = (last_visible_col >= col_count - 1)
            if row_type == 'header':
                bf = bf_hdr_r if is_last_col else bf_hdr_l
            else:
                bf = bf_body_r if is_last_col else bf_body_l
            tc.set('borderFillIDRef', str(bf))

            addr = etree.SubElement(tc, f'{{{HP}}}cellAddr')
            addr.set('colAddr', str(col_idx))
            addr.set('rowAddr', str(row_idx))

            span_el = etree.SubElement(tc, f'{{{HP}}}cellSpan')
            span_el.set('colSpan', str(cs))
            span_el.set('rowSpan', str(rs))

            cell_w = sum(col_widths[col_idx:col_idx + cs])
            csz = etree.SubElement(tc, f'{{{HP}}}cellSz')
            csz.set('width', str(cell_w))
            csz.set('height', '1200')

            s_cell = header_style if row_type == 'header' else body_style
            sub = etree.SubElement(tc, f'{{{HP}}}subList')
            sub.set('id', '')
            sub.set('textDirection', 'HORIZONTAL')
            sub.set('lineWrap', 'BREAK')
            sub.set('vertAlign', s_cell.get('vertAlign', 'CENTER'))
            sub.set('linkListIDRef', '0')
            sub.set('linkListNextIDRef', '0')
            sub.set('textWidth', '0')
            sub.set('textHeight', '0')
            sub.set('hasTextRef', '0')
            sub.set('hasNumRef', '0')

            cp = etree.SubElement(sub, f'{{{HP}}}p')
            cp.set('id', next_id())
            cp.set('paraPrIDRef', s_cell['pPr'])
            cp.set('styleIDRef', s_cell['sRef'])
            cp.set('pageBreak', '0')
            cp.set('columnBreak', '0')
            cp.set('merged', '0')

            cr = etree.SubElement(cp, f'{{{HP}}}run')
            cr.set('charPrIDRef', s_cell['cPr'])
            ct = etree.SubElement(cr, f'{{{HP}}}t')
            ct.text = str(cell_text)

    add_linesegarray(p)
    return p


# ─── Figure / Image ──────────────────────────────────────────────


def make_figure_box(width: int, height: int, border_fill: str,
                    para_pr: str = '11', style_ref: str = '0',
                    char_pr: str = '4') -> etree._Element:
    """Create empty figure box (1x1 table, treatAsChar).

    Args:
        para_pr: paraPrIDRef for the wrapping paragraph (use CENTER-aligned for centering)
        style_ref: styleIDRef for the wrapping paragraph
        char_pr: charPrIDRef for the run containing the table
    """
    p = etree.Element(f'{{{HP}}}p')
    p.set('id', next_id())
    p.set('paraPrIDRef', str(para_pr))
    p.set('styleIDRef', str(style_ref))
    p.set('pageBreak', '0')
    p.set('columnBreak', '0')
    p.set('merged', '0')

    run = etree.SubElement(p, f'{{{HP}}}run')
    run.set('charPrIDRef', str(char_pr))

    tbl = etree.SubElement(run, f'{{{HP}}}tbl')
    tbl.set('id', next_id())
    tbl.set('rowCnt', '1')
    tbl.set('colCnt', '1')
    tbl.set('cellSpacing', '0')
    tbl.set('borderFillIDRef', '2')
    tbl.set('zOrder', '0')
    tbl.set('textWrap', 'TOP_AND_BOTTOM')
    tbl.set('textFlow', 'BOTH_SIDES')
    tbl.set('noAdjust', '0')

    sz = etree.SubElement(tbl, f'{{{HP}}}sz')
    sz.set('width', str(width))
    sz.set('widthRelTo', 'ABSOLUTE')
    sz.set('height', str(height))
    sz.set('heightRelTo', 'ABSOLUTE')
    sz.set('protect', '0')

    pos = etree.SubElement(tbl, f'{{{HP}}}pos')
    for k, v in [('treatAsChar', '1'), ('affectLSpacing', '0'),
                 ('flowWithText', '1'), ('allowOverlap', '0'),
                 ('holdAnchorAndSO', '0'), ('vertRelTo', 'PARA'),
                 ('horzRelTo', 'PARA'), ('vertAlign', 'TOP'),
                 ('horzAlign', 'LEFT'), ('vertOffset', '0'),
                 ('horzOffset', '0')]:
        pos.set(k, v)

    for mn in ['outMargin', 'inMargin']:
        m = etree.SubElement(tbl, f'{{{HP}}}{mn}')
        m.set('left', '0'); m.set('right', '0')
        m.set('top', '0'); m.set('bottom', '0')

    tr = etree.SubElement(tbl, f'{{{HP}}}tr')
    tc = etree.SubElement(tr, f'{{{HP}}}tc')
    tc.set('borderFillIDRef', str(border_fill))
    tc.set('name', ''); tc.set('header', '0'); tc.set('hasMargin', '0')
    tc.set('protect', '0'); tc.set('editable', '0'); tc.set('dirty', '0')

    addr = etree.SubElement(tc, f'{{{HP}}}cellAddr')
    addr.set('colAddr', '0'); addr.set('rowAddr', '0')
    span = etree.SubElement(tc, f'{{{HP}}}cellSpan')
    span.set('colSpan', '1'); span.set('rowSpan', '1')
    csz = etree.SubElement(tc, f'{{{HP}}}cellSz')
    csz.set('width', str(width)); csz.set('height', str(height))

    sub = etree.SubElement(tc, f'{{{HP}}}subList')
    sub.set('id', ''); sub.set('textDirection', 'HORIZONTAL')
    sub.set('lineWrap', 'BREAK'); sub.set('vertAlign', 'CENTER')
    sub.set('linkListIDRef', '0'); sub.set('linkListNextIDRef', '0')
    sub.set('textWidth', '0'); sub.set('textHeight', '0')
    sub.set('hasTextRef', '0'); sub.set('hasNumRef', '0')
    cp = etree.SubElement(sub, f'{{{HP}}}p')
    cp.set('id', next_id())
    cp.set('paraPrIDRef', '1'); cp.set('styleIDRef', '0')
    cp.set('pageBreak', '0'); cp.set('columnBreak', '0'); cp.set('merged', '0')
    cr = etree.SubElement(cp, f'{{{HP}}}run')
    cr.set('charPrIDRef', '4')
    ct = etree.SubElement(cr, f'{{{HP}}}t')
    ct.text = ''

    add_linesegarray(p)
    return p


def make_image_pic(bin_id: str, org_w: int, org_h: int,
                   cur_w: int, cur_h: int,
                   figure_box_width: int = 34016,
                   figure_box_bf: str = '20') -> etree._Element:
    """Create <hp:p> containing figure box with inline <hp:pic>.

    Uses raw XML string for correct hc: namespace on renderingInfo, img, pt elements.
    """
    xml = f'''<hp:p xmlns:hp="{HP}" xmlns:hc="{HC}"
      id="{next_id()}" paraPrIDRef="11" styleIDRef="0"
      pageBreak="0" columnBreak="0" merged="0">
      <hp:run charPrIDRef="4">
        <hp:tbl id="{next_id()}" rowCnt="1" colCnt="1"
          cellSpacing="0" borderFillIDRef="2" zOrder="0"
          textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" noAdjust="0">
          <hp:sz width="{figure_box_width}" widthRelTo="ABSOLUTE"
            height="{cur_h}" heightRelTo="ABSOLUTE" protect="0"/>
          <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1"
            allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA"
            horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT"
            vertOffset="0" horzOffset="0"/>
          <hp:outMargin left="0" right="0" top="0" bottom="0"/>
          <hp:inMargin left="0" right="0" top="0" bottom="0"/>
          <hp:tr>
            <hp:tc borderFillIDRef="{figure_box_bf}" name="" header="0"
              hasMargin="0" protect="0" editable="0" dirty="0">
              <hp:cellAddr colAddr="0" rowAddr="0"/>
              <hp:cellSpan colSpan="1" rowSpan="1"/>
              <hp:cellSz width="{figure_box_width}" height="{cur_h}"/>
              <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK"
                vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0"
                textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
                <hp:p id="{next_id()}" paraPrIDRef="1" styleIDRef="0"
                  pageBreak="0" columnBreak="0" merged="0">
                  <hp:run charPrIDRef="4">
                    <hp:pic id="{next_id()}" zOrder="0"
                      numberingType="PICTURE" textWrap="TOP_AND_BOTTOM"
                      textFlow="BOTH_SIDES" lock="0" dropcapstyle="None"
                      href="" groupLevel="0" instid="{next_id()}"
                      reverse="0">
                      <hp:offset x="0" y="0"/>
                      <hp:orgSz width="{org_w}" height="{org_h}"/>
                      <hp:curSz width="{cur_w}" height="{cur_h}"/>
                      <hp:flip horizontal="0" vertical="0"/>
                      <hp:rotationInfo angle="0" centerX="{cur_w//2}"
                        centerY="{cur_h//2}" rotateimage="1"/>
                      <hp:renderingInfo>
                        <hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>
                        <hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>
                        <hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>
                      </hp:renderingInfo>
                      <hc:img binaryItemIDRef="{bin_id}" bright="0"
                        contrast="0" effect="REAL_PIC" alpha="0"/>
                      <hp:imgRect>
                        <hc:pt0 x="0" y="0"/>
                        <hc:pt1 x="{org_w}" y="0"/>
                        <hc:pt2 x="{org_w}" y="{org_h}"/>
                        <hc:pt3 x="0" y="{org_h}"/>
                      </hp:imgRect>
                      <hp:imgClip left="0" right="{org_w}" top="0" bottom="{org_h}"/>
                      <hp:inMargin left="0" right="0" top="0" bottom="0"/>
                      <hp:imgDim dimwidth="{org_w}" dimheight="{org_h}"/>
                      <hp:effects/>
                      <hp:sz width="{cur_w}" widthRelTo="ABSOLUTE"
                        height="{cur_h}" heightRelTo="ABSOLUTE" protect="0"/>
                      <hp:pos treatAsChar="1" affectLSpacing="0"
                        flowWithText="1" allowOverlap="0"
                        holdAnchorAndSO="0" vertRelTo="PARA"
                        horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT"
                        vertOffset="0" horzOffset="0"/>
                      <hp:outMargin left="0" right="0" top="0" bottom="0"/>
                      <hp:shapeComment/>
                    </hp:pic>
                  </hp:run>
                </hp:p>
              </hp:subList>
            </hp:tc>
          </hp:tr>
        </hp:tbl>
      </hp:run>
      <hp:linesegarray>
        <hp:lineseg textpos="0" vertpos="0" vertsize="1200"
          textheight="1200" baseline="1020" spacing="840"
          horzpos="0" horzsize="46488" flags="393216"/>
      </hp:linesegarray>
    </hp:p>'''
    return etree.fromstring(xml)


# ─── Utility ──────────────────────────────────────────────────────


def xml_escape(text: str) -> str:
    """Escape XML special characters."""
    return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;')
                .replace("'", '&apos;'))


def get_para_text(para: etree._Element) -> str:
    """Get combined text from all runs in a paragraph."""
    texts = []
    for t in para.iter(f'{{{HP}}}t'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)


def set_para_text(para: etree._Element, new_text: str) -> None:
    """Set text in first run, clear extra runs."""
    runs = para.findall(f'{{{HP}}}run')
    if not runs:
        return
    first_run = runs[0]
    t_elements = first_run.findall(f'{{{HP}}}t')
    if t_elements:
        t_elements[0].text = new_text
        for extra_t in t_elements[1:]:
            first_run.remove(extra_t)
    else:
        t = etree.SubElement(first_run, f'{{{HP}}}t')
        t.text = new_text
    for extra_run in runs[1:]:
        para.remove(extra_run)
