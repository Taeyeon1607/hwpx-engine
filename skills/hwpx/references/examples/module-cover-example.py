"""표지 처리 모듈 예시.

모든 모듈은 process(xml_data, content, ctx, **kwargs) → bytes 시그니처를 따른다.
ctx에서 ctx.styles, ctx.metadata, ctx.charpr_mgr를 사용한다.
"""
from lxml import etree
from hwpx_engine.xml_primitives import make_para, HP, get_para_text


def process(xml_data: bytes, content: dict, ctx, **kwargs) -> bytes:
    """표지 섹션 처리: 텍스트 교체 + 요약 불릿 재생성."""

    # Phase 1: 표지/판권 텍스트 교체 (raw string replacement)
    # 런 경계를 넘는 텍스트도 안전하게 교체됨
    replacements = content.get('cover_replacements', {})
    if replacements:
        text = xml_data.decode('utf-8')
        # 긴 키부터 매칭 (부분문자열 충돌 방지)
        for old, new in sorted(replacements.items(), key=lambda x: len(x[0]), reverse=True):
            text = text.replace(old, new)
        xml_data = text.encode('utf-8')

    # Phase 2: 요약/정책건의 불릿 재생성
    tree = etree.fromstring(xml_data)
    proposal_items = content.get('summary', [])  # 또는 'policy_proposal' 등 양식에 따라
    if proposal_items:
        _rebuild_summary(tree, proposal_items, ctx)

    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8')


def _rebuild_summary(sec_root, items, ctx):
    """기존 BULLET 문단 제거 후 새 요약 항목 삽입."""
    styles = ctx.styles
    all_paras = list(sec_root.findall(f'{{{HP}}}p'))

    # 양식에서 사용하는 불릿 스타일의 pPr/sRef 집합 (metadata.json에서 로드됨)
    bullet_pprs = {styles['bullet_1']['pPr'],
                   styles['bullet_2']['pPr'],
                   styles['bullet_3']['pPr']}
    bullet_srefs = {styles['bullet_1']['sRef'],
                    styles['bullet_2']['sRef'],
                    styles['bullet_3']['sRef']}

    # 헤더 테이블이 포함된 문단은 유지, 일반 불릿만 제거
    header_para = None
    to_remove = []
    for para in all_paras:
        ppr = para.get('paraPrIDRef', '0')
        sref = para.get('styleIDRef', '0')
        if ppr in bullet_pprs and sref in bullet_srefs:
            has_tbl = para.find(f'.//{{{HP}}}tbl') is not None
            if has_tbl:
                header_para = para  # 구조 요소 — 유지
            else:
                to_remove.append(para)

    for para in to_remove:
        try:
            sec_root.remove(para)
        except ValueError:
            pass

    # 새 항목 삽입
    if header_para is not None:
        anchor = header_para
        charpr_resolver = ctx.charpr_mgr.find_or_create_charpr_from_part
        for item in items:
            style = item.get('style', 'bullet_2')
            s = styles.get(style, styles['bullet_2'])
            new_para = make_para(
                item.get('text', ''), s['cPr'], s['pPr'], s['sRef'],
                parts=item.get('parts'),
                charpr_resolver=charpr_resolver,
            )
            anchor.addnext(new_para)
            anchor = new_para
