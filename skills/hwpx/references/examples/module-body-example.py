"""본문 처리 모듈 예시 (다중 섹션 양식용).

각 본문 섹션은 간지(있는 경우) + 본문 내용으로 구성된다.
sec_num kwargs로 어떤 장인지 식별한다.
"""
import re
from lxml import etree
from hwpx_engine.xml_primitives import (
    make_para, make_two_run_para, make_table_xml, make_figure_box,
    make_image_pic, HP, next_id,
)

HP_NS = 'http://www.hancom.co.kr/hwpml/2011/paragraph'
NS = {'hp': HP_NS}


def process(xml_data: bytes, content: dict, ctx, sec_num: int = 2, **kwargs) -> bytes:
    """본문 섹션 처리: 간지 교체 + 본문 재생성."""
    chapters = content.get('chapters', [])
    chap_idx = sec_num - 2  # section2=chapter[0]

    if chap_idx >= len(chapters):
        return xml_data  # 이 섹션에 해당하는 장 없음

    chapter = chapters[chap_idx]
    tree = etree.fromstring(xml_data)
    styles = ctx.styles

    # Phase 1: 간지 교체 (양식에 간지가 있는 경우)
    # 간지 모듈을 별도 파일로 분리하여 호출하거나, 여기서 직접 처리
    # _replace_interleaf(tree, chapter, styles)

    # Phase 2: 본문 내용 교체
    # ... (기존 내용 제거 + 새 내용 삽입)

    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8')


# ─── 아이템 생성 헬퍼 ─────────────────────────────

def make_table_items(item: dict, chapter_num: int, ctx) -> list:
    """표 아이템 → 문단 리스트.

    양식별로 캡션 처리가 다를 수 있다:
    - 자동번호 양식: styleIDRef의 NUMBER 속성이 '[표 N-M]' 자동 생성 → 제목만 입력
    - 수동번호 양식: 대괄호 포함하여 직접 입력 → '[표제목]'
    캡션 sRef를 장 번호에 따라 조정하는 양식도 있음 (예: base_sRef - (chapter_num - 1))
    """
    styles = ctx.styles
    paras = []

    # 캡션에서 자동번호 접두사 제거 (사용자가 실수로 포함한 경우)
    caption_text = re.sub(r'^\[(?:표|그림)\s*\d+-\d+\]\s*', '', item.get('caption', ''))

    # 스타일 조회 (모든 값은 metadata.json에서)
    s_cap = styles.get('table_caption', {})
    s_hdr = styles.get('table_header_cell', {})
    s_body = styles.get('table_body_cell', {})

    headers = item.get('headers', [])
    rows = item.get('rows', [])
    page_width = int(ctx.metadata.get('table_page_width', '34016'))

    # ... (표 생성 로직)

    return paras


def make_figure_items(item: dict, chapter_num: int, ctx) -> list:
    """그림 아이템 → 캡션 + 빈 박스 + 출처 문단 리스트."""
    styles = ctx.styles
    paras = []

    # ... (그림 생성 로직)

    return paras
