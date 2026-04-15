"""이슈페이퍼 A4 — section pipeline 예시.

단일 본문 섹션 구조: 표지(section0) + 본문 전체(section1).
본문 안에서 장 구분은 1x1 테이블 + pageBreak로 구현.
"""

sections = [
    # section0: 표지 RECT drawText 텍스트 교체 + TOC 바이트패딩 교체
    {'file': 'section0.xml', 'handler': 'cover'},

    # section1: 본문 전체 (요약 + 장×N + 참고문헌 + 부록)
    {'file': 'section1.xml', 'handler': 'body_single'},
]
