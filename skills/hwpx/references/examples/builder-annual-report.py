"""연간보고서 A4 — section pipeline 예시.

다중 섹션 구조: 표지, 목차, 본문 장×3, 참고문헌.
각 section 파일에 대해 어떤 모듈의 process()를 실행할지 선언한다.

handler 값은 modules/ 디렉토리 안의 파일명(문자열).
build.py가 로드 시 자동으로 modules/{handler}.py → process() 함수로 연결한다.
"""

sections = [
    # section0: 표지 텍스트 교체 + 요약 불릿 재생성
    {'file': 'section0.xml', 'handler': 'cover'},

    # section1: 목차 구조 유지, 항목만 교체
    {'file': 'section1.xml', 'handler': 'toc'},

    # section2~4: 본문 장. sec_num으로 content['chapters'][sec_num-2]에 매핑
    {'file': 'section2.xml', 'handler': 'body', 'sec_num': 2},
    {'file': 'section3.xml', 'handler': 'body', 'sec_num': 3},
    {'file': 'section4.xml', 'handler': 'body', 'sec_num': 4},

    # section5: 참고문헌 + Abstract + 키워드
    {'file': 'section5.xml', 'handler': 'references'},
]
