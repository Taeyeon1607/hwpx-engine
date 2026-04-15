"""Core utilities: namespace fixing and section extraction."""

import os
import re
import zipfile
from pathlib import Path

NS_DECL = (
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
    'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" '
    'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
    'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
    'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" '
    'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:opf="http://www.idpf.org/2007/opf/" '
    'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
    'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" '
    'xmlns:epub="http://www.idpf.org/2007/ops" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)

_EXPECTED_NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
    "hp10": "http://www.hancom.co.kr/hwpml/2016/paragraph",
}

def fix_namespaces(hwpx_path: str) -> list[str]:
    hwpx_path = str(hwpx_path)
    tmp = hwpx_path + ".ns_tmp"
    fixed = []
    with zipfile.ZipFile(hwpx_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("Contents/") and item.filename.endswith(".xml"):
                    text = data.decode("utf-8")
                    original = text
                    for prefix, uri in _EXPECTED_NS.items():
                        text = re.sub(
                            rf'xmlns:ns\d+="{re.escape(uri)}"',
                            f'xmlns:{prefix}="{uri}"',
                            text,
                        )
                    if "section" in item.filename:
                        for prefix, uri in _EXPECTED_NS.items():
                            decl = f'xmlns:{prefix}="{uri}"'
                            if decl not in text:
                                text = text.replace("<hs:sec ", f"<hs:sec {decl} ", 1)
                    if text != original:
                        fixed.append(item.filename)
                    data = text.encode("utf-8")
                if item.filename == "mimetype":
                    zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                else:
                    zout.writestr(item, data)
    os.replace(tmp, hwpx_path)
    return fixed

def extract_secpr(hwpx_path: str) -> tuple[str, str]:
    with zipfile.ZipFile(hwpx_path, "r") as z:
        data = z.read("Contents/section0.xml").decode("utf-8")
    secpr_match = re.search(r"<hp:secPr.*?</hp:secPr>", data, re.DOTALL)
    secpr = secpr_match.group() if secpr_match else ""
    end = secpr_match.end() if secpr_match else 0
    ctrl_match = re.search(r"<hp:ctrl>.*?</hp:ctrl>", data[end:end + 500], re.DOTALL)
    colpr = ctrl_match.group() if ctrl_match else ""
    return secpr, colpr
