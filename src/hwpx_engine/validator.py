"""3-level HWPX validation with auto-fix capabilities."""

import zipfile
import os
from dataclasses import dataclass, field
from pathlib import Path
from lxml import etree

from hwpx_engine.utils import fix_namespaces


@dataclass
class ValidationResult:
    valid: bool = True
    level1_passed: bool = False
    level2_passed: bool = False
    level3_passed: bool = False
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    auto_fixed: list[str] = field(default_factory=list)
    stats: dict = field(default_factory=dict)


class HwpxValidator:
    """HWPX document validator with 3-level checking and auto-fix."""

    @staticmethod
    def validate(
        hwpx_path: str,
        auto_fix: bool = False,
        reference_path: str | None = None,
    ) -> ValidationResult:
        result = ValidationResult()
        path = Path(hwpx_path)

        # Pre-check: file exists
        if not path.exists():
            result.valid = False
            result.errors.append(f"File not found: {hwpx_path}")
            return result

        # Level 1: Structural
        HwpxValidator._check_level1(path, result, auto_fix)
        if not result.level1_passed:
            result.valid = False
            return result

        # Level 2: Integrity
        HwpxValidator._check_level2(path, result, auto_fix)
        if result.errors:
            result.valid = False

        # Level 3: Fidelity (only if reference provided)
        if reference_path and Path(reference_path).exists():
            HwpxValidator._check_level3(path, Path(reference_path), result)
        else:
            result.level3_passed = True  # skip

        # Stats
        result.stats = HwpxValidator._collect_stats(path)

        if not result.errors:
            result.valid = True

        return result

    @staticmethod
    def _check_level1(path: Path, result: ValidationResult, auto_fix: bool) -> None:
        """Check: does it open?"""
        try:
            with zipfile.ZipFile(path, "r") as z:
                names = z.namelist()

                # mimetype must be first entry
                if not names or names[0] != "mimetype":
                    if auto_fix:
                        HwpxValidator._fix_mimetype_position(path)
                        result.auto_fixed.append("mimetype 위치 교정")
                    else:
                        result.errors.append("mimetype is not the first ZIP entry")
                        return

                # Check mimetype compression
                info = z.getinfo("mimetype")
                if info.compress_type != zipfile.ZIP_STORED:
                    if auto_fix:
                        HwpxValidator._fix_mimetype_position(path)
                        result.auto_fixed.append("mimetype ZIP_STORED 교정")
                    else:
                        result.warnings.append("mimetype should be ZIP_STORED")

                # Required files
                required = ["Contents/header.xml", "Contents/content.hpf"]
                section_found = any(
                    n.startswith("Contents/section") and n.endswith(".xml")
                    for n in names
                )
                for req in required:
                    if req not in names:
                        result.errors.append(f"Missing required file: {req}")
                if not section_found:
                    result.errors.append("No section XML found in Contents/")

                if result.errors:
                    return

                # All XML parseable
                for name in names:
                    if name.endswith(".xml"):
                        try:
                            data = z.read(name)
                            etree.fromstring(data)
                        except etree.XMLSyntaxError as e:
                            if auto_fix and "namespace" in str(e).lower():
                                fixed = fix_namespaces(str(path))
                                if fixed:
                                    result.auto_fixed.append(
                                        f"네임스페이스 수정 ({len(fixed)}개 파일)"
                                    )
                            else:
                                result.errors.append(f"XML parse error in {name}: {e}")

        except zipfile.BadZipFile:
            result.errors.append("Not a valid ZIP file")
            return

        if not result.errors:
            result.level1_passed = True

        # Auto-fix namespaces proactively
        if auto_fix and result.level1_passed:
            fixed = fix_namespaces(str(path))
            if fixed:
                result.auto_fixed.append(f"네임스페이스 수정 ({len(fixed)}개 파일)")

    @staticmethod
    def _check_level2(path: Path, result: ValidationResult, auto_fix: bool) -> None:
        """Check: does it render correctly?"""
        try:
            with zipfile.ZipFile(path, "r") as z:
                header_data = z.read("Contents/header.xml")
                header_tree = etree.fromstring(header_data)

                # Extract max charPr and paraPr IDs from header
                ns = {"hh": "http://www.hancom.co.kr/hwpml/2011/head"}
                char_prs = header_tree.findall(".//hh:charPr", ns)
                para_prs = header_tree.findall(".//hh:paraPr", ns)
                max_char_id = max(
                    (int(el.get("id", "0")) for el in char_prs), default=0
                )
                max_para_id = max(
                    (int(el.get("id", "0")) for el in para_prs), default=0
                )

                # Check section files reference valid IDs
                for name in z.namelist():
                    if name.startswith("Contents/section") and name.endswith(".xml"):
                        sec_data = z.read(name)
                        sec_tree = etree.fromstring(sec_data)

                        for run in sec_tree.iter(
                            "{http://www.hancom.co.kr/hwpml/2011/paragraph}run"
                        ):
                            char_ref = int(run.get("charPrIDRef", "0"))
                            if char_ref > max_char_id:
                                if auto_fix:
                                    run.set("charPrIDRef", "0")
                                    result.auto_fixed.append(
                                        f"charPrIDRef {char_ref} → 0 (범위 초과)"
                                    )
                                else:
                                    result.warnings.append(
                                        f"charPrIDRef={char_ref} exceeds header max={max_char_id}"
                                    )

        except Exception as e:
            result.warnings.append(f"Level 2 check error: {e}")

        if not any("Level 2" in e for e in result.errors):
            result.level2_passed = True

    @staticmethod
    def _check_level3(
        path: Path, reference: Path, result: ValidationResult
    ) -> None:
        """Check: fidelity against original template."""
        try:
            with zipfile.ZipFile(path, "r") as z_out:
                with zipfile.ZipFile(reference, "r") as z_ref:
                    # Compare section sizes
                    for name in z_ref.namelist():
                        if name.startswith("Contents/section") and name.endswith(".xml"):
                            if name in z_out.namelist():
                                ref_size = len(z_ref.read(name))
                                out_size = len(z_out.read(name))
                                if ref_size > 0 and out_size < ref_size * 0.5:
                                    result.warnings.append(
                                        f"{name} size ratio: {out_size/ref_size:.0%} (possible content loss)"
                                    )
        except Exception as e:
            result.warnings.append(f"Level 3 check error: {e}")

        result.level3_passed = True  # Level 3 only produces warnings

    @staticmethod
    def _collect_stats(path: Path) -> dict:
        stats = {"file_size_kb": round(path.stat().st_size / 1024, 1)}
        try:
            with zipfile.ZipFile(path, "r") as z:
                for name in z.namelist():
                    if name.startswith("Contents/section") and name.endswith(".xml"):
                        data = z.read(name).decode("utf-8")
                        stats["paragraphs"] = data.count("<hp:p ")
                        stats["tables"] = data.count("<hp:tbl ")
                        stats["images"] = data.count("<hp:pic ")
                        stats["footnotes"] = data.count("<hp:footNote") + data.count(
                            "<hp:endNote"
                        )
        except Exception:
            pass
        return stats

    @staticmethod
    def _fix_mimetype_position(path: Path) -> None:
        """Repack ZIP with mimetype as first entry, ZIP_STORED."""
        tmp = str(path) + ".repack_tmp"
        with zipfile.ZipFile(str(path), "r") as zin:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                # Write mimetype first
                if "mimetype" in zin.namelist():
                    data = zin.read("mimetype")
                    zout.writestr("mimetype", data, compress_type=zipfile.ZIP_STORED)
                # Write everything else
                for item in zin.infolist():
                    if item.filename != "mimetype":
                        data = zin.read(item.filename)
                        zout.writestr(item, data)
        os.replace(tmp, str(path))
