"""Semantic style mapping — translate logical style names to template charPr/paraPr/styleIDRef IDs."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, Optional, Tuple

DEFAULT_STYLE_MAP: Dict[str, Dict[str, int]] = {
    "title": {"charPr": 7, "paraPr": 20},
    "heading1": {"charPr": 8, "paraPr": 0},
    "heading2": {"charPr": 13, "paraPr": 27},
    "heading3": {"charPr": 9, "paraPr": 0},
    "body": {"charPr": 0, "paraPr": 0},
    "body_bold": {"charPr": 9, "paraPr": 0},
    "caption": {"charPr": 11, "paraPr": 0},
    "table_header": {"charPr": 9, "paraPr": 21},
    "table_cell": {"charPr": 0, "paraPr": 22},
    "footnote": {"charPr": 3, "paraPr": 0},
    "page_number": {"charPr": 1, "paraPr": 0},
}


class StyleMapper:
    """Resolve semantic style names to (charPr, paraPr, styleIDRef) ID tuples."""

    def __init__(self, style_map: Dict[str, Dict] = None) -> None:
        self._map = style_map or DEFAULT_STYLE_MAP

    def resolve(self, name: str) -> Tuple[int, int, Optional[int]]:
        """Return (charPr, paraPr, styleIDRef) for *name*.

        Falls back to 'body'/'empty' for unknowns.
        styleIDRef is None if not present in the map (legacy DEFAULT_STYLE_MAP).
        """
        entry = self._map.get(name) or self._map.get("body") or self._map.get("empty")
        if entry is None:
            return 0, 0, None
        # Support both metadata.json format (cPr/pPr/sRef) and legacy (charPr/paraPr)
        cpr = int(entry.get("cPr", entry.get("charPr", 0)))
        ppr = int(entry.get("pPr", entry.get("paraPr", 0)))
        sref = entry.get("sRef")
        sref = int(sref) if sref is not None else None
        return cpr, ppr, sref

    def has_style(self, name: str) -> bool:
        return name in self._map

    @classmethod
    def from_metadata_path(cls, metadata_path: str) -> "StyleMapper":
        """Create a StyleMapper from a template metadata.json file."""
        with open(metadata_path, encoding="utf-8") as f:
            data = json.load(f)
        return cls(data.get("styles", {}))

    @classmethod
    def from_template_id(cls, template_id: str, base_dir: str = None) -> "StyleMapper":
        """Create a StyleMapper from a registered template ID.

        Search order: project-local → global → plugin-internal.
        """
        if base_dir is not None:
            metadata_path = str(Path(base_dir) / template_id / "metadata.json")
            return cls.from_metadata_path(metadata_path)

        search_paths = [
            Path.home() / '.claude' / 'hwpx-engine' / 'registered',
            Path.cwd() / 'assets' / 'registered',
        ]
        for search_dir in search_paths:
            metadata_path = search_dir / template_id / 'metadata.json'
            if metadata_path.exists():
                return cls.from_metadata_path(str(metadata_path))

        raise FileNotFoundError(f"Template '{template_id}' metadata.json not found")

    # Legacy compat
    @classmethod
    def from_metadata(cls, path: str) -> "StyleMapper":
        return cls.from_metadata_path(path)
