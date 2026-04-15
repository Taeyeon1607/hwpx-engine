"""Build orchestrator for HWPX document generation.

Flow: template_id → load builder.py → execute section handlers → save

Template-specific logic lives
in per-template builder.py + modules/, while this file handles the generic
build pipeline: copy template, process sections, embed images, validate.
"""
import importlib.util
import json
import os
import re
import shutil
import zipfile
from pathlib import Path

from lxml import etree

from hwpx_engine.charpr_manager import CharPrManager
from hwpx_engine.utils import fix_namespaces
from hwpx_engine.validator import HwpxValidator
from hwpx_engine.xml_primitives import reset_id

HP = 'http://www.hancom.co.kr/hwpml/2011/paragraph'

_REQUIRED_METADATA_FIELDS = ["id", "styles", "sections"]


def _validate_metadata(metadata: dict, context: str = "build"):
    """Validate metadata.json has required fields."""
    missing = [f for f in _REQUIRED_METADATA_FIELDS if f not in metadata]
    if missing:
        raise ValueError(
            f"metadata.json is missing required fields for {context}: {', '.join(missing)}. "
            f"Required: {', '.join(_REQUIRED_METADATA_FIELDS)}"
        )


class BuildContext:
    """Shared state passed to all template module handlers during a build."""

    def __init__(self, styles: dict, metadata: dict, charpr_mgr: CharPrManager):
        self.styles = styles
        self.metadata = metadata
        self.charpr_mgr = charpr_mgr
        self.pending_images = []  # [(bin_id, image_bytes, media_type, ext)]
        self.image_counter = 100

    def resolve_style(self, name: str) -> dict:
        """Get style dict by semantic name. Returns {'pPr':, 'sRef':, 'cPr':, ...}."""
        return self.styles.get(name, {'pPr': '0', 'sRef': '0', 'cPr': '0'})

    def register_image(self, image_path: str) -> tuple:
        """Register an image for embedding. Returns (bin_id, org_w, org_h)."""
        from PIL import Image
        img = Image.open(image_path)
        w_px, h_px = img.size
        img_format = img.format or 'JPEG'
        ext = {'JPEG': 'jpg', 'PNG': 'png', 'BMP': 'bmp', 'GIF': 'gif'}.get(img_format, 'jpg')
        media = {'JPEG': 'image/jpg', 'PNG': 'image/png', 'BMP': 'image/bmp'}.get(img_format, 'image/jpg')

        self.image_counter += 1
        bin_id = f'inserted_image{self.image_counter}'
        with open(image_path, 'rb') as f:
            data = f.read()

        self.pending_images.append((bin_id, data, media, ext))

        px_to_hwp = 75  # 7200 / 96 DPI
        org_w = w_px * px_to_hwp
        org_h = h_px * px_to_hwp
        return bin_id, org_w, org_h


def build(template_id: str, content: dict, output_path: str,
          base_dir: str = None) -> dict:
    """Build an HWPX document from a registered template.

    Args:
        template_id: registered template ID (directory name under assets/registered/)
        content: content dict as defined by the template's usage-guide.md
        output_path: where to write the output .hwpx file
        base_dir: override base directory for registered templates

    Returns:
        dict with 'success', 'output_path', 'validation', 'stats'
    """
    template_dir = _resolve_template_dir(template_id, base_dir)
    metadata = json.loads((template_dir / 'metadata.json').read_text(encoding='utf-8'))
    _validate_metadata(metadata)
    styles = metadata.get('styles', {})

    # Ensure output directory exists
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(str(template_dir / 'template.hwpx'), output_path)

    # Reset ID counter for this build
    reset_id(900000)

    # Load template's builder.py (declares sections pipeline)
    # If builder has a prepare() function, call it BEFORE resolving handlers
    # (prepare may rebuild the sections list with new string handler references)
    builder_path = Path(str(template_dir)) / 'builder.py'
    spec = importlib.util.spec_from_file_location('builder', str(builder_path))
    builder_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(builder_module)

    if hasattr(builder_module, 'prepare'):
        builder_module.prepare(output_path, content, metadata)

    # Resolve handler strings to actual module.process functions
    _resolve_handlers(str(template_dir), builder_module.sections)

    sections = builder_module.sections

    # Pre-load header.xml for CharPrManager
    with zipfile.ZipFile(output_path, 'r') as z:
        header_root = etree.fromstring(z.read('Contents/header.xml'))
    charpr_mgr = CharPrManager(header_root)
    ctx = BuildContext(styles, metadata, charpr_mgr)

    # Process ZIP: execute section handlers
    _process_zip(output_path, sections, content, ctx)

    # Embed images if any were registered
    if ctx.pending_images:
        _embed_images(output_path, ctx.pending_images)

    # Namespace fix + validation
    fix_namespaces(output_path)
    result = HwpxValidator.validate(output_path, auto_fix=True)

    return {
        'success': result.valid,
        'output_path': str(output_path),
        'validation': {
            'valid': result.valid,
            'level1_passed': result.level1_passed,
            'level2_passed': result.level2_passed,
            'level3_passed': result.level3_passed,
            'auto_fixed': result.auto_fixed,
            'warnings': result.warnings,
            'errors': result.errors,
        },
        'stats': result.stats,
    }


GLOBAL_REGISTERED = Path.home() / '.claude' / 'hwpx-engine' / 'registered'


def register_template(template_id: str, source_dir: str) -> Path:
    """Register a template to the global persistent path.

    Copies all template files to ~/.claude/hwpx-engine/registered/{template_id}/.
    This path survives plugin updates.

    Args:
        template_id: unique template identifier
        source_dir: directory containing template.hwpx, metadata.json, etc.

    Returns:
        Path to the registered template directory
    """
    source = Path(source_dir)
    if not (source / 'template.hwpx').exists():
        raise FileNotFoundError(f"template.hwpx not found in {source}")
    if not (source / 'metadata.json').exists():
        raise FileNotFoundError(f"metadata.json not found in {source}")

    metadata = json.loads((source / 'metadata.json').read_text(encoding='utf-8'))
    _validate_metadata(metadata, "register")

    target = GLOBAL_REGISTERED / template_id

    if target.exists():
        shutil.rmtree(target)
    shutil.copytree(str(source), str(target))

    return target


# ─── Internal helpers ─────────────────────────────────────────────


def _resolve_template_dir(template_id: str, base_dir: str = None) -> Path:
    """Find the registered template directory.

    Search order (first match wins):
    1. base_dir (if explicitly provided)
    2. Global: ~/.claude/hwpx-engine/registered/{template_id}/  ← 유일한 영구 저장소
    3. Project-local: ./assets/registered/{template_id}/  ← 프로젝트 전용 양식
    """
    if base_dir is not None:
        template_dir = Path(base_dir) / template_id
        if template_dir.exists():
            return template_dir
        raise FileNotFoundError(f"Template not found: {template_dir}")

    search_paths = [
        # Global (persistent, survives plugin updates) — primary
        GLOBAL_REGISTERED,
        # Project-local (project-specific overrides)
        Path.cwd() / 'assets' / 'registered',
    ]

    for search_dir in search_paths:
        template_dir = search_dir / template_id
        if template_dir.exists():
            return template_dir

    searched = '\n  '.join(str(p / template_id) for p in search_paths)
    raise FileNotFoundError(
        f"Template '{template_id}' not found. Searched:\n  {searched}"
    )


def _resolve_handlers(template_dir: str, sections: list):
    """Resolve handler string references to actual module.process functions.

    Each section dict may have ``handler`` as a string (e.g. ``'cover'``).
    This resolves it to ``modules/cover.py -> process()`` function.
    """
    modules_dir = Path(template_dir) / 'modules'
    for sec in sections:
        handler = sec['handler']
        if isinstance(handler, str):
            mod_path = modules_dir / f'{handler}.py'
            if not mod_path.exists():
                raise FileNotFoundError(f"Module not found: {mod_path}")
            mod_spec = importlib.util.spec_from_file_location(
                f'modules.{handler}', str(mod_path))
            handler_mod = importlib.util.module_from_spec(mod_spec)
            mod_spec.loader.exec_module(handler_mod)
            sec['handler'] = handler_mod.process


def _load_builder(template_dir: str):
    """Load builder.py and resolve module handler references.

    builder.py declares sections with handler names as strings:
        sections = [
            {'file': 'section0.xml', 'handler': 'cover'},
            {'file': 'section1.xml', 'handler': 'toc'},
        ]

    This function resolves each handler string to the actual .process function
    from the corresponding modules/<name>.py file.
    """
    builder_path = Path(template_dir) / 'builder.py'
    if not builder_path.exists():
        raise FileNotFoundError(f"builder.py not found in {template_dir}")

    spec = importlib.util.spec_from_file_location('builder', str(builder_path))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    _resolve_handlers(template_dir, module.sections)
    return module


def _process_zip(hwpx_path: str, sections: list, content: dict, ctx: BuildContext):
    """Process all section XMLs via the handler pipeline.

    Header.xml is deferred until all sections are processed (so charPr
    additions from section processing are captured).
    """
    # Build a lookup: filename -> section dict (handler + extra metadata)
    section_map = {}
    for sec in sections:
        section_map[f"Contents/{sec['file']}"] = sec

    # Global replacements (applied to all section XMLs)
    global_reps = content.get('global_replacements', {})

    tmp = hwpx_path + '.build_tmp'
    with zipfile.ZipFile(hwpx_path, 'r') as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            pending_header = None

            for item in zin.infolist():
                data = zin.read(item.filename)

                # Defer header.xml
                if item.filename == 'Contents/header.xml':
                    pending_header = (item, data)
                    continue

                # Execute section handler if registered
                if item.filename in section_map:
                    sec = section_map[item.filename]
                    handler = sec['handler']
                    # Pass section metadata (sec_num, etc.) as kwargs
                    kwargs = {k: v for k, v in sec.items() if k not in ('file', 'handler')}
                    try:
                        data = handler(data, content, ctx, **kwargs)
                    except Exception as e:
                        sec_file = sec['file']
                        raise RuntimeError(
                            f"Error processing section '{sec_file}' with handler "
                            f"'{handler.__module__}.{handler.__name__}': {e}"
                        ) from e

                # Apply global replacements to all section XMLs
                if (item.filename.startswith('Contents/section') and
                        item.filename.endswith('.xml') and global_reps):
                    data = _apply_global_replacements(data, global_reps)

                # mimetype must be ZIP_STORED (not deflated)
                if item.filename == 'mimetype':
                    zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                else:
                    zout.writestr(item, data)

            # Write header.xml last (with any new charPr/borderFill entries)
            if pending_header:
                item, data = pending_header
                if ctx.charpr_mgr.modified:
                    data = ctx.charpr_mgr.serialize()
                zout.writestr(item, data)

    os.replace(tmp, hwpx_path)


def _apply_global_replacements(xml_data: bytes, replacements: dict) -> bytes:
    """Apply text replacements across section XML (running headers, etc)."""
    text = xml_data.decode('utf-8')
    for old, new in sorted(replacements.items(), key=lambda x: len(x[0]), reverse=True):
        text = text.replace(old, new)
    return text.encode('utf-8')


def _embed_images(hwpx_path: str, pending_images: list):
    """Add images to the HWPX ZIP and update content.hpf manifest."""
    tmp = hwpx_path + '.img_tmp'
    with zipfile.ZipFile(hwpx_path, 'r') as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'Contents/content.hpf':
                    text = data.decode('utf-8')
                    insert_pos = text.find('</opf:manifest>')
                    if insert_pos > 0:
                        entries = ''
                        for bin_id, _, media, ext in pending_images:
                            entries += (f'<opf:item id="{bin_id}" '
                                        f'href="BinData/{bin_id}.{ext}" '
                                        f'media-type="{media}" isEmbeded="1"/>')
                        text = text[:insert_pos] + entries + text[insert_pos:]
                    data = text.encode('utf-8')
                if item.filename == 'mimetype':
                    zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                else:
                    zout.writestr(item, data)

            for bin_id, img_data, _, ext in pending_images:
                zout.writestr(f'BinData/{bin_id}.{ext}', img_data)

    os.replace(tmp, hwpx_path)
