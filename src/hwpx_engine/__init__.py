"""HWPX Engine — reliable HWPX document creation and editing."""

from hwpx_engine.utils import fix_namespaces, NS_DECL
from hwpx_engine.validator import HwpxValidator, ValidationResult
from hwpx_engine.formatter import StyleMapper, DEFAULT_STYLE_MAP
from hwpx_engine.editor import HwpxEditor, TextNotFoundError
from hwpx_engine.build import build, BuildContext, register_template
from hwpx_engine.charpr_manager import CharPrManager
from hwpx_engine.xml_primitives import (
    make_para, make_run, make_two_run_para, add_linesegarray,
    make_table_xml, make_figure_box, make_image_pic,
    next_id, reset_id, xml_escape, has_part_overrides,
    get_para_text, set_para_text,
)
from hwpx_engine.registry import (
    list_templates,
    unregister_template,
    repair_template_metadata,
    validate_template_id,
    TemplateError,
    TemplateAlreadyExistsError,
    TemplateNotFoundError,
    InvalidTemplateIdError,
)


__all__ = [
    # Core
    "build", "BuildContext", "register_template",
    "HwpxEditor", "TextNotFoundError",
    "CharPrManager",
    # XML primitives
    "make_para", "make_run", "make_two_run_para", "add_linesegarray",
    "make_table_xml", "make_figure_box", "make_image_pic",
    "next_id", "reset_id", "xml_escape", "has_part_overrides",
    "get_para_text", "set_para_text",
    # Support
    "HwpxValidator", "ValidationResult",
    "StyleMapper", "DEFAULT_STYLE_MAP",
    "fix_namespaces", "NS_DECL",
    # Registry API
    "list_templates", "unregister_template", "repair_template_metadata",
    "validate_template_id",
    "TemplateError", "TemplateAlreadyExistsError",
    "TemplateNotFoundError", "InvalidTemplateIdError",
]
