"""Template registry API — list/register/unregister/repair templates.

Filesystem is the single source of truth. No cache index file.
"""
from __future__ import annotations

import re
from pathlib import Path

# ─── Exceptions ───────────────────────────────────────────────────

class TemplateError(Exception):
    """Base class for template registry errors."""


class InvalidTemplateIdError(TemplateError):
    """Raised when a template ID fails validation."""


class TemplateAlreadyExistsError(TemplateError):
    """Raised when registering an ID that is already taken."""


class TemplateNotFoundError(TemplateError):
    """Raised when operating on a non-existent template ID."""


# ─── Constants ────────────────────────────────────────────────────

GLOBAL_REGISTERED = Path.home() / '.claude' / 'hwpx-engine' / 'registered'
TRASH_DIR = Path.home() / '.claude' / 'hwpx-engine' / '.trash'

_ID_PATTERN = re.compile(r'^[a-z0-9_]+$')
_ID_MAX_LEN = 64
_RESERVED_IDS = {'registry', '.trash', 'trash', '__pycache__'}


# ─── Public API ───────────────────────────────────────────────────

def validate_template_id(template_id: str) -> None:
    """Raise InvalidTemplateIdError if id is unsafe to use as a folder name."""
    if not template_id:
        raise InvalidTemplateIdError("Template ID cannot be empty")
    if len(template_id) > _ID_MAX_LEN:
        raise InvalidTemplateIdError(
            f"Template ID exceeds max length {_ID_MAX_LEN}: {len(template_id)}"
        )
    if template_id in _RESERVED_IDS:
        raise InvalidTemplateIdError(
            f"'{template_id}' is a reserved name"
        )
    if not _ID_PATTERN.match(template_id):
        raise InvalidTemplateIdError(
            f"Template ID must be lowercase alphanumeric with underscores "
            f"(got {template_id!r}). Allowed characters: a-z 0-9 _"
        )
