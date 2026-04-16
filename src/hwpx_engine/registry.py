"""Template registry API — list/register/unregister/repair templates.

Filesystem is the single source of truth. No cache index file.
"""
from __future__ import annotations

import json
import re
import shutil
import time
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

_DISPLAY_FIELDS = ["display_name", "summary"]


def list_templates() -> list[dict]:
    """Scan registered/ and return all templates with status.

    Each entry has the shape:
        {id, display_name, summary, description, path, status, [missing_fields], [error]}

    Status values:
        - 'ok': all required display fields (display_name, summary) present
        - 'incomplete': one or more display fields missing. `missing_fields`
          lists which of {display_name, summary} are absent or empty.
        - 'invalid': metadata.json failed to parse or is not a JSON object.
          `error` contains the exception message.

    `display_name` in the returned entry always falls back to the folder name
    when metadata lacks it; `missing_fields` still reports it as missing so
    the caller can trigger `repair_template_metadata`.

    Side effect: creates `~/.claude/hwpx-engine/registered/` on first call if
    it does not yet exist.
    """
    GLOBAL_REGISTERED.mkdir(parents=True, exist_ok=True)
    results = []
    for sub in sorted(GLOBAL_REGISTERED.iterdir()):
        if not sub.is_dir():
            continue
        if sub.name.startswith('.'):
            continue
        meta_path = sub / 'metadata.json'
        if not meta_path.exists():
            continue

        entry = {"id": sub.name, "path": str(sub)}
        try:
            meta = json.loads(meta_path.read_text(encoding='utf-8'))
        except (json.JSONDecodeError, OSError) as e:
            entry["status"] = "invalid"
            entry["error"] = f"{type(e).__name__}: {e}"
            results.append(entry)
            continue

        if not isinstance(meta, dict):
            entry["status"] = "invalid"
            entry["error"] = f"metadata.json is {type(meta).__name__}, not dict"
            results.append(entry)
            continue

        entry["display_name"] = meta.get("display_name", sub.name)
        entry["summary"] = meta.get("summary", "")
        entry["description"] = meta.get("description", "")

        missing = [f for f in _DISPLAY_FIELDS if not meta.get(f)]
        if missing:
            entry["status"] = "incomplete"
            entry["missing_fields"] = missing
        else:
            entry["status"] = "ok"

        results.append(entry)
    return results


def unregister_template(template_id: str, backup: bool = True) -> str:
    """Remove a registered template. By default moves folder to .trash/.

    Returns the new path as a string (trash path if backup=True, else '').
    Raises TemplateNotFoundError if the template doesn't exist.
    Raises InvalidTemplateIdError if the id is malformed.
    """
    validate_template_id(template_id)

    target = GLOBAL_REGISTERED / template_id
    if not target.exists():
        raise TemplateNotFoundError(f"Template '{template_id}' not registered")

    if backup:
        TRASH_DIR.mkdir(parents=True, exist_ok=True)
        stamp = time.strftime("%Y%m%d_%H%M%S")
        dest = TRASH_DIR / f"{template_id}_{stamp}"
        shutil.move(str(target), str(dest))
        return str(dest)
    else:
        shutil.rmtree(target)
        return ""


def _humanize_id(template_id: str) -> str:
    """Convert 'my_template_a4' → 'My Template A4'."""
    return " ".join(word.capitalize() for word in template_id.split("_"))


def repair_template_metadata(template_id: str) -> dict:
    """Auto-fill ONLY display_name (deterministically derivable from id).

    Deliberately does NOT touch summary or description — those require
    human judgment and a placeholder would mask the 'incomplete' status
    that the SKILL.md disambiguation rule depends on.

    Returns: {
        "filled": list[str],            # fields written this call
        "skipped": list[str],            # fields that were already present
        "warnings_summary_missing": list[str],  # fields that still need human input
    }
    Raises TemplateNotFoundError if the template is not registered.
    Raises ValueError if metadata.json is unparseable.

    Note: writes JSON with indent=2 / ensure_ascii=False. Existing files
    using the same format (the 7 bundled GRI templates confirmed to match)
    produce minimal diffs. If a third-party template uses a different
    format, the rewrite will normalize it — acceptable tradeoff.
    """
    target = GLOBAL_REGISTERED / template_id
    if not target.exists():
        raise TemplateNotFoundError(f"Template '{template_id}' not registered")

    meta_path = target / "metadata.json"
    try:
        metadata = json.loads(meta_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise ValueError(f"metadata.json for '{template_id}' is not valid JSON: {e}") from e

    filled = []
    skipped = []
    warnings_summary_missing = []

    if not metadata.get("display_name"):
        metadata["display_name"] = _humanize_id(template_id)
        filled.append("display_name")
    else:
        skipped.append("display_name")

    # summary / description: do NOT auto-fill. Report as needing human input.
    if not metadata.get("summary"):
        warnings_summary_missing.append("summary")
    if not metadata.get("description"):
        warnings_summary_missing.append("description")

    if filled:
        meta_path.write_text(
            json.dumps(metadata, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    return {
        "filled": filled,
        "skipped": skipped,
        "warnings_summary_missing": warnings_summary_missing,
    }
