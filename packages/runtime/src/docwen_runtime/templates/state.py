"""Persistent user-template identity and presentation state.

The packaged ``templates`` directory is read-only in MSIX installs.  This
module therefore keeps user-owned templates and their UI state under the
platform user-data directory, outside the application package.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import unicodedata
import uuid
from pathlib import Path
from typing import Any

from platformdirs import user_data_dir

_STATE_VERSION = 1
_CANONICAL_TEMPLATE_ID_PATTERN = re.compile(r"^template\.(?:docx|xlsx)\.[0-9a-f]{64}$")


def user_templates_dir() -> Path:
    """Return the writable directory used for user-managed templates."""

    return Path(user_data_dir("docwen", appauthor=False)) / "templates"


def template_state_path() -> Path:
    """Return the persistent template-management state file."""

    return Path(user_data_dir("docwen", appauthor=False)) / "template-state.json"


def _identity_key(filename: str) -> str:
    return unicodedata.normalize("NFC", filename).casefold()


def _new_template_id(target: str) -> str:
    stable_seed = uuid.uuid4().hex.encode("ascii")
    digest = hashlib.sha256(stable_seed).hexdigest()
    return f"template.{target}.{digest}"


class TemplateStateStore:
    """Small JSON store for stable user identities, enablement and ordering."""

    def __init__(self, path: Path | str, *, user_dir: Path | str) -> None:
        self.path = Path(path)
        self.user_dir = Path(user_dir)

    @classmethod
    def default(cls) -> "TemplateStateStore":
        return cls(template_state_path(), user_dir=user_templates_dir())

    def _empty_state(self) -> dict[str, Any]:
        return {
            "version": _STATE_VERSION,
            "user_identities": {},
            "enabled": {},
            "order": {"docx": [], "xlsx": []},
        }

    def load(self) -> dict[str, Any]:
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError):
            return self._empty_state()
        if not isinstance(raw, dict):
            return self._empty_state()

        state = self._empty_state()
        identities = raw.get("user_identities")
        enabled = raw.get("enabled")
        order = raw.get("order")
        if isinstance(identities, dict):
            state["user_identities"] = dict(identities)
        if isinstance(enabled, dict):
            state["enabled"] = {str(key): bool(value) for key, value in enabled.items()}
        if isinstance(order, dict):
            for target in ("docx", "xlsx"):
                values = order.get(target)
                if isinstance(values, list):
                    state["order"][target] = [str(value) for value in values if isinstance(value, str)]
        return state

    def save(self, state: dict[str, Any]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(state, ensure_ascii=False, indent=2, sort_keys=True)
        temporary = self.path.with_name(f".{self.path.name}.{uuid.uuid4().hex}.tmp")
        try:
            temporary.write_text(payload + "\n", encoding="utf-8")
            os.replace(temporary, self.path)
        finally:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass

    def ensure_user_identity(self, path: Path, target: str) -> str:
        """Return a stable canonical ID for one user-owned template path."""

        state = self.load()
        identities = state["user_identities"]
        key = _identity_key(path.name)
        existing = identities.get(key)
        if isinstance(existing, dict):
            existing_id = existing.get("id")
            existing_target = existing.get("target")
            if (
                isinstance(existing_id, str)
                and _CANONICAL_TEMPLATE_ID_PATTERN.fullmatch(existing_id)
                and existing_target == target
            ):
                return existing_id

        template_id = _new_template_id(target)
        identities[key] = {
            "id": template_id,
            "target": target,
            "filename": path.name,
        }
        self.save(state)
        return template_id

    def rename_user_identity(self, old_name: str, new_name: str) -> None:
        """Move a user identity to a new filename without changing its ID."""

        state = self.load()
        identities = state["user_identities"]
        old_key = _identity_key(old_name)
        new_key = _identity_key(new_name)
        entry = identities.pop(old_key, None)
        if isinstance(entry, dict):
            entry = dict(entry)
            entry["filename"] = new_name
            identities[new_key] = entry
            self.save(state)

    def forget_user_identity(self, filename: str) -> None:
        state = self.load()
        identities = state["user_identities"]
        if identities.pop(_identity_key(filename), None) is not None:
            self.save(state)

    def is_enabled(self, template_id: str) -> bool:
        state = self.load()
        return bool(state["enabled"].get(template_id, True))

    def set_enabled(self, template_id: str, enabled: bool) -> None:
        state = self.load()
        state["enabled"][template_id] = bool(enabled)
        self.save(state)

    def ordered_ids(self, target: str, discovered_ids: list[str]) -> list[str]:
        """Return discovered IDs in persisted order, appending new IDs at the end."""

        state = self.load()
        persisted = state["order"].setdefault(target, [])
        discovered_set = set(discovered_ids)
        result: list[str] = []
        for template_id in persisted:
            if template_id in discovered_set and template_id not in result:
                result.append(template_id)
        for template_id in discovered_ids:
            if template_id not in result:
                result.append(template_id)
        if result != persisted:
            state["order"][target] = result
            self.save(state)
        return result

    def set_order(self, target: str, ordered_ids: list[str]) -> None:
        state = self.load()
        deduplicated: list[str] = []
        for template_id in ordered_ids:
            if template_id not in deduplicated:
                deduplicated.append(template_id)
        state["order"][target] = deduplicated
        self.save(state)


__all__ = [
    "TemplateStateStore",
    "template_state_path",
    "user_templates_dir",
]
