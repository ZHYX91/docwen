"""Persistent user-template identity and presentation state.

The packaged ``templates`` directory is read-only in MSIX installs.  This
module therefore keeps user-owned templates and their UI state under the
platform user-data directory, outside the application package.
"""

from __future__ import annotations

import hashlib
import json
import re
import threading
import unicodedata
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from docwen_runtime import file_transactions as transaction
from docwen_runtime.file_transactions import process_config_lock

_STATE_VERSION = 1
_STATE_LOCK = threading.RLock()
_LOCK_DEPTH = threading.local()
_CANONICAL_TEMPLATE_ID_PATTERN = re.compile(r"^template\.(?:docx|xlsx)\.[0-9a-f]{64}$")


def user_templates_dir() -> Path:
    """Return the writable directory used for user-managed templates."""

    return template_data_root() / "templates"


def template_data_root() -> Path:
    """Read the template component of the same bound startup profile."""
    from docwen_runtime.profile_paths import current_profile_paths

    return current_profile_paths().data_dir


def template_state_path() -> Path:
    """Return the persistent template-management state file."""

    return template_data_root() / "template-state.json"


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
    def default(cls) -> TemplateStateStore:
        return cls(template_state_path(), user_dir=user_templates_dir())

    def _empty_state(self) -> dict[str, Any]:
        return {
            "version": _STATE_VERSION,
            "defaults_initialized": False,
            "user_identities": {},
            "enabled": {},
            "order": {"docx": [], "xlsx": []},
            "defaults": {"docx": None, "xlsx": None},
        }

    def load(self) -> dict[str, Any]:
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
        except FileNotFoundError:
            return self._empty_state()
        except (OSError, ValueError) as exc:
            raise ValueError(f"Cannot read template state: {self.path}") from exc
        if not isinstance(raw, dict) or raw.get("version") != _STATE_VERSION:
            raise ValueError(f"Unsupported template state: {self.path}")

        state = self._empty_state()
        if type(raw.get("defaults_initialized")) is not bool:
            raise ValueError("Invalid template defaults initialization state")
        state["defaults_initialized"] = raw["defaults_initialized"]
        identities = raw.get("user_identities")
        enabled = raw.get("enabled")
        order = raw.get("order")
        if not isinstance(identities, dict) or not isinstance(enabled, dict) or not isinstance(order, dict):
            raise ValueError("Invalid template state structure")
        if any(type(value) is not bool for value in enabled.values()):
            raise ValueError("Invalid template enablement")
        for key, entry in identities.items():
            if not isinstance(entry, dict) or not isinstance(entry.get("filename"), str):
                raise ValueError("Invalid template identity")
            filename = entry["filename"]
            if Path(filename).name != filename or _identity_key(filename) != key:
                raise ValueError("Invalid template filename")
            target = entry.get("target")
            identity = entry.get("id")
            if (
                target not in {"docx", "xlsx"}
                or not isinstance(identity, str)
                or not re.fullmatch(rf"template\.{target}\.[0-9a-f]{{64}}", identity)
            ):
                raise ValueError("Invalid template identity")
        defaults = raw.get("defaults")
        if not isinstance(defaults, dict) or set(defaults) != {"docx", "xlsx"}:
            raise ValueError("Invalid template defaults")
        for target in ("docx", "xlsx"):
            selected = defaults[target]
            if selected is not None and (
                not isinstance(selected, str) or not re.fullmatch(rf"template\.{target}\.[0-9a-f]{{64}}", selected)
            ):
                raise ValueError("Invalid default template identity")
            values = order.get(target)
            if (
                not isinstance(values, list)
                or any(
                    not isinstance(value, str) or not re.fullmatch(rf"template\.{target}\.[0-9a-f]{{64}}", value)
                    for value in values
                )
                or len(set(values)) != len(values)
            ):
                raise ValueError("Invalid template order")
        state["defaults"] = defaults
        state["user_identities"] = dict(identities)
        state["enabled"] = dict(enabled)
        state["order"] = {target: list(order[target]) for target in ("docx", "xlsx")}
        return state

    def save(self, state: dict[str, Any]) -> None:
        from docwen_runtime.toml_io import atomic_write_bytes

        payload = json.dumps(state, ensure_ascii=False, indent=2, sort_keys=True)
        atomic_write_bytes(self.path, (payload + "\n").encode("utf-8"))

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
                stat = path.stat()
                file_key = f"{stat.st_dev}:{stat.st_ino}"
                if existing.get("file_key") != file_key:
                    existing["file_key"] = file_key
                    self.save(state)
                return existing_id

        stat = path.stat()
        file_key = f"{stat.st_dev}:{stat.st_ino}"
        for old_key, entry in list(identities.items()):
            if (
                entry.get("file_key") == file_key
                and entry.get("target") == target
                and not (self.user_dir / entry["filename"]).exists()
            ):
                entry["filename"] = path.name
                identities[key] = identities.pop(old_key)
                self.save(state)
                return entry["id"]
        template_id = _new_template_id(target)
        identities[key] = {
            "id": template_id,
            "target": target,
            "filename": path.name,
            "file_key": file_key,
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
        if not enabled:
            for target, selected in state["defaults"].items():
                if selected == template_id:
                    state["defaults"][target] = None
        self.save(state)

    @contextmanager
    def locked(self):
        """Serialize complete catalog operations across GUI and CLI processes."""
        with _STATE_LOCK, process_config_lock(self.path.parent):
            depth = getattr(_LOCK_DEPTH, "value", 0)
            _LOCK_DEPTH.value = depth + 1
            try:
                if depth == 0:
                    tracked = self.load()["user_identities"]
                    allowed = [self.path, *(self.user_dir / item["filename"] for item in tracked.values())]
                    if self.user_dir.exists():
                        allowed.extend(self.user_dir.iterdir())
                    journal = self.path.parent / transaction.CONFIG_JOURNAL_NAME
                    if journal.exists():
                        envelope = json.loads(journal.read_text(encoding="utf-8"))
                        for item in envelope.get("payload", {}).get("preimages", []):
                            candidate = self.path.parent / item.get("path", "")
                            if candidate == self.path or candidate.parent == self.user_dir:
                                allowed.append(candidate)
                    transaction.recover_transaction_journal(
                        self.path.parent,
                        allowed,
                        lambda before, operation: transaction.restore_user_file_preimage(before, operation=operation),
                    )
                yield
            finally:
                _LOCK_DEPTH.value = depth

    @contextmanager
    def file_transaction(self, *paths: Path):
        """Journal template bytes and metadata together before publishing changes."""
        with self.locked():
            before = [transaction.capture_user_file_preimage(path) for path in (self.path, *paths)]
            transaction.write_transaction_journal(self.path.parent, "templates", before, state="PREPARED")
            try:
                yield
                transaction.mark_transaction_committed(self.path.parent, "templates", before)
            except BaseException:
                for item in reversed(before):
                    transaction.restore_user_file_preimage(item, operation="templates")
                transaction.remove_transaction_journal(self.path.parent)
                raise
            transaction.remove_transaction_journal(self.path.parent)

    def default_id(self, target: str) -> str | None:
        return self.load()["defaults"].get(target)

    def set_default(self, target: str, template_id: str | None) -> None:
        with self.locked():
            state = self.load()
            state["defaults"][target] = template_id
            self.save(state)

    def ordered_ids(self, target: str, discovered_ids: list[str]) -> list[str]:
        """Return discovered IDs in persisted order and persist newly discovered IDs."""

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
        appended = [identity for identity in discovered_ids if identity not in persisted]
        if appended:
            state["order"][target] = [*persisted, *appended]
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
