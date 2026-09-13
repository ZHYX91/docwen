"""User-facing template management operations.

All writes are confined to the platform user-data template directory.  Bundled
application templates are treated as immutable resources and can only be copied
into the user directory before editing.
"""

from __future__ import annotations

import unicodedata
from collections.abc import Callable, Iterable
from functools import wraps
from pathlib import Path
from typing import Concatenate

from docwen_core.detection import inspect_file
from docwen_core.models import StructureStatus
from docwen_runtime.toml_io import atomic_write_bytes

from .registry import TemplateInfo, TemplateNotFoundError, TemplateRegistry
from .state import TemplateStateStore, user_templates_dir


class TemplateManagementError(RuntimeError):
    """A requested template-management operation cannot be completed safely."""


def _safe_display_name(value: str) -> str:
    name = unicodedata.normalize("NFC", value.strip())
    if not name or name in {".", ".."}:
        raise TemplateManagementError("Template name cannot be empty")
    if Path(name).name != name or any(character in name for character in '<>:"/\\|?*'):
        raise TemplateManagementError("Template name contains invalid path characters")
    if (
        name.endswith(".")
        or any(ord(character) < 32 for character in name)
        or name.split(".")[0].upper()
        in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}
    ):
        raise TemplateManagementError("Template name is not a portable filename")
    return name


def _detect_valid_target(path: Path) -> str:
    candidate = path.expanduser().resolve(strict=False)
    if not candidate.is_file():
        raise TemplateManagementError(f"Template file not found: {path}")
    try:
        inspection = inspect_file(str(candidate))
    except Exception as exc:
        raise TemplateManagementError(f"Template could not be inspected: {candidate}: {exc}") from exc
    target = inspection.detected_format
    if target not in {"docx", "xlsx"} or inspection.structure_status is not StructureStatus.VALID:
        raise TemplateManagementError(f"File is not a structurally valid DOCX/XLSX template: {candidate}")
    return target


def _move_to_recycle_bin(path: Path) -> None:
    """Use the declared cross-platform trash backend; never permanently unlink."""
    from send2trash import send2trash

    send2trash(str(path))


def serialized[**P, R](
    method: Callable[Concatenate[TemplateManager, P], R],
) -> Callable[Concatenate[TemplateManager, P], R]:
    """Keep discovery and mutations in one cooperative catalog transaction."""

    @wraps(method)
    def call(self: TemplateManager, *args: P.args, **kwargs: P.kwargs) -> R:
        with self.state_store.locked():
            return method(self, *args, **kwargs)

    return call


class TemplateManager:
    """Manage built-in and user templates without writing into the app package."""

    def __init__(
        self,
        registry: TemplateRegistry,
        *,
        state_store: TemplateStateStore,
        user_dir: Path | str,
    ) -> None:
        self.registry = registry
        self.state_store = state_store
        self.user_dir = Path(user_dir)

    @classmethod
    def default(cls) -> TemplateManager:
        state = TemplateStateStore.default()
        return cls(
            TemplateRegistry.default(state_store=state),
            state_store=state,
            user_dir=user_templates_dir(),
        )

    def ensure_user_directory(self) -> Path:
        self.user_dir.mkdir(parents=True, exist_ok=True)
        return self.user_dir

    @serialized
    def list_templates(self, target_type: str | None = None, *, include_disabled: bool = True) -> list[TemplateInfo]:
        return self.registry.list_templates(target_type, include_disabled=include_disabled)

    def is_custom(self, template: TemplateInfo) -> bool:
        try:
            return template.path.parent.resolve(strict=False) == self.user_dir.resolve(strict=False)
        except OSError:
            return template.path.parent == self.user_dir

    def is_enabled(self, template_id: str) -> bool:
        return self.state_store.is_enabled(template_id)

    @serialized
    def set_enabled(self, template_id: str, enabled: bool) -> None:
        self._find(template_id)
        self.state_store.set_enabled(template_id, enabled)

    @serialized
    def set_default(self, template_id: str) -> None:
        template = self._find(template_id)
        if not self.is_enabled(template_id):
            raise TemplateManagementError("Enable the template before setting it as default")
        self.state_store.set_default(template.target, template.id)

    @serialized
    def set_order(self, target: str, template_ids: list[str]) -> None:
        expected = {item.id for item in self.list_templates(target)}
        if len(template_ids) != len(expected) or set(template_ids) != expected:
            raise TemplateManagementError("Template catalog changed; refresh before sorting")
        self.state_store.set_order(target, template_ids)

    def import_templates(self, paths: Iterable[Path | str]) -> list[TemplateInfo]:
        imported: list[TemplateInfo] = []
        for path in paths:
            imported.append(self.import_template(path))
        return imported

    @serialized
    def import_template(self, path: Path | str, *, replace_id: str | None = None) -> TemplateInfo:
        source = Path(path).expanduser().resolve()
        target = _detect_valid_target(source)
        if replace_id is not None:
            existing = self._find(replace_id)
            self._require_custom(existing)
            if existing.target != target:
                raise TemplateManagementError("Template format does not match the replacement")
            if source != existing.path.resolve():
                with self.state_store.file_transaction(existing.path):
                    atomic_write_bytes(existing.path, source.read_bytes())
            return self._find(replace_id)
        destination = self._unique_destination(source.stem, target)
        self.ensure_user_directory()
        with self.state_store.file_transaction(destination):
            atomic_write_bytes(destination, source.read_bytes())
            template_id = self.state_store.ensure_user_identity(destination, target)
            self.state_store.set_enabled(template_id, True)
            self._place_at_end(target, template_id)
        return self._find(template_id)

    @serialized
    def import_conflict(self, path: Path | str) -> TemplateInfo | None:
        source = Path(path)
        target = _detect_valid_target(source)
        return next(
            (
                item
                for item in self.list_templates(target)
                if self.is_custom(item) and item.name.casefold() == source.stem.casefold()
            ),
            None,
        )

    @serialized
    def copy_builtin_as_custom(self, template_id: str, *, custom_name: str | None = None) -> TemplateInfo:
        template = self._find(template_id)
        if self.is_custom(template):
            raise TemplateManagementError("Only built-in templates use the copy-as-custom operation")
        base_name = _safe_display_name(custom_name) if custom_name is not None else f"{template.name} - custom"
        destination = self._unique_destination(base_name, template.target)
        self.ensure_user_directory()
        with self.state_store.file_transaction(destination):
            atomic_write_bytes(destination, template.path.read_bytes())
            custom_id = self.state_store.ensure_user_identity(destination, template.target)
            self.state_store.set_enabled(custom_id, True)
            self._place_after(template.target, custom_id, after_id=template.id)
        return self._find(custom_id)

    @serialized
    def rename_custom(self, template_id: str, new_name: str) -> TemplateInfo:
        template = self._find(template_id)
        self._require_custom(template)
        safe_name = _safe_display_name(new_name)
        suffix = f".{template.target}"
        if safe_name.casefold().endswith(suffix):
            safe_name = safe_name[: -len(suffix)]
            safe_name = _safe_display_name(safe_name)
        for candidate in self.list_templates(template.target, include_disabled=True):
            if candidate.id != template.id and candidate.name.casefold() == safe_name.casefold():
                raise TemplateManagementError(f"A template named {safe_name!r} already exists")
        destination = template.path.with_name(f"{safe_name}{suffix}")
        if destination.exists() and destination.resolve(strict=False) != template.path.resolve(strict=False):
            raise TemplateManagementError(f"A template named {destination.name!r} already exists")
        old_name = template.path.name
        if destination == template.path:
            return template
        with self.state_store.file_transaction(template.path, destination):
            template.path.rename(destination)
            self.state_store.rename_user_identity(old_name, destination.name)
        return self._find(template_id)

    @serialized
    def delete_custom(self, template_id: str) -> None:
        """Move a custom template to the OS recycle bin, never hard-delete it."""

        template = self._find(template_id)
        self._require_custom(template)
        filename = template.path.name
        try:
            with self.state_store.file_transaction(template.path):
                self.state_store.set_enabled(template.id, False)
                self._remove_from_order(template.target, template.id)
                self.state_store.forget_user_identity(filename)
                _move_to_recycle_bin(template.path)
        except TemplateManagementError:
            raise
        except Exception as exc:
            raise TemplateManagementError(f"Could not move template to the recycle bin: {template.path.name}") from exc

    @serialized
    def export_custom(self, template_id: str, destination: Path | str) -> Path:
        template = self._find(template_id)
        self._require_custom(template)
        destination_path = Path(destination).expanduser()
        if destination_path.exists() and destination_path.is_dir():
            destination_path = destination_path / template.path.name
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        atomic_write_bytes(destination_path, template.path.read_bytes())
        return destination_path

    @serialized
    def move(self, template_id: str, offset: int) -> None:
        if offset == 0:
            return
        template = self._find(template_id)
        ordered = [item.id for item in self.list_templates(template.target, include_disabled=True)]
        try:
            index = ordered.index(template_id)
        except ValueError as exc:
            raise TemplateNotFoundError(template_id) from exc
        destination = max(0, min(len(ordered) - 1, index + offset))
        if destination == index:
            return
        ordered.pop(index)
        ordered.insert(destination, template_id)
        self.state_store.set_order(template.target, ordered)

    def _place_at_end(self, target: str, template_id: str) -> None:
        ordered = [item.id for item in self.list_templates(target, include_disabled=True)]
        if template_id in ordered:
            ordered.remove(template_id)
        ordered.append(template_id)
        self.state_store.set_order(target, ordered)

    def _place_after(self, target: str, template_id: str, *, after_id: str) -> None:
        ordered = [item.id for item in self.list_templates(target, include_disabled=True)]
        if template_id in ordered:
            ordered.remove(template_id)
        try:
            index = ordered.index(after_id) + 1
        except ValueError:
            index = len(ordered)
        ordered.insert(index, template_id)
        self.state_store.set_order(target, ordered)

    def _remove_from_order(self, target: str, template_id: str) -> None:
        ordered = [item.id for item in self.list_templates(target, include_disabled=True) if item.id != template_id]
        self.state_store.set_order(target, ordered)

    def _find(self, template_id: str) -> TemplateInfo:
        for template in self.registry.list_templates(include_disabled=True):
            if template.id == template_id:
                return template
        raise TemplateNotFoundError(f"Template resource ID not found: {template_id}")

    def _require_custom(self, template: TemplateInfo) -> None:
        if not self.is_custom(template):
            raise TemplateManagementError("Built-in templates are read-only; copy the template before editing it")

    def _unique_destination(self, stem: str, target: str) -> Path:
        safe_stem = _safe_display_name(stem or "template")
        directory = self.ensure_user_directory()
        existing_names = {template.name.casefold() for template in self.list_templates(target, include_disabled=True)}
        candidate_stem = safe_stem
        counter = 2
        while candidate_stem.casefold() in existing_names or (directory / f"{candidate_stem}.{target}").exists():
            candidate_stem = f"{safe_stem} ({counter})"
            counter += 1
        return directory / f"{candidate_stem}.{target}"


__all__ = ["TemplateManagementError", "TemplateManager"]
