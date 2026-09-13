"""User-facing template management operations.

All writes are confined to the platform user-data template directory.  Bundled
application templates are treated as immutable resources and can only be copied
into the user directory before editing.
"""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Iterable

from docwen_core.detection import inspect_file
from docwen_core.models import StructureStatus

from .registry import TemplateInfo, TemplateNotFoundError, TemplateRegistry
from .state import TemplateStateStore, user_templates_dir


class TemplateManagementError(RuntimeError):
    """A requested template-management operation cannot be completed safely."""


def _safe_display_name(value: str) -> str:
    name = value.strip()
    if not name or name in {".", ".."}:
        raise TemplateManagementError("Template name cannot be empty")
    if Path(name).name != name or any(character in name for character in '<>:"/\\|?*'):
        raise TemplateManagementError("Template name contains invalid path characters")
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
    def default(cls) -> "TemplateManager":
        state = TemplateStateStore.default()
        return cls(
            TemplateRegistry.default(state_store=state),
            state_store=state,
            user_dir=user_templates_dir(),
        )

    def ensure_user_directory(self) -> Path:
        self.user_dir.mkdir(parents=True, exist_ok=True)
        return self.user_dir

    def list_templates(self, target_type: str | None = None, *, include_disabled: bool = True) -> list[TemplateInfo]:
        return self.registry.list_templates(target_type, include_disabled=include_disabled)

    def is_custom(self, template: TemplateInfo) -> bool:
        try:
            return template.path.parent.resolve(strict=False) == self.user_dir.resolve(strict=False)
        except OSError:
            return template.path.parent == self.user_dir

    def is_enabled(self, template_id: str) -> bool:
        return self.state_store.is_enabled(template_id)

    def set_enabled(self, template_id: str, enabled: bool) -> None:
        self._find(template_id)
        self.state_store.set_enabled(template_id, enabled)

    def import_templates(self, paths: Iterable[Path | str]) -> list[TemplateInfo]:
        imported: list[TemplateInfo] = []
        for path in paths:
            imported.append(self.import_template(path))
        return imported

    def import_template(self, path: Path | str) -> TemplateInfo:
        source = Path(path)
        target = _detect_valid_target(source)
        destination = self._unique_destination(source.stem, target)
        self.ensure_user_directory()
        shutil.copy2(source, destination)
        template_id = self.state_store.ensure_user_identity(destination, target)
        self.state_store.set_enabled(template_id, True)
        self._place_at_end(target, template_id)
        return self._find(template_id)

    def copy_builtin_as_custom(self, template_id: str, *, custom_name: str | None = None) -> TemplateInfo:
        template = self._find(template_id)
        if self.is_custom(template):
            raise TemplateManagementError("Only built-in templates use the copy-as-custom operation")
        base_name = _safe_display_name(custom_name) if custom_name is not None else f"{template.name} - custom"
        destination = self._unique_destination(base_name, template.target)
        self.ensure_user_directory()
        shutil.copy2(template.path, destination)
        custom_id = self.state_store.ensure_user_identity(destination, template.target)
        self.state_store.set_enabled(custom_id, True)
        self._place_after(template.target, custom_id, after_id=template.id)
        return self._find(custom_id)

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
        template.path.rename(destination)
        self.state_store.rename_user_identity(old_name, destination.name)
        return self._find(template_id)

    def delete_custom(self, template_id: str) -> None:
        """Move a custom template to the OS recycle bin, never hard-delete it."""

        template = self._find(template_id)
        self._require_custom(template)
        filename = template.path.name
        try:
            from send2trash import send2trash  # type: ignore[import-not-found]
        except ImportError as exc:
            raise TemplateManagementError(
                "Safe recycle-bin support is unavailable; the template was not deleted"
            ) from exc
        try:
            send2trash(str(template.path))
        except Exception as exc:
            raise TemplateManagementError(
                f"Could not move template to the recycle bin: {template.path.name}"
            ) from exc
        self.state_store.forget_user_identity(filename)
        self._remove_from_order(template.target, template.id)

    def export_custom(self, template_id: str, destination: Path | str) -> Path:
        template = self._find(template_id)
        self._require_custom(template)
        destination_path = Path(destination).expanduser()
        if destination_path.exists() and destination_path.is_dir():
            destination_path = destination_path / template.path.name
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template.path, destination_path)
        return destination_path

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
        existing_names = {
            template.name.casefold()
            for template in self.list_templates(target, include_disabled=True)
        }
        candidate_stem = safe_stem
        counter = 2
        while candidate_stem.casefold() in existing_names or (directory / f"{candidate_stem}.{target}").exists():
            candidate_stem = f"{safe_stem} ({counter})"
            counter += 1
        return directory / f"{candidate_stem}.{target}"


__all__ = ["TemplateManagementError", "TemplateManager"]
