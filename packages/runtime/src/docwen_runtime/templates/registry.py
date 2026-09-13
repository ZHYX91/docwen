"""Content-validated template discovery and resolution."""

from __future__ import annotations

import hashlib
import logging
import re
import unicodedata
from contextlib import nullcontext
from dataclasses import dataclass
from pathlib import Path

from docwen_core.detection import inspect_file
from docwen_core.models import StructureStatus
from docwen_runtime.resources import ResourceRegistry

from .state import TemplateStateStore, user_templates_dir

logger = logging.getLogger(__name__)

_CANONICAL_TEMPLATE_ID_PATTERN = re.compile(r"^template\.(?:docx|xlsx)\.[0-9a-f]{64}$")


@dataclass(frozen=True)
class TemplateInfo:
    id: str
    name: str
    target: str
    description: str
    path: Path
    size_bytes: int
    modified_ns: int

    def to_dict(self) -> dict[str, object]:
        return {
            "id": self.id,
            "name": self.name,
            "target": self.target,
            "description": self.description,
            "path": str(self.path),
            "size_bytes": self.size_bytes,
            "modified_ns": self.modified_ns,
        }


class TemplateNotFoundError(LookupError):
    pass


class TemplateIdentityConflictError(RuntimeError):
    """Two discovered templates claim the same canonical resource identity."""

    diagnostic_code = "TEMPLATE_ID_CONFLICT"


class TemplateResolutionError(ValueError):
    """A user-selected template cannot be resolved or safely consumed."""

    def __init__(self, message: str, *, diagnostic_code: str) -> None:
        super().__init__(message)
        self.diagnostic_code = diagnostic_code


def validate_template_path(path: Path | str, *, expected_target: str) -> Path:
    """Validate one template by content and return its absolute path.

    Template filenames are presentation metadata only. A valid DOCX/XLSX
    package remains usable when renamed, while a misleading suffix never
    authorizes the wrong package type.
    """

    candidate = Path(path).expanduser().resolve(strict=False)
    if not candidate.is_file():
        raise TemplateResolutionError(
            f"Template file not found: {path}",
            diagnostic_code="TEMPLATE_NOT_FOUND",
        )
    try:
        inspection = inspect_file(str(candidate))
    except Exception as exc:
        raise TemplateResolutionError(
            f"Template could not be inspected: {candidate}: {exc}",
            diagnostic_code="TEMPLATE_INSPECTION_FAILED",
        ) from exc
    if inspection.detected_format != expected_target:
        raise TemplateResolutionError(
            (
                f"Template content is {inspection.detected_format or 'unknown'}, "
                f"but {expected_target} is required: {candidate}"
            ),
            diagnostic_code="TEMPLATE_FORMAT_MISMATCH",
        )
    if inspection.structure_status is not StructureStatus.VALID:
        raise TemplateResolutionError(
            f"Template package is not structurally valid: {candidate}",
            diagnostic_code="TEMPLATE_STRUCTURE_INVALID",
        )
    return candidate


class TemplateRegistry:
    """Discover structurally valid DOCX/XLSX templates by content.

    ``TemplateRegistry.default()`` merges immutable bundled templates with the
    writable per-user template directory. Explicit roots and an optional state
    store support isolated catalogs and external resource discovery.
    """

    def __init__(
        self,
        templates_dir: Path | str,
        extra_paths: list[Path] | None = None,
        *,
        state_store: TemplateStateStore | None = None,
        managed_user_dir: Path | str | None = None,
    ) -> None:
        self._dirs = [Path(templates_dir)] + (extra_paths or [])
        self._state_store = state_store
        self._managed_user_dir = Path(managed_user_dir) if managed_user_dir is not None else None

    @classmethod
    def default(
        cls,
        extra_paths: list[Path] | None = None,
        *,
        state_store: TemplateStateStore | None = None,
    ) -> TemplateRegistry:
        user_dir = user_templates_dir()
        state = state_store or TemplateStateStore.default()
        combined_paths: list[Path] = [user_dir]
        for path in extra_paths or []:
            candidate = Path(path)
            if candidate not in combined_paths:
                combined_paths.append(candidate)
        return cls(
            ResourceRegistry.default().templates_dir(),
            extra_paths=combined_paths,
            state_store=state,
            managed_user_dir=user_dir,
        )

    def list_templates(
        self,
        target_type: str | None = None,
        *,
        include_disabled: bool = False,
    ) -> list[TemplateInfo]:
        with self._state_store.locked() if self._state_store else nullcontext():
            return self._list_templates(target_type, include_disabled=include_disabled)

    def _list_templates(self, target_type: str | None, *, include_disabled: bool) -> list[TemplateInfo]:
        templates: list[TemplateInfo] = []
        identities: dict[str, Path] = {}
        for templates_dir in self._dirs:
            if not templates_dir.exists():
                continue
            for path in sorted(templates_dir.iterdir(), key=lambda item: item.name.casefold()):
                if path.is_symlink() or not path.is_file():
                    continue
                try:
                    inspection = inspect_file(str(path))
                except Exception as exc:
                    logger.debug("Skipping unreadable template candidate %s: %s", path, exc)
                    continue
                target = inspection.detected_format
                if target not in {"docx", "xlsx"} or inspection.structure_status is not StructureStatus.VALID:
                    continue
                if self._state_store is not None and self._is_managed_user_template(path):
                    template_id = self._state_store.ensure_user_identity(path, target)
                else:
                    template_id = _canonical_template_id(path.stem, target)
                conflicting_path = identities.get(template_id)
                if conflicting_path is not None:
                    raise TemplateIdentityConflictError(
                        f"Conflicting canonical template identity {template_id}: {conflicting_path} and {path}"
                    )
                identities[template_id] = path
                stat = path.stat()
                templates.append(
                    TemplateInfo(
                        id=template_id,
                        name=path.stem,
                        target=target,
                        description=_template_description(path.stem, target),
                        path=path,
                        size_bytes=stat.st_size,
                        modified_ns=stat.st_mtime_ns,
                    )
                )

        if self._state_store is not None:
            templates = self._apply_persisted_order(templates)
            if not include_disabled:
                templates = [template for template in templates if self._state_store.is_enabled(template.id)]

        if target_type is None:
            return templates
        return [template for template in templates if template.target == target_type]

    def get_template(self, identifier: str, target_type: str | None = None) -> TemplateInfo:
        """Resolve the exact canonical template resource ID."""

        if not is_canonical_template_id(identifier):
            raise TemplateNotFoundError(f"模板资源 ID 无效: {identifier}")

        templates = self.list_templates(target_type)
        if not templates:
            raise TemplateNotFoundError("没有可用模板")

        for template in templates:
            if template.id == identifier:
                return template
        raise TemplateNotFoundError(f"模板资源 ID 不存在或目标类型不匹配: {identifier}")

    def _is_managed_user_template(self, path: Path) -> bool:
        if self._managed_user_dir is None:
            return False
        try:
            return path.parent.resolve(strict=False) == self._managed_user_dir.resolve(strict=False)
        except OSError:
            return path.parent == self._managed_user_dir

    def _apply_persisted_order(self, templates: list[TemplateInfo]) -> list[TemplateInfo]:
        if self._state_store is None:
            return templates
        ordered: list[TemplateInfo] = []
        for target in ("docx", "xlsx"):
            candidates = [template for template in templates if template.target == target]
            by_id = {template.id: template for template in candidates}
            ids = self._state_store.ordered_ids(target, [template.id for template in candidates])
            ordered.extend(by_id[template_id] for template_id in ids if template_id in by_id)
        ordered_ids = {template.id for template in ordered}
        ordered.extend(template for template in templates if template.id not in ordered_ids)
        return ordered


def _canonical_template_id(name: str, target: str) -> str:
    normalized_name = unicodedata.normalize("NFC", name).strip().casefold()
    identity = f"{target}\0{normalized_name}".encode()
    return f"template.{target}.{hashlib.sha256(identity).hexdigest()}"


def is_canonical_template_id(value: str) -> bool:
    """Return whether *value* is a protocol 3 canonical template resource ID."""

    return _CANONICAL_TEMPLATE_ID_PATTERN.fullmatch(value) is not None


def _template_description(name: str, target: str) -> str:
    if target == "xlsx":
        return f"{name} Excel 模板"
    return f"{name} DOCX 模板"
