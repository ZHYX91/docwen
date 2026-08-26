"""InputAreaViewModel — observable state for the file drop / input area.

This is the single source of truth for the InputArea widget's observable
state.  Widgets bind to its signals and properties; user actions flow
through method calls that delegate to ``MainWindowViewModel``.

Widgets never call runtime/plugins directly — they go through this
ViewModel which delegates to the parent ``MainWindowViewModel``.
"""

from __future__ import annotations

import logging
import re
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

from PySide6.QtCore import QObject, QUrl, Signal

from docwen_core.paths import scan_input_directory
from docwen_gui.file_types import FILE_CATEGORY_ORDER, FILE_EXTENSIONS_BY_CATEGORY
from docwen_gui.i18n import t as _t

if TYPE_CHECKING:
    from docwen_core.models.file_ref import FileRef

    from .main_window_vm import MainWindowViewModel

# Maximum number of paths to parse from a text drag payload
_TEXT_PAYLOAD_MAX_PATHS = 64

# Maximum files to scan in drag-preview folder recursion
_BATCH_SCAN_LIMIT = 200
_DRAG_PREVIEW_SKIPPED_SAMPLE_LIMIT = 3

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class DragPreview:
    """Preview message for drag-enter feedback."""

    message: str
    tooltip: str
    tone: str
    added_count: int = 0
    skipped_count: int = 0
    has_recursive_scan: bool = False
    has_degraded_preview: bool = False


@dataclass(frozen=True)
class _BatchCollection:
    files: list[str]
    skipped_count: int = 0


def _read_default_mode(main_vm: MainWindowViewModel) -> str:
    controller = getattr(main_vm, "controller", None)
    cfg = getattr(controller, "config_port", None) if controller is not None else None
    try:
        mode = cfg.get("gui.window.default_mode", "single") if cfg is not None else "single"
        if mode in ("single", "batch"):
            return mode
    except Exception:
        logger.warning("Default input mode could not be read; using single mode", exc_info=True)
        return "single"
    return "single"


class InputAreaViewModel(QObject):
    """Observable state for the InputArea widget.

    Signals:
        mode_changed: emitted when single/batch mode changes.
        files_added: emitted with list of file paths selected/dropped.
        files_cleared: emitted when clear is requested.
        selection_message_changed: emitted with (message, tone) for user feedback.
    """

    # ── Signals ──────────────────────────────────────────────────────

    mode_changed = Signal(str)
    """Emitted when single/batch mode changes."""

    files_added = Signal(list)
    """Emitted when files are added (payload: list[str] of absolute paths)."""

    files_cleared = Signal()
    """Emitted when all files are cleared."""

    selection_message_changed = Signal(str, str)
    """Emitted with (message, tone) for selection feedback."""

    # ── Construction ─────────────────────────────────────────────────

    def __init__(
        self,
        main_vm: MainWindowViewModel,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._main_vm = main_vm
        self._selection_message: str = ""
        self._selection_detail: str = ""
        self._selection_tone: str = "secondary"
        self._file_filter: Callable[[str], bool] | None = None

        # Read default mode from config, then honor the owning VM's current mode.
        default_mode = _read_default_mode(main_vm)
        if main_vm.mode != default_mode:
            # Main VM takes precedence if it has a different value (e.g. set by IPC)
            default_mode = main_vm.mode
        self._mode: str = default_mode

    # ── Observable properties ────────────────────────────────────────

    @property
    def mode(self) -> str:
        """Current mode: ``"single"`` or ``"batch"``."""
        return self._mode

    @property
    def selection_message(self) -> str:
        """Current selection feedback message (empty if none)."""
        return self._selection_message

    @property
    def selection_tone(self) -> str:
        """Current selection feedback tone: success/warning/danger/info/secondary."""
        return self._selection_tone

    @property
    def selection_detail(self) -> str:
        """Secondary selection detail shown below the primary feedback."""
        return self._selection_detail

    @property
    def file_filter(self) -> Callable[[str], bool] | None:
        """Optional custom file filter for validation."""
        return self._file_filter

    @file_filter.setter
    def file_filter(self, fn: Callable[[str], bool] | None) -> None:
        self._file_filter = fn

    # ── Command methods ──────────────────────────────────────────────

    def set_mode(self, mode: str) -> None:
        """Set single/batch mode.

        Args:
            mode: ``"single"`` or ``"batch"``.
        """
        if mode not in ("single", "batch"):
            raise ValueError(f"Invalid mode: {mode!r}")
        if mode == self._mode:
            return
        self._mode = mode
        self.mode_changed.emit(mode)
        self._main_vm.set_mode(mode)

    def add_files(self, paths: list[str]) -> None:
        """Validate and add file paths.

        For single mode: rejects folders, requires exactly 1 supported file.
        For batch mode: accepts multiple files/folders with recursive collection.

        Args:
            paths: List of absolute file/folder paths.
        """
        normalized = [str(Path(p)) for p in paths if p]
        if not normalized:
            return

        if self._mode == "single":
            self._add_single(normalized)
        else:
            self._add_batch(normalized)

    def _add_single(self, paths: list[str]) -> None:
        folder_paths = [p for p in paths if Path(p).is_dir()]
        if folder_paths:
            self._emit_message(
                _t("messages.no_folder_in_single_mode", "Single mode does not support folders"),
                "warning",
            )
            return

        file_paths = [p for p in paths if Path(p).is_file()]
        if len(file_paths) != 1:
            self._emit_message(
                _t("components.file_drop.single_mode_only_one", "Please select exactly one file in single mode"),
                "warning",
            )
            return

        file_path = file_paths[0]
        supported = self._is_supported(file_path)
        if not supported:
            self._emit_message(
                _t(
                    "components.file_drop.unsupported_type_msg",
                    "Unsupported file type: {filename}",
                    filename=Path(file_path).name,
                ),
                "danger",
            )
            return

        # MainWindowViewModel owns the one canonical content inspection and
        # stores the result on FileRef; this renderer never inspects twice.
        self._emit_files_added([file_path])

    def _add_batch(self, paths: list[str]) -> None:
        collection = self._collect_batch_files_with_feedback(paths)
        if not collection.files:
            self._emit_message(
                _t("components.file_drop.no_supported_files", "No supported files found"),
                "warning",
            )
            return
        self._emit_files_added(collection.files, skipped_count=collection.skipped_count)

    def _collect_batch_files(self, paths: list[str]) -> list[str]:
        return self._collect_batch_files_with_feedback(paths).files

    def _collect_batch_files_with_feedback(self, paths: list[str]) -> _BatchCollection:
        collected: list[str] = []
        seen: set[str] = set()
        skipped_count = 0
        for raw_path in paths:
            path = Path(raw_path)
            candidates: list[Path] = [path]
            if path.is_dir():
                scan = scan_input_directory(path)
                candidates = list(scan.files)
                skipped_count += len(scan.unreadable_paths)
            for candidate in candidates:
                normalized = str(candidate)
                key = normalized.casefold()
                if key in seen:
                    continue
                seen.add(key)
                if self._is_supported(normalized):
                    collected.append(normalized)
                elif candidate.exists():
                    skipped_count += 1
                else:
                    skipped_count += 1
        return _BatchCollection(collected, skipped_count)

    def clear_files(self) -> None:
        """Clear all files and reset selection state."""
        self._emit_message("", "secondary")
        self.files_cleared.emit()

    def sync_selection(self, file_refs: Sequence[FileRef]) -> None:
        """Synchronize externally-added files into the visual selection state.

        Unlike :meth:`add_files`, this renderer-facing method never validates,
        re-adds, or emits ``files_added``.  It is used for startup/IPC paths
        that already entered the owning ``MainWindowViewModel``.
        """
        refs = [ref for ref in file_refs if ref.path]
        normalized = [str(Path(ref.path)) for ref in refs]
        if not normalized:
            self._emit_message("", "secondary")
            return
        warning_message = self._selection_warning_from_refs(refs)
        if warning_message:
            detail = str(Path(normalized[0]).parent) if self._mode == "single" else ""
            self._emit_message(warning_message, "warning", detail=detail)
            return
        if self._mode == "single":
            file_path = normalized[0]
            message = _t(
                "components.file_drop.file_selected_msg",
                "Selected: {filename}",
                filename=Path(file_path).name,
            )
            self._emit_message(message, "success", detail=str(Path(file_path).parent))
            return
        message = _t(
            "components.file_drop.files_added_msg",
            "Added {count} file(s)",
            count=len(normalized),
        )
        self._emit_message(message, "success")

    @staticmethod
    def _selection_warning_from_refs(file_refs: Sequence[FileRef]) -> str:
        from docwen_core.models import FILE_INSPECTION_METADATA_KEY

        for ref in file_refs:
            warning = str(ref.warning_message or "").strip()
            if warning:
                return warning
            inspection = ref.metadata.get(FILE_INSPECTION_METADATA_KEY)
            if not isinstance(inspection, dict):
                continue
            for key in ("warning_message", "reason_message"):
                message = str(inspection.get(key) or "").strip()
                if message:
                    return message
        return ""

    def request_add_dialog(self, *, force_batch_mode: bool = False) -> None:
        """Request that the add-file dialog be opened.

        The actual dialog display is handled by the widget layer.
        This signals intent.

        Args:
            force_batch_mode: If True, switch to batch mode before opening.
        """
        if force_batch_mode and self._mode != "batch":
            self.set_mode("batch")

    def request_add_folder_dialog(self, *, force_batch_mode: bool = False) -> None:
        """Request that the add-folder dialog be opened.

        Args:
            force_batch_mode: If True, switch to batch mode before opening.
        """
        if force_batch_mode and self._mode != "batch":
            self.set_mode("batch")

    def build_drag_preview(self, paths: list[str]) -> DragPreview:
        """Build the feedback shown while files or folders hover over the drop area."""
        normalized = [str(Path(path)) for path in paths if path]
        if not normalized:
            return DragPreview(_t("components.file_drop.no_supported_files", "No supported files found"), "", "warning")

        if self._mode == "single":
            return self._build_single_drag_preview(normalized)
        return self._build_batch_drag_preview(normalized)

    # ── Drag-and-drop text parsing ───────────────────────────────────

    @staticmethod
    def extract_paths_from_text_payload(text: str) -> list[str]:
        """Parse file paths from a text drag-and-drop payload.

        Supports:
        - {braced paths} for paths containing spaces
        - space-separated paths
        - file:// URLs
        - quote-wrapped paths
        - deduplication (case-insensitive)

        Args:
            text: Raw text from drag-and-drop MIME data.

        Returns:
            List of resolved absolute file paths.
        """
        raw = (text or "").strip()
        if not raw:
            return []

        candidates: list[str] = []
        brace_pattern = re.compile(r"\{([^}]+)\}")
        candidates.extend(brace_pattern.findall(raw))
        remaining = brace_pattern.sub("\n", raw)
        candidates.extend([line.strip() for line in remaining.splitlines() if line.strip()])

        paths: list[str] = []
        seen: set[str] = set()

        def _strip_quotes(value: str) -> str:
            value = value.strip()
            if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
                return value[1:-1].strip()
            return value

        def _maybe_add(value: str) -> None:
            nonlocal paths
            if len(paths) >= _TEXT_PAYLOAD_MAX_PATHS:
                return
            value = _strip_quotes(value)
            if not value:
                return
            if value.lower().startswith("file:"):
                url = QUrl(value)
                if url.isValid() and url.isLocalFile():
                    value = url.toLocalFile()
            candidate_path = Path(value)
            if not candidate_path.is_absolute():
                return
            try:
                exists = candidate_path.exists()
            except OSError:
                # A text payload may contain an oversized or otherwise invalid
                # path-shaped token. Treat it as unusable input instead of
                # letting the drag/drop event escape with an OS path error.
                return
            if not exists:
                return
            normalized = str(candidate_path)
            key = normalized.casefold()
            if key in seen:
                return
            seen.add(key)
            paths.append(normalized)

        for candidate in candidates:
            candidate = candidate.strip()
            if not candidate:
                continue
            _maybe_add(candidate)
            if len(paths) >= _TEXT_PAYLOAD_MAX_PATHS:
                break
            if (" " in candidate or "\t" in candidate) and len(candidate.split()) > 1:
                for token in candidate.split():
                    _maybe_add(token)
                    if len(paths) >= _TEXT_PAYLOAD_MAX_PATHS:
                        break

        return paths

    @staticmethod
    def extract_urls_from_mime_data(urls: list[QUrl]) -> list[str]:
        """Extract local file paths from MIME URL data.

        Args:
            urls: List of QUrl objects from MIME data.

        Returns:
            List of local file paths.
        """
        return [url.toLocalFile() for url in urls if url.isLocalFile()]

    # ── File dialog filter construction ──────────────────────────────

    def build_file_dialog_filter(self) -> str:
        """Build the Qt file dialog filter string.

        Returns:
            Filter string for QFileDialog (e.g. "All Supported (*.docx ...);;...;;All Files (*.*)").
        """
        all_patterns = self._build_dialog_patterns(
            [extension for category in FILE_CATEGORY_ORDER for extension in FILE_EXTENSIONS_BY_CATEGORY[category]]
        )
        filetypes: list[tuple[str, str]] = [(_t("components.file_drop.all_supported_filter"), all_patterns)]
        for category in FILE_CATEGORY_ORDER:
            label = _t(f"file_types.{category}", default=category)
            patterns = self._build_dialog_patterns(list(FILE_EXTENSIONS_BY_CATEGORY[category]))
            filetypes.append((label, patterns))
        filetypes.append((_t("components.file_drop.all_files_filter"), "*.*"))
        return ";;".join([f"{label} ({patterns})" for label, patterns in filetypes])

    @staticmethod
    def _build_dialog_patterns(extensions: list[str]) -> str:
        unique: list[str] = []
        seen: set[str] = set()
        for ext in extensions:
            normalized = ext.strip().lower()
            if not normalized or normalized in seen:
                continue
            seen.add(normalized)
            unique.append(f"*{normalized}")
        return " ".join(unique)

    # ── Internal helpers ─────────────────────────────────────────────

    def _is_supported(self, file_path: str) -> bool:
        if self._file_filter is not None:
            return self._file_filter(file_path)
        # A filename suffix is only a declaration. The owning MainWindow VM
        # performs content inspection and decides ALLOW / explicit acceptance /
        # BLOCK, so the drop surface must forward every regular file.
        return Path(file_path).is_file()

    def _build_single_drag_preview(self, paths: list[str]) -> DragPreview:
        folder_paths = [path for path in paths if Path(path).is_dir()]
        if folder_paths:
            return DragPreview(
                _t("messages.no_folder_in_single_mode", "Single mode does not support folders"),
                "\n".join(folder_paths),
                "warning",
            )

        file_paths = [path for path in paths if Path(path).is_file()]
        if len(file_paths) != 1:
            return DragPreview(
                _t("components.file_drop.single_mode_only_one", "Please select exactly one file in single mode"),
                "\n".join(paths),
                "warning",
            )

        file_path = file_paths[0]
        if not self._is_supported(file_path):
            return DragPreview(
                _t(
                    "components.file_drop.unsupported_type_msg",
                    "Unsupported file type: {filename}",
                    filename=Path(file_path).name,
                ),
                file_path,
                "danger",
                skipped_count=1,
            )

        return DragPreview(
            _t("components.file_drop.drag_preview_single", "Will select: {filename}", filename=Path(file_path).name),
            file_path,
            "info",
            added_count=1,
        )

    def _build_batch_drag_preview(self, paths: list[str]) -> DragPreview:
        added: list[str] = []
        skipped_samples: list[str] = []
        skipped_count = 0
        seen: set[str] = set()
        scanned = 0
        has_recursive_scan = False
        has_degraded_preview = False

        for raw_path in paths:
            if scanned >= _BATCH_SCAN_LIMIT:
                has_degraded_preview = True
                break
            path = Path(raw_path)
            if not path.exists():
                skipped_count += 1
                if len(skipped_samples) < _DRAG_PREVIEW_SKIPPED_SAMPLE_LIMIT:
                    skipped_samples.append(
                        _t("components.file_drop.drag_preview_skipped_unreadable", "Unreadable: {path}", path=raw_path)
                    )
                continue

            candidates: list[Path] = [path]
            if path.is_dir():
                has_recursive_scan = True
                scan = scan_input_directory(path, limit=_BATCH_SCAN_LIMIT - scanned)
                candidates = list(scan.files)
                scanned += len(candidates)
                has_degraded_preview = has_degraded_preview or scan.truncated
                skipped_count += len(scan.unreadable_paths)
                for unreadable in scan.unreadable_paths:
                    if len(skipped_samples) >= _DRAG_PREVIEW_SKIPPED_SAMPLE_LIMIT:
                        break
                    skipped_samples.append(
                        _t(
                            "components.file_drop.drag_preview_skipped_unreadable",
                            "Unreadable: {path}",
                            path=str(unreadable),
                        )
                    )

            for candidate in candidates:
                normalized = str(candidate)
                key = normalized.casefold()
                if key in seen:
                    continue
                seen.add(key)
                if candidate.is_file() and self._is_supported(normalized):
                    added.append(normalized)
                elif candidate.is_file():
                    skipped_count += 1
                    if len(skipped_samples) < _DRAG_PREVIEW_SKIPPED_SAMPLE_LIMIT:
                        skipped_samples.append(
                            _t(
                                "components.file_drop.drag_preview_skipped_unsupported",
                                "Unsupported: {name}",
                                name=candidate.name,
                            )
                        )

        tooltip_parts: list[str] = []
        if added:
            tooltip_parts.extend(added)
        if skipped_samples:
            tooltip_parts.append(_t("components.file_drop.drag_preview_skipped_title", "Will skip:"))
            tooltip_parts.extend(skipped_samples)

        if not added:
            return DragPreview(
                _t("components.file_drop.no_supported_files", "No supported files found"),
                "\n".join(tooltip_parts or paths),
                "warning",
                skipped_count=skipped_count,
                has_recursive_scan=has_recursive_scan,
                has_degraded_preview=has_degraded_preview,
            )

        summary = _t(
            "components.file_drop.drag_preview_summary",
            "Will add {added}, skip {skipped}",
            added=len(added),
            skipped=skipped_count,
        )
        if has_recursive_scan:
            summary = _t(
                "components.file_drop.drag_preview_summary_recursive",
                "{summary}, including recursive folder scan",
                summary=summary,
            )
        if has_degraded_preview:
            summary = _t(
                "components.file_drop.drag_preview_summary_degraded",
                "{summary}; large batch preview is summarized",
                summary=summary,
            )

        return DragPreview(
            summary,
            "\n".join(tooltip_parts),
            "info",
            added_count=len(added),
            skipped_count=skipped_count,
            has_recursive_scan=has_recursive_scan,
            has_degraded_preview=has_degraded_preview,
        )

    def _emit_message(self, message: str, tone: str, *, detail: str = "") -> None:
        self._selection_message = message
        self._selection_detail = detail
        self._selection_tone = tone
        self.selection_message_changed.emit(message, tone)

    def _emit_files_added(
        self,
        paths: list[str],
        *,
        skipped_count: int = 0,
        warning_message: str = "",
    ) -> None:
        outcome = self._main_vm.add_files(paths)
        admitted_paths = [ref.path for ref in outcome.added]
        if not admitted_paths:
            if outcome.rejected:
                self._emit_message(outcome.rejected[0][1], "danger")
            return
        paths = admitted_paths
        skipped_count += len(outcome.rejected)
        normalized_paths = {str(Path(path)) for path in admitted_paths}
        warnings = [
            str(ref.warning_message)
            for ref in outcome.added
            if str(Path(ref.path)) in normalized_paths and ref.warning_message
        ]
        if warnings:
            warning_message = warnings[0]

        file_count = len(paths)
        if warning_message and self._mode == "single":
            msg = warning_message
            tone = "warning"
        elif self._mode == "single":
            msg = _t(
                "components.file_drop.file_selected_msg",
                "Selected: {filename}",
                filename=Path(paths[0]).name if paths else "",
            )
            tone = "success"
        elif skipped_count > 0:
            msg = _t(
                "components.file_drop.files_added_with_skipped_msg",
                "Added {added} file(s), skipped {skipped}",
                added=file_count,
                skipped=skipped_count,
            )
            tone = "warning"
        elif warning_message:
            msg = _t(
                "components.file_drop.files_added_msg",
                "Added {count} file(s)",
                count=file_count,
            )
            tone = "warning"
        else:
            msg = _t(
                "components.file_drop.files_added_msg",
                "Added {count} file(s)",
                count=file_count,
            )
            tone = "success"
        detail = str(Path(paths[0]).parent) if self._mode == "single" and paths else warning_message
        self._emit_message(msg, tone, detail=detail)
        self.files_added.emit(paths)


__all__ = [
    "_TEXT_PAYLOAD_MAX_PATHS",
    "DragPreview",
    "InputAreaViewModel",
]
