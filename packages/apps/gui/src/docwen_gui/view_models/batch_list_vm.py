"""BatchListViewModel — observable state for the BatchList widget.

This is the single source of truth for the BatchList widget's observable
state.  Widgets bind to its signals and properties; user actions flow
through method calls that delegate to ``MainWindowViewModel``.

Widgets never call runtime/plugins directly — they go through this
ViewModel which delegates to the parent ``MainWindowViewModel``.

State managed:
  - 6 category tabs (text/spreadsheet/document/image/layout/other)
  - File entries with per-entry status, metadata
  - Filter (all/pending/processing/completed/failed/skipped/cancelled)
  - Sort (custom/name/type/size/mtime, asc/desc)
  - Current tab, selection, custom ordering
"""

from __future__ import annotations

import contextlib
import os
from collections.abc import Callable, Mapping
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Any

from PySide6.QtCore import QObject, Signal

from docwen_core.formats.categories import ALL_CATEGORIES
from docwen_core.models import FILE_INSPECTION_METADATA_KEY, AdmissionDecision
from docwen_gui.file_admission_i18n import render_file_inspection_message


def _normalize_path(file_path: str) -> str:
    """Normalize a file path to a consistent key format (forward slashes)."""
    return str(Path(file_path)).replace(os.sep, "/")


if TYPE_CHECKING:
    from .main_window_vm import MainWindowViewModel

# ── Category constants ──────────────────────────────────────────────────

CATEGORY_ORDER = ["text", "spreadsheet", "document", "image", "layout", "other"]

# ── Filter constants ────────────────────────────────────────────────────

FILTER_OPTIONS: tuple[tuple[str, str, tuple[str, ...]], ...] = (
    ("all", "All", ()),
    ("pending", "Pending", ("pending",)),
    ("processing", "Processing", ("processing",)),
    ("completed", "Completed", ("completed",)),
    ("failed", "Failed", ("failed",)),
    ("skipped", "Skipped", ("skipped",)),
    ("cancelled", "Cancelled", ("cancelled",)),
)

# This is both the runtime state domain and the finite i18n suffix contract
# consumed by ``BatchList``.  Unknown states must never reach a dynamic
# translation lookup or become an unfilterable row.
BATCH_FILE_STATUSES = frozenset(status for _key, _label, statuses in FILTER_OPTIONS for status in statuses)

# ── Sort constants ──────────────────────────────────────────────────────

SORT_KEYS = ("custom", "name", "type", "size", "mtime")

# ── Pulse animation limit ───────────────────────────────────────────────

BATCH_STATUS_PULSE_ENTRY_LIMIT = 40

# ── Sentinel for unset parameters ───────────────────────────────────────

_UNSET = object()

# ── Data classes ────────────────────────────────────────────────────────


@dataclass
class BatchFileEntry:
    """Represents a single file in the batch list.

    Mirrors the old ``batch_widgets.entry_widget.BatchFileEntry`` data class.
    """

    file_path: str
    file_name: str
    detected_format: str
    workflow_category: str
    warning_message: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)
    size_bytes: int = 0
    status: str = "pending"  # pending/processing/completed/skipped/failed/cancelled
    output_path: str | None = None
    skip_reason: str | None = None
    error_message: str | None = None
    error_count: int = 0  # Number of diagnostics/errors for this file
    operation_id: str | None = None

    def __post_init__(self) -> None:
        detected_format = str(self.detected_format).strip().lower()
        workflow_category = str(self.workflow_category).strip().lower()
        status = str(self.status).strip().lower()
        if not detected_format:
            raise ValueError("detected_format must not be empty")
        if workflow_category not in ALL_CATEGORIES:
            raise ValueError("workflow_category must be canonical")
        if status not in BATCH_FILE_STATUSES:
            raise ValueError(f"status must be one of {sorted(BATCH_FILE_STATUSES)!r}")
        self.detected_format = detected_format
        self.workflow_category = workflow_category
        self.status = status


@dataclass
class BatchSelectionSnapshot:
    """Captured selection state for restore after filter/sort changes."""

    current_file: str | None
    selected_files: list[str] = field(default_factory=list)


# ── Helpers ─────────────────────────────────────────────────────────────


def format_size(size_bytes: int) -> str:
    """Format byte count into human-readable string."""
    value = float(max(0, size_bytes))
    units = ["B", "KB", "MB", "GB"]
    for unit in units:
        if value < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(value)} {unit}"
            return f"{value:.1f} {unit}"
        value /= 1024.0
    return f"{int(size_bytes)} B"


def _sort_value(entry: BatchFileEntry, sort_key: str) -> float | str:
    """Compute sort value for an entry given a sort key."""
    if sort_key == "name":
        return str(entry.file_name or "").casefold()
    if sort_key == "type":
        return str(entry.detected_format or "").casefold()
    if sort_key == "size":
        return int(entry.size_bytes) if entry.size_bytes >= 0 else 0
    if sort_key == "mtime":
        try:
            return float(Path(entry.file_path).stat().st_mtime)
        except OSError:
            return 0.0
    return 0  # custom


def should_pulse_processing_transition(entry_count: int) -> bool:
    """Whether pulse animation is allowed for the given entry count."""
    return entry_count <= BATCH_STATUS_PULSE_ENTRY_LIMIT


# ── ViewModel ───────────────────────────────────────────────────────────


class BatchListViewModel(QObject):
    """Observable state for the BatchList widget.

    Signals:
        files_added: emitted with (added_paths, failed_paths_reasons).
        files_removed: emitted with removed file path.
        status_changed: emitted when a file status changes.
        filter_changed: emitted with new filter key.
        sort_changed: emitted with (sort_key, ascending).
        current_category_changed: emitted with new category.
        selection_changed: emitted with current file path (or None).
        entry_count_changed: emitted with total entry count.
        pulse_requested: emitted when a processing pulse animation is needed.
    """

    # ── Signals ────────────────────────────────────────────────────────

    files_added = Signal(list, list)
    """Emitted with (added: list[str], failed: list[tuple[str, str]])."""

    files_removed = Signal(str)
    """Emitted when a file is removed (payload: file_path)."""

    files_cleared = Signal()
    """Emitted when all files are cleared."""

    status_changed = Signal(str, str)
    """Emitted when file status changes (file_path, new_status)."""

    filter_changed = Signal(str)
    """Emitted when filter state changes (payload: filter_key)."""

    sort_changed = Signal(str, bool)
    """Emitted when sort changes (payload: sort_key, ascending)."""

    current_category_changed = Signal(str)
    """Emitted when current tab changes (payload: category)."""

    selection_changed = Signal(object)
    """Emitted when selection changes (payload: current_file_path or None)."""

    entry_count_changed = Signal(int)
    """Emitted when total entry count changes."""

    pulse_requested = Signal(str)
    """Emitted when a processing pulse animation should play (payload: file_path)."""

    reorder_requested = Signal(str)
    """Emitted when manual reorder occurs (payload: category)."""

    # ── Construction ───────────────────────────────────────────────────

    def __init__(
        self,
        main_vm: MainWindowViewModel | None = None,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._main_vm = main_vm
        self._entries: dict[str, tuple[str, BatchFileEntry]] = {}
        """Map file_path -> (category, BatchFileEntry)."""
        self._current_category: str = CATEGORY_ORDER[0]
        self._active_filter: str = "all"
        self._sort_key: str = "custom"
        self._sort_ascending: bool = True
        self._custom_order_by_category: dict[str, list[str]] = {cat: [] for cat in CATEGORY_ORDER}
        self._suspended: bool = False

    # ── Properties ─────────────────────────────────────────────────────

    @property
    def current_category(self) -> str:
        return self._current_category

    @property
    def active_filter(self) -> str:
        return self._active_filter

    @property
    def sort_key(self) -> str:
        return self._sort_key

    @property
    def sort_ascending(self) -> bool:
        return self._sort_ascending

    @property
    def entry_count(self) -> int:
        return len(self._entries)

    @property
    def is_suspended(self) -> bool:
        return self._suspended

    # ── File operations ────────────────────────────────────────────────

    def add_files(
        self,
        file_paths: list[str],
        *,
        file_resolver: Callable[[str], Mapping[str, Any] | None] | None = None,
    ) -> tuple[list[str], list[tuple[str, str]]]:
        """Add files to the batch list.

        ``file_resolver`` is an optional callable ``(path) -> dict | None``
        that resolves ``detected_format`` and Core-owned
        ``workflow_category`` metadata. Without a resolver, the real file is
        inspected by Core; missing paths are never classified from a suffix.
        """
        added: list[str] = []
        failed: list[tuple[str, str]] = []

        for file_path in file_paths:
            normalized = _normalize_path(file_path)
            if normalized in self._entries:
                continue

            info: Mapping[str, Any] | None = None
            if file_resolver is not None:
                try:
                    info = file_resolver(normalized)
                except Exception as exc:
                    failed.append((normalized, str(exc)))
                    continue
                if info is None:
                    failed.append((normalized, "Unsupported file type"))
                    continue
            else:
                # Use content-based format detection
                check_path = str(Path(file_path))
                try:
                    from docwen_core.detection import inspect_file

                    inspection = inspect_file(check_path)
                    if inspection.decision is AdmissionDecision.BLOCK:
                        failed.append(
                            (
                                normalized,
                                render_file_inspection_message(inspection, prefer_reason=True)
                                or "Unsupported file content",
                            )
                        )
                        continue
                    detected_format = inspection.detected_format
                    validation = {"warning_message": render_file_inspection_message(inspection)}
                except (OSError, ValueError) as exc:
                    failed.append((normalized, str(exc) or "File is unavailable"))
                    continue

                workflow_category = inspection.workflow_category
                info = {
                    "detected_format": detected_format,
                    "workflow_category": workflow_category,
                    "warning_message": validation.get("warning_message") or None,
                    "metadata": {FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
                }

            detected_format = info.get("detected_format")
            if not isinstance(detected_format, str) or not detected_format.strip():
                failed.append((normalized, "File detected format is unavailable"))
                continue
            detected_format = detected_format.strip().lower()
            workflow_category = str(info.get("workflow_category") or "").strip().lower()
            if workflow_category not in ALL_CATEGORIES:
                failed.append((normalized, "File workflow category is unavailable"))
                continue
            display_category = self._display_category(workflow_category)

            size_bytes = 0
            # Use original file_path for filesystem operations
            fs_path = str(Path(file_path))
            with contextlib.suppress(OSError):
                size_bytes = Path(fs_path).stat().st_size

            entry = BatchFileEntry(
                file_path=normalized,
                file_name=Path(fs_path).name,
                detected_format=detected_format,
                workflow_category=workflow_category,
                warning_message=info.get("warning_message") or None,
                metadata=dict(info.get("metadata") or {}),
                size_bytes=size_bytes,
            )
            self._entries[normalized] = (display_category, entry)
            added.append(normalized)

        self._maybe_sort_after_add(added)
        self.files_added.emit(added, failed)
        self.entry_count_changed.emit(len(self._entries))
        self._activate_optimal_category_for_added(added)
        return added, failed

    @staticmethod
    def _display_category(category: str | None) -> str:
        """Map the canonical workflow category to the batch tab vocabulary."""
        workflow_category = str(category or "").strip().lower()
        if workflow_category == "markdown":
            return "text"
        if workflow_category in CATEGORY_ORDER:
            return str(workflow_category)
        return "other"

    def get_file_display_category(self, file_path: str) -> str | None:
        """Return the GUI tab category without mutating execution metadata."""
        record = self._entries.get(_normalize_path(file_path))
        return record[0] if record is not None else None

    def _maybe_sort_after_add(self, added: list[str]) -> None:
        """Apply current sort to affected categories after adding files."""
        if self._sort_key == "custom":
            for path in added:
                entry_info = self._entries.get(path)
                if entry_info is None:
                    continue
                category, _entry = entry_info
                order = self._custom_order_by_category.setdefault(category, [])
                if path not in order:
                    order.append(path)

    def _activate_optimal_category_for_added(self, added: list[str]) -> None:
        """Activate the most relevant category after a batch add.

        Prefer the category with the most newly
        added files, using the stable category order as the tie-breaker.
        """
        if not added:
            return
        counts = dict.fromkeys(CATEGORY_ORDER, 0)
        for path in added:
            record = self._entries.get(path)
            if record is None:
                continue
            category = record[0]
            if category in counts:
                counts[category] += 1
        max_count = max(counts.values(), default=0)
        if max_count <= 0:
            return
        for category in CATEGORY_ORDER:
            if counts[category] == max_count:
                self.activate_tab(category)
                return

    def remove_file(self, file_path: str) -> bool:
        """Remove a single file. Returns True if found and removed."""
        normalized = _normalize_path(file_path)
        record = self._entries.pop(normalized, None)
        if record is None:
            return False
        category = record[0]
        if self._sort_key == "custom":
            order = self._custom_order_by_category.get(category, [])
            if normalized in order:
                order.remove(normalized)
        self.files_removed.emit(normalized)
        self.entry_count_changed.emit(len(self._entries))
        return True

    def remove_files(self, file_paths: list[str]) -> list[str]:
        """Remove multiple files. Returns list of successfully removed paths."""
        removed: list[str] = []
        for path in file_paths:
            if self.remove_file(path):
                removed.append(path)
        return removed

    def clear_files(self) -> None:
        """Remove all files and reset state."""
        self._entries.clear()
        self._custom_order_by_category = {cat: [] for cat in CATEGORY_ORDER}
        self._sort_key = "custom"
        self._sort_ascending = True
        self._active_filter = "all"
        self.files_cleared.emit()
        self.entry_count_changed.emit(0)
        self.filter_changed.emit("all")

    def get_files(self) -> list[str]:
        """Return all file paths in category order, respecting current ordering."""
        files: list[str] = []
        for category in CATEGORY_ORDER:
            files.extend(self._get_ordered_paths_for_category(category))
        return files

    def get_files_for_category(self, category: str) -> list[str]:
        """Return file paths for a specific category."""
        if category not in CATEGORY_ORDER:
            return []
        ordered = self._get_ordered_paths_for_category(category)
        return ordered

    def get_file_count(self, category: str | None = None) -> int:
        """Return file count for a category or total if None."""
        if category is None:
            return len(self._entries)
        return sum(1 for _cat, _ in self._entries.values() if _cat == category)

    def get_visible_count_for_category(self, category: str) -> int:
        """Return count of entries matching the active filter for a category."""
        return sum(1 for _c, entry in self._entries.values() if _c == category and self._entry_matches_filter(entry))

    def get_file_entry(self, file_path: str) -> BatchFileEntry | None:
        """Get the BatchFileEntry for a file path."""
        normalized = _normalize_path(file_path)
        record = self._entries.get(normalized)
        if record is None:
            return None
        return record[1]

    def set_file_status(
        self,
        file_path: str,
        status: str,
        output_path: str | None = None,
        skip_reason: str | None = None,
        error_message: str | None = None,
        error_count: int | None = None,
        operation_id: str | None = None,
    ) -> bool:
        """Update the status of a file entry. Returns True if entry was found."""
        normalized_status = str(status).strip().lower()
        if normalized_status not in BATCH_FILE_STATUSES:
            raise ValueError(f"status must be one of {sorted(BATCH_FILE_STATUSES)!r}")
        normalized = _normalize_path(file_path)
        record = self._entries.get(normalized)
        if record is None:
            return False
        entry = record[1]
        previous_status = entry.status
        # Apply the complete payload before publishing the new status.  A
        # terminal status is the observable commit point for the row; callers
        # must never see it paired with stale artifact or diagnostic fields.
        if output_path is not None:
            entry.output_path = output_path
        if skip_reason is not None:
            entry.skip_reason = skip_reason
        if error_message is not None:
            entry.error_message = error_message
        if error_count is not None:
            entry.error_count = error_count
        if operation_id is not None:
            entry.operation_id = operation_id
        entry.status = normalized_status

        self.status_changed.emit(normalized, normalized_status)

        if (
            previous_status == "pending"
            and normalized_status == "processing"
            and should_pulse_processing_transition(len(self._entries))
        ):
            self.pulse_requested.emit(normalized)

        return True

    # ── Filter operations ──────────────────────────────────────────────

    def set_status_filter(self, filter_key: str) -> None:
        """Set the active status filter."""
        available = {opt[0] for opt in FILTER_OPTIONS}
        if filter_key not in available:
            filter_key = "all"
        if filter_key == self._active_filter:
            return
        self._active_filter = filter_key
        self.filter_changed.emit(filter_key)

    def get_status_filter(self) -> str:
        """Return the current filter key."""
        return self._active_filter

    def focus_failed_items(self) -> bool:
        """Set filter to 'failed' and return whether any failed items exist."""
        self.set_status_filter("failed")
        return self.get_failed_file_count() > 0

    def _entry_matches_filter(self, entry: BatchFileEntry) -> bool:
        """Check if an entry matches the active filter."""
        for fkey, _label, statuses in FILTER_OPTIONS:
            if fkey == self._active_filter:
                if not statuses:  # "all" — show everything
                    return True
                return entry.status in statuses
        return True  # fallback

    def _get_matching_filter_option(self) -> tuple[str, str, tuple[str, ...]]:
        """Get the filter option tuple for the active filter."""
        for opt in FILTER_OPTIONS:
            if opt[0] == self._active_filter:
                return opt
        return FILTER_OPTIONS[0]

    def get_failed_file_count(self) -> int:
        """Return count of failed files across all categories."""
        return sum(1 for _cat, entry in self._entries.values() if entry.status == "failed")

    def get_failed_files(self, category: str | None = None) -> list[str]:
        """Return file paths of all failed entries."""
        categories = [category] if category in CATEGORY_ORDER else CATEGORY_ORDER
        failed: list[str] = []
        for cat in categories:
            failed.extend(
                path for path, (ecat, entry) in self._entries.items() if ecat == cat and entry.status == "failed"
            )
        return failed

    def reset_failed_files(self, file_paths: list[str]) -> list[str]:
        """Reset failed files back to pending state."""
        reset_files: list[str] = []
        for path in file_paths:
            normalized = _normalize_path(path)
            record = self._entries.get(normalized)
            if record is None:
                continue
            _cat, entry = record
            if entry.status != "failed":
                continue
            self.set_file_status(normalized, "pending", output_path="", skip_reason="", error_message="")
            reset_files.append(normalized)
        return reset_files

    # ── Category / Tab operations ──────────────────────────────────────

    def activate_tab(self, category: str) -> bool:
        """Switch to the given category tab. Returns True if tab changed."""
        if category not in CATEGORY_ORDER:
            return False
        if category == self._current_category:
            return False
        self._current_category = category
        self.current_category_changed.emit(category)
        return True

    def get_current_category(self) -> str:
        """Return the current category."""
        return self._current_category

    # ── Sort operations ────────────────────────────────────────────────

    def set_sort_state(self, sort_key: str, ascending: bool = True) -> None:
        """Set sort key and direction."""
        if sort_key not in SORT_KEYS:
            sort_key = "custom"
        ascending = bool(ascending)
        if sort_key != "custom" and self._sort_key == "custom":
            for cat in CATEGORY_ORDER:
                self._custom_order_by_category[cat] = self._get_ordered_paths_for_category(cat)
        self._sort_key = sort_key
        self._sort_ascending = ascending
        self.sort_changed.emit(sort_key, ascending)

    def get_sort_state(self) -> tuple[str, bool]:
        """Return current (sort_key, ascending) tuple."""
        return self._sort_key, self._sort_ascending

    def _get_ordered_paths_for_category(self, category: str) -> list[str]:
        """Get file paths for a category, ordered by current sort state."""
        # Gather entries for this category
        entries_in_cat: list[tuple[str, BatchFileEntry]] = [
            (path, entry) for path, (cat, entry) in self._entries.items() if cat == category
        ]

        if self._sort_key == "custom":
            custom_order = self._custom_order_by_category.get(category, [])
            path_set: set[str] = set()
            ordered: list[str] = []
            for path in custom_order:
                if path in self._entries and self._entries[path][0] == category:
                    ordered.append(path)
                    path_set.add(path)
            for path, _entry in entries_in_cat:
                if path not in path_set:
                    ordered.append(path)
            return ordered

        # Sort by specified key
        entries_in_cat.sort(
            key=lambda pair: _sort_value(pair[1], self._sort_key),
            reverse=not self._sort_ascending,
        )
        return [path for path, _entry in entries_in_cat]

    def reorder_manual(self, category: str, ordered_paths: list[str]) -> None:
        """Record a manual reorder (from drag or Ctrl+Up/Down)."""
        self._custom_order_by_category[category] = list(ordered_paths)
        if self._sort_key != "custom":
            self._sort_key = "custom"
            self._sort_ascending = True
            self.sort_changed.emit("custom", True)
        self.reorder_requested.emit(category)

    # ── Selection operations ───────────────────────────────────────────

    def get_current_file(self, category: str | None = None) -> str | None:
        """Return a suggested current file — first in category if no explicit selection."""
        cat = category or self._current_category
        paths = self._get_ordered_paths_for_category(cat)
        # Filter by active filter
        filtered = [p for p in paths if self._entries.get(p) and self._entry_matches_filter(self._entries[p][1])]
        return filtered[0] if filtered else None

    def locate_file_entry(self, file_path: str) -> tuple[bool, str | None]:
        """Find which category a file belongs to and activate that tab.

        Returns (found, category).
        """
        normalized = _normalize_path(file_path)
        record = self._entries.get(normalized)
        if record is None:
            return False, None
        category = record[0]
        self.activate_tab(category)
        return True, category

    # ── Build retry targets ────────────────────────────────────────────

    def build_retry_targets(self, category: str, file_path: str | None) -> tuple[list[str], list[str]]:
        """Build retry targets: (selected_failed, category_failed)."""
        selected_failed: list[str] = []
        if file_path:
            normalized = _normalize_path(file_path)
            entry = self.get_file_entry(normalized)
            if entry is not None and entry.status == "failed":
                selected_failed.append(normalized)
        category_failed = self.get_failed_files(category)
        return selected_failed, category_failed

    # ── Aggregate file collection ────────────────────────────────────────

    # Map aggregate action → expected file categories
    _AGGREGATE_ACTION_CATEGORIES: dict[str, frozenset[str]] = {  # noqa: RUF012
        "merge_pdfs": frozenset({"layout"}),
        "merge_tables": frozenset({"spreadsheet"}),
        "merge_images_to_tiff": frozenset({"image"}),
    }
    _AGGREGATE_ACTION_FORMATS: dict[str, frozenset[str]] = {  # noqa: RUF012
        "merge_pdfs": frozenset({"pdf"}),
    }

    def get_aggregate_file_list(self, action_name: str) -> list[str]:
        """Return all file paths suitable for the given aggregate action.

        Args:
            action_name: One of ``"merge_pdfs"``, ``"merge_tables"``,
                ``"merge_images_to_tiff"``.

        Returns:
            List of matching file paths across all categories, ordered
            by the current sort.  Returns an empty list for unrecognized
            aggregate actions.
        """
        allowed = self._AGGREGATE_ACTION_CATEGORIES.get(action_name)
        if allowed is None:
            return []
        allowed_formats = self._AGGREGATE_ACTION_FORMATS.get(action_name)
        result: list[str] = []
        for category in CATEGORY_ORDER:
            if category not in allowed:
                continue
            for path in self._get_ordered_paths_for_category(category):
                entry_info = self._entries.get(path)
                if entry_info is None:
                    continue
                _cat, entry = entry_info
                if allowed_formats is not None and entry.detected_format.strip().lower() not in allowed_formats:
                    continue
                result.append(path)
        return result

    def has_aggregate_targets(self, action_name: str) -> bool:
        """Return True if at least two files match the aggregate action."""
        return len(self.get_aggregate_file_list(action_name)) >= 2
