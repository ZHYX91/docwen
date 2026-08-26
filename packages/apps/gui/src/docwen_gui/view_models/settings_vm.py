"""SettingsViewModel — observable state and user-action delegation for Settings.

This is the single source of truth for the Settings dialog's observable
state.  Widgets bind to its signals/properties; user actions flow through
method calls that delegate to ``ApplicationController``.

Design rules:
- ViewModel owns a ``SettingsConfig`` dataclass (typed config model).
- Every field change emits a ``config_changed(section, key, value)`` signal.
- Dirty tracking is automatic — ``is_dirty`` and ``dirty_sections`` are
  computed from config comparisons.
- Apply / Cancel / Reset go through the ViewModel, which delegates to
  ``ApplicationController``.
- ``apply_settings`` validates input before committing.

Thread-safety: all mutable state is protected by ``QMutex``.
"""

from __future__ import annotations

import logging
from collections.abc import Mapping
from copy import deepcopy
from dataclasses import fields, is_dataclass
from typing import TYPE_CHECKING, Any

from PySide6.QtCore import QMutex, QMutexLocker, QObject, Signal

from ..models.settings_config import (
    DEFAULT_HEADING_MERGE_PUNCTUATION,
    ConversionDefaultsConfig,
    ExportConfig,
    FormattingConfig,
    GUIConfig,
    LinkConfig,
    LoggingConfig,
    OutputConfig,
    ProofreadConfig,
    SettingsConfig,
    SoftwarePriorityConfig,
    TextConfig,
)
from ._optimization_filter import OptimizationChoicesResult, OptimizationSource, discover_optimization_choices

logger = logging.getLogger(__name__)

if TYPE_CHECKING:
    from docwen_application.controller import ApplicationController


# ── Settings section identifiers ─────────────────────────────────────
SECTION_GUI = "gui"
SECTION_TEXT = "text"
SECTION_PROOFREAD = "proofread"
SECTION_CONVERSION_DEFAULTS = "conversion_defaults"
SECTION_SOFTWARE_PRIORITY = "software_priority"
SECTION_LINK = "link"
SECTION_FORMATTING = "formatting"
SECTION_OUTPUT = "output"
SECTION_EXPORT = "export"
SECTION_LOGGING = "logging"

_SPECIAL_CONVERSION_ALLOWED_BACKENDS: dict[str, tuple[str, ...]] = {
    "odt_conversion": ("msoffice_word", "libreoffice"),
    "ods_conversion": ("msoffice_excel", "libreoffice"),
    "pdf_to_office": ("msoffice_word", "libreoffice"),
}

# GUI-model leaves owned by each Reset Tab action. Runtime remains the source
# of truth for physical files/dotted keys; this map only describes which draft
# values belong to each visible tab so a reset can preserve unrelated drafts.
_RESET_GROUP_DRAFT_PATHS: dict[str, tuple[tuple[str, ...], ...]] = {
    "general": (
        ("gui", "language"),
        ("gui", "theme"),
        ("gui", "transparency_enabled"),
        ("gui", "transparency_value"),
        ("gui", "remember_gui_state"),
        ("gui", "auto_center"),
        ("gui", "expand_side_panels"),
        ("gui", "default_mode"),
    ),
    "text": (
        ("text", "remove_numbering"),
        ("text", "add_numbering"),
        ("text", "default_scheme"),
        ("text", "numbering_schemes", "settings", "default_scheme"),
        ("text", "heading_numbering_render_mode"),
        ("gui", "md_default_template"),
    ),
    "proofread": (
        ("proofread", "symbol_pairing"),
        ("proofread", "symbol_correction"),
        ("proofread", "typos_rule"),
        ("proofread", "sensitive_word"),
        ("proofread", "skip_code_blocks"),
        ("proofread", "skip_quote_blocks"),
    ),
    "document": (
        ("conversion_defaults", "document", "to_md_keep_images"),
        ("conversion_defaults", "document", "to_md_enable_ocr"),
        ("conversion_defaults", "document", "to_md_table_merge_export_strategy"),
        ("conversion_defaults", "document", "to_md_remove_numbering"),
        ("conversion_defaults", "document", "to_md_add_numbering"),
        ("conversion_defaults", "document", "to_md_default_scheme"),
        ("conversion_defaults", "document", "to_md_enable_optimization"),
        ("conversion_defaults", "document", "to_md_optimization_type"),
        ("software_priority", "word_processors"),
        ("software_priority", "odt_conversion"),
        ("software_priority", "document_to_pdf"),
    ),
    "spreadsheet": (
        ("conversion_defaults", "spreadsheet", "to_md_keep_images"),
        ("conversion_defaults", "spreadsheet", "to_md_enable_ocr"),
        ("conversion_defaults", "spreadsheet", "to_md_table_merge_export_strategy"),
        ("conversion_defaults", "spreadsheet", "merge_mode"),
        ("software_priority", "spreadsheet_processors"),
        ("software_priority", "ods_conversion"),
        ("software_priority", "spreadsheet_to_pdf"),
    ),
    "image": (("conversion_defaults", "image"),),
    "layout": (
        ("conversion_defaults", "layout", "to_md_keep_images"),
        ("conversion_defaults", "layout", "to_md_enable_ocr"),
        ("conversion_defaults", "layout", "to_md_enable_optimization"),
        ("conversion_defaults", "layout", "to_md_optimization_type"),
        ("conversion_defaults", "layout", "render_dpi"),
        ("software_priority", "pdf_to_office"),
    ),
    "link": (("link",),),
    "formatting": (("formatting",),),
    "output": (("output",),),
    "export": (("conversion_defaults", "export"), ("export",)),
    "logging": (("logging",),),
    "other": (
        ("conversion_defaults", "other", "to_md_keep_images"),
        ("conversion_defaults", "other", "to_md_enable_ocr"),
    ),
    "conversion_defaults": (
        ("conversion_defaults",),
        ("text", "remove_numbering"),
        ("text", "add_numbering"),
        ("text", "default_scheme"),
        ("text", "heading_numbering_render_mode"),
        ("export", "image_mode"),
        ("export", "ocr_mode"),
        ("formatting", "table_style_mode"),
        ("formatting", "builtin_table_style"),
        ("formatting", "custom_table_style_name"),
    ),
    "software_priority": (("software_priority",),),
    "software": (("software_priority",),),
}

# Child editors commit immediately.  These paths are reconciled into both the
# live draft and its persisted baseline after the owning ConfigPort reloads so
# a later parent Apply/Cancel cannot replay stale data over that committed edit.
_EDITOR_FILE_MODEL_PATHS: dict[str, tuple[tuple[str, ...], ...]] = {
    "numbering/add.toml": (
        ("text", "numbering_schemes"),
        ("text", "default_scheme"),
    ),
    "numbering/cleanup.toml": (("text", "numbering_clean_rules"),),
    "proofread/pairs.toml": (),
    "proofread/symbol_map.toml": (("proofread", "symbol_mappings"),),
    "proofread/typos.toml": (("proofread", "typos_dict"),),
    "proofread/sensitive_words.toml": (("proofread", "sensitive_words"),),
}

_MISSING = object()


def _normalize_priority(values: object, allowed_backends: tuple[str, ...]) -> list[str]:
    """Keep GUI priority choices aligned with each route's backend contract."""
    normalized: list[str] = []
    if isinstance(values, (list, tuple)):
        for item in values:
            if isinstance(item, str) and item in allowed_backends and item not in normalized:
                normalized.append(item)
    for backend in allowed_backends:
        if backend not in normalized:
            normalized.append(backend)
    return normalized


def _normalize_settings_config(config: SettingsConfig) -> SettingsConfig:
    for key, allowed_backends in _SPECIAL_CONVERSION_ALLOWED_BACKENDS.items():
        setattr(
            config.software_priority,
            key,
            _normalize_priority(getattr(config.software_priority, key), allowed_backends),
        )
    return config


def _normalize_field_value(section: str, key: str, value: object) -> object:
    if section == SECTION_SOFTWARE_PRIORITY and key in _SPECIAL_CONVERSION_ALLOWED_BACKENDS:
        return _normalize_priority(value, _SPECIAL_CONVERSION_ALLOWED_BACKENDS[key])
    return value


def _plain_mapping(value: object) -> dict[str, Any]:
    """Normalize a dynamically assigned Settings draft value to a plain mapping."""
    if not isinstance(value, Mapping):
        return {}
    return {str(key): _plain_data(item) for key, item in value.items()}


def _changed_model_paths(before: Any, after: Any, path: tuple[str, ...] = ()) -> set[tuple[str, ...]]:
    """Return precise model leaves changed between two persisted snapshots."""
    if is_dataclass(before) and is_dataclass(after):
        changed: set[tuple[str, ...]] = set()
        for model_field in fields(before):
            if model_field.name.startswith("_"):
                continue
            changed.update(
                _changed_model_paths(
                    getattr(before, model_field.name),
                    getattr(after, model_field.name),
                    (*path, model_field.name),
                )
            )
        return changed

    if isinstance(before, Mapping) and isinstance(after, Mapping):
        changed = set()
        for key in set(before) | set(after):
            key_path = (*path, str(key))
            if key not in before or key not in after:
                changed.add(key_path)
                continue
            changed.update(_changed_model_paths(before[key], after[key], key_path))
        return changed

    return {path} if before != after else set()


def _model_paths_overlap(left: tuple[str, ...], right: tuple[str, ...]) -> bool:
    shared_length = min(len(left), len(right))
    return left[:shared_length] == right[:shared_length]


def _read_model_path(model: Any, path: tuple[str, ...]) -> Any:
    current = model
    for part in path:
        if is_dataclass(current):
            if not hasattr(current, part):
                return _MISSING
            current = getattr(current, part)
        elif isinstance(current, Mapping):
            if part not in current:
                return _MISSING
            current = current[part]
        else:
            return _MISSING
    return deepcopy(current)


def _write_model_path(model: Any, path: tuple[str, ...], value: Any) -> None:
    if not path:
        return
    current = model
    for part in path[:-1]:
        if is_dataclass(current):
            if not hasattr(current, part):
                return
            current = getattr(current, part)
        elif isinstance(current, dict):
            if part not in current:
                if value is _MISSING:
                    return
                current[part] = {}
            current = current[part]
        else:
            return

    leaf = path[-1]
    if is_dataclass(current):
        if value is not _MISSING and hasattr(current, leaf):
            setattr(current, leaf, deepcopy(value))
    elif isinstance(current, dict):
        if value is _MISSING:
            current.pop(leaf, None)
        else:
            current[leaf] = deepcopy(value)


def _mark_model_dirty_against(config: SettingsConfig, baseline: SettingsConfig) -> None:
    config.mark_clean()
    for model_field in fields(config):
        if model_field.name.startswith("_"):
            continue
        if getattr(config, model_field.name) != getattr(baseline, model_field.name):
            config.mark_dirty(model_field.name)


class SettingsViewModel(QObject):
    """Observable state for the Settings dialog.

    Signals:
        config_changed(section, key, value): emitted when any field changes.
        dirty_state_changed(is_dirty): emitted when dirty-state toggles.
        status_changed(message, is_error): emitted for Apply/Reset feedback.
        config_reloaded: emitted after a successful Reset or initial load.
    """

    # ── Signals ────────────────────────────────────────────────────────────

    config_changed = Signal(str, str, object)
    """Emitted when any config field changes: (section, key, value)."""

    dirty_state_changed = Signal(bool)
    """Emitted when the overall dirty-state flips."""

    status_changed = Signal(str, bool)
    """Emitted to show a status message: (message, is_error)."""

    config_reloaded = Signal()
    """Emitted after config is reloaded from source (after Reset)."""

    template_lists_changed = Signal(object)
    """Emitted when template lists are loaded/updated: dict of type -> [names]."""

    template_selection_changed = Signal(str, str)
    """Emitted when a template is selected: (template_type, template_name)."""

    optimization_types_changed = Signal(dict)
    """Emitted when localization-resolved optimization types are ready:
    {type_id: localized_display_name}."""

    # ── Construction ────────────────────────────────────────────────────────

    def __init__(
        self,
        config: SettingsConfig | None = None,
        controller: ApplicationController | None = None,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._mutex = QMutex()
        self._controller = controller
        self._config = _normalize_settings_config(config or SettingsConfig())
        # Snapshot used for dirty-state comparison and Cancel discard
        self._snapshot: SettingsConfig = deepcopy(self._config)
        self._is_dirty = False
        # ── Three-layer state (persisted / draft / preview) ──────────────
        # ''persisted'' — the locked baseline captured at begin_session()
        # (also updated after every apply_changes()).  Cancel restores here.
        self._persisted_baseline: SettingsConfig | None = None
        # ''preview theme'' track the live visual preview state so
        # cancel_changes() can revert it to the persisted baseline without
        # writing through _config (which holds the draft).
        self._preview_theme: str | None = None
        # Template lists and concrete template-name selections are session state.
        # Only the default template type is persisted as gui.md_default_template.
        self._templates: dict[str, list[str]] = {}
        self._selected_templates: dict[str, str] = {}
        self._initial_tab_key: str | None = None
        if controller is not None:
            self.load_from_controller_config()

    # ── Observable properties ───────────────────────────────────────────────

    @property
    def config(self) -> SettingsConfig:
        """The current (possibly unsaved) settings config."""
        with QMutexLocker(self._mutex):
            return deepcopy(self._config)

    @property
    def persisted_config(self) -> SettingsConfig:
        """The effective persisted baseline used by Cancel in this session."""
        with QMutexLocker(self._mutex):
            baseline = self._persisted_baseline if self._persisted_baseline is not None else self._snapshot
            return deepcopy(baseline)

    @property
    def is_dirty(self) -> bool:
        """Whether any field differs from the last-saved snapshot."""
        with QMutexLocker(self._mutex):
            return self._is_dirty

    def get_section(self, section: str) -> Any:
        """Return a dataclass section reference for a given section name.

        Returns a copy, not the live object, to prevent accidental mutation.
        """
        with QMutexLocker(self._mutex):
            return deepcopy(getattr(self._config, section, None))

    def get_field(self, section: str, key: str, default: object = None) -> object:
        """Return a single config field value.

        Args:
            section: Which config section (e.g. ``"text"``).
            key: Field name within the section dataclass.
            default: Value returned when section or key is not found.
        """
        with QMutexLocker(self._mutex):
            section_obj = getattr(self._config, section, None)
            if section_obj is None:
                return default
            return getattr(section_obj, key, default)

    # ── Command methods ─────────────────────────────────────────────────────

    def set_field(self, section: str, key: str, value: object) -> None:
        """Set a single config field and emit signals.

        Args:
            section: Which config section (e.g. ``"gui"``).
            key: Field name within the section dataclass.
            value: New value.
        """
        with QMutexLocker(self._mutex):
            section_obj = getattr(self._config, section, None)
            if section_obj is None:
                return
            if not hasattr(section_obj, key):
                return
            value = _normalize_field_value(section, key, value)
            old = getattr(section_obj, key)
            if old == value:
                return
            setattr(section_obj, key, value)
            self._config.mark_dirty(section)
            dirty = self._recompute_dirty()
        self.config_changed.emit(section, key, value)
        if dirty != self._is_dirty:
            self._is_dirty = dirty
            self.dirty_state_changed.emit(dirty)

    def set_conversion_default(self, category: str, key: str, value: object) -> None:
        """Set a field inside ``conversion_defaults.<category>``."""
        if value is None:
            return
        with QMutexLocker(self._mutex):
            defaults = self._config.conversion_defaults
            category_data = getattr(defaults, category, None)
            if not isinstance(category_data, dict):
                return
            old = category_data.get(key)
            if old == value:
                return
            category_data[key] = value
            self._config.mark_dirty(SECTION_CONVERSION_DEFAULTS)
            dirty = self._recompute_dirty()
        self.config_changed.emit(SECTION_CONVERSION_DEFAULTS, f"{category}.{key}", value)
        if dirty != self._is_dirty:
            self._is_dirty = dirty
            self.dirty_state_changed.emit(dirty)

    def get_available_field_processors(self) -> list[dict[str, Any]]:
        """Return Text-tab field processors visible for the current GUI locale."""
        with QMutexLocker(self._mutex):
            field_processors = deepcopy(self._config.text.field_processors)
            locale = self._config.gui.language
        return self._list_field_processors(field_processors, locale)

    def set_field_processor_enabled(self, processor_id: str, enabled: bool) -> bool:
        """Update a field processor enabled flag in the Settings draft."""
        with QMutexLocker(self._mutex):
            text = self._config.text
            field_processors = _plain_mapping(text.field_processors)
            processors = field_processors.get("processors", {})
            if not isinstance(processors, dict):
                return False
            processor = processors.get(processor_id)
            if not isinstance(processor, dict):
                return False
            old = bool(processor.get("enabled", True))
            if old == enabled:
                return True
            processor["enabled"] = bool(enabled)
            text.field_processors = field_processors
            self._config.mark_dirty(SECTION_TEXT)
            dirty = self._recompute_dirty()
        self.config_changed.emit(SECTION_TEXT, f"field_processors.{processor_id}.enabled", bool(enabled))
        if dirty != self._is_dirty:
            self._is_dirty = dirty
            self.dirty_state_changed.emit(dirty)
        return True

    def set_field_batch(self, section: str, updates: dict[str, object]) -> None:
        """Set multiple fields within one section atomically.

        Only one ``config_reloaded`` signal is emitted afterward.
        """
        with QMutexLocker(self._mutex):
            section_obj = getattr(self._config, section, None)
            if section_obj is None:
                return
            changed = False
            for key, value in updates.items():
                if not hasattr(section_obj, key):
                    continue
                value = _normalize_field_value(section, key, value)
                if getattr(section_obj, key) != value:
                    setattr(section_obj, key, value)
                    changed = True
            if not changed:
                return
            self._config.mark_dirty(section)
            dirty = self._recompute_dirty()
        self.config_reloaded.emit()
        if dirty != self._is_dirty:
            self._is_dirty = dirty
            self.dirty_state_changed.emit(dirty)

    def load_full_config(self, config: SettingsConfig) -> None:
        """Replace the entire config (e.g. on initial open or after Reset).

        Does NOT emit per-field signals — only ``config_reloaded``.
        """
        normalized = _normalize_settings_config(deepcopy(config))
        with QMutexLocker(self._mutex):
            self._config = normalized
            self._snapshot = deepcopy(normalized)
            self._is_dirty = False
        self.config_reloaded.emit()
        self.dirty_state_changed.emit(False)

    def apply_settings(self) -> bool:
        """Validate and commit all changes.

        Returns:
            True if all settings were accepted.
        """
        result = self.apply_changes()
        return result

    # ── Three-layer state machine ─────────────────────────────────────────

    def begin_session(self) -> None:
        """Capture the current live config as the persisted baseline.

        Must be called when the Settings dialog opens, BEFORE any user edits.
        The baseline is locked and used as the restore target for
        :meth:`cancel_changes`.
        """
        with QMutexLocker(self._mutex):
            self._persisted_baseline = deepcopy(self._config)
            self._snapshot = deepcopy(self._config)
            self._preview_theme = None
            self._is_dirty = False

    def apply_changes(self) -> bool:
        """Persist the current draft to disk and update the locked baseline.

        After this call the persisted baseline equals the draft, so any
        subsequent Cancel would still restore to this post-Apply state.

        Returns:
            True if all settings were accepted.
        """
        with QMutexLocker(self._mutex):
            errors = self._validate()
            if errors:
                self.status_changed.emit("\n".join(errors), True)
                return False
            draft = deepcopy(self._config)
            baseline = deepcopy(self._snapshot)

        try:
            persisted = self._persist_to_controller_config(draft, baseline)
        except Exception as exc:
            logger.exception("Settings persistence raised unexpectedly")
            persisted = False
            failure_detail = f" ({exc})"
        else:
            failure_detail = ""

        if not persisted:
            self._refresh_persisted_baseline_from_controller()
            self.status_changed.emit(f"Settings could not be fully applied{failure_detail}.", True)
            return False

        with QMutexLocker(self._mutex):
            self._config.mark_clean()
            self._snapshot = deepcopy(self._config)
            self._persisted_baseline = deepcopy(self._config)
            self._is_dirty = False
        self.status_changed.emit("Settings applied.", False)
        self.dirty_state_changed.emit(False)
        return True

    def cancel_changes(self) -> None:
        """Discard all draft edits and revert preview to the persisted baseline.

        Uses the locked baseline captured at :meth:`begin_session`
        (or last updated at :meth:`apply_changes`), NOT a stale snapshot
        from before the dialog opened.
        """
        with QMutexLocker(self._mutex):
            target = self._persisted_baseline if self._persisted_baseline is not None else self._snapshot
            self._config = deepcopy(target)
            self._config.mark_clean()
            self._preview_theme = None
            self._is_dirty = False
        self.config_reloaded.emit()
        self.dirty_state_changed.emit(False)

    def ok_changes(self) -> bool:
        """Apply changes and return True (for OK button — Apply + close)."""
        return self.apply_changes()

    def preview_theme(self, theme: str) -> None:
        """Apply *theme* as a visual preview without persisting or dirtying.

        The preview state is tracked separately from the draft config so that
        ``is_dirty`` stays False and :meth:`cancel_changes` can revert the
        visual state to the persisted baseline.

        Args:
            theme: Theme name (e.g. ``"dark"``, ``"light"``).
        """
        with QMutexLocker(self._mutex):
            self._preview_theme = str(theme)
        # NOTE: actual ThemeManager application happens outside the VM
        # (the dialog connects this to the ThemeManager).  The VM only
        # tracks the preview value so cancel can revert it.

    def preview_opacity(self, opacity: float) -> None:
        """Apply *opacity* as a visual preview without persisting.

        Same semantics as :meth:`preview_theme`: the value is tracked for
        Cancel revert but does not go through the draft/persist pipeline.

        Args:
            opacity: Opacity value (0.1 – 1.0).
        """
        with QMutexLocker(self._mutex):
            self._preview_theme = None  # opacity preview is independent
        # NOTE: actual window opacity is applied by the dialog, not the VM.

    def cancel(self) -> None:
        """Discard unsaved changes by restoring the last snapshot."""
        with QMutexLocker(self._mutex):
            self._config = deepcopy(self._snapshot)
            self._config.mark_clean()
            self._is_dirty = False
        self.config_reloaded.emit()
        self.dirty_state_changed.emit(False)

    def get_change_summary(self) -> list[dict[str, object]]:
        """Return field-level changes as ``{field, old, new}`` records.

        Compares the current draft against the persisted baseline.
        Only returns summary when ``begin_session()`` has been called
        (dialog is open), otherwise falls back to the internal snapshot.

        Returns:
            List of change records, each with ``field`` (dotted path),
            ``old`` (baseline value) and ``new`` (draft value).
        """
        with QMutexLocker(self._mutex):
            baseline = self._persisted_baseline if self._persisted_baseline is not None else self._snapshot
            live = self._config
            changes: list[dict[str, object]] = []
            _diff_configs(baseline, live, "", changes)
        return changes

    def set_templates(self, data: dict[str, list[str]]) -> None:
        """Inject template lists from an external data source.

        Called by the application layer (e.g. dialog opener) after
        querying the template registry.  Emits ``template_lists_changed``.

        Args:
            data: ``{template_type: [template_name, ...]}``.
        """
        with QMutexLocker(self._mutex):
            self._templates = deepcopy(data)
        self.template_lists_changed.emit(deepcopy(self._templates))

    def get_templates(self) -> dict[str, list[str]]:
        """Return the current template lists: ``{template_type: [name, ...]}``."""
        with QMutexLocker(self._mutex):
            return deepcopy(self._templates)

    def select_template(self, template_type: str, name: str) -> None:
        """Record a session template-name selection.

        The persistent setting is the template type preference
        (``gui.md_default_template``), which callers update through
        :meth:`set_field` when the user changes the active template type.
        Concrete template names are kept as view/session state so unavailable
        or refreshed template lists do not write stale names into GUI config.

        Args:
            template_type: ``"docx"`` or ``"xlsx"``.
            name: The selected template name.
        """
        with QMutexLocker(self._mutex):
            self._selected_templates[template_type] = name
        self.template_selection_changed.emit(template_type, name)

    @property
    def selected_templates(self) -> dict[str, str]:
        """Get current template selections: ``{template_type: template_name}``."""
        with QMutexLocker(self._mutex):
            return dict(self._selected_templates)

    # ── Initial tab activation ──────────────────────────────────────────────

    @property
    def initial_tab_key(self) -> str | None:
        """The settings tab key, if any, that should be activated on dialog open.

        Set by the caller (e.g. MainWindow) before constructing the dialog
        so the settings entry auto-focuses the tab most relevant to the
        user's current context.
        """
        with QMutexLocker(self._mutex):
            return self._initial_tab_key

    def set_initial_tab(self, tab_key: str) -> None:
        """Set the settings tab key to activate on dialog open.

        Args:
            tab_key: One of the ``TAB_KEYS`` values (e.g. ``"document"``).
        """
        with QMutexLocker(self._mutex):
            self._initial_tab_key = tab_key

    # ── Localized optimization types ────────────────────────────────────────

    def get_optimization_choices_result(
        self,
        locale: str | None = None,
        source_category: str | None = None,
    ) -> OptimizationChoicesResult:
        """Return the explicit Runtime-backed optimization discovery result."""

        from docwen_gui.i18n import get_locale

        return discover_optimization_choices(
            self._controller,
            locale=locale or get_locale(),
            sources=(OptimizationSource(detected_format="", source_category=source_category),)
            if source_category
            else (),
        )

    @staticmethod
    def _list_field_processors(field_processors: object, locale: str) -> list[dict[str, Any]]:
        field_processor_map = _plain_mapping(field_processors)
        processors = field_processor_map.get("processors", {})
        if not isinstance(processors, dict):
            return []
        settings = field_processor_map.get("settings", {})
        configured_order = settings.get("order", []) if isinstance(settings, dict) else []
        ordered_ids: list[str] = []
        for item in configured_order:
            if isinstance(item, str) and item in processors and item not in ordered_ids:
                ordered_ids.append(item)
        for processor_id in processors:
            if isinstance(processor_id, str) and processor_id not in ordered_ids:
                ordered_ids.append(processor_id)

        result: list[dict[str, Any]] = []
        for processor_id in ordered_ids:
            processor = processors.get(processor_id, {})
            if not isinstance(processor, dict):
                continue
            locales = processor.get("locales", ["*"])
            locale_list = locales if isinstance(locales, list) else ["*"]
            if "*" not in locale_list and locale not in locale_list:
                continue
            module = processor.get("module", "")
            item: dict[str, Any] = {
                "id": processor_id,
                "enabled": bool(processor.get("enabled", True)),
                "locales": list(locale_list),
                "module": module if isinstance(module, str) else "",
                "name": str(processor.get("name", "")),
                "name_key": str(processor.get("name_key", "")),
                "description": str(processor.get("description", "")),
                "is_system": bool(processor.get("is_system", False)),
            }
            if not item["module"]:
                item["load_error"] = "missing module"
            result.append(item)
        return result

    # ── Section → registry group mapping for reset operations ──────────────
    # Each GUI section maps to one runtime registry group.  Runtime decides
    # whether that logical owner uses complete files, precise values, or both.
    # Tab-level callers use reset_group() directly when they already know the
    # more precise group, such as "document" or "image".
    _SECTION_GROUP_MAP: dict[str, str] = {  # noqa: RUF012
        SECTION_GUI: "general",
        SECTION_TEXT: "text",
        SECTION_PROOFREAD: "proofread",
        SECTION_CONVERSION_DEFAULTS: "conversion_defaults",
        SECTION_SOFTWARE_PRIORITY: "software_priority",
        SECTION_LINK: "link",
        SECTION_FORMATTING: "formatting",
        SECTION_OUTPUT: "output",
        SECTION_EXPORT: "export",
        SECTION_LOGGING: "logging",
    }

    def load_from_controller_config(self) -> bool:
        controller = self._controller
        if controller is None:
            return False
        cfg_port = getattr(controller, "config_port", None)
        if cfg_port is None:
            return False
        try:
            raw = cfg_port.snapshot()
        except Exception:
            return False
        return self._replace_current_config_from_raw(raw)

    def _replace_current_config_from_raw(self, raw: dict[str, Any]) -> bool:
        """Replace draft state from one already-reconciled source snapshot."""
        mapped = self._map_raw_to_config(raw)
        with QMutexLocker(self._mutex):
            self._config = mapped
            self._snapshot = deepcopy(self._config)
            self._is_dirty = False
        self.config_reloaded.emit()
        self.dirty_state_changed.emit(False)
        return True

    @staticmethod
    def _try_snapshot_config_port(cfg_port: Any) -> dict[str, Any] | None:
        try:
            raw = cfg_port.snapshot()
        except Exception:
            logger.exception("Failed to snapshot config around reset")
            return None
        return deepcopy(raw) if isinstance(raw, dict) else None

    def _refresh_persisted_baseline_from_controller(self) -> bool:
        """Reconcile a possibly partial write while preserving the user's draft."""
        controller = self._controller
        cfg_port = getattr(controller, "config_port", None) if controller is not None else None
        if cfg_port is None:
            return False

        reload_ok = True
        reload_config = getattr(cfg_port, "reload", None)
        if callable(reload_config):
            try:
                reload_config()
            except Exception:
                logger.exception("Failed to reload config after a persistence attempt")
                reload_ok = False
        try:
            persisted = self._map_raw_to_config(cfg_port.snapshot())
        except Exception:
            logger.exception("Failed to snapshot config after a persistence attempt")
            return False

        with QMutexLocker(self._mutex):
            self._snapshot = deepcopy(persisted)
            if self._persisted_baseline is not None:
                self._persisted_baseline = deepcopy(persisted)
            self._is_dirty = self._recompute_dirty()
            dirty = self._is_dirty
        self.dirty_state_changed.emit(dirty)
        return reload_ok

    def _reconcile_reset_attempt(
        self,
        cfg_port: Any,
        before_source: dict[str, Any] | None,
        *,
        operation_succeeded: bool,
        group: str | None,
    ) -> bool:
        """Refresh a reset source without discarding drafts on a no-op failure."""
        reload_ok = True
        reload_config = getattr(cfg_port, "reload", None)
        if callable(reload_config):
            try:
                reload_config()
            except Exception:
                logger.exception("Failed to reload config after reset")
                reload_ok = False
        after_source = self._try_snapshot_config_port(cfg_port)
        if after_source is None:
            return False

        source_may_have_changed = before_source is None or after_source != before_source
        if operation_succeeded or source_may_have_changed:
            if group is None or (operation_succeeded and group not in _RESET_GROUP_DRAFT_PATHS):
                self._replace_current_config_from_raw(after_source)
                self._mark_current_config_as_persisted_baseline()
            else:
                self._merge_group_reset_from_raw(
                    group,
                    before_source,
                    after_source,
                    operation_succeeded=operation_succeeded,
                )
        return reload_ok

    def _merge_group_reset_from_raw(
        self,
        group: str,
        before_source: dict[str, Any] | None,
        after_source: dict[str, Any],
        *,
        operation_succeeded: bool,
    ) -> None:
        """Merge one immediate Reset Tab result without losing other drafts.

        The persisted baseline always follows the reconciled source. The draft
        takes actual before/after source changes plus, on success, every visible
        model leaf owned by the target group. Unrelated draft leaves survive.
        """
        after = self._map_raw_to_config(after_source)
        if before_source is not None:
            before = self._map_raw_to_config(before_source)
        else:
            with QMutexLocker(self._mutex):
                baseline = self._persisted_baseline if self._persisted_baseline is not None else self._snapshot
                before = deepcopy(baseline)

        source_paths = _changed_model_paths(before, after)
        reset_paths = set(source_paths)
        if operation_succeeded:
            reset_paths.update(_RESET_GROUP_DRAFT_PATHS.get(group, ()))

        with QMutexLocker(self._mutex):
            previous_draft = deepcopy(self._config)
            previous_snapshot = deepcopy(self._snapshot)
            draft_paths = _changed_model_paths(previous_snapshot, previous_draft)
            merged = deepcopy(after)
            for path in sorted(draft_paths, key=lambda item: (len(item), item)):
                if any(_model_paths_overlap(path, reset_path) for reset_path in reset_paths):
                    continue
                _write_model_path(merged, path, _read_model_path(previous_draft, path))
            merged = _normalize_settings_config(merged)
            _mark_model_dirty_against(merged, after)
            self._config = merged
            self._snapshot = deepcopy(after)
            if self._persisted_baseline is not None:
                self._persisted_baseline = deepcopy(after)
            self._is_dirty = self._recompute_dirty()
            dirty = self._is_dirty
        self.config_reloaded.emit()
        self.dirty_state_changed.emit(dirty)

    def _mark_current_config_as_persisted_baseline(self) -> None:
        """Treat the currently loaded config as persisted inside an open session."""
        with QMutexLocker(self._mutex):
            self._snapshot = deepcopy(self._config)
            if self._persisted_baseline is not None:
                self._persisted_baseline = deepcopy(self._config)
            self._config.mark_clean()
            self._is_dirty = False

    @staticmethod
    def _map_raw_to_config(raw: dict[str, Any]) -> SettingsConfig:
        """Convert the flat config dict into a SettingsConfig dataclass.

        Maps the registry-driven three-layer config keys (each file mounts
        under its declared namespace: ``text``, ``document``, ``proofread.*``,
        etc.) back to the SettingsConfig section names used by the GUI tabs.
        """
        config = SettingsConfig()

        # gui section
        gui = raw.get("gui", {})
        if isinstance(gui, dict):
            config.gui = GUIConfig(
                language=gui.get("language", {}).get("locale", "zh_CN")
                if isinstance(gui.get("language"), dict)
                else "zh_CN",
                theme=gui.get("theme", {}).get("default_theme", "light")
                if isinstance(gui.get("theme"), dict)
                else "light",
                transparency_enabled=gui.get("transparency", {}).get("enabled", False)
                if isinstance(gui.get("transparency"), dict)
                else False,
                transparency_value=gui.get("transparency", {}).get("default_value", 1.0)
                if isinstance(gui.get("transparency"), dict)
                else 1.0,
                remember_gui_state=gui.get("window", {}).get("remember_gui_state", True)
                if isinstance(gui.get("window"), dict)
                else True,
                auto_center=gui.get("window", {}).get("auto_center", False)
                if isinstance(gui.get("window"), dict)
                else False,
                expand_side_panels=gui.get("window", {}).get("expand_side_panels", False)
                if isinstance(gui.get("window"), dict)
                else False,
                default_mode=gui.get("window", {}).get("default_mode", "single")
                if isinstance(gui.get("window"), dict)
                else "single",
                md_default_template=gui.get("template", {}).get("md_default_template", "docx")
                if isinstance(gui.get("template"), dict)
                else "docx",
            )

        # text section — text.toml (text defaults) + numbering/add.toml (schemes)
        text_raw = raw.get("text", {}) if isinstance(raw.get("text"), dict) else {}
        numbering_raw = raw.get("numbering", {}) if isinstance(raw.get("numbering"), dict) else {}
        numbering_add = numbering_raw.get("add", {}) if isinstance(numbering_raw, dict) else {}
        numbering_cleanup = numbering_raw.get("cleanup", {}) if isinstance(numbering_raw, dict) else {}
        config.text = TextConfig(
            remove_numbering=text_raw.get("remove_numbering", True),
            add_numbering=text_raw.get("add_numbering", False),
            default_scheme=text_raw.get("numbering_scheme", "hierarchical_standard"),
            numbering_schemes=deepcopy(numbering_add) if isinstance(numbering_add, dict) else {},
            numbering_clean_rules=deepcopy(numbering_cleanup) if isinstance(numbering_cleanup, dict) else {},
            field_processors=_plain_data(raw.get("field_processors", {}))
            if isinstance(raw.get("field_processors"), Mapping)
            else {},
            heading_numbering_render_mode=text_raw.get("heading_numbering_render_mode", "text"),
        )

        # proofread section — proofread/{engine,skip,symbol_map,typos,sensitive_words}.toml
        proofread_raw = raw.get("proofread", {}) if isinstance(raw.get("proofread"), dict) else {}
        engine = proofread_raw.get("engine", {}) if isinstance(proofread_raw.get("engine"), dict) else {}
        skip = proofread_raw.get("skip", {}) if isinstance(proofread_raw.get("skip"), dict) else {}
        symbol_map = proofread_raw.get("symbol_map", {}) if isinstance(proofread_raw.get("symbol_map"), dict) else {}
        typos = proofread_raw.get("typos", {}) if isinstance(proofread_raw.get("typos"), dict) else {}
        sensitive = (
            proofread_raw.get("sensitive_words", {}) if isinstance(proofread_raw.get("sensitive_words"), dict) else {}
        )
        config.proofread = ProofreadConfig(
            symbol_pairing=engine.get("enable_symbol_pairing", True),
            symbol_correction=engine.get("enable_symbol_correction", True),
            typos_rule=engine.get("enable_typos_rule", True),
            sensitive_word=engine.get("enable_sensitive_word", True),
            skip_code_blocks=skip.get("code_blocks", True),
            skip_quote_blocks=skip.get("quote_blocks", False),
            symbol_mappings=symbol_map.get("entries", {}) if isinstance(symbol_map.get("entries"), dict) else {},
            typos_dict=typos.get("entries", {}) if isinstance(typos.get("entries"), dict) else {},
            sensitive_words=sensitive.get("entries", {}) if isinstance(sensitive.get("entries"), dict) else {},
        )

        # conversion_defaults — document/spreadsheet/image/layout/other/export.toml
        config.conversion_defaults = ConversionDefaultsConfig(
            document=_plain_data(raw.get("document", {})) if isinstance(raw.get("document"), Mapping) else {},
            spreadsheet=_plain_data(raw.get("spreadsheet", {})) if isinstance(raw.get("spreadsheet"), Mapping) else {},
            image=_plain_data(raw.get("image", {})) if isinstance(raw.get("image"), Mapping) else {},
            layout=_plain_data(raw.get("layout", {})) if isinstance(raw.get("layout"), Mapping) else {},
            other=_plain_data(raw.get("other", {})) if isinstance(raw.get("other"), Mapping) else {},
            export=_plain_data(raw.get("export", {})) if isinstance(raw.get("export"), Mapping) else {},
        )

        # software_priority
        soft = raw.get("software", {})
        if isinstance(soft, dict):
            dp = soft.get("default_priority", {})
            sc = soft.get("special_conversions", {})
            config.software_priority = SoftwarePriorityConfig(
                word_processors=dp.get("word_processors", ["wps_writer", "msoffice_word", "libreoffice"])
                if isinstance(dp, dict)
                else [],
                odt_conversion=sc.get("odt", ["msoffice_word", "libreoffice"]) if isinstance(sc, dict) else [],
                document_to_pdf=sc.get("document_to_pdf", ["wps_writer", "msoffice_word", "libreoffice"])
                if isinstance(sc, dict)
                else [],
                spreadsheet_processors=dp.get(
                    "spreadsheet_processors", ["wps_spreadsheets", "msoffice_excel", "libreoffice"]
                )
                if isinstance(dp, dict)
                else [],
                ods_conversion=sc.get("ods", ["msoffice_excel", "libreoffice"]) if isinstance(sc, dict) else [],
                spreadsheet_to_pdf=sc.get("spreadsheet_to_pdf", ["wps_spreadsheets", "msoffice_excel", "libreoffice"])
                if isinstance(sc, dict)
                else [],
                pdf_to_office=sc.get("pdf_to_office", ["msoffice_word", "libreoffice"]) if isinstance(sc, dict) else [],
            )

        # link
        link = raw.get("link", {})
        if isinstance(link, dict):
            fmt = link.get("format", {})
            non_embed = link.get("non_embed_links", {})
            embed = link.get("embed_links", {})
            emb_cfg = link.get("embedding", {})
            config.link = LinkConfig(
                image_link_style=fmt.get("image_link_style", "wiki_embed") if isinstance(fmt, dict) else "wiki_embed",
                md_file_link_style=fmt.get("md_file_link_style", "wiki_embed")
                if isinstance(fmt, dict)
                else "wiki_embed",
                wiki_link_mode=non_embed.get("wiki_mode", "hyperlink") if isinstance(non_embed, dict) else "hyperlink",
                markdown_link_mode=non_embed.get("markdown_mode", "hyperlink")
                if isinstance(non_embed, dict)
                else "hyperlink",
                wiki_embed_image_mode=embed.get("wiki_image_mode", "embed") if isinstance(embed, dict) else "embed",
                markdown_embed_image_mode=embed.get("markdown_image_mode", "embed")
                if isinstance(embed, dict)
                else "embed",
                embed_md_file_mode=embed.get("md_file_mode", "embed") if isinstance(embed, dict) else "embed",
                max_depth=emb_cfg.get("max_depth", 3) if isinstance(emb_cfg, dict) else 3,
            )

        # formatting
        conv = raw.get("conversion", {})
        if isinstance(conv, dict):
            d2m = conv.get("docx_to_md", {})
            m2d = conv.get("md_to_docx", {})
            syntax = conv.get("syntax", {})
            hr = conv.get("horizontal_rule", {})
            hr_d2m = hr.get("docx_to_md", {}) if isinstance(hr, dict) else {}
            hr_m2d = hr.get("md_to_docx", {}) if isinstance(hr, dict) else {}
            document = raw.get("document", {})
            style = document.get("style", {}) if isinstance(document, dict) else {}
            table_style = style.get("table", {}) if isinstance(style, dict) else {}
            table_m2d = table_style.get("md_to_docx", {}) if isinstance(table_style, dict) else {}
            config.formatting = FormattingConfig(
                body_format="preserve"
                if (isinstance(d2m, dict) and d2m.get("preserve_formatting", True))
                else "discard",
                heading_format="preserve"
                if (isinstance(d2m, dict) and d2m.get("preserve_heading_formatting", False))
                else "discard",
                table_header_format="preserve"
                if (isinstance(d2m, dict) and d2m.get("preserve_table_header_formatting", False))
                else "discard",
                page_break_sep=hr_d2m.get("page_break", "---") if isinstance(hr_d2m, dict) else "---",
                section_break_sep=hr_d2m.get("section_break", "***") if isinstance(hr_d2m, dict) else "***",
                horizontal_rule_sep=hr_d2m.get("horizontal_rule", "___") if isinstance(hr_d2m, dict) else "___",
                bold_syntax=syntax.get("bold", "asterisk") if isinstance(syntax, dict) else "asterisk",
                italic_syntax=syntax.get("italic", "asterisk") if isinstance(syntax, dict) else "asterisk",
                strikethrough_syntax=syntax.get("strikethrough", "extended")
                if isinstance(syntax, dict)
                else "extended",
                highlight_syntax=syntax.get("highlight", "extended") if isinstance(syntax, dict) else "extended",
                superscript_syntax=syntax.get("superscript", "html") if isinstance(syntax, dict) else "html",
                subscript_syntax=syntax.get("subscript", "html") if isinstance(syntax, dict) else "html",
                unordered_list_syntax=syntax.get("unordered_list", "dash") if isinstance(syntax, dict) else "dash",
                indent_spaces=syntax.get("indent_spaces", 4) if isinstance(syntax, dict) else 4,
                md_body_format=m2d.get("formatting_mode", "apply") if isinstance(m2d, dict) else "apply",
                md_heading_format=m2d.get("heading_formatting_mode", "remove") if isinstance(m2d, dict) else "remove",
                md_table_header_format=m2d.get("table_header_formatting_mode", "remove")
                if isinstance(m2d, dict)
                else "remove",
                heading_merge_mode=m2d.get("heading_merge_mode", "punct_required")
                if isinstance(m2d, dict)
                else "punct_required",
                heading_merge_punctuation=(
                    str(m2d.get("heading_merge_punctuation", DEFAULT_HEADING_MERGE_PUNCTUATION))
                    if isinstance(m2d, dict)
                    and m2d.get("heading_merge_punctuation", DEFAULT_HEADING_MERGE_PUNCTUATION) is not None
                    else DEFAULT_HEADING_MERGE_PUNCTUATION
                ),
                list_separator=(
                    str(m2d.get("list_separator", "、"))
                    if isinstance(m2d, dict) and m2d.get("list_separator", "、") is not None
                    else "、"
                ),
                table_style_mode=table_m2d.get("table_style_mode", "builtin")
                if isinstance(table_m2d, dict)
                else "builtin",
                builtin_table_style=table_m2d.get("builtin_style_key", "three_line_table")
                if isinstance(table_m2d, dict)
                else "three_line_table",
                custom_table_style_name=table_m2d.get("custom_style_name", "") if isinstance(table_m2d, dict) else "",
                dash_sep=hr_m2d.get("dash", "page_break") if isinstance(hr_m2d, dict) else "page_break",
                asterisk_sep=hr_m2d.get("asterisk", "section_break") if isinstance(hr_m2d, dict) else "section_break",
                underscore_sep=hr_m2d.get("underscore", "horizontal_rule_1")
                if isinstance(hr_m2d, dict)
                else "horizontal_rule_1",
            )

        # output
        out = raw.get("output", {})
        if isinstance(out, dict):
            odir = out.get("directory", {})
            obeh = out.get("behavior", {})
            oint = out.get("intermediate_files", {})
            config.output = OutputConfig(
                output_mode=odir.get("mode", "source") if isinstance(odir, dict) else "source",
                custom_path=odir.get("custom_path", "") if isinstance(odir, dict) else "",
                create_date_subfolder=odir.get("create_date_subfolder", False) if isinstance(odir, dict) else False,
                date_folder_format=odir.get("date_folder_format", "%Y-%m-%d") if isinstance(odir, dict) else "%Y-%m-%d",
                auto_open_folder=obeh.get("auto_open_folder", False) if isinstance(obeh, dict) else False,
                save_intermediate_files=oint.get("save_to_output", False) if isinstance(oint, dict) else False,
            )

        # export
        export_defaults = raw.get("export", {}) if isinstance(raw.get("export"), dict) else {}
        conv_export = conv.get("export", {}) if isinstance(conv, dict) else {}
        ocr_output = conv.get("ocr_output", {}) if isinstance(conv, dict) else {}
        title_overrides = (
            ocr_output.get("blockquote_title_override_by_locale", {}) if isinstance(ocr_output, dict) else {}
        )
        try:
            from docwen_gui.i18n import get_locale

            current_locale = get_locale()
        except Exception:
            current_locale = config.gui.language
        title_text = ""
        if isinstance(title_overrides, dict):
            title_text = str(title_overrides.get(current_locale) or title_overrides.get(config.gui.language) or "")
        config.export = ExportConfig(
            image_mode=export_defaults.get("to_md_image_extraction_mode", "file"),
            ocr_mode=export_defaults.get("to_md_ocr_placement_mode", "image_md"),
            ocr_title_enabled=ocr_output.get("show_blockquote_title", True) if isinstance(ocr_output, dict) else True,
            ocr_title_text=title_text,
            base64_compress_enabled=conv_export.get("base64_compress_enabled", True)
            if isinstance(conv_export, dict)
            else True,
            base64_compress_threshold_kb=conv_export.get("base64_compress_threshold_kb", 100)
            if isinstance(conv_export, dict)
            else 100,
        )

        # logging
        log = raw.get("logger", {})
        if isinstance(log, dict):
            config.logging = LoggingConfig(
                enable=log.get("enable", True),
                level=log.get("level", "debug"),
                file_prefix=log.get("file_prefix", "docwen"),
                retention_days=log.get("retention_days", 30),
                console_enable=log.get("console_enable", True),
                console_level=log.get("console_level", "info"),
                console_format=log.get("console_format", ""),
                console_colorize=log.get("console_colorize", "auto"),
                directory_mode=log.get("directory_mode", "user"),
                directory=log.get("directory", ""),
            )

        return _normalize_settings_config(config)

    def reset_section(self, section: str) -> bool:
        """Reset a GUI section through its runtime-owned logical group plan."""
        group = self._SECTION_GROUP_MAP.get(section)
        if group is None:
            self.status_changed.emit(f"Unknown section: {section}", True)
            return False
        return self.reset_group(group)

    def reset_group(self, group: str) -> bool:
        """Reset one logical config group through the injected config port.

        Runtime owns the declarative file/key plan.  Delegating through the
        controller keeps a custom composition root and the GUI on the same
        user-config directory instead of consulting a process-global loader.
        """

        controller = self._controller
        cfg_port = getattr(controller, "config_port", None) if controller is not None else None
        reset_group = getattr(cfg_port, "reset_group", None)
        if not callable(reset_group):
            self.status_changed.emit(f"Unknown config group: {group}", True)
            return False
        before_source = self._try_snapshot_config_port(cfg_port)
        try:
            ok = bool(reset_group(group))
        except Exception:
            logger.exception("Config group reset raised: %s", group)
            ok = False

        reconciled = self._reconcile_reset_attempt(
            cfg_port,
            before_source,
            operation_succeeded=ok,
            group=group,
        )
        ok = ok and reconciled
        if ok:
            self.status_changed.emit(f"Config group '{group}' reset to defaults.", False)
        else:
            self.status_changed.emit(f"Config group '{group}' reset partially failed.", True)
        return ok

    def reset_all(self) -> bool:
        """Reset all non-protected settings through the injected config port."""
        controller = self._controller
        cfg_port = getattr(controller, "config_port", None) if controller is not None else None
        if cfg_port is None:
            ok = False
            reconciled = False
        else:
            before_source = self._try_snapshot_config_port(cfg_port)
            try:
                ok = bool(cfg_port.reset_all())
            except Exception:
                logger.exception("Reset all settings raised")
                ok = False
            reconciled = self._reconcile_reset_attempt(
                cfg_port,
                before_source,
                operation_succeeded=ok,
                group=None,
            )
        ok = ok and reconciled
        if ok:
            self.status_changed.emit("All settings reset to defaults.", False)
        else:
            self.status_changed.emit("Some settings could not be reset (excluded files).", True)
        return ok

    def _persist_to_controller_config(
        self,
        config: SettingsConfig,
        baseline: SettingsConfig | None = None,
    ) -> bool:
        controller = self._controller
        cfg_port = getattr(controller, "config_port", None) if controller is not None else None
        if cfg_port is None:
            return True

        values: dict[str, Any] = {}

        def put(key: str, value: Any) -> None:
            values[key] = value

        gui = config.gui
        put("gui.language.locale", gui.language)
        put("gui.theme.default_theme", gui.theme)
        put("gui.transparency.enabled", bool(gui.transparency_enabled))
        put("gui.transparency.default_value", float(gui.transparency_value))
        put("gui.window.remember_gui_state", bool(gui.remember_gui_state))
        put("gui.window.auto_center", bool(gui.auto_center))
        put("gui.window.expand_side_panels", bool(gui.expand_side_panels))
        put("gui.window.default_mode", gui.default_mode)
        put("gui.template.md_default_template", gui.md_default_template)

        link = config.link
        put("link.format.image_link_style", link.image_link_style)
        put("link.format.md_file_link_style", link.md_file_link_style)
        put("link.non_embed_links.wiki_mode", link.wiki_link_mode)
        put("link.non_embed_links.markdown_mode", link.markdown_link_mode)
        put("link.embed_links.wiki_image_mode", link.wiki_embed_image_mode)
        put("link.embed_links.markdown_image_mode", link.markdown_embed_image_mode)
        put("link.embed_links.md_file_mode", link.embed_md_file_mode)
        put("link.embedding.max_depth", int(link.max_depth))

        text = config.text
        numbering_schemes = _plain_mapping(text.numbering_schemes)
        settings = numbering_schemes.setdefault("settings", {})
        if isinstance(settings, dict):
            settings["default_scheme"] = text.default_scheme

        self._collect_markdown_numbering_values(
            values,
            numbering_schemes,
            {
                "remove_numbering": bool(text.remove_numbering),
                "add_numbering": bool(text.add_numbering),
                "numbering_scheme": text.default_scheme,
                "heading_numbering_render_mode": text.heading_numbering_render_mode,
            },
        )
        self._collect_numbering_clean_values(values, _plain_mapping(text.numbering_clean_rules))
        self._collect_field_processor_enabled_flags(values, text.field_processors)
        out = config.output
        put("output.directory.mode", out.output_mode)
        put("output.directory.custom_path", out.custom_path)
        put("output.directory.create_date_subfolder", bool(out.create_date_subfolder))
        put("output.directory.date_folder_format", out.date_folder_format)
        put("output.behavior.auto_open_folder", bool(out.auto_open_folder))
        put("output.intermediate_files.save_to_output", bool(out.save_intermediate_files))

        log = config.logging
        put("logger.enable", bool(log.enable))
        put("logger.level", log.level)
        put("logger.file_prefix", log.file_prefix)
        put("logger.retention_days", int(log.retention_days))
        put("logger.console_enable", bool(log.console_enable))
        put("logger.console_level", log.console_level)
        put("logger.console_format", log.console_format)
        put("logger.console_colorize", log.console_colorize)
        put("logger.directory_mode", log.directory_mode)
        put("logger.directory", log.directory)

        sp = config.software_priority
        put("software.default_priority.word_processors", list(sp.word_processors))
        put("software.default_priority.spreadsheet_processors", list(sp.spreadsheet_processors))
        put("software.special_conversions.odt", list(sp.odt_conversion))
        put("software.special_conversions.ods", list(sp.ods_conversion))
        put("software.special_conversions.document_to_pdf", list(sp.document_to_pdf))
        put("software.special_conversions.spreadsheet_to_pdf", list(sp.spreadsheet_to_pdf))
        put("software.special_conversions.pdf_to_office", list(sp.pdf_to_office))

        self._collect_conversion_defaults(values, config.conversion_defaults)

        fmt = config.formatting
        put("conversion.docx_to_md.preserve_formatting", fmt.body_format == "preserve")
        put("conversion.docx_to_md.preserve_heading_formatting", fmt.heading_format == "preserve")
        put("conversion.docx_to_md.preserve_table_header_formatting", fmt.table_header_format == "preserve")
        put("conversion.syntax.bold", fmt.bold_syntax)
        put("conversion.syntax.italic", fmt.italic_syntax)
        put("conversion.syntax.strikethrough", fmt.strikethrough_syntax)
        put("conversion.syntax.highlight", fmt.highlight_syntax)
        put("conversion.syntax.superscript", fmt.superscript_syntax)
        put("conversion.syntax.subscript", fmt.subscript_syntax)
        put("conversion.syntax.unordered_list", fmt.unordered_list_syntax)
        put("conversion.syntax.indent_spaces", int(fmt.indent_spaces))
        put("conversion.md_to_docx.formatting_mode", fmt.md_body_format)
        put("conversion.md_to_docx.heading_formatting_mode", fmt.md_heading_format)
        put("conversion.md_to_docx.table_header_formatting_mode", fmt.md_table_header_format)
        put("conversion.md_to_docx.heading_merge_mode", fmt.heading_merge_mode)
        put("conversion.md_to_docx.heading_merge_punctuation", fmt.heading_merge_punctuation)
        put("conversion.md_to_docx.list_separator", fmt.list_separator)
        put("document.style.table.md_to_docx.table_style_mode", fmt.table_style_mode)
        put("document.style.table.md_to_docx.builtin_style_key", fmt.builtin_table_style)
        put("document.style.table.md_to_docx.custom_style_name", fmt.custom_table_style_name)
        put("conversion.horizontal_rule.docx_to_md.page_break", fmt.page_break_sep)
        put("conversion.horizontal_rule.docx_to_md.section_break", fmt.section_break_sep)
        put("conversion.horizontal_rule.docx_to_md.horizontal_rule", fmt.horizontal_rule_sep)
        put("conversion.horizontal_rule.md_to_docx.dash", fmt.dash_sep)
        put("conversion.horizontal_rule.md_to_docx.asterisk", fmt.asterisk_sep)
        put("conversion.horizontal_rule.md_to_docx.underscore", fmt.underscore_sep)

        proof = config.proofread
        put("proofread.engine.enable_symbol_pairing", proof.symbol_pairing)
        put("proofread.engine.enable_symbol_correction", proof.symbol_correction)
        put("proofread.engine.enable_typos_rule", proof.typos_rule)
        put("proofread.engine.enable_sensitive_word", proof.sensitive_word)
        put("proofread.skip.code_blocks", proof.skip_code_blocks)
        put("proofread.skip.quote_blocks", proof.skip_quote_blocks)
        baseline_proof = baseline.proofread if baseline is not None else None
        if baseline_proof is None or proof.symbol_mappings != baseline_proof.symbol_mappings:
            put("proofread.symbol_map.entries", proof.symbol_mappings)
        if baseline_proof is None or proof.typos_dict != baseline_proof.typos_dict:
            put("proofread.typos.entries", proof.typos_dict)
        if baseline_proof is None or proof.sensitive_words != baseline_proof.sensitive_words:
            put("proofread.sensitive_words.entries", proof.sensitive_words)

        exp = config.export
        put("export.to_md_image_extraction_mode", exp.image_mode)
        put("export.to_md_ocr_placement_mode", exp.ocr_mode)
        put("conversion.export.base64_compress_enabled", bool(exp.base64_compress_enabled))
        put("conversion.export.base64_compress_threshold_kb", int(exp.base64_compress_threshold_kb))
        put("conversion.ocr_output.show_blockquote_title", bool(exp.ocr_title_enabled))
        try:
            from docwen_gui.i18n import get_locale

            title_locale = get_locale()
        except Exception:
            title_locale = config.gui.language
        put(f"conversion.ocr_output.blockquote_title_override_by_locale.{title_locale}", exp.ocr_title_text)

        return self._persist_config_values(cfg_port, values)

    @staticmethod
    def _collect_conversion_defaults(values: dict[str, Any], defaults: ConversionDefaultsConfig) -> None:
        """Collect GUI-owned conversion default keys for persistence."""
        gui_owned_keys: dict[str, tuple[str, ...]] = {
            "document": (
                "to_md_keep_images",
                "to_md_enable_ocr",
                "to_md_remove_numbering",
                "to_md_add_numbering",
                "to_md_default_scheme",
                "to_md_enable_optimization",
                "to_md_optimization_type",
                "to_md_table_merge_export_strategy",
            ),
            "spreadsheet": (
                "to_md_keep_images",
                "to_md_enable_ocr",
                "to_md_table_merge_export_strategy",
                "merge_mode",
            ),
            "image": (
                "to_md_keep_images",
                "to_md_enable_ocr",
                "to_md_enable_optimization",
                "to_md_optimization_type",
                "ocr_language",
                "compress_mode",
                "size_limit",
                "size_unit",
                "pdf_quality",
                "tiff_mode",
            ),
            "layout": (
                "to_md_keep_images",
                "to_md_enable_ocr",
                "to_md_enable_optimization",
                "to_md_optimization_type",
                "render_dpi",
            ),
            "other": (
                "to_md_keep_images",
                "to_md_enable_ocr",
            ),
        }
        for category, keys in gui_owned_keys.items():
            category_data = getattr(defaults, category, None)
            if not isinstance(category_data, dict):
                continue
            for key in keys:
                if key in category_data:
                    values[f"{category}.{key}"] = category_data[key]

    @staticmethod
    def _collect_field_processor_enabled_flags(values: dict[str, Any], field_processors: object) -> None:
        processors = _plain_mapping(field_processors).get("processors", {})
        if not isinstance(processors, dict):
            return
        for processor_id, processor in processors.items():
            if not isinstance(processor_id, str) or not isinstance(processor, dict):
                continue
            if "enabled" in processor:
                values[f"field_processors.processors.{processor_id}.enabled"] = bool(processor["enabled"])

    @staticmethod
    def _persist_config_values(cfg_port: Any, values: dict[str, Any]) -> bool:
        if not values:
            return True
        set_many = getattr(cfg_port, "set_many", None)
        if not callable(set_many):
            logger.error("Settings persistence requires transactional ConfigPort.set_many")
            return False
        return bool(set_many(values))

    @staticmethod
    def _collect_markdown_numbering_values(
        values: dict[str, Any],
        numbering_schemes: dict[str, Any],
        defaults_text_updates: dict[str, Any] | None = None,
    ) -> None:
        """Collect complete editor-owned numbering sections for one port write."""
        for section in ("settings", "number_styles", "schemes"):
            if section in numbering_schemes:
                values[f"numbering.add.{section}"] = deepcopy(numbering_schemes[section])
        for key, value in (defaults_text_updates or {}).items():
            values[f"text.{key}"] = deepcopy(value)

    @staticmethod
    def _collect_numbering_clean_values(values: dict[str, Any], data: dict[str, Any]) -> None:
        if "settings" in data:
            values["numbering.cleanup.settings"] = deepcopy(data["settings"])
        if "rules" in data:
            values["numbering.cleanup.rules"] = deepcopy(data["rules"])

    def _config_port(self) -> Any | None:
        controller = self._controller
        return getattr(controller, "config_port", None) if controller is not None else None

    def _reconcile_editor_file_source(self, rel_path: str, raw: dict[str, Any] | None = None) -> bool:
        """Adopt one immediate child-editor commit without losing other drafts."""
        paths = _EDITOR_FILE_MODEL_PATHS.get(rel_path)
        if paths is None:
            return False
        cfg_port = self._config_port()
        if cfg_port is None:
            return False
        if raw is None:
            try:
                raw = cfg_port.snapshot()
            except Exception:
                logger.exception("Failed to snapshot config after editor save: %s", rel_path)
                return False
        if not isinstance(raw, dict):
            return False
        if not paths:
            return True

        persisted = self._map_raw_to_config(raw)
        with QMutexLocker(self._mutex):
            for path in paths:
                value = _read_model_path(persisted, path)
                _write_model_path(self._config, path, value)
                _write_model_path(self._snapshot, path, value)
                if self._persisted_baseline is not None:
                    _write_model_path(self._persisted_baseline, path, value)
            _mark_model_dirty_against(self._config, self._snapshot)
            self._is_dirty = self._recompute_dirty()
            dirty = self._is_dirty
        self.config_reloaded.emit()
        self.dirty_state_changed.emit(dirty)
        return True

    def _persist_editor_values(self, rel_path: str, values: dict[str, Any]) -> bool:
        cfg_port = self._config_port()
        if cfg_port is None:
            return False
        before = self._try_snapshot_config_port(cfg_port)
        try:
            ok = self._persist_config_values(cfg_port, values)
        except Exception:
            logger.exception("Editor config persistence raised: %s", rel_path)
            ok = False
        after = self._try_snapshot_config_port(cfg_port)
        source_changed = after is not None and (before is None or after != before)
        reconciled = True
        if ok or source_changed:
            reconciled = after is not None and self._reconcile_editor_file_source(rel_path, after)
        return bool(ok and reconciled)

    def persist_numbering_schemes_source(self, data: dict[str, Any]) -> bool:
        """Commit numbering-add schemes through this VM's injected config port."""
        values: dict[str, Any] = {}
        default_scheme = data.get("settings", {}).get("default_scheme", "")
        self._collect_markdown_numbering_values(
            values,
            _plain_mapping(data),
            {"numbering_scheme": default_scheme} if isinstance(default_scheme, str) and default_scheme else None,
        )
        return self._persist_editor_values("numbering/add.toml", values)

    def persist_numbering_clean_rules_source(self, data: dict[str, Any]) -> bool:
        """Commit numbering-clean rules through this VM's injected config port."""
        values: dict[str, Any] = {}
        self._collect_numbering_clean_values(values, _plain_mapping(data))
        return self._persist_editor_values("numbering/cleanup.toml", values)

    def read_config_file_text(self, config_name: str) -> str | None:
        """Read a VM-owned editor source through the injected config port."""
        if config_name not in _EDITOR_FILE_MODEL_PATHS:
            logger.error("Settings editor does not own config source: %s", config_name)
            return None
        cfg_port = self._config_port()
        get_file_text = getattr(cfg_port, "get_file_text", None)
        if not callable(get_file_text):
            return None
        try:
            content = get_file_text(config_name)
        except Exception:
            logger.exception("Config editor source read raised: %s", config_name)
            return None
        return content if isinstance(content, str) else None

    def save_config_file_text(self, config_name: str, content: str) -> bool:
        """Save VM-owned TOML through the port and reconcile editor state."""
        if config_name not in _EDITOR_FILE_MODEL_PATHS:
            logger.error("Settings editor does not own config source: %s", config_name)
            return False
        cfg_port = self._config_port()
        save_file_text = getattr(cfg_port, "save_file_text", None)
        if not callable(save_file_text):
            return False
        before = self._try_snapshot_config_port(cfg_port)
        try:
            ok = bool(save_file_text(config_name, content))
        except Exception:
            logger.exception("Config editor source save raised: %s", config_name)
            ok = False
        after = self._try_snapshot_config_port(cfg_port)
        source_changed = after is not None and (before is None or after != before)
        reconciled = True
        if ok or source_changed:
            reconciled = after is not None and self._reconcile_editor_file_source(config_name, after)
        return bool(ok and reconciled)

    def make_save_config_text_callback(self) -> Any:
        """Build a ``save_callback`` for :class:`TomlEditorWidget`.

        Returns a ``(config_name, content) -> bool`` callable that delegates
        to the injected config port. Use this when constructing a
        ``TomlEditorWidget`` so raw TOML saves go through the composition root
        (registry validation + in-memory wiring) instead of bypassing
        the configured source with a direct ``path.write_text``.

        Example::

            TomlEditorWidget(
                parent,
                config_name="numbering/cleanup.toml",
                path_resolver=self._resolve_config_path,
                reload_callback=self._on_config_saved,
                save_callback=self.make_save_config_text_callback(),
            )
        """

        return self.save_config_file_text

    # ── Internal helpers ────────────────────────────────────────────────────

    def _recompute_dirty(self) -> bool:
        """Compare live config against snapshot (must hold _mutex)."""
        # Use deepcopy to avoid false-sharing of mutable sub-objects
        live = deepcopy(self._config)
        snap = self._snapshot
        if live.gui != snap.gui:
            return True
        if live.text != snap.text:
            return True
        if live.proofread != snap.proofread:
            return True
        if live.conversion_defaults != snap.conversion_defaults:
            return True
        if live.software_priority != snap.software_priority:
            return True
        if live.link != snap.link:
            return True
        if live.formatting != snap.formatting:
            return True
        if live.output != snap.output:
            return True
        if live.export != snap.export:
            return True
        return live.logging != snap.logging

    @staticmethod
    def _validate() -> list[str]:
        """Return a list of human-readable validation error messages."""
        # Validation is section-specific; base VM has no mandatory checks.
        # LoggingTab validation is handled by the tab's own validate method.
        return []


# ── Diff helper ──────────────────────────────────────────────────────────


def _diff_configs(
    baseline: SettingsConfig | Any,
    live: SettingsConfig | Any,
    path: str,
    out: list[dict[str, object]],
) -> None:
    """Walk two SettingsConfig trees and record field-level differences.

    Uses ``__dataclass_fields__`` introspection to enumerate leaf fields
    without hardcoding section names.
    """
    from dataclasses import fields, is_dataclass

    if not is_dataclass(baseline) or not is_dataclass(live):
        return
    for f in fields(baseline):
        name = f.name
        if name.startswith("_"):
            continue
        old_val = getattr(baseline, name)
        new_val = getattr(live, name)
        field_path = f"{path}.{name}" if path else name
        if is_dataclass(old_val) and is_dataclass(new_val):
            _diff_configs(old_val, new_val, field_path, out)
        elif old_val != new_val:
            out.append({"field": field_path, "old": old_val, "new": new_val})


def _plain_data(value: Any) -> Any:
    """Convert tomlkit-backed containers into plain Python data."""
    if isinstance(value, Mapping):
        return _plain_mapping(value)
    if isinstance(value, list):
        return [_plain_data(item) for item in value]
    return deepcopy(value)
