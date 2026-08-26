"""Strict, fail-closed locale reference audit used by repository tests.

Only shipped Python source and explicitly registered declarative producers are
trusted.  Tests, docs, generated trees, arbitrary strings, and generic ``.t``
methods can never make a catalogue key look live.
"""

from __future__ import annotations

import ast
import tomllib
import warnings
from collections import Counter
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any

STRUCTURAL_SECTIONS = frozenset({"meta", "styles", "style_formats", "placeholders", "yaml_keys"})
_BANNED_SOURCE_PARTS = frozenset({".tmp", ".pytest_cache", "build", "dist", "tmp", "__pycache__"})


@dataclass(frozen=True)
class DynamicCallContract:
    """Exact finite key set for one non-literal translator expression."""

    expected_count: int
    keys: frozenset[str]
    rationale: str


@dataclass(frozen=True)
class LiteralFallbackContract:
    """Exact reviewed fallback for one literal key absent from locale TOML."""

    expected_count: int
    default: str
    rationale: str


@dataclass(frozen=True)
class LocaleReferenceAudit:
    defined: frozenset[str]
    used: frozenset[str]
    unused: frozenset[str]
    unresolved: tuple[str, ...]
    undefined_literal_keys: tuple[str, ...]
    contract_mismatches: tuple[str, ...]
    undefined_contract_keys: frozenset[str]
    source_files: tuple[str, ...]
    source_roots: tuple[str, ...]


def _keys(prefix: str, suffixes: set[str] | frozenset[str] | tuple[str, ...]) -> frozenset[str]:
    return frozenset(f"{prefix}{suffix}" for suffix in suffixes)


_FILE_ADMISSION_KEYS = _keys(
    "file_admission.",
    {
        "compatible_text",
        "same_family_mismatch",
        "cross_family_mismatch",
        "unknown_extension",
        "empty",
        "container_invalid",
        "container_unsupported",
        "container_unrecognized",
        "content_unrecognized",
        "read_error",
        "unsupported_format",
    },
)
_ABOUT_TOOL_KEYS = _keys(
    "about.tools.",
    {
        "python_docx",
        "openpyxl",
        "pymupdf",
        "pymupdf4llm",
        "pdf2docx",
        "easyofd",
        "rapidocr",
        "paddleocr",
        "onnxruntime",
        "pyside6",
        "pillow",
        "pillow_heif",
        "img2pdf",
        "pywin32",
        "lxml",
        "latex2mathml",
        "pyyaml",
        "tomlkit",
        "pandas",
        "numpy",
        "olefile",
        "emoji",
    },
)
_FONT_SIZE_KEYS = _keys("components.font_size.", {"small", "default", "large", "xlarge"})
_TASK_NOTIFICATION_KEYS = _keys(
    "components.info_area.task_completion_notification_", {"success", "partial", "failed", "cancelled"}
)
_TASK_STATE_KEYS = _keys("info_area.task_state_", {"active", "success", "partial", "failed", "cancelled"})
_FILE_TYPE_KEYS = _keys("file_types.", {"text", "document", "spreadsheet", "layout", "image", "other"})
_BATCH_STATUS_KEYS = _keys(
    "components.file_drop.status.", {"pending", "processing", "completed", "failed", "skipped", "cancelled"}
)
_BATCH_FILTER_KEYS = _keys(
    "components.file_drop.batch_list.filter_",
    {"pending", "processing", "completed", "failed", "skipped", "cancelled"},
)
_INPUT_AREA_KEYS = frozenset(
    {
        "components.file_drop.batch_mode",
        "components.file_drop.single_mode",
        "components.file_drop.add_button",
        "components.file_drop.clear_button",
        "components.file_drop.add_file_action",
        "components.file_drop.add_folder_action",
        "components.file_drop.recent_files_action",
        "components.file_drop.clear_recent_files_action",
        "components.file_drop.select_file_dialog",
        "components.file_drop.select_folder_dialog",
        "components.file_drop.empty_hint_single",
        "components.file_drop.empty_hint_batch",
        "info_area.transient_title",
        "file_category.text_short",
        "file_category.layout_short",
        "file_category.spreadsheet_short",
        "file_category.document_short",
        "file_category.image_short",
        "file_category.other_short",
    }
)
_NUMBERING_ADD_NAME_KEYS = _keys("editors.numbering_add.names.", {"hierarchical_standard", "hierarchical_h2_start"})
_NUMBERING_ADD_DESCRIPTION_KEYS = _keys(
    "editors.numbering_add.descriptions.", {"hierarchical_standard_desc", "hierarchical_h2_start_desc"}
)
_NUMBERING_LEVEL_KEYS = _keys("editors.numbering_add.level_", tuple(str(value) for value in range(1, 10)))
_NUMBERING_CLEAN_NAMES = {
    "chinese_unit_suffix",
    "number_separator",
    "bracket_number",
    "circled_numbers",
    "hierarchical",
    "chinese_unit_prefix",
    "legal_english",
    "letter_number",
}
_NUMBERING_CLEAN_KEYS = _keys("editors.numbering_clean.names.", _NUMBERING_CLEAN_NAMES) | _keys(
    "editors.numbering_clean.descriptions.", {f"{name}_desc" for name in _NUMBERING_CLEAN_NAMES}
)
_FORMATTING_KEYS = frozenset(
    f"settings.formatting.syntax.{kind}_{form}"
    for kind in ("strikethrough", "highlight", "superscript", "subscript")
    for form in ("extended", "html")
)
_TEMPLATE_TAB_KEYS = _keys("components.template_selector_tabbed.", {"document_templates", "spreadsheet_templates"})


def _contract(expected_count: int, keys: frozenset[str], rationale: str) -> DynamicCallContract:
    return DynamicCallContract(expected_count=expected_count, keys=keys, rationale=rationale)


def _literal_fallback(expected_count: int, default: str, rationale: str) -> LiteralFallbackContract:
    return LiteralFallbackContract(
        expected_count=expected_count,
        default=default,
        rationale=rationale,
    )


# Keyed by the shipped source path and Python's canonical AST expression.
# This is deliberately exhaustive: a new non-literal translator call fails
# until its finite producer and exact suffix set are reviewed here.
DYNAMIC_CALL_CONTRACTS: Mapping[tuple[str, str], DynamicCallContract] = {
    (
        "packages/apps/cli/src/docwen_cli/commands/execution_request.py",
        "key",
    ): _contract(1, frozenset(), "transparent wrapper; only structural yaml_keys.title/subtitle call sites"),
    (
        "packages/apps/cli/src/docwen_cli/file_admission_i18n.py",
        "key",
    ): _contract(
        2,
        _FILE_ADMISSION_KEYS | {"main_window.file_admission_confirm_action"},
        "Core diagnostic producer plus the local acceptance-action constant",
    ),
    (
        "packages/apps/gui/src/docwen_gui/file_admission_i18n.py",
        "key",
    ): _contract(1, _FILE_ADMISSION_KEYS, "Core FILE_ADMISSION_MESSAGE_KEYS producer"),
    (
        "packages/apps/gui/src/docwen_gui/dialogs/about.py",
        "f'about.tools.{tooltip_key}'",
    ): _contract(2, _ABOUT_TOOL_KEYS, "finite _TOOLS_LEFT/_TOOLS_RIGHT suffixes"),
    (
        "packages/apps/gui/src/docwen_gui/main_window.py",
        "f'components.font_size.{preset}'",
    ): _contract(1, _FONT_SIZE_KEYS, "FONT_SIZE_PRESETS loop"),
    (
        "packages/apps/gui/src/docwen_gui/main_window.py",
        "f'components.font_size.{normalized}'",
    ): _contract(1, _FONT_SIZE_KEYS, "validated font-size normalizer"),
    (
        "packages/apps/gui/src/docwen_gui/main_window.py",
        "state_key",
    ): _contract(1, _TASK_NOTIFICATION_KEYS, "finite terminal notification map"),
    (
        "packages/apps/gui/src/docwen_gui/numbering_schemes.py",
        "f'editors.numbering_add.names.{name_key}'",
    ): _contract(1, _NUMBERING_ADD_NAME_KEYS, "configs/numbering/add.toml name_key fields"),
    (
        "packages/apps/gui/src/docwen_gui/view_models/_optimization_filter.py",
        "f'cli.interactive.optimization_types.{resource.id}'",
    ): _contract(
        1,
        _keys("cli.interactive.optimization_types.", {"gongwen", "invoice_cn"}),
        "literal OptimizationResourceSpec ids in the two bundled optimizer manifests",
    ),
    (
        "packages/apps/gui/src/docwen_gui/view_models/action_area_vm.py",
        "key",
    ): _contract(
        1,
        frozenset(
            {
                "action_area.document.export_markdown_tooltip",
                "action_area.spreadsheet.export_markdown_tooltip",
                "action_area.image.export_markdown_tooltip",
                "action_area.layout.export_markdown_tooltip",
                "action_area.md_to_document.generate_tooltip",
                "action_area.md_to_spreadsheet.generate_tooltip",
            }
        ),
        "finite action-tooltip map",
    ),
    (
        "packages/apps/gui/src/docwen_gui/view_models/info_area_vm.py",
        "i18n_key",
    ): _contract(
        1,
        _keys(
            "info_area.task_guide_",
            {"open_output_dir", "view_failed_details", "retry_failed", "add_more_files"},
        )
        | {"common.ok"},
        "finite _TASK_GUIDE_LABELS map and default",
    ),
    (
        "packages/apps/gui/src/docwen_gui/view_models/info_area_vm.py",
        "f'info_area.task_state_{state}'",
    ): _contract(1, _TASK_STATE_KEYS, "TaskSummaryState validates TASK_SUMMARY_STATES"),
    (
        "packages/apps/gui/src/docwen_gui/view_models/input_area_vm.py",
        "f'file_types.{category}'",
    ): _contract(1, _FILE_TYPE_KEYS, "FILE_CATEGORY_ORDER"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/action_area.py",
        "label_key_map.get(ft, 'action_area.layout.optimize_for_type')",
    ): _contract(
        1,
        _keys(
            "action_area.",
            {f"{category}.optimize_for_type" for category in ("document", "spreadsheet", "image", "layout")},
        ),
        "finite optimization label map",
    ),
    (
        "packages/apps/gui/src/docwen_gui/widgets/batch_list.py",
        "f'components.file_drop.status.{filter_key}'",
    ): _contract(1, _BATCH_STATUS_KEYS, "FILTER_OPTIONS finite status domain"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/batch_list.py",
        "f'components.file_drop.batch_list.filter_{filter_key}'",
    ): _contract(1, _BATCH_FILTER_KEYS, "FILTER_OPTIONS finite filter domain"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/batch_list.py",
        "f'components.file_drop.status.{entry.status}'",
    ): _contract(2, _BATCH_STATUS_KEYS, "BatchFileEntry validates BATCH_FILE_STATUSES"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/batch_list.py",
        "f'file_types.{category}'",
    ): _contract(1, _FILE_TYPE_KEYS, "FILE_CATEGORY_ORDER"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/input_area.py",
        "key",
    ): _contract(1, _INPUT_AREA_KEYS, "transparent _i18n wrapper with finite widget constants"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/document_tab.py",
        "_SOFTWARE_LABEL_KEYS.get(sid, '')",
    ): _contract(
        1,
        _keys(
            "settings.document.software.",
            {"wps_writer", "msoffice_word", "libreoffice"},
        ),
        "finite document software label map",
    ),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/formatting_tab.py",
        "f'settings.formatting.syntax.{kind}_extended'",
    ): _contract(1, _FORMATTING_KEYS, "four literal formatting-kind call sites"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/formatting_tab.py",
        "f'settings.formatting.syntax.{kind}_html'",
    ): _contract(1, _FORMATTING_KEYS, "four literal formatting-kind call sites"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/layout_tab.py",
        "_LAYOUT_SOFTWARE_LABEL_KEYS.get(sid, '')",
    ): _contract(
        1,
        _keys("settings.document.software.", {"msoffice_word", "libreoffice"}),
        "finite layout software label map",
    ),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/numbering_add_editor.py",
        "f'editors.numbering_add.names.{self.name_key}'",
    ): _contract(1, _NUMBERING_ADD_NAME_KEYS, "allowlisted add.toml name_key fields"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/numbering_add_editor.py",
        "f'editors.numbering_add.descriptions.{self.description_key}'",
    ): _contract(1, _NUMBERING_ADD_DESCRIPTION_KEYS, "allowlisted add.toml description_key fields"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/numbering_add_editor.py",
        "f'editors.numbering_add.level_{idx}'",
    ): _contract(2, _NUMBERING_LEVEL_KEYS, "literal range(1, 10) level editor"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/numbering_clean_editor.py",
        "key",
    ): _contract(2, _NUMBERING_CLEAN_KEYS, "allowlisted cleanup.toml name/description fields"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/spreadsheet_tab.py",
        "_SS_SOFTWARE_LABEL_KEYS.get(sid, '')",
    ): _contract(
        1,
        _keys("settings.spreadsheet.software.", {"wps_spreadsheets", "excel", "libreoffice"}),
        "finite spreadsheet software label map",
    ),
    (
        "packages/apps/gui/src/docwen_gui/widgets/template_selector_tabbed.py",
        "'components.template_selector_tabbed.document_templates' if template_type == 'docx' else "
        "'components.template_selector_tabbed.spreadsheet_templates'",
    ): _contract(1, _TEMPLATE_TAB_KEYS, "literal two-way template type projection"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/template_selector_tabbed.py",
        "label_key",
    ): _contract(1, _TEMPLATE_TAB_KEYS, "literal two-row template tab list"),
}


# These call sites intentionally supply bounded English fallback copy while the
# catalogue migration remains incomplete. The path, key, call count, and exact
# fallback are all frozen: deleting any other live literal locale key cannot be
# hidden by merely passing a default string.
LITERAL_FALLBACK_CONTRACTS: Mapping[tuple[str, str], LiteralFallbackContract] = {
    (
        "packages/apps/gui/src/docwen_gui/dialogs/feedback.py",
        "common.copy",
    ): _literal_fallback(2, "Copy", "feedback detail copy action"),
    (
        "packages/apps/gui/src/docwen_gui/main_window.py",
        "main_window.conversion_warning",
    ): _literal_fallback(1, "Conversion completed with a warning", "diagnostic fallback"),
    (
        "packages/apps/gui/src/docwen_gui/main_window.py",
        "main_window.add_file",
    ): _literal_fallback(1, "Add File", "window shortcut label"),
    (
        "packages/apps/gui/src/docwen_gui/main_window.py",
        "main_window.start_conversion",
    ): _literal_fallback(1, "Start Conversion", "window shortcut label"),
    (
        "packages/apps/gui/src/docwen_gui/main_window.py",
        "main_window.cancel",
    ): _literal_fallback(1, "Cancel", "window shortcut label"),
    (
        "packages/apps/gui/src/docwen_gui/main_window.py",
        "main_window.aggregate_need_two",
    ): _literal_fallback(1, "At least two matching files are required.", "aggregate admission"),
    (
        "packages/apps/gui/src/docwen_gui/main_window.py",
        "main_window.thread_start_uncertain",
    ): _literal_fallback(
        1,
        "The task worker started but startup reporting failed; cancellation was requested.",
        "worker startup recovery",
    ),
    (
        "packages/apps/gui/src/docwen_gui/main_window.py",
        "main_window.close_waiting_for_tasks",
    ): _literal_fallback(1, "Cancelling active tasks before closing...", "close lifecycle"),
    (
        "packages/apps/gui/src/docwen_gui/main_window.py",
        "main_window.close_cancel_failed",
    ): _literal_fallback(
        1,
        "Could not request cancellation for {task_id}: {message}",
        "close lifecycle failure",
    ),
    (
        "packages/apps/gui/src/docwen_gui/main_window.py",
        "main_window.close_wait_timeout",
    ): _literal_fallback(
        1,
        "A task is still stopping. DocWen will remain open until cleanup finishes.",
        "close lifecycle timeout",
    ),
    (
        "packages/apps/gui/src/docwen_gui/view_models/main_window_vm.py",
        "components.file_drop.file_unavailable",
    ): _literal_fallback(1, "File is unavailable", "file admission recovery"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/dialog.py",
        "settings.unsaved_changes.message",
    ): _literal_fallback(1, "You have unsaved changes. Close without saving?", "settings close guard"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/logging_tab.py",
        "settings.logging.browse_title",
    ): _literal_fallback(1, "Select Log Directory", "log directory chooser"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/numbering_clean_editor.py",
        "editors.numbering_add.level",
    ): _literal_fallback(1, "Level:", "numbering cleanup editor field"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/numbering_clean_editor.py",
        "common.type_warning",
    ): _literal_fallback(1, "Warning", "typed warning dialog"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/proofread_tab.py",
        "editors.mapping.save_symbol_mapping_failed",
    ): _literal_fallback(1, "Could not save symbol pairing entries.", "mapping persistence error"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/proofread_tab.py",
        "editors.mapping.load_failed",
    ): _literal_fallback(1, "Could not load the editable configuration source.", "mapping load error"),
    (
        "packages/apps/gui/src/docwen_gui/widgets/settings/toml_editor.py",
        "settings.toml_editor.save_failed_message",
    ): _literal_fallback(2, "The configuration could not be saved.", "TOML persistence error"),
}


def _flatten(prefix: str, value: Any, out: set[str]) -> None:
    if isinstance(value, dict):
        for key, nested in value.items():
            child = f"{prefix}.{key}" if prefix else str(key)
            _flatten(child, nested, out)
        return
    if value is not None:
        out.add(prefix)


def locale_leaf_keys(path: Path) -> set[str]:
    data = tomllib.loads(path.read_text(encoding="utf-8"))
    out: set[str] = set()
    _flatten("", data, out)
    return out


def defined_locale_keys(reference: Path) -> set[str]:
    data = tomllib.loads(reference.read_text(encoding="utf-8"))
    out: set[str] = set()
    for top_key, top_value in data.items():
        if str(top_key) not in STRUCTURAL_SECTIONS:
            _flatten(str(top_key), top_value, out)
    return out


def _product_source_files(project_root: Path) -> tuple[tuple[str, ...], tuple[Path, ...]]:
    package_root = project_root / "packages"
    roots: list[Path] = []
    files: set[Path] = set()
    for candidate in package_root.rglob("src"):
        if not candidate.is_dir():
            continue
        relative_parts = set(candidate.relative_to(package_root).parts)
        if relative_parts & _BANNED_SOURCE_PARTS:
            continue
        roots.append(candidate)
        for path in candidate.rglob("*.py"):
            if not (set(path.relative_to(package_root).parts) & _BANNED_SOURCE_PARTS):
                files.add(path)
    relative_roots = tuple(sorted(path.relative_to(project_root).as_posix() for path in roots))
    return relative_roots, tuple(sorted(files))


def _translator_aliases(tree: ast.AST, relative_path: str) -> set[str]:
    aliases: set[str] = set()
    is_gui_source = "/apps/gui/src/" in f"/{relative_path}"
    for node in ast.walk(tree):
        if not isinstance(node, ast.ImportFrom):
            continue
        module = node.module or ""
        for imported in node.names:
            canonical = (module, imported.name)
            is_cli = canonical == ("docwen_cli.i18n", "cli_t")
            is_gui = canonical == ("docwen_gui.i18n", "t")
            is_relative_gui = is_gui_source and node.level > 0 and canonical == ("i18n", "t")
            if is_cli or is_gui or is_relative_gui:
                aliases.add(imported.asname or imported.name)
    return aliases


def _parse_python(path: Path) -> ast.AST:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", SyntaxWarning)
        return ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))


def semantic_numbering_keys(project_root: Path) -> set[str]:
    """Read only the two reviewed numbering translation-key fields."""

    used: set[str] = set()
    add_path = project_root / "configs" / "numbering" / "add.toml"
    if add_path.exists():
        add_doc = tomllib.loads(add_path.read_text(encoding="utf-8"))
        schemes = add_doc.get("schemes", {})
        if isinstance(schemes, Mapping):
            for scheme in schemes.values():
                if not isinstance(scheme, Mapping):
                    continue
                name_key = scheme.get("name_key")
                description_key = scheme.get("description_key")
                if isinstance(name_key, str) and name_key:
                    used.add(f"editors.numbering_add.names.{name_key}")
                if isinstance(description_key, str) and description_key:
                    used.add(f"editors.numbering_add.descriptions.{description_key}")

    cleanup_path = project_root / "configs" / "numbering" / "cleanup.toml"
    if cleanup_path.exists():
        cleanup_doc = tomllib.loads(cleanup_path.read_text(encoding="utf-8"))
        rules = cleanup_doc.get("rules", [])
        if isinstance(rules, list):
            for rule in rules:
                if not isinstance(rule, Mapping):
                    continue
                name_key = rule.get("name_key")
                description_key = rule.get("description_key")
                if isinstance(name_key, str) and name_key:
                    used.add(f"editors.numbering_clean.names.{name_key}")
                if isinstance(description_key, str) and description_key:
                    used.add(f"editors.numbering_clean.descriptions.{description_key}")
    return used


def audit_locale_references(
    project_root: Path,
    *,
    locale_dir: Path,
    contracts: Mapping[tuple[str, str], DynamicCallContract] | None = None,
    literal_contracts: Mapping[tuple[str, str], LiteralFallbackContract] | None = None,
) -> LocaleReferenceAudit:
    defined = defined_locale_keys(locale_dir / "zh_CN.toml")
    available = locale_leaf_keys(locale_dir / "zh_CN.toml")
    active_contracts = DYNAMIC_CALL_CONTRACTS if contracts is None else contracts
    active_literal_contracts = LITERAL_FALLBACK_CONTRACTS if literal_contracts is None else literal_contracts
    source_roots, paths = _product_source_files(project_root)
    used: set[str] = semantic_numbering_keys(project_root) & defined
    unresolved: list[str] = []
    undefined_literal_keys: list[str] = []
    seen_dynamic: Counter[tuple[str, str]] = Counter()
    seen_literal_fallback: Counter[tuple[str, str]] = Counter()

    for path in paths:
        relative_path = path.relative_to(project_root).as_posix()
        tree = _parse_python(path)
        aliases = _translator_aliases(tree, relative_path)
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            if not isinstance(node.func, ast.Name) or node.func.id not in aliases:
                continue
            key_keywords = [keyword.value for keyword in node.keywords if keyword.arg == "key"]
            default_keywords = [keyword.value for keyword in node.keywords if keyword.arg == "default"]
            binding_errors: list[str] = []
            if len(node.args) > 2:
                binding_errors.append("more than two positional arguments")
            if any(isinstance(argument, ast.Starred) for argument in node.args):
                binding_errors.append("starred positional arguments")
            if len(key_keywords) > 1 or (node.args and key_keywords):
                binding_errors.append("multiple values for key")
            if len(default_keywords) > 1 or (len(node.args) > 1 and default_keywords):
                binding_errors.append("multiple values for default")
            if binding_errors:
                unresolved.append(
                    f"{relative_path}:{node.lineno}: invalid {node.func.id} argument binding "
                    f"({', '.join(binding_errors)}): {ast.unparse(node)}"
                )
                continue
            key_arguments = ([node.args[0]] if node.args else []) + key_keywords
            if len(key_arguments) != 1:
                unresolved.append(
                    f"{relative_path}:{node.lineno}: {node.func.id} call must provide "
                    f"exactly one explicit key argument: {ast.unparse(node)}"
                )
                continue
            argument = key_arguments[0]
            if isinstance(argument, ast.Constant) and isinstance(argument.value, str):
                if argument.value in defined:
                    used.add(argument.value)
                    continue
                if argument.value in available:
                    continue
                signature = (relative_path, argument.value)
                contract = active_literal_contracts.get(signature)
                default_arguments = ([node.args[1]] if len(node.args) > 1 else []) + default_keywords
                default = None
                if len(default_arguments) == 1:
                    default_argument = default_arguments[0]
                    if isinstance(default_argument, ast.Constant) and isinstance(default_argument.value, str):
                        default = default_argument.value
                if contract is not None and default == contract.default:
                    seen_literal_fallback[signature] += 1
                    continue
                undefined_literal_keys.append(f"{relative_path}:{node.lineno}: {argument.value}")
                continue
            expression = ast.unparse(argument)
            signature = (relative_path, expression)
            seen_dynamic[signature] += 1
            contract = active_contracts.get(signature)
            if contract is None:
                unresolved.append(f"{relative_path}:{node.lineno}: {node.func.id}({expression})")
                continue
            used.update(contract.keys & defined)

    mismatches: list[str] = []
    for signature, contract in sorted(active_contracts.items()):
        actual = seen_dynamic.get(signature, 0)
        if actual != contract.expected_count:
            mismatches.append(f"{signature[0]} :: {signature[1]} expected {contract.expected_count}, observed {actual}")
    for signature, contract in sorted(active_literal_contracts.items()):
        actual = seen_literal_fallback.get(signature, 0)
        if actual != contract.expected_count:
            mismatches.append(
                f"{signature[0]} :: literal {signature[1]} expected {contract.expected_count}, observed {actual}"
            )

    contract_keys = (
        set().union(*(contract.keys for contract in active_contracts.values())) if active_contracts else set()
    )
    undefined_contract_keys = contract_keys - defined
    source_files = tuple(path.relative_to(project_root).as_posix() for path in paths)
    return LocaleReferenceAudit(
        defined=frozenset(defined),
        used=frozenset(used),
        unused=frozenset(defined - used),
        unresolved=tuple(sorted(unresolved)),
        undefined_literal_keys=tuple(sorted(undefined_literal_keys)),
        contract_mismatches=tuple(mismatches),
        undefined_contract_keys=frozenset(undefined_contract_keys),
        source_files=source_files,
        source_roots=source_roots,
    )


__all__ = [
    "DYNAMIC_CALL_CONTRACTS",
    "LITERAL_FALLBACK_CONTRACTS",
    "STRUCTURAL_SECTIONS",
    "DynamicCallContract",
    "LiteralFallbackContract",
    "LocaleReferenceAudit",
    "audit_locale_references",
    "defined_locale_keys",
    "locale_leaf_keys",
    "semantic_numbering_keys",
]
