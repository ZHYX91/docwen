"""I18n key completeness tests.

Validates that all shipped locale files have the same GUI key set for current GUI
namespaces, following ``docs/configuration.md``.

Policy: code-referenced literal keys and every GUI key in zh_CN must exist
in all locales, including zh_CN. Dynamic call sites require explicit review.
Non-GUI keys (CLI, conversion progress messages, style formats) are allowed
to differ since they are not part of the GUI refactor scope.
"""

from __future__ import annotations

import ast
import string
import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

LOCALES_DIR = Path(__file__).resolve().parents[4] / "i18n" / "locales"
GUI_SOURCE = Path(__file__).resolve().parents[1] / "src" / "docwen_gui"


def _source_translation_keys(source: str) -> tuple[set[str], set[str]]:
    """Inspect actual GUI translation imports, aliases and simple forwarding wrappers."""
    tree = ast.parse(source)
    names: set[str] = set()
    modules: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and (
            node.module == "docwen_gui.i18n" or (node.level and node.module == "i18n")
        ):
            names.update(item.asname or item.name for item in node.names if item.name == "t")
        elif isinstance(node, ast.ImportFrom) and node.module == "docwen_gui":
            modules.update(item.asname or item.name for item in node.names if item.name == "i18n")
        elif isinstance(node, ast.Import):
            modules.update(item.asname or item.name for item in node.names if item.name == "docwen_gui.i18n")

    def translated(call: ast.Call) -> bool:
        return (isinstance(call.func, ast.Name) and call.func.id in names) or (
            isinstance(call.func, ast.Attribute) and call.func.attr == "t" and ast.unparse(call.func.value) in modules
        )

    forwarded: set[int] = set()
    for node in tree.body:
        if not isinstance(node, ast.FunctionDef) or not node.args.args:
            continue
        body = [statement for statement in node.body if not isinstance(statement, ast.Expr)]
        if len(body) == 1 and isinstance(body[0], ast.Return) and isinstance(body[0].value, ast.Call):
            call = body[0].value
            if (
                translated(call)
                and call.args
                and isinstance(call.args[0], ast.Name)
                and call.args[0].id == node.args.args[0].arg
            ):
                names.add(node.name)
                forwarded.add(id(call))

    constants = {
        target.id: node.value.value
        for node in tree.body
        if isinstance(node, ast.Assign) and isinstance(node.value, ast.Constant) and isinstance(node.value.value, str)
        for target in node.targets
        if isinstance(target, ast.Name)
    }

    def literal_keys(value: ast.expr) -> set[str] | None:
        if isinstance(value, ast.Constant) and isinstance(value.value, str):
            return {value.value}
        if isinstance(value, ast.Name) and value.id in constants:
            return {constants[value.id]}
        if isinstance(value, ast.IfExp):
            left, right = literal_keys(value.body), literal_keys(value.orelse)
            if left is not None and right is not None:
                return left | right
        return None

    literals: set[str] = set()
    dynamic: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Call) and translated(node) and id(node) not in forwarded:
            argument = (
                node.args[0] if node.args else next((item.value for item in node.keywords if item.arg == "key"), None)
            )
            assert argument is not None, "Translation call must supply a key"
            keys = literal_keys(argument)
            if keys is None:
                dynamic.add(ast.unparse(argument))
            else:
                literals.update(keys)
    return literals, dynamic


# This is a call-site review inventory, not a claim that arbitrary external keys
# exist. Product-owned finite vocabularies have their normal widget/config tests;
# user/plugin-provided names deliberately retain the documented text fallback.
# New or changed dynamic expressions must be reviewed instead of being ignored.
REVIEWED_DYNAMIC_KEYS = {
    ("dialogs/about.py", "f'about.tools.{tooltip_key}'"): "About tool table",
    ("file_admission_i18n.py", "key"): "Core admission-code to GUI-key mapping",
    ("main_window.py", "f'components.font_size.{normalized}'"): "Font preset vocabulary",
    ("main_window.py", "f'components.font_size.{preset}'"): "Font preset vocabulary",
    ("main_window.py", "state_key"): "Task-state mapping",
    ("numbering_schemes.py", "f'editors.numbering_add.names.{name_key}'"): "Configured scheme with name fallback",
    (
        "view_models/_optimization_filter.py",
        "f'cli.interactive.optimization_types.{resource.id}'",
    ): "Plugin resource with name fallback",
    ("view_models/action_area_vm.py", "key"): "Action label mapping",
    ("view_models/info_area_vm.py", "i18n_key"): "Transient feedback mapping",
    ("view_models/info_area_vm.py", "f'info_area.task_state_{state}'"): "Task-state vocabulary",
    ("view_models/input_area_vm.py", "f'file_types.{category}'"): "File category vocabulary",
    (
        "widgets/action_area.py",
        "label_key_map.get(ft, 'action_area.layout.optimize_for_type')",
    ): "Layout format label map",
    ("widgets/batch_list.py", "f'components.file_drop.status.{filter_key}'"): "Batch status vocabulary",
    ("widgets/batch_list.py", "f'components.file_drop.batch_list.filter_{filter_key}'"): "Batch filter vocabulary",
    ("widgets/batch_list.py", "f'components.file_drop.status.{entry.status}'"): "Validated batch status vocabulary",
    ("widgets/batch_list.py", "f'file_types.{category}'"): "File category vocabulary",
    ("widgets/input_area.py", "label_key"): "File-input action label map passed through the local translation wrapper",
    ("widgets/settings/dialog.py", "value_key"): "Settings change-summary values",
    ("widgets/settings/dialog.py", "key"): "Settings change-summary field map",
    ("widgets/settings/dialog.py", "f'settings.markdown_extensions.{direction}'"): "Markdown input/output directions",
    ("widgets/settings/dialog.py", "f'settings.markdown_extensions.{name}'"): "Markdown extension vocabulary",
    (
        "widgets/settings/formatting_tab.py",
        "f'settings.formatting.syntax.{kind}_extended'",
    ): "Formatting syntax vocabulary",
    ("widgets/settings/formatting_tab.py", "f'settings.formatting.syntax.{kind}_html'"): "Formatting syntax vocabulary",
    (
        "widgets/settings/formatting_tab.py",
        "f'settings.markdown_extensions.{direction}'",
    ): "Markdown input/output directions",
    (
        "widgets/settings/formatting_tab.py",
        "f'settings.markdown_extensions.{direction}_hint'",
    ): "Markdown input/output directions",
    ("widgets/settings/formatting_tab.py", "f'settings.markdown_extensions.{name}'"): "Markdown extension vocabulary",
    (
        "widgets/settings/formatting_tab.py",
        "f'settings.markdown_extensions.{direction}_help.{name}'",
    ): "Markdown extension/direction vocabulary",
    ("widgets/settings/general_tab.py", "f'components.font_size.{preset}'"): "Font preset vocabulary",
    (
        "widgets/settings/numbering_add_editor.py",
        "f'editors.numbering_add.names.{self.name_key}'",
    ): "Configured name with text fallback",
    (
        "widgets/settings/numbering_add_editor.py",
        "f'editors.numbering_add.descriptions.{self.description_key}'",
    ): "Configured description with text fallback",
    ("widgets/settings/numbering_add_editor.py", "f'editors.numbering_add.level_{idx}'"): "Heading levels 1 through 9",
    ("widgets/settings/numbering_clean_editor.py", "key"): "Configured name/description with text fallback",
    ("widgets/settings/proofread_transfer.py", "_CHANGE_LABEL_KEYS[change.kind]"): "Proofreading change-kind map",
    ("widgets/settings/software_tab.py", "label_key"): "Software priority and known-software maps",
    ("widgets/template_selector_tabbed.py", "label_key"): "Template format label map",
}

ALL_LOCALES = [
    "en_US",
    "zh_CN",
    "zh_TW",
    "de_DE",
    "es_ES",
    "fr_FR",
    "ja_JP",
    "ko_KR",
    "pt_BR",
    "ru_RU",
    "vi_VN",
    "ar_SA",
]

GUI_KEY_PREFIXES = (
    "main_window.",
    "components.",
    "conversion_panel.",
    "action_area.",
    "settings.",
    "about.",
    "common.",
    "messages.",
    "editors.",
    "file_types.",
    "file_category.",
    "file_format_validation.",
    "conversion.",
    "components.template_selector",
    "components.template_selector_tabbed",
    "components.file_drop",
    "components.file_selector",
    "components.info_area",
    "components.action_area",
    "components.conversion_panel",
    "info_area.",
)


def _flatten(data: dict, prefix: str = "") -> dict[str, str]:
    result: dict[str, str] = {}
    for key, value in data.items():
        full_key = f"{prefix}.{key}" if prefix else key
        if isinstance(value, dict):
            result.update(_flatten(value, full_key))
        elif value is not None:
            result[full_key] = str(value)
    return result


def _load_locale(locale: str) -> dict[str, str]:
    path = LOCALES_DIR / f"{locale}.toml"
    if not path.exists():
        return {}
    return _flatten(tomllib.loads(path.read_text(encoding="utf-8")))


def _is_gui_key(key: str) -> bool:
    return any(key.startswith(prefix) for prefix in GUI_KEY_PREFIXES)


# ── Tests ───────────────────────────────────────────────────────────────


@pytest.fixture(scope="module")
def source_keys() -> tuple[set[str], set[tuple[str, str]]]:
    literal_keys: set[str] = set()
    dynamic_keys: set[tuple[str, str]] = set()
    for file in GUI_SOURCE.rglob("*.py"):
        literals, dynamic = _source_translation_keys(file.read_text(encoding="utf-8-sig"))
        literal_keys.update(literals)
        dynamic_keys.update((file.relative_to(GUI_SOURCE).as_posix(), expression) for expression in dynamic)
    return literal_keys, dynamic_keys


@pytest.mark.parametrize("locale", ALL_LOCALES)
def test_code_referenced_keys_exist_in_every_locale(locale, source_keys) -> None:
    missing = source_keys[0] - _load_locale(locale).keys()
    assert not missing, f"{locale} is missing code-referenced keys: {sorted(missing)}"


def test_dynamic_translation_calls_require_review(source_keys) -> None:
    actual = source_keys[1]
    reviewed = set(REVIEWED_DYNAMIC_KEYS)
    assert actual == reviewed, f"Unreviewed: {sorted(actual - reviewed)}; stale: {sorted(reviewed - actual)}"
    assert all(REVIEWED_DYNAMIC_KEYS.values())


def test_source_key_scanner_follows_import_aliases_and_forwarding_wrappers() -> None:
    literals, dynamic = _source_translation_keys("""
from docwen_gui.i18n import t as translate
import docwen_gui.i18n as locale
from elsewhere import t
LABEL = "components.example"
def label(key, default=""):
    return translate(key, default)
label(LABEL)
translate("main_window.example")
locale.t(key="common.example")
translate("common.first" if choice else "common.second")
translate(f"common.{suffix}")
t("not.a.gui.key")
""")
    assert literals == {"components.example", "main_window.example", "common.example", "common.first", "common.second"}
    assert dynamic == {"f'common.{suffix}'"}


@pytest.mark.parametrize("locale", ALL_LOCALES)
def test_cancellation_failure_translation_keeps_substitutions(locale) -> None:
    text = _load_locale(locale)["main_window.close_cancel_failed"]
    fields = {field for _, field, _, _ in string.Formatter().parse(text) if field is not None}
    assert fields == {"task_id", "message"}
    rendered = text.format(task_id="fixture-id", message="fixture-error")
    assert "fixture-id" in rendered and "fixture-error" in rendered


class TestLocaleKeyCompleteness:
    """All locale files must have identical sets of GUI keys."""

    @pytest.fixture(scope="class")
    def all_tables(self) -> dict[str, dict[str, str]]:
        return {locale: _load_locale(locale) for locale in ALL_LOCALES}

    @pytest.fixture(scope="class")
    def zh_cn_gui_keys(self, all_tables: dict[str, dict[str, str]]) -> set[str]:
        return {k for k in all_tables["zh_CN"] if _is_gui_key(k)}

    @pytest.mark.parametrize(
        "locale",
        [loc for loc in ALL_LOCALES if loc != "zh_CN"],
    )
    def test_locale_has_all_zh_cn_gui_keys(
        self, locale: str, all_tables: dict[str, dict[str, str]], zh_cn_gui_keys: set[str]
    ) -> None:
        locale_keys = set(all_tables[locale].keys())
        missing = zh_cn_gui_keys - locale_keys
        assert not missing, (
            f"{locale} is missing {len(missing)} GUI keys vs zh_CN:\n"
            + "\n".join(f"  - {k}" for k in sorted(missing)[:20])
            + ("\n  ..." if len(missing) > 20 else "")
        )

    def test_all_locales_have_same_gui_key_count(self, all_tables: dict[str, dict[str, str]]) -> None:
        zh_cn_gui = {k for k in all_tables["zh_CN"] if _is_gui_key(k)}
        deviations: list[tuple[str, int, int]] = []
        for locale in ALL_LOCALES:
            if locale == "zh_CN":
                continue
            locale_gui = {k for k in all_tables[locale] if _is_gui_key(k)}
            if len(locale_gui) != len(zh_cn_gui):
                deviations.append((locale, len(locale_gui), len(zh_cn_gui)))
        assert not deviations, "Some locales have different GUI key counts vs zh_CN:\n" + "\n".join(
            f"  {loc}: {cnt} vs {ref}" for loc, cnt, ref in deviations
        )


class TestNoMissingEnUS:
    """en_US is the English fallback — it must have all keys zh_CN has."""

    def test_en_us_not_missing_any_zh_cn_key(self) -> None:
        zh_cn = _load_locale("zh_CN")
        en_us = _load_locale("en_US")
        missing = set(zh_cn.keys()) - set(en_us.keys())
        assert not missing, f"en_US is missing {len(missing)} keys vs zh_CN:\n" + "\n".join(
            f"  - {k}" for k in sorted(missing)
        )


class TestKeyNaming:
    """GUI keys must follow the naming conventions from the i18n plan."""

    def test_settings_tab_keys_exist(self) -> None:
        zh_cn = _load_locale("zh_CN")
        expected = {
            "settings.tabs.general",
            "settings.tabs.text",
            "settings.tabs.proofread",
            "settings.tabs.document",
            "settings.tabs.spreadsheet",
            "settings.tabs.image",
            "settings.tabs.layout",
            "settings.tabs.link",
            "settings.tabs.formatting",
            "settings.tabs.output",
            "settings.tabs.export",
            "settings.tabs.logging",
            "settings.tabs.other",
        }
        missing = expected - set(zh_cn.keys())
        assert not missing, f"Missing settings tab keys: {missing}"

    def test_main_window_keys_exist(self) -> None:
        zh_cn = _load_locale("zh_CN")
        expected = {
            "main_window.window_title",
            "main_window.settings_tooltip",
            "main_window.about_tooltip",
            "main_window.font_size_tooltip",
            "main_window.version_offline",
        }
        missing = expected - set(zh_cn.keys())
        assert not missing, f"Missing main_window keys: {missing}"

    def test_common_keys_exist(self) -> None:
        zh_cn = _load_locale("zh_CN")
        expected = {"common.ok", "common.cancel", "common.apply"}
        missing = expected - set(zh_cn.keys())
        assert not missing, f"Missing common keys: {missing}"

    def test_settings_reset_keys_exist(self) -> None:
        zh_cn = _load_locale("zh_CN")
        expected = {
            "settings.reset.tab_button",
            "settings.reset.all_button",
            "settings.reset.tab_confirm_title",
            "settings.reset.all_confirm_title",
        }
        missing = expected - set(zh_cn.keys())
        assert not missing, f"Missing settings.reset keys: {missing}"

    @pytest.mark.parametrize("locale", ALL_LOCALES)
    def test_general_theme_keys_are_not_nested_under_languages(self, locale: str) -> None:
        table = _load_locale(locale)
        expected = {
            "settings.general.theme_section",
            "settings.general.theme_description",
            "settings.general.theme_label",
            "settings.general.theme_tooltip",
            "settings.general.theme_preview",
            "settings.general.sample_button",
            "settings.general.sample_text",
        }
        misplaced = {f"settings.general.languages.{key.rsplit('.', 1)[-1]}" for key in expected}

        missing = expected - set(table)
        nested = misplaced & set(table)
        assert not missing, f"{locale} missing settings.general theme keys: {sorted(missing)}"
        assert not nested, f"{locale} has theme keys nested under settings.general.languages: {sorted(nested)}"


@pytest.mark.parametrize("locale", ALL_LOCALES)
def test_link_mode_help_is_multiline_and_covers_all_choices(locale: str) -> None:
    table = _load_locale(locale)
    for kind in ("wiki", "markdown"):
        main_help = table[f"settings.link.{kind}_link_tooltip"]
        fallback_help = table[f"settings.link.{kind}_hyperlink_fallback_tooltip"]
        assert table[f"settings.link.{kind}_hyperlink_fallback"].strip()
        assert len(main_help.splitlines()) >= 7
        assert len(fallback_help.splitlines()) >= 6
        for mode in ("keep", "extract_text", "extract_address", "remove", "hyperlink"):
            assert table[f"settings.link.modes.{mode}"] in main_help
        for mode in ("keep", "extract_text", "extract_address"):
            assert table[f"settings.link.modes.{mode}"] in fallback_help
