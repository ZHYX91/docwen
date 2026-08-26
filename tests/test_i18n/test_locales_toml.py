"""i18n 单元测试。"""

from __future__ import annotations

import re
import tomllib
from functools import cache
from pathlib import Path
from typing import Any

import pytest

from docwen_gui import i18n as gui_i18n

LOCALES_DIR = Path(__file__).resolve().parents[2] / "i18n" / "locales"
PROJECT_ROOT = LOCALES_DIR.parents[1]

NON_EN_UI_LOCALES = (
    "de_DE",
    "es_ES",
    "fr_FR",
    "ja_JP",
    "ko_KR",
    "pt_BR",
    "ru_RU",
    "vi_VN",
)

# These are ordinary user-facing phrases, not product names, file formats,
# protocol tokens, theme names, or words that are legitimately spelled the
# same in another language. Keeping the list explicit makes the guard stable
# while still preventing a newly copied English block from passing unnoticed.
LOCALIZED_PROSE_KEYS = (
    "settings.text.add_numbering_tooltip",
    "settings.text.scheme_tooltip",
    "settings.text.render_mode_label",
    "settings.text.render_mode_tooltip",
    "settings.text.render_mode_hint",
    "settings.text.save_numbering_schemes_failed",
    "settings.text.save_numbering_clean_rules_failed",
    "settings.proofread.validation_section",
    "settings.proofread.enable_symbol_pairing",
    "settings.proofread.enable_symbol_correction",
    "settings.proofread.enable_typos_rule",
    "settings.proofread.enable_sensitive_word",
    "settings.proofread.skip_code_blocks",
    "settings.proofread.skip_quote_blocks",
    "settings.proofread.symbol_mapping_section",
    "settings.proofread.symbol_mapping_desc",
    "settings.proofread.symbol_correction_section",
    "settings.proofread.symbol_correction_desc",
    "settings.proofread.typos_section",
    "settings.proofread.typos_desc",
    "settings.proofread.sensitive_words_section",
    "settings.proofread.sensitive_words_desc",
    "settings.proofread.edit",
    "settings.logging.system_desc",
    "settings.logging.file_desc",
    "settings.logging.console_desc",
    "settings.logging.validation_custom_directory_required",
    "settings.logging.validation_prefix_invalid",
    "settings.logging.validation_prefix_required",
    "settings.logging.directory_auto_notice",
    "settings.logging.directory_env_override_notice",
    "settings.logging.console_format_tooltip",
    "settings.logging.console_format_label",
    "settings.logging.console_colorize_label",
    "settings.logging.console_colorize_tooltip",
    "settings.logging.open_directory",
    "settings.logging.directory_notice_label",
    "settings.logging.override_source_label",
    "settings.logging.fallback_reason_label",
    "settings.logging.actual_log_file_label",
    "settings.logging.custom_directory_tooltip",
    "settings.logging.custom_directory_label",
    "settings.logging.directory_mode_tooltip",
    "settings.logging.directory_mode_label",
    "settings.logging.location_desc",
    "settings.logging.location_section",
    "settings.logging.dir_modes.user",
    "settings.logging.dir_modes.temp",
    "settings.logging.dir_modes.custom",
    "settings.logging.console_colorize.auto",
    "settings.logging.console_colorize.always",
    "settings.logging.console_colorize.never",
    "settings.image.optimization_section",
    "settings.image.enable_optimization",
    "settings.image.optimization_type_label",
    "editors.numbering_add.word_native_unavailable",
    "editors.numbering_add.word_native_full",
    "editors.numbering_add.word_native_approximate",
    "editors.numbering_add.word_native_incompatible",
    "editors.mapping.source_symbol",
    "editors.mapping.target_symbol",
    "editors.mapping.typo",
    "editors.mapping.multi_value_hint",
    "editors.mapping.save_typos_failed",
    "editors.mapping.save_symbol_correction_failed",
    "editors.mapping.save_sensitive_words_failed",
)


pytestmark = pytest.mark.unit


def _get_nested_value(data: dict, dotted_key: str) -> str:
    """按点分路径读取 TOML 嵌套键值。"""
    current = data
    for part in dotted_key.split("."):
        current = current[part]
    assert isinstance(current, str)
    return current


@cache
def _read_locale_text(path: Path) -> str:
    """Read each immutable shipped locale once per pytest worker."""
    return path.read_text(encoding="utf-8")


def _validate_toml_syntax(path: Path) -> bool:
    try:
        _read_toml_file(path)
    except Exception:
        return False
    return True


@cache
def _read_toml_file(path: Path) -> dict[str, Any]:
    """Parse each immutable shipped locale/config once per pytest worker."""
    return tomllib.loads(_read_locale_text(path))


def _reference_locale_path() -> Path:
    return LOCALES_DIR / "zh_CN.toml"


def test_locales_dir_exists() -> None:
    assert LOCALES_DIR.is_dir()


def test_all_locales_toml_parse_and_have_meta() -> None:
    toml_files = sorted(LOCALES_DIR.glob("*.toml"))
    assert toml_files, "No locale TOML files found"

    for toml_path in toml_files:
        assert _validate_toml_syntax(toml_path) is True, f"Invalid TOML syntax: {toml_path.name}"
        data = _read_toml_file(toml_path)
        assert "meta" in data, f"Missing [meta] in {toml_path.name}"
        assert "name" in data["meta"], f"Missing meta.name in {toml_path.name}"
        assert "native_name" in data["meta"], f"Missing meta.native_name in {toml_path.name}"


def test_image_optimization_settings_never_fall_back_to_english() -> None:
    keys = (
        "settings.image.optimization_section",
        "settings.image.enable_optimization",
        "settings.image.optimization_type_label",
    )
    forbidden_fallbacks = {
        "Optimization",
        "Enable content optimization",
        "Optimization Type:",
    }

    for toml_path in sorted(LOCALES_DIR.glob("*.toml")):
        data = _read_toml_file(toml_path)
        values = [_get_nested_value(data, key).strip() for key in keys]
        assert all(values), f"{toml_path.name} has an empty image optimization label"
        assert not (set(values) & forbidden_fallbacks), (
            f"{toml_path.name} still exposes the generic English image optimization fallback"
        )


def test_numbering_cleanup_display_keys_resolve_in_every_locale() -> None:
    """The shipped cleanup editor must not fall back to raw rule IDs or Chinese-only descriptions."""
    cleanup = _read_toml_file(PROJECT_ROOT / "configs" / "numbering" / "cleanup.toml")
    rules = cleanup.get("rules", [])
    assert isinstance(rules, list) and rules

    key_pairs: list[tuple[str, str]] = []
    for rule in rules:
        assert isinstance(rule, dict)
        name_key = str(rule.get("name_key", ""))
        description_key = str(rule.get("description_key", ""))
        assert name_key, f"cleanup rule {rule.get('id')} has no localized name_key"
        assert description_key, f"cleanup rule {rule.get('id')} has no localized description_key"
        key_pairs.append((name_key, description_key))

    for toml_path in sorted(LOCALES_DIR.glob("*.toml")):
        data = _read_toml_file(toml_path)
        numbering_clean = data["editors"]["numbering_clean"]
        names = numbering_clean["names"]
        descriptions = numbering_clean["descriptions"]
        for name_key, description_key in key_pairs:
            assert str(names.get(name_key, "")).strip(), f"{toml_path.name} lacks cleanup name {name_key}"
            assert str(descriptions.get(description_key, "")).strip(), (
                f"{toml_path.name} lacks cleanup description {description_key}"
            )


def test_about_subtitle_exists_in_all_locales() -> None:
    """AboutDialog uses about.subtitle, so every locale should provide it."""
    reference_path = _reference_locale_path()
    reference_text = _read_locale_text(reference_path)

    assert "about.subtitle" in reference_text or re.search(
        r"^\s*subtitle\s*=",
        reference_text,
        re.MULTILINE,
    ), "zh_CN.toml 应包含 about.subtitle"

    lines = reference_text.splitlines()
    subtitle_line_idx = None
    for i, line in enumerate(lines):
        if re.match(r"^\s*subtitle\s*=", line):
            subtitle_line_idx = i
            break

    assert subtitle_line_idx is not None, "zh_CN.toml 中应存在 subtitle 键"

    toml_files = sorted(LOCALES_DIR.glob("*.toml"))
    for toml_path in toml_files:
        text = _read_locale_text(toml_path)
        about_match = re.search(r"\[about\](.*?)(?=\n\[|\Z)", text, re.DOTALL)
        assert about_match, f"{toml_path.name} 应包含 [about] 段"
        about_section = about_match.group(1)
        assert re.search(r"^\s*subtitle\s*=", about_section, re.MULTILINE), (
            f"{toml_path.name} 的 [about] 段应包含 subtitle"
        )


def test_key_groups_order_and_styles_key_order() -> None:
    reference_path = _reference_locale_path()
    reference_doc = _read_toml_file(reference_path)
    reference_styles_keys = list(reference_doc.get("styles", {}).keys())

    key_sections = ["styles", "style_formats", "placeholders", "yaml_keys"]
    missing_reference_sections = [name for name in key_sections if name not in reference_doc]
    assert not missing_reference_sections, (
        f"Missing required sections in {reference_path.name}: {missing_reference_sections}"
    )

    reference_section_order = [k for k in reference_doc if k in key_sections]
    assert reference_section_order == key_sections

    toml_files = sorted(LOCALES_DIR.glob("*.toml"))
    for toml_path in toml_files:
        actual_doc = _read_toml_file(toml_path)
        missing = [name for name in key_sections if name not in actual_doc]
        assert not missing, f"Missing sections in {toml_path.name}: {missing}"

        actual_section_order = [k for k in actual_doc if k in key_sections]
        assert actual_section_order == reference_section_order, (
            f"Section order mismatch in {toml_path.name}: {actual_section_order} != {reference_section_order}"
        )

        actual_styles_table = actual_doc.get("styles", {})
        actual_styles_keys = list(actual_styles_table.keys())
        if toml_path == reference_path:
            assert actual_styles_keys == reference_styles_keys
            continue

        positions_in_styles = {k: actual_styles_keys.index(k) for k in reference_styles_keys if k in actual_styles_keys}
        missing_style_keys = [k for k in reference_styles_keys if k not in positions_in_styles]
        assert not missing_style_keys, f"Missing styles keys in {toml_path.name}: {missing_style_keys}"

        last = -1
        for key in reference_styles_keys:
            idx = positions_in_styles[key]
            assert idx > last, f"Styles key order incorrect in {toml_path.name}: {key}"
            last = idx

        extras = [k for k in actual_styles_keys if k not in set(reference_styles_keys)]
        if extras:
            extra_positions = [actual_styles_keys.index(k) for k in extras]
            assert all(p > last for p in extra_positions), f"Extra styles keys must be appended in {toml_path.name}"


@pytest.mark.parametrize(
    "key",
    [
        "conversion_panel.layout.invalid_page_range",
        "conversion_panel.layout.invalid_page",
    ],
)
def test_gui_layout_page_errors_are_not_left_as_en_us_placeholders(key: str) -> None:
    """非 en_US 语言包不应把新增 GUI 页码错误文案直接留成英文。"""
    en_data = _read_toml_file(LOCALES_DIR / "en_US.toml")
    en_value = _get_nested_value(en_data, key)

    for toml_path in sorted(LOCALES_DIR.glob("*.toml")):
        if toml_path.name == "en_US.toml":
            continue
        locale_data = _read_toml_file(toml_path)
        locale_value = _get_nested_value(locale_data, key)
        assert locale_value != en_value, f"{toml_path.name} still uses en_US placeholder for {key}"


def test_key_settings_and_editors_do_not_copy_english_prose() -> None:
    """关键设置和编辑器文案不能整段复制英文，同时必须保留占位符契约。"""
    en_data = _read_toml_file(LOCALES_DIR / "en_US.toml")

    for locale in NON_EN_UI_LOCALES:
        locale_data = _read_toml_file(LOCALES_DIR / f"{locale}.toml")
        for key in LOCALIZED_PROSE_KEYS:
            en_value = _get_nested_value(en_data, key)
            locale_value = _get_nested_value(locale_data, key)
            assert locale_value != en_value, f"{locale}.toml still copies English prose for {key}"
            assert sorted(re.findall(r"\{[^{}]+\}", locale_value)) == sorted(re.findall(r"\{[^{}]+\}", en_value)), (
                f"{locale}.toml changed placeholders for {key}"
            )


def test_locale_short_labels_are_not_left_as_english_placeholders() -> None:
    """已收尾的短标签不应再回退成英文缩写占位。"""
    expected_values = {
        "ko_KR.toml": {
            "components.template_selector_tabbed.document_templates": "문서",
            "components.template_selector_tabbed.spreadsheet_templates": "시트",
            "file_types.document": "문서",
            "file_types.image": "이미지",
            "file_category.document_short": "문서",
            "file_category.spreadsheet_short": "시트",
            "file_category.layout_short": "레이아웃",
        },
        "pt_BR.toml": {
            "action_area.export_options": "Opções de exportação",
            "editors.common.description": "Descrição:",
            "file_category.document_short": "Documento",
            "file_category.spreadsheet_short": "Planilha",
        },
    }

    for locale_name, key_values in expected_values.items():
        locale_data = _read_toml_file(LOCALES_DIR / locale_name)
        for key, expected_value in key_values.items():
            assert _get_nested_value(locale_data, key) == expected_value, (
                f"{locale_name} should use localized value for {key}"
            )


def test_zh_tw_gui_status_messages_use_traditional_chinese_wording() -> None:
    """繁中 GUI 状态消息应保持繁体用词，不回退到简体口径。"""
    locale_data = _read_toml_file(LOCALES_DIR / "zh_TW.toml")

    expected_values = {
        "components.file_drop.files_added_msg": "已新增 {count} 個檔案",
        "components.file_drop.files_added_with_skipped_msg": "已新增 {added} 個檔案，跳過 {skipped} 個",
        "components.file_drop.file_selected_msg": "已選取：{filename}",
        "components.file_drop.unsupported_type_msg": "不支援的檔案類型：{filename}",
        "components.template_selector.auto_selected_reason": "已依預設設定自動選取 {template_kind} 中第一個可用範本。",
    }

    for key, expected_value in expected_values.items():
        assert _get_nested_value(locale_data, key) == expected_value, (
            f"zh_TW.toml should keep Traditional Chinese wording for {key}"
        )


def test_locale_markdown_term_uses_consistent_spelling() -> None:
    """locale 中用户可见术语统一使用 Markdown，不回退到旧的 MarkDown 拼写。"""
    for toml_path in sorted(LOCALES_DIR.glob("*.toml")):
        text = _read_locale_text(toml_path)
        assert "MarkDown" not in text, f"{toml_path.name} still contains legacy MarkDown spelling"


def test_language_section_does_not_keep_english_suffix_in_non_en_locales() -> None:
    """非 en_US 语言包不应把语言分区标题保留成“本地词 / Language”混合占位。"""
    en_data = _read_toml_file(LOCALES_DIR / "en_US.toml")
    assert _get_nested_value(en_data, "settings.general.language_section") == "Language"

    for toml_path in sorted(LOCALES_DIR.glob("*.toml")):
        if toml_path.name == "en_US.toml":
            continue
        locale_data = _read_toml_file(toml_path)
        value = _get_nested_value(locale_data, "settings.general.language_section")
        assert "/ Language" not in value, f"{toml_path.name} still keeps English suffix in language_section"


def test_theme_semantic_labels_exist_in_all_locales() -> None:
    expected_values = {
        "zh_CN.toml": {
            "settings.general.themes.light": "浅色",
            "settings.general.themes.dark": "深色",
            "settings.general.themes.system": "跟随系统",
        },
        "zh_TW.toml": {
            "settings.general.themes.light": "淺色",
            "settings.general.themes.dark": "深色",
            "settings.general.themes.system": "跟隨系統",
        },
        "en_US.toml": {
            "settings.general.themes.light": "Light",
            "settings.general.themes.dark": "Dark",
            "settings.general.themes.system": "Follow System",
        },
    }

    for toml_path in sorted(LOCALES_DIR.glob("*.toml")):
        locale_data = _read_toml_file(toml_path)
        for key in (
            "settings.general.themes.light",
            "settings.general.themes.dark",
            "settings.general.themes.system",
        ):
            value = _get_nested_value(locale_data, key)
            assert value.strip(), f"{toml_path.name} should define non-empty theme semantic label for {key}"

        if toml_path.name in expected_values:
            for key, expected_value in expected_values[toml_path.name].items():
                assert _get_nested_value(locale_data, key) == expected_value


def test_zh_tw_batch_retry_actions_use_traditional_chinese_wording() -> None:
    """繁中批量失败重试文案应保持繁体口径。"""
    locale_data = _read_toml_file(LOCALES_DIR / "zh_TW.toml")

    expected_values = {
        "components.file_drop.batch_list.retry_selected_failed": "重試當前失敗項",
        "components.file_drop.batch_list.retry_all_failed": "重試當前分頁全部失敗項",
    }

    for key, expected_value in expected_values.items():
        assert _get_nested_value(locale_data, key) == expected_value, (
            f"zh_TW.toml should keep Traditional Chinese wording for {key}"
        )


def test_chinese_file_selected_message_uses_full_width_colon() -> None:
    """中文文件选择提示统一使用全角冒号。"""
    zh_cn = _read_toml_file(LOCALES_DIR / "zh_CN.toml")
    zh_tw = _read_toml_file(LOCALES_DIR / "zh_TW.toml")

    assert _get_nested_value(zh_cn, "components.file_drop.file_selected_msg") == "已选择：{filename}"
    assert _get_nested_value(zh_tw, "components.file_drop.file_selected_msg") == "已選取：{filename}"


def test_chinese_status_progress_and_failure_messages_use_full_width_colon() -> None:
    """中文 GUI 高频状态消息统一使用全角冒号。"""
    zh_cn = _read_toml_file(LOCALES_DIR / "zh_CN.toml")
    zh_tw = _read_toml_file(LOCALES_DIR / "zh_TW.toml")

    expected_values = {
        "info_area.task_current_file": {
            "zh_CN": "当前文件：{name}",
            "zh_TW": "目前檔案：{name}",
        },
        "info_area.task_progress_detail": {
            "zh_CN": "已完成 {completed}/{total}，失败 {failed}，任务 ID：{operation_id}",
            "zh_TW": "已完成 {completed}/{total}，失敗 {failed}，任務 ID：{operation_id}",
        },
        "components.info_area.task_completion_notification_failed": {
            "zh_CN": "任务已结束：共 {total} 个文件，失败 {failed} 个。",
            "zh_TW": "任務已結束：共 {total} 個檔案，失敗 {failed} 個。",
        },
    }

    for key, locale_values in expected_values.items():
        assert _get_nested_value(zh_cn, key) == locale_values["zh_CN"]
        assert _get_nested_value(zh_tw, key) == locale_values["zh_TW"]


def test_chinese_gui_operation_failure_messages_use_full_width_colon() -> None:
    """中文 GUI 操作失败提示统一使用全角冒号。"""
    zh_cn = _read_toml_file(LOCALES_DIR / "zh_CN.toml")
    zh_tw = _read_toml_file(LOCALES_DIR / "zh_TW.toml")

    expected_values = {
        "main_window.open_path_failed": {
            "zh_CN": "无法打开路径：{path}",
            "zh_TW": "無法開啟路徑：{path}",
        },
        "settings.errors.tab_load_failed_message": {
            "zh_CN": "{tab}加载失败\n\n错误：{error}",
            "zh_TW": "無法載入 {tab}\n\n錯誤：{error}",
        },
    }

    for key, locale_values in expected_values.items():
        assert _get_nested_value(zh_cn, key) == locale_values["zh_CN"]
        assert _get_nested_value(zh_tw, key) == locale_values["zh_TW"]


def test_chinese_open_location_label_uses_full_width_colon() -> None:
    """中文状态栏“打开文件位置”统一使用全角冒号。"""
    zh_cn = _read_toml_file(LOCALES_DIR / "zh_CN.toml")
    zh_tw = _read_toml_file(LOCALES_DIR / "zh_TW.toml")

    assert _get_nested_value(zh_cn, "info_area.open_location") == "打开文件位置：{path}"
    assert _get_nested_value(zh_tw, "info_area.open_location") == "打開檔案位置：{path}"


def test_chinese_batch_progress_messages_use_full_width_colon() -> None:
    """中文批量进度消息统一使用全角冒号。"""
    zh_cn = _read_toml_file(LOCALES_DIR / "zh_CN.toml")
    zh_tw = _read_toml_file(LOCALES_DIR / "zh_TW.toml")

    expected_values = {
        "components.info_area.batch_completed": {
            "zh_CN": "批量处理结束：成功 {success} 个，失败 {failed} 个，跳过 {skipped} 个，取消 {cancelled} 个",
            "zh_TW": "批次處理結束：成功 {success} 個，失敗 {failed} 個，略過 {skipped} 個，取消 {cancelled} 個",
        },
        "info_area.task_progress_detail": {
            "zh_CN": "已完成 {completed}/{total}，失败 {failed}，任务 ID：{operation_id}",
            "zh_TW": "已完成 {completed}/{total}，失敗 {failed}，任務 ID：{operation_id}",
        },
    }

    for key, locale_values in expected_values.items():
        assert _get_nested_value(zh_cn, key) == locale_values["zh_CN"]
        assert _get_nested_value(zh_tw, key) == locale_values["zh_TW"]


def test_batch_completion_locales_preserve_all_terminal_count_placeholders() -> None:
    """Every locale must expose the same four terminal batch outcomes."""
    expected_placeholders = ["{cancelled}", "{failed}", "{skipped}", "{success}"]

    for toml_path in sorted(LOCALES_DIR.glob("*.toml")):
        locale_data = _read_toml_file(toml_path)
        completed = _get_nested_value(locale_data, "components.info_area.batch_completed")
        assert sorted(re.findall(r"\{[^{}]+\}", completed)) == expected_placeholders, toml_path.name
        skipped = _get_nested_value(locale_data, "info_area.task_skipped_count")
        cancelled = _get_nested_value(locale_data, "info_area.task_cancelled_count")
        assert re.findall(r"\{[^{}]+\}", skipped) == ["{skipped}"], toml_path.name
        assert re.findall(r"\{[^{}]+\}", cancelled) == ["{cancelled}"], toml_path.name


def test_chinese_main_window_progress_messages_use_full_width_colon() -> None:
    """中文主窗口进度消息统一使用全角冒号。"""
    zh_cn = _read_toml_file(LOCALES_DIR / "zh_CN.toml")
    zh_tw = _read_toml_file(LOCALES_DIR / "zh_TW.toml")

    expected_values = {
        "main_window.task_processing_prefix": {
            "zh_CN": "正在处理：",
            "zh_TW": "正在處理：",
        },
        "main_window.task_progress_prefix": {
            "zh_CN": "进度：",
            "zh_TW": "進度：",
        },
    }

    for key, locale_values in expected_values.items():
        assert _get_nested_value(zh_cn, key) == locale_values["zh_CN"]
        assert _get_nested_value(zh_tw, key) == locale_values["zh_TW"]


def test_gui_i18n_can_load_locale_file() -> None:
    prev = gui_i18n.get_locale()
    try:
        gui_i18n.set_locale("zh_CN")
        assert gui_i18n.t("meta.name", default="") != ""
    finally:
        gui_i18n.set_locale(prev)
