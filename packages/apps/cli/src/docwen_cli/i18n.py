"""CLI i18n adapter with inline fallbacks and an optional runtime bridge."""

from __future__ import annotations

from typing import Any

# ── Optional runtime i18n bridge ───────────────────────────────────

_runtime_i18n: Any = None  # I18nManager | None


def _try_runtime_i18n() -> Any:
    """Attempt to create an ``I18nManager`` from the runtime package.

    Falls back to ``None`` when the runtime (or its locale directory) is
    not available so that callers degrade gracefully.
    """
    try:
        from docwen_runtime.i18n import I18nManager
        from docwen_runtime.resources.registry import ResourceRegistry

        locales_dir = ResourceRegistry.default().locales_dir()
        return I18nManager(locales_dir)
    except Exception:
        return None


# Lazy-init singleton.
def _get_runtime_i18n() -> Any:
    global _runtime_i18n
    if _runtime_i18n is None:
        _runtime_i18n = _try_runtime_i18n()
    return _runtime_i18n


# ── Inline translation tables ─────────────────────────────────────

_DEFAULT_LOCALE = "zh_CN"
_current_locale = _DEFAULT_LOCALE

# Key → locale → translation
_TRANSLATIONS: dict[str, dict[str, str]] = {
    # CLI description
    "cli.description": {
        "zh_CN": "DocWen — 文档转换工具",
        "en_US": "DocWen — Document conversion tool",
    },
    "cli.example": {
        "zh_CN": "示例",
        "en_US": "Example",
    },
    # Actions
    "cli.actions.convert": {
        "zh_CN": "格式转换",
        "en_US": "Format conversion",
    },
    "cli.actions.validate": {
        "zh_CN": "文档校对",
        "en_US": "Document proofread",
    },
    "cli.actions.inspect": {
        "zh_CN": "文件信息查询",
        "en_US": "File information inspection",
    },
    # Help texts
    "cli.help.files": {
        "zh_CN": "输入文件路径（支持通配符）",
        "en_US": "Input file paths (supports glob patterns)",
    },
    "cli.help.template": {
        "zh_CN": "模板资源 ID（必须是 resources list templates 返回的精确 canonical ID）",
        "en_US": "Exact canonical template resource ID returned by resources list templates",
    },
    "cli.help.check_punct": {
        "zh_CN": "检测项: punct(标点), typo(错别字), symbol(符号), sensitive(敏感词), all, none",
        "en_US": "Check items: punct, typo, symbol, sensitive, all, none",
    },
    "cli.help.extract_img": {
        "zh_CN": "从文档中提取图片（仅 --to md）",
        "en_US": "Extract images from document (--to md only)",
    },
    "cli.help.ocr": {
        "zh_CN": "对图片启用 OCR 文本识别（仅 --to md）",
        "en_US": "Enable OCR for images (--to md only)",
    },
    "cli.help.ocr_language": {
        "zh_CN": "OCR 识别语言 (auto/chinese/chinese_cht/english/japanese/korean/latin/cyrillic)（需 --ocr）",
        "en_US": "OCR recognition language (auto/chinese/chinese_cht/english/japanese/korean/latin/cyrillic) (requires --ocr)",
    },
    "cli.help.image_mode": {
        "zh_CN": "图片导出模式 (file/base64/embed/omit)（仅 --to md）",
        "en_US": "Image export mode (file/base64/embed/omit) (--to md only)",
    },
    "cli.help.image_link_style": {
        "zh_CN": "Markdown 图片链接样式 (wiki_embed/wiki_link/markdown_embed/markdown_link)（仅 --to md）",
        "en_US": "Markdown image link style (wiki_embed/wiki_link/markdown_embed/markdown_link) (--to md only)",
    },
    "cli.help.table_merge_strategy": {
        "zh_CN": "Markdown 合并单元格导出策略 (fill/empty/marker)（仅 --to md）",
        "en_US": "Markdown merged-cell export strategy (fill/empty/marker) (--to md only)",
    },
    "cli.help.ocr_placement": {
        "zh_CN": "OCR 文本放置位置 (image_md/main_md)（需 --ocr）",
        "en_US": "OCR text placement (image_md/main_md) (requires --ocr)",
    },
    "cli.help.list_optimizations": {
        "zh_CN": "列出可用的 action 型优化项及其适用 scope",
        "en_US": "List available action-based optimization types and their scopes",
    },
    "cli.help.clean_numbering": {
        "zh_CN": "清理 MD 小标题序号: default/remove/keep",
        "en_US": "Clean MD heading numbering: default/remove/keep",
    },
    "cli.help.add_numbering": {
        "zh_CN": "为 MD 小标题添加序号（指定方案 ID）",
        "en_US": "Add numbering to MD headings (scheme ID)",
    },
    "cli.help.heading_merge_mode": {
        "zh_CN": "标题与正文合并模式 (punct_required/always/never)（仅 --to docx）",
        "en_US": "Heading merge mode (punct_required/always/never) (--to docx only)",
    },
    "cli.help.dry_run": {
        "zh_CN": "仅预演检测、归一化与路由，不执行转换",
        "en_US": "Dry run: detect, normalize and route without converting",
    },
    "cli.help.lang": {
        "zh_CN": "界面语言",
        "en_US": "Interface language",
    },
    "cli.help.json": {
        "zh_CN": "JSON 输出模式",
        "en_US": "JSON output mode",
    },
    "cli.help.quiet": {
        "zh_CN": "最小输出",
        "en_US": "Minimal output",
    },
    "cli.help.verbose": {
        "zh_CN": "详细输出",
        "en_US": "Verbose output",
    },
    "cli.help.timing": {
        "zh_CN": "在 JSON 输出中包含每文件耗时",
        "en_US": "Include per-file timing in JSON output",
    },
    "cli.help.batch": {
        "zh_CN": "强制批处理模式",
        "en_US": "Force batch processing mode",
    },
    "cli.help.yes": {
        "zh_CN": "跳过所有确认提示",
        "en_US": "Skip all confirmation prompts",
    },
    "cli.help.continue_on_error": {
        "zh_CN": "出错后继续处理剩余文件",
        "en_US": "Continue processing after errors",
    },
    "cli.help.jobs": {
        "zh_CN": "批处理并发数（默认：1）",
        "en_US": "Number of concurrent batch jobs (default: 1)",
    },
    "cli.help.inspect": {
        "zh_CN": "查询文件可执行的操作",
        "en_US": "Query supported actions for a file",
    },
    "cli.help.list": {
        "zh_CN": "列出可用资源",
        "en_US": "List available resources",
    },
    "cli.help.list_actions": {
        "zh_CN": "列出支持的操作",
        "en_US": "List supported actions",
    },
    "cli.help.list_formats": {
        "zh_CN": "列出可用目标格式",
        "en_US": "List available target formats",
    },
    "cli.help.list_templates": {
        "zh_CN": "列出模板",
        "en_US": "List templates",
    },
    "cli.help.list_numbering_schemes": {
        "zh_CN": "列出编号方案",
        "en_US": "List numbering schemes",
    },
    "cli.help.show_template": {
        "zh_CN": "查看模板详情",
        "en_US": "Show template details",
    },
    "cli.help.show_numbering_scheme": {
        "zh_CN": "查看序号方案详情",
        "en_US": "Show numbering scheme details",
    },
    "cli.help.scope": {
        "zh_CN": "过滤范围",
        "en_US": "Filter scope",
    },
    "cli.help.format_source": {
        "zh_CN": "按源类别过滤",
        "en_US": "Filter by source category",
    },
    "cli.help.schema": {
        "zh_CN": "导出机器可读参数协议",
        "en_US": "Export machine-readable parameter contract",
    },
    "cli.help.schema_convert": {
        "zh_CN": "导出 convert 命令参数协议",
        "en_US": "Export convert command parameter contract",
    },
    "cli.help.doctor": {
        "zh_CN": "环境诊断",
        "en_US": "Environment diagnostics",
    },
    "cli.help.settings": {
        "zh_CN": "设置管理",
        "en_US": "Settings management",
    },
    "cli.help.settings_reset": {
        "zh_CN": "还原设置为默认值",
        "en_US": "Reset settings to defaults",
    },
    "cli.help.settings_reset_tab": {
        "zh_CN": "仅还原指定选项卡的设置",
        "en_US": "Reset only the specified tab",
    },
    "cli.help.output": {
        "zh_CN": "输出目录",
        "en_US": "Output directory",
    },
    "cli.help.target": {
        "zh_CN": "目标格式（docx/xlsx 等）",
        "en_US": "Target format (docx/xlsx etc.)",
    },
    "cli.help.scheme_id": {
        "zh_CN": "序号方案 ID",
        "en_US": "Numbering scheme ID",
    },
    # Messages
    "cli.messages.error_prefix": {
        "zh_CN": "错误",
        "en_US": "Error",
    },
    "cli.messages.error_file_not_found": {
        "zh_CN": "未找到有效文件",
        "en_US": "No valid files found",
    },
    "cli.messages.error_no_valid_files": {
        "zh_CN": "所有文件均无效或格式不支持",
        "en_US": "All files are invalid or unsupported",
    },
    "cli.messages.warning_invalid_files": {
        "zh_CN": "检测到 {count} 个无效文件",
        "en_US": "Detected {count} invalid file(s)",
    },
    "cli.messages.program_interrupted": {
        "zh_CN": "程序已被用户中断",
        "en_US": "Program interrupted by user",
    },
    # Config reset messages are duplicated here intentionally.  Locale TOML
    # files provide the full 11-locale catalogue; this table is the documented
    # zh/en fallback when the runtime package or locale directory is missing.
    "cli.messages.cancelled": {
        "zh_CN": "操作已取消",
        "en_US": "Operation cancelled",
    },
    "cli.messages.invalid_tab": {
        "zh_CN": "未知设置选项卡：{tab}",
        "en_US": "Unknown settings tab: {tab}",
    },
    "cli.messages.confirm_required": {
        "zh_CN": "JSON 模式下重置设置需要 --yes",
        "en_US": "JSON mode requires --yes to reset settings",
    },
    "settings.reset.tab_confirm_message": {
        "zh_CN": "确定要将「{tab_name}」选项卡的设置还原为默认值吗？此操作不可撤销。",
        "en_US": 'Reset all settings in "{tab_name}" tab to defaults? This cannot be undone.',
    },
    "settings.reset.all_confirm_message": {
        "zh_CN": "确定要将所有设置还原为默认值吗？（词库数据不受影响）此操作不可撤销。",
        "en_US": "Reset all settings to defaults? (Dictionary data is not affected.) This cannot be undone.",
    },
    "settings.reset.success": {
        "zh_CN": "所有设置已还原为默认值",
        "en_US": "All settings have been reset to defaults",
    },
    "settings.reset.tab_success": {
        "zh_CN": "「{tab_name}」选项卡设置已还原为默认值",
        "en_US": '"{tab_name}" tab settings have been reset to defaults',
    },
    "settings.reset.failed": {
        "zh_CN": "设置还原失败，请查看日志",
        "en_US": "Failed to reset settings. Please check the logs.",
    },
    "common.error": {
        "zh_CN": "错误",
        "en_US": "Error",
    },
}


def init_cli_locale(lang: str | None = None) -> None:
    """Initialize the CLI locale.  Called early, before arg parsing."""
    global _current_locale
    if lang and lang in _AVAILABLE_CODES:
        _current_locale = lang
    # Also sync the runtime i18n instance if available
    runtime = _get_runtime_i18n()
    if runtime is not None:
        runtime.set_locale(_current_locale)


def get_cli_locale() -> str:
    """Return the currently active CLI locale code."""
    return _current_locale


_AVAILABLE_CODES = frozenset(
    {
        "zh_CN",
        "en_US",
        "de_DE",
        "es_ES",
        "fr_FR",
        "ja_JP",
        "ko_KR",
        "pt_BR",
        "ru_RU",
        "vi_VN",
        "zh_TW",
    }
)


def cli_t(key: str, default: str = "", **fmt_kwargs: Any) -> str:
    """Translate *key* to the current locale.

    Tries the runtime ``I18nManager`` first; falls back to the inline
    translation table when the runtime is not available.

    Args:
        key: Translation key (e.g. ``"cli.help.target"``).
        default: Fallback when key or locale is missing.
        **fmt_kwargs: Format values for string interpolation.

    Returns:
        The translated string, or *key* itself if nothing is found.
    """
    # Try runtime i18n first
    runtime = _get_runtime_i18n()
    if runtime is not None:
        result = runtime.t(key, **fmt_kwargs)
        if result != key or default:
            return result

    # Fall back to inline table
    table = _TRANSLATIONS.get(key)
    if table is None:
        return default or key
    text = table.get(_current_locale) or table.get(_DEFAULT_LOCALE)
    if text is None:
        return default or key
    if fmt_kwargs:
        try:
            return text.format(**fmt_kwargs)
        except (KeyError, ValueError):
            return text
    return text
