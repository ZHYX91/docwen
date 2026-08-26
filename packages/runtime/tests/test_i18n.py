"""Unit tests for docwen_runtime.i18n locale table loading."""

from __future__ import annotations

from typing import Any

import pytest

pytestmark = pytest.mark.unit


class TestLoadLocaleTable:
    def test_reads_toml_file_as_dict(self, tmp_path):
        from docwen_runtime.i18n import load_locale_table

        path = tmp_path / "zh_CN.toml"
        path.write_text(
            '[cli.interactive]\ntitle = "标题"\n',
            encoding="utf-8",
        )
        table = load_locale_table(path)
        assert isinstance(table, dict)
        assert table["cli"]["interactive"]["title"] == "标题"

    def test_missing_file_returns_empty_dict(self, tmp_path):
        from docwen_runtime.i18n import load_locale_table

        table = load_locale_table(tmp_path / "nonexistent.toml")
        assert table == {}

    def test_invalid_toml_returns_empty_dict(self, tmp_path):
        from docwen_runtime.i18n import load_locale_table

        path = tmp_path / "bad.toml"
        path.write_text("not = valid = toml = !!!\n", encoding="utf-8")
        table = load_locale_table(path)
        assert table == {}


class TestI18nManager:
    @pytest.mark.parametrize("locale_state", ["missing", "invalid"])
    def test_unreadable_locale_falls_back_to_key(self, tmp_path, locale_state: str) -> None:
        from docwen_runtime.i18n import I18nManager

        if locale_state == "invalid":
            (tmp_path / "zh_CN.toml").write_text(
                "not = valid = toml = !!!\n",
                encoding="utf-8",
            )

        manager = I18nManager(tmp_path)

        assert manager.t("cli.errors.unavailable") == "cli.errors.unavailable"

    def test_localized_options_honor_wildcard_and_locale_metadata(self, tmp_path) -> None:
        from docwen_runtime.i18n import I18nManager

        (tmp_path / "en_US.toml").write_text(
            """
[choices]
universal = "all locales"
english = "current locale"
chinese = "other locale"
unrestricted = "no metadata"
_internal = "hidden"

[_locales.choices]
universal = ["*"]
english = ["en_US"]
chinese = ["zh_CN"]
""".strip()
            + "\n",
            encoding="utf-8",
        )

        manager = I18nManager(tmp_path, default_locale="en_US")

        assert manager.get_localized_options("choices") == {
            "universal": "all locales",
            "english": "current locale",
            "unrestricted": "no metadata",
        }

    def test_localized_options_resolve_dotted_nested_sections(self, tmp_path) -> None:
        from docwen_runtime.i18n import I18nManager

        (tmp_path / "en_US.toml").write_text(
            """
[editors.numbering_add.names]
universal = "all locales"
english = "current locale"
chinese = "other locale"
unrestricted = "no metadata"
_internal = "hidden"

[_locales.editors.numbering_add.names]
universal = ["*"]
english = ["en_US"]
chinese = ["zh_CN"]
""".strip()
            + "\n",
            encoding="utf-8",
        )

        manager = I18nManager(tmp_path, default_locale="en_US")

        assert manager.get_localized_options("editors.numbering_add.names") == {
            "universal": "all locales",
            "english": "current locale",
            "unrestricted": "no metadata",
        }

    def test_localized_options_reject_literal_dotted_table_spelling(self, tmp_path) -> None:
        from docwen_runtime.i18n import I18nManager

        (tmp_path / "en_US.toml").write_text(
            """
["editors.numbering_add.names"]
visible = "Visible"
hidden = "Hidden"

[_locales."editors.numbering_add.names"]
visible = ["en_US"]
hidden = ["zh_CN"]
""".strip()
            + "\n",
            encoding="utf-8",
        )

        manager = I18nManager(tmp_path, default_locale="en_US")

        assert manager.get_localized_options("editors.numbering_add.names") == {}

    def test_style_format_uses_selected_locale_and_totalizes_invalid_entries(self, tmp_path) -> None:
        from docwen_runtime.i18n import I18nManager

        (tmp_path / "en_US.toml").write_text(
            """
[style_formats.body]
ascii_font = "Times New Roman"
font_size_pt = 12

[style_formats.invalid]
value = "valid table"
""".strip()
            + "\n",
            encoding="utf-8",
        )
        (tmp_path / "zh_CN.toml").write_text(
            """
[style_formats.body]
east_asia_font = "仿宋"
font_size_pt = 16

[style_formats]
invalid = "not a table"
""".strip()
            + "\n",
            encoding="utf-8",
        )

        manager = I18nManager(tmp_path, default_locale="en_US")
        assert manager.get_style_format("body") == {
            "ascii_font": "Times New Roman",
            "font_size_pt": 12,
        }
        assert manager.get_style_format("missing") is None
        assert manager.get_style_format("invalid") == {"value": "valid table"}

        manager.set_locale("zh_CN")
        assert manager.get_style_format("body") == {
            "east_asia_font": "仿宋",
            "font_size_pt": 16,
        }
        assert manager.get_style_format("invalid") is None

    def test_clear_cache_reloads_changed_locale_file(self, tmp_path) -> None:
        from docwen_runtime.i18n import I18nManager

        locale_path = tmp_path / "en_US.toml"
        locale_path.write_text('[messages]\nstatus = "before"\n', encoding="utf-8")
        manager = I18nManager(tmp_path, default_locale="en_US")

        assert manager.t("messages.status") == "before"
        locale_path.write_text('[messages]\nstatus = "after"\n', encoding="utf-8")
        assert manager.t("messages.status") == "before"

        manager.clear_cache()

        assert manager.t("messages.status") == "after"


class TestRequestOcrBlockquoteTitle:
    def test_locale_override_precedes_shipped_fallback(self, tmp_path) -> None:
        from docwen_runtime.config import build_ocr_blockquote_title

        (tmp_path / "en_US.toml").write_text(
            '[conversion.ocr_output]\nblockquote_prefix = "English fallback"\n',
            encoding="utf-8",
        )
        snapshot = {
            "conversion": {
                "ocr_output": {
                    "show_blockquote_title": True,
                    "blockquote_title_override_by_locale": {
                        "zh_CN": "中文覆盖",
                        "en_US": "English override",
                    },
                }
            }
        }

        assert (
            build_ocr_blockquote_title(
                snapshot,
                requested_locale="en_US",
                locales_dir=tmp_path,
            )
            == "English override"
        )

    def test_empty_override_uses_requested_locale_resource(self, tmp_path) -> None:
        from docwen_runtime.config import build_ocr_blockquote_title

        (tmp_path / "ja_JP.toml").write_text(
            '[conversion.ocr_output]\nblockquote_prefix = "🖼️ **画像 OCR**:"\n',
            encoding="utf-8",
        )

        assert (
            build_ocr_blockquote_title(
                {
                    "conversion": {
                        "ocr_output": {
                            "show_blockquote_title": True,
                            "blockquote_title_override_by_locale": {},
                        }
                    }
                },
                requested_locale="ja_JP",
                locales_dir=tmp_path,
            )
            == "🖼️ **画像 OCR**:"
        )

    @pytest.mark.parametrize("requested_locale", ["../secret", r"C:\\secret"])
    def test_request_locale_cannot_escape_locale_directory(
        self,
        tmp_path,
        requested_locale: str,
    ) -> None:
        from docwen_runtime.config import build_ocr_blockquote_title

        locales_dir = tmp_path / "locales"
        locales_dir.mkdir()
        (locales_dir / "zh_CN.toml").write_text(
            '[conversion.ocr_output]\nblockquote_prefix = "safe fallback"\n',
            encoding="utf-8",
        )
        (tmp_path / "secret.toml").write_text(
            '[conversion.ocr_output]\nblockquote_prefix = "escaped"\n',
            encoding="utf-8",
        )

        assert (
            build_ocr_blockquote_title(
                {"conversion": {"ocr_output": {"show_blockquote_title": True}}},
                requested_locale=requested_locale,
                locales_dir=locales_dir,
            )
            == "safe fallback"
        )

    def test_disabled_title_does_not_read_locale_resource(self, tmp_path) -> None:
        from docwen_runtime.config import build_ocr_blockquote_title

        assert (
            build_ocr_blockquote_title(
                {
                    "conversion": {
                        "ocr_output": {
                            "show_blockquote_title": False,
                            "blockquote_title_override_by_locale": {"en_US": "ignored"},
                        }
                    }
                },
                requested_locale="en_US",
                locales_dir=tmp_path / "missing",
            )
            == ""
        )


def test_task_manager_injects_request_ocr_title_into_execution_context(tmp_path) -> None:
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.manifest import PluginManifest, RouteSpec
    from docwen_core.models.request import ConversionRequest
    from docwen_core.models.result import ConversionResult
    from docwen_runtime.engine.route_resolver import RouteResolver
    from docwen_runtime.engine.task_manager import TaskManager
    from docwen_runtime.output.finalizer import OutputFinalizer
    from docwen_runtime.plugin_registry.registry import PluginRegistry
    from docwen_runtime.workspace.manager import WorkspaceManager

    class _CaptureOcrTitlePlugin:
        def __init__(self) -> None:
            self.seen: dict[str, str] = {}
            self._manifest = PluginManifest(
                plugin_id="request_ocr_title_probe",
                name="Request OCR title probe",
                version="1.0",
                description="captures request-owned OCR title presentation",
                routes=[
                    RouteSpec(
                        source_format="markdown",
                        target_format="md",
                        action_name="request_ocr_title_probe",
                        label="request OCR title probe",
                    )
                ],
            )

        @property
        def manifest(self) -> PluginManifest:
            return self._manifest

        def can_handle(
            self,
            source_format: str,
            target_format: str,
            action_name: str = "",
        ) -> bool:
            return source_format == "markdown" and target_format == "md" and action_name == "request_ocr_title_probe"

        def convert(self, context: Any) -> ConversionResult:
            self.seen[context.request.request_id] = context.ocr_blockquote_title
            return ConversionResult(task_id=context.request.request_id, success=False)

    source = tmp_path / "probe.md"
    source.write_text("probe\n", encoding="utf-8")
    plugin = _CaptureOcrTitlePlugin()
    plugins = PluginRegistry()
    plugins.register(plugin)
    manager = TaskManager(
        plugins,
        RouteResolver(plugins),
        WorkspaceManager(root_dir=str(tmp_path / "workspace")),
        OutputFinalizer(),
    )

    cases = {
        "override": {
            "show_blockquote_title": True,
            "blockquote_title_override_by_locale": {"en_US": "request override"},
        },
        "fallback": {
            "show_blockquote_title": True,
            "blockquote_title_override_by_locale": {},
        },
        "disabled": {
            "show_blockquote_title": False,
            "blockquote_title_override_by_locale": {"en_US": "ignored"},
        },
    }
    for request_id, ocr_output in cases.items():
        manager.execute_single(
            ConversionRequest(
                request_id=request_id,
                input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
                target_format="md",
                action_name="request_ocr_title_probe",
                options={"locale": "en_US"},
                config_snapshot={"conversion": {"ocr_output": ocr_output}},
            )
        )

    assert plugin.seen == {
        "override": "request override",
        "fallback": "🖼️ **Image OCR**:",
        "disabled": "",
    }
