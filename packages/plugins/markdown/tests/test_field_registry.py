from __future__ import annotations

import importlib
import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

from docwen_plugin_markdown.field_processors.gongwen import process_yaml
from docwen_plugin_markdown.field_registry import (
    collect_placeholder_rules,
    collect_special_placeholder_handlers,
    get_available_processors_from_config,
    run_yaml_processors,
)
from docwen_plugin_markdown.template_filler import apply_special_placeholder_handlers

PROJECT_ROOT = Path(__file__).resolve().parents[4]


@pytest.mark.parametrize("invalid_config", [None, "invalid", ["invalid"]])
def test_registry_treats_non_mapping_runtime_config_as_empty(invalid_config: object) -> None:
    yaml_data = {"标题": "公文"}

    run_yaml_processors(yaml_data, invalid_config, current_locale="zh_CN")

    assert yaml_data == {"标题": "公文"}
    assert get_available_processors_from_config(invalid_config, current_locale="zh_CN") == []
    assert collect_placeholder_rules(invalid_config, current_locale="zh_CN") == []
    assert collect_special_placeholder_handlers(invalid_config, current_locale="zh_CN") == {}


def test_gongwen_yaml_processor_applies_canonical_field_semantics() -> None:
    yaml_data = {
        "附件说明": ["1. 材料清单", "（二）办理依据"],
        "抄送机关": ["省政府办公厅", "市政府办公室"],
        "附注": "（此件公开发布）",
        "印发日期": "2026-06-28",
        "成文日期": "2026/6/27",
    }

    process_yaml(yaml_data)

    assert yaml_data["附件说明"] == ["附件：1. 材料清单", "\u3000\u3000\u30002. 办理依据"]
    assert yaml_data["抄送机关"] == "省政府办公厅，市政府办公室"
    assert yaml_data["附注"] == "此件公开发布"
    assert yaml_data["印发日期"] == "2026年6月28日印发"
    assert yaml_data["成文日期"] == "2026年6月27日"


def test_registry_lists_processors_by_order_and_locale() -> None:
    config = {
        "settings": {"order": ["b", "a"]},
        "processors": {
            "a": {"module": "demo.a", "name": "A", "locales": ["*"], "enabled": True},
            "b": {"module": "demo.b", "name": "B", "locales": ["zh_CN"], "enabled": False},
            "c": {"module": "demo.c", "name": "C", "locales": ["en_US"], "enabled": True},
        },
    }

    processors = get_available_processors_from_config(config, current_locale="zh_CN")

    assert [item["id"] for item in processors] == ["b", "a"]
    assert processors[0]["enabled"] is False


def test_run_yaml_processors_uses_enabled_locale_matching_modules(tmp_path, monkeypatch) -> None:
    module_dir = tmp_path / "mods"
    module_dir.mkdir()
    (module_dir / "demo_processor.py").write_text(
        "def process_yaml(data):\n    data['标题'] = data.get('标题', '') + '-processed'\n",
        encoding="utf-8",
    )
    monkeypatch.syspath_prepend(str(module_dir))

    yaml_data = {"标题": "公文"}
    config = {
        "settings": {"order": ["demo", "disabled"]},
        "processors": {
            "demo": {
                "module": "demo_processor",
                "locales": ["zh_CN"],
                "enabled": True,
            },
            "disabled": {
                "module": "demo_processor",
                "locales": ["zh_CN"],
                "enabled": False,
            },
        },
    }

    run_yaml_processors(yaml_data, config, current_locale="zh_CN")

    assert yaml_data["标题"] == "公文-processed"


def test_default_field_processor_config_uses_current_process_yaml_entrypoint() -> None:
    config_path = PROJECT_ROOT / "configs" / "field_processors.toml"
    config = tomllib.loads(config_path.read_text(encoding="utf-8"))
    processor = config["processors"]["gongwen"]
    module = importlib.import_module(processor["module"])

    assert processor["module"] == "docwen_plugin_markdown.field_processors.gongwen"
    assert hasattr(module, "process_yaml")
    assert not hasattr(module, "process_gongwen_yaml")


def test_default_gongwen_field_processor_exposes_placeholder_cleanup_rules() -> None:
    config_path = PROJECT_ROOT / "configs" / "field_processors.toml"
    config = tomllib.loads(config_path.read_text(encoding="utf-8"))

    rules = collect_placeholder_rules(config, current_locale="zh_CN")

    assert rules
    assert ["抄送机关"] in rules[0]["delete_row_if_empty"]
    assert ["印发机关", "印发日期"] in rules[0]["delete_row_if_empty"]
    assert ["附件说明"] in rules[0]["delete_paragraph_if_empty"]


def test_default_gongwen_field_processor_exposes_attachment_special_handler() -> None:
    config_path = PROJECT_ROOT / "configs" / "field_processors.toml"
    config = tomllib.loads(config_path.read_text(encoding="utf-8"))

    handlers = collect_special_placeholder_handlers(config, current_locale="zh_CN")

    assert callable(handlers["附件说明"])
    assert handlers["附件说明"].__name__ == "process_attachment_placeholder"


def test_special_placeholder_dispatch_keeps_non_callable_handler_guard() -> None:
    calls: list[str] = []

    def _handle(_doc: object, _yaml: dict[str, object]) -> bool:
        calls.append("handled")
        return True

    handled = apply_special_placeholder_handlers(
        object(),
        {},
        {"valid": _handle, "invalid": object()},
    )

    assert handled == {"valid"}
    assert calls == ["handled"]
