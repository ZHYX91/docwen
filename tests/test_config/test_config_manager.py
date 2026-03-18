"""
config_manager 单元测试

覆盖 ConfigManager 的核心逻辑：深度合并、路径获取、配置更新与多实例隔离。
使用临时目录避免影响真实配置文件。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.config.config_manager import ConfigManager, deep_merge_dicts
from docwen.config.schemas import CONFIG_FILES, DEFAULT_CONFIG

pytestmark = pytest.mark.unit


# ============================================================
# deep_merge_dicts（核心合并逻辑）
# ============================================================


class TestDeepMerge:
    """深度合并两个字典"""

    def test_user_overrides_default(self) -> None:
        default = {"a": 1, "b": 2}
        user = {"b": 99}
        result = deep_merge_dicts(default, user)
        assert result == {"a": 1, "b": 99}

    def test_user_adds_new_keys(self) -> None:
        default = {"a": 1}
        user = {"b": 2}
        result = deep_merge_dicts(default, user)
        assert result == {"a": 1, "b": 2}

    def test_nested_dict_recursive_merge(self) -> None:
        default = {"section": {"x": 1, "y": 2}}
        user = {"section": {"y": 99, "z": 3}}
        result = deep_merge_dicts(default, user)
        assert result == {"section": {"x": 1, "y": 99, "z": 3}}

    def test_user_replaces_non_dict_with_dict(self) -> None:
        """用户配置类型不同时直接覆盖"""
        default = {"a": "string"}
        user = {"a": {"nested": True}}
        result = deep_merge_dicts(default, user)
        assert result == {"a": {"nested": True}}

    def test_empty_user(self) -> None:
        default = {"a": 1, "b": {"c": 2}}
        result = deep_merge_dicts(default, {})
        assert result == default

    def test_empty_default(self) -> None:
        user = {"a": 1}
        result = deep_merge_dicts({}, user)
        assert result == {"a": 1}

    def test_both_empty(self) -> None:
        assert deep_merge_dicts({}, {}) == {}

    def test_does_not_mutate_default(self) -> None:
        default = {"a": 1, "nested": {"b": 2}}
        user = {"nested": {"b": 99}}
        original_default = {"a": 1, "nested": {"b": 2}}
        deep_merge_dicts(default, user)
        assert default == original_default


# ============================================================
# get_config_file_path
# ============================================================


class TestGetConfigFilePath:
    """获取配置文件路径"""

    def test_known_config_returns_path(self, tmp_path: Path) -> None:
        """已知配置名称返回完整路径"""
        cm = ConfigManager(config_dir=str(tmp_path))
        path = cm.get_config_file_path("gui_config")
        assert path.endswith("gui_config.toml")
        assert Path(path).is_absolute()

    def test_unknown_config_raises(self, tmp_path: Path) -> None:
        """未知配置名称抛出 ValueError"""
        cm = ConfigManager(config_dir=str(tmp_path))
        with pytest.raises(ValueError, match="未知的配置名称"):
            cm.get_config_file_path("nonexistent_config")


# ============================================================
# CONFIG_FILES 与 DEFAULT_CONFIG 完整性
# ============================================================


class TestConfigIntegrity:
    """配置完整性校验"""

    def test_config_files_mapping_not_empty(self) -> None:
        """CONFIG_FILES 不为空且值都是 .toml 文件名"""
        assert len(CONFIG_FILES) > 0
        for name, filename in CONFIG_FILES.items():
            assert filename.endswith(".toml"), f"CONFIG_FILES['{name}'] 不是 TOML 文件: {filename}"

    def test_default_config_values_are_dicts(self) -> None:
        """DEFAULT_CONFIG 中的每个值都是字典"""
        for name, value in DEFAULT_CONFIG.items():
            assert isinstance(value, dict), f"DEFAULT_CONFIG['{name}'] 不是字典: {type(value)}"

    def test_initialized_configs_cover_all_config_files(self) -> None:
        """DEFAULT_CONFIG 与 CONFIG_FILES 的键集合一致"""
        assert set(CONFIG_FILES.keys()) == set(DEFAULT_CONFIG.keys())


# ============================================================
# update_config_value
# ============================================================


class TestUpdateConfigValue:
    """更新配置值"""

    def test_unknown_config_name_returns_false(self, tmp_path: Path) -> None:
        cm = ConfigManager(config_dir=str(tmp_path))
        result = cm.update_config_value("nonexistent", "section", "key", "value")
        assert result is False


class TestDefaultsMergeAndReload:
    def test_reload_config_block_applies_defaults_when_file_missing(self, tmp_path: Path) -> None:
        cm = ConfigManager(config_dir=str(tmp_path))
        assert cm.get_min_height() == 720

    def test_update_config_value_keeps_defaults_in_memory(self, tmp_path: Path) -> None:
        cm = ConfigManager(config_dir=str(tmp_path))
        assert cm.update_config_value("gui_config", "window", "center_panel_width", 123) is True
        assert cm.get_center_panel_width() == 123
        assert cm.get_min_height() == 720


class TestMultiInstanceIsolation:
    def test_instances_do_not_share_config_dir(self, tmp_path: Path) -> None:
        d1 = tmp_path / "c1"
        d2 = tmp_path / "c2"
        d1.mkdir()
        d2.mkdir()

        cm1 = ConfigManager(config_dir=str(d1))
        cm2 = ConfigManager(config_dir=str(d2))

        assert cm1.update_config_value("gui_config", "window", "center_panel_width", 123) is True
        assert cm1.get_center_panel_width() == 123
        assert cm2.get_center_panel_width() != 123
