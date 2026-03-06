"""
config_manager 单元测试

覆盖 ConfigManager 的核心逻辑：深度合并、配置加载、路径获取、配置更新。
使用临时目录避免影响真实配置。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.config.config_manager import ConfigManager
from docwen.config.schemas import CONFIG_FILES, DEFAULT_CONFIG

pytestmark = pytest.mark.unit


# ============================================================
# _deep_merge（核心合并逻辑）
# ============================================================


class TestDeepMerge:
    """深度合并两个字典"""

    def _merge(self, default: dict, user: dict) -> dict:
        """通过临时实例调用 _deep_merge"""
        # 直接访问未绑定方法（_deep_merge 是普通方法）
        cm = ConfigManager.__new__(ConfigManager)
        return cm._deep_merge(default, user)

    def test_user_overrides_default(self) -> None:
        default = {"a": 1, "b": 2}
        user = {"b": 99}
        result = self._merge(default, user)
        assert result == {"a": 1, "b": 99}

    def test_user_adds_new_keys(self) -> None:
        default = {"a": 1}
        user = {"b": 2}
        result = self._merge(default, user)
        assert result == {"a": 1, "b": 2}

    def test_nested_dict_recursive_merge(self) -> None:
        default = {"section": {"x": 1, "y": 2}}
        user = {"section": {"y": 99, "z": 3}}
        result = self._merge(default, user)
        assert result == {"section": {"x": 1, "y": 99, "z": 3}}

    def test_user_replaces_non_dict_with_dict(self) -> None:
        """用户配置类型不同时直接覆盖"""
        default = {"a": "string"}
        user = {"a": {"nested": True}}
        result = self._merge(default, user)
        assert result == {"a": {"nested": True}}

    def test_empty_user(self) -> None:
        default = {"a": 1, "b": {"c": 2}}
        result = self._merge(default, {})
        assert result == default

    def test_empty_default(self) -> None:
        user = {"a": 1}
        result = self._merge({}, user)
        assert result == {"a": 1}

    def test_both_empty(self) -> None:
        assert self._merge({}, {}) == {}

    def test_does_not_mutate_default(self) -> None:
        default = {"a": 1, "nested": {"b": 2}}
        user = {"nested": {"b": 99}}
        original_default = {"a": 1, "nested": {"b": 2}}
        self._merge(default, user)
        assert default == original_default


# ============================================================
# get_config_file_path
# ============================================================


class TestGetConfigFilePath:
    """获取配置文件路径"""

    def test_known_config_returns_path(self) -> None:
        """已知配置名称返回完整路径"""
        cm = ConfigManager()
        path = cm.get_config_file_path("gui_config")
        assert path.endswith("gui_config.toml")
        assert Path(path).is_absolute()

    def test_unknown_config_raises(self) -> None:
        """未知配置名称抛出 ValueError"""
        cm = ConfigManager()
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
        """初始化后 _configs 覆盖 CONFIG_FILES 中的所有配置名"""
        cm = ConfigManager()
        for name in CONFIG_FILES:
            assert name in cm._configs, f"初始化后缺少配置块: '{name}'"


# ============================================================
# ConfigManager 单例
# ============================================================


class TestConfigManagerSingleton:
    """单例行为"""

    def test_same_instance(self) -> None:
        cm1 = ConfigManager()
        cm2 = ConfigManager()
        assert cm1 is cm2

    def test_configs_loaded(self) -> None:
        """初始化后应加载了所有配置块"""
        cm = ConfigManager()
        for name in CONFIG_FILES:
            assert name in cm._configs, f"配置块 '{name}' 未加载"


# ============================================================
# _load_single_config
# ============================================================


class TestLoadSingleConfig:
    """加载单个配置文件"""

    def test_nonexistent_file_returns_empty(self, tmp_path: Path) -> None:
        """不存在的文件返回空字典"""
        cm = ConfigManager.__new__(ConfigManager)
        cm._config_dir = str(tmp_path)
        result = cm._load_single_config("not_exist.toml")
        assert result == {}

    def test_empty_file_returns_empty(self, tmp_path: Path) -> None:
        """空文件返回空字典"""
        empty_file = tmp_path / "empty.toml"
        empty_file.write_text("", encoding="utf-8")
        cm = ConfigManager.__new__(ConfigManager)
        cm._config_dir = str(tmp_path)
        result = cm._load_single_config("empty.toml")
        assert result == {}

    def test_valid_toml_loaded(self, tmp_path: Path) -> None:
        """有效 TOML 文件正确加载"""
        toml_file = tmp_path / "test.toml"
        toml_file.write_text('[section]\nkey = "value"\n', encoding="utf-8")
        cm = ConfigManager.__new__(ConfigManager)
        cm._config_dir = str(tmp_path)
        result = cm._load_single_config("test.toml")
        assert result == {"section": {"key": "value"}}


# ============================================================
# update_config_value
# ============================================================


class TestUpdateConfigValue:
    """更新配置值"""

    def test_unknown_config_name_returns_false(self) -> None:
        cm = ConfigManager()
        result = cm.update_config_value("nonexistent", "section", "key", "value")
        assert result is False


class TestDefaultsMergeAndReload:
    def _new_cm(self, tmp_path: Path) -> ConfigManager:
        cm = ConfigManager.__new__(ConfigManager)
        cm._config_dir = str(tmp_path)
        cm._configs = {}
        return cm

    def test_reload_config_block_applies_defaults_when_file_missing(self, tmp_path: Path) -> None:
        cm = self._new_cm(tmp_path)
        cm._reload_config_block("gui_config")
        assert cm._configs["gui_config"]["window"]["min_height"] == 720

    def test_update_config_value_keeps_defaults_in_memory(self, tmp_path: Path) -> None:
        cm = self._new_cm(tmp_path)
        assert cm.update_config_value("gui_config", "window", "center_panel_width", 123) is True
        assert cm._configs["gui_config"]["window"]["center_panel_width"] == 123
        assert cm._configs["gui_config"]["window"]["min_height"] == 720
