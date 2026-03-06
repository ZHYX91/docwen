"""
配置模块集合

导出所有配置默认值和 Mixin 类，供 ConfigManager 使用。

架构说明：
    每个配置模块包含两部分：
    1. DEFAULT_XXX_CONFIG: 对应 TOML 文件的默认配置字典
    2. XxxConfigMixin: 提供该配置域的 getter 方法

    ConfigManager 通过多重继承组合所有 Mixin 类，
    实现代码的模块化组织，同时保持对外接口不变。

使用方式：
    from docwen.config.schemas import (
        DEFAULT_CONFIG,
        CONFIG_FILES,
        LoggerConfigMixin,
        GUIConfigMixin,
        # ... 其他 Mixin
    )
"""

# ==============================================================================
#                              导入各模块
# ==============================================================================

from .conversion import (
    DEFAULT_CONVERSION_CONFIG,
    DEFAULT_CONVERSION_DEFAULTS_CONFIG,
    ConversionConfigMixin,
)
from .gui import DEFAULT_GUI_CONFIG, GUIConfigMixin
from .link import DEFAULT_LINK_CONFIG, LinkConfigMixin
from .logger import DEFAULT_LOGGING_CONFIG, LoggerConfigMixin
from .numbering import (
    DEFAULT_HEADING_NUMBERING_CONFIG,
    DEFAULT_NUMBERING_PATTERNS_CONFIG,
    NumberingConfigMixin,
)
from .optimization import (
    DEFAULT_OPTIMIZATION_CONFIG,
    OptimizationConfigMixin,
)
from .output import DEFAULT_OUTPUT_CONFIG, OutputConfigMixin
from .proofread import (
    DEFAULT_PROOFREAD_CONFIG,
    DEFAULT_PROOFREAD_SENSITIVE,
    DEFAULT_PROOFREAD_SYMBOLS,
    DEFAULT_PROOFREAD_TYPOS,
    ProofreadConfigMixin,
)
from .software import DEFAULT_SOFTWARE_PRIORITY_CONFIG, SOFTWARE_ID_MAPPING, SoftwareConfigMixin
from .style import (
    DEFAULT_STYLE_CODE_CONFIG,
    DEFAULT_STYLE_FORMULA_CONFIG,
    DEFAULT_STYLE_QUOTE_CONFIG,
    DEFAULT_STYLE_TABLE_CONFIG,
    StyleConfigMixin,
)

# ==============================================================================
#                              合并的默认配置
# ==============================================================================


def _unwrap_default_block(config_name: str, data: dict) -> dict:
    if len(data) == 1 and config_name in data and isinstance(data[config_name], dict):
        return data[config_name]
    return data


DEFAULT_CONFIG = {
    "logger_config": _unwrap_default_block("logger_config", DEFAULT_LOGGING_CONFIG),
    "gui_config": _unwrap_default_block("gui_config", DEFAULT_GUI_CONFIG),
    "proofread_config": _unwrap_default_block("proofread_config", DEFAULT_PROOFREAD_CONFIG),
    "proofread_symbols": _unwrap_default_block("proofread_symbols", DEFAULT_PROOFREAD_SYMBOLS),
    "proofread_typos": _unwrap_default_block("proofread_typos", DEFAULT_PROOFREAD_TYPOS),
    "proofread_sensitive": _unwrap_default_block("proofread_sensitive", DEFAULT_PROOFREAD_SENSITIVE),
    "output_config": _unwrap_default_block("output_config", DEFAULT_OUTPUT_CONFIG),
    "software_priority": _unwrap_default_block("software_priority", DEFAULT_SOFTWARE_PRIORITY_CONFIG),
    "link_config": _unwrap_default_block("link_config", DEFAULT_LINK_CONFIG),
    "heading_numbering_add": _unwrap_default_block("heading_numbering_add", DEFAULT_HEADING_NUMBERING_CONFIG),
    "heading_numbering_clean": _unwrap_default_block("heading_numbering_clean", DEFAULT_NUMBERING_PATTERNS_CONFIG),
    "style_code": _unwrap_default_block("style_code", DEFAULT_STYLE_CODE_CONFIG),
    "style_quote": _unwrap_default_block("style_quote", DEFAULT_STYLE_QUOTE_CONFIG),
    "style_formula": _unwrap_default_block("style_formula", DEFAULT_STYLE_FORMULA_CONFIG),
    "style_table": _unwrap_default_block("style_table", DEFAULT_STYLE_TABLE_CONFIG),
    "conversion_defaults": _unwrap_default_block("conversion_defaults", DEFAULT_CONVERSION_DEFAULTS_CONFIG),
    "conversion_config": _unwrap_default_block("conversion_config", DEFAULT_CONVERSION_CONFIG),
    "optimization_config": _unwrap_default_block("optimization_config", DEFAULT_OPTIMIZATION_CONFIG),
}

# ==============================================================================
#                              配置文件映射
# ==============================================================================

CONFIG_FILES = {
    "logger_config": "logger_config.toml",
    "gui_config": "gui_config.toml",
    "proofread_config": "proofread_config.toml",
    "proofread_symbols": "proofread_symbols.toml",
    "proofread_typos": "proofread_typos.toml",
    "proofread_sensitive": "proofread_sensitive.toml",
    "output_config": "output_config.toml",
    "software_priority": "software_priority.toml",
    "link_config": "link_config.toml",
    "heading_numbering_add": "heading_numbering_add.toml",
    "heading_numbering_clean": "heading_numbering_clean.toml",
    "style_code": "style_code.toml",
    "style_quote": "style_quote.toml",
    "style_formula": "style_formula.toml",
    "style_table": "style_table.toml",
    "conversion_defaults": "conversion_defaults.toml",
    "conversion_config": "conversion_config.toml",
    "optimization_config": "optimization_config.toml",
}

# ==============================================================================
#                              导出列表
# ==============================================================================

__all__ = [
    "CONFIG_FILES",
    # 合并配置
    "DEFAULT_CONFIG",
    "DEFAULT_CONVERSION_CONFIG",
    "DEFAULT_CONVERSION_DEFAULTS_CONFIG",
    "DEFAULT_GUI_CONFIG",
    "DEFAULT_HEADING_NUMBERING_CONFIG",
    "DEFAULT_LINK_CONFIG",
    # 单独的默认配置（供直接访问）
    "DEFAULT_LOGGING_CONFIG",
    "DEFAULT_NUMBERING_PATTERNS_CONFIG",
    "DEFAULT_OPTIMIZATION_CONFIG",
    "DEFAULT_OUTPUT_CONFIG",
    "DEFAULT_PROOFREAD_CONFIG",
    "DEFAULT_PROOFREAD_SENSITIVE",
    "DEFAULT_PROOFREAD_SYMBOLS",
    "DEFAULT_PROOFREAD_TYPOS",
    "DEFAULT_SOFTWARE_PRIORITY_CONFIG",
    "DEFAULT_STYLE_CODE_CONFIG",
    "DEFAULT_STYLE_FORMULA_CONFIG",
    "DEFAULT_STYLE_QUOTE_CONFIG",
    "DEFAULT_STYLE_TABLE_CONFIG",
    "SOFTWARE_ID_MAPPING",
    "ConversionConfigMixin",
    "GUIConfigMixin",
    "LinkConfigMixin",
    # Mixin 类
    "LoggerConfigMixin",
    "NumberingConfigMixin",
    "OptimizationConfigMixin",
    "OutputConfigMixin",
    "ProofreadConfigMixin",
    "SoftwareConfigMixin",
    "StyleConfigMixin",
]
