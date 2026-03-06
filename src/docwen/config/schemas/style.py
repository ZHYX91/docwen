"""
样式配置模块

对应配置文件：
    - style_code.toml: 代码样式配置
    - style_quote.toml: 引用样式配置
    - style_formula.toml: 公式样式配置
    - style_table.toml: 表格样式配置

包含：
    - DEFAULT_STYLE_CODE_CONFIG: 默认代码样式配置
    - DEFAULT_STYLE_QUOTE_CONFIG: 默认引用样式配置
    - DEFAULT_STYLE_FORMULA_CONFIG: 默认公式样式配置
    - DEFAULT_STYLE_TABLE_CONFIG: 默认表格样式配置
    - StyleConfigMixin: 样式配置获取方法
"""

from typing import Any

from ..safe_logger import safe_log

# ==============================================================================
#                              默认配置 - 代码样式
# ==============================================================================

DEFAULT_STYLE_CODE_CONFIG = {
    "style_code": {
        "docx_to_md": {
            "paragraph_styles": ["HTML Preformatted", "Code Block", "代码块"],
            "character_styles": [
                "HTML Code",
                "HTML Typewriter",
                "HTML Keyboard",
                "HTML Sample",
                "HTML Variable",
                "HTML Definition",
                "HTML Cite",
                "HTML Address",
                "HTML Acronym",
                "Inline Code",
                "Code",
                "Source Code",
                "行内代码",
                "代码",
                "源代码",
                "源码",
            ],
            "full_paragraph_as_block": True,
            "fuzzy_match_enabled": True,
            "fuzzy_keywords": ["code", "代码", "源码"],
            "shading": {"wps_enabled": True, "word_enabled": True},
        },
        "md_to_docx": {"inline_code_style": "Inline Code", "code_block_style": "Code Block"},
    }
}

# ==============================================================================
#                              默认配置 - 引用样式
# ==============================================================================

DEFAULT_STYLE_QUOTE_CONFIG = {
    "style_quote": {
        "docx_to_md": {
            "level_styles": {
                "Quote 1": 1,
                "Quote 2": 2,
                "Quote 3": 3,
                "Quote 4": 4,
                "Quote 5": 5,
                "Quote 6": 6,
                "Quote 7": 7,
                "Quote 8": 8,
                "Quote 9": 9,
                "引用 1": 1,
                "引用 2": 2,
                "引用 3": 3,
                "引用 4": 4,
                "引用 5": 5,
                "引用 6": 6,
                "引用 7": 7,
                "引用 8": 8,
                "引用 9": 9,
            },
            "paragraph_styles": ["Quote", "Block Text", "Intense Quote", "引用", "明显引用"],
            "character_styles": ["Quote Char", "引用字符"],
            "full_paragraph_as_block": True,
            "fuzzy_match_enabled": True,
            "fuzzy_keywords": ["quote", "引用"],
        },
        "md_to_docx": {
            "level_1_style": "Quote 1",
            "level_2_style": "Quote 2",
            "level_3_style": "Quote 3",
            "level_4_style": "Quote 4",
            "level_5_style": "Quote 5",
            "level_6_style": "Quote 6",
            "level_7_style": "Quote 7",
            "level_8_style": "Quote 8",
            "level_9_style": "Quote 9",
        },
    }
}

# ==============================================================================
#                              默认配置 - 公式样式
# ==============================================================================

DEFAULT_STYLE_FORMULA_CONFIG = {
    "style_formula": {"md_to_docx": {"inline_formula_style": "Inline Formula", "formula_block_style": "Formula Block"}}
}

# ==============================================================================
#                              默认配置 - 表格样式
# ==============================================================================

DEFAULT_STYLE_TABLE_CONFIG = {
    "style_table": {
        "md_to_docx": {
            "table_style": "Three Line Table",
            "table_content_style": "Table Content",
            "table_style_mode": "builtin",  # builtin/custom
            "builtin_style_key": "three_line_table",  # three_line_table/table_grid
            "custom_style_name": "",
        }
    }
}

# 配置文件名
CONFIG_FILES = {
    "style_code": "style_code.toml",
    "style_quote": "style_quote.toml",
    "style_formula": "style_formula.toml",
    "style_table": "style_table.toml",
}

# ==============================================================================
#                              Mixin 类
# ==============================================================================


class StyleConfigMixin:
    """
    样式配置获取方法 Mixin

    提供代码、引用、公式、表格样式相关配置的访问方法。
    """

    _configs: dict[str, dict[str, Any]]

    # --------------------------------------------------------------------------
    # 第一层：配置块
    # --------------------------------------------------------------------------

    def get_style_code_block(self) -> dict[str, Any]:
        """
        获取代码样式配置块

        返回:
            Dict[str, Any]: 代码样式配置字典
        """
        return self._configs.get("style_code", {})

    def get_style_quote_block(self) -> dict[str, Any]:
        """
        获取引用样式配置块

        返回:
            Dict[str, Any]: 引用样式配置字典
        """
        return self._configs.get("style_quote", {})

    def get_style_formula_block(self) -> dict[str, Any]:
        """
        获取公式样式配置块

        返回:
            Dict[str, Any]: 公式样式配置字典
        """
        return self._configs.get("style_formula", {})

    def get_style_table_block(self) -> dict[str, Any]:
        """
        获取表格样式配置块

        返回:
            Dict[str, Any]: 表格样式配置字典
        """
        return self._configs.get("style_table", {})

    # ==========================================================================
    #                           代码样式配置
    # ==========================================================================

    def is_code_style_detection_enabled(self) -> bool:
        """
        是否启用代码样式检测（DOCX→MD）

        返回:
            bool: 是否启用代码样式检测（始终返回 True，因为样式列表控制具体行为）
        """
        # 新配置结构中没有 enabled 开关，始终返回 True
        return True

    def get_code_paragraph_styles(self) -> list[str]:
        """
        获取代码块段落样式列表（DOCX→MD）

        返回:
            List[str]: 段落样式名称列表，匹配后转为代码块 ```
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        styles = config.get("paragraph_styles", ["HTML Preformatted", "Code Block", "代码块"])
        safe_log.debug("代码块段落样式列表: %s", styles)
        return styles

    def get_code_character_styles(self) -> list[str]:
        """
        获取行内代码字符样式列表（DOCX→MD）

        返回:
            List[str]: 字符样式名称列表，匹配后转为行内代码 `
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        styles = config.get(
            "character_styles",
            [
                "HTML Code",
                "HTML Typewriter",
                "HTML Keyboard",
                "HTML Sample",
                "HTML Variable",
                "HTML Definition",
                "HTML Cite",
                "HTML Address",
                "HTML Acronym",
                "Inline Code",
                "Code",
                "Source Code",
                "行内代码",
                "代码",
                "源代码",
                "源码",
            ],
        )
        safe_log.debug("行内代码字符样式列表: %s", styles)
        return styles

    def get_code_full_paragraph_as_block(self) -> bool:
        """
        获取是否将整段代码样式的段落视为代码块

        返回:
            bool: 是否整段转代码块
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        return config.get("full_paragraph_as_block", True)

    def get_code_fuzzy_match_enabled(self) -> bool:
        """
        是否启用代码样式模糊匹配

        返回:
            bool: 是否启用模糊匹配
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        enabled = config.get("fuzzy_match_enabled", True)
        safe_log.debug("代码样式模糊匹配启用: %s", enabled)
        return enabled

    def get_code_fuzzy_keywords(self) -> list[str]:
        """
        获取代码样式模糊匹配关键词（不区分大小写）

        返回:
            List[str]: 模糊匹配关键词列表
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        keywords = config.get("fuzzy_keywords", ["code", "代码", "源码"])
        safe_log.debug("代码模糊匹配关键词: %s", keywords)
        return keywords

    # --------------------------------------------------------------------------
    # 底纹检测配置
    # --------------------------------------------------------------------------

    def is_wps_shading_enabled(self) -> bool:
        """
        是否启用 WPS 底纹检测（DOCX→MD）

        WPS 底纹使用纯色填充方式（w:val="clear" + w:fill="D9D9D9"）

        返回:
            bool: 是否启用 WPS 底纹检测
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        shading = config.get("shading", {})
        enabled = shading.get("wps_enabled", True)
        safe_log.debug("WPS 底纹检测启用: %s", enabled)
        return enabled

    def is_word_shading_enabled(self) -> bool:
        """
        是否启用 Word 底纹检测（DOCX→MD）

        Word 底纹使用百分比图案填充方式（w:val="pct15" + w:fill="FFFFFF"）

        返回:
            bool: 是否启用 Word 底纹检测
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        shading = config.get("shading", {})
        enabled = shading.get("word_enabled", True)
        safe_log.debug("Word 底纹检测启用: %s", enabled)
        return enabled

    # --------------------------------------------------------------------------
    # 代码样式 MD→DOCX
    # --------------------------------------------------------------------------

    def get_inline_code_style(self) -> str:
        """
        获取行内代码样式配置（MD→DOCX）

        返回:
            str: 样式名
        """
        config = self.get_style_code_block().get("md_to_docx", {})
        style = config.get("inline_code_style", "Inline Code")
        safe_log.debug("行内代码样式: %s", style)
        return style

    def get_code_block_style(self) -> str:
        """
        获取代码块样式配置（MD→DOCX）

        返回:
            str: 样式名
        """
        config = self.get_style_code_block().get("md_to_docx", {})
        style = config.get("code_block_style", "Code Block")
        safe_log.debug("代码块样式: %s", style)
        return style

    # ==========================================================================
    #                           引用样式配置
    # ==========================================================================

    def is_quote_style_detection_enabled(self) -> bool:
        """
        是否启用引用样式检测（DOCX→MD）

        返回:
            bool: 是否启用引用样式检测（始终返回 True，因为样式列表控制具体行为）
        """
        # 新配置结构中没有 enabled 开关，始终返回 True
        return True

    def get_quote_level_styles(self) -> dict[str, int]:
        """
        获取分级引用样式映射（DOCX→MD）

        返回:
            Dict[str, int]: 样式名到引用级别的映射 {"quote 1": 1, "引用 1": 1, ...}
        """
        config = self.get_style_quote_block().get("docx_to_md", {})
        level_styles = config.get(
            "level_styles",
            {
                "Quote 1": 1,
                "Quote 2": 2,
                "Quote 3": 3,
                "Quote 4": 4,
                "Quote 5": 5,
                "Quote 6": 6,
                "Quote 7": 7,
                "Quote 8": 8,
                "Quote 9": 9,
                "引用 1": 1,
                "引用 2": 2,
                "引用 3": 3,
                "引用 4": 4,
                "引用 5": 5,
                "引用 6": 6,
                "引用 7": 7,
                "引用 8": 8,
                "引用 9": 9,
            },
        )
        safe_log.debug("分级引用样式映射: %d 个样式", len(level_styles))
        return level_styles

    def get_quote_paragraph_styles(self) -> list[str]:
        """
        获取通用引用段落样式列表（DOCX→MD）

        返回:
            List[str]: 通用引用段落样式列表（无法确定级别时默认为1级）
        """
        config = self.get_style_quote_block().get("docx_to_md", {})
        styles = config.get("paragraph_styles", ["Quote", "Block Text", "Intense Quote", "引用", "明显引用"])
        safe_log.debug("通用引用段落样式列表: %s", styles)
        return styles

    def get_quote_character_styles(self) -> list[str]:
        """
        获取引用字符样式列表（DOCX→MD）

        返回:
            List[str]: 引用字符样式列表（转为行内代码）
        """
        config = self.get_style_quote_block().get("docx_to_md", {})
        styles = config.get("character_styles", ["Quote Char", "引用字符"])
        safe_log.debug("引用字符样式列表: %s", styles)
        return styles

    def get_quote_full_paragraph_as_block(self) -> bool:
        """
        获取是否将整段引用样式的段落视为引用块

        返回:
            bool: 是否整段转引用块
        """
        config = self.get_style_quote_block().get("docx_to_md", {})
        return config.get("full_paragraph_as_block", True)

    def get_quote_fuzzy_match_enabled(self) -> bool:
        """
        是否启用引用样式模糊匹配

        返回:
            bool: 是否启用模糊匹配
        """
        config = self.get_style_quote_block().get("docx_to_md", {})
        enabled = config.get("fuzzy_match_enabled", True)
        safe_log.debug("引用样式模糊匹配启用: %s", enabled)
        return enabled

    def get_quote_fuzzy_keywords(self) -> list[str]:
        """
        获取引用样式模糊匹配关键词（不区分大小写）

        返回:
            List[str]: 模糊匹配关键词列表
        """
        config = self.get_style_quote_block().get("docx_to_md", {})
        keywords = config.get("fuzzy_keywords", ["quote", "引用"])
        safe_log.debug("引用模糊匹配关键词: %s", keywords)
        return keywords

    def get_quote_style_for_level(self, level: int) -> str:
        """
        根据引用级别获取样式配置（MD→DOCX）

        参数:
            level: 引用级别 (1-9)

        返回:
            str: 样式名，如 "Quote 1"
        """
        # 确保级别在有效范围内
        level = max(1, min(9, level))

        config = self.get_style_quote_block().get("md_to_docx", {})

        style_key = f"level_{level}_style"
        style = config.get(style_key, f"Quote {level}")

        safe_log.debug("引用级别 %d 样式: %s", level, style)
        return style

    # ==========================================================================
    #                           公式样式配置
    # ==========================================================================

    def get_inline_formula_style(self) -> str:
        """
        获取行内公式样式配置（MD→DOCX）

        返回:
            str: 样式名
        """
        config = self.get_style_formula_block().get("md_to_docx", {})
        style = config.get("inline_formula_style", "Inline Formula")
        safe_log.debug("行内公式样式: %s", style)
        return style

    def get_formula_block_style(self) -> str:
        """
        获取公式块样式配置（MD→DOCX）

        返回:
            str: 样式名
        """
        config = self.get_style_formula_block().get("md_to_docx", {})
        style = config.get("formula_block_style", "Formula Block")
        safe_log.debug("公式块样式: %s", style)
        return style

    # ==========================================================================
    #                           表格样式配置
    # ==========================================================================

    def get_table_style(self) -> str:
        """
        获取表格样式配置（MD→DOCX）

        返回:
            str: 样式名
        """
        config = self.get_style_table_block().get("md_to_docx", {})
        style = config.get("table_style", "Three Line Table")
        safe_log.debug("表格样式: %s", style)
        return style

    def get_table_content_style(self) -> str:
        """
        获取表格内容样式配置（MD→DOCX）

        返回:
            str: 样式名
        """
        config = self.get_style_table_block().get("md_to_docx", {})
        style = config.get("table_content_style", "Table Content")
        safe_log.debug("表格内容样式: %s", style)
        return style

    def get_table_style_mode(self) -> str:
        """
        获取表格样式模式（MD→DOCX）

        返回:
            str: 模式 ("builtin" 使用内置样式, "custom" 使用自定义样式名)
        """
        config = self.get_style_table_block().get("md_to_docx", {})
        mode = config.get("table_style_mode", "builtin")
        if mode not in ["builtin", "custom"]:
            mode = "builtin"
        safe_log.debug("表格样式模式: %s", mode)
        return mode

    def get_builtin_table_style_key(self) -> str:
        """
        获取内置表格样式键名（MD→DOCX）

        返回:
            str: 内置样式键名 ("three_line_table" 或 "table_grid")
        """
        config = self.get_style_table_block().get("md_to_docx", {})
        key = config.get("builtin_style_key", "three_line_table")
        if key not in ["three_line_table", "table_grid"]:
            key = "three_line_table"
        safe_log.debug("内置表格样式键名: %s", key)
        return key

    def get_custom_table_style_name(self) -> str:
        """
        获取自定义表格样式名称（MD→DOCX）

        返回:
            str: 用户自定义的样式名称
        """
        config = self.get_style_table_block().get("md_to_docx", {})
        name = config.get("custom_style_name", "")
        safe_log.debug("自定义表格样式名: %s", name)
        return name
