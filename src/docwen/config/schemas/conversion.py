"""
转换配置模块

对应配置文件：
    - conversion_defaults.toml: 转换默认值配置（控制 GUI 界面的默认值设置）
    - conversion_config.toml: 转换行为配置（直接控制转换引擎的行为规则）

包含：
    - DEFAULT_CONVERSION_DEFAULTS_CONFIG: 默认转换默认值配置
    - DEFAULT_CONVERSION_CONFIG: 默认转换行为配置
    - ConversionConfigMixin: 转换配置获取方法
"""

from typing import Any

from ..safe_logger import safe_log
from .numbering import NumberingConfigMixin

# ==============================================================================
#                              默认配置 - 转换默认值
# ==============================================================================

DEFAULT_CONVERSION_DEFAULTS_CONFIG = {
    "conversion_defaults": {
        "export": {
            "to_md_image_extraction_mode": "base64",
            "to_md_ocr_placement_mode": "main_md",
        },
        "document": {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "to_md_image_extraction_mode": "file",
            "to_md_ocr_placement_mode": "image_md",
            "to_md_remove_numbering": True,
            "to_md_add_numbering": False,
            "to_md_default_scheme": "gongwen_standard",
            "to_md_enable_optimization": False,
            "to_md_optimization_type": "gongwen",
            "enable_symbol_pairing": True,
            "enable_symbol_correction": True,
            "enable_typos_rule": True,
            "enable_sensitive_word": True,
        },
        "spreadsheet": {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "to_md_image_extraction_mode": "file",
            "to_md_ocr_placement_mode": "image_md",
            "merge_mode": 3,
        },
        "image": {
            "to_md_keep_images": True,
            "to_md_enable_ocr": True,
            "to_md_image_extraction_mode": "file",
            "to_md_ocr_placement_mode": "image_md",
            "ocr_language": "auto",  # OCR语言：auto/chinese/japanese/latin/cyrillic
            "compress_mode": "lossless",
            "size_limit": 200,
            "size_unit": "KB",
            "pdf_quality": "original",
            "tiff_mode": "smart",
        },
        "layout": {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "to_md_image_extraction_mode": "file",
            "to_md_ocr_placement_mode": "image_md",
            "to_md_enable_optimization": False,
            "to_md_optimization_type": "invoice_cn",
            "render_dpi": 300,
        },
        "text": {
            "to_docx_remove_numbering": True,
            "to_docx_add_numbering": True,
            "to_docx_default_scheme": "gongwen_standard",
            "to_xlsx_remove_numbering": True,
            "to_xlsx_add_numbering": False,
            "to_xlsx_default_scheme": "hierarchical_standard",
        },
    }
}

# ==============================================================================
#                              默认配置 - 转换行为
# ==============================================================================

DEFAULT_CONVERSION_CONFIG = {
    "conversion_config": {
        # DOCX → MD 格式设置
        "docx_to_md": {
            "preserve_formatting": True,
            "preserve_heading_formatting": False,
            "preserve_table_header_formatting": False,
        },
        # MD → DOCX 格式设置
        "md_to_docx": {
            "formatting_mode": "apply",
            "heading_formatting_mode": "remove",
            "table_header_formatting_mode": "remove",
            "list_separator": "、",
        },
        # Markdown格式语法配置（输出时使用）
        "syntax": {
            "bold": "asterisk",
            "italic": "asterisk",
            "strikethrough": "extended",
            "highlight": "extended",
            "superscript": "html",
            "subscript": "html",
            "unordered_list": "dash",
            "ordered_list": "restart",
            "indent_spaces": 2,
        },
        # 代码识别配置
        "code_detection": {"code_font": "Consolas", "code_background_color": "E7E6E6"},
        "export": {
            "base64_compress_enabled": True,
            "base64_compress_threshold_kb": 200,
        },
        # 分隔符/分页符/分隔线双向转换配置
        "horizontal_rule": {
            "enabled": True,
            # 文档转MD：Word分隔符 → MD分隔符
            # 可选值: "---", "***", "___", "ignore"
            "docx_to_md": {
                "page_break": "---",  # 分页符 → ---
                "section_break": "***",  # 分节符（所有类型统一）→ ***
                "horizontal_rule": "___",  # 分隔线（Horizontal Rule 1/2/3 样式）→ ___
            },
            # MD转文档：MD分隔符 → Word分隔符
            # 可选值: "page_break", "section_break", "horizontal_rule_1", "horizontal_rule_2", "horizontal_rule_3", "ignore"
            "md_to_docx": {
                "dash": "page_break",  # --- → 分页符
                "asterisk": "section_break",  # *** → 分节符（下一页）
                "underscore": "horizontal_rule_1",  # ___ → 分隔线（Horizontal Rule 1 样式）
            },
        },
        # OCR输出格式配置
        "ocr_output": {
            "show_blockquote_title": True,
            "blockquote_title_override_by_locale": {},
        },
    }
}

# 配置文件名
CONFIG_FILES = {"conversion_defaults": "conversion_defaults.toml", "conversion_config": "conversion_config.toml"}

# ==============================================================================
#                              Mixin 类
# ==============================================================================


class ConversionConfigMixin(NumberingConfigMixin):
    """
    转换配置获取方法 Mixin

    提供转换默认值和转换行为相关配置的访问方法。
    """

    _configs: dict[str, dict[str, Any]]

    # ==========================================================================
    #                           第一层：配置块
    # ==========================================================================

    def get_conversion_defaults_block(self) -> dict[str, Any]:
        """
        获取整个转换默认值配置块（控制 GUI 界面的默认值设置）

        返回:
            Dict[str, Any]: 转换默认值配置字典
        """
        return self._configs.get("conversion_defaults", {})

    def get_conversion_defaults(self, section: str) -> dict[str, Any]:
        """
        获取转换默认值配置的指定子表

        Args:
            section: 子表名称（如 document/spreadsheet/layout/image/text）

        Returns:
            Dict[str, Any]: 对应子表的默认配置字典
        """
        return self.get_conversion_defaults_block().get(str(section or ""), {})

    def get_conversion_config_block(self) -> dict[str, Any]:
        """
        获取整个转换行为配置块（直接控制转换引擎的行为规则）

        返回:
            Dict[str, Any]: 转换行为配置字典
        """
        return self._configs.get("conversion_config", {})

    # ==========================================================================
    #                           第二层：子表 - conversion_defaults
    # ==========================================================================

    def get_document_defaults(self) -> dict[str, Any]:
        """
        获取文档文件默认设置子表

        返回:
            Dict[str, Any]: 文档默认设置字典
        """
        return self.get_conversion_defaults_block().get("document", {})

    def get_spreadsheet_defaults(self) -> dict[str, Any]:
        """
        获取表格文件默认设置子表

        返回:
            Dict[str, Any]: 表格默认设置字典
        """
        return self.get_conversion_defaults_block().get("spreadsheet", {})

    def get_image_defaults(self) -> dict[str, Any]:
        """
        获取图片文件默认设置子表

        返回:
            Dict[str, Any]: 图片默认设置字典
        """
        return self.get_conversion_defaults_block().get("image", {})

    def get_layout_defaults(self) -> dict[str, Any]:
        """
        获取版式文件默认设置子表

        返回:
            Dict[str, Any]: 版式默认设置字典
        """
        return self.get_conversion_defaults_block().get("layout", {})

    def get_export_defaults(self) -> dict[str, Any]:
        return self.get_conversion_defaults_block().get("export", {})

    def get_text_defaults(self) -> dict[str, Any]:
        """
        获取文本文件默认设置子表

        返回:
            Dict[str, Any]: 文本默认设置字典
        """
        return self.get_conversion_defaults_block().get("text", {})

    # ==========================================================================
    #                           第二层：子表 - conversion_config
    # ==========================================================================

    def get_export_config(self) -> dict[str, Any]:
        return self.get_conversion_config_block().get("export", {})

    def get_docx_to_md_config(self) -> dict[str, Any]:
        """
        获取DOCX转MD配置子表

        返回:
            Dict[str, Any]: DOCX转MD配置字典
        """
        return self.get_conversion_config_block().get("docx_to_md", {})

    def get_md_to_docx_config(self) -> dict[str, Any]:
        """
        获取MD转DOCX配置子表

        返回:
            Dict[str, Any]: MD转DOCX配置字典
        """
        return self.get_conversion_config_block().get("md_to_docx", {})

    def get_formatting_syntax_config(self) -> dict[str, Any]:
        """
        获取Markdown语法配置子表

        返回:
            Dict[str, Any]: Markdown语法配置字典
        """
        return self.get_conversion_config_block().get("syntax", {})

    def get_code_detection_config(self) -> dict[str, Any]:
        """
        获取代码识别配置子表

        返回:
            Dict[str, Any]: 代码识别配置字典
        """
        return self.get_conversion_config_block().get("code_detection", {})

    def get_horizontal_rule_config(self) -> dict[str, Any]:
        """
        获取分隔符/分页符转换配置子表

        返回:
            Dict[str, Any]: 分隔符转换配置字典
        """
        return self.get_conversion_config_block().get("horizontal_rule", {})

    def get_ocr_output_config(self) -> dict[str, Any]:
        """
        获取 OCR 输出格式配置子表

        返回:
            Dict[str, Any]: OCR 输出配置字典
        """
        return self.get_conversion_config_block().get("ocr_output", {})

    def get_export_to_md_image_extraction_mode(self) -> str:
        export_config = self.get_export_defaults()
        if "to_md_image_extraction_mode" in export_config:
            return export_config.get("to_md_image_extraction_mode", "file")
        return "file"

    def get_export_to_md_ocr_placement_mode(self) -> str:
        export_config = self.get_export_defaults()
        if "to_md_ocr_placement_mode" in export_config:
            return export_config.get("to_md_ocr_placement_mode", "image_md")
        return "image_md"

    def get_export_base64_compress_enabled(self) -> bool:
        export_config = self.get_export_config()
        return bool(export_config.get("base64_compress_enabled", True))

    def get_export_base64_compress_threshold_kb(self) -> int:
        export_config = self.get_export_config()
        try:
            return int(export_config.get("base64_compress_threshold_kb", 200))
        except Exception:
            return 200

    def get_ocr_blockquote_title_override_text(self) -> str | None:
        config = self.get_ocr_output_config()
        overrides = config.get("blockquote_title_override_by_locale", {})
        if not isinstance(overrides, dict):
            return None
        locale = None
        try:
            from typing import Protocol, cast

            class _LocaleProvider(Protocol):
                def get_locale(self) -> str: ...

            locale = cast(_LocaleProvider, self).get_locale()
        except Exception:
            locale = None
        if not locale:
            return None
        value = overrides.get(locale)
        if value is None:
            return None
        return str(value)

    # ==========================================================================
    #                           第三层：文档转换默认值
    # ==========================================================================

    def get_docx_to_md_keep_images(self) -> bool:
        """
        获取文档转MD时是否保留图片

        返回:
            bool: 是否保留图片
        """
        document_config = self.get_document_defaults()
        keep_images = document_config.get("to_md_keep_images", True)
        safe_log.debug("文档转MD保留图片: %s", keep_images)
        return keep_images

    def get_docx_to_md_enable_ocr(self) -> bool:
        """
        获取文档转MD时是否启用OCR

        返回:
            bool: 是否启用OCR
        """
        document_config = self.get_document_defaults()
        enable_ocr = document_config.get("to_md_enable_ocr", False)
        safe_log.debug("文档转MD启用OCR: %s", enable_ocr)
        return enable_ocr

    def get_docx_to_md_image_extraction_mode(self) -> str:
        """
        获取文档转MD时图片提取方式
        返回:
            str: "file" 或 "base64"
        """
        export_config = self.get_export_defaults()
        if "to_md_image_extraction_mode" in export_config:
            mode = export_config.get("to_md_image_extraction_mode", "file")
        else:
            document_config = self.get_document_defaults()
            mode = document_config.get("to_md_image_extraction_mode", "file")
        safe_log.debug("文档转MD图片提取方式: %s", mode)
        return mode

    def get_docx_to_md_ocr_placement_mode(self) -> str:
        """
        获取文档转MD时OCR位置
        返回:
            str: "image_md" 或 "main_md"
        """
        export_config = self.get_export_defaults()
        if "to_md_ocr_placement_mode" in export_config:
            mode = export_config.get("to_md_ocr_placement_mode", "image_md")
        else:
            document_config = self.get_document_defaults()
            mode = document_config.get("to_md_ocr_placement_mode", "image_md")
        safe_log.debug("文档转MD OCR位置: %s", mode)
        return mode

    def get_docx_to_md_remove_numbering(self) -> bool:
        """
        获取文档转MD时是否默认清除原序号

        返回:
            bool: 是否清除原序号
        """
        document_config = self.get_document_defaults()
        remove = document_config.get("to_md_remove_numbering", True)
        safe_log.debug("文档转MD清除原序号: %s", remove)
        return remove

    def get_docx_to_md_add_numbering(self) -> bool:
        """
        获取文档转MD时是否默认添加新序号

        返回:
            bool: 是否添加新序号
        """
        document_config = self.get_document_defaults()
        add = document_config.get("to_md_add_numbering", False)
        safe_log.debug("文档转MD添加新序号: %s", add)
        return add

    def get_docx_to_md_default_scheme(self) -> str:
        """
        获取文档转MD的默认序号方案

        返回:
            str: 序号方案名称，如果配置的方案不存在则返回全局默认方案
        """
        try:
            document_config = self.get_document_defaults()
            scheme_name = document_config.get("to_md_default_scheme", "gongwen_standard")

            # 验证方案是否存在
            heading_config = self.get_heading_numbering_config_block()
            schemes = heading_config.get("schemes", {})

            if scheme_name not in schemes:
                safe_log.warning("配置的序号方案 '%s' 不存在，使用全局默认方案 'gongwen_standard'", scheme_name)
                return "gongwen_standard"

            safe_log.debug("文档转MD默认序号方案: %s", scheme_name)
            return scheme_name
        except Exception as e:
            safe_log.error("获取文档转MD默认序号方案失败: %s", str(e))
            return "gongwen_standard"

    def get_docx_to_md_enable_optimization(self) -> bool:
        """
        获取文档转MD时是否默认启用针对性优化

        返回:
            bool: 是否启用针对文档类型优化
        """
        document_config = self.get_document_defaults()
        enable = document_config.get("to_md_enable_optimization", True)
        safe_log.debug("文档转MD启用优化: %s", enable)
        return enable

    def get_docx_to_md_optimization_type(self) -> str:
        """
        获取文档转MD的默认优化类型ID

        返回:
            str: 优化类型ID（如 gongwen）
        """
        document_config = self.get_document_defaults()
        opt_type = document_config.get("to_md_optimization_type", "gongwen")
        safe_log.debug("文档转MD优化类型: %s", opt_type)
        return opt_type

    def get_document_compress_mode(self) -> str:
        """获取文档压缩模式（暂未使用）"""
        document_config = self.get_document_defaults()
        return document_config.get("compress_mode", "lossless")

    def get_document_enable_symbol_pairing(self) -> bool:
        """获取文档校对-标点配对默认值"""
        document_config = self.get_document_defaults()
        enabled = document_config.get("enable_symbol_pairing", True)
        safe_log.debug("文档标点配对: %s", enabled)
        return enabled

    def get_document_enable_symbol_correction(self) -> bool:
        """获取文档校对-符号校对默认值"""
        document_config = self.get_document_defaults()
        enabled = document_config.get("enable_symbol_correction", True)
        safe_log.debug("文档符号校对: %s", enabled)
        return enabled

    def get_document_enable_typos_rule(self) -> bool:
        """获取文档校对-错别字默认值"""
        document_config = self.get_document_defaults()
        enabled = document_config.get("enable_typos_rule", True)
        safe_log.debug("文档错别字校对: %s", enabled)
        return enabled

    def get_document_enable_sensitive_word(self) -> bool:
        """获取文档校对-敏感词默认值"""
        document_config = self.get_document_defaults()
        enabled = document_config.get("enable_sensitive_word", True)
        safe_log.debug("文档敏感词匹配: %s", enabled)
        return enabled

    # ==========================================================================
    #                           第三层：表格转换默认值
    # ==========================================================================

    def get_xlsx_to_md_keep_images(self) -> bool:
        """
        获取表格转MD时是否保留图片

        返回:
            bool: 是否保留图片
        """
        spreadsheet_config = self.get_spreadsheet_defaults()
        keep_images = spreadsheet_config.get("to_md_keep_images", True)
        safe_log.debug("表格转MD保留图片: %s", keep_images)
        return keep_images

    def get_xlsx_to_md_enable_ocr(self) -> bool:
        """
        获取表格转MD时是否启用OCR

        返回:
            bool: 是否启用OCR
        """
        spreadsheet_config = self.get_spreadsheet_defaults()
        enable_ocr = spreadsheet_config.get("to_md_enable_ocr", False)
        safe_log.debug("表格转MD启用OCR: %s", enable_ocr)
        return enable_ocr

    def get_xlsx_to_md_image_extraction_mode(self) -> str:
        """
        获取表格转MD时图片提取方式
        返回:
            str: "file" 或 "base64"
        """
        export_config = self.get_export_defaults()
        if "to_md_image_extraction_mode" in export_config:
            mode = export_config.get("to_md_image_extraction_mode", "file")
        else:
            spreadsheet_config = self.get_spreadsheet_defaults()
            mode = spreadsheet_config.get("to_md_image_extraction_mode", "file")
        safe_log.debug("表格转MD图片提取方式: %s", mode)
        return mode

    def get_xlsx_to_md_ocr_placement_mode(self) -> str:
        """
        获取表格转MD时OCR位置
        返回:
            str: "image_md" 或 "main_md"
        """
        export_config = self.get_export_defaults()
        if "to_md_ocr_placement_mode" in export_config:
            mode = export_config.get("to_md_ocr_placement_mode", "image_md")
        else:
            spreadsheet_config = self.get_spreadsheet_defaults()
            mode = spreadsheet_config.get("to_md_ocr_placement_mode", "image_md")
        safe_log.debug("表格转MD OCR位置: %s", mode)
        return mode

    def get_spreadsheet_merge_mode(self) -> int:
        """获取表格汇总模式默认值"""
        spreadsheet_config = self.get_spreadsheet_defaults()
        mode = spreadsheet_config.get("merge_mode", 3)
        safe_log.debug("表格汇总模式: %d", mode)
        return mode

    # ==========================================================================
    #                           第三层：图片转换默认值
    # ==========================================================================

    def get_image_to_md_keep_images(self) -> bool:
        """
        获取图片文件转MD时是否保留图片

        返回:
            bool: 是否保留图片
        """
        image_config = self.get_image_defaults()
        keep_images = image_config.get("to_md_keep_images", True)
        safe_log.debug("图片转MD保留图片: %s", keep_images)
        return keep_images

    def get_image_to_md_enable_ocr(self) -> bool:
        """
        获取图片文件转MD时是否启用OCR

        返回:
            bool: 是否启用OCR
        """
        image_config = self.get_image_defaults()
        enable_ocr = image_config.get("to_md_enable_ocr", True)
        safe_log.debug("图片转MD启用OCR: %s", enable_ocr)
        return enable_ocr

    def get_image_to_md_image_extraction_mode(self) -> str:
        """
        获取图片转MD时图片提取方式
        返回:
            str: "file" 或 "base64"
        """
        export_config = self.get_export_defaults()
        if "to_md_image_extraction_mode" in export_config:
            mode = export_config.get("to_md_image_extraction_mode", "file")
        else:
            image_config = self.get_image_defaults()
            mode = image_config.get("to_md_image_extraction_mode", "file")
        safe_log.debug("图片转MD图片提取方式: %s", mode)
        return mode

    def get_image_to_md_ocr_placement_mode(self) -> str:
        """
        获取图片转MD时OCR位置
        返回:
            str: "image_md" 或 "main_md"
        """
        export_config = self.get_export_defaults()
        if "to_md_ocr_placement_mode" in export_config:
            mode = export_config.get("to_md_ocr_placement_mode", "image_md")
        else:
            image_config = self.get_image_defaults()
            mode = image_config.get("to_md_ocr_placement_mode", "image_md")
        safe_log.debug("图片转MD OCR位置: %s", mode)
        return mode

    def get_image_compress_mode(self) -> str:
        """获取图片压缩模式默认值"""
        image_config = self.get_image_defaults()
        mode = image_config.get("compress_mode", "lossless")
        safe_log.debug("图片压缩模式: %s", mode)
        return mode

    def get_image_size_limit(self) -> int:
        """获取图片大小限制默认值"""
        image_config = self.get_image_defaults()
        limit = image_config.get("size_limit", 200)
        safe_log.debug("图片大小限制: %d", limit)
        return limit

    def get_image_size_unit(self) -> str:
        """获取图片大小单位默认值"""
        image_config = self.get_image_defaults()
        unit = image_config.get("size_unit", "KB")
        safe_log.debug("图片大小单位: %s", unit)
        return unit

    def get_image_pdf_quality(self) -> str:
        """获取图片转PDF质量默认值"""
        image_config = self.get_image_defaults()
        quality = image_config.get("pdf_quality", "original")
        safe_log.debug("图片转PDF质量: %s", quality)
        return quality

    def get_image_tiff_mode(self) -> str:
        """获取图片转TIFF模式默认值"""
        image_config = self.get_image_defaults()
        mode = image_config.get("tiff_mode", "smart")
        safe_log.debug("图片转TIFF模式: %s", mode)
        return mode

    def get_ocr_language(self) -> str:
        """
        获取OCR识别语言设置

        返回:
            str: OCR语言配置值
                - "auto": 跟随界面语言自动选择
                - "chinese": 中文/英文
                - "japanese": 日语
                - "latin": 拉丁语系（德语/法语/葡萄牙语等）
                - "cyrillic": 西里尔语系（俄语等）
        """
        image_config = self.get_image_defaults()
        language = image_config.get("ocr_language", "auto")
        # 验证配置值有效性
        valid_languages = ["auto", "chinese", "japanese", "latin", "cyrillic"]
        if language not in valid_languages:
            safe_log.warning("无效的OCR语言配置 '%s'，使用默认值 'auto'", language)
            language = "auto"
        safe_log.debug("OCR语言: %s", language)
        return language

    # ==========================================================================
    #                           第三层：版式转换默认值
    # ==========================================================================

    def get_layout_to_md_keep_images(self) -> bool:
        """
        获取版式文件转MD时是否保留图片

        返回:
            bool: 是否保留图片
        """
        layout_config = self.get_layout_defaults()
        keep_images = layout_config.get("to_md_keep_images", True)
        safe_log.debug("版式转MD保留图片: %s", keep_images)
        return keep_images

    def get_layout_to_md_enable_ocr(self) -> bool:
        """
        获取版式文件转MD时是否启用OCR

        返回:
            bool: 是否启用OCR
        """
        layout_config = self.get_layout_defaults()
        enable_ocr = layout_config.get("to_md_enable_ocr", False)
        safe_log.debug("版式转MD启用OCR: %s", enable_ocr)
        return enable_ocr

    def get_layout_to_md_image_extraction_mode(self) -> str:
        """
        获取版式转MD时图片提取方式
        返回:
            str: "file" 或 "base64"
        """
        export_config = self.get_export_defaults()
        if "to_md_image_extraction_mode" in export_config:
            mode = export_config.get("to_md_image_extraction_mode", "file")
        else:
            layout_config = self.get_layout_defaults()
            mode = layout_config.get("to_md_image_extraction_mode", "file")
        safe_log.debug("版式转MD图片提取方式: %s", mode)
        return mode

    def get_layout_to_md_ocr_placement_mode(self) -> str:
        """
        获取版式转MD时OCR位置
        返回:
            str: "image_md" 或 "main_md"
        """
        export_config = self.get_export_defaults()
        if "to_md_ocr_placement_mode" in export_config:
            mode = export_config.get("to_md_ocr_placement_mode", "image_md")
        else:
            layout_config = self.get_layout_defaults()
            mode = layout_config.get("to_md_ocr_placement_mode", "image_md")
        safe_log.debug("版式转MD OCR位置: %s", mode)
        return mode

    def get_layout_to_md_enable_optimization(self) -> bool:
        """
        获取版式文件转MD时是否启用针对类型优化

        返回:
            bool: 是否启用优化
        """
        layout_config = self.get_layout_defaults()
        enable_optimization = layout_config.get("to_md_enable_optimization", False)
        safe_log.debug("版式转MD启用优化: %s", enable_optimization)
        return enable_optimization

    def get_layout_to_md_optimization_type(self) -> str:
        """
        获取版式文件转MD的优化类型ID

        返回:
            str: 优化类型ID（如 invoice_cn）
        """
        layout_config = self.get_layout_defaults()
        optimization_type = layout_config.get("to_md_optimization_type", "invoice_cn")
        safe_log.debug("版式转MD优化类型: %s", optimization_type)
        return optimization_type

    def get_layout_render_dpi(self) -> int:
        """获取版式文件渲染DPI默认值"""
        layout_config = self.get_layout_defaults()
        dpi = layout_config.get("render_dpi", 300)
        safe_log.debug("版式渲染DPI: %d", dpi)
        return dpi

    # ==========================================================================
    #                           第三层：文本转换默认值
    # ==========================================================================

    def get_md_to_docx_remove_numbering(self) -> bool:
        """
        获取MD转DOCX时是否默认清除原序号

        返回:
            bool: 是否清除原序号
        """
        text_config = self.get_text_defaults()
        remove = text_config.get("to_docx_remove_numbering", True)
        safe_log.debug("MD转DOCX清除原序号: %s", remove)
        return remove

    def get_md_to_docx_add_numbering(self) -> bool:
        """
        获取MD转DOCX时是否默认添加新序号

        返回:
            bool: 是否添加新序号
        """
        text_config = self.get_text_defaults()
        add = text_config.get("to_docx_add_numbering", True)
        safe_log.debug("MD转DOCX添加新序号: %s", add)
        return add

    def get_md_to_docx_default_scheme(self) -> str:
        """
        获取MD转DOCX的默认序号方案

        返回:
            str: 序号方案名称，如果配置的方案不存在则返回全局默认方案
        """
        try:
            text_config = self.get_text_defaults()
            scheme_name = text_config.get("to_docx_default_scheme", "gongwen_standard")

            # 验证方案是否存在
            heading_config = self.get_heading_numbering_config_block()
            schemes = heading_config.get("schemes", {})

            if scheme_name not in schemes:
                safe_log.warning("配置的序号方案 '%s' 不存在，使用全局默认方案 'gongwen_standard'", scheme_name)
                return "gongwen_standard"

            safe_log.debug("MD转DOCX默认序号方案: %s", scheme_name)
            return scheme_name
        except Exception as e:
            safe_log.error("获取MD转DOCX默认序号方案失败: %s", str(e))
            return "gongwen_standard"

    # ==========================================================================
    #                           第三层：转换行为配置
    # ==========================================================================

    def get_docx_to_md_preserve_formatting(self) -> bool:
        """
        获取DOCX转MD时是否保留正文格式

        返回:
            bool: 是否保留格式（转为Markdown标记）
        """
        config = self.get_docx_to_md_config()
        preserve = config.get("preserve_formatting", True)
        safe_log.debug("DOCX转MD保留正文格式: %s", preserve)
        return preserve

    def get_docx_to_md_preserve_heading_formatting(self) -> bool:
        """
        获取DOCX转MD时是否保留小标题格式

        返回:
            bool: 是否保留小标题内的格式标记
        """
        config = self.get_docx_to_md_config()
        preserve = config.get("preserve_heading_formatting", False)
        safe_log.debug("DOCX转MD保留小标题格式: %s", preserve)
        return preserve

    def get_docx_to_md_preserve_table_header_formatting(self) -> bool:
        """
        获取DOCX转MD时是否保留表头格式

        返回:
            bool: 是否保留表头单元格中的格式标记
        """
        config = self.get_docx_to_md_config()
        preserve = config.get("preserve_table_header_formatting", False)
        safe_log.debug("DOCX转MD保留表头格式: %s", preserve)
        return preserve

    def get_md_to_docx_formatting_mode(self) -> str:
        """
        获取MD转DOCX正文格式处理模式

        返回:
            str: 处理模式 ("apply", "keep", "remove")
        """
        config = self.get_md_to_docx_config()
        mode = config.get("formatting_mode", "apply")
        if mode not in ["apply", "keep", "remove"]:
            mode = "apply"
        safe_log.debug("MD转DOCX正文格式处理模式: %s", mode)
        return mode

    def get_md_to_docx_heading_formatting_mode(self) -> str:
        """
        获取MD转DOCX小标题格式处理模式

        返回:
            str: 处理模式
                - "apply": 应用格式，覆盖样式默认格式（未标记文字显式不加粗/不斜体）
                - "keep": 保留标记原样
                - "remove": 清理标记，让Word标题样式格式自然生效
        """
        config = self.get_md_to_docx_config()
        mode = config.get("heading_formatting_mode", "remove")
        if mode not in ["apply", "keep", "remove"]:
            mode = "remove"
        safe_log.debug("MD转DOCX小标题格式处理模式: %s", mode)
        return mode

    def get_md_to_docx_table_header_formatting_mode(self) -> str:
        """
        获取MD转DOCX表头格式处理模式

        返回:
            str: 处理模式
                - "apply": 应用格式（**粗体** → 真正的粗体，可能与表格样式重复）
                - "keep": 保留标记原样
                - "remove": 清理标记，让表格样式（如 firstRow 加粗）自然生效
        """
        config = self.get_md_to_docx_config()
        mode = config.get("table_header_formatting_mode", "remove")
        if mode not in ["apply", "keep", "remove"]:
            mode = "remove"
        safe_log.debug("MD转DOCX表头格式处理模式: %s", mode)
        return mode

    def get_yaml_list_separator(self) -> str:
        """
        获取YAML列表值的拼接符

        当YAML字段值为列表时，使用此分隔符将元素拼接为字符串。
        用于文档（DOCX）和表格（XLSX）的占位符替换。

        返回:
            str: 拼接符字符串，默认为顿号 "、"
        """
        config = self.get_md_to_docx_config()
        separator = config.get("list_separator", "、")
        safe_log.debug("YAML列表拼接符: %s", separator)
        return separator

    # ==========================================================================
    #                           第三层：Markdown语法配置
    # ==========================================================================

    def get_syntax_setting(self, key: str, default: str | None = None) -> str:
        """
        获取特定格式的Markdown语法配置

        参数:
            key: 格式键名（bold, italic, strikethrough, highlight, superscript, subscript）
            default: 默认值

        返回:
            str: 语法选项（如 "asterisk", "underscore", "extended", "html"）
        """
        config = self.get_formatting_syntax_config()
        value = config.get(key, default)
        safe_log.debug("获取语法配置 %s: %s", key, value)
        if value is None:
            return ""
        return value

    def get_all_syntax_settings(self) -> dict[str, str]:
        """
        获取所有Markdown语法配置

        返回:
            Dict[str, str]: 语法配置字典
        """
        default_syntax = {
            "bold": "asterisk",
            "italic": "asterisk",
            "strikethrough": "extended",
            "highlight": "extended",
            "superscript": "html",
            "subscript": "html",
        }
        config = self.get_formatting_syntax_config()
        # 合并默认值和用户配置
        result = {**default_syntax, **config}
        safe_log.debug("获取所有语法配置: %s", result)
        return result

    def get_code_font(self) -> str:
        """
        获取代码字体

        返回:
            str: 字体名称
        """
        config = self.get_code_detection_config()
        font = config.get("code_font", "Consolas")
        return font

    def get_code_background_color(self) -> str:
        """
        获取代码背景颜色

        返回:
            str: 背景颜色（十六进制）
        """
        config = self.get_code_detection_config()
        color = config.get("code_background_color", "E7E6E6")
        return color

    def get_unordered_list_marker(self) -> str:
        """
        获取无序列表标记

        返回:
            str: 标记类型 ("dash", "asterisk", "plus")
        """
        config = self.get_formatting_syntax_config()
        marker = config.get("unordered_list", "dash")
        if marker not in ["dash", "asterisk", "plus"]:
            marker = "dash"
        safe_log.debug("获取无序列表标记: %s", marker)
        return marker

    def get_ordered_list_mode(self) -> str:
        """
        获取有序列表编号模式（暂未实现，预留）

        返回:
            str: 模式 ("restart", "preserve")
        """
        config = self.get_formatting_syntax_config()
        mode = config.get("ordered_list", "restart")
        if mode not in ["restart", "preserve"]:
            mode = "restart"
        safe_log.debug("获取有序列表模式: %s", mode)
        return mode

    def get_list_indent_spaces(self) -> int:
        """
        获取列表每级缩进的空格数

        用于 MD → DOCX 转换时识别多级列表嵌套级别。

        返回:
            int: 每级缩进的空格数，默认为 2
        """
        config = self.get_formatting_syntax_config()
        spaces = config.get("indent_spaces", 2)
        # 确保在有效范围内
        if not isinstance(spaces, int) or spaces < 1:
            spaces = 2
        safe_log.debug("获取列表缩进空格数: %d", spaces)
        return spaces

    # ==========================================================================
    #                           第三层：分隔符/分页符配置
    # ==========================================================================

    def is_horizontal_rule_enabled(self) -> bool:
        """
        是否启用分隔符转换

        返回:
            bool: 是否启用
        """
        config = self.get_horizontal_rule_config()
        enabled = config.get("enabled", True)
        safe_log.debug("分隔符转换启用: %s", enabled)
        return enabled

    def get_horizontal_rule_docx_to_md_config(self) -> dict[str, str]:
        """
        获取文档转MD的分隔符配置

        返回:
            Dict[str, str]: {Word分隔符类型: MD分隔符}
                - page_break: 分页符对应的MD分隔符
                - section_break: 分节符对应的MD分隔符（所有类型统一）
                - horizontal_rule: 分隔线对应的MD分隔符（Horizontal Rule 1/2/3 样式统一转换）
        """
        config = self.get_horizontal_rule_config()
        docx_to_md = config.get("docx_to_md", {})
        defaults = {"page_break": "---", "section_break": "***", "horizontal_rule": "___"}
        result = {
            "page_break": docx_to_md.get("page_break", defaults["page_break"]),
            "section_break": docx_to_md.get("section_break", defaults["section_break"]),
            "horizontal_rule": docx_to_md.get("horizontal_rule", defaults["horizontal_rule"]),
        }
        safe_log.debug("文档转MD分隔符配置: %s", result)
        return result

    def get_horizontal_rule_md_to_docx_config(self) -> dict[str, str]:
        """
        获取MD转文档的分隔符配置

        返回:
            Dict[str, str]: {MD分隔符类型: Word分隔符}
                - dash: --- 对应的Word元素
                - asterisk: *** 对应的Word元素
                - underscore: ___ 对应的Word元素（可选 horizontal_rule_1/2/3）
        """
        config = self.get_horizontal_rule_config()
        md_to_docx = config.get("md_to_docx", {})
        defaults = {"dash": "page_break", "asterisk": "section_break", "underscore": "horizontal_rule_1"}
        result = {
            "dash": md_to_docx.get("dash", defaults["dash"]),
            "asterisk": md_to_docx.get("asterisk", defaults["asterisk"]),
            "underscore": md_to_docx.get("underscore", defaults["underscore"]),
        }
        safe_log.debug("MD转文档分隔符配置: %s", result)
        return result

    def get_horizontal_rule_mapping(self, hr_type: str) -> str:
        """
        获取指定分隔符类型的转换目标（MD→DOCX）

        参数:
            hr_type: 分隔符类型 ("dash", "asterisk", "underscore")

        返回:
            str: 转换目标 ("page_break", "section_continuous", "section_break", "ignore")
        """
        config = self.get_horizontal_rule_md_to_docx_config()
        target = config.get(hr_type, "page_break")
        safe_log.debug("分隔符 %s 映射目标: %s", hr_type, target)
        return target

    def get_all_horizontal_rule_mappings(self) -> dict[str, str]:
        """
        获取所有分隔符映射配置（MD→DOCX方向）

        返回:
            Dict[str, str]: {分隔符类型: 转换目标}
        """
        return self.get_horizontal_rule_md_to_docx_config()

    def get_md_separator_for_break_type(self, break_type: str) -> str | None:
        """
        根据 Word 分页/分节/分隔线类型获取对应的 MD 分隔符（DOCX→MD）

        参数:
            break_type: Word 元素类型
                - "page_break": 分页符
                - "section_next", "section_continuous", "section_even", "section_odd": 分节符
                - "horizontal_rule": 分隔线（Horizontal Rule 1/2/3 样式）

        返回:
            str: MD 分隔符 ("---", "***", "___")，未找到或忽略时返回 None
        """
        config = self.get_horizontal_rule_docx_to_md_config()

        # 将所有分节符类型统一映射到 section_break 配置键
        type_to_config_key = {
            "page_break": "page_break",
            "section_continuous": "section_break",
            "section_next": "section_break",
            "section_even": "section_break",
            "section_odd": "section_break",
            "horizontal_rule": "horizontal_rule",
        }
        config_key = type_to_config_key.get(break_type, break_type)

        # 获取配置的MD分隔符
        md_separator = config.get(config_key)

        if md_separator and md_separator != "ignore":
            safe_log.debug("分隔符类型 %s (配置键: %s) 映射为: %s", break_type, config_key, md_separator)
            return md_separator

        safe_log.debug("分隔符类型 %s 映射为忽略或未配置", break_type)
        return None

    def get_ocr_blockquote_title_enabled(self) -> bool:
        """
        获取 OCR 输出到主 MD 时是否显示引用块标题行

        返回:
            bool: True 显示标题行，False 不显示标题行
        """
        config = self.get_ocr_output_config()
        enabled = config.get("show_blockquote_title", True)
        return bool(enabled)
