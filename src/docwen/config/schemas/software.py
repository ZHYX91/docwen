"""
软件优先级配置模块

对应配置文件：software_priority.toml

包含：
    - DEFAULT_SOFTWARE_PRIORITY_CONFIG: 默认软件优先级配置
    - SOFTWARE_ID_MAPPING: 软件标识符到COM对象的映射
    - SoftwareConfigMixin: 软件优先级配置获取方法
"""

from typing import Any

from ..safe_logger import safe_log

# ==============================================================================
#                              默认配置
# ==============================================================================

DEFAULT_SOFTWARE_PRIORITY_CONFIG = {
    "software_priority": {
        "default_priority": {
            "word_processors": ["wps_writer", "msoffice_word", "libreoffice"],
            "spreadsheet_processors": ["wps_spreadsheets", "msoffice_excel", "libreoffice"],
        },
        "special_conversions": {
            "odt": ["msoffice_word", "libreoffice"],
            "ods": ["msoffice_excel", "libreoffice"],
            "pdf_to_office": ["msoffice_word", "libreoffice"],
            "document_to_pdf": ["wps_writer", "msoffice_word", "libreoffice"],
            "spreadsheet_to_pdf": ["wps_spreadsheets", "msoffice_excel", "libreoffice"],
        },
    }
}

# 软件标识符到COM对象的映射
SOFTWARE_ID_MAPPING = {
    "wps_writer": "Kwps.Application",
    "wps_spreadsheets": "Ket.Application",
    "msoffice_word": "Word.Application",
    "msoffice_excel": "Excel.Application",
    "libreoffice": "soffice",
}

# 配置文件名
CONFIG_FILE = "software_priority.toml"

# ==============================================================================
#                              Mixin 类
# ==============================================================================


class SoftwareConfigMixin:
    """
    软件优先级配置获取方法 Mixin

    提供软件优先级相关配置的访问方法。
    """

    _configs: dict[str, dict[str, Any]]

    # --------------------------------------------------------------------------
    # 第一层：配置块
    # --------------------------------------------------------------------------

    def get_software_priority_block(self) -> dict[str, Any]:
        """
        获取整个软件优先级配置块

        返回:
            Dict[str, Any]: 软件优先级配置字典
        """
        return self._configs.get("software_priority", {})

    # --------------------------------------------------------------------------
    # 第二层：子表
    # --------------------------------------------------------------------------

    def get_default_priority_config(self) -> dict[str, Any]:
        """
        获取默认优先级配置子表

        返回:
            Dict[str, Any]: 默认优先级配置字典
        """
        return self.get_software_priority_block().get("default_priority", {})

    def get_special_conversions_config(self) -> dict[str, Any]:
        """
        获取特殊转换配置子表

        返回:
            Dict[str, Any]: 特殊转换配置字典
        """
        return self.get_software_priority_block().get("special_conversions", {})

    # --------------------------------------------------------------------------
    # 第三层：具体配置值
    # --------------------------------------------------------------------------

    def get_word_processors_priority(self) -> list[str]:
        """
        获取文档处理软件优先级列表

        返回:
            List[str]: 软件标识符列表，按优先级排序
        """
        default_priority = self.get_default_priority_config()
        priority = default_priority.get("word_processors", ["wps_writer", "msoffice_word", "libreoffice"])
        safe_log.debug("文档处理软件优先级: %s", priority)
        return priority

    def get_spreadsheet_processors_priority(self) -> list[str]:
        """
        获取表格处理软件优先级列表

        返回:
            List[str]: 软件标识符列表，按优先级排序
        """
        default_priority = self.get_default_priority_config()
        priority = default_priority.get("spreadsheet_processors", ["wps_spreadsheets", "msoffice_excel", "libreoffice"])
        safe_log.debug("表格处理软件优先级: %s", priority)
        return priority

    def get_special_conversion_priority(self, conversion_type: str) -> list[str]:
        """
        获取特定特殊转换的软件优先级列表

        参数:
            conversion_type: 转换类型标识符
                - "odt": ODT文件转换
                - "ods": ODS文件转换
                - "pdf_to_office": PDF转Office格式
                - "document_to_pdf": 文档转PDF
                - "spreadsheet_to_pdf": 表格转PDF

        返回:
            List[str]: 软件标识符列表，按优先级排序
        """
        special_conversions = self.get_special_conversions_config()
        priority = special_conversions.get(conversion_type, [])
        safe_log.debug("特殊转换 %s 软件优先级: %s", conversion_type, priority)
        return priority

    def get_document_to_pdf_priority(self) -> list[str]:
        """
        获取文档转PDF的软件优先级列表

        返回:
            List[str]: 软件标识符列表，按优先级排序
        """
        special_conversions = self.get_special_conversions_config()
        priority = special_conversions.get("document_to_pdf", ["wps_writer", "msoffice_word", "libreoffice"])
        safe_log.debug("文档转PDF软件优先级: %s", priority)
        return priority

    def get_spreadsheet_to_pdf_priority(self) -> list[str]:
        """
        获取表格转PDF的软件优先级列表

        返回:
            List[str]: 软件标识符列表，按优先级排序
        """
        special_conversions = self.get_special_conversions_config()
        priority = special_conversions.get("spreadsheet_to_pdf", ["wps_spreadsheets", "msoffice_excel", "libreoffice"])
        safe_log.debug("表格转PDF软件优先级: %s", priority)
        return priority
