"""
标题序号配置模块

对应配置文件：
    - heading_numbering_add.toml: 序号添加方案
    - heading_numbering_clean.toml: 序号清理规则

包含：
    - DEFAULT_HEADING_NUMBERING_CONFIG: 默认序号添加配置
    - DEFAULT_NUMBERING_PATTERNS_CONFIG: 默认序号清理规则配置
    - NumberingConfigMixin: 标题序号配置获取方法
"""

from typing import Dict, Any, List
from ..safe_logger import safe_log

# ==============================================================================
#                              默认配置 - 序号添加
# ==============================================================================

DEFAULT_HEADING_NUMBERING_CONFIG = {
    "heading_numbering_add": {
        "settings": {
            "default_scheme": "gongwen_standard"
        },
        # 数字样式定义（供模板占位符使用）
        "number_styles": {
            "chinese_lower": {
                "name": "小写中文数字",
                "description": "一、二、三、四、五、六、七、八、九、十、十一、十二..."
            },
            "chinese_upper": {
                "name": "大写中文数字",
                "description": "壹、贰、叁、肆、伍、陆、柒、捌、玖、拾、拾壹、拾贰..."
            },
            "arabic_half": {
                "name": "半角阿拉伯数字",
                "description": "1, 2, 3, 4, 5, 6, 7, 8, 9, 10..."
            },
            "arabic_full": {
                "name": "全角阿拉伯数字",
                "description": "１、２、３、４、５、６、７、８、９、１０..."
            },
            "arabic_circled": {
                "name": "带圈阿拉伯数字",
                "description": "①②③④⑤⑥⑦⑧⑨⑩⑪⑫... (最多50)"
            },
            "letter_upper": {
                "name": "大写拉丁字母",
                "description": "A, B, C, D, E, F, G..."
            },
            "letter_lower": {
                "name": "小写拉丁字母",
                "description": "a, b, c, d, e, f, g..."
            },
            "roman_upper": {
                "name": "大写罗马数字",
                "description": "I, II, III, IV, V, VI, VII..."
            },
            "roman_lower": {
                "name": "小写罗马数字",
                "description": "i, ii, iii, iv, v, vi, vii..."
            }
        },
        "schemes": {
            # 方案1：公文标准
            "gongwen_standard": {
                "name": "公文标准",
                "description": "中国公文标准序号格式",
                "level_1": {"format": "{1.chinese_lower}、"},
                "level_2": {"format": "（{2.chinese_lower}）"},
                "level_3": {"format": "{3.arabic_half}. "},
                "level_4": {"format": "（{4.arabic_half}）"},
                "level_5": {"format": "{5.arabic_circled}"},
                "level_6": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half} "},
                "level_7": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half} "},
                "level_8": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half} "},
                "level_9": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half}.{9.arabic_half} "}
            },
            # 方案2：层级数字标准
            "hierarchical_standard": {
                "name": "层级数字标准",
                "description": "层级递进格式",
                "level_1": {"format": "{1.arabic_half} "},
                "level_2": {"format": "{1.arabic_half}.{2.arabic_half} "},
                "level_3": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half} "},
                "level_4": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half} "},
                "level_5": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half} "},
                "level_6": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half} "},
                "level_7": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half} "},
                "level_8": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half} "},
                "level_9": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half}.{9.arabic_half} "}
            },
            # 方案3：法律条文标准
            "legal_standard": {
                "name": "法律条文标准",
                "description": "中国法律文本格式",
                "level_1": {"format": "第{1.chinese_lower}编　"},
                "level_2": {"format": "第{2.chinese_lower}章　"},
                "level_3": {"format": "第{3.chinese_lower}节　"},
                "level_4": {"format": "第{4.chinese_lower}条　"},
                "level_5": {"format": "（{5.chinese_lower}）"},
                "level_6": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half} "},
                "level_7": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half} "},
                "level_8": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half} "},
                "level_9": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half}.{9.arabic_half} "}
            }
        }
    }
}

# ==============================================================================
#                              默认配置 - 序号清理规则
# ==============================================================================

# 占位符定义在 heading_utils.py 中（使用 raw string）
DEFAULT_NUMBERING_PATTERNS_CONFIG = {
    "heading_numbering_clean": {
        "settings": {
            "order": [
                "chinese_unit_suffix",
                "chinese_unit_prefix",
                "circled_numbers",
                "bracket_number",
                "hierarchical",
                "number_separator",
                "legal_english",
                "letter_number"
            ]
        },
        "rules": {
            "chinese_unit_suffix": {
                "name": "中文【第X单位】格式",
                "description": "匹配：第X回/篇/册/卷/部/集/期/编/章/节/条/款/项/目 + 空格",
                "enabled": True,
                "is_system": True,
                "regex": "^{space}*第{space}*{num}+{space}*[回篇册卷部集期编章节条款项目]{space}+"
            },
            "chinese_unit_prefix": {
                "name": "中文【单位X】格式",
                "description": "匹配：卷/篇/册/部/集/期/编/章/节/回 + 数字 + 空格",
                "enabled": True,
                "is_system": True,
                "regex": "^{space}*[卷篇册部集期编章节回]{space}*{num}+{space}+"
            },
            "circled_numbers": {
                "name": "带圈数字",
                "description": "匹配：①-㊿、❶-⓴、⓵-⓾、㈠-㈩ 等带圈/特殊数字 + 可选分隔符",
                "enabled": True,
                "is_system": True,
                "regex": "^{space}*{circled}{sep}?{space}*"
            },
            "bracket_number": {
                "name": "括号数字",
                "description": "匹配：(1)、（一）、1)、一）等 + 可选分隔符",
                "enabled": True,
                "is_system": True,
                "regex": "^{space}*{bracket_open}?{num}+{bracket_close}{sep}?{space}*"
            },
            "hierarchical": {
                "name": "层级数字",
                "description": "匹配：1.1 / 1.1.1 / 1.1.1.1 等 + 空格",
                "enabled": True,
                "is_system": True,
                "regex": "^{space}*{num_arab}+(?:\\.{num_arab}+)+{sep}?{space}+"
            },
            "number_separator": {
                "name": "数字分隔符",
                "description": "匹配：1. / 2、/ 一、/ 二，等 + 可选空格",
                "enabled": True,
                "is_system": True,
                "regex": "^{space}*{num}+{sep}{space}*"
            },
            "legal_english": {
                "name": "英文章节格式",
                "description": "匹配：Chapter/Part/Volume/Book/Section + 数字/罗马数字",
                "enabled": False,
                "is_system": False,
                "regex": "^{space}*(?:Chapter|Part|Volume|Book|Section)\\s+[0-9IVXLCDMivxlcdm]+[\\s\\.:\\-]+"
            },
            "letter_number": {
                "name": "字母序号",
                "description": "匹配：A. / A) / (A) / a. / a) 等",
                "enabled": False,
                "is_system": False,
                "regex": "^{space}*(?:{bracket_open}{letter}{bracket_close}|{letter}{bracket_close}|{letter}{sep}){space}*"
            }
        }
    }
}

# 配置文件名
CONFIG_FILES = {
    "heading_numbering_add": "heading_numbering_add.toml",
    "heading_numbering_clean": "heading_numbering_clean.toml"
}

# ==============================================================================
#                              Mixin 类
# ==============================================================================

class NumberingConfigMixin:
    """
    标题序号配置获取方法 Mixin
    
    提供序号添加方案和序号清理规则相关配置的访问方法。
    """
    
    # --------------------------------------------------------------------------
    # 第一层：配置块
    # --------------------------------------------------------------------------
    
    def get_heading_numbering_config_block(self) -> Dict[str, Any]:
        """
        获取整个标题序号配置块（序号添加方案）
        
        返回:
            Dict[str, Any]: 标题序号配置字典
        """
        return self._configs.get("heading_numbering_add", {})
    
    def get_numbering_patterns_block(self) -> Dict[str, Any]:
        """
        获取整个序号清理规则配置块
        
        返回:
            Dict[str, Any]: 序号清理规则配置字典
        """
        # 优先从配置文件加载，否则使用默认配置
        config = self._configs.get("heading_numbering_clean", {})
        if not config:
            config = DEFAULT_NUMBERING_PATTERNS_CONFIG.get("heading_numbering_clean", {})
        return config
    
    # --------------------------------------------------------------------------
    # 第二层：子表
    # --------------------------------------------------------------------------
    
    def get_numbering_settings(self) -> Dict[str, Any]:
        """
        获取序号添加设置子表
        
        返回:
            Dict[str, Any]: 序号添加设置字典
        """
        return self.get_heading_numbering_config_block().get("settings", {})
    
    def get_number_styles(self) -> Dict[str, Any]:
        """
        获取数字样式定义
        
        返回:
            Dict[str, Any]: 数字样式字典
        """
        return self.get_heading_numbering_config_block().get("number_styles", {})
    
    # --------------------------------------------------------------------------
    # 第三层：具体配置值 - 序号方案
    # --------------------------------------------------------------------------
    
    def get_heading_schemes(self) -> Dict[str, Dict]:
        """
        获取所有序号方案
        
        返回:
            Dict[str, Dict]: 方案字典，键为方案ID，值为方案配置
        """
        heading_config = self.get_heading_numbering_config_block()
        schemes = heading_config.get("schemes", {})
        safe_log.debug("获取序号方案列表，共 %d 个方案", len(schemes))
        return schemes
    
    def get_scheme_names(self, locale: str = None) -> List[str]:
        """
        获取当前语言下可用的序号方案名称列表（供GUI下拉框使用）
        
        参数:
            locale: 语言代码，如 "zh_CN"、"en_US"，默认使用当前语言
        
        返回:
            List[str]: 方案显示名称列表（已过滤不适用的方案）
        """
        localized_schemes = self.get_localized_numbering_schemes(locale)
        return list(localized_schemes.values())
    
    def get_localized_numbering_schemes(self, locale: str = None, include_description: bool = False) -> Dict[str, Any]:
        """
        获取当前语言下可用的序号方案字典（供GUI使用）
        
        参数:
            locale: 语言代码，如 "zh_CN"、"en_US"，默认使用当前语言
            include_description: 是否包含描述信息，默认 False
        
        返回:
            Dict[str, Any]: 方案字典
                - include_description=False: 键为方案ID，值为显示名称（字符串）
                - include_description=True: 键为方案ID，值为 {"name": str, "description": str}
        """
        from docwen.i18n import t, get_current_locale
        
        if locale is None:
            locale = get_current_locale()
        
        schemes = self.get_heading_schemes()
        # 获取方案顺序
        settings = self.get_numbering_settings()
        order = settings.get("order", list(schemes.keys()))
        
        result = {}
        
        # 按顺序遍历方案
        for scheme_id in order:
            scheme_config = schemes.get(scheme_id)
            if not isinstance(scheme_config, dict):
                continue
            
            # 检查 locales 过滤
            locales = scheme_config.get("locales", ["*"])
            if "*" not in locales and locale not in locales:
                safe_log.debug("跳过方案 %s (locales=%s, 当前=%s)", scheme_id, locales, locale)
                continue
            
            # 获取名称（支持 name_key 国际化）
            if "name_key" in scheme_config:
                name_key = scheme_config["name_key"]
                name = t(f"editors.numbering_add.names.{name_key}")
                # 如果翻译失败，回退到 name 或 ID
                if name.startswith("[") and name.endswith("]"):
                    name = scheme_config.get("name", scheme_id)
            else:
                name = scheme_config.get("name", scheme_id)
            
            if include_description:
                # 获取描述（支持 description_key 国际化）
                if "description_key" in scheme_config:
                    desc_key = scheme_config["description_key"]
                    description = t(f"editors.numbering_add.descriptions.{desc_key}")
                    # 如果翻译失败，回退到 description 或空字符串
                    if description.startswith("[") and description.endswith("]"):
                        description = scheme_config.get("description", "")
                else:
                    description = scheme_config.get("description", "")
                
                result[scheme_id] = {"name": name, "description": description}
            else:
                result[scheme_id] = name
        
        safe_log.debug("获取序号方案字典 (locale=%s): %s", locale, result)
        return result
    
    def get_default_numbering_scheme(self) -> str:
        """
        获取默认序号方案ID
        
        返回:
            str: 默认方案ID
        """
        settings = self.get_numbering_settings()
        scheme = settings.get("default_scheme", "gongwen_standard")
        safe_log.debug("获取默认序号方案: %s", scheme)
        return scheme
    
    def get_scheme_by_id(self, scheme_id: str) -> Dict[str, Any]:
        """
        根据方案ID获取方案配置
        
        参数:
            scheme_id: 方案ID
            
        返回:
            Dict[str, Any]: 方案配置字典，未找到返回空字典
        """
        schemes = self.get_heading_schemes()
        scheme = schemes.get(scheme_id, {})
        if not scheme:
            safe_log.warning("未找到序号方案: %s", scheme_id)
        return scheme
    
    def get_scheme_by_name(self, scheme_name: str) -> Dict[str, Any]:
        """
        根据方案名称获取方案配置
        
        参数:
            scheme_name: 方案显示名称
            
        返回:
            Dict[str, Any]: 方案配置字典，未找到返回空字典
        """
        schemes = self.get_heading_schemes()
        for scheme_id, scheme_config in schemes.items():
            if isinstance(scheme_config, dict):
                if scheme_config.get("name") == scheme_name:
                    return scheme_config
        safe_log.warning("未找到名称为 '%s' 的序号方案", scheme_name)
        return {}
    
    # --------------------------------------------------------------------------
    # 第三层：具体配置值 - 序号清理规则
    # --------------------------------------------------------------------------
    
    def get_cleaning_rules(self, localized: bool = False) -> List[Dict[str, Any]]:
        """
        获取启用的清理规则列表（按优先级排序，已替换占位符）
        
        参数:
            localized: 是否返回国际化的名称和描述，默认 False
        
        返回:
            List[Dict]: 规则列表，每个规则包含:
                - rule_id: str - 规则ID
                - name: str - 规则名称
                - description: str - 规则说明
                - enabled: bool - 是否启用
                - is_system: bool - 是否系统规则
                - regex: str - 原始正则表达式
                - compiled_regex: str - 替换占位符后的正则表达式
        """
        # 从 heading_utils.py 导入占位符定义（硬编码，使用 raw string）
        from docwen.utils.heading_utils import NUMBERING_PLACEHOLDERS
        
        if localized:
            from docwen.i18n import t
        
        config = self.get_numbering_patterns_block()
        placeholders = NUMBERING_PLACEHOLDERS  # 使用代码中的占位符
        settings = config.get("settings", {})
        rules_dict = config.get("rules", {})
        order = settings.get("order", list(rules_dict.keys()))
        
        result = []
        for rule_id in order:
            if rule_id not in rules_dict:
                continue
            
            rule = rules_dict[rule_id]
            if not rule.get("enabled", True):
                continue
            
            # 替换占位符
            regex = rule.get("regex", "")
            compiled_regex = self._replace_placeholders(regex, placeholders)
            
            # 获取名称（支持国际化）
            if localized and "name_key" in rule:
                name_key = rule["name_key"]
                name = t(f"editors.numbering_clean.names.{name_key}")
                if name.startswith("[") and name.endswith("]"):
                    name = rule.get("name", rule_id)
            else:
                name = rule.get("name", rule_id)
            
            # 获取描述（支持国际化）
            if localized and "description_key" in rule:
                desc_key = rule["description_key"]
                description = t(f"editors.numbering_clean.descriptions.{desc_key}")
                if description.startswith("[") and description.endswith("]"):
                    description = rule.get("description", "")
            else:
                description = rule.get("description", "")
            
            result.append({
                "rule_id": rule_id,
                "name": name,
                "description": description,
                "enabled": rule.get("enabled", True),
                "is_system": rule.get("is_system", False),
                "regex": regex,
                "compiled_regex": compiled_regex
            })
        
        safe_log.debug("获取清理规则列表，共 %d 条启用的规则", len(result))
        return result
    
    def _replace_placeholders(self, regex: str, placeholders: Dict[str, str]) -> str:
        """
        替换正则表达式中的占位符
        
        参数:
            regex: 原始正则表达式（含 {placeholder} 格式的占位符）
            placeholders: 占位符字典
            
        返回:
            str: 替换后的正则表达式
        """
        result = regex
        for name, value in placeholders.items():
            result = result.replace(f"{{{name}}}", value)
        return result
    
    def get_all_cleaning_rules(self) -> Dict[str, Dict[str, Any]]:
        """
        获取所有清理规则（包括禁用的）
        
        返回:
            Dict[str, Dict]: 规则字典，键为规则ID
        """
        config = self.get_numbering_patterns_block()
        rules = config.get("rules", {})
        safe_log.debug("获取所有清理规则，共 %d 条", len(rules))
        return rules
    
    def get_cleaning_rule_order(self) -> List[str]:
        """
        获取清理规则执行顺序
        
        返回:
            List[str]: 规则ID列表，按执行顺序排列
        """
        config = self.get_numbering_patterns_block()
        settings = config.get("settings", {})
        order = settings.get("order", [])
        safe_log.debug("获取清理规则执行顺序: %s", order)
        return order
