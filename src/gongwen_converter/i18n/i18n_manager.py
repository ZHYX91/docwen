"""
国际化管理器模块

提供语言加载、翻译查找和语言切换功能。
采用单例模式，确保全局只有一个实例。
"""

import os
import logging
from typing import Dict, Any, List, Optional

logger = logging.getLogger(__name__)


class I18nManager:
    """
    国际化管理器类
    
    负责加载和管理多语言翻译文本。
    使用单例模式，全局共享一个实例。
    
    支持的语言文件格式：TOML
    语言文件位置：src/gongwen_converter/i18n/locales/
    
    使用示例：
        from gongwen_converter.i18n import t
        
        # 简单翻译
        text = t("common.close")  # 返回 "关闭"
        
        # 带参数的翻译
        text = t("common.version", version="1.0.0")  # 返回 "版本 1.0.0"
    """
    
    _instance = None
    _initialized = False
    
    # 支持的语言列表
    AVAILABLE_LOCALES = [
        {"code": "zh_CN", "name": "简体中文", "native_name": "简体中文"},
        {"code": "en_US", "name": "English (US)", "native_name": "English (US)"},
    ]
    
    # 默认语言
    DEFAULT_LOCALE = "zh_CN"
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(I18nManager, cls).__new__(cls)
        return cls._instance
    
    def __init__(self):
        """初始化国际化管理器"""
        if self._initialized:
            return
        
        self._locale = self.DEFAULT_LOCALE
        self._translations: Dict[str, Any] = {}
        self._locales_dir = self._get_locales_dir()
        
        # 加载当前语言
        self._load_locale_from_config()
        self._load_translations()
        
        self._initialized = True
        logger.info("I18nManager 初始化完成，当前语言: %s", self._locale)
    
    def _get_locales_dir(self) -> str:
        """获取语言文件目录路径"""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        locales_dir = os.path.join(current_dir, "locales")
        return locales_dir
    
    def _load_locale_from_config(self):
        """从配置文件加载语言设置"""
        try:
            from gongwen_converter.config.config_manager import config_manager
            locale = config_manager.get_locale()
            if locale and self._is_valid_locale(locale):
                self._locale = locale
                logger.debug("从配置加载语言设置: %s", locale)
            else:
                logger.debug("配置中无有效语言设置，使用默认值: %s", self.DEFAULT_LOCALE)
        except Exception as e:
            logger.warning("加载语言配置失败，使用默认值: %s", str(e))
    
    def _is_valid_locale(self, locale: str) -> bool:
        """检查语言代码是否有效"""
        return any(loc["code"] == locale for loc in self.AVAILABLE_LOCALES)
    
    def _load_translations(self):
        """加载当前语言的翻译文件"""
        locale_file = os.path.join(self._locales_dir, f"{self._locale}.toml")
        
        if not os.path.exists(locale_file):
            logger.warning("语言文件不存在: %s，尝试加载默认语言", locale_file)
            # 尝试加载默认语言
            if self._locale != self.DEFAULT_LOCALE:
                locale_file = os.path.join(self._locales_dir, f"{self.DEFAULT_LOCALE}.toml")
                self._locale = self.DEFAULT_LOCALE
        
        if not os.path.exists(locale_file):
            logger.error("默认语言文件也不存在: %s", locale_file)
            self._translations = {}
            return
        
        try:
            from gongwen_converter.config.toml_operations import read_toml_file
            self._translations = read_toml_file(locale_file) or {}
            logger.info("成功加载语言文件: %s，共 %d 个顶级键", 
                       locale_file, len(self._translations))
        except Exception as e:
            logger.error("加载语言文件失败: %s | 错误: %s", locale_file, str(e))
            self._translations = {}
    
    def t(self, key: str, default: Optional[str] = None, **kwargs) -> str:
        """
        翻译函数
        
        根据键名查找翻译文本，支持嵌套键和参数替换。
        
        参数：
            key: 翻译键，支持点号分隔的嵌套键（如 "common.close"）
            default: 默认值，当找不到翻译时返回
            **kwargs: 用于格式化字符串的参数
            
        返回：
            str: 翻译后的文本
            
        示例：
            t("common.close")  # 返回 "关闭"
            t("common.version", version="1.0.0")  # 返回 "版本 1.0.0"
            t("not.exist", default="默认值")  # 返回 "默认值"
        """
        # 查找翻译
        value = self._get_nested_value(key)
        
        if value is None:
            if default is not None:
                value = default
            else:
                logger.warning("翻译键不存在: %s", key)
                return f"[{key}]"  # 返回键名，方便调试
        
        # 参数替换
        if kwargs:
            try:
                value = value.format(**kwargs)
            except KeyError as e:
                logger.warning("翻译参数缺失: %s | 键: %s", str(e), key)
            except Exception as e:
                logger.warning("翻译格式化失败: %s | 键: %s", str(e), key)
        
        return value
    
    def _get_nested_value(self, key: str) -> Optional[str]:
        """
        获取嵌套字典中的值
        
        参数：
            key: 点号分隔的键路径，如 "common.close"
            
        返回：
            Optional[str]: 找到的值，或 None
        """
        keys = key.split(".")
        value = self._translations
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return None
        
        return value if isinstance(value, str) else None
    
    def get_current_locale(self) -> str:
        """获取当前语言代码"""
        return self._locale
    
    def get_current_locale_name(self) -> str:
        """获取当前语言的显示名称"""
        for locale in self.AVAILABLE_LOCALES:
            if locale["code"] == self._locale:
                return locale["native_name"]
        return self._locale
    
    def get_available_locales(self) -> List[Dict[str, str]]:
        """
        获取可用语言列表
        
        返回：
            List[Dict[str, str]]: 语言列表，每项包含：
                - code: 语言代码（如 "zh_CN"）
                - name: 语言名称（如 "简体中文"）
                - native_name: 本地化名称
        """
        return self.AVAILABLE_LOCALES.copy()
    
    def set_locale(self, locale: str) -> bool:
        """
        设置语言（保存到配置，下次启动生效）
        
        参数：
            locale: 语言代码（如 "zh_CN" 或 "en_US"）
            
        返回：
            bool: 是否设置成功
        """
        if not self._is_valid_locale(locale):
            logger.error("无效的语言代码: %s", locale)
            return False
        
        try:
            from gongwen_converter.config.config_manager import config_manager
            success = config_manager.update_config_value(
                "gui_config", "language", "locale", locale
            )
            if success:
                logger.info("语言设置已保存: %s，将在下次启动时生效", locale)
            return success
        except Exception as e:
            logger.error("保存语言设置失败: %s", str(e))
            return False
    
    def reload_translations(self):
        """重新加载翻译文件"""
        self._load_locale_from_config()
        self._load_translations()
        logger.info("翻译文件已重新加载，当前语言: %s", self._locale)
    
    def t_locale(self, key: str, locale: str, **kwargs) -> str:
        """
        获取指定语言的翻译（不改变当前语言）
        
        用于获取其他语言版本的样式名/占位符，
        而不影响当前语言设置。
        
        参数：
            key: 翻译键，支持点号分隔的嵌套键（如 "styles.code_block"）
            locale: 目标语言代码（如 "zh_CN" 或 "en_US"）
            **kwargs: 用于格式化字符串的参数
            
        返回：
            str: 指定语言的翻译文本，如果找不到返回 [key] 格式
            
        示例：
            t_locale("styles.code_block", "en_US")  # 返回 "Code Block"
            t_locale("styles.code_block", "zh_CN")  # 返回 "代码块"
        """
        logger.debug("获取指定语言翻译: key=%s, locale=%s", key, locale)
        
        # 加载指定语言的翻译文件
        translations = self._load_locale_translations(locale)
        if not translations:
            logger.warning("无法加载语言 %s 的翻译文件", locale)
            return f"[{key}]"
        
        # 查找翻译
        value = self._get_nested_value_from_dict(key, translations)
        
        if value is None:
            logger.warning("翻译键 %s 在语言 %s 中不存在", key, locale)
            return f"[{key}]"
        
        # 参数替换
        if kwargs:
            try:
                value = value.format(**kwargs)
            except KeyError as e:
                logger.warning("翻译参数缺失: %s | 键: %s | 语言: %s", str(e), key, locale)
            except Exception as e:
                logger.warning("翻译格式化失败: %s | 键: %s | 语言: %s", str(e), key, locale)
        
        return value
    
    def t_all_locales(self, key: str) -> Dict[str, str]:
        """
        获取所有语言版本的翻译
        
        用于获取某个翻译键在所有支持语言中的翻译文本。
        主要用于样式名检测，需要匹配所有可能的语言版本。
        
        参数：
            key: 翻译键，支持点号分隔的嵌套键
            
        返回：
            Dict[str, str]: 字典，键为语言代码，值为翻译文本
            如 {"zh_CN": "代码块", "en_US": "Code Block"}
            如果某语言没有该翻译，则不包含在结果中
            
        示例：
            t_all_locales("styles.code_block")
            # 返回 {"zh_CN": "代码块", "en_US": "Code Block"}
        """
        logger.debug("获取所有语言版本翻译: key=%s", key)
        
        result = {}
        for locale_info in self.AVAILABLE_LOCALES:
            locale_code = locale_info["code"]
            translations = self._load_locale_translations(locale_code)
            
            if translations:
                value = self._get_nested_value_from_dict(key, translations)
                if value is not None:
                    result[locale_code] = value
        
        logger.debug("所有语言版本翻译结果: %s -> %s", key, result)
        return result
    
    def get_detection_priority(self) -> List[str]:
        """
        获取样式检测时的语言优先级
        
        返回语言代码列表，当前语言自动置顶。
        用于确定样式检测时的匹配优先级，优先使用当前语言的样式。
        
        返回：
            List[str]: 语言代码列表（去重后）
            如 ["zh_CN", "en_US"]（当前语言为中文时）
            如 ["en_US", "zh_CN"]（当前语言为英文时）
            
        示例：
            # 当前语言为 zh_CN 时
            get_detection_priority()  # 返回 ["zh_CN", "en_US"]
        """
        # 将当前语言置顶
        priority = [self._locale]
        
        # 添加其他语言
        for locale_info in self.AVAILABLE_LOCALES:
            locale_code = locale_info["code"]
            if locale_code not in priority:
                priority.append(locale_code)
        
        logger.debug("检测优先级: %s", priority)
        return priority
    
    def _load_locale_translations(self, locale: str) -> Optional[Dict[str, Any]]:
        """
        加载指定语言的翻译文件（带缓存）
        
        参数：
            locale: 语言代码
            
        返回：
            Optional[Dict[str, Any]]: 翻译字典，加载失败返回 None
        """
        # 如果是当前语言，直接返回已加载的翻译
        if locale == self._locale:
            return self._translations
        
        # 检查缓存
        if not hasattr(self, '_locale_cache'):
            self._locale_cache: Dict[str, Dict[str, Any]] = {}
        
        if locale in self._locale_cache:
            return self._locale_cache[locale]
        
        # 加载翻译文件
        locale_file = os.path.join(self._locales_dir, f"{locale}.toml")
        
        if not os.path.exists(locale_file):
            logger.warning("语言文件不存在: %s", locale_file)
            return None
        
        try:
            from gongwen_converter.config.toml_operations import read_toml_file
            translations = read_toml_file(locale_file) or {}
            self._locale_cache[locale] = translations
            logger.debug("已加载语言文件到缓存: %s", locale)
            return translations
        except Exception as e:
            logger.error("加载语言文件失败: %s | 错误: %s", locale_file, str(e))
            return None
    
    def _get_nested_value_from_dict(self, key: str, data: Dict[str, Any]) -> Optional[str]:
        """
        从字典中获取嵌套键的值
        
        参数：
            key: 点号分隔的键路径
            data: 数据字典
            
        返回：
            Optional[str]: 找到的字符串值，或 None
        """
        keys = key.split(".")
        value = data
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return None
        
        return value if isinstance(value, str) else None
