"""
国际化管理器模块

提供语言加载、翻译查找和语言切换功能。
采用单例模式，确保全局只有一个实例。
"""

import logging
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


class I18nManager:
    """
    国际化管理器类

    负责加载和管理多语言翻译文本。
    使用单例模式，全局共享一个实例。

    支持的语言文件格式：TOML
    语言文件位置：src/docwen/i18n/locales/

    使用示例：
        from docwen.i18n import t

        # 简单翻译
        text = t("common.close")  # 返回 "关闭"

        # 带参数的翻译
        text = t("common.version", version="1.0.0")  # 返回 "版本 1.0.0"
    """

    _instance = None
    _initialized = False

    # 支持的语言列表
    AVAILABLE_LOCALES: list[dict[str, Any]]

    # 默认语言
    DEFAULT_LOCALE = "zh_CN"

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        """初始化国际化管理器"""
        if self._initialized:
            return

        self._locale = self.DEFAULT_LOCALE
        self._translations: dict[str, Any] = {}
        self._locales_dir = self._get_locales_dir()
        self._available_locales_scanned = False

        # 加载当前语言
        self._load_locale_from_config()
        self._load_translations()

        self._initialized = True
        logger.info("I18nManager 初始化完成，当前语言: %s", self._locale)

    def _ensure_available_locales_scanned(self):
        if self._available_locales_scanned:
            return
        self._scan_available_locales()

    def _get_locales_dir(self) -> str:
        """获取语言文件目录路径"""
        return str(Path(__file__).resolve().parent / "locales")

    def _scan_available_locales(self):
        """扫描语言目录，动态加载可用语言列表"""
        self.AVAILABLE_LOCALES = []

        # 必须确保默认语言存在，即使文件扫描失败
        default_locale_info = {"code": self.DEFAULT_LOCALE, "name": "Simplified Chinese", "native_name": "简体中文"}

        try:
            import tomllib

            toml_files = list(Path(self._locales_dir).glob("*.toml"))

            for file_path in toml_files:
                try:
                    # 获取文件名作为语言代码
                    filename = file_path.name
                    code = file_path.stem

                    meta_lines: list[str] = []
                    in_meta = False
                    with file_path.open(encoding="utf-8") as f:
                        for line in f:
                            stripped = line.strip()
                            if stripped.startswith("[") and stripped.endswith("]"):
                                if stripped == "[meta]":
                                    in_meta = True
                                    meta_lines.append(line)
                                    continue
                                if in_meta:
                                    break
                            if in_meta:
                                meta_lines.append(line)

                    if not meta_lines:
                        logger.warning("语言文件 %s 缺少 [meta] 元数据，跳过", filename)
                        continue

                    data = tomllib.loads("".join(meta_lines))
                    meta = data.get("meta")
                    if not isinstance(meta, dict):
                        logger.warning("语言文件 %s 的 [meta] 元数据无效，跳过", filename)
                        continue

                    name = meta.get("name", code)
                    native_name = meta.get("native_name", name)
                    self.AVAILABLE_LOCALES.append({"code": code, "name": name, "native_name": native_name})
                except Exception as e:
                    logger.error("扫描语言文件 %s 失败: %s", str(file_path), str(e))

            # 如果没有扫描到任何语言，添加默认语言作为兜底
            if not self.AVAILABLE_LOCALES:
                logger.warning("未扫描到任何有效语言文件，使用默认列表")
                self.AVAILABLE_LOCALES.append(default_locale_info)
            else:
                # 按语言代码排序，确保顺序稳定
                self.AVAILABLE_LOCALES.sort(key=lambda x: x["code"])

            logger.info(
                "已加载 %d 种语言: %s",
                len(self.AVAILABLE_LOCALES),
                ", ".join([loc["code"] for loc in self.AVAILABLE_LOCALES]),
            )

        except Exception as e:
            logger.error("扫描语言目录失败: %s", str(e))
            self.AVAILABLE_LOCALES = [default_locale_info]
        finally:
            self._available_locales_scanned = True

    def _load_locale_from_config(self):
        """从配置文件加载语言设置"""
        try:
            from docwen.config.config_manager import config_manager

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
        if not locale:
            return False
        return (Path(self._locales_dir) / f"{locale}.toml").exists()

    def _load_translations(self):
        """加载当前语言的翻译文件"""
        locale_file = Path(self._locales_dir) / f"{self._locale}.toml"

        if not locale_file.exists():
            logger.warning("语言文件不存在: %s，尝试加载默认语言", locale_file)
            # 尝试加载默认语言
            if self._locale != self.DEFAULT_LOCALE:
                locale_file = Path(self._locales_dir) / f"{self.DEFAULT_LOCALE}.toml"
                self._locale = self.DEFAULT_LOCALE

        if not locale_file.exists():
            logger.error("默认语言文件也不存在: %s", locale_file)
            self._translations = {}
            return

        try:
            from docwen.config.toml_operations import read_toml_file

            self._translations = read_toml_file(str(locale_file)) or {}
            logger.info("成功加载语言文件: %s，共 %d 个顶级键", locale_file, len(self._translations))
        except Exception as e:
            logger.error("加载语言文件失败: %s | 错误: %s", locale_file, str(e))
            self._translations = {}

    def t(self, key: str, default: str | None = None, **kwargs) -> str:
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
            t("<missing.key>", default="默认值")  # 返回 "默认值"
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

    def _get_nested_value(self, key: str) -> str | None:
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
        self._ensure_available_locales_scanned()
        for locale in self.AVAILABLE_LOCALES:
            if locale["code"] == self._locale:
                return locale["native_name"]
        return self._locale

    def get_available_locales(self) -> list[dict[str, str]]:
        """
        获取可用语言列表

        返回：
            List[Dict[str, str]]: 语言列表，每项包含：
                - code: 语言代码（如 "zh_CN"）
                - name: 语言名称（如 "简体中文"）
                - native_name: 本地化名称
        """
        self._ensure_available_locales_scanned()
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
            from docwen.config.config_manager import config_manager

            success = config_manager.update_config_value("gui_config", "language", "locale", locale)
            if success:
                logger.info("语言设置已保存: %s，将在下次启动时生效", locale)
            return success
        except Exception as e:
            logger.error("保存语言设置失败: %s", str(e))
            return False

    def apply_locale(self, locale: str) -> bool:
        if not self._is_valid_locale(locale):
            logger.error("无效的语言代码: %s", locale)
            return False

        self._locale = locale
        self._load_translations()
        logger.info("语言已应用: %s", locale)
        return True

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

    def t_all_locales(self, key: str) -> dict[str, str]:
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
        self._ensure_available_locales_scanned()
        for locale_info in self.AVAILABLE_LOCALES:
            locale_code = locale_info["code"]
            translations = self._load_locale_translations(locale_code)

            if translations:
                value = self._get_nested_value_from_dict(key, translations)
                if value is not None:
                    result[locale_code] = value

        logger.debug("所有语言版本翻译结果: %s -> %s", key, result)
        return result

    def get_detection_priority(self) -> list[str]:
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
        self._ensure_available_locales_scanned()
        for locale_info in self.AVAILABLE_LOCALES:
            locale_code = locale_info["code"]
            if locale_code not in priority:
                priority.append(locale_code)

        logger.debug("检测优先级: %s", priority)
        return priority

    def _load_locale_translations(self, locale: str) -> dict[str, Any] | None:
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
        if not hasattr(self, "_locale_cache"):
            self._locale_cache: dict[str, dict[str, Any]] = {}

        if locale in self._locale_cache:
            return self._locale_cache[locale]

        # 加载翻译文件
        locale_file = Path(self._locales_dir) / f"{locale}.toml"

        if not locale_file.exists():
            logger.warning("语言文件不存在: %s", str(locale_file))
            return None

        try:
            from docwen.config.toml_operations import read_toml_file

            translations = read_toml_file(str(locale_file)) or {}
            self._locale_cache[locale] = translations
            logger.debug("已加载语言文件到缓存: %s", locale)
            return translations
        except Exception as e:
            logger.error("加载语言文件失败: %s | 错误: %s", str(locale_file), str(e))
            return None

    def _get_nested_value_from_dict(self, key: str, data: dict[str, Any]) -> str | None:
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

    def get_localized_options(self, section_key: str) -> dict[str, str]:
        """
        获取适用于当前语言的选项列表

        根据 _locales 元数据过滤选项，只返回适用于当前语言的选项。
        _locales 元数据格式：
        - ["zh_CN"]: 只在中文界面显示
        - ["*"]: 所有语言通用
        - ["zh_CN", "en_US"]: 在中文和英文界面显示

        参数：
            section_key: 节键，如 "editors.numbering_add.names"

        返回：
            Dict[str, str]: 过滤后的选项字典，键为选项ID，值为翻译后的显示名称
            如 {"hierarchical_standard": "层级数字标准"}

        示例：
            # 中文界面
            get_localized_options("editors.numbering_add.names")
            # 返回 {"gongwen_standard": "公文标准", "hierarchical_standard": "层级数字标准", "legal_standard": "法律条文标准"}

            # 英文界面
            get_localized_options("editors.numbering_add.names")
            # 返回 {"hierarchical_standard": "Hierarchical Number Standard"}
        """
        current_locale = self._locale
        result = {}

        # 获取选项节
        options_data = self._get_section_data(section_key)
        if not options_data:
            logger.warning("选项节不存在: %s", section_key)
            return result

        # 获取语言标记元数据节
        locales_key = f"{section_key}._locales"
        locales_data = self._get_section_data(locales_key)

        # 遍历选项，检查语言标记
        for option_id, option_name in options_data.items():
            # 跳过元数据节和非字符串值
            if option_id.startswith("_") or not isinstance(option_name, str):
                continue

            # 获取该选项的语言标记
            if locales_data and option_id in locales_data:
                allowed_locales = locales_data[option_id]
                if isinstance(allowed_locales, list):
                    # 检查是否适用于当前语言
                    if "*" in allowed_locales or current_locale in allowed_locales:
                        result[option_id] = option_name
                        logger.debug("选项 %s 适用于当前语言 %s", option_id, current_locale)
                    else:
                        logger.debug(
                            "选项 %s 不适用于当前语言 %s（允许: %s）", option_id, current_locale, allowed_locales
                        )
            else:
                # 没有语言标记的选项默认显示（向后兼容）
                result[option_id] = option_name
                logger.debug("选项 %s 无语言标记，默认显示", option_id)

        logger.info("获取本地化选项: %s -> %d 个选项（当前语言: %s）", section_key, len(result), current_locale)
        return result

    def _get_section_data(self, section_key: str) -> dict[str, Any] | None:
        """
        获取指定节的数据

        参数：
            section_key: 节键，如 "editors.numbering_add.names"

        返回：
            Optional[Dict[str, Any]]: 节数据字典，或 None
        """
        keys = section_key.split(".")
        value = self._translations

        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return None

        return value if isinstance(value, dict) else None

    def has_localized_options(self, section_key: str) -> bool:
        """
        检查当前语言是否有可用的选项

        用于决定是否显示某个选项区域（如优化类型选择）。
        如果当前语言没有任何可用选项，可以隐藏整个区域。

        参数：
            section_key: 节键，如 "action_panel.optimization_types"

        返回：
            bool: 是否有可用选项
        """
        options = self.get_localized_options(section_key)
        return len(options) > 0

    def get_style_format(self, style_key: str) -> dict[str, Any] | None:
        """
        获取样式格式配置

        从当前语言的 [style_formats] 节中读取样式格式配置。
        用于样式注入时设置字体、字号、缩进等格式属性。

        参数：
            style_key: 样式键名，如 "body_paragraph"、"heading_1"、"heading_3_9"

        返回：
            Optional[Dict[str, Any]]: 格式配置字典，包含：
                - east_asia_font: 东亚字体（如 "仿宋_GB2312"）
                - ascii_font: 西文字体（如 "Times New Roman"）
                - font_size_pt: 字号（磅值，如 16）
                - first_line_indent_chars: 首行缩进（1/100字符，如 200 表示 2 字符）
                - first_line_indent_cm: 首行缩进（厘米，用于俄语等）
                - spacing_after_twip: 段后间距（twip，1pt=20twip）
                - spacing_before_twip: 段前间距（twip）
                - bold: 是否加粗
                - justification: 对齐方式（"both" 或 "left"）
            如果找不到返回 None

        示例：
            format_config = i18n.get_style_format("heading_1")
            # 返回 {"east_asia_font": "黑体", "ascii_font": "Times New Roman", ...}
        """
        logger.debug("获取样式格式配置: style_key=%s", style_key)

        # 从 style_formats 节获取配置
        format_data = self._get_section_data(f"style_formats.{style_key}")

        if format_data is None:
            logger.debug("样式格式配置不存在: style_formats.%s", style_key)
            return None

        logger.debug("样式格式配置: %s -> %s", style_key, format_data)
        return format_data
