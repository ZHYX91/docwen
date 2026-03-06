"""
docx_spell 文本验证器模块
实现错别字、标点符号、敏感词等校对功能
支持通过构造参数覆盖配置设置
"""

import logging
import re
from typing import NamedTuple

from docwen.config.config_manager import config_manager
from docwen.translation import t

# 配置日志
logger = logging.getLogger("DocxSpellValidator")


# 定义错误类型
class TextError(NamedTuple):
    """
    文本错误信息容器
    Attributes:
        start_pos: 错误开始位置
        end_pos: 错误结束位置
        error_text: 错误文本
        suggestion: 修改建议
        error_type: 错误类型
        source: 错误来源 (custom/symbol)
    """

    start_pos: int
    end_pos: int
    error_text: str
    suggestion: str
    error_type: str
    source: str


class TextValidator:
    """
    文本验证器核心类
    实现多种校对方式：
    1. 标点配对
    2. 符号校对
    3. 错别字校对
    4. 敏感词匹配
    支持通过构造参数覆盖配置设置
    """

    def __init__(self, symbol_pairing=None, symbol_correction=None, typos_rule=None, sensitive_word=None):
        """
        初始化文本验证器（支持配置覆盖）
        采用延迟加载策略，避免不必要的资源消耗

        参数:
            symbol_pairing: 是否启用标点配对（None=使用配置）
            symbol_correction: 是否启用符号校对（None=使用配置）
            typos_rule: 是否启用错别字校对（None=使用配置）
            sensitive_word: 是否启用敏感词匹配（None=使用配置）
        """
        logger.info("初始化文本验证器（支持配置覆盖）")

        # 存储构造函数传入的覆盖配置
        self.overrides = {
            "enable_symbol_pairing": symbol_pairing,
            "enable_symbol_correction": symbol_correction,
            "enable_typos_rule": typos_rule,
            "enable_sensitive_word": sensitive_word,
        }

        logger.debug(f"接收覆盖配置: {self.overrides}")

        # 初始化状态
        self._initialized = False
        self.symbol_pairing_enabled = False
        self.symbol_correction_enabled = False
        self.typos_rule_enabled = False
        self.sensitive_word_enabled = False
        self.symbol_pairs = []
        self.symbol_map = {}
        self.typos_map = {}
        self.sensitive_words_map = {}

        logger.debug("验证器已创建，等待首次调用时初始化")

    def _initialize(self):
        """内部初始化方法（延迟加载配置和资源）"""
        if self._initialized:
            return

        logger.info("开始初始化文本验证器...")

        # 统一配置和加载资源
        self._configure_engine()

        self._initialized = True
        logger.info("文本验证器初始化完成")

    def _configure_engine(self):
        """
        统一配置校对引擎，建立清晰的优先级：
        构造函数参数 > 配置文件 > `constants.py` 默认值
        """
        logger.debug("开始配置校对引擎...")

        try:
            # 1. 从 config_manager 获取校对引擎配置（集中配置）
            settings = config_manager.get_proofread_engine_config()
            logger.debug(f"基础配置 (默认+文件): {settings}")

            # 2. 应用构造函数传入的覆盖值
            for key, value in self.overrides.items():
                if value is not None:
                    settings[key] = value
                    logger.info(f"应用构造函数覆盖: {key} = {value}")

            # 3. 设置最终的引擎状态
            self.symbol_pairing_enabled = settings.get("enable_symbol_pairing")
            self.symbol_correction_enabled = settings.get("enable_symbol_correction")
            self.typos_rule_enabled = settings.get("enable_typos_rule")
            self.sensitive_word_enabled = settings.get("enable_sensitive_word")

            logger.info(
                f"最终引擎配置: "
                f"标点配对={self.symbol_pairing_enabled}, "
                f"标点校对={self.symbol_correction_enabled}, "
                f"错别字校对={self.typos_rule_enabled}, "
                f"敏感词匹配={self.sensitive_word_enabled}"
            )

            # 4. 加载校对所需资源
            self.symbol_pairs = config_manager.get_symbol_pairs()
            self.symbol_map = config_manager.get_symbol_map()
            self.typos_map = config_manager.get_typos()
            self.sensitive_words_map = config_manager.get_sensitive_words()
            logger.debug("校对资源加载完成")

            # 5. 验证敏感词配置
            if self.sensitive_word_enabled and self.sensitive_words_map:
                logger.info("开始验证敏感词配置...")
                invalid_count = 0

                for sensitive_word, exceptions in self.sensitive_words_map.items():
                    if not exceptions or not isinstance(exceptions, list):
                        continue

                    for exception in exceptions:
                        if not exception:
                            continue

                        # 检查例外词是否包含敏感词
                        if sensitive_word not in exception:
                            logger.warning(
                                f"配置警告：例外词 '{exception}' 不包含敏感词 '{sensitive_word}'，"
                                f"这可能导致意外的匹配行为，建议修正配置"
                            )
                            invalid_count += 1

                if invalid_count > 0:
                    logger.warning(f"发现 {invalid_count} 个配置警告，建议检查 proofread_sensitive.toml")
                else:
                    logger.info("敏感词配置验证通过")

        except Exception as e:
            logger.error(f"引擎配置失败: {e!s}", exc_info=True)
            # 配置失败时，禁用所有功能以保证安全
            self.symbol_pairing_enabled = False
            self.symbol_correction_enabled = False
            self.typos_rule_enabled = False
            self.sensitive_word_enabled = False
            self.symbol_pairs = []
            self.symbol_map = {}
            self.typos_map = {}
            self.sensitive_words_map = {}
            logger.warning("所有校对功能已禁用")

    def validate_text(self, text: str) -> list[TextError]:
        """
        执行所有启用的校对并返回错误列表
        自动执行延迟初始化

        参数:
            text: 要校对的文本

        返回:
            List[TextError]: 错误列表
        """
        # 确保验证器已初始化
        if not self._initialized:
            self._initialize()

        logger.info(f"开始校对文本: {text[:50]}... (长度: {len(text)})")
        errors = []

        # 1. 执行错别字校对
        if self.typos_rule_enabled:
            logger.debug("执行错别字校对...")
            typo_errors = self.typos_check(text)
            errors.extend(typo_errors)
            logger.info(f"错别字校对完成，发现 {len(typo_errors)} 个错误")

        # 2. 执行符号校对（独立于错别字校对）
        if self.symbol_correction_enabled:
            logger.debug("执行符号校对...")
            symbol_correction_errors = self.symbol_correction_check(text)
            errors.extend(symbol_correction_errors)
            logger.info(f"符号校对完成，发现 {len(symbol_correction_errors)} 个错误")

        # 3. 执行符号配对检查
        if self.symbol_pairing_enabled:
            logger.debug("执行符号配对检查...")
            symbol_pairing_errors = self.symbol_pairing_check(text)
            errors.extend(symbol_pairing_errors)
            logger.info(f"符号配对检查完成，发现 {len(symbol_pairing_errors)} 个错误")

        # 4. 执行敏感词匹配
        if self.sensitive_word_enabled:
            logger.debug("执行敏感词匹配...")
            sensitive_errors = self.sensitive_word_check(text)
            errors.extend(sensitive_errors)
            logger.info(f"敏感词匹配完成，发现 {len(sensitive_errors)} 个错误")

        # 不再去重处理，保留所有错误
        logger.info(f"文本校对完成，共发现 {len(errors)} 个错误")
        return errors

    def sensitive_word_check(self, text: str) -> list[TextError]:
        """
        敏感词匹配逻辑

        检查文本中是否包含敏感词，支持例外词配置。

        参数:
            text: 要校对的文本

        返回:
            List[TextError]: 错误列表
        """
        errors = []
        logger.debug("开始执行敏感词匹配...")

        if not self.sensitive_words_map:
            logger.info("敏感词词典为空，跳过检查。")
            return errors

        for sensitive_word, exceptions in self.sensitive_words_map.items():
            if not sensitive_word:
                continue

            # 查找所有敏感词出现的位置
            try:
                for match in re.finditer(re.escape(sensitive_word), text, re.IGNORECASE):
                    start, end = match.span()
                    is_exception = False

                    # 检查是否属于例外情况
                    if exceptions and isinstance(exceptions, list):
                        # 过滤出有效的例外词（必须包含敏感词）
                        valid_exceptions = [ex for ex in exceptions if ex and sensitive_word in ex]

                        if not valid_exceptions:
                            logger.debug(f"敏感词 '{sensitive_word}' 没有有效的例外情况")
                        else:
                            # 计算所有例外词的最大长度
                            max_exception_len = max(len(ex) for ex in valid_exceptions)

                            # 统一计算扩展范围
                            context_start = max(0, start - max_exception_len)
                            context_end = min(len(text), end + max_exception_len)
                            context = text[context_start:context_end]

                            logger.debug(f"检查上下文 [{context_start}:{context_end}]: '{context}'")

                            # 在统一的上下文中查找例外词
                            for exception in valid_exceptions:
                                if exception in context:
                                    # 精确验证敏感词是否在例外词内部
                                    exception_idx = context.find(exception)
                                    exception_start_in_text = context_start + exception_idx
                                    exception_end_in_text = exception_start_in_text + len(exception)

                                    # 检查敏感词是否完全在例外词范围内
                                    if exception_start_in_text <= start and end <= exception_end_in_text:
                                        is_exception = True
                                        logger.debug(
                                            f"命中敏感词 '{sensitive_word}' (位置 {start}-{end})，"
                                            f"但属于例外情况 '{exception}' (位置 {exception_start_in_text}-{exception_end_in_text})，跳过。"
                                        )
                                        break

                    if not is_exception:
                        errors.append(
                            TextError(
                                start_pos=start,
                                end_pos=end,
                                error_text=match.group(0),
                                suggestion=t("proofread.suggestion.sensitive"),
                                error_type=t("proofread.error_type.sensitive"),
                                source=t("proofread.source.sensitive"),
                            )
                        )
                        logger.warning(f"发现敏感词: '{match.group(0)}' (位置: {start}-{end})")

            except re.error as e:
                logger.error(f"处理敏感词 '{sensitive_word}' 的正则表达式时出错: {e}")

        logger.info(f"敏感词匹配完成，发现 {len(errors)} 个潜在问题。")
        return errors

    def typos_check(self, text: str) -> list[TextError]:
        """
        错别字校对逻辑

        参数:
            text: 要校对的文本

        返回:
            List[TextError]: 错误列表

        功能说明:
            1. 处理错别字
            2. 处理标点符号错误
            3. 避免在同一位置重复报告错误
        """
        logger.debug("开始执行错别字校对...")
        logger.debug(f"输入文本长度: {len(text)} 字符")

        errors = []
        typo_errors = 0

        try:
            # ================== 步骤1：构建错误写法到正确写法的映射表 ==================
            logger.debug("构建错误写法到正确写法的映射表...")

            # 1.1 构建错别字映射（仅包含非标点错误）
            typo_correction_map = {}
            for correct_word, wrong_words in self.typos_map.items():
                # 跳过空值
                if not correct_word or not wrong_words:
                    logger.debug(f"跳过空值: correct_word={correct_word}, wrong_words={wrong_words}")
                    continue

                # 确保wrong_words是可迭代对象
                if not isinstance(wrong_words, list):
                    logger.warning(f"错别字配置类型错误: {wrong_words}, 跳过")
                    continue

                for wrong_word in wrong_words:
                    # 跳过空值
                    if not wrong_word:
                        logger.debug("跳过空错误词")
                        continue

                    # 添加到映射表
                    typo_correction_map[wrong_word] = correct_word
                    logger.debug(f"添加错别字映射: '{wrong_word}' -> '{correct_word}'")

            # 1.2 获取标点符号映射（新格式：正确标点→错误标点列表）
            symbol_map = self.symbol_map.copy()
            logger.debug(f"加载标点符号映射: {len(symbol_map)} 项")

            logger.info(f"映射表构建完成 | 错别字: {len(typo_correction_map)} | 标点: {len(symbol_map)}")

            # ================== 步骤2：处理错别字（非标点错误） ==================
            logger.debug("开始错别字检测...")

            # 2.1 遍历所有可能的错别字
            for wrong_word, correct_word in typo_correction_map.items():
                # 2.2 在文本中查找所有错别字出现的位置
                # 中文无需单词边界检测，直接搜索字符串
                pattern = re.escape(wrong_word)

                # 2.3 查找所有匹配
                for match in re.finditer(pattern, text):
                    start, end = match.span()
                    error_text = text[start:end]

                    # 2.4 创建错误记录
                    errors.append(
                        TextError(
                            start_pos=start,
                            end_pos=end,
                            error_text=error_text,
                            suggestion=f"{correct_word}",
                            error_type=t("proofread.error_type.typo"),
                            source=t("proofread.source.typo"),
                        )
                    )
                    typo_errors += 1
                    logger.debug(f"发现错别字: '{error_text}' -> '{correct_word}' (位置: {start}-{end})")

            # ================== 结果统计 ==================
            logger.info(f"错别字校对完成，发现 {typo_errors} 个错误")

        except Exception as e:
            # 记录详细错误信息
            logger.error(f"错别字校对失败: {e!s}", exc_info=True)
            # 返回已发现的错误（可能不完整）
            logger.warning("返回已发现的错误列表（可能不完整）")

        return errors

    def symbol_correction_check(self, text: str) -> list[TextError]:
        """
        符号校对逻辑（独立方法）

        检查文本中的错误标点符号，如英文标点应改为中文标点等。

        参数:
            text: 要校对的文本

        返回:
            List[TextError]: 错误列表
        """
        logger.debug("开始执行符号校对...")
        logger.debug(f"输入文本长度: {len(text)} 字符")

        errors = []
        symbol_errors = 0

        try:
            # 获取标点符号映射（正确标点→错误标点列表）
            if not self.symbol_map:
                logger.warning("符号校对已启用，但符号映射表为空")
                return errors

            logger.debug(f"加载标点符号映射: {len(self.symbol_map)} 项")

            # 遍历标点符号映射表（正确标点→错误标点列表）
            for correct_punc, wrong_puncs in self.symbol_map.items():
                logger.debug(f"处理正确标点: '{correct_punc}' 的 {len(wrong_puncs) if wrong_puncs else 0} 个错误形式")

                # 确保错误列表是有效的
                if not wrong_puncs or not isinstance(wrong_puncs, list):
                    logger.warning(f"标点错误列表无效: {type(wrong_puncs)}, 跳过")
                    continue

                # 遍历每个可能的错误标点
                for wrong_punc in wrong_puncs:
                    if not wrong_punc:
                        logger.debug("跳过空错误标点")
                        continue

                    logger.debug(f"检查错误标点: '{wrong_punc}' -> 正确应为: '{correct_punc}'")

                    # 在文本中查找所有错误标点出现的位置
                    pattern = re.escape(wrong_punc)
                    for match in re.finditer(pattern, text):
                        start, end = match.span()
                        error_text = text[start:end]

                        # 添加标点错误
                        errors.append(
                            TextError(
                                start_pos=start,
                                end_pos=end,
                                error_text=error_text,
                                suggestion=f"{correct_punc}",
                                error_type=t("proofread.error_type.punctuation"),
                                source=t("proofread.source.symbol"),
                            )
                        )
                        symbol_errors += 1
                        logger.debug(f"发现标点错误: '{error_text}' -> '{correct_punc}' (位置: {start}-{end})")

            logger.info(f"符号校对完成，发现 {symbol_errors} 个标点错误")

        except Exception as e:
            logger.error(f"符号校对失败: {e!s}", exc_info=True)
            logger.warning("返回已发现的错误列表（可能不完整）")

        return errors

    def symbol_pairing_check(self, text: str) -> list[TextError]:
        """
        检查符号配对情况

        参数:
            text: 要检查的文本

        返回:
            List[TextError]: 错误列表
        """
        errors = []
        logger.debug("开始符号配对检查...")

        if not self.symbol_pairs:
            logger.warning("符号配对检查已启用，但符号对列表为空")
            return errors

        # 为每种符号对创建栈
        stacks = {pair[0]: [] for pair in self.symbol_pairs}
        closing_map = {pair[1]: pair[0] for pair in self.symbol_pairs}

        logger.debug(f"检查的符号对: {self.symbol_pairs}")

        # 第一遍扫描：检查未闭合的符号
        for i, char in enumerate(text):
            # 如果是开符号，入栈
            if char in stacks:
                stacks[char].append(i)
                logger.debug(f"开符号入栈: '{char}' (位置: {i})")
            # 如果是闭符号，检查是否匹配
            elif char in closing_map:
                opening_char = closing_map[char]
                if stacks[opening_char]:
                    # 匹配成功，出栈
                    stacks[opening_char].pop()
                    logger.debug(f"闭符号匹配: '{char}' (位置: {i})")
                else:
                    # 多余的闭符号
                    errors.append(
                        TextError(
                            start_pos=i,
                            end_pos=i + 1,
                            error_text=char,
                            suggestion=t("proofread.suggestion.extra_close"),
                            error_type=t("proofread.error_type.unmatched"),
                            source=t("proofread.source.pairing"),
                        )
                    )
                    logger.debug(f"发现多余的闭符号: '{char}' (位置: {i})")

        # 第二遍扫描：处理未闭合的开符号
        for opening_char, stack in stacks.items():
            for pos in stack:
                # 未闭合的开符号
                errors.append(
                    TextError(
                        start_pos=pos,
                        end_pos=pos + 1,
                        error_text=opening_char,
                        suggestion=t("proofread.suggestion.unclosed"),
                        error_type=t("proofread.error_type.unmatched"),
                        source=t("proofread.source.pairing"),
                    )
                )
                logger.debug(f"发现未闭合的开符号: '{opening_char}' (位置: {pos})")

        logger.info(f"符号配对检查完成，发现 {len(errors)} 个错误")
        return errors

    def __str__(self):
        """返回验证器状态信息"""
        status = "已初始化" if self._initialized else "未初始化"
        return f"文本验证器({status})"
