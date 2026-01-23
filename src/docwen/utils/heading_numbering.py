"""
标题序号格式化模块
提供标题序号的生成、格式化功能，支持配置化的序号方案
"""

import re
import logging
from typing import Dict, Any, List, Optional

from .number_utils import (
    number_to_chinese,
    number_to_chinese_upper,
    number_to_circled,
    number_to_arabic_full,
    number_to_letter_upper,
    number_to_letter_lower,
    number_to_roman_upper,
    number_to_roman_lower
)

# 配置日志
logger = logging.getLogger(__name__)


class HeadingFormatter:
    """
    标题序号格式化器
    
    负责根据配置的序号方案生成标题序号。
    支持：
    - 模板占位符解析（如 "{1.chinese_lower}、"）
    - 多级标题计数器维护
    - 跨级引用（如三级标题引用一级计数器）
    
    使用示例:
        >>> config = {
        ...     "level_1": {"format": "{1.chinese_lower}、"},
        ...     "level_2": {"format": "（{2.chinese_lower}）"}
        ... }
        >>> formatter = HeadingFormatter(config)
        >>> formatter.increment_level(1)
        >>> print(formatter.format_heading(1))  # 输出: 一、
        >>> formatter.increment_level(2)
        >>> print(formatter.format_heading(2))  # 输出: （一）
    """
    
    # 数字样式到转换函数的映射
    STYLE_CONVERTERS = {
        'chinese_lower': number_to_chinese,
        'chinese_upper': number_to_chinese_upper,
        'arabic_half': str,  # 半角阿拉伯数字直接转字符串
        'arabic_full': number_to_arabic_full,
        'arabic_circled': number_to_circled,
        'letter_upper': number_to_letter_upper,
        'letter_lower': number_to_letter_lower,
        'roman_upper': number_to_roman_upper,
        'roman_lower': number_to_roman_lower
    }
    
    def __init__(self, scheme_config: Dict[str, Any]):
        """
        初始化标题序号格式化器
        
        参数:
            scheme_config: 序号方案配置，包含level_1到level_9的格式定义
                          格式：{"level_1": {"format": "{1.chinese_lower}、"}, ...}
        
        异常:
            ValueError: 如果配置格式不正确
        """
        logger.debug("初始化标题序号格式化器")
        
        if not isinstance(scheme_config, dict):
            logger.error("序号方案配置必须是字典类型")
            raise ValueError("序号方案配置格式错误")
        
        self.scheme_config = scheme_config
        
        # 初始化各级标题计数器（1-9级）
        self.counters = [0] * 9
        
        # 预编译模板（缓存解析结果，提高性能）
        self._template_cache = {}
        
        logger.info(f"标题序号格式化器初始化完成，方案配置包含 {len(scheme_config)} 个级别")
    
    def reset_counters(self):
        """
        重置所有级别的计数器
        
        用于处理新文档时重置状态。
        """
        logger.debug("重置所有标题计数器")
        self.counters = [0] * 9
    
    def increment_level(self, level: int):
        """
        递增指定级别的计数器，并重置其下级计数器
        
        参数:
            level: 标题级别 (1-9)
        
        示例:
            当前计数器: [1, 2, 3, 0, 0, 0, 0, 0, 0]
            increment_level(2) 后: [1, 3, 0, 0, 0, 0, 0, 0, 0]
            （二级计数器+1，三级及以下重置为0）
        """
        if not (1 <= level <= 9):
            logger.warning(f"无效的标题级别: {level}，应在1-9之间")
            return
        
        # 递增当前级别计数器
        self.counters[level - 1] += 1
        
        # 重置所有下级计数器
        for i in range(level, 9):
            self.counters[i] = 0
        
        logger.debug(
            f"递增标题级别 {level}，计数器状态: {self.counters[:level]}"
        )
    
    def format_heading(self, level: int) -> str:
        """
        生成指定级别的标题序号
        
        参数:
            level: 标题级别 (1-9)
            
        返回:
            str: 格式化后的序号字符串，如 "一、"、"（1）"、"第一章　"
                 如果该级别无格式定义，返回空字符串
        
        示例:
            >>> formatter.increment_level(1)
            >>> formatter.format_heading(1)  # "一、"
            >>> formatter.increment_level(2)
            >>> formatter.format_heading(2)  # "（一）"
        """
        if not (1 <= level <= 9):
            logger.warning(f"无效的标题级别: {level}，返回空字符串")
            return ""
        
        # 获取该级别的格式模板
        level_key = f"level_{level}"
        level_config = self.scheme_config.get(level_key, {})
        template = level_config.get("format", "")
        
        if not template:
            logger.debug(f"标题级别 {level} 无格式定义，返回空字符串")
            return ""
        
        # 解析并格式化模板
        try:
            result = self._parse_template(template, level)
            logger.debug(f"标题级别 {level} 格式化结果: '{result}'")
            return result
        except Exception as e:
            logger.error(f"格式化标题级别 {level} 失败: {e}")
            return ""
    
    def _parse_template(self, template: str, current_level: int) -> str:
        """
        解析模板字符串中的占位符
        
        参数:
            template: 模板字符串，如 "{1.chinese_lower}、"
            current_level: 当前标题级别，用于日志记录
            
        返回:
            str: 解析后的字符串，所有占位符被替换为实际序号
        
        占位符格式: {级别.样式}
        - 级别: 1-9
        - 样式: chinese_lower/arabic_half/arabic_circled等
        
        示例:
            >>> self.counters = [1, 2, 3, 0, 0, 0, 0, 0, 0]
            >>> self._parse_template("{1.arabic_half}.{2.arabic_half}.{3.arabic_half} ", 3)
            "1.2.3 "
        """
        # 使用缓存提高性能
        if template in self._template_cache:
            placeholder_pattern = self._template_cache[template]
        else:
            # 编译正则表达式：匹配 {级别.样式} 格式
            placeholder_pattern = re.compile(r'\{(\d+)\.(\w+)\}')
            self._template_cache[template] = placeholder_pattern
        
        def replace_placeholder(match):
            """替换单个占位符"""
            ref_level = int(match.group(1))  # 引用的级别
            style = match.group(2)           # 数字样式
            
            # 验证引用级别
            if not (1 <= ref_level <= 9):
                logger.warning(
                    f"模板中引用的级别 {ref_level} 无效（应在1-9之间），"
                    f"位于级别{current_level}的模板: {template}"
                )
                return str(ref_level)
            
            # 获取对应级别的计数器值
            counter_value = self.counters[ref_level - 1]
            
            # 验证计数器是否已初始化（防止引用未出现的级别）
            if counter_value == 0:
                logger.warning(
                    f"级别{current_level}的模板引用了级别{ref_level}，"
                    f"但该级别计数器为0，使用默认值1"
                )
                counter_value = 1
            
            # 获取样式转换函数
            if style not in self.STYLE_CONVERTERS:
                logger.warning(
                    f"未知的数字样式: {style}，使用半角阿拉伯数字"
                )
                return str(counter_value)
            
            converter = self.STYLE_CONVERTERS[style]
            
            # 执行转换
            try:
                converted = converter(counter_value)
                logger.debug(
                    f"占位符替换: {{level={ref_level}, style={style}}} "
                    f"-> {counter_value} -> '{converted}'"
                )
                return converted
            except Exception as e:
                logger.error(
                    f"转换数字 {counter_value} 为样式 {style} 失败: {e}，"
                    f"使用原始数字"
                )
                return str(counter_value)
        
        # 替换所有占位符
        result = placeholder_pattern.sub(replace_placeholder, template)
        return result
    
    def get_scheme_info(self) -> Dict[str, Any]:
        """
        获取当前方案的信息
        
        返回:
            Dict[str, Any]: 包含方案名称、描述等元信息
        """
        return {
            "name": self.scheme_config.get("name", "未命名方案"),
            "description": self.scheme_config.get("description", ""),
            "levels_defined": len([k for k in self.scheme_config.keys() if k.startswith("level_")])
        }


def get_formatter_from_config(config_manager, scheme_name: str) -> Optional[HeadingFormatter]:
    """
    从配置管理器创建HeadingFormatter实例
    
    参数:
        config_manager: 配置管理器实例
        scheme_name: 序号方案名称（如 "gongwen_standard"）
        
    返回:
        Optional[HeadingFormatter]: 格式化器实例，如果方案不存在则返回None
    
    示例:
        >>> from docwen.config.config_manager import config_manager
        >>> formatter = get_formatter_from_config(config_manager, "gongwen_standard")
        >>> formatter.increment_level(1)
        >>> print(formatter.format_heading(1))  # "一、"
    """
    logger.debug(f"从配置创建序号格式化器，方案: {scheme_name}")
    
    try:
        # 获取所有方案
        all_schemes = config_manager.get_heading_schemes()
        
        # 查找指定方案
        if scheme_name not in all_schemes:
            logger.error(f"序号方案 '{scheme_name}' 不存在")
            return None
        
        scheme_config = all_schemes[scheme_name]
        
        # 创建格式化器
        formatter = HeadingFormatter(scheme_config)
        logger.info(f"成功创建序号格式化器，方案: {scheme_name}")
        return formatter
        
    except Exception as e:
        logger.error(f"创建序号格式化器失败: {e}", exc_info=True)
        return None


# ==============================================
# Markdown文件序号处理函数
# ==============================================

# Markdown标题正则表达式
_MD_HEADING_PATTERN = re.compile(r'^(#{1,6})\s+(.*)$')

# 代码块状态检测
_CODE_BLOCK_PATTERN = re.compile(r'^```')


def remove_numbering_from_md(content: str) -> str:
    """
    从Markdown内容中去除所有标题的序号
    
    处理逻辑:
    1. 逐行扫描，识别Markdown标题行（# ## ### 等）
    2. 跳过代码块内的内容
    3. 对每个标题行，调用 remove_heading_numbering() 去除序号
    4. 保持非标题行不变
    
    参数:
        content: Markdown文件内容
        
    返回:
        str: 去除序号后的内容
        
    示例:
        >>> content = "# 一、标题内容\\n正文"
        >>> remove_numbering_from_md(content)
        "# 标题内容\\n正文"
    """
    from .heading_utils import remove_heading_numbering
    
    lines = content.split('\n')
    result_lines = []
    in_code_block = False
    
    for line in lines:
        # 检测代码块开始/结束
        if _CODE_BLOCK_PATTERN.match(line):
            in_code_block = not in_code_block
            result_lines.append(line)
            continue
        
        # 代码块内不处理
        if in_code_block:
            result_lines.append(line)
            continue
        
        # 检测Markdown标题
        match = _MD_HEADING_PATTERN.match(line)
        if match:
            hashes = match.group(1)  # # 或 ## 等
            heading_text = match.group(2)  # 标题内容
            
            # 去除序号
            cleaned_text = remove_heading_numbering(heading_text)
            
            # 重新组合
            result_lines.append(f"{hashes} {cleaned_text}")
            logger.debug(f"去除序号: '{heading_text}' -> '{cleaned_text}'")
        else:
            result_lines.append(line)
    
    return '\n'.join(result_lines)


def add_numbering_to_md(content: str, scheme_id: str = 'gongwen_standard') -> str:
    """
    为Markdown内容的所有标题添加序号
    
    处理逻辑:
    1. 创建 HeadingFormatter 实例
    2. 逐行扫描，识别Markdown标题行
    3. 跳过代码块内的内容
    4. 根据 # 的数量确定标题级别
    5. 递增计数器并生成序号
    6. 在标题文本前插入序号
    
    参数:
        content: Markdown文件内容
        scheme_id: 序号方案ID
            - 'gongwen_standard': 公文标准（一、（一）1.（1）①）
            - 'hierarchical_standard': 层级数字标准（1 1.1 1.1.1）
            - 'legal_standard': 法律条文标准（第一编 第一章 第一节 第一条）
        
    返回:
        str: 添加序号后的内容
        
    示例:
        >>> content = "# 标题一\\n## 子标题"
        >>> add_numbering_to_md(content, 'gongwen_standard')
        "# 一、标题一\\n## （一）子标题"
    """
    from docwen.config.config_manager import config_manager
    
    # 获取序号方案配置
    try:
        all_schemes = config_manager.get_heading_schemes()
        if scheme_id not in all_schemes:
            logger.error(f"序号方案 '{scheme_id}' 不存在，使用默认方案")
            scheme_id = 'gongwen_standard'
            if scheme_id not in all_schemes:
                logger.error("默认方案也不存在，返回原内容")
                return content
        
        scheme_config = all_schemes[scheme_id]
    except Exception as e:
        logger.error(f"获取序号方案配置失败: {e}")
        return content
    
    # 创建格式化器
    formatter = HeadingFormatter(scheme_config)
    
    lines = content.split('\n')
    result_lines = []
    in_code_block = False
    
    for line in lines:
        # 检测代码块开始/结束
        if _CODE_BLOCK_PATTERN.match(line):
            in_code_block = not in_code_block
            result_lines.append(line)
            continue
        
        # 代码块内不处理
        if in_code_block:
            result_lines.append(line)
            continue
        
        # 检测Markdown标题
        match = _MD_HEADING_PATTERN.match(line)
        if match:
            hashes = match.group(1)  # # 或 ## 等
            heading_text = match.group(2).strip()  # 标题内容
            level = len(hashes)  # 标题级别
            
            # 递增计数器
            formatter.increment_level(level)
            
            # 生成序号
            numbering = formatter.format_heading(level)
            
            if numbering:
                # 添加序号
                new_heading = f"{hashes} {numbering}{heading_text}"
                logger.debug(f"添加序号: '{heading_text}' -> '{numbering}{heading_text}'")
            else:
                # 该级别无序号定义，保持原样
                new_heading = line
                logger.debug(f"级别 {level} 无序号定义，保持原样")
            
            result_lines.append(new_heading)
        else:
            result_lines.append(line)
    
    return '\n'.join(result_lines)


def process_md_numbering(
    content: str, 
    remove_existing: bool = True, 
    add_new: bool = True, 
    scheme_id: str = 'gongwen_standard'
) -> str:
    """
    处理Markdown内容的标题序号（组合函数）
    
    这是一个便捷函数，组合了去除和添加序号的操作。
    典型用法是先去除原有序号，再按指定方案添加新序号。
    
    参数:
        content: Markdown文件内容
        remove_existing: 是否去除原有序号（默认True）
        add_new: 是否添加新序号（默认True）
        scheme_id: 序号方案ID（仅当add_new=True时生效）
        
    返回:
        str: 处理后的内容
        
    示例:
        # 规范化序号（先去除再添加）
        >>> content = "# 1. 标题\\n## A. 子标题"
        >>> process_md_numbering(content, scheme_id='gongwen_standard')
        "# 一、标题\\n## （一）子标题"
        
        # 只去除序号
        >>> process_md_numbering(content, remove_existing=True, add_new=False)
        "# 标题\\n## 子标题"
        
        # 只添加序号（假设原内容无序号）
        >>> process_md_numbering(content, remove_existing=False, add_new=True)
        "# 一、1. 标题\\n## （一）A. 子标题"  # 注意：原序号会保留
    """
    result = content
    
    if remove_existing:
        logger.info("开始去除Markdown标题序号")
        result = remove_numbering_from_md(result)
    
    if add_new:
        logger.info(f"开始添加Markdown标题序号，方案: {scheme_id}")
        result = add_numbering_to_md(result, scheme_id)
    
    return result
