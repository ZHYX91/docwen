"""
YAML头部生成模块

负责生成简化模式Markdown文件的YAML Front Matter头部。

主要功能：
- build_yaml_header(): 根据元数据生成YAML头部行列表
- 处理aliases列表
- 处理纯数字值的引号包裹

国际化说明：
YAML 键名通过 i18n 模块获取当前语言版本：
- 中文环境输出: 标题: xxx, 副标题: xxx
- 英文环境输出: title: xxx, subtitle: xxx

使用示例：
    from .yaml_builder import build_yaml_header
    
    metadata = {
        'aliases': ['文档标题'],
        '标题': '文档标题',
        '副标题': '副标题内容'
    }
    
    yaml_lines = build_yaml_header(metadata)
    # 返回: ['---', 'aliases:', '  - 文档标题', '标题: 文档标题', ...]  # 中文环境
    # 返回: ['---', 'aliases:', '  - 文档标题', 'title: 文档标题', ...]  # 英文环境
"""

import logging
from docwen.utils.text_utils import is_pure_number, format_yaml_value
from docwen.i18n import t

logger = logging.getLogger(__name__)


def build_yaml_header(metadata: dict) -> list:
    """
    根据元数据生成YAML头部行列表
    
    生成符合Markdown Front Matter规范的YAML头部，包括：
    - aliases列表（用于Obsidian等笔记软件）
    - 标题字段
    - 副标题字段
    
    参数:
        metadata: dict - 元数据字典，包含以下键：
            - aliases: list[str] - 别名列表（通常包含标题）
            - 标题: str - 文档标题
            - 副标题: str - 文档副标题（可选）
    
    返回:
        list[str]: YAML头部行列表，包含起始和结束的 '---' 标记
    
    示例:
        >>> metadata = {'aliases': ['测试'], '标题': '测试', '副标题': ''}
        >>> build_yaml_header(metadata)
        ['---', 'aliases:', '  - 测试', '标题: 测试', '副标题: ', '---', '']
    
    注意:
        - 使用 format_yaml_value() 安全处理特殊字符（[] {} : # ' " 等）
        - aliases为空时输出 'aliases: []'
        - 标题和副标题始终输出，即使为空
    """
    lines = []
    
    # 开始标记
    lines.append("---")
    
    # 提取字段
    aliases = metadata.get('aliases', [])
    title = metadata.get('标题', '')
    subtitle = metadata.get('副标题', '')
    
    # 输出aliases（列表格式）
    if aliases:
        lines.append("aliases:")
        for alias in aliases:
            # 使用 format_yaml_value 安全处理（处理特殊字符、数字、引号等）
            safe_alias = format_yaml_value(alias)
            lines.append(f"  - {safe_alias}")
    else:
        lines.append("aliases: []")
    
    # 获取国际化的 YAML 键名
    title_key = t("yaml_keys.title")
    subtitle_key = t("yaml_keys.subtitle")
    
    # 输出标题（始终输出，即使为空）
    if title:
        # 使用 format_yaml_value 安全处理
        safe_title = format_yaml_value(title)
        lines.append(f"{title_key}: {safe_title}")
    else:
        lines.append(f"{title_key}: ")
    
    # 输出副标题（始终输出，即使为空）
    if subtitle:
        # 使用 format_yaml_value 安全处理
        safe_subtitle = format_yaml_value(subtitle)
        lines.append(f"{subtitle_key}: {safe_subtitle}")
    else:
        lines.append(f"{subtitle_key}: ")
    
    # 结束标记
    lines.append("---")
    lines.append("")  # 空行分隔YAML和正文
    
    logger.debug(f"YAML头部生成完成: 标题='{title}', 副标题='{subtitle}', aliases数量={len(aliases)}")
    
    return lines


def build_yaml_header_string(metadata: dict) -> str:
    """
    根据元数据生成YAML头部字符串
    
    这是 build_yaml_header() 的便捷方法，直接返回拼接后的字符串。
    
    参数:
        metadata: dict - 元数据字典（同 build_yaml_header）
    
    返回:
        str: YAML头部字符串，各行用换行符连接
    
    示例:
        >>> metadata = {'aliases': ['测试'], '标题': '测试', '副标题': ''}
        >>> print(build_yaml_header_string(metadata))
        ---
        aliases:
          - 测试
        标题: 测试
        副标题: 
        ---
    """
    return "\n".join(build_yaml_header(metadata))
