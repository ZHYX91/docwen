"""
Markdown处理工具模块
包含Markdown文件解析和处理工具函数
"""

import re
import logging
from .heading_utils import detect_heading_level  # 导入标题处理工具

# 配置日志
logger = logging.getLogger(__name__)

def extract_yaml(content: str) -> tuple:
    """提取YAML头部和MD正文"""
    logger.debug("提取YAML内容")
    
    # 使用多行匹配模式，提高兼容性
    pattern = r'^---\s*\n(.*?)\n---\s*\n(.*)$'
    match = re.search(pattern, content, re.DOTALL)
    
    if match:
        yaml_content = match.group(1)
        md_body = match.group(2)
        logger.debug(f"成功提取YAML, 长度: {len(yaml_content)} 字符")
        return yaml_content, md_body
    
    logger.warning("未找到YAML头部, 返回全部内容作为正文")
    return "", content

def clean_heading(text: str) -> str:
    """清理标题中的序号（兼容中英文符号混用）"""
    logger.debug(f"清理标题: {text}")
    
    # 使用标题工具模块的功能
    cleaned, _ = detect_heading_level(text)
    
    logger.debug(f"清理后: {cleaned}")
    return cleaned


def format_image_link(filename: str, style: str = "wiki_embed") -> str:
    """
    格式化图片链接
    
    参数:
        filename: 图片文件名
        style: 链接样式，可选值：
            - "markdown_embed": ![image.png](image.png) - 嵌入显示
            - "markdown_link": [image.png](image.png) - 仅链接
            - "wiki_embed": ![[image.png]] - Wiki嵌入显示
            - "wiki_link": [[image.png]] - Wiki仅链接
    
    返回:
        格式化后的链接字符串
    
    示例:
        >>> format_image_link("image.png", "markdown_embed")
        '![image.png](image.png)'
        >>> format_image_link("image.png", "markdown_link")
        '[image.png](image.png)'
        >>> format_image_link("image.png", "wiki_embed")
        '![[image.png]]'
        >>> format_image_link("image.png", "wiki_link")
        '[[image.png]]'
    """
    style_mapping = {
        "markdown_embed": f"![{filename}]({filename})",
        "markdown_link": f"[{filename}]({filename})",
        "wiki_embed": f"![[{filename}]]",
        "wiki_link": f"[[{filename}]]"
    }
    
    result = style_mapping.get(style, style_mapping["wiki_embed"])
    logger.debug(f"格式化图片链接: {filename} -> {result} (style={style})")
    return result


def format_md_file_link(filename: str, style: str = "wiki_embed") -> str:
    """
    格式化MD文件链接
    
    参数:
        filename: MD文件名
        style: 链接样式，可选值：
            - "markdown_link": [file.md](file.md) - Markdown链接（固定为链接形式）
            - "wiki_embed": ![[file.md]] - Wiki嵌入显示
            - "wiki_link": [[file.md]] - Wiki仅链接
    
    返回:
        格式化后的链接字符串
    
    示例:
        >>> format_md_file_link("file.md", "markdown_link")
        '[file.md](file.md)'
        >>> format_md_file_link("file.md", "wiki_embed")
        '![[file.md]]'
        >>> format_md_file_link("file.md", "wiki_link")
        '[[file.md]]'
    
    注意:
        Markdown格式不支持嵌入MD文件内容，"markdown_embed"会被视为"markdown_link"
    """
    style_mapping = {
        "markdown_link": f"[{filename}]({filename})",
        "markdown_embed": f"[{filename}]({filename})",  # Markdown不支持嵌入，等同于link
        "wiki_embed": f"![[{filename}]]",
        "wiki_link": f"[[{filename}]]"
    }
    
    result = style_mapping.get(style, style_mapping["wiki_embed"])
    logger.debug(f"格式化MD文件链接: {filename} -> {result} (style={style})")
    return result


def _clean_table_line(line: str) -> str:
    """
    清理表格行的前导字符（引用标记、空格、制表符等）
    
    参数:
        line: 原始表格行
        
    返回:
        str: 清理后的表格行（以 | 开头）
        
    示例:
        >>> _clean_table_line("> | 数据 |")
        '| 数据 |'
        >>> _clean_table_line("   | 数据 |")
        '| 数据 |'
    """
    # 移除前导的引用标记(>)、空格和制表符
    cleaned = re.sub(r'^[> \t]+', '', line)
    return cleaned


def parse_table_row(line: str) -> list:
    """
    解析Markdown表格单行
    
    参数:
        line: 表格行字符串，如 "| 姓名 | 年龄 |"
        
    返回:
        list: 单元格内容列表
        
    示例:
        >>> parse_table_row("| 张三 | 25 |")
        ['张三', '25']
    """
    # 清理前导字符（支持引用块和缩进中的表格）
    line = _clean_table_line(line)
    
    # 处理转义的竖线符号
    line = line.replace('\\|', '[[PIPE]]')
    
    # 分割单元格（忽略首尾的空列）
    cells = line.split('|')[1:-1]
    
    # 清理每个单元格内容并恢复转义的竖线
    cleaned_cells = [c.strip().replace('[[PIPE]]', '|') for c in cells]
    
    logger.debug(f"解析表格行: {len(cleaned_cells)} 个单元格")
    return cleaned_cells


def is_table_separator(line: str) -> bool:
    """
    判断是否是Markdown表格分隔行
    
    有效的分隔行格式：
    - |---|---|  （基本格式）
    - | :--- | :---: | ---: |  （带对齐标记）
    - |:--|:--:|--:|  （紧凑格式）
    
    参数:
        line: 待判断的行
        
    返回:
        bool: 是否为分隔行
        
    示例:
        >>> is_table_separator("|------|------|")
        True
        >>> is_table_separator("| --- | --- |")
        True
        >>> is_table_separator("| :--- | :---: | ---: |")
        True
        >>> is_table_separator("| 数据 | 数据 |")
        False
        >>> is_table_separator("| --- | 这是数据 |")
        False
    """
    # 清理前导的引用标记和空格
    line = _clean_table_line(line)
    line = line.strip()
    
    # 必须以 | 开头和结尾
    if not line.startswith('|') or not line.endswith('|'):
        return False
    
    # 分割各列（去掉首尾空列）
    cells = line.split('|')[1:-1]
    
    if not cells:
        return False
    
    # 分隔行单元格的有效模式：只能包含 - : 和空格
    separator_cell_pattern = r'^[\s\-:]+$'
    
    for cell in cells:
        cell_stripped = cell.strip()
        # 每列至少要有一个 -
        if '-' not in cell_stripped:
            return False
        # 只能包含 - : 和空格
        if not re.match(separator_cell_pattern, cell):
            return False
    
    return True


def extract_markdown_tables(md_body: str) -> list:
    """
    从Markdown文本中提取所有表格
    
    支持以下格式：
    - 标准表格：| 姓名 | 年龄 |
    - 引用块中的表格：> | 姓名 | 年龄 |
    - 缩进的表格：   | 姓名 | 年龄 |
    
    参数:
        md_body: Markdown正文内容
        
    返回:
        list: 表格列表，每个元素为字典：
        {
            'raw_text': '原始表格文本',
            'headers': ['表头1', '表头2', ...],
            'rows': [
                ['数据1', '数据2', ...],
                ['数据3', '数据4', ...]
            ],
            'start_line': 起始行号,
            'end_line': 结束行号
        }
        
    示例:
        >>> md = "# 标题\\n\\n| 姓名 | 年龄 |\\n|------|------|\\n| 张三 | 25 |\\n"
        >>> tables = extract_markdown_tables(md)
        >>> len(tables)
        1
        >>> tables[0]['headers']
        ['姓名', '年龄']
    """
    logger.info("开始提取Markdown表格...")
    
    tables = []
    
    try:
        # 使用正则匹配所有表格
        # 支持引用块(>)和缩进(空格/制表符)中的表格
        # 注意：使用 [ \t] 而不是 \s，避免匹配换行符导致行号计算错误
        table_pattern = r'([> \t]*\|.*\|(?:\n[> \t]*\|.*\|)+)'
        matches = re.finditer(table_pattern, md_body)
        
        for match_idx, match in enumerate(matches):
            try:
                table_text = match.group(1)
                
                # 分割表格行
                lines = table_text.strip().split('\n')
                
                if len(lines) < 3:  # 至少需要表头、分隔行和一行数据
                    logger.debug(f"跳过表格 {match_idx+1}: 行数不足 ({len(lines)}行)")
                    continue
                
                # 解析表头（第一行）
                header_line = lines[0]
                headers = parse_table_row(header_line)
                
                if not headers:
                    logger.debug(f"跳过表格 {match_idx+1}: 无法解析表头")
                    continue
                
                # 验证分隔行（第二行）
                if not is_table_separator(lines[1]):
                    logger.debug(f"跳过表格 {match_idx+1}: 第二行不是有效的分隔行")
                    continue
                
                # 解析数据行（第三行及以后）
                rows = []
                for line_idx, line in enumerate(lines[2:], start=2):
                    row_data = parse_table_row(line)
                    if row_data:
                        # 确保行的列数与表头一致（补齐或截断）
                        if len(row_data) < len(headers):
                            # 补齐空值
                            row_data.extend([''] * (len(headers) - len(row_data)))
                        elif len(row_data) > len(headers):
                            # 截断多余列
                            row_data = row_data[:len(headers)]
                        rows.append(row_data)
                
                if not rows:
                    logger.debug(f"跳过表格 {match_idx+1}: 没有数据行")
                    continue
                
                # 计算表格在原文中的行号
                start_pos = match.start()
                start_line = md_body[:start_pos].count('\n')
                end_line = start_line + len(lines)
                
                # 构建表格数据
                table_data = {
                    'raw_text': table_text,
                    'headers': headers,
                    'rows': rows,
                    'start_line': start_line,
                    'end_line': end_line
                }
                
                tables.append(table_data)
                logger.info(f"提取表格 {match_idx+1}: {len(headers)}列 x {len(rows)}行")
                
            except Exception as e:
                logger.warning(f"跳过表格 {match_idx+1}: {str(e)}")
                continue
        
        logger.info(f"共提取到 {len(tables)} 个表格")
        return tables
        
    except Exception as e:
        logger.error(f"表格提取失败: {str(e)}", exc_info=True)
        return []


# 模块测试代码
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(level=logging.DEBUG)
    logger.info("Markdown工具模块测试")
    
    # 测试YAML提取
    test_content = """---
title: 测试文档
date: 2023-01-15
---
# 标题内容
正文内容"""
    
    yaml_part, md_part = extract_yaml(test_content)
    assert yaml_part == "title: 测试文档\ndate: 2023-01-15"
    assert md_part == "# 标题内容\n正文内容"
    logger.debug("YAML提取测试通过")
    
    # 测试标题清理
    assert clean_heading("1. 测试标题") == "测试标题"
    assert clean_heading("（一）二级标题") == "二级标题"
    logger.debug("标题清理测试通过")
    
    logger.info("所有测试通过!")
