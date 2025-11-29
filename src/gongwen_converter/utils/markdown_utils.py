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


def format_image_link(filename: str, link_format: str = "wiki", embed: bool = True) -> str:
    """
    格式化图片链接
    
    参数:
        filename: 图片文件名
        link_format: "markdown" 或 "wiki"
        embed: 是否嵌入显示
    
    返回:
        格式化后的链接字符串
    
    示例:
        >>> format_image_link("image.png", "markdown", True)
        '![image.png](image.png)'
        >>> format_image_link("image.png", "markdown", False)
        '[image.png](image.png)'
        >>> format_image_link("image.png", "wiki", True)
        '![[image.png]]'
        >>> format_image_link("image.png", "wiki", False)
        '[[image.png]]'
    """
    if link_format == "wiki":
        result = f"![[{filename}]]" if embed else f"[[{filename}]]"
    else:  # markdown
        result = f"![{filename}]({filename})" if embed else f"[{filename}]({filename})"
    
    logger.debug(f"格式化图片链接: {filename} -> {result} (format={link_format}, embed={embed})")
    return result


def format_md_file_link(filename: str, link_format: str = "wiki", embed: bool = True) -> str:
    """
    格式化MD文件链接
    
    参数:
        filename: MD文件名
        link_format: "markdown" 或 "wiki"
        embed: 是否嵌入（仅wiki有效）
    
    返回:
        格式化后的链接字符串
    
    示例:
        >>> format_md_file_link("file.md", "markdown", True)
        '[file.md](file.md)'
        >>> format_md_file_link("file.md", "markdown", False)
        '[file.md](file.md)'
        >>> format_md_file_link("file.md", "wiki", True)
        '![[file.md]]'
        >>> format_md_file_link("file.md", "wiki", False)
        '[[file.md]]'
    
    注意:
        Markdown格式固定为链接形式，embed参数被忽略
    """
    if link_format == "wiki":
        result = f"![[{filename}]]" if embed else f"[[{filename}]]"
    else:  # markdown
        # Markdown格式固定为链接形式
        result = f"[{filename}]({filename})"
    
    logger.debug(f"格式化MD文件链接: {filename} -> {result} (format={link_format}, embed={embed})")
    return result


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
    
    参数:
        line: 待判断的行
        
    返回:
        bool: 是否为分隔行
        
    示例:
        >>> is_table_separator("|------|------|")
        True
        >>> is_table_separator("| --- | --- |")
        True
        >>> is_table_separator("| 数据 | 数据 |")
        False
    """
    # 移除首尾空白
    line = line.strip()
    
    # 检查是否匹配表格分隔行模式
    # 格式：|------|------| 或 | --- | --- |
    pattern = r'^\s*\|[\s\-:]+\|[\s\-:]*'
    
    if re.match(pattern, line):
        # 进一步检查是否包含足够的 - 字符（至少3个）
        if line.count('-') >= 3:
            return True
    
    return False


def extract_markdown_tables(md_body: str) -> list:
    """
    从Markdown文本中提取所有表格
    
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
        # 表格格式：至少两行，每行都有竖线分隔
        table_pattern = r'(\|.*\|(?:\n\|.*\|)+)'
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
