"""
CLI辅助函数模块

提供CLI所需的各种辅助功能：
- 文件路径解析（支持拖放格式）
- 文件类型检测和分类
- 通配符扩展
- 进度显示回调
- 用户输入辅助函数
"""

import os
import glob
import logging
from typing import List, Dict, Optional, Callable
from pathlib import Path

from docwen.cli.i18n import cli_t

logger = logging.getLogger(__name__)

# ==================== 文件路径处理 ====================

def parse_file_path(user_input: str) -> str:
    """
    解析用户输入的文件路径，支持多种格式
    
    支持格式：
    1. PowerShell拖放: & 'd:\测试.md'
    2. 标准拖放: "d:\测试.md"
    3. 带空格未加引号: d:\测试.md
    
    Args:
        user_input: 用户输入的路径字符串
        
    Returns:
        str: 解析后的文件路径
    """
    logger.debug(f"解析路径: {user_input}")
    
    # 1. 处理PowerShell拖放格式 (& 'path')
    if user_input.startswith('& '):
        path_part = user_input[2:].strip()
        if path_part.startswith("'") and path_part.endswith("'"):
            return path_part[1:-1]
        elif path_part.startswith('"') and path_part.endswith('"'):
            return path_part[1:-1]
        else:
            return path_part
    
    # 2. 处理标准拖放格式 (双引号包裹)
    elif user_input.startswith('"') and user_input.endswith('"'):
        return user_input[1:-1]
    
    # 3. 直接返回（可能包含空格）
    return user_input.strip()


def expand_paths(paths: List[str]) -> List[str]:
    """
    展开路径列表，处理通配符和目录
    
    Args:
        paths: 路径列表（可能包含通配符）
        
    Returns:
        List[str]: 展开后的文件路径列表
    """
    expanded = []
    
    for path in paths:
        # 解析路径格式
        path = parse_file_path(path)
        
        # 检查是否为目录
        if os.path.isdir(path):
            # 递归查找所有文件
            for root, _, files in os.walk(path):
                for file in files:
                    if is_supported_file(file):
                        expanded.append(os.path.join(root, file))
        
        # 检查是否包含通配符
        elif '*' in path or '?' in path:
            # 使用glob展开
            matches = glob.glob(path, recursive=True)
            expanded.extend([m for m in matches if os.path.isfile(m) and is_supported_file(m)])
        
        # 普通文件
        elif os.path.isfile(path):
            expanded.append(path)
        
        else:
            logger.warning(f"路径不存在或无效: {path}")
    
    # 去重并排序
    expanded = sorted(set(expanded))
    logger.debug(f"展开后得到 {len(expanded)} 个文件")
    
    return expanded


def is_supported_file(filename: str) -> bool:
    """
    检查文件是否为支持的类型
    
    Args:
        filename: 文件名
        
    Returns:
        bool: 是否支持
    """
    supported_exts = {
        '.md', '.txt',
        '.docx', '.doc', '.wps', '.odt', '.rtf',
        '.xlsx', '.xls', '.et', '.csv', '.ods',
        '.pdf', '.ofd', '.xps',
        '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tif', '.tiff', '.webp', '.heic', '.heif'
    }
    
    ext = os.path.splitext(filename)[1].lower()
    return ext in supported_exts


# ==================== 文件分类 ====================

def detect_category(file_path: str) -> str:
    """
    检测文件类别
    
    Args:
        file_path: 文件路径
        
    Returns:
        str: 类别名称 (markdown, document, spreadsheet, image, layout)
    """
    ext = os.path.splitext(file_path)[1].lower()
    
    # Markdown
    if ext in ['.md', '.txt']:
        return 'markdown'
    
    # Document
    elif ext in ['.docx', '.doc', '.wps', '.odt', '.rtf']:
        return 'document'
    
    # Spreadsheet
    elif ext in ['.xlsx', '.xls', '.et', '.csv', '.ods']:
        return 'spreadsheet'
    
    # Layout
    elif ext in ['.pdf', '.ofd', '.xps']:
        return 'layout'
    
    # Image
    elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tif', '.tiff', '.webp', '.heic', '.heif']:
        return 'image'
    
    else:
        return 'unknown'


def categorize_files(files: List[str]) -> Dict[str, List[str]]:
    """
    将文件列表按类别分组
    
    Args:
        files: 文件路径列表
        
    Returns:
        Dict[str, List[str]]: 类别 -> 文件列表的映射
    """
    categories = {}
    
    for file in files:
        category = detect_category(file)
        if category != 'unknown':
            if category not in categories:
                categories[category] = []
            categories[category].append(file)
    
    return categories


def detect_format(file_path: str) -> str:
    """
    检测文件格式（不带点号）
    
    Args:
        file_path: 文件路径
        
    Returns:
        str: 格式名称（小写，无点号）
    """
    ext = os.path.splitext(file_path)[1].lower()
    return ext[1:] if ext.startswith('.') else ext


# ==================== 进度显示 ====================

def create_progress_callback(quiet: bool = False, verbose: bool = False) -> Callable[[str], None]:
    """
    创建进度回调函数
    
    Args:
        quiet: 安静模式
        verbose: 详细模式
        
    Returns:
        Callable: 进度回调函数
    """
    def callback(message: str):
        if not quiet:
            if verbose:
                print(f"[进度] {message}")
            else:
                print(f"... {message}")
    
    return callback


def print_separator(char: str = '=', length: int = 60):
    """打印分隔线"""
    print(char * length)


def print_header(text: str, char: str = '=', length: int = 60):
    """打印标题"""
    print_separator(char, length)
    print(text)
    print_separator(char, length)


# ==================== 用户输入 ====================

def prompt_yes_no(question: str, default: bool = True) -> bool:
    """
    询问是/否问题
    
    Args:
        question: 问题文本
        default: 默认值
        
    Returns:
        bool: 用户选择
    """
    suffix = " [Y/n]: " if default else " [y/N]: "
    while True:
        response = input(question + suffix).strip().lower()
        
        if not response:
            return default
        
        if response in ['y', 'yes', '是']:
            return True
        elif response in ['n', 'no', '否']:
            return False
        else:
            print(cli_t("cli.prompts.invalid_yes_no", default="无效输入，请输入 y/n"))


def prompt_choice(options: List[str], prompt_text: str = None, default: int = 1) -> int:
    """
    让用户从选项列表中选择
    
    Args:
        options: 选项列表
        prompt_text: 提示文本（可选，默认使用国际化文本）
        default: 默认选项（1-based）
        
    Returns:
        int: 用户选择的索引（1-based）
    """
    if prompt_text is None:
        prompt_text = cli_t("cli.prompts.select", default="请选择")
    
    # 显示选项
    for i, option in enumerate(options, 1):
        print(f"[{i}] {option}")
    
    # 获取输入
    default_hint = cli_t("cli.prompts.default", default="默认")
    while True:
        try:
            response = input(f"{prompt_text} [1-{len(options)}, {default_hint}{default}]: ").strip()
            
            if not response:
                return default
            
            choice = int(response)
            if 1 <= choice <= len(options):
                return choice
            else:
                print(cli_t("cli.prompts.enter_number_range", default="请输入 1-{max} 之间的数字", max=len(options)))
        
        except ValueError:
            print(cli_t("cli.prompts.invalid_number", default="无效输入，请输入数字"))


def prompt_input(prompt_text: str, default: str = "", allow_empty: bool = True) -> str:
    """
    获取用户文本输入
    
    Args:
        prompt_text: 提示文本
        default: 默认值
        allow_empty: 是否允许空输入
        
    Returns:
        str: 用户输入
    """
    default_label = cli_t("cli.prompts.default", default="默认")
    suffix = f" [{default_label}: {default}]: " if default else ": "
    
    while True:
        response = input(prompt_text + suffix).strip()
        
        if not response:
            if default:
                return default
            elif allow_empty:
                return ""
            else:
                print(cli_t("cli.prompts.cannot_be_empty", default="此项不能为空"))
        else:
            return response


def prompt_files(prompt_text: str = "请输入文件路径或拖入文件") -> List[str]:
    """
    获取用户输入的文件列表
    
    Args:
        prompt_text: 提示文本
        
    Returns:
        List[str]: 文件路径列表
    """
    user_input = input(f"{prompt_text}: ").strip()
    
    if not user_input:
        return []
    
    # 分割多个路径（可能用空格或逗号分隔）
    if ',' in user_input:
        paths = [p.strip() for p in user_input.split(',')]
    else:
        paths = [user_input]
    
    # 展开路径
    return expand_paths(paths)


# ==================== 文件信息显示 ====================

def format_file_size(size_bytes: int) -> str:
    """
    格式化文件大小
    
    Args:
        size_bytes: 字节数
        
    Returns:
        str: 格式化后的大小
    """
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f}{unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f}TB"


def get_file_info(file_path: str) -> Dict[str, any]:
    """
    获取文件信息
    
    Args:
        file_path: 文件路径
        
    Returns:
        Dict: 文件信息字典
    """
    stat = os.stat(file_path)
    
    return {
        'name': os.path.basename(file_path),
        'path': file_path,
        'size': stat.st_size,
        'size_str': format_file_size(stat.st_size),
        'category': detect_category(file_path),
        'format': detect_format(file_path)
    }


def print_file_list(files: List[str], title: str = None):
    """
    打印文件列表
    
    Args:
        files: 文件路径列表
        title: 标题（可选，默认使用国际化文本）
    """
    if title is None:
        title = cli_t("cli.messages.file_list", default="文件列表")
    
    count_label = cli_t("cli.messages.file_count_suffix", default="个")
    print(f"\n{title} ({len(files)}{count_label}):")
    print_separator('-')
    
    for i, file in enumerate(files, 1):
        info = get_file_info(file)
        print(f"  [{i}] {info['name']} ({info['size_str']})")


# ==================== 验证函数 ====================

def validate_file(file_path: str) -> tuple[bool, str]:
    """
    验证文件是否有效
    
    Args:
        file_path: 文件路径
        
    Returns:
        tuple: (是否有效, 错误/警告信息)
    """
    if not os.path.exists(file_path):
        return False, cli_t("cli.validation.file_not_exists", default="文件不存在")
    
    if not os.path.isfile(file_path):
        return False, cli_t("cli.validation.not_a_file", default="路径不是文件")
    
    if os.path.getsize(file_path) == 0:
        return False, cli_t("cli.validation.file_empty", default="文件为空")
    
    # 使用实际格式检测（与GUI保持一致）
    try:
        from docwen.utils.file_type_utils import validate_file_format, ACTUAL_FORMAT_TO_CATEGORY
        
        validation = validate_file_format(file_path)
        actual_format = validation['actual_format']
        
        # 检查实际格式是否支持
        if actual_format == 'unknown' or actual_format not in ACTUAL_FORMAT_TO_CATEGORY:
            return False, cli_t("cli.validation.unsupported_type", default="不支持的文件类型")
        
        # 格式不匹配时返回警告（但仍接受文件，与GUI一致）
        if not validation['is_match']:
            return True, validation['warning_message']
        
        return True, ""
        
    except Exception as e:
        logger.debug(f"格式检测失败，回退到扩展名检查: {e}")
        # 回退到原有的扩展名检查
        if not is_supported_file(file_path):
            return False, cli_t("cli.validation.unsupported_type", default="不支持的文件类型")
        return True, ""


def validate_files(files: List[str]) -> tuple[List[str], List[tuple[str, str]], List[tuple[str, str]]]:
    """
    批量验证文件
    
    Args:
        files: 文件路径列表
        
    Returns:
        tuple: (有效文件列表, 无效文件列表[(路径, 原因)], 警告文件列表[(路径, 警告)])
    """
    valid = []
    invalid = []
    warnings = []
    
    for file in files:
        is_valid, message = validate_file(file)
        if is_valid:
            valid.append(file)
            # 如果有警告信息（格式不匹配但仍接受）
            if message:
                warnings.append((file, message))
        else:
            invalid.append((file, message))
    
    return valid, invalid, warnings
