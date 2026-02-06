"""
TOML 文件读写操作

提供 TOML 文件的读取、写入和更新功能，支持保留注释和格式。

详细说明：
    本模块封装了 tomlkit 库的操作，提供统一的 TOML 文件处理接口。
    支持两种模式：
    1. 普通模式：将 TOML 转换为普通 Python 字典，用于配置读取
    2. 文档模式：保留 TOMLDocument 对象，用于保留注释的更新操作

主要功能：
    - read_toml_file: 读取 TOML 文件为字典
    - write_toml_file: 将字典写入 TOML 文件
    - update_toml_value: 更新单个配置值（保留注释）
    - get_toml_value: 获取单个配置值
    - extract_inline_comments: 提取行内注释
    - save_mapping_with_comments: 保存映射数据和注释

依赖：
    - tomlkit: TOML 解析和序列化库（保留注释）
    - safe_logger: 安全日志记录

使用方式：
    from docwen.config.toml_operations import (
        read_toml_file,
        write_toml_file,
        update_toml_value,
        get_toml_value,
    )
    
    # 读取配置
    config = read_toml_file("configs/gui_config.toml")
    
    # 获取单个值
    theme = get_toml_value("configs/gui_config.toml", "gui_config.theme", "default_theme")
    
    # 更新单个值（保留注释）
    update_toml_value("configs/gui_config.toml", "gui_config.theme", "default_theme", "dark")
"""

import os
import shutil
from typing import Dict, Any, Optional, Union
from pathlib import Path

import tomlkit
from tomlkit import parse, document, table, TOMLDocument

from .safe_logger import safe_log


# ==============================================================================
#                              基础读写函数
# ==============================================================================

def read_toml_file(filepath: Union[str, Path]) -> Dict[str, Any]:
    """
    读取 TOML 文件内容并转换为字典
    
    参数：
        filepath: TOML 文件路径，可以是字符串或 Path 对象
        
    返回：
        Dict[str, Any]: 解析后的配置字典，如果文件不存在或解析失败返回空字典
    
    示例：
        config = read_toml_file("configs/gui_config.toml")
        window_width = config.get("gui_config", {}).get("window", {}).get("width")
    """
    filepath = Path(filepath)
    
    if not filepath.is_file():
        safe_log.warning("TOML文件不存在: %s", filepath)
        return {}
    
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read()
        
        # 解析 TOML 内容
        parsed = parse(content)
        
        # 转换为普通字典
        return tomlkit_to_dict(parsed)
        
    except Exception as e:
        safe_log.error("解析TOML文件失败: %s | 错误: %s", filepath, str(e))
        return {}


def write_toml_file(filepath: Union[str, Path], data: Dict[str, Any]) -> bool:
    """
    将字典数据写入 TOML 文件
    
    参数：
        filepath: 目标文件路径，可以是字符串或 Path 对象
        data: 要写入的数据字典
        
    返回：
        bool: 写入是否成功
    
    注意：
        此函数不保留原文件的注释，如需保留注释请使用 update_toml_value
    """
    filepath = Path(filepath)
    
    try:
        # 确保目录存在
        filepath.parent.mkdir(parents=True, exist_ok=True)
        
        # 创建 TOML 文档
        doc = document()
        
        # 递归添加数据
        def add_to_doc(doc_obj, data_dict, path=""):
            for key, value in data_dict.items():
                current_path = f"{path}.{key}" if path else key
                
                if isinstance(value, dict):
                    # 创建新表并递归添加
                    new_table = table()
                    add_to_doc(new_table, value, current_path)
                    doc_obj[key] = new_table
                elif isinstance(value, list):
                    # 处理列表，尝试保持格式
                    doc_obj[key] = value
                else:
                    # 直接添加简单值
                    doc_obj[key] = value
        
        add_to_doc(doc, data)
        
        # 写入文件
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(tomlkit.dumps(doc))
        
        safe_log.info("成功写入TOML文件: %s", filepath)
        return True
        
    except Exception as e:
        safe_log.error("写入TOML文件失败: %s | 错误: %s", filepath, str(e))
        return False


# ==============================================================================
#                              文档模式函数（保留注释）
# ==============================================================================

def read_toml_document(filepath: Union[str, Path]) -> Optional[TOMLDocument]:
    """
    读取 TOML 文件内容并返回完整的 TOMLDocument 对象（保留注释）
    
    参数：
        filepath: TOML 文件路径，可以是字符串或 Path 对象
        
    返回：
        Optional[TOMLDocument]: 解析后的 TOML 文档对象，
                                如果文件不存在或解析失败返回 None
    
    用途：
        当需要修改 TOML 文件并保留原有注释时使用此函数
    """
    filepath = Path(filepath)
    
    if not filepath.is_file():
        safe_log.warning("TOML文件不存在: %s", filepath)
        return None
    
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read()
        
        # 解析 TOML 内容，保留注释
        return parse(content)
        
    except Exception as e:
        safe_log.error("解析TOML文档失败: %s | 错误: %s", filepath, str(e))
        return None


def write_toml_document(filepath: Union[str, Path], doc: TOMLDocument) -> bool:
    """
    将 TOMLDocument 对象写入文件（保留注释）
    
    参数：
        filepath: 目标文件路径，可以是字符串或 Path 对象
        doc: TOML 文档对象
        
    返回：
        bool: 写入是否成功
    """
    filepath = Path(filepath)
    
    try:
        # 确保目录存在
        filepath.parent.mkdir(parents=True, exist_ok=True)
        
        # 写入文件
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(tomlkit.dumps(doc))
        
        safe_log.info("成功写入TOML文档: %s", filepath)
        return True
        
    except Exception as e:
        safe_log.error("写入TOML文档失败: %s | 错误: %s", filepath, str(e))
        return False


# ==============================================================================
#                              配置更新函数
# ==============================================================================

def update_toml_value(
    filepath: Union[str, Path], 
    section: str, 
    key: str, 
    value: Any, 
    create_missing: bool = True
) -> bool:
    """
    更新 TOML 文件中的特定值（保留注释）
    
    参数：
        filepath: TOML 文件路径，可以是字符串或 Path 对象
        section: 节名称，支持多级，如 "gui_config.window"
        key: 键名称
        value: 新值
        create_missing: 是否创建不存在的节或键，默认 True
        
    返回：
        bool: 更新是否成功
    
    示例：
        # 更新窗口宽度
        update_toml_value("configs/gui_config.toml", "gui_config.window", "width", 1024)
        
        # 更新主题
        update_toml_value("configs/gui_config.toml", "gui_config.theme", "default_theme", "dark")
    """
    filepath = Path(filepath)
    
    try:
        # 读取现有内容或创建新文档（保留注释）
        if filepath.exists():
            doc = read_toml_document(filepath)
            if doc is None:
                doc = document()
        else:
            doc = document()
            safe_log.info("创建新的TOML文件: %s", filepath)
        
        # 分割多级节名称
        section_parts = section.split('.')
        
        # 导航到目标节
        current = doc
        for part in section_parts:
            if part not in current:
                if create_missing:
                    current[part] = table()
                    current = current[part]
                else:
                    safe_log.error("节不存在且不允许创建: %s", section)
                    return False
            else:
                current = current[part]
        
        # 更新值
        current[key] = value
        
        # 写回文件（保留注释）
        success = write_toml_document(filepath, doc)
        if success:
            safe_log.info("更新TOML值成功: %s -> %s.%s = %s", filepath, section, key, value)
        
        return success
        
    except Exception as e:
        safe_log.error("更新TOML值失败: %s | 错误: %s", filepath, str(e))
        return False


# ==============================================================================
#                              工具函数
# ==============================================================================

def tomlkit_to_dict(toml_data: Union[TOMLDocument, Any]) -> Dict[str, Any]:
    """
    将 tomlkit 对象转换为普通字典
    
    参数：
        toml_data: tomlkit 对象（TOMLDocument、Table 等）
        
    返回：
        Dict[str, Any]: 转换后的普通 Python 字典
    
    说明：
        tomlkit 的解析结果包含特殊类型（如 Integer、String 等），
        此函数将其转换为普通 Python 类型，便于后续处理。
    """
    if isinstance(toml_data, dict):
        return {k: tomlkit_to_dict(v) for k, v in toml_data.items()}
    elif hasattr(toml_data, 'unwrap'):
        # 处理 tomlkit 的特殊类型
        return tomlkit_to_dict(toml_data.unwrap())
    elif isinstance(toml_data, list):
        return [tomlkit_to_dict(item) for item in toml_data]
    else:
        return toml_data


def get_toml_value(
    filepath: Union[str, Path], 
    section: str, 
    key: str, 
    default: Any = None
) -> Any:
    """
    获取 TOML 文件中的特定值
    
    参数：
        filepath: TOML 文件路径，可以是字符串或 Path 对象
        section: 节名称，支持多级，如 "gui_config.window"
        key: 键名称
        default: 如果值不存在时返回的默认值，默认 None
        
    返回：
        Any: 找到的值或默认值
    
    示例：
        width = get_toml_value("configs/gui_config.toml", "gui_config.window", "width", 800)
    """
    filepath = Path(filepath)
    
    if not filepath.exists():
        return default
    
    try:
        # 读取文件内容
        config_data = read_toml_file(filepath)
        if not config_data:
            return default
        
        # 分割多级节名称
        section_parts = section.split('.')
        
        # 导航到目标节
        current = config_data
        for part in section_parts:
            if part not in current:
                return default
            current = current[part]
        
        # 返回值
        return current.get(key, default)
        
    except Exception as e:
        safe_log.error("获取TOML值失败: %s -> %s.%s | 错误: %s", 
                     filepath, section, key, str(e))
        return default


def validate_toml_syntax(filepath: Union[str, Path]) -> bool:
    """
    验证 TOML 文件语法是否正确
    
    参数：
        filepath: TOML 文件路径，可以是字符串或 Path 对象
        
    返回：
        bool: 语法是否正确
    
    用途：
        在写入配置前验证文件格式，或检查用户编辑的配置文件
    """
    filepath = Path(filepath)
    
    if not filepath.exists():
        return False
    
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read()
        
        # 尝试解析
        parse(content)
        return True
        
    except Exception as e:
        safe_log.error("TOML语法验证失败: %s | 错误: %s", filepath, str(e))
        return False


# ==============================================================================
#                              注释处理函数
# ==============================================================================

def extract_inline_comments(filepath: Union[str, Path], section: str) -> Dict[str, str]:
    """
    从 TOML 文件中提取指定 section 的行内注释
    
    参数：
        filepath: TOML 文件路径，可以是字符串或 Path 对象
        section: 节名称，如 "sensitive_words"
        
    返回：
        Dict[str, str]: 键到注释内容的映射，格式为 {键: 注释内容}
    
    用途：
        用于映射编辑器，保留用户添加的注释说明
    """
    filepath = Path(filepath)
    comments_dict = {}
    
    try:
        doc = read_toml_document(filepath)
        if doc is None:
            return comments_dict
        
        # 分割多级节名称
        section_parts = section.split('.')
        
        # 导航到目标节
        current = doc
        for part in section_parts:
            if part not in current:
                safe_log.warning("节不存在: %s", section)
                return comments_dict
            current = current[part]
        
        # 提取每个键的行内注释
        for key in current:
            if hasattr(current[key], 'trivia') and current[key].trivia.comment:
                # 移除注释标记 "#" 和首尾空格
                comment_text = current[key].trivia.comment.strip()
                if comment_text.startswith('#'):
                    comment_text = comment_text[1:].strip()
                comments_dict[key] = comment_text
        
        safe_log.debug("提取了 %d 条注释从 %s.%s", len(comments_dict), filepath, section)
        return comments_dict
        
    except Exception as e:
        safe_log.error("提取注释失败: %s -> %s | 错误: %s", filepath, section, str(e))
        return comments_dict


def save_mapping_with_comments(
    filepath: Union[str, Path], 
    section: str, 
    mapping_data: Dict[str, list],
    comments_data: Dict[str, str]
) -> bool:
    """
    保存映射数据和注释到 TOML 文件
    
    参数：
        filepath: TOML 文件路径，可以是字符串或 Path 对象
        section: 节名称，如 "sensitive_words"
        mapping_data: 映射数据，格式为 {键: [值列表]}
        comments_data: 注释数据，格式为 {键: 注释文本}
        
    返回：
        bool: 保存是否成功
    
    用途：
        用于映射编辑器保存数据，同时保留用户添加的注释
    
    示例：
        mapping = {"敏感词1": ["替换词1", "替换词2"]}
        comments = {"敏感词1": "这是一条说明"}
        save_mapping_with_comments("config.toml", "sensitive_words", mapping, comments)
    """
    filepath = Path(filepath)
    
    try:
        # 读取现有文档或创建新文档
        if filepath.exists():
            doc = read_toml_document(filepath)
            if doc is None:
                doc = document()
        else:
            doc = document()
        
        # 分割多级节名称
        section_parts = section.split('.')
        
        # 导航到目标节，如果不存在则创建
        current = doc
        for part in section_parts:
            if part not in current:
                current[part] = table()
            current = current[part]
        
        # 清空现有数据
        keys_to_remove = list(current.keys())
        for key in keys_to_remove:
            del current[key]
        
        # 添加新数据和注释
        for key, values in mapping_data.items():
            # 将键转为字符串
            key_str = str(key)
            
            # 添加键值对
            current[key_str] = values
            
            # 如果有注释，添加行内注释
            if key_str in comments_data and comments_data[key_str]:
                comment_text = comments_data[key_str].strip()
                if not comment_text.startswith('#'):
                    comment_text = f"# {comment_text}"
                else:
                    comment_text = f"# {comment_text[1:].strip()}"
                current[key_str].comment(comment_text)
        
        # 写回文件
        success = write_toml_document(filepath, doc)
        if success:
            safe_log.info("保存映射数据和注释成功: %s -> %s (共 %d 条)", 
                         filepath, section, len(mapping_data))
        
        return success
        
    except Exception as e:
        safe_log.error("保存映射数据和注释失败: %s -> %s | 错误: %s", 
                      filepath, section, str(e))
        return False
