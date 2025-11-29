"""
工作空间管理器

提供临时文件管理和安全保存工具函数。

核心功能：
1. 扩展名修正 - 处理扩展名不匹配的文件
2. 文件保存 - 带降级策略的安全保存
3. 配置读取 - 输出目录和中间文件配置

设计原则：
- 简单直接的工具函数
- 不创建临时目录（由调用者管理）
- 不负责文件命名（由调用者决定）
"""

import os
import shutil
import time
import logging
from typing import Optional, Tuple
from datetime import datetime

logger = logging.getLogger(__name__)


# ==================== 格式到标准扩展名的映射 ====================

FORMAT_EXTENSION_MAP = {
    # 文档格式
    'doc': '.doc',
    'wps': '.doc',
    'docx': '.docx',
    'odt': '.odt',
    'rtf': '.rtf',
    
    # 表格格式
    'xls': '.xls',
    'et': '.xls',
    'xlsx': '.xlsx',
    'ods': '.ods',
    'csv': '.csv',
    
    # 版式格式
    'pdf': '.pdf',
    'ofd': '.ofd',
    'caj': '.caj',
    
    # Markdown
    'md': '.md',
}


def get_standard_extension(actual_format: str) -> str:
    """
    获取格式的标准扩展名
    
    参数:
        actual_format: 文件的真实格式（如 'xls', 'doc', 'md'）
    
    返回:
        str: 标准扩展名（如 '.xls', '.doc', '.md'）
    
    示例:
        >>> get_standard_extension('xls')
        '.xls'
        >>> get_standard_extension('et')
        '.xls'
    """
    standard_ext = FORMAT_EXTENSION_MAP.get(actual_format.lower(), f'.{actual_format}')
    logger.debug(f"格式映射: '{actual_format}' → '{standard_ext}'")
    return standard_ext


# ==================== 临时文件准备 ====================

def prepare_input_file(input_path: str, temp_dir: str, actual_format: str) -> str:
    """
    在临时目录中准备输入文件（修正扩展名）
    
    功能：
    - 复制原文件到临时目录
    - 如果扩展名不匹配，自动修正为标准扩展名
    
    参数:
        input_path: 原始文件路径
        temp_dir: 临时目录路径（由调用者创建）
        actual_format: 文件的真实格式
        
    返回:
        str: 临时目录中的文件路径（扩展名已修正）
    
    示例:
        >>> prepare_input_file("报表.doc", "C:/Temp/tmp123", "xls")
        "C:/Temp/tmp123/input.xls"
    
    说明:
        - 文件名固定为 "input{扩展名}"
        - 由临时目录唯一性保证无碰撞
        - 调试友好，文件名固定易识别
    """
    std_ext = get_standard_extension(actual_format)
    temp_file = os.path.join(temp_dir, f"input{std_ext}")
    
    logger.debug(f"准备输入文件")
    logger.debug(f"  原始: {os.path.basename(input_path)}")
    logger.debug(f"  临时: {os.path.basename(temp_file)}")
    
    shutil.copy2(input_path, temp_file)
    
    logger.info(f"✓ 输入文件已准备: {os.path.basename(temp_file)}")
    return temp_file


# ==================== 文件保存工具 ====================

def _generate_unique_filename(filepath: str) -> str:
    """
    生成唯一的文件名（当文件已存在时）
    
    策略：在文件名后添加数字后缀
    
    参数:
        filepath: 原始文件路径
        
    返回:
        str: 唯一的文件路径
        
    示例:
        报告.pdf → 报告_1.pdf → 报告_2.pdf ...
    """
    if not os.path.exists(filepath):
        return filepath
    
    directory = os.path.dirname(filepath)
    filename = os.path.basename(filepath)
    name, ext = os.path.splitext(filename)
    
    counter = 1
    while True:
        new_filename = f"{name}_{counter}{ext}"
        new_filepath = os.path.join(directory, new_filename)
        if not os.path.exists(new_filepath):
            logger.debug(f"生成唯一文件名: {new_filename}")
            return new_filepath
        counter += 1


def move_file_with_retry(
    source: str,
    destination: str,
    max_retries: int = 3,
    retry_delay: float = 0.5
) -> bool:
    """
    带重试机制的文件移动
    
    处理常见的文件移动失败情况：
    - 权限错误（自动重试）
    - 文件已存在（自动生成新名字）
    - 目录不存在（自动创建）
    
    参数:
        source: 源文件路径
        destination: 目标文件路径
        max_retries: 最大重试次数
        retry_delay: 重试间隔（秒）
        
    返回:
        bool: 移动是否成功
    """
    logger.debug(f"准备移动文件: {os.path.basename(source)} → {destination}")
    
    for attempt in range(max_retries):
        try:
            # 确保目标目录存在
            target_dir = os.path.dirname(destination)
            if target_dir:
                os.makedirs(target_dir, exist_ok=True)
                logger.debug(f"确保目标目录存在: {target_dir}")
            
            # 处理文件已存在的情况
            if os.path.exists(destination):
                logger.warning(f"目标文件已存在，生成新文件名")
                destination = _generate_unique_filename(destination)
            
            # 移动文件
            shutil.move(source, destination)
            logger.info(f"✓ 文件移动成功: {os.path.basename(destination)}")
            return True
            
        except PermissionError as e:
            logger.warning(f"权限错误（尝试 {attempt + 1}/{max_retries}）: {e}")
            if attempt < max_retries - 1:
                time.sleep(retry_delay)
                
        except Exception as e:
            logger.error(f"移动文件失败: {e}")
            if attempt < max_retries - 1:
                time.sleep(retry_delay)
    
    logger.error(f"✗ 文件移动失败（已重试{max_retries}次）")
    return False


def save_output_with_fallback(
    temp_file: str,
    target_path: str,
    original_input_file: str,
    max_retries: int = 3,
    retry_delay: float = 0.5
) -> Tuple[Optional[str], str]:
    """
    三级降级保存策略
    
    降级顺序：
    1. 目标位置（用户指定或配置）
    2. 原文件所在目录
    3. 用户桌面
    4. 临时救援目录
    
    参数:
        temp_file: 临时文件路径
        target_path: 目标文件路径
        original_input_file: 原始输入文件路径
        max_retries: 最大重试次数
        retry_delay: 重试间隔（秒）
        
    返回:
        Tuple[Optional[str], str]: (实际保存路径, 保存位置说明)
    """
    logger.info(f"开始保存文件: {os.path.basename(temp_file)}")
    
    # 尝试1：目标位置
    logger.debug(f"尝试保存到目标位置: {target_path}")
    if move_file_with_retry(temp_file, target_path, max_retries, retry_delay):
        logger.info("✓ 文件已保存到目标位置")
        return target_path, "目标位置"
    
    # 尝试2：原文件所在目录
    source_dir = os.path.dirname(original_input_file)
    source_dir_path = os.path.join(source_dir, os.path.basename(target_path))
    
    logger.warning("目标位置保存失败，尝试备选位置1：原文件所在目录")
    if move_file_with_retry(temp_file, source_dir_path, max_retries, retry_delay):
        logger.info(f"✓ 文件已保存到原文件目录: {source_dir}")
        return source_dir_path, "原文件所在目录（目标位置不可写）"
    
    # 尝试3：用户桌面
    try:
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        if os.path.exists(desktop):
            desktop_path = os.path.join(desktop, os.path.basename(target_path))
            
            logger.warning("原文件目录保存失败，尝试备选位置2：用户桌面")
            if move_file_with_retry(temp_file, desktop_path, max_retries, retry_delay):
                logger.info(f"✓ 文件已保存到桌面: {desktop}")
                return desktop_path, "桌面（原位置和目标位置均不可写）"
    except Exception as e:
        logger.warning(f"尝试保存到桌面失败: {e}")
    
    # 尝试4：临时救援目录
    try:
        import tempfile
        rescue_dir = os.path.join(tempfile.gettempdir(), "gongwen_converter_rescue")
        os.makedirs(rescue_dir, exist_ok=True)
        rescue_path = os.path.join(rescue_dir, os.path.basename(target_path))
        
        logger.error("所有常规位置保存失败，使用临时救援目录")
        shutil.copy2(temp_file, rescue_path)
        logger.warning(f"⚠ 文件已复制到临时救援目录: {rescue_dir}")
        
        return rescue_path, f"临时救援目录（所有位置均不可写）\n路径: {rescue_dir}"
        
    except Exception as e:
        logger.critical(f"复制到救援目录失败: {e}")
        return temp_file, "临时文件（请手动保存）"


# ==================== 配置读取辅助函数 ====================

def get_output_directory(input_file: str, custom_dir: Optional[str] = None) -> str:
    """
    根据配置决定输出目录
    
    决策优先级：
    1. 函数参数 custom_dir（最高优先级）
    2. 配置文件中的自定义目录
    3. 原文件所在目录（默认）
    
    参数:
        input_file: 输入文件路径
        custom_dir: 自定义目录（可选，如提供则直接使用）
        
    返回:
        str: 输出目录路径
    """
    # 优先级1：函数参数
    if custom_dir:
        logger.debug(f"使用参数指定的输出目录: {custom_dir}")
        return custom_dir
    
    # 优先级2：读取配置并确定基础输出目录
    base_dir = None
    output_settings = None
    
    try:
        from gongwen_converter.config.config_manager import config_manager
        output_settings = config_manager.get_output_directory_settings()
        
        mode = output_settings.get("mode", "source")
        logger.debug(f"配置的输出模式: {mode}")
        
        if mode == "custom":
            custom_path = output_settings.get("custom_path", "")
            if custom_path:
                # 处理相对路径
                if not os.path.isabs(custom_path):
                    try:
                        from gongwen_converter.utils.path_utils import get_project_root
                        custom_path = os.path.join(get_project_root(), custom_path)
                    except:
                        pass
                
                base_dir = custom_path
                logger.debug(f"使用配置的自定义目录: {custom_path}")
        
    except Exception as e:
        logger.warning(f"读取输出目录配置失败: {e}")
    
    # 如果没有设置自定义目录，使用原文件所在目录
    if not base_dir:
        base_dir = os.path.dirname(input_file)
        logger.debug(f"使用原文件所在目录: {base_dir}")
    
    # 统一处理日期子文件夹（对两种模式都有效）
    if output_settings and output_settings.get("create_date_subfolder", False):
        date_format = output_settings.get("date_folder_format", "%Y-%m-%d")
        date_folder = datetime.now().strftime(date_format)
        final_path = os.path.join(base_dir, date_folder)
        logger.debug(f"创建日期子文件夹: {date_folder}")
        
        # 确保日期子文件夹存在
        os.makedirs(final_path, exist_ok=True)
        logger.debug(f"已确保目录存在: {final_path}")
        
        return final_path
    
    return base_dir


def should_save_intermediate_files() -> bool:
    """
    检查是否应保存中间文件到输出目录
    
    返回:
        bool: True表示保存中间文件，False表示只保存最终文件
    """
    try:
        from gongwen_converter.config.config_manager import config_manager
        save_intermediate = config_manager.get_save_intermediate_files()
        
        logger.debug(f"中间文件保存配置: {save_intermediate}")
        return save_intermediate
        
    except Exception as e:
        logger.warning(f"读取中间文件配置失败，默认不保存: {e}")
        return False


# ==================== 模块导出 ====================

__all__ = [
    'get_standard_extension',
    'prepare_input_file',
    'move_file_with_retry',
    'save_output_with_fallback',
    'get_output_directory',
    'should_save_intermediate_files',
]
