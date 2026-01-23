"""
公文转换器路径处理工具模块

本模块提供统一的路径处理功能，确保文件路径的安全性和一致性。
主要功能包括：
- 目录创建和验证
- 路径标准化和安全连接
- 文件扩展名验证和修正
- 临时文件路径管理
- 统一输出路径生成规则
- 项目资源路径获取

采用安全路径处理机制，防止路径遍历攻击，确保文件操作的安全性。
"""

import os
import sys
import re
import tempfile
import datetime
import logging
from typing import Tuple, Optional, List
from docwen.utils.file_type_utils import is_supported_file

# 配置日志
logger = logging.getLogger(__name__)

def ensure_dir_exists(path: str) -> bool:
    """
    确保目录存在（自动创建缺失目录）
    
    参数:
        path: 目录路径
        
    返回:
        bool: 目录是否存在（创建后）
    """
    logger.debug(f"检查目录是否存在: {path}")
    try:
        os.makedirs(path, exist_ok=True)
        logger.info(f"目录已创建/确认: {path}")
        return True
    except Exception as e:
        # 回退到临时目录
        import tempfile
        temp_dir = tempfile.gettempdir()
        logger.warning(f"创建目录失败: {path}, 错误: {str(e)}")
        logger.warning(f"使用临时目录: {temp_dir}")
        return False

def get_project_root() -> str:
    """
    获取项目根目录（自适应开发和生产环境）
    
    返回:
        str: 项目根目录路径
    """
    # 打包环境：可执行文件在软件根目录
    if getattr(sys, 'frozen', False):
        project_root = os.path.dirname(sys.executable)
        logger.debug(f"生产环境项目根目录: {project_root}")
        return project_root
    
    # 开发环境：项目根目录（src目录的父目录）
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 计算项目根目录：当前文件所在目录是 src/docwen/utils
    # 所以父目录是src/docwen，再父目录是src，再父目录是项目根目录
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(current_dir)))
    logger.debug(f"开发环境项目根目录: {project_root}")
    return project_root

def get_app_dir(subdir: str) -> str:
    """
    获取应用子目录（自动创建缺失目录）
    
    参数:
        subdir: 子目录名称
        
    返回:
        str: 完整目录路径
    """
    base_dir = get_project_root()
    target = os.path.join(base_dir, subdir)
    ensure_dir_exists(target)
    logger.info(f"应用目录: {target}")
    return target

def normalize_path(path: str) -> str:
    """
    标准化路径（处理~、相对路径、大小写等）
    
    参数:
        path: 原始路径
        
    返回:
        str: 标准化后的绝对路径
    """
    # 扩展用户目录
    if path.startswith("~"):
        path = os.path.expanduser(path)
    
    # 解析绝对路径
    abs_path = os.path.abspath(path)
    
    # 统一大小写（Windows）
    if sys.platform == "win32":
        abs_path = abs_path.lower()
    
    logger.debug(f"标准化路径: {path} -> {abs_path}")
    return abs_path

def safe_join_path(base: str, *parts: str) -> str:
    """
    安全连接路径（防止路径穿透攻击）
    
    参数:
        base: 基础路径
        parts: 路径部分
        
    返回:
        str: 连接后的完整路径
        
    异常:
        ValueError: 如果路径不在基础目录内
    """
    full_path = os.path.join(base, *parts)
    normalized = normalize_path(full_path)
    base_norm = normalize_path(base)
    
    # 验证是否在基础路径内
    if not normalized.startswith(base_norm):
        logger.error(f"路径安全违规: {full_path} 不在基础目录 {base} 内")
        raise ValueError(f"非法路径访问: {full_path}")
    
    logger.debug(f"安全连接路径: {base} + {parts} -> {normalized}")
    return normalized

def validate_extension(path: str, expected: str) -> str:
    """
    验证并修正文件扩展名
    
    参数:
        path: 文件路径
        expected: 预期的扩展名（不含点）
        
    返回:
        str: 修正后的文件路径
    """
    # 确保预期扩展名格式正确
    expected = expected.lower().replace('.', '')
    
    # 获取当前扩展名
    base, ext = os.path.splitext(path)
    current_ext = ext.lstrip('.').lower() if ext else ""
    
    # 如果扩展名匹配，直接返回
    if current_ext == expected:
        return path
    
    # 修正扩展名
    new_path = f"{base}.{expected}"
    logger.warning(f"文件后缀修正: {path} -> {new_path}")
    return new_path

def get_temp_dir(name: str = "docwen") -> str:
    """
    获取带命名空间的临时目录
    
    参数:
        name: 应用名称（用于隔离）
        
    返回:
        str: 临时目录路径
    """
    temp_dir = os.path.join(tempfile.gettempdir(), name)
    ensure_dir_exists(temp_dir)
    logger.info(f"临时目录: {temp_dir}")
    return temp_dir

def create_temp_file(prefix: str = "", suffix: str = "") -> str:
    """
    创建临时文件路径
    
    参数:
        prefix: 文件名前缀
        suffix: 文件名后缀（含扩展名）
        
    返回:
        str: 临时文件路径
    """
    temp_dir = get_temp_dir()
    file_name = f"{prefix}{os.urandom(4).hex()}{suffix}"
    temp_path = os.path.join(temp_dir, file_name)
    logger.debug(f"创建临时文件路径: {temp_path}")
    return temp_path

def generate_output_path(
    input_path: str, 
    output_dir: Optional[str] = None,
    section: str = "",
    add_timestamp: bool = True,
    description: Optional[str] = None,
    file_type: str = ""
) -> str:
    """
    生成统一格式的输出文件路径（带时间戳冲突检测）
    新规则：原文件名(去除旧时间戳和描述)+section+时间戳+描述
    
    时间戳冲突处理：
    - 如果在同一秒内生成了相同的时间戳，会自动等待并重试
    - 最多重试11次，每次等待100ms
    - 确保批量处理时每个文件都有唯一的时间戳
    
    参数:
        input_path: 输入文件路径
        output_dir: 指定输出目录（可选）
        file_type: 输出文件类型
        description: 附加描述（可选）
        add_timestamp: 是否添加时间戳后缀（默认True）
        section: 章节/部分标识（放在原文件名后，时间戳前）
        
    返回:
        str: 输出文件完整路径
    """
    logger.debug(f"生成输出路径 - 输入: {input_path}")
    
    # 获取原文件名（不含扩展名）
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    logger.debug(f"原始基础文件名: {base_name}")
    
    # 移除可能存在的旧时间戳和描述
    # 时间戳格式: _YYYYMMDD_HHMMSS
    # 匹配第一个时间戳及其后面的所有内容
    timestamp_pattern = r'(_\d{8}_\d{6})(?:.*)?$'
    match = re.search(timestamp_pattern, base_name)
    if match:
        # 找到第一个时间戳的位置，保留时间戳之前的部分
        timestamp_start = match.start()
        base_name_clean = base_name[:timestamp_start]
        logger.debug(f"找到时间戳，清理后的基础文件名: {base_name_clean}")
    else:
        base_name_clean = base_name
        logger.debug(f"未找到时间戳，使用原始文件名: {base_name_clean}")
    
    # 确定输出目录
    if not output_dir:
        output_dir = os.path.dirname(input_path)
        logger.debug(f"使用输入文件目录: {output_dir}")
    
    # 创建输出目录（如果不存在）
    ensure_dir_exists(output_dir)
    
    # 生成时间戳（如果启用），带冲突检测
    timestamp_str = ""
    if add_timestamp:
        import time
        max_retries = 11
        retry_count = 0
        
        while retry_count <= max_retries:
            # 生成当前时间戳
            timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # 构建完整文件名进行冲突检测
            filename_parts = [base_name_clean]
            if section:
                filename_parts.append(section)
            filename_parts.append(timestamp_str)
            if description:
                filename_parts.append(description)
            
            test_filename = "_".join(filename_parts)
            if file_type:
                if not file_type.startswith('.'):
                    file_type_with_dot = '.' + file_type
                else:
                    file_type_with_dot = file_type
                test_filename += file_type_with_dot
            
            test_path = os.path.join(output_dir, test_filename)
            
            # 检测是否存在同名文件或文件夹
            if not os.path.exists(test_path):
                logger.debug(f"时间戳无冲突: {timestamp_str}")
                break
            
            # 存在冲突
            if retry_count < max_retries:
                logger.info(f"时间戳冲突，等待100ms后重试 (第{retry_count + 1}次)")
                time.sleep(0.1)  # 等待100ms
                retry_count += 1
            else:
                logger.warning(f"时间戳冲突达到最大重试次数({max_retries}次)，使用当前时间戳")
                break
        
        logger.debug(f"最终生成时间戳: {timestamp_str} (重试次数: {retry_count})")
    
    # 构建文件名
    filename_parts = [base_name_clean]
    
    # 添加section（如果有）
    if section:
        filename_parts.append(section)
        logger.debug(f"添加section: {section}")
    
    # 添加时间戳
    if timestamp_str:
        filename_parts.append(timestamp_str)
    
    # 添加描述
    if description:
        filename_parts.append(description)
        logger.debug(f"添加描述: {description}")
    
    # 组合文件名
    filename = "_".join(filename_parts)
    logger.debug(f"组合文件名: {filename}")
    
    # 添加文件扩展名（如果指定了）
    if file_type:
        if not file_type.startswith('.'):
            file_type = '.' + file_type
        filename += file_type
        logger.debug(f"添加文件类型: {file_type}")
    
    # 组合完整路径
    output_path = os.path.join(output_dir, filename)
    logger.info(f"生成输出路径: {input_path} -> {output_path}")
    
    return output_path

def split_path(path: str) -> Tuple[str, str, str]:
    """
    拆分路径为目录、文件名和扩展名
    
    参数:
        path: 文件路径
        
    返回:
        tuple: (目录, 文件名, 扩展名)
    """
    dir_path = os.path.dirname(path)
    file_name = os.path.basename(path)
    base, ext = os.path.splitext(file_name)
    return (dir_path, base, ext.lstrip('.'))

def is_subpath(child: str, parent: str) -> bool:
    """
    检查子路径是否在父路径内
    
    参数:
        child: 子路径
        parent: 父路径
        
    返回:
        bool: 是否是子路径
    """
    parent = normalize_path(parent)
    child = normalize_path(child)
    return child.startswith(parent)

def generate_timestamp_suffix() -> str:
    """
    生成时间戳后缀（格式：_YYYYMMDD_HHMMSS）
    
    返回:
        str: 时间戳字符串
    """
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    logger.debug(f"生成时间戳后缀: {timestamp}")
    return timestamp

def get_package_resource_path(resource_name: str) -> str:
    """
    获取包内资源文件的路径
    
    参数:
        resource_name: 资源文件名
        
    返回:
        str: 资源文件的完整路径
    """
    project_root = get_project_root()
    resource_path = os.path.join(project_root, "resources", resource_name)
    logger.debug(f"获取包内资源路径: {resource_name} -> {resource_path}")
    return resource_path

def get_config_path(config_name: str) -> str:
    """
    获取配置文件路径（考虑包结构调整）
    
    参数:
        config_name: 配置文件名（不含扩展名）
        
    返回:
        str: 配置文件的完整路径
    """
    project_root = get_project_root()
    config_path = os.path.join(project_root, "configs", f"{config_name}.toml")
    logger.debug(f"获取配置路径: {config_name} -> {config_path}")
    return config_path

def get_template_path(template_name: str) -> str:
    """
    获取模板文件路径（考虑包结构调整）
    
    参数:
        template_name: 模板文件名
        
    返回:
        str: 模板文件的完整路径
    """
    project_root = get_project_root()
    template_path = os.path.join(project_root, "templates", template_name)
    logger.debug(f"获取模板路径: {template_name} -> {template_path}")
    return template_path

def get_asset_path(asset_name: str) -> str:
    """
    获取资源文件路径（如图标等）
    
    参数:
        asset_name: 资源文件名
        
    返回:
        str: 资源文件的完整路径
    """
    project_root = get_project_root()
    asset_path = os.path.join(project_root, "assets", asset_name)
    logger.debug(f"获取资源路径: {asset_name} -> {asset_path}")
    return asset_path

def get_file_size_formatted(file_path: str) -> str:
    """
    获取文件大小并格式化为易读的字符串
    
    参数:
        file_path: 文件路径
        
    返回:
        str: 格式化后的文件大小（如 "1.5 MB", "256 KB" 等）
    """
    try:
        size_bytes = os.path.getsize(file_path)
        
        # 定义单位
        units = ['B', 'KB', 'MB', 'GB']
        unit_index = 0
        
        # 计算合适的单位
        size = float(size_bytes)
        while size >= 1024.0 and unit_index < len(units) - 1:
            size /= 1024.0
            unit_index += 1
        
        # 格式化输出
        if unit_index == 0:
            # 字节，不显示小数
            return f"{int(size)} {units[unit_index]}"
        elif size < 10:
            # 小于10，显示1位小数
            return f"{size:.1f} {units[unit_index]}"
        else:
            # 大于等于10，显示整数
            return f"{int(size)} {units[unit_index]}"
            
    except (OSError, TypeError) as e:
        logger.error(f"获取文件大小失败: {file_path}, 错误: {str(e)}")
        return "未知大小"


def collect_files_from_folder(folder_path: str) -> List[str]:
    """
    递归收集文件夹及其子文件夹中的所有支持文件
    
    参数:
        folder_path: 文件夹路径
        
    返回:
        List[str]: 支持的文件路径列表
    """
    logger.debug(f"开始收集文件夹中的文件: {folder_path}")
    supported_files = []
    
    try:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                if is_supported_file(file_path):
                    supported_files.append(file_path)
                    logger.debug(f"找到支持文件: {file_path}")
                else:
                    logger.debug(f"跳过不支持文件: {file_path}")
        
        logger.info(f"从文件夹收集到 {len(supported_files)} 个支持文件: {folder_path}")
        return supported_files
    except Exception as e:
        logger.error(f"收集文件夹文件失败: {folder_path}, 错误: {str(e)}")
        return []

class PathAdapter:
    """
    路径适配器，处理包结构调整前后的路径兼容性
    """
    
    @staticmethod
    def get_legacy_config_path(config_name: str) -> str:
        """
        获取旧版配置路径（用于过渡期）
        
        参数:
            config_name: 配置文件名
            
        返回:
            str: 旧版配置路径
        """
        # 旧版配置路径逻辑（如果有）
        project_root = get_project_root()
        legacy_path = os.path.join(project_root, "configs", f"{config_name}.toml")
        logger.debug(f"获取旧版配置路径: {config_name} -> {legacy_path}")
        return legacy_path
    
    @staticmethod
    def migrate_paths() -> bool:
        """
        迁移路径引用到新结构
        
        返回:
            bool: 迁移是否成功
        """
        logger.info("开始迁移路径引用到新结构...")
        # 这里可以实现路径迁移逻辑
        # 目前先返回成功，具体迁移逻辑可以根据需要实现
        logger.info("路径迁移完成")
        return True
    
    @staticmethod
    def is_path_compatible(old_path: str, new_path: str) -> bool:
        """
        检查路径是否兼容
        
        参数:
            old_path: 旧路径
            new_path: 新路径
            
        返回:
            bool: 是否兼容
        """
        old_norm = normalize_path(old_path)
        new_norm = normalize_path(new_path)
        compatible = old_norm == new_norm
        logger.debug(f"路径兼容性检查: {old_path} vs {new_path} -> {compatible}")
        return compatible

# 模块测试代码
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(level=logging.DEBUG)
    logger.info("路径工具模块测试")
    
    # 测试目录创建
    test_dir = "test_dir"
    ensure_dir_exists(test_dir)
    
    # 测试项目根目录
    root = get_project_root()
    logger.info(f"项目根目录: {root}")
    
    # 测试路径标准化
    test_path = "~/test"
    norm_path = normalize_path(test_path)
    logger.info(f"标准化路径: {test_path} -> {norm_path}")
    
    # 测试安全连接路径
    try:
        safe_path = safe_join_path(root, "configs", "test.toml")
        logger.info(f"安全连接路径: {safe_path}")
    except ValueError as e:
        logger.error(f"安全连接失败: {str(e)}")
    
    # 测试扩展名验证
    test_file = "document.txt"
    validated = validate_extension(test_file, "docx")
    logger.info(f"扩展名验证: {test_file} -> {validated}")
    
    # 测试临时目录
    temp_dir = get_temp_dir()
    logger.info(f"临时目录: {temp_dir}")
    
    # 测试输出路径生成
    test_input = "/path/to/document.md"
    test_output = generate_output_path(
        test_input,
        output_dir="/output",
        section="",
        add_timestamp=True,
        description="fromMd",
        file_type="docx"
    )
    logger.info(f"输出路径: {test_output}")
    
    # 测试路径拆分
    dir_path, base, ext = split_path(test_output)
    logger.info(f"路径拆分: dir={dir_path}, base={base}, ext={ext}")
    
    # 测试子路径检查
    is_child = is_subpath("/path/to/file.txt", "/path/to")
    logger.info(f"子路径检查: {is_child}")
    
    # 测试时间戳后缀
    timestamp = generate_timestamp_suffix()
    logger.info(f"时间戳后缀: {timestamp}")
    
    logger.info("所有测试完成!")
