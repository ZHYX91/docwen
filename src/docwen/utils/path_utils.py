"""
路径处理工具模块

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

import datetime
import logging
import os
import re
import sys
import tempfile
from pathlib import Path

from docwen.formats import category_from_actual_format
from docwen.utils.file_type_utils import is_supported_file, validate_file_format

# 配置日志
logger = logging.getLogger(__name__)

_TIMESTAMP_SUFFIX_RE = re.compile(r"(_\d{8}_\d{6})(?:.*)?$")


def strip_timestamp_suffix(name: str) -> str:
    match = _TIMESTAMP_SUFFIX_RE.search(name)
    if match:
        return name[: match.start()]
    return name


def generate_timestamp() -> str:
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def make_output_stem(
    input_path: str,
    *,
    section: str = "",
    add_timestamp: bool = True,
    description: str | None = None,
    timestamp_override: str | None = None,
) -> str:
    base_name = Path(input_path).stem
    base_name_clean = strip_timestamp_suffix(base_name)

    parts: list[str] = [base_name_clean]
    if section:
        parts.append(section)
    if add_timestamp:
        parts.append(timestamp_override or generate_timestamp())
    if description:
        parts.append(description)
    return "_".join(parts)


def ensure_unique_file_path(path: str, *, max_counter: int = 999) -> str:
    target = Path(path)
    if not target.exists():
        return str(target)

    directory = target.parent
    name = target.stem
    ext = target.suffix

    for counter in range(1, max_counter + 1):
        candidate = directory / f"{name}_{counter:03d}{ext}"
        if not candidate.exists():
            return str(candidate)

    raise FileExistsError(f"无法生成唯一输出文件名: {path}")


def make_named_output_stem(
    base_name: str,
    *,
    section: str = "",
    add_timestamp: bool = True,
    description: str | None = None,
    timestamp_override: str | None = None,
) -> str:
    base_name_clean = strip_timestamp_suffix(base_name)

    parts: list[str] = [base_name_clean]
    if section:
        parts.append(section)
    if add_timestamp:
        parts.append(timestamp_override or generate_timestamp())
    if description:
        parts.append(description)
    return "_".join(parts)


def generate_named_output_path(
    *,
    output_dir: str,
    base_name: str,
    file_type: str,
    section: str = "",
    add_timestamp: bool = True,
    description: str | None = None,
    timestamp_override: str | None = None,
    strict_dir: bool = False,
) -> str:
    if strict_dir:
        ensure_dir_exists_or_raise(output_dir)
    else:
        ensure_dir_exists(output_dir)

    output_ext = file_type if file_type.startswith(".") else f".{file_type}"
    stem = make_named_output_stem(
        base_name,
        section=section,
        add_timestamp=add_timestamp,
        description=description,
        timestamp_override=timestamp_override,
    )
    return ensure_unique_file_path(str(Path(output_dir) / f"{stem}{output_ext}"))


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
        Path(path).mkdir(parents=True, exist_ok=True)
        logger.info(f"目录已创建/确认: {path}")
        return True
    except Exception as e:
        logger.warning(f"创建目录失败: {path}, 错误: {e!s}")
        return False


def ensure_dir_exists_or_raise(path: str) -> None:
    try:
        Path(path).mkdir(parents=True, exist_ok=True)
    except Exception as e:
        raise OSError(f"无法创建目录: {path}") from e


def get_project_root() -> str:
    """
    获取项目根目录（自适应开发和生产环境）

    返回:
        str: 项目根目录路径
    """
    # 打包环境：可执行文件在软件根目录
    if getattr(sys, "frozen", False):
        project_root = str(Path(sys.executable).resolve().parent)
        logger.debug(f"生产环境项目根目录: {project_root}")
        return project_root

    # 开发环境：项目根目录（src目录的父目录）
    project_root = str(Path(__file__).resolve().parents[3])
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
    target = str(Path(base_dir) / subdir)
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
    abs_path = str(Path(path).expanduser().resolve())

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
    base_path = Path(base).expanduser().resolve()
    full_path = base_path.joinpath(*parts).resolve()

    if not full_path.is_relative_to(base_path):
        logger.error(f"路径安全违规: {full_path} 不在基础目录 {base_path} 内")
        raise ValueError(f"非法路径访问: {full_path}")

    normalized = normalize_path(str(full_path))
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
    expected = expected.lower().replace(".", "")

    target = Path(path)
    current_ext = target.suffix.lstrip(".").lower() if target.suffix else ""

    # 如果扩展名匹配，直接返回
    if current_ext == expected:
        return path

    # 修正扩展名
    new_path = str(target.with_suffix(f".{expected}"))
    logger.warning(f"文件后缀修正: {path} -> {new_path}")
    return new_path


def ensure_unique_directory_path(target_dir: str, max_counter: int = 100, max_ms_attempts: int = 10) -> str:
    """
    生成一个不存在的目录路径，用于避免输出目录冲突。

    参数:
        target_dir: 期望使用的目录路径
        max_counter: 计数后缀最大尝试次数
        max_ms_attempts: 毫秒级时间戳兜底尝试次数

    返回:
        str: 不存在的目录路径

    异常:
        FileExistsError: 多次尝试后仍无法生成唯一目录名
    """
    if not target_dir:
        raise ValueError("target_dir不能为空")

    target = Path(target_dir)
    if not target.exists():
        return str(target)

    for i in range(1, max_counter + 1):
        candidate = Path(f"{target_dir}_{i:03d}")
        if not candidate.exists():
            return str(candidate)

    for _ in range(max_ms_attempts):
        ms = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
        candidate = Path(f"{target_dir}_{ms}")
        if not candidate.exists():
            return str(candidate)

    raise FileExistsError(f"无法生成唯一输出目录名: {target_dir}")


def get_temp_dir(name: str = "docwen") -> str:
    """
    获取带命名空间的临时目录

    参数:
        name: 应用名称（用于隔离）

    返回:
        str: 临时目录路径
    """
    temp_dir = Path(tempfile.gettempdir()) / name
    ensure_dir_exists(str(temp_dir))
    logger.info(f"临时目录: {temp_dir}")
    return str(temp_dir)


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
    temp_path = str(Path(temp_dir) / file_name)
    logger.debug(f"创建临时文件路径: {temp_path}")
    return temp_path


def generate_output_path(
    input_path: str,
    output_dir: str | None = None,
    section: str = "",
    add_timestamp: bool = True,
    description: str | None = None,
    file_type: str = "",
    strict_dir: bool = False,
) -> str:
    """
    生成统一格式的输出文件路径（带冲突检测）
    新规则：原文件名(去除旧时间戳和描述)+section+时间戳+描述

    冲突处理：
    - 默认使用秒级时间戳
    - 如发生同名冲突，在文件名末尾追加 _001、_002... 直到唯一

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
    base_name = Path(input_path).stem
    logger.debug(f"原始基础文件名: {base_name}")
    base_name_clean = strip_timestamp_suffix(base_name)
    logger.debug(f"清理后的基础文件名: {base_name_clean}")

    # 确定输出目录
    if not output_dir:
        output_dir = str(Path(input_path).parent)
        logger.debug(f"使用输入文件目录: {output_dir}")

    # 创建输出目录（如果不存在）
    if strict_dir:
        ensure_dir_exists_or_raise(output_dir)
    else:
        ensure_dir_exists(output_dir)

    output_ext = ""
    if file_type:
        output_ext = file_type if file_type.startswith(".") else f".{file_type}"

    base_filename = make_output_stem(
        input_path,
        section=section,
        add_timestamp=add_timestamp,
        description=description,
    )
    logger.debug(f"组合文件名: {base_filename}")

    candidate = f"{base_filename}{output_ext}"
    output_path = Path(output_dir) / candidate
    if not output_path.exists():
        logger.info(f"生成输出路径: {input_path} -> {output_path}")
        return str(output_path)

    for counter in range(1, 1000):
        candidate = f"{base_filename}_{counter:03d}{output_ext}"
        output_path = Path(output_dir) / candidate
        if not output_path.exists():
            logger.info(f"生成输出路径: {input_path} -> {output_path}")
            return str(output_path)

    raise FileExistsError(f"无法生成唯一输出文件名: {Path(output_dir) / (base_filename + output_ext)}")


def split_path(path: str) -> tuple[str, str, str]:
    """
    拆分路径为目录、文件名和扩展名

    参数:
        path: 文件路径

    返回:
        tuple: (目录, 文件名, 扩展名)
    """
    target = Path(path)
    return (str(target.parent), target.stem, target.suffix.lstrip("."))


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
    timestamp = generate_timestamp()
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
    resource_path = str(Path(project_root) / "resources" / resource_name)
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
    config_path = str(Path(project_root) / "configs" / f"{config_name}.toml")
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
    template_path = str(Path(project_root) / "templates" / template_name)
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
    asset_path = str(Path(project_root) / "assets" / asset_name)
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
        size_bytes = Path(file_path).stat().st_size

        # 定义单位
        units = ["B", "KB", "MB", "GB"]
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
        logger.error(f"获取文件大小失败: {file_path}, 错误: {e!s}")
        return "未知大小"


def collect_files_from_folder(folder_path: str) -> list[str]:
    """
    递归收集文件夹及其子文件夹中的所有支持文件

    参数:
        folder_path: 文件夹路径

    返回:
        List[str]: 支持的文件路径列表
    """
    logger.debug(f"开始收集文件夹中的文件: {folder_path}")
    supported_files = []
    seen = set()
    ignored_dir_names = {
        ".git",
        "__pycache__",
        ".pytest_cache",
        "node_modules",
        ".venv",
        "venv",
        ".idea",
    }

    try:
        for root, dirs, files in os.walk(folder_path):
            dirs[:] = [d for d in dirs if d not in ignored_dir_names]
            for file in files:
                file_path = str(Path(root) / file)
                if file_path in seen:
                    continue

                is_supported = False
                try:
                    validation = validate_file_format(file_path)
                    actual_format = validation["actual_format"]
                    is_supported = category_from_actual_format(actual_format) != "unknown"
                except Exception as e:
                    logger.debug(f"实际格式检测失败，回退到扩展名检查: {file_path}, {e}")
                    is_supported = is_supported_file(file_path)

                if is_supported:
                    supported_files.append(file_path)
                    seen.add(file_path)
                    logger.debug(f"找到支持文件: {file_path}")
                else:
                    logger.debug(f"跳过不支持文件: {file_path}")

        supported_files = sorted(supported_files, key=lambda p: p.lower())
        logger.info(f"从文件夹收集到 {len(supported_files)} 个支持文件: {folder_path}")
        return supported_files
    except Exception as e:
        logger.error(f"收集文件夹文件失败: {folder_path}, 错误: {e!s}")
        return []
