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

import contextlib
import json
import logging
import shutil
import tempfile
import time
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class IntermediateItem:
    kind: str
    path: str


def mask_input_path(path: str) -> str:
    return str(Path("<redacted>") / Path(path).name)


def to_jsonable(value: object) -> object:
    if value is None:
        return None
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, dict):
        result: dict = {}
        for k, v in value.items():
            if not isinstance(k, str):
                continue
            result[k] = to_jsonable(v)
        return result
    if isinstance(value, (list, tuple)):
        return [to_jsonable(v) for v in value]
    return str(value)


def write_manifest_json(output_dir: str, manifest: dict, *, filename: str = "manifest.json") -> str | None:
    try:
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        target_path = str(Path(output_dir) / filename)

        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            delete=False,
            dir=output_dir,
            suffix=".tmp",
        ) as f:
            json.dump(to_jsonable(manifest), f, ensure_ascii=False, indent=2)
            temp_path = f.name

        return replace_file_atomic(temp_path, target_path)
    except Exception as e:
        logger.warning(f"写入 manifest 失败: {e}")
        return None


def build_manifest(
    *,
    file_path: str,
    actual_format: str | None,
    preprocess_chain: list | None,
    saved_intermediate_items: Iterable[tuple[IntermediateItem, str]] | None,
    options: dict | None,
    success: bool,
    message: str,
    output_path: str | None,
    mask_input: bool,
    schema_version: int = 1,
) -> dict:
    manifest = {
        "schema_version": schema_version,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "input": {
            "path": mask_input_path(file_path) if mask_input else file_path,
        },
        "actual_format": actual_format,
        "preprocess_chain": preprocess_chain,
        "intermediates": [
            {"kind": item.kind, "path": saved_path} for item, saved_path in (saved_intermediate_items or [])
        ],
        "options": {k: v for k, v in (options or {}).items() if k != "cancel_event"},
        "result": {
            "success": success,
            "message": message,
            "output_path": output_path,
        },
    }
    return manifest


# ==================== 格式到标准扩展名的映射 ====================

FORMAT_EXTENSION_MAP = {
    # 文档格式
    "doc": ".doc",
    "wps": ".doc",
    "docx": ".docx",
    "odt": ".odt",
    "rtf": ".rtf",
    # 表格格式
    "xls": ".xls",
    "et": ".xls",
    "xlsx": ".xlsx",
    "ods": ".ods",
    "csv": ".csv",
    # 版式格式
    "pdf": ".pdf",
    "ofd": ".ofd",
    # Markdown
    "md": ".md",
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
    standard_ext = FORMAT_EXTENSION_MAP.get(actual_format.lower(), f".{actual_format}")
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
    temp_file = str(Path(temp_dir) / f"input{std_ext}")

    logger.debug("准备输入文件")
    logger.debug(f"  原始: {Path(input_path).name}")
    logger.debug(f"  临时: {Path(temp_file).name}")

    shutil.copy2(input_path, temp_file)

    logger.info(f"✓ 输入文件已准备: {Path(temp_file).name}")
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
        报告.pdf → 报告_001.pdf → 报告_002.pdf ...
    """
    target = Path(filepath)
    if not target.exists():
        return str(target)

    directory = target.parent
    name = target.stem
    ext = target.suffix

    counter = 1
    while True:
        new_filename = f"{name}_{counter:03d}{ext}"
        new_filepath = directory / new_filename
        if not new_filepath.exists():
            logger.debug(f"生成唯一文件名: {new_filename}")
            return str(new_filepath)
        counter += 1


def _ensure_unique_path(path: str) -> str:
    target = Path(path)
    if not target.exists():
        return str(target)

    if target.is_dir():
        from docwen.utils.path_utils import ensure_unique_directory_path

        return ensure_unique_directory_path(str(target))

    return _generate_unique_filename(path)


def copy_item_to_unique_destination(source: str, destination: str) -> str | None:
    try:
        target_dir = Path(destination).parent
        if target_dir != Path():
            target_dir.mkdir(parents=True, exist_ok=True)

        final_destination = _ensure_unique_path(destination)
        if Path(source).is_dir():
            shutil.copytree(source, final_destination)
        else:
            shutil.copy2(source, final_destination)

        return final_destination
    except Exception as e:
        logger.error(f"复制失败: {e}")
        return None


def save_intermediate_item(source: str, output_dir: str, *, move: bool = False) -> str | None:
    destination = str(Path(output_dir) / Path(source).name)
    if move:
        return move_file_with_retry(source, destination)
    return copy_item_to_unique_destination(source, destination)


def save_intermediate_items(
    items: Iterable[IntermediateItem], output_dir: str, *, move: bool = False
) -> list[tuple[IntermediateItem, str]]:
    saved_items: list[tuple[IntermediateItem, str]] = []
    for item in items:
        filename = Path(item.path).name
        if filename.startswith("input."):
            continue
        saved_path = save_intermediate_item(item.path, output_dir, move=move)
        if saved_path:
            saved_items.append((item, saved_path))
    return saved_items


def collect_intermediate_items_from_dir(
    temp_dir: str,
    *,
    exclude_filenames: Iterable[str] | None = None,
    exclude_prefixes: tuple[str, ...] = ("input.",),
    include_dirs: bool = True,
    include_files: bool = True,
) -> list[IntermediateItem]:
    excluded = set(exclude_filenames or [])
    items: list[IntermediateItem] = []

    temp_dir_path = Path(temp_dir)
    if not temp_dir_path.is_dir():
        return items

    try:
        entries = list(temp_dir_path.iterdir())
    except Exception:
        return items

    for entry in entries:
        name = entry.name
        if name in excluded:
            continue
        if any(name.startswith(prefix) for prefix in exclude_prefixes):
            continue

        if entry.is_dir():
            if not include_dirs:
                continue
            kind = "intermediate_dir"
        else:
            if not include_files:
                continue
            ext = entry.suffix.lstrip(".").lower()
            kind = f"intermediate_{ext}" if ext else "intermediate_file"

        items.append(IntermediateItem(kind=kind, path=str(entry)))

    return items


def save_intermediates_from_temp_dir(
    temp_dir: str,
    output_dir: str,
    *,
    move: bool = True,
    exclude_filenames: Iterable[str] | None = None,
    exclude_prefixes: tuple[str, ...] = ("input.",),
    include_dirs: bool = True,
    include_files: bool = True,
) -> list[tuple[IntermediateItem, str]]:
    items = collect_intermediate_items_from_dir(
        temp_dir,
        exclude_filenames=exclude_filenames,
        exclude_prefixes=exclude_prefixes,
        include_dirs=include_dirs,
        include_files=include_files,
    )
    return save_intermediate_items(items, output_dir, move=move)


def move_file_with_retry(source: str, destination: str, max_retries: int = 3, retry_delay: float = 0.5) -> str | None:
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
        Optional[str]: 实际目标路径；失败时返回None
    """
    logger.debug(f"准备移动文件: {Path(source).name} → {destination}")

    for attempt in range(max_retries):
        try:
            # 确保目标目录存在
            destination_path = Path(destination)
            target_dir = destination_path.parent
            if target_dir != Path():
                target_dir.mkdir(parents=True, exist_ok=True)
                logger.debug(f"确保目标目录存在: {target_dir}")

            # 处理文件已存在的情况
            final_destination = destination
            if Path(final_destination).exists():
                logger.warning("目标文件已存在，生成新文件名")
                final_destination = _ensure_unique_path(final_destination)

            # 移动文件
            shutil.move(source, final_destination)
            logger.info(f"✓ 文件移动成功: {Path(final_destination).name}")
            return final_destination

        except PermissionError as e:
            logger.warning(f"权限错误（尝试 {attempt + 1}/{max_retries}）: {e}")
            if attempt < max_retries - 1:
                time.sleep(retry_delay)

        except Exception as e:
            logger.error(f"移动文件失败: {e}")
            if attempt < max_retries - 1:
                time.sleep(retry_delay)

    logger.error(f"✗ 文件移动失败（已重试{max_retries}次）")
    return None


def replace_file_atomic(
    temp_file: str,
    target_path: str,
    *,
    create_backup: bool = False,
) -> str:
    target = Path(target_path)
    target_dir = target.parent
    if target_dir != Path():
        target_dir.mkdir(parents=True, exist_ok=True)

    backup_path = None
    if create_backup and target.exists():
        backup_path = _ensure_unique_path(f"{target}.bak")
        shutil.copy2(str(target), backup_path)

    try:
        Path(temp_file).replace(target)
        return str(target)
    except Exception:
        if backup_path and Path(backup_path).exists():
            with contextlib.suppress(Exception):
                Path(backup_path).unlink()
        raise


def save_output_with_fallback(
    temp_file: str, target_path: str, original_input_file: str, max_retries: int = 3, retry_delay: float = 0.5
) -> tuple[str | None, str]:
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
    logger.info(f"开始保存文件: {Path(temp_file).name}")

    # 尝试1：目标位置
    logger.debug(f"尝试保存到目标位置: {target_path}")
    saved_path = move_file_with_retry(temp_file, target_path, max_retries, retry_delay)
    if saved_path:
        logger.info("✓ 文件已保存到目标位置")
        return saved_path, "目标位置"

    # 尝试2：原文件所在目录
    source_dir = Path(original_input_file).parent
    source_dir_path = str(source_dir / Path(target_path).name)

    logger.warning("目标位置保存失败，尝试备选位置1：原文件所在目录")
    saved_path = move_file_with_retry(temp_file, source_dir_path, max_retries, retry_delay)
    if saved_path:
        logger.info(f"✓ 文件已保存到原文件目录: {source_dir}")
        return saved_path, "原文件所在目录（目标位置不可写）"

    # 尝试3：用户桌面
    try:
        desktop = Path.home() / "Desktop"
        if desktop.exists():
            desktop_path = str(desktop / Path(target_path).name)

            logger.warning("原文件目录保存失败，尝试备选位置2：用户桌面")
            saved_path = move_file_with_retry(temp_file, desktop_path, max_retries, retry_delay)
            if saved_path:
                logger.info(f"✓ 文件已保存到桌面: {desktop}")
                return saved_path, "桌面（原位置和目标位置均不可写）"
    except Exception as e:
        logger.warning(f"尝试保存到桌面失败: {e}")

    # 尝试4：临时救援目录
    try:
        rescue_dir = Path(tempfile.gettempdir()) / "docwen_rescue"
        rescue_dir.mkdir(parents=True, exist_ok=True)
        rescue_path = str(rescue_dir / Path(target_path).name)

        logger.error("所有常规位置保存失败，使用临时救援目录")
        rescue_path = _ensure_unique_path(rescue_path)
        if Path(temp_file).is_dir():
            shutil.copytree(temp_file, rescue_path)
            with contextlib.suppress(Exception):
                shutil.rmtree(temp_file, ignore_errors=True)
        else:
            shutil.copy2(temp_file, rescue_path)
            with contextlib.suppress(Exception):
                Path(temp_file).unlink()
        logger.warning(f"⚠ 文件已复制到临时救援目录: {rescue_dir}")

        return rescue_path, f"临时救援目录（所有位置均不可写）\n路径: {rescue_dir}"

    except Exception as e:
        logger.critical(f"复制到救援目录失败: {e}")
        raise RuntimeError("复制到救援目录失败，无法保存输出文件") from e


def finalize_output(
    temp_item: str,
    target_path: str,
    *,
    original_input_file: str,
    max_retries: int = 3,
    retry_delay: float = 0.5,
) -> tuple[str | None, str]:
    return save_output_with_fallback(
        temp_item,
        target_path,
        original_input_file=original_input_file,
        max_retries=max_retries,
        retry_delay=retry_delay,
    )


def write_temp_file_then_finalize(
    *,
    temp_dir: str,
    target_path: str,
    original_input_file: str,
    writer: Callable[[str], None],
    max_retries: int = 3,
    retry_delay: float = 0.5,
) -> tuple[str | None, str]:
    Path(temp_dir).mkdir(parents=True, exist_ok=True)
    temp_output_path = str(Path(temp_dir) / Path(target_path).name)
    writer(temp_output_path)
    return finalize_output(
        temp_output_path,
        target_path,
        original_input_file=original_input_file,
        max_retries=max_retries,
        retry_delay=retry_delay,
    )


# ==================== 配置读取辅助函数 ====================


def get_output_directory(input_file: str, custom_dir: str | None = None) -> str:
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
        from docwen.config.config_manager import config_manager

        output_settings = config_manager.get_output_directory_settings()

        mode = output_settings.get("mode", "source")
        logger.debug(f"配置的输出模式: {mode}")

        if mode == "custom":
            custom_path = output_settings.get("custom_path", "")
            if custom_path:
                # 处理相对路径
                if not Path(custom_path).is_absolute():
                    try:
                        from docwen.utils.path_utils import get_project_root

                        custom_path = str(Path(get_project_root()) / custom_path)
                    except Exception:
                        pass

                base_dir = custom_path
                logger.debug(f"使用配置的自定义目录: {custom_path}")

    except Exception as e:
        logger.warning(f"读取输出目录配置失败: {e}")

    # 如果没有设置自定义目录，使用原文件所在目录
    if not base_dir:
        base_dir = str(Path(input_file).parent)
        logger.debug(f"使用原文件所在目录: {base_dir}")

    # 统一处理日期子文件夹（对两种模式都有效）
    if output_settings and output_settings.get("create_date_subfolder", False):
        date_format = output_settings.get("date_folder_format", "%Y-%m-%d")
        date_folder = datetime.now().strftime(date_format)
        final_path = str(Path(base_dir) / date_folder)
        logger.debug(f"创建日期子文件夹: {date_folder}")

        # 确保日期子文件夹存在
        Path(final_path).mkdir(parents=True, exist_ok=True)
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
        from docwen.config.config_manager import config_manager

        save_intermediate = config_manager.get_save_intermediate_files()

        logger.debug(f"中间文件保存配置: {save_intermediate}")
        return save_intermediate

    except Exception as e:
        logger.warning(f"读取中间文件配置失败，默认不保存: {e}")
        return False


# ==================== 模块导出 ====================

__all__ = [
    "finalize_output",
    "get_output_directory",
    "get_standard_extension",
    "move_file_with_retry",
    "prepare_input_file",
    "replace_file_atomic",
    "save_intermediate_item",
    "save_output_with_fallback",
    "should_save_intermediate_files",
    "write_temp_file_then_finalize",
]
