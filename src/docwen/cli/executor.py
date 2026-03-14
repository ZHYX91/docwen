"""
Strategy执行器模块

统一调用Strategy层，处理所有文件转换操作：
- 根据action和文件类型获取Strategy
- 执行Strategy并处理结果
- 支持JSON输出模式
- 批量执行支持
"""

import json
import logging
import sys
import time
import traceback
from collections.abc import Callable
from difflib import get_close_matches
from pathlib import Path

from docwen.cli.i18n import cli_t
from docwen.errors import DocWenError, ExitCode, InvalidInputError, exit_code_from_error_code
from docwen.services.error_codes import ERROR_CODE_INVALID_INPUT, ERROR_CODE_UNKNOWN_ERROR
from docwen.services.requests import BatchRequest, ConversionRequest
from docwen.services.result import ConversionResult
from docwen.services.result_presentation import (
    format_result_message,
    normalize_result_error_fields,
)
from docwen.services.strategies import get_strategy
from docwen.services.use_cases import ConversionService
from docwen.utils.file_type_utils import detect_actual_file_format, get_strategy_file_category

logger = logging.getLogger(__name__)

_IMAGE_TARGET_FORMATS: set[str] = {
    "jpg",
    "jpeg",
    "png",
    "gif",
    "bmp",
    "tiff",
    "tif",
    "webp",
    "heic",
    "heif",
}

_FORMAT_GROUPS: dict[str, set[str]] = {
    "document": {"md", "pdf", "ofd"},
    "spreadsheet": {"md", "pdf", "xlsx", "xls", "ods", "csv"},
    "layout": {"md", "pdf", "docx", "doc", "odt", "rtf", "png", "jpg", "jpeg", "tif", "tiff"},
    "image": {"md", "pdf", *_IMAGE_TARGET_FORMATS},
    "markdown": {"docx", "doc", "odt", "rtf", "xlsx", "xls", "ods", "csv"},
}


def _get_format_groups() -> dict[str, set[str]]:
    try:
        from docwen.formats import CATEGORY_UNKNOWN, get_strategy_category_from_format
        from docwen.services import strategies as strategies_pkg

        strategies_pkg.load_all()
        registry = strategies_pkg.get_conversion_registry()

        category_names = {"document", "spreadsheet", "layout", "image", "markdown"}
        groups: dict[str, set[str]] = {}
        for src, tgt in registry:
            src_norm = (src or "").strip().lower()
            if not src_norm:
                continue
            if src_norm in category_names:
                src_cat = src_norm
            else:
                src_cat = get_strategy_category_from_format(src_norm)
                if src_cat == CATEGORY_UNKNOWN:
                    continue
            tgt_norm = (tgt or "").strip().lower()
            if not tgt_norm:
                continue
            if tgt_norm in category_names:
                continue
            groups.setdefault(src_cat, set()).add(tgt_norm)

        return groups or _FORMAT_GROUPS
    except Exception:
        return _FORMAT_GROUPS


def get_strategy_for_action(action: str, file_path: str, options: dict | None = None):
    options = options or {}
    actual_format = options.get("actual_format")
    if not actual_format:
        actual_format = detect_actual_file_format(file_path)
    category = options.get("category") or get_strategy_file_category(file_path, actual_format)

    if action == "convert":
        target_format = options.get("target_format")
        if not target_format:
            raise InvalidInputError("需要指定目标格式")
        target_format = str(target_format).lower()
        if target_format == "md":
            return get_strategy(action_type=None, source_format=category, target_format="md")
        return get_strategy(action_type=None, source_format=actual_format, target_format=target_format)

    return get_strategy(action_type=action, source_format=None, target_format=None)


# ==================== 执行器 ====================


def _execute_action_core(
    action: str,
    file_path: str,
    options: dict | None = None,
    progress_callback: Callable[[str], None] | None = None,
) -> tuple[ConversionResult, float]:
    start_time = time.time()
    options = options or {}

    if not Path(file_path).exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")

    logger.info(f"执行操作: {action} on {file_path}")
    service = ConversionService()
    req = ConversionRequest(file_path=file_path, action_type=action, options=options)
    result = service.execute(req, progress_callback=progress_callback)
    duration = time.time() - start_time
    return result, duration


def execute_action(
    action: str,
    file_path: str,
    options: dict | None = None,
    json_mode: bool = False,
    progress_callback: Callable[[str], None] | None = None,
    *,
    include_timing: bool = False,
) -> int:
    """
    执行单个文件操作

    Args:
        action: 操作名称
        file_path: 文件路径
        options: 选项字典
        json_mode: 是否使用JSON输出
        progress_callback: 进度回调函数

    Returns:
        int: 退出码 (0=成功, 1=失败)
    """
    try:
        result, duration = _execute_action_core(
            action=action, file_path=file_path, options=options, progress_callback=progress_callback
        )
        result = normalize_result_error_fields(result)

        # 输出结果
        if json_mode:
            print_json_result(result, action, file_path, duration, include_timing=include_timing)
        else:
            print_text_result(result)

        if result.success:
            return int(ExitCode.OK)
        return int(exit_code_from_error_code(result.error_code))

    except FileNotFoundError as e:
        error_msg = str(e)
        if json_mode:
            print_json_error(action, file_path, error_msg, error_code=ERROR_CODE_INVALID_INPUT)
        else:
            print(f"错误: {error_msg}", file=sys.stderr)
        return int(ExitCode.INVALID_INPUT)

    except KeyboardInterrupt:
        logger.info("用户中断操作")
        if json_mode:
            print_json_error(action, file_path, "用户中断", details={"interrupted": True})
        else:
            print("\n操作已中断", file=sys.stderr)
        return 130  # SIGINT

    except DocWenError as e:
        logger.error(f"执行失败: {e}", exc_info=True)
        if json_mode:
            print_json_error(action, file_path, e.user_message, error_code=e.code, details=e.details)
        else:
            print(f"错误: {e}", file=sys.stderr)
        return int(exit_code_from_error_code(e.code))

    except Exception as e:
        logger.error(f"执行失败: {e}", exc_info=True)
        if json_mode:
            print_json_error(
                action,
                file_path,
                str(e),
                error_code=ERROR_CODE_UNKNOWN_ERROR,
                details={"traceback": traceback.format_exc()},
            )
        else:
            print(f"错误: {e}", file=sys.stderr)
        return int(ExitCode.UNKNOWN_ERROR)


def execute_batch(
    action: str,
    files: list[str],
    options: dict | None = None,
    json_mode: bool = False,
    continue_on_error: bool = False,
    progress_callback: Callable[[str], None] | None = None,
    *,
    max_workers: int | None = None,
    include_timing: bool = False,
) -> int:
    """
    批量执行操作

    Args:
        action: 操作名称
        files: 文件路径列表
        options: 选项字典
        json_mode: 是否使用JSON输出
        continue_on_error: 出错时是否继续
        progress_callback: 进度回调函数

    Returns:
        int: 退出码 (0=全部成功, 1=部分或全部失败)
    """
    requested_total = len(files)
    options_base = dict(options or {})
    options_base.setdefault("headless", True)
    results: list[dict] = []
    failed_files: list[str] = []
    interrupted = False
    batch_exit_code = ExitCode.OK

    aggregate_actions = {"merge_pdfs", "merge_images_to_tiff", "merge_tables"}
    service = ConversionService()

    if action in aggregate_actions and files:
        start_time = time.time()
        file_list = list(files)
        base_file = options_base.get("base_table") if action == "merge_tables" else None
        base_file = str(base_file) if base_file else file_list[0]
        if action == "merge_tables" and base_file not in file_list:
            file_list = [base_file, *file_list]
        req = ConversionRequest(file_path=base_file, action_type=action, file_list=file_list, options=options_base)
        result = normalize_result_error_fields(service.execute(req, progress_callback=progress_callback))
        duration = time.time() - start_time
        duration_out = round(duration, 2) if include_timing else 0.0

        per_file_success = bool(result.success)
        for f in files:
            if per_file_success:
                results.append(
                    {
                        "file": f,
                        "success": True,
                        "output_file": result.output_path,
                        "message": result.message,
                        "duration": duration_out,
                        "metadata": getattr(result, "metadata", None) or {},
                    }
                )
            else:
                failed_files.append(f)
                results.append(
                    {
                        "file": f,
                        "success": False,
                        "error": result.message or "操作失败",
                        "error_code": result.error_code,
                        "details": result.details,
                        "duration": duration_out,
                        "metadata": getattr(result, "metadata", None) or {},
                    }
                )

        success_count = requested_total if per_file_success else 0
        processed_count = requested_total
        if not per_file_success:
            batch_exit_code = exit_code_from_error_code(result.error_code)

        if json_mode:
            print_json_batch_summary(
                action,
                results,
                requested_total=requested_total,
                processed_total=processed_count,
                success=success_count,
                failed=failed_files,
                interrupted=interrupted,
            )
        else:
            print_text_batch_summary(
                requested_total=requested_total,
                processed_total=processed_count,
                success=success_count,
                failed=failed_files,
                interrupted=interrupted,
            )

        if success_count == requested_total:
            return int(ExitCode.OK)
        return int(batch_exit_code)

    try:
        batch_request = BatchRequest(
            requests=[ConversionRequest(file_path=f, action_type=action, options=dict(options_base)) for f in files],
            continue_on_error=bool(continue_on_error),
            max_workers=int(max_workers) if max_workers else 1,
        )
        batch_result = service.execute_batch(
            batch_request,
            progress_callback=progress_callback,
        )
    except KeyboardInterrupt:
        interrupted = True
        processed_total = 1 if files else 0
        results = (
            [
                {
                    "file": files[0],
                    "success": False,
                    "error": "用户中断",
                    "interrupted": True,
                    "duration": 0.0,
                    "metadata": {},
                }
            ]
            if files
            else []
        )
        failed_files = [files[0]] if files else []
        if json_mode:
            print_json_batch_summary(
                action,
                results,
                requested_total=requested_total,
                processed_total=processed_total,
                success=0,
                failed=failed_files,
                interrupted=True,
            )
        else:
            print_text_batch_summary(
                requested_total=requested_total,
                processed_total=processed_total,
                success=0,
                failed=failed_files,
                interrupted=True,
            )
        return 130

    success_count = 0
    processed_count = batch_result.processed_count
    for file_result in batch_result.results:
        if file_result.success:
            success_count += 1
            results.append(
                {
                    "file": file_result.file_path,
                    "success": True,
                    "output_file": file_result.output_path,
                    "message": file_result.message,
                    "duration": round(float(file_result.duration_s or 0.0), 2) if include_timing else 0.0,
                    "metadata": {},
                }
            )
        else:
            failed_files.append(file_result.file_path)
            results.append(
                {
                    "file": file_result.file_path,
                    "success": False,
                    "error": file_result.message or "操作失败",
                    "error_code": file_result.error_code,
                    "details": file_result.details,
                    "duration": round(float(file_result.duration_s or 0.0), 2) if include_timing else 0.0,
                    "metadata": {},
                }
            )
            mapped = exit_code_from_error_code(file_result.error_code)
            if mapped != ExitCode.UNKNOWN_ERROR:
                batch_exit_code = mapped
            elif batch_exit_code == ExitCode.OK:
                batch_exit_code = ExitCode.UNKNOWN_ERROR
            if not continue_on_error:
                break

    # 输出汇总
    if json_mode:
        print_json_batch_summary(
            action,
            results,
            requested_total=requested_total,
            processed_total=processed_count,
            success=success_count,
            failed=failed_files,
            interrupted=interrupted,
        )
    else:
        print_text_batch_summary(
            requested_total=requested_total,
            processed_total=processed_count,
            success=success_count,
            failed=failed_files,
            interrupted=interrupted,
        )

    if success_count == requested_total:
        return int(ExitCode.OK)
    return int(batch_exit_code)


# ==================== 结果输出 ====================


def print_text_result(result: ConversionResult):
    """打印文本格式的结果"""
    if result.success:
        default_msg = cli_t("cli.result.success", default="操作成功")
        print(f"\n✓ {result.message or default_msg}")
        if result.output_path:
            output_label = cli_t("cli.result.output", default="输出")
            print(f"{output_label}: {result.output_path}")
    else:
        default_msg = cli_t("cli.result.failed", default="操作失败")
        print(f"\n✗ {format_result_message(result, default_msg)}", file=sys.stderr)


def print_json_result(
    result: ConversionResult, action: str, input_file: str, duration: float, *, include_timing: bool = False
):
    """打印JSON格式的结果"""
    metadata = getattr(result, "metadata", None) or {}
    duration_out = round(duration, 2) if include_timing else 0.0
    data = {
        "action": action,
        "input_file": input_file,
        "output_file": result.output_path,
        "message": result.message,
        "error_code": result.error_code,
        "details": result.details,
        "duration": duration_out,
        "metadata": metadata,
    }
    output = make_json_envelope(
        command=action,
        success=bool(result.success),
        data=data,
        error_code=(result.error_code if not result.success else None),
        message=(result.message if not result.success else None),
        details=(result.details if not result.success else None),
    )
    print(json.dumps(output, ensure_ascii=False, indent=2))


def make_json_envelope(
    *,
    command: str,
    success: bool,
    data: dict | None = None,
    error_code: str | None = None,
    message: str | None = None,
    details: object | None = None,
    warnings: list[object] | None = None,
    timing: dict | None = None,
) -> dict:
    return {
        "schema_version": 2,
        "success": bool(success),
        "command": str(command),
        "data": data or {},
        "error": (
            None
            if success
            else {
                "error_code": error_code,
                "message": message,
                "details": details,
            }
        ),
        "warnings": list(warnings or []),
        "timing": timing,
    }


def print_json_error(
    command: str,
    input_file: str,
    error: str,
    *,
    error_code: str | None = None,
    details: object | None = None,
) -> None:
    """打印JSON格式的错误（v2 envelope）"""
    output = make_json_envelope(
        command=command,
        success=False,
        data={"input_file": input_file},
        error_code=error_code,
        message=error,
        details=details,
    )
    print(json.dumps(output, ensure_ascii=False, indent=2))


def print_text_batch_summary(
    requested_total: int, processed_total: int, success: int, failed: list[str], interrupted: bool = False
):
    """打印文本格式的批量操作汇总"""
    print(f"\n{'=' * 60}")
    batch_completed = cli_t(
        "cli.messages.batch_completed",
        default="批量操作完成: {success}/{total} 成功",
        success=success,
        total=processed_total,
    )
    print(batch_completed)

    if processed_total != requested_total:
        processed_line = cli_t(
            "cli.messages.batch_processed",
            default="已处理: {processed}/{total}",
            processed=processed_total,
            total=requested_total,
        )
        print(processed_line)

    if interrupted:
        interrupted_line = cli_t("cli.messages.batch_interrupted", default="状态: 已中断")
        print(interrupted_line)

    if failed:
        failed_msg = cli_t("cli.messages.failed_count", default="失败: {count} 个文件", count=len(failed))
        print(failed_msg)
        for file in failed:
            print(f"  ✗ {Path(file).name}")

    print("=" * 60)


def print_json_batch_summary(
    action: str,
    results: list[dict],
    requested_total: int,
    processed_total: int,
    success: int,
    failed: list[str],
    interrupted: bool = False,
):
    """打印JSON格式的批量操作汇总"""
    ok = (success == requested_total) and (processed_total == requested_total) and (not interrupted)
    data = {
        "action": action,
        "total": requested_total,
        "processed_count": processed_total,
        "success_count": success,
        "failed_count": len(failed),
        "interrupted": interrupted,
        "results": results,
    }
    error_message = None
    if not ok:
        error_message = "用户中断" if interrupted else f"{len(failed)}/{requested_total} 文件处理失败"
    output = make_json_envelope(command=str(action), success=ok, data=data, message=error_message)
    print(json.dumps(output, ensure_ascii=False, indent=2))


# ==================== AI功能：自描述 ====================


def inspect_file(file_path: str, json_mode: bool = True) -> int:
    """
    查询文件支持的操作（AI自描述）

    Args:
        file_path: 文件路径
        json_mode: 是否JSON输出

    Returns:
        int: 退出码
    """
    from docwen.cli.utils import detect_category, detect_format

    try:
        category = detect_category(file_path)
        fmt = detect_format(file_path)

        # 根据类别定义支持的操作
        actions = get_supported_actions(category, fmt)

        if json_mode:
            output = make_json_envelope(
                command="inspect",
                success=True,
                data={"file": file_path, "category": category, "format": fmt, "supported_actions": actions},
            )
            print(json.dumps(output, ensure_ascii=False, indent=2))
        else:
            file_label = cli_t("cli.messages.file_info", default="文件")
            category_label = cli_t("cli.messages.category_info", default="类别")
            format_label = cli_t("cli.messages.format_info", default="格式")
            actions_label = cli_t("cli.messages.supported_actions", default="支持的操作")

            print(f"\n{file_label}: {file_path}")
            print(f"{category_label}: {category}")
            print(f"{format_label}: {fmt}")
            print(f"\n{actions_label}:")
            for action in actions:
                print(f"  - {action['name']}: {action['description']}")

        return 0

    except Exception as e:
        logger.error(f"查询文件信息失败: {e}")
        if json_mode:
            print_json_error("inspect", file_path, str(e))
        else:
            error_label = cli_t("cli.messages.error_prefix", default="错误")
            print(f"{error_label}: {e}")
        return 1


def get_supported_actions(category: str, fmt: str) -> list[dict]:
    """
    获取文件类别支持的操作列表

    Args:
        category: 文件类别
        fmt: 文件格式

    Returns:
        List[Dict]: 操作列表
    """
    actions_by_category: dict[str, list[dict]] = {
        "markdown": [
            {
                "name": "convert",
                "description": "格式转换",
                "parameters": {
                    "target_format": {
                        "type": "choice",
                        "choices": ["docx", "doc", "odt", "rtf", "xlsx", "xls", "ods", "csv"],
                        "required": True,
                    },
                    "template": {"type": "string", "required": True, "description": "模板名称"},
                    "check": {
                        "type": "choice",
                        "choices": ["punct", "typo", "symbol", "sensitive", "all", "none"],
                        "repeatable": True,
                    },
                },
            },
        ],
        "document": [
            {
                "name": "convert",
                "description": "格式转换",
                "parameters": {
                    "target_format": {
                        "type": "choice",
                        "choices": ["md", "docx", "doc", "odt", "rtf"],
                        "required": True,
                    }
                },
            },
            {
                "name": "validate",
                "description": "文档校对",
                "parameters": {
                    "check_punct": {"type": "boolean", "default": False},
                    "check_typo": {"type": "boolean", "default": False},
                    "check_symbol": {"type": "boolean", "default": False},
                    "check_sensitive": {"type": "boolean", "default": False},
                    "check_none": {"type": "boolean", "default": False},
                },
            },
        ],
        "spreadsheet": [
            {
                "name": "convert",
                "description": "格式转换",
                "parameters": {
                    "target_format": {
                        "type": "choice",
                        "choices": ["md", "xlsx", "xls", "ods", "csv"],
                        "required": True,
                    }
                },
            },
            {
                "name": "merge_tables",
                "description": "表格汇总",
                "parameters": {
                    "mode": {"type": "choice", "choices": ["row", "col", "cell"], "required": True},
                    "base_table": {"type": "string", "description": "基准表格"},
                },
            },
        ],
        "layout": [
            {
                "name": "convert",
                "description": "格式转换",
                "parameters": {"target_format": {"type": "choice", "choices": ["md"], "required": True}},
            },
            {"name": "merge_pdfs", "description": "合并PDF"},
            {
                "name": "split_pdf",
                "description": "拆分PDF",
                "parameters": {"pages": {"type": "string", "required": True, "description": "页码范围"}},
            },
        ],
        "image": [
            {
                "name": "convert",
                "description": "格式转换",
                "parameters": {
                    "target_format": {
                        "type": "choice",
                        "choices": ["md", "png", "jpg", "bmp", "gif", "tif", "webp"],
                        "required": True,
                    }
                },
            },
        ],
    }
    return actions_by_category.get(category, [])


def list_all_actions(json_mode: bool = True) -> int:
    """
    列出所有可用操作

    Args:
        json_mode: 是否JSON输出

    Returns:
        int: 退出码
    """
    actions = [
        {
            "name": "convert",
            "description": cli_t("cli.actions.convert", default="格式转换"),
            "categories": ["document", "spreadsheet", "image", "layout"],
        },
        {
            "name": "validate",
            "description": cli_t("cli.actions.validate", default="文档校对"),
            "categories": ["document"],
        },
        {
            "name": "merge_tables",
            "description": cli_t("cli.actions.merge_tables", default="表格汇总"),
            "categories": ["spreadsheet"],
        },
        {
            "name": "merge_pdfs",
            "description": cli_t("cli.actions.merge_pdfs", default="合并PDF"),
            "categories": ["layout"],
        },
        {
            "name": "split_pdf",
            "description": cli_t("cli.actions.split_pdf", default="拆分PDF"),
            "categories": ["layout"],
        },
        {
            "name": "merge_images_to_tiff",
            "description": cli_t("cli.actions.merge_tiff", default="合并为TIF"),
            "categories": ["image"],
        },
        {
            "name": "process_md_numbering",
            "description": cli_t("cli.actions.process_numbering", default="处理MD小标题序号"),
            "categories": ["markdown"],
        },
    ]

    if json_mode:
        output = make_json_envelope(command="actions", success=True, data={"actions": actions})
        print(json.dumps(output, ensure_ascii=False, indent=2))
    else:
        actions_label = cli_t("cli.messages.available_actions", default="可用操作")
        print(f"\n{actions_label}:")
        for action in actions:
            cats = ", ".join(action["categories"])
            print(f"  {action['name']}: {action['description']} ({cats})")

    return 0


def list_numbering_schemes(json_mode: bool = True) -> int:
    """
    列出所有可用的序号方案

    Args:
        json_mode: 是否JSON输出

    Returns:
        int: 退出码
    """
    from docwen.config.config_manager import config_manager

    localized = config_manager.get_localized_numbering_schemes(include_description=True)
    schemes = [{"id": scheme_id, **info} for scheme_id, info in localized.items()]

    if json_mode:
        output = make_json_envelope(command="numbering-schemes", success=True, data={"schemes": schemes})
        print(json.dumps(output, ensure_ascii=False, indent=2))
    else:
        schemes_label = cli_t("cli.messages.available_schemes", default="可用序号方案")
        example_label = cli_t("cli.messages.example_format", default="示例")
        print(f"\n{schemes_label}:")
        for scheme in schemes:
            print(f"  {scheme['id']}: {scheme['name']}")
            if scheme.get("description"):
                print(f"    {example_label}: {scheme['description']}")

    return 0


def list_templates(json_mode: bool = True, target: str | None = None) -> int:
    """
    列出所有可用模板

    Args:
        json_mode: 是否JSON输出

    Returns:
        int: 退出码
    """
    try:
        from docwen.template.loader import TemplateLoader

        template_loader = TemplateLoader()

        # 获取文档模板和表格模板
        docx_templates = template_loader.get_available_templates("docx")
        xlsx_templates = template_loader.get_available_templates("xlsx")

        templates = {"docx": docx_templates or [], "xlsx": xlsx_templates or []}

        if target:
            target_lower = str(target).lower()
            if target_lower in {"docx", "doc"}:
                templates = {"docx": templates.get("docx") or [], "xlsx": []}
            elif target_lower in {"xlsx", "xls"}:
                templates = {"docx": [], "xlsx": templates.get("xlsx") or []}

        if json_mode:
            items: list[dict] = []
            for name in templates.get("docx") or []:
                items.append({"id": name, "name": name, "target": "docx", "description": None, "example": None})
            for name in templates.get("xlsx") or []:
                items.append({"id": name, "name": name, "target": "xlsx", "description": None, "example": None})
            output = make_json_envelope(command="templates", success=True, data={"templates": items})
            print(json.dumps(output, ensure_ascii=False, indent=2))
        else:
            templates_label = cli_t("cli.messages.available_templates", default="可用模板")
            docx_label = cli_t("cli.messages.document_templates", default="文档模板 (DOCX)")
            xlsx_label = cli_t("cli.messages.spreadsheet_templates", default="表格模板 (XLSX)")
            none_label = cli_t("cli.messages.no_templates", default="(无)")

            print(f"\n{templates_label}:")
            print(f"\n  {docx_label}:")
            if templates["docx"]:
                for name in templates["docx"]:
                    print(f"    • {name}")
            else:
                print(f"    {none_label}")

            print(f"\n  {xlsx_label}:")
            if templates["xlsx"]:
                for name in templates["xlsx"]:
                    print(f"    • {name}")
            else:
                print(f"    {none_label}")

        return 0

    except Exception as e:
        logger.error(f"获取模板列表失败: {e}")
        if json_mode:
            print_json_error("templates", "", str(e))
        else:
            error_msg = cli_t("cli.messages.template_list_error", default="错误: 无法获取模板列表")
            print(f"{error_msg} - {e}")
        return 1


def list_optimizations(json_mode: bool = True, scope: str | None = None) -> int:
    """
    列出当前语言下可用的优化类型

    Args:
        json_mode: 是否JSON输出
        scope: 作用域过滤（如 "document_to_md"、"layout_to_md"）

    Returns:
        int: 退出码
    """
    try:
        from docwen.config import config_manager
        from docwen.i18n import get_current_locale

        locale = get_current_locale()
        types = config_manager.get_optimization_types()
        settings = config_manager.get_optimization_settings()
        order = settings.get("order", list(types.keys()))

        optimizations: list[dict] = []
        for type_id in order:
            type_config = types.get(type_id)
            if not isinstance(type_config, dict):
                continue
            if not type_config.get("enabled", True):
                continue

            locales = type_config.get("locales", ["*"])
            if "*" not in locales and locale not in locales:
                continue

            if scope is not None:
                scopes = type_config.get("scopes", ["*"])
                if "*" not in scopes and scope not in scopes:
                    continue

            name = type_config.get("name", type_id)
            desc = type_config.get("description", "")
            scopes_out = type_config.get("scopes", ["*"])
            optimizations.append({"id": type_id, "name": name, "description": desc, "scopes": scopes_out})

        if json_mode:
            items = [{**opt, "examples": [], "recommended_for": []} for opt in optimizations]
            output = make_json_envelope(command="optimizations", success=True, data={"optimizations": items})
            print(json.dumps(output, ensure_ascii=False, indent=2))
        else:
            label = cli_t("cli.messages.available_optimizations", default="可用优化类型")
            none_label = cli_t("cli.messages.no_optimizations", default="(无)")
            print(f"\n{label}:")
            if optimizations:
                for opt in optimizations:
                    print(f"  • {opt['id']}: {opt['name']}")
                    if opt.get("description"):
                        print(f"    {opt['description']}")
            else:
                print(f"  {none_label}")

        return 0
    except Exception as e:
        logger.error(f"获取优化类型列表失败: {e}")
        if json_mode:
            print_json_error("optimizations", "", str(e))
        else:
            print(f"错误: {e}", file=sys.stderr)
        return 1


def get_supported_convert_targets() -> list[str]:
    targets: set[str] = set()
    for items in _get_format_groups().values():
        targets.update(items)
    return sorted({t.strip().lower() for t in targets if t and t.strip()})


def list_formats(json_mode: bool = True, source: str | None = None) -> int:
    try:
        groups = {k: set(v) for k, v in _get_format_groups().items()}
        if source:
            src_filter = (source or "").strip().lower()
            if src_filter in groups:
                groups = {src_filter: groups.get(src_filter) or set()}
            else:
                suggestions = get_close_matches(src_filter, sorted(groups.keys()), n=3, cutoff=0.2)
                hint = f"；可选：{', '.join(suggestions)}" if suggestions else ""
                raise InvalidInputError(f"无效类别: {source}{hint}")

        formats_out = [{"source": src, "targets": sorted(targets)} for src, targets in groups.items() if targets]

        if json_mode:
            output = make_json_envelope(command="formats", success=True, data={"formats": formats_out})
            print(json.dumps(output, ensure_ascii=False, indent=2))
        else:
            label = cli_t("cli.messages.available_formats", default="可用目标格式")
            print(f"\n{label}:")
            for item in formats_out:
                targets = ", ".join(item["targets"])
                print(f"  {item['source']} -> {targets}")

        return 0
    except DocWenError as e:
        if json_mode:
            print_json_error("formats", "", e.user_message, error_code=e.code, details=e.details)
        else:
            print(f"{cli_t('cli.messages.error_prefix')}: {e.user_message}", file=sys.stderr)
        return int(exit_code_from_error_code(e.code))
    except Exception as e:
        logger.error(f"获取格式列表失败: {e}")
        if json_mode:
            print_json_error("formats", "", str(e))
        else:
            print(f"{cli_t('cli.messages.error_prefix')}: {e}", file=sys.stderr)
        return 1
