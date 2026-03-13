from __future__ import annotations

from collections.abc import Callable
import logging
from pathlib import Path
from typing import Any

from docwen.errors import DocWenError, InvalidInputError
from docwen.proofread_keys import PROOFREAD_OPTION_KEYS, SENSITIVE_WORD, SYMBOL_CORRECTION, SYMBOL_PAIRING, TYPOS_RULE
from .batch import BatchEvent, BatchResult, execute_batch
from .cancellation import CancellationToken
from .context import AppContext, get_default_context
from .requests import BatchRequest, ConversionRequest
from .error_codes import ERROR_CODE_SKIPPED_SAME_FORMAT, ERROR_CODE_UNKNOWN_ERROR
from .result import ConversionResult
from .result_presentation import normalize_result_error_fields


logger = logging.getLogger(__name__)


def _normalize_convert_action(action_type: str) -> tuple[str | None, str | None] | None:
    if not action_type:
        return None
    if "convert_" not in action_type or "_to_" not in action_type:
        return None
    try:
        _, rest = action_type.split("convert_", 1)
        src, tgt = rest.split("_to_", 1)
        src = src.strip().lower() or None
        tgt = tgt.strip().lower() or None
        if not tgt:
            return None
        return src, tgt
    except Exception:
        return None


def _merge_proofread_options(options: dict[str, Any], *, ctx: AppContext) -> dict[str, bool]:
    try:
        engine_config = ctx.config_manager.get_proofread_engine_config()
    except Exception:
        engine_config = {}

    defaults = {
        SYMBOL_PAIRING: bool(engine_config.get("enable_symbol_pairing", True)),
        SYMBOL_CORRECTION: bool(engine_config.get("enable_symbol_correction", True)),
        TYPOS_RULE: bool(engine_config.get("enable_typos_rule", True)),
        SENSITIVE_WORD: bool(engine_config.get("enable_sensitive_word", True)),
    }

    if "proofread_options" in options:
        provided = options.get("proofread_options")
        if isinstance(provided, dict):
            if not provided:
                return {}
            merged = defaults.copy()
            for key in PROOFREAD_OPTION_KEYS:
                if key in provided:
                    merged[key] = bool(provided.get(key))
            return merged

    if not any(key in options for key in PROOFREAD_OPTION_KEYS):
        return defaults

    merged = defaults.copy()
    for key in PROOFREAD_OPTION_KEYS:
        if key in options:
            merged[key] = bool(options.get(key))
    return merged


def _map_merge_tables_mode(value: Any) -> int:
    if isinstance(value, int):
        if value in (1, 2, 3):
            return value
        raise InvalidInputError("表格汇总模式不合法", details=str(value))
    if isinstance(value, str):
        v = value.strip().lower()
        mapping = {"row": 1, "col": 2, "cell": 3}
        if v in mapping:
            return mapping[v]
        raise InvalidInputError("表格汇总模式不合法", details=value)
    raise InvalidInputError("表格汇总模式不合法", details=str(value))


def _coerce_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, int) and value in (0, 1):
        return bool(value)
    if isinstance(value, str):
        v = value.strip().lower()
        if v in {"1", "true", "yes", "y", "on"}:
            return True
        if v in {"0", "false", "no", "n", "off", ""}:
            return False
    return bool(value)


def _normalize_md_options(options: dict[str, Any], *, category: str, ctx: AppContext) -> None:
    if "extract_image" not in options or options.get("extract_image") is None:
        options["extract_image"] = category in {"document", "markdown", "image"}
    else:
        options["extract_image"] = _coerce_bool(options.get("extract_image"))

    if "extract_ocr" not in options or options.get("extract_ocr") is None:
        options["extract_ocr"] = False
    else:
        options["extract_ocr"] = _coerce_bool(options.get("extract_ocr"))

    if "optimize_for_type" not in options or options.get("optimize_for_type") is None:
        options["optimize_for_type"] = ""
    else:
        options["optimize_for_type"] = str(options.get("optimize_for_type") or "")

    if "to_md_image_extraction_mode" not in options or options.get("to_md_image_extraction_mode") is None:
        try:
            options["to_md_image_extraction_mode"] = str(
                ctx.config_manager.get_export_to_md_image_extraction_mode() or "file"
            )
        except Exception:
            options["to_md_image_extraction_mode"] = "file"
    else:
        options["to_md_image_extraction_mode"] = str(options.get("to_md_image_extraction_mode") or "file")

    if "to_md_ocr_placement_mode" not in options or options.get("to_md_ocr_placement_mode") is None:
        try:
            options["to_md_ocr_placement_mode"] = str(
                ctx.config_manager.get_export_to_md_ocr_placement_mode() or "image_md"
            )
        except Exception:
            options["to_md_ocr_placement_mode"] = "image_md"
    else:
        options["to_md_ocr_placement_mode"] = str(options.get("to_md_ocr_placement_mode") or "image_md")

    if str(options.get("to_md_image_extraction_mode") or "").lower() == "base64":
        options["to_md_image_extraction_mode"] = "base64"
        options["to_md_ocr_placement_mode"] = "main_md"
    else:
        options["to_md_image_extraction_mode"] = str(options.get("to_md_image_extraction_mode") or "file").lower()
        options["to_md_ocr_placement_mode"] = str(options.get("to_md_ocr_placement_mode") or "image_md").lower()


def _normalize_numbering_options(
    options: dict[str, Any],
    *,
    category: str,
    target_format: str | None,
    action_type: str | None,
    ctx: AppContext,
) -> None:
    """
    将 CLI 传入的“模式参数”归一化为转换器所需的最终序号选项

    - 输入：clean_numbering / add_numbering_mode（可选）
    - 输出：remove_numbering / add_numbering / numbering_scheme
    - GUI 直设最终键时，旁路不干预

    Args:
        options: 可变 options 字典（就地修改）
        category: 输入文件类别（document/markdown/...）
        target_format: 目标格式（convert 场景）
        action_type: 动作类型（如 process_md_numbering）
        ctx: 应用上下文（用于读取配置与序号方案）
    """
    if any(k in options for k in ("remove_numbering", "add_numbering", "numbering_scheme")):
        logger.debug("序号归一化：检测到GUI已设最终序号键，跳过归一化")
        return

    user_provided = ("clean_numbering" in options) or ("add_numbering_mode" in options)
    clean_mode = str(options.get("clean_numbering") or "default").strip().lower()
    add_mode = str(options.get("add_numbering_mode") or "default").strip()

    ctx_pair: tuple[str, str] | None = None
    tf = str(target_format or "").strip().lower()
    if action_type == "process_md_numbering":
        ctx_pair = ("text", "to_docx")
    elif action_type is None:
        if category == "document" and tf == "md":
            ctx_pair = ("document", "to_md")
        elif category == "markdown" and tf == "docx":
            ctx_pair = ("text", "to_docx")
        elif category == "markdown" and tf == "xlsx":
            ctx_pair = ("text", "to_xlsx")

    if ctx_pair is None:
        if user_provided:
            logger.debug(
                f"序号归一化：当前输入类型不支持序号参数，category={category}, target_format={target_format}, action_type={action_type}"
            )
            raise InvalidInputError(
                "当前输入类型不支持序号参数",
                details={"category": category, "target_format": target_format, "action_type": action_type},
            )
        logger.debug(
            f"序号归一化：未命中支持场景且用户未显式传参，直接跳过，category={category}, target_format={target_format}, action_type={action_type}"
        )
        return

    if clean_mode not in {"default", "remove", "keep"}:
        logger.debug(f"序号归一化：清理序号模式不合法: {clean_mode}")
        raise InvalidInputError("清理序号模式不合法", details=clean_mode)

    section, prefix = ctx_pair
    defaults = ctx.config_manager.get_conversion_defaults(section)
    logger.debug(
        f"序号归一化：命中场景 section={section}, prefix={prefix}, clean_mode={clean_mode}, add_mode={add_mode}"
    )

    if clean_mode == "default":
        remove_numbering = _coerce_bool(defaults.get(f"{prefix}_remove_numbering", False))
    elif clean_mode == "remove":
        remove_numbering = True
    else:
        remove_numbering = False

    add_mode_lower = add_mode.lower()
    if add_mode_lower == "default":
        add_numbering = _coerce_bool(defaults.get(f"{prefix}_add_numbering", False))
        scheme_id = (
            str(defaults.get(f"{prefix}_default_scheme", "gongwen_standard") or "gongwen_standard").strip().lower()
        )
    elif add_mode_lower == "none":
        add_numbering = False
        scheme_id = ""
    else:
        add_numbering = True
        scheme_id = add_mode.strip().lower()

    if add_numbering:
        schemes = ctx.config_manager.get_heading_schemes()
        if scheme_id not in schemes:
            from difflib import get_close_matches

            keyword_matches = get_close_matches(scheme_id, ["default", "none"], n=1, cutoff=0.8)
            if keyword_matches:
                logger.debug(
                    f"序号归一化：疑似关键字拼写错误 add_numbering={add_mode}, did_you_mean={keyword_matches[0]}"
                )
                raise InvalidInputError(
                    "新增序号参数疑似拼写错误",
                    details={"add_numbering": add_mode, "did_you_mean": keyword_matches[0]},
                )

            matches = get_close_matches(scheme_id, list(schemes.keys()), n=5, cutoff=0.5)
            details: dict[str, Any] = {"scheme_id": scheme_id}
            if matches:
                details["close_matches"] = matches
            logger.debug(
                f"序号归一化：序号方案不存在 scheme_id={scheme_id}, close_matches={details.get('close_matches')}"
            )
            raise InvalidInputError("序号方案不存在", details=details)

    options.pop("clean_numbering", None)
    options.pop("add_numbering_mode", None)
    options["remove_numbering"] = bool(remove_numbering)
    options["add_numbering"] = bool(add_numbering)
    options["numbering_scheme"] = scheme_id if add_numbering else ""
    logger.debug(
        f"序号归一化：归一化完成 remove_numbering={options['remove_numbering']}, add_numbering={options['add_numbering']}, numbering_scheme={options['numbering_scheme']}"
    )


class ConversionService:
    def __init__(self, *, ctx: AppContext | None = None) -> None:
        self._ctx = ctx or get_default_context()

    def execute(
        self,
        request: ConversionRequest,
        *,
        progress_callback: Callable[[str], None] | None = None,
        cancel_token: CancellationToken | None = None,
    ) -> ConversionResult:
        cancel_token = cancel_token or CancellationToken()
        try:
            if request.action_type:
                request.action_type = request.action_type.strip()
            normalized_action = request.action_type if request.action_type else None
            original_action = normalized_action
            options: dict[str, Any] = dict(request.options or {})
            if request.file_list is None and isinstance(options.get("file_list"), list):
                request.file_list = list(options.get("file_list") or [])

            if normalized_action == "convert":
                target = options.get("target_format")
                request.action_type = None
                request.target_format = str(target).lower() if target else None
                normalized_action = None

            convert_pair = _normalize_convert_action(normalized_action or "")
            if convert_pair:
                src, tgt = convert_pair
                request.action_type = None
                request.source_format = request.source_format or src
                request.target_format = request.target_format or tgt
                normalized_action = None

            request.validate()
            file_path = request.file_path
            if file_path is None:
                raise InvalidInputError("缺少输入文件路径")

            if request.actual_format:
                actual_format = request.actual_format
            else:
                try:
                    actual_format = self._ctx.detect_actual_file_format(file_path)
                except Exception:
                    ext = Path(file_path).suffix.lower().lstrip(".")
                    if ext == "markdown":
                        ext = "md"
                    actual_format = ext or "unknown"
                request.actual_format = actual_format

            if request.category:
                category = request.category
            else:
                try:
                    category = self._ctx.get_actual_file_category(file_path, actual_format)
                except Exception:
                    from docwen.formats import category_from_actual_format

                    category = category_from_actual_format(actual_format)
                request.category = category

            options.setdefault("headless", True)
            options["actual_format"] = actual_format
            options.setdefault("cancel_event", cancel_token.event)

            if request.action_type == "merge_tables" and "mode" in options:
                options["mode"] = _map_merge_tables_mode(options.get("mode"))

            if request.file_list:
                options.setdefault("file_list", list(request.file_list))
                options.setdefault("selected_file", request.file_path)

            if category in {"markdown", "document"}:
                options["proofread_options"] = _merge_proofread_options(options, ctx=self._ctx)

            _normalize_numbering_options(
                options,
                category=category,
                target_format=request.target_format,
                action_type=request.action_type,
                ctx=self._ctx,
            )

            if request.action_type is None:
                target_format = request.target_format
                if target_format == "md":
                    source_format = (request.source_format or category or "").strip().lower() or None
                    request.source_format = source_format
                else:
                    source_format = (request.source_format or actual_format or "").strip().lower() or None
                if not source_format:
                    raise InvalidInputError("无法确定源格式", details="source_format")

                if target_format == "md":
                    _normalize_md_options(options, category=category, ctx=self._ctx)

                is_same_format = actual_format and target_format and actual_format.lower() == target_format.lower()
                is_image_compress = category == "image" and (
                    options.get("size_limit") is not None or options.get("limit_size") is not None
                )
                if is_same_format and not is_image_compress:
                    return ConversionResult.ok(
                        error_code=ERROR_CODE_SKIPPED_SAME_FORMAT,
                        details=(actual_format or "").upper() or None,
                    )

                strategy_class = self._ctx.get_strategy(
                    action_type=None,
                    source_format=source_format,
                    target_format=target_format,
                )
            else:
                strategy_class = self._ctx.get_strategy(
                    action_type=request.action_type, source_format=None, target_format=None
                )

            strategy = strategy_class()
            result = strategy.execute(
                file_path=request.file_path,
                options=options,
                progress_callback=progress_callback,
            )
            return normalize_result_error_fields(result)
        except DocWenError as e:
            return normalize_result_error_fields(
                ConversionResult.fail(
                    message=e.user_message,
                    error=e,
                    error_code=e.code,
                    details=e.details,
                )
            )
        except Exception as e:
            return normalize_result_error_fields(
                ConversionResult.fail(
                    message=str(e),
                    error=e,
                    error_code=ERROR_CODE_UNKNOWN_ERROR,
                    details=str(e),
                )
            )

    def execute_batch(
        self,
        batch_request: BatchRequest,
        *,
        progress_callback: Callable[[str], None] | None = None,
        cancel_token: CancellationToken | None = None,
        event_sink: Callable[[BatchEvent], None] | None = None,
    ) -> BatchResult:
        cancel_token = cancel_token or CancellationToken()

        def _execute_one(req: ConversionRequest) -> ConversionResult:
            return self.execute(req, progress_callback=progress_callback, cancel_token=cancel_token)

        return execute_batch(
            batch_request,
            execute_one=_execute_one,
            cancel_token=cancel_token,
            event_sink=event_sink,
        )
