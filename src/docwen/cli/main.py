"""
CLI主入口模块

提供命令行参数解析和程序入口：
- argparse参数定义
- 子命令（subcommands）路由
- 查询类命令（inspect/list 等）
- Windows控制台Unicode支持
"""

import argparse
import logging
import os
import sys
from difflib import get_close_matches

from docwen.cli.i18n import cli_t, init_cli_locale
from docwen.errors import ExitCode
from docwen.proofread_keys import SENSITIVE_WORD, SYMBOL_CORRECTION, SYMBOL_PAIRING, TYPOS_RULE

logger = logging.getLogger(__name__)


# ==================== 控制台编码 ====================


def setup_console_encoding():
    """
    设置Windows控制台编码为UTF-8

    解决Windows cmd/PowerShell中文乱码问题：
    1. 设置控制台代码页为65001 (UTF-8)
    2. 重新配置stdout/stderr的编码
    """
    if sys.platform != "win32":
        return

    # 设置控制台代码页为UTF-8
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        kernel32.SetConsoleOutputCP(65001)
        kernel32.SetConsoleCP(65001)
        logger.debug("已设置控制台代码页为UTF-8")
    except Exception as e:
        logger.debug(f"设置控制台代码页失败: {e}")

    # 重新配置stdout/stderr编码
    try:
        import io

        stdout_reconfigure = getattr(sys.stdout, "reconfigure", None)
        stderr_reconfigure = getattr(sys.stderr, "reconfigure", None)

        if callable(stdout_reconfigure) and callable(stderr_reconfigure):
            stdout_reconfigure(encoding="utf-8", errors="replace")
            stderr_reconfigure(encoding="utf-8", errors="replace")
            logger.debug("已重新配置stdout/stderr编码为UTF-8")
        elif isinstance(sys.stdout, io.TextIOWrapper) and isinstance(sys.stderr, io.TextIOWrapper):
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True)
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace", line_buffering=True)
            logger.debug("已包装stdout/stderr为UTF-8 TextIOWrapper")
    except Exception as e:
        logger.debug(f"重新配置stdout/stderr失败: {e}")


# ==================== 语言预解析 ====================


def get_available_locale_codes() -> list[str]:
    """
    获取当前可用的语言代码列表

    用于：
    - argparse choices
    - 预解析 --lang 以便 help 文本使用正确语言
    """
    try:
        from pathlib import Path

        locales_dir = Path(__file__).resolve().parents[1] / "i18n" / "locales"
        codes = sorted({p.stem for p in locales_dir.glob("*.toml") if p.stem})
        return list(codes) if codes else ["zh_CN", "en_US"]
    except Exception:
        return ["zh_CN", "en_US"]


def pre_parse_lang() -> str | None:
    """
    预解析 --lang 参数

    在创建 argparse 之前先获取语言设置，
    以便 help 文本可以使用正确的语言。

    Returns:
        Optional[str]: 语言代码或 None
    """
    available = set(get_available_locale_codes())
    for i, arg in enumerate(sys.argv):
        if arg == "--lang" and i + 1 < len(sys.argv):
            lang = sys.argv[i + 1]
            if lang in available:
                return lang
        elif arg.startswith("--lang="):
            lang = arg.split("=", 1)[1]
            if lang in available:
                return lang
    return None


# ==================== 参数解析器 ====================


def create_argument_parser() -> argparse.ArgumentParser:
    """
    创建命令行参数解析器（已国际化）

    所有帮助文本使用 cli_t() 获取翻译，
    支持 --lang 参数切换语言。

    Returns:
        argparse.ArgumentParser: 配置好的解析器
    """
    # ========== 通用选项 ==========
    # 通过 parents 让通用选项既可放在子命令前，也可放在子命令后：
    #   docwen --json convert --to md file.docx
    #   docwen convert --to md file.docx --json
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--lang", choices=get_available_locale_codes(), help=cli_t("cli.help.lang"))
    common.add_argument("--json", action="store_true", help=cli_t("cli.help.json"))
    common.add_argument("--quiet", "-q", action="store_true", help=cli_t("cli.help.quiet"))
    common.add_argument("--verbose", "-v", action="store_true", help=cli_t("cli.help.verbose"))
    common.add_argument(
        "--timing",
        action="store_true",
        help=cli_t("cli.help.timing", default="输出每个文件的耗时字段（JSON 中 duration）"),
    )
    common.add_argument("--batch", action="store_true", help=cli_t("cli.help.batch"))
    common.add_argument("--yes", "-y", action="store_true", help=cli_t("cli.help.yes"))
    common.add_argument("--continue-on-error", action="store_true", help=cli_t("cli.help.continue_on_error"))
    common.add_argument("--jobs", type=int, default=1, help=cli_t("cli.help.jobs", default="批处理并发数（默认：1）"))

    parser = argparse.ArgumentParser(
        prog="docwen",
        parents=[common],
        description=cli_t("cli.description"),
        epilog=f"{cli_t('cli.example')}: %(prog)s convert --to md document.docx --extract-img",
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    # ---------------------------------------------------------------------
    # 转换/处理类命令
    # ---------------------------------------------------------------------
    convert = subparsers.add_parser("convert", parents=[common], help=cli_t("cli.actions.convert"))
    convert.add_argument("files", nargs="+", help=cli_t("cli.help.files"))
    convert.add_argument("--to", required=True, help=cli_t("cli.help.target"))
    convert.add_argument("--template", help=cli_t("cli.help.template"))
    convert.add_argument("--extract-img", action="store_true", help=cli_t("cli.help.extract_img"))
    convert.add_argument(
        "--no-extract-img",
        action="store_true",
        help=cli_t("cli.help.no_extract_img", default="不提取图片"),
    )
    convert.add_argument("--ocr", action="store_true", help=cli_t("cli.help.ocr"))
    convert.add_argument("--optimize-for", help=cli_t("cli.help.optimize_for"))
    convert.add_argument(
        "--clean-numbering",
        choices=["default", "remove", "keep"],
        default=None,
        help=cli_t("cli.help.clean_numbering"),
    )
    convert.add_argument("--add-numbering", default=None, help=cli_t("cli.help.add_numbering"))

    validate = subparsers.add_parser("validate", parents=[common], help=cli_t("cli.actions.validate"))
    validate.add_argument("files", nargs="+", help=cli_t("cli.help.files"))
    validate.add_argument(
        "--check",
        action="append",
        choices=["punct", "typo", "symbol", "sensitive", "all", "none"],
        help=cli_t("cli.help.check_punct"),
    )

    merge_tables = subparsers.add_parser("merge-tables", parents=[common], help=cli_t("cli.actions.merge_tables"))
    merge_tables.add_argument("files", nargs="+", help=cli_t("cli.help.files"))
    merge_tables.add_argument(
        "--mode", required=True, choices=["row", "col", "cell"], help=cli_t("cli.help.merge_mode")
    )
    merge_tables.add_argument("--base-table", help=cli_t("cli.help.base_table"))

    merge_pdfs = subparsers.add_parser("merge-pdfs", parents=[common], help=cli_t("cli.actions.merge_pdfs"))
    merge_pdfs.add_argument("files", nargs="+", help=cli_t("cli.help.files"))

    split_pdf = subparsers.add_parser("split-pdf", parents=[common], help=cli_t("cli.actions.split_pdf"))
    split_pdf.add_argument("file", help=cli_t("cli.help.file"))
    split_pdf.add_argument("--pages", required=True, help=cli_t("cli.help.pages"))
    split_pdf.add_argument("--dpi", type=int, choices=[150, 300, 600], help=cli_t("cli.help.dpi"))

    merge_tiff = subparsers.add_parser("merge-images-to-tiff", parents=[common], help=cli_t("cli.actions.merge_tiff"))
    merge_tiff.add_argument("files", nargs="+", help=cli_t("cli.help.files"))
    merge_tiff.add_argument("--keep-alpha", action="store_true", help=cli_t("cli.help.keep_alpha"))
    merge_tiff.add_argument("--compress", choices=["lossless", "limit_size"], help=cli_t("cli.help.compress"))
    merge_tiff.add_argument("--size-limit", type=int, help=cli_t("cli.help.size_limit"))
    merge_tiff.add_argument("--quality-mode", choices=["original", "a4", "a3"], help=cli_t("cli.help.quality_mode"))

    md_numbering = subparsers.add_parser("md-numbering", parents=[common], help=cli_t("cli.actions.process_numbering"))
    md_numbering.add_argument("files", nargs="+", help=cli_t("cli.help.files"))
    md_numbering.add_argument(
        "--clean-numbering",
        choices=["default", "remove", "keep"],
        default=None,
        help=cli_t("cli.help.clean_numbering"),
    )
    md_numbering.add_argument("--add-numbering", default=None, help=cli_t("cli.help.add_numbering"))

    # ---------------------------------------------------------------------
    # 查询类命令
    # ---------------------------------------------------------------------
    inspect = subparsers.add_parser("inspect", parents=[common], help=cli_t("cli.help.inspect"))
    inspect.add_argument("file", help=cli_t("cli.help.files"))

    subparsers.add_parser("doctor", parents=[common], help=cli_t("cli.help.doctor"))

    actions = subparsers.add_parser("actions", parents=[common], help=cli_t("cli.help.list_actions"))
    actions_sub = actions.add_subparsers(dest="actions_command", required=True)
    actions_sub.add_parser("list", parents=[common], help=cli_t("cli.help.list_actions"))

    numbering = subparsers.add_parser(
        "numbering-schemes", parents=[common], help=cli_t("cli.help.list_numbering_schemes")
    )
    numbering_sub = numbering.add_subparsers(dest="numbering_command", required=True)
    numbering_sub.add_parser("list", parents=[common], help=cli_t("cli.help.list_numbering_schemes"))

    templates = subparsers.add_parser("templates", parents=[common], help=cli_t("cli.help.list_templates"))
    templates_sub = templates.add_subparsers(dest="templates_command", required=True)
    templates_list = templates_sub.add_parser("list", parents=[common], help=cli_t("cli.help.list_templates"))
    templates_list.add_argument("--for", dest="for_target", choices=["docx", "xlsx"], help=cli_t("cli.help.target"))

    optimizations = subparsers.add_parser("optimizations", parents=[common], help=cli_t("cli.help.optimize_for"))
    optimizations_sub = optimizations.add_subparsers(dest="optimizations_command", required=True)
    optimizations_list = optimizations_sub.add_parser("list", parents=[common], help=cli_t("cli.help.optimize_for"))
    optimizations_list.add_argument(
        "--scope",
        choices=["document_to_md", "layout_to_md", "image_to_md", "spreadsheet_to_md"],
        help=cli_t("cli.help.scope"),
    )

    formats = subparsers.add_parser(
        "formats", parents=[common], help=cli_t("cli.help.list_formats", default="列出可用目标格式")
    )
    formats_sub = formats.add_subparsers(dest="formats_command", required=True)
    formats_list = formats_sub.add_parser(
        "list", parents=[common], help=cli_t("cli.help.list_formats", default="列出可用目标格式")
    )
    formats_list.add_argument(
        "--for",
        dest="for_source",
        choices=["document", "spreadsheet", "layout", "image", "markdown"],
        help=cli_t("cli.help.format_source", default="按源类别过滤（document/spreadsheet/layout/image/markdown）"),
    )

    return parser


def _run_action_for_files(action: str, raw_files: list[str], options: dict, args) -> int:
    from docwen.cli import executor, utils

    files = utils.expand_paths(raw_files)
    if not files:
        if getattr(args, "json", False):
            from docwen.cli.executor import print_json_error
            from docwen.services.error_codes import ERROR_CODE_INVALID_INPUT

            print_json_error(
                action,
                "",
                cli_t("cli.messages.error_file_not_found"),
                error_code=ERROR_CODE_INVALID_INPUT,
            )
            return int(ExitCode.INVALID_INPUT)
        print(cli_t("cli.messages.error_file_not_found"), file=sys.stderr)
        return int(ExitCode.INVALID_INPUT)

    valid_files, invalid_files, _warning_files = utils.validate_files(files)

    if invalid_files and not getattr(args, "quiet", False):
        out = sys.stderr if getattr(args, "json", False) else sys.stdout
        print(cli_t("cli.messages.warning_invalid_files", count=len(invalid_files)), file=out)
        for file, reason in invalid_files:
            print(f"  - {file}: {reason}", file=out)

    if not valid_files:
        if getattr(args, "json", False):
            from docwen.cli.executor import print_json_error
            from docwen.services.error_codes import ERROR_CODE_INVALID_INPUT

            print_json_error(
                action,
                "",
                cli_t("cli.messages.error_no_valid_files"),
                error_code=ERROR_CODE_INVALID_INPUT,
            )
            return int(ExitCode.INVALID_INPUT)
        print(cli_t("cli.messages.error_no_valid_files"), file=sys.stderr)
        return int(ExitCode.INVALID_INPUT)

    progress_callback = utils.create_progress_callback(
        quiet=getattr(args, "quiet", False), verbose=getattr(args, "verbose", False)
    )
    include_timing = bool(getattr(args, "timing", False))

    if len(valid_files) == 1 and not getattr(args, "batch", False):
        return executor.execute_action(
            action,
            valid_files[0],
            options,
            json_mode=getattr(args, "json", False),
            progress_callback=progress_callback,
            include_timing=include_timing,
        )

    return executor.execute_batch(
        action,
        valid_files,
        options,
        json_mode=getattr(args, "json", False),
        continue_on_error=bool(getattr(args, "continue_on_error", False)),
        progress_callback=progress_callback,
        max_workers=min(
            max(1, int(getattr(args, "jobs", 1) or 1)),
            min(4, max(1, (os.cpu_count() or 2))),
        ),
        include_timing=include_timing,
    )


def _print_invalid_input(action: str, args, message: str, input_file: str = "") -> int:
    prefix = cli_t("cli.messages.error_prefix")
    out_message = message if str(message).startswith(str(prefix)) else f"{prefix}: {message}"
    if getattr(args, "json", False):
        from docwen.cli.executor import print_json_error
        from docwen.services.error_codes import ERROR_CODE_INVALID_INPUT

        print_json_error(
            action,
            input_file,
            out_message,
            error_code=ERROR_CODE_INVALID_INPUT,
        )
        return int(ExitCode.INVALID_INPUT)
    print(out_message, file=sys.stderr)
    return int(ExitCode.INVALID_INPUT)


def build_options(action: str, args) -> dict:
    """
    从命令行参数构建选项字典

    Args:
        action: 动作名称（与 services/use_cases 保持一致）
        args: 解析后的参数（子命令解析结果）

    Returns:
        dict: 选项字典
    """
    options = {}

    # convert 目标格式
    to = getattr(args, "to", None)
    if to:
        options["target_format"] = str(to).lower()

    # 导出选项
    extract_img = bool(getattr(args, "extract_img", False))
    no_extract_img = bool(getattr(args, "no_extract_img", False))
    if extract_img and no_extract_img:
        raise ValueError("--extract-img 不能与 --no-extract-img 同时使用")
    if no_extract_img:
        options["extract_image"] = False
    elif extract_img:
        options["extract_image"] = True
    if getattr(args, "ocr", False):
        options["extract_ocr"] = True
    optimize_for = getattr(args, "optimize_for", None)
    if optimize_for:
        options["optimize_for_type"] = str(optimize_for)
    else:
        options["optimize_for_type"] = None

    # 模板
    template = getattr(args, "template", None)
    if template:
        options["template_name"] = str(template)

    checks = getattr(args, "check", None) or []
    if "none" in checks and len(checks) > 1:
        raise ValueError("--check none 不能与其它 --check 同时使用")

    if checks:
        if "none" in checks:
            options["proofread_options"] = {}
        else:
            enabled = set(checks)
            if "all" in enabled:
                enabled.update({"punct", "typo", "symbol", "sensitive"})
            options["proofread_options"] = {
                SYMBOL_PAIRING: ("punct" in enabled),
                SYMBOL_CORRECTION: ("symbol" in enabled),
                TYPOS_RULE: ("typo" in enabled),
                SENSITIVE_WORD: ("sensitive" in enabled),
            }

    # 表格汇总
    merge_mode = getattr(args, "mode", None) or getattr(args, "merge_mode", None)
    if merge_mode:
        options["mode"] = str(merge_mode)
    base_table = getattr(args, "base_table", None)
    if base_table:
        options["base_table"] = str(base_table)

    # PDF选项
    pages = getattr(args, "pages", None)
    if pages:
        options["pages"] = str(pages)
    dpi = getattr(args, "dpi", None)
    if dpi:
        options["dpi"] = dpi

    # 图片选项
    compress = getattr(args, "compress", None)
    if compress:
        options["compress_mode"] = str(compress)
    size_limit = getattr(args, "size_limit", None)
    if size_limit:
        options["size_limit"] = size_limit
    quality_mode = getattr(args, "quality_mode", None)
    if quality_mode:
        options["quality_mode"] = str(quality_mode)
    if getattr(args, "keep_alpha", False):
        options["keep_alpha"] = True

    # 序号处理选项
    clean_mode = getattr(args, "clean_numbering", None)
    if clean_mode is not None:
        options["clean_numbering"] = str(clean_mode).strip().lower()

    add_mode = getattr(args, "add_numbering", None)
    if add_mode is not None:
        options["add_numbering_mode"] = str(add_mode).strip()

    return options


# ==================== 主入口 ====================


def main() -> int:
    """
    CLI主入口

    执行流程：
    1. 设置控制台编码
    2. 预解析 --lang 参数并初始化语言
    3. 创建带国际化帮助文本的 argparse
    4. 解析参数并执行相应模式

    Returns:
        int: 退出码
    """
    # 初始化控制台编码（解决Windows中文乱码）
    setup_console_encoding()

    try:
        # 预解析 --lang 参数，在创建 argparse 前初始化语言
        # 这样 --help 输出才能使用正确的语言
        lang = pre_parse_lang()
        init_cli_locale(lang)
        from docwen.translation import set_translator

        set_translator(cli_t)

        # 创建带国际化帮助文本的解析器
        parser = create_argument_parser()
        args = parser.parse_args()

        from docwen.cli import executor

        cmd = getattr(args, "command", None)
        if int(getattr(args, "jobs", 1) or 1) < 1:
            return _print_invalid_input(action=str(cmd or ""), args=args, message="--jobs 必须 >= 1")
        if cmd == "inspect":
            return executor.inspect_file(args.file, json_mode=bool(args.json))
        if cmd == "doctor":
            from docwen.cli.doctor import run_doctor

            return run_doctor(json_mode=bool(args.json))
        if cmd == "actions" and getattr(args, "actions_command", None) == "list":
            return executor.list_all_actions(json_mode=bool(args.json))
        if cmd == "numbering-schemes" and getattr(args, "numbering_command", None) == "list":
            return executor.list_numbering_schemes(json_mode=bool(args.json))
        if cmd == "templates" and getattr(args, "templates_command", None) == "list":
            return executor.list_templates(
                json_mode=bool(args.json),
                target=getattr(args, "for_target", None),
            )
        if cmd == "optimizations" and getattr(args, "optimizations_command", None) == "list":
            return executor.list_optimizations(
                json_mode=bool(args.json),
                scope=getattr(args, "scope", None),
            )
        if cmd == "formats" and getattr(args, "formats_command", None) == "list":
            return executor.list_formats(json_mode=bool(args.json), source=getattr(args, "for_source", None))

        if cmd == "convert":
            action = "convert"
            try:
                options = build_options(action, args)
            except ValueError as e:
                return _print_invalid_input(
                    action=action, args=args, message=str(e), input_file=(args.files[0] if args.files else "")
                )
            supported = set(executor.get_supported_convert_targets())
            target_fmt = str(options.get("target_format") or "").strip().lower()
            if target_fmt and target_fmt not in supported:
                suggestions = get_close_matches(target_fmt, sorted(supported), n=5, cutoff=0.2)
                hint = f"；你可能想要：{', '.join(suggestions)}" if suggestions else ""
                return _print_invalid_input(
                    action=action,
                    args=args,
                    message=f"--to {target_fmt} 不支持{hint}；可用列表见：docwen formats list",
                    input_file=(args.files[0] if args.files else ""),
                )
            has_md_only_flags = any(
                [
                    bool(getattr(args, "extract_img", False)),
                    bool(getattr(args, "no_extract_img", False)),
                    bool(getattr(args, "ocr", False)),
                    bool(getattr(args, "optimize_for", None)),
                ]
            )
            if target_fmt != "md" and has_md_only_flags:
                return _print_invalid_input(
                    action=action,
                    args=args,
                    message="--extract-img/--no-extract-img/--ocr/--optimize-for 仅在 --to md 时可用",
                    input_file=(args.files[0] if args.files else ""),
                )

            template_name = options.get("template_name")
            if target_fmt == "md" and template_name:
                return _print_invalid_input(
                    action=action,
                    args=args,
                    message="--template 仅在 Markdown 转文档/表格时使用（--to 不是 md）",
                    input_file=(args.files[0] if args.files else ""),
                )

            if target_fmt != "md" and not template_name:
                exts_need_template = {".md", ".markdown", ".txt"}
                targets_need_template = {"docx", "doc", "odt", "rtf", "xlsx", "xls", "ods", "csv"}
                if target_fmt in targets_need_template and any(
                    str(f).lower().endswith(tuple(exts_need_template)) for f in (args.files or [])
                ):
                    return _print_invalid_input(
                        action=action,
                        args=args,
                        message="Markdown 转文档/表格需要指定 --template",
                        input_file=(args.files[0] if args.files else ""),
                    )

            if target_fmt == "md":
                optimize_for_type = options.get("optimize_for_type")
                if optimize_for_type:
                    try:
                        from docwen.config import config_manager

                        types = config_manager.get_optimization_types()
                        entry = types.get(str(optimize_for_type))
                        if not isinstance(entry, dict) or not entry.get("enabled", True):
                            raise KeyError(str(optimize_for_type))
                    except KeyError:
                        available = []
                        try:
                            from docwen.config import config_manager

                            types = config_manager.get_optimization_types()
                            available = sorted(
                                [k for k, v in (types or {}).items() if isinstance(v, dict) and v.get("enabled", True)]
                            )
                        except Exception:
                            available = []
                        suggestions = get_close_matches(str(optimize_for_type), available, n=5, cutoff=0.2)
                        hint = f"；你可能想要：{', '.join(suggestions)}" if suggestions else ""
                        return _print_invalid_input(
                            action=action,
                            args=args,
                            message=f"--optimize-for {optimize_for_type} 不存在或未启用{hint}；可用列表见：docwen optimizations list",
                            input_file=(args.files[0] if args.files else ""),
                        )
                    except Exception:
                        return _print_invalid_input(
                            action=action,
                            args=args,
                            message="无法读取优化类型配置，请先运行：docwen doctor",
                            input_file=(args.files[0] if args.files else ""),
                        )
            return _run_action_for_files(action, args.files, options, args)
        if cmd == "validate":
            action = "validate"
            try:
                options = build_options(action, args)
            except ValueError as e:
                return _print_invalid_input(
                    action=action, args=args, message=str(e), input_file=(args.files[0] if args.files else "")
                )
            return _run_action_for_files(action, args.files, options, args)
        if cmd == "merge-tables":
            action = "merge_tables"
            try:
                options = build_options(action, args)
            except ValueError as e:
                return _print_invalid_input(
                    action=action, args=args, message=str(e), input_file=(args.files[0] if args.files else "")
                )
            return _run_action_for_files(action, args.files, options, args)
        if cmd == "merge-pdfs":
            action = "merge_pdfs"
            try:
                options = build_options(action, args)
            except ValueError as e:
                return _print_invalid_input(
                    action=action, args=args, message=str(e), input_file=(args.files[0] if args.files else "")
                )
            return _run_action_for_files(action, args.files, options, args)
        if cmd == "split-pdf":
            action = "split_pdf"
            try:
                options = build_options(action, args)
            except ValueError as e:
                return _print_invalid_input(
                    action=action, args=args, message=str(e), input_file=str(getattr(args, "file", "") or "")
                )
            return _run_action_for_files(action, [args.file], options, args)
        if cmd == "merge-images-to-tiff":
            action = "merge_images_to_tiff"
            try:
                options = build_options(action, args)
            except ValueError as e:
                return _print_invalid_input(
                    action=action, args=args, message=str(e), input_file=(args.files[0] if args.files else "")
                )
            return _run_action_for_files(action, args.files, options, args)
        if cmd == "md-numbering":
            action = "process_md_numbering"
            try:
                options = build_options(action, args)
            except ValueError as e:
                return _print_invalid_input(
                    action=action, args=args, message=str(e), input_file=(args.files[0] if args.files else "")
                )
            return _run_action_for_files(action, args.files, options, args)

        raise ValueError(f"未知命令: {cmd}")

    except KeyboardInterrupt:
        logger.info("用户中断程序")
        print(f"\n\n{cli_t('cli.messages.program_interrupted')}", file=sys.stderr)
        return 130

    except Exception as e:
        logger.error(f"程序异常: {e}", exc_info=True)
        print(f"{cli_t('common.error')}: {e}", file=sys.stderr)
        return int(ExitCode.UNKNOWN_ERROR)


if __name__ == "__main__":
    sys.exit(main())
