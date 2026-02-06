"""
CLI主入口模块

提供命令行参数解析和程序入口：
- argparse参数定义
- Headless模式执行
- 交互模式路由
- AI功能支持
- Windows控制台Unicode支持
"""

import sys
import argparse
import logging
from typing import Optional, List

from docwen.cli.i18n import cli_t, init_cli_locale
from docwen.proofread_keys import SYMBOL_CORRECTION, SYMBOL_PAIRING, SENSITIVE_WORD, TYPOS_RULE

logger = logging.getLogger(__name__)


# ==================== 控制台编码 ====================

def setup_console_encoding():
    """
    设置Windows控制台编码为UTF-8
    
    解决Windows cmd/PowerShell中文乱码问题：
    1. 设置控制台代码页为65001 (UTF-8)
    2. 重新配置stdout/stderr的编码
    """
    if sys.platform != 'win32':
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

        if isinstance(sys.stdout, io.TextIOWrapper) and isinstance(sys.stderr, io.TextIOWrapper):
            sys.stdout.reconfigure(encoding='utf-8', errors='replace')
            sys.stderr.reconfigure(encoding='utf-8', errors='replace')
            logger.debug("已重新配置stdout/stderr编码为UTF-8")
        else:
            # Python 3.6兼容
            sys.stdout = io.TextIOWrapper(
                sys.stdout.buffer, 
                encoding='utf-8', 
                errors='replace',
                line_buffering=True
            )
            sys.stderr = io.TextIOWrapper(
                sys.stderr.buffer, 
                encoding='utf-8', 
                errors='replace',
                line_buffering=True
            )
            logger.debug("已包装stdout/stderr为UTF-8 TextIOWrapper")
    except Exception as e:
        logger.debug(f"重新配置stdout/stderr失败: {e}")


# ==================== 语言预解析 ====================

def get_available_locale_codes() -> List[str]:
    """
    获取当前可用的语言代码列表
    
    用于：
    - argparse choices
    - 预解析 --lang 以便 help 文本使用正确语言
    """
    try:
        from docwen.i18n.i18n_manager import I18nManager
        codes: List[str] = []
        for locale in I18nManager().get_available_locales():
            code = locale.get("code")
            if isinstance(code, str) and code:
                codes.append(code)
        return codes
    except Exception:
        return ['zh_CN', 'en_US']


def pre_parse_lang() -> Optional[str]:
    """
    预解析 --lang 参数
    
    在创建 argparse 之前先获取语言设置，
    以便 help 文本可以使用正确的语言。
    
    Returns:
        Optional[str]: 语言代码或 None
    """
    available = set(get_available_locale_codes())
    for i, arg in enumerate(sys.argv):
        if arg == '--lang' and i + 1 < len(sys.argv):
            lang = sys.argv[i + 1]
            if lang in available:
                return lang
        elif arg.startswith('--lang='):
            lang = arg.split('=', 1)[1]
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
    parser = argparse.ArgumentParser(
        prog='docwen',
        description=cli_t('cli.description'),
        epilog=f"{cli_t('cli.example')}: %(prog)s document.docx --action export_md --extract-img"
    )
    
    # ========== 基础参数 ==========
    parser.add_argument(
        'files',
        nargs='*',
        help=cli_t('cli.help.files')
    )
    
    parser.add_argument(
        '--action',
        choices=[
            'export_md', 'convert', 'validate',
            'merge_tables', 'merge_pdfs', 'split_pdf',
            'merge_images_to_tiff',
            'process_md_numbering'
        ],
        help=cli_t('cli.help.action')
    )
    
    parser.add_argument(
        '--target',
        help=cli_t('cli.help.target')
    )
    
    # ========== 导出选项 ==========
    export_group = parser.add_argument_group(
        cli_t('cli.groups.export_options'),
        cli_t('cli.groups.export_options_desc')
    )
    
    export_group.add_argument(
        '--extract-img',
        action='store_true',
        help=cli_t('cli.help.extract_img')
    )
    
    export_group.add_argument(
        '--ocr',
        action='store_true',
        help=cli_t('cli.help.ocr')
    )
    
    export_group.add_argument(
        '--optimize-for',
        choices=['gongwen', 'contract', 'thesis'],
        help=cli_t('cli.help.optimize_for')
    )
    
    # ========== 模板选项 ==========
    parser.add_argument(
        '--template',
        help=cli_t('cli.help.template')
    )
    
    # ========== 文档校对选项 ==========
    validate_group = parser.add_argument_group(
        cli_t('cli.groups.validate_options'),
        cli_t('cli.groups.validate_options_desc')
    )
    
    validate_group.add_argument(
        '--check-punct',
        action='store_true',
        help=cli_t('cli.help.check_punct')
    )
    
    validate_group.add_argument(
        '--check-typo',
        action='store_true',
        help=cli_t('cli.help.check_typo')
    )
    
    validate_group.add_argument(
        '--check-symbol',
        action='store_true',
        help=cli_t('cli.help.check_symbol')
    )
    
    validate_group.add_argument(
        '--check-sensitive',
        action='store_true',
        help=cli_t('cli.help.check_sensitive')
    )

    validate_group.add_argument(
        '--check-none',
        action='store_true',
        help=cli_t('cli.help.check_none')
    )
    
    # ========== 表格汇总选项 ==========
    merge_group = parser.add_argument_group(
        cli_t('cli.groups.merge_options'),
        cli_t('cli.groups.merge_options_desc')
    )
    
    merge_group.add_argument(
        '--merge-mode',
        choices=['row', 'col', 'cell'],
        help=cli_t('cli.help.merge_mode')
    )
    
    merge_group.add_argument(
        '--base-table',
        help=cli_t('cli.help.base_table')
    )
    
    # ========== PDF操作选项 ==========
    pdf_group = parser.add_argument_group(
        cli_t('cli.groups.pdf_options'),
        cli_t('cli.groups.pdf_options_desc')
    )
    
    pdf_group.add_argument(
        '--pages',
        help=cli_t('cli.help.pages')
    )
    
    pdf_group.add_argument(
        '--dpi',
        type=int,
        choices=[150, 300, 600],
        help=cli_t('cli.help.dpi')
    )
    
    # ========== 图片压缩选项 ==========
    image_group = parser.add_argument_group(
        cli_t('cli.groups.image_options'),
        cli_t('cli.groups.image_options_desc')
    )
    
    image_group.add_argument(
        '--compress',
        choices=['lossless', 'limit_size'],
        help=cli_t('cli.help.compress')
    )
    
    image_group.add_argument(
        '--size-limit',
        help=cli_t('cli.help.size_limit')
    )
    
    image_group.add_argument(
        '--quality-mode',
        choices=['original', 'a4', 'a3'],
        help=cli_t('cli.help.quality_mode')
    )
    
    # ========== 批处理参数 ==========
    batch_group = parser.add_argument_group(
        cli_t('cli.groups.batch_options'),
        cli_t('cli.groups.batch_options_desc')
    )
    
    batch_group.add_argument(
        '--batch',
        action='store_true',
        help=cli_t('cli.help.batch')
    )
    
    batch_group.add_argument(
        '--yes', '-y',
        action='store_true',
        help=cli_t('cli.help.yes')
    )
    
    batch_group.add_argument(
        '--continue-on-error',
        action='store_true',
        help=cli_t('cli.help.continue_on_error')
    )
    
    # ========== 输出控制 ==========
    output_group = parser.add_argument_group(
        cli_t('cli.groups.output_options'),
        cli_t('cli.groups.output_options_desc')
    )
    
    output_group.add_argument(
        '--json',
        action='store_true',
        help=cli_t('cli.help.json')
    )
    
    output_group.add_argument(
        '--quiet', '-q',
        action='store_true',
        help=cli_t('cli.help.quiet')
    )
    
    output_group.add_argument(
        '--verbose', '-v',
        action='store_true',
        help=cli_t('cli.help.verbose')
    )
    
    # ========== 序号处理选项 ==========
    numbering_group = parser.add_argument_group(
        cli_t('cli.groups.numbering_options'),
        cli_t('cli.groups.numbering_options_desc')
    )
    
    numbering_group.add_argument(
        '--remove-numbering',
        action='store_true',
        help=cli_t('cli.help.remove_numbering')
    )
    
    numbering_group.add_argument(
        '--add-numbering',
        action='store_true',
        help=cli_t('cli.help.add_numbering')
    )
    
    numbering_group.add_argument(
        '--numbering-scheme',
        default='gongwen_standard',
        help=cli_t('cli.help.numbering_scheme')
    )
    
    # ========== AI功能 ==========
    ai_group = parser.add_argument_group(
        cli_t('cli.groups.ai_options'),
        cli_t('cli.groups.ai_options_desc')
    )
    
    ai_group.add_argument(
        '--inspect',
        action='store_true',
        help=cli_t('cli.help.inspect')
    )
    
    ai_group.add_argument(
        '--list-actions',
        action='store_true',
        help=cli_t('cli.help.list_actions')
    )
    
    ai_group.add_argument(
        '--list-numbering-schemes',
        action='store_true',
        help=cli_t('cli.help.list_numbering_schemes')
    )
    
    ai_group.add_argument(
        '--list-templates',
        action='store_true',
        help=cli_t('cli.help.list_templates')
    )
    
    # ========== 版本和语言 ==========
    parser.add_argument(
        '--version', '-V',
        action='version',
        version='%(prog)s 0.1.0'
    )
    
    parser.add_argument(
        '--lang',
        choices=get_available_locale_codes(),
        help=cli_t('cli.help.lang')
    )
    
    # ========== 图片高级选项 ==========
    image_group.add_argument(
        '--keep-alpha',
        action='store_true',
        help=cli_t('cli.help.keep_alpha')
    )
    
    return parser


# ==================== Headless模式 ====================

def execute_headless(args) -> int:
    """
    Headless模式执行
    
    Args:
        args: 解析后的参数
        
    Returns:
        int: 退出码
    """
    from docwen.cli import executor, utils
    
    # 处理查询功能
    if args.list_actions:
        return executor.list_all_actions(json_mode=args.json)
    
    if args.list_numbering_schemes:
        return executor.list_numbering_schemes(json_mode=args.json)
    
    if args.list_templates:
        return executor.list_templates(json_mode=args.json)
    
    if args.inspect:
        if not args.files:
            print(cli_t('cli.messages.error_inspect_no_file'))
            return 1
        return executor.inspect_file(args.files[0], json_mode=args.json)
    
    # 验证必需参数
    if not args.action:
        print(cli_t('cli.messages.error_no_action'))
        return 1
    
    if not args.files:
        print(cli_t('cli.messages.error_no_files'))
        return 1
    
    # 展开文件路径
    files = utils.expand_paths(args.files)
    
    if not files:
        print(cli_t('cli.messages.error_file_not_found'))
        return 1
    
    # 验证文件（现在返回三个值：有效文件、无效文件列表、警告文件列表）
    valid_files, invalid_files, warning_files = utils.validate_files(files)
    
    if invalid_files and not args.quiet:
        print(cli_t('cli.messages.warning_invalid_files', count=len(invalid_files)))
        for file, reason in invalid_files:
            print(f"  - {file}: {reason}")
    
    if not valid_files:
        print(cli_t('cli.messages.error_no_valid_files'))
        return 1
    
    # 构建选项字典
    options = build_options(args)
    
    # 创建进度回调
    progress_callback = utils.create_progress_callback(
        quiet=args.quiet,
        verbose=args.verbose
    )
    
    # 执行操作
    if len(valid_files) == 1:
        # 单文件
        return executor.execute_action(
            args.action,
            valid_files[0],
            options,
            json_mode=args.json,
            progress_callback=progress_callback
        )
    else:
        # 批量文件
        return executor.execute_batch(
            args.action,
            valid_files,
            options,
            json_mode=args.json,
            continue_on_error=args.continue_on_error,
            progress_callback=progress_callback
        )


def build_options(args) -> dict:
    """
    从命令行参数构建选项字典
    
    Args:
        args: 解析后的参数
        
    Returns:
        dict: 选项字典
    """
    options = {}
    
    # 目标格式
    if args.target:
        options['target_format'] = args.target.lower()
    
    # 导出选项
    if args.extract_img:
        options['extract_image'] = True
    if args.ocr:
        options['extract_ocr'] = True
    if args.optimize_for:
        options['optimize_for_type'] = args.optimize_for
    else:
        # 不提供--optimize-for时，不启用优化（简化模式）
        options['optimize_for_type'] = None
    
    # 模板
    if args.template:
        options['template_name'] = args.template
    
    # 校对选项 - 使用布尔字典格式
    has_any_check = args.check_punct or args.check_typo or args.check_symbol or args.check_sensitive
    if getattr(args, "check_none", False) and has_any_check:
        raise ValueError("--check-none 不能与 --check-* 同时使用")

    if getattr(args, "check_none", False):
        options['proofread_options'] = {}
    elif has_any_check:
        options['proofread_options'] = {
            SYMBOL_PAIRING: args.check_punct,
            SYMBOL_CORRECTION: args.check_symbol,
            TYPOS_RULE: args.check_typo,
            SENSITIVE_WORD: args.check_sensitive,
        }
    
    # 表格汇总
    if args.merge_mode:
        options['mode'] = args.merge_mode
    if args.base_table:
        options['base_table'] = args.base_table
    
    # PDF选项
    if args.pages:
        options['pages'] = args.pages
    if args.dpi:
        options['dpi'] = args.dpi
    
    # 图片选项
    if args.compress:
        options['compress_mode'] = args.compress
    if args.size_limit:
        options['size_limit'] = args.size_limit
    if args.quality_mode:
        options['quality_mode'] = args.quality_mode
    
    # 序号处理选项（根据action决定前缀）
    if args.remove_numbering:
        if args.action == 'export_md':
            options['doc_remove_numbering'] = True
        elif args.action == 'convert':
            options['md_remove_numbering'] = True
        elif args.action == 'process_md_numbering':
            options['remove_numbering'] = True
    
    if args.add_numbering:
        if args.action == 'export_md':
            options['doc_add_numbering'] = True
        elif args.action == 'convert':
            options['md_add_numbering'] = True
        elif args.action == 'process_md_numbering':
            options['add_numbering'] = True
    
    if args.numbering_scheme:
        if args.action == 'export_md':
            options['doc_numbering_scheme'] = args.numbering_scheme
        elif args.action == 'convert':
            options['md_numbering_scheme'] = args.numbering_scheme
        elif args.action == 'process_md_numbering':
            options['numbering_scheme'] = args.numbering_scheme
    
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
        
        # 创建带国际化帮助文本的解析器
        parser = create_argument_parser()
        args = parser.parse_args()
        
        # 判断模式
        if args.action or args.inspect or args.list_actions or args.list_numbering_schemes or args.list_templates:
            # Headless模式
            return execute_headless(args)
        else:
            # 交互模式
            from docwen.cli.interactive import interactive_mode
            return interactive_mode(args)
    
    except KeyboardInterrupt:
        logger.info("用户中断程序")
        print(f"\n\n{cli_t('cli.messages.program_interrupted')}")
        return 130
    
    except Exception as e:
        logger.error(f"程序异常: {e}", exc_info=True)
        print(f"{cli_t('common.error')}: {e}")
        return 1


if __name__ == '__main__':
    sys.exit(main())
