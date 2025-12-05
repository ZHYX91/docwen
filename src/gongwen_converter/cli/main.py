"""
CLI主入口模块

提供命令行参数解析和程序入口：
- argparse参数定义
- Headless模式执行
- 交互模式路由
- AI功能支持
"""

import sys
import argparse
import logging
from typing import Optional, List

logger = logging.getLogger(__name__)

# ==================== 参数解析器 ====================

def create_argument_parser() -> argparse.ArgumentParser:
    """
    创建命令行参数解析器
    
    Returns:
        argparse.ArgumentParser: 配置好的解析器
    """
    parser = argparse.ArgumentParser(
        prog='gongwen-converter',
        description='公文转换器命令行工具 - 支持文档/表格/图片/版式文件的转换',
        epilog='示例: %(prog)s document.docx --action export_md --extract-img'
    )
    
    # ========== 基础参数 ==========
    parser.add_argument(
        'files',
        nargs='*',
        help='输入文件路径（支持通配符，如 *.docx）'
    )
    
    parser.add_argument(
        '--action',
        choices=[
            'export_md', 'convert', 'validate',
            'merge_tables', 'merge_pdfs', 'split_pdf',
            'merge_images_to_tiff'
        ],
        help='操作类型'
    )
    
    parser.add_argument(
        '--target',
        help='目标格式（用于convert操作，如 docx, pdf, jpg）'
    )
    
    # ========== 导出选项 ==========
    export_group = parser.add_argument_group('导出选项', '用于export_md操作')
    
    export_group.add_argument(
        '--extract-img',
        action='store_true',
        help='提取图片'
    )
    
    export_group.add_argument(
        '--ocr',
        action='store_true',
        help='OCR识别图片文字'
    )
    
    export_group.add_argument(
        '--optimize-for',
        choices=['gongwen', 'contract', 'thesis'],
        help='针对优化类型：gongwen（公文）、contract（合同）、thesis（论文）。不提供此参数则不启用优化（简化模式）'
    )
    
    # ========== 模板选项 ==========
    parser.add_argument(
        '--template',
        help='模板名称（用于MD转文档/表格）'
    )
    
    # ========== 文档校对选项 ==========
    validate_group = parser.add_argument_group('校对选项', '用于validate操作')
    
    validate_group.add_argument(
        '--check-punct',
        action='store_true',
        help='检查标点配对'
    )
    
    validate_group.add_argument(
        '--check-typo',
        action='store_true',
        help='检查错别字'
    )
    
    validate_group.add_argument(
        '--check-symbol',
        action='store_true',
        help='检查符号规范'
    )
    
    validate_group.add_argument(
        '--check-sensitive',
        action='store_true',
        help='检查敏感词'
    )
    
    # ========== 表格汇总选项 ==========
    merge_group = parser.add_argument_group('汇总选项', '用于merge_tables操作')
    
    merge_group.add_argument(
        '--merge-mode',
        choices=['row', 'col', 'cell'],
        help='汇总模式：按行/按列/按单元格'
    )
    
    merge_group.add_argument(
        '--base-table',
        help='基准表格文件路径'
    )
    
    # ========== PDF操作选项 ==========
    pdf_group = parser.add_argument_group('PDF选项', '用于PDF合并/拆分')
    
    pdf_group.add_argument(
        '--pages',
        help='PDF拆分页码范围（如 1-3,5,7-10）'
    )
    
    pdf_group.add_argument(
        '--dpi',
        type=int,
        choices=[150, 300, 600],
        help='图片DPI质量'
    )
    
    # ========== 图片压缩选项 ==========
    image_group = parser.add_argument_group('图片选项', '用于图片格式转换')
    
    image_group.add_argument(
        '--compress',
        choices=['lossless', 'limit_size'],
        help='压缩模式'
    )
    
    image_group.add_argument(
        '--size-limit',
        help='文件大小限制（如 200KB, 2MB）'
    )
    
    image_group.add_argument(
        '--quality-mode',
        choices=['original', 'a4', 'a3'],
        help='PDF质量模式（图片转PDF时使用）'
    )
    
    # ========== 批处理参数 ==========
    batch_group = parser.add_argument_group('批处理选项', '批量文件处理')
    
    batch_group.add_argument(
        '--batch',
        action='store_true',
        help='批量模式（所有文件使用相同配置）'
    )
    
    batch_group.add_argument(
        '--yes', '-y',
        action='store_true',
        help='跳过所有确认提示'
    )
    
    batch_group.add_argument(
        '--continue-on-error',
        action='store_true',
        help='出错时继续处理剩余文件'
    )
    
    # ========== 输出控制 ==========
    output_group = parser.add_argument_group('输出控制', '控制输出格式和详细程度')
    
    output_group.add_argument(
        '--json',
        action='store_true',
        help='JSON格式输出（适合AI/脚本解析）'
    )
    
    output_group.add_argument(
        '--quiet', '-q',
        action='store_true',
        help='安静模式（最小输出）'
    )
    
    output_group.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='详细模式（显示更多信息）'
    )
    
    # ========== AI功能 ==========
    ai_group = parser.add_argument_group('AI功能', 'AI自描述和查询功能')
    
    ai_group.add_argument(
        '--inspect',
        action='store_true',
        help='查询文件支持的操作（AI自描述）'
    )
    
    ai_group.add_argument(
        '--list-actions',
        action='store_true',
        help='列出所有可用操作'
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
    from gongwen_converter.cli import executor, utils
    
    # 处理AI功能
    if args.list_actions:
        return executor.list_all_actions(json_mode=args.json)
    
    if args.inspect:
        if not args.files:
            print("错误: --inspect 需要指定文件")
            return 1
        return executor.inspect_file(args.files[0], json_mode=args.json)
    
    # 验证必需参数
    if not args.action:
        print("错误: Headless模式需要指定 --action")
        return 1
    
    if not args.files:
        print("错误: 需要指定输入文件")
        return 1
    
    # 展开文件路径
    files = utils.expand_paths(args.files)
    
    if not files:
        print("错误: 未找到有效文件")
        return 1
    
    # 验证文件（现在返回三个值：有效文件、无效文件列表、警告文件列表）
    valid_files, invalid_files, warning_files = utils.validate_files(files)
    
    if invalid_files and not args.quiet:
        print(f"警告: {len(invalid_files)} 个文件无效")
        for file, reason in invalid_files:
            print(f"  - {file}: {reason}")
    
    if not valid_files:
        print("错误: 没有有效文件")
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
    
    # 校对选项
    if args.check_punct or args.check_typo or args.check_symbol or args.check_sensitive:
        # 构建spell_check_options位掩码
        spell_check = 0
        if args.check_punct:
            spell_check |= 1
        if args.check_symbol:
            spell_check |= 2
        if args.check_typo:
            spell_check |= 4
        if args.check_sensitive:
            spell_check |= 8
        options['spell_check_options'] = spell_check
    
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
    
    return options


# ==================== 主入口 ====================

def main() -> int:
    """
    CLI主入口
    
    Returns:
        int: 退出码
    """
    try:
        # 解析参数
        parser = create_argument_parser()
        args = parser.parse_args()
        
        # 判断模式
        if args.action or args.inspect or args.list_actions:
            # Headless模式
            return execute_headless(args)
        else:
            # 交互模式
            from gongwen_converter.cli.interactive import interactive_mode
            return interactive_mode(args)
    
    except KeyboardInterrupt:
        logger.info("用户中断程序")
        print("\n\n程序已中断")
        return 130
    
    except Exception as e:
        logger.error(f"程序异常: {e}", exc_info=True)
        print(f"错误: {e}")
        return 1


if __name__ == '__main__':
    sys.exit(main())
