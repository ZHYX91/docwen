"""
交互式菜单模块

提供友好的交互式菜单系统：
- 5类文件菜单（Markdown/Document/Spreadsheet/Image/Layout）
- 批量文件处理
- 用户配置收集
- 调用executor执行
- 序号处理选项
- 优化类型选项
"""

import os
import logging
from typing import List, Dict, Optional, Tuple

from docwen.cli import utils, executor
from docwen.cli.i18n import cli_t

logger = logging.getLogger(__name__)

# ==================== 主交互模式 ====================

def interactive_mode(args) -> int:
    """
    交互模式主入口
    
    Args:
        args: 命令行参数
        
    Returns:
        int: 退出码
    """
    utils.print_header(cli_t("cli.interactive.header"))
    
    # 1. 获取输入文件
    if args.files:
        files = utils.expand_paths(args.files)
    else:
        files = prompt_for_files()
    
    if not files:
        print(cli_t("cli.messages.no_files_selected"))
        return 0
    
    # 2. 验证文件
    valid_files, invalid_files, warning_files = utils.validate_files(files)
    
    if invalid_files:
        print(f"\n{cli_t('cli.messages.warning_invalid_files', count=len(invalid_files))}")
        for file, reason in invalid_files:
            print(f"  ✗ {os.path.basename(file)}: {reason}")
    
    if warning_files:
        print(f"\n{cli_t('cli.messages.warning_format_mismatch', count=len(warning_files))}")
        for file, warning in warning_files:
            print(f"  ⚠ {os.path.basename(file)}: {warning}")
    
    if not valid_files:
        print(f"\n{cli_t('cli.messages.error_no_valid_files')}")
        return 1
    
    # 3. 分类文件
    categories = utils.categorize_files(valid_files)
    
    # 4. 场景判断
    if len(valid_files) == 1:
        # 单文件
        return handle_single_file(valid_files[0])
    
    elif len(categories) == 1:
        # 多个同类型文件
        category = list(categories.keys())[0]
        return handle_same_type_batch(category, valid_files)
    
    else:
        # 多类型文件
        return handle_mixed_types(categories)


def prompt_for_files() -> List[str]:
    """提示用户输入文件"""
    print(f"\n{cli_t('cli.interactive.input_files')}")
    
    return utils.prompt_files()


# ==================== 单文件处理 ====================

def handle_single_file(file_path: str) -> int:
    """
    处理单个文件
    
    Args:
        file_path: 文件路径
        
    Returns:
        int: 退出码
    """
    # 使用实际格式检测（而非扩展名）以支持无扩展名或伪装文件
    from docwen.utils.file_type_utils import get_actual_file_category
    category = get_actual_file_category(file_path)
    
    utils.print_header(f"{cli_t('cli.messages.file_info')}: {os.path.basename(file_path)}")
    
    if category == 'text':
        # text类别对应markdown菜单
        return show_markdown_menu(file_path)
    elif category == 'document':
        return show_document_menu(file_path)
    elif category == 'spreadsheet':
        return show_spreadsheet_menu(file_path)
    elif category == 'image':
        return show_image_menu(file_path)
    elif category == 'layout':
        return show_layout_menu(file_path)
    else:
        print(cli_t("messages.unsupported_file_type"))
        return 1


# ==================== Markdown文件菜单 ====================

def show_markdown_menu(file_path: str) -> int:
    """Markdown文件菜单"""
    print(f"\n{cli_t('cli.interactive.select_operation')}:")
    choice = utils.prompt_choice([
        cli_t("cli.interactive.menus.markdown_to_docx"),
        cli_t("cli.interactive.menus.markdown_to_xlsx"),
        cli_t("cli.interactive.menus.process_numbering"),
        cli_t("cli.interactive.menus.cancel")
    ])
    
    if choice == 1:
        return handle_md_to_document(file_path)
    elif choice == 2:
        return handle_md_to_spreadsheet(file_path)
    elif choice == 3:
        return handle_md_numbering(file_path)
    else:
        return 0


def handle_md_to_document(file_path: str) -> int:
    """处理MD转文档"""
    print(f"\n{cli_t('cli.interactive.select_format')}:")
    format_choice = utils.prompt_choice([
        cli_t("cli.interactive.formats.docx_recommended"),
        cli_t("cli.interactive.formats.doc_office"), 
        cli_t("cli.interactive.formats.odt_libre"),
        cli_t("cli.interactive.formats.rtf_office")
    ])
    
    format_map = {1: 'docx', 2: 'doc', 3: 'odt', 4: 'rtf'}
    target_format = format_map[format_choice]
    
    # 选择模板
    template = select_template('docx')
    if not template:
        return 1
    
    # 序号处理选项
    numbering_options = get_numbering_options('md_to_doc')
    
    # 校对选项
    proofread_opts = get_proofread_options()
    
    # 执行 - 使用convert action
    options = {
        'target_format': target_format,
        'template_name': template,
        'proofread_options': proofread_opts,
        **numbering_options
    }
    
    return executor.execute_action('convert', file_path, options)


def handle_md_numbering(file_path: str) -> int:
    """处理MD文件序号（不转换格式）"""
    print(f"\n{cli_t('cli.interactive.numbering_processing', default='小标题序号处理')}:")
    
    # 获取序号选项
    numbering_options = get_numbering_options('md_only')
    
    if not numbering_options:
        print(cli_t("cli.messages.no_operation_selected", default="未选择任何操作"))
        return 0
    
    return executor.execute_action('process_md_numbering', file_path, numbering_options)


def handle_md_to_spreadsheet(file_path: str) -> int:
    """处理MD转表格"""
    print(f"\n{cli_t('cli.interactive.select_format')}:")
    format_choice = utils.prompt_choice([
        cli_t("cli.interactive.formats.xlsx_recommended"),
        cli_t("cli.interactive.formats.xls_office"),
        cli_t("cli.interactive.formats.ods_libre"),
        cli_t("cli.interactive.formats.csv_text")
    ])
    
    format_map = {1: 'xlsx', 2: 'xls', 3: 'ods', 4: 'csv'}
    target_format = format_map[format_choice]
    
    # 选择模板
    template = select_template('xlsx')
    if not template:
        return 1
    
    # 执行 - 使用convert action
    options = {
        'target_format': target_format,
        'template_name': template
    }
    
    return executor.execute_action('convert', file_path, options)


# ==================== Document文件菜单 ====================

def show_document_menu(file_path: str) -> int:
    """文档文件菜单"""
    print(f"\n{cli_t('cli.interactive.select_operation')}:")
    choice = utils.prompt_choice([
        cli_t("cli.interactive.menus.export_markdown"),
        cli_t("cli.interactive.menus.format_convert"),
        cli_t("cli.interactive.menus.save_as_pdf"),
        cli_t("cli.interactive.menus.document_validate"),
        cli_t("cli.interactive.menus.cancel")
    ])
    
    if choice == 1:
        return handle_document_export_md(file_path)
    elif choice == 2:
        return handle_document_convert(file_path)
    elif choice == 3:
        return handle_document_to_pdf(file_path)
    elif choice == 4:
        return handle_document_validate(file_path)
    else:
        return 0


def handle_document_export_md(file_path: str) -> int:
    """处理文档导出Markdown"""
    print(f"\n{cli_t('cli.interactive.export_options', default='导出选项')}:")
    extract_img = utils.prompt_yes_no(cli_t("cli.interactive.prompts.extract_images"), default=True)
    
    extract_ocr = False
    if extract_img:
        extract_ocr = utils.prompt_yes_no(cli_t("cli.interactive.prompts.ocr_recognition"), default=False)
    
    # 优化选项
    optimize_options = get_optimize_options()
    
    # 序号处理选项
    numbering_options = get_numbering_options('doc_to_md')
    
    options = {
        'extract_image': extract_img,
        'extract_ocr': extract_ocr,
        **optimize_options,
        **numbering_options
    }
    
    return executor.execute_action('export_md', file_path, options)


def handle_document_convert(file_path: str) -> int:
    """处理文档格式转换"""
    print(f"\n{cli_t('cli.interactive.select_target_format', default='请选择目标格式')}:")
    choice = utils.prompt_choice([
        "DOCX",
        cli_t("cli.interactive.formats.doc_office"),
        cli_t("cli.interactive.formats.odt_libre"),
        cli_t("cli.interactive.formats.rtf_office")
    ])
    
    format_map = {1: 'docx', 2: 'doc', 3: 'odt', 4: 'rtf'}
    target_format = format_map[choice]
    
    options = {'target_format': target_format}
    return executor.execute_action('convert', file_path, options)


def handle_document_to_pdf(file_path: str) -> int:
    """处理文档另存为PDF"""
    if utils.prompt_yes_no(cli_t("cli.interactive.confirm_convert_pdf", default="确认转换为PDF?"), default=True):
        options = {'target_format': 'pdf'}
        return executor.execute_action('convert', file_path, options)
    return 0


def handle_document_validate(file_path: str) -> int:
    """处理文档校对"""
    proofread_opts = get_proofread_options()
    
    # 检查是否至少启用了一个校对规则
    if not any(proofread_opts.values()):
        print(cli_t("cli.messages.no_proofread_selected", default="未选择任何校对选项"))
        return 0
    
    options = {'proofread_options': proofread_opts}
    return executor.execute_action('validate', file_path, options)


# ==================== Spreadsheet文件菜单 ====================

def show_spreadsheet_menu(file_path: str) -> int:
    """表格文件菜单"""
    print(f"\n{cli_t('cli.interactive.select_operation')}:")
    choice = utils.prompt_choice([
        cli_t("cli.interactive.menus.export_markdown"),
        cli_t("cli.interactive.menus.format_convert"),
        cli_t("cli.interactive.menus.save_as_pdf"),
        cli_t("cli.interactive.menus.cancel")
    ])
    
    if choice == 1:
        return handle_spreadsheet_export_md(file_path)
    elif choice == 2:
        return handle_spreadsheet_convert(file_path)
    elif choice == 3:
        return handle_spreadsheet_to_pdf(file_path)
    else:
        return 0


def handle_spreadsheet_export_md(file_path: str) -> int:
    """处理表格导出Markdown"""
    print(f"\n{cli_t('cli.interactive.export_options', default='导出选项')}:")
    extract_img = utils.prompt_yes_no(cli_t("cli.interactive.prompts.extract_images"), default=False)
    
    extract_ocr = False
    if extract_img:
        extract_ocr = utils.prompt_yes_no(cli_t("cli.interactive.prompts.ocr_recognition"), default=False)
    
    options = {
        'extract_image': extract_img,
        'extract_ocr': extract_ocr
    }
    
    return executor.execute_action('export_md', file_path, options)


def handle_spreadsheet_convert(file_path: str) -> int:
    """处理表格格式转换"""
    print(f"\n{cli_t('cli.interactive.select_target_format', default='请选择目标格式')}:")
    choice = utils.prompt_choice([
        "XLSX",
        cli_t("cli.interactive.formats.xls_office"),
        cli_t("cli.interactive.formats.ods_libre"),
        cli_t("cli.interactive.formats.csv_text")
    ])
    
    format_map = {1: 'xlsx', 2: 'xls', 3: 'ods', 4: 'csv'}
    target_format = format_map[choice]
    
    options = {'target_format': target_format}
    return executor.execute_action('convert', file_path, options)


def handle_spreadsheet_to_pdf(file_path: str) -> int:
    """处理表格另存为PDF"""
    if utils.prompt_yes_no(cli_t("cli.interactive.confirm_convert_pdf", default="确认转换为PDF?"), default=True):
        options = {'target_format': 'pdf'}
        return executor.execute_action('convert', file_path, options)
    return 0


# ==================== Image文件菜单 ====================

def show_image_menu(file_path: str) -> int:
    """图片文件菜单"""
    print(f"\n{cli_t('cli.interactive.select_operation')}:")
    choice = utils.prompt_choice([
        cli_t("cli.interactive.menus.export_md_ocr", default="导出 Markdown (OCR)"),
        cli_t("cli.interactive.menus.image_format_convert", default="格式转换 (PNG/JPG/BMP/GIF/TIF/WebP)"),
        cli_t("cli.interactive.menus.save_as_pdf"),
        cli_t("cli.interactive.menus.cancel")
    ])
    
    if choice == 1:
        return handle_image_export_md(file_path)
    elif choice == 2:
        return handle_image_convert(file_path)
    elif choice == 3:
        return handle_image_to_pdf(file_path)
    else:
        return 0


def handle_image_export_md(file_path: str) -> int:
    """处理图片导出Markdown"""
    print(f"\n{cli_t('cli.interactive.export_options', default='导出选项')}:")
    print(cli_t("cli.interactive.image_export_hint", default="注意: 图片文件至少需要选择一个选项"))
    
    extract_img = utils.prompt_yes_no(cli_t("cli.interactive.prompts.embed_image", default="嵌入图片?"), default=True)
    extract_ocr = utils.prompt_yes_no(cli_t("cli.interactive.prompts.ocr_simple", default="OCR识别?"), default=False)
    
    if not extract_img and not extract_ocr:
        print(cli_t("cli.messages.error_need_option", default="错误: 至少需要选择一个选项"))
        return 1
    
    options = {
        'extract_image': extract_img,
        'extract_ocr': extract_ocr
    }
    
    return executor.execute_action('export_md', file_path, options)


def handle_image_convert(file_path: str) -> int:
    """处理图片格式转换"""
    print(f"\n{cli_t('cli.interactive.select_target_format', default='请选择目标格式')}:")
    choice = utils.prompt_choice([
        cli_t("cli.interactive.formats.png_lossless"),
        cli_t("cli.interactive.formats.jpg_lossy", default="JPG (有损压缩)"),
        cli_t("cli.interactive.formats.bmp", default="BMP (位图)"),
        cli_t("cli.interactive.formats.gif", default="GIF (动画)"),
        cli_t("cli.interactive.formats.tif", default="TIF (专业)"),
        cli_t("cli.interactive.formats.webp", default="WebP (现代格式)")
    ])
    
    format_map = {1: 'png', 2: 'jpg', 3: 'bmp', 4: 'gif', 5: 'tif', 6: 'webp'}
    target_format = format_map[choice]
    
    # 压缩选项（仅对JPG/WebP）
    options = {'target_format': target_format}
    
    if target_format in ['jpg', 'webp']:
        compress_choice = utils.prompt_choice([
            cli_t("cli.interactive.compress.highest_quality"),
            cli_t("cli.interactive.compress.limit_size")
        ], prompt_text=cli_t("cli.interactive.select_compress_mode", default="压缩模式"))
        
        if compress_choice == 2:
            size_limit = utils.prompt_input(cli_t("cli.interactive.input_size_limit"))
            options['compress_mode'] = 'limit_size'
            options['size_limit'] = size_limit
        else:
            options['compress_mode'] = 'lossless'
    
    return executor.execute_action('convert', file_path, options)


def handle_image_to_pdf(file_path: str) -> int:
    """处理图片另存为PDF"""
    print(f"\n{cli_t('cli.interactive.select_pdf_size', default='请选择PDF尺寸')}:")
    quality_choice = utils.prompt_choice([
        cli_t("cli.interactive.pdf_sizes.original", default="原图尺寸"),
        cli_t("cli.interactive.pdf_sizes.fit_a4", default="适合A4"),
        cli_t("cli.interactive.pdf_sizes.fit_a3", default="适合A3")
    ])
    
    quality_map = {1: 'original', 2: 'a4', 3: 'a3'}
    quality_mode = quality_map[quality_choice]
    
    options = {'quality_mode': quality_mode}
    return executor.execute_action('convert_image_to_pdf', file_path, options)


# ==================== Layout文件菜单 ====================

def show_layout_menu(file_path: str) -> int:
    """版式文件菜单"""
    print(f"\n{cli_t('cli.interactive.select_operation')}:")
    choice = utils.prompt_choice([
        cli_t("cli.interactive.menus.export_markdown"),
        cli_t("cli.interactive.menus.convert_to_document", default="转换为文档 (DOCX/DOC)"),
        cli_t("cli.interactive.menus.render_to_image"),
        cli_t("cli.interactive.menus.pdf_tools"),
        cli_t("cli.interactive.menus.cancel")
    ])
    
    if choice == 1:
        return handle_layout_export_md(file_path)
    elif choice == 2:
        return handle_layout_to_document(file_path)
    elif choice == 3:
        return handle_layout_to_image(file_path)
    elif choice == 4:
        return handle_pdf_tools(file_path)
    else:
        return 0


def handle_layout_export_md(file_path: str) -> int:
    """处理版式导出Markdown"""
    print(f"\n{cli_t('cli.interactive.extraction_options', default='提取选项')}:")
    extract_images = utils.prompt_yes_no(cli_t("cli.interactive.prompts.extract_images"), default=False)
    
    extract_ocr = False
    if extract_images:
        extract_ocr = utils.prompt_yes_no(cli_t("cli.interactive.prompts.ocr_recognition"), default=False)
    
    options = {
        'extract_images': extract_images,
        'extract_ocr': extract_ocr
    }
    
    return executor.execute_action('export_md', file_path, options)


def handle_layout_to_document(file_path: str) -> int:
    """处理版式转文档"""
    print(f"\n{cli_t('cli.interactive.select_target_format', default='请选择目标格式')}:")
    choice = utils.prompt_choice([
        cli_t("cli.interactive.formats.docx_recommended"),
        cli_t("cli.interactive.formats.doc_office")
    ])
    
    target_format = 'docx' if choice == 1 else 'doc'
    options = {'target_format': target_format}
    
    return executor.execute_action('convert', file_path, options)


def handle_layout_to_image(file_path: str) -> int:
    """处理版式渲染为图片"""
    print(f"\n{cli_t('cli.interactive.select_image_format', default='请选择图片格式')}:")
    choice = utils.prompt_choice([
        cli_t("cli.interactive.formats.tif_multipage"),
        cli_t("cli.interactive.formats.jpg_jpeg")
    ])
    
    target_format = 'tif' if choice == 1 else 'jpg'
    
    print(f"\n{cli_t('cli.interactive.select_dpi')}:")
    dpi_choice = utils.prompt_choice([
        cli_t("cli.interactive.dpi.dpi_150"),
        cli_t("cli.interactive.dpi.dpi_300"),
        cli_t("cli.interactive.dpi.dpi_600")
    ], default=2)
    
    dpi_map = {1: 150, 2: 300, 3: 600}
    dpi = dpi_map[dpi_choice]
    
    options = {
        'target_format': target_format,
        'dpi': dpi
    }
    
    return executor.execute_action('convert', file_path, options)


def handle_pdf_tools(file_path: str) -> int:
    """PDF工具箱"""
    print(f"\n{cli_t('cli.interactive.menus.pdf_tools')}:")
    choice = utils.prompt_choice([
        cli_t("cli.interactive.menus.merge_pdfs"),
        cli_t("cli.interactive.menus.split_pdf"),
        cli_t("cli.interactive.menus.back")
    ])
    
    if choice == 1:
        return handle_merge_pdfs(file_path)
    elif choice == 2:
        return handle_split_pdf(file_path)
    else:
        return 0


def handle_merge_pdfs(base_file: str) -> int:
    """处理PDF合并"""
    base_label = cli_t("cli.messages.base_file", default="基准文件")
    print(f"\n{base_label}: {os.path.basename(base_file)}")
    print(cli_t("cli.interactive.input_other_pdfs", default="请输入要合并的其他PDF文件路径（支持通配符）"))
    
    other_files = utils.prompt_files(cli_t("cli.messages.other_pdf_files", default="其他PDF文件"))
    
    if not other_files:
        print(cli_t("cli.messages.no_other_files", default="未选择其他文件"))
        return 0
    
    # 合并文件列表
    all_files = [base_file] + other_files
    
    will_merge = cli_t("cli.messages.will_merge_files", default="将合并 {count} 个文件", count=len(all_files))
    print(f"\n{will_merge}:")
    for i, f in enumerate(all_files, 1):
        print(f"  {i}. {os.path.basename(f)}")
    
    if utils.prompt_yes_no(cli_t("cli.interactive.confirm_merge"), default=True):
        return executor.execute_batch('merge_pdfs', all_files, {})
    return 0


def handle_split_pdf(file_path: str) -> int:
    """处理PDF拆分"""
    print(f"\n{cli_t('cli.interactive.input_pages')}")
    print(cli_t("cli.messages.page_range_hint"))
    
    pages = utils.prompt_input(cli_t("cli.interactive.page_range_label", default="页码范围"), allow_empty=False)
    
    if pages:
        options = {'pages': pages}
        return executor.execute_action('split_pdf', file_path, options)
    return 0


# ==================== 批量处理 ====================

def handle_same_type_batch(category: str, files: List[str]) -> int:
    """处理多个同类型文件"""
    count = len(files)
    
    detected_msg = cli_t("cli.batch.detected_same_type", default="检测到 {count} 个{category}文件", count=count, category=get_category_name(category))
    utils.print_header(detected_msg)
    utils.print_file_list(files)
    
    print(f"\n{cli_t('cli.batch.batch_mode', default='批处理模式')}:")
    mode_choice = utils.prompt_choice([
        cli_t("cli.batch.batch_apply_same"),
        cli_t("cli.batch.process_one_by_one"),
        cli_t("cli.interactive.menus.cancel")
    ])
    
    if mode_choice == 3:
        return 0
    
    if mode_choice == 1:
        # 批量模式：用第一个文件配置，应用到所有文件
        config_msg = cli_t("cli.batch.configure_operation", default="配置操作")
        print(f"\n{config_msg} ({cli_t('cli.messages.example_format', default='示例')}: {os.path.basename(files[0])})")
        
        # 显示菜单获取action和options
        action, options = get_batch_action_config(category, files[0])
        
        if not action:
            return 0
        
        # 确认
        apply_msg = cli_t("cli.interactive.prompts.apply_to_all", default="将此操作应用到所有 {count} 个文件?")
        if utils.prompt_yes_no(apply_msg.format(count=count), default=True):
            return executor.execute_batch(action, files, options)
        return 0
    
    else:
        # 逐个模式
        for i, file in enumerate(files):
            current_msg = cli_t("cli.messages.current_file", default="当前")
            print(f"\n{current_msg}: {os.path.basename(file)} ({i+1}/{count})")
            
            result = handle_single_file(file)
            
            if i < count - 1:
                if not utils.prompt_yes_no(cli_t("cli.interactive.prompts.continue_next"), default=True):
                    break
        
        return 0


def get_batch_action_config(category: str, sample_file: str) -> tuple:
    """
    获取批量操作配置
    
    Returns:
        tuple: (action, options)
    """
    # 简化的菜单配置
    if category == 'document':
        print(f"\n{cli_t('cli.interactive.select_operation')}:")
        choice = utils.prompt_choice([
            cli_t("cli.interactive.menus.export_markdown"),
            cli_t("cli.interactive.menus.format_convert"),
            cli_t("cli.interactive.menus.cancel")
        ])
        
        if choice == 1:
            extract_img = utils.prompt_yes_no(cli_t("cli.interactive.prompts.extract_images"), default=True)
            extract_ocr = utils.prompt_yes_no(cli_t("cli.interactive.prompts.ocr_simple", default="OCR识别?"), default=False) if extract_img else False
            return 'export_md', {'extract_image': extract_img, 'extract_ocr': extract_ocr}
        
        elif choice == 2:
            target = utils.prompt_input(cli_t("cli.interactive.input_target_format", default="目标格式 (docx/pdf/...)"), default='pdf')
            return 'convert', {'target_format': target}
    
    # 其他类别类似处理...
    return None, {}


def handle_mixed_types(categories: Dict[str, List[str]]) -> int:
    """处理多类型文件"""
    utils.print_header(cli_t("cli.batch.detected_mixed_types"))
    
    count_suffix = cli_t("cli.messages.file_count_suffix", default="个")
    for category, files in categories.items():
        print(f"\n{get_category_name(category)} ({len(files)}{count_suffix}):")
        for file in files[:3]:  # 只显示前3个
            print(f"  • {os.path.basename(file)}")
        if len(files) > 3:
            more_msg = cli_t("cli.messages.more_files", default="还有 {count} 个文件", count=len(files)-3)
            print(f"  ... {more_msg}")
    
    print(f"\n{cli_t('cli.batch.process_mode', default='处理方式')}:")
    choice = utils.prompt_choice([
        cli_t("cli.batch.process_by_type"),
        cli_t("cli.batch.export_all_to_md"),
        cli_t("cli.interactive.menus.cancel")
    ])
    
    if choice == 1:
        # 逐个类别处理
        for category, files in categories.items():
            processing_msg = cli_t("cli.batch.processing_category", default="处理{category}文件...", category=get_category_name(category))
            print(f"\n{processing_msg}")
            result = handle_same_type_batch(category, files)
            if result != 0:
                break
        return 0
    
    elif choice == 2:
        # 统一导出Markdown
        all_files = []
        for files in categories.values():
            all_files.extend(files)
        
        extract_img = utils.prompt_yes_no(cli_t("cli.interactive.prompts.extract_images"), default=True)
        extract_ocr = utils.prompt_yes_no(cli_t("cli.interactive.prompts.ocr_simple", default="OCR识别?"), default=False) if extract_img else False
        options = {'extract_image': extract_img, 'extract_ocr': extract_ocr}
        
        return executor.execute_batch('export_md', all_files, options, continue_on_error=True)
    
    return 0


# ==================== 序号选项辅助函数 ====================

def get_numbering_options(direction: str) -> Dict[str, any]:
    """
    获取序号处理选项
    
    Args:
        direction: 转换方向
            - 'md_to_doc': MD转文档
            - 'doc_to_md': 文档转MD
            - 'md_only': 纯MD处理
            
    Returns:
        Dict: 序号选项字典
    """
    print(f"\n{cli_t('cli.interactive.numbering_options', default='序号处理选项')}:")
    
    remove_numbering = utils.prompt_yes_no(cli_t("cli.interactive.prompts.remove_numbering"), default=False)
    add_numbering = utils.prompt_yes_no(cli_t("cli.interactive.prompts.add_numbering"), default=False)
    
    options = {}
    
    # 根据方向确定选项键前缀
    if direction == 'md_to_doc':
        prefix = 'md_'
    elif direction == 'doc_to_md':
        prefix = 'doc_'
    else:
        prefix = ''
    
    if remove_numbering:
        options[f'{prefix}remove_numbering'] = True
    
    if add_numbering:
        options[f'{prefix}add_numbering'] = True
        
        # 选择序号方案
        scheme = select_numbering_scheme()
        if scheme:
            options[f'{prefix}numbering_scheme'] = scheme
    
    return options


def select_numbering_scheme() -> Optional[str]:
    """
    选择序号方案（从配置动态读取）
    
    Returns:
        Optional[str]: 方案ID
    """
    try:
        from docwen.config.config_manager import config_manager
        
        # 从配置读取方案列表（包含描述）
        schemes = config_manager.get_localized_numbering_schemes(include_description=True)
        
        if not schemes:
            raise ValueError("没有可用的序号方案")
        
        # 构建选项列表
        options = []
        scheme_ids = []
        for scheme_id, info in schemes.items():
            # 格式：方案名称（描述）
            if info.get('description'):
                display = f"{info['name']}（{info['description']}）"
            else:
                display = info['name']
            options.append(display)
            scheme_ids.append(scheme_id)
        
        print(f"\n{cli_t('cli.interactive.select_numbering_scheme')}:")
        choice = utils.prompt_choice(options)
        
        return scheme_ids[choice - 1]
        
    except Exception as e:
        logger.warning(f"从配置读取序号方案失败: {e}，使用默认方案列表")
        
        # 降级：使用硬编码方案
        print(f"\n{cli_t('cli.interactive.select_numbering_scheme')}:")
        choice = utils.prompt_choice([
            cli_t("cli.interactive.numbering_schemes.gongwen_standard"),
            cli_t("cli.interactive.numbering_schemes.hierarchical_standard"),
            cli_t("cli.interactive.numbering_schemes.hierarchical_h2_start", default="层级数字(H2起)（H1无序号，H2起为 1 1.1 1.1.1）"),
            cli_t("cli.interactive.numbering_schemes.legal_standard")
        ])
        
        scheme_map = {
            1: 'gongwen_standard',
            2: 'hierarchical_standard',
            3: 'hierarchical_h2_start',
            4: 'legal_standard'
        }
        
        return scheme_map.get(choice)


def get_optimize_options() -> Dict[str, any]:
    """
    获取文档优化选项
    
    Returns:
        Dict: 优化选项字典
    """
    options = {}
    
    enable_optimize = utils.prompt_yes_no(cli_t("cli.interactive.prompts.optimize_for_type"), default=False)
    
    if enable_optimize:
        print(f"\n{cli_t('cli.interactive.select_optimization_type')}:")
        choice = utils.prompt_choice([
            cli_t("cli.interactive.optimization_types.gongwen"),
            cli_t("cli.interactive.optimization_types.contract"),
            cli_t("cli.interactive.optimization_types.thesis")
        ])
        
        type_map = {1: 'gongwen', 2: 'contract', 3: 'thesis'}
        options['optimize_for_type'] = type_map.get(choice)
    
    return options


# ==================== 辅助函数 ====================

def select_template(template_type: str) -> Optional[str]:
    """
    选择模板
    
    Args:
        template_type: 模板类型 ('docx' 或 'xlsx')
        
    Returns:
        Optional[str]: 模板名称
    """
    try:
        from docwen.template.loader import TemplateLoader
        template_loader = TemplateLoader()
        templates = template_loader.get_available_templates(template_type)
        
        if not templates:
            error_msg = cli_t("cli.messages.error_no_template", default="错误: 未找到{type}模板", type=template_type.upper())
            print(error_msg)
            return None
        
        available_msg = cli_t("cli.messages.available_type_templates", default="可用{type}模板", type=template_type.upper())
        print(f"\n{available_msg}:")
        choice = utils.prompt_choice(templates)
        
        return templates[choice - 1]
    
    except Exception as e:
        logger.error(f"获取模板失败: {e}")
        print(cli_t("cli.messages.template_list_error", default="错误: 无法加载模板列表"))
        return None


def get_proofread_options() -> Dict[str, bool]:
    """
    获取校对选项（交互式多选）
    
    Returns:
        Dict[str, bool]: 校对选项字典，None表示使用默认配置
    """
    # 定义校对规则
    rules = [
        ("symbol_pairing", cli_t('cli.interactive.proofread.symbol_pairing')),
        ("symbol_correction", cli_t('cli.interactive.proofread.symbol_correction')),
        ("typos_rule", cli_t('cli.interactive.proofread.typos_rule')),
        ("sensitive_word", cli_t('cli.interactive.proofread.sensitive_word')),
    ]
    
    # 从配置获取默认值
    try:
        from docwen.config.config_manager import config_manager
        engine_config = config_manager.get_proofread_engine_config()
        default_states = {
            "symbol_pairing": engine_config.get("enable_symbol_pairing", True),
            "symbol_correction": engine_config.get("enable_symbol_correction", True),
            "typos_rule": engine_config.get("enable_typos_rule", True),
            "sensitive_word": engine_config.get("enable_sensitive_word", False),
        }
    except Exception:
        default_states = {
            "symbol_pairing": True,
            "symbol_correction": True,
            "typos_rule": True,
            "sensitive_word": False,
        }
    
    # 当前选中状态（初始化为默认值）
    selected = default_states.copy()
    
    while True:
        # 显示当前状态
        print(f"\n{cli_t('cli.interactive.proofread_options', default='校对选项')}:")
        for i, (key, label) in enumerate(rules, 1):
            status = "[✓]" if selected[key] else "[ ]"
            print(f"  [{i}] {label} {'.' * (30 - len(label))} {status}")
        
        print(cli_t("cli.interactive.proofread_toggle_hint", 
                   default="\n输入序号切换状态（如: 1 3 4），直接回车确认"))
        
        user_input = utils.prompt_input(
            cli_t("cli.interactive.proofread_input", default="校对选项"), 
            default="", 
            allow_empty=True
        )
        
        # 直接回车确认当前选择
        if not user_input.strip():
            return selected
        
        # 解析用户输入的序号
        try:
            indices = [int(x.strip()) for x in user_input.split()]
            for idx in indices:
                if 1 <= idx <= len(rules):
                    key = rules[idx - 1][0]
                    selected[key] = not selected[key]  # 切换状态
                else:
                    print(cli_t("cli.messages.invalid_index", default="无效序号: {idx}", idx=idx))
        except ValueError:
            print(cli_t("cli.messages.invalid_input", default="输入无效，请输入数字序号"))


def get_category_name(category: str) -> str:
    """获取类别本地化名称"""
    # 使用 cli.categories 节的翻译
    key = f"cli.categories.{category}"
    name = cli_t(key)
    # 如果翻译键不存在（返回 [key] 格式），则使用原始类别名
    if name.startswith('[') and name.endswith(']'):
        return category
    return name
