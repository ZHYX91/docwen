"""
交互式菜单模块

提供友好的交互式菜单系统：
- 5类文件菜单（Markdown/Document/Spreadsheet/Image/Layout）
- 批量文件处理
- 用户配置收集
- 调用executor执行
"""

import os
import logging
from typing import List, Dict, Optional

from gongwen_converter.cli import utils, executor

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
    utils.print_header("公文转换器 - 交互模式")
    
    # 1. 获取输入文件
    if args.files:
        files = utils.expand_paths(args.files)
    else:
        files = prompt_for_files()
    
    if not files:
        print("未选择文件，退出")
        return 0
    
    # 2. 验证文件
    valid_files, invalid_files, warning_files = utils.validate_files(files)
    
    if invalid_files:
        print(f"\n错误: {len(invalid_files)} 个文件无效")
        for file, reason in invalid_files:
            print(f"  ✗ {os.path.basename(file)}: {reason}")
    
    if warning_files:
        print(f"\n警告: {len(warning_files)} 个文件格式不匹配")
        for file, warning in warning_files:
            print(f"  ⚠ {os.path.basename(file)}: {warning}")
    
    if not valid_files:
        print("\n错误: 没有有效文件")
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
    print("\n请拖入文件或输入文件路径（支持通配符）")
    print("提示: 可以输入多个路径，用逗号分隔")
    print("      支持通配符，如 *.docx")
    print()
    
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
    from gongwen_converter.utils.file_type_utils import get_actual_file_category
    category = get_actual_file_category(file_path)
    
    utils.print_header(f"文件: {os.path.basename(file_path)}")
    
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
        print(f"错误: 不支持的文件类型")
        return 1


# ==================== Markdown文件菜单 ====================

def show_markdown_menu(file_path: str) -> int:
    """Markdown文件菜单"""
    print("\n请选择转换类型:")
    choice = utils.prompt_choice([
        "转换为 Word 文档 (DOCX/DOC/ODT/RTF)",
        "转换为 Excel 表格 (XLSX/XLS/ODS/CSV)",
        "取消"
    ])
    
    if choice == 1:
        return handle_md_to_document(file_path)
    elif choice == 2:
        return handle_md_to_spreadsheet(file_path)
    else:
        return 0


def handle_md_to_document(file_path: str) -> int:
    """处理MD转文档"""
    print("\n请选择输出格式:")
    format_choice = utils.prompt_choice([
        "DOCX (推荐)",
        "DOC (需要Office)", 
        "ODT (需要LibreOffice)",
        "RTF (需要Office)"
    ])
    
    format_map = {1: 'docx', 2: 'doc', 3: 'odt', 4: 'rtf'}
    target_format = format_map[format_choice]
    
    # 选择模板
    template = select_template('docx')
    if not template:
        return 1
    
    # 校对选项
    spell_check = get_spell_check_options()
    
    # 执行 - 使用convert action
    options = {
        'target_format': target_format,
        'template_name': template,
        'spell_check_options': spell_check
    }
    
    return executor.execute_action('convert', file_path, options)


def handle_md_to_spreadsheet(file_path: str) -> int:
    """处理MD转表格"""
    print("\n请选择输出格式:")
    format_choice = utils.prompt_choice([
        "XLSX (推荐)",
        "XLS (需要Office)",
        "ODS (需要LibreOffice)",
        "CSV (纯文本)"
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
    print("\n请选择操作:")
    choice = utils.prompt_choice([
        "导出 Markdown",
        "格式转换 (DOCX/DOC/ODT/RTF)",
        "另存为 PDF",
        "文档校对",
        "取消"
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
    print("\n导出选项:")
    extract_img = utils.prompt_yes_no("提取图片?", default=True)
    
    extract_ocr = False
    if extract_img:
        extract_ocr = utils.prompt_yes_no("OCR识别图片文字?", default=False)
    
    options = {
        'extract_image': extract_img,
        'extract_ocr': extract_ocr
    }
    
    return executor.execute_action('export_md', file_path, options)


def handle_document_convert(file_path: str) -> int:
    """处理文档格式转换"""
    print("\n请选择目标格式:")
    choice = utils.prompt_choice([
        "DOCX",
        "DOC (需要Office)",
        "ODT (需要LibreOffice)",
        "RTF (需要Office)"
    ])
    
    format_map = {1: 'docx', 2: 'doc', 3: 'odt', 4: 'rtf'}
    target_format = format_map[choice]
    
    options = {'target_format': target_format}
    return executor.execute_action('convert', file_path, options)


def handle_document_to_pdf(file_path: str) -> int:
    """处理文档另存为PDF"""
    if utils.prompt_yes_no("确认转换为PDF?", default=True):
        options = {'target_format': 'pdf'}
        return executor.execute_action('convert', file_path, options)
    return 0


def handle_document_validate(file_path: str) -> int:
    """处理文档校对"""
    spell_check = get_spell_check_options()
    
    if spell_check == 0:
        print("未选择任何校对选项")
        return 0
    
    options = {'spell_check_options': spell_check}
    return executor.execute_action('validate', file_path, options)


# ==================== Spreadsheet文件菜单 ====================

def show_spreadsheet_menu(file_path: str) -> int:
    """表格文件菜单"""
    print("\n请选择操作:")
    choice = utils.prompt_choice([
        "导出 Markdown",
        "格式转换 (XLSX/XLS/ODS/CSV)",
        "另存为 PDF",
        "取消"
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
    print("\n导出选项:")
    extract_img = utils.prompt_yes_no("提取图片?", default=False)
    
    extract_ocr = False
    if extract_img:
        extract_ocr = utils.prompt_yes_no("OCR识别图片文字?", default=False)
    
    options = {
        'extract_image': extract_img,
        'extract_ocr': extract_ocr
    }
    
    return executor.execute_action('export_md', file_path, options)


def handle_spreadsheet_convert(file_path: str) -> int:
    """处理表格格式转换"""
    print("\n请选择目标格式:")
    choice = utils.prompt_choice([
        "XLSX",
        "XLS (需要Office)",
        "ODS (需要LibreOffice)",
        "CSV (纯文本)"
    ])
    
    format_map = {1: 'xlsx', 2: 'xls', 3: 'ods', 4: 'csv'}
    target_format = format_map[choice]
    
    options = {'target_format': target_format}
    return executor.execute_action('convert', file_path, options)


def handle_spreadsheet_to_pdf(file_path: str) -> int:
    """处理表格另存为PDF"""
    if utils.prompt_yes_no("确认转换为PDF?", default=True):
        options = {'target_format': 'pdf'}
        return executor.execute_action('convert', file_path, options)
    return 0


# ==================== Image文件菜单 ====================

def show_image_menu(file_path: str) -> int:
    """图片文件菜单"""
    print("\n请选择操作:")
    choice = utils.prompt_choice([
        "导出 Markdown (OCR)",
        "格式转换 (PNG/JPG/BMP/GIF/TIF/WebP)",
        "另存为 PDF",
        "取消"
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
    print("\n导出选项:")
    print("注意: 图片文件至少需要选择一个选项")
    
    extract_img = utils.prompt_yes_no("嵌入图片?", default=True)
    extract_ocr = utils.prompt_yes_no("OCR识别?", default=False)
    
    if not extract_img and not extract_ocr:
        print("错误: 至少需要选择一个选项")
        return 1
    
    options = {
        'extract_image': extract_img,
        'extract_ocr': extract_ocr
    }
    
    return executor.execute_action('export_md', file_path, options)


def handle_image_convert(file_path: str) -> int:
    """处理图片格式转换"""
    print("\n请选择目标格式:")
    choice = utils.prompt_choice([
        "PNG (无损)",
        "JPG (有损压缩)",
        "BMP (位图)",
        "GIF (动画)",
        "TIF (专业)",
        "WebP (现代格式)"
    ])
    
    format_map = {1: 'png', 2: 'jpg', 3: 'bmp', 4: 'gif', 5: 'tif', 6: 'webp'}
    target_format = format_map[choice]
    
    # 压缩选项（仅对JPG/WebP）
    options = {'target_format': target_format}
    
    if target_format in ['jpg', 'webp']:
        compress_choice = utils.prompt_choice([
            "最高质量",
            "限制文件大小"
        ], prompt_text="压缩模式")
        
        if compress_choice == 2:
            size_limit = utils.prompt_input("文件大小上限 (如 200KB, 2MB)")
            options['compress_mode'] = 'limit_size'
            options['size_limit'] = size_limit
        else:
            options['compress_mode'] = 'lossless'
    
    return executor.execute_action('convert', file_path, options)


def handle_image_to_pdf(file_path: str) -> int:
    """处理图片另存为PDF"""
    print("\n请选择PDF尺寸:")
    quality_choice = utils.prompt_choice([
        "原图尺寸",
        "适合A4",
        "适合A3"
    ])
    
    quality_map = {1: 'original', 2: 'a4', 3: 'a3'}
    quality_mode = quality_map[quality_choice]
    
    options = {'quality_mode': quality_mode}
    return executor.execute_action('convert_image_to_pdf', file_path, options)


# ==================== Layout文件菜单 ====================

def show_layout_menu(file_path: str) -> int:
    """版式文件菜单"""
    print("\n请选择操作:")
    choice = utils.prompt_choice([
        "导出 Markdown",
        "转换为文档 (DOCX/DOC)",
        "渲染为图片 (TIF/JPG)",
        "PDF工具箱 (合并/拆分)",
        "取消"
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
    print("\n提取选项:")
    extract_images = utils.prompt_yes_no("提取图片?", default=False)
    
    extract_ocr = False
    if extract_images:
        extract_ocr = utils.prompt_yes_no("OCR识别图片文字?", default=False)
    
    options = {
        'extract_images': extract_images,
        'extract_ocr': extract_ocr
    }
    
    return executor.execute_action('export_md', file_path, options)


def handle_layout_to_document(file_path: str) -> int:
    """处理版式转文档"""
    print("\n请选择目标格式:")
    choice = utils.prompt_choice([
        "DOCX (推荐)",
        "DOC (需要Office)"
    ])
    
    target_format = 'docx' if choice == 1 else 'doc'
    options = {'target_format': target_format}
    
    return executor.execute_action('convert', file_path, options)


def handle_layout_to_image(file_path: str) -> int:
    """处理版式渲染为图片"""
    print("\n请选择图片格式:")
    choice = utils.prompt_choice([
        "TIF (多页TIFF)",
        "JPG (JPEG)"
    ])
    
    target_format = 'tif' if choice == 1 else 'jpg'
    
    print("\n请选择DPI质量:")
    dpi_choice = utils.prompt_choice([
        "150 DPI (最小)",
        "300 DPI (适中，推荐)",
        "600 DPI (高清)"
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
    print("\nPDF工具箱:")
    choice = utils.prompt_choice([
        "合并PDF文件",
        "拆分PDF文件",
        "返回"
    ])
    
    if choice == 1:
        return handle_merge_pdfs(file_path)
    elif choice == 2:
        return handle_split_pdf(file_path)
    else:
        return 0


def handle_merge_pdfs(base_file: str) -> int:
    """处理PDF合并"""
    print(f"\n基准文件: {os.path.basename(base_file)}")
    print("请输入要合并的其他PDF文件路径（支持通配符）")
    
    other_files = utils.prompt_files("其他PDF文件")
    
    if not other_files:
        print("未选择其他文件")
        return 0
    
    # 合并文件列表
    all_files = [base_file] + other_files
    
    print(f"\n将合并 {len(all_files)} 个文件:")
    for i, f in enumerate(all_files, 1):
        print(f"  {i}. {os.path.basename(f)}")
    
    if utils.prompt_yes_no("确认合并?", default=True):
        return executor.execute_batch('merge_pdfs', all_files, {})
    return 0


def handle_split_pdf(file_path: str) -> int:
    """处理PDF拆分"""
    print("\n请输入要拆分的页码范围")
    print("支持格式: 1-3,5,7-10 或 1~3;5;7至10")
    
    pages = utils.prompt_input("页码范围", allow_empty=False)
    
    if pages:
        options = {'pages': pages}
        return executor.execute_action('split_pdf', file_path, options)
    return 0


# ==================== 批量处理 ====================

def handle_same_type_batch(category: str, files: List[str]) -> int:
    """处理多个同类型文件"""
    count = len(files)
    
    utils.print_header(f"检测到 {count} 个{get_category_name(category)}文件")
    utils.print_file_list(files)
    
    print("\n批处理模式:")
    mode_choice = utils.prompt_choice([
        "批量应用相同操作 (推荐)",
        "逐个处理",
        "取消"
    ])
    
    if mode_choice == 3:
        return 0
    
    if mode_choice == 1:
        # 批量模式：用第一个文件配置，应用到所有文件
        print(f"\n配置操作 (示例: {os.path.basename(files[0])})")
        
        # 显示菜单获取action和options
        action, options = get_batch_action_config(category, files[0])
        
        if not action:
            return 0
        
        # 确认
        if utils.prompt_yes_no(f"将此操作应用到所有 {count} 个文件?", default=True):
            return executor.execute_batch(action, files, options)
        return 0
    
    else:
        # 逐个模式
        for i, file in enumerate(files):
            print(f"\n当前: {os.path.basename(file)} ({i+1}/{count})")
            
            result = handle_single_file(file)
            
            if i < count - 1:
                if not utils.prompt_yes_no("继续处理下一个?", default=True):
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
        print("\n请选择操作:")
        choice = utils.prompt_choice([
            "导出Markdown",
            "格式转换",
            "取消"
        ])
        
        if choice == 1:
            extract_img = utils.prompt_yes_no("提取图片?", default=True)
            extract_ocr = utils.prompt_yes_no("OCR识别?", default=False) if extract_img else False
            return 'export_md', {'extract_image': extract_img, 'extract_ocr': extract_ocr}
        
        elif choice == 2:
            target = utils.prompt_input("目标格式 (docx/pdf/...)", default='pdf')
            return 'convert', {'target_format': target}
    
    # 其他类别类似处理...
    return None, {}


def handle_mixed_types(categories: Dict[str, List[str]]) -> int:
    """处理多类型文件"""
    utils.print_header("检测到多种类型文件")
    
    for category, files in categories.items():
        print(f"\n{get_category_name(category)} ({len(files)}个):")
        for file in files[:3]:  # 只显示前3个
            print(f"  • {os.path.basename(file)}")
        if len(files) > 3:
            print(f"  ... 还有 {len(files)-3} 个文件")
    
    print("\n处理方式:")
    choice = utils.prompt_choice([
        "按类型分别处理 (推荐)",
        "全部导出为Markdown",
        "取消"
    ])
    
    if choice == 1:
        # 逐个类别处理
        for category, files in categories.items():
            print(f"\n处理{get_category_name(category)}文件...")
            result = handle_same_type_batch(category, files)
            if result != 0:
                break
        return 0
    
    elif choice == 2:
        # 统一导出Markdown
        all_files = []
        for files in categories.values():
            all_files.extend(files)
        
        extract_img = utils.prompt_yes_no("提取图片?", default=True)
        extract_ocr = utils.prompt_yes_no("OCR识别?", default=False) if extract_img else False
        options = {'extract_image': extract_img, 'extract_ocr': extract_ocr}
        
        return executor.execute_batch('export_md', all_files, options, continue_on_error=True)
    
    return 0


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
        from gongwen_converter.template.loader import get_available_templates
        templates = get_available_templates(template_type)
        
        if not templates:
            print(f"错误: 未找到{template_type.upper()}模板")
            return None
        
        print(f"\n可用{template_type.upper()}模板:")
        choice = utils.prompt_choice(templates)
        
        return templates[choice - 1]
    
    except Exception as e:
        logger.error(f"获取模板失败: {e}")
        print(f"错误: 无法加载模板列表")
        return None


def get_spell_check_options() -> int:
    """
    获取校对选项
    
    Returns:
        int: 校对选项位掩码
    """
    print("\n校对选项 (可多选):")
    print("  1 = 标点配对")
    print("  2 = 符号校对")
    print("  4 = 错别字检查")
    print("  8 = 敏感词匹配")
    print("输入数字组合（如 7 = 1+2+4），或直接回车使用默认配置")
    
    user_input = utils.prompt_input("校对选项", default="", allow_empty=True)
    
    if not user_input:
        return -1  # 使用默认配置
    
    try:
        return int(user_input)
    except ValueError:
        print("输入无效，使用默认配置")
        return -1


def get_category_name(category: str) -> str:
    """获取类别中文名"""
    names = {
        'markdown': 'Markdown',
        'document': '文档',
        'spreadsheet': '表格',
        'image': '图片',
        'layout': '版式'
    }
    return names.get(category, category)
