"""
Strategy执行器模块

统一调用Strategy层，处理所有文件转换操作：
- 根据action和文件类型获取Strategy
- 执行Strategy并处理结果
- 支持JSON输出模式
- 批量执行支持
"""

import os
import json
import time
import logging
import traceback
from typing import Dict, List, Optional, Callable, Any, Type

from docwen.services.result import ConversionResult
from docwen.services.strategies.base_strategy import BaseStrategy
from docwen.cli.i18n import cli_t

# 导入Strategy查找函数
try:
    from docwen.services.strategies import get_strategy
except ImportError:
    # 如果__init__.py没有导出，尝试直接导入
    get_strategy = None

logger = logging.getLogger(__name__)

# ==================== Strategy获取 ====================

def get_strategy_for_action(
    action: str,
    file_path: str,
    options: Optional[Dict] = None
) -> Type[BaseStrategy]:
    """
    根据action和文件类型获取对应的Strategy类
    
    Args:
        action: 操作名称
        file_path: 文件路径
        options: 额外选项
        
    Returns:
        Type[BaseStrategy]: 策略类（需要实例化后使用）
        
    Raises:
        ValueError: 未找到对应的Strategy
    """
    from docwen.cli.utils import detect_format, detect_category
    
    options = options or {}
    
    # 优先使用actual_format（与GUI保持一致）
    if 'actual_format' in options:
        actual_format = options['actual_format']
        source_fmt = actual_format
        
        # 根据actual_format获取类别
        from docwen.utils.file_type_utils import ACTUAL_FORMAT_TO_CATEGORY
        category = ACTUAL_FORMAT_TO_CATEGORY.get(actual_format, 'unknown')
        logger.debug(f"使用实际格式: format={actual_format}, category={category}")
    else:
        # 回退到扩展名检测
        source_fmt = detect_format(file_path)
        category = detect_category(file_path)
        logger.debug(f"使用扩展名检测: format={source_fmt}, category={category}")
    
    logger.debug(f"查找Strategy: action={action}, format={source_fmt}, category={category}")
    
    # 使用Strategy层的get_strategy函数
    if get_strategy:
        try:
            # 特殊处理：export_md实际是category→md的转换操作
            if action == 'export_md':
                logger.debug(f"导出Markdown: {category} -> md")
                return get_strategy(
                    action_type=None,
                    source_format=category,  # 使用类别而不是具体格式
                    target_format='md'
                )
            
            # convert操作：使用具体格式
            elif action == 'convert':
                target_fmt = options.get('target_format', '').lower()
                if not target_fmt:
                    raise ValueError("格式转换需要指定目标格式 (--target)")
                
                logger.debug(f"格式转换: {source_fmt} -> {target_fmt}")
                return get_strategy(
                    action_type=None,
                    source_format=source_fmt,
                    target_format=target_fmt
                )
            
            # 其他命名操作（validate, merge_pdfs等）
            else:
                logger.debug(f"命名操作: {action}")
                return get_strategy(
                    action_type=action,
                    source_format=None,
                    target_format=None
                )
        
        except Exception as e:
            logger.error(f"获取Strategy失败: {e}")
            raise ValueError(f"未找到操作 '{action}' 的实现: {e}")
    
    # 降级：手动导入常用Strategy
    else:
        return _get_strategy_fallback(action, file_path, options)


def _get_strategy_fallback(action: str, file_path: str, options: Dict) -> Type[BaseStrategy]:
    """
    降级方案：手动导入Strategy类
    
    当get_strategy函数不可用时使用
    
    Returns:
        Type[BaseStrategy]: 策略类（需要实例化后使用）
    """
    from docwen.cli.utils import detect_format
    
    # 根据action导入对应的Strategy类
    if action == 'export_md':
        category = detect_format(file_path)
        if category in ['docx', 'doc']:
            from docwen.services.strategies.document import DocxToMdStrategy
            return DocxToMdStrategy
        elif category in ['xlsx', 'xls', 'et', 'csv']:
            from docwen.services.strategies.spreadsheet import SpreadsheetToMarkdownStrategy
            return SpreadsheetToMarkdownStrategy
        elif category in ['pdf', 'ofd']:
            from docwen.services.strategies.layout import LayoutToMarkdownPymupdf4llmStrategy
            return LayoutToMarkdownPymupdf4llmStrategy
    
    elif action == 'validate':
        from docwen.services.strategies.document import DocxValidationStrategy
        return DocxValidationStrategy
    
    elif action == 'merge_tables':
        from docwen.services.strategies.operations import MergeTablesStrategy
        return MergeTablesStrategy
    
    # 更多Strategy...
    
    raise ValueError(f"未实现的操作: {action}")


# ==================== 执行器 ====================

def execute_action(
    action: str,
    file_path: str,
    options: Optional[Dict] = None,
    json_mode: bool = False,
    progress_callback: Optional[Callable[[str], None]] = None
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
    start_time = time.time()
    options = options or {}
    
    try:
        # 验证文件
        if not os.path.exists(file_path):
            error_msg = f"文件不存在: {file_path}"
            if json_mode:
                print_json_error(action, file_path, error_msg)
            else:
                print(f"错误: {error_msg}")
            return 1
        
        # 检测实际文件格式（与GUI保持一致）
        if 'actual_format' not in options:
            try:
                from docwen.utils.file_type_utils import detect_actual_file_format
                actual_format = detect_actual_file_format(file_path)
                options['actual_format'] = actual_format
                logger.debug(f"检测到实际格式: {actual_format}")
            except Exception as e:
                logger.debug(f"实际格式检测失败: {e}")
        
        # 获取Strategy类
        logger.info(f"执行操作: {action} on {file_path}")
        strategy_class = get_strategy_for_action(action, file_path, options)
        
        # 实例化Strategy
        strategy = strategy_class()
        
        # 执行
        result = strategy.execute(
            file_path=file_path,
            options=options,
            progress_callback=progress_callback
        )
        
        # 计算耗时
        duration = time.time() - start_time
        
        # 输出结果
        if json_mode:
            print_json_result(result, action, file_path, duration)
        else:
            print_text_result(result)
        
        return 0 if result.success else 1
    
    except KeyboardInterrupt:
        logger.info("用户中断操作")
        if json_mode:
            print_json_error(action, file_path, "用户中断", interrupted=True)
        else:
            print("\n操作已中断")
        return 130  # SIGINT
    
    except Exception as e:
        logger.error(f"执行失败: {e}", exc_info=True)
        if json_mode:
            print_json_error(action, file_path, str(e), traceback=traceback.format_exc())
        else:
            print(f"错误: {e}")
        return 1


def execute_batch(
    action: str,
    files: List[str],
    options: Optional[Dict] = None,
    json_mode: bool = False,
    continue_on_error: bool = False,
    progress_callback: Optional[Callable[[str], None]] = None
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
    total = len(files)
    success_count = 0
    failed_files = []
    results = []
    
    for i, file in enumerate(files, 1):
        if not json_mode:
            print(f"[{i}/{total}] {os.path.basename(file)}...", end=' ', flush=True)
        
        try:
            exit_code = execute_action(
                action, file, options,
                json_mode=False,  # 批量模式下单个文件不用JSON
                progress_callback=progress_callback
            )
            
            if exit_code == 0:
                if not json_mode:
                    print("✓")
                success_count += 1
                results.append({"file": file, "success": True})
            else:
                if not json_mode:
                    print("✗")
                failed_files.append(file)
                results.append({"file": file, "success": False, "error": "操作失败"})
                
                if not continue_on_error:
                    break
        
        except KeyboardInterrupt:
            logger.info("用户中断批量操作")
            if not json_mode:
                print("\n\n批量操作已中断")
            break
        
        except Exception as e:
            logger.error(f"处理文件失败: {file}, {e}")
            if not json_mode:
                print(f"✗ {e}")
            failed_files.append(file)
            results.append({"file": file, "success": False, "error": str(e)})
            
            if not continue_on_error:
                break
    
    # 输出汇总
    if json_mode:
        print_json_batch_summary(results, total, success_count, failed_files)
    else:
        print_text_batch_summary(total, success_count, failed_files)
    
    return 0 if success_count == total else 1


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
        print(f"\n✗ {result.message or default_msg}")


def print_json_result(result: ConversionResult, action: str, input_file: str, duration: float):
    """打印JSON格式的结果"""
    output = {
        "success": result.success,
        "action": action,
        "input_file": input_file,
        "output_file": result.output_path,
        "message": result.message,
        "duration": round(duration, 2),
        "metadata": result.metadata or {}
    }
    
    print(json.dumps(output, ensure_ascii=False, indent=2))


def print_json_error(action: str, input_file: str, error: str, **kwargs):
    """打印JSON格式的错误"""
    output = {
        "success": False,
        "action": action,
        "input_file": input_file,
        "error": error,
        **kwargs
    }
    
    print(json.dumps(output, ensure_ascii=False, indent=2))


def print_text_batch_summary(total: int, success: int, failed: List[str]):
    """打印文本格式的批量操作汇总"""
    print(f"\n{'='*60}")
    batch_completed = cli_t("cli.messages.batch_completed", default="批量操作完成: {success}/{total} 成功", success=success, total=total)
    print(batch_completed)
    
    if failed:
        failed_msg = cli_t("cli.messages.failed_count", default="失败: {count} 个文件", count=len(failed))
        print(failed_msg)
        for file in failed:
            print(f"  ✗ {os.path.basename(file)}")
    
    print('='*60)


def print_json_batch_summary(results: List[Dict], total: int, success: int, failed: List[str]):
    """打印JSON格式的批量操作汇总"""
    output = {
        "success": success == total,
        "total": total,
        "success_count": success,
        "failed_count": len(failed),
        "results": results
    }
    
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
            output = {
                "file": file_path,
                "category": category,
                "format": fmt,
                "supported_actions": actions
            }
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


def get_supported_actions(category: str, fmt: str) -> List[Dict]:
    """
    获取文件类别支持的操作列表
    
    Args:
        category: 文件类别
        fmt: 文件格式
        
    Returns:
        List[Dict]: 操作列表
    """
    # 根据类别返回支持的操作
    if category == 'markdown':
        return [
            {
                "name": "convert_md_to_docx",
                "description": "转换为DOCX文档",
                "parameters": {
                    "template": {"type": "string", "description": "模板名称"},
                    "check_typo": {"type": "boolean", "default": False}
                }
            },
            {
                "name": "convert_md_to_xlsx",
                "description": "转换为XLSX表格",
                "parameters": {
                    "template": {"type": "string", "description": "模板名称"}
                }
            }
        ]
    
    elif category == 'document':
        return [
            {
                "name": "export_md",
                "description": "导出为Markdown",
                "parameters": {
                    "extract_image": {"type": "boolean", "default": True},
                    "extract_ocr": {"type": "boolean", "default": False}
                }
            },
            {
                "name": "convert",
                "description": "格式转换",
                "parameters": {
                    "target_format": {
                        "type": "choice",
                        "choices": ["docx", "doc", "odt", "rtf"],
                        "required": True
                    }
                }
            },
            {
                "name": "validate",
                "description": "文档校对",
                "parameters": {
                    "check_punct": {"type": "boolean", "default": True},
                    "check_typo": {"type": "boolean", "default": True}
                }
            }
        ]
    
    elif category == 'spreadsheet':
        return [
            {
                "name": "export_md",
                "description": "导出为Markdown",
                "parameters": {
                    "extract_image": {"type": "boolean", "default": False}
                }
            },
            {
                "name": "convert",
                "description": "格式转换",
                "parameters": {
                    "target_format": {
                        "type": "choice",
                        "choices": ["xlsx", "xls", "ods", "csv"],
                        "required": True
                    }
                }
            },
            {
                "name": "merge_tables",
                "description": "表格汇总",
                "parameters": {
                    "mode": {
                        "type": "choice",
                        "choices": ["row", "col", "cell"],
                        "required": True
                    },
                    "base_table": {"type": "string", "description": "基准表格"}
                }
            }
        ]
    
    elif category == 'layout':
        return [
            {
                "name": "export_md",
                "description": "导出为Markdown",
                "parameters": {
                    "extract_images": {"type": "boolean", "default": False},
                    "extract_ocr": {"type": "boolean", "default": False}
                }
            },
            {
                "name": "merge_pdfs",
                "description": "合并PDF"
            },
            {
                "name": "split_pdf",
                "description": "拆分PDF",
                "parameters": {
                    "pages": {"type": "string", "required": True, "description": "页码范围"}
                }
            }
        ]
    
    elif category == 'image':
        return [
            {
                "name": "export_md",
                "description": "导出为Markdown(OCR)",
                "parameters": {
                    "extract_ocr": {"type": "boolean", "default": True}
                }
            },
            {
                "name": "convert",
                "description": "格式转换",
                "parameters": {
                    "target_format": {
                        "type": "choice",
                        "choices": ["png", "jpg", "bmp", "gif", "tif", "webp"],
                        "required": True
                    }
                }
            }
        ]
    
    else:
        return []


def list_all_actions(json_mode: bool = True) -> int:
    """
    列出所有可用操作
    
    Args:
        json_mode: 是否JSON输出
        
    Returns:
        int: 退出码
    """
    actions = [
        {"name": "export_md", "description": cli_t("cli.actions.export_md", default="导出为Markdown"), "categories": ["document", "spreadsheet", "layout", "image"]},
        {"name": "convert", "description": cli_t("cli.actions.convert", default="格式转换"), "categories": ["document", "spreadsheet", "image", "layout"]},
        {"name": "validate", "description": cli_t("cli.actions.validate", default="文档校对"), "categories": ["document"]},
        {"name": "merge_tables", "description": cli_t("cli.actions.merge_tables", default="表格汇总"), "categories": ["spreadsheet"]},
        {"name": "merge_pdfs", "description": cli_t("cli.actions.merge_pdfs", default="合并PDF"), "categories": ["layout"]},
        {"name": "split_pdf", "description": cli_t("cli.actions.split_pdf", default="拆分PDF"), "categories": ["layout"]},
        {"name": "merge_images_to_tiff", "description": cli_t("cli.actions.merge_tiff", default="合并为TIF"), "categories": ["image"]},
        {"name": "process_md_numbering", "description": cli_t("cli.actions.process_numbering", default="处理MD小标题序号"), "categories": ["markdown"]},
    ]
    
    if json_mode:
        output = {"actions": actions}
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
    # 尝试从配置文件读取序号方案
    try:
        from docwen.config.config_manager import config_manager
        all_schemes = config_manager.get_heading_schemes()
        
        # 构建方案信息列表
        schemes = []
        for scheme_id, scheme_config in all_schemes.items():
            schemes.append({
                "id": scheme_id,
                "name": scheme_config.get("name", scheme_id),
                "description": scheme_config.get("description", "")
            })
    except Exception as e:
        logger.warning(f"从配置读取序号方案失败: {e}，使用默认方案列表")
        # 降级到硬编码
        schemes = [
            {"id": "gongwen_standard", "name": cli_t("cli.numbering_schemes.gongwen_standard", default="公文标准"), "description": "一、（一）1.（1）①"},
            {"id": "hierarchical_standard", "name": cli_t("cli.numbering_schemes.hierarchical_standard", default="层级数字标准"), "description": "1 1.1 1.1.1"},
            {"id": "hierarchical_h2_start", "name": cli_t("cli.numbering_schemes.hierarchical_h2_start", default="层级数字(H2起)"), "description": "H1无序号，H2起为 1 1.1 1.1.1"},
            {"id": "legal_standard", "name": cli_t("cli.numbering_schemes.legal_standard", default="法律条文标准"), "description": "第一编 第一章 第一节 第一条"},
        ]
    
    if json_mode:
        output = {"schemes": schemes}
        print(json.dumps(output, ensure_ascii=False, indent=2))
    else:
        schemes_label = cli_t("cli.messages.available_schemes", default="可用序号方案")
        example_label = cli_t("cli.messages.example_format", default="示例")
        print(f"\n{schemes_label}:")
        for scheme in schemes:
            print(f"  {scheme['id']}: {scheme['name']}")
            if scheme.get('description'):
                print(f"    {example_label}: {scheme['description']}")
    
    return 0


def list_templates(json_mode: bool = True) -> int:
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
        docx_templates = template_loader.get_available_templates('docx')
        xlsx_templates = template_loader.get_available_templates('xlsx')
        
        templates = {
            "docx": docx_templates or [],
            "xlsx": xlsx_templates or []
        }
        
        if json_mode:
            output = {"templates": templates}
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
            print_json_error("list_templates", "", str(e))
        else:
            error_msg = cli_t("cli.messages.template_list_error", default="错误: 无法获取模板列表")
            print(f"{error_msg} - {e}")
        return 1
