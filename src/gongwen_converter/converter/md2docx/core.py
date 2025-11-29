"""
MD转DOCX转换器核心模块

统一的临时目录管理和转换流程
"""

import os
import re
import datetime
import logging
import shutil
import tempfile
import threading
from typing import Callable, Optional
from docx import Document

from . import md_processor, docx_processor
from gongwen_converter.template.loader import load_docx_template
from gongwen_converter.utils.validation_utils import is_value_empty
from gongwen_converter.utils.heading_utils import convert_to_halfwidth
from gongwen_converter.docx_spell.core import process_docx
from gongwen_converter.utils.workspace_manager import prepare_input_file, should_save_intermediate_files
from gongwen_converter.utils.path_utils import generate_output_path
from gongwen_converter.utils import yaml_utils

logger = logging.getLogger(__name__)

# 附件序号清理正则
ATTACH_NUM_PATTERN = re.compile(
    r'^[一二三四五六七八九十]+、|'  # 中文序号
    r'^（[一二三四五六七八九十]+）|'  # 带括号中文序号
    r'^\d+[\.．]\s*|'  # 半角数字 + 点
    r'^[０１２３４５６７８９]+[\.．]\s*|'  # 全角数字 + 点
    r'^（\d+）\s*|'  # 带括号半角数字
    r'^（[０１２３４５６７８９]+）\s*|'  # 带括号全角数字
    r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟㊱㊲㊳㊴㊵㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿]\s*'  # 带圈数字
)


def convert(
    md_path: str,
    output_path: str,
    *,
    template_name: str = "公文通用",
    spell_check_option: int = 0,
    progress_callback: Optional[Callable[[str], None]] = None,
    cancel_event: Optional[threading.Event] = None,
    original_source_path: Optional[str] = None
) -> Optional[str]:
    """
    将Markdown文件转换为DOCX文档
    
    统一的临时目录管理流程：
    1. 创建临时目录
    2. 准备input.md副本
    3. 对副本进行转换
    4. 保存到指定输出路径
    5. 自动清理临时目录
    
    参数:
        md_path: Markdown文件路径
        output_path: 输出文件完整路径
        template_name: 模板名称，默认"公文通用"
        spell_check_option: 拼写检查选项，0表示不检查
        progress_callback: 进度回调函数
        cancel_event: 取消事件
        original_source_path: 原始源文件路径（用于嵌入功能的路径解析，可选）
    
    返回:
        str: 输出文件完整路径，失败时返回None
    """
    try:
        logger.info("=" * 60)
        logger.info(f"开始MD转DOCX: {os.path.basename(md_path)}")
        logger.info(f"使用模板: {template_name}")
        logger.info(f"错别字检查选项: {spell_check_option}")
        logger.info("=" * 60)
        
        with tempfile.TemporaryDirectory() as temp_dir:
            # 1. 创建input.md副本
            temp_input = prepare_input_file(md_path, temp_dir, 'md')
            logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
            
            # 2. 读取和解析MD文件（传递原始路径用于路径解析）
            if progress_callback:
                progress_callback("正在解析Markdown文件...")
            
            # 如果提供了original_source_path，使用它；否则使用md_path
            path_for_resolve = original_source_path if original_source_path else md_path
            yaml_data, md_body = _read_and_parse_md(temp_input, path_for_resolve)
            
            if cancel_event and cancel_event.is_set():
                logger.info("操作已取消")
                return None
            
            # 3. 在临时目录生成临时输出文件
            temp_output = os.path.join(temp_dir, "temp_output.docx")
            
            # 4. 在临时目录转换
            success = _convert_internal(
                yaml_data, md_body, temp_output,
                template_name, spell_check_option,
                progress_callback, cancel_event
            )
            
            if not success:
                logger.error("内部转换失败")
                return None
            
            if cancel_event and cancel_event.is_set():
                logger.info("操作已取消")
                return None
            
            # 5. 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            
            # 6. 移动到最终输出路径
            shutil.move(temp_output, output_path)
            
            # 8. 保留中间文件（如需要）
            if should_save_intermediate_files():
                logger.debug("保留中间文件到输出目录")
                for filename in os.listdir(temp_dir):
                    src = os.path.join(temp_dir, filename)
                    if os.path.isfile(src):
                        dst = os.path.join(output_dir, filename)
                        shutil.copy(src, dst)
                        logger.debug(f"保留中间文件: {filename}")
            
            logger.info(f"MD转DOCX成功: {output_path}")
            return output_path
        
    except Exception as e:
        logger.error(f"MD转DOCX失败: {e}", exc_info=True)
        return None


def _read_and_parse_md(temp_md_path: str, original_md_path: str = None) -> tuple:
    """
    读取并解析Markdown文件，返回YAML数据和正文
    
    新增功能：
    - 在返回前展开MD正文中的嵌入链接（如果启用）
    
    参数:
        temp_md_path: 临时Markdown文件路径（用于读取内容）
        original_md_path: 原始Markdown文件路径（用于路径解析，可选）
    
    返回:
        tuple: (yaml_data字典, md_body字符串)
    """
    logger.debug(f"读取Markdown文件: {temp_md_path}")
    with open(temp_md_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 匹配YAML格式部分
    yaml_match = re.search(r'---\n(.*?)\n---', content, re.DOTALL)
    
    if yaml_match:
        yaml_content = yaml_match.group(1).replace('\t', '  ')
        md_body = content[yaml_match.end():].strip()
        logger.debug(f"提取YAML内容: {len(yaml_content)} 字符")
        logger.debug(f"提取正文内容: {len(md_body)} 字符")
    else:
        logger.warning("未找到YAML头部，使用空YAML")
        yaml_content = ""
        md_body = content.strip()
    
    yaml_data = yaml_utils.parse_yaml(yaml_content)
    
    # 新增：处理所有Markdown链接（嵌入和非嵌入，如果启用）
    from gongwen_converter.config.config_manager import config_manager
    if config_manager.is_embedding_enabled():
        try:
            from gongwen_converter.utils.link_embedding import process_markdown_links
            
            # 使用原始文件路径进行路径解析（如果提供），否则使用临时路径
            source_path_for_resolve = original_md_path if original_md_path else temp_md_path
            
            # 1. 处理YAML字段值中的链接
            logger.info("开始处理YAML字段值中的链接...")
            
            def process_yaml_links(data, source_path):
                """递归处理YAML数据结构中的所有链接"""
                if isinstance(data, dict):
                    return {k: process_yaml_links(v, source_path) for k, v in data.items()}
                elif isinstance(data, list):
                    return [process_yaml_links(item, source_path) for item in data]
                elif isinstance(data, str):
                    # 对字符串调用链接处理（depth=0，允许嵌套展开）
                    return process_markdown_links(data, source_path, set(), depth=0)
                else:
                    return data
            
            yaml_data = process_yaml_links(yaml_data, source_path_for_resolve)
            logger.info("YAML字段链接处理完成")
            
            # 2. 处理正文中的链接
            logger.info("开始处理正文中的链接...")
            original_length = len(md_body)
            md_body = process_markdown_links(md_body, source_path_for_resolve)
            new_length = len(md_body)
            logger.info(f"正文链接处理完成 | 长度变化: {original_length} → {new_length}")
        except Exception as e:
            logger.error(f"处理Markdown链接失败: {e}", exc_info=True)
            logger.warning("将继续使用原始MD内容")
    else:
        logger.debug("嵌入功能未启用，跳过链接处理")
    
    return yaml_data, md_body


def _convert_internal(
    yaml_data: dict,
    md_body: str,
    output_path: str,
    template_name: str,
    spell_check_option: int,
    progress_callback,
    cancel_event
) -> bool:
    """
    内部转换实现
    
    复用现有的md2docx模块代码
    
    参数:
        yaml_data: YAML数据字典
        md_body: Markdown正文
        output_path: 输出文件路径
        template_name: 模板名称
        spell_check_option: 拼写检查选项
        progress_callback: 进度回调
        cancel_event: 取消事件
    
    返回:
        bool: 转换是否成功
    """
    try:
        # 1. 预处理DOCX相关的YAML数据
        logger.debug("预处理DOCX相关的YAML数据...")
        process_docx_specific_yaml(yaml_data)
        
        # 2. 获取公文标题
        yaml_data['公文标题'] = _get_title_from_yaml(yaml_data)
        logger.debug(f"设置公文标题: {yaml_data['公文标题']}")
        
        # 3. 处理Markdown正文
        if progress_callback:
            progress_callback("正在处理Markdown正文...")
        logger.debug("处理Markdown正文...")
        processed_body = md_processor.process_md_body(md_body)
        logger.info(f"正文处理完成, 生成 {len(processed_body)} 个段落")
        
        # 4. 加载模板
        if progress_callback:
            progress_callback("正在加载Word模板...")
        logger.debug(f"加载Word模板: {template_name}")
        doc = load_docx_template(template_name)
        logger.info("模板加载成功")
        
        # 5. 替换占位符
        if progress_callback:
            progress_callback("正在填充模板内容...")
        logger.debug("处理模板占位符...")
        success = _process_docx_template(doc, output_path, yaml_data, processed_body)
        
        if not success:
            logger.error("在模板处理过程中发生错误")
            return False
            
        if cancel_event and cancel_event.is_set():
            logger.info("操作被用户取消")
            return False
        
        # 6. 根据选项执行错别字检查
        if spell_check_option is None:
            logger.info("使用配置文件默认设置进行错别字检查")
            process_docx(output_path)
        elif spell_check_option > 0:
            if progress_callback:
                progress_callback("正在执行最终校对...")
            logger.info(f"使用用户指定的规则进行错别字检查: {spell_check_option}")
            process_docx(
                output_path, 
                spell_check_options=spell_check_option, 
                progress_callback=progress_callback, 
                cancel_event=cancel_event
            )
        else:
            logger.info("用户选择不进行错别字检查")
        
        logger.info(f"转换成功! DOCX文件已保存到: {output_path}")
        return True
        
    except Exception as e:
        logger.error(f"内部转换失败: {e}", exc_info=True)
        return False


def _get_title_from_yaml(yaml_data: dict) -> str:
    """从YAML数据中获取公文标题"""
    if '公文标题' in yaml_data and yaml_data['公文标题']:
        return str(yaml_data['公文标题']).strip()
    if 'aliases' in yaml_data:
        aliases = yaml_data['aliases']
        if isinstance(aliases, str):
            return aliases.strip()
        elif isinstance(aliases, list) and aliases:
            return str(aliases[0]).strip()
    logger.warning("未设置公文标题，使用默认标题")
    return "公文标题"


def _process_docx_template(doc: Document, output_path: str, yaml_data: dict, body_data: list) -> bool:
    """处理DOCX模板并保存结果"""
    try:
        logger.debug("替换模板中的占位符...")
        temp_path = docx_processor.replace_placeholders(doc, yaml_data, body_data)
        
        output_dir = os.path.dirname(output_path)
        os.makedirs(output_dir, exist_ok=True)
        
        shutil.move(temp_path, output_path)
        
        if not os.path.exists(output_path):
            logger.error(f"移动文件到最终位置失败: {output_path}")
            return False
            
        return True
    except Exception as e:
        logger.error(f"处理DOCX模板时发生错误: {str(e)}", exc_info=True)
        return False


# --- DOCX相关的YAML处理函数 ---

def process_docx_specific_yaml(data: dict):
    """处理DOCX输出特有的YAML字段"""
    logger.info("处理DOCX特定的YAML字段...")
    process_attachment_description(data)
    process_cc_orgs(data)
    process_special_fields(data)


def process_attachment_description(data: dict):
    """
    处理附件说明字段
    第2行起添加全角空格对齐，配合悬挂缩进使用
    """
    if '附件说明' not in data:
        return
    
    attachments = data['附件说明']
    if is_value_empty(attachments):
        data['附件说明'] = []
        return
    
    if not isinstance(attachments, list):
        attachments = [attachments]
    
    # 清理序号
    cleaned_attachments = []
    for item in attachments:
        content = str(item).strip() if item is not None else ""
        normalized_content = convert_to_halfwidth(content)
        cleaned = ATTACH_NUM_PATTERN.sub('', normalized_content).strip()
        if cleaned == "" and content != "":
            cleaned = content
        cleaned_attachments.append(cleaned)
    
    # 构建附件说明列表
    formatted = []
    for i, content in enumerate(cleaned_attachments, 1):
        if len(cleaned_attachments) == 1:
            # 单附件：附件：<内容>
            formatted.append(f"附件：{content}")
        else:
            # 多附件
            if i == 1:
                formatted.append(f"附件：{i}. {content}")
            else:
                # 第2行起：加3个全角空格（序号10开始，加两个全角空格和一个半角空格）
                indent = "\u3000\u3000\u3000" if i < 10 else "\u3000\u3000 "
                formatted.append(f"{indent}{i}. {content}")
    
    data['附件说明'] = formatted
    logger.debug(f"处理附件说明完成，共 {len(formatted)} 项")


def process_cc_orgs(data: dict):
    """处理抄送机关字段"""
    if '抄送机关' in data:
        cc_orgs = data['抄送机关']
        if is_value_empty(cc_orgs):
            data['抄送机关'] = ""
        else:
            if not isinstance(cc_orgs, list):
                cc_orgs = [cc_orgs]
            
            # 使用format_display_value处理可能的嵌套列表
            from gongwen_converter.utils.text_utils import format_display_value
            valid_orgs = [format_display_value(org).strip() for org in cc_orgs if not is_value_empty(org)]
            data['抄送机关'] = "，".join(valid_orgs)


def process_special_fields(data: dict):
    """处理特殊字段"""
    if '附注' in data:
        data['附注'] = process_notes(data['附注'])
    if '印发日期' in data:
        data['印发日期'] = format_date(data['印发日期'], suffix="印发")
    if '成文日期' in data:
        data['成文日期'] = format_date(data['成文日期'])


def process_notes(notes: str) -> str:
    """处理附注字段"""
    if not notes:
        return ""
    if re.match(r'^[（(].*?[)）]$', notes):
        return notes[1:-1]
    return notes


def format_date(date_str, suffix: str = "") -> str:
    """格式化日期字段"""
    if not date_str:
        return ""
    if isinstance(date_str, (datetime.date, datetime.datetime)):
        date_obj = date_str
    else:
        date_formats = ["%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日", "%Y.%m.%d", "%Y年%m月%d号"]
        date_obj = None
        for fmt in date_formats:
            try:
                date_obj = datetime.datetime.strptime(str(date_str), fmt)
                break
            except ValueError:
                continue
    if date_obj:
        formatted = f"{date_obj.year}年{date_obj.month}月{date_obj.day}日"
        if suffix:
            formatted += suffix
        return formatted
    return str(date_str)
