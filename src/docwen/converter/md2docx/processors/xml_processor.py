"""
XML处理器模块
处理Word文档的XML内容，深度替换占位符
"""

import os
import zipfile
import tempfile
import logging
import glob
import lxml.etree as etree
from docwen.utils.path_utils import ensure_dir_exists
from docwen.utils import text_utils
from docwen.utils import docx_utils

# 配置日志
logger = logging.getLogger(__name__)

# 使用docx_utils中的命名空间
NAMESPACES = docx_utils.NAMESPACES

def process_docx_file(docx_path, yaml_data):
    """
    处理DOCX文件中的所有占位符
    :param docx_path: DOCX文件路径
    :param yaml_data: 包含替换数据的字典
    :return: 处理后的文件路径
    """
    try:
        logger.info(f"开始处理DOCX文件: {docx_path}")
        
        # 创建临时工作目录
        with tempfile.TemporaryDirectory() as tmpdir:
            logger.info(f"创建临时工作目录: {tmpdir}")
            
            # 解压DOCX文件
            extract_docx(docx_path, tmpdir)
            logger.info(f"解压DOCX文件完成: {docx_path}")
            
            # 处理所有XML文件
            xml_files = glob.glob(os.path.join(tmpdir, "**", "*.xml"), recursive=True)
            logger.info(f"找到 {len(xml_files)} 个XML文件需要处理")
            
            for xml_path in xml_files:
                try:
                    process_xml_file(xml_path, yaml_data)
                    logger.debug(f"成功处理XML文件: {xml_path}")
                except Exception as e:
                    logger.error(f"处理XML文件失败: {xml_path}, 错误: {str(e)}")
            
            # 创建输出文件名
            base_name = os.path.splitext(os.path.basename(docx_path))[0]
            output_name = f"{base_name}_processed.docx"
            output_path = os.path.join(os.path.dirname(docx_path), output_name)
            ensure_dir_exists(os.path.dirname(output_path))
            
            # 删除可能存在的旧文件
            if os.path.exists(output_path):
                logger.warning(f"检测到旧文件存在: {output_path}")
                try:
                    os.remove(output_path)
                    logger.info(f"成功删除旧文件: {output_path}")
                except Exception as e:
                    logger.error(f"删除旧文件失败: {str(e)}")
                    # 尝试重命名避免冲突
                    backup_path = f"{output_path}.backup"
                    os.rename(output_path, backup_path)
                    logger.warning(f"已将旧文件重命名为: {backup_path}")
            
            # 重新打包DOCX文件
            repack_docx(tmpdir, output_path)
            logger.info(f"重新打包DOCX文件完成: {output_path}")
            
            # 检查未替换的占位符
            if has_unreplaced_placeholders(output_path, yaml_data):
                logger.warning("检测到未替换的占位符，执行全局文本替换")
                return global_text_replacement(output_path, yaml_data)
            
            return output_path
    except Exception as e:
        logger.error(f"处理DOCX文件失败: {str(e)}", exc_info=True)
        raise

def extract_docx(docx_path, output_dir):
    """
    解压DOCX文件到指定目录
    :param docx_path: DOCX文件路径
    :param output_dir: 输出目录
    """
    try:
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(output_dir)
        logger.debug(f"解压DOCX文件: {docx_path} -> {output_dir}")
    except Exception as e:
        logger.error(f"解压DOCX文件失败: {str(e)}")
        raise

def repack_docx(source_dir, output_path):
    """
    将处理后的文件重新打包为DOCX
    :param source_dir: 源目录
    :param output_path: 输出文件路径
    """
    try:
        # 确保输出目录存在
        ensure_dir_exists(os.path.dirname(output_path))
        
        # 创建新的ZIP文件
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(source_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    
                    # 计算在ZIP中的相对路径
                    arcname = os.path.relpath(file_path, source_dir)
                    
                    # 写入ZIP文件
                    zipf.write(file_path, arcname)
                    logger.debug(f"添加文件到ZIP: {arcname}")
        
        logger.info(f"成功重新打包DOCX文件: {output_path}")
    except Exception as e:
        logger.error(f"重新打包DOCX文件失败: {str(e)}")
        raise

def process_xml_file(xml_path, yaml_data):
    """
    处理单个XML文件中的占位符
    :param xml_path: XML文件路径
    :param yaml_data: 包含替换数据的字典
    """
    try:
        # 读取XML内容
        with open(xml_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        logger.debug(f"读取XML文件: {xml_path}, 大小: {len(content)} 字节")
        
        # 解析XML
        parser = etree.XMLParser(remove_blank_text=True)
        root = etree.fromstring(content.encode('utf-8'), parser)
        
        # 处理所有可能的文本框类型
        process_all_textboxes(root, yaml_data)
        
        # 查找所有文本节点
        text_nodes = root.xpath('//w:t', namespaces=NAMESPACES)
        logger.debug(f"找到 {len(text_nodes)} 个文本节点")
        
        # 替换占位符
        for node in text_nodes:
            if node.text:
                # 使用统一的占位符处理函数
                new_text, should_delete, should_clear = text_utils.replace_placeholder(node.text, yaml_data)
                node.text = new_text
                
                # 记录需要特殊处理的情况
                if should_delete or should_clear:
                    logger.warning(f"在 {xml_path} 中检测到需要特殊处理的占位符")
        
        # 保存修改后的XML
        with open(xml_path, 'w', encoding='utf-8') as f:
            f.write(etree.tostring(root, pretty_print=True, encoding='unicode'))
        
        logger.info(f"成功处理XML文件: {xml_path}")
    except Exception as e:
        logger.error(f"处理XML文件失败: {xml_path}, 错误: {str(e)}")
        raise

def process_all_textboxes(root, yaml_data):
    """
    处理所有类型的文本框，返回是否进行了替换
    """
    replaced = False
    
    # 处理DrawingML文本框
    if process_drawingml_textboxes(root, yaml_data):
        replaced = True
        logger.debug("已处理DrawingML文本框")
    
    # 处理VML文本框
    if process_vml_textboxes(root, yaml_data):
        replaced = True
        logger.debug("已处理VML文本框")
    
    # 处理AlternateContent结构中的文本框
    if process_alternate_content(root, yaml_data):
        replaced = True
        logger.debug("已处理AlternateContent文本框")
        
    return replaced

def process_drawingml_textboxes(root, yaml_data):
    """
    处理DrawingML格式的文本框
    """
    textboxes = root.xpath('//wps:txbx', namespaces=NAMESPACES)
    logger.debug(f"找到 {len(textboxes)} 个DrawingML文本框")
    
    for textbox in textboxes:
        text_nodes = textbox.xpath('.//w:t', namespaces=NAMESPACES)
        for node in text_nodes:
            if node.text:
                new_text, _, _ = text_utils.replace_placeholder(node.text, yaml_data)
                node.text = new_text
                logger.debug(f"替换DrawingML文本框文本: {new_text[:50]}...")
    
    return len(textboxes) > 0

def process_vml_textboxes(root, yaml_data):
    """
    处理VML格式的文本框
    """
    textboxes = root.xpath('//v:textbox', namespaces=NAMESPACES)
    logger.debug(f"找到 {len(textboxes)} 个VML文本框")
    
    for textbox in textboxes:
        text_nodes = textbox.xpath('.//w:t', namespaces=NAMESPACES)
        for node in text_nodes:
            if node.text:
                new_text, _, _ = text_utils.replace_placeholder(node.text, yaml_data)
                node.text = new_text
                logger.debug(f"替换VML文本框文本: {new_text[:50]}...")
    
    return len(textboxes) > 0

def process_alternate_content(root, yaml_data):
    """
    处理mc:AlternateContent结构中的文本框
    """
    alt_content_nodes = root.xpath('//mc:AlternateContent', namespaces=NAMESPACES)
    logger.debug(f"找到 {len(alt_content_nodes)} 个AlternateContent元素")
    
    replaced = False
    
    for alt_content in alt_content_nodes:
        # 处理Choice部分 (wps:txbx)
        choice_nodes = alt_content.xpath('.//mc:Choice', namespaces=NAMESPACES)
        for choice in choice_nodes:
            txbx_nodes = choice.xpath('.//wps:txbx', namespaces=NAMESPACES)
            for txbx in txbx_nodes:
                text_nodes = txbx.xpath('.//w:t', namespaces=NAMESPACES)
                for node in text_nodes:
                    if node.text:
                        new_text, _, _ = text_utils.replace_placeholder(node.text, yaml_data)
                        node.text = new_text
                        replaced = True
                        logger.debug(f"替换Choice文本框文本: {new_text[:50]}...")
        
        # 处理Fallback部分 (v:textbox)
        fallback_nodes = alt_content.xpath('.//mc:Fallback', namespaces=NAMESPACES)
        for fallback in fallback_nodes:
            v_textbox_nodes = fallback.xpath('.//v:textbox', namespaces=NAMESPACES)
            for v_textbox in v_textbox_nodes:
                text_nodes = v_textbox.xpath('.//w:t', namespaces=NAMESPACES)
                for node in text_nodes:
                    if node.text:
                        new_text, _, _ = text_utils.replace_placeholder(node.text, yaml_data)
                        node.text = new_text
                        replaced = True
                        logger.debug(f"替换Fallback文本框文本: {new_text[:50]}...")
    
    return replaced

def has_unreplaced_placeholders(docx_path, yaml_data):
    """
    检查DOCX文件中是否还有未替换的占位符
    :param docx_path: DOCX文件路径
    :param yaml_data: 包含替换数据的字典
    :return: 如果找到未替换的占位符返回True，否则False
    """
    try:
        logger.debug(f"检查未替换的占位符: {docx_path}")
        with tempfile.TemporaryDirectory() as tmpdir:
            extract_docx(docx_path, tmpdir)
            
            xml_files = glob.glob(os.path.join(tmpdir, "**", "*.xml"), recursive=True)
            for xml_path in xml_files:
                with open(xml_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    for key in yaml_data.keys():
                        placeholder = f'{{{{{key}}}}}'
                        if placeholder in content:
                            logger.warning(f"在 {xml_path} 中检测到未替换的占位符: {placeholder}")
                            return True
        return False
    except Exception as e:
        logger.error(f"检查未替换占位符失败: {str(e)}")
        return False

def global_text_replacement(docx_path, yaml_data):
    """
    全局文本替换作为兜底方案
    :param docx_path: DOCX文件路径
    :param yaml_data: 包含替换数据的字典
    """
    try:
        logger.warning("启动全局文本替换作为兜底方案")
        with tempfile.TemporaryDirectory() as tmpdir:
            extract_docx(docx_path, tmpdir)
            
            xml_files = glob.glob(os.path.join(tmpdir, "**", "*.xml"), recursive=True)
            for xml_path in xml_files:
                try:
                    with open(xml_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    original_content = content
                    replacements_done = 0
                    
                    for key, value in yaml_data.items():
                        placeholder = f'{{{{{key}}}}}'
                        if placeholder in content:
                            display_value = text_utils.format_display_value(value)
                            content = content.replace(placeholder, str(display_value))
                            replacements_done += 1
                            logger.info(f"全局替换占位符: {placeholder} -> {display_value}")
                    
                    # 只有实际进行了替换才保存
                    if replacements_done > 0:
                        with open(xml_path, 'w', encoding='utf-8') as f:
                            f.write(content)
                        logger.info(f"在 {xml_path} 中完成 {replacements_done} 处替换")
                    
                except Exception as e:
                    logger.error(f"全局替换失败: {xml_path}, 错误: {str(e)}")
            
            # 创建输出文件名
            base_name = os.path.splitext(os.path.basename(docx_path))[0]
            output_name = f"{base_name}_global_replaced.docx"
            output_path = os.path.join(os.path.dirname(docx_path), output_name)
            ensure_dir_exists(os.path.dirname(output_path))
            
            # 删除可能存在的旧文件
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                    logger.info(f"删除旧全局替换文件: {output_path}")
                except:
                    pass
            
            # 重新打包DOCX文件
            repack_docx(tmpdir, output_path)
            logger.info(f"全局兜底替换完成: {output_path}")
            
            return output_path
    except Exception as e:
        logger.error(f"全局文本替换失败: {str(e)}")
        return docx_path

