"""
numbering.xml 处理器模块

确保 Word 文档包含 numbering 部分（word/numbering.xml）。
某些简单模板可能没有 numbering.xml，需要动态创建。

参照 notes_part_handler.py 实现。
"""

import logging
import shutil
import tempfile
import zipfile
from pathlib import Path

import lxml.etree as etree

logger = logging.getLogger(__name__)

# 命名空间定义
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

NUMBERING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
NUMBERING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"

# 空的 numbering.xml 模板
NUMBERING_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
             xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
</w:numbering>"""


def ensure_numbering_part(template_path: str, cleanup_callback=None) -> str:
    """
    确保模板包含 numbering 部分

    如果模板已有 numbering.xml，直接返回原路径。
    如果没有，创建临时副本并添加 numbering.xml。

    参数:
        template_path: 模板文件路径
        cleanup_callback: 清理回调函数，用于注册临时文件清理

    返回:
        str: 处理后的模板路径（可能是临时副本）
    """
    logger.debug(f"检查模板 numbering 部分: {template_path}")

    try:
        # 检查是否已有 numbering.xml
        with zipfile.ZipFile(template_path, "r") as zf:
            has_numbering = "word/numbering.xml" in zf.namelist()

        if has_numbering:
            logger.debug("模板已包含 numbering.xml")
            return template_path

        # 需要添加 numbering.xml
        logger.info("模板缺少 numbering.xml，创建临时副本并添加")

        # 创建临时副本
        temp_dir = tempfile.mkdtemp(prefix="docwen_numbering_")
        temp_path = str(Path(temp_dir) / Path(template_path).name)
        shutil.copy2(template_path, temp_path)

        # 注册清理回调
        if cleanup_callback:
            cleanup_callback(temp_dir)

        # 添加 numbering.xml
        _add_numbering_part(temp_path)

        logger.info(f"已创建包含 numbering.xml 的临时模板: {temp_path}")
        return temp_path

    except Exception as e:
        logger.error(f"确保 numbering 部分时出错: {e}", exc_info=True)
        return template_path


def _add_numbering_part(docx_path: str):
    """
    向 DOCX 文件添加 numbering.xml 部分

    步骤:
    1. 创建 word/numbering.xml
    2. 更新 word/_rels/document.xml.rels 添加关系
    3. 更新 [Content_Types].xml 添加内容类型
    """
    # 读取现有内容
    with zipfile.ZipFile(docx_path, "r") as zf:
        file_contents = {}
        for name in zf.namelist():
            file_contents[name] = zf.read(name)

    # 1. 添加 numbering.xml
    file_contents["word/numbering.xml"] = NUMBERING_XML_TEMPLATE.encode("utf-8")
    logger.debug("添加 word/numbering.xml")

    # 2. 更新 document.xml.rels
    rels_path = "word/_rels/document.xml.rels"
    if rels_path in file_contents:
        rels_content = file_contents[rels_path]
        rels_root = etree.fromstring(rels_content)

        # 检查是否已有 numbering 关系
        existing = rels_root.find(f'.//*[@Type="{NUMBERING_REL_TYPE}"]')
        if existing is None:
            # 生成新的 rId
            existing_ids = [el.get("Id") for el in rels_root if el.get("Id")]
            new_id = _generate_unique_rid(existing_ids)

            # 添加关系
            rel_elem = etree.SubElement(rels_root, f"{{{REL_NS}}}Relationship")
            rel_elem.set("Id", new_id)
            rel_elem.set("Type", NUMBERING_REL_TYPE)
            rel_elem.set("Target", "numbering.xml")

            file_contents[rels_path] = etree.tostring(
                rels_root, xml_declaration=True, encoding="UTF-8", standalone="yes"
            )
            logger.debug(f"在 document.xml.rels 中添加 numbering 关系: {new_id}")

    # 3. 更新 [Content_Types].xml
    ct_path = "[Content_Types].xml"
    if ct_path in file_contents:
        ct_content = file_contents[ct_path]
        ct_root = etree.fromstring(ct_content)

        # 检查是否已有 numbering 内容类型
        existing = ct_root.find('.//*[@PartName="/word/numbering.xml"]')
        if existing is None:
            # 添加 Override
            override_elem = etree.SubElement(ct_root, f"{{{CT_NS}}}Override")
            override_elem.set("PartName", "/word/numbering.xml")
            override_elem.set("ContentType", NUMBERING_CONTENT_TYPE)

            file_contents[ct_path] = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            logger.debug("在 [Content_Types].xml 中添加 numbering 内容类型")

    # 写回文件
    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in file_contents.items():
            zf.writestr(name, content)

    logger.info("numbering.xml 部分添加完成")


def _generate_unique_rid(existing_ids: list) -> str:
    """
    生成唯一的 rId

    参数:
        existing_ids: 现有的 rId 列表

    返回:
        str: 新的唯一 rId
    """
    # 提取现有 ID 的数字部分
    max_num = 0
    for rid in existing_ids:
        if rid and rid.startswith("rId"):
            try:
                num = int(rid[3:])
                max_num = max(max_num, num)
            except ValueError:
                continue

    return f"rId{max_num + 1}"


def get_numbering_xml_root(doc):
    """
    获取文档的 numbering.xml 根元素

    参数:
        doc: python-docx Document 对象

    返回:
        lxml.etree.Element: numbering 根元素，如果不存在返回 None
    """
    try:
        numbering_part = doc.part.numbering_part
        if numbering_part is not None:
            return numbering_part._element
    except (AttributeError, KeyError, NotImplementedError):
        # NotImplementedError: python-docx 在文档没有 numbering.xml 时抛出此异常
        pass

    return None


def get_max_abstract_num_id(numbering_root) -> int:
    """
    获取 numbering.xml 中最大的 abstractNumId

    参数:
        numbering_root: numbering.xml 根元素

    返回:
        int: 最大的 abstractNumId，如果没有返回 -1
    """
    if numbering_root is None:
        return -1

    max_id = -1
    for abstract_num in numbering_root.findall(f".//{{{WORD_NS}}}abstractNum"):
        abstract_num_id = abstract_num.get(f"{{{WORD_NS}}}abstractNumId")
        if abstract_num_id:
            try:
                max_id = max(max_id, int(abstract_num_id))
            except ValueError:
                continue

    return max_id


def get_max_num_id(numbering_root) -> int:
    """
    获取 numbering.xml 中最大的 numId

    参数:
        numbering_root: numbering.xml 根元素

    返回:
        int: 最大的 numId，如果没有返回 0
    """
    if numbering_root is None:
        return 0

    max_id = 0
    for num in numbering_root.findall(f".//{{{WORD_NS}}}num"):
        num_id = num.get(f"{{{WORD_NS}}}numId")
        if num_id:
            try:
                max_id = max(max_id, int(num_id))
            except ValueError:
                continue

    return max_id
