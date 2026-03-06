"""
脚注/尾注 Part 处理模块

确保 Word 模板包含脚注和尾注部分。如果模板没有相关文件，
将创建 footnotes.xml / endnotes.xml 并更新文档关系。

主要功能:
- ensure_notes_parts(): 确保模板包含脚注和尾注部分

技术实现:
- 检查 ZIP 包中是否存在 word/footnotes.xml 或 word/endnotes.xml
- 如果不存在，创建基础 XML 文件（含 separator 和 continuationSeparator）
- 更新 word/_rels/document.xml.rels 添加关系声明
- 更新 [Content_Types].xml 添加内容类型声明
"""

import logging
import tempfile
import zipfile
from pathlib import Path

import lxml.etree as etree

logger = logging.getLogger(__name__)

# OOXML 命名空间
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

# 关系类型
FOOTNOTES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
ENDNOTES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"

# 内容类型
FOOTNOTES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
ENDNOTES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"

# ==============================================================================
# 基础 XML 模板（采用 Word 标准：系统保留 ID 为 -1 和 0）
# ==============================================================================

FOOTNOTES_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:footnote w:type="separator" w:id="-1">
        <w:p>
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
            </w:pPr>
            <w:r>
                <w:separator/>
            </w:r>
        </w:p>
    </w:footnote>
    <w:footnote w:type="continuationSeparator" w:id="0">
        <w:p>
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
            </w:pPr>
            <w:r>
                <w:continuationSeparator/>
            </w:r>
        </w:p>
    </w:footnote>
</w:footnotes>"""

ENDNOTES_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:endnote w:type="separator" w:id="-1">
        <w:p>
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
            </w:pPr>
            <w:r>
                <w:separator/>
            </w:r>
        </w:p>
    </w:endnote>
    <w:endnote w:type="continuationSeparator" w:id="0">
        <w:p>
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
            </w:pPr>
            <w:r>
                <w:continuationSeparator/>
            </w:r>
        </w:p>
    </w:endnote>
</w:endnotes>"""


# ==============================================================================
# 主入口函数
# ==============================================================================


def ensure_notes_parts(template_path: str, cleanup_callback=None) -> str:
    """
    确保模板包含脚注和尾注部分

    参数:
        template_path: 模板文件路径
        cleanup_callback: 可选的清理回调函数，用于注册临时文件以便后续清理

    返回:
        str: 处理后的文件路径（可能是临时文件）
    """
    logger.info("检查模板脚注/尾注部分...")

    try:
        with zipfile.ZipFile(template_path, "r") as zf:
            file_list = zf.namelist()
            has_footnotes = "word/footnotes.xml" in file_list
            has_endnotes = "word/endnotes.xml" in file_list

            if has_footnotes and has_endnotes:
                logger.info("模板已包含脚注和尾注部分")
                return template_path

            # 需要创建缺失的部分
            logger.info(f"需要创建: 脚注={not has_footnotes}, 尾注={not has_endnotes}")

            # 读取需要修改的文件
            rels_xml = zf.read("word/_rels/document.xml.rels")
            content_types_xml = zf.read("[Content_Types].xml")

            # 解析 XML
            rels_root = etree.fromstring(rels_xml)
            ct_root = etree.fromstring(content_types_xml)

            files_to_add = {}

            # 处理脚注
            if not has_footnotes:
                files_to_add["word/footnotes.xml"] = FOOTNOTES_XML_TEMPLATE.encode("utf-8")
                _add_relationship(rels_root, "footnotes.xml", FOOTNOTES_REL_TYPE)
                _add_content_type(ct_root, "/word/footnotes.xml", FOOTNOTES_CONTENT_TYPE)
                logger.debug("准备创建 footnotes.xml")

            # 处理尾注
            if not has_endnotes:
                files_to_add["word/endnotes.xml"] = ENDNOTES_XML_TEMPLATE.encode("utf-8")
                _add_relationship(rels_root, "endnotes.xml", ENDNOTES_REL_TYPE)
                _add_content_type(ct_root, "/word/endnotes.xml", ENDNOTES_CONTENT_TYPE)
                logger.debug("准备创建 endnotes.xml")

            # 创建临时文件
            temp_dir = tempfile.mkdtemp()
            temp_path = str(Path(temp_dir) / Path(template_path).name)

            # 如果提供了清理回调，注册临时目录
            if cleanup_callback:
                cleanup_callback(temp_dir)

            # 写入修改后的文件
            with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
                for item in zf.infolist():
                    if item.filename == "word/_rels/document.xml.rels":
                        modified_xml = etree.tostring(
                            rels_root, xml_declaration=True, encoding="UTF-8", standalone="yes"
                        )
                        zf_out.writestr(item, modified_xml)
                    elif item.filename == "[Content_Types].xml":
                        modified_xml = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
                        zf_out.writestr(item, modified_xml)
                    else:
                        zf_out.writestr(item, zf.read(item.filename))

                # 添加新文件
                for filename, content in files_to_add.items():
                    zf_out.writestr(filename, content)

            logger.info(f"已创建包含脚注/尾注部分的临时模板: {temp_path}")
            return temp_path

    except Exception as e:
        logger.error(f"处理脚注/尾注部分失败: {e}", exc_info=True)
        return template_path


def _add_relationship(rels_root, target: str, rel_type: str):
    """
    向关系文件添加新关系

    参数:
        rels_root: document.xml.rels 的根元素
        target: 目标文件名
        rel_type: 关系类型 URI
    """
    # 获取现有最大 rId
    max_id = 0
    for rel in rels_root.findall(f"{{{RELS_NS}}}Relationship"):
        rid = rel.get("Id", "")
        if rid.startswith("rId"):
            try:
                num = int(rid[3:])
                max_id = max(max_id, num)
            except ValueError:
                pass

    new_id = f"rId{max_id + 1}"

    rel_elem = etree.SubElement(rels_root, f"{{{RELS_NS}}}Relationship")
    rel_elem.set("Id", new_id)
    rel_elem.set("Type", rel_type)
    rel_elem.set("Target", target)

    logger.debug(f"添加关系: {new_id} -> {target}")


def _add_content_type(ct_root, part_name: str, content_type: str):
    """
    向 Content_Types 添加新条目

    参数:
        ct_root: [Content_Types].xml 的根元素
        part_name: 部件名称（如 /word/footnotes.xml）
        content_type: 内容类型
    """
    # 检查是否已存在
    for override in ct_root.findall(f"{{{CONTENT_TYPES_NS}}}Override"):
        if override.get("PartName") == part_name:
            logger.debug(f"内容类型已存在: {part_name}")
            return

    override_elem = etree.SubElement(ct_root, f"{{{CONTENT_TYPES_NS}}}Override")
    override_elem.set("PartName", part_name)
    override_elem.set("ContentType", content_type)

    logger.debug(f"添加内容类型: {part_name}")


def check_notes_parts(docx_path: str) -> dict:
    """
    检查文档是否包含脚注/尾注部分

    参数:
        docx_path: 文档路径

    返回:
        dict: {'has_footnotes': bool, 'has_endnotes': bool}
    """
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            file_list = zf.namelist()
            return {
                "has_footnotes": "word/footnotes.xml" in file_list,
                "has_endnotes": "word/endnotes.xml" in file_list,
            }
    except Exception as e:
        logger.warning(f"检查脚注/尾注部分失败: {e}")
        return {"has_footnotes": False, "has_endnotes": False}
