"""
DOCX公式提取处理器

从Word文档中提取Office Math公式并转换为LaTeX格式。

主要功能：
- 检测段落中的OMML（Office Math XML）对象
- 将OMML转换为LaTeX
- 在Markdown中正确插入行内公式和块公式
"""

import logging

import lxml.etree as etree

from docwen.utils.docx_utils import is_inside_fallback
from docwen.utils.formula_utils import OMML_NS, is_formula_supported, mathml_to_latex, omml_to_mathml

logger = logging.getLogger(__name__)

# Word 命名空间
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _is_boolean_format_enabled(rPr, tag_name: str) -> bool:
    """
    检测布尔型格式属性是否启用（如粗体、斜体、删除线）

    Word XML 中的布尔型格式属性规则：
    - <w:b/> 或 <w:b w:val="true"/> 或 <w:b w:val="1"/> = 启用
    - <w:b w:val="false"/> 或 <w:b w:val="0"/> = 显式禁用
    - 元素不存在 = 继承样式（这里视为禁用）

    参数:
        rPr: run properties 元素
        tag_name: 标签名（不含命名空间）

    返回:
        bool: 是否启用
    """
    elem = rPr.find(f"{{{WORD_NS}}}{tag_name}")
    if elem is None:
        return False

    val = elem.get(f"{{{WORD_NS}}}val")
    # 没有 val 属性 = 启用，val="true"/"1"/其他 = 启用
    # val="false"/"0" = 显式禁用
    return val not in ("false", "0")


def _apply_format_markers_from_xml(
    run_elem, text: str, preserve_formatting: bool = True, syntax_config: dict | None = None
) -> str:
    """
    从 XML 层面检测 run 元素的格式属性，应用 Markdown 格式标记

    参数:
        run_elem: lxml 的 run 元素 (<w:r>)
        text: 原始文本
        preserve_formatting: 是否保留格式
        syntax_config: 语法配置字典

    返回:
        str: 添加了 Markdown 标记的文本
    """
    if not preserve_formatting or not text:
        return text

    if syntax_config is None:
        syntax_config = {}

    try:
        # 查找 rPr 元素（run properties）
        rPr = run_elem.find(f"{{{WORD_NS}}}rPr")
        if rPr is None:
            return text

        # 检测各种格式属性（使用正确的布尔检测逻辑）
        is_bold = _is_boolean_format_enabled(rPr, "b")
        is_italic = _is_boolean_format_enabled(rPr, "i")
        is_strike = _is_boolean_format_enabled(rPr, "strike")

        # 下划线检测：<w:u w:val="single"/> = 启用，<w:u w:val="none"/> = 禁用
        is_underline = False
        u_elem = rPr.find(f"{{{WORD_NS}}}u")
        if u_elem is not None:
            u_val = u_elem.get(f"{{{WORD_NS}}}val")
            # val="none" 表示无下划线，其他值（如 "single"）表示有下划线
            is_underline = u_val is not None and u_val != "none"

        is_superscript = False
        is_subscript = False
        is_highlight = rPr.find(f"{{{WORD_NS}}}highlight") is not None

        # 检测上下标
        vertAlign = rPr.find(f"{{{WORD_NS}}}vertAlign")
        if vertAlign is not None:
            val = vertAlign.get(f"{{{WORD_NS}}}val")
            if val == "superscript":
                is_superscript = True
            elif val == "subscript":
                is_subscript = True

        # 按优先级应用格式（从内到外）
        result = text

        # 上标
        if is_superscript:
            result = f"^{result}^" if syntax_config.get("superscript") == "extended" else f"<sup>{result}</sup>"

        # 下标
        if is_subscript:
            result = f"~{result}~" if syntax_config.get("subscript") == "extended" else f"<sub>{result}</sub>"

        # 删除线
        if is_strike:
            result = f"~~{result}~~" if syntax_config.get("strikethrough") == "extended" else f"<del>{result}</del>"

        # 下划线
        if is_underline:
            result = f"<u>{result}</u>"

        # 斜体
        if is_italic:
            result = f"*{result}*" if syntax_config.get("italic") == "asterisk" else f"_{result}_"

        # 粗体
        if is_bold:
            result = f"**{result}**" if syntax_config.get("bold") == "asterisk" else f"__{result}__"

        # 高亮
        if is_highlight:
            result = f"=={result}==" if syntax_config.get("highlight") == "extended" else f"<mark>{result}</mark>"

        return result

    except Exception as e:
        logger.debug(f"从 XML 检测格式时出错: {e}")
        return text


def has_formulas_in_paragraph(paragraph) -> bool:
    """
    检查段落中是否包含公式

    参数:
        paragraph: python-docx的Paragraph对象

    返回:
        bool: 是否包含公式
    """
    try:
        # 访问底层XML
        p_elem = paragraph._element

        # 查找所有oMath元素
        omaths = p_elem.findall(".//{{{}}}oMath".format(OMML_NS["m"]))

        return len(omaths) > 0
    except Exception as e:
        logger.debug(f"检查段落公式时出错: {e}")
        return False


def extract_formulas_from_paragraph(paragraph) -> list[dict]:
    """
    从段落中提取所有公式

    参数:
        paragraph: python-docx的Paragraph对象

    返回:
        list: 公式列表，每项包含:
            - 'omml': OMML XML字符串
            - 'latex': LaTeX字符串
            - 'position': 在段落中的位置索引
            - 'is_inline': 是否为行内公式（通过上下文判断）
    """
    if not is_formula_supported():
        logger.warning("公式转换功能不可用，需要安装latex2mathml")
        return []

    formulas = []

    try:
        # 访问底层XML
        p_elem = paragraph._element

        # 查找所有oMath元素
        all_omaths = p_elem.findall(".//{{{}}}oMath".format(OMML_NS["m"]))

        # 过滤掉位于 mc:Fallback 中的公式（AlternateContent中Choice和Fallback内容相同，只处理Choice）
        omaths = [omath for omath in all_omaths if not is_inside_fallback(omath)]

        if len(omaths) != len(all_omaths):
            logger.debug(f"过滤Fallback公式: 总数{len(all_omaths)} -> 有效{len(omaths)}")

        for idx, omath in enumerate(omaths):
            # 序列化OMML为字符串
            omml_str = etree.tostring(omath, encoding="unicode")

            # OMML → MathML → LaTeX
            mathml_str = omml_to_mathml(omml_str)
            if mathml_str:
                latex_str = mathml_to_latex(mathml_str)
                if latex_str:
                    # 判断是否为行内公式
                    # 简单策略：检查段落文本长度，如果段落有其他文字内容，认为是行内
                    para_text = paragraph.text.strip()
                    is_inline = len(para_text) > 20  # 如果段落文本较长，很可能包含其他内容

                    formulas.append({"omml": omml_str, "latex": latex_str, "position": idx, "is_inline": is_inline})

                    logger.debug(f"提取公式 {idx + 1}: {latex_str[:50]}... (行内: {is_inline})")

    except Exception as e:
        logger.error(f"提取段落公式时出错: {e}", exc_info=True)

    return formulas


def replace_formulas_in_text(text: str, formulas: list[dict]) -> str:
    """
    在文本中替换公式占位符为LaTeX语法

    注意：这个函数假设text中已经用特殊标记（如[[FORMULA_0]]）标记了公式位置

    参数:
        text: 段落文本
        formulas: 公式列表

    返回:
        str: 替换后的文本
    """
    result = text

    for idx, formula in enumerate(formulas):
        placeholder = f"[[FORMULA_{idx}]]"

        latex_syntax = f"${formula['latex']}$"

        result = result.replace(placeholder, latex_syntax)

    return result


def process_paragraph_with_formulas(
    paragraph, preserve_formatting: bool = True, syntax_config: dict | None = None
) -> str | None:
    """
    处理包含公式的段落，返回Markdown格式的文本

    策略：
    1. 遍历段落中的所有 run，按顺序处理
    2. 文本 run 检测格式属性，添加 Markdown 格式标记
    3. 公式对象转换为 LaTeX 语法
    4. 保持公式在文本中的原始位置

    参数:
        paragraph: python-docx的Paragraph对象
        preserve_formatting: 是否保留格式（粗体、斜体等）
        syntax_config: 语法配置字典（控制格式标记语法）

    返回:
        str: Markdown格式的段落文本，失败返回None
    """
    if not is_formula_supported():
        logger.warning("公式转换功能不可用")
        return None

    if syntax_config is None:
        syntax_config = {}

    try:
        # 提取所有公式（用于统计）
        formulas = extract_formulas_from_paragraph(paragraph)

        if not formulas:
            return None

        logger.debug(
            f"段落包含 {len(formulas)} 个公式，开始按顺序处理 runs (preserve_formatting={preserve_formatting})"
        )

        # 检查段落文本长度，判断是纯公式段落还是混合段落
        para_text = paragraph.text.strip()
        text_length = len(para_text)

        # 按顺序遍历段落中的所有子元素（包括 run 和 oMath）
        result_parts = []
        p_elem = paragraph._element

        # 定义命名空间
        w_ns = OMML_NS["w"]
        m_ns = OMML_NS["m"]

        formula_index = 0
        has_text = False  # 标记是否有普通文本

        def process_element(elem):
            """递归处理元素，提取文本和公式"""
            nonlocal formula_index, has_text

            tag = elem.tag

            # 处理文本 run 元素
            if tag == f"{{{w_ns}}}r":
                # 遍历 run 的所有子元素，按顺序处理文本和可能嵌套的公式
                for run_child in elem:
                    if run_child.tag == f"{{{w_ns}}}t":
                        # 处理文本，应用格式标记
                        if run_child.text:
                            formatted_text = _apply_format_markers_from_xml(
                                elem, run_child.text, preserve_formatting, syntax_config
                            )
                            result_parts.append(formatted_text)
                            has_text = True
                            logger.debug(f"添加文本: {run_child.text} -> {formatted_text}")
                    elif run_child.tag == f"{{{m_ns}}}oMath":
                        # 处理嵌套在 run 中的公式（非标准但常见）
                        _process_omath(run_child)

            # 处理公式元素（独立的 oMath）
            elif tag == f"{{{m_ns}}}oMath":
                _process_omath(elem)

            # 处理公式段落元素（oMathPara 包含 oMath）
            elif tag == f"{{{m_ns}}}oMathPara":
                # oMathPara 内部的 oMath 会被单独处理
                for child in elem:
                    process_element(child)

            # 处理超链接、智能标签等容器元素
            elif tag in [f"{{{w_ns}}}hyperlink", f"{{{w_ns}}}smartTag", f"{{{w_ns}}}sdt", f"{{{w_ns}}}sdtContent"]:
                for child in elem:
                    process_element(child)

        def _process_omath(omath_elem):
            """处理单个 oMath 元素"""
            nonlocal formula_index, has_text

            if formula_index < len(formulas):
                formula = formulas[formula_index]
                latex_str = formula["latex"]

                # 判断是否为行内公式：段落文本长度 > 20 或者已有普通文本
                is_inline = text_length > 20 or has_text

                if is_inline:
                    result_parts.append(f"${latex_str}$")
                    logger.debug(f"添加行内公式: ${latex_str[:30]}...")
                else:
                    result_parts.append(f"$$\n{latex_str}\n$$")
                    logger.debug(f"添加块公式: $${latex_str[:30]}...$$")

                formula_index += 1

        # 遍历段落的直接子元素，按顺序处理
        for child in p_elem:
            process_element(child)

        # 如果通过run遍历没有提取到内容，但知道有公式，使用fallback方案
        if not result_parts and formulas:
            logger.warning(f"Run遍历未提取到内容，使用fallback方案处理 {len(formulas)} 个公式")
            # 直接使用提取的公式列表
            for formula in formulas:
                latex_str = formula["latex"]
                # 纯公式段落使用块公式格式（单行格式，避免被分割）
                if text_length <= 20:
                    result_parts.append(f"$${latex_str}$$")
                    logger.debug(f"Fallback: 添加块公式: $${latex_str[:30]}...$$")
                else:
                    result_parts.append(f"${latex_str}$")
                    logger.debug(f"Fallback: 添加行内公式: ${latex_str[:30]}...")

        # 如果仍然没有提取到任何内容，尝试使用段落文本
        if not result_parts and para_text:
            result_parts.append(para_text)
            logger.warning(f"使用段落文本作为最后fallback: {para_text[:50]}...")

        result = "".join(result_parts).strip()

        if result:
            logger.info(f"段落处理完成，包含 {formula_index} 个公式，总长度 {len(result)}")
            return result
        else:
            # 即使result为空，如果有公式也应该返回公式
            if formulas:
                logger.warning(f"结果为空但有 {len(formulas)} 个公式，强制返回第一个公式")
                return f"$${formulas[0]['latex']}$$"
            return None

    except Exception as e:
        logger.error(f"处理公式段落时出错: {e}", exc_info=True)
        return None


def extract_standalone_formulas(paragraph) -> list[str]:
    """
    提取独立的公式（没有其他文本的段落）

    返回块公式的Markdown格式列表

    参数:
        paragraph: python-docx的Paragraph对象

    返回:
        list: 块公式的Markdown字符串列表
    """
    if not is_formula_supported():
        return []

    try:
        # 检查段落是否只包含公式
        para_text = paragraph.text.strip()

        # 提取公式
        formulas = extract_formulas_from_paragraph(paragraph)

        if not formulas:
            return []

        # 如果段落文本很短或为空，认为是独立公式
        if not para_text or len(para_text) < 5:
            result = []
            for formula in formulas:
                # 使用块公式语法
                result.append(f"$$\n{formula['latex']}\n$$")
            return result

        return []

    except Exception as e:
        logger.error(f"提取独立公式时出错: {e}", exc_info=True)
        return []
