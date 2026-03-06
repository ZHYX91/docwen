"""
Markdown处理器模块
处理MD文件中YAML头部之后的内容，转换为结构化的段落数据
包括：小标题、正文文本、表格、列表、脚注/尾注等
支持标题序号配置化处理
"""

import hashlib
import logging
import re

from docwen.utils.heading_numbering import HeadingFormatter
from docwen.utils.heading_utils import remove_heading_numbering
from docwen.utils.markdown_utils import extract_markdown_tables

from ..handlers.formula_handler import has_latex_formulas
from ..handlers.note_handler import extract_notes

# 配置日志
logger = logging.getLogger(__name__)

# 半角标点集合（英文标点，后面需要加空格）
# 用于标题与正文合并时判断是否需要添加空格
HALFWIDTH_PUNCT_NEED_SPACE = {".", ",", ";", ":", "!", "?"}

# 定义符号列表，用于判断标题末尾是否有符号（支持多语言）
# 包含：中文、英文、日文、全角、阿拉伯文、希腊文、西班牙文等语言的常用标点
PUNCTUATION_SET = {
    # 中文标点
    "。",
    "，",
    "；",
    "：",
    "！",
    "？",
    "、",
    "—",  # 全角破折号
    "～",  # 全角波浪号
    "…",  # 省略号
    # 英文/拉丁语系标点
    ".",
    ",",
    ";",
    ":",
    "!",
    "?",
    "–",  # 短破折号（en dash）
    "-",  # 连字符/半角减号
    # 全角英文标点（日文/韩文等常用）
    "．",
    # 阿拉伯文标点
    "؟",
    "؛",
    "،",
    "۔",
    # 希腊文标点（; 作为问号）
    "·",
    # 西班牙文倒置标点（虽然通常在句首，但也可能出现在句末）
    "¿",
    "¡",
    # 俄语标点（与英文相同，省略号已在中文部分添加）
    # 亚美尼亚文标点
    "։",
    "՞",
    "՜",
    # 缅甸文标点
    "။",
    "၊",
    # 埃塞俄比亚文标点
    "።",
    "፣",
    "፧",
    # 藏文标点
    "།",
    "༎",
    # 泰文标点
    "ฯ",
    # 高棉文标点
    "។",
    "៕",
}

# 标题匹配正则（支持1-9级）
HEADING_REGEX = re.compile(r"^(#{1,9})\s+(.*)")

# 引用块匹配正则（匹配任意数量的 >，支持紧凑格式 >>> 和分隔格式 > > >）
# 超过9级在代码中处理为正文
QUOTE_REGEX = re.compile(r"^((?:>\s*)+)(.*)")

# 分隔符/水平线匹配正则（三种类型）
HR_DASH_REGEX = re.compile(r"^-{3,}\s*$")  # --- 短横线（可能是Setext二级标题或分隔符）
HR_ASTERISK_REGEX = re.compile(r"^\*{3,}\s*$")  # *** 星号
HR_UNDERSCORE_REGEX = re.compile(r"^_{3,}\s*$")  # ___ 下划线

# Setext 标题匹配正则
SETEXT_H1_REGEX = re.compile(r"^=+\s*$")  # === 一级标题
SETEXT_H2_REGEX = re.compile(r"^-+\s*$")  # --- 二级标题（与HR_DASH_REGEX相同，需要上下文判断）

# 列表项匹配正则
# 无序列表：以 -、*、+ 开头，前面可以有空格（每2个空格或1个制表符表示一级缩进）
UNORDERED_LIST_REGEX = re.compile(r"^(\s*)([-*+])\s+(.*)$")
# 有序列表：以数字加点开头，前面可以有空格
ORDERED_LIST_REGEX = re.compile(r"^(\s*)(\d+)\.\s+(.*)$")

# 列表续行内容检测：只有空格/制表符开头的行
LIST_CONTINUATION_REGEX = re.compile(r"^(\s+)(.*)$")


def _calculate_list_level_relative(indent_str: str, indent_stack: list) -> int:
    """
    根据相对缩进计算列表级别

    使用相对缩进检测，而不是固定空格数，以适应用户不同的缩进习惯。

    策略：
    - 缩进为0 → level 0
    - 缩进比栈顶大 → 新级别 = 栈顶级别 + 1
    - 缩进与栈中某项相同 → 返回该级别
    - 缩进比栈中所有项都小 → 重置为 level 0

    参数:
        indent_str: 缩进字符串（空格或制表符）
        indent_stack: 缩进栈，用于追踪 [(缩进长度, 级别), ...]

    返回:
        int: 列表级别（0为第一级，1为第二级，以此类推，最大8）
    """
    # 将制表符转换为4空格（标准处理）
    indent_len = len(indent_str.replace("\t", "    "))

    if indent_len == 0:
        # 无缩进，重置栈，返回 level 0
        indent_stack.clear()
        indent_stack.append((0, 0))
        return 0

    # 从栈底向栈顶检查，找到匹配或合适的位置
    for i, (stack_indent, stack_level) in enumerate(indent_stack):
        if indent_len == stack_indent:
            # 缩进相同，返回该级别，并截断栈（移除更深的级别）
            del indent_stack[i + 1 :]
            return stack_level
        elif indent_len < stack_indent:
            # 缩进更小，说明回退到更浅的级别
            if i == 0:
                # 比栈中最小的还小，重置为 level 0
                indent_stack.clear()
                indent_stack.append((indent_len, 0))
                return 0
            else:
                # 回退到上一个级别
                del indent_stack[i:]
                # 添加新的缩进点（可能是不同的缩进值但相同级别）
                prev_level = indent_stack[-1][1]
                new_level = prev_level + 1
                if new_level > 8:
                    new_level = 8
                indent_stack.append((indent_len, new_level))
                return new_level

    # 缩进比栈顶大，创建新级别
    new_level = indent_stack[-1][1] + 1 if indent_stack else 0
    if new_level > 8:
        new_level = 8  # 最大9级（0-8）
    indent_stack.append((indent_len, new_level))
    return new_level


def _calculate_list_level(indent_str: str, indent_spaces: int = 2) -> int:
    """
    根据缩进字符串计算列表级别（旧版本，保留兼容）

    参数:
        indent_str: 缩进字符串（空格或制表符）
        indent_spaces: 每级缩进的空格数（默认2，可从配置获取）

    返回:
        int: 列表级别（0为第一级，1为第二级，以此类推）
    """
    if not indent_str:
        return 0

    # 将制表符转换为相应空格数
    normalized = indent_str.replace("\t", " " * indent_spaces)

    # 根据配置的空格数计算级别
    level = len(normalized) // indent_spaces

    return level


def _determine_hr_attach_mode(lines: list, current_idx: int, total_lines: int) -> str:
    """
    判断分隔符/水平线应该附加到哪个相邻元素

    用于实现 MD→DOCX 转换时，根据分隔符前后是否有空行决定：
    - 如果前面没有空行，附加到前一个段落（与前一段落同一 Word 段落）
    - 否则，作为独立段落

    参数:
        lines: 所有行的列表
        current_idx: 当前分隔符所在行的索引
        total_lines: 总行数

    返回:
        str: 附加模式
            - 'previous': 附加到前一个段落末尾（前面没有空行）
            - 'none': 作为独立段落（前面有空行或位于开头）
    """
    # 检查前一行是否为空行（或者当前是第一行）
    # 向前查找第一个非空行
    prev_non_empty_idx = current_idx - 1
    while prev_non_empty_idx >= 0 and not lines[prev_non_empty_idx].strip():
        prev_non_empty_idx -= 1

    # 如果当前分隔符紧跟在非空行之后（没有空行间隔）
    has_blank_before = (prev_non_empty_idx < 0) or (prev_non_empty_idx < current_idx - 1)

    if not has_blank_before:
        # 前面没有空行，应该附加到前一个段落
        return "previous"
    else:
        # 前面有空行或位于开头，作为独立段落
        return "none"


def process_md_body(
    md_body: str,
    remove_numbering: bool = False,
    add_numbering: bool = False,
    formatter: HeadingFormatter | None = None,
) -> list:
    """
    处理YAML后的Markdown内容 - 支持序号配置化

    参数:
        md_body: YAML头部之后的所有Markdown内容（包含小标题、正文文本、表格等）
        remove_numbering: 是否清除Markdown中的原有序号（默认False）
        add_numbering: 是否添加新序号（默认False）
        formatter: HeadingFormatter实例，用于生成序号（当add_numbering=True时使用）

    返回:
        list: 处理后的段落列表 [{
            'text': 文本内容,
            'level': 标题级别 (0表示正文文本),
            'type': 段落类型 ('heading'/'heading_with_content'/'content'/'table'),
            'table_data': 表格数据 (仅type='table'时存在),
            'has_formula': 是否包含公式 (可选)
        }]

    使用说明:
        - remove_numbering=False, add_numbering=False: 保持原样（默认行为）
        - remove_numbering=True, add_numbering=False: 只清除序号
        - remove_numbering=True, add_numbering=True: 清除旧序号，添加新序号
        - remove_numbering=False, add_numbering=True: 保留原序号，添加新序号（可能重复）
    """
    # 统一换行符为 Unix 格式（支持 Windows \r\n 和旧版 Mac \r）
    # 这对于表格行号计算和正则匹配至关重要
    md_body = md_body.replace("\r\n", "\n").replace("\r", "\n")

    logger.info(f"开始处理Markdown内容（序号配置：清除={remove_numbering}, 添加={add_numbering}）...")

    if add_numbering and formatter:
        logger.info("将使用序号方案生成标题序号")
        formatter.reset_counters()  # 重置计数器

    # 记录Markdown内容初始状态
    body_hash = hashlib.md5(md_body.encode("utf-8")).hexdigest()[:8]
    logger.debug(f"Markdown内容初始状态 | 长度: {len(md_body)} 字符 | 哈希: {body_hash}")

    if not md_body:
        logger.warning("Markdown内容为空")
        return []

    # 列表缩进栈，用于相对级别计算
    # 格式: [(缩进长度, 级别), ...]
    list_indent_stack = []

    # 提取所有表格
    tables = extract_markdown_tables(md_body)
    logger.info(f"提取到 {len(tables)} 个表格")

    # 创建表格行号映射（用于跳过表格行）
    table_line_ranges = {}
    for table_idx, table in enumerate(tables):
        start_line = table["start_line"]
        end_line = table["end_line"]
        for line_no in range(start_line, end_line):
            table_line_ranges[line_no] = table_idx
        logger.debug(f"表格 {table_idx + 1} 占据行 {start_line}-{end_line - 1}")

    processed = []  # 处理后的段落列表

    # 分割为行，保留空行
    lines = md_body.split("\n")
    logger.debug(f"Markdown内容总行数: {len(lines)}")

    i = 0
    n = len(lines)

    # 状态跟踪变量
    heading_count = 0
    content_count = 0
    combined_count = 0

    # 记录上一个有效（非空）行的行号，用于 Setext 标题检测
    # Setext 标题要求文本和 ===/-—- 之间不能有空行
    last_content_line_no = -1

    while i < n:
        line = lines[i].strip()

        # 检查当前行是否是表格的一部分
        if i in table_line_ranges:
            table_idx = table_line_ranges[i]
            table = tables[table_idx]

            # 如果是表格的第一行，插入表格段落
            if i == table["start_line"]:
                processed.append({"type": "table", "level": 0, "table_data": table})
                logger.info(f"在行 {i} 插入表格 {table_idx + 1}: {len(table['headers'])}列 x {len(table['rows'])}行")

            # 跳过表格行
            i += 1
            continue

        # 跳过空行
        if not line:
            i += 1
            continue

        # 检查是否是代码块的开始（``` 或 ~~~，可能有缩进）
        original_line = lines[i]
        code_fence_match = re.match(r"^(\s*)(```|~~~)(.*?)$", original_line)
        if code_fence_match:
            code_indent_str = code_fence_match.group(1)
            fence_char = code_fence_match.group(2)
            # 提取语言标记（如果有）
            language = code_fence_match.group(3).strip()

            # 收集代码块内容
            code_lines = []
            i += 1

            while i < n:
                current_line = lines[i]
                # 检查是否是代码块结束（同样可能有缩进）
                if current_line.strip().startswith(fence_char[:3]):
                    break
                # 保留原始行（包括缩进）
                code_lines.append(current_line)
                i += 1

            # 去除代码内容的公共前导空格（dedent）
            # 在 MD 中，列表内的代码块通常有缩进以保持对齐
            # 转换时需要去除这些缩进空格
            if code_lines:
                # 计算所有非空行的最小前导空格数
                min_indent = float("inf")
                for line in code_lines:
                    if line.strip():  # 只检查非空行
                        leading_spaces = len(line) - len(line.lstrip())
                        min_indent = min(min_indent, leading_spaces)

                # 如果有公共缩进，去除它
                if min_indent > 0 and min_indent != float("inf"):
                    code_lines = [line[min_indent:] if len(line) >= min_indent else line for line in code_lines]
                    logger.debug(f"去除代码块公共缩进: {min_indent} 空格")

            # 合并代码内容
            code_content = "\n".join(code_lines)

            # 检查是否在列表上下文中（根据代码块的缩进）
            code_block_para = {
                "text": code_content,
                "level": 0,
                "type": "code_block",
                "language": language if language else None,
            }

            # 如果有缩进且在列表上下文中，记录列表上下文信息
            if list_indent_stack and code_indent_str:
                indent_len = len(code_indent_str.replace("\t", "    "))
                # 找到该缩进对应的列表级别
                for stack_indent, stack_level in reversed(list_indent_stack):
                    if indent_len >= stack_indent:
                        code_block_para["in_list_context"] = True
                        code_block_para["list_level"] = stack_level
                        logger.debug(f"代码块在列表上下文中: level={stack_level}, indent={indent_len}")
                        break

            # 创建代码块段落
            processed.append(code_block_para)
            content_count += 1
            logger.debug(f"添加代码块段落（{len(code_lines)}行，语言: {language or '无'}）")

            i += 1
            continue

        # 处理Markdown标题行 (使用正则匹配)
        heading_match = HEADING_REGEX.match(line)
        if heading_match:
            hashes = heading_match.group(1)
            title_text = heading_match.group(2).strip()
            heading_level = len(hashes)

            # 处理标题文本
            final_text = title_text

            # 步骤1：清除原序号（如果配置）
            if remove_numbering:
                cleaned_text = remove_heading_numbering(title_text)
                if not cleaned_text.strip():
                    # 如果清除后为空，保留原标题
                    cleaned_text = title_text
                    logger.warning(f"标题清除序号后为空，保留原标题: '{title_text}'")
                else:
                    logger.debug(f"清除标题序号: '{title_text}' -> '{cleaned_text}'")
                final_text = cleaned_text

            # 步骤2：添加新序号（如果配置）
            if add_numbering and formatter:
                formatter.increment_level(heading_level)
                numbering = formatter.format_heading(heading_level)
                final_text = numbering + final_text
                logger.debug(f"添加序号: -> '{final_text}'")

            heading_count += 1

            # 检查标题末尾是否有符号（用于判断是否与正文合并）
            if final_text and final_text[-1] in PUNCTUATION_SET:
                # 有符号，检查下一行是否是普通正文文本（没有空行，且不是特殊Markdown元素）
                can_merge = False
                if i + 1 < n and lines[i + 1].strip():
                    next_line_content = lines[i + 1].strip()
                    next_line_raw = lines[i + 1]  # 保留原始行（列表检测需要缩进信息）

                    # 排除所有特殊 Markdown 元素
                    is_next_heading = next_line_content.startswith("#")
                    is_next_formula = next_line_content.startswith("$$")
                    is_next_table = next_line_content.startswith("|")
                    is_next_quote = next_line_content.startswith(">")
                    is_next_code_block = next_line_content.startswith(("```", "~~~"))
                    is_next_unordered_list = UNORDERED_LIST_REGEX.match(next_line_raw)
                    is_next_ordered_list = ORDERED_LIST_REGEX.match(next_line_raw)
                    is_next_hr = bool(re.match(r"^[-*_]{3,}\s*$", next_line_content))

                    is_special_element = (
                        is_next_heading
                        or is_next_formula
                        or is_next_table
                        or is_next_quote
                        or is_next_code_block
                        or is_next_unordered_list
                        or is_next_ordered_list
                        or is_next_hr
                    )

                    can_merge = not is_special_element

                if can_merge:
                    # 获取正文内容
                    content_text = lines[i + 1].strip()

                    # 如果标题末尾是半角标点（英文），在正文前添加空格
                    if final_text[-1] in HALFWIDTH_PUNCT_NEED_SPACE:
                        content_text = " " + content_text
                        logger.debug("半角标点后自动添加空格")

                    processed.append(
                        {
                            "text": final_text,
                            "level": heading_level,
                            "type": "heading_with_content",
                            "content": content_text,
                        }
                    )
                    i += 1  # 跳过正文行
                    combined_count += 1
                    logger.debug(f"创建组合标题段落 ({heading_level}级): {final_text[:30]}... + 正文")
                else:
                    processed.append({"text": final_text, "level": heading_level, "type": "heading"})
                    logger.debug(f"创建独立标题段落 ({heading_level}级): {final_text[:30]}...")
            else:
                # 没有符号，即使下一行是正文也不合并
                processed.append({"text": final_text, "level": heading_level, "type": "heading"})
                logger.debug(f"创建独立标题段落 ({heading_level}级): {final_text[:30]}...")

                # 下一行是正文时，单独处理（但需要排除特殊Markdown元素，让它们走正常处理流程）
                if i + 1 < n:
                    next_line_content = lines[i + 1].strip()
                    next_line_raw = lines[i + 1]  # 保留原始行（列表检测需要缩进信息）

                    # 排除所有特殊 Markdown 元素
                    is_next_heading = next_line_content.startswith("#")
                    is_next_formula = next_line_content.startswith("$$")
                    is_next_table = next_line_content.startswith("|")
                    is_next_quote = next_line_content.startswith(">")
                    is_next_code_block = next_line_content.startswith(("```", "~~~"))
                    is_next_unordered_list = UNORDERED_LIST_REGEX.match(next_line_raw)
                    is_next_ordered_list = ORDERED_LIST_REGEX.match(next_line_raw)
                    # 分隔符/水平线：三个或更多的 -, *, _（可能有空格）
                    is_next_hr = bool(re.match(r"^[-*_]{3,}\s*$", next_line_content))

                    is_special_element = (
                        is_next_heading
                        or is_next_formula
                        or is_next_table
                        or is_next_quote
                        or is_next_code_block
                        or is_next_unordered_list
                        or is_next_ordered_list
                        or is_next_hr
                    )

                    if next_line_content and not is_special_element:
                        processed.append({"text": next_line_content, "level": 0, "type": "content"})
                        i += 1
                        content_count += 1
                        logger.debug("标题后添加正文文本")

        else:  # 正文文本段落
            line_content = line.strip()

            # ==== Setext 标题检测 ====
            # Setext 标题语法：前一行是普通文本，当前行是 === 或 ---
            # === → 一级标题, --- → 二级标题

            # 检查是否是 Setext 一级标题 (===)
            if SETEXT_H1_REGEX.match(line_content):
                # 检查前一行是否是可转换为标题的普通文本
                # 【重要】Setext 标题要求文本和 === 之间不能有空行
                # 通过检查 last_content_line_no 是否等于 i-1 来判断
                if processed and processed[-1].get("type") == "content" and last_content_line_no == i - 1:
                    # 将前一个 content 段落转换为 heading
                    prev_para = processed[-1]
                    prev_text = prev_para.get("text", "")

                    # 处理标题文本
                    final_text = prev_text

                    # 清除原序号（如果配置）
                    if remove_numbering:
                        cleaned_text = remove_heading_numbering(prev_text)
                        if cleaned_text.strip():
                            final_text = cleaned_text
                            logger.debug(f"Setext H1 清除序号: '{prev_text}' -> '{cleaned_text}'")

                    # 添加新序号（如果配置）
                    if add_numbering and formatter:
                        formatter.increment_level(1)
                        numbering = formatter.format_heading(1)
                        final_text = numbering + final_text
                        logger.debug(f"Setext H1 添加序号: -> '{final_text}'")

                    # 更新段落类型为 heading
                    prev_para["type"] = "heading"
                    prev_para["level"] = 1
                    prev_para["text"] = final_text

                    heading_count += 1
                    content_count -= 1  # 之前计为 content，现在转为 heading
                    logger.debug(f"Setext 一级标题: {final_text[:50]}...")
                    i += 1
                    continue
                else:
                    # 前面没有可转换的文本，=== 作为普通文本处理
                    logger.debug("=== 前面没有可转换的普通文本，当作普通内容处理")
                    # 继续后续处理（会被当作普通文本）

            # 检查是否是 --- （可能是 Setext 二级标题或分隔符）
            if HR_DASH_REGEX.match(line_content):
                # 检查前一行是否是可转换为标题的普通文本
                # 【重要】Setext 标题要求文本和 --- 之间不能有空行
                # 通过检查 last_content_line_no 是否等于 i-1 来判断
                if processed and processed[-1].get("type") == "content" and last_content_line_no == i - 1:
                    # 将前一个 content 段落转换为 Setext 二级标题
                    prev_para = processed[-1]
                    prev_text = prev_para.get("text", "")

                    # 处理标题文本
                    final_text = prev_text

                    # 清除原序号（如果配置）
                    if remove_numbering:
                        cleaned_text = remove_heading_numbering(prev_text)
                        if cleaned_text.strip():
                            final_text = cleaned_text
                            logger.debug(f"Setext H2 清除序号: '{prev_text}' -> '{cleaned_text}'")

                    # 添加新序号（如果配置）
                    if add_numbering and formatter:
                        formatter.increment_level(2)
                        numbering = formatter.format_heading(2)
                        final_text = numbering + final_text
                        logger.debug(f"Setext H2 添加序号: -> '{final_text}'")

                    # 更新段落类型为 heading
                    prev_para["type"] = "heading"
                    prev_para["level"] = 2
                    prev_para["text"] = final_text

                    heading_count += 1
                    content_count -= 1  # 之前计为 content，现在转为 heading
                    logger.debug(f"Setext 二级标题: {final_text[:50]}...")
                    i += 1
                    continue
                else:
                    # 前面没有可转换的文本，--- 作为分隔符处理
                    attach_to = _determine_hr_attach_mode(lines, i, n)
                    processed.append({"type": "horizontal_rule", "hr_type": "dash", "level": 0, "attach_to": attach_to})
                    content_count += 1
                    logger.debug(f"添加分隔符段落（dash类型）: ---, attach_to={attach_to}")
                    i += 1
                    continue
            elif HR_ASTERISK_REGEX.match(line_content):
                attach_to = _determine_hr_attach_mode(lines, i, n)
                processed.append({"type": "horizontal_rule", "hr_type": "asterisk", "level": 0, "attach_to": attach_to})
                content_count += 1
                logger.debug(f"添加分隔符段落（asterisk类型）: ***, attach_to={attach_to}")
                i += 1
                continue
            elif HR_UNDERSCORE_REGEX.match(line_content):
                attach_to = _determine_hr_attach_mode(lines, i, n)
                processed.append(
                    {"type": "horizontal_rule", "hr_type": "underscore", "level": 0, "attach_to": attach_to}
                )
                content_count += 1
                logger.debug(f"添加分隔符段落（underscore类型）: ___, attach_to={attach_to}")
                i += 1
                continue

            # 检查是否是块公式的开始
            # 情况1: 单独的 $$ 行（多行块公式的起始）
            # 情况2: $$xxx 但不以 $$ 结尾（如 $$\int 形式的多行块公式）

            if line_content.startswith("$$"):
                # 判断是否是完整的单行块公式 $$...$$（长度>4且内容以$$结尾）
                is_complete_single_line = (
                    len(line_content) > 4
                    and line_content.endswith("$$")
                    and line_content[2:-2].strip()  # 中间有实际内容
                )

                if is_complete_single_line:
                    # 单行完整块公式 $$...$$
                    processed.append({"text": line_content, "level": 0, "type": "content", "has_formula": True})
                    content_count += 1
                    logger.debug(f"添加单行块公式段落: {line_content[:50]}...")
                else:
                    # 多行块公式开始（包括单独的$$行或$$xxx形式）
                    formula_lines = [line_content]
                    i += 1

                    # 收集所有行直到遇到$$结束
                    while i < n:
                        next_line = lines[i].strip()
                        formula_lines.append(next_line)

                        if next_line.endswith("$$"):
                            # 块公式结束
                            break
                        i += 1

                    # 合并为一个完整的块公式段落
                    full_formula = "\n".join(formula_lines)

                    processed.append({"text": full_formula, "level": 0, "type": "content", "has_formula": True})
                    content_count += 1
                    logger.debug(f"添加多行块公式段落（{len(formula_lines)}行）: {formula_lines[0][:50]}...")

            elif line_content:  # 普通文本段落（包括列表和引用）
                # 检查是否是引用块
                quote_match = QUOTE_REGEX.match(line_content)
                if quote_match:
                    quote_markers = quote_match.group(1)
                    quote_content = quote_match.group(2).strip()
                    # 计算 > 的数量（忽略空格），支持 >>> 和 > > > 两种格式
                    quote_level = quote_markers.count(">")

                    # 检查引用内容是否包含LaTeX公式
                    has_formula = has_latex_formulas(quote_content)

                    # 超过9级的引用当作普通正文处理（与标题超过9级的行为统一）
                    if quote_level > 9:
                        processed.append(
                            {"text": quote_content, "level": 0, "type": "content", "has_formula": has_formula}
                        )
                        content_count += 1
                        logger.debug(f"超过9级引用（{quote_level}级），当作正文处理: {quote_content[:50]}...")
                    else:
                        # 1-9级正常处理为引用块
                        processed.append(
                            {"text": quote_content, "level": quote_level, "type": "quote", "has_formula": has_formula}
                        )
                        content_count += 1
                        logger.debug(f"添加引用块段落（{quote_level}级）: {quote_content[:50]}...")
                else:
                    # 检查是否是列表项
                    original_line = lines[i]

                    # 无序列表检测
                    unordered_match = UNORDERED_LIST_REGEX.match(original_line)
                    if unordered_match:
                        indent_str = unordered_match.group(1)
                        list_content = unordered_match.group(3).strip()
                        # 使用相对级别计算（自动适应用户的缩进习惯）
                        list_level = _calculate_list_level_relative(indent_str, list_indent_stack)

                        has_formula = has_latex_formulas(list_content)
                        processed.append(
                            {
                                "text": list_content,
                                "level": list_level,
                                "type": "list_item",
                                "list_type": "unordered",
                                "has_formula": has_formula,
                            }
                        )
                        content_count += 1
                        logger.debug(f"添加无序列表项（{list_level}级）: {list_content[:50]}...")
                        i += 1
                        continue

                    # 有序列表检测
                    ordered_match = ORDERED_LIST_REGEX.match(original_line)
                    if ordered_match:
                        indent_str = ordered_match.group(1)
                        list_content = ordered_match.group(3).strip()
                        # 使用相对级别计算（自动适应用户的缩进习惯）
                        list_level = _calculate_list_level_relative(indent_str, list_indent_stack)

                        has_formula = has_latex_formulas(list_content)
                        processed.append(
                            {
                                "text": list_content,
                                "level": list_level,
                                "type": "list_item",
                                "list_type": "ordered",
                                "has_formula": has_formula,
                            }
                        )
                        content_count += 1
                        logger.debug(f"添加有序列表项（{list_level}级）: {list_content[:50]}...")
                        i += 1
                        continue

                    # ==== 列表续行内容检测 ====
                    # 如果当前在列表上下文中，且该行有适当缩进，视为列表续行
                    continuation_match = LIST_CONTINUATION_REGEX.match(original_line)
                    if list_indent_stack and continuation_match:
                        indent_str = continuation_match.group(1)
                        continuation_text = continuation_match.group(2).strip()

                        # 计算缩进级别（与列表项同样的逻辑）
                        indent_len = len(indent_str.replace("\t", "    "))

                        # 找到该缩进对应的列表级别
                        # 续行内容的缩进应该与某个列表项的内容位置对齐
                        # 或者至少比某个列表项的缩进更深
                        continuation_level = None
                        for stack_indent, stack_level in reversed(list_indent_stack):
                            if indent_len >= stack_indent:
                                continuation_level = stack_level
                                break

                        if continuation_level is not None:
                            # 是有效的列表续行
                            has_formula = has_latex_formulas(continuation_text)
                            processed.append(
                                {
                                    "text": continuation_text,
                                    "level": continuation_level,
                                    "type": "list_continuation",
                                    "has_formula": has_formula,
                                }
                            )
                            content_count += 1
                            logger.debug(f"添加列表续行内容（{continuation_level}级）: {continuation_text[:50]}...")
                            i += 1
                            continue

                    # 非列表项且非续行，重置缩进栈（列表被中断）
                    if list_indent_stack:
                        list_indent_stack.clear()
                        logger.debug("非列表内容，重置列表缩进栈")

                    # 普通段落
                    text_content = line_content

                    # 检查是否包含LaTeX公式
                    has_formula = has_latex_formulas(text_content)

                    processed.append(
                        {
                            "text": text_content,
                            "level": 0,
                            "type": "content",
                            "has_formula": has_formula,  # 标记是否包含公式
                        }
                    )
                    content_count += 1
                    # 记录当前行号，用于 Setext 标题检测
                    last_content_line_no = i
                    if has_formula:
                        logger.debug(f"添加正文文本段落（含公式）: {text_content[:50]}...")
                    else:
                        logger.debug(f"添加正文文本段落: {text_content[:50]}...")

        i += 1

    # 统计处理结果
    total_paragraphs = len(processed)
    logger.info(f"处理完成 | 总段落数: {total_paragraphs}")
    logger.info(f"段落统计 | 标题: {heading_count} | 正文: {content_count} | 组合段落: {combined_count}")

    return processed


def process_md_body_with_notes(
    md_body: str,
    remove_numbering: bool = False,
    add_numbering: bool = False,
    formatter: HeadingFormatter | None = None,
) -> tuple[list[dict], dict[str, str], dict[str, str]]:
    """
    处理YAML后的Markdown内容，同时提取脚注和尾注

    此函数先提取脚注和尾注定义，然后处理剩余的Markdown内容。

    参数:
        md_body: YAML头部之后的所有Markdown内容
        remove_numbering: 是否清除Markdown中的原有序号（默认False）
        add_numbering: 是否添加新序号（默认False）
        formatter: HeadingFormatter实例，用于生成序号

    返回:
        Tuple[List[dict], Dict[str, str], Dict[str, str]]:
            - 处理后的段落列表
            - 脚注字典 {md_id: content}
            - 尾注字典 {md_id: content}（ID不含endnote-前缀）

    示例:
        >>> paragraphs, footnotes, endnotes = process_md_body_with_notes(md_body)
        >>> print(f"提取到 {len(footnotes)} 个脚注, {len(endnotes)} 个尾注")
    """
    logger.info("开始处理Markdown内容（含脚注/尾注提取）...")

    # 1. 先提取脚注和尾注定义
    footnotes, endnotes, cleaned_body = extract_notes(md_body)
    logger.info(f"脚注/尾注提取 | 脚注: {len(footnotes)} 个, 尾注: {len(endnotes)} 个")

    # 2. 处理清理后的Markdown内容
    processed = process_md_body(
        cleaned_body, remove_numbering=remove_numbering, add_numbering=add_numbering, formatter=formatter
    )

    return processed, footnotes, endnotes
