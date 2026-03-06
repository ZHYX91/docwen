"""
Markdown链接处理模块

处理Markdown中的各类链接，支持：
1. 嵌入链接（带!前缀）：
   - Wiki嵌入图片：![[image.png]] 或 ![[image.png|alt]]
   - Wiki嵌入MD文件：![[file.md]] → 递归展开内容
   - Markdown嵌入图片：![alt](image.png)
2. 普通链接：
   - Wiki链接：[[link]] 或 [[link|text]]
   - Markdown链接：[text](url)
3. 精确嵌入（Obsidian兼容）：
   - 章节嵌入：![[file.md#标题名]] → 嵌入指定标题的章节
   - 块嵌入：![[file.md#^block-id]] → 嵌入带有块ID的段落

处理模式（可配置）：
- embed：嵌入文件内容/生成占位符
- keep：保留原链接
- extract_text：提取显示文本
- remove：完全移除

特性：
- 路径解析：支持绝对路径、相对路径、仅文件名
- URL 解码：支持 %20、%E4%B8%AD 等编码字符
- 精确嵌入：支持 #标题 和 #^块ID 锚点
- 查询参数处理：自动移除 ?v=1 等查询参数
- Obsidian 兼容：不带扩展名时自动添加 .md
- YAML 移除：自动移除嵌入MD文件的 YAML front matter
- 循环检测：防止嵌入链接无限递归
- 深度限制：限制嵌入嵌套层数

章节嵌入规则：
- ![[file.md#二级标题]] → 提取从 ## 二级标题 到下一个同级或更高级标题的内容
- 包含该标题下的所有子标题内容
- 标题匹配忽略大小写和空格差异

块嵌入规则：
- 块ID格式：段落末尾的 ^block-id 或单独一行的 ^block-id
- ![[file.md#^important]] → 提取带有 ^important 标记的段落
- 会自动移除块ID标记，只返回内容
"""

import base64
import contextlib
import logging
import os
import re
import tempfile
from pathlib import Path
from urllib.parse import unquote

from docwen.config.config_manager import config_manager
from docwen.translation import t

logger = logging.getLogger(__name__)

TABLE_CELL_BR_TOKEN = "{{DOCWEN_BR}}"

_MAX_DATA_URI_IMAGE_BYTES = 10 * 1024 * 1024

_DATA_URI_IMAGE_MIME_TO_EXT = {
    "png": ".png",
    "jpeg": ".jpg",
    "jpg": ".jpg",
    "gif": ".gif",
    "bmp": ".bmp",
    "tiff": ".tiff",
}


# ============================================================
# 工具函数
# ============================================================


def _unescape_pipe(text: str | None) -> str | None:
    """
    还原Markdown表格中转义的竖线

    在Markdown表格中，竖线需要转义为 \\|，此函数将其还原为 |

    参数:
        text: 可能包含转义竖线的文本

    返回:
        还原后的文本，如果输入为None则返回None
    """
    return text.replace(r"\|", "|") if text else text


def _normalize_link_target(link_target: str) -> str:
    """
    规范化链接目标路径（不含锚点部分）

    处理步骤：
    1. URL 解码：%20 → 空格，%E4%B8%AD → 中
    2. 移除锚点：file.md#section → file.md
    3. 移除查询参数：file.png?v=1 → file.png

    参数:
        link_target: 原始链接目标

    返回:
        str: 规范化后的路径（不含锚点）
    """
    # 1. URL 解码
    result = unquote(link_target)

    # 2. 移除锚点 (#) - 取 # 前面的部分
    if "#" in result:
        result = result.split("#")[0]

    # 3. 移除查询参数 (?) - 取 ? 前面的部分
    if "?" in result:
        result = result.split("?")[0]

    return result.strip()


def _is_data_uri_image(link_target: str) -> bool:
    """
    判断链接目标是否为 base64 编码的 data URI 图片。

    仅支持形如：
    - data:image/<subtype>;base64,<payload>

    参数:
        link_target: 原始链接目标字符串

    返回:
        bool: 是否为 data:image 的 base64 data URI
    """
    return link_target.startswith("data:image/") and ";base64," in link_target


def _estimate_base64_decoded_size(payload: str) -> int:
    """
    估算 base64 payload 解码后的字节大小。

    用于在实际解码前快速判断是否超出大小上限，降低内存峰值风险。
    该值为估算值：会根据末尾 '=' padding 做简单修正。

    参数:
        payload: base64 字符串（不含 data URI header）

    返回:
        int: 估算的解码后字节数
    """
    payload_len = len(payload)
    if payload_len == 0:
        return 0
    padding = 2 if payload.endswith("==") else 1 if payload.endswith("=") else 0
    return max(0, (payload_len * 3) // 4 - padding)


def _resolve_data_uri_image_to_temp_file(data_uri: str, *, temp_dir: str | None) -> str | None:
    """
    将 base64 data URI 图片解码为临时文件，并返回文件路径。

    支持：
    - data:image/<subtype>[;param...];base64,<payload>

    参数:
        data_uri: 完整 data URI 字符串
        temp_dir: 临时目录（为 None 时使用系统临时目录）

    返回:
        str | None: 临时文件路径；解析/解码/写入失败返回 None
    """
    try:
        header, payload = data_uri.split(",", 1)
    except ValueError:
        return None

    if not header.startswith("data:image/") or ";base64" not in header:
        return None

    mime_subtype = header[len("data:image/") :].split(";", 1)[0].strip().lower()
    ext = _DATA_URI_IMAGE_MIME_TO_EXT.get(mime_subtype)
    if not ext:
        cleaned = re.sub(r"[^a-z0-9]+", "", mime_subtype)
        ext = f".{cleaned}" if cleaned else ".img"

    payload = payload.strip()
    payload = payload.translate({ord(c): None for c in " \r\n\t"})
    if _estimate_base64_decoded_size(payload) > _MAX_DATA_URI_IMAGE_BYTES:
        logger.warning("data URI 图片超过大小上限，已跳过（上限: %.2f MB）", _MAX_DATA_URI_IMAGE_BYTES / (1024 * 1024))
        return None

    try:
        image_bytes = base64.b64decode(payload, validate=True)
    except Exception as e:
        logger.warning("data URI 图片 base64 解码失败: %s", e)
        return None

    if len(image_bytes) > _MAX_DATA_URI_IMAGE_BYTES:
        logger.warning("data URI 图片超过大小上限，已跳过（上限: %.2f MB）", _MAX_DATA_URI_IMAGE_BYTES / (1024 * 1024))
        return None

    fd = None
    temp_path = None
    try:
        fd, temp_path = tempfile.mkstemp(suffix=ext, prefix="docwen_data_uri_", dir=temp_dir)
        with os.fdopen(fd, "wb") as f:
            fd = None
            f.write(image_bytes)
        logger.info("data URI 图片已解码为临时文件: %s (%.2f KB)", Path(temp_path).name, len(image_bytes) / 1024)
        return temp_path
    except Exception as e:
        logger.error("data URI 图片写入临时文件失败: %s", e)
        if fd is not None:
            with contextlib.suppress(Exception):
                os.close(fd)
        if temp_path:
            with contextlib.suppress(Exception):
                Path(temp_path).unlink(missing_ok=True)
        return None


def _parse_anchor(link_target: str) -> tuple:
    """
    解析链接目标，分离文件路径、标题锚点、块ID

    支持的格式：
    - file.md → (file.md, None, None)
    - file.md#标题名 → (file.md, "标题名", None)
    - file.md#^block-id → (file.md, None, "block-id")

    参数:
        link_target: 原始链接目标

    返回:
        tuple: (file_path, heading, block_id)
    """
    # URL 解码
    decoded = unquote(link_target)

    # 移除查询参数
    if "?" in decoded:
        decoded = decoded.split("?")[0]

    # 检查是否有锚点
    if "#" not in decoded:
        return (decoded.strip(), None, None)

    # 分离文件路径和锚点
    parts = decoded.split("#", 1)
    file_path = parts[0].strip()
    anchor = parts[1].strip() if len(parts) > 1 else None

    if not anchor:
        return (file_path, None, None)

    # 判断是块ID还是标题
    if anchor.startswith("^"):
        # 块ID格式：^block-id
        block_id = anchor[1:]  # 移除 ^ 前缀
        return (file_path, None, block_id)
    else:
        # 标题格式
        return (file_path, anchor, None)


def _extract_section_by_heading(content: str, heading: str) -> str | None:
    """
    从内容中提取指定标题的章节

    提取从指定标题到下一个同级或更高级标题之间的所有内容。

    参数:
        content: Markdown 文件内容
        heading: 要提取的标题文本（不含 # 前缀）

    返回:
        Optional[str]: 提取的章节内容，未找到返回 None
    """
    lines = content.split("\n")

    # 标准化目标标题（去除首尾空格，用于匹配）
    target_heading = heading.strip()

    # 查找目标标题
    start_index = None
    start_level: int | None = None

    for i, line in enumerate(lines):
        # 检查是否是标题行
        match = re.match(r"^(#{1,9})\s+(.*)$", line)
        if match:
            level = len(match.group(1))
            title = match.group(2).strip()

            # 匹配标题（忽略大小写和空格差异）
            if title.lower().replace(" ", "") == target_heading.lower().replace(" ", ""):
                start_index = i
                start_level = level
                logger.debug(f"找到目标标题: '{title}' (级别 {level}) 在行 {i}")
                break

    if start_index is None:
        logger.warning(f"未找到标题: '{heading}'")
        return None
    if start_level is None:
        return None

    # 查找章节结束位置（下一个同级或更高级标题）
    end_index = len(lines)

    for i in range(start_index + 1, len(lines)):
        line = lines[i]
        match = re.match(r"^(#{1,9})\s+", line)
        if match:
            level = len(match.group(1))
            if level <= start_level:
                # 遇到同级或更高级标题，结束
                end_index = i
                logger.debug(f"章节结束于行 {i} (遇到级别 {level} 标题)")
                break

    # 提取章节内容（包含标题行）
    section_lines = lines[start_index:end_index]

    # 移除末尾空行
    while section_lines and not section_lines[-1].strip():
        section_lines.pop()

    result = "\n".join(section_lines)
    logger.debug(f"提取章节内容: {len(section_lines)} 行")
    return result


def _extract_block_by_id(content: str, block_id: str) -> str | None:
    """
    从内容中提取带有指定块ID的段落

    块ID格式：
    - 行内：这是一段文字 ^block-id
    - 单独行：^block-id（提取其前一个非空段落）

    参数:
        content: Markdown 文件内容
        block_id: 块ID（不含 ^ 前缀）

    返回:
        Optional[str]: 提取的段落内容，未找到返回 None
    """
    lines = content.split("\n")

    # 构建块ID正则：匹配 ^block-id（可能在行尾或单独一行）
    # 注意：块ID后面可能有空格
    block_pattern = re.compile(r"\^" + re.escape(block_id) + r"\s*$")

    for i, line in enumerate(lines):
        if block_pattern.search(line):
            # 找到包含块ID的行
            line_stripped = line.strip()

            if line_stripped == f"^{block_id}":
                # 单独一行的块ID，提取前一个非空段落
                logger.debug(f"找到单独块ID行: {i}")

                # 向上查找前一个非空段落
                paragraph_lines = []
                j = i - 1

                # 跳过空行
                while j >= 0 and not lines[j].strip():
                    j -= 1

                # 收集段落（直到遇到空行或文件开头）
                while j >= 0 and lines[j].strip():
                    paragraph_lines.insert(0, lines[j])
                    j -= 1

                if paragraph_lines:
                    result = "\n".join(paragraph_lines)
                    logger.debug(f"提取块（单独行模式）: {len(paragraph_lines)} 行")
                    return result
            else:
                # 块ID在行内，移除块ID标记并返回该行/段落
                # 移除块ID标记
                clean_line = block_pattern.sub("", line).rstrip()

                # 如果这行是段落的一部分，尝试提取完整段落
                paragraph_lines = [clean_line]

                # 向上扩展段落（如果不是空行）
                j = i - 1
                while j >= 0 and lines[j].strip() and not lines[j].strip().startswith("#"):
                    paragraph_lines.insert(0, lines[j])
                    j -= 1

                # 向下扩展段落（直到空行或标题）
                j = i + 1
                while j < len(lines) and lines[j].strip() and not lines[j].strip().startswith("#"):
                    # 检查是否有另一个块ID
                    if re.search(r"\^\w+", lines[j]):
                        break
                    paragraph_lines.append(lines[j])
                    j += 1

                result = "\n".join(paragraph_lines)
                logger.debug(f"提取块（行内模式）: {len(paragraph_lines)} 行")
                return result

    logger.warning(f"未找到块ID: ^{block_id}")
    return None


def _strip_yaml_front_matter(content: str) -> str:
    """
    移除 Markdown 文件开头的 YAML front matter

    YAML front matter 格式：
    ---
    title: xxx
    date: xxx
    ---

    参数:
        content: 原始文件内容

    返回:
        str: 移除 YAML 后的内容
    """
    if not content.startswith("---"):
        return content

    # 查找第二个 ---（从位置3开始搜索，跳过开头的 ---）
    end_pos = content.find("---", 3)
    if end_pos == -1:
        return content

    # 跳过 YAML 部分和紧随的换行符
    result = content[end_pos + 3 :].lstrip("\r\n")
    logger.debug("已移除 YAML front matter")
    return result


# 支持的文件扩展名
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg", ".webp", ".ico"}
MD_EXTENSIONS = {".md", ".markdown"}

# 链接识别正则表达式
# 支持 | 和 \| 作为分隔符（\| 用于Markdown表格中）
WIKI_EMBED_PATTERN = r"!\[\[((?:[^|\]\\]|\\(?![|]).)*)(?:(?:\\)?\|((?:[^\]\\]|\\(?![|]).)*)?)?\]\]"  # ![[target]] 或 ![[target|text]] 或 ![[target\|text]]
WIKI_EMBED_SIZE_PATTERN = r"!\[\[([^|\]]+)\|(\d+)(?:x(\d+))?\]\]"  # ![[target|300]] 或 ![[target|300x200]]
MD_EMBED_IMAGE_PATTERN = (
    r"!\[([^\]]*)\]\(([^)\s]+)(?:\s*=\s*(\d*)x(\d*))?\s*\)"  # ![alt](url) 或 ![alt](url =300x200) 或 ![alt](url =300x)
)


def _format_image_placeholder(image_path: str, width: int | None = None, height: int | None = None) -> str:
    if width is None and height is None:
        return f"{{{{IMAGE:{image_path}}}}}"
    w = "" if width is None else str(width)
    h = "" if height is None else str(height)
    return f"{{{{IMAGE:{image_path}\\|{w}\\|{h}}}}}"


def _split_alt_text_and_size(alt_text: str | None) -> tuple[str | None, int | None, int | None]:
    if not alt_text:
        return None, None, None
    if "|" not in alt_text:
        return alt_text, None, None
    left, right = alt_text.rsplit("|", 1)
    right = right.strip()
    if not right:
        return alt_text, None, None
    if "x" in right:
        w_str, h_str = right.split("x", 1)
        if w_str.isdigit() and (h_str.isdigit() or h_str == ""):
            width = int(w_str)
            height = int(h_str) if h_str.isdigit() else None
            return left, width, height
        return alt_text, None, None
    if right.isdigit():
        return left, int(right), None
    return alt_text, None, None


def _is_probable_table_row_line(line: str) -> bool:
    from docwen.utils.markdown_utils import is_table_separator

    cleaned = re.sub(r"^[> \t]+", "", line).strip()
    if not cleaned.startswith("|") or not cleaned.endswith("|"):
        return False
    if cleaned.count("|") < 2:
        return False
    return not is_table_separator(cleaned)


def _is_table_context(text: str, start_index: int) -> bool:
    line_start = text.rfind("\n", 0, start_index) + 1
    line_end = text.find("\n", start_index)
    if line_end == -1:
        line_end = len(text)
    return _is_probable_table_row_line(text[line_start:line_end])


def _flatten_for_table_cell(text: str) -> str:
    if not text:
        return text
    flattened = re.sub(r"\r?\n+", TABLE_CELL_BR_TOKEN, text.strip())
    flattened = re.sub(r"(?<!\\)\|", r"\\|", flattened)
    return flattened


def process_markdown_links(
    text: str,
    source_file_path: str,
    visited_files: set[str] | None = None,
    depth: int = 0,
    table_safe: bool = False,
    temp_dir: str | None = None,
) -> str:
    """
    处理Markdown中的所有链接（嵌入和非嵌入）

    处理流程：
    1. 检查深度限制
    2. 初始化visited集合（防止循环引用）
    3. 处理嵌入链接（带!前缀）：
       - ![[image.png]] → {{IMAGE:path}}占位符
       - ![[file.md]] → 读取并递归处理
    4. 处理普通链接（不带!前缀）：
       - [[link]] 或 [[link|text]] → 根据配置处理
       - [text](url) → 根据配置处理
    5. 返回处理后的文本

    参数:
        text: 要处理的文本
        source_file_path: 源MD文件的绝对路径
        visited_files: 已访问文件集合（用于循环检测）
        depth: 当前递归深度

    返回:
        str: 处理后的文本（所有链接已处理）
    """
    # 规范化源文件路径
    source_file_path = str(Path(source_file_path).resolve())

    logger.info("展开嵌入内容 | 源文件: %s | 深度: %d", Path(source_file_path).name, depth)

    # 检查深度限制
    max_depth = config_manager.get_max_embed_depth()
    if depth >= max_depth:
        logger.warning("达到最大嵌入深度 %d，停止展开", max_depth)
        return _handle_max_depth_error()

    # 初始化visited集合
    if visited_files is None:
        visited_files = set()

    # 添加当前文件到visited（使用规范化路径）
    visited_files.add(source_file_path)

    if not text:
        return text

    result = text
    embed_count = {"images": 0, "markdown": 0}

    # 0. 处理带尺寸的 Wiki 嵌入（优先于通用模式）
    wiki_size_embeds = list(re.finditer(WIKI_EMBED_SIZE_PATTERN, result))
    for match in wiki_size_embeds:
        original_link = match.group(0)
        link_target = match.group(1).strip()
        start_index = result.find(original_link)
        in_table = table_safe and start_index != -1 and _is_table_context(result, start_index)
        width = int(match.group(2)) if match.group(2) else None
        height = int(match.group(3)) if match.group(3) else None
        logger.debug("识别到Wiki嵌入链接(带尺寸): %s", original_link)
        replacement = _process_single_embed(
            link_target,
            original_link,
            source_file_path,
            visited_files,
            depth,
            embed_count,
            link_type="wiki",
            display_text=None,
            width=width,
            height=height,
            table_safe=in_table,
            temp_dir=temp_dir,
        )
        if replacement is not None:
            result = result.replace(original_link, replacement, 1)

    # 1. 处理Wiki嵌入链接：![[target]] 或 ![[target|display]]
    wiki_embeds = list(re.finditer(WIKI_EMBED_PATTERN, result))
    for match in wiki_embeds:
        original_link = match.group(0)  # 完整的 ![[target]]
        link_target = _unescape_pipe(match.group(1).strip()) or ""  # target，还原转义的竖线
        # 提取显示文本（group(2)是 | 后面的部分）
        display_text = _unescape_pipe(match.group(2).strip()) if match.group(2) else None

        start_index = result.find(original_link)
        in_table = table_safe and start_index != -1 and _is_table_context(result, start_index)

        width = None
        height = None
        if display_text:
            dt = display_text.strip()
            if dt.isdigit():
                width = int(dt)
                display_text = None
            elif "x" in dt:
                w_str, h_str = dt.split("x", 1)
                if w_str.isdigit() and (h_str.isdigit() or h_str == ""):
                    width = int(w_str)
                    height = int(h_str) if h_str.isdigit() else None
                    display_text = None

        logger.debug("识别到Wiki嵌入链接: %s", original_link)

        # 处理嵌入（Wiki类型）
        replacement = _process_single_embed(
            link_target,
            original_link,
            source_file_path,
            visited_files,
            depth,
            embed_count,
            link_type="wiki",
            display_text=display_text,
            width=width,
            height=height,
            table_safe=in_table,
            temp_dir=temp_dir,
        )

        if replacement is not None:
            result = result.replace(original_link, replacement, 1)

    # 2. 处理Markdown嵌入图片：![alt](url)
    md_embeds = list(re.finditer(MD_EMBED_IMAGE_PATTERN, result))
    for match in md_embeds:
        original_link = match.group(0)  # 完整的 ![alt](url)
        alt_text = match.group(1)  # alt（作为显示文本）
        link_target = match.group(2).strip()  # url
        md_width = int(match.group(3)) if match.group(3) and match.group(3).isdigit() else None
        md_height = int(match.group(4)) if match.group(4) and match.group(4).isdigit() else None
        display_text, alt_width, alt_height = _split_alt_text_and_size(alt_text)
        width = md_width if (md_width is not None or md_height is not None) else alt_width
        height = md_height if (md_width is not None or md_height is not None) else alt_height

        logger.debug("识别到Markdown嵌入图片: %s", original_link)

        # 处理嵌入（Markdown类型，alt_text作为显示文本）
        replacement = _process_single_embed(
            link_target,
            original_link,
            source_file_path,
            visited_files,
            depth,
            embed_count,
            link_type="markdown",
            display_text=display_text,
            width=width,
            height=height,
            temp_dir=temp_dir,
        )

        if replacement is not None:
            result = result.replace(original_link, replacement, 1)

    logger.info("嵌入处理完成 | 图片: %d个 | MD文件: %d个", embed_count["images"], embed_count["markdown"])

    # 处理普通链接（非嵌入，根据配置）
    result = _process_non_embed_links(result)
    logger.debug("普通链接处理完成")

    return result


def _process_non_embed_links(text: str) -> str:
    """
    处理普通链接（非嵌入）

    根据配置处理Wiki链接和Markdown链接：
    - Wiki: [[link]] 或 [[link|text]]
    - Markdown: [text](url)

    参数:
        text: 要处理的文本

    返回:
        str: 处理后的文本
    """
    logger.debug(f"_process_non_embed_links 输入: {text[:200] if text else 'None'}")

    if not text:
        logger.debug("文本为空，直接返回")
        return text

    result = text

    # 获取非嵌入链接处理配置
    wiki_mode = config_manager.get_wiki_link_mode()
    markdown_mode = config_manager.get_markdown_link_mode()

    logger.debug("处理普通链接 | Wiki模式: %s | Markdown模式: %s", wiki_mode, markdown_mode)

    # 1. 处理Wiki链接：[[link]] 或 [[link|text]]
    # 使用 (?<!!) 排除嵌入链接，支持 | 和 \| 作为分隔符
    wiki_pattern = r"(?<!!)\[\[(?:((?:[^|\]\\]|\\[|\]])+)(?:\\)?\|)?((?:[^\]\\]|\\[|\]])+)\]\]"

    # 查找所有匹配
    wiki_matches = list(re.finditer(wiki_pattern, text))
    logger.debug(f"找到 {len(wiki_matches)} 个Wiki链接")
    for match in wiki_matches:
        logger.debug(f"  Wiki链接: {match.group(0)}")

    if wiki_mode == "keep":
        # 保留原样
        logger.debug("Wiki链接模式: keep（保留原样）")
    elif wiki_mode == "extract_text":
        # 提取显示文本（group(2)是显示文本），还原转义的竖线
        logger.debug("Wiki链接模式: extract_text（提取文本）")
        result = re.sub(wiki_pattern, lambda m: _unescape_pipe(m.group(2)) or "", result)
        logger.debug(f"Wiki链接处理后: {result[:200]}")
    elif wiki_mode == "remove":
        # 完全移除
        logger.debug("Wiki链接模式: remove（移除）")
        result = re.sub(wiki_pattern, "", result)

    # 2. 处理Markdown链接：[text](url)
    # 使用 (?<!!) 排除嵌入链接（![alt](url)）
    markdown_pattern = r"(?<!!)\[([^\]]+)\]\([^)]+\)"

    # 查找所有匹配
    md_matches = list(re.finditer(markdown_pattern, result))
    logger.debug(f"找到 {len(md_matches)} 个Markdown链接")
    for match in md_matches:
        logger.debug(f"  Markdown链接: {match.group(0)}")

    if markdown_mode == "keep":
        # 保留原样
        logger.debug("Markdown链接模式: keep（保留原样）")
    elif markdown_mode == "extract_text":
        # 提取显示文本（group(1)是显示文本）
        logger.debug("Markdown链接模式: extract_text（提取文本）")
        result = re.sub(markdown_pattern, r"\1", result)
        logger.debug(f"Markdown链接处理后: {result[:200]}")
    elif markdown_mode == "remove":
        # 完全移除
        logger.debug("Markdown链接模式: remove（移除）")
        result = re.sub(markdown_pattern, "", result)

    logger.debug(f"_process_non_embed_links 输出: {result[:200] if result else 'None'}")
    return result


def _process_single_embed(
    link_target: str,
    original_link: str,
    source_file_path: str,
    visited_files: set[str],
    depth: int,
    embed_count: dict,
    link_type: str = "wiki",
    display_text: str | None = None,
    width: int | None = None,
    height: int | None = None,
    table_safe: bool = False,
    temp_dir: str | None = None,
) -> str | None:
    """
    处理单个嵌入链接

    支持精确嵌入：
    - ![[file.md#标题名]] → 嵌入指定章节
    - ![[file.md#^block-id]] → 嵌入指定块

    参数:
        link_target: 链接目标（文件名或路径，可能含锚点）
        original_link: 原始链接文本
        source_file_path: 源文件路径
        visited_files: 已访问文件集合
        depth: 当前深度
        embed_count: 统计计数器
        link_type: 链接类型（'wiki' 或 'markdown'）
        display_text: 显示文本（用于extract_text模式，优先于文件名）

    返回:
        Optional[str]: 替换文本，None表示不替换
    """
    if _is_data_uri_image(link_target):
        temp_path = _resolve_data_uri_image_to_temp_file(link_target, temp_dir=temp_dir)
        if temp_path:
            embed_count["images"] += 1
            return process_embedded_image(temp_path, original_link, link_type, display_text, width=width, height=height)
        if link_type == "wiki":
            mode = config_manager.get_wiki_embed_image_mode()
        else:
            mode = config_manager.get_markdown_embed_image_mode()
        if mode == "keep":
            return original_link
        if mode == "extract_text":
            return display_text or ""
        if mode == "remove":
            return ""
        return display_text or ""

    # 解析锚点（提取文件路径、标题、块ID）
    file_path_part, heading, block_id = _parse_anchor(link_target)

    if heading:
        logger.debug("检测到章节嵌入: %s#%s", file_path_part, heading)
    elif block_id:
        logger.debug("检测到块嵌入: %s#^%s", file_path_part, block_id)

    # 解析文件路径（使用不含锚点的路径）
    resolved_path = resolve_file_path(file_path_part, source_file_path)

    if resolved_path is None:
        # 文件未找到
        logger.warning("文件未找到: %s（来自: %s）", link_target, Path(source_file_path).name)
        return _handle_not_found_error(link_target)

    logger.debug("路径解析成功: %s → %s", file_path_part, resolved_path)

    # 判断文件类型
    file_type = get_file_type(resolved_path)

    if file_type == "image":
        embed_count["images"] += 1
        return process_embedded_image(resolved_path, original_link, link_type, display_text, width=width, height=height)

    elif file_type == "markdown":
        embed_count["markdown"] += 1
        return process_embedded_md_file(
            resolved_path,
            source_file_path,
            visited_files,
            depth,
            display_text,
            heading=heading,
            block_id=block_id,
            table_safe=table_safe,
        )

    else:
        # 未知文件类型
        logger.warning("未知文件类型: %s", resolved_path)
        return None


def resolve_file_path(link_target: str, source_file_path: str) -> str | None:
    r"""
    解析嵌入文件的路径

    解析优先级：
    1. 绝对路径 → 直接使用
    2. 相对路径（含/或\）→ 相对于源文件目录
    3. 仅文件名 → 在search_dirs中按顺序搜索

    特性：
    - URL 解码：支持 %20、%E4%B8%AD 等编码字符
    - 移除锚点：file.md#section → file.md
    - 移除查询参数：file.png?v=1 → file.png
    - 自动添加扩展名：文档 → 文档.md（Obsidian 兼容）

    参数:
        link_target: 链接目标（从![[这里]]或![]()中提取）
        source_file_path: 源MD文件的绝对路径

    返回:
        Optional[str]: 解析后的绝对路径，或None（未找到文件）
    """
    logger.debug("解析文件路径: %s", link_target)

    # 规范化链接目标（URL解码、移除锚点和查询参数）
    link_target = _normalize_link_target(link_target)
    logger.debug("规范化后: %s", link_target)

    # 获取源文件所在目录
    source_dir = Path(source_file_path).parent

    # 辅助函数：尝试查找文件（支持不带扩展名的 Obsidian 格式）
    def try_find_file(path: str) -> str | None:
        """尝试查找文件，如果不存在且无扩展名则尝试添加 .md"""
        candidate = Path(path)
        if candidate.exists():
            return str(candidate)
        # 无扩展名时，尝试添加 .md（Obsidian 兼容）
        if not candidate.suffix:
            md_path = candidate.with_suffix(".md")
            if md_path.exists():
                logger.debug("自动添加 .md 扩展名: %s → %s", candidate, md_path)
                return str(md_path)
        return None

    # 1. 绝对路径：直接使用
    if Path(link_target).is_absolute():
        normalized = os.path.normpath(link_target)
        result = try_find_file(normalized)
        if result:
            logger.debug("解析为绝对路径: %s", result)
            return result
        else:
            logger.debug("绝对路径文件不存在: %s", normalized)
            return None

    # 2. 相对路径（包含路径分隔符）
    if "/" in link_target or "\\" in link_target:
        full_path = os.path.normpath(str(source_dir / link_target))
        result = try_find_file(full_path)
        if result:
            logger.debug("解析为相对路径: %s", result)
            return result
        else:
            logger.debug("相对路径文件不存在: %s", full_path)
            return None

    # 3. 仅文件名：先搜索同名文件夹，再在search_dirs中搜索

    # 3.1 搜索同名文件夹（如 B.md 对应 B/ 文件夹）
    # 获取源文件名（不含扩展名）
    source_basename = Path(source_file_path).stem
    same_name_folder = source_dir / source_basename

    if same_name_folder.is_dir():
        search_path = os.path.normpath(str(same_name_folder / link_target))
        result = try_find_file(search_path)
        if result:
            logger.debug("在同名文件夹 '%s' 中找到文件: %s", source_basename, result)
            return result

    # 3.2 在配置的search_dirs中搜索
    search_dirs = config_manager.get_search_dirs()
    logger.debug("搜索文件: %s（在目录: %s）", link_target, search_dirs)

    for search_dir in search_dirs:
        search_path = os.path.normpath(str(source_dir / search_dir / link_target))
        result = try_find_file(search_path)
        if result:
            logger.debug("在搜索目录 '%s' 中找到文件: %s", search_dir, result)
            return result

    logger.debug("在所有搜索目录中都未找到文件: %s", link_target)
    return None


def get_file_type(file_path: str) -> str:
    """
    判断文件类型

    参数:
        file_path: 文件路径

    返回:
        str: 文件类型 ('image', 'markdown', 'unknown')
    """
    ext = Path(file_path).suffix.lower()

    if ext in IMAGE_EXTENSIONS:
        logger.debug("文件类型: 图片 (%s)", ext)
        return "image"
    elif ext in MD_EXTENSIONS:
        logger.debug("文件类型: Markdown (%s)", ext)
        return "markdown"
    else:
        logger.debug("文件类型: 未知 (%s)", ext)
        return "unknown"


def process_embedded_image(
    image_path: str,
    original_link: str,
    link_type: str = "wiki",
    display_text: str | None = None,
    *,
    width: int | None = None,
    height: int | None = None,
) -> str:
    """
    处理嵌入图片

    根据配置决定如何处理嵌入的图片链接。

    参数:
        image_path: 图片文件的绝对路径
        original_link: 原始链接文本（用于keep模式）
        link_type: 链接类型（'wiki' 或 'markdown'）
        display_text: 显示文本（用于extract_text模式，优先于文件名）

    返回:
        str: 处理后的文本
    """
    # 根据链接类型获取对应的处理模式
    if link_type == "wiki":
        mode = config_manager.get_wiki_embed_image_mode()
    else:
        mode = config_manager.get_markdown_embed_image_mode()

    logger.debug("处理嵌入图片: %s | 类型: %s | 模式: %s", Path(image_path).name, link_type, mode)

    if mode == "embed":
        # 生成图片占位符（后续会被docx_processor替换为实际图片）
        placeholder = _format_image_placeholder(image_path, width=width, height=height)
        logger.info("生成图片占位符: %s", placeholder)
        return placeholder

    elif mode == "keep":
        # 保留原链接
        logger.debug("保留原图片链接: %s", original_link)
        return original_link

    elif mode == "extract_text":
        # 优先返回显示文本，没有则返回文件名
        if display_text:
            logger.debug("提取图片显示文本: %s", display_text)
            return display_text
        else:
            filename = Path(image_path).name
            logger.debug("提取图片文件名: %s", filename)
            return filename

    elif mode == "remove":
        # 完全移除
        logger.debug("移除图片链接")
        return ""

    else:
        # 未知模式，使用默认（embed）
        logger.warning("未知的图片嵌入模式: %s，使用默认embed", mode)
        return _format_image_placeholder(image_path, width=width, height=height)


def process_embedded_md_file(
    md_path: str,
    source_file_path: str,
    visited_files: set[str],
    depth: int,
    display_text: str | None = None,
    heading: str | None = None,
    block_id: str | None = None,
    table_safe: bool = False,
) -> str:
    """
    处理嵌入MD文件（递归）

    根据配置决定如何处理嵌入的MD文件。
    支持精确嵌入：
    - heading: 嵌入指定标题的章节
    - block_id: 嵌入指定块ID的段落

    参数:
        md_path: MD文件的绝对路径
        source_file_path: 当前源文件路径（用于日志）
        visited_files: 已访问文件集合
        depth: 当前递归深度
        display_text: 显示文本（用于extract_text模式，优先于文件名）
        heading: 要嵌入的章节标题（可选）
        block_id: 要嵌入的块ID（可选）

    返回:
        str: 处理后的内容
    """
    mode = config_manager.get_embed_md_file_mode()

    # 构建日志描述
    embed_desc = Path(md_path).name
    if heading:
        embed_desc += f"#{heading}"
    elif block_id:
        embed_desc += f"#^{block_id}"

    logger.debug("处理嵌入MD文件: %s | 模式: %s", embed_desc, mode)

    # 规范化路径用于循环检测
    normalized_path = str(Path(md_path).resolve())

    # 检查循环引用
    if config_manager.is_circular_detection_enabled() and normalized_path in visited_files:
        logger.error("检测到循环引用: %s（被 %s 嵌入）", Path(md_path).name, Path(source_file_path).name)
        return _handle_circular_error(Path(md_path).name)

    if mode == "embed":
        # 读取并递归处理
        try:
            logger.info("读取MD文件: %s（深度: %d）", embed_desc, depth + 1)

            with Path(md_path).open(encoding="utf-8") as f:
                content = f.read()

            logger.debug("读取到 %d 字符", len(content))

            # 硬编码：默认移除 YAML front matter
            # 如需保留 YAML，手动改为 False
            strip_yaml = True
            if strip_yaml:
                content = _strip_yaml_front_matter(content)

            # 精确提取：根据 heading 或 block_id 提取特定内容
            if heading:
                extracted = _extract_section_by_heading(content, heading)
                if extracted is None:
                    logger.warning("章节未找到: %s#%s", Path(md_path).name, heading)
                    return _handle_section_not_found_error(Path(md_path).name, heading)
                content = extracted
                logger.info("已提取章节: %s#%s（%d 字符）", Path(md_path).name, heading, len(content))
            elif block_id:
                extracted = _extract_block_by_id(content, block_id)
                if extracted is None:
                    logger.warning("块未找到: %s#^%s", Path(md_path).name, block_id)
                    return _handle_block_not_found_error(Path(md_path).name, block_id)
                content = extracted
                logger.info("已提取块: %s#^%s（%d 字符）", Path(md_path).name, block_id, len(content))

            # 将当前文件添加到visited（递归前）
            visited_files.add(normalized_path)

            try:
                # 递归展开内容（提取的内容中可能还有嵌入链接）
                expanded = process_markdown_links(
                    content,
                    md_path,
                    visited_files,
                    depth + 1,
                    table_safe=table_safe,
                )
            finally:
                # 递归返回后移除（允许其他路径再次访问）
                visited_files.discard(normalized_path)

            logger.info("MD文件展开完成: %s", embed_desc)
            return _flatten_for_table_cell(expanded) if table_safe else expanded

        except Exception as e:
            logger.error("读取MD文件失败: %s | 错误: %s", md_path, str(e))
            return _handle_not_found_error(Path(md_path).name)

    elif mode == "keep":
        # 保留原链接（包含锚点）
        link_text = Path(md_path).name
        if heading:
            link_text += f"#{heading}"
        elif block_id:
            link_text += f"#^{block_id}"
        original = f"![[{link_text}]]"
        logger.debug("保留原MD链接: %s", original)
        return original

    elif mode == "extract_text":
        # 优先返回显示文本，没有则返回文件名（不含扩展名）
        if display_text:
            logger.debug("提取MD显示文本: %s", display_text)
            return display_text
        else:
            # 返回文件名（不含扩展名）
            filename = Path(md_path).stem
            logger.debug("提取MD文件名: %s", filename)
            return filename

    elif mode == "remove":
        # 完全移除
        logger.debug("移除MD链接")
        return ""

    else:
        # 未知模式，使用默认（embed）
        logger.warning("未知的MD嵌入模式: %s，使用默认embed", mode)
        try:
            with Path(md_path).open(encoding="utf-8") as f:
                content = f.read()
            expanded = process_markdown_links(content, md_path, visited_files, depth + 1, table_safe=table_safe)
            return _flatten_for_table_cell(expanded) if table_safe else expanded
        except Exception as e:
            logger.error("读取MD文件失败: %s", str(e))
            return _handle_not_found_error(Path(md_path).name)


def _handle_section_not_found_error(filename: str, heading: str) -> str:
    """
    处理章节未找到错误

    参数:
        filename: 文件名
        heading: 未找到的章节标题

    返回:
        str: 根据配置返回的文本
    """
    mode = config_manager.get_file_not_found_mode()

    if mode == "ignore":
        logger.debug("章节未找到，静默忽略: %s#%s", filename, heading)
        return ""

    elif mode == "keep":
        original = f"![[{filename}#{heading}]]"
        logger.debug("章节未找到，保留原链接: %s", original)
        return original

    elif mode == "placeholder":
        placeholder = t("link_processing.section_not_found", filename=filename, heading=heading)
        logger.debug("章节未找到，插入占位符: %s", placeholder)
        return placeholder

    else:
        return t("link_processing.section_not_found", filename=filename, heading=heading)


def _handle_block_not_found_error(filename: str, block_id: str) -> str:
    """
    处理块未找到错误

    参数:
        filename: 文件名
        block_id: 未找到的块ID

    返回:
        str: 根据配置返回的文本
    """
    mode = config_manager.get_file_not_found_mode()

    if mode == "ignore":
        logger.debug("块未找到，静默忽略: %s#^%s", filename, block_id)
        return ""

    elif mode == "keep":
        original = f"![[{filename}#^{block_id}]]"
        logger.debug("块未找到，保留原链接: %s", original)
        return original

    elif mode == "placeholder":
        placeholder = t("link_processing.block_not_found", filename=filename, block_id=block_id)
        logger.debug("块未找到，插入占位符: %s", placeholder)
        return placeholder

    else:
        return t("link_processing.block_not_found", filename=filename, block_id=block_id)


def _handle_not_found_error(filename: str) -> str:
    """
    处理文件未找到错误

    参数:
        filename: 未找到的文件名

    返回:
        str: 根据配置返回的文本
    """
    mode = config_manager.get_file_not_found_mode()

    if mode == "ignore":
        # 静默忽略（删除链接）
        logger.debug("文件未找到，静默忽略: %s", filename)
        return ""

    elif mode == "keep":
        # 保留原链接
        original = f"![[{filename}]]"
        logger.debug("文件未找到，保留原链接: %s", original)
        return original

    elif mode == "placeholder":
        # 插入占位符（使用国际化）
        placeholder = t("link_processing.file_not_found", filename=filename)
        logger.debug("文件未找到，插入占位符: %s", placeholder)
        return placeholder

    else:
        # 默认：占位符
        return t("link_processing.file_not_found", filename=filename)


def _handle_circular_error(filename: str) -> str:
    """
    处理循环引用错误

    参数:
        filename: 导致循环的文件名

    返回:
        str: 根据配置返回的文本
    """
    mode = config_manager.get_circular_reference_mode()

    if mode == "ignore":
        # 静默忽略（停止嵌入）
        logger.debug("循环引用，静默忽略: %s", filename)
        return ""

    elif mode == "keep":
        # 保留原链接
        original = f"![[{filename}]]"
        logger.debug("循环引用，保留原链接: %s", original)
        return original

    elif mode == "placeholder":
        # 插入占位符（使用国际化）
        placeholder = t("link_processing.circular_reference", filename=filename)
        logger.debug("循环引用，插入占位符: %s", placeholder)
        return placeholder

    else:
        # 默认：占位符
        return t("link_processing.circular_reference", filename=filename)


def _handle_max_depth_error() -> str:
    """
    处理达到最大深度错误

    返回:
        str: 根据配置返回的文本
    """
    mode = config_manager.get_max_depth_reached_mode()

    if mode == "ignore":
        # 静默忽略
        return ""

    elif mode == "placeholder":
        # 插入占位符（使用国际化）
        text = t("link_processing.max_depth_reached")
        logger.debug("达到最大深度，插入占位符: %s", text)
        return text

    else:
        # 默认：占位符
        return t("link_processing.max_depth_reached")
