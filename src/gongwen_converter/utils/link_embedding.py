"""
链接嵌入处理模块

提供嵌入链接的展开功能，支持：
- 图片嵌入：![[image.png]] 或 ![](image.png) → {{IMAGE:path}}
- MD文件嵌入：![[file.md]] → 文件内容（递归处理）
- 路径解析：绝对路径、相对路径、仅文件名
- 循环引用检测：防止无限递归
- 错误处理：文件未找到、循环引用、最大深度等
"""

import os
import re
import logging
from typing import Optional, Set

from gongwen_converter.config.config_manager import config_manager

logger = logging.getLogger(__name__)

# 支持的文件扩展名
IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp', '.ico'}
MD_EXTENSIONS = {'.md', '.markdown'}

# 链接识别正则表达式
# 支持转义的竖线 \|（用于Markdown表格中的Wiki链接）
WIKI_EMBED_PATTERN = r'!\[\[((?:[^|\]\\]|\\[|\]])+)(?:(?<!\\)\|((?:[^\]\\]|\\[|\]])+))?\]\]'  # ![[target]] 或 ![[target|text]]
MD_EMBED_IMAGE_PATTERN = r'!\[([^\]]*)\]\(([^)]+)\)'  # ![alt](url)


def process_markdown_links(
    text: str,
    source_file_path: str,
    visited_files: Optional[Set[str]] = None,
    depth: int = 0
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
    source_file_path = os.path.abspath(source_file_path)
    
    logger.info("展开嵌入内容 | 源文件: %s | 深度: %d", os.path.basename(source_file_path), depth)
    
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
    embed_count = {'images': 0, 'markdown': 0}
    
    # 1. 处理Wiki嵌入链接：![[target]]
    wiki_embeds = re.finditer(WIKI_EMBED_PATTERN, text)
    for match in wiki_embeds:
        original_link = match.group(0)  # 完整的 ![[target]]
        link_target = match.group(1).strip().replace(r'\|', '|')  # target，还原转义的竖线
        
        logger.debug("识别到Wiki嵌入链接: %s", original_link)
        
        # 处理嵌入
        replacement = _process_single_embed(
            link_target, original_link, source_file_path,
            visited_files, depth, embed_count
        )
        
        if replacement is not None:
            result = result.replace(original_link, replacement, 1)
    
    # 2. 处理Markdown嵌入图片：![alt](url)
    md_embeds = re.finditer(MD_EMBED_IMAGE_PATTERN, text)
    for match in md_embeds:
        original_link = match.group(0)  # 完整的 ![alt](url)
        alt_text = match.group(1)  # alt
        link_target = match.group(2).strip()  # url
        
        logger.debug("识别到Markdown嵌入图片: %s", original_link)
        
        # 处理嵌入
        replacement = _process_single_embed(
            link_target, original_link, source_file_path,
            visited_files, depth, embed_count
        )
        
        if replacement is not None:
            result = result.replace(original_link, replacement, 1)
    
    logger.info("嵌入处理完成 | 图片: %d个 | MD文件: %d个", embed_count['images'], embed_count['markdown'])
    
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
    
    # 1. 处理Wiki链接：[[link]] 或 [[link|text]]（支持转义竖线 \|）
    wiki_pattern = r'\[\[(?:((?:[^|\]\\]|\\[|\]])+)(?<!\\)\|)?((?:[^\]\\]|\\[|\]])+)\]\]'
    
    # 查找所有匹配
    wiki_matches = list(re.finditer(wiki_pattern, text))
    logger.debug(f"找到 {len(wiki_matches)} 个Wiki链接")
    for match in wiki_matches:
        logger.debug(f"  Wiki链接: {match.group(0)}")
    
    if wiki_mode == "keep":
        # 保留原样
        logger.debug("Wiki链接模式: keep（保留原样）")
    elif wiki_mode == "extract_text":
        # 提取显示文本（group(2)是显示文本）
        # 先替换，然后还原转义的竖线
        logger.debug("Wiki链接模式: extract_text（提取文本）")
        result = re.sub(wiki_pattern, lambda m: m.group(2).replace(r'\|', '|'), result)
        logger.debug(f"Wiki链接处理后: {result[:200]}")
    elif wiki_mode == "remove":
        # 完全移除
        logger.debug("Wiki链接模式: remove（移除）")
        result = re.sub(wiki_pattern, '', result)
    
    # 2. 处理Markdown链接：[text](url)
    markdown_pattern = r'\[([^\]]+)\]\([^)]+\)'
    
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
        result = re.sub(markdown_pattern, r'\1', result)
        logger.debug(f"Markdown链接处理后: {result[:200]}")
    elif markdown_mode == "remove":
        # 完全移除
        logger.debug("Markdown链接模式: remove（移除）")
        result = re.sub(markdown_pattern, '', result)
    
    logger.debug(f"_process_non_embed_links 输出: {result[:200] if result else 'None'}")
    return result


def _process_single_embed(
    link_target: str,
    original_link: str,
    source_file_path: str,
    visited_files: Set[str],
    depth: int,
    embed_count: dict
) -> Optional[str]:
    """
    处理单个嵌入链接
    
    参数:
        link_target: 链接目标（文件名或路径）
        original_link: 原始链接文本
        source_file_path: 源文件路径
        visited_files: 已访问文件集合
        depth: 当前深度
        embed_count: 统计计数器
    
    返回:
        Optional[str]: 替换文本，None表示不替换
    """
    # 解析文件路径
    resolved_path = resolve_file_path(link_target, source_file_path)
    
    if resolved_path is None:
        # 文件未找到
        logger.warning("文件未找到: %s（来自: %s）", link_target, os.path.basename(source_file_path))
        return _handle_not_found_error(link_target)
    
    logger.debug("路径解析成功: %s → %s", link_target, resolved_path)
    
    # 判断文件类型
    file_type = get_file_type(resolved_path)
    
    if file_type == 'image':
        embed_count['images'] += 1
        return process_embedded_image(resolved_path, original_link)
    
    elif file_type == 'markdown':
        embed_count['markdown'] += 1
        return process_embedded_md_file(
            resolved_path, source_file_path,
            visited_files, depth
        )
    
    else:
        # 未知文件类型
        logger.warning("未知文件类型: %s", resolved_path)
        return None


def resolve_file_path(
    link_target: str,
    source_file_path: str
) -> Optional[str]:
    """
    解析嵌入文件的路径
    
    解析优先级：
    1. 绝对路径 → 直接使用
    2. 相对路径（含/或\）→ 相对于源文件目录
    3. 仅文件名 → 在search_dirs中按顺序搜索
    
    参数:
        link_target: 链接目标（从![[这里]]或![]()中提取）
        source_file_path: 源MD文件的绝对路径
    
    返回:
        Optional[str]: 解析后的绝对路径，或None（未找到文件）
    """
    logger.debug("解析文件路径: %s", link_target)
    
    # 获取源文件所在目录
    source_dir = os.path.dirname(source_file_path)
    
    # 1. 绝对路径：直接使用
    if os.path.isabs(link_target):
        normalized = os.path.normpath(link_target)
        if os.path.exists(normalized):
            logger.debug("解析为绝对路径: %s", normalized)
            return normalized
        else:
            logger.debug("绝对路径文件不存在: %s", normalized)
            return None
    
    # 2. 相对路径（包含路径分隔符）
    if '/' in link_target or '\\' in link_target:
        full_path = os.path.normpath(os.path.join(source_dir, link_target))
        if os.path.exists(full_path):
            logger.debug("解析为相对路径: %s", full_path)
            return full_path
        else:
            logger.debug("相对路径文件不存在: %s", full_path)
            return None
    
    # 3. 仅文件名：在search_dirs中搜索
    search_dirs = config_manager.get_search_dirs()
    logger.debug("搜索文件: %s（在目录: %s）", link_target, search_dirs)
    
    for search_dir in search_dirs:
        search_path = os.path.normpath(os.path.join(source_dir, search_dir, link_target))
        if os.path.exists(search_path):
            logger.debug("在搜索目录 '%s' 中找到文件: %s", search_dir, search_path)
            return search_path
    
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
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext in IMAGE_EXTENSIONS:
        logger.debug("文件类型: 图片 (%s)", ext)
        return 'image'
    elif ext in MD_EXTENSIONS:
        logger.debug("文件类型: Markdown (%s)", ext)
        return 'markdown'
    else:
        logger.debug("文件类型: 未知 (%s)", ext)
        return 'unknown'


def process_embedded_image(
    image_path: str,
    original_link: str
) -> str:
    """
    处理嵌入图片
    
    根据配置决定如何处理嵌入的图片链接。
    
    参数:
        image_path: 图片文件的绝对路径
        original_link: 原始链接文本（用于keep模式）
    
    返回:
        str: 处理后的文本
    """
    mode = config_manager.get_embed_image_mode()
    logger.debug("处理嵌入图片: %s | 模式: %s", os.path.basename(image_path), mode)
    
    if mode == "embed":
        # 生成图片占位符（后续会被docx_processor替换为实际图片）
        placeholder = f"{{{{IMAGE:{image_path}}}}}"
        logger.info("生成图片占位符: %s", placeholder)
        return placeholder
    
    elif mode == "keep":
        # 保留原链接
        logger.debug("保留原图片链接: %s", original_link)
        return original_link
    
    elif mode == "extract_text":
        # 提取文件名
        filename = os.path.basename(image_path)
        logger.debug("提取图片文件名: %s", filename)
        return filename
    
    elif mode == "remove":
        # 完全移除
        logger.debug("移除图片链接")
        return ""
    
    else:
        # 未知模式，使用默认（embed）
        logger.warning("未知的图片嵌入模式: %s，使用默认embed", mode)
        return f"{{{{IMAGE:{image_path}}}}}"


def process_embedded_md_file(
    md_path: str,
    source_file_path: str,
    visited_files: Set[str],
    depth: int
) -> str:
    """
    处理嵌入MD文件（递归）
    
    根据配置决定如何处理嵌入的MD文件。
    
    参数:
        md_path: MD文件的绝对路径
        source_file_path: 当前源文件路径（用于日志）
        visited_files: 已访问文件集合
        depth: 当前递归深度
    
    返回:
        str: 处理后的内容
    """
    mode = config_manager.get_embed_md_file_mode()
    logger.debug("处理嵌入MD文件: %s | 模式: %s", os.path.basename(md_path), mode)
    
    # 规范化路径用于循环检测
    normalized_path = os.path.abspath(md_path)
    
    # 检查循环引用
    if config_manager.is_circular_detection_enabled():
        if normalized_path in visited_files:
            logger.error("检测到循环引用: %s（被 %s 嵌入）", 
                        os.path.basename(md_path), 
                        os.path.basename(source_file_path))
            return _handle_circular_error(os.path.basename(md_path))
    
    if mode == "embed":
        # 读取并递归处理
        try:
            logger.info("读取MD文件: %s（深度: %d）", os.path.basename(md_path), depth + 1)
            
            with open(md_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            logger.debug("读取到 %d 字符", len(content))
            
            # 将当前文件添加到visited（递归前）
            visited_files.add(normalized_path)
            
            try:
                # 递归展开内容
                expanded = process_markdown_links(
                    content,
                    md_path,
                    visited_files,
                    depth + 1
                )
            finally:
                # 递归返回后移除（允许其他路径再次访问）
                visited_files.discard(normalized_path)
            
            logger.info("MD文件展开完成: %s", os.path.basename(md_path))
            return expanded
            
        except Exception as e:
            logger.error("读取MD文件失败: %s | 错误: %s", md_path, str(e))
            return _handle_not_found_error(os.path.basename(md_path))
    
    elif mode == "keep":
        # 保留原链接
        original = f"![[{os.path.basename(md_path)}]]"
        logger.debug("保留原MD链接: %s", original)
        return original
    
    elif mode == "extract_text":
        # 提取文件名
        filename = os.path.basename(md_path)
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
            with open(md_path, 'r', encoding='utf-8') as f:
                content = f.read()
            return process_markdown_links(content, md_path, visited_files, depth + 1)
        except Exception as e:
            logger.error("读取MD文件失败: %s", str(e))
            return _handle_not_found_error(os.path.basename(md_path))


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
        # 插入占位符
        template = config_manager.get_file_not_found_text()
        placeholder = template.format(filename=filename)
        logger.debug("文件未找到，插入占位符: %s", placeholder)
        return placeholder
    
    else:
        # 默认：占位符
        return f"⚠️ 文件未找到: {filename}"


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
        # 插入占位符
        template = config_manager.get_circular_text()
        placeholder = template.format(filename=filename)
        logger.debug("循环引用，插入占位符: %s", placeholder)
        return placeholder
    
    else:
        # 默认：占位符
        return f"⚠️ 检测到循环引用: {filename}"


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
        # 插入占位符
        text = config_manager.get_max_depth_text()
        logger.debug("达到最大深度，插入占位符: %s", text)
        return text
    
    else:
        # 默认：占位符
        return "⚠️ 达到最大嵌入深度"
