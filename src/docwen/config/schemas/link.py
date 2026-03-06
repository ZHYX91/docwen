"""
链接配置模块

对应配置文件：link_config.toml

包含：
    - DEFAULT_LINK_CONFIG: 默认链接处理配置
    - LinkConfigMixin: 链接配置获取方法
"""

from typing import Any

from ..safe_logger import safe_log

# ==============================================================================
#                              默认配置
# ==============================================================================

DEFAULT_LINK_CONFIG = {
    "link_config": {
        # 格式设置（生成MD时的链接格式）
        "format": {
            # 图片链接样式: markdown_embed, markdown_link, wiki_embed, wiki_link
            "image_link_style": "wiki_embed",
            # MD文件链接样式: markdown_link, wiki_embed, wiki_link
            "md_file_link_style": "wiki_embed",
        },
        # 非嵌入链接处理（MD转文档/表格时）
        "non_embed_links": {
            "wiki_mode": "extract_text",  # keep, extract_text, remove
            "markdown_mode": "extract_text",  # keep, extract_text, remove
        },
        # 嵌入链接处理（MD转文档/表格时）
        "embed_links": {
            "wiki_image_mode": "embed",  # keep, extract_text, remove, embed
            "markdown_image_mode": "embed",  # keep, extract_text, remove, embed
            "md_file_mode": "embed",  # keep, extract_text, remove, embed
        },
        # 嵌入详细配置
        "embedding": {
            "max_depth": 3  # 最大递归嵌入深度
        },
        # 路径解析配置
        "path_resolution": {"search_dirs": [".", "assets", "images", "attachments"]},
        # 错误处理配置
        "error_handling": {
            "file_not_found": "placeholder",  # ignore, keep, placeholder
            "file_not_found_text": "⚠️ 文件未找到: {filename}",
            "detect_circular": True,
            "circular_reference": "placeholder",  # ignore, keep, placeholder
            "circular_text": "⚠️ 检测到循环引用: {filename}",
            "max_depth_reached": "placeholder",  # ignore, keep, placeholder
            "max_depth_text": "⚠️ 达到最大嵌入深度",
        },
    }
}

# 配置文件名
CONFIG_FILE = "link_config.toml"

# ==============================================================================
#                              Mixin 类
# ==============================================================================


class LinkConfigMixin:
    """
    链接配置获取方法 Mixin

    提供链接处理相关配置的访问方法，包括格式、嵌入、路径解析和错误处理。
    """

    _configs: dict[str, dict[str, Any]]

    # --------------------------------------------------------------------------
    # 第一层：配置块
    # --------------------------------------------------------------------------

    def get_link_config_block(self) -> dict[str, Any]:
        """
        获取整个链接配置块

        返回:
            Dict[str, Any]: 链接配置字典
        """
        return self._configs.get("link_config", {})

    # --------------------------------------------------------------------------
    # 第二层：子表
    # --------------------------------------------------------------------------

    def get_link_format_config(self) -> dict[str, Any]:
        """
        获取链接格式配置子表

        返回:
            Dict[str, Any]: 链接格式配置字典
        """
        return self.get_link_config_block().get("format", {})

    def get_non_embed_links_config(self) -> dict[str, Any]:
        """
        获取非嵌入链接配置子表

        返回:
            Dict[str, Any]: 非嵌入链接配置字典
        """
        return self.get_link_config_block().get("non_embed_links", {})

    def get_embed_links_config(self) -> dict[str, Any]:
        """
        获取嵌入链接配置子表

        返回:
            Dict[str, Any]: 嵌入链接配置字典
        """
        return self.get_link_config_block().get("embed_links", {})

    def get_embedding_config(self) -> dict[str, Any]:
        """
        获取嵌入详细配置子表

        返回:
            Dict[str, Any]: 嵌入详细配置字典
        """
        return self.get_link_config_block().get("embedding", {})

    def get_path_resolution_config(self) -> dict[str, Any]:
        """
        获取路径解析配置子表

        返回:
            Dict[str, Any]: 路径解析配置字典
        """
        return self.get_link_config_block().get("path_resolution", {})

    def get_link_error_handling_config(self) -> dict[str, Any]:
        """
        获取链接错误处理配置子表

        返回:
            Dict[str, Any]: 错误处理配置字典
        """
        return self.get_link_config_block().get("error_handling", {})

    # --------------------------------------------------------------------------
    # 第三层：具体配置值 - 链接格式
    # --------------------------------------------------------------------------

    def get_markdown_link_style_settings(self) -> dict[str, Any]:
        """
        获取Markdown链接格式设置

        返回:
            Dict[str, Any]: 包含以下键的字典：
                - image_link_style: str - 图片链接样式
                    可选值: "markdown_embed", "markdown_link", "wiki_embed", "wiki_link"
                - md_file_link_style: str - MD文件链接样式
                    可选值: "markdown_link", "wiki_embed", "wiki_link"
        """
        link_format_config = self.get_link_format_config()
        settings = {
            "image_link_style": link_format_config.get("image_link_style", "wiki_embed"),
            "md_file_link_style": link_format_config.get("md_file_link_style", "wiki_embed"),
        }
        safe_log.debug("获取Markdown链接格式设置: %s", settings)
        return settings

    def get_image_link_style(self) -> str:
        """
        获取图片链接样式

        返回:
            str: 样式值 ("markdown_embed", "markdown_link", "wiki_embed", "wiki_link")
        """
        config = self.get_link_format_config()
        style = config.get("image_link_style", "wiki_embed")
        safe_log.debug("获取图片链接样式: %s", style)
        return style

    def get_md_file_link_style(self) -> str:
        """
        获取MD文件链接样式

        返回:
            str: 样式值 ("markdown_link", "wiki_embed", "wiki_link")
        """
        config = self.get_link_format_config()
        style = config.get("md_file_link_style", "wiki_embed")
        safe_log.debug("获取MD文件链接样式: %s", style)
        return style

    # --------------------------------------------------------------------------
    # 第三层：具体配置值 - 非嵌入链接
    # --------------------------------------------------------------------------

    def get_wiki_link_mode(self) -> str:
        """
        获取Wiki链接处理模式

        返回:
            str: 处理模式 ("keep", "extract_text", "remove")
        """
        config = self.get_non_embed_links_config()
        mode = config.get("wiki_mode", "extract_text")
        safe_log.debug("获取Wiki链接处理模式: %s", mode)
        return mode

    def get_markdown_link_mode(self) -> str:
        """
        获取Markdown链接处理模式

        返回:
            str: 处理模式 ("keep", "extract_text", "remove")
        """
        config = self.get_non_embed_links_config()
        mode = config.get("markdown_mode", "extract_text")
        safe_log.debug("获取Markdown链接处理模式: %s", mode)
        return mode

    # --------------------------------------------------------------------------
    # 第三层：具体配置值 - 嵌入链接
    # --------------------------------------------------------------------------

    def get_wiki_embed_image_mode(self) -> str:
        """
        获取Wiki嵌入图片处理模式（![[image.png]]）

        返回:
            str: 处理模式 ("keep", "extract_text", "remove", "embed")
        """
        config = self.get_embed_links_config()
        mode = config.get("wiki_image_mode", "embed")
        safe_log.debug("获取Wiki嵌入图片处理模式: %s", mode)
        return mode

    def get_markdown_embed_image_mode(self) -> str:
        """
        获取Markdown嵌入图片处理模式（![alt](image.png)）

        返回:
            str: 处理模式 ("keep", "extract_text", "remove", "embed")
        """
        config = self.get_embed_links_config()
        mode = config.get("markdown_image_mode", "embed")
        safe_log.debug("获取Markdown嵌入图片处理模式: %s", mode)
        return mode

    def get_embed_md_file_mode(self) -> str:
        """
        获取嵌入MD文件处理模式

        返回:
            str: 处理模式 ("keep", "extract_text", "remove", "embed")
        """
        config = self.get_embed_links_config()
        mode = config.get("md_file_mode", "embed")
        safe_log.debug("获取嵌入MD文件处理模式: %s", mode)
        return mode

    # --------------------------------------------------------------------------
    # 第三层：具体配置值 - 嵌入详细配置
    # --------------------------------------------------------------------------

    def get_max_embed_depth(self) -> int:
        """
        获取最大嵌入深度

        返回:
            int: 最大递归嵌入深度
        """
        config = self.get_embedding_config()
        depth = config.get("max_depth", 3)
        safe_log.debug("获取最大嵌入深度: %d", depth)
        return depth

    # --------------------------------------------------------------------------
    # 第三层：具体配置值 - 路径解析
    # --------------------------------------------------------------------------

    def get_search_dirs(self) -> list[str]:
        """
        获取搜索目录列表

        返回:
            List[str]: 文件搜索子目录列表
        """
        config = self.get_path_resolution_config()
        dirs = config.get("search_dirs", [".", "assets", "images", "attachments"])
        safe_log.debug("获取搜索目录列表: %s", dirs)
        return dirs

    # --------------------------------------------------------------------------
    # 第三层：具体配置值 - 错误处理
    # --------------------------------------------------------------------------

    def get_file_not_found_mode(self) -> str:
        """
        获取文件未找到处理模式

        返回:
            str: 处理模式 ("ignore", "keep", "placeholder")
        """
        config = self.get_link_error_handling_config()
        mode = config.get("file_not_found", "placeholder")
        safe_log.debug("获取文件未找到处理模式: %s", mode)
        return mode

    def get_file_not_found_text(self) -> str:
        """
        获取文件未找到占位文本

        返回:
            str: 占位文本模板，支持 {filename} 占位符
        """
        config = self.get_link_error_handling_config()
        text = config.get("file_not_found_text", "⚠️ 文件未找到: {filename}")
        safe_log.debug("获取文件未找到占位文本: %s", text)
        return text

    def is_circular_detection_enabled(self) -> bool:
        """
        是否启用循环引用检测

        返回:
            bool: 是否检测循环引用
        """
        config = self.get_link_error_handling_config()
        enabled = config.get("detect_circular", True)
        safe_log.debug("循环引用检测启用状态: %s", enabled)
        return enabled

    def get_circular_reference_mode(self) -> str:
        """
        获取循环引用处理模式

        返回:
            str: 处理模式 ("ignore", "keep", "placeholder")
        """
        config = self.get_link_error_handling_config()
        mode = config.get("circular_reference", "placeholder")
        safe_log.debug("获取循环引用处理模式: %s", mode)
        return mode

    def get_circular_text(self) -> str:
        """
        获取循环引用占位文本

        返回:
            str: 占位文本模板，支持 {filename} 占位符
        """
        config = self.get_link_error_handling_config()
        text = config.get("circular_text", "⚠️ 检测到循环引用: {filename}")
        safe_log.debug("获取循环引用占位文本: %s", text)
        return text

    def get_max_depth_reached_mode(self) -> str:
        """
        获取达到最大深度处理模式

        返回:
            str: 处理模式 ("ignore", "keep", "placeholder")
        """
        config = self.get_link_error_handling_config()
        mode = config.get("max_depth_reached", "placeholder")
        safe_log.debug("获取最大深度处理模式: %s", mode)
        return mode

    def get_max_depth_text(self) -> str:
        """
        获取最大深度警告文本

        返回:
            str: 警告文本
        """
        config = self.get_link_error_handling_config()
        text = config.get("max_depth_text", "⚠️ 达到最大嵌入深度")
        safe_log.debug("获取最大深度警告文本: %s", text)
        return text
