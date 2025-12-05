"""
配置常量定义
"""

# 默认日志配置
DEFAULT_LOGGING_CONFIG = {
    "logging": {
        "enable": True,
        "level": "info",
        "file_prefix": "gongwen",
        "retention_days": 7,
        "console_enable": True,
        "console_level": "info"
    }
}

# 默认符号校对配置
DEFAULT_SYMBOL_SETTINGS = {
    "symbol_settings": {
        "engine_settings": {
            "enable_symbol_pairing": True,
            "enable_symbol_correction": True,
        },
        "symbol_pairing": {
            "pairs": []
        },
        "symbol_map": {},
    }
}

# 默认错别字配置
DEFAULT_TYPOS_SETTINGS = {
    "typos_settings": {
        "engine_settings": {
            "enable_typos_rule": True,
        },
        "typos": {}
    }
}

# 默认敏感词配置
DEFAULT_SENSITIVE_WORDS_SETTINGS = {
    "sensitive_words_settings": {
        "engine_settings": {
            "enable_sensitive_word": True,
        },
        "sensitive_words": {}
    }
}


# 默认GUI配置
DEFAULT_GUI_CONFIG = {
    "gui_config": {
        "window": {
            "center_panel_width": 400,
            "left_panel_width": 400,          # 批量面板宽度
            "right_panel_width": 300,         # 模板面板宽度
            "center_panel_screen_x": 0,
            "window_y": 0,
            "default_mode": "single",
            "default_height": 740,
            "min_height": 720,
            "auto_center": True,
            "remember_gui_state": True
        },
        "component": {
            "file_drop_height": 200
        },
        "theme": {
            "default_theme": "morph"
        },
        "transparency": {
            "enabled": True,
            "default_value": 0.95,
            "min_value": 0.8,
            "max_value": 1.0
        },
        "template": {
            "md_default_template": "docx"
        }
    }
}

# 默认链接配置
DEFAULT_LINK_CONFIG = {
    "link_config": {
        # 格式设置（生成MD时的链接格式）
        "format": {
            "image_link_format": "wiki",
            "image_embed": True,
            "md_file_link_format": "wiki",
            "md_file_embed": True
        },
        # 非嵌入链接处理（MD转文档/表格时）
        "non_embed_links": {
            "wiki_mode": "extract_text",
            "markdown_mode": "extract_text"
        },
        # 嵌入链接处理（MD转文档/表格时）
        "embed_links": {
            "enabled": True,
            "image_mode": "embed",
            "md_file_mode": "embed"
        },
        # 嵌入详细配置
        "embedding": {
            "max_depth": 3
        },
        # 路径解析配置
        "path_resolution": {
            "search_dirs": [".", "assets", "images", "attachments"]
        },
        # 错误处理配置
        "error_handling": {
            "file_not_found": "placeholder",
            "file_not_found_text": "⚠️ 文件未找到: {filename}",
            "detect_circular": True,
            "circular_reference": "placeholder",
            "circular_text": "⚠️ 检测到循环引用: {filename}",
            "max_depth_reached": "placeholder",
            "max_depth_text": "⚠️ 达到最大嵌入深度"
        }
    }
}

# 默认文件处理配置
DEFAULT_FILE_DEFAULTS_CONFIG = {
    "file_defaults": {
        "document": {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "enable_symbol_pairing": True,
            "enable_symbol_correction": True,
            "enable_typos_rule": True,
            "enable_sensitive_word": True
        },
        "spreadsheet": {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "merge_mode": 3
        },
        "image": {
            "to_md_keep_images": True,
            "to_md_enable_ocr": True,
            "compress_mode": "lossless",
            "size_limit": 200,
            "size_unit": "KB",
            "pdf_quality": "original",
            "tiff_mode": "smart"
        },
        "layout": {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "render_dpi": 300
        }
    }
}

# 默认输出配置
DEFAULT_OUTPUT_CONFIG = {
    "output_config": {
        "directory": {
            "mode": "source",
            "custom_path": "",
            "create_date_subfolder": False,
            "date_folder_format": "%Y-%m-%d"
        },
        "behavior": {
            "auto_open_folder": False
        },
        "intermediate_files": {
            "save_to_output": True
        }
    }
}

# 合并的默认配置常量
DEFAULT_CONFIG = {
    **DEFAULT_LOGGING_CONFIG,
    **DEFAULT_GUI_CONFIG,
    **DEFAULT_LINK_CONFIG,
    **DEFAULT_FILE_DEFAULTS_CONFIG,
    **DEFAULT_OUTPUT_CONFIG,
    **DEFAULT_SYMBOL_SETTINGS,
    **DEFAULT_TYPOS_SETTINGS,
    **DEFAULT_SENSITIVE_WORDS_SETTINGS
}

# 默认软件优先级配置
DEFAULT_SOFTWARE_PRIORITY_CONFIG = {
    "software_priority": {
        "default_priority": {
            "word_processors": ["wps_writer", "msoffice_word", "libreoffice"],
            "spreadsheet_processors": ["wps_spreadsheets", "msoffice_excel", "libreoffice"]
        },
        "special_conversions": {
            "odt": ["msoffice_word", "libreoffice"],
            "ods": ["msoffice_excel", "libreoffice"],
            "pdf_to_office": ["msoffice_word", "libreoffice"],
            "document_to_pdf": ["wps_writer", "msoffice_word", "libreoffice"],
            "spreadsheet_to_pdf": ["wps_spreadsheets", "msoffice_excel", "libreoffice"]
        }
    }
}

# 软件标识符到COM对象的映射
SOFTWARE_ID_MAPPING = {
    "wps_writer": "Kwps.Application",
    "wps_spreadsheets": "Ket.Application",
    "msoffice_word": "Word.Application",
    "msoffice_excel": "Excel.Application",
    "libreoffice": "soffice"
}

# 配置文件映射
CONFIG_FILES = {
    "logger_config": "logger_config.toml",
    "gui_config": "gui_config.toml",
    "link_config": "link_config.toml",
    "file_defaults": "file_defaults.toml",
    "output_config": "output_config.toml",
    "symbol_settings": "symbol_settings.toml",
    "typos_settings": "typos_settings.toml",
    "sensitive_words_settings": "sensitive_words.toml",
    "software_priority": "software_priority.toml"
}
