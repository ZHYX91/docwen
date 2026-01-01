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

# 默认校对主配置（开关和跳过规则）
DEFAULT_PROOFREAD_CONFIG = {
    "proofread_config": {
        "engine": {
            "enable_typos_rule": True,
            "enable_symbol_pairing": True,
            "enable_symbol_correction": True,
            "enable_sensitive_word": True,
            "skip_code_blocks": True,
            "skip_quote_blocks": False,
        }
    }
}

# 默认符号校对配置（符号配对和映射表）
DEFAULT_PROOFREAD_SYMBOLS = {
    "proofread_symbols": {
        "symbol_pairing": {
            "pairs": []
        },
        "symbol_map": {},
    }
}

# 默认错别字配置（错别字映射表）
DEFAULT_PROOFREAD_TYPOS = {
    "proofread_typos": {
        "typos": {}
    }
}

# 默认敏感词配置（敏感词映射表）
DEFAULT_PROOFREAD_SENSITIVE = {
    "proofread_sensitive": {
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
        },
        "language": {
            "locale": "zh_CN"  # 默认语言: zh_CN（简体中文）, en_US（英文）
        }
    }
}

# 默认链接配置
DEFAULT_LINK_CONFIG = {
    "link_config": {
        # 格式设置（生成MD时的链接格式）
        "format": {
            # 图片链接样式: markdown_embed, markdown_link, wiki_embed, wiki_link
            "image_link_style": "wiki_embed",
            # MD文件链接样式: markdown_link, wiki_embed, wiki_link
            "md_file_link_style": "wiki_embed"
        },
        # 非嵌入链接处理（MD转文档/表格时）
        "non_embed_links": {
            "wiki_mode": "extract_text",
            "markdown_mode": "extract_text"
        },
        # 嵌入链接处理（MD转文档/表格时）
        "embed_links": {
            "wiki_image_mode": "embed",
            "markdown_image_mode": "embed",
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

# 默认转换默认值配置（控制 GUI 界面的默认值设置）
DEFAULT_CONVERSION_DEFAULTS_CONFIG = {
    "conversion_defaults": {
        "document": {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "to_md_remove_numbering": True,
            "to_md_add_numbering": False,
            "to_md_default_scheme": "gongwen_standard",
            "to_md_enable_optimization": False,
            "to_md_optimization_type": "公文",
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
        },
        "text": {
            "to_docx_remove_numbering": True,
            "to_docx_add_numbering": True,
            "to_docx_default_scheme": "gongwen_standard",
            "to_xlsx_remove_numbering": True,
            "to_xlsx_add_numbering": False,
            "to_xlsx_default_scheme": "hierarchical_standard"
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

# 默认标题序号配置
DEFAULT_HEADING_NUMBERING_CONFIG = {
    "heading_numbering_add": {
        "settings": {
            "default_scheme": "gongwen_standard"
        },
        # 数字样式定义（供模板占位符使用）
        "number_styles": {
            "chinese_lower": {
                "name": "小写中文数字",
                "description": "一、二、三、四、五、六、七、八、九、十、十一、十二..."
            },
            "chinese_upper": {
                "name": "大写中文数字",
                "description": "壹、贰、叁、肆、伍、陆、柒、捌、玖、拾、拾壹、拾贰..."
            },
            "arabic_half": {
                "name": "半角阿拉伯数字",
                "description": "1, 2, 3, 4, 5, 6, 7, 8, 9, 10..."
            },
            "arabic_full": {
                "name": "全角阿拉伯数字",
                "description": "１、２、３、４、５、６、７、８、９、１０..."
            },
            "arabic_circled": {
                "name": "带圈阿拉伯数字",
                "description": "①②③④⑤⑥⑦⑧⑨⑩⑪⑫... (最多50)"
            },
            "letter_upper": {
                "name": "大写拉丁字母",
                "description": "A, B, C, D, E, F, G..."
            },
            "letter_lower": {
                "name": "小写拉丁字母",
                "description": "a, b, c, d, e, f, g..."
            },
            "roman_upper": {
                "name": "大写罗马数字",
                "description": "I, II, III, IV, V, VI, VII..."
            },
            "roman_lower": {
                "name": "小写罗马数字",
                "description": "i, ii, iii, iv, v, vi, vii..."
            }
        },
        "schemes": {
            # 方案1：公文标准
            "gongwen_standard": {
                "name": "公文标准",
                "description": "中国公文标准序号格式",
                "level_1": {"format": "{1.chinese_lower}、"},
                "level_2": {"format": "（{2.chinese_lower}）"},
                "level_3": {"format": "{3.arabic_half}. "},
                "level_4": {"format": "（{4.arabic_half}）"},
                "level_5": {"format": "{5.arabic_circled}"},
                "level_6": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half} "},
                "level_7": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half} "},
                "level_8": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half} "},
                "level_9": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half}.{9.arabic_half} "}
            },
            # 方案2：层级数字标准
            "hierarchical_standard": {
                "name": "层级数字标准",
                "description": "层级递进格式",
                "level_1": {"format": "{1.arabic_half} "},
                "level_2": {"format": "{1.arabic_half}.{2.arabic_half} "},
                "level_3": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half} "},
                "level_4": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half} "},
                "level_5": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half} "},
                "level_6": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half} "},
                "level_7": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half} "},
                "level_8": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half} "},
                "level_9": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half}.{9.arabic_half} "}
            },
            # 方案3：法律条文标准
            "legal_standard": {
                "name": "法律条文标准",
                "description": "中国法律文本格式",
                "level_1": {"format": "第{1.chinese_lower}编　"},
                "level_2": {"format": "第{2.chinese_lower}章　"},
                "level_3": {"format": "第{3.chinese_lower}节　"},
                "level_4": {"format": "第{4.chinese_lower}条　"},
                "level_5": {"format": "（{5.chinese_lower}）"},
                "level_6": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half} "},
                "level_7": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half} "},
                "level_8": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half} "},
                "level_9": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half}.{9.arabic_half} "}
            }
        }
    }
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

# 默认代码样式配置
DEFAULT_STYLE_CODE_CONFIG = {
    "style_code": {
        "docx_to_md": {
            "paragraph_styles": ["HTML Preformatted", "Code Block", "代码块"],
            "character_styles": [
                "HTML Code", "HTML Typewriter", "HTML Keyboard", "HTML Sample",
                "HTML Variable", "HTML Definition", "HTML Cite", "HTML Address",
                "HTML Acronym", "Inline Code", "Code", "Source Code",
                "行内代码", "代码", "源代码", "源码"
            ],
            "full_paragraph_as_block": True,
            "fuzzy_match_enabled": True,
            "fuzzy_keywords": ["code", "代码", "源码"],
            "shading": {
                "wps_enabled": True,
                "word_enabled": True
            }
        },
        "md_to_docx": {
            "inline_code_style": "Inline Code",
            "code_block_style": "Code Block"
        }
    }
}

# 默认公式样式配置
DEFAULT_STYLE_FORMULA_CONFIG = {
    "style_formula": {
        "md_to_docx": {
            "inline_formula_style": "Inline Formula",
            "formula_block_style": "Formula Block"
        }
    }
}

# 默认表格样式配置
DEFAULT_STYLE_TABLE_CONFIG = {
    "style_table": {
        "md_to_docx": {
            "table_style": "Three Line Table",
            "table_content_style": "Table Content"
        }
    }
}

# 默认引用样式配置
DEFAULT_STYLE_QUOTE_CONFIG = {
    "style_quote": {
        "docx_to_md": {
            "level_styles": {
                "Quote 1": 1, "Quote 2": 2, "Quote 3": 3, "Quote 4": 4, "Quote 5": 5,
                "Quote 6": 6, "Quote 7": 7, "Quote 8": 8, "Quote 9": 9,
                "引用 1": 1, "引用 2": 2, "引用 3": 3, "引用 4": 4, "引用 5": 5,
                "引用 6": 6, "引用 7": 7, "引用 8": 8, "引用 9": 9
            },
            "paragraph_styles": ["Quote", "Block Text", "Intense Quote", "引用", "明显引用"],
            "character_styles": ["Quote Char", "引用字符"],
            "full_paragraph_as_block": True,
            "fuzzy_match_enabled": True,
            "fuzzy_keywords": ["quote", "引用"]
        },
        "md_to_docx": {
            "level_1_style": "Quote 1",
            "level_2_style": "Quote 2",
            "level_3_style": "Quote 3",
            "level_4_style": "Quote 4",
            "level_5_style": "Quote 5",
            "level_6_style": "Quote 6",
            "level_7_style": "Quote 7",
            "level_8_style": "Quote 8",
            "level_9_style": "Quote 9"
        }
    }
}

# 默认转换行为配置（直接控制转换引擎的行为规则）
DEFAULT_CONVERSION_CONFIG = {
    "conversion_config": {
        # DOCX → MD 格式设置
        "docx_to_md": {
            "preserve_formatting": True,
            "preserve_heading_formatting": False,
            "preserve_table_header_formatting": False
        },
        # MD → DOCX 格式设置
        "md_to_docx": {
            "formatting_mode": "apply",
            "heading_formatting_mode": "remove",
            "table_header_formatting_mode": "remove"
        },
        # Markdown格式语法配置（输出时使用）
        "syntax": {
            "bold": "asterisk",
            "italic": "asterisk",
            "strikethrough": "extended",
            "highlight": "extended",
            "superscript": "html",
            "subscript": "html",
            "unordered_list": "dash",
            "ordered_list": "restart",
            "indent_spaces": 2
        },
        # 分隔符/分页符/分隔线双向转换配置
        "horizontal_rule": {
            "enabled": True,
            # 文档转MD：Word分隔符 → MD分隔符
            # 可选值: "---", "***", "___", "ignore"
            "docx_to_md": {
                "page_break": "---",           # 分页符 → ---
                "section_break": "***",        # 分节符（所有类型统一）→ ***
                "horizontal_rule": "___"       # 分隔线（Horizontal Rule 1/2/3 样式）→ ___
            },
            # MD转文档：MD分隔符 → Word分隔符
            # 可选值: "page_break", "section_break", "horizontal_rule_1", "horizontal_rule_2", "horizontal_rule_3", "ignore"
            "md_to_docx": {
                "dash": "page_break",          # --- → 分页符
                "asterisk": "section_break",   # *** → 分节符（下一页）
                "underscore": "horizontal_rule_1"  # ___ → 分隔线（Horizontal Rule 1 样式）
            }
        }
    }
}

# 默认序号清理规则配置（简化版后备）
# 占位符定义在 heading_utils.py 中（使用 raw string）
DEFAULT_NUMBERING_PATTERNS_CONFIG = {
    "heading_numbering_clean": {
        "settings": {
            "order": [
                "chinese_unit_suffix",
                "chinese_unit_prefix",
                "circled_numbers",
                "bracket_number",
                "hierarchical",
                "number_separator",
                "legal_english",
                "letter_number"
            ]
        },
        "rules": {
            "chinese_unit_suffix": {
                "name": "中文【第X单位】格式",
                "description": "匹配：第X回/篇/册/卷/部/集/期/编/章/节/条/款/项/目 + 空格",
                "enabled": True,
                "is_system": True,
                "regex": "^{space}*第{space}*{num}+{space}*[回篇册卷部集期编章节条款项目]{space}+"
            },
            "chinese_unit_prefix": {
                "name": "中文【单位X】格式",
                "description": "匹配：卷/篇/册/部/集/期/编/章/节/回 + 数字 + 空格",
                "enabled": True,
                "is_system": True,
                "regex": "^{space}*[卷篇册部集期编章节回]{space}*{num}+{space}+"
            },
            "circled_numbers": {
                "name": "带圈数字",
                "description": "匹配：①-㊿、❶-⓴、⓵-⓾、㈠-㈩ 等带圈/特殊数字 + 可选分隔符",
                "enabled": True,
                "is_system": True,
                "regex": "^{space}*{circled}{sep}?{space}*"
            },
            "bracket_number": {
                "name": "括号数字",
                "description": "匹配：(1)、（一）、1)、一）等 + 可选分隔符",
                "enabled": True,
                "is_system": True,
                "regex": "^{space}*{bracket_open}?{num}+{bracket_close}{sep}?{space}*"
            },
            "hierarchical": {
                "name": "层级数字",
                "description": "匹配：1.1 / 1.1.1 / 1.1.1.1 等 + 空格",
                "enabled": True,
                "is_system": True,
                "regex": "^{space}*{num_arab}+(?:\\.{num_arab}+)+{sep}?{space}+"
            },
            "number_separator": {
                "name": "数字分隔符",
                "description": "匹配：1. / 2、/ 一、/ 二，等 + 可选空格",
                "enabled": True,
                "is_system": True,
                "regex": "^{space}*{num}+{sep}{space}*"
            },
            "legal_english": {
                "name": "英文章节格式",
                "description": "匹配：Chapter/Part/Volume/Book/Section + 数字/罗马数字",
                "enabled": False,
                "is_system": False,
                "regex": "^{space}*(?:Chapter|Part|Volume|Book|Section)\\s+[0-9IVXLCDMivxlcdm]+[\\s\\.:\\-]+"
            },
            "letter_number": {
                "name": "字母序号",
                "description": "匹配：A. / A) / (A) / a. / a) 等",
                "enabled": False,
                "is_system": False,
                "regex": "^{space}*(?:{bracket_open}{letter}{bracket_close}|{letter}{bracket_close}|{letter}{sep}){space}*"
            }
        }
    }
}

# 合并的默认配置常量（必须在所有子配置定义之后）
DEFAULT_CONFIG = {
    **DEFAULT_LOGGING_CONFIG,
    **DEFAULT_GUI_CONFIG,
    **DEFAULT_LINK_CONFIG,
    **DEFAULT_CONVERSION_DEFAULTS_CONFIG,
    **DEFAULT_OUTPUT_CONFIG,
    **DEFAULT_PROOFREAD_CONFIG,
    **DEFAULT_PROOFREAD_SYMBOLS,
    **DEFAULT_PROOFREAD_TYPOS,
    **DEFAULT_PROOFREAD_SENSITIVE,
    **DEFAULT_HEADING_NUMBERING_CONFIG,
    **DEFAULT_NUMBERING_PATTERNS_CONFIG,
    **DEFAULT_SOFTWARE_PRIORITY_CONFIG,
    **DEFAULT_CONVERSION_CONFIG,
    **DEFAULT_STYLE_CODE_CONFIG,
    **DEFAULT_STYLE_QUOTE_CONFIG,
    **DEFAULT_STYLE_FORMULA_CONFIG,
    **DEFAULT_STYLE_TABLE_CONFIG,
}

# 配置文件映射
CONFIG_FILES = {
    "logger_config": "logger_config.toml",
    "gui_config": "gui_config.toml",
    "link_config": "link_config.toml",
    "conversion_defaults": "conversion_defaults.toml",
    "output_config": "output_config.toml",
    "proofread_config": "proofread_config.toml",
    "proofread_symbols": "proofread_symbols.toml",
    "proofread_typos": "proofread_typos.toml",
    "proofread_sensitive": "proofread_sensitive.toml",
    "software_priority": "software_priority.toml",
    "heading_numbering_add": "heading_numbering_add.toml",
    "heading_numbering_clean": "heading_numbering_clean.toml",
    "conversion_config": "conversion_config.toml",
    "style_code": "style_code.toml",
    "style_quote": "style_quote.toml",
    "style_formula": "style_formula.toml",
    "style_table": "style_table.toml",
}
