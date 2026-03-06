"""
格式转换模块

提供各种文件格式之间的转换功能。

子模块：
- common: 公共模块（软件检测、容错框架）
- document: 文档格式转换（DOC/DOCX/RTF/ODT）
- spreadsheet: 表格格式转换（XLS/XLSX/ODS）
- layout: 版式文件转换（PDF/OFD/XPS/CAJ）
- image: 图片格式转换
- pdf_export: PDF 导出

架构说明：
    本模块遵循三层架构设计：
    1. strategies（策略层）→ 协调转换流程
    2. formats（格式转换层）→ 具体格式转换实现（本模块）
    3. docx2md/md2docx 等（核心转换层）→ MD 互转

使用方式:
    # 推荐：从子模块导入
    from docwen.converter.formats.document import office_to_docx
    from docwen.converter.formats.layout import ofd_to_pdf
"""
