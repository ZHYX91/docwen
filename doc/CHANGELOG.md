# Changelog / 更新日志

## v0.8.1 (2026-03-06)

### 更新日志（中文）

- 扩展 service 层，新增批量并发、错误注册表与统一转换请求模型
- 新增 CLI JSON 输出规范与 doctor 环境诊断命令
- 重构策略加载为显式注册表 + 按需加载
- 优化高 DPI 适配与 GUI 导出设置
- 移除试用期检查，统一结构化错误码体系
- 修复图片、版式文件、MD↔DOCX/XLSX 等大量转换 bug
- 大幅提升单元测试覆盖率（新增 100+ 测试文件）

### Changelog (English)

- Expanded service layer with batch concurrency, error registry and unified conversion request models.
- Added CLI JSON output schema and doctor diagnostics command.
- Refactored strategy loading to explicit registry + on-demand imports.
- Improved high-DPI adaptation and GUI export settings.
- Removed trial expiration check; unified structured error code system.
- Fixed numerous conversion bugs across image, layout, MD↔DOCX/XLSX, etc.
- Significantly expanded unit test coverage (100+ new test files).

## v0.7.0 (2025-02-06)

### 更新日志（中文）

- 修复校对规则与跳过逻辑，优化校对选项联动
- 增强表格样式注入与对齐，新增图片样式支持
- 优化 Emoji 和换行符处理
- 完善多语言翻译文件，新增 locale 验证器
- 修复模板选择与图片路径查找
- 优化 GUI 设置面板与界面交互
- 大幅提升单元测试覆盖率
- README 重构并添加界面截图

### Changelog (English)

- Fixed proofreading rules and skip logic, improved option linkage.
- Enhanced table style injection and alignment, added image style support.
- Improved Emoji and line break handling.
- Completed multilingual translation files, added locale validator.
- Fixed template selection and image path lookup.
- Optimized GUI settings panel and interface interaction.
- Significantly improved unit test coverage.
- Restructured README and added UI screenshots.

## v0.6.0 (2025-01-20)

### 更新日志（中文）

- 完整的国际化支持（GUI 和 CLI 支持 11 种语言）
- 使用 RapidOCR 替代 PaddleOCR，提升兼容性
- 新增多语言 Word/Excel 模板
- 模板样式自动检测与注入
- 其他优化和修复

### Changelog (English)

- Full internationalization support (GUI and CLI support 11 languages).
- Replaced PaddleOCR with RapidOCR for better compatibility.
- Added multilingual Word/Excel templates.
- Automatic template style detection and injection.
- Other optimizations and fixes.

## v0.5.1 (2025-01-01)

### 更新日志（中文）

- 新增数学公式双向转换（Word OMML ↔ Markdown LaTeX）
- 新增脚注/尾注双向转换
- 新增代码、引用等字符和段落样式
- 增强列表处理（多级嵌套、自动编号）
- 增强表格功能（样式检测/注入、三线表等）
- 优化小标题序号清理和添加
- 改进界面交互和设置联动

### Changelog (English)

- Added bidirectional math formula conversion (Word OMML ↔ Markdown LaTeX).
- Added bidirectional footnote/endnote conversion.
- Added character and paragraph styles for code, quotes, etc.
- Enhanced list processing (multi-level nesting, automatic numbering).
- Enhanced table functions (style detection/injection, three-line tables, etc.).
- Optimized cleaning and adding of subheading numbers.
- Improved interface interaction and settings linkage.

## v0.4.1 (2025-12-05)

### 更新日志（中文）

- 重构命令行界面，提升用户体验
- 添加对非公文文档转换的支持
- 实现更多选项配置化

### Changelog (English)

- Refactored CLI to improve user experience.
- Added support for more document types.
- Implemented more configurable options.
