# Changelog / 更新日志

- 本文档只记录版本演进和对外可见变更，不作为当前功能边界、架构约束或规范条款的事实源。
- 长期有效的项目说明以 README 体系为准，设计与维护边界以 `docs/architecture.md` 为准，强约束规则以 `docs/specs/` 为准。

## 0.9.0 (Unreleased)

### 破坏性变化（中文）

- CLI JSON 展示协议升级为 protocol 3；它独立于新的 Machine Protocol v1。旧命令、别名、flag 和 protocol 2 输出不再兼容。
  Markdown 目标只接受精确标识 `md`；`markdown` 不再重写。
- 顶层命令统一为 `info`、`inspect`、`doctor`、`resources`、`schema`、`convert`、`validate`、
  `number`、`merge`、`split`、`batch`、`gui`、`config`；其中编号入口为 `number markdown`，
  GUI 控制入口为 `gui open|activate|status`。
- JSON 输出采用严格 typed envelope、稳定错误码和退出码；stdout 只输出机器结果，诊断进入 stderr。
- 新增 `docwen.machine.v1`、`docwen.artifact_bundle.v2` 与 `docwen.proofread_report.v2` 正式契约；不接受旧 Bundle 或旧校对报告形状。
- Windows 与 Ubuntu 正式附件改用双干净构建逐字节比较、固定资产清单、证明和不可变 Release 读回校验；只有明确的 HTTP 404 才允许创建 Release，并在发布前后重验精确数字标签。
- GUI control 收敛为本机受控 runtime/control 边界，删除旧文件 IPC、PID 探测和双入口实现。
- CLI、GUI 与应用服务共享同一执行链；拆分请求构造、执行、展示和路径策略所有权。
- 明确移除 0.8 的 Python 导入面：`OfficeSoftwareNotFoundError`、`docwen.cli.api` 和
  `docwen.converter.docx.to_md.shared.docx_utils` 不再提供兼容 shim。Python 集成应迁移到当前
  `docwen_cli`/公开 protocol 3 边界与 typed outcomes（`dependency_missing`、`conversion_failed`、
  `operation_cancelled`），不再直接导入旧内部模块。
- Windows 打包 CLI 与源码 CLI 使用同一 fixture 验证；DocWen Assistant 2.0 和 OpenClaw 2.0
  仍须在最终 0.9 候选冻结后分别完成组合验证，当前未发布章节不把旧候选证据当作最终结论。
- 脚注/尾注 Markdown 合同统一为 `[^id]`、`[^footnote:id]` 与 `[^endnote:id]`；已废弃的
  `[^endnote-id]` 输入会被拒绝。两类注释按首次引用独立编号，缺失、重复或规范化冲突安全失败。

从 `0.8.x` 升级时，请更新所有 CLI 脚本和集成；不存在旧命令兼容层。正式附件、支持平台、
代码签名和已知边界以 `0.9.0` Release notes 为准。

### 其他变化（中文）

- 采用新的 DocWen 品牌图标：双文档与双向转换箭头、小型 60° 镜像折角；SVG 是唯一设计源，
  Windows ICO 与跨平台 PNG 由仓库脚本确定性生成并接受一致性检查。

### Breaking changes (English)

- The CLI JSON presentation protocol is now protocol 3 and is distinct from Machine Protocol v1. Legacy commands, aliases, flags, and protocol 2
  output are not supported. The Markdown target accepts only the exact `md` identifier; `markdown`
  is not rewritten.
- The canonical top-level commands are `info`, `inspect`, `doctor`, `resources`, `schema`, `convert`,
  `validate`, `number`, `merge`, `split`, `batch`, `gui`, and `config`. Numbering is exposed through
  `number markdown`; GUI control is exposed through `gui open|activate|status`.
- JSON mode uses a strict typed envelope with stable error and exit codes. Stdout is reserved for the
  machine result; diagnostics go to stderr.
- The formal contracts now include `docwen.machine.v1`, `docwen.artifact_bundle.v2`, and
  `docwen.proofread_report.v2`; old Bundle and proofread-report shapes are rejected.
- Windows and Ubuntu release assets use two byte-compared clean builds, a fixed asset inventory,
  provenance attestations, and immutable Release readback. Only an explicit HTTP 404 permits Release
  creation, and the exact numeric tag is revalidated immediately before and after publication.
- GUI control now uses one bounded local runtime/control boundary. Legacy file IPC, PID probing, and
  dual entry paths were removed.
- CLI, GUI, and application services share one execution path with explicit ownership for request
  construction, execution, presentation, and path policy.
- The 0.8 Python import surfaces `OfficeSoftwareNotFoundError`, `docwen.cli.api`, and
  `docwen.converter.docx.to_md.shared.docx_utils` were removed without compatibility shims. Python
  integrations must migrate to the current `docwen_cli`/public protocol 3 boundary and typed outcomes
  (`dependency_missing`, `conversion_failed`, and `operation_cancelled`) instead of importing legacy
  internal modules.
- Source and packaged CLIs use the same contract fixtures. DocWen Assistant 2.0 and OpenClaw 2.0
  still require separate combination validation after the final 0.9 candidate is frozen; evidence from
  superseded local candidates is not a final result.
- Markdown notes now use `[^id]` or `[^footnote:id]` for footnotes and `[^endnote:id]` for endnotes;
  the retired `[^endnote-id]` input is rejected. Each note domain numbers by first reference and
  fails closed on missing, duplicate, or normalized-collision definitions.

Update all CLI scripts and integrations when upgrading from `0.8.x`; no legacy command compatibility
layer is provided. The final artifacts, supported platforms, code-signing status, and known boundaries
will be listed in the `0.9.0` Release notes.

### Other changes (English)

- Adopted a new DocWen brand icon with two documents, bidirectional conversion arrows, and a small
  mirrored 60-degree fold. The SVG is the sole design source; the Windows ICO and cross-platform PNG
  are generated deterministically and checked for drift.

## v0.8.5 (2026-03-20)

### 更新日志（中文）

- 新增"一键还原设置"功能：设置对话框新增按选项卡或整体恢复默认值的按钮
- CLI 新增 `settings reset` 子命令，支持 `--tab`、`--yes`、`--json` 参数
- 修复 Windows CLI 中文输出乱码问题

### Changelog (English)

- Added one-click settings reset: new buttons in the settings dialog to reset the current tab or all settings to defaults.
- Added CLI `settings reset` subcommand with `--tab`, `--yes`, and `--json` support.
- Fixed garbled output in Windows CLI (UTF-8 encoding).

## v0.8.4 (2026-03-19)

### 更新日志（中文）

- 修复校对模块 `has_non_text_content()` 的 xpath 调用兼容性，避免所有段落误走降级路径
- 新增批注锚点分析报告工具（`anchor_report`），支持跨段/未闭合锚点检测与多语言脱敏
- 新增"标题+正文合并模式"设置，支持三种行为：有标点才合并（默认）、紧邻正文就合并、永不合并
- 修正 README 中对标题+正文合并行为的错误描述
- macOS 跨平台支持：LibreOffice 路径自动查找（适配 macOS `.app` 安装）、Release CI 补充 python-tk 安装、README 补充 macOS/Linux GUI 使用说明
- 公共 API 分层重构：各域新增 `api.py` 入口层，tests/tools 统一通过公开 API 访问，消除 reportPrivateUsage 告警
- ConfigManager 去除单例代理，改为普通类 + 模块级实例，提取 `deep_merge_dicts()` 为模块级函数
- `tools/qa.py` 新增私有符号边界扫描（AST 级别，检测 tests/tools 是否直接访问 `_xxx`）
- 新增 `docs/specs/public-api-boundaries.md` 规范文档

### Changelog (English)

- Fixed proofreading `has_non_text_content()` xpath call compatibility with python-docx, preventing all paragraphs from falling through to the degraded path.
- Added comment anchor analysis report tool (`anchor_report`) with cross-paragraph/unclosed anchor detection and multilingual redaction.
- Added `heading_merge_mode` setting for MD-to-DOCX conversion with three options: merge only when heading ends with punctuation (default), always merge when adjacent, or never merge.
- Fixed incorrect README description of heading+body merge behavior.
- macOS cross-platform: auto-locate LibreOffice soffice in macOS .app paths; fixed python-tk in CI; updated READMEs with macOS and Linux GUI notes.
- Public API layer refactor: added `api.py` entry point per domain; tests/tools now import only public APIs, eliminating reportPrivateUsage noise.
- ConfigManager: removed singleton proxy, converted to plain class with module-level instance; extracted `deep_merge_dicts()` as module-level function.
- `tools/qa.py`: added AST-based private symbol boundary scan for tests/tools directories.
- Added `docs/specs/public-api-boundaries.md` coding convention document.

## v0.8.3 (2026-03-15)

### 更新日志（中文）

- 修复 Release CI：Windows CLI 部署目录生成；Linux tar 校验 Broken pipe 假失败
- 构建去重：减少 Windows GUI/CLI 构建产物重复，降低包体积
- 日志系统增强：文件日志创建失败时自动重试并回退到临时目录
- 修复 Linux/macOS 窗口图标（改用 `iconphoto` + PNG）
- Release CI 新增 Linux GUI 发布产物，构建结构与 Windows 对称

### Changelog (English)

- Fixed Release CI: Windows CLI deploy dir generation; Linux tar verification false failure (Broken pipe).
- Build deduplication: reduced duplicate resources in Windows GUI/CLI outputs and shrank release size.
- Logging improvements: file logging retries on failure and falls back to a temp directory.
- Fixed Linux/macOS window icon (switched to `iconphoto` + PNG).
- Added Linux GUI release artifact; Linux build structure now mirrors Windows.

## v0.8.2 (2026-03-13)

### 更新日志（中文）

- 新增字段注册表与公文字段处理器，支持 YAML 处理器/占位符规则/特殊处理器的注册式管理
- 补充 OpenClaw 配套集成说明（CLI wrapper / Plugin + Skill）
- 新增 GitHub Actions CI/CD（release 自动构建 + tests lint/typecheck/pytest）
- 版本号迁移至纯 SemVer 格式

### Changelog (English)

- Added field registry and gongwen field processors with registration-based management for YAML processors, placeholder rules and special handlers.
- Added OpenClaw companion integration docs (CLI wrapper / Plugin + Skill).
- Added GitHub Actions CI/CD (release asset builds + tests with lint/typecheck/pytest).
- Migrated version numbering to pure SemVer format.

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

## v0.4.1 (2024-12-05)

### 更新日志（中文）

- 重构命令行界面，提升用户体验
- 添加对非公文文档转换的支持
- 实现更多选项配置化

### Changelog (English)

- Refactored CLI to improve user experience.
- Added support for more document types.
- Implemented more configurable options.
