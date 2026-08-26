# Overview / 概览

DocWen is a local-first desktop and command-line document conversion toolkit. It provides GUI and CLI entry points over one application/runtime/plugin composition and does not require a network service for normal conversion.

DocWen 是本地优先的桌面与命令行文档转换工具。GUI 与 CLI 复用同一套应用、运行时和插件组合；常规转换不依赖网络服务。

## Scope / 范围

- Documents: Markdown, DOCX, DOC, ODT, RTF, WPS and related PDF export routes.
- Spreadsheets: XLSX, XLS, ODS, CSV, TSV and ET routes.
- Layout and images: PDF, XPS, OFD and common raster formats.
- Imports: HTML, MHTML, EPUB, ENEX, PPTX and PPT to Markdown where declared.
- OCR, proofreading, numbering, template rendering, batch conversion and aggregate actions.

- 文档：Markdown、DOCX、DOC、ODT、RTF、WPS 及相关 PDF 导出路线。
- 表格：XLSX、XLS、ODS、CSV、TSV、ET。
- 版式与图片：PDF、XPS、OFD 和常见位图格式。
- 导入：按 route 声明支持 HTML、MHTML、EPUB、ENEX、PPTX、PPT 到 Markdown。
- OCR、校对、编号、模板渲染、批处理和聚合操作。

The exact supported surface is defined by [Capabilities](capabilities.md) and [Routes and actions](specs/routes-and-actions.md). Unsupported or best-effort behavior must be reported truthfully by the owning route.

精确支持面以[能力矩阵](capabilities.md)和[路线与操作](specs/routes-and-actions.md)为准。不支持或尽力而为的行为必须由所属 route 如实反馈。

## Safety / 安全

Supported GUI and CLI composition roots activate one CPython audit guard for the full lifetime of the main DocWen Python process. It denies all DNS/name resolution and AF_INET/AF_INET6 `bind`, `connect`, `connect_ex`, `sendto`, and `sendmsg` operations used by in-process Python dependencies while leaving Windows named pipes and Unix-domain sockets available. Native code that bypasses CPython audit events, separately launched processes, external Office/WPS/LibreOffice applications and the dedicated Office helper are outside this boundary. The guard is defence in depth against accidental dependency egress, not an operating-system sandbox. Source files are treated as immutable inputs; final outputs are placed through the runtime finalizer with collision and containment checks.
An inherited or already-connected IP socket can perform an ordinary payload write without a `socket.send` audit event, so that case is explicitly outside the guarantee; supported packaged entries activate before product/Qt imports and do not intentionally establish one.

受支持的 GUI 和 CLI 组合入口会在 DocWen 的 Python 主进程整个生命周期内启用唯一的 CPython 审计守卫。它阻止进程内 Python 依赖执行全部 DNS/名称解析以及 AF_INET/AF_INET6 的 `bind`、`connect`、`connect_ex`、`sendto`、`sendmsg` 操作，同时保留 Windows 命名管道与 Unix 域套接字。绕过 CPython 审计事件的原生代码、单独启动的进程、外部 Office/WPS/LibreOffice 以及专用 Office helper 不在该边界内。该守卫用于纵深防御依赖意外联网，不是操作系统级沙箱。源文件按不可变输入处理；最终输出通过 Runtime Finalizer 完成命名冲突和路径边界检查后落盘。
继承或在守卫启用前已经连接的 IP 套接字可执行不产生 `socket.send` 审计事件的普通载荷写入，因此该情形明确不在保证范围内；受支持的打包入口会在产品和 Qt 模块导入前启用守卫，也不会有意建立这类套接字。
