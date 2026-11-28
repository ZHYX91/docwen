# 概览

[English](overview.md)

DocWen 是本地优先的桌面与命令行文档转换工具。GUI 与 CLI 复用同一套应用、运行时和插件组合；常规转换不依赖网络服务。

## 范围

- 文档：Markdown、DOCX、DOC、ODT、RTF、WPS 及相关 PDF 导出路线。
- 表格：XLSX、XLS、ODS、CSV、TSV、ET。
- 版式与图片：PDF、XPS、OFD 和常见位图格式。
- 导入：按 route 声明支持 HTML、MHTML、EPUB、ENEX、PPTX、PPT 到 Markdown。
- OCR、校对、编号、模板渲染、批处理和聚合操作。

精确支持面以[能力矩阵](capabilities.md)和[路线与操作](specs/routes-and-actions.md)为准。不支持或尽力而为的行为必须由所属 route 如实反馈。

## 安全

受支持的 GUI 和 CLI 组合入口会在 DocWen 的 Python 主进程整个生命周期内启用唯一的 CPython 审计守卫。它阻止进程内 Python 依赖执行全部 DNS/名称解析以及 AF_INET/AF_INET6 的 `bind`、`connect`、`connect_ex`、`sendto`、`sendmsg` 操作，同时保留 Windows 命名管道与 Unix 域套接字。绕过 CPython 审计事件的原生代码、单独启动的进程、外部 Office/WPS/LibreOffice 以及专用 Office helper 不在该边界内。该守卫用于纵深防御依赖意外联网，不是操作系统级沙箱。源文件按不可变输入处理；最终输出通过 Runtime Finalizer 完成命名冲突和路径边界检查后落盘。
继承或在守卫启用前已经连接的 IP 套接字可执行不产生 `socket.send` 审计事件的普通载荷写入，因此该情形明确不在保证范围内；受支持的打包入口会在产品和 Qt 模块导入前启用守卫，也不会有意建立这类套接字。
