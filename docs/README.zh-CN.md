# DocWen 文档

[English](README.md)

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.vi-VN.md) · [العربية](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.ar.md)

本索引收录当前行为、支持范围和维护说明。

## 使用指南

- [产品手册](i18n/README.zh-CN.md)
- [概览](overview.zh-CN.md)
- [能力矩阵](capabilities.md)
- [命令行](cli.md)
- [配置](configuration.md)
- [表格汇总](guides/table-merge.zh-CN.md)
- [macOS 首次打开](guides/macos.zh-CN.md)
- [Mermaid 图片](mermaid.md)
- [故障排查](maintenance/troubleshooting.md)

## 开发与发布

- [开发指南](development.zh-CN.md)
- [架构](architecture.md)
- [运行产物](runtime-artifacts.md)
- [外部依赖](external-dependencies.md)
- [测试](testing.md)
- [打包与发布](packaging.md)
- [文档规范](maintenance/docs-style-guide.zh-CN.md)
- [版本记录](../CHANGELOG.md)
- [漏洞报告](../SECURITY.md)

## 规格

- [Rejected artifact cleanup](specs/artifact-discard.md)
- [Delimited spreadsheet fidelity](specs/delimited-spreadsheets.md)
- [Conversion result-directory output](specs/document-node-output.md)
- [DOCX proofreading comments](specs/docx-proofread-comments.md)
- [Golden regression suite / Golden 回归套件](specs/golden-regression-suite.md)
- [GUI behavior / GUI 行为](specs/gui-behavior.md)
- [JSON contracts / JSON 契约](specs/json-contracts.md)
- [Machine Protocol v2 and Artifact Bundle v3](specs/machine-protocol-v2.md)
- [Markdown compatibility / Markdown 兼容](specs/markdown-compatibility.md)
- [Offline OCR table recognition](specs/ocr-tables.md)
- [Physical-page OCR and artifact relations](specs/physical-page-ocr.md)
- [Plugin manifest](specs/plugin-manifest.md)
- [Public API boundaries](specs/public-api-boundaries.md)
- [Routes and actions](specs/routes-and-actions.md)
- [Resolved structured numbering and export plan](specs/structured-numbering-phases.md)
- [Display and task languages](specs/task-language.md)
- [Templates and styles](specs/templates-and-styles.md)
- [User profile paths](specs/user-profile.md)
- [Web archives and HTML image references](specs/web-archives.md)

## 设计决策

- [Each line as a Word paragraph](architecture/decisions/each-line-paragraph.md)
- [Raw Markdown HTML in DOCX exports](architecture/decisions/html-block-text-projection.md)
- [OFD/XPS normalization for PDF actions](architecture/decisions/layout-pdf-action-composition.md)
- [Office bridge stage integrity](architecture/decisions/office-bridge-stage-integrity.md)
- [Markdown document proofreading before legacy Office output](architecture/decisions/proofread-legacy-document-pipeline.md)
- [Settings information architecture](architecture/decisions/settings-information-architecture.md)
- [Template file location](architecture/decisions/template-location-affordance.md)
