# DocWen Documentation / DocWen 文档

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

DocWen documentation describes the current product, its supported contracts, and the checks required to maintain it. Historical implementation and acceptance records are available from Git history rather than the working tree.

DocWen 文档只描述当前产品、受支持契约和维护门禁。实现与验收历史由 Git 历史追溯，不在工作树中保留阶段报告。

## User guides / 用户手册

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md)

## Product and development / 产品与开发

- [Overview / 概览](overview.md)
- [Capabilities / 能力矩阵](capabilities.md)
- [CLI](cli.md)
- [Configuration / 配置](configuration.md)
- [Architecture / 架构](architecture.md)
- [Runtime artifacts / 运行产物](runtime-artifacts.md)
- [External dependencies / 外部依赖](external-dependencies.md)
- [Testing / 测试](testing.md)
- [Packaging / 打包](packaging.md)
- [Development / 开发](development.md)

## Specifications / 规格

- [Machine Protocol v1 and Artifact Bundle v2](specs/machine-protocol-v1.md)
- [Markdown document-node output](specs/document-node-output.md)
- [Routes and actions](specs/routes-and-actions.md)
- [Plugin manifest](specs/plugin-manifest.md)
- [JSON contracts](specs/json-contracts.md)
- [Golden regression suite](specs/golden-regression-suite.md)
- [GUI behavior](specs/gui-behavior.md)
- [Markdown compatibility](specs/markdown-compatibility.md)
- [Templates and styles](specs/templates-and-styles.md)
- [Physical-page OCR and artifact relations](specs/physical-page-ocr.md)
- [Resolved structured numbering and export plan](specs/structured-numbering-phases.md)
- [Public API boundaries](specs/public-api-boundaries.md)

The resolved-plan capability rejects `remove_numbering`, `add_numbering`, and `numbering_scheme`, never parses a
manual Heading prefix, and is guarded by an explicit negative capability gate.

## Maintenance / 维护

- [Troubleshooting / 故障排查](maintenance/troubleshooting.md)
- [Documentation style / 文档规范](maintenance/docs-style-guide.md)
