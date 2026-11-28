# DocWen documentation

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

[简体中文](README.zh-CN.md)

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.vi-VN.md) · [العربية](https://github.com/ZHYX91/docwen/blob/main/docs/i18n/README.ar.md)

Use this index for current behavior, supported contracts and maintenance instructions.

## Use DocWen

- [Product guide](../README.md)
- [Overview](overview.md)
- [Capabilities](capabilities.md)
- [CLI](cli.md)
- [Configuration](configuration.md)
- [Heading numbering and markers](guides/heading-numbering.md)
- [Merge spreadsheets](guides/table-merge.md)
- [macOS first launch](guides/macos.md)
- [Mermaid images](mermaid.md)
- [Troubleshooting](maintenance/troubleshooting.md)

## Develop and release

- [Development](development.md)
- [Architecture](architecture.md)
- [Runtime artifacts](runtime-artifacts.md)
- [External dependencies](external-dependencies.md)
- [Testing](testing.md)
- [Packaging and release](packaging.md)
- [Documentation style](maintenance/docs-style-guide.md)
- [Changelog](../CHANGELOG.md)
- [Security reporting](../SECURITY.md)

## Specifications

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

## Design decisions

- [Each line as a Word paragraph](architecture/decisions/each-line-paragraph.md)
- [Raw Markdown HTML in DOCX exports](architecture/decisions/html-block-text-projection.md)
- [OFD/XPS normalization for PDF actions](architecture/decisions/layout-pdf-action-composition.md)
- [Office bridge stage integrity](architecture/decisions/office-bridge-stage-integrity.md)
- [Markdown document proofreading before legacy Office output](architecture/decisions/proofread-legacy-document-pipeline.md)
- [Settings information architecture](architecture/decisions/settings-information-architecture.md)
- [Template file location](architecture/decisions/template-location-affordance.md)
