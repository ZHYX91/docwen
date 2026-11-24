# Development / 开发

DocWen requires Python 3.12. The repository uses a uv workspace containing core, application, runtime, apps, plugins and bundle packages.

DocWen 要求 Python 3.12。仓库使用 uv workspace 管理 core、application、runtime、apps、plugins 和 bundle 包。

## Setup / 环境

```powershell
uv sync --frozen --all-extras  # CI/release pin uv 0.12.0
.\.venv\Scripts\python.exe tools\qa.py --suite fast
docwen --help
docwen-gui
```

For dependency and lockfile updates, use a uv version allowed by `tool.uv.required-version` in [pyproject.toml](../pyproject.toml).

## Application icon / 应用图标

`assets/icon.svg` is the sole design source. Regenerate the committed PNG and ICO derivatives after
changing it; use `--check` to verify that they have not drifted from the SVG.

`assets/icon.svg` 是唯一设计源。修改后重新生成提交到仓库的 PNG 与 ICO 派生资源；使用
`--check` 验证派生资源与 SVG 保持一致。

```powershell
.\.venv\Scripts\python.exe scripts\maintenance\generate_app_icons.py
.\.venv\Scripts\python.exe scripts\maintenance\generate_app_icons.py --check
```

## UI dimensions and resources / 界面尺寸与资源

See [GUI behavior and visual ownership](specs/gui-behavior.md#visual-system--视觉规范). Add baseline dimensions to shared tokens, bind static geometry through `set_metric`, and use font metrics for responsive text layout. Keep new assets in the existing licensed icon collection or the canonical brand source. MSIX builds require the Windows SDK MakePri tool beside MakeAppx; the package must include `resources.pri`.

## Contribution rules / 开发规则

- Put shared contracts, parsing and document semantics in core. Declare its direct library dependencies
  in `packages/core/pyproject.toml`; shared DOCX semantics use `python-docx` and `lxml` without depending
  on plugin installation. Keep conversion routes in plugins.
- Keep orchestration in application/runtime and format behavior in plugins.
- Add or update the owning tests with every behavioral change.
- Update current documentation when routes, configuration, output semantics, UI behavior or release gates change.
- Do not add compatibility shims or second state sources without an explicit public contract.
- Preserve source files and unrelated user changes during tests and maintenance.

共享契约、解析与文档语义属于 Core；其直接库依赖在 `packages/core/pyproject.toml` 声明。
共享 DOCX 语义直接使用 `python-docx` 和 `lxml`，不依赖插件安装补齐；转换路线仍属于插件。

See [Architecture](architecture.md), [Testing](testing.md) and [Documentation style](maintenance/docs-style-guide.md).
