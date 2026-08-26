# Development / 开发

DocWen requires Python 3.12. The repository uses a uv workspace containing core, application, runtime, apps, plugins and bundle packages.

DocWen 要求 Python 3.12。仓库使用 uv workspace 管理 core、application、runtime、apps、plugins 和 bundle 包。

## Setup / 环境

```powershell
uv sync --frozen --all-extras  # requires uv 0.12.0
.\.venv\Scripts\python.exe -m pytest
docwen --help
docwen-gui
```

## Application icon / 应用图标

`assets/icon.svg` is the sole design source. Regenerate the committed PNG and ICO derivatives after
changing it; use `--check` to verify that they have not drifted from the SVG.

`assets/icon.svg` 是唯一设计源。修改后重新生成提交到仓库的 PNG 与 ICO 派生资源；使用
`--check` 验证派生资源与 SVG 保持一致。

```powershell
.\.venv\Scripts\python.exe scripts\maintenance\generate_app_icons.py
.\.venv\Scripts\python.exe scripts\maintenance\generate_app_icons.py --check
```

## Contribution rules / 开发规则

- Put shared contracts and pure logic in core.
- Keep orchestration in application/runtime and format behavior in plugins.
- Add or update the owning tests with every behavioral change.
- Update current documentation when routes, configuration, output semantics, UI behavior or release gates change.
- Do not add compatibility shims or second state sources without an explicit public contract.
- Preserve source files and unrelated user changes during tests and maintenance.

See [Architecture](architecture.md), [Testing](testing.md) and [Documentation style](maintenance/docs-style-guide.md).
