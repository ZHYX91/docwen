# 开发指南

[English](development.md)

DocWen 使用 Python 3.12，通过 uv workspace 管理 core、application、runtime、apps、plugins 和 bundle 包。以下命令在仓库根目录运行。

## 准备环境

CI 与生产构建固定 **uv 0.12.0**。依赖和锁文件维护可以使用兼容的 **uv 0.12.x**，不改变生产工具链。

```sh
uv sync --frozen --all-extras
```

Windows 使用 `.venv\Scripts\activate` 激活环境；macOS/Linux 使用 `source .venv/bin/activate`。不要使用 `pip install -e`：pip 不应用本工作区的 scoped dependency exclusion，可能将两个 OpenCV 分发包安装到同一套 `cv2` 文件。见[外部依赖](external-dependencies.md)。

## 源码运行与检查

QA 和源码启动入口把生成数据放入明确归属的工程目录。维护者已有工作区直接复用；普通 clone 首次准备时，在仓库外初始化新目录：

```sh
python tools/workspace_root.py --init ../docwen-workspace --repository .
python tools/qa.py --workspace-root ../docwen-workspace
python tools/run_source.py --workspace-root ../docwen-workspace gui
python tools/run_source.py --workspace-root ../docwen-workspace cli -- --version
```

也可将 `DOCWEN_WORKSPACE_ROOT` 设为该目录的绝对路径。标准 `repos/docwen` 布局及关联 worktree 能发现已登记工作区。初始化仅创建明确选择的目录或向既有根登记仓库，不依赖维护者私有计划。

入口租用运行目录，通过 `DOCWEN_RUNTIME_ROOT` 传给产品。正常安装的程序使用应用缓存。CI 可显式指定 `--pytest-runtime-root`、`--own-pytest-runtime` 和报告目录。

`python tools/qa.py` 执行格式、lint、类型检查和快速测试；确需全量测试时选择 `--suite full`。小改动优先运行所属检查，打包及交互验收范围见[测试指南](testing.md)。

```sh
ruff format --check .
ruff check .
python tools/typecheck.py
```

本地格式化使用 `ruff format .`；保留 `ruff check . --fix` 的结果前检查差异。Pyright 检查 `packages/`，排除测试、构建及资源数据。可选 pre-commit hooks 使用 `pre-commit install` 安装，用 `pre-commit run --all-files` 手动运行。

## 应用资源

`assets/icon.svg` 是应用图标的唯一设计源。修改后重新生成并检查已提交的 PNG 和 ICO 派生资源：

```sh
python scripts/maintenance/generate_app_icons.py
python scripts/maintenance/generate_app_icons.py --check
```

遵循 [GUI 视觉规范](specs/gui-behavior.md#visual-system--视觉规范)：基准尺寸使用共享 token，静态几何通过 `set_metric` 绑定，响应式文字布局使用字体度量。资源来自既有授权图标集或规范品牌源。MSIX 打包需要 Windows SDK 的 MakePri 与 MakeAppx，包内须包含 `resources.pri`。

## 修改边界

- 共享契约和纯逻辑放 core，编排放 application/runtime，格式行为放 plugins。
- 行为、配置、输出或发布要求改变时，同步所属测试和文档；兼容性决策明确记录，新增兼容层或第二状态源须有明确的公开契约。
- 保留输入和无关本地修改。批量文本修改应有明确范围并审阅，避免破坏 Windows 路径、编码或换行的原地 shell 替换。
- 重命名同步更新导入、引用和打包。让 Git 识别重命名，不为整理差异删除、重建无关内容。

另见[贡献说明](../CONTRIBUTING.md)、[架构](architecture.md)和[文档规范](maintenance/docs-style-guide.zh-CN.md)。
