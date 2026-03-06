# 开发指南

## 环境准备
- 推荐 Python 3.12
- 建议使用虚拟环境（conda/venv 均可）

安装项目（含测试与 lint 工具）：

```bash
python -m pip install -U pip
python -m pip install -e ".[test,lint]"
```

如果需要打包相关工具（含 pre-commit）：

```bash
python -m pip install -e ".[dev]"
```

## 代码规范（Ruff）
项目在 CI 中启用 ruff 门禁：

```bash
ruff format --check .
ruff check .
```

在本地自动修复/格式化：

```bash
ruff format .
ruff check . --fix
```

## 一键质量检查（推荐）
在本地一次跑完格式化校验、静态检查、类型检查与快速测试：

```bash
python tools/qa.py
```

仅跑快速测试（默认值）以外的全量测试：

```bash
python tools/qa.py --suite full
```

## 提交前自动检查（pre-commit）
首次使用需要安装 git hooks：

```bash
pre-commit install
```

手动对全仓执行一次：

```bash
pre-commit run --all-files
```

## 测试
运行测试：

```bash
python -m pytest
```

## 类型检查（Pyright）
运行类型检查（当前为渐进落地的低噪音配置，默认只检查 `src/docwen`）：

```bash
pyright
```

按仓库约定一次跑 core + GUI + optional：

```bash
python tools/typecheck.py
```
