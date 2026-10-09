# 开发指南

## 环境准备

### 推荐方案：使用 uv（快速、可靠）

安装生产环境固定的 `uv 0.12.0`，然后：

```bash
# 安装所有依赖（包括测试、lint、打包工具）
uv sync --frozen --all-extras

# 激活虚拟环境
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # macOS/Linux
```

### 安装器边界

DocWen 的 CI 与生产构建仍严格固定 `uv 0.12.0`；依赖与锁文件维护允许兼容的 `uv 0.12.x`，以便自动安全更新刷新仓库内 `uv.lock`，而不改变生产工具链。
不要使用 `pip install -e`：pip 不读取项目的 uv scoped dependency exclusion，会同时安装
两个互斥且覆盖相同 `cv2` 文件的 OpenCV 分发包。

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
QA 和源码启动都通过显式工程目录保存运行数据。普通 clone 首次准备时，在仓库外初始化一个工作区；已有维护者工作区不重复初始化：

```bash
python tools/workspace_root.py --init ../.workspace --repository .
python tools/qa.py --workspace-root ../.workspace
python tools/run_source.py --workspace-root ../.workspace gui
python tools/run_source.py --workspace-root ../.workspace cli -- --version
```

也可设置 `DOCWEN_WORKSPACE_ROOT` 为该目录的绝对路径。标准 `repos/docwen` 布局和关联 Git worktree 可自动定位既有工作区；普通 clone 使用显式参数。初始化只创建明确指定的新目录，或为既有治理根添加仓库登记，不读取维护者私有规划。

运行目录由工程入口租用并通过 `DOCWEN_RUNTIME_ROOT` 传给产品；正常安装默认使用应用缓存。产品不解析工程目录布局或 README。CI 保留显式 `--pytest-runtime-root`、`--own-pytest-runtime` 与报告目录的入口。

在本地一次跑完格式化校验、静态检查、类型检查与快速测试：

```bash
python tools/qa.py
```

仅跑快速测试（默认值）以外的全量测试：

```bash
python tools/qa.py --suite full
```

## 应用图标

`assets/icon.svg` 是唯一设计源。修改它以后，重新生成提交到仓库的 PNG 与 ICO 派生资源：

```bash
python scripts/maintenance/generate_app_icons.py
```

CI 和本地检查使用以下命令验证派生资源没有漂移：

```bash
python scripts/maintenance/generate_app_icons.py --check
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
python tools/qa.py --suite fast
```

## 批量 Import 替换

**禁止用 `sed -i`**。Windows + git-bash 下 `sed -i` 对路径分隔符、BOM、CRLF 不友好。统一用 Python 脚本：

```bash
rg "old_pattern" src/ tests/ --files-with-matches | xargs python -c "
import sys; [open(f,'r+').write(open(f).read().replace('old','new')) for f in sys.argv[1:]]
"
```

每次批改后立即 `pytest -q` 确认绿再继续下一批。

## 文件重命名

移动文件用 `git mv`（保留历史，`git log --follow` 可追溯），不要先 `git rm` 再新建。子目录整理优先用 `git mv` 一次性搬到位，再批量修 import。

## 类型检查（Pyright）
运行类型检查。仓库根目录的 Pyright 配置覆盖当前 `packages/` 工作区，并排除测试、构建产物、资源和配置数据：

```bash
pyright
```

也可使用仓库包装脚本执行同一门禁：

```bash
python tools/typecheck.py
```
