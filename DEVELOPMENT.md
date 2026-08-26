# 开发指南

## 环境准备

### 推荐方案：使用 uv（快速、可靠）

安装固定版本 `uv 0.12.0`，然后：

```bash
# 安装所有依赖（包括测试、lint、打包工具）
uv sync --frozen --all-extras

# 激活虚拟环境
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # macOS/Linux
```

### 安装器边界

DocWen 0.9 的源码、测试与构建合同是 `uv 0.12.0` 加仓库内的 `uv.lock`。
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
python -m pytest
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
