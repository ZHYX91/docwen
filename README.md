# 公文转换器 (Gongwen Converter)

公文图表格式转换软件 - 支持 Word/Markdown/Excel 互转，完全本地运行，数据安全可靠。

## ✨ 核心功能

- **📄 Word 转 Markdown** - 智能解析公文格式，提取关键信息。输出的MD文件带有 YAML front matter，可以用 Obsidian 或 Typora 管理。
- **📝 Markdown 转 Word** - 基于模板生成格式规范的公文。
- **📊 表格文件处理** - 支持 Excel/CSV 转 Markdown，以及 Markdown 转 Excel 报表。
- **🔄 多格式自动转换** - 自动处理 DOC、WPS、XLS、ET 等旧格式，转为标准格式后处理。
- **🔍 智能校对** - 自动检查错别字、标点配对、符号混用和敏感词问题。
- **💻 双模式操作** - 图形界面 + 命令行界面。
- **🔒 本地运行** - 完全离线，数据安全，内置网络隔离机制。
- **🔗 单实例运行** - 智能管理程序实例，支持与 Obsidian 插件无缝集成。

## 🚀 快速开始

### 安装步骤

1. **克隆或下载项目**

   ```bash
   git clone https://github.com/ZHYX91/gongwen-converter.git
   cd gongwen-converter
   ```

2. **安装依赖**

   ```bash
   # 从 pyproject.toml 安装所有依赖
   pip install .
   
   # 或使用开发模式（可编辑安装）
   pip install -e .
   
   # 如需构建工具，安装开发依赖
   pip install -e ".[dev]"
   ```

3. **配置OCR模型（可选）**

   如需使用OCR文字识别功能，请下载PaddleOCR模型文件并放置到 `models/paddleocr/` 目录：
   
   ```
   models/paddleocr/
   ├── det/           # 文本检测模型
   ├── rec/           # 文本识别模型
   ├── cls/           # 文本方向分类模型
   └── doc_ori/       # 文档方向检测模型
   ```
   
   模型下载地址：https://github.com/PaddlePaddle/PaddleOCR/blob/main/doc/doc_ch/models_list.md
   
   **注意**：程序会自动检测OCR模型，如果模型不存在，OCR功能将不可用，但不影响其他转换功能。

4. **启动程序**

   ```bash
   # 图形界面
   python src/gui_run.py
   
   # 命令行界面
   python src/cli_run.py
   ```

### 快速体验

1. **准备一个 Markdown 文件**：

   ```markdown
   ---
   公文标题: 测试文档
   发文字号: 测试〔2024〕1号
   主送机关: 测试部门
   成文日期: 2024年10月26日
   ---
   
   ## 测试标题
   
   这是测试正文内容。
   ```

2. **拖拽转换**：
   - 启动 GUI 程序
   - 将 .md 文件拖入窗口
   - 选择模板
   - 点击"转换为 DOCX"

3. **获取结果**：
   - 在相同目录下生成格式规范的 Word 文档

## 📖 详细使用指南

### Word 转 Markdown

1. 将 .docx 文件拖入程序窗口
2. 程序自动分析文档结构
3. 生成包含 YAML 元数据的 .md 文件

**支持的格式**：
- `.docx` - 标准 Word 文档
- `.doc` - 自动转换为 DOCX 后处理
- `.wps` - WPS 文档自动转换

### Markdown 转 Word

1. 准备包含 YAML 头部的 .md 文件
2. 拖入程序窗口并选择对应的 Word 模板
3. 程序自动填充模板并生成文档

**注意**：如果公文中有小标题和正文混合的段落，在MD文件中需要保持严格换行。

### 表格文件处理

1. **Excel/CSV 转 Markdown**：拖入 .xlsx 或 .csv 文件，自动转换为 Markdown 表格
2. **Markdown 转 Excel**：准备好 MD 文件，选择 Excel 模板进行转换

**支持的格式**：
- `.xlsx` - 标准 Excel 文档
- `.xls` - 自动转换为 XLSX 后处理
- `.et` - WPS 表格自动转换
- `.csv` - CSV 文本表格

### 智能校对功能

程序提供四种校对规则：

1. **标点配对检查** - 检测括号、引号等成对标点是否匹配
2. **符号校对** - 检测中英文标点混用问题
3. **错别字检查** - 基于自定义词典检查常见错别字
4. **敏感词匹配** - 检测文档中的敏感词使用

使用方法：
1. 将需要校对的 Word 文档拖入程序
2. 勾选需要的校对规则
3. 点击"智能校对"按钮
4. 校对结果以批注形式显示在文档中

### 支持的公文元素

- 公文标题、发文字号、密级
- 主送机关、抄送机关、附件
- 签发人、成文日期、印发机关
- 以及其他标准公文要素

## ⚙️ 配置说明

程序行为可以通过 `configs/` 目录下的配置文件调整：

| 配置文件 | 功能说明 |
|---------|---------|
| `gui_config.toml` | 界面设置（主题、窗口大小、透明度等） |
| `logger_config.toml` | 日志系统配置 |
| `link_config.toml` | 链接嵌入和格式配置 |
| `image_config.toml` | 图片提取默认设置 |
| `output_config.toml` | 输出目录和行为配置 |
| `symbol_settings.toml` | 符号校对规则配置 |
| `typos_settings.toml` | 错别字映射表配置 |
| `sensitive_words.toml` | 敏感词匹配规则配置 |
| `software_priority.toml` | Office软件优先级配置 |

详细配置说明请参考 `doc/技术文档.md`。

## 🛠️ 模板系统

### 使用现有模板

程序自带多个常用公文模板：

- `公文通用.docx` - 标准公文格式
- `白头文.docx` - 无红头公文
- `表格测试.xlsx` - Excel 报表模板

### 自定义模板

- **方法一**：使用 Word 或 WPS 创建模板文件
  1. 参考现有模板，在需要填充的位置插入占位符：`{{公文标题}}`、`{{发文字号}}` 等
  2. 模板中，内置的标题1~标题5，需要手动修改样式
  3. 将模板保存到 `templates/` 目录

- **方法二**：复制现有模板，修改后重命名

## 🔧 命令行使用

```bash
# 转换单个文件
python src/cli_run.py document.docx

# 批量转换
python src/cli_run.py file1.docx file2.docx

# 交互模式
python src/cli_run.py
```

## 🔌 Obsidian 插件

项目包含配套的 Obsidian 插件（位于 `obsidian-plugin/` 目录），实现与转换器的智能集成：

### 核心特性

- **🚀 一键启动** - 侧边栏图标快速启动转换器
- **📂 智能传递** - 自动传递当前打开的文件路径
- **🔄 单实例管理** - 程序已运行时自动发送文件，无需重复启动
- **💪 崩溃恢复** - 智能检测进程状态，自动清理残留文件

### 工作原理

插件通过**基于文件系统的 IPC（进程间通信）**与转换器交互：

1. **首次点击** → 启动转换器并传入当前文件
2. **再次点击（有文件）** → 替换为新文件（单文件模式）
3. **再次点击（无文件）** → 激活转换器窗口

### 安装方法

1. 构建插件：
   ```bash
   cd obsidian-plugin
   npm install
   npm run release
   ```

2. 复制到 Obsidian 插件目录：
   ```
   <Vault>/.obsidian/plugins/gongwen-converter-assistant/
   ```

3. 在 Obsidian 设置中启用插件并配置转换器路径

详细说明请参考 `obsidian-plugin/README.md`。

## 🏗️ 构建可执行文件

如需打包为独立的可执行程序：

```bash
# 安装构建依赖
pip install -e ".[dev]"

# 运行构建脚本
python scripts/build/build.py
```

构建产物位于 `dist/GongwenConverter_v<版本号>/` 目录。

## ❓ 常见问题

### 转换失败怎么办？

- 检查文件是否被其他程序占用
- 确认文件格式正确
- 查看 `logs/` 目录下的错误日志

### 模板不显示？

- 确认模板文件在 `templates/` 目录中
- 检查模板文件是否损坏
- 重启程序重新加载模板

### 程序无法启动？

- 确认 Python 版本为 3.12+
- 检查所有依赖是否安装成功：`pip install .`
- 查看日志文件了解详细错误信息

### 校对功能不工作？

- 确认文档为 .docx 格式
- 检查文档是否包含可编辑文本
- 在设置中确认校对规则已启用

## 📞 获取帮助

- **详细技术文档**：查看 `doc/技术文档.md`（面向开发者）
- **软件使用说明**：查看 `doc/软件说明书.md`（面向用户）
- **配置示例**：参考 `configs/` 目录中的配置文件

## 🔒 安全特性

- **完全本地运行**：所有处理在本地完成，不依赖网络
- **网络隔离**：内置网络隔离机制，防止数据泄露
- **无数据上传**：用户文件不会上传到任何服务器

## 📜 许可证

本项目采用 **GNU Affero General Public License v3.0 (AGPL-3.0)** 许可证。

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

### 重要说明

- 本项目使用了 PyMuPDF（采用AGPL-3.0许可证），因此整个项目也采用AGPL-3.0许可证
- 您可以自由地使用、修改和分发本软件
- 如果您修改本软件并通过网络提供服务，必须向用户提供修改后的源代码
- 详细许可证信息请参阅 [LICENSE](LICENSE) 文件
- 第三方组件许可证信息请参阅 [LICENSE_THIRD_PARTY.txt](LICENSE_THIRD_PARTY.txt)

### 源代码获取

- **GitHub**: https://github.com/ZHYX91/gongwen-converter
- **联系作者**: zhengyx91@hotmail.com

## 💡 技术特点

- **模块化架构**：清晰的分层设计，易于维护和扩展
- **配置驱动**：灵活的 TOML 配置系统
- **多轮评分算法**：智能识别公文元素
- **策略模式**：可扩展的转换策略体系
- **MVP模式GUI**：视图与逻辑分离的界面设计
- **文件系统IPC**：基于文件监控的进程间通信，完全不依赖网络
- **单实例锁**：跨平台的文件锁机制，确保程序单实例运行

---

**作者**：ZhengYX
**版本**：从 `src/gongwen_converter/__init__.py` 读取  
**Python要求**：>=3.12
