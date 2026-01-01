# 公文转换器 (Gongwen Converter)

公文图表格式转换软件 - 支持 Word/Markdown/Excel 互转，完全本地运行，数据安全可靠。

## 📖 项目背景

本软件最初为文印室日常工作设计，解决以下痛点：
- 各科室发来的文档格式混乱，需要整理为规范格式
- 文档类型繁多，每种类型有不同的固定格式要求
- 需要完全离线运行，适配内网环境和老旧设备

**设计理念**：降低学习成本，不依赖pandoc/LaTeX等技术工具，提供傻瓜式操作界面。

## ✨ 核心功能

- **📄 文档格式转换** - Word ↔ Markdown 互转，智能识别公文元素，支持数学公式转换、分隔符双向转换（`---` ↔ 分页符，`***` ↔ 分节符，`___` ↔ 分隔线）。支持 DOCX/DOC/WPS/RTF/ODT 等格式。
- **📊 表格格式转换** - Excel ↔ Markdown 互转，支持 XLSX/XLS/ET/ODS/CSV 等格式。包含表格汇总工具。
- **📑 PDF与版式文件** - PDF/XPS/OFD 转 Markdown 或 DOCX，支持 PDF 合并、拆分等操作。
- **🖼️ 图片处理** - 支持 JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC 等格式互转和智能压缩。
- **🔍 OCR文字识别** - 集成 PaddleOCR，从图片和 PDF 中提取文字。
- **✏️ 文本校对** - 基于自定义词库检查错别字、标点、符号和敏感词。可在设置界面编辑规则。
- **📝 模板系统** - 灵活的模板机制，支持自定义公文和报表格式。
- **💻 双模式操作** - 图形界面 + 命令行界面。
- **🔒 完全本地运行** - 离线运行，数据安全，内置网络隔离机制。
- **🔗 单实例运行** - 智能管理程序实例，支持与 Obsidian 插件无缝集成。

## 更新日志

### v0.5.1 (2025-01-01)

- 新增数学公式双向转换（Word OMML ↔ Markdown LaTeX）
- 新增脚注/尾注双向转换
- 新增代码、引用等字符和段落样式
- 增强列表处理（多级嵌套、自动编号）
- 增强表格功能（样式检测/注入、三线表等）
- 优化小标题序号清理和添加
- 改进界面交互和设置联动

### v0.4.1 (2025-12-05)

- 重构命令行界面，提升用户体验
- 添加对非公文文档转换的支持
- 实现更多选项配置化

## 🚀 快速开始

### 启动程序

双击 `GongwenConverter.exe` 启动图形界面。

### 快速入门指南

1. **准备一个 Markdown 文件**：

   ```markdown
   ---
   标题: 测试文档
   发文字号: 测试〔2024〕1号
   主送机关: 测试部门
   成文日期: 2024年10月26日
   ---
   
   ## 测试标题
   
   这是测试正文内容。
   ```

2. **拖拽转换**：
   - 启动程序
   - 将 .md 文件拖入窗口
   - 选择模板
   - 点击"转换为 DOCX"

3. **获取结果**：
   - 在相同目录下生成格式规范的 Word 文档

## 📝 Markdown 语法约定

### 标题级别映射

为方便无背景知识的同事记忆，本软件的Markdown标题与Word标题**一一对应**：
- 文档的标题（title）和副标题（subtitle）放在YAML元数据中
- Markdown的 `# 一级标题` 对应 Word的"标题1"
- Markdown的 `## 二级标题` 对应 Word的"标题2"
- 以此类推，最多支持9级标题

### 换行与分段

**基本规则**：每个非空行默认作为独立段落处理。

**混合段落**：当小标题需要与正文混合在同一段时，需满足以下条件：
1. 小标题末尾有标点符号（支持：。，；：！？ 及对应的英文标点）
2. 正文文本位于小标题的**紧邻下一行**
3. 正文行不能是标题（不以 `#` 开头）

**示例**：
```markdown
## 一、工作要求。
本次会议要求各单位认真落实...
```
上述两行会被合并为同一段落，其中"一、工作要求。"保持小标题格式，"本次会议..."为正文格式。

**注意**：
- 小标题和正文之间不能有空行，否则会被识别为独立段落
- 如果小标题末尾没有标点符号，即使没有空行也会被识别为独立段落

### 分隔线双向转换

支持 Markdown 分隔线与 Word 分页符/分节符/分隔线的双向转换：

- **DOCX → MD**：Word 中的分页符、分节符、分隔线自动转换为 Markdown 分隔线
- **MD → DOCX**：Markdown 中的 `---`、`***`、`___` 自动转换为对应的 Word 元素
- **可配置**：具体映射关系可在设置界面自定义

## 📖 详细使用指南

### Word 转 Markdown

1. 将 .docx 文件拖入程序窗口
2. 程序自动分析文档结构
3. 生成包含 YAML 元数据的 .md 文件

**支持的格式**：
- `.docx` - 标准 Word 文档
- `.doc` - 自动转换为 DOCX 后处理
- `.wps` - WPS 文档自动转换

**导出选项说明**：

| 选项 | 说明 |
|-----|------|
| **提取图片** | 勾选后，将文档中的图片提取到输出文件夹，MD文件中插入图片链接 |
| **图片文字识别** | 勾选后，对图片进行OCR识别，创建图片.md文件（包含识别的文字） |
| **针对优化（公文）** | 勾选后，使用三轮评分算法识别公文元素，生成包含14个公文专用字段的YAML元数据；不勾选则使用简化模式，YAML只包含标题和副标题两个基本字段 |
| **清理小标题序号** | 勾选后，移除小标题前的序号（如"一、""（一）""1."等），转换为纯标题文本 |
| **添加小标题序号** | 勾选后，根据标题层级自动添加序号（可在设置中配置序号方案） |

### Markdown 转 Word

1. 准备包含 YAML 头部的 .md 文件
2. 拖入程序窗口并选择对应的 Word 模板
3. 程序自动填充模板并生成文档

**转换选项说明**：

| 选项 | 说明 |
|-----|------|
| **清理小标题序号** | 勾选后，移除小标题前的序号（如"一、""（一）""1."等），转换为纯标题文本 |
| **添加小标题序号** | 勾选后，根据标题层级自动添加序号（可在设置中配置序号方案） |

**注意**：如果公文中有小标题和正文文本混合的段落，在MD文件中需要保持严格换行（参见上方"换行与分段"说明）。

### 模板样式智能处理

转换器在 Markdown → DOCX 转换时会自动检测和处理模板样式：

| 样式类型 | 检测行为 | 缺失时注入 |
|---------|---------|-----------|
| 标题 (heading 1~9) | 检测段落样式 | 公文风格标题样式 |
| 引用 (Quote 1~9) | 检测段落样式 | 灰色背景 + 左边框 |
| 代码块/行内代码 | 检测对应类型 | Consolas 字体 + 灰色背景 |
| 表格样式 | 用户配置优先，不存在则注入 | 三线表/网格表 |
| 列表定义 | 优先使用模板最新定义 | decimal/bullet 预设 |
| 分隔线 | 检测 Horizontal Rule | 底部边框段落样式 |

**使用建议**：在模板中自定义样式后，转换器会自动使用您的样式；如果模板中没有，会使用内置预设样式。

### 表格文件处理

1. **Excel/CSV 转 Markdown**：拖入 .xlsx 或 .csv 文件，自动转换为 Markdown 表格
2. **Markdown 转 Excel**：准备好 MD 文件，选择 Excel 模板进行转换

**支持的格式**：
- `.xlsx` - 标准 Excel 文档
- `.xls` - 自动转换为 XLSX 后处理
- `.et` - WPS 表格自动转换
- `.csv` - CSV 文本表格

### 文本校对功能

程序提供四种可自定义的校对规则：

1. **标点配对检查** - 检测括号、引号等成对标点是否匹配
2. **符号校对** - 检测中英文标点混用问题
3. **错别字检查** - 基于自定义词库检查常见错别字
4. **敏感词检测** - 基于自定义词库检测敏感词

**自定义词库**：在程序的"设置"界面中可视化编辑错别字库和敏感词库。

**使用方法**：
1. 将需要校对的 Word 文档拖入程序
2. 勾选需要的校对规则
3. 点击"文本校对"按钮
4. 校对结果以批注形式显示在文档中

### 支持的公文元素

- 标题、发文字号、密级
- 主送机关、抄送机关、附件
- 签发人、成文日期、印发机关
- 以及其他标准公文要素

## ⚙️ 配置说明

程序行为可以通过 `configs/` 目录下的配置文件（共17个）调整：

| 配置文件 | 功能说明 |
|---------|---------|
| `gui_config.toml` | 界面设置（主题、窗口大小、透明度等） |
| `logger_config.toml` | 日志系统配置 |
| `conversion_config.toml` | 格式转换配置（DOCX↔MD格式保留、分隔符转换） |
| `conversion_defaults.toml` | 文件处理默认设置（提取、OCR、校对、汇总、压缩、DPI等） |
| `heading_numbering_add.toml` | 标题序号方案（添加序号） |
| `heading_numbering_clean.toml` | 序号清理规则（清除序号） |
| `link_config.toml` | 链接嵌入和格式配置 |
| `output_config.toml` | 输出目录和行为配置 |
| `proofread_config.toml` | 校对主配置 |
| `proofread_symbols.toml` | 符号校对规则配置 |
| `proofread_typos.toml` | 错别字映射表配置 |
| `proofread_sensitive.toml` | 敏感词匹配规则配置 |
| `software_priority.toml` | Office软件优先级配置 |
| `style_code.toml` | 代码样式配置（Code Block/Inline Code） |
| `style_formula.toml` | 公式样式配置（Formula Block/Inline Formula） |
| `style_quote.toml` | 引用样式配置（Quote 1-9） |
| `style_table.toml` | 表格样式配置（Three Line Table/Table Content） |

详细配置说明请参考 `doc/技术文档.md`。

## 🛠️ 模板系统

### 使用现有模板

程序自带多个常用公文模板：

- `公文通用.docx` - 标准公文格式
- `白头文.docx` - 无红头公文
- `表格测试.xlsx` - Excel 报表模板

### 自定义模板

1. 使用 Word 或 WPS 创建模板文件
2. 参考现有模板，在需要填充的位置插入占位符：`{{标题}}`、`{{发文字号}}` 等
3. 模板中，内置的标题1~标题5，需要手动修改样式
4. 将模板保存到 `templates/` 目录
5. 重启程序，新模板自动加载

也可以复制现有模板，修改后重命名。

## 🔧 命令行使用

除了图形界面，程序还提供功能完整的命令行界面（CLI），适合批量处理。

### 两种运行模式

1. **交互模式** - 友好的菜单引导，类似GUI操作
2. **Headless模式** - 直接执行命令，适合自动化脚本

### 快速开始

```bash
# 交互模式：拖拽文件到终端即可
GongwenConverter.exe document.docx

# Headless模式：导出为Markdown
GongwenConverter.exe document.docx --action export_md --extract-img

# 批量处理：转换所有Word文档
GongwenConverter.exe *.docx --action export_md --batch --yes
```

### 常用操作示例

**导出Markdown**：
```bash
# 导出Word为MD（提取图片）
GongwenConverter.exe report.docx --action export_md --extract-img

# 导出并启用OCR
GongwenConverter.exe report.docx --action export_md --extract-img --ocr
```

**格式转换**：
```bash
# Word转PDF
GongwenConverter.exe report.docx --action convert --target pdf

# Markdown转Word（指定模板）
GongwenConverter.exe document.md --action convert --target docx --template 公文通用
```

**文档校对**：
```bash
# 启用所有校对规则
GongwenConverter.exe document.docx --action validate --check-punct --check-typo --check-symbol --check-sensitive
```

**批量处理**：
```bash
# 批量转换Word为MD
GongwenConverter.exe *.docx --action export_md --extract-img --batch --yes

# 批量转换为PDF（遇错继续）
GongwenConverter.exe *.xlsx --action convert --target pdf --batch --continue-on-error
```

**PDF操作**：
```bash
# 合并PDF文件
GongwenConverter.exe file1.pdf file2.pdf file3.pdf --action merge_pdfs

# 拆分PDF（指定页码）
GongwenConverter.exe report.pdf --action split_pdf --pages "1-3,5,7-10"
```

### 主要参数说明

| 参数 | 说明 | 示例 |
|-----|------|------|
| `--action` | 操作类型 | `export_md`, `convert`, `validate` |
| `--target` | 目标格式 | `pdf`, `docx`, `xlsx` |
| `--template` | 模板名称 | `公文通用` |
| `--extract-img` | 提取图片 | - |
| `--ocr` | OCR识别 | - |
| `--batch` | 批量模式 | - |
| `--yes` / `-y` | 跳过确认 | - |
| `--json` | JSON输出 | - |
| `--quiet` / `-q` | 安静模式 | - |

## 🔌 Obsidian 插件

项目包含配套的 Obsidian 插件，实现与转换器的智能集成：

### 核心特性

- **🚀 一键启动** - 侧边栏图标快速启动转换器
- **📂 智能传递** - 自动传递当前打开的文件路径
- **🔄 单实例管理** - 程序已运行时自动发送文件，无需重复启动
- **💪 崩溃恢复** - 智能检测进程状态，自动清理残留文件

### 工作原理

插件通过基于文件系统的进程间通信与转换器交互：

1. **首次点击** → 启动转换器并传入当前文件
2. **再次点击（有文件）** → 替换为新文件（单文件模式）
3. **再次点击（无文件）** → 激活转换器窗口

### 安装方法

1. 将 `obsidian-plugin/` 目录中的插件文件复制到 Obsidian 插件目录：
   ```
   <Vault>/.obsidian/plugins/gongwen-converter-assistant/
   ```

2. 在 Obsidian 设置中启用插件并配置转换器路径

详细说明请参考 `obsidian-plugin/README.md`。

## ❓ 常见问题

### 转换失败怎么办？

- 检查文件是否被其他程序占用
- 确认文件格式正确
- 查看 `logs/` 目录下的错误日志

### 模板不显示？

- 确认模板文件在 `templates/` 目录中
- 检查模板文件是否损坏
- 重启程序重新加载模板

### 校对功能不工作？

- 确认文档为 .docx 格式
- 检查文档是否包含可编辑文本
- 在设置中确认校对规则已启用

## 📞 获取帮助

- **详细技术文档**：查看 `doc/技术文档.md`
- **配置示例**：参考 `configs/` 目录中的配置文件

## 🔒 安全特性

- **完全本地运行**：所有处理在本地完成，不依赖网络
- **网络隔离**：内置网络隔离机制，防止数据泄露
- **无数据上传**：用户文件不会上传到任何服务器

## 📜 许可证

本项目采用 **GNU Affero General Public License v3.0 (AGPL-3.0)** 许可证。

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

- 本项目使用了 PyMuPDF（采用AGPL-3.0许可证），因此整个项目也采用AGPL-3.0许可证
- 您可以自由地使用、修改和分发本软件
- 如果您修改本软件并通过网络提供服务，必须向用户提供修改后的源代码
- 详细许可证信息请参阅 [LICENSE](LICENSE) 文件

### 联系方式

- **GitHub**: https://github.com/ZHYX91/gongwen-converter
- **联系作者**: zhengyx91@hotmail.com

---

**作者**：ZhengYX
