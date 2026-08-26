# DocWen

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md)

文檔圖表格式轉換軟體 - 支持 Word/Markdown/Excel 互轉，完全本地運行，數據安全可靠。

## 📖 專案背景

本軟體最初為文印室日常工作設計，解決以下問題：
- 各科室發來的文檔格式混亂，需要整理為規範格式
- 文檔類型繁多，每種類型有不同的固定格式要求
- 需要離線運行，適配內網環境和老舊設備

**設計理念**：本軟體定位為輕量級傻瓜式工具，在專業性和功能完整性上無法與 LaTeX、Pandoc 等專業工具相比，但勝在零學習成本、開箱即用，適合對格式要求不高的日常辦公場景。

## ✨ 核心功能

- **📄 文檔格式轉換** - Word ↔ Markdown 互轉，支持數學公式轉換、分隔符雙向轉換（Markdown 的三種分隔線與文檔中的分頁符、分節符、分隔線），以及將 Markdown 表格顯式 `<` / `^` marker 恢復為 Word 矩形合併。支持 DOCX/DOC/WPS/RTF/ODT 等格式。
- **📊 表格格式轉換** - Excel ↔ Markdown 互轉，支持 XLSX/XLS/ET/ODS/CSV/TSV 等格式；支持合併單元格導出策略（`fill / empty / marker`）和表格匯總工具。Markdown→XLSX 範本已恢復 YAML 欄位、縱向和橫向欄占位符核心能力，完整 Excel 範本區域/圖片/合併恢復仍是對照審計中的遷移目標。
- **📑 PDF與版式文件** - PDF/XPS/OFD 轉 Markdown 或 DOCX，支持 PDF 合併、拆分等操作。
- **🖼️ 圖片處理** - 支持 JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC 等格式互轉和壓縮。
- **📥 其他格式導入** - 支持 HTML/MHTML/ENEX/EPUB/PPTX/PPT 單向轉換為 Markdown。
- **🔍 OCR文字識別** - 集成 RapidOCR，從圖片和 PDF 中提取文字。
- **✏️ 文本校對** - 基於自定義詞庫檢查錯別字、標點、符號和敏感詞。支持 Word (.docx) 和 Markdown (.md) 文件。可在設定介面編輯規則。
- **📝 範本系統** - 靈活的範本機制，支持自定義文檔和報表格式。
- **💻 雙模式操作** - 圖形介面 + 命令行介面。
- **🔒 本地處理與依賴出站保護** - 轉換不依賴線上服務。DocWen 運行期間會阻止其 Python 進程內依賴使用 DNS 和 IPv4/IPv6；外部啟動的 Office/WPS/LibreOffice 仍遵循自身及系統網絡策略。
- **🔗 單實例運行** - 自動管理程式實例，支持與配套 Obsidian 插件集成。

## 📸 介面截圖

| 批次處理 | Markdown |
| --- | --- |
| ![批次處理介面](../assets/screenshots/batch-light.png) | ![主視窗](../assets/screenshots/main-light.png) |

| 文檔 | 表格 |
| --- | --- |
| ![文檔介面](../assets/screenshots/conversion-document-light.png) | ![表格介面](../assets/screenshots/conversion-spreadsheet-light.png) |

| 圖片 | 版式檔案 |
| --- | --- |
| ![圖片介面](../assets/screenshots/conversion-image-light.png) | ![版式檔案介面](../assets/screenshots/conversion-layout-light.png) |

更新日誌：見 [CHANGELOG.md](../CHANGELOG.md)

## 🚀 快速開始

### 從原始碼安裝

**前置條件**：Python 3.12

**0.9 目標邊界**：目前原始碼建置 Windows x64 與 Ubuntu 24.04 x64 正式套件。其他 Linux
發行版與 macOS 仍屬於原始碼／開發路徑，不在 Ubuntu 套件的支援承諾內。

**方式一：使用 uv（推薦）**

安裝 [uv](https://docs.astral.sh/uv/getting-started/)，然後：

```bash
git clone https://github.com/ZHYX91/docwen.git
cd docwen
uv sync --frozen --all-extras
```

DocWen 0.9 原始碼、測試與建置僅支援倉庫內鎖定檔及 `uv 0.12.0`；不支援 `pip install -e`。

### 啟動程式

Windows 打包版可雙擊 `DocWen.exe` 啟動圖形介面。從原始碼安裝後可執行：

```bash
docwen-gui  # 圖形介面
docwen      # 命令列
```

### macOS 安裝說明

**目前限制**：macOS 上的 `convert`、`validate`、`number`、`merge`、`split` capability 目前均為
unavailable。以下只記錄開發實驗所需的選用相依套件。

**LibreOffice 支援（可選）**

如需轉換 `.doc`、`.xls` 等舊格式，請安裝 LibreOffice：  
下載地址：https://www.libreoffice.org/download/

**HEIC 圖片支援（可選）**

如需處理 HEIC/HEIF 格式圖片：

```bash
brew install libheif
pip install pillow-heif
```

### Linux GUI 版本前置條件

**支援的打包目標**：DocWen 0.9 支援 Ubuntu 24.04 x64 打包 GUI 與 CLI。以下前置條件
不會把支援承諾擴展到其他發行版或架構。

- 已安裝桌面環境（GNOME、KDE、XFCE 等均可）
- GUI 基於 PySide6（Qt6），不再依賴 Python Tk；如啟動時提示缺少系統函式庫，請依錯誤訊息安裝對應的 Qt 執行期依賴（常見為 OpenGL/X11 相關函式庫）。
- 純伺服器（無顯示環境）建議優先使用 `docwen` CLI 入口而非 GUI；Windows 打包版同時提供 `DocWenCLI.exe`。

### 快速入門指南

1. **準備一個 Markdown 文件**：

   ```markdown
   ---
   標題: 測試文檔
   ---
   
   ## 測試標題
   
   這是測試正文內容。
   ```

2. **拖拽轉換**：
   - 啟動程式
   - 將 .md 文件拖入窗口
   - 選擇範本
   - 點擊"轉換為 DOCX"

3. **獲取結果**：
   - 在相同目錄下生成格式規範的 Word 文檔

**提示**：新手可以使用 `samples/` 目錄中的範例文件快速體驗軟體功能。

## 🖥️ 圖形介面使用

大部分用戶通過圖形介面使用本軟體，以下是詳細的操作指南。

### 介面概覽

程式採用**自適應三欄佈局**設計：

| 區域 | 說明 | 顯示時機 |
|-----|------|---------|
| **中欄（主區域）** | 文件拖拽區、操作面板、狀態欄 | 始終顯示 |
| **右欄** | 範本選擇器 / 格式轉換面板 | 選擇文件後自動展開 |
| **左欄** | 批量文件列表（按類型分組） | 切換到批量模式時顯示 |

### 基本操作流程

1. **啟動程式**：雙擊 `DocWen.exe`（Windows 打包版）或執行 `docwen-gui`
2. **導入文件**：
   - 方式一：直接拖拽文件到窗口
   - 方式二：點擊拖拽區域的"添加"按鈕選擇文件
3. **選擇範本**（如需轉換）：右側範本面板自動展開，選擇合適的範本
4. **配置選項**：在操作面板中勾選需要的轉換/導出選項
5. **執行操作**：點擊對應的功能按鈕（如"導出MD"、"轉換為DOCX"等）
6. **查看結果**：狀態欄顯示處理進度和結果，可點擊右側的「開啟輸出」操作打開輸出位置

### 單文件模式與批量模式

程式支持兩種處理模式，可通過文件拖拽區域的切換按鈕切換：

**單文件模式**（默認）：
- 一次處理一個文件
- 介面簡潔，適合日常使用

**批量模式**：
- 可同時導入多個文件
- 左欄顯示分類文件列表（按文檔/表格/圖片等分組）
- 支持批量添加、移除、排序
- 點擊列表中的文件可切換當前操作對象

### 操作面板功能

操作面板根據文件類型自動調整可用選項：

| 文件類型 | 可用操作 |
|---------|---------|
| Word文檔 | 導出MD、轉換PDF、文本校對、OCR |
| Markdown | 轉換DOCX、轉換PDF、文本校對 |
| Excel表格 | 導出MD、轉換PDF、表格匯總 |
| PDF文件 | 導出MD、合併、拆分、OCR |
| 圖片文件 | 格式轉換、壓縮、OCR |
| HTML/EPUB/PPTX等 | 導出MD |

### 設定介面

點擊操作區標題列中的「設定」按鈕打開設定介面，可配置：

設定按選項卡組織：**通用**、**文字**、**校對**、**文件**、**表格**、**圖片**、**版式**、**連結**、**格式**、**輸出**、**匯出**、**日誌**、**其他**。

### 快捷操作

- **拖拽外部文件**：直接拖入窗口即可導入
- **開啟輸出**：點擊狀態欄右側的「開啟輸出」操作打開輸出位置
- **右鍵範本項**：打開範本文件位置

---

## 🔧 命令列使用

除了圖形介面，程式也提供命令列介面（CLI），適合自動化腳本、批次處理與外部整合情境。

### 建議的自動化呼叫順序

對於腳本、Agent 或外掛整合，建議依照以下順序呼叫：

1. `inspect <file> [--json]`：先識別檔案的真實類別、格式與可執行動作。
2. `schema convert`：讀取 `convert` 的機器可讀參數契約與條件約束。
3. `convert <file> --to <fmt> --output <path> --dry-run --json`：先預演檢測、正規化與路由結果，不直接落地轉換。
4. `convert <file> --to <fmt> --output <path> ...`：確認後再執行正式轉換。

### 常用範例

```bash
# Windows 打包版
DocWenCLI.exe inspect document.docx --json

# 查看 convert 參數契約（適合腳本 / Agent 組裝請求）
DocWenCLI.exe schema convert

# 預演本次轉換會如何執行，但不實際寫出結果
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --ocr --dry-run --json

# 匯出 Word 為 Markdown（提取圖片 + OCR）
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --ocr

# Markdown 轉 Word（指定模板，並設定標題 + 正文合併模式）
DocWenCLI.exe convert document.md --to docx --output document.docx --template template.docx.926a8bcb579f16e796662bff35edb7bb437aa718b7b7739b2461f4641128001e --heading-merge-mode punct_required

# 控制 Markdown 匯出圖片與 OCR 文字落位
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --image-mode file --ocr --ocr-placement image_md

# 查看目前環境下的執行期能力與依賴門控
DocWenCLI.exe doctor --json
DocWenCLI.exe resources list formats --json

# 文件校對
DocWenCLI.exe validate document.docx --check typo --check punct
DocWenCLI.exe validate input.md --check typo --check punct

# 原始碼 / uv 安裝
# inspect -> schema -> dry-run -> convert
# docwen inspect document.docx --json
# docwen schema convert
# docwen convert document.docx --to md --output document.md --dry-run --json
# docwen convert document.docx --to md --output document.md
```

### 常用命令與選項

下表只列出常用命令；完整命令面請以 `docwen --help`（原始碼 / uv 安裝）或 `DocWenCLI --help`（打包版）為準。

| 命令/選項 | 說明 |
| --- | --- |
| `convert <file> --to <fmt> --output <path>` | 統一轉換入口。 |
| `convert <file> --to <fmt> --output <path> --dry-run --json` | 只預演檢測、正規化、路由與生效參數，不執行實際轉換。 |
| `schema convert` | 匯出 `convert` 的機器可讀參數契約、預設值、條件約束與規範鍵。 |
| `validate <file> --check ...` | 文件校對（`typo/punct/symbol/sensitive/all/none`）。需要 CLI 頂層 Envelope JSON 時使用 `--json`；`--report` 是可選的報告檔案路徑。 |
| `inspect <file> [--json]` | 查詢檔案類別/格式、建議動作，以及副檔名與內容不一致警告。 |
| `doctor --json` | 輸出診斷結果，並附帶執行期能力摘要與依賴門控資訊。 |
| `resources list formats --json` | 按來源類別列出可用目標格式，並附帶執行期依賴門控 / 限制摘要。 |
| `resources list templates` | 列出可用模板。 |
| `resources list numbering-schemes` | 列出可用編號方案。 |
| `--template <id>` | 原樣使用 `resources list templates` 回傳的 canonical 資源 ID；顯示名稱、檔名與路徑直接拒絕。DOCX ID 用於 `docx/doc/odt/rtf/wps/pdf`，XLSX ID 用於 `xlsx/xls/ods/csv`。 |
| `--extract-img` / `--no-extract-img` / `--ocr` | `convert --to md` 的圖片提取與 OCR 選項。 |
| `--image-mode file|base64` | 控制 Markdown 匯出時圖片的落地方式。 |
| `--ocr-placement image_md|main_md` | 控制 OCR 文字寫入圖片配套 Markdown 或主 Markdown。 |
| `--heading-merge-mode punct_required|always|never` | 控制 `convert --to docx` 時「標題 + 正文」段落合併策略。 |
| `--optimization <id>` | 明確啟用某個最佳化配置（可用列表見 `resources list optimizations`）。 |
| `batch convert|validate ... --jobs <n> [--continue-on-error]` | 批次處理控制。 |
| `--json` / `--quiet` / `--timing` | 結構化輸出、壓縮日誌與耗時資訊，適合腳本或外掛呼叫。 |


## 📝 Markdown 語法約定

### 標題級別映射

為方便無背景知識的同事記憶，本軟體的Markdown標題與Word標題**一一對應**：
- 文檔的標題（title）和副標題（subtitle）放在YAML元數據中
- Markdown的 `# 一級標題` 對應 Word的"標題1"
- Markdown的 `## 二級標題` 對應 Word的"標題2"
- 以此類推，最多支持9級標題

**提示**：如果您習慣用 Markdown 的一級標題（`#`）表示文檔標題（title），從二級標題（`##`）開始表示正文小標題（heading），可以在 Word 範本中將「標題1」樣式調整為文檔標題的外觀（如置中、粗體、較大字號），並在設定的序號方案中選擇「跳過一級標題編號」的方案。這樣，一級標題就能呈現為文檔標題的效果。

### 換行與分段

**基本規則**：每個非空行默認作為獨立段落處理。

**混合段落**：當小標題需要與正文混合在同一段時（預設「有標點才合併」模式），需滿足以下條件：
1. 小標題末尾是設定中允許的觸發符號。預設值精確為 `。：！？.:!?`，即全形或半形的句號、冒號、問號、驚嘆號
2. 正文文本位於小標題的**緊鄰下一行**
3. 正文行不能是特殊 Markdown 元素（如標題、代碼塊、表格、列表、引用、公式塊、分隔線等）

**範例**：
```markdown
## 一、工作要求。
本次會議要求各單位認真落實...
```
上述兩行會被合併為同一段落，其中"一、工作要求。"保持小標題格式，"本次會議..."為正文格式。

**注意**：
- 小標題和正文之間不能有空行，否則會被識別為獨立段落
- 預設情況下（有標點才合併模式），若小標題末尾沒有結束標點符號，即使與正文之間無空行，也會被識別為獨立段落，不會合併
- 可在設定介面「格式」→「Markdown 轉為文件」中調整「標題 + 正文合併模式」和觸發符號。觸發符號留空時，「有標點才合併」模式不會合併任何段落
- 逗號、分號、頓號、破折號和刪節號預設不觸發合併；確有需要時可由使用者明確加入

### 分隔線雙向轉換

支持 Markdown 分隔線與 Word 分頁符/分節符/分隔線的雙向轉換：

- **DOCX → MD**：Word 中的分頁符、分節符、分隔線自動轉換為 Markdown 分隔線
- **MD → DOCX**：Markdown 中的 `---`、`***`、`___` 自動轉換為對應的 Word 元素
- **可配置**：具體映射關係可在設定介面自定義

### 任務列表

支援 GFM 任務列表的雙向轉換：

```markdown
- [ ] 待辦事項
- [x] 已完成
```

- **MD → DOCX**：渲染為無序列表，文字前添加 `☐` / `☑` 前綴
- **DOCX → MD**：識別列表項中的 `☐` / `☑` / `☒` 前綴，還原為 `- [ ]` / `- [x]`
- **字型提示**：`☐`/`☑` 在部分字型下可能無法顯示，如有需要請在 Word 範本中使用 "Segoe UI Symbol" 等支援該字元的字型

### 圖片嵌入與尺寸

支持 Obsidian/Wiki 與標準 Markdown 圖片嵌入，並可指定寬高（單位：px）：

```markdown
![[image.png]]
![[image.png|300]]
![[image.png\|300]]
![alt](image.png =300x200)
![alt](image.png =300x)
![alt|300](image.png)
```

- 未指定尺寸：使用圖片原始尺寸，但不超過頁面/單元格可用寬度
- 指定尺寸：允許放大，但仍受可用寬度上限限制
- 純圖片段落：自動使用「圖片」段落樣式（置中、單倍行距）

### 連結處理

支援 Markdown -> DOCX 的可點擊連結：

```markdown
[Docwen](https://example.com)
[[Target]]
[[Target|Open target]]
<https://example.com>
<user@example.com>
```

- Markdown 連結與 Wiki 連結預設寫入為 Word 超連結
- Wiki 連結在找到目標檔案時會解析為本機 `file:///` 連結
- 尖括號自動連結支援 `https://...` 與電子郵件 `mailto:...`
- 裸 URL 自動連結按 Markdown -> DOCX 請求獨立套用，預設關閉，可透過 `configs/link.toml` 的 `[non_embed_links].auto_link_bare_url` 開啟
- Markdown -> XLSX 不會產生 DOCX 超連結佔位符，會保留原始連結語法

## 📖 詳細使用指南

### Word 轉 Markdown

1. 將 .docx 文件拖入程式窗口
2. 程式自動分析文檔結構
3. 生成包含 YAML 元數據的 .md 文件

**支持的格式**：
- `.docx` - 標準 Word 文檔
- `.doc` - 自動轉換為 DOCX 後處理
- `.wps` - WPS 文檔自動轉換

**導出選項說明**：

| 選項 | 說明 |
|-----|------|
| **提取圖片** | 勾選後，將文檔中的圖片提取到輸出文件夾，MD文件中插入圖片連結 |
| **圖片文字識別** | 勾選後，對圖片進行OCR識別，創建圖片.md文件（包含識別的文字） |
| **進階欄位優化** | 勾選後，提取更豐富的結構化元數據；不勾選則使用簡化模式，YAML 只包含標題和副標題兩個基本字段 |
| **清理小標題序號** | 勾選後，移除小標題前的序號（如"一、""（一）""1."等），轉換為純標題文本 |
| **添加小標題序號** | 勾選後，根據標題層級自動添加序號（可在設定中配置序號方案） |

補充：DOCX -> MD 現已支援還原透過段落樣式（pStyle）關聯到 numbering.xml 的多級列表序號，因此 Word/WPS 中以多級列表實現的標題前綴（如「一、」「（一）」「1．」「（1）」「①」）在簡化模式與進階欄位模式下都能保留；勾選「清理小標題序號」後仍會正確識別標題層級。

### Markdown 轉 Word

1. 準備包含 YAML 頭部的 .md 文件
2. 拖入程式窗口並選擇對應的 Word 範本
3. 程式自動填充範本並生成文檔

**轉換選項說明**：

| 選項 | 說明 |
|-----|------|
| **清理小標題序號** | 勾選後，移除小標題前的序號（如"一、""（一）""1."等），轉換為純標題文本 |
| **添加小標題序號** | 勾選後，根據標題層級自動添加序號（可在設定中配置序號方案） |

**注意**：如果文檔中有小標題和正文文本混合的段落，在MD文件中需要保持嚴格換行（參見上方"換行與分段"說明）。

### 範本樣式自動處理

轉換器在 Markdown → DOCX 轉換時會自動檢測和處理範本樣式：

#### 樣式分類

**段落樣式（Paragraph Style）**：應用到整個段落

| 樣式 | 檢測行為 | 缺失時注入 | 來源 |
|-----|---------|-----------|-----|
| 標題 (heading 1~9) | 檢測段落樣式 | 模板標題樣式 | Word 內置 |
| 代碼塊 | 檢測段落樣式 | Consolas 字體 + 灰色背景 | 本軟體定義 |
| 引用 (Quote 1~9) | 檢測段落樣式 | 灰色背景 + 左邊框 | 本軟體定義 |
| 公式塊 | 檢測段落樣式 | 公式專用樣式 | 本軟體定義 |
| 分隔線 (1~3) | 檢測段落樣式 | 底部邊框段落樣式 | 本軟體定義 |

**字符樣式（Character Style）**：應用到選中的文字

| 樣式 | 檢測行為 | 缺失時注入 | 來源 |
|-----|---------|-----------|-----|
| 行內代碼 | 檢測字符樣式 | Consolas 字體 + 灰色底紋 | 本軟體定義 |
| 行內公式 | 檢測字符樣式 | 公式專用樣式 | 本軟體定義 |

**表格樣式（Table Style）**：應用到整個表格

| 樣式 | 檢測行為 | 缺失時注入 | 來源 |
|-----|---------|-----------|-----|
| 三線表 | 用戶配置優先 | 三線表樣式定義 | 本軟體定義 |
| 網格表 | 用戶配置優先 | 網格表樣式定義 | 本軟體定義 |

**編號定義（Numbering Definition）**：用於列表格式

| 類型 | 檢測行為 | 缺失時處理 |
|-----|---------|-----------|
| 列表編號 | 掃描範本中已有的有序/無序列表定義 | 使用 decimal/bullet 預設 |

#### 樣式名國際化說明

- **Word 內置樣式**（標題 heading 1~9）：
  - 樣式名使用 Word 標準英文名稱（如 `heading 1`）
  - Word 根據系統語言自動本地化顯示（中文系統顯示為"標題 1"）
  
- **本軟體定義的樣式**（代碼塊、引用、公式、分隔線、表格等）：
  - 根據本軟體的介面語言設定，注入相應語言的樣式名
  - 中文介面：注入"代碼塊"、"引用 1"、"三線表"等
  - 英文介面：注入 "Code Block"、"Quote 1"、"Three Line Table" 等

**使用建議**：在範本中自定義樣式後，轉換器會自動使用您的樣式；如果範本中沒有，會使用內置預設樣式。

### 表格文件處理

1. **Excel/CSV 轉 Markdown**：拖入 .xlsx 或 .csv 文件，自動轉換為 Markdown 表格
2. **Markdown 轉 Excel**：可將 Markdown 表格匯出為 XLSX；XLSX 範本支援 YAML 欄位、縱向/橫向表格欄占位符、圖片占位符、合併/保護儲存格處理和 Markdown 表格合併恢復。

**支持的格式**：
- `.xlsx` - 標準 Excel 文檔
- `.xls` - 自動轉換為 XLSX 後處理
- `.et` - WPS 表格自動轉換
- `.csv` - CSV 文本表格
- `.tsv` - TSV 製表符分隔表格

### 文本校對功能

程式提供四種可自定義的校對規則：

1. **標點配對檢查** - 檢測括號、引號等成對標點是否匹配
2. **符號校對** - 檢測中英文標點混用問題
3. **錯別字檢查** - 基於自定義詞庫檢查常見錯別字
4. **敏感詞檢測** - 基於自定義詞庫檢測敏感詞

**自定義詞庫**：在程式的"設定"介面中可視化編輯錯別字庫和敏感詞庫。

**使用方法**：
1. 將需要校對的 Word 或 Markdown 文檔拖入程式
2. 勾選需要的校對規則
3. 點擊"文本校對"按鈕
4. Word 文檔的校對結果以批註形式顯示在文檔中；Markdown 文件的校對結果輸出為結構化 JSON 報告

補充說明（Markdown 校對 JSON 報告）：
- 引擎：`text_rules` + Markdown 適配層 `md_spell`
- 輸出方式：目前 CLI 校對入口為 `validate`；需要 CLI 頂層 Envelope JSON 時使用 `--json`。`--report` 是可選的報告檔案路徑。

## 🛠️ 範本系統

### 使用現有範本

程式自帶多種範本，包含多語言版本，可根據需要選擇使用。範本文件位於 `templates/` 目錄。

### 自定義範本

1. 使用 Word 或 WPS 創建範本文件
2. 參考現有範本，在需要填充的位置插入占位符：`{{標題}}` 等
3. 範本中，內置的標題1~標題5，需要手動修改樣式
4. 將範本保存到 `templates/` 目錄
5. 重啟程式，新範本自動加載

也可以複製現有範本，修改後重命名。

### 占位符使用說明

#### Word 範本占位符

**YAML 字段占位符**：在範本中使用 `{{字段名}}` 格式，轉換時會被 Markdown 文件 YAML 頭部的對應值替換。

| 占位符 | 說明 |
|-------|------|
| `{{標題}}` | 文檔標題（獲取規則見下方說明） |
| `{{正文}}` | Markdown 正文內容插入位置 |
| 其他 | 支持任意自定義字段 |

**標題獲取優先級**：

| 優先級 | 來源 | 說明 |
|-------|------|------|
| 1 | YAML `標題` 字段 | 最高優先級 |
| 2 | YAML `aliases` 字段 | 取列表第一個元素，或字符串值 |
| 3 | 文件名 | 去除 `.md` 擴展名後的文件名 |

**多語言支援**：標題和正文占位符支援多語言寫法，如標題可使用 `{{標題}}`、`{{title}}`、`{{Titel}}` 等，正文可使用 `{{正文}}`、`{{body}}`、`{{Inhalt}}` 等。

#### Excel 範本占位符（舊專案能力對照目標）

XLSX 範本支援 YAML 欄位占位符、縱向 `{{↓欄位}}`、橫向 `{{→欄位}}` 表格欄占位符、圖片占位符、合併/保護儲存格處理和 Markdown 表格合併恢復。

**1. YAML 字段占位符** `{{字段名}}`

用於填充 Markdown 文件 YAML 頭部的單一值：

```markdown
---
報表名稱: 2024年度銷售統計表
編制單位: 銷售部
編制日期: 2024年12月31日
---
```

範本中的 `{{報表名稱}}`、`{{編制單位}}` 等會被替換為對應值。標題字段同樣按優先級獲取（YAML標題 → aliases → 文件名）。

**2. 列填充占位符** `{{↓字段名}}`

從 Markdown 表格提取數據，從占位符位置開始**向下**逐行填充：

```markdown
| 產品名稱 | 銷售數量 |
|---------|---------|
| 產品A | 100 |
| 產品B | 200 |
```

Excel 範本中的 `{{↓產品名稱}}` 會被替換為"產品A"，下一行填充"產品B"。

**3. 行填充占位符** `{{→字段名}}`

從 Markdown 表格提取數據，從占位符位置開始**向右**逐列填充：

```markdown
| 月份 |
|------|
| 1月 |
| 2月 |
| 3月 |
```

Excel 範本中的 `{{→月份}}` 會依次向右填充"1月"、"2月"、"3月"。

**合併儲存格處理**：

- Markdown -> Excel 會繼續保留範本原有的 merged ranges。
- 對由連續整格 `{{↓欄位名}}` 組成的已知列式範本區域，支援根據 Markdown 表格中的顯式 `<` / `^` marker 還原矩形合併。
- 僅當儲存格去掉首尾空白後精確等於 `<` 或 `^` 時才參與 merge 識別；`\<`、`\^` 會保留為字面文字。
- 非法矩形或與範本原有 merged ranges 衝突時，預設降級為普通文字並記錄警告，不會強制覆蓋範本結構。

**多表格數據合併**：如果 Markdown 中有多個表格使用相同的表頭名，數據會按順序合併後依次填充。

## 🔌 Obsidian 插件

配套 Obsidian 插件已在獨立倉庫發布，可與轉換器聯動使用：

### 核心特性

- **🚀 一鍵啟動** - 側邊欄圖標快速啟動轉換器
- **📂 自動傳遞** - 自動傳遞當前打開的文件路徑
- **🔄 單實例管理** - 程式已運行時自動發送文件，無需重複啟動
- **🔒 有界本機控制** - 使用有類型的 `status`、`open`、`activate` 請求，不按進程名稱探測，也不使用命令檔或狀態檔

### 工作原理

DocWen Core 的 runtime/control transport 可在 Windows 使用命名管道，在 Linux/macOS 使用
AF_UNIX 通訊端。檔案鎖只負責單一實例所有權，控制命令不透過檔案傳輸。這只是 Core 能力
說明；DocWen Assistant 2.0 仍僅限 Windows 桌面端，尚無 Linux/macOS 組合驗收。

1. **首次點擊** → 啟動轉換器並傳入當前文件
2. **再次點擊（有文件）** → 替換為新文件（單文件模式）
3. **再次點擊（無文件）** → 激活轉換器窗口

### 安裝方法

DocWen Assistant 2.0 使用 DocWen Machine Protocol v1 與唯一的 Artifact Bundle v2 合同。原始碼版本不能
證明已經發布；請只安裝明確標識了相容且已發布 DocWen 版本的數字版本 Release。

## 🔌 OpenClaw（插件 + Skill）

OpenClaw 2.0 使用 DocWen Machine Protocol v1 與唯一的 Artifact Bundle v2 合同。原始碼版本不能證明已經
發布；請以數字版本 Release 頁面為準，並只在不可變發布閘門成功後安裝。

## ❓ 常見問題

### 轉換失敗怎麼辦？

- 檢查文件是否被其他程式占用
- 確認文件格式正確
- 在設定頁查看「目前實際日誌檔案路徑」，或到系統使用者日誌目錄中查看錯誤日誌；若打包驗證使用了 `DOCWEN_LOG_DIR`，則查看對應的覆寫目錄

### 範本不顯示？

- 確認範本文件在 `templates/` 目錄中
- 檢查範本文件是否損壞
- 重啟程式重新加載範本

### 校對功能不工作？

- 確認文檔為 .docx 或 .md 格式
- 檢查文檔是否包含可編輯文本
- 在設定中確認校對規則已啟用

### 輸出格式不符合預期？

- 程式根據範本樣式生成文檔，如需調整輸出格式，請直接修改範本文件中的樣式定義
- 範本文件位於 `templates/` 目錄
- 修改範本樣式後，所有使用該範本轉換的文檔都會應用新樣式

### Excel轉Markdown後公式儲存格為空？

這是正常現象，原因是程式讀取的是儲存格的**快取值**而非公式本身。

**技術原因**：
- Excel 文件中，公式儲存格同時存儲公式和上次計算的結果（快取值）
- 程式使用 `data_only=True` 模式讀取，只獲取快取值
- 如果文件從未在 Excel 中打開過（如由程式生成），或編輯後未重新保存，快取值就是空的

**解決方法**：
1. 在 Excel 中打開文件
2. 等待公式計算完成
3. 保存文件
4. 重新轉換

## 🔒 安全特性

- **完全本地運行**：所有處理預設在本地完成，不依賴線上服務
- **依賴出站保護**：受支援的 GUI/CLI 入口會在整個 Python 主進程生命週期內啟用 CPython 稽核守衛；守衛阻止進程內依賴執行所有 DNS／名稱解析以及 AF_INET/AF_INET6 的 `bind`、`connect`、`connect_ex`、`sendto`、`sendmsg` 操作，同時保留 Windows 命名管道與 Unix 域套接字
- **明確邊界**：單獨啟動的進程（包括 Office/WPS/LibreOffice 及專用 Office helper）不受該守衛管理；它用於防止依賴意外連網，不是作業系統級沙箱
- **無數據上傳**：DocWen 不提供上傳、遙測、線上下載或網絡服務功能
- **嚴格安全模式**：預設啟用；核心安全檢查失敗時程式將終止。進階設定/排障請參閱 [故障排查](../maintenance/troubleshooting.md)。

## 📜 許可證

本項目採用 **GNU Affero General Public License v3.0 (AGPL-3.0)** 許可證。

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

- 本項目使用了 PyMuPDF（採用AGPL-3.0許可證），因此整個項目也採用AGPL-3.0許可證
- 當前 GUI 在支援的宿主路徑中可使用 `PySide6-Fluent-Widgets`（QFluentWidgets）；其採用 `GPLv3 / 商業授權` 雙軌模式，當前倉庫繼續按 AGPL 開源分發
- 您可以自由地使用、修改和分發本軟體
- 如果您修改本軟體並通過網絡提供服務，必須向用戶提供修改後的源代碼
- 詳細許可證資訊請參閱 [LICENSE](../../LICENSE) 文件
- 第三方元件許可證登記請參閱 [LICENSE_THIRD_PARTY.txt](../../LICENSE_THIRD_PARTY.txt)；分發摘要見 [NOTICE.txt](../../NOTICE.txt)

### 聯繫方式

- **GitHub**: https://github.com/ZHYX91/docwen
- **聯繫作者**: zhengyx91@hotmail.com

---

**作者**：ZhengYX
