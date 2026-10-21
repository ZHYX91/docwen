# DocWen

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md) · [العربية](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ar.md)

DocWen 在本機轉換辦公檔案。它能把 Word 轉成 Markdown、把 Markdown 匯出為 Word、從 PDF 中擷取文字和表格，也能處理試算表、圖片和常見的檔案整理工作。日常操作使用圖形介面；需要批次自動化時使用命令列。

介面語言僅控制應用程式文字與版面；一般設定中的「工作輸出語言」獨立控制產生的標籤、文件樣式與校對註解，不翻譯內文。切換介面不會更改已儲存的輸出語言或範本。OCR「自動」跟隨工作輸出語言。Markdown 語法中的每行獨立成段選項預設關閉，可將一般內文逐行轉成 Word 段落。

## 適合做什麼

- **Word 與 Markdown 互轉**：處理標題、表格和公式；特殊的 Obsidian 語法需要按需開啟對應擴充。
- **表格轉換**：把 Excel 等表格匯出為 Markdown，或從 Markdown 產生試算表檔案。
- **PDF 與版式檔案**：從 PDF、XPS、OFD 擷取文字或轉換檔案，也能合併、拆分 PDF。
- **圖片與 OCR**：轉換常見圖片格式，離線辨識圖片中的文字和部分表格；辨識結果需要人工核對。
- **範本與校對**：使用自訂範本產生文件，依自己的詞庫檢查錯別字、標點與符號。
- **單檔或批次處理**：桌面介面適合日常操作，CLI 適合重複工作。

DocWen 支援 DOCX、XLSX、PDF 和一些舊版辦公、圖片格式，但**並非所有格式之間都能互轉**。實際取決於操作、系統平台及安裝的轉換軟體，參見[功能說明](../capabilities.md)。

## 離線處理

檔案由本機處理，不需要上傳到線上轉換服務。DocWen 限制自身 Python 程序內相依套件的網路存取；外部啟動的 Office、WPS、LibreOffice 仍遵循它們自身及系統的網路原則。處理敏感檔案時，也應檢查這些外部軟體的設定。

## 📸 介面截圖

為 Markdown 文件選擇範本和產生選項。

![Markdown 產生文件](../assets/screenshots/markdown-generation-en.png)


更新日誌：見 [CHANGELOG.md](../CHANGELOG.md)

## 🚀 快速開始

### 從原始碼安裝

**前置條件**：Python 3.12

**安裝套件**：提供 Windows x64 與 Ubuntu 24.04 x64 版本。其他 Linux 發行版與 macOS 需要從原始碼執行，目前沒有正式安裝套件。

**使用 uv 安裝**

安裝 [uv](https://docs.astral.sh/uv/getting-started/)，然後：

```bash
git clone https://github.com/ZHYX91/docwen.git
cd docwen
uv sync --frozen --all-extras
```

請使用 `uv 0.12.0` 和倉庫中的鎖定檔安裝相依套件。測試與正式建置使用此版本；維護相依套件時可使用相容的 `uv 0.12.x`。`pip install -e` 無法套用本倉庫的相依套件排除規則，因此不支援這種安裝方式。

### 啟動程式

從 [Releases 頁面](https://github.com/ZHYX91/docwen/releases) 下載對應套件。Windows 版雙擊 `DocWen.exe`；Ubuntu 24.04 x64 圖形版解壓縮後執行 `./DocWen`，無桌面環境時使用獨立的 `DocWenCLI` 套件。從原始碼安裝後可執行：

```bash
docwen-gui  # 圖形介面
docwen      # 命令列
```

### macOS 安裝說明

**目前限制**：macOS 上的 `convert`、`validate`、`number`、`merge`、`split` 功能目前都無法使用。以下只記錄開發實驗所需的選用相依套件。

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

**支援的打包目標**：DocWen 支援 Ubuntu 24.04 x64 打包 GUI 與 CLI。以下前置條件
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
   - 點擊"生成 DOCX"

3. **獲取結果**：
   - Word 文件會儲存在輸出位置的結果資料夾中；點擊「開啟輸出」即可找到。輸出位置可在「設定 → 檔案儲存」中更改。

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
5. **執行操作**：點擊對應的功能按鈕（如「導出 Markdown」、"生成 DOCX"等）
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

點擊視窗底部的設定圖示或按 **Ctrl+,** 開啟設定。左側導覽按**基本設定、輸入類型、內容處理、專項最佳化**分組；視窗較矮時可捲動側欄，只有寬度不足時才切換為頁面下拉選單。

**OCR 與圖片**管理共用辨識與 Markdown 圖文輸出；各輸入類型的 OCR 預設仍在對應的輸入頁，圖片轉檔與壓縮則在**傳入圖片**中。**Markdown 語法**分別說明共用輸出約定、輸入讀取規則與 Word 專用映射；部分來源及專項最佳化尚未套用全部語法設定。

**公文**與**發票**各有獨立設定頁。公文分為**文件 → Markdown**辨識預設和**Markdown → 文件**範本欄位格式化，資訊按鈕說明附件說明、抄送機關、日期及附註等規則。發票分別設定**PDF / OFD → Markdown**及**圖片 → Markdown**的預設。這些預設只決定相容檔案的初始選擇，不會自動判定公文或發票；按**套用／確定**儲存，**取消**捨棄未套用的變更。

### 快捷操作

- **拖拽外部文件**：直接拖入窗口即可導入
- **開啟輸出**：點擊狀態欄右側的「開啟輸出」操作打開輸出位置
- **右鍵範本項**：打開範本文件位置

---

## 🔧 命令列使用

使用命令列（CLI）可以在終端機檢查檔案、執行轉換，或編寫指令碼處理重複工作。與其他軟體整合時，請使用下文的 Machine Protocol 介面。

### 建議的自動化呼叫順序

編寫本機自動化指令碼時，可以依照以下順序檢查並執行轉換：

1. `inspect <file> [--json]`：先識別檔案的真實類別、格式與可執行動作。
2. `resources list formats --json`：查看可用的轉換格式，以及所需軟體是否已安裝。
3. `schema convert`：查看 `convert` 接受的參數、預設值和使用條件。
4. `convert <file> --to <fmt> (--output-dir <dir> | --output <path>) --dry-run --json`：檢查辨識到的格式、轉換方法和參數，不產生檔案。
5. `convert <file> --to <fmt> (--output-dir <dir> | --output <path>) ...`：確認後再執行正式轉換。

其他軟體或外掛透過 `serve --stdio` 使用 Machine Protocol v2：查詢可用操作、規劃工作，再執行。這是用於軟體整合的穩定介面；CLI 的 `--json` 供終端機指令碼讀取結果。需要開啟桌面介面或將檔案傳給 DocWen 時，使用 `gui open|activate|status`。詳見[命令列說明](../cli.md)和[整合協定](../specs/machine-protocol-v2.md)。

### 常用範例

傳入或產生 Markdown 的轉換統一在 `--output-dir DIR` 指定的父目錄下建立結果資料夾，即使只有一個輸出檔案；名稱包含原始檔名、時間戳記和本次傳入格式。其他僅產生單一檔案的轉換才使用 `--output PATH` 指定精確檔案路徑。

```bash
# Windows 打包版
DocWenCLI.exe inspect document.docx --json

# 查看 convert 接受的參數
DocWenCLI.exe schema convert

# 預演本次轉換會如何執行，但不實際寫出結果
DocWenCLI.exe convert report.docx --to md --output-dir exports --extract-img --ocr --dry-run --json

# 匯出 Word 為 Markdown（提取圖片 + OCR）
DocWenCLI.exe convert report.docx --to md --output-dir exports --extract-img --ocr

# Markdown 轉 Word（指定模板，並設定標題 + 正文合併模式）
DocWenCLI.exe convert document.md --to docx --output-dir exports --template template.docx.926a8bcb579f16e796662bff35edb7bb437aa718b7b7739b2461f4641128001e --heading-merge-mode punct_required

# Markdown 轉 Word，並在儲存前校對產生的 DOCX
DocWenCLI.exe convert document.md --to docx --output-dir exports --proofread --check typo --check punct

# 選擇圖片儲存方式和 OCR 文字位置
DocWenCLI.exe convert report.docx --to md --output-dir exports --extract-img --image-mode file --ocr --ocr-placement image_md

# 檢查可用轉換和所需軟體
DocWenCLI.exe doctor --json
DocWenCLI.exe resources list formats --json

# 文件校對
DocWenCLI.exe validate document.docx --check typo --check punct
DocWenCLI.exe validate input.md --check typo --check punct

# 原始碼 / uv 安裝
# inspect -> schema -> dry-run -> convert
# docwen inspect document.docx --json
# docwen schema convert
# docwen convert document.docx --to md --output-dir exports --dry-run --json
# docwen convert document.docx --to md --output-dir exports
```

### 常用命令與選項

下表只列出常用命令；完整命令請以 `docwen --help`（原始碼 / uv 安裝）或 `DocWenCLI --help`（打包版）為準。

| 命令/選項 | 說明 |
| --- | --- |
| `convert <file> --to <fmt> (--output-dir <dir> \| --output <path>)` | 轉換一個檔案。涉及 Markdown 時使用 `--output-dir` 儲存結果資料夾；其他單檔轉換使用 `--output` 指定檔案路徑。 |
| `convert <markdown> --to docx --output-dir <dir> --proofread [--check ...]` | Markdown→DOCX 轉換後校對；`--check` 必須與 `--proofread` 同時使用，省略 `--check` 時使用已設定的校對預設值。 |
| `convert <file> --to <fmt> (--output-dir <dir> \| --output <path>) --dry-run --json` | 檢查檔案格式、轉換方法和參數，不產生檔案。 |
| `schema convert` | 列出轉換參數、預設值和使用條件，方便指令碼讀取。 |
| `validate <file> --check ...` | 預設只檢查文字；加上 `--report <path>` 才會儲存校對結果。可選規則為 `typo/punct/symbol/sensitive/all/none`。 |
| `inspect <file> [--json]` | 查詢檔案類別/格式、建議動作，以及副檔名與內容不一致警告。 |
| `doctor --json` | 檢查安裝情況、可用轉換和所需軟體。 |
| `resources list formats --json` | 列出可用轉換格式、所需軟體和目前限制。 |
| `resources list optimizations --json` | 列出可用最佳化類型及其適用的轉換。 |
| `resources list templates` | 列出可用模板。 |
| `resources list numbering-schemes` | 列出可用編號方案。 |
| `gui open [ABSOLUTE_FILE] [--timeout <seconds>] [--json]` | 開啟 DocWen 視窗，並可同時傳入一個檔案。檔案路徑必須是絕對路徑。 |
| `--template <id>` | 請填寫 `resources list templates` 回傳的資源 ID，不能使用範本名稱或路徑。DOCX ID 用於 `docx/doc/odt/rtf/wps/pdf`，XLSX ID 用於 `xlsx/xls/ods/csv`。 |
| `--extract-img` / `--no-extract-img` / `--ocr` | `convert --to md` 的圖片提取與 OCR 選項。 |
| `--image-mode file\|base64\|embed\|omit` | 選擇匯出圖片的儲存或嵌入方式。 |
| `--ocr-placement image_md\|main_md` | 控制 OCR 文字寫入圖片配套 Markdown 或主 Markdown。 |
| `--heading-merge-mode punct_required\|always\|never` | 控制 `convert --to docx` 時「標題 + 正文」段落合併策略。 |
| `--optimization <id>` | 明確啟用某個最佳化配置（可用列表見 `resources list optimizations`）。 |
| `batch convert\|validate ... --jobs <n> [--continue-on-error]` | 批次處理控制。 |
| `--json` / `--quiet` / `--timing` | 在終端機指令碼中輸出 JSON、減少日誌，或顯示耗時。 |


## 📝 Markdown 語法約定

**需要識別 Obsidian 擴充語法嗎？** 新安裝 DocWen 時，四種選用擴充預設**關閉**：`structural_tables`（合併儲存格、多列表頭）、`captions_references`（Number Suite 題註與交互參照）、`extended_headings`（第七至九級標題）、`typed_endnotes`（類型化尾註）。

- **Markdown → Word**：在 **設定 → Markdown 語法 → 輸入** 啟用所需擴充，或套用輸入方向的 **Obsidian 擴充**預設。未啟用時，部分標記會保留為一般文字。
- **Word → Markdown**：另在**輸出**方向啟用對應擴充（或輸出預設），否則可能簡化結構、正規化語法並提示警告。

保留原始 Markdown 並檢查輸出的 DOCX。Word 編號由 DocWen 轉換選項決定，往返轉換不保證逐字還原原始標記及空白。參閱 [Markdown 相容規範](../specs/markdown-compatibility.md)。

### 標題級別映射

Markdown 標題與 Word 標題級別直接對應：
- 文檔的標題（title）和副標題（subtitle）放在YAML元數據中
- Markdown的 `# 一級標題` 對應 Word的"標題1"
- Markdown的 `## 二級標題` 對應 Word的"標題2"
- 第一至六級使用標準 Markdown 語法。第七至九級需先在 Markdown 輸入設定中啟用「第七至九級標題」，再使用七至九個開頭的 `#`

**提示**：如果您習慣用 Markdown 的一級標題（`#`）表示文檔標題（title），從二級標題（`##`）開始表示正文小標題（heading），可以在 Word 範本中將「標題1」樣式調整為文檔標題的外觀（如置中、粗體、較大字號），並在設定的序號方案中選擇「跳過一級標題編號」的方案。這樣，一級標題就能呈現為文檔標題的效果。

### 換行與分段

**基本規則**：一般 Markdown 內文的軟換行預設不會建立獨立的 Word 段落。需要逐行分段時，可在 **設定 → Markdown 語法** 啟用 **每行獨立成段**。清單、引言、程式碼等結構與明確硬換行仍依各自規則處理。

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
- 可在「設定 → 傳入文字 → Markdown 轉為文件 — 格式處理」中調整「標題 + 正文合併模式」和觸發符號。觸發符號留空時，「有標點才合併」模式不會合併任何段落
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
2. 選擇匯出選項，點擊「導出 Markdown」
3. 生成包含 YAML 元數據的 .md 文件

**支持的格式**：
- `.docx` - 標準 Word 文檔
- `.doc` - 自動生成 DOCX 後處理
- `.wps` - WPS 文檔自動轉換

**導出選項說明**：

| 選項 | 說明 |
|-----|------|
| **提取圖片** | 勾選後，將文檔中的圖片提取到輸出文件夾，MD文件中插入圖片連結 |
| **圖片文字識別** | 辨識圖片中的文字。依圖文輸出位置設定，儲存在每張圖片配套的 Markdown 中，或寫入主 Markdown 檔案。 |
| **針對文件類型優化** | 選擇適合文件類型的最佳化，例如從公文中擷取欄位；不啟用時使用標準轉換。 |
| **清除原有文件小標題序號** | 勾選後，移除小標題前的序號（如"一、""（一）""1."等），轉換為純標題文本 |
| **新增小標題序號** | 勾選後，根據標題層級自動添加序號（可在設定中配置序號方案） |

Word/WPS 多層次清單產生的標題編號在匯出時會保留。勾選「清除原有文件小標題序號」可移除這些編號，標題級別仍會保留。

### Markdown 轉 Word

1. 準備 .md 檔案；如需填寫標題或自訂欄位，可在開頭加入 YAML 資訊
2. 拖入程式窗口並選擇對應的 Word 範本
3. 將目標格式設為 DOCX，點擊「生成 DOCX」，填入範本並儲存文件

**轉換選項說明**：

| 選項 | 說明 |
|-----|------|
| **清除原有Markdown小標題序號** | 勾選後，移除小標題前的序號（如"一、""（一）""1."等），轉換為純標題文本 |
| **新增小標題序號** | 勾選後，根據標題層級自動添加序號（可在設定中配置序號方案） |

**注意**：如果文檔中有小標題和正文文本混合的段落，在MD文件中需要保持嚴格換行（參見上方"換行與分段"說明）。

### 範本樣式自動處理

轉換器在 Markdown → DOCX 轉換時會自動檢測和處理範本樣式：

#### 樣式分類

**段落樣式（Paragraph Style）**：應用到整個段落

| 樣式 | 檢測行為 | 缺少時新增 | 來源 |
|-----|---------|-----------|-----|
| 標題 (heading 1~9) | 檢測段落樣式 | 模板標題樣式 | Word 內置 |
| 代碼塊 | 檢測段落樣式 | Consolas 字體 + 灰色背景 | 本軟體定義 |
| 引用 (Quote 1~9) | 檢測段落樣式 | 灰色背景 + 左邊框 | 本軟體定義 |
| 公式塊 | 檢測段落樣式 | 公式專用樣式 | 本軟體定義 |
| 分隔線 (1~3) | 檢測段落樣式 | 底部邊框段落樣式 | 本軟體定義 |

**字符樣式（Character Style）**：應用到選中的文字

| 樣式 | 檢測行為 | 缺少時新增 | 來源 |
|-----|---------|-----------|-----|
| 行內代碼 | 檢測字符樣式 | Consolas 字體 + 灰色底紋 | 本軟體定義 |
| 行內公式 | 檢測字符樣式 | 公式專用樣式 | 本軟體定義 |

**表格樣式（Table Style）**：應用到整個表格

| 樣式 | 檢測行為 | 缺少時新增 | 來源 |
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
  - 新增受管理的樣式名稱由**一般 → 工作輸出語言**決定，與介面語言無關；切換介面語言不會重新命名範本中的既有樣式。

**使用建議**：在範本中自定義樣式後，轉換器會自動使用您的樣式；如果範本中沒有，會使用內置預設樣式。

### 表格文件處理

1. **Excel/CSV 轉 Markdown**：新增 `.xlsx` 或 `.csv` 檔案，點擊「導出 Markdown」。
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
3. 點擊「文件校對」按鈕
4. Word 文檔的校對結果以批註形式顯示在文檔中；Markdown 文件的校對結果輸出為結構化 JSON 報告

命令列使用 `validate <file>` 檢查檔案，預設不修改原始檔。加上 `--report <path>` 可儲存校對結果：Markdown 產生 JSON 報告，Word 及支援的舊版文件格式產生附註解的 DOCX。`--json` 用於把命令的執行結果列印為 JSON。

## 🛠️ 範本系統

### 使用現有範本

在**設定 → 範本**中統一管理 DOCX/XLSX 範本，可啟用、停用、排序及設定預設項目。內建範本唯讀，自訂範本儲存在使用者可寫目錄；MSIX 使用者無須修改 WindowsApps。

### 自定義範本

1. 選擇內建範本並**複製並編輯**，或匯入現有 DOCX/XLSX 檔案。
2. 在 Word、Excel 或 WPS 中修改自訂副本的內容、樣式與預留位置，規則見下文。
3. 儲存檔案，返回 DocWen 並**重新整理**，無須重新啟動。

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

- Markdown → Excel 會保留範本原有的合併儲存格。
- 對由連續整格 `{{↓欄位名}}` 組成的已知列式範本區域，支援根據 Markdown 表格中的顯式 `<` / `^` 合併標記還原矩形合併。
- 只有整個儲存格的內容為 `<` 或 `^`（忽略首尾空白）時才會辨識為合併標記；`\<`、`\^` 會保留為字面文字。
- 合併區域無法組成矩形，或與範本已有合併儲存格衝突時，會保留為一般文字並提示警告，範本結構保持不變。

**多表格數據合併**：如果 Markdown 中有多個表格使用相同的表頭名，數據會按順序合併後依次填充。

## 🔌 Obsidian 插件

配套 Obsidian 插件已在獨立倉庫發布，可與轉換器聯動使用：

### 核心特性

- **🚀 一鍵啟動** - 側邊欄圖標快速啟動轉換器
- **📂 自動傳遞** - 自動傳遞當前打開的文件路徑
- **🔄 使用已開啟的視窗** - 轉換器正在執行時，直接開啟新檔案，無須重新啟動

### 使用方式

DocWen Assistant 3.3 支援 Windows 與 Linux 桌面端；Windows Microsoft Store 安裝版可使用自動偵測，Linux 使用手動選取的相容解壓縮套件。Assistant 目前不支援 macOS。

1. **首次點選** → 啟動轉換器並開啟目前的檔案
2. **再次點選（有開啟的檔案）** → 以該檔案取代目前的輸入檔案（單一檔案模式）
3. **再次點選（沒有開啟的檔案）** → 將轉換器視窗切換到前景

### 安裝方法

安裝 [DocWen Assistant](https://github.com/ZHYX91/obsidian-docwen-assistant) 前，需要先安裝相容的 DocWen。Assistant 3.3 要求 DocWen 0.17.0 或更新版本；外掛不會代為安裝主程式。請依發行說明選擇相容版本。開發整合時，這些版本使用 Machine Protocol v2 通訊，並以 Artifact Bundle v3 回傳轉換結果。

## 🔌 OpenClaw（插件 + Skill）

[docwen-openclaw](https://github.com/ZHYX91/docwen-openclaw) 讓 OpenClaw 呼叫本機 DocWen 轉換和校對檔案。需要另外安裝 DocWen，並依外掛發行說明選擇相容版本。外掛 3.0 使用 Machine Protocol v2 和 Artifact Bundle v3。請下載正式發行的安裝檔；原始碼中的版本號不代表該版本已可下載。

## ❓ 常見問題

### 轉換失敗怎麼辦？

- 檢查文件是否被其他程式占用
- 確認文件格式正確
- 開啟「設定 → 日誌」，找到實際日誌檔案並查看錯誤詳情

### 範本不顯示？

- 在**設定 → 範本**中確認範本已啟用，且符合目前輸出格式。
- 確認 DOCX/XLSX 檔案可以正常開啟。
- 外部修改後重新整理；新增檔案請使用匯入功能。

### 校對功能不工作？

- 確認文檔為 .docx 或 .md 格式
- 檢查文檔是否包含可編輯文本
- 在設定中確認校對規則已啟用

### 輸出格式不符合預期？

- 開啟**設定 → 範本**：內建範本先複製並編輯，自訂範本可直接編輯。
- 在 Word、Excel 或 WPS 中調整樣式，儲存後重新整理清單。
- 產生新檔案前選擇該自訂範本。

### Excel轉Markdown後公式儲存格為空？

DocWen 匯出的是公式儲存格**上次儲存的計算結果**。如果檔案裡沒有儲存這個結果，匯出的儲存格就會為空；由其他程式產生的檔案，或修改後沒有重新計算並儲存的檔案，都可能出現這種情況。

**解決方法**：
1. 在 Excel 中打開文件
2. 等待公式計算完成
3. 保存文件
4. 重新轉換

## 🔒 安全特性

DocWen 在本機處理檔案，不提供檔案上傳、遙測、線上下載或網路服務功能。程式會限制自身 Python 元件的網路存取；如果這項防護無法啟用，DocWen 不會繼續執行。

另外啟動的 Word/Office、WPS、LibreOffice 等程式及相關輔助程序，仍遵循各自和系統的網路設定，不受這項限制。處理敏感檔案時，也請檢查這些程式的設定。

防護的實作和限制請參閱[網路防護說明](../architecture.md)。如果程式無法啟動或轉換失敗，可查閱[疑難排解](../maintenance/troubleshooting.md)。

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
