[English](README.md) | [简体中文](README_zh-CN.md) | [繁體中文](README_zh-TW.md) | [Deutsch](README_de-DE.md) | [Français](README_fr-FR.md) | [Русский](README_ru-RU.md) | [Português](README_pt-BR.md) | [日本語](README_ja-JP.md) | [한국어](README_ko-KR.md) | [Español](README_es-ES.md) | [Tiếng Việt](README_vi-VN.md)

# DocWen

A document and chart format conversion tool supporting Word/Markdown/Excel bidirectional conversion. Runs completely locally, ensuring data security and reliability.

## 📖 Project Background

This software was originally designed for the daily work of the printing office to solve the following problems:
- Document formats sent by various departments are chaotic and need to be organized into standardized formats.
- There are many types of documents, each with different fixed format requirements.
- Needs to run offline, adapting to intranet environments and legacy equipment.

**Design Philosophy**: This software is positioned as a lightweight, fool-proof tool. While it cannot compare with professional tools like LaTeX or Pandoc in terms of professionalism and functional completeness, it excels in zero learning cost and out-of-the-box usability, making it suitable for daily office scenarios where format requirements are not extremely strict.

## ✨ Core Features

- **📄 Document Format Conversion** - Bidirectional Word ↔ Markdown conversion. Supports mathematical formula conversion, and bidirectional separator conversion (Markdown's three types of separators vs. Word's page breaks, section breaks, and horizontal lines). Supports formats like DOCX/DOC/WPS/RTF/ODT.
- **📊 Spreadsheet Format Conversion** - Bidirectional Excel ↔ Markdown conversion. Supports XLSX/XLS/ET/ODS/CSV formats. Includes table summary tools.
- **📑 PDF and Layout Files** - PDF/XPS/OFD to Markdown or DOCX conversion. Supports PDF merging, splitting, and other operations.
- **🖼️ Image Processing** - Supports bidirectional conversion and compression of JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC formats.
- **🔍 OCR Text Recognition** - Integrated RapidOCR to extract text from images and PDFs.
- **✏️ Text Proofreading** - Checks for typos, punctuation, symbols, and sensitive words based on custom dictionaries. Rules can be edited in the settings interface.
- **📝 Template System** - Flexible template mechanism supporting custom document and report formats.
- **💻 Dual Mode Operation** - Graphical User Interface (GUI) + Command Line Interface (CLI).
- **🔒 Completely Local Operation** - Runs offline, ensuring data security with built-in network isolation mechanisms.
- **🔗 Single Instance Operation** - Automatically manages program instances and supports integration with the accompanying Obsidian plugin.

## 📸 Screenshots

| Batch | Markdown |
| --- | --- |
| ![Batch panel](assets/screenshots/batch.png) | ![Markdown panel](assets/screenshots/markdown.png) |

| Document | Spreadsheet |
| --- | --- |
| ![Document panel](assets/screenshots/document.png) | ![Spreadsheet panel](assets/screenshots/spreadsheet.png) |

| Image | Layout |
| --- | --- |
| ![Image panel](assets/screenshots/image.png) | ![Layout panel](assets/screenshots/layout.png) |

Changelog: see [doc/CHANGELOG.md](doc/CHANGELOG.md)

## 🚀 Quick Start

### Launch Program

On the Windows packaged release, double-click `DocWen.exe` to start the graphical interface. If installed from source / pip, run `docwen-gui`.

### Quick Start Guide

1.  **Prepare a Markdown File**:

    ```markdown
    ---
    title: Test Document
    ---
    
    ## Test Title
    
    This is the test body content.
    ```

2.  **Drag and Drop Conversion**:
    - Launch the program.
    - Drag the `.md` file into the window.
    - Select a template.
    - Click "Convert to DOCX".

3.  **Get Results**:
    - A standardized Word document will be generated in the same directory.

**Tip**: You can use the sample files in the `samples/` directory to quickly try out the software's features.

## 📝 Markdown Syntax Conventions

### Heading Level Mapping

To make it easier for colleagues without background knowledge to remember, the Markdown headings in this software correspond **one-to-one** with Word headings:
- Document title and subtitle are placed in YAML metadata.
- Markdown `# Heading 1` corresponds to Word "Heading 1".
- Markdown `## Heading 2` corresponds to Word "Heading 2".
- And so on, supporting up to 9 levels of headings.

**Tip**: If you prefer using Markdown's first-level heading (`#`) as the document title, starting from second-level headings (`##`) for body headings, you can style "Heading 1" in the Word template to look like a document title (e.g., centered, bold, larger font size), and select a numbering scheme that skips first-level heading numbering in the settings. This way, your first-level headings will appear as document titles.

### Line Breaks and Paragraphs

**Basic Rule**: Every non-empty line is treated as a separate paragraph by default.

**Mixed Paragraphs**: When a subheading needs to be mixed with the body text in the same paragraph, the following conditions must be met:
1.  The subheading ends with a terminating punctuation mark (supports multilingual punctuation, including periods, question marks, exclamation marks, and other common terminating punctuation).
2.  The body text is located on the **immediate next line** of the subheading.
3.  The body text line cannot be a special Markdown element (such as headings, code blocks, tables, lists, quotes, formula blocks, separators, etc.).

**Example**:
```markdown
## I. Work Requirements.
This meeting requires all units to earnestly implement...
```
The above two lines will be merged into the same paragraph, where "I. Work Requirements." keeps the subheading format, and "This meeting..." keeps the body text format.

**Note**:
- There cannot be an empty line between the subheading and the body text; otherwise, they will be recognized as separate paragraphs.
- If the subheading does not end with a punctuation mark and has no empty line before the body text, the body text will be merged into the heading line with adjusted formatting.

### Bidirectional Separator Conversion

Supports bidirectional conversion between Markdown separators and Word page breaks/section breaks/horizontal lines:

-   **DOCX → MD**: Word page breaks, section breaks, and horizontal lines are automatically converted to Markdown separators.
-   **MD → DOCX**: Markdown `---`, `***`, `___` are automatically converted to corresponding Word elements.
-   **Configurable**: Specific mapping relationships can be customized in the settings interface.

### Image Embeds and Sizing

Supports Obsidian/Wiki and standard Markdown image embeds, with optional sizing (px):

```markdown
![[image.png]]
![[image.png|300]]
![[image.png\|300]]
![alt](image.png =300x200)
![alt](image.png =300x)
![alt|300](image.png)
```

- No size: uses the original image size, capped by available page/cell width
- With size: allows upscaling, still capped by available width
- Image-only paragraph: uses the Image paragraph style (centered, single spacing)

## 📖 Detailed Usage Guide

### Word to Markdown

1.  Drag the `.docx` file into the program window.
2.  The program automatically analyzes the document structure.
3.  Generates a `.md` file containing YAML metadata.

**Supported Formats**:
-   `.docx` - Standard Word document.
-   `.doc` - Automatically converted to DOCX for processing.
-   `.wps` - WPS document automatically converted.

**Export Options**:

| Option | Description |
| :--- | :--- |
| **Extract Images** | If checked, images in the document are extracted to the output folder, and image links are inserted into the MD file. |
| **Image OCR** | If checked, performs OCR on images and creates an image `.md` file (containing recognized text). |
| **Clean Subheading Numbers** | If checked, removes numbers before subheadings (e.g., "一、", "（一）", "1.", etc.) and converts them to pure title text. |
| **Add Subheading Numbers** | If checked, automatically adds numbers based on heading levels (numbering scheme can be configured in settings). |

### Markdown to Word

1.  Prepare a `.md` file with a YAML header.
2.  Drag it into the program window and select the corresponding Word template.
3.  The program automatically fills the template and generates the document.

**Conversion Options**:

| Option | Description |
| :--- | :--- |
| **Clean Subheading Numbers** | If checked, removes numbers before subheadings. |
| **Add Subheading Numbers** | If checked, automatically adds numbers based on heading levels. |

**Note**: If there are paragraphs where subheadings and body text are mixed, strict line breaks must be maintained in the MD file (see "Line Breaks and Paragraphs" above).

### Automatic Template Style Processing

The converter automatically detects and processes template styles during Markdown → DOCX conversion:

#### Style Classification

**Paragraph Style**: Applied to the entire paragraph.

| Style | Detection Behavior | Injection when Missing | Source |
| :--- | :--- | :--- | :--- |
| Heading (1~9) | Detects paragraph style | Template heading styles | Word Built-in |
| Code Block | Detects paragraph style | Consolas font + Gray background | Defined by Software |
| Quote (1~9) | Detects paragraph style | Gray background + Left border | Defined by Software |
| Formula Block | Detects paragraph style | Formula specific style | Defined by Software |
| Separator (1~3) | Detects paragraph style | Bottom border paragraph style | Defined by Software |

**Character Style**: Applied to selected text.

| Style | Detection Behavior | Injection when Missing | Source |
| :--- | :--- | :--- | :--- |
| Inline Code | Detects character style | Consolas font + Gray shading | Defined by Software |
| Inline Formula | Detects character style | Formula specific style | Defined by Software |

**Table Style**: Applied to the entire table.

| Style | Detection Behavior | Injection when Missing | Source |
| :--- | :--- | :--- | :--- |
| Three-Line Table | User config priority | Three-line table style definition | Defined by Software |
| Grid Table | User config priority | Grid table style definition | Defined by Software |

**Numbering Definition**: Used for list formats.

| Type | Detection Behavior | Handling when Missing |
| :--- | :--- | :--- |
| List Numbering | Scans existing ordered/unordered list definitions in template | Uses decimal/bullet preset |

#### Style Name Internationalization

-   **Word Built-in Styles** (heading 1~9):
    -   Style names use Word standard English names (e.g., `heading 1`).
    -   Word automatically displays localized names based on system language (e.g., "标题 1" on Chinese systems).
-   **Software Defined Styles** (Code Block, Quote, Formula, Separator, Table, etc.):
    -   Injects corresponding language style names based on the software's interface language setting.
    -   Chinese Interface: Injects "代码块", "引用 1", "三线表", etc.
    -   English Interface: Injects "Code Block", "Quote 1", "Three Line Table", etc.

**Suggestion**: After customizing styles in the template, the converter will automatically use your styles; if not present in the template, it will use built-in preset styles.

### Spreadsheet File Processing

1.  **Excel/CSV to Markdown**: Drag `.xlsx` or `.csv` files to automatically convert to Markdown tables.
2.  **Markdown to Excel**: Prepare an MD file and select an Excel template for conversion.

**Supported Formats**:
-   `.xlsx` - Standard Excel document.
-   `.xls` - Automatically converted to XLSX for processing.
-   `.et` - WPS spreadsheet automatically converted.
-   `.csv` - CSV text table.

### Text Proofreading Function

The program provides four customizable proofreading rules:

1.  **Punctuation Pairing Check** - Detects if paired punctuation like parentheses and quotes match.
2.  **Symbol Proofreading** - Detects mixed use of Chinese and English punctuation.
3.  **Typo Check** - Checks for common typos based on a custom dictionary.
4.  **Sensitive Word Detection** - Detects sensitive words based on a custom dictionary.

**Custom Dictionaries**: Visually edit typo and sensitive word dictionaries in the "Settings" interface.

**Usage**:
1.  Drag the Word document to be proofread into the program.
2.  Check the required proofreading rules.
3.  Click the "Text Proofreading" button.
4.  Proofreading results are displayed as comments in the document.

## 🛠️ Template System

### Using Existing Templates

The program comes with various templates, including multilingual versions. You can select and use them as needed. Template files are located in the `templates/` directory.

### Custom Templates

1.  Create a template file using Word or WPS.
2.  Refer to existing templates and insert placeholders like `{{Title}}`, `{{DocumentNumber}}`, etc., where filling is needed.
3.  In the template, built-in Heading 1 ~ Heading 5 styles need to be manually modified.
4.  Save the template to the `templates/` directory.
5.  Restart the program, and the new template will be automatically loaded.

You can also copy an existing template, modify it, and rename it.

### Placeholder Usage

#### Word Template Placeholders

**YAML Field Placeholders**: Use `{{Field Name}}` format in the template, which will be replaced by the corresponding value in the Markdown file's YAML header during conversion.

| Placeholder | Description |
| :--- | :--- |
| `{{Title}}` | Document title (Retrieval rules see below) |
| `{{Body}}` | Markdown body content insertion position |
| Others | Supports any custom field |

**Title Retrieval Priority**:

| Priority | Source | Description |
| :--- | :--- | :--- |
| 1 | YAML `Title` field | Highest priority |
| 2 | YAML `aliases` field | Takes the first element of the list, or string value |
| 3 | Filename | Filename without `.md` extension |

**Multilingual Support**: The title and body placeholders support multiple languages, e.g., title can be `{{title}}`, `{{标题}}`, `{{Titel}}`, etc., body can be `{{body}}`, `{{正文}}`, `{{Inhalt}}`, etc.

#### Excel Template Placeholders

Excel templates support three types of placeholders:

**1. YAML Field Placeholder** `{{Field Name}}`

Used to fill a single value from the Markdown file's YAML header:

```markdown
---
ReportName: 2024 Annual Sales Statistics
Unit: Sales Dept
---
```

`{{ReportName}}`, `{{Unit}}` in the template will be replaced with corresponding values. The title field also follows the priority rules.

**2. Column Fill Placeholder** `{{↓Field Name}}`

Extracts data from the Markdown table and fills **downwards** row by row starting from the placeholder position:

```markdown
| ProductName | Quantity |
|:--- |:--- |
| Product A | 100 |
| Product B | 200 |
```

`{{↓ProductName}}` in the Excel template will be replaced by "Product A", and the next row will be filled with "Product B".

**3. Row Fill Placeholder** `{{→Field Name}}`

Extracts data from the Markdown table and fills **rightwards** column by column starting from the placeholder position:

```markdown
| Month |
|:--- |
| Jan |
| Feb |
| Mar |
```

`{{→Month}}` in the Excel template will be filled with "Jan", "Feb", "Mar" sequentially to the right.

**Merged Cell Handling**: The program automatically skips non-first cells of merged cells to ensure correct data filling.

**Multi-table Data Merge**: If there are multiple tables in Markdown using the same header name, data will be merged in order and filled sequentially.

## 🖥️ Graphical Interface Usage

Most users use this software through the graphical interface. Here is the detailed operation guide.

### Interface Overview

The program uses an **adaptive three-column layout**:

| Area | Description | Display Timing |
| :--- | :--- | :--- |
| **Center Column (Main Area)** | File drag-and-drop area, operation panel, status bar | Always shown |
| **Right Column** | Template selector / Format conversion panel | Automatically expands after selecting a file |
| **Left Column** | Batch file list (grouped by type) | Shown when switching to batch mode |

### Basic Operation Flow

1.  **Launch Program**: Double-click `DocWen.exe` (Windows packaged release) or run `docwen-gui`.
2.  **Import File**:
    -   Method 1: Drag and drop files directly into the window.
    -   Method 2: Click the "Add" button in the drag-and-drop area to select files.
3.  **Select Template** (if conversion is needed): The right template panel expands automatically; select a suitable template.
4.  **Configure Options**: Check the required conversion/export options in the operation panel.
5.  **Execute Operation**: Click the corresponding function button (e.g., "Export MD", "Convert to DOCX", etc.).
6.  **View Result**: The status bar shows progress and results; click the 📍 icon to locate the output file.

### Single File Mode vs. Batch Mode

The program supports two processing modes, switchable via the toggle button in the file drag-and-drop area:

**Single File Mode** (Default):
-   Process one file at a time.
-   Simple interface, suitable for daily use.

**Batch Mode**:
-   Import multiple files simultaneously.
-   Left column shows categorized file list (grouped by document/spreadsheet/image, etc.).
-   Supports batch adding, removing, and sorting.
-   Clicking a file in the list switches the current operation target.

### Operation Panel Functions

The operation panel automatically adjusts available options based on file type:

| File Type | Available Operations |
| :--- | :--- |
| Word Document | Export MD, Convert PDF, Text Proofreading, OCR |
| Markdown | Convert DOCX, Convert PDF |
| Excel Spreadsheet | Export MD, Convert PDF, Table Summary |
| PDF File | Export MD, Merge, Split, OCR |
| Image File | Format Conversion, Compression, OCR |

### Settings Interface

Click the ⚙️ button in the bottom right corner of the window to open settings:

-   **General**: Interface theme, language, window opacity.
-   **Conversion**: Default values for various conversion options.
-   **Output**: Default output directory, file naming rules.
-   **Proofread**: Edit typo and sensitive word dictionaries.
-   **Style**: Code block, quote, table style configurations.

### Shortcuts

-   **Drag External File**: Drag directly into the window to import.
-   **Double-click Status Bar Result**: Quickly open the output file directory.
-   **Right-click Template Item**: Open template file location.

---

## 🔧 Command Line Usage

In addition to the GUI, the program provides a Command Line Interface (CLI), suitable for automation scripts and batch processing scenarios.

### Running Modes

-   **CLI Mode**: Use subcommands (e.g. `convert`, `validate`) for automation scripts and batch processing.

### Common Examples

```bash
# Packaged release (Windows)
DocWenCLI.exe convert document.docx --to md

# Export Word to Markdown (Extract Images + OCR)
DocWenCLI.exe convert report.docx --to md --extract-img --ocr

# Markdown to Word (Specify Template)
DocWenCLI.exe convert document.md --to docx --template "Template Name"

# Batch Conversion (Skip confirmation, continue on error)
DocWenCLI.exe convert *.docx --to md --batch --yes --continue-on-error

# Document Proofreading
DocWenCLI.exe validate document.docx --check typo --check punct

# PDF Merge/Split
DocWenCLI.exe merge-pdfs *.pdf
DocWenCLI.exe split-pdf report.pdf --pages "1-3,5,7-10"

# From source / pip
docwen convert document.docx --to md
docwen convert report.docx --to md --extract-img --ocr
```

### Main Commands & Options

| Command / Option | Description |
| :--- | :--- |
| `convert <files...> --to <fmt>` | Convert files to target format (including `md`) |
| `validate <files...> --check ...` | Proofread documents (`--check typo/punct/symbol/sensitive/all/none`) |
| `merge-pdfs <files...>` | Merge PDF/OFD/XPS files |
| `split-pdf <file> --pages ...` | Split PDF by page ranges |
| `merge-tables <files...> --mode row|col|cell` | Merge spreadsheet tables |
| `merge-images-to-tiff <files...>` | Merge images into TIFF |
| `md-numbering <files...>` | Process Markdown heading numbering |
| `templates list [--for docx|xlsx]` | List available templates |
| `optimizations list [--scope ...]` | List available optimization types |
| `formats list [--for-source document|spreadsheet|layout|image|markdown]` | List available target formats |
| `inspect <file>` | Inspect file category/format and supported actions |
| `--template <name>` | Template name (used with `convert`) |
| `--extract-img` / `--no-extract-img` / `--ocr` | Options for `convert --to md` |
| `--optimize-for <id>` | Enable optimization explicitly (e.g. `gongwen`, `invoice_cn`) |
| `--batch` / `--jobs` / `--continue-on-error` | Batch processing controls |
| `--json` | Output result in JSON format |
| `--quiet` / `-q` | Quiet mode, reduce output |
| `--lang` | Switch UI language (affects help/messages) |

## 🔌 Obsidian Plugin

The project includes a matching Obsidian plugin to work in tandem with the converter:

### Core Features

-   **🚀 One-Click Launch** - Sidebar icon to quickly launch the converter.
-   **📂 Automatic Handover** - Automatically passes the currently open file path.
-   **🔄 Single Instance Management** - Automatically sends file if the program is already running, no need to restart.
-   **💪 Crash Recovery** - Automatically detects process status and automatically cleans up residual files.

### Working Principle

The plugin interacts with the converter via file system-based IPC:

1.  **First Click** → Launch converter and pass current file.
2.  **Click Again (With File)** → Replace with new file (Single File Mode).
3.  **Click Again (No File)** → Activate converter window.

### Installation

The plugin has been released to a separate repository. Please visit [docwen-obsidian](https://github.com/ZHYX91/docwen-obsidian) for installation instructions and the latest version.

## ❓ FAQ

### What if conversion fails?

-   Check if the file is occupied by another program.
-   Confirm the file format is correct.
-   Check error logs in the `logs/` directory.

### Template not showing?

-   Confirm template files are in the `templates/` directory.
-   Check if the template file is corrupted.
-   Restart the program to reload templates.

### Proofreading function not working?

-   Confirm the document is in .docx format.
-   Check if the document contains editable text.
-   Confirm proofreading rules are enabled in settings.

### Output format not as expected?

-   The program generates documents based on template styles. To adjust output format, modify the style definitions in the template file directly.
-   Template files are located in the `templates/` directory.
-   After modifying template styles, all documents converted using that template will apply the new styles.

### Formula cells are empty after Excel to Markdown conversion?

This is expected behavior. The program reads the **cached values** of cells rather than the formulas themselves.

**Technical Reason**:
-   In Excel files, formula cells store both the formula and the last calculated result (cached value).
-   The program uses `data_only=True` mode, which only retrieves cached values.
-   If the file has never been opened in Excel (e.g., generated by a program), or was edited but not re-saved, the cached value will be empty.

**Solution**:
1.  Open the file in Excel.
2.  Wait for formula calculation to complete.
3.  Save the file.
4.  Convert again.

## 🔒 Security Features

-   **Completely Local Operation**: All processing is done locally, no network dependency.
-   **Network Isolation**: Built-in network isolation mechanism prevents data leakage.
-   **No Data Upload**: User files are never uploaded to any server.
-   **Strict Security Mode**: Enabled by default; the app exits if security checks fail. See [doc/技术文档.md](doc/技术文档.md).

## 📜 License

This project is licensed under the **GNU Affero General Public License v3.0 (AGPL-3.0)**.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

-   This project uses PyMuPDF (licensed under AGPL-3.0), so the entire project is also licensed under AGPL-3.0.
-   You are free to use, modify, and distribute this software.
-   If you modify this software and provide services over a network, you must provide the modified source code to users.
-   For detailed license information, please see the [LICENSE](LICENSE) file.

### Contact

-   **GitHub**: https://github.com/ZHYX91/docwen
-   **Contact Author**: zhengyx91@hotmail.com

---

**Author**: ZhengYX
