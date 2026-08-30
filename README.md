# DocWen

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md)

A document and chart format conversion tool supporting Word/Markdown/Excel bidirectional conversion. Runs completely locally, ensuring data security and reliability.

## 📖 Project Background

This software was originally designed for the daily work of the printing office to solve the following problems:
- Document formats sent by various departments are chaotic and need to be organized into standardized formats.
- There are many types of documents, each with different fixed format requirements.
- Needs to run offline, adapting to intranet environments and legacy equipment.

**Design Philosophy**: This software is positioned as a lightweight, fool-proof tool. While it cannot compare with professional tools like LaTeX or Pandoc in terms of professionalism and functional completeness, it excels in zero learning cost and out-of-the-box usability, making it suitable for daily office scenarios where format requirements are not extremely strict.

## ✨ Core Features

- **📄 Document Format Conversion** - Bidirectional Word ↔ Markdown conversion. Supports mathematical formula conversion, bidirectional separator conversion (Markdown's three types of separators vs. Word's page breaks, section breaks, and horizontal lines), and restoring explicit Markdown table `<` / `^` markers to rectangular Word table merges. Supports formats like DOCX/DOC/WPS/RTF/ODT.
- **📊 Spreadsheet Format Conversion** - Bidirectional Excel ↔ Markdown conversion. Supports XLSX/XLS/ET/ODS/CSV/TSV formats, configurable merged-cell export strategies (`fill / empty / marker`), table summary tools, and the template placeholder surface described below.
- **📑 PDF and Layout Files** - PDF/XPS/OFD to Markdown or DOCX conversion. Supports PDF merging, splitting, and other operations.
- **🖼️ Image Processing** - Supports bidirectional conversion and compression of JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC formats.
- **📥 Other Format Import** - Supports one-way conversion of HTML/MHTML/ENEX/EPUB/PPTX/PPT to Markdown.
- **🔍 OCR Text Recognition** - Integrated RapidOCR to extract text from images and PDFs.
- **✏️ Text Proofreading** - Checks for typos, punctuation, symbols, and sensitive words based on custom dictionaries. Supports Word (.docx) and Markdown (.md) files. Rules can be edited in the settings interface.
- **📝 Template System** - Flexible template mechanism supporting custom document and report formats.
- **💻 Dual Mode Operation** - Graphical User Interface (GUI) + Command Line Interface (CLI).
- **🔒 Local processing with dependency egress protection** - Conversion does not depend on online services. While DocWen runs, its Python process blocks DNS and IPv4/IPv6 use by in-process dependencies; externally launched Office suites keep their own system network policy.
- **🔗 Single Instance Operation** - Automatically manages program instances and supports integration with the accompanying Obsidian plugin.

## 📸 Screenshots

| Main window | Batch |
| --- | --- |
| ![Main window](docs/assets/screenshots/main-light.png) | ![Batch panel](docs/assets/screenshots/batch-light.png) |

| Document | Spreadsheet |
| --- | --- |
| ![Document panel](docs/assets/screenshots/conversion-document-light.png) | ![Spreadsheet panel](docs/assets/screenshots/conversion-spreadsheet-light.png) |

| Image | Layout |
| --- | --- |
| ![Image panel](docs/assets/screenshots/conversion-image-light.png) | ![Layout panel](docs/assets/screenshots/conversion-layout-light.png) |

Changelog: see [CHANGELOG.md](docs/CHANGELOG.md)

## 🚀 Quick Start

### Installation from Source

**Prerequisites**: Python 3.12

**0.9 release boundary**: The [0.9.1 Release](https://github.com/ZHYX91/docwen/releases/tag/0.9.1) publishes one Windows x64 GUI+CLI package and separate
Ubuntu 24.04 x64 GUI+CLI and CLI-only packages. Other Linux
distributions and macOS remain source/development paths and are not implied by the Ubuntu package.

**Option 1: Using uv (Recommended)**

Install `uv 0.12.0`, then:

```bash
git clone https://github.com/ZHYX91/docwen.git
cd docwen
uv sync --frozen --all-extras
```

DocWen 0.9's source/test/build contract is the checked-in lock with exactly `uv 0.12.0`.
`pip install -e` is unsupported because pip cannot apply the repository's scoped dependency exclusion.

### Launch Program

On the Windows packaged release, double-click `DocWen.exe` to start the graphical interface. On
Ubuntu 24.04 x64, extract `DocWen-0.9.1-linux-x64.tar.gz` and run `./DocWen`; the companion
`DocWenCLI-0.9.1-linux-x64.tar.gz` is the CLI-only package. These assets are installable from the
immutable 0.9.1 Release. If installed from source, run:

```bash
docwen-gui  # GUI mode
docwen      # CLI mode
```

### macOS Notes

**Current limitation**: On macOS, `convert`, `validate`, `number`, `merge`, and `split` are currently
unavailable. The notes below only document optional dependencies for development experiments.

**LibreOffice support (Optional)**

To convert legacy formats like `.doc` and `.xls`, install LibreOffice:  
Download: https://www.libreoffice.org/download/

**HEIC image support (Optional)**

To process HEIC/HEIF images:

```bash
brew install libheif
pip install pillow-heif
```

### Linux GUI Prerequisites

**Supported package target**: DocWen 0.9 supports the packaged GUI and CLI on Ubuntu 24.04 x64.
These prerequisites do not extend that support claim to another Linux distribution or architecture.

- Desktop environment installed (GNOME, KDE, XFCE, etc.)
- GUI uses PySide6 (Qt6) and no longer depends on Python Tk. If startup fails due to missing system libraries, install the Qt runtime dependencies indicated by the error (commonly OpenGL/X11 related).
- For headless Ubuntu 24.04 systems, use `DocWenCLI` from the CLI-only archive instead of the GUI.

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
6.  **View Result**: The status bar shows progress and results; click the "Open Output" action on the right to open the output location.

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
| Markdown | Convert DOCX, Convert PDF, Text Proofreading |
| Excel Spreadsheet | Export MD, Convert PDF, Table Summary |
| PDF File | Export MD, Merge, Split, OCR |
| Image File | Format Conversion, Compression, OCR |
| HTML/EPUB/PPTX etc. | Export MD |

### Settings Interface

Click the "Settings" button in the operation header to open settings:

Settings are organized into tabs: **General**, **Text**, **Proofread**, **Document**, **Spreadsheet**, **Image**, **Layout**, **Link**, **Formatting**, **Output**, **Export**, **Logging**, **Other**.

### Shortcuts

-   **Drag External File**: Drag directly into the window to import.
-   **Open Output**: Click the "Open Output" action on the right side of the status bar to open the output location.
-   **Right-click Template Item**: Open template file location.

---

## 🔧 Command Line Usage

In addition to the GUI, DocWen provides a Command Line Interface (CLI) for automation scripts, batch processing, and external integrations.

### Recommended Automation Flow

For scripts, agents, or plugin integrations, use this order:

1. `inspect <file> [--json]`: detect the real file category, format, and supported actions first.
2. `resources list formats --json`: read the loaded Runtime's available routes and dependency gates.
3. `schema convert`: read the machine-readable conversion contract and conditional rules.
4. `convert <file> --to <fmt> --output <path> --dry-run --json`: preview detection, normalization, and routing without writing outputs.
5. `convert <file> --to <fmt> --output <path> ...`: run the actual conversion after the preview is acceptable.

### Common Examples

```bash
# Packaged release (Windows)
DocWenCLI.exe inspect document.docx --json

# Export the conversion contract for scripts / agents
DocWenCLI.exe schema convert

# Preview how the conversion would run without writing files
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --ocr --dry-run --json

# Export Word to Markdown (extract images + OCR)
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --ocr

# Markdown to Word (select a template and heading/body merge mode)
DocWenCLI.exe convert document.md --to docx --output document.docx --template template.docx.f1eeb0a008ce3eae0619ecae6e185ab132a3ee0abdd382c8863481d9af1dc77f --heading-merge-mode punct_required

# Control Markdown image export mode and OCR placement
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --image-mode file --ocr --ocr-placement image_md

# Chinese official-document optimization selected by public resource ID
DocWenCLI.exe convert report.docx --to md --output report.md --optimization gongwen

# Chinese invoice optimization selected by public resource ID
DocWenCLI.exe convert invoice.pdf --to md --output invoice.md --optimization invoice_cn

# Check runtime capability summary and dependency gates
DocWenCLI.exe doctor --json
DocWenCLI.exe resources list formats --json

# Read-only proofreading; add --report only when an output is wanted
DocWenCLI.exe validate document.docx --check typo --check punct --json
DocWenCLI.exe validate document.docx --check typo --check punct --report reviewed.docx
DocWenCLI.exe validate input.md --check typo --check punct --report report.json

# From source / uv
# inspect -> resources -> schema -> dry-run -> convert
# docwen inspect document.docx --json
# docwen resources list formats --json
# docwen schema convert
# docwen convert document.docx --to md --output document.md --dry-run --json
# docwen convert document.docx --to md --output document.md
```

### Common Commands & Options

The table below lists common commands only. For the full command surface, use `docwen --help` (source / uv) or `DocWenCLI --help` (packaged release).

| Command / Option | Description |
| --- | --- |
| `convert <file> --to <fmt> --output <path>` | Convert one file to one exact output path. Existing targets require explicit `--overwrite`. |
| `convert <file> --to <fmt> --output <path> --dry-run --json` | Preview detection, normalization, routing, and effective options without writing the output. |
| `convert <file> --to <fmt> --output <path> --optimization <id>` | Select an optimization by the public resource ID returned by `resources list optimizations`. |
| `validate <file> --check ... [--report <path>]` | Proofread DOCX, Markdown, or legacy Word-family content. The default is read-only; `--report` writes annotated DOCX for document input or JSON for Markdown. |
| `number markdown <file> --operation add\|remove (--output <path> \| --in-place)` | Explicitly add or remove Markdown heading numbering. |
| `merge pdf\|tables\|images <files...> --output <path>` | Run an explicit aggregate operation. |
| `split pdf <file> --pages <range> --output-dir <dir>` | Split selected PDF pages into an explicit directory. |
| `batch convert\|validate <files...>` | Run an explicit multi-file conversion or proofreading operation. |
| `schema convert` | Export the machine-readable conversion contract, defaults, conditions, and canonical keys. |
| `inspect <file> [--json]` | Inspect file category/format, recommended actions, and extension/content mismatch warnings. |
| `doctor --json` | Output diagnostics together with runtime capability summaries and dependency gates. |
| `resources list formats --json` | List canonical Runtime routes, dependency gates, availability, and limitations. |
| `resources list optimizations --json` | List typed optimization resources and their canonical route bindings. |
| `resources list templates [--target docx\|xlsx]` | List available templates. |
| `resources list numbering-schemes` | List available numbering schemes. |
| `--template <id>` | Exact canonical resource ID returned by `resources list templates`; display names, filenames, and paths are rejected. DOCX IDs apply to `docx/doc/odt/rtf/wps/pdf`, XLSX IDs to `xlsx/xls/ods/csv`. CSV without a template remains direct table export; CSV with an XLSX template uses the MD→XLSX template workbook → per-sheet CSV artifact chain. |
| `--extract-img` / `--no-extract-img` / `--ocr` | Image extraction and OCR options for `convert --to md`. |
| `--image-mode file|base64|embed|omit` | Control how images are emitted during Markdown export. |
| `--ocr-placement image_md|main_md` | Control whether OCR text is written to image-side Markdown or the main Markdown file. |
| `--heading-merge-mode punct_required|always|never` | Control the heading + body merge strategy for `convert --to docx`. |
| `--optimization <id>` | Select a typed optimization resource for conversion; internal Runtime action names are not public CLI options. |
| `--jobs` / `--continue-on-error` | Controls provided by explicit `batch` subcommands. |
| `--json` / `--quiet` / `--timing` | Structured output, reduced logs, and timing data for scripts or plugins. |


## 📝 Markdown Syntax Conventions

### Heading Level Mapping

To make it easier for colleagues without background knowledge to remember, the Markdown headings in this software correspond **one-to-one** with Word headings:
- Document title and subtitle are placed in YAML metadata.
- Markdown `# Heading 1` corresponds to Word "Heading 1".
- Markdown `## Heading 2` corresponds to Word "Heading 2".
- And so on, supporting all 9 Word heading levels. Levels 1-6 use standard CommonMark ATX syntax; levels 7-9 use
  DocWen's compatibility extension with seven to nine leading `#` markers.

**Tip**: If you prefer using Markdown's first-level heading (`#`) as the document title, starting from second-level headings (`##`) for body headings, you can style "Heading 1" in the Word template to look like a document title (e.g., centered, bold, larger font size), and select a numbering scheme that skips first-level heading numbering in the settings. This way, your first-level headings will appear as document titles.

### Line Breaks and Paragraphs

**Basic Rule**: Every non-empty line is treated as a separate paragraph by default.

**Mixed Paragraphs**: When a subheading needs to be mixed with the body text in the same paragraph (default mode: "Punctuation required"), the following conditions must be met:
1.  The subheading ends with one of the configured trigger characters. The exact default is `。：！？.:!?` (full-width or half-width period, colon, question mark, or exclamation mark).
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
- By default ("Punctuation required" mode), if the subheading does not end with a terminating punctuation mark, it will not merge with the next line even without an empty line.
- The trigger-character field is editable in Formatting settings. An empty value disables merging in "Punctuation required" mode. Commas, semicolons, enumeration commas, dashes, and ellipses are deliberately excluded from the default, but users may add them explicitly.
- You can change this in Settings → Formatting → "MarkDown to Document" → "Heading + body merge mode".

### Bidirectional Separator Conversion

Supports bidirectional conversion between Markdown separators and Word page breaks/section breaks/horizontal lines:

-   **DOCX → MD**: Word page breaks, section breaks, and horizontal lines are automatically converted to Markdown separators.
-   **MD → DOCX**: Markdown `---`, `***`, `___` are automatically converted to corresponding Word elements.
-   **Configurable**: Specific mapping relationships can be customized in the settings interface.

### Task List Items

Supports GFM task list items in both directions:

```markdown
- [ ] Todo
- [x] Done
```

-   **MD → DOCX**: Renders as a bullet list with `☐` / `☑` text prefix.
-   **DOCX → MD**: Converts list items starting with `☐` / `☑` / `☒` back to `- [ ]` / `- [x]`.
-   **Font note**: `☐`/`☑` may not render in some fonts. If needed, use fonts like "Segoe UI Symbol" in your Word template.

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

### Link Handling

Supports clickable links in Markdown -> DOCX:

```markdown
[Docwen](https://example.com)
[[Target]]
[[Target|Open target]]
<https://example.com>
<user@example.com>
```

- Markdown links and Wiki links are written as Word hyperlinks by default
- Wiki links resolve to local `file:///` links when the target file is found
- Angle-bracket autolinks support `https://...` and email `mailto:...`
- Bare URL autolinking is request-scoped for Markdown -> DOCX, defaults to off, and is enabled by `[non_embed_links].auto_link_bare_url` in `configs/link.toml`
- Markdown -> XLSX keeps the original link syntax instead of emitting DOCX hyperlink placeholders

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
| **Advanced Field Optimization** | If checked, extracts richer structured metadata; otherwise uses simplified mode with only title and subtitle fields. |
| **Clean Subheading Numbers** | If checked, removes numbers before subheadings (e.g., "一、", "（一）", "1.", etc.) and converts them to pure title text. |
| **Add Subheading Numbers** | If checked, automatically adds numbers based on heading levels (numbering scheme can be configured in settings). |

Note: DOCX -> MD now restores multilevel numbering linked from paragraph styles (pStyle) in numbering.xml, so heading prefixes created by Word/WPS multilevel lists such as "一、", "（一）", "1．", "（1）", and "①" are preserved in both simplified and advanced-field modes; heading levels are still detected correctly when "Clean Subheading Numbers" is enabled.

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
2.  **Markdown to Excel/CSV**: Basic Markdown tables can be exported directly. XLSX templates consume YAML fields, vertical/horizontal table placeholders, image placeholders, protected-cell notes, existing merged ranges, Markdown `<` / `^` merge markers, and the template-backed CSV sheet artifact chain.

**Supported Formats**:
-   `.xlsx` - Standard Excel document.
-   `.xls` - Automatically converted to XLSX for processing.
-   `.et` - WPS spreadsheet automatically converted.
-   `.csv` - CSV text table.
-   `.tsv` - TSV tab-separated table.

### Text Proofreading

The program provides four customizable proofreading rules:

1.  **Punctuation Pairing Check** - Detects if paired punctuation like parentheses and quotes match.
2.  **Symbol Proofreading** - Detects mixed use of Chinese and English punctuation.
3.  **Typo Check** - Checks for common typos based on a custom dictionary.
4.  **Sensitive Word Detection** - Detects sensitive words based on a custom dictionary.

**Custom Dictionaries**: Visually edit typo and sensitive word dictionaries in the "Settings" interface.

**Usage**:
1.  Drag the Word or Markdown document to be proofread into the program.
2.  Check the required proofreading rules.
3.  Click the "Text Proofreading" button.
4.  For Word documents, proofreading results are displayed as comments in the document. For Markdown files, results are output as a structured JSON report.

Note (Markdown proofreading report):
- Engine: `text_rules` + Markdown adapter `md_spell`
- Output: use `validate <file>`. The command is read-only unless `--report <path>` is supplied. Markdown reports are JSON; DOCX and pre-converted DOC/WPS/RTF/ODT reports are annotated DOCX files. `--json` controls the CLI envelope, not the report format.

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

#### Excel Template Placeholders (Focused Restore and Parity Boundary)

Markdown/TXT → `xlsx/xls/ods/csv` supports YAML field placeholders, vertical and horizontal table placeholders, image placeholders, protected-cell notes, existing template merged ranges, Markdown `<` / `^` merge markers, and the CSV per-sheet artifact chain.

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

**Merged Cell Handling**:

- Markdown -> Excel continues to preserve the template's original merged ranges.
- For known column-oriented template regions composed of contiguous `{{↓Field Name}}` placeholders, the program can restore rectangular merges from explicit Markdown table `<` / `^` markers.
- Only cells whose trimmed content is exactly `<` or `^` participate in merge detection; `\<` and `\^` remain literal text.
- Invalid rectangles or conflicts with existing template merged ranges are downgraded to plain text with a warning instead of forcibly overwriting template structure.

**Multi-table Data Merge**: If there are multiple tables in Markdown using the same header name, data will be merged in order and filled sequentially.

## 🔌 Obsidian Plugin

A companion Obsidian plugin is published separately and works in tandem with the converter:

### Core Features

-   **🚀 One-Click Launch** - Sidebar icon to quickly launch the converter.
-   **📂 Automatic Handover** - Automatically passes the currently open file path.
-   **🔄 Single Instance Management** - Automatically sends file if the program is already running, no need to restart.
-   **🔒 Bounded Local Control** - Uses typed `status`, `open`, and `activate` requests without process-name probing or command/status files.

### Working Principle

DocWen Core's runtime/control transport uses a Windows named pipe or an AF_UNIX socket on
Linux/macOS. A file lock only establishes single-instance ownership; files are not used to transport
control commands. This describes the Core transport only. DocWen Assistant 2.0 remains Windows
desktop-only and has no Linux/macOS combination acceptance.

1.  **First Click** → Launch converter and pass current file.
2.  **Click Again (With File)** → Replace with new file (Single File Mode).
3.  **Click Again (No File)** → Activate converter window.

### Installation

DocWen Assistant 2.0 uses DocWen Machine Protocol v1 and the single Artifact Bundle v2 contract. Its
source version does not prove publication; install only a numeric release that explicitly identifies a
compatible published DocWen release.

## 🔌 OpenClaw (Plugin + Skill)

OpenClaw 2.0 uses DocWen Machine Protocol v1 and the single Artifact Bundle v2 contract. Its source
version does not prove publication; follow the numeric release page and install only after its immutable
release gate succeeds.

## ❓ FAQ

### What if conversion fails?

-   Check if the file is occupied by another program.
-   Confirm the file format is correct.
-   Check the current log file path in the settings dialog, or inspect the system user log directory. Packaged verification flows may override the directory with `DOCWEN_LOG_DIR`.

### Template not showing?

-   Confirm template files are in the `templates/` directory.
-   Check if the template file is corrupted.
-   Restart the program to reload templates.

### Proofreading function not working?

-   Confirm the document is in .docx or .md format.
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

-   **Completely Local Operation**: Processing runs locally by default and does not depend on online services.
-   **Dependency Egress Protection**: Supported GUI/CLI entry points fail closed unless a process-lifetime CPython audit guard is active. It blocks all DNS/name resolution and AF_INET/AF_INET6 `bind`, `connect`, `connect_ex`, `sendto`, and `sendmsg` operations from in-process Python dependencies while preserving Windows named pipes and Unix-domain sockets.
-   **Explicit Boundary**: Separately launched processes, including Office/WPS/LibreOffice and the dedicated Office helper, are not managed. This protection is defense in depth against accidental dependency egress, not an operating-system sandbox.
-   **No Data Upload**: DocWen has no upload, telemetry, online-download, or network-service feature.
-   **Strict Security Mode**: Enabled by default; the app exits if core security checks fail. See [开发调试与故障排查.md](docs/maintenance/troubleshooting.md).

## 📜 License

This project is licensed under the **GNU Affero General Public License v3.0 (AGPL-3.0)**.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

-   This project uses PyMuPDF (licensed under AGPL-3.0), so the entire project is also licensed under AGPL-3.0.
-   The current GUI can use `PySide6-Fluent-Widgets` (QFluentWidgets) on supported host paths; this dependency follows a `GPLv3 / commercial` dual-license model, while this repository continues to distribute under AGPL.
-   You are free to use, modify, and distribute this software.
-   If you modify this software and provide services over a network, you must provide the modified source code to users.
-   For detailed license information, please see the [LICENSE](LICENSE) file.
-   For third-party component notices, see [LICENSE_THIRD_PARTY.txt](LICENSE_THIRD_PARTY.txt) and [NOTICE.txt](NOTICE.txt).

### Contact

-   **GitHub**: https://github.com/ZHYX91/docwen
-   **Contact Author**: zhengyx91@hotmail.com

---

**Author**: ZhengYX
