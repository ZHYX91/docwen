# DocWen

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md) · [العربية](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ar.md)

DocWen converts office files locally. Use it to turn Word files into Markdown, export Markdown to Word, extract text and tables from PDFs, or clean up files for a report. It has a desktop interface for everyday work and a command line for repeatable tasks.

## What DocWen is for

Documents arrive in different formats, and reformatting them by hand takes time. DocWen helps when you need to:

- Convert Word and Markdown in either direction, including headings, tables and formulas where the selected route supports them.
- Turn Excel spreadsheets into Markdown tables, or produce spreadsheets from Markdown.
- Extract text and images from PDFs, XPS and OFD files; use offline OCR where available. OCR can make mistakes, so check the original.
- Merge or split PDFs, convert common image formats, and use your own Word or spreadsheet templates.
- Check Markdown or Word text with your own spelling, symbol and punctuation rules.
- Process one file in the desktop app, or script a batch using the CLI.

DocWen handles DOCX, XLSX, PDF and several older office and image formats. Availability depends on the operation, platform and installed conversion software: not every format converts to every other format. Check [the capability reference](docs/capabilities.md) for details.

## Working offline

DocWen processes files on your computer rather than uploading them to a conversion service. Its Python process blocks dependency network calls, while external Office, WPS and LibreOffice processes follow their own system network policies. For sensitive work, keep those programs and your operating system configured accordingly.

## 📸 Screenshots

Select a template and generation options for Markdown documents.

![Markdown to document](docs/assets/screenshots/markdown-generation-en.png)

Choose attachment alignment in Official documents settings:

![Official document field formatting](docs/assets/screenshots/official-documents-en.png)


Changelog: see [CHANGELOG.md](docs/CHANGELOG.md)

## 🚀 Quick Start

### Installation from Source

**Prerequisites**: Python 3.12

**Package targets**: DocWen provides one Windows x64 GUI+CLI package and separate
Ubuntu 24.04 x64 GUI+CLI and CLI-only packages. Other Linux
distributions and macOS remain source/development paths and are not implied by the Ubuntu package.

**Option 1: Using uv (Recommended)**

Install the production-pinned `uv 0.12.0`, then:

```bash
git clone https://github.com/ZHYX91/docwen.git
cd docwen
uv sync --frozen --all-extras
```

DocWen CI and production builds use exactly `uv 0.12.0`; dependency/lock maintenance may use a compatible `uv 0.12.x` so automated security updates can refresh the checked-in lock without changing the production toolchain.
`pip install -e` is unsupported because pip cannot apply the repository's scoped dependency exclusion.

### Launch Program

On the Windows packaged release, double-click `DocWen.exe` to start the graphical interface. On
Ubuntu 24.04 x64, extract `DocWen-<version>-linux-x64.tar.gz` and run `./DocWen`; the companion
`DocWenCLI-<version>-linux-x64.tar.gz` is the CLI-only package. Download the matching assets from [DocWen Releases](https://github.com/ZHYX91/docwen/releases). If installed from source, run:

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

**Supported package target**: DocWen supports the packaged GUI and CLI on Ubuntu 24.04 x64.
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
    - Click "Generate DOCX".

3.  **Get Results**:
    - DocWen saves the Word document in a result folder at the configured output location. Use **Open Output** to find it.

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
5.  **Execute Operation**: Click the corresponding function button (e.g., "Export Markdown", "Generate DOCX", etc.).
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

Click the settings icon at the bottom of the window or press **Ctrl+,**.

The sidebar groups related settings under headings. In a short window it scrolls vertically, keeping every page reachable; a narrow window uses a page selector instead. Switching layouts preserves the selected page. When the sidebar has keyboard focus, **Up / Down** switch pages, **Home / End** select the first or last page, and **Tab** enters the selected page. Group headings are not selectable.

Under **General → Appearance**, choose **Small / Standard / Large** text and **90 / 100 / 110 / 125 / 150%** interface scale. Scaling is relative to the system setting: system 200% with application 90% gives a nominal 180%. Changes preview immediately; Apply or OK saves them, and Cancel restores the last saved appearance.

Settings are organized into **Basics**, **Input types**, **Content** and **Specialized**. The **OCR and images** page controls shared recognition and Markdown image output; per-input OCR defaults stay on the corresponding input pages. Image format conversion and compression remain on Incoming images.

The **Specialized** group has dedicated **Official documents** and **Invoices** pages. Official document settings separate **Document → Markdown** recognition defaults from **Markdown → Document** template field formatting. The field toggle has an information button describing attachment notes, copied recipients, dates, annotations and supported placeholder positions.

Invoices have separate defaults for **PDF / OFD → Markdown** and **Image → Markdown**. A default only sets the initial choice for compatible files; it does not classify a file as an official document or invoice. Each input category has one default: **Set as default** replaces its previous choice, while **Clear default** removes the current page's choice. Apply or OK saves these changes; Cancel discards the draft. Shared extraction and OCR settings remain on their own pages, with links from the optimization pages. Invoice images always use OCR; scanned PDFs require OCR enabled in the incoming fixed-layout settings, and OFD reads existing text and structure without OCR for scanned content.

**Interface language** controls application text and layout. **Task output language** independently controls generated labels, document styles and proofreading comments; it does not translate body text. Changing the interface keeps saved template choices and output language. OCR **Auto** follows the task output language; choose an OCR language explicitly when needed. See [language settings](docs/specs/task-language.md).

**Markdown syntax** separates shared output conventions from input recognition, with Word-specific mappings listed below. The visible help describes current support: output extensions and preferred syntax are not yet implemented uniformly across all source formats or specialized optimizers. Enabling an extension does not add recognition capabilities.

Under **Markdown syntax**, **Each line as a paragraph** optionally turns ordinary prose lines into separate Word paragraphs. It is off by default; lists, quotes, code, semantic structures and explicit Markdown hard breaks retain their own boundaries.

**Scope:** The specialized field formatter targets Chinese official-document conventions, and the invoice extraction optimization targets Chinese invoice formats. Other documents can still use the regular conversion paths.

### Shortcuts

-   **Drag External File**: Drag directly into the window to import.
-   **Open Output**: Click the "Open Output" action on the right side of the status bar to open the output location.
-   **Right-click Template Item**: Choose **Open file location** to show the template in its folder.

---

## 🔧 Command Line Usage

Use the command line (CLI) to inspect files, convert them from a terminal, or automate repeated tasks. For integrations with other apps, use the Machine Protocol described below.

### Recommended CLI Automation Flow

For terminal use and local automation scripts, use this order:

1. `inspect <file> [--json]`: detect the real file category, format, and supported actions first.
2. `resources list formats --json`: check which conversions are available and whether required software is installed.
3. `schema convert`: read the accepted conversion options, defaults and requirements.
4. `convert <file> --to <fmt> (--output-dir <dir> | --output <path>) --dry-run --json`: check the detected format, conversion method and options without writing files.
5. `convert <file> --to <fmt> (--output-dir <dir> | --output <path>) ...`: run the actual conversion after the preview is acceptable.

Apps and plugins connect through `serve --stdio` using Machine Protocol v2 to discover operations, plan a task and run it. This is the supported API for integrations; CLI `--json` is intended for terminal scripts. Use `gui open|activate|status` to open files in the desktop app or bring its window to the front. See the [CLI guide](docs/cli.md) and [integration protocol](docs/specs/machine-protocol-v2.md) for details.

### Common Examples

Conversions from or to Markdown publish a result folder inside the parent selected with `--output-dir DIR`, even when there is only one output. Names include the source name, timestamp and current input format. Use `--output PATH` only for other conversions that produce one exact file.

```bash
# Packaged release (Windows)
DocWenCLI.exe inspect document.docx --json

# List the accepted conversion options
DocWenCLI.exe schema convert

# Preview how the conversion would run without writing files
DocWenCLI.exe convert report.docx --to md --output-dir exports --extract-img --ocr --dry-run --json

# Export Word to Markdown (extract images + OCR)
DocWenCLI.exe convert report.docx --to md --output-dir exports --extract-img --ocr

# Markdown to Word (select a template and heading/body merge mode)
DocWenCLI.exe convert document.md --to docx --output-dir exports --template template.docx.f1eeb0a008ce3eae0619ecae6e185ab132a3ee0abdd382c8863481d9af1dc77f --heading-merge-mode punct_required

# Markdown to Word, then proofread the generated DOCX before saving it
# --check is valid on convert only together with --proofread.
DocWenCLI.exe convert document.md --to docx --output-dir exports --proofread --check typo --check punct

# Control Markdown image export mode and OCR placement
DocWenCLI.exe convert report.docx --to md --output-dir exports --extract-img --image-mode file --ocr --ocr-placement image_md

# Chinese official-document optimization selected by public resource ID
DocWenCLI.exe convert report.docx --to md --output-dir exports --optimization gongwen

# Chinese invoice optimization selected by public resource ID
DocWenCLI.exe convert invoice.pdf --to md --output-dir exports --optimization invoice_cn

# Check available conversions and required software
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
# docwen convert document.docx --to md --output-dir exports --dry-run --json
# docwen convert document.docx --to md --output-dir exports
```

### Common Commands & Options

The table below lists common commands only. For the full command surface, use `docwen --help` (source / uv) or `DocWenCLI --help` (packaged release).

| Command / Option | Description |
| --- | --- |
| `convert <file> --to <fmt> (--output-dir <dir> \| --output <path>)` | Convert one file. Conversions from or to Markdown require `--output-dir`; other one-file conversions use exact `--output`. Existing targets require explicit `--overwrite`. |
| `convert <file> --to <fmt> ... --dry-run --json` | Check the detected format, conversion method and options without writing files. |
| `convert <markdown> --to docx --output-dir <dir> --proofread [--check ...]` | Generate DOCX, proofread it, then save the result. `--check` without `--proofread` is rejected; omitting `--check` uses configured proofreading defaults. |
| `convert <file> --to <fmt> ... --optimization <id>` | Select an optimization by the public resource ID returned by `resources list optimizations`. |
| `validate <file> --check ... [--report <path>]` | Proofread DOCX, Markdown, or legacy Word-family content. The default is read-only; `--report` writes annotated DOCX for document input or JSON for Markdown. |
| `number markdown <file> --operation add\|remove (--output <path> \| --in-place)` | Explicitly add or remove Markdown heading numbering. |
| `merge pdf\|tables\|images <files...> --output <path>` | Combine PDFs, spreadsheets or images into one output. |
| `split pdf <file> --pages <range> --output-dir <dir>` | Split selected PDF pages into an explicit directory. |
| `batch convert\|validate <files...>` | Run an explicit multi-file conversion or proofreading operation. |
| `schema convert` | List accepted conversion options, defaults and requirements in a form scripts can read. |
| `inspect <file> [--json]` | Inspect file category/format, recommended actions, and extension/content mismatch warnings. |
| `doctor --json` | Check the installation, available conversions and required software. |
| `resources list formats --json` | List supported conversions, required software and current limitations. |
| `resources list optimizations --json` | List optional optimizations and the conversions they apply to. |
| `resources list templates [--target docx\|xlsx]` | List available templates. |
| `resources list numbering-schemes` | List available numbering schemes. |
| `gui open [ABSOLUTE_FILE] [--timeout <seconds>] [--json]` | Open the DocWen window and optionally load a file. Use an absolute file path. |
| `--template <id>` | Use the resource ID returned by `resources list templates`, rather than a template name or path. DOCX IDs apply to `docx/doc/odt/rtf/wps/pdf`, XLSX IDs to `xlsx/xls/ods/csv`. CSV export writes table data directly when no template is selected. With an XLSX template, DocWen fills the workbook first and creates one CSV per worksheet. |
| `--extract-img` / `--no-extract-img` / `--ocr` | Image extraction and OCR options for `convert --to md`. |
| `--image-mode file\|base64\|embed\|omit` | Control how images are emitted during Markdown export. |
| `--ocr-placement image_md\|main_md` | Control whether OCR text is written to image-side Markdown or the main Markdown file. |
| `--heading-merge-mode punct_required\|always\|never` | Control the heading + body merge strategy for `convert --to docx`. |
| `--optimization <id>` | Apply an optimization from `resources list optimizations`. |
| `--jobs` / `--continue-on-error` | Controls provided by explicit `batch` subcommands. |
| `--json` / `--quiet` / `--timing` | Show JSON output, reduce logs, or include elapsed time in terminal scripts. |


## 📝 Markdown Syntax Conventions

**Using Obsidian Markdown extensions?** Four extensions are optional and **disabled on a new DocWen installation**: `structural_tables` (merged cells and multi-row/row headers), `captions_references` (Number Suite captions and references), `extended_headings` (H7–H9), and `typed_endnotes` (typed endnotes).

- **Markdown → Word:** Open **Settings → Markdown syntax → Input** and enable the extensions your notes need, or click **Use Obsidian extensions** for the input direction. Otherwise special markers may remain literal text.
- **Word → Markdown:** Enable the corresponding **Output** extensions separately (or use the output preset). With them off, structures may be flattened or normalized and a warning reported.

Keep the original Markdown and inspect the produced DOCX. Word numbering follows DocWen's conversion settings; the round trip does not promise identical source spelling or whitespace. See the [Markdown compatibility reference](https://github.com/ZHYX91/docwen/blob/main/docs/specs/markdown-compatibility.md) for exact boundaries.

### Heading Level Mapping

Markdown headings map directly to Word heading levels:
- Document title and subtitle are placed in YAML metadata.
- Markdown `# Heading 1` corresponds to Word "Heading 1".
- Markdown `## Heading 2` corresponds to Word "Heading 2".
- Levels 1–6 use standard Markdown syntax. For levels 7–9, enable **Extended headings** in the Markdown input settings and use seven to nine leading `#` markers.

**Tip**: If you prefer using Markdown's first-level heading (`#`) as the document title, starting from second-level headings (`##`) for body headings, you can style "Heading 1" in the Word template to look like a document title (e.g., centered, bold, larger font size), and select a numbering scheme that skips first-level heading numbering in the settings. This way, your first-level headings will appear as document titles.

### Line Breaks and Paragraphs

**Basic Rule**: Ordinary Markdown soft line breaks do not create separate Word paragraphs by default. Enable **Each line as a paragraph** under **Settings → Markdown syntax** to split ordinary prose lines. Lists, quotes, code, other semantic structures and explicit hard breaks retain their own rules.

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
- The trigger-character field is editable in the Incoming text settings. An empty value disables merging in "Punctuation required" mode. Commas, semicolons, enumeration commas, dashes, and ellipses are deliberately excluded from the default, but users may add them explicitly.
- You can change this in Settings → Incoming text → "Markdown to Document — Format Processing" → "Heading + body merge mode".

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
- Bare URLs become clickable in Word only when `[non_embed_links].auto_link_bare_url` is enabled in `configs/link.toml`; this option is off by default
- Markdown -> XLSX keeps the original link syntax instead of emitting DOCX hyperlink placeholders

## 📖 Detailed Usage Guide

### Word to Markdown

1.  Drag the `.docx` file into the program window.
2. Choose the export options, then click **Export Markdown**.
3.  Generates a `.md` file containing YAML metadata.

**Supported Formats**:
-   `.docx` - Standard Word document.
-   `.doc` - Automatically converted to DOCX for processing.
-   `.wps` - WPS document automatically converted.

**Export Options**:

| Option | Description |
| :--- | :--- |
| **Extract Images** | If checked, images in the document are extracted to the output folder, and image links are inserted into the MD file. |
| **Image OCR** | Recognize text in images. Save it in a Markdown file for each image or in the main Markdown file, according to the OCR output setting. |
| **Optimize for document type** | Apply an optimization for the selected document type, such as extracting fields from Chinese official documents. Leave this off to use the standard conversion. |
| **Remove existing heading numbers** | If checked, removes numbers before subheadings (e.g., "一、", "（一）", "1.", etc.) and converts them to pure title text. |
| **Add heading numbers** | If checked, automatically adds numbers based on heading levels (numbering scheme can be configured in settings). |

Word/WPS heading numbers created with multilevel lists are preserved during export. Enable **Remove existing heading numbers** to remove those numbers while keeping the heading levels.

### Markdown to Word

1. Prepare a `.md` file. Add a YAML header if you want to supply a title or custom fields.
2.  Drag it into the program window and select the corresponding Word template.
3. Click **Generate DOCX** to fill the template and save the document.

**Conversion Options**:

| Option | Description |
| :--- | :--- |
| **Remove existing heading numbers** | If checked, removes numbers before subheadings. |
| **Add heading numbers** | If checked, automatically adds numbers based on heading levels. |

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
    -   New managed style names use **Task output language** in General settings, not interface language. Switching interface language does not rename styles already in the selected template.

**Suggestion**: After customizing styles in the template, the converter will automatically use your styles; if not present in the template, it will use built-in preset styles.

### Spreadsheet File Processing

1. **Excel/CSV to Markdown**: Add an `.xlsx` or `.csv` file and click **Export Markdown**.
2. **Markdown to Excel/CSV**: Export Markdown tables directly, or choose an XLSX template to fill YAML fields, tables and images. Template formatting, existing merged cells and cell protection are handled during conversion. With a template, CSV export creates one file per worksheet.

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
3.  Click the "Proofread" button.
4.  For Word documents, proofreading results are displayed as comments in the document. For Markdown files, results are output as a structured JSON report.

From the command line, use `validate <file>` to check a file without changing it. Add `--report <path>` to save the findings: Markdown produces a JSON report; Word and supported legacy document formats produce a DOCX with comments. `--json` changes how the command prints its result.

## 🛠️ Template System

### Using Existing Templates

Manage DOCX/XLSX templates in **Settings → Templates**. Built-in templates are read-only; custom templates use writable user storage. Enable, disable, sort and choose defaults here. MSIX users do not need to modify WindowsApps.

### Custom Templates

1. Select a built-in template and choose **Copy and edit**, or import an existing DOCX/XLSX file.
2. Edit the custom copy in Word, Excel or WPS; adjust styles and placeholders as described below.
3. Save the file, return to DocWen and choose **Refresh**. No restart is needed.

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

Use an XLSX template to place YAML values, table rows or columns, and images in a spreadsheet. DocWen preserves existing merged cells and can recreate merges marked with `<` / `^` in Markdown. Check conversion warnings for protected cells. When exporting to CSV with a template, each worksheet becomes a separate CSV file.

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
-   **🔄 Reuse the open window** - Open another file in the running converter without restarting it.

### Using the plugin

DocWen Assistant 3.3 supports Windows and Linux desktop hosts. Windows Microsoft Store installations can use automatic detection; Linux uses manual package selection. The Assistant does not currently support macOS.

1.  **First click** → Launch the converter and open the current file.
2.  **Click again with a file open** → Replace the current input with that file (single-file mode).
3.  **Click again without a file open** → Bring the converter window to the front.

### Installation

Install [DocWen Assistant](https://github.com/ZHYX91/obsidian-docwen-assistant) and a compatible DocWen release. Assistant 3.3 requires DocWen 0.17.0 or later. The plugin does not install DocWen for you. Check the release notes for compatible versions before downloading. For integration development, these versions use Machine Protocol v2 and Artifact Bundle v3 for conversion results.

## 🔌 OpenClaw (Plugin + Skill)

[docwen-openclaw](https://github.com/ZHYX91/docwen-openclaw) lets OpenClaw use DocWen to convert and check local files. Install DocWen separately and choose a plugin release that lists a compatible DocWen version. Plugin 3.0 uses Machine Protocol v2 and Artifact Bundle v3. Use the published release assets; a version number in source code does not mean that a release is available.

## ❓ FAQ

### What if conversion fails?

-   Check if the file is occupied by another program.
-   Confirm the file format is correct.
-   Open **Settings → Logging** to find the current log file and read the error details.

### Template not showing?

- Check that the template is enabled in **Settings → Templates** and matches the selected output format.
- Confirm that the DOCX/XLSX file can be opened normally.
- Choose **Refresh** after external changes; use **Import templates** to add a file.

### Proofreading function not working?

-   Confirm the document is in .docx or .md format.
-   Check if the document contains editable text.
-   Confirm proofreading rules are enabled in settings.

### Output format not as expected?

- Open **Settings → Templates**. For a built-in template, use **Copy and edit**; for a custom template, use **Edit**.
- Adjust its styles in Word, Excel or WPS, then save and refresh the list.
- Select the custom template before generating new documents.

### Formula cells are empty after Excel to Markdown conversion?

DocWen exports the **last saved calculation result** of a formula cell. If the spreadsheet has no saved result, the exported cell will be empty. This can happen with files created by another program or edited without recalculating and saving in Excel.

**Solution**:
1.  Open the file in Excel.
2.  Wait for formula calculation to complete.
3.  Save the file.
4.  Convert again.

## 🔒 Security Features

DocWen processes files on your computer and has no file-upload, telemetry, online-download or network-service feature. It restricts network access by its own Python components. If this protection cannot start, DocWen does not continue running.

Separately launched programs such as Word/Office, WPS and LibreOffice, including helper processes, follow their own and the operating system's network settings. They are outside this restriction, so check their settings as well when working with sensitive files.

See [how network protection works](docs/architecture.md) for its implementation and limits, or [troubleshooting](docs/maintenance/troubleshooting.md) if the program cannot start or a conversion fails.

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
