# DocWen

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md)

Công cụ chuyển đổi định dạng tài liệu và bảng biểu: hỗ trợ chuyển đổi hai chiều Word/Markdown/Excel. Chạy trên máy (offline), đảm bảo an toàn dữ liệu.

## 📖 Bối cảnh dự án

Phần mềm được tạo ra để giải quyết các vấn đề thường gặp trong công việc văn phòng:
- Tài liệu từ nhiều nguồn có định dạng không thống nhất, cần chuẩn hoá.
- Nhiều loại file với yêu cầu định dạng khác nhau.
- Cần chạy offline trong môi trường intranet/thiết bị cũ.

**Triết lý thiết kế**: công cụ nhẹ, dễ dùng, chi phí học thấp. Không nhằm thay thế các công cụ chuyên nghiệp như LaTeX/Pandoc.

## ✨ Tính năng chính

- **📄 Chuyển đổi tài liệu** - Word ↔ Markdown, hỗ trợ công thức, ánh xạ dấu phân cách (---/***/___) với ngắt trang/ngắt mục/dòng kẻ và khôi phục marker bảng Markdown `<` / `^` thành gộp ô hình chữ nhật trong Word. DOCX/DOC/WPS/RTF/ODT.
- **📊 Chuyển đổi bảng tính** - Excel ↔ Markdown. XLSX/XLS/ET/ODS/CSV/TSV. Có chiến lược xuất ô gộp cấu hình được (`fill / empty / marker`), công cụ tóm tắt bảng và các placeholder template được mô tả bên dưới.
- **📑 PDF & file bố cục** - PDF/XPS/OFD → Markdown hoặc DOCX. Hỗ trợ gộp/tách PDF.
- **🖼️ Ảnh** - Chuyển đổi và nén JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC.
- **📥 Nhập định dạng khác** - Hỗ trợ chuyển đổi một chiều HTML/MHTML/ENEX/EPUB/PPTX/PPT sang Markdown.
- **🔍 Nhận dạng văn bản OCR** - Tích hợp RapidOCR để trích xuất văn bản từ ảnh và PDF.
- **✏️ Kiểm tra lỗi** - Quy tắc tuỳ chỉnh cho dấu câu, ký hiệu, lỗi chính tả và từ nhạy cảm trong file Word (.docx) và Markdown (.md). Có thể chỉnh sửa quy tắc trong giao diện cài đặt.
- **📝 Mẫu (Template)** - Cơ chế template linh hoạt cho tài liệu/báo cáo.
- **💻 GUI + CLI** - Giao diện đồ hoạ và dòng lệnh.
- **🔒 Xử lý cục bộ với bảo vệ kết nối đi của thư viện** - Việc chuyển đổi không phụ thuộc dịch vụ trực tuyến. Khi DocWen chạy, tiến trình Python chặn DNS và IPv4/IPv6 của thư viện trong tiến trình; các ứng dụng Office bên ngoài vẫn theo chính sách mạng của hệ thống.
- **🔗 Chạy đơn phiên bản** - Tự động quản lý instance của chương trình và hỗ trợ tích hợp với plugin Obsidian đi kèm.

## 📸 Ảnh chụp màn hình

| Hàng loạt | Markdown |
| --- | --- |
| ![Bảng hàng loạt](../assets/screenshots/batch-light.png) | ![Cửa sổ chính](../assets/screenshots/main-light.png) |

| Tài liệu | Bảng tính |
| --- | --- |
| ![Bảng tài liệu](../assets/screenshots/conversion-document-light.png) | ![Bảng bảng tính](../assets/screenshots/conversion-spreadsheet-light.png) |

| Ảnh | Tệp bố cục |
| --- | --- |
| ![Bảng ảnh](../assets/screenshots/conversion-image-light.png) | ![Bảng bố cục](../assets/screenshots/conversion-layout-light.png) |

Nhật ký thay đổi: xem [CHANGELOG.md](../CHANGELOG.md)

## 🚀 Bắt đầu nhanh

### Cài đặt từ mã nguồn

**Yêu cầu**: Python 3.12

**Phạm vi mục tiêu 0.9**: Mã nguồn này tạo gói cho Windows x64 và Ubuntu 24.04 x64. Các bản phân
phối Linux khác và macOS vẫn là đường dẫn mã nguồn/phát triển, không thuộc cam kết của gói Ubuntu.

**Cách 1: Sử dụng uv (khuyến nghị)**

Cài đặt [uv](https://docs.astral.sh/uv/getting-started/), sau đó:

```bash
git clone https://github.com/ZHYX91/docwen.git
cd docwen
uv sync --frozen --all-extras
```

Mã nguồn, kiểm thử và bản dựng DocWen 0.9 chỉ hỗ trợ tệp khóa trong kho với `uv 0.12.0`; `pip install -e` không được hỗ trợ.

### Khởi chạy chương trình

Với bản đóng gói trên Windows: nhấp đúp `DocWen.exe` để mở GUI. Sau khi cài từ mã nguồn:

```bash
docwen-gui  # Chế độ GUI
docwen      # Chế độ CLI
```

### Ghi chú cho macOS

**Giới hạn hiện tại**: Trên macOS, các capability `convert`, `validate`, `number`, `merge`, `split`
hiện không khả dụng. Phần dưới chỉ ghi các phụ thuộc tùy chọn cho thử nghiệm phát triển.

**Hỗ trợ LibreOffice (Tùy chọn)**

Nếu cần chuyển đổi các định dạng cũ như `.doc`, `.xls`, hãy cài LibreOffice:  
Tải về: https://www.libreoffice.org/download/

**Hỗ trợ ảnh HEIC (Tùy chọn)**

Để xử lý ảnh HEIC/HEIF:

```bash
brew install libheif
pip install pillow-heif
```

### Yêu cầu cho bản GUI trên Linux

**Đích gói được hỗ trợ**: DocWen 0.9 hỗ trợ GUI và CLI trong gói Ubuntu 24.04 x64. Các yêu cầu này
không mở rộng cam kết sang bản phân phối hoặc kiến trúc khác.

- Có môi trường desktop (GNOME, KDE, XFCE, ...)
- GUI dùng PySide6 (Qt6) và không còn phụ thuộc vào Python Tk. Nếu khởi động lỗi vì thiếu thư viện hệ thống, hãy cài các phụ thuộc runtime của Qt theo thông báo lỗi (thường liên quan OpenGL/X11).
- Với máy chủ headless, hãy ưu tiên entry CLI `docwen` thay vì GUI; bản đóng gói Windows cũng cung cấp `DocWenCLI.exe`.

### Hướng dẫn nhanh

1.  **Chuẩn bị file Markdown**:

    ```markdown
    ---
    title: Test Document
    ---
    
    ## Test Title
    
    This is the test body content.
    ```

2.  **Chuyển đổi bằng kéo thả**:
    - Khởi chạy chương trình.
    - Kéo file `.md` vào cửa sổ.
    - Chọn template.
    - Nhấn "Convert to DOCX".

3.  **Nhận kết quả**:
    - Tài liệu Word chuẩn hoá sẽ được tạo trong cùng thư mục.

**Mẹo**: Có thể dùng các file mẫu trong thư mục `samples/` để trải nghiệm nhanh.

## 🖥️ Sử dụng giao diện đồ hoạ

Hầu hết người dùng sử dụng phần mềm qua giao diện đồ hoạ. Dưới đây là hướng dẫn chi tiết.

### Tổng quan giao diện

Chương trình dùng **bố cục 3 cột thích ứng**:

| Khu vực | Mô tả | Khi hiển thị |
| :--- | :--- | :--- |
| **Cột giữa (khu vực chính)** | Khu kéo thả file, panel thao tác, thanh trạng thái | Luôn hiển thị |
| **Cột phải** | Bộ chọn template / panel chuyển đổi | Tự mở sau khi chọn file |
| **Cột trái** | Danh sách file theo lô (nhóm theo loại) | Hiển thị khi bật chế độ theo lô |

### Quy trình thao tác cơ bản

1.  **Khởi chạy**: Nhấp đúp `DocWen.exe` (Windows đóng gói) hoặc chạy `docwen-gui`.
2.  **Nhập file**:
    -   Cách 1: Kéo thả file vào cửa sổ.
    -   Cách 2: Nhấn nút "Add" trong vùng kéo thả để chọn file.
3.  **Chọn template** (nếu cần): Panel template bên phải tự mở, chọn template phù hợp.
4.  **Chọn tuỳ chọn**: Tick các tuỳ chọn xuất/chuyển đổi trong panel thao tác.
5.  **Thực thi**: Nhấn nút chức năng tương ứng (ví dụ: "Export MD", "Convert to DOCX", ...).
6.  **Xem kết quả**: Thanh trạng thái hiển thị tiến độ và kết quả; nhấn thao tác "Mở đầu ra" ở bên phải để mở vị trí đầu ra.

### Chế độ 1 file vs chế độ theo lô

Chương trình hỗ trợ hai chế độ xử lý, chuyển đổi bằng nút trong vùng kéo thả:

**Chế độ 1 file** (mặc định):
-   Xử lý từng file một.
-   Giao diện gọn, phù hợp sử dụng hằng ngày.

**Chế độ theo lô**:
-   Nhập nhiều file cùng lúc.
-   Cột trái hiển thị danh sách file theo nhóm.
-   Hỗ trợ thêm/xoá/sắp xếp theo lô.
-   Nhấp vào file trong danh sách để đổi mục tiêu thao tác.

### Chức năng panel thao tác

Panel thao tác tự điều chỉnh chức năng theo loại file:

| Loại file | Thao tác khả dụng |
| :--- | :--- |
| Tài liệu Word | Export MD, Convert PDF, Text Proofreading, OCR |
| Markdown | Convert DOCX, Convert PDF, Text Proofreading |
| Bảng tính Excel | Export MD, Convert PDF, Table Summary |
| PDF | Export MD, Merge, Split, OCR |
| Ảnh | Chuyển đổi định dạng, Nén, OCR |
| HTML/EPUB/PPTX v.v. | Export MD |

### Màn hình cài đặt

Nhấn nút "Cài đặt" trong phần tiêu đề của khu thao tác để mở cài đặt:

Cài đặt được tổ chức theo tab: **Chung**, **Văn bản**, **Soát lỗi**, **Tài liệu**, **Bảng tính**, **Hình ảnh**, **Bố cục**, **Liên kết**, **Định dạng**, **Đầu ra**, **Xuất**, **Ghi log**, **Khác**.

### Phím tắt

-   **Kéo file ngoài**: Kéo trực tiếp vào cửa sổ để nhập
-   **Mở đầu ra**: Nhấn thao tác "Mở đầu ra" ở bên phải thanh trạng thái để mở vị trí đầu ra.
-   **Chuột phải vào template**: Mở vị trí file template

---

## 🔧 Sử dụng CLI

Ngoài giao diện đồ họa, DocWen còn cung cấp giao diện dòng lệnh (CLI) cho tự động hóa, xử lý hàng loạt và tích hợp bên ngoài.

### Luồng gọi khuyến nghị cho tự động hóa

Đối với script, agent hoặc plugin, nên dùng thứ tự sau:

1. `inspect <file> [--json]`: trước tiên nhận diện loại tệp thực tế, định dạng và các thao tác được hỗ trợ.
2. `schema convert`: đọc hợp đồng máy đọc được và các ràng buộc điều kiện của `convert`.
3. `convert <file> --to <fmt> --output <path> --dry-run --json`: xem trước quá trình nhận diện, chuẩn hóa và định tuyến mà không ghi tệp đầu ra.
4. `convert <file> --to <fmt> --output <path> ...`: sau khi xác nhận, mới chạy chuyển đổi thật.

### Ví dụ thường dùng

```bash
# Bản đóng gói Windows
DocWenCLI.exe inspect document.docx --json

# Xuất hợp đồng convert cho script / agent
DocWenCLI.exe schema convert

# Xem trước cách chuyển đổi sẽ chạy mà không ghi kết quả
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --ocr --dry-run --json

# Xuất Word sang Markdown (trích ảnh + OCR)
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --ocr

# Markdown sang Word (mẫu + chế độ gộp tiêu đề/nội dung)
DocWenCLI.exe convert document.md --to docx --output document.docx --template template.docx.da28ee624892975bc590fd419880875136f22e0edcd878bca69472e81297c0bc --heading-merge-mode punct_required

# Điều khiển chế độ ảnh và vị trí văn bản OCR trong Markdown
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --image-mode file --ocr --ocr-placement image_md

# Kiểm tra khả năng chạy và cổng phụ thuộc
DocWenCLI.exe doctor --json
DocWenCLI.exe resources list formats --json

# Soát lỗi tài liệu
DocWenCLI.exe validate document.docx --check typo --check punct
DocWenCLI.exe validate input.md --check typo --check punct

# Từ mã nguồn / uv
# inspect -> schema -> dry-run -> convert
# docwen inspect document.docx --json
# docwen schema convert
# docwen convert document.docx --to md --output document.md --dry-run --json
# docwen convert document.docx --to md --output document.md
```

### Lệnh và tùy chọn thông dụng

Bảng dưới đây chỉ liệt kê các lệnh thông dụng. Để xem đầy đủ toàn bộ lệnh, dùng `docwen --help` (mã nguồn / uv) hoặc `DocWenCLI --help` (bản đóng gói).

| Lệnh / tùy chọn | Mô tả |
| --- | --- |
| `convert <file> --to <fmt> --output <path>` | Điểm vào thống nhất cho chuyển đổi. |
| `convert <file> --to <fmt> --output <path> --dry-run --json` | Xem trước nhận diện, chuẩn hóa, định tuyến và các tùy chọn hiệu lực mà không thực hiện chuyển đổi thật. |
| `schema convert` | Xuất hợp đồng máy đọc được, giá trị mặc định, điều kiện và khóa chuẩn của `convert`. |
| `validate <file> --check ...` | Soát lỗi tài liệu (`typo/punct/symbol/sensitive/all/none`). Dùng `--json` cho envelope của CLI; `--report` là đường dẫn tệp báo cáo tùy chọn. |
| `inspect <file> [--json]` | Kiểm tra loại/định dạng tệp, hành động gợi ý và cảnh báo khi phần mở rộng không khớp nội dung. |
| `doctor --json` | Trả về chẩn đoán cùng với phần tóm tắt khả năng chạy và cổng phụ thuộc. |
| `resources list formats --json` | Liệt kê định dạng đích theo loại nguồn và kèm cổng phụ thuộc / bản tóm tắt giới hạn. |
| `resources list templates` | Liệt kê các mẫu có sẵn. |
| `resources list numbering-schemes` | Liệt kê các sơ đồ đánh số có sẵn. |
| `--template <id>` | ID tài nguyên chuẩn chính xác từ `resources list templates`; tên hiển thị, tên tệp và đường dẫn đều bị từ chối. ID DOCX dùng cho `docx/doc/odt/rtf/wps/pdf`, ID XLSX cho `xlsx/xls/ods/csv`. |
| `--extract-img` / `--no-extract-img` / `--ocr` | Tùy chọn trích ảnh và OCR cho `convert --to md`. |
| `--image-mode file|base64` | Kiểm soát cách ảnh được xuất ra khi xuất Markdown. |
| `--ocr-placement image_md|main_md` | Kiểm soát việc ghi văn bản OCR vào Markdown phụ của ảnh hay Markdown chính. |
| `--heading-merge-mode punct_required|always|never` | Kiểm soát chiến lược gộp "tiêu đề + nội dung" cho `convert --to docx`. |
| `--optimization <id>` | Bật rõ ràng một hồ sơ tối ưu hóa (xem `resources list optimizations`). |
| `batch convert|validate ... --jobs <n> [--continue-on-error]` | Điều khiển xử lý hàng loạt. |
| `--json` / `--quiet` / `--timing` | Đầu ra có cấu trúc, giảm log và dữ liệu thời gian cho script hoặc plugin. |

Trong chế độ `punct_required`, danh sách mặc định chính xác là `。：！？.:!?`. Có thể chỉnh sửa danh sách này trong phần cài đặt định dạng; để trống sẽ tắt việc gộp trong chế độ này. Dấu phẩy, dấu chấm phẩy, dấu liệt kê, dấu gạch ngang và dấu chấm lửng không được bật mặc định.


## 📝 Quy ước Markdown

### Ánh xạ cấp tiêu đề

Để dễ ghi nhớ, tiêu đề Markdown trong phần mềm tương ứng **1:1** với tiêu đề Word:
- Tiêu đề và phụ đề của tài liệu đặt trong YAML metadata.
- Markdown `# Heading 1` tương ứng Word "Heading 1".
- Markdown `## Heading 2` tương ứng Word "Heading 2".
- Và tiếp tục như vậy, hỗ trợ tối đa 9 cấp.

**Mẹo**: Nếu bạn muốn dùng `#` làm tiêu đề tài liệu và dùng `##` trở đi cho tiêu đề nội dung, hãy chỉnh style "Heading 1" trong template Word để trông giống tiêu đề (ví dụ: căn giữa, in đậm, cỡ chữ lớn) và chọn scheme đánh số bỏ qua cấp 1 trong phần cài đặt.

### Xuống dòng và đoạn văn

**Quy tắc cơ bản**: Mỗi dòng không rỗng được xem là một đoạn riêng theo mặc định.

**Đoạn trộn**: Khi một tiêu đề phụ cần trộn với nội dung trong cùng một đoạn (mặc định: chế độ "Cần dấu kết câu"), phải thỏa các điều kiện:
1.  Tiêu đề phụ kết thúc bằng dấu kết câu (hỗ trợ dấu kết câu đa ngôn ngữ).
2.  Nội dung nằm ở **dòng ngay bên dưới** tiêu đề phụ.
3.  Dòng nội dung không được là phần tử Markdown đặc biệt (tiêu đề, code block, bảng, danh sách, trích dẫn, khối công thức, dấu phân cách, ...).

**Ví dụ**:
```markdown
## I. Work Requirements.
This meeting requires all units to earnestly implement...
```
Hai dòng trên sẽ được gộp thành một đoạn: "I. Work Requirements." giữ định dạng tiêu đề phụ và "This meeting..." giữ định dạng nội dung.

**Lưu ý**:
- Không được có dòng trống giữa tiêu đề phụ và nội dung; nếu có sẽ bị nhận diện thành hai đoạn riêng.
- Theo mặc định (chế độ "Cần dấu kết câu"), nếu tiêu đề phụ không kết thúc bằng dấu kết câu, sẽ không gộp với dòng kế tiếp dù không có dòng trống.
- Bạn có thể đổi trong Cài đặt → Định dạng → "MarkDown sang Word" → "Heading + body merge mode".

### Chuyển đổi dấu phân cách hai chiều

Hỗ trợ chuyển đổi hai chiều giữa dấu phân cách Markdown và ngắt trang/ngắt mục/dòng kẻ trong Word:

-   **DOCX → MD**: Tự động chuyển ngắt trang, ngắt mục và dòng kẻ của Word thành dấu phân cách Markdown.
-   **MD → DOCX**: Tự động chuyển `---`, `***`, `___` thành phần tử Word tương ứng.
-   **Có thể cấu hình**: Quan hệ ánh xạ có thể tuỳ chỉnh trong phần cài đặt.

### Danh sách công việc

Hỗ trợ chuyển đổi hai chiều danh sách công việc GFM:

```markdown
- [ ] Cần làm
- [x] Hoàn thành
```

-   **MD → DOCX**: Hiển thị dưới dạng danh sách có dấu đầu dòng với tiền tố `☐` / `☑`.
-   **DOCX → MD**: Chuyển đổi các mục danh sách bắt đầu bằng `☐` / `☑` / `☒` thành `- [ ]` / `- [x]`.
-   **Lưu ý về phông chữ**: `☐`/`☑` có thể không hiển thị trên một số phông chữ. Nếu cần, hãy sử dụng phông chữ như "Segoe UI Symbol" trong mẫu Word.

### Nhúng ảnh và kích thước

Hỗ trợ nhúng ảnh kiểu Obsidian/Wiki và Markdown chuẩn, có thể chỉ định kích thước (px):

```markdown
![[image.png]]
![[image.png|300]]
![[image.png\|300]]
![alt](image.png =300x200)
![alt](image.png =300x)
![alt|300](image.png)
```

- Không chỉ định: dùng kích thước gốc, giới hạn bởi chiều rộng khả dụng (trang/ô)
- Có chỉ định: cho phép phóng to, nhưng vẫn giới hạn bởi chiều rộng khả dụng
- Đoạn chỉ có ảnh: dùng style đoạn “Image” (căn giữa, giãn dòng đơn)

### Xử lý liên kết

Hỗ trợ liên kết có thể bấm trong Markdown -> DOCX:

```markdown
[Docwen](https://example.com)
[[Target]]
[[Target|Open target]]
<https://example.com>
<user@example.com>
```

- Liên kết Markdown và Wiki mặc định được ghi thành siêu liên kết Word
- Liên kết Wiki được phân giải thành liên kết cục bộ `file:///` khi tìm thấy tệp đích
- Liên kết tự động trong dấu `< >` hỗ trợ `https://...` và email `mailto:...`
- Tự động liên kết URL trần được áp dụng theo từng yêu cầu Markdown -> DOCX, mặc định tắt và được bật bằng `[non_embed_links].auto_link_bare_url` trong `configs/link.toml`
- Markdown -> XLSX không tạo placeholder siêu liên kết cho DOCX và giữ nguyên cú pháp liên kết gốc

## 📖 Hướng dẫn sử dụng chi tiết

### Word sang Markdown

1.  Kéo file `.docx` vào cửa sổ chương trình.
2.  Chương trình tự phân tích cấu trúc tài liệu.
3.  Tạo file `.md` có chứa YAML metadata.

**Định dạng hỗ trợ**:
-   `.docx` - Tài liệu Word chuẩn.
-   `.doc` - Tự chuyển sang DOCX để xử lý.
-   `.wps` - Tự chuyển tài liệu WPS để xử lý.

**Tuỳ chọn xuất**:

| Tuỳ chọn | Mô tả |
| :--- | :--- |
| **Trích xuất ảnh** | Nếu bật, ảnh trong tài liệu sẽ được xuất ra thư mục và chèn link ảnh vào file MD. |
| **OCR ảnh** | Nếu bật, chạy OCR trên ảnh và tạo file ảnh `.md` (chứa văn bản nhận dạng). |
| **Tối ưu trường nâng cao** | Nếu bật, trích xuất siêu dữ liệu có cấu trúc phong phú hơn; nếu không sẽ sử dụng chế độ đơn giản chỉ có tiêu đề và phụ đề. |
| **Dọn số tiêu đề phụ** | Nếu bật, xoá số trước tiêu đề phụ (ví dụ: "一、", "（一）", "1.", ...). |
| **Thêm số tiêu đề phụ** | Nếu bật, tự thêm số theo cấp tiêu đề (có thể cấu hình). |

Lưu ý: DOCX -> MD hiện cũng khôi phục đánh số đa cấp được liên kết trong numbering.xml thông qua kiểu đoạn (pStyle). Vì vậy, các tiền tố tiêu đề do danh sách đa cấp của Word/WPS tạo ra như "一、", "（一）", "1．", "（1）" và "①" đều được giữ lại trong cả chế độ đơn giản lẫn chế độ trường nâng cao; cấp độ tiêu đề vẫn được nhận diện chính xác khi bật tuỳ chọn "Dọn số tiêu đề phụ".

### Markdown sang Word

1.  Chuẩn bị file `.md` có YAML header.
2.  Kéo vào cửa sổ và chọn template Word tương ứng.
3.  Chương trình tự điền template và tạo tài liệu.

**Tuỳ chọn chuyển đổi**:

| Tuỳ chọn | Mô tả |
| :--- | :--- |
| **Dọn số tiêu đề phụ** | Nếu bật, xoá số trước tiêu đề phụ. |
| **Thêm số tiêu đề phụ** | Nếu bật, tự thêm số theo cấp tiêu đề. |

**Lưu ý**: Nếu có đoạn trộn giữa tiêu đề phụ và nội dung, cần giữ xuống dòng nghiêm ngặt trong file MD (xem "Xuống dòng và đoạn văn").

### Tự động xử lý style của template

Trong quá trình Markdown → DOCX, chương trình tự phát hiện và xử lý style của template:

#### Phân loại style

**Style đoạn (Paragraph Style)**: Áp dụng cho toàn bộ đoạn.

| Style | Hành vi phát hiện | Chèn khi thiếu | Nguồn |
| :--- | :--- | :--- | :--- |
| Heading (1~9) | Phát hiện style đoạn | Style heading trong template | Word built-in |
| Code Block | Phát hiện style đoạn | Font Consolas + nền xám | Định nghĩa bởi phần mềm |
| Quote (1~9) | Phát hiện style đoạn | Nền xám + viền trái | Định nghĩa bởi phần mềm |
| Formula Block | Phát hiện style đoạn | Style công thức | Định nghĩa bởi phần mềm |
| Separator (1~3) | Phát hiện style đoạn | Style đoạn có viền dưới | Định nghĩa bởi phần mềm |

**Style ký tự (Character Style)**: Áp dụng cho vùng chữ được chọn.

| Style | Hành vi phát hiện | Chèn khi thiếu | Nguồn |
| :--- | :--- | :--- | :--- |
| Inline Code | Phát hiện style ký tự | Font Consolas + shading xám | Định nghĩa bởi phần mềm |
| Inline Formula | Phát hiện style ký tự | Style công thức | Định nghĩa bởi phần mềm |

**Style bảng (Table Style)**: Áp dụng cho toàn bộ bảng.

| Style | Hành vi phát hiện | Chèn khi thiếu | Nguồn |
| :--- | :--- | :--- | :--- |
| Three-Line Table | Ưu tiên cấu hình người dùng | Định nghĩa style bảng 3 dòng | Định nghĩa bởi phần mềm |
| Grid Table | Ưu tiên cấu hình người dùng | Định nghĩa style bảng lưới | Định nghĩa bởi phần mềm |

**Định nghĩa đánh số (Numbering Definition)**: Dùng cho định dạng danh sách.

| Loại | Hành vi phát hiện | Xử lý khi thiếu |
| :--- | :--- | :--- |
| List Numbering | Quét các định nghĩa danh sách trong template | Dùng preset decimal/bullet |

#### Quốc tế hoá tên style

-   **Style Word built-in** (heading 1~9):
    -   Tên style dùng tên chuẩn tiếng Anh (ví dụ: `heading 1`).
    -   Word sẽ hiển thị tên đã được bản địa hoá theo ngôn ngữ hệ thống.
-   **Style do phần mềm định nghĩa** (Code Block, Quote, Formula, Separator, Table, ...):
    -   Chèn tên style theo ngôn ngữ giao diện của phần mềm.

**Gợi ý**: Sau khi bạn tuỳ chỉnh style trong template, chương trình sẽ ưu tiên dùng style đó; nếu template không có thì dùng preset mặc định.

### Xử lý file bảng tính

1.  **Excel/CSV sang Markdown**: Kéo file `.xlsx` hoặc `.csv` để tự chuyển sang bảng Markdown.
2.  **Markdown sang Excel**: Bảng Markdown có thể xuất sang XLSX. Template hỗ trợ trường YAML, placeholder cột và ảnh, cùng ô gộp hoặc ô được bảo vệ.

**Định dạng hỗ trợ**:
-   `.xlsx` - Excel chuẩn.
-   `.xls` - Tự chuyển sang XLSX để xử lý.
-   `.et` - Tự chuyển bảng tính WPS để xử lý.
-   `.csv` - Bảng văn bản CSV.
-   `.tsv` - Bảng phân tách bằng tab TSV.


### Chức năng kiểm tra lỗi văn bản

Chương trình cung cấp 4 quy tắc kiểm tra có thể tuỳ chỉnh:

1.  **Kiểm tra cặp dấu câu** - Kiểm tra ngoặc kép, ngoặc đơn, ... có khớp cặp không.
2.  **Kiểm tra ký hiệu** - Phát hiện dùng lẫn dấu câu tiếng Trung/tiếng Anh.
3.  **Kiểm tra lỗi chính tả** - Dựa trên từ điển tuỳ chỉnh.
4.  **Phát hiện từ nhạy cảm** - Dựa trên từ điển tuỳ chỉnh.

**Từ điển tuỳ chỉnh**: Chỉnh sửa trực quan từ điển lỗi chính tả và từ nhạy cảm trong "Cài đặt".

**Cách dùng**:
1.  Kéo tài liệu Word hoặc file Markdown cần kiểm tra vào chương trình.
2.  Chọn các quy tắc cần dùng.
3.  Nhấn nút "Text Proofreading".
4.  Kết quả hiển thị dưới dạng comment trong tài liệu. Đối với file Markdown, kết quả được xuất dưới dạng báo cáo JSON.

Ghi chú (báo cáo JSON khi kiểm tra Markdown):
- Engine: `text_rules` + adapter Markdown `md_spell`
- Đầu ra: lối vào soát lỗi hiện tại của CLI là `validate`; dùng `--json` cho envelope của CLI. `--report` là đường dẫn tệp báo cáo tùy chọn.

- Khác với `--json` (lớp bao JSON của CLI)

## 🛠️ Hệ thống template

### Dùng template có sẵn

Chương trình có sẵn nhiều template (bao gồm đa ngôn ngữ). File template nằm trong thư mục `templates/`.

### Template tuỳ chỉnh

1.  Tạo file template bằng Word hoặc WPS.
2.  Tham khảo template có sẵn và chèn placeholder như `{{Title}}`, `{{DocumentNumber}}`, ... vào vị trí cần điền.
3.  Trong template, các style Heading 1 ~ Heading 5 built-in cần chỉnh sửa thủ công.
4.  Lưu template vào thư mục `templates/`.
5.  Khởi động lại chương trình, template mới sẽ tự được tải.

Bạn cũng có thể copy một template có sẵn, chỉnh sửa và đổi tên.

### Cách dùng placeholder

#### Placeholder trong template Word

**Placeholder theo trường YAML**: Dùng dạng `{{Field Name}}` trong template, sẽ được thay thế bằng giá trị tương ứng trong YAML header của file Markdown.

| Placeholder | Mô tả |
| :--- | :--- |
| `{{Title}}` | Tiêu đề tài liệu (xem thứ tự ưu tiên bên dưới) |
| `{{Body}}` | Vị trí chèn nội dung Markdown |
| Khác | Hỗ trợ mọi trường tuỳ chỉnh |

**Thứ tự ưu tiên lấy tiêu đề**:

| Ưu tiên | Nguồn | Mô tả |
| :--- | :--- | :--- |
| 1 | YAML `Title` | Cao nhất |
| 2 | YAML `aliases` | Lấy phần tử đầu tiên của danh sách hoặc chuỗi |
| 3 | Tên file | Tên file không gồm đuôi `.md` |

**Hỗ trợ đa ngôn ngữ**: Placeholder tiêu đề và nội dung hỗ trợ nhiều ngôn ngữ, ví dụ tiêu đề `{{title}}`, `{{标题}}`, `{{Titel}}`, ...; nội dung `{{body}}`, `{{正文}}`, `{{Inhalt}}`, ...

#### Placeholder trong template Excel

Template XLSX hỗ trợ trường YAML, placeholder cột dọc `{{↓Field}}` và ngang `{{→Field}}`, placeholder ảnh, cùng ô gộp hoặc ô được bảo vệ.

**1. Placeholder theo trường YAML** `{{Field Name}}`

Điền một giá trị từ YAML header:

```markdown
---
ReportName: 2024 Annual Sales Statistics
Unit: Sales Dept
---
```

`{{ReportName}}`, `{{Unit}}` sẽ được thay bằng giá trị tương ứng. Trường Title cũng theo cùng thứ tự ưu tiên.

**2. Placeholder điền theo cột** `{{↓Field Name}}`

Trích dữ liệu từ bảng Markdown và điền **xuống dưới** theo từng dòng từ vị trí placeholder:

```markdown
| ProductName | Quantity |
|:--- |:--- |
| Product A | 100 |
| Product B | 200 |
```

`{{↓ProductName}}` sẽ được thay bằng "Product A", dòng tiếp theo điền "Product B".

**3. Placeholder điền theo hàng** `{{→Field Name}}`

Trích dữ liệu từ bảng Markdown và điền **sang phải** theo từng cột từ vị trí placeholder:

```markdown
| Month |
|:--- |
| Jan |
| Feb |
| Mar |
```

`{{→Month}}` sẽ được điền "Jan", "Feb", "Mar" sang phải.

**Xử lý ô gộp**:

- Markdown -> Excel tiếp tục giữ nguyên các merged ranges có sẵn của mẫu.
- Với các vùng mẫu dạng cột đã biết, được tạo bởi các placeholder `{{↓Field Name}}` liên tiếp, chương trình có thể khôi phục gộp hình chữ nhật từ marker `<` / `^` tường minh trong bảng Markdown.
- Chỉ những ô có nội dung sau khi bỏ khoảng trắng đầu/cuối chính xác là `<` hoặc `^` mới tham gia nhận diện gộp; `\<` và `\^` được giữ lại như văn bản literal.
- Hình chữ nhật không hợp lệ hoặc xung đột với merged ranges có sẵn của mẫu sẽ bị hạ cấp thành văn bản thường kèm cảnh báo, thay vì cưỡng ép ghi đè cấu trúc mẫu.

**Gộp dữ liệu nhiều bảng**: Nếu Markdown có nhiều bảng dùng cùng tiêu đề cột, dữ liệu sẽ được gộp theo thứ tự và điền liên tục.

## 🔌 Plugin Obsidian

Plugin Obsidian đồng hành được phát hành ở repo riêng và hoạt động cùng bộ chuyển đổi:

### Tính năng cốt lõi

-   **🚀 Khởi chạy 1 lần nhấn** - Icon ở sidebar để mở nhanh bộ chuyển đổi.
-   **📂 Bàn giao tự động** - Tự truyền đường dẫn file đang mở.
-   **🔄 Quản lý đơn phiên bản** - Nếu chương trình đang chạy, chỉ gửi file, không cần khởi chạy lại.
-   **🔒 Điều khiển cục bộ có giới hạn** - Dùng các yêu cầu có kiểu `status`, `open`, `activate`, không dò tên tiến trình và không dùng file lệnh/trạng thái.

### Nguyên lý hoạt động

Runtime/control transport của DocWen Core có thể dùng named pipe trên Windows hoặc socket AF_UNIX trên
Linux/macOS. Khóa file chỉ xác lập quyền sở hữu một phiên bản đang chạy; file không được dùng để truyền
lệnh điều khiển. Đây chỉ là mô tả capability của Core. DocWen Assistant 2.0 vẫn chỉ dành cho Windows
desktop và chưa có nghiệm thu kết hợp trên Linux/macOS.

1.  **Nhấn lần đầu** → Khởi chạy bộ chuyển đổi và truyền file hiện tại.
2.  **Nhấn lại (có file)** → Thay file mới (chế độ 1 file).
3.  **Nhấn lại (không có file)** → Kích hoạt cửa sổ bộ chuyển đổi.

### Cài đặt

DocWen Assistant 2.0 dùng DocWen Machine Protocol v1 và hợp đồng Artifact Bundle v2 duy nhất. Phiên bản mã nguồn
không chứng minh rằng sản phẩm đã được phát hành; chỉ cài bản phát hành dạng số xác định rõ một bản DocWen đã phát
hành và tương thích.

## 🔌 OpenClaw (Plugin + Skill)

OpenClaw 2.0 dùng DocWen Machine Protocol v1 và hợp đồng Artifact Bundle v2 duy nhất. Phiên bản mã nguồn không chứng
minh rằng sản phẩm đã được phát hành; hãy theo dõi trang phát hành dạng số và chỉ cài sau khi cổng phát hành bất biến
thành công.

## ❓ Câu hỏi thường gặp (FAQ)

### Nếu chuyển đổi thất bại thì sao?

-   Kiểm tra file có đang bị ứng dụng khác mở/chiếm dụng không.
-   Xác nhận định dạng file đúng.
-   Xem mục "Đường dẫn file log thực tế hiện tại" trong phần cài đặt, hoặc kiểm tra lỗi trong thư mục log người dùng của hệ thống; nếu bước kiểm thử gói dùng `DOCWEN_LOG_DIR` thì hãy kiểm tra thư mục bị ghi đè đó.

### Template không hiển thị?

-   Xác nhận file template nằm trong `templates/`.
-   Kiểm tra template có bị hỏng không.
-   Khởi động lại chương trình để tải lại template.

### Chức năng kiểm tra lỗi không hoạt động?

-   Xác nhận tài liệu là `.docx` hoặc `.md`.
-   Kiểm tra tài liệu có chứa text có thể chỉnh sửa không.
-   Xác nhận các quy tắc kiểm tra đã bật trong cài đặt.

### Định dạng đầu ra không như mong đợi?

-   Chương trình tạo tài liệu dựa trên style của template. Nếu muốn điều chỉnh đầu ra, hãy sửa trực tiếp style trong file template.
-   Template nằm trong `templates/`.
-   Sau khi sửa style, mọi tài liệu chuyển đổi với template đó sẽ áp dụng style mới.

### Ô công thức bị trống sau khi chuyển Excel sang Markdown?

Đây là hành vi dự kiến. Chương trình đọc **giá trị cache** của ô thay vì công thức.

**Lý do kỹ thuật**:
-   Ô công thức trong Excel lưu cả công thức và kết quả tính gần nhất (cache).
-   Chương trình dùng `data_only=True` nên chỉ đọc cache.
-   Nếu file chưa từng mở trong Excel hoặc chưa lưu lại sau khi tính, cache có thể trống.

**Giải pháp**:
1.  Mở file trong Excel.
2.  Đợi tính toán hoàn tất.
3.  Lưu file.
4.  Chuyển đổi lại.

## 🔒 Tính năng bảo mật

-   **Chạy hoàn toàn cục bộ**: Việc xử lý mặc định diễn ra trên máy và không phụ thuộc dịch vụ trực tuyến.
-   **Bảo vệ kết nối đi của thư viện**: Điểm vào GUI/CLI được hỗ trợ bật bộ bảo vệ audit của CPython trong toàn bộ vòng đời tiến trình Python chính. Bộ bảo vệ chặn mọi hoạt động phân giải DNS/tên và các thao tác AF_INET/AF_INET6 `bind`, `connect`, `connect_ex`, `sendto`, `sendmsg`, đồng thời giữ nguyên named pipe của Windows và Unix-domain socket.
-   **Ranh giới rõ ràng**: Tiến trình khởi chạy riêng, gồm Office/WPS/LibreOffice và Office helper chuyên dụng, không bị quản lý. Đây là lớp phòng thủ trước kết nối ngoài ý muốn của thư viện, không phải sandbox của hệ điều hành.
-   **Không upload dữ liệu**: Mặc định không chủ động tải file người dùng lên máy chủ bên ngoài.
-   **Chế độ bảo mật nghiêm ngặt**: bật mặc định; ứng dụng sẽ thoát nếu các kiểm tra bảo mật cốt lõi thất bại. Xem [Troubleshooting](../maintenance/troubleshooting.md).

## 📜 Giấy phép

Dự án dùng giấy phép **GNU Affero General Public License v3.0 (AGPL-3.0)**.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

-   Dự án sử dụng PyMuPDF (AGPL-3.0), nên toàn bộ dự án cũng theo AGPL-3.0.
- GUI hiện tại có thể sử dụng `PySide6-Fluent-Widgets` (QFluentWidgets) trên các đường host được hỗ trợ; phụ thuộc này dùng mô hình cấp phép kép `GPLv3 / thương mại`, còn kho này vẫn tiếp tục được phân phối theo AGPL.
-   Bạn có thể tự do sử dụng, sửa đổi và phân phối phần mềm.
-   Nếu bạn sửa phần mềm và cung cấp dịch vụ qua mạng, bạn phải cung cấp mã nguồn đã sửa cho người dùng.
-   Xem thêm trong [LICENSE](../../LICENSE).
- Thong bao ve thanh phan ben thu ba nam trong [LICENSE_THIRD_PARTY.txt](../../LICENSE_THIRD_PARTY.txt); tom tat phan phoi nam trong [NOTICE.txt](../../NOTICE.txt).

### Liên hệ

-   **GitHub**: https://github.com/ZHYX91/docwen
-   **Email**: zhengyx91@hotmail.com

---

**Tác giả**: ZhengYX
